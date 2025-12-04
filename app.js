// Personnel Planning Calendar Application

let calendar = null;
let gapiClient = null;
let tokenClient = null;
let accessToken = null;
let isSignedIn = false;
let allEvents = [];
let currentSelectInfo = null; // Store selected date range for event creation

// Load config from localStorage or use default
function loadConfig() {
  const savedConfig = localStorage.getItem('personnelCalendarConfig');
  if (savedConfig) {
    const parsed = JSON.parse(savedConfig);
    // Merge with default CONFIG (preserve calendarId and oauthClientId)
    
    // Migrate projects from old format (objects with colors) to new format (strings)
    if (parsed.projects) {
      CONFIG.projects = parsed.projects.map(project => {
        if (typeof project === 'string') {
          return project;
        }
        // Old format: extract just the name
        return project.name || project;
      });
    } else {
      // Migrate default CONFIG.projects if they're objects
      if (CONFIG.projects && CONFIG.projects.length > 0 && typeof CONFIG.projects[0] === 'object') {
        CONFIG.projects = CONFIG.projects.map(project => project.name || project);
      }
    }
    
    // Migrate people from old format (strings) to new format (objects with colors)
    if (parsed.people) {
      const defaultColors = ['#4285f4', '#ea4335', '#fbbc04', '#34a853', '#9c27b0', '#ff9800', '#00bcd4', '#795548', '#607d8b', '#e91e63'];
      CONFIG.people = parsed.people.map((person, index) => {
        if (typeof person === 'string') {
          // Old format: migrate to new format
          return { name: person, color: defaultColors[index % defaultColors.length] };
        }
        // New format: ensure it has name and color
        return { name: person.name || person, color: person.color || defaultColors[index % defaultColors.length] };
      });
    } else {
      // Migrate default CONFIG.people if they're strings
      if (CONFIG.people && CONFIG.people.length > 0 && typeof CONFIG.people[0] === 'string') {
        const defaultColors = ['#4285f4', '#ea4335', '#fbbc04', '#34a853', '#9c27b0', '#ff9800', '#00bcd4', '#795548', '#607d8b', '#e91e63'];
        CONFIG.people = CONFIG.people.map((person, index) => ({
          name: person,
          color: defaultColors[index % defaultColors.length]
        }));
      }
    }
  } else {
    // Migrate default CONFIG.people if they're strings
    if (CONFIG.people && CONFIG.people.length > 0 && typeof CONFIG.people[0] === 'string') {
      const defaultColors = ['#4285f4', '#ea4335', '#fbbc04', '#34a853', '#9c27b0', '#ff9800', '#00bcd4', '#795548', '#607d8b', '#e91e63'];
      CONFIG.people = CONFIG.people.map((person, index) => ({
        name: person,
        color: defaultColors[index % defaultColors.length]
      }));
    }
  }
}

// Save config to localStorage
function saveConfig() {
  const configToSave = {
    people: CONFIG.people,
    projects: CONFIG.projects
  };
  localStorage.setItem('personnelCalendarConfig', JSON.stringify(configToSave));
}

// Initialize config on load
loadConfig();

// Initialize the application
async function init() {
  showStatus('Loading Google API...', 'loading');
  
  try {
    // Wait for Google Identity Services to load
    await waitForGoogleIdentityServices();
    
    // Initialize Google API client (without auth2)
    await loadGAPI();
    
    // Initialize Google API client with discovery docs
    await gapi.client.init({
      discoveryDocs: ['https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest']
    });
    
    gapiClient = gapi.client;
    
    // Initialize OAuth 2.0 token client with new Google Identity Services
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CONFIG.oauthClientId,
      scope: 'https://www.googleapis.com/auth/calendar',
      callback: (tokenResponse) => {
        if (tokenResponse.error) {
          console.error('Token error:', tokenResponse);
          showStatus('Authentication failed: ' + tokenResponse.error, 'error');
          return;
        }
        accessToken = tokenResponse.access_token;
        // Set token on gapi client
        gapi.client.setToken({ access_token: accessToken });
        onSignInSuccess();
      },
      error_callback: (error) => {
        console.error('OAuth error:', error);
        showStatus('Authentication error: ' + (error.message || 'Unknown error'), 'error');
      }
    });
    
    // Check if we have a stored token
    const storedToken = localStorage.getItem('google_access_token');
    if (storedToken) {
      try {
        // Set token and verify it's still valid
        gapi.client.setToken({ access_token: storedToken });
        accessToken = storedToken;
        // Make a test request to verify token
        await gapi.client.calendar.calendarList.list({ maxResults: 1 });
        isSignedIn = true;
        await onSignInSuccess();
        return;
      } catch (error) {
        // Token invalid or expired, clear it
        console.log('Stored token invalid, clearing...');
        localStorage.removeItem('google_access_token');
        gapi.client.setToken(null);
      }
    }
    
    // Show sign-in button
    showSignInButton();
    showStatus('Please sign in to continue', 'info');
    
  } catch (error) {
    console.error('Initialization error:', error);
    showStatus('Error: ' + error.message, 'error');
  }
}

// Wait for Google Identity Services to load
function waitForGoogleIdentityServices() {
  return new Promise((resolve) => {
    if (window.google && window.google.accounts) {
      resolve();
      return;
    }
    
    const checkInterval = setInterval(() => {
      if (window.google && window.google.accounts) {
        clearInterval(checkInterval);
        resolve();
      }
    }, 100);
    
    // Timeout after 10 seconds
    setTimeout(() => {
      clearInterval(checkInterval);
      if (!window.google || !window.google.accounts) {
        throw new Error('Google Identity Services failed to load');
      }
      resolve();
    }, 10000);
  });
}

// Load Google API client library (without auth2)
function loadGAPI() {
  return new Promise((resolve, reject) => {
    gapi.load('client', {
      callback: resolve,
      onerror: reject
    });
  });
}

// Handle successful sign-in
async function onSignInSuccess() {
  isSignedIn = true;
  hideSignInButton();
  
  // Store token for future use
  if (accessToken) {
    localStorage.setItem('google_access_token', accessToken);
  }
  
  showStatus('Signed in successfully', 'success');
  
  // Initialize UI
  initializeUI();
  
  // Update personnel legend
  updatePersonnelLegend();
  
  // Load calendar events
  await loadEvents();
  
  showStatus('Calendar loaded successfully', 'success');
  setTimeout(() => hideStatus(), 3000);
}

// Sign in handler
function handleSignIn() {
  try {
    showStatus('Signing in...', 'loading');
    // Clear any stale tokens first
    localStorage.removeItem('google_access_token');
    gapi.client.setToken(null);
    accessToken = null;
    
    // Request access token using the new Google Identity Services
    // This will open the modern OAuth consent screen
    if (tokenClient) {
      // Use prompt: 'select_account' to force fresh login and avoid legacy flow
      tokenClient.requestAccessToken({ prompt: 'select_account' });
    } else {
      showStatus('Authentication not initialized. Please refresh the page.', 'error');
    }
  } catch (error) {
    console.error('Sign-in error:', error);
    showStatus('Sign-in failed: ' + error.message, 'error');
  }
}

// Sign out handler
function handleSignOut() {
  if (accessToken) {
    // Revoke the token
    google.accounts.oauth2.revoke(accessToken, () => {
      console.log('Token revoked');
    });
  }
  
  // Clear stored token
  localStorage.removeItem('google_access_token');
  accessToken = null;
  gapi.client.setToken(null);
  
  isSignedIn = false;
  showSignInButton();
  const signOutBtn = document.getElementById('signOutBtn');
  if (signOutBtn) {
    signOutBtn.style.display = 'none';
  }
  if (calendar) {
    calendar.destroy();
    calendar = null;
  }
  allEvents = [];
  showStatus('Signed out', 'info');
}

// Show sign-in button
function showSignInButton() {
  const signInBtn = document.getElementById('signInBtn');
  if (signInBtn) {
    signInBtn.style.display = 'block';
    signInBtn.onclick = handleSignIn;
  }
}

// Hide sign-in button
function hideSignInButton() {
  const signInBtn = document.getElementById('signInBtn');
  if (signInBtn) {
    signInBtn.style.display = 'none';
  }
}

// Load events from calendar
async function loadEvents() {
  showStatus('Loading events...', 'loading');
  
  try {
    const timeMin = new Date(new Date().getFullYear(), 0, 1).toISOString();
    const timeMax = new Date(new Date().getFullYear() + 1, 11, 31).toISOString();
    
    // Use gapi.client directly for better error handling
    const response = await gapiClient.calendar.events.list({
      calendarId: CONFIG.calendarId,
      timeMin: timeMin,
      timeMax: timeMax,
      singleEvents: true,
      orderBy: 'startTime'
    });
    
    allEvents = response.result.items || [];
    
    // Update calendar display
    updateCalendar();
    
    console.log(`Loaded ${allEvents.length} events`);
  } catch (error) {
    console.error('Error loading events:', error);
    // Check if token expired
    if (error.status === 401) {
      // Token expired, request new one
      showStatus('Session expired. Please sign in again.', 'error');
      handleSignOut();
      showSignInButton();
    } else {
      showStatus('Error loading events: ' + error.message, 'error');
    }
  }
}

// Parse event to extract person and project
function parseEvent(event) {
  const summary = event.summary || '';
  const match = summary.match(/^(.+?)\s*-\s*(.+)$/);
  
  if (match) {
    return {
      person: match[1].trim(),
      project: match[2].trim()
    };
  }
  
  return { person: '', project: summary };
}

// Get person color
function getPersonColor(personName) {
  const personConfig = CONFIG.people.find(p => p.name === personName);
  return personConfig ? personConfig.color : '#9aa0a6';
}

// Convert Google Calendar event to FullCalendar event
function toFullCalendarEvent(gcalEvent) {
  const { person, project } = parseEvent(gcalEvent);
  const color = getPersonColor(person);
  
  const start = gcalEvent.start.dateTime || gcalEvent.start.date;
  const end = gcalEvent.end.dateTime || gcalEvent.end.date;
  
  return {
    id: gcalEvent.id,
    title: `${person} - ${project}`,
    start: start,
    end: end,
    backgroundColor: color,
    borderColor: color,
    extendedProps: {
      person: person,
      project: project,
      gcalEvent: gcalEvent
    }
  };
}

// Initialize UI components
function initializeUI() {
  // Show sign-out button and hide sign-in button
  const signOutBtn = document.getElementById('signOutBtn');
  if (signOutBtn) {
    signOutBtn.style.display = 'block';
    signOutBtn.onclick = handleSignOut;
  }
  hideSignInButton();
  
  // Populate person filter
  const personFilter = document.getElementById('personFilter');
  CONFIG.people.forEach(person => {
    const option = document.createElement('option');
    option.value = person.name || person;
    option.textContent = person.name || person;
    personFilter.appendChild(option);
  });
  
  // Populate project filter
  const projectFilter = document.getElementById('projectFilter');
  CONFIG.projects.forEach(project => {
    const option = document.createElement('option');
    const projectName = project.name || project;
    option.value = projectName;
    option.textContent = projectName;
    projectFilter.appendChild(option);
  });
  
  // Initialize FullCalendar
  const calendarEl = document.getElementById('calendar');
  calendar = new FullCalendar.Calendar(calendarEl, {
    initialView: 'dayGridMonth',
    firstDay: 1, // Start week on Monday
    weekNumbers: true, // Display week numbers
    weekNumberCalculation: 'ISO', // ISO week numbering (Monday as first day)
    editable: true,
    droppable: false,
    selectable: true,
    selectMirror: true,
    dayMaxEvents: true,
    headerToolbar: {
      left: 'prev,next today',
      center: 'title',
      right: 'dayGridMonth,timeGridWeek,listYear'
    },
    events: [],
    select: handleDateSelect,
    eventDrop: handleEventDrop,
    eventResize: handleEventResize,
    eventClick: handleEventClick
  });
  
  calendar.render();
  
  // Event listeners
  document.getElementById('personFilter').addEventListener('change', updateCalendar);
  document.getElementById('projectFilter').addEventListener('change', updateCalendar);
  document.getElementById('refreshBtn').addEventListener('click', loadEvents);
  
  // Overview toggle button
  let overviewActive = false;
  document.getElementById('overviewToggleBtn').addEventListener('click', () => {
    overviewActive = !overviewActive;
    const btn = document.getElementById('overviewToggleBtn');
    if (overviewActive) {
      showCompactYearView();
      btn.textContent = 'Hide Overview';
      btn.classList.add('btn-primary');
      btn.classList.remove('btn-secondary');
    } else {
      hideCompactYearView();
      btn.textContent = 'Overview';
      btn.classList.remove('btn-primary');
      btn.classList.add('btn-secondary');
    }
  });
  
  // Modal event listeners
  document.getElementById('eventCreateBtn').addEventListener('click', handleEventCreate);
  document.getElementById('eventCancelBtn').addEventListener('click', closeEventModal);
  document.querySelector('#eventModal .modal-close').addEventListener('click', closeEventModal);
  
  // Management buttons
  document.getElementById('managePeopleBtn').addEventListener('click', () => {
    showPeopleModal();
  });
  document.getElementById('manageProjectsBtn').addEventListener('click', () => {
    showProjectsModal();
  });
  
  // People management
  document.getElementById('addPersonBtn').addEventListener('click', addPerson);
  document.getElementById('peopleCloseBtn').addEventListener('click', () => {
    document.getElementById('peopleModal').style.display = 'none';
  });
  document.querySelector('#peopleModal .modal-close').addEventListener('click', () => {
    document.getElementById('peopleModal').style.display = 'none';
  });
  
  // Projects management
  document.getElementById('addProjectBtn').addEventListener('click', addProject);
  document.getElementById('projectsCloseBtn').addEventListener('click', () => {
    document.getElementById('projectsModal').style.display = 'none';
  });
  document.querySelector('#projectsModal .modal-close').addEventListener('click', () => {
    document.getElementById('projectsModal').style.display = 'none';
  });
  
  // Close modals when clicking outside
  window.addEventListener('click', (e) => {
    const eventModal = document.getElementById('eventModal');
    const peopleModal = document.getElementById('peopleModal');
    const projectsModal = document.getElementById('projectsModal');
    
    if (e.target === eventModal) {
      closeEventModal();
    }
    if (e.target === peopleModal) {
      peopleModal.style.display = 'none';
    }
    if (e.target === projectsModal) {
      projectsModal.style.display = 'none';
    }
  });
}

// Update calendar display with filtered events
function updateCalendar() {
  const personFilter = document.getElementById('personFilter').value;
  const projectFilter = document.getElementById('projectFilter').value;
  
  let filteredEvents = allEvents.map(toFullCalendarEvent);
  
  if (personFilter) {
    filteredEvents = filteredEvents.filter(e => e.extendedProps.person === personFilter);
  }
  
  if (projectFilter) {
    filteredEvents = filteredEvents.filter(e => e.extendedProps.project === projectFilter);
  }
  
  calendar.removeAllEvents();
  calendar.addEventSource(filteredEvents);
  
  // Also update compact year view if it's active
  const overviewBtn = document.getElementById('overviewToggleBtn');
  if (overviewBtn && overviewBtn.textContent === 'Hide Overview') {
    renderCompactYearView();
  }
}

// Handle date selection (create new event)
function handleDateSelect(selectInfo) {
  currentSelectInfo = selectInfo;
  
  // Populate dropdowns
  const personSelect = document.getElementById('eventPerson');
  const projectSelect = document.getElementById('eventProject');
  
  // Clear existing options (except first)
  personSelect.innerHTML = '<option value="">Select a person...</option>';
  projectSelect.innerHTML = '<option value="">Select a project...</option>';
  
  // Populate person dropdown
  CONFIG.people.forEach(person => {
    const option = document.createElement('option');
    option.value = person.name || person;
    option.textContent = person.name || person;
    personSelect.appendChild(option);
  });
  
  // Populate project dropdown
  CONFIG.projects.forEach(project => {
    const option = document.createElement('option');
    const projectName = project.name || project;
    option.value = projectName;
    option.textContent = projectName;
    projectSelect.appendChild(option);
  });
  
  // Show modal
  document.getElementById('eventModal').style.display = 'block';
}

// Close event modal
function closeEventModal() {
  document.getElementById('eventModal').style.display = 'none';
  currentSelectInfo = null;
  if (calendar) {
    calendar.unselect();
  }
}

// Handle event creation from modal
async function handleEventCreate() {
  const person = document.getElementById('eventPerson').value;
  const project = document.getElementById('eventProject').value;
  
  if (!person || !project) {
    showStatus('Please select both person and project', 'error');
    return;
  }
  
  if (!currentSelectInfo) {
    showStatus('No date range selected', 'error');
    return;
  }
  
  try {
    await createEvent(person, project, currentSelectInfo.start, currentSelectInfo.end);
    closeEventModal();
  } catch (error) {
    console.error('Error creating event:', error);
    showStatus('Error creating event: ' + error.message, 'error');
  }
}

// Format date as YYYY-MM-DD in local timezone (not UTC)
function formatLocalDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// Create new event
async function createEvent(person, project, start, end) {
  showStatus('Creating event...', 'loading');
  
  // Use local date formatting to avoid timezone issues
  // For all-day events, Google Calendar expects dates in YYYY-MM-DD format
  // We need to use local date, not UTC, to prevent day shifts
  const startDate = formatLocalDate(start);
  // End date should be exclusive (day after last day), so add 1 day
  const endDateObj = new Date(end);
  endDateObj.setDate(endDateObj.getDate() + 1);
  const endDate = formatLocalDate(endDateObj);
  
  const event = {
    summary: `${person} - ${project}`,
    start: {
      date: startDate
    },
    end: {
      date: endDate
    }
  };
  
  try {
    const response = await gapiClient.calendar.events.insert({
      calendarId: CONFIG.calendarId,
      resource: event
    });
    
    const createdEvent = response.result;
    allEvents.push(createdEvent);
    updateCalendar();
    refreshOverviewIfVisible();
    
    showStatus('Event created successfully', 'success');
    setTimeout(() => hideStatus(), 2000);
  } catch (error) {
    console.error('Error creating event:', error);
    if (error.status === 401) {
      handleSignOut();
      showSignInButton();
      throw new Error('Session expired. Please sign in again.');
    }
    throw error;
  }
}

// Handle event drop (move event)
async function handleEventDrop(dropInfo) {
  const event = dropInfo.event;
  const gcalEvent = event.extendedProps.gcalEvent;
  
  try {
    await updateEvent(gcalEvent.id, dropInfo.event.start, dropInfo.event.end);
    await loadEvents(); // Reload events to sync with Google Calendar
  } catch (error) {
    console.error('Error updating event:', error);
    showStatus('Error updating event: ' + error.message, 'error');
    // Revert the change
    dropInfo.revert();
  }
}

// Handle event resize
async function handleEventResize(resizeInfo) {
  const event = resizeInfo.event;
  const gcalEvent = event.extendedProps.gcalEvent;
  
  try {
    await updateEvent(gcalEvent.id, resizeInfo.event.start, resizeInfo.event.end);
    await loadEvents(); // Reload events to sync with Google Calendar
    refreshOverviewIfVisible();
  } catch (error) {
    console.error('Error updating event:', error);
    showStatus('Error updating event: ' + error.message, 'error');
    resizeInfo.revert();
  }
}

// Update event
async function updateEvent(eventId, start, end) {
  showStatus('Updating event...', 'loading');
  
  // Find the original event
  const gcalEvent = allEvents.find(e => e.id === eventId);
  if (!gcalEvent) {
    throw new Error('Event not found');
  }
  
  // Use local date formatting to avoid timezone issues
  const startDate = formatLocalDate(start);
  // End date should be exclusive (day after last day), so add 1 day
  const endDateObj = new Date(end);
  endDateObj.setDate(endDateObj.getDate() + 1);
  const endDate = formatLocalDate(endDateObj);
  
  const update = {
    ...gcalEvent,
    start: {
      date: startDate
    },
    end: {
      date: endDate
    }
  };
  
  try {
    const response = await gapiClient.calendar.events.update({
      calendarId: CONFIG.calendarId,
      eventId: eventId,
      resource: update
    });
    
    const updatedEvent = response.result;
    const index = allEvents.findIndex(e => e.id === eventId);
    if (index !== -1) {
      allEvents[index] = updatedEvent;
    }
    
    updateCalendar();
    refreshOverviewIfVisible();
    
    showStatus('Event updated successfully', 'success');
    setTimeout(() => hideStatus(), 2000);
  } catch (error) {
    console.error('Error updating event:', error);
    if (error.status === 401) {
      handleSignOut();
      showSignInButton();
      throw new Error('Session expired. Please sign in again.');
    }
    throw error;
  }
}

// Handle event click (delete)
async function handleEventClick(clickInfo) {
  if (confirm('Delete this event?')) {
    const eventId = clickInfo.event.extendedProps.gcalEvent.id;
    
    try {
      await deleteEvent(eventId);
    } catch (error) {
      console.error('Error deleting event:', error);
      showStatus('Error deleting event: ' + error.message, 'error');
    }
  }
}

// Delete event
async function deleteEvent(eventId) {
  showStatus('Deleting event...', 'loading');
  
  try {
    await gapiClient.calendar.events.delete({
      calendarId: CONFIG.calendarId,
      eventId: eventId
    });
    
    allEvents = allEvents.filter(e => e.id !== eventId);
    updateCalendar();
    
    showStatus('Event deleted successfully', 'success');
    setTimeout(() => hideStatus(), 2000);
  } catch (error) {
    console.error('Error deleting event:', error);
    if (error.status === 401) {
      handleSignOut();
      showSignInButton();
      throw new Error('Session expired. Please sign in again.');
    }
    throw error;
  }
}

// Status message helpers
function showStatus(message, type = 'loading') {
  const statusEl = document.getElementById('status');
  statusEl.textContent = message;
  statusEl.className = `status ${type}`;
  statusEl.style.display = 'block';
}

function hideStatus() {
  document.getElementById('status').style.display = 'none';
}

// Show people management modal
function showPeopleModal() {
  updatePeopleList();
  document.getElementById('peopleModal').style.display = 'block';
}

// Update people list display
function updatePeopleList() {
  const peopleList = document.getElementById('peopleList');
  peopleList.innerHTML = '';
  
  if (CONFIG.people.length === 0) {
    peopleList.innerHTML = '<p style="color: #999; padding: 10px;">No people added yet.</p>';
    return;
  }
  
  CONFIG.people.forEach((person, index) => {
    const personName = person.name || person;
    const personColor = person.color || '#9aa0a6';
    const item = document.createElement('div');
    item.className = 'item-list-item';
    item.innerHTML = `
      <input type="color" class="person-color-picker" value="${personColor}" data-index="${index}" style="width: 30px; height: 30px; border: none; cursor: pointer; border-radius: 4px;">
      <span class="item-name">${personName}</span>
      <div class="item-actions">
        <button class="btn btn-small btn-secondary" onclick="removePerson(${index})">Remove</button>
      </div>
    `;
    peopleList.appendChild(item);
    
    // Add color change handler
    const colorPicker = item.querySelector('.person-color-picker');
    colorPicker.addEventListener('change', (e) => {
      CONFIG.people[index].color = e.target.value;
      saveConfig();
      updateCalendar();
      if (document.getElementById('overviewToggleBtn')?.textContent === 'Hide Overview') {
        renderCompactYearView();
      }
    });
  });
}

// Add person
function addPerson() {
  const nameInput = document.getElementById('newPersonName');
  const colorInput = document.getElementById('newPersonColor');
  const name = nameInput.value.trim();
  const color = colorInput.value;
  
  if (!name) {
    showStatus('Please enter a person name', 'error');
    return;
  }
  
  // Check if person already exists (handle both old string format and new object format)
  const exists = CONFIG.people.some(p => (p.name || p) === name);
  if (exists) {
    showStatus('Person already exists', 'error');
    return;
  }
  
  CONFIG.people.push({ name: name, color: color });
  saveConfig();
  updatePeopleList();
  updateFilters();
  updatePersonnelLegend();
  
  // Re-render compact year view if it's currently visible
  const compactYearView = document.getElementById('compactYearView');
  if (compactYearView && compactYearView.style.display !== 'none') {
    renderCompactYearView();
  }
  
  nameInput.value = '';
  colorInput.value = '#4285f4';
  showStatus('Person added successfully', 'success');
  setTimeout(() => hideStatus(), 2000);
}

// Remove person
function removePerson(index) {
  const person = CONFIG.people[index];
  const personName = person.name || person;
  if (confirm(`Remove "${personName}"?`)) {
    CONFIG.people.splice(index, 1);
    saveConfig();
    updatePeopleList();
    updateFilters();
    updatePersonnelLegend();
    
    // Re-render compact year view if it's currently visible
    const compactYearView = document.getElementById('compactYearView');
    if (compactYearView && compactYearView.style.display !== 'none') {
      renderCompactYearView();
    }
    
    showStatus('Person removed', 'success');
    setTimeout(() => hideStatus(), 2000);
  }
}

// Show projects management modal
function showProjectsModal() {
  updateProjectsList();
  document.getElementById('projectsModal').style.display = 'block';
}

// Update projects list display
function updateProjectsList() {
  const projectsList = document.getElementById('projectsList');
  projectsList.innerHTML = '';
  
  if (CONFIG.projects.length === 0) {
    projectsList.innerHTML = '<p style="color: #999; padding: 10px;">No projects added yet.</p>';
    return;
  }
  
  CONFIG.projects.forEach((project, index) => {
    const projectName = project.name || project;
    const item = document.createElement('div');
    item.className = 'item-list-item';
    item.innerHTML = `
      <span class="item-name">${projectName}</span>
      <div class="item-actions">
        <button class="btn btn-small btn-secondary" onclick="removeProject(${index})">Remove</button>
      </div>
    `;
    projectsList.appendChild(item);
  });
}

// Add project
function addProject() {
  const nameInput = document.getElementById('newProjectName');
  const name = nameInput.value.trim();
  
  if (!name) {
    showStatus('Please enter a project name', 'error');
    return;
  }
  
  // Check if project already exists (handle both old object format and new string format)
  const exists = CONFIG.projects.some(p => (p.name || p) === name);
  if (exists) {
    showStatus('Project already exists', 'error');
    return;
  }
  
  CONFIG.projects.push(name);
  saveConfig();
  updateProjectsList();
  updateFilters();
  nameInput.value = '';
  showStatus('Project added successfully', 'success');
  setTimeout(() => hideStatus(), 2000);
}

// Remove project
function removeProject(index) {
  const project = CONFIG.projects[index];
  const projectName = project.name || project;
  if (confirm(`Remove "${projectName}"?`)) {
    CONFIG.projects.splice(index, 1);
    saveConfig();
    updateProjectsList();
    updateFilters();
    showStatus('Project removed', 'success');
    setTimeout(() => hideStatus(), 2000);
  }
}

// Update filter dropdowns
function updatePersonnelLegend() {
  const legendContainer = document.getElementById('personnelLegend');
  if (!legendContainer) return;
  
  const legendItems = CONFIG.people.map(person => {
    const personName = person.name || person;
    const personColor = person.color || '#9aa0a6';
    return `<div class="personnel-legend-item">
      <div class="personnel-legend-color" style="background-color: ${personColor};"></div>
      <span class="personnel-legend-name">${personName}</span>
    </div>`;
  }).join('');
  
  legendContainer.innerHTML = legendItems;
}

function updateFilters() {
  // Update person filter
  const personFilter = document.getElementById('personFilter');
  const currentPersonValue = personFilter.value;
  personFilter.innerHTML = '<option value="">All People</option>';
  CONFIG.people.forEach(person => {
    const personName = person.name || person;
    const option = document.createElement('option');
    option.value = personName;
    option.textContent = personName;
    if (personName === currentPersonValue) {
      option.selected = true;
    }
    personFilter.appendChild(option);
  });
  
  // Update project filter
  const projectFilter = document.getElementById('projectFilter');
  const currentProjectValue = projectFilter.value;
  projectFilter.innerHTML = '<option value="">All Projects</option>';
  CONFIG.projects.forEach(project => {
    const option = document.createElement('option');
    const projectName = project.name || project;
    option.value = projectName;
    option.textContent = projectName;
    if (projectName === currentProjectValue) {
      option.selected = true;
    }
    projectFilter.appendChild(option);
  });
  
  // Update calendar display
  updateCalendar();
}

// Make remove functions globally accessible
window.removePerson = removePerson;
window.removeProject = removeProject;

// Refresh overview if it's currently visible
function refreshOverviewIfVisible() {
  const compactYearView = document.getElementById('compactYearView');
  const overviewBtn = document.getElementById('overviewToggleBtn');
  if (compactYearView && overviewBtn && overviewBtn.textContent === 'Hide Overview') {
    renderCompactYearView();
  }
}

// Compact Year View - Months as columns, weeks as rows
function showCompactYearView() {
  const calendarEl = document.getElementById('calendar');
  const compactViewEl = document.getElementById('compactYearView');
  
  if (calendarEl) {
    calendarEl.style.display = 'none';
  }
  if (compactViewEl) {
    compactViewEl.style.display = 'block';
    compactViewEl.style.visibility = 'visible';
    renderCompactYearView();
  }
}

function hideCompactYearView() {
  const calendarEl = document.getElementById('calendar');
  const compactViewEl = document.getElementById('compactYearView');
  
  if (compactViewEl) {
    compactViewEl.style.display = 'none';
    compactViewEl.style.visibility = 'hidden';
    compactViewEl.innerHTML = ''; // Clear content to prevent any rendering issues
  }
  if (calendarEl) {
    calendarEl.style.display = 'block';
    calendarEl.style.visibility = 'visible';
    // Force calendar to refresh/redraw
    if (calendar) {
      calendar.render();
    }
  }
}

// Calculate ISO week number
function getISOWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function renderCompactYearView() {
  const container = document.getElementById('compactYearView');
  const currentYear = calendar ? calendar.getDate().getFullYear() : new Date().getFullYear();
  const nextYear = currentYear + 1;
  
  // Get filtered events
  const personFilter = document.getElementById('personFilter').value;
  const projectFilter = document.getElementById('projectFilter').value;
  
  let filteredEvents = allEvents.map(toFullCalendarEvent);
  
  if (personFilter) {
    filteredEvents = filteredEvents.filter(e => e.extendedProps.person === personFilter);
  }
  
  if (projectFilter) {
    filteredEvents = filteredEvents.filter(e => e.extendedProps.project === projectFilter);
  }
  
  // Create events map by date
  const eventsByDate = {};
  filteredEvents.forEach(event => {
    const start = new Date(event.start);
    const end = new Date(event.end);
    const current = new Date(start);
    
    while (current <= end) {
      const dateKey = current.toISOString().split('T')[0];
      if (!eventsByDate[dateKey]) {
        eventsByDate[dateKey] = [];
      }
      eventsByDate[dateKey].push(event);
      current.setDate(current.getDate() + 1);
    }
  });
  
  // Create person assignments map by date for bar rendering
  const personAssignmentsByDate = {};
  filteredEvents.forEach(event => {
    const start = new Date(event.start);
    const end = new Date(event.end);
    const current = new Date(start);
    const person = event.extendedProps.person;
    
    while (current <= end) {
      const dateKey = current.toISOString().split('T')[0];
      if (!personAssignmentsByDate[dateKey]) {
        personAssignmentsByDate[dateKey] = new Set();
      }
      personAssignmentsByDate[dateKey].add(person);
      current.setDate(current.getDate() + 1);
    }
  });
  
  // Determine slot assignment based on chronological first appearance
  // Scan all dates chronologically, assign slots 0, 1, 2... as persons first appear
  const personSlotMap = new Map(); // Maps person name to their slot index (0 = bottom)
  const personFirstAppearance = new Map(); // Maps person name to their first date
  let nextSlotIndex = 0;
  
  // Get all dates in chronological order
  const allDateKeys = Object.keys(personAssignmentsByDate).sort();
  
  // First pass: determine first appearance and assign slots
  allDateKeys.forEach(dateKey => {
    const personsOnDate = Array.from(personAssignmentsByDate[dateKey] || []);
    // Sort persons by CONFIG.people order for consistent ordering when multiple appear on same date
    const sortedPersons = personsOnDate.sort((a, b) => {
      const indexA = CONFIG.people.findIndex(p => (p.name || p) === a);
      const indexB = CONFIG.people.findIndex(p => (p.name || p) === b);
      return indexA - indexB;
    });
    
    sortedPersons.forEach(personName => {
      if (!personSlotMap.has(personName)) {
        // First appearance: assign next available slot (from bottom up)
        personSlotMap.set(personName, nextSlotIndex);
        personFirstAppearance.set(personName, dateKey);
        nextSlotIndex++;
      }
    });
  });
  
  // Helper function to check if a date is part of a consecutive range for a person
  function getPersonBarInfo(dateKey, personName) {
    if (!personAssignmentsByDate[dateKey] || !personAssignmentsByDate[dateKey].has(personName)) {
      return null;
    }
    
    const date = new Date(dateKey);
    const prevDate = new Date(date);
    prevDate.setDate(prevDate.getDate() - 1);
    const nextDate = new Date(date);
    nextDate.setDate(nextDate.getDate() + 1);
    
    const prevKey = prevDate.toISOString().split('T')[0];
    const nextKey = nextDate.toISOString().split('T')[0];
    
    const hasPrev = personAssignmentsByDate[prevKey] && personAssignmentsByDate[prevKey].has(personName);
    const hasNext = personAssignmentsByDate[nextKey] && personAssignmentsByDate[nextKey].has(personName);
    
    let position = 'single';
    if (hasPrev && hasNext) {
      position = 'middle';
    } else if (hasPrev) {
      position = 'end';
    } else if (hasNext) {
      position = 'start';
    }
    
    const slotIndex = personSlotMap.get(personName) ?? 999; // Use slot index based on first appearance
    
    return { position, color: getPersonColor(personName), slotIndex, personName };
  }
  
  // Generate HTML
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const dayNames = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  
  let html = `<div class="compact-year-container">
    <div class="compact-year-header">
      <h2>${currentYear} - ${nextYear}</h2>
    </div>
    <div class="compact-year-grid">`;
  
  // Render both years (24 months total)
  const years = [currentYear, nextYear];
  years.forEach(year => {
    // Create each month as a column
    for (let month = 0; month < 12; month++) {
      const firstDay = new Date(year, month, 1);
      const lastDay = new Date(year, month + 1, 0);
      const daysInMonth = lastDay.getDate();
      const firstDayOfWeek = firstDay.getDay(); // 0 = Sunday, 1 = Monday, etc.
      
      // Convert Sunday (0) to 7 for easier calculation
      const startOffset = firstDayOfWeek === 0 ? 6 : firstDayOfWeek - 1;
      
      // Always use 6 weeks for proper alignment
      const weeksInMonth = 6;
      
      html += `<div class="compact-month-column">
        <div class="compact-month-header">${monthNames[month]} ${year}</div>
        <div class="compact-month-days-header">
          <div class="compact-week-number-header">KW</div>
          ${dayNames.map((day, idx) => {
            const isWeekend = idx >= 5; // Sat (5) and Sun (6)
            return `<div class="compact-day-name ${isWeekend ? 'weekend' : ''}">${day}</div>`;
          }).join('')}
        </div>`;
      
      // Create week rows (always 6 weeks for alignment)
      for (let week = 0; week < weeksInMonth; week++) {
        // Calculate week number for the first day of this week
        const weekStartDate = new Date(year, month, week * 7 - startOffset + 1);
        const weekNumber = getISOWeekNumber(weekStartDate);
        
        html += `<div class="compact-week-row">`;
        html += `<div class="compact-week-number">${weekNumber}</div>`;
        
        // Create 7 day cells for this week
        for (let dayOfWeek = 0; dayOfWeek < 7; dayOfWeek++) {
          const cellIndex = week * 7 + dayOfWeek;
          const dayNumber = cellIndex - startOffset + 1;
          const isWeekend = dayOfWeek >= 5; // Sat (5) and Sun (6)
          
          if (dayNumber > 0 && dayNumber <= daysInMonth) {
            // Valid day in this month
            const date = new Date(year, month, dayNumber);
            const dateKey = date.toISOString().split('T')[0];
            const dayEvents = eventsByDate[dateKey] || [];
            const isToday = dateKey === new Date().toISOString().split('T')[0];
            const hasPersonnel = dayEvents.length > 0;
            
            // Get unique personnel with their colors
            const personnelList = dayEvents.map(e => e.extendedProps.person).filter(p => p);
            const uniquePersonnel = [...new Set(personnelList)];
            
            // Get bar info for each person and sort by slot index (based on first appearance)
            const personBars = uniquePersonnel
              .map(personName => getPersonBarInfo(dateKey, personName))
              .filter(b => b)
              .sort((a, b) => a.slotIndex - b.slotIndex); // Sort by slot index (chronological first appearance)
            
            // Calculate available space for bars
            // Cell is 80px high, with 2px padding top/bottom = 4px total padding
            // Reserve ~20px at top for day number to prevent overlap
            // Bars are absolutely positioned from bottom: 0
            const cellHeight = 80;
            const cellPaddingBottom = 2; // Bottom padding
            const reservedTopSpace = 20; // Space to keep clear for day number
            const availableHeight = cellHeight - cellPaddingBottom - reservedTopSpace; // 80 - 2 - 20 = 58px
            
            // Calculate bar dimensions: 1px margin/gap around each bar, make bars thicker
            const barMargin = 1; // 1px margin/gap
            const maxSlots = 15; // Maximum slots available
            const minBarHeight = 4; // Minimum bar height for visibility
            
            // Calculate how many slots we can actually fit with minimum bar height
            // Formula: availableHeight = (numSlots * barHeight) + (numSlots * gap)
            // Rearranged: numSlots = availableHeight / (barHeight + gap)
            const effectiveSlots = Math.min(maxSlots, Math.floor(availableHeight / (minBarHeight + barMargin)));
            
            // Calculate bar height based on effective slots
            const totalGapSpace = effectiveSlots * barMargin; // 1px gap for each slot
            let barHeight = Math.floor((availableHeight - totalGapSpace) / effectiveSlots); // Height per bar
            barHeight = Math.max(barHeight, minBarHeight); // Ensure minimum thickness
            
            // Use effective slots for positioning
            const slotsToUse = effectiveSlots;
            
            // Create tooltip content with person names and colors
            const personNamesHtml = personBars.map(barInfo => {
              return `<span class="compact-hover-name" style="color: ${barInfo.color}; font-weight: 600;">${barInfo.personName}</span>`;
            }).join('');
            
            html += `<div class="compact-day-cell ${isToday ? 'today' : ''} ${hasPersonnel ? 'has-personnel' : ''}" 
              data-date="${dateKey}"
              data-personnel="${personBars.map(b => b.personName).join(',')}"
              title="${date.toLocaleDateString()}">
              <div class="compact-day-number">${dayNumber}</div>
              <div class="compact-personnel-hover-tooltip">${personNamesHtml}</div>
              <div class="compact-personnel-bars">`;
            
            // Create slot array and place bars in their assigned slots based on first appearance
            // Slot 0 = first person to appear (chronologically) = bottom
            // Additional persons stack above in order of first appearance
            // Exception: If only one person on this date, always place at bottom slot (0)
            const barSlots = new Array(slotsToUse).fill(null);
            
            if (personBars.length === 1) {
              // Single person: always place at bottom slot (0) regardless of chronological slot index
              barSlots[0] = personBars[0];
            } else {
              // Multiple people: use their chronological first appearance slots
              personBars.forEach(barInfo => {
                if (barInfo.slotIndex < slotsToUse) {
                  barSlots[barInfo.slotIndex] = barInfo;
                }
              });
            }
            
            // Render bars from bottom up, ordered by slot index (chronological first appearance)
            // Each bar: bottom position = slotIdx * (barHeight + 1px gap) + 1px bottom margin
            barSlots.forEach((barInfo, slotIdx) => {
              if (barInfo) {
                const barClass = `compact-personnel-bar compact-bar-${barInfo.position}`;
                // Calculate bottom position: slot 0 at bottom with 1px margin, each slot above with bar height + 1px gap
                const bottomPosition = slotIdx * (barHeight + barMargin) + barMargin;
                html += `<div class="${barClass}" style="background-color: ${barInfo.color}; bottom: ${bottomPosition}px; height: ${barHeight}px;" data-slot="${slotIdx}"></div>`;
              }
            });
            
            html += `</div></div>`;
          } else {
            // Empty cell (before or after month)
            html += `<div class="compact-day-cell empty"></div>`;
          }
        }
        
        html += `</div>`;
      }
      
      html += `</div>`;
    }
  });
  
  html += `</div></div>`;
  
  container.innerHTML = html;
  
  // Scroll to today
  setTimeout(() => {
    const todayCell = container.querySelector('.compact-day-cell.today');
    if (todayCell) {
      todayCell.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
    }
  }, 100);
  
  // No click handlers - this is a read-only presentation view
}

// Initialize when page loads
document.addEventListener('DOMContentLoaded', init);

