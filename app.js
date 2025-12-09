// Personnel Planning Calendar Application

let calendar = null;
let gapiClient = null;
let tokenClient = null;
let accessToken = null;
let isSignedIn = false;
let allEvents = [];
let currentSelectInfo = null; // Store selected date range for event creation

// Get default config data (used when creating new config event or when not signed in)
function getDefaultConfigData() {
  return {
    personnel: [],
    projects: [],
    roles: [
      { name: 'Project-Manager', color: '#4285f4' },
      { name: 'Foreman', color: '#ea4335' },
      { name: 'Shaper', color: '#fbbc04' },
      { name: 'Operator-Shaper', color: '#34a853' }
    ]
  };
}

// Initialize config arrays (data will be loaded from calendar)
function loadConfig() {
  // Initialize empty arrays - data comes from calendar event
  CONFIG.personnel = [];
  CONFIG.projects = [];
  CONFIG.roles = [];
}

// Find or create config event in calendar
async function findOrCreateConfigEvent() {
  const CONFIG_EVENT_DATE = '2000-01-01';
  const CONFIG_EVENT_TITLE = '__PERSONNEL_CONFIG__';
  
  try {
    // Search for existing config event
    const response = await gapiClient.calendar.events.list({
      calendarId: CONFIG.calendarId,
      timeMin: CONFIG_EVENT_DATE + 'T00:00:00Z',
      timeMax: CONFIG_EVENT_DATE + 'T23:59:59Z',
      singleEvents: true,
      q: CONFIG_EVENT_TITLE
    });
    
    const events = response.result.items || [];
    const configEvent = events.find(e => e.summary === CONFIG_EVENT_TITLE);
    
    if (configEvent) {
      return configEvent;
    }
    
    // Create new config event with defaults
    const defaultData = getDefaultConfigData();
    const newEvent = {
      summary: CONFIG_EVENT_TITLE,
      start: { date: CONFIG_EVENT_DATE },
      end: { date: '2000-01-02' },
      extendedProperties: {
        shared: {
          personnelConfig: JSON.stringify(defaultData)
        }
      }
    };
    
    const createResponse = await gapiClient.calendar.events.insert({
      calendarId: CONFIG.calendarId,
      resource: newEvent
    });
    
    return createResponse.result;
  } catch (error) {
    console.error('Error finding/creating config event:', error);
    throw error;
  }
}

// Load config from calendar
async function loadConfigFromCalendar() {
  if (!isSignedIn || !gapiClient) {
    // If not signed in, use defaults
    const defaultData = getDefaultConfigData();
    CONFIG.personnel = defaultData.personnel;
    CONFIG.projects = defaultData.projects;
    CONFIG.roles = defaultData.roles;
    return;
  }
  
  try {
    const configEvent = await findOrCreateConfigEvent();
    
    if (configEvent?.extendedProperties?.shared?.personnelConfig) {
      const configData = JSON.parse(configEvent.extendedProperties.shared.personnelConfig);
      
      // Migrate roles from old format (strings) to new format (objects with colors)
      if (configData.roles && Array.isArray(configData.roles)) {
        const defaultColors = ['#4285f4', '#ea4335', '#fbbc04', '#34a853', '#9c27b0', '#ff9800', '#00bcd4', '#795548', '#607d8b', '#e91e63'];
        CONFIG.roles = configData.roles.map((role, index) => {
          if (typeof role === 'string') {
            return { name: role, color: defaultColors[index % defaultColors.length] };
          }
          return { name: role.name || role, color: role.color || defaultColors[index % defaultColors.length] };
        });
      } else {
        // Use defaults if no roles
        const defaultData = getDefaultConfigData();
        CONFIG.roles = defaultData.roles;
      }
      
      // Migrate personnel from old format (objects with colors) to new format (strings)
      // Handle both old "people" and new "personnel" keys for backward compatibility
      const personnelData = configData.personnel || configData.people;
      if (personnelData && Array.isArray(personnelData)) {
        CONFIG.personnel = personnelData.map(person => {
          return typeof person === 'string' ? person : (person.name || person);
        });
      } else {
        CONFIG.personnel = [];
      }
      
      if (configData.projects && Array.isArray(configData.projects)) {
        CONFIG.projects = configData.projects.map(project => {
          return typeof project === 'string' ? project : (project.name || project);
        });
      } else {
        CONFIG.projects = [];
      }
    } else {
      // No config data, use defaults
      const defaultData = getDefaultConfigData();
      CONFIG.personnel = defaultData.personnel;
      CONFIG.projects = defaultData.projects;
      CONFIG.roles = defaultData.roles;
    }
  } catch (error) {
    console.error('Error loading config from calendar:', error);
    // Fall back to defaults
    const defaultData = getDefaultConfigData();
    CONFIG.personnel = defaultData.personnel;
    CONFIG.projects = defaultData.projects;
    CONFIG.roles = defaultData.roles;
  }
}

// Save config to calendar
async function saveConfigToCalendar() {
  if (!isSignedIn || !gapiClient) {
    return; // Can't save to calendar if not signed in
  }
  
  try {
    const configEvent = await findOrCreateConfigEvent();
    
    const configData = {
      personnel: CONFIG.personnel,
      projects: CONFIG.projects,
      roles: CONFIG.roles
    };
    
    await gapiClient.calendar.events.update({
      calendarId: CONFIG.calendarId,
      eventId: configEvent.id,
      resource: {
        ...configEvent,
        extendedProperties: {
          shared: {
            personnelConfig: JSON.stringify(configData)
          }
        }
      }
    });
  } catch (error) {
    console.error('Error saving config to calendar:', error);
    throw error;
  }
}

// Save config to calendar
async function saveConfig() {
  // Only save to calendar - no localStorage
  if (isSignedIn && gapiClient) {
    try {
      await saveConfigToCalendar();
    } catch (error) {
      console.error('Failed to save config to calendar:', error);
      throw error;
    }
  } else {
    console.warn('Cannot save config: not signed in');
  }
}

// Initialize config on load
loadConfig();

// Load defaults if not signed in (will be overridden when signed in)
if (!isSignedIn) {
  const defaultData = getDefaultConfigData();
  CONFIG.personnel = defaultData.personnel;
  CONFIG.projects = defaultData.projects;
  CONFIG.roles = defaultData.roles;
}

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
        // Token invalid or expired, clear it
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
  
  // Load config from calendar (will override localStorage)
  try {
    await loadConfigFromCalendar();
  } catch (error) {
    console.error('Error loading config from calendar:', error);
  }
  
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
      // Token revoked
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
    
    // Filter out config event
    allEvents = (response.result.items || []).filter(e => 
      !e.summary || e.summary !== '__PERSONNEL_CONFIG__'
    );
    
    // Update calendar display
    updateCalendar();
    
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

// Parse event to extract person, project, and role
function parseEvent(event) {
  const summary = event.summary || '';
  // Try to match "Person - Project - Role" format
  const matchWithRole = summary.match(/^(.+?)\s*-\s*(.+?)\s*-\s*(.+)$/);
  
  if (matchWithRole) {
    return {
      person: matchWithRole[1].trim(),
      project: matchWithRole[2].trim(),
      role: matchWithRole[3].trim()
    };
  }
  
  // Fallback to old format "Person - Project" (backward compatibility)
  const match = summary.match(/^(.+?)\s*-\s*(.+)$/);
  if (match) {
    return {
      person: match[1].trim(),
      project: match[2].trim(),
      role: '' // No role in old format
    };
  }
  
  return { person: '', project: summary, role: '' };
}

// Get role color
function getRoleColor(roleName) {
  const roleConfig = CONFIG.roles.find(r => (r.name || r) === roleName);
  return roleConfig ? (roleConfig.color || '#9aa0a6') : '#9aa0a6';
}

// Check if person, project, or role exists in CONFIG
function isValidEvent(person, project, role) {
  // If person is empty, consider it valid (might be old format event)
  // If person has a value, it must exist in CONFIG
  const personExists = !person || person.trim() === '' || CONFIG.personnel.some(p => {
    const personName = typeof p === 'string' ? p : p.name;
    return personName === person;
  });
  
  // If project is empty, consider it valid (might be old format event)
  // If project has a value, it must exist in CONFIG
  const projectExists = !project || project.trim() === '' || CONFIG.projects.some(p => {
    const projectName = typeof p === 'string' ? p : p.name;
    return projectName === project;
  });
  
  // If role is empty, consider it valid (might be old format event)
  // If role has a value, it must exist in CONFIG
  const roleExists = !role || role.trim() === '' || CONFIG.roles.some(r => {
    const roleName = typeof r === 'string' ? r : r.name;
    return roleName === role;
  });
  
  return personExists && projectExists && roleExists;
}

// Convert Google Calendar event to FullCalendar event
function toFullCalendarEvent(gcalEvent) {
  const { person, project, role } = parseEvent(gcalEvent);
  
  // Check if event references removed items - if so, use grey color
  const isValid = isValidEvent(person, project, role);
  const color = isValid ? getRoleColor(role) : '#9aa0a6';
  
  const start = gcalEvent.start.dateTime || gcalEvent.start.date;
  const end = gcalEvent.end.dateTime || gcalEvent.end.date;
  
  const title = role ? `${person} - ${project} - ${role}` : `${person} - ${project}`;
  
  return {
    id: gcalEvent.id,
    title: title,
    start: start,
    end: end,
    backgroundColor: color,
    borderColor: color,
    extendedProps: {
      person: person,
      project: project,
      role: role,
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
  CONFIG.personnel.forEach(person => {
    const option = document.createElement('option');
    const personName = typeof person === 'string' ? person : person.name;
    option.value = personName;
    option.textContent = personName;
    personFilter.appendChild(option);
  });
  
  // Populate project filter
  const projectFilter = document.getElementById('projectFilter');
  CONFIG.projects.forEach(project => {
    const option = document.createElement('option');
    const projectName = typeof project === 'string' ? project : project.name;
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
    headerToolbar: false, // No navigation buttons - month view only
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
      btn.textContent = 'Scheduling';
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
  const manageRolesBtn = document.getElementById('manageRolesBtn');
  if (manageRolesBtn) {
    manageRolesBtn.addEventListener('click', () => {
      showRolesModal();
    });
  }
  
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
  
  // Roles management
  const addRoleBtn = document.getElementById('addRoleBtn');
  const rolesCloseBtn = document.getElementById('rolesCloseBtn');
  const rolesModalClose = document.querySelector('#rolesModal .modal-close');
  if (addRoleBtn) {
    addRoleBtn.addEventListener('click', addRole);
  }
  if (rolesCloseBtn) {
    rolesCloseBtn.addEventListener('click', () => {
      document.getElementById('rolesModal').style.display = 'none';
    });
  }
  if (rolesModalClose) {
    rolesModalClose.addEventListener('click', () => {
      document.getElementById('rolesModal').style.display = 'none';
    });
  }
  
  // Close modals when clicking outside
  window.addEventListener('click', (e) => {
    const eventModal = document.getElementById('eventModal');
    const peopleModal = document.getElementById('peopleModal');
    const projectsModal = document.getElementById('projectsModal');
    const rolesModal = document.getElementById('rolesModal');
    
    if (e.target === eventModal) {
      closeEventModal();
    }
    if (e.target === peopleModal) {
      peopleModal.style.display = 'none';
    }
    if (e.target === projectsModal) {
      projectsModal.style.display = 'none';
    }
    if (rolesModal && e.target === rolesModal) {
      rolesModal.style.display = 'none';
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
  if (overviewBtn && overviewBtn.textContent === 'Scheduling') {
    renderCompactYearView();
  }
}

// Handle date selection (create new event)
function handleDateSelect(selectInfo) {
  currentSelectInfo = selectInfo;
  
  // Populate dropdowns
  const personSelect = document.getElementById('eventPerson');
  const projectSelect = document.getElementById('eventProject');
  const roleSelect = document.getElementById('eventRole');
  
  if (!personSelect || !projectSelect || !roleSelect) {
    console.error('Modal dropdown elements not found');
    return;
  }
  
  // Clear existing options (except first)
  personSelect.innerHTML = '<option value="">Select personnel...</option>';
  projectSelect.innerHTML = '<option value="">Select a project...</option>';
  roleSelect.innerHTML = '<option value="">Select a role...</option>';
  
  // Populate person dropdown
  if (CONFIG.personnel && CONFIG.personnel.length > 0) {
    CONFIG.personnel.forEach(person => {
      const option = document.createElement('option');
      const personName = typeof person === 'string' ? person : person.name;
      option.value = personName;
      option.textContent = personName;
      personSelect.appendChild(option);
    });
  }
  
  // Populate project dropdown
  if (CONFIG.projects && CONFIG.projects.length > 0) {
    CONFIG.projects.forEach(project => {
      const option = document.createElement('option');
      const projectName = typeof project === 'string' ? project : project.name;
      option.value = projectName;
      option.textContent = projectName;
      projectSelect.appendChild(option);
    });
  }
  
  // Populate role dropdown
  if (CONFIG.roles && CONFIG.roles.length > 0) {
    if (roleSelect) {
      roleSelect.innerHTML = '<option value="">Select a role...</option>';
      
      CONFIG.roles.forEach(role => {
        const roleName = typeof role === 'string' ? role : role.name;
        if (roleName && roleName.trim()) {
          const option = document.createElement('option');
          option.value = roleName;
          option.textContent = roleName;
          roleSelect.appendChild(option);
        }
      });
    }
  }
  
  // Show modal
  const eventModal = document.getElementById('eventModal');
  if (eventModal) {
    eventModal.style.display = 'block';
    
    // Double-check role dropdown after modal is shown
    setTimeout(() => {
      const roleSelectCheck = document.getElementById('eventRole');
      if (roleSelectCheck && roleSelectCheck.options.length <= 1) {
        // Repopulate if still empty
        roleSelectCheck.innerHTML = '<option value="">Select a role...</option>';
        const rolesToAdd = CONFIG.roles && CONFIG.roles.length > 0 
          ? CONFIG.roles 
          : ['Project-Manager', 'Foreman', 'Shaper', 'Operator-Shaper'];
        rolesToAdd.forEach(role => {
          if (role && role.trim()) {
            const option = document.createElement('option');
            option.value = role;
            option.textContent = role;
            roleSelectCheck.appendChild(option);
          }
        });
      }
    }, 100);
  }
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
  const role = document.getElementById('eventRole').value;
  
  if (!person || !project || !role) {
    showStatus('Please select personnel, project, and role', 'error');
    return;
  }
  
  if (!currentSelectInfo) {
    showStatus('No date range selected', 'error');
    return;
  }
  
  try {
    await createEvent(person, project, role, currentSelectInfo.start, currentSelectInfo.end);
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
async function createEvent(person, project, role, start, end) {
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
    summary: `${person} - ${project} - ${role}`,
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
  
  if (CONFIG.personnel.length === 0) {
    peopleList.innerHTML = '<p style="color: #999; padding: 10px;">No personnel added yet.</p>';
    return;
  }
  
  CONFIG.personnel.forEach((person, index) => {
    const personName = typeof person === 'string' ? person : person.name;
    const item = document.createElement('div');
    item.className = 'item-list-item';
    item.innerHTML = `
      <span class="item-name">${personName}</span>
      <div class="item-actions">
        <button class="btn btn-small btn-secondary" onclick="removePerson(${index})">Remove</button>
      </div>
    `;
    peopleList.appendChild(item);
  });
}

// Add person
function addPerson() {
  const nameInput = document.getElementById('newPersonName');
  const name = nameInput.value.trim();
  
  if (!name) {
    showStatus('Please enter a personnel name', 'error');
    return;
  }
  
  // Check if person already exists
  const exists = CONFIG.personnel.some(p => (typeof p === 'string' ? p : p.name) === name);
  if (exists) {
    showStatus('Personnel already exists', 'error');
    return;
  }
  
  CONFIG.personnel.push(name);
  saveConfig();
  updatePeopleList();
  updateFilters();
  
  // Re-render compact year view if it's currently visible
  const compactYearView = document.getElementById('compactYearView');
  if (compactYearView && compactYearView.style.display !== 'none') {
    renderCompactYearView();
  }
  
  nameInput.value = '';
  showStatus('Personnel added successfully', 'success');
  setTimeout(() => hideStatus(), 2000);
}

// Remove person
function removePerson(index) {
  const person = CONFIG.personnel[index];
  const personName = typeof person === 'string' ? person : person.name;
  if (confirm(`Remove "${personName}"?`)) {
    CONFIG.personnel.splice(index, 1);
    saveConfig();
    updatePeopleList();
    updateFilters();
    
    // Re-render compact year view if it's currently visible
    const compactYearView = document.getElementById('compactYearView');
    if (compactYearView && compactYearView.style.display !== 'none') {
      renderCompactYearView();
    }
    
    showStatus('Personnel removed', 'success');
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
    const projectName = typeof project === 'string' ? project : project.name;
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
  
  // Check if project already exists
  const exists = CONFIG.projects.some(p => (typeof p === 'string' ? p : p.name) === name);
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
  const projectName = typeof project === 'string' ? project : project.name;
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
  
  const legendItems = CONFIG.roles.map(role => {
    const roleName = typeof role === 'string' ? role : role.name;
    const roleColor = typeof role === 'string' ? '#9aa0a6' : (role.color || '#9aa0a6');
    return `<div class="personnel-legend-item">
      <div class="personnel-legend-color" style="background-color: ${roleColor};"></div>
      <span class="personnel-legend-name">${roleName}</span>
    </div>`;
  }).join('');
  
  legendContainer.innerHTML = legendItems;
}

function updateFilters() {
  // Update person filter
  const personFilter = document.getElementById('personFilter');
  const currentPersonValue = personFilter.value;
  personFilter.innerHTML = '<option value="">All Personnel</option>';
  CONFIG.personnel.forEach(person => {
    const personName = typeof person === 'string' ? person : person.name;
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
    const projectName = typeof project === 'string' ? project : project.name;
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

// Show roles management modal
function showRolesModal() {
  updateRolesList();
  document.getElementById('rolesModal').style.display = 'block';
}

// Update roles list display
function updateRolesList() {
  const rolesList = document.getElementById('rolesList');
  rolesList.innerHTML = '';
  
  if (CONFIG.roles.length === 0) {
    rolesList.innerHTML = '<p style="color: #999; padding: 10px;">No roles added yet.</p>';
    return;
  }
  
  CONFIG.roles.forEach((role, index) => {
    const roleName = typeof role === 'string' ? role : role.name;
    const roleColor = typeof role === 'string' ? '#9aa0a6' : (role.color || '#9aa0a6');
    const item = document.createElement('div');
    item.className = 'item-list-item';
    item.innerHTML = `
      <input type="color" class="role-color-picker" value="${roleColor}" data-index="${index}" style="width: 30px; height: 30px; border: none; cursor: pointer; border-radius: 4px;">
      <span class="item-name">${roleName}</span>
      <div class="item-actions">
        <button class="btn btn-small btn-secondary" onclick="removeRole(${index})">Remove</button>
      </div>
    `;
    rolesList.appendChild(item);
    
    // Add color change handler
    const colorPicker = item.querySelector('.role-color-picker');
    colorPicker.addEventListener('change', (e) => {
      if (typeof CONFIG.roles[index] === 'string') {
        CONFIG.roles[index] = { name: CONFIG.roles[index], color: e.target.value };
      } else {
        CONFIG.roles[index].color = e.target.value;
      }
      saveConfig();
      updateCalendar();
      updatePersonnelLegend();
      if (document.getElementById('overviewToggleBtn')?.textContent === 'Scheduling') {
        renderCompactYearView();
      }
    });
  });
}

// Add role
function addRole() {
  const nameInput = document.getElementById('newRoleName');
  const colorInput = document.getElementById('newRoleColor');
  const name = nameInput.value.trim();
  const color = colorInput.value;
  
  if (!name) {
    showStatus('Please enter a role name', 'error');
    return;
  }
  
  // Check if role already exists
  const exists = CONFIG.roles.some(r => (typeof r === 'string' ? r : r.name) === name);
  if (exists) {
    showStatus('Role already exists', 'error');
    return;
  }
  
  CONFIG.roles.push({ name: name, color: color });
  saveConfig();
  updateRolesList();
  updatePersonnelLegend();
  
  // Re-render compact year view if it's currently visible
  const compactYearView = document.getElementById('compactYearView');
  if (compactYearView && compactYearView.style.display !== 'none') {
    renderCompactYearView();
  }
  
  nameInput.value = '';
  colorInput.value = '#4285f4';
  showStatus('Role added successfully', 'success');
  setTimeout(() => hideStatus(), 2000);
}

// Remove role
function removeRole(index) {
  const role = CONFIG.roles[index];
  const roleName = typeof role === 'string' ? role : role.name;
  if (confirm(`Remove "${roleName}"?`)) {
    CONFIG.roles.splice(index, 1);
    saveConfig();
    updateRolesList();
    updatePersonnelLegend();
    
    // Re-render compact year view if it's currently visible
    const compactYearView = document.getElementById('compactYearView');
    if (compactYearView && compactYearView.style.display !== 'none') {
      renderCompactYearView();
    }
    
    showStatus('Role removed', 'success');
    setTimeout(() => hideStatus(), 2000);
  }
}

// Make remove functions globally accessible
window.removePerson = removePerson;
window.removeProject = removeProject;
window.removeRole = removeRole;

// Refresh overview if it's currently visible
function refreshOverviewIfVisible() {
  const compactYearView = document.getElementById('compactYearView');
  const overviewBtn = document.getElementById('overviewToggleBtn');
  if (compactYearView && overviewBtn && overviewBtn.textContent === 'Scheduling') {
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
  
  // Group events by project, then by person+role combination
  // Only include events with valid person, project, and role
  const eventsByProjectPersonRole = {};
  filteredEvents.forEach(event => {
    const project = event.extendedProps.project || '';
    const person = event.extendedProps.person || '';
    const role = event.extendedProps.role || '';
    
    // Skip events without valid person, project, or role
    if (!project || !person || !role || project.trim() === '' || person.trim() === '' || role.trim() === '' || project === 'Unassigned') {
      return;
    }
    
    const personRoleKey = `${person}|||${role}`;
    
    if (!eventsByProjectPersonRole[project]) {
      eventsByProjectPersonRole[project] = {};
    }
    if (!eventsByProjectPersonRole[project][personRoleKey]) {
      eventsByProjectPersonRole[project][personRoleKey] = {
        person: person,
        role: role,
        events: []
      };
    }
    eventsByProjectPersonRole[project][personRoleKey].events.push(event);
  });
  
  // Create a map of all dates for each person+role in each project
  const personRoleDatesByProject = {};
  // Build a map to detect conflicts: dateKey -> Map of personnel -> count
  const personnelCountByDate = {};
  
  Object.keys(eventsByProjectPersonRole).forEach(project => {
    personRoleDatesByProject[project] = {};
    Object.keys(eventsByProjectPersonRole[project]).forEach(personRoleKey => {
      const { person, role, events } = eventsByProjectPersonRole[project][personRoleKey];
      const dateSet = new Set();
      
      events.forEach(event => {
        const start = new Date(event.start);
        const end = new Date(event.end);
        const current = new Date(start);
        
        while (current <= end) {
          // Use local date formatting to avoid timezone issues
          const dateKey = formatLocalDate(current);
          dateSet.add(dateKey);
          
          // Track personnel assignments per date for conflict detection
          if (!personnelCountByDate[dateKey]) {
            personnelCountByDate[dateKey] = {};
          }
          if (!personnelCountByDate[dateKey][person]) {
            personnelCountByDate[dateKey][person] = 0;
          }
          personnelCountByDate[dateKey][person]++;
          
          current.setDate(current.getDate() + 1);
        }
      });
      
      personRoleDatesByProject[project][personRoleKey] = {
        person,
        role,
        dates: Array.from(dateSet).sort()
      };
    });
  });
  
  // Find dates with conflicts (same personnel assigned more than once)
  const conflictedDates = new Set();
  Object.keys(personnelCountByDate).forEach(dateKey => {
    const personnelCounts = personnelCountByDate[dateKey];
    // Check if any personnel appears more than once on this date
    Object.keys(personnelCounts).forEach(person => {
      if (personnelCounts[person] > 1) {
        conflictedDates.add(dateKey);
      }
    });
  });
  
  // Get projects that have valid person+role assignments
  const projectsWithAssignments = new Set();
  Object.keys(personRoleDatesByProject).forEach(project => {
    const combos = personRoleDatesByProject[project] || {};
    const validKeys = Object.keys(combos).filter(key => {
      const combo = combos[key];
      return combo && combo.person && combo.role && combo.dates && combo.dates.length > 0;
    });
    if (validKeys.length > 0) {
      projectsWithAssignments.add(project);
    }
  });
  
  // Get projects to show: ONLY show projects with valid assignments (no empty rows from CONFIG)
  let projectsToShow = projectFilter 
    ? [projectFilter]
    : Array.from(projectsWithAssignments).sort();
  
  // Remove duplicates and filter invalid
  projectsToShow = [...new Set(projectsToShow)].filter(p => p && p.trim() !== '' && p !== '...');
  
  
  // Generate HTML
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const dayNames = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
  
  // Generate all days for current year and next year
  const allDays = [];
  const years = [currentYear, nextYear];
  years.forEach(year => {
    for (let month = 0; month < 12; month++) {
      const lastDay = new Date(year, month + 1, 0);
      const daysInMonth = lastDay.getDate();
      for (let day = 1; day <= daysInMonth; day++) {
        const date = new Date(year, month, day);
        // Use local date formatting to avoid timezone issues
        const dateKey = formatLocalDate(date);
        const dayOfWeek = date.getDay();
        const dayName = dayNames[dayOfWeek === 0 ? 6 : dayOfWeek - 1];
        const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
        const weekNumber = getISOWeekNumber(date);
        
        allDays.push({
          date,
          dateKey,
          day,
          month,
          year,
          dayName,
          isWeekend,
          weekNumber,
          monthName: monthNames[month]
        });
      }
    }
  });
  
  // Build proper table structure like CSV
  let html = `<div class="project-overview-container">
    <div class="project-overview-header">
      <h2>${currentYear} - ${nextYear}</h2>
      <button id="scrollToTodayBtn" class="btn btn-secondary" style="margin-left: 10px;">Today</button>
    </div>
    <div class="overview-table-wrapper">
      <table class="overview-table">
      <thead>
        <!-- Row 1: Month headers -->
        <tr class="header-row-month">
          <th class="project-col-header" rowspan="3">Project</th>
          <th class="role-col-header" rowspan="3">Role</th>
          <th class="person-col-header" rowspan="3">Personnel</th>`;
  
  // Month header row - group days by month
  let currentMonth = null;
  let monthStartIndex = 0;
  allDays.forEach((dayInfo, index) => {
    const monthKey = `${dayInfo.year}-${dayInfo.month}`;
    if (currentMonth !== monthKey) {
      if (currentMonth !== null) {
        // Close previous month header
        const monthSpan = index - monthStartIndex;
        html += `<th class="month-header" colspan="${monthSpan}">${allDays[monthStartIndex].monthName} ${allDays[monthStartIndex].year}</th>`;
      }
      // Start new month header
      monthStartIndex = index;
      currentMonth = monthKey;
    }
  });
  // Close last month header
  if (currentMonth !== null) {
    const monthSpan = allDays.length - monthStartIndex;
    html += `<th class="month-header" colspan="${monthSpan}">${allDays[monthStartIndex].monthName} ${allDays[monthStartIndex].year}</th>`;
  }
  html += `</tr>`;
  
  // Row 2: Week number headers
  html += `<tr class="header-row-week">`;
  let currentWeek = null;
  let weekStartIndex = 0;
  allDays.forEach((dayInfo, index) => {
    const weekKey = `${dayInfo.year}-W${dayInfo.weekNumber}`;
    if (currentWeek !== weekKey) {
      if (currentWeek !== null) {
        // Close previous week header
        const weekSpan = index - weekStartIndex;
        html += `<th class="week-header" colspan="${weekSpan}">KW${allDays[weekStartIndex].weekNumber}</th>`;
      }
      weekStartIndex = index;
      currentWeek = weekKey;
    }
  });
  // Close last week header
  if (currentWeek !== null) {
    const weekSpan = allDays.length - weekStartIndex;
    html += `<th class="week-header" colspan="${weekSpan}">KW${allDays[weekStartIndex].weekNumber}</th>`;
  }
  html += `</tr>`;
  
  // Row 3: Day name headers (weekday only)
  html += `<tr class="header-row-day">`;
  allDays.forEach(dayInfo => {
    const hasConflict = conflictedDates.has(dayInfo.dateKey);
    const conflictClass = hasConflict ? 'conflict' : '';
    html += `<th class="day-header ${dayInfo.isWeekend ? 'weekend' : ''} ${conflictClass}" title="${dayInfo.dateKey}">
      ${dayInfo.dayName}
    </th>`;
  });
  html += `</tr>
      </thead>
      <tbody>`;
  
  // Group data rows by project (one card per project)
  projectsToShow.forEach(project => {
    const projectName = typeof project === 'string' ? project : project.name;
    
    // Skip if project name is invalid
    if (!projectName || projectName.trim() === '' || projectName === '...') {
      return;
    }
    
    // Get all person+role combinations for this project
    const personRoleCombos = personRoleDatesByProject[project] || {};
    const personRoleKeys = Object.keys(personRoleCombos).filter(key => {
      const combo = personRoleCombos[key];
      return combo && combo.person && combo.role && combo.dates && combo.dates.length > 0;
    });
    
    // Sort by CONFIG.roles order first, then by CONFIG.personnel order
    const rolesOrder = CONFIG.roles.map(r => typeof r === 'string' ? r : r.name) || ['Project-Manager', 'Foreman', 'Shaper', 'Operator-Shaper'];
    const peopleOrder = CONFIG.personnel.map(p => typeof p === 'string' ? p : p.name);
    
    personRoleKeys.sort((a, b) => {
      const comboA = personRoleCombos[a];
      const comboB = personRoleCombos[b];
      
      // First sort by role order
      const roleIndexA = rolesOrder.indexOf(comboA.role);
      const roleIndexB = rolesOrder.indexOf(comboB.role);
      if (roleIndexA !== roleIndexB) {
        // If role not found in config, put it at the end
        if (roleIndexA === -1) return 1;
        if (roleIndexB === -1) return -1;
        return roleIndexA - roleIndexB;
      }
      
      // Then sort by person order
      const personIndexA = peopleOrder.indexOf(comboA.person);
      const personIndexB = peopleOrder.indexOf(comboB.person);
      if (personIndexA !== personIndexB) {
        // If person not found in config, put it at the end
        if (personIndexA === -1) return 1;
        if (personIndexB === -1) return -1;
        return personIndexA - personIndexB;
      }
      
      return 0;
    });
    
    // Skip projects with no valid assignments
    if (personRoleKeys.length === 0) {
      return;
    }
    
    // For each person+role combination, create a row
    personRoleKeys.forEach((personRoleKey, index) => {
      const { person, role, dates } = personRoleCombos[personRoleKey];
      
      if (!person || !role || !dates || dates.length === 0) {
        return;
      }
      
      const roleColor = getRoleColor(role);
      const dateSet = new Set(dates);
      const isFirstRow = index === 0;
      const rowSpan = personRoleKeys.length;
      
      // Check if this person has conflicts on any of their assigned dates
      const hasPersonConflict = dates.some(dateKey => {
        return conflictedDates.has(dateKey) && 
               personnelCountByDate[dateKey] && 
               personnelCountByDate[dateKey][person] > 1;
      });
      const personColClass = hasPersonConflict ? 'person-col conflict' : 'person-col';
      
      html += `<tr class="data-row">`;
      
      // Project column (only in first row, spans all rows for this project, rotated 90°)
      if (isFirstRow) {
        html += `<td class="project-col" rowspan="${rowSpan}">
          <div class="project-name-rotated">${projectName}</div>
        </td>`;
      }
      
      // Role column (with role color background)
      html += `<td class="role-col" style="background-color: ${roleColor};">${role}</td>`;
      
      // Person column (diagonal stripe pattern if this person has conflicts)
      html += `<td class="${personColClass}">${person}</td>`;
      
      // Day cells
      allDays.forEach((dayInfo, dayIndex) => {
        const hasPerson = dateSet.has(dayInfo.dateKey);
        // Use local date to avoid timezone issues
        const today = new Date();
        const todayKey = formatLocalDate(today);
        const isToday = dayInfo.dateKey === todayKey;
        
        // Check if this is start, middle, or end of a bar span
        const prevDay = dayIndex > 0 ? allDays[dayIndex - 1] : null;
        const nextDay = dayIndex < allDays.length - 1 ? allDays[dayIndex + 1] : null;
        const hasPrev = prevDay && dateSet.has(prevDay.dateKey);
        const hasNext = nextDay && dateSet.has(nextDay.dateKey);
        
        let barClass = '';
        if (hasPerson) {
          if (hasPrev && hasNext) {
            barClass = 'bar-middle';
          } else if (hasPrev) {
            barClass = 'bar-end';
          } else if (hasNext) {
            barClass = 'bar-start';
          } else {
            barClass = 'bar-single';
          }
        }
        
        html += `<td class="day-cell ${isToday ? 'today' : ''} ${hasPerson ? 'has-personnel' : ''} ${dayInfo.isWeekend ? 'weekend' : ''} ${barClass}" 
          data-date="${dayInfo.dateKey}"
          data-person="${person}"
          title="${dayInfo.date.toLocaleDateString()} - ${projectName} - ${person} (${role})">`;
        
        html += `<div class="day-number-in-cell">${dayInfo.day}</div>`;
        
        if (hasPerson) {
          html += `<div class="person-role-bar" style="background-color: ${roleColor};"></div>`;
        }
        
        html += `</td>`;
      });
      
      html += `</tr>`;
    });
    
    // End project card group - add separator row
    html += `<tr class="project-group-end">
      <td class="project-separator-left" colspan="3"></td>`;
    
    allDays.forEach(() => {
      html += `<td class="project-separator-cell"></td>`;
    });
    
    html += `</tr>`;
  });
  
  html += `</tbody>
      </table>
    </div>
  </div>`;
  
  container.innerHTML = html;
  
  // Add scroll to today button handler
  const scrollToTodayBtn = container.querySelector('#scrollToTodayBtn');
  if (scrollToTodayBtn) {
    scrollToTodayBtn.addEventListener('click', () => {
      const todayCell = container.querySelector('.day-cell.today');
      if (todayCell) {
        todayCell.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
      }
    });
  }
  
  // Auto-scroll to today on initial load
  setTimeout(() => {
    const todayCell = container.querySelector('.day-cell.today');
    if (todayCell) {
      todayCell.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
    }
  }, 100);
  
  // No click handlers - this is a read-only presentation view
}

// Initialize when page loads
document.addEventListener('DOMContentLoaded', init);

