# Personnel Planning Calendar

A drag-and-drop web application for visual personnel allocation planning using Google Calendar. This app uses a shared dummy calendar to track who is working on which projects during specific date spans.

## Features

- **Visual Overview**: Month, week, and year views of personnel allocations
- **Table-Based Overview**: Comprehensive table view with project/role/person columns and sticky headers
- **Role-Based Assignments**: Assign specific roles (Project-Manager, Foreman, Shaper, Operator-Shaper) to personnel
- **Drag & Drop**: Easily move assignments by dragging events
- **Create Assignments**: Click empty slots to create new person-project-role assignments
- **Filtering**: Filter by person or project
- **Color Coding**: Projects are color-coded for easy visual identification
- **Secure Authentication**: Uses OAuth 2.0 with Google Workspace accounts
- **Today Navigation**: Quick scroll to today's date in the overview

## Prerequisites

- Google Workspace account
- A shared dummy calendar in Google Calendar (separate from real person calendars)
- Basic knowledge of Google Cloud Console

## Setup Instructions

### Step 1: Create a Shared Dummy Calendar

1. Open Google Calendar
2. Click the "+" next to "Other calendars" on the left sidebar
3. Select "Create new calendar"
4. Name it "Team Personnel Planning" (or any name you prefer)
5. Click "Create calendar"
6. Note the Calendar ID (found in Settings > Integrate calendar > Calendar ID)
   - Format: `xxxxx@group.calendar.google.com`

### Step 2: Google Cloud Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project (or select existing)
3. Enable the Calendar API:
   - Navigate to "APIs & Services" > "Library"
   - Search for "Google Calendar API"
   - Click "Enable"

### Step 3: Create OAuth 2.0 Credentials

1. In Google Cloud Console, go to "APIs & Services" > "Credentials"
2. Click "Create Credentials" > "OAuth client ID"
3. If prompted, configure the OAuth consent screen:
   - Choose "Internal" (for Google Workspace) or "External"
   - Fill in the required fields (App name, User support email, Developer contact)
   - Click "Save and Continue"
   - Add scopes: `https://www.googleapis.com/auth/calendar`
   - Click "Save and Continue"
   - Add test users if using "External" (or skip if "Internal")
   - Click "Save and Continue" then "Back to Dashboard"
4. Select "Web application" as the application type
5. Enter a name (e.g., "Personnel Calendar App")
6. Under "Authorized JavaScript origins", add:
   - `http://localhost:8000` (for local development)
   - Your production URL if deploying (e.g., `https://yourdomain.com`)
7. Under "Authorized redirect URIs", add:
   - `http://localhost:8000` (for local development)
   - Your production URL if deploying
8. Click "Create"
9. **Copy the Client ID** (you'll need this for `config.js`)

### Step 4: Configure the App

1. Open `config.js`
2. Update `calendarId` with your calendar ID from Step 1
3. Update `oauthClientId` with your OAuth Client ID from Step 3
4. Update `people` array with your team member names and colors
5. Update `projects` array with your project names
6. Update `roles` array with your role definitions:

```javascript
const CONFIG = {
  calendarId: 'your-calendar-id@group.calendar.google.com',
  oauthClientId: 'YOUR_CLIENT_ID.apps.googleusercontent.com',
  people: [
    { name: 'John', color: '#4285f4' },
    { name: 'Sarah', color: '#ea4335' },
    { name: 'Mike', color: '#fbbc04' },
    // ... add your people with colors
  ],
  projects: [
    'Project Alpha',
    'Project Beta',
    // ... add your projects
  ],
  roles: [
    'Project-Manager',
    'Foreman',
    'Shaper',
    'Operator-Shaper'
    // ... add your roles
  ]
};
```

### Step 5: Deploy or Run Locally

#### Option A: Local Development Server

Since the app uses OAuth 2.0, you need a local server (file:// won't work):

**Quick Start (Recommended):**

**Windows:**
```bash
run-local.bat
```

**Mac/Linux or if you have Node.js:**
```bash
node run-local.js
```

The script will automatically:
- Start a local HTTP server on port 8000
- Open your browser to `http://localhost:8000`

**Manual Start:**

**Using Python:**
```bash
python -m http.server 8000
```

**Using Node.js (http-server):**
```bash
npm install -g http-server
http-server -p 8000
```

Then open: `http://localhost:8000`

**First time**: Click "Sign in with Google" and authorize the app to access your calendar.

#### Option B: Deploy to GitHub Pages

1. Create a GitHub repository
2. Push your code
3. Go to repository Settings > Pages
4. Select source branch and folder
5. Your app will be available at `https://yourusername.github.io/repository-name`
6. **Important**: Update the OAuth Client ID in Google Cloud Console to include your GitHub Pages URL in "Authorized JavaScript origins"

#### Option C: Deploy to Netlify/Vercel

1. Sign up for [Netlify](https://netlify.com) or [Vercel](https://vercel.com)
2. Connect your GitHub repository
3. Deploy
4. **Important**: Update the OAuth Client ID in Google Cloud Console to include your deployment URL in "Authorized JavaScript origins"

## Usage

1. Open the app in your browser
2. Click "Sign in with Google" (first time only, or if signed out)
3. Authorize the app to access your Google Calendar
4. The calendar will load events from your shared dummy calendar
5. **Create assignment**: Click an empty date slot, select person, project, and role
6. **Move assignment**: Drag an event to a different date
7. **Resize assignment**: Drag the edges of an event to change duration
8. **Delete assignment**: Click an event and confirm deletion
9. **Filter**: Use the dropdowns to filter by person or project
10. **Change view**: Use the view selector or calendar toolbar buttons
11. **Overview mode**: Click the "Overview" button to see a table-based view with:
    - Projects grouped vertically (each project has its own section)
    - Each project shows rows for each person-role combination with assignments
    - Days as columns (grouped by month and week)
    - Three sticky left columns: Project name, Role, and Person
    - Sticky top header with year display and "Today" button
    - Color-coded bars showing assignments spanning multiple days
    - Day numbers displayed in cells, weekday names in column headers
    - Empty rows between projects have no grid lines for visual separation
12. **Sign out**: Click "Sign out" button when done

## Overview Structure

The Overview mode provides a comprehensive table view of all personnel assignments:

- **Left Columns (Sticky)**: 
  - **Project**: Project name displayed in the first row of each project group
  - **Role**: Role assigned (e.g., Project-Manager, Foreman)
  - **Person**: Person assigned to the role

- **Top Headers (Sticky)**:
  - **Year and Today Button**: Always visible at the top
  - **Month Headers**: Group columns by month
  - **Week Headers**: Show work week numbers
  - **Day Headers**: Display weekday names (Mon, Tue, Wed, etc.)

- **Data Cells**:
  - Day numbers appear in each cell
  - Color-coded bars span across days for multi-day assignments
  - Bars show the project color for easy identification

## Event Format

Events are stored in Google Calendar with the format:
```
[Person Name] - [Project Name] - [Role]
```

Examples:
- `John - Project Alpha - Project-Manager`
- `Sarah - Project Beta - Foreman`
- `Mike - Project Alpha - Shaper`

## Security Notes

- **OAuth 2.0 is secure**: No private keys are stored in the app
- Users authenticate with their own Google Workspace accounts
- Access tokens are short-lived and scoped to calendar access only
- Each user can only access calendars they have permission to view/edit

## Troubleshooting

**"Sign in failed" or OAuth errors**
- Verify the OAuth Client ID in `config.js` matches your Google Cloud Console credentials
- Check that your domain is added to "Authorized JavaScript origins" in Google Cloud Console
- Ensure the Calendar API is enabled in Google Cloud Console
- Check browser console for detailed error messages

**"Calendar ID not found" or "Events not showing"**
- Verify the calendar ID in `config.js` matches your calendar
- Ensure you're signed in with a Google account that has access to the calendar
- Check that the calendar is shared with your Google account
- Verify you have "Make changes to events" permission on the calendar

**"Not signed in" errors**
- Click "Sign in with Google" button
- Make sure popup blockers aren't blocking the OAuth window
- Check browser console for authentication errors

## Customization

### Adding People
Edit `config.js` and add people to the `people` array with names and colors:
```javascript
people: [
  { name: 'John', color: '#4285f4' },
  { name: 'Sarah', color: '#ea4335' },
  // ... add more people
]
```

### Adding Projects
Edit `config.js` and add project names to the `projects` array:
```javascript
projects: [
  'Project Alpha',
  'Project Beta',
  // ... add more projects
]
```

### Adding Roles
Edit `config.js` and add roles to the `roles` array:
```javascript
roles: [
  'Project-Manager',
  'Foreman',
  'Shaper',
  'Operator-Shaper',
  // ... add more roles
]
```

### Changing Person Colors
Update the `color` property in the `people` array. Use hex colors (e.g., `#4285f4`).

### Managing People and Projects
Use the "Manage People" and "Manage Projects" buttons in the interface to add, edit, or remove entries without editing the config file directly.

## License

This is a simple utility tool. Use as needed for your organization.

