// Configuration for Personnel Planning Calendar
const CONFIG = {
  // Calendar ID of the shared dummy calendar
  // Get this from Google Calendar settings > Integrate calendar > Calendar ID
  calendarId: 'c_8a82d2d6d54ec545f6019870cde156c3b4f0b338d760091f88c329358e0de867@group.calendar.google.com',
  
  // List of people with colors
  people: [
    { name: 'Cody', color: '#4285f4' },
    { name: 'Kay', color: '#ea4335' },
    { name: 'Rene', color: '#fbbc04' },
    { name: 'Gary', color: '#34a853' },
    { name: 'Dimitri', color: '#9c27b0' },
    { name: 'Alec', color: '#ff9800' },
    { name: 'Allen', color: '#00bcd4' },
    { name: 'Paul S', color: '#795548' },
    { name: 'Paul V', color: '#607d8b' },
    { name: 'Viktor', color: '#e91e63' }
    // Add more people as needed
  ],
  
  // List of projects
  projects: [
    'Rossau II',
    'Klaus Äule',
    'Alberschwende',
    '...'
    // Add more projects as needed
  ],
  
  // OAuth 2.0 Client ID from Google Cloud Console
  // Get this from: APIs & Services > Credentials > OAuth 2.0 Client IDs
  oauthClientId: '581029517390-ofv5l3p97p5ikikp59d99o10065gc7eb.apps.googleusercontent.com'
};

