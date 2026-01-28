// Google Sheets API Configuration
let CLIENT_ID = '';
let CLIENT_SECRET = '';
let API_KEY = '';
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly';

// Global state
let tokenClient;
let gapiInited = false;
let gisInited = false;

// Application state
const state = {
  url: "",
  mode: "flash",
  outputFormat: "HTML",
  options: {
    responsive: false,
    styling: false,
    headers: false,
    search: false,
    sort: false,
    export: false,
  },
  isHovering: false,
  lastGeneratedHtml: null,
  lastRawData: null,
  backgroundData: null,
  customSettings: {
    theme: 'light',
    accentPrimary: '#667eea',
    padding: 'normal',
    fontSize: 14,
    fontFamily: 'Inter',
    roundness: 'md',
    shadow: 'md',
    showKPIs: true,
    showSearch: true,
    showExport: true,
    showPrint: true,
    striped: false,
    hover: true,
    borders: false,
    fadeIn: true,
    cardHover: true
  }
};
