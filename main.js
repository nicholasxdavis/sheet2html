// Main Application Initialization and Startup

// Load the libraries
// Note: We need to trigger these manually or ensure the script tags in HTML call them. 
// Since we added script tags with async defer, we can just wait for them or use window.onload.
// Or effectively, just set global callbacks if the script tags have `onload` attributes.
// Current script tags in index.html don't have onload. Let's force load or wait.
// easiest way:
window.onload = async function() {
    // Load credentials first
    const creds = await loadCredentials();
    CLIENT_ID = creds.clientId;
    CLIENT_SECRET = creds.clientSecret;
    API_KEY = creds.apiKey;
    
    // Check if google libraries are present
    if (typeof gapi !== 'undefined') gapiLoaded();
    if (typeof google !== 'undefined') gisLoaded();
    
    // Initialize Logic
    if (state.mode === 'flash') {
        sheetInput.disabled = false;
        pasteBtn.disabled = false;
        sheetInput.placeholder = "https://docs.google.com/spreadsheets/d/...";
    } else {
        // Should not happen on load given default, but just in case
        sheetInput.disabled = true;
        pasteBtn.disabled = true;
        sheetInput.placeholder = "Please connect to Google Sheets first";
    }
}
