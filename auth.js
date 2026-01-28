// Google Authentication Logic

/**
 * Callback after Google Identity Services are loaded.
 */
function gapiLoaded() {
  gapi.load('client', initializeGapiClient);
}

/**
 * Callback after the API client is loaded. Loads the
 * discovery doc to initialize the API.
 */
async function initializeGapiClient() {
  // We only need the discovery doc for the Sheets API
  await gapi.client.init({
    apiKey: API_KEY, // Use the API Key for unauthenticated access (Flash Mode)
    discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4'],
  });
  gapiInited = true;
  maybeEnableButtons();
}

/**
 * Callback after Google Identity Services are loaded.
 */
function gisLoaded() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES,
    callback: '', // defined later
  });
  gisInited = true;
  maybeEnableButtons();
}

/**
 * Enables user interaction after all libraries are loaded.
 */
function maybeEnableButtons() {
  if (gapiInited && gisInited) {
    authBtn.style.display = 'inline-flex';
    authBtn.onclick = handleAuthClick;
  }
}

/**
 *  Sign in the user upon button click.
 */
function handleAuthClick() {
  tokenClient.callback = async (resp) => {
    if (resp.error !== undefined) {
      throw (resp);
    }
    authBtn.innerText = 'Connected (Recommended)';
    authBtn.style.backgroundColor = 'var(--accent-mint)'; // Active state
    authBtn.style.color = 'var(--text-primary)';
    authBtn.style.borderColor = 'transparent';
    authBtn.onclick = handleSignoutClick;
    
    // Set Mode
    state.mode = 'recommended';
    flashBtn.classList.remove('active');

    // Unlock Input
    sheetInput.disabled = false;
    pasteBtn.disabled = false;
    sheetInput.placeholder = "https://docs.google.com/spreadsheets/d/...";
    
    // Check if we have a URL waiting
    if (state.url) {
        // Future: trigger conversion automatically?
        console.log("Ready to access: " + state.url);
    }
  };

  if (gapi.client.getToken() === null) {
    // Prompt the user to select a Google Account and ask for consent to share their data
    // when establishing a new session.
    tokenClient.requestAccessToken({prompt: 'consent'});
  } else {
    // Skip display of account chooser and consent dialog for an existing session.
    tokenClient.requestAccessToken({prompt: ''});
  }
}

/**
 *  Sign out the user upon button click.
 */
function handleSignoutClick() {
  const token = gapi.client.getToken();
  if (token !== null) {
    google.accounts.oauth2.revoke(token.access_token);
    gapi.client.setToken('');
    authBtn.innerText = 'Connect (Recommended)';
    authBtn.style.backgroundColor = ''; 
    authBtn.style.color = '';
    authBtn.style.borderColor = '';
    authBtn.onclick = handleAuthClick;
    
    // Fallback to Flash mode or Lock Input?
    // If we sign out, we can't use Recommended.
    // Let's switch to Flash? Or just lock?
    // User wants Flash default.
    state.mode = 'flash';
    updateModeUI();

    // Note: inputs might be locked by updateModeUI if mode was recommended, 
    // but we just set it to flash, so they will be unlocked.
    // However, if we wanted to stay in "Connect" mode but locked...
    // Let's assume sign out -> return to default Flash state.
  }
}
