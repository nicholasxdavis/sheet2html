/**
 * Fetches API credentials from a local file (devkeys.json) if available,
 * otherwise falls back to build-time/environment defaults ("GitHub Secrets").
 */
async function loadCredentials() {
    // Default / "GitHub Secrets" placeholders
    // These strings are expected to be replaced during the build process
    // or handled via environment variable injection if applicable.
    let creds = {
        clientId: '38772390105-8v86sf9k06or2i72uha5qrhhckq6ltq8.apps.googleusercontent.com',
        clientSecret: '', // Not exposed on client side
        apiKey: 'AIzaSyB46Fth--G1-nFy0KXyJbPZN1D71wI4TKw'
    };

    try {
        console.log('Attempting to fetch devkeys.json...');
        const response = await fetch('devkeys.json');
        
        if (response.ok) {
            const json = await response.json();
            console.log('Successfully loaded credentials from devkeys.json');
            
            // Override defaults if keys exist in JSON
            // Note: devkeys.json uses GitHub Secret names as keys
            if (json.GOOGLE0CLIENT0ID) creds.clientId = json.GOOGLE0CLIENT0ID;
            if (json.GOOGLE0SECRET) creds.clientSecret = json.GOOGLE0SECRET;
            if (json.GOOGLE0API0KEY) creds.apiKey = json.GOOGLE0API0KEY;
        } else {
            console.log('devkeys.json not found/ok, using default credentials');
        }
    } catch (e) {
        // Fallback silently or log warning
        console.warn('Error fetching devkeys.json, using default credentials:', e);
    }

    return creds;
}

/**
 * Fetches sheet data using the Recommended method (OAuth + Full API).
 * @param {string} spreadsheetId 
 * @returns {Promise<Object>} The raw result from the API.
 */
async function fetchSheetDataRecommended(spreadsheetId) {
    const response = await gapi.client.sheets.spreadsheets.get({
        spreadsheetId: spreadsheetId,
        includeGridData: true,
        ranges: [], 
    });

    const result = response.result;
    
    if (!result.sheets || result.sheets.length === 0) {
        throw new Error("No sheets found in this spreadsheet.");
    }
    
    return result; 
}

/**
 * Fetches sheet data using the Flash method (API Key -> Fallback to GViz).
 * @param {string} spreadsheetId 
 * @param {string} [gid] Optional GID for GViz fallback
 * @returns {Promise<{mode: string, data: any}>} The result object containing mode and data.
 */
async function fetchSheetDataFlash(spreadsheetId, gid) {
    // Attempt 1: High Fidelity via API Key
    let apiKeyError = null;
    try {
        console.log("Attempting Flash conversion via API Key...");
        
        const response = await gapi.client.sheets.spreadsheets.get({
            spreadsheetId: spreadsheetId,
            includeGridData: true,
            ranges: [], 
        });

        const result = response.result;
        
        if (!result.sheets || result.sheets.length === 0) {
            throw new Error("No sheets found in this spreadsheet.");
        }

        return {
            mode: 'recommended', // It returns the same structure as recommended
            data: result.sheets
        };

    } catch (err) {
        console.warn("Flash API Key method failed or not available. Using fallback (GViz). Error:", err);
        apiKeyError = err;
        // Fallback to Legacy GViz
        try {
            return await fetchSheetDataGViz(spreadsheetId, gid);
        } catch (gvizErr) {
            // Combined Error Analysis
            if (apiKeyError && apiKeyError.status === 403) {
                // If API key failed with 403 AND GViz failed, it's definitely a permissions issue
                throw new Error("Access Denied. Flash Mode requires the Sheet to be 'Anyone with the link can view'.");
            }
            throw gvizErr;
        }
    }
}

/**
 * Fetches sheet data using the Legacy GViz method.
 * @param {string} spreadsheetId 
 * @param {string} [gid] 
 * @returns {Promise<{mode: string, data: any}>}
 */
function fetchSheetDataGViz(spreadsheetId, gid) {
    return new Promise((resolve, reject) => {
        console.log("Running Legacy Flash Fetch (GViz)...");
        
        try {
            google.charts.load('current', {packages: ['table']});
            google.charts.setOnLoadCallback(() => {
                const queryUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?headers=1${gid ? '&gid=' + gid : ''}`;
                const query = new google.visualization.Query(queryUrl);
                
                // Set a timeout since GViz doesn't handle CORS hangs well
                const timer = setTimeout(() => {
                    reject(new Error("Request timed out. Please check if the sheet is Public."));
                }, 10000);

                query.send((response) => {
                    clearTimeout(timer);
                    if (response.isError()) {
                        const msg = response.getMessage();
                        // Status 0 or "Request Failed" usually means CORS/Network
                        if (msg.includes('Request Failed') || msg.includes('status=0')) {
                             reject(new Error("Connection Failed. The sheet might be Private or the link is invalid."));
                        } else {
                             reject(new Error(`Flash Conversion Failed: ${msg}`));
                        }
                    } else {
                        resolve({
                            mode: 'flash',
                            data: response.getDataTable()
                        });
                    }
                });
            });
        } catch (e) {
            reject(new Error("Failed to initialize Google Charts: " + e.message));
        }
    });
}
