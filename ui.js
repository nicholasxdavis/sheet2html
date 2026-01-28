// UI Event Handlers and Interactions

let bgFetchTimeout = null;


// DOM Elements
const authBtn = document.getElementById("authBtn");
const flashBtn = document.getElementById("flashBtn");
const sheetInput = document.getElementById("sheetUrl");
const convertBtn = document.getElementById("convertBtn");
const previewBtn = document.getElementById("previewBtn");
const outputPreview = document.getElementById("outputPreview");
const optionCards = document.querySelectorAll(".option-card");
const skeletonLoader = document.getElementById("skeleton-loader");
const mainCard = document.querySelector(".main-card");
const comparisonSection = document.querySelector(".comparison-section");
const pasteBtn = document.getElementById("pasteBtn");

// Modal Elements
const downloadHtmlBtn = document.getElementById('downloadHtmlBtn');
const exportJsonBtn = document.getElementById('exportJsonBtn');
const resultModal = document.getElementById('resultModal');
const modalBody = document.getElementById('modalBody');
const closeModalBtn = document.getElementById('closeModalBtn');
const downloadResultBtn = document.getElementById('downloadResultBtn');
const copyResultBtn = document.getElementById('copyResultBtn');
const outputFormatSwitch = document.getElementById('outputFormatSwitch');
const heroTitle = document.getElementById('heroTitle');
const conversionOptionsSection = document.getElementById('conversionOptionsSection');
const customizationSection = document.getElementById('customizationSection');

// Output Format Toggle Logic
if (outputFormatSwitch) {
    outputFormatSwitch.addEventListener("click", (e) => {
        const option = e.target.closest(".switch-option");
        if (!option) return;

        const value = option.dataset.value;
        state.outputFormat = value;
        outputFormatSwitch.dataset.value = value;

        // Update Hero Title
        heroTitle.innerHTML = `Transform Spreadsheets into Beautiful ${value}`;

        // Toggle Visibility of Customization Sections
        if (value === 'HTML') {
            conversionOptionsSection.style.display = 'block';
            customizationSection.style.display = state.options.styling ? 'block' : 'none';
            convertBtn.querySelector('span').textContent = 'Convert to HTML';
            previewBtn.querySelector('span').textContent = 'Preview';
        } else {
            conversionOptionsSection.style.display = 'none';
            customizationSection.style.display = 'none';
            convertBtn.querySelector('span').textContent = `Convert to ${value}`;
            previewBtn.querySelector('span').textContent = `Preview ${value}`;
        }

        console.log("Output Format Selected:", value);
    });
}

// Comparison Slider Elements
const slider = document.getElementById('comparisonSlider');
const sliderOverlay = document.getElementById('imageOverlay');
const sliderHandle = document.getElementById('sliderHandle');

// Input validation and button state
sheetInput.addEventListener("input", (e) => {
  const value = e.target.value.trim();
  state.url = value;

  // Simple validation for Google Sheets URL
  const isValid = value.includes("docs.google.com/spreadsheets");
  convertBtn.disabled = !isValid;
  previewBtn.disabled = !isValid;

  // Visual feedback
  if (value && !isValid) {
    sheetInput.style.borderColor = "#ff6b6b";
  } else if (isValid) {
    sheetInput.style.borderColor = "var(--accent-mint)";
    
    // Trigger Background Fetch (Debounced)
    if (bgFetchTimeout) clearTimeout(bgFetchTimeout);
    state.backgroundData = null; // Clear old cache on new input
    
    bgFetchTimeout = setTimeout(() => {
        triggerBackgroundFetch(value);
    }, 500);

  } else {
    sheetInput.style.borderColor = "var(--border-color)";
    state.backgroundData = null;
  }
});

// Background Fetch Logic
async function triggerBackgroundFetch(url) {
    const { id: spreadsheetId } = extractSheetInfo(url);
    if (!spreadsheetId) return;

    if (state.mode === 'recommended' && !gapi.client.getToken()) {
        console.log("Skipping background fetch: Not signed in.");
        return;
    }

    console.log(`Triggering background fetch for ${spreadsheetId} (${state.mode})...`);
    
    try {
        let data;
        if (state.mode === 'recommended') {
            data = await fetchSheetDataRecommended(spreadsheetId);
        } else {
            data = await fetchSheetDataFlash(spreadsheetId);
        }
        
        // Check if user hasn't changed input while fetching
        const currentId = extractSheetInfo(sheetInput.value).id;
        if (currentId === spreadsheetId) {
            state.backgroundData = {
                id: spreadsheetId,
                mode: state.mode,
                content: data
            };
            console.log("Background fetch cached successfully.");
        }
    } catch (err) {
        // Suppress known "Private Sheet" errors to avoid console noise during typing
        if (err.message && (err.message.includes('Access Denied') || err.message.includes('Connection Failed'))) {
             console.log("Background Check: Sheet appears private or inaccessible (Flash mode). Waiting for explicit user action.");
        } else {
             console.warn("Background fetch warning:", err);
        }
    }
}

// Option card toggles
optionCards.forEach((card) => {
  card.addEventListener("click", () => {
    const option = card.dataset.option;
    
    // Don't allow clicking disabled cards (except styling)
    if (card.classList.contains('disabled') && option !== 'styling') {
      return;
    }
    
    card.classList.toggle("active");
    state.options[option] = !state.options[option];

    // Show/hide Customization Options section when Custom Styling is toggled
    if (option === 'styling') {
      const customizationSection = document.getElementById('customizationSection');
      if (customizationSection) {
        customizationSection.style.display = state.options.styling ? 'block' : 'none';
      }
      // Update dependent options visibility/state
      updateDependentOptions();
    }

    // Sync Standard Options with Custom Settings if styling is active
    if (state.options.styling) {
        if(option === 'search') state.customSettings.showSearch = state.options.search;
        if(option === 'export') state.customSettings.showExport = state.options.export;
        if(option === 'sort') state.customSettings.sortable = state.options.sort;
        if(option === 'headers') state.customSettings.useHeaderRow = state.options.headers;
        if(option === 'responsive') state.customSettings.responsive = state.options.responsive;
    }

    // Micro-interaction: slight bounce
    card.style.transform = "scale(0.98)";
    setTimeout(() => {
      card.style.transform = "";
    }, 100);
  });
});

// Function to update dependent options based on custom styling
function updateDependentOptions() {
  const dependentOptions = ['responsive', 'headers', 'search', 'sort', 'export'];
  const isStylingEnabled = state.options.styling;
  
  dependentOptions.forEach(option => {
    const card = document.querySelector(`[data-option="${option}"]`);
    if (card) {
      if (isStylingEnabled) {
        card.classList.remove('disabled');
      } else {
        card.classList.add('disabled');
        // Turn off the option if it was active
        if (state.options[option]) {
          state.options[option] = false;
          card.classList.remove('active');
        }
      }
    }
  });
}

// Initialize dependent options state on page load
document.addEventListener('DOMContentLoaded', () => {
  updateDependentOptions();
  setupCustomizationListeners();
});

// Customization Listeners
function setupCustomizationListeners() {
    // Theme Buttons
    document.querySelectorAll('.theme-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
             document.querySelectorAll('.theme-btn').forEach(b => b.classList.remove('active'));
             btn.classList.add('active');
             state.customSettings.theme = btn.dataset.theme;
        });
    });

    // Color Dots
    document.querySelectorAll('.color-dot').forEach(btn => {
        btn.addEventListener('click', () => {
             document.querySelectorAll('.color-dot').forEach(b => b.classList.remove('active'));
             btn.classList.add('active');
             state.customSettings.accentPrimary = btn.dataset.color;
             
             // Reset custom picker button if present
             const pickerBtn = document.getElementById('picker-btn');
             if (pickerBtn) {
                 pickerBtn.style.background = '';
                 pickerBtn.style.color = '';
                 pickerBtn.classList.remove('active');
             }
        });
    });

    // Coloris Initialization
    if (window.Coloris) {
        Coloris({
            theme: 'polaroid',
            themeMode: 'light',
            alpha: false,
            format: 'hex',
            clearButton: true
        });
    }

    // Trigger Coloris on Button Click
    const pickerBtn = document.getElementById('picker-btn');
    const customAccentPicker = document.getElementById('customAccentPicker');
    if (pickerBtn && customAccentPicker) {
        pickerBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            customAccentPicker.click(); // Trigger Coloris popup
        });
    }

    // Coloris Change Event (using 'input' or 'change')
    if (customAccentPicker) {
         customAccentPicker.addEventListener('input', (e) => {
             const color = e.target.value;
             state.customSettings.accentPrimary = color;
             
             // Deselect dots
             document.querySelectorAll('.color-dot').forEach(b => b.classList.remove('active'));
             
             // Update picker button color
             if (pickerBtn) {
                 pickerBtn.style.background = color;
                 // Dynamic icon color (white or black based on luminance)
                 const r = parseInt(color.slice(1, 3), 16);
                 const g = parseInt(color.slice(3, 5), 16);
                 const b = parseInt(color.slice(5, 7), 16);
                 const brightness = (r * 299 + g * 587 + b * 114) / 1000;
                 pickerBtn.style.color = brightness > 155 ? '#333' : '#fff';
                 pickerBtn.classList.add('active'); // Style as chosen
             }
         });
    }

    // Sliders & Selects
    const bindInput = (id, key, transform = v => v) => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('input', (e) => {
                 state.customSettings[key] = transform(e.target.value);
                 if (id === 'setting-font-size') document.getElementById('font-size-val').textContent = e.target.value + 'px';
            });
        }
    };

    bindInput('setting-font-size', 'fontSize', Number);
    bindInput('setting-font-family', 'fontFamily');
    bindInput('setting-roundness', 'roundness');
    bindInput('setting-shadow', 'shadow');
    bindInput('setting-padding', 'padding');
    bindInput('setting-rows-per-page', 'rowsPerPage');
    bindInput('dashboard-name-input', 'dashboardName');
    
    // Toggles
    const bindCheckbox = (id, key) => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('change', (e) => state.customSettings[key] = e.target.checked);
    };

    bindCheckbox('setting-kpis', 'showKPIs');
    bindCheckbox('setting-search', 'showSearch');
    bindCheckbox('setting-export', 'showExport');
    bindCheckbox('setting-print', 'showPrint');
    bindCheckbox('setting-add-entry', 'showAddEntry');
    bindCheckbox('setting-striped', 'striped');
    bindCheckbox('setting-hover', 'hover');
    bindCheckbox('setting-borders', 'borders');
    bindCheckbox('setting-index', 'showIndex');
    bindCheckbox('setting-truncate', 'truncate');
    bindCheckbox('setting-show-empty', 'showEmpty');
}

// Global Reset Function
window.resetCustomSettings = function() {
    state.customSettings = {
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
         showAddEntry: true,
         striped: false,
         hover: true,
         borders: false,
         showIndex: true,
         truncate: true,
         showEmpty: true,
         dashboardName: 'Data Dashboard',
         rowsPerPage: 'all'
    };
    
    // Reset UI
    document.querySelectorAll('.theme-btn').forEach(b => b.classList.toggle('active', b.dataset.theme === 'light'));
    document.querySelectorAll('.color-dot').forEach(b => b.classList.toggle('active', b.dataset.color === '#667eea'));
    
    document.getElementById('setting-font-size').value = 14;
    document.getElementById('font-size-val').textContent = '14px';
    document.getElementById('setting-font-family').value = 'Inter';
    document.getElementById('setting-roundness').value = 'md';
    document.getElementById('setting-shadow').value = 'md';
    document.getElementById('setting-padding').value = 'normal';
    document.getElementById('setting-rows-per-page').value = 'all';
    document.getElementById('dashboard-name-input').value = 'Data Dashboard';
    
    document.getElementById('setting-kpis').checked = true;
    document.getElementById('setting-search').checked = true;
    document.getElementById('setting-export').checked = true;
    document.getElementById('setting-print').checked = true;
    document.getElementById('setting-add-entry').checked = true;
    document.getElementById('setting-striped').checked = false;
    document.getElementById('setting-hover').checked = true;
    document.getElementById('setting-borders').checked = false;
    document.getElementById('setting-index').checked = true;
    document.getElementById('setting-truncate').checked = true;
    document.getElementById('setting-show-empty').checked = true;
};

// Convert button handler
convertBtn.addEventListener("click", handleConvert);

async function handleConvert() {
  const url = sheetInput.value.trim();
  const { id: spreadsheetId } = extractSheetInfo(url);

  if (!spreadsheetId) {
    alert("Invalid Google Sheets URL. Could not find Spreadsheet ID.");
    return;
  }

  // Visual local feedback
  convertBtn.innerHTML = '<span class="loading"></span>';
  convertBtn.disabled = true;

  try {
      // 1. Ensure Data is Ready
      let inputData = null;
      if (state.backgroundData && state.backgroundData.id === spreadsheetId && state.backgroundData.mode === state.mode) {
          inputData = state.backgroundData.content;
      }

      if (!inputData) {
          if (state.mode === 'recommended') {
              if (!gapi.client.getToken()) { authBtn.click(); throw new Error("Please sign in first."); }
              inputData = await fetchSheetDataRecommended(spreadsheetId);
          } else {
              inputData = await fetchSheetDataFlash(spreadsheetId);
          }
          state.backgroundData = { id: spreadsheetId, mode: state.mode, content: inputData };
      }

      // 2. Normalize Data
      let normalizedData = {};
      const actualData = inputData.content || inputData.data || inputData;
      
      let sheetsArray = [];
      if (Array.isArray(actualData)) {
          sheetsArray = actualData;
      } else if (actualData.sheets && Array.isArray(actualData.sheets)) {
          sheetsArray = actualData.sheets;
      }
      
      if (sheetsArray.length > 0) {
          sheetsArray.forEach(sheet => {
              if (sheet.properties && sheet.data) {
                  const title = sheet.properties.title || 'Sheet1';
                  const rows = sheet.data?.[0]?.rowData || [];
                  normalizedData[title] = normalizeRowsToJson(rows, 'recommended', state.customSettings);
              }
          });
      } else if (inputData.mode === 'flash' || (actualData.getNumberOfColumns && typeof actualData.getNumberOfColumns === 'function')) {
           const dt = actualData.data || actualData;
           if (dt.getNumberOfColumns) normalizedData['Sheet1'] = normalizeRowsToJson(dt, 'flash', state.customSettings);
      } else if (Array.isArray(actualData) && actualData.length > 0) {
           normalizedData = { 'Sheet1': actualData };
      } else if (typeof actualData === 'object' && actualData !== null) {
           normalizedData = actualData;
      }

      if (Object.keys(normalizedData).length === 0) throw new Error("No data found.");

      // 3. Generate Output based on Format
      let finalOutput = "";
      if (state.outputFormat === 'HTML') {
          finalOutput = CustomGenerator.generate(normalizedData, state.customSettings);
      } else if (state.outputFormat === 'JSON') {
          finalOutput = JSON.stringify(normalizedData, null, 2);
      } else if (state.outputFormat === 'CSV') {
          finalOutput = jsonToCsv(normalizedData);
      }

      // 4. Update UI
      updatePreviewState(finalOutput);
      openModal(finalOutput);

      // Reset Button
      convertBtn.innerHTML = `
            <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <polyline points="16 18 22 12 16 6"/>
                <polyline points="8 6 2 12 8 18"/>
            </svg>
            <span>Convert to ${state.outputFormat}</span>
        `;
      convertBtn.disabled = false;

  } catch (e) {
      console.error(e);
      alert("Conversion failed: " + e.message);
      convertBtn.innerHTML = `<span>Convert to ${state.outputFormat}</span>`;
      convertBtn.disabled = false;
  }
}

// Mode UI Helper
function updatePlaceholder() {
    if (state.mode === 'flash') {
        if (state.isHovering) {
             sheetInput.placeholder = "https://docs.google.com/spreadsheets/d/...";
        } else {
             sheetInput.placeholder = "Spreadsheet must be Public (Anyone with the link)";
        }
    }
}

function updateModeUI() {
    if (state.mode === 'flash') {
        flashBtn.classList.add('active');
        // Do not touch authBtn style here as it is managed by auth status
        
        // Enable Inputs
        sheetInput.disabled = false;
        pasteBtn.disabled = false;
        updatePlaceholder();
    } else {
        flashBtn.classList.remove('active');
        
        // Input depends on Auth
        if (!gapi.client.getToken()) {
            sheetInput.disabled = true;
            pasteBtn.disabled = true;
            sheetInput.placeholder = "Please connect to Google Sheets first";
        } else {
            sheetInput.disabled = false;
            pasteBtn.disabled = false;
            sheetInput.placeholder = "https://docs.google.com/spreadsheets/d/...";
        }
    }
}

// Flash Button Listener
if (flashBtn) {
    flashBtn.addEventListener('click', () => {
        state.mode = 'flash';
        updateModeUI();
    });
}

// Hover Listeners for Placeholder
if (mainCard) {
    mainCard.addEventListener('mouseenter', () => {
        state.isHovering = true;
        updatePlaceholder();
    });
    mainCard.addEventListener('mouseleave', () => {
        state.isHovering = false;
        updatePlaceholder();
    });
}

// Preview button handler (Re-opens Modal)
previewBtn.addEventListener("click", () => {
  if (state.lastGeneratedHtml) {
      openModal(state.lastGeneratedHtml);
  } else {
      // Logic for when no conversion has happened yet (Placeholder behavior)
      previewBtn.innerHTML = '<span class="loading"></span>';
      previewBtn.disabled = true;

      setTimeout(() => {
        previewBtn.innerHTML = `
                        <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
                            <circle cx="12" cy="12" r="3"/>
                        </svg>
                        <span>Preview</span>
                    `;
        previewBtn.disabled = false;

        outputPreview.classList.add("active");
        outputPreview.innerHTML = `
                        <div style="text-align: center; color: var(--text-secondary);">
                            <div style="margin-bottom: 1rem;">
                                <svg style="width: 64px; height: 64px; stroke: var(--accent-mint); fill: none;" viewBox="0 0 24 24" stroke-width="2">
                                    <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
                                    <circle cx="12" cy="12" r="3"/>
                                </svg>
                            </div>
                            <p style="font-weight: 600; color: var(--text-primary); margin-bottom: 0.5rem;">Preview Ready</p>
                            <p style="font-size: 0.9rem;">Table preview will appear here</p>
                        </div>
                    `;
      }, 500);
  }
});

// Paste button handler
pasteBtn.addEventListener("click", async () => {
  try {
    const text = await navigator.clipboard.readText();
    sheetInput.value = text;
    // Trigger input event to run validation
    sheetInput.dispatchEvent(new Event("input"));
    
    // Visual feedback
    const originalText = pasteBtn.textContent;
    pasteBtn.textContent = "Pasted!";
    setTimeout(() => {
        pasteBtn.textContent = originalText;
    }, 1000);
  } catch (err) {
    console.error("Failed to read clipboard contents: ", err);
    pasteBtn.textContent = "Error";
    setTimeout(() => {
        pasteBtn.textContent = "Paste";
    }, 1000);
  }
});

// Modal Logic
function openModal(content) {
    if (!resultModal) return;
    
    // Check if it's full HTML
    if (content.trim().startsWith('<!DOCTYPE') || content.trim().startsWith('<html')) {
        modalBody.innerHTML = `<iframe id="modalIframe" style="width:100%; height:100%; border:none; background:white; min-height: 80vh;"></iframe>`;
        const iframe = document.getElementById('modalIframe');
        const doc = iframe.contentWindow.document;
        doc.open();
        doc.write(content);
        doc.close();
    } else if (state.outputFormat === 'JSON' || state.outputFormat === 'CSV') {
        // Code/Data display
        modalBody.innerHTML = `<pre style="background: #f8f9fa; padding: 1.5rem; border-radius: 8px; overflow: auto; max-height: 70vh; font-size: 0.85rem; border: 1px solid #eee; margin: 0; color: #333; text-align: left;"><code>${escapeHtml(content)}</code></pre>`;
    } else {
        // Standard Fragment
        modalBody.innerHTML = content;
    }

    resultModal.classList.add('active');
    document.body.style.overflow = 'hidden';

    // Update Labeling for Download/Copy buttons
    const format = state.outputFormat || 'HTML';
    if (downloadResultBtn) {
        downloadResultBtn.querySelector('span') || (downloadResultBtn.innerHTML = `
            <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                <polyline points="7 10 12 15 17 10"></polyline>
                <line x1="12" y1="15" x2="12" y2="3"></line>
            </svg>
            <span>Download ${format}</span>
        `);
        downloadResultBtn.querySelector('span').textContent = `Download ${format}`;
    }
    if (copyResultBtn) {
        copyResultBtn.querySelector('span') || (copyResultBtn.innerHTML = `
            <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
                <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
            </svg>
            <span>Copy ${format}</span>
        `);
        copyResultBtn.querySelector('span').textContent = `Copy ${format}`;
    }
}

// Helper to escape HTML for code display
function escapeHtml(text) {
    if (typeof text !== 'string') return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function closeModal() {
    if (!resultModal) return;
    resultModal.classList.remove('active');
    document.body.style.overflow = '';
}

if (closeModalBtn) closeModalBtn.addEventListener('click', closeModal);

// Helper to update persistent preview state
function updatePreviewState(content) {
    state.lastGeneratedHtml = content;
    outputPreview.classList.add("active");
    
    const header = generatePreviewHeader(); // Always generate header
    
    // Full HTML Document -> Use Iframe
    if (content.trim().startsWith('<!DOCTYPE') || content.trim().startsWith('<html')) {
        outputPreview.innerHTML = header + `<div style="height: 600px; border-radius: 8px; overflow: hidden; border: 1px solid var(--border-color);"><iframe id="previewIframe" style="width:100%; height:100%; border:none; background:white;"></iframe></div>`;
        const iframe = document.getElementById('previewIframe');
        const doc = iframe.contentWindow.document;
        doc.open();
        doc.write(content);
        doc.close();
    } else if (state.outputFormat === 'JSON' || state.outputFormat === 'CSV') {
        // Code/Data display
        outputPreview.innerHTML = header + `
            <div style="background: #f8f9fa; padding: 2.5rem 1.5rem; border-radius: 12px; border: 2px dashed #e9ecef; text-align: center;">
                <div style="margin-bottom: 1.5rem; color: var(--accent-mint);">
                    <svg style="width: 48px; height: 48px;" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                        <polyline points="14 2 14 8 20 8"></polyline>
                        <line x1="16" y1="13" x2="8" y2="13"></line>
                        <line x1="16" y1="17" x2="8" y2="17"></line>
                        <polyline points="10 9 9 9 8 9"></polyline>
                    </svg>
                </div>
                <p style="font-weight: 600; color: var(--text-primary); margin-bottom: 0.5rem; font-size: 1.1rem;">${state.outputFormat} Ready</p>
                <p style="font-size: 0.9rem; color: var(--text-secondary); margin-bottom: 1.5rem;">The data has been structured and is ready for export.</p>
                <div style="background: white; padding: 1.25rem; border-radius: 8px; border: 1px solid #ddd; max-height: 120px; overflow: hidden; font-family: 'Cascadia Code', 'Consolas', monospace; font-size: 11px; opacity: 0.7; text-align: left; line-height: 1.5;">
                    ${escapeHtml(content.substring(0, 300))}...
                </div>
            </div>`;
    } else {
        // Standard Fragment (Table)
        outputPreview.innerHTML = header + content;
    }
    
    // Update Preview Button to "Open Result"
    previewBtn.innerHTML = `
        <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="15 3 21 3 21 9"></polyline>
            <polyline points="9 21 3 21 3 15"></polyline>
            <line x1="21" y1="3" x2="14" y2="10"></line>
            <line x1="3" y1="21" x2="10" y2="14"></line>
        </svg>
        <span>Open Result</span>
    `;
    previewBtn.disabled = false;
}

// Generate Preview Header HTML
function generatePreviewHeader() {
    const format = state.outputFormat || 'HTML';
    return `
    <div class="preview-header-bar" style="display: flex; gap: 0.5rem; margin: 0 0 0.5rem 0; padding: 0.5rem 0; border-bottom: 1px solid var(--border-color); justify-content: space-between; align-items: center;">
         <div style="display: flex; gap: 0.5rem; flex-wrap: wrap; margin-left: auto;">
             <button class="btn-micro" onclick="window.downloadResult()">
                Download HTML
                <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px;height:14px;margin-left:4px;">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                    <polyline points="7 10 12 15 17 10"></polyline>
                    <line x1="12" y1="15" x2="12" y2="3"></line>
                </svg>
             </button>
             <button class="btn-micro" onclick="window.exportJson()">
                Export JSON
                <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px;height:14px;margin-left:4px;">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                    <polyline points="14 2 14 8 20 8"></polyline>
                </svg>
             </button>
             <button class="btn-micro" onclick="window.copyResultToClipboard()">
                Copy
                <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="width:14px;height:14px;margin-left:4px;">
                    <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
                    <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
                </svg>
             </button>
         </div>
    </div>`;
}

// --- REFACTORED EXPORT LOGIC ---
async function downloadResult() {
    if (!state.lastGeneratedHtml) return;
    const format = state.outputFormat || 'HTML';
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const extension = format.toLowerCase();
    const type = format === 'HTML' ? 'text/html' : (format === 'JSON' ? 'application/json' : 'text/csv');
    downloadFile(state.lastGeneratedHtml, `conversion-${timestamp}.${extension}`, type);
}

async function copyResultToClipboard() {
    if (!state.lastGeneratedHtml) {
        alert("No content available to copy.");
        return;
    }
    
    try {
        await navigator.clipboard.writeText(state.lastGeneratedHtml);
        // Feedback
        const buttons = document.querySelectorAll('.btn-micro[onclick*="copyResultToClipboard"], #copyResultBtn');
        buttons.forEach(btn => {
            const originalHTML = btn.innerHTML;
            btn.innerHTML = 'Copied!';
            btn.style.background = 'var(--accent-mint)';
            setTimeout(() => {
                btn.innerHTML = originalHTML;
                btn.style.background = '';
            }, 2000);
        });
    } catch (err) {
        console.error("Copy failed:", err);
    }
}

// Global Exposure
window.downloadResult = downloadResult;
window.copyResultToClipboard = copyResultToClipboard;
window.downloadHtml = downloadResult; // Fallback
window.copyHtmlToClipboard = copyResultToClipboard; // Fallback

// Helper: Normalize Rows to JSON Array (Headers = Keys)
function normalizeRowsToJson(rows, sourceType, options = {}) {
    let jsonArray = [];
    const useHeaderRow = options.useHeaderRow !== false; // Default true
    
    if (sourceType === 'recommended') {
        if (!rows || rows.length === 0) return [];
        
        // Determine Headers
        let headers = [];
        let startIndex = 0;

        if (useHeaderRow) {
            headers = rows[0]?.values?.map((cell, index) => {
                return cell.formattedValue || `Column_${index}`;
            }) || [];
            startIndex = 1;
        } else {
            // Generate generic headers based on max columns of first few rows
            const maxCols = rows.slice(0, 10).reduce((max, r) => Math.max(max, r.values?.length || 0), 0);
            for(let i=0; i<maxCols; i++) headers.push(`Column_${i}`);
            startIndex = 0;
        }
        
        // Process Data Rows
        for (let i = startIndex; i < rows.length; i++) {
            const row = rows[i];
            let obj = {};
            if(row.values) {
                row.values.forEach((cell, colIndex) => {
                    const key = headers[colIndex] || `Column_${colIndex}`;
                    obj[key] = cell.formattedValue || null; 
                });
            }
            jsonArray.push(obj);
        }
    } else if (sourceType === 'flash') {
        const numCols = rows.getNumberOfColumns();
        const numRows = rows.getNumberOfRows();
        
        let headers = [];
        let startIndex = 0;

        if (useHeaderRow) {
            for (let c = 0; c < numCols; c++) {
                headers.push(rows.getColumnLabel(c) || `Column_${c}`);
            }
            // GViz often treats labels separate from rows if headers=1 was passed in query
            // But if we want to force NO header row, we might need to be careful.
            // In Flash mode, we queried with `headers=1`? 
            // The `rows` object is a DataTable.
            // If the user says "Don't use header row", we can't easily "un-header" the DataTable labels.
            // But we can just use generic keys.
        } else {
             for (let c = 0; c < numCols; c++) headers.push(`Column_${c}`);
        }
        
        // Rows
        for (let r = 0; r < numRows; r++) {
            let obj = {};
            for (let c = 0; c < numCols; c++) {
                 const key = headers[c];
                 obj[key] = rows.getFormattedValue(r, c) || null;
            }
            jsonArray.push(obj);
        }
    }
    
    return jsonArray;
}

window.exportJson = () => {
    if (!state.lastRawData) { alert("No data to export."); return; }
    
    let exportData = {};
    
    // Normalization logic based on stored mode
    const options = state.customSettings || {}; 
    if (state.lastRawData.mode === 'recommended') {
        // Recommended structure: array of sheet objects
        const sheets = state.lastRawData.data;
        if (Array.isArray(sheets)) {
            sheets.forEach(sheet => {
                const title = sheet.properties.title;
                const rows = sheet.data?.[0]?.rowData || [];
                exportData[title] = normalizeRowsToJson(rows, 'recommended', options);
            });
        }
    } else {
        // Flash structure: GViz DataTable
        const dt = state.lastRawData.data;
        exportData['Sheet1'] = normalizeRowsToJson(dt, 'flash', options);
    }
    
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    downloadFile(JSON.stringify(exportData, null, 2), `sheet-data-${timestamp}.json`, 'application/json');
};

window.exportCsv = () => {
    state.outputFormat = 'CSV';
    handleConvert();
};

// Wire up Buttons
if (downloadResultBtn) downloadResultBtn.addEventListener('click', downloadResult);
if (copyResultBtn) copyResultBtn.addEventListener('click', copyResultToClipboard);
if (downloadHtmlBtn) downloadHtmlBtn.addEventListener('click', downloadResult);
if (exportJsonBtn) exportJsonBtn.addEventListener('click', window.exportJson);

// Close on backdrop click
if (resultModal) {
    resultModal.addEventListener('click', (e) => {
        if (e.target === resultModal) closeModal();
    });
    
    // Close on Escape
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && resultModal.classList.contains('active')) {
            closeModal();
        }
    });
}

// Comparison Slider Logic
let isDragging = false;

function moveSlider(x) {
    const sliderRect = slider.getBoundingClientRect();
    let position = ((x - sliderRect.left) / sliderRect.width) * 100;
    
    // Clamp between 0 and 100
    position = Math.max(0, Math.min(100, position));
    
    sliderOverlay.style.width = position + '%';
    sliderHandle.style.left = position + '%';
}

if (slider) {
    slider.addEventListener('mousedown', (e) => {
        isDragging = true;
        moveSlider(e.clientX);
    });

    // Touch support
    slider.addEventListener('touchstart', (e) => {
        isDragging = true;
        moveSlider(e.touches[0].clientX);
    });
}

window.addEventListener('mouseup', () => {
    isDragging = false;
});

window.addEventListener('mousemove', (e) => {
    if (!isDragging) return;
    moveSlider(e.clientX);
});

window.addEventListener('touchend', () => {
    isDragging = false;
});

window.addEventListener('touchmove', (e) => {
    if (!isDragging) return;
    moveSlider(e.touches[0].clientX);
});

// Set initial overlay image width to match parent container width
// This ensures the image doesn't scale down when the overlay shrinks
function updateOverlayImageWidth() {
    // If not visible yet, do nothing (will be called again by skeleton loader)
    if (slider && slider.offsetParent === null) return;

    const sliderWidth = slider.offsetWidth;
    const overlayImg = sliderOverlay.querySelector('img');
    if (overlayImg) {
        overlayImg.style.width = sliderWidth + 'px';
    }
}

// Update on load and resize
window.addEventListener('load', updateOverlayImageWidth);
window.addEventListener('resize', updateOverlayImageWidth);

// Keyboard shortcuts
document.addEventListener("keydown", (e) => {
  // Ctrl/Cmd + Enter to convert
  if ((e.ctrlKey || e.metaKey) && e.key === "Enter" && !convertBtn.disabled) {
    convertBtn.click();
  }

  // Ctrl/Cmd + P to preview
  if ((e.ctrlKey || e.metaKey) && e.key === "p" && !previewBtn.disabled) {
    e.preventDefault();
    previewBtn.click();
  }
});

// Parallax effect on scroll
let ticking = false;
window.addEventListener("scroll", () => {
  if (!ticking) {
    window.requestAnimationFrame(() => {
      const scrolled = window.pageYOffset;
      const hero = document.querySelector(".hero");
      if (hero) {
        hero.style.transform = `translateY(${scrolled * 0.3}px)`;
        hero.style.opacity = 1 - scrolled / 500;
      }
      ticking = false;
    });
    ticking = true;
  }
});
