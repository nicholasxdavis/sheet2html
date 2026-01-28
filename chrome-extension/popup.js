/**
 * sheets2json Extension Popup Logic - Simplified
 * Always uses API Key (Flash Mode) and First Row Headers
 */

const API_KEY = 'AIzaSyB46Fth--G1-nFy0KXyJbPZN1D71wI4TKw';

const state = {
    spreadsheetId: null,
    gid: null,
    sheetTitle: null,
    lastResult: null, // JSON object
    format: 'json' // 'json' or 'csv'
};

// DOM Elements
const sheetTitleEl = document.getElementById('sheetTitle');
const sheetIdEl = document.getElementById('sheetId');
const convertBtn = document.getElementById('convertBtn');
const previewArea = document.getElementById('previewArea');
const copyBtn = document.getElementById('copyBtn');
const downloadBtn = document.getElementById('downloadBtn');
const resetBtn = document.getElementById('resetBtn');
const toggleBtns = document.querySelectorAll('.toggle-btn');
const inputSection = document.getElementById('inputSection');
const previewSection = document.getElementById('previewSection');

// Initialize
document.addEventListener('DOMContentLoaded', async () => {
    initDetection();
    setupEventListeners();
});

// 1. Detect Sheet from Active Tab
async function initDetection() {
    try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (tab && tab.url.includes('docs.google.com/spreadsheets/d/')) {
            const matches = tab.url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
            const gidMatch = tab.url.match(/[#&]gid=([0-9]+)/);
            
            state.spreadsheetId = matches ? matches[1] : null;
            state.gid = gidMatch ? gidMatch[1] : null;
            state.sheetTitle = tab.title.split(' - Google Sheets')[0];

            sheetTitleEl.textContent = state.sheetTitle;
            sheetIdEl.textContent = state.spreadsheetId;
            convertBtn.disabled = false;
        } else {
            sheetTitleEl.textContent = 'No Sheet Detected';
            sheetIdEl.textContent = 'Open a Google Sheet';
            convertBtn.disabled = true;
        }
    } catch (err) {
        console.error('Detection error:', err);
    }
}

// 2. Event Listeners
function setupEventListeners() {
    convertBtn.addEventListener('click', handleExport);
    
    toggleBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            toggleBtns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.format = btn.dataset.format;
        });
    });

    copyBtn.addEventListener('click', () => {
        if (!state.lastResult) return;
        const content = getFormattedContent();
        navigator.clipboard.writeText(content);
        const originalHtml = copyBtn.innerHTML;
        copyBtn.innerHTML = '✓';
        setTimeout(() => copyBtn.innerHTML = originalHtml, 2000);
    });

    downloadBtn.addEventListener('click', () => {
        if (!state.lastResult) return;
        const content = getFormattedContent();
        const type = state.format === 'json' ? 'application/json' : 'text/csv';
        const ext = state.format;
        
        const blob = new Blob([content], { type: type });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${state.sheetTitle || 'sheet'}.${ext}`;
        a.click();
        URL.revokeObjectURL(url);
    });

    resetBtn.addEventListener('click', handleReset);
}

// 3. UI State Management
function showPreview() {
    inputSection.classList.add('hidden');
    previewSection.classList.remove('hidden');
    renderPreview();
}

function handleReset() {
    state.lastResult = null;
    previewArea.innerHTML = '<div class="placeholder">Click Export to generate...</div>';
    previewSection.classList.add('hidden');
    inputSection.classList.remove('hidden');
}

// 4. Fetch & Export Logic
async function handleExport() {
    convertBtn.disabled = true;
    const originalContent = convertBtn.innerHTML;
    convertBtn.innerHTML = '<div class="loading-spinner"></div>';
    
    try {
        // Always Fast/API Key Fetch
        const data = await fetchWithApiKey(state.spreadsheetId);
        const normalized = normalizeData(data);
        state.lastResult = normalized;
        showPreview();

    } catch (err) {
        console.error('Export error:', err);
        // On error, we might want to stay on input screen or show error in a toast
        // For now, let's briefly show error on button or alert
        alert(`Error: ${err.message}. Ensure the sheet is Public.`);
    } finally {
        convertBtn.disabled = false;
        convertBtn.innerHTML = originalContent;
    }
}

async function fetchWithApiKey(spreadsheetId) {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?includeGridData=true&key=${API_KEY}`;
    const response = await fetch(url);
    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error?.message || 'API Request Failed');
    }
    const result = await response.json();
    return result.sheets;
}

// 5. Normalization Logic
function normalizeData(sheets) {
    const result = {};
    
    sheets.forEach(sheet => {
        const title = sheet.properties.title;
        const rows = sheet.data?.[0]?.rowData || [];
        
        if (rows.length === 0) {
            result[title] = [];
            return;
        }

        // Always use first row as headers
        const headers = rows[0]?.values?.map((cell, index) => {
            return cell.formattedValue || `Column_${index}`;
        }) || [];

        const jsonArray = [];
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            let obj = {};
            if (row.values) {
                row.values.forEach((cell, colIndex) => {
                    const key = headers[colIndex] || `Column_${colIndex}`;
                    obj[key] = cell.formattedValue !== undefined ? cell.formattedValue : null;
                });
                
                // Only push if not completely empty
                if (Object.values(obj).some(v => v !== null && v !== "")) {
                    jsonArray.push(obj);
                }
            }
        }
        result[title] = jsonArray;
    });

    return result;
}

// 6. CSV Logic
function jsonToCsv(data) {
    let allRows = [];
    let headers = new Set();

    // Flatten multi-sheet structure
    const sheets = typeof data === 'object' && !Array.isArray(data) ? Object.values(data) : [data];
    
    sheets.forEach(sheetRows => {
        if (!Array.isArray(sheetRows)) return;
        sheetRows.forEach(row => {
            Object.keys(row).forEach(h => headers.add(h));
            allRows.push(row);
        });
    });

    if (allRows.length === 0) return "";

    const headerArray = Array.from(headers);
    const csvRows = [];
    
    // Add Header
    csvRows.push(headerArray.join(','));

    // Add Data
    allRows.forEach(row => {
        const values = headerArray.map(h => {
            const val = row[h] === undefined || row[h] === null ? "" : String(row[h]);
            // Escape quotes and wrap in quotes if contains comma
            const escaped = val.replace(/"/g, '""');
            return `"${escaped}"`;
        });
        csvRows.push(values.join(','));
    });

    return csvRows.join('\n');
}

function getFormattedContent() {
    if (state.format === 'json') {
        return JSON.stringify(state.lastResult, null, 2);
    } else {
        return jsonToCsv(state.lastResult);
    }
}

function renderPreview() {
    const content = getFormattedContent();
    previewArea.innerHTML = `<pre style="margin: 0; white-space: pre-wrap;">${escapeHtml(content)}</pre>`;
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
