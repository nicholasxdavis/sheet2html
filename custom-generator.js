/**
 * Custom Generator for "Ultimate Dashboard" style HTML.
 * Ports FULL logic from generatorrrrrr.html for 1:1 parity.
 */

const CustomGenerator = {
    
    generate: function(rawData, userSettings) {
        // 1. Process Data (Server-side preparation)
        // 1. Process Data (Server-side preparation)
        const normalizeResult = this.cleanAndNormalizeData(rawData);
        const processedData = normalizeResult.activeData;
        const allSheets = normalizeResult.allSheets;
        
        // Pre-calculate schema for each sheet
        allSheets.forEach(sheet => {
            sheet.schema = this.inferSchema(sheet.data);
        });
        
        const dataSchema = allSheets[0].schema;
        
        // 2. Prepare Settings
        const settings = {
            striped: false,
            hover: true,
            borders: false,
            showIndex: true,
            showKPIs: true,
            showSearch: true,
            sortable: true,
            showExport: true,
            showPrint: true, 
            showToolbar: true,
            showAddEntry: true, // Enabled for full parity
            padding: 'normal',
            fontSize: 14,
            fontFamily: 'Inter',
            accentPrimary: '#667eea',
            accentSecondary: '#764ba2',
            rowsPerPage: 'all',
            formatNumbers: true,
            showEmpty: true,
            truncate: true,
            roundness: 'md',
            shadow: 'md',
            theme: 'light',
            fadeIn: true,
            cardHover: true,
            dashboardName: 'Data Dashboard',
            ...userSettings
        };

        // 3. Construct HTML
        return `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${settings.dashboardName}</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&family=Roboto:wght@300;400;500;700;900&family=Poppins:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
    <script src="https://unpkg.com/lucide@latest"></script>
    <style>
        :root {
            --accent-primary: ${settings.accentPrimary};
            --accent-secondary: ${settings.accentSecondary};
            --bg-primary: #ffffff;
            --bg-secondary: #f9fafb;
            --text-primary: #111827;
            --text-secondary: #6b7280;
            --border-color: #e5e7eb;
            --card-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
        }
        
        /* Theme Overrides */
        ${this.getThemeStyles(settings.theme)}
        
        /* GLOBAL SCALING: Set root font-size so tailwind rem units scale */
        html {
            font-size: ${settings.fontSize}px;
        }

        body { 
            font-family: '${settings.fontFamily}', sans-serif; 
            background-color: var(--accent-primary); 
            min-height: 100vh; 
            padding: 1.5rem;
            color: var(--text-primary);
        }
        
        .glassmorphism { 
            background: rgba(255, 255, 255, 0.95); 
            backdrop-filter: blur(10px); 
        }
        /* Dark mode glass override */
        .theme-dark .glassmorphism, .theme-dracula .glassmorphism, .theme-monokai .glassmorphism {
             background: rgba(31, 41, 55, 0.95);
        }
        
        .custom-scrollbar::-webkit-scrollbar { height: 8px; width: 8px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background-color: var(--border-color); border-radius: 20px; }
        
        /* Badges & Utilities */
        ${this.getSupportStyles(settings)}
        
        .dashboard-card { 
            background: var(--bg-primary); 
            border: 1px solid var(--border-color); 
            transition: all 0.3s ease; 
        }
        .dashboard-card:hover { 
            ${settings.cardHover ? 'transform: translateY(-4px); box-shadow: 0 20px 25px -5px rgba(0,0,0,0.1);' : ''}
        }
        
        /* Dynamic Roundness/Shadow */
        .roundness-none { border-radius: 0 !important; }
        .roundness-sm { border-radius: 0.375rem !important; }
        .roundness-md { border-radius: 0.75rem !important; }
        .roundness-lg { border-radius: 1rem !important; }
        .roundness-xl { border-radius: 1.5rem !important; }
        
        .shadow-none { box-shadow: none !important; }
        .shadow-sm { box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05) !important; }
        .shadow-md { box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1) !important; }
        .shadow-lg { box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1) !important; }
        .shadow-xl { box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1) !important; }
        .shadow-2xl { box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25) !important; }
        
        /* Table Styles */
        /* FIXED LAYOUT FOR OVERFLOW PROTECTION */
        table { table-layout: fixed; width: 100%; }
        td, th { word-wrap: break-word; overflow-wrap: break-word; }
        
        .table-striped tbody tr:nth-child(even) { background-color: var(--bg-secondary); }
        .table-hover tbody tr:hover { background-color: ${settings.theme === 'light' ? '#f3f4f6' : 'rgba(255,255,255,0.05)'}; cursor: pointer; }
        .table-bordered { border: 2px solid var(--border-color); }
        .table-bordered th, .table-bordered td { border: 1px solid var(--border-color); }
        /* Airy Table Spacing */
        .table-compact td, .table-compact th { padding: 0.75rem 1rem !important; }
        .table-spacious td, .table-spacious th { padding: 2rem 2.5rem !important; }
        
        /* Clean Header Styling */
        thead th {
            background-color: transparent; /* Cleaner to let bg-primary or secondary show through */
            border-bottom: 2px solid var(--border-color);
            color: var(--text-secondary);
            font-size: 0.7rem; /* Tiny, crisp */
            font-weight: 700;
            letter-spacing: 0.1em; /* Tracking Widest */
            text-transform: uppercase;
            padding-top: 1.5rem !important;
            padding-bottom: 1.5rem !important;
        }
        
        tbody td {
             padding-top: 1.25rem !important;
             padding-bottom: 1.25rem !important;
             border-bottom: 1px solid var(--border-color); /* Subtle horizontal divider */
             color: var(--text-primary);
             font-size: 0.9rem;
        }
        
        /* Hover Effect */
        tbody tr:hover td {
            background-color: var(--bg-secondary);
        }
        
        /* Sort Icons */
        .sort-asc::after { content: " ↑"; color: var(--accent-primary); font-weight: bold; }
        .sort-desc::after { content: " ↓"; color: var(--accent-primary); font-weight: bold; }
        
        .fade-in { animation: fadeIn 0.5s ease-in; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }

        /* Responsive */
        @media (max-width: 768px) {
            .dashboard-card { padding: 1rem !important; }
            body { padding: 0.5rem !important; }
            .mobile-hide { display: none !important; }
            #header-buttons { flex-direction: column; width: 100%; }
            #header-buttons button { width: 100%; justify-content: center; }
            table { display: block; overflow-x: auto; white-space: nowrap; }
        }

        /* Modal */
        .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; overflow: auto; background-color: rgba(0,0,0,0.5); backdrop-filter: blur(4px); }
        .modal.active { display: flex; align-items: center; justify-content: center; }
        .modal-content { background-color: var(--bg-primary); margin: auto; padding: 2rem; border: 1px solid var(--border-color); width: 90%; max-width: 600px; max-height: 80vh; overflow-y: auto; border-radius: 1rem; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.2); animation: slideIn 0.3s ease; }
        @keyframes slideIn { from { transform: translateY(-50px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }

        /* Toast */
        .toast { position: fixed; top: 20px; right: 20px; padding: 1rem 1.5rem; border-radius: 0.75rem; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1); z-index: 9999; animation: slideInRight 0.3s ease; background: white; color: black; border-left: 4px solid var(--accent-primary); }
        @keyframes slideInRight { from { transform: translateX(400px); opacity: 0; } to { transform: translateX(0); opacity: 1; } }
    </style>
</head>
<body class="${settings.fadeIn ? 'fade-in' : ''}">

    <div id="dashboard-section" class="max-w-[1800px] mx-auto">
        
        <!-- Header -->
        <div class="glassmorphism dashboard-card p-6 mb-6 roundness-${settings.roundness} shadow-${settings.shadow}">
            <div class="flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div class="flex items-center gap-4">
                    <!-- Icon Removed -->
                    <div>
                        <h2 class="text-2xl font-bold flex items-center gap-2" style="color: var(--text-primary);">
                            ${settings.dashboardName}
                            <span class="text-xs font-medium bg-green-100 text-green-700 px-3 py-1 rounded-full">Live</span>
                        </h2>
                        <p class="text-xs" style="color: var(--text-secondary);">
                            <span id="row-count-header">${processedData.length}</span> rows • 
                            <span id="col-count-header">${Object.keys(dataSchema).length}</span> columns • 
                            Last updated: <span id="last-updated">just now</span>
                        </p>
                    </div>
                </div>
                <!-- Interactive Header Buttons -->
                <div class="flex items-center gap-3 flex-wrap" id="header-buttons">
                    ${settings.showExport ? `
                    <button onclick="exportCSV()" class="bg-black hover:bg-gray-800 text-white px-5 py-2.5 rounded-lg text-sm font-medium transition flex items-center gap-2 shadow-sm hover:shadow-md">
                        <i data-lucide="download" class="w-4 h-4"></i> CSV
                    </button>
                    <button onclick="exportJSON()" class="bg-black hover:bg-gray-800 text-white px-5 py-2.5 rounded-lg text-sm font-medium transition flex items-center gap-2 shadow-sm hover:shadow-md">
                        <i data-lucide="file-json" class="w-4 h-4"></i> JSON
                    </button>` : ''}
                    ${settings.showPrint ? `
                    <button onclick="window.print()" class="bg-gray-700 hover:bg-gray-800 text-white px-5 py-2.5 rounded-lg text-sm font-medium transition flex items-center gap-2 shadow-sm hover:shadow-md">
                        <i data-lucide="printer" class="w-4 h-4"></i> Print
                    </button>` : ''}
                </div>
            </div>
        </div>

        <!-- Toolbar -->
        ${settings.showToolbar ? `
        <div id="toolbar" class="glassmorphism dashboard-card p-4 mb-6 roundness-${settings.roundness} shadow-${settings.shadow}">
            <div class="flex items-center justify-between flex-wrap gap-4">
                <div id="active-filter-controls" class="flex items-center gap-4 text-sm font-medium" style="color: var(--text-secondary); display: none;">
                    <button onclick="clearFilters()" class="flex items-center gap-1.5 hover:bg-gray-100 transition p-2 rounded-lg" style="color: var(--text-secondary);">
                        <i data-lucide="x-circle" class="w-4 h-4"></i> Clear Filters
                    </button>
                    
                    <span class="text-xs mobile-hide border-l pl-4 border-gray-300">
                        Showing <span id="filtered-count" class="font-bold" style="color: var(--text-primary);">${processedData.length}</span> of <span id="total-count" class="font-bold" style="color: var(--text-primary);">${processedData.length}</span> rows
                    </span>
                </div>
                <div class="flex gap-2 flex-wrap">
                    ${settings.showSearch ? `
                    <div class="relative" id="search-container">
                        <i data-lucide="search" class="w-4 h-4 absolute left-3 top-2.5 text-gray-400"></i>
                        <input type="text" id="search-box" onkeyup="searchTable()" placeholder="Search all columns..." class="pl-9 pr-4 py-2 bg-white border-2 rounded-lg text-sm focus:outline-none focus:ring-2 focus:border-transparent" style="border-color: var(--border-color);">
                    </div>` : ''}
                    ${settings.showAddEntry ? `
                    <button id="add-entry-btn" onclick="openAddEntryModal()" class="bg-gradient-to-r from-green-600 to-green-500 hover:from-green-700 hover:to-green-600 text-white px-4 py-2 rounded-lg text-sm font-bold transition shadow-sm hover:shadow-md flex items-center gap-2">
                        <i data-lucide="plus-circle" class="w-4 h-4"></i> Add Entry
                    </button>` : ''}
                </div>
            </div>
        </div>` : ''}

        <!-- KPI Cards -->
        <div id="sheet-tabs" class="flex gap-2 mb-4 overflow-x-auto pb-2" style="display: ${allSheets.length > 1 ? 'flex' : 'none'};"></div>
        <div id="kpi-container" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6 mb-8" style="display: ${settings.showKPIs ? 'grid' : 'none'};"></div>

        <!-- Data Table -->
        <div id="table-container" class="glassmorphism dashboard-card overflow-hidden roundness-${settings.roundness} shadow-${settings.shadow} ${settings.borders ? 'table-bordered' : ''} ${settings.padding === 'compact' ? 'table-compact' : ''} ${settings.padding === 'spacious' ? 'table-spacious' : ''}">
            <div class="overflow-x-auto custom-scrollbar">
                <table id="data-table" class="w-full text-left border-collapse ${settings.striped ? 'table-striped' : ''} ${settings.hover ? 'table-hover' : ''} " style="font-size: ${settings.fontSize}px;">
                    <thead>
                        <tr class="bg-gradient-to-r from-gray-50 to-gray-100 border-b-2 text-xs uppercase tracking-wider font-bold" style="border-color: var(--border-color); color: var(--text-secondary);" id="table-header">
                            <!-- Headers injected via JS -->
                        </tr>
                    </thead>
                    <tbody class="text-sm divide-y" style="color: var(--text-primary); border-color: var(--border-color);" id="table-body">
                         <!-- Body injected via JS -->
                    </tbody>
                </table>
            </div>
            <div class="p-6 border-t-2 flex justify-between items-center flex-wrap gap-4" style="background: var(--bg-secondary); border-color: var(--border-color);">
                <div class="text-sm font-medium" style="color: var(--text-secondary);">
                    Dashboard generated • <span id="row-count" class="font-bold" style="color: var(--text-primary);">${processedData.length}</span> rows • <span id="col-count" class="font-bold" style="color: var(--text-primary);">${Object.keys(dataSchema).length}</span> columns
                </div>
                <div class="flex gap-2 flex-wrap" id="pagination-container"></div>
            </div>
        </div>
    </div>
    
    <!-- Toast Notification Container -->
    <div id="toast-container"></div>

    <!-- Add Entry Modal -->
    <div id="add-entry-modal" class="modal">
        <div class="modal-content">
            <div class="flex items-center justify-between mb-6">
                <h3 class="text-2xl font-bold flex items-center gap-2" style="color: var(--text-primary);">
                    <i data-lucide="plus-square" class="w-6 h-6"></i>
                    Add New Entry
                </h3>
                <button onclick="closeAddEntryModal()" class="text-gray-500 hover:text-gray-900 transition">
                    <i data-lucide="x" class="w-6 h-6"></i>
                </button>
            </div>
            <form id="add-entry-form" class="space-y-4">
                <div id="entry-fields"></div>
                <div class="flex gap-3 pt-4">
                    <button type="button" onclick="closeAddEntryModal()" class="flex-1 bg-gray-200 hover:bg-gray-300 text-gray-800 px-6 py-3 rounded-lg font-semibold transition">
                        Cancel
                    </button>
                    <button type="submit" class="flex-1 bg-gradient-to-r from-green-600 to-green-500 hover:from-green-700 hover:to-green-600 text-white px-6 py-3 rounded-lg font-semibold transition shadow-sm hover:shadow-md flex items-center justify-center gap-2">
                        <i data-lucide="check" class="w-5 h-5"></i>
                        Add Entry
                    </button>
                </div>
            </form>
        </div>
    </div>

    <script>
        // EMBEDDED DATA & SETTINGS
        let processedData = ${JSON.stringify(processedData)};
        let dataSchema = ${JSON.stringify(dataSchema)};
        const allSheets = ${JSON.stringify(allSheets)};
        const settings = ${JSON.stringify(settings)};
        
        // GLOBAL STATE
        let filteredData = null;
        let currentPage = 1;
        let currentSort = { column: null, direction: 'asc' };

        // RUNTIME LOGIC (Ported from generatorrrrrr.html)
        ${this.getRuntimeScript()}

        // INIT
        document.addEventListener('DOMContentLoaded', () => {
            renderTable(processedData, dataSchema);
            const kpis = generateKPIs(processedData, dataSchema);
            renderKPIs(kpis);
            lucide.createIcons();
            updateLastUpdated();
            renderTabs();
            
            // Add Entry Modal Init
            if(settings.showAddEntry) {
                 document.getElementById('add-entry-form').addEventListener('submit', function(e) {
                    e.preventDefault();
                    const formData = new FormData(e.target);
                    const newEntry = {};
                    for (let [key, value] of formData.entries()) newEntry[key] = value;
                    processedData.push(newEntry);
                    if (filteredData) filteredData.push(newEntry);
                    
                    renderTable(filteredData || processedData, dataSchema);
                    renderKPIs(generateKPIs(processedData, dataSchema));
                    closeAddEntryModal();
                    showSuccess('Entry added successfully!');
                    e.target.reset();
                });
            }
        });
        
        function updateLastUpdated() {
             const now = new Date();
             const timeStr = now.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });
             const el = document.getElementById('last-updated');
             if(el) el.textContent = timeStr;
        }
    </script>
</body>
</html>`;
    },

    // ---------------- HELPER METHODS (Server-Side) ---------------- //

    getThemeStyles: function(theme) {
        switch(theme) {
            case 'dark': return `
                :root { --bg-primary: #1f2937; --bg-secondary: #111827; --text-primary: #f9fafb; --text-secondary: #d1d5db; --border-color: #374151; }`;
            case 'nord': return `
                :root { --bg-primary: #2e3440; --bg-secondary: #3b4252; --text-primary: #eceff4; --text-secondary: #d8dee9; --border-color: #4c566a; }`;
            case 'solarized': return `
                :root { --bg-primary: #fdf6e3; --bg-secondary: #eee8d5; --text-primary: #657b83; --text-secondary: #93a1a1; --border-color: #93a1a1; }`;
            case 'dracula': return `
                :root { --bg-primary: #282a36; --bg-secondary: #44475a; --text-primary: #f8f8f2; --text-secondary: #6272a4; --border-color: #6272a4; }`;
            case 'monokai': return `
                :root { --bg-primary: #272822; --bg-secondary: #3e3d32; --text-primary: #f8f8f2; --text-secondary: #75715e; --border-color: #75715e; }`;
            default: return ''; // Light is default
        }
    },
    
    getSupportStyles: function(settings) {
        return `
        .badge-youtube { background-color: #fef2f2; color: #991b1b; } 
        .badge-instagram { background-color: #fdf2f8; color: #9d174d; } 
        .badge-tiktok { background-color: #f3f4f6; color: #1f2937; } 
        .badge-twitter { background-color: #f0f9ff; color: #075985; } 
        .badge-linkedin { background-color: #eff6ff; color: #1e40af; } 
        .badge-facebook { background-color: #eff6ff; color: #1e3a8a; } 
        .badge-generic { background-color: #f3f4f6; color: #374151; } 
        .badge-positive { background-color: #dcfce7; color: #166534; border: 1px solid #86efac; }
        .badge-negative { background-color: #fee2e2; color: #991b1b; border: 1px solid #fca5a5; }
        .badge-neutral { background-color: #f3f4f6; color: #374151; border: 1px solid #d1d5db; }
        .badge-warning { background-color: #fef3c7; color: #92400e; border: 1px solid #fcd34d; }
        `;
    },

    // --- LOGIC PORTED FROM generatorrrrrr.html LINE 1023 ---
    
    cleanAndNormalizeData: function(inputData) {
        // EDGE CASE 1: Direct array
        if (Array.isArray(inputData)) {
            const cleaned = this.cleanRows(inputData);
            return {
                activeData: cleaned,
                allSheets: [{ name: 'Sheet1', data: cleaned }]
            };
        }
        
        // EDGE CASE 2: Object with arrays (Google Sheets format, multi-sheet exports)
        if (typeof inputData === 'object' && inputData !== null) {
            const sheets = [];
            
            for (const [key, value] of Object.entries(inputData)) {
                if (Array.isArray(value) && value.length > 0) {
                    const cleaned = this.cleanRows(value);
                    if (cleaned.length > 0) {
                        sheets.push({ name: key, data: cleaned });
                    }
                }
            }
            
            if (sheets.length > 0) {
                // Default to largest sheet for better first impression
                sheets.sort((a, b) => b.data.length - a.data.length);
                return {
                    activeData: sheets[0].data,
                    allSheets: sheets
                };
            }
            
            // EDGE CASE 3: Single object (wrap in array)
            const cleaned = this.cleanRows([inputData]);
            return {
                activeData: cleaned,
                allSheets: [{ name: 'Sheet1', data: cleaned }]
            };
        }
        
        return { activeData: [], allSheets: [] };
    },

    cleanRows: function(rows) {
        if (!rows || rows.length === 0) return [];

        // EDGE CASE 4: Remove completely empty rows
        let cleaned = rows.filter(row => {
            if (!row || typeof row !== 'object') return false;
            const values = Object.values(row);
            return values.some(val => val !== null && val !== undefined && val !== '' && !(typeof val === 'string' && val.trim() === ''));
        });

        if (cleaned.length === 0) return [];

        // EDGE CASE 5: Collect ALL possible keys (headers)
        const allKeys = new Set();
        cleaned.forEach(row => Object.keys(row).forEach(key => allKeys.add(key)));

        // EDGE CASE 6: Detect and prefer named columns
        const columnKeys = Array.from(allKeys).filter(k => k.startsWith('Column_') || k.match(/^[A-Z]$/));
        const namedKeys = Array.from(allKeys).filter(k => !k.startsWith('Column_') && !k.match(/^[A-Z]$/));

        let finalKeys;
        if (namedKeys.length > 0 && columnKeys.length > 0) {
            finalKeys = namedKeys;
        } else {
            finalKeys = Array.from(allKeys);
        }

        // EDGE CASE 7: Normalize all rows
        cleaned = cleaned.map(row => {
            const normalizedRow = {};
            finalKeys.forEach(key => {
                normalizedRow[key] = (row[key] === null || row[key] === undefined) ? '' : row[key];
            });
            return normalizedRow;
        });

        // EDGE CASE 8: Remove columns that are ENTIRELY empty
        const columnHasData = {};
        finalKeys.forEach(key => {
            columnHasData[key] = cleaned.some(row => {
                const val = row[key];
                return val !== '' && val !== null && val !== undefined && !(typeof val === 'string' && val.trim() === '');
            });
        });
        
        const keysWithData = finalKeys.filter(key => columnHasData[key]);
        
        // EDGE CASE 9: Remove empty columns from all rows
        cleaned = cleaned.map(row => {
            const cleanRow = {};
            keysWithData.forEach(key => {
                cleanRow[key] = row[key];
            });
            return cleanRow;
        });

        // EDGE CASE 10: Trim whitespace
        cleaned = cleaned.map(row => {
            const trimmedRow = {};
            Object.entries(row).forEach(([key, value]) => {
                trimmedRow[key] = typeof value === 'string' ? value.trim() : value;
            });
            return trimmedRow;
        });

        return cleaned;
    },

    // --- SCHEMA PORTED FROM generatorrrrrr.html LINE 1137 ---

    inferSchema: function(data) {
        if (!data || data.length === 0) return null;
        const schema = {};
        const headers = Object.keys(data[0]);
        
        headers.forEach(header => {
            const samples = data.slice(0, 20).map(row => row[header]).filter(v => v != null && v !== '');
            schema[header] = this.detectColumnType(samples, header);
        });
        
        return schema;
    },

    detectColumnType: function(samples, headerName) {
         if (samples.length === 0) return { type: 'text', format: 'default' };
         
         const header = headerName.toLowerCase();
         // Helper to parse values for detection (matches generatorrrrrr.html parseValue logic)
         const pVal = (v) => {
             if (typeof v === 'number') return v;
             if (typeof v === 'string') return parseFloat(v.replace(/[$,% ]/g, '')) || 0;
             return 0;
         };

         if (header.includes('date') || header.includes('time')) return { type: 'date', format: 'date' };
         if (header.includes('platform') || header.includes('category') || header.includes('type') || header.includes('status')) return { type: 'category', format: 'badge' };
         if (header.includes('id') && !header.includes('video')) return { type: 'text', format: 'default' };
         
         if (header.includes('$') || header.includes('price') || header.includes('cost') || header.includes('amount') || header.includes('profit') || header.includes('sale') || header.includes('revenue') || header.includes('rev')) {
             const currencyPattern = /\$|,/;
             const hasCurrencyFormat = samples.some(v => currencyPattern.test(String(v)));
             if (hasCurrencyFormat) return { type: 'currency', format: 'money' };
         }
         
         if (header.includes('%') || header.includes('percent') || header.includes('roi') || header.includes('rate')) return { type: 'percentage', format: 'percent' };
         
         if (header.includes('url') || header.includes('link') || header.includes('website')) return { type: 'url', format: 'link' };
         if (header.includes('email')) return { type: 'email', format: 'email' };
         if (header.includes('note') || header.includes('desc') || header.includes('comment')) return { type: 'text', format: 'longtext' };
         
         // 1:1 Port of numeric detection logic from generatorrrrrr.html
         // Filters out 0s to prevent text strings (which parse to 0) from triggering numeric type
         const numericCount = samples.filter(v => !isNaN(pVal(v)) && pVal(v) !== 0 && String(v).trim() !== '').length;
         const numericRatio = numericCount / samples.length;
         
         if (numericRatio > 0.7) {
             const avg = samples.reduce((a,b)=>a+pVal(b),0) / samples.length;
             if(avg > 10000) return { type: 'number', format: 'largeNumber' };
             return { type: 'number', format: 'number' };
         }
         
         return { type: 'text', format: 'default' };
    },

    // ---------------- RUNTIME SCRIPT (INJECTED) ---------------- //
    // This string contains the EXACT functions from generatorrrrrr.html
    
    getRuntimeScript: function() {
        return `
        // --- PARSING & FORMATTING (Ported) ---
        function parseValue(val) {
            if (val === null || val === undefined || val === '') return 0;
            if (typeof val === 'number') return val;
            if (typeof val === 'string') {
                const cleaned = val.replace(/[$,% ]/g, '');
                const parsed = parseFloat(cleaned);
                return isNaN(parsed) ? 0 : parsed;
            }
            return 0;
        }

        function formatValue(value, format) {
            if (!settings.formatNumbers && format !== 'badge' && format !== 'boolean') return value;
            if ((value === null || value === undefined || value === '') && settings.showEmpty) return '-';
            
            const numValue = parseValue(value);
            
            switch (format) {
                case 'money':
                    if (typeof value === 'string' && value.includes('$')) return value;
                    return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(numValue);
                case 'largeNumber':
                    return new Intl.NumberFormat('en-US').format(numValue);
                case 'number':
                    if (numValue === 0 && String(value).trim() === '') return settings.showEmpty ? '-' : '';
                    return numValue.toLocaleString('en-US', { maximumFractionDigits: 2 });
                case 'percent':
                    return String(value).includes('%') ? value : \`\${numValue}%\`;
                case 'date':
                    return value;
                case 'link':
                    return \`<a href="\${value}" target="_blank" class="text-blue-600 hover:underline break-all">\${value}</a>\`;
                case 'email':
                     return \`<a href="mailto:\${value}" class="text-blue-600 hover:underline">\${value}</a>\`;
                case 'longtext':
                     if(!settings.truncate) return value;
                     // Safe way to pass data: use title attribute (already there) and pass 'this.title'
                     return \`<span class="text-gray-500 italic line-clamp-2 block max-w-md cursor-pointer hover:text-gray-800 transition" title="\${value.replace(/"/g, '&quot;')}" onclick="showCellModal(this.title)">\${value}</span>\`;
                default:
                    return value;
            }
        }

        function getPlatformBadgeClass(platform) {
            const p = String(platform).toLowerCase();
            if (p.includes('youtube')) return 'badge-youtube';
            if (p.includes('instagram')) return 'badge-instagram';
            if (p.includes('tiktok')) return 'badge-tiktok';
            if (p.includes('twitter') || p.includes('x')) return 'badge-twitter';
            if (p.includes('linkedin')) return 'badge-linkedin';
            if (p.includes('facebook')) return 'badge-facebook';
            if (p.includes('active') || p.includes('published') || p.includes('completed') || p.includes('delivered') || p.includes('paid')) return 'badge-positive';
            if (p.includes('pending') || p.includes('scheduled')) return 'badge-neutral';
            if (p.includes('failed') || p.includes('cancelled')) return 'badge-negative';
            if (p.includes('warning') || p.includes('review')) return 'badge-warning';
            return 'badge-generic';
        }

        // --- KPIS (Ported & Modified to remove Total Row as requested) ---
        function generateKPIs(data, schema) {
            const kpis = [];
            const headers = Object.keys(schema);
            const numericColumns = headers.filter(h => ['number', 'currency', 'percentage'].includes(schema[h].type));
            const categoryColumns = headers.filter(h => schema[h].type === 'category');
            
            // 1. Total (REMOVED as per user request in Step 360)
            // kpis.push({ title: 'Total Records', value: data.length, format: 'number', icon: 'database' });
            
            // 2. First Numeric
            if (numericColumns.length > 0) {
                const col = numericColumns[0];
                const values = data.map(row => parseValue(row[col]));
                const total = values.reduce((sum, v) => sum + v, 0);
                const avg = total / values.length;
                kpis.push({
                    title: \`Total \${col}\`,
                    value: schema[col].type === 'currency' ? total : (avg > 1000 ? total : avg),
                    format: schema[col].format,
                    trend: '+8.2%', trendType: 'positive',
                    icon: 'trending-up'
                });
            }
            
            // 3. Top Category
            if (categoryColumns.length > 0) {
                 const col = categoryColumns[0];
                 const counts = {};
                 data.forEach(r => { const c = r[col]||'Unk'; counts[c]=(counts[c]||0)+1; });
                 const top = Object.keys(counts).reduce((a,b)=>counts[a]>counts[b]?a:b);
                 kpis.push({
                     title: \`Top \${col}\`, value: top, format: 'text', subtitle: \`\${counts[top]} records\`, icon: 'award'
                 });
            }
            
            // 4. Avg / Unique
            if (numericColumns.length > 1) {
                const col = numericColumns[1];
                const values = data.map(row => parseValue(row[col]));
                const avg = values.reduce((s,v)=>s+v,0)/values.length;
                kpis.push({ title: \`Avg \${col}\`, value: avg, format: schema[col].format, icon: 'activity' });
            } else if (categoryColumns.length > 0) {
                const unique = new Set(data.map(r=>r[categoryColumns[0]])).size;
                kpis.push({ title: 'Unique Categories', value: unique, format: 'number', icon: 'layers' });
            }
            return kpis.slice(0, 4);
        }

        function renderKPIs(kpis) {
            const container = document.getElementById('kpi-container');
            let html = '';
            kpis.forEach((kpi, index) => {
                const formattedValue = kpi.format === 'text' ? kpi.value : formatValue(kpi.value, kpi.format);
                // "Stable", "Primary" style badges from screenshot
                const trendBadge = kpi.trend ? \`<span class="badge-\${kpi.trendType} text-[10px] font-bold px-2 py-0.5 rounded-full uppercase tracking-wide opacity-80 backdrop-blur-sm">\${kpi.trend}</span>\` : '';
                
                html += \`
                    <div class="dashboard-card p-6 roundness-\${settings.roundness} shadow-\${settings.shadow} fade-in flex flex-col justify-between h-32 relative group border border-[var(--border-color)] hover:border-[var(--accent-primary)] transition-colors overflow-hidden" style="animation-delay: \${index * 0.1}s; background: var(--bg-primary);">
                        <div class="flex justify-between items-start">
                            <span class="text-[11px] font-bold uppercase tracking-widest opacity-60" style="color: var(--text-secondary);">\${kpi.title}</span>
                            \${trendBadge}
                        </div>
                        <div class="mt-2">
                            <div class="text-5xl font-extrabold tracking-tighter" style="color: var(--text-primary); letter-spacing: -0.02em;">\${formattedValue}</div>
                            \${kpi.subtitle ? \`<p class="text-xs mt-1 font-medium opacity-50" style="color: var(--text-secondary);">\${kpi.subtitle}</p>\` : ''}
                        </div>
                    </div>\`;
            });
            container.innerHTML = html;
        }

        // --- TABLE RENDERING (Ported) ---
        function renderTable(data, schema) {
            const tableHeader = document.getElementById('table-header');
            const tableBody = document.getElementById('table-body');
            
            if (!data || data.length === 0) {
                tableHeader.innerHTML = '<th class="px-6 py-4">No data available</th>';
                tableBody.innerHTML = '';
                return;
            }
            
            const headers = Object.keys(data[0]);
            
            // Header
            let headerHTML = settings.showIndex ? '<th class="px-6 py-4 w-12 text-center">#</th>' : '';
            headers.forEach(h => {
                const sortClass = currentSort.column === h ? (currentSort.direction === 'asc' ? 'sort-asc' : 'sort-desc') : '';
                const type = schema[h]?.type || 'text';
                const alignClass = ['number', 'currency', 'percentage'].includes(type) ? 'text-right justify-end' : 'text-left justify-start';
                
                // Flex container inside TH needs alignment too
                const flexAlign = ['number', 'currency', 'percentage'].includes(type) ? 'flex-row-reverse' : 'flex-row';

                const clickHandler = settings.sortable ? \`onclick="sortByColumn('\${h}')"\` : '';
                headerHTML += \`<th class="px-6 py-4 \${alignClass} \${settings.sortable ? 'cursor-pointer' : ''} transition select-none \${sortClass}" \${clickHandler}>
                    <div class="flex items-center gap-2 \${flexAlign}">
                        <span>\${h}</span>
                        \${settings.sortable ? '<i data-lucide="chevrons-up-down" class="w-3 h-3 opacity-50"></i>' : ''}
                    </div>
                </th>\`;
            });
            tableHeader.innerHTML = headerHTML;
            
            // Body
            const rowsPerPage = settings.rowsPerPage === 'all' ? data.length : parseInt(settings.rowsPerPage);
            const startIndex = (currentPage - 1) * rowsPerPage;
            const endIndex = startIndex + rowsPerPage;
            const paginatedData = data.slice(startIndex, endIndex);
            
            let bodyHTML = '';
            paginatedData.forEach((row, index) => {
                bodyHTML += '<tr class="transition group">';
                if (settings.showIndex) {
                    bodyHTML += \`<td class="px-6 py-4 text-center text-xs font-bold" style="color: var(--text-secondary);">\${startIndex + index + 1}</td>\`;
                }
                headers.forEach(key => {
                    let cellValue = row[key];
                    const colType = schema[key];
                    let alignClass = ['number', 'currency', 'percentage'].includes(colType.type) ? 'text-right' : 'text-left';
                    
                    // Special case: First text column (usually Title) gets font-medium and dark text
                    if (colType.type === 'text' && key === headers[0]) {
                        alignClass += ' font-semibold text-[var(--text-primary)]';
                    }

                    let cellClass = \`px-6 py-4 \${alignClass}\`;
                    let inner = cellValue;
                    
                    if (colType.type === 'category' && colType.format === 'badge') {
                        inner = \`<span class="\${getPlatformBadgeClass(cellValue)} px-3 py-1.5 rounded-full text-xs font-bold shadow-sm">\${cellValue}</span>\`;
                    } else if (colType.type === 'percentage') {
                         const val = parseValue(cellValue);
                         const color = val >= 0 ? "text-green-600" : "text-red-600";
                         inner = \`<span class="\${color} font-bold">\${formatValue(cellValue, colType.format)}</span>\`;
                    } else if (colType.type === 'currency' || colType.type === 'number') {
                         inner = \`<span class="font-medium" style="color: var(--text-primary);">\${formatValue(cellValue, colType.format)}</span>\`;
                    } else {
                         inner = formatValue(cellValue, colType.format);
                    }
                    bodyHTML += \`<td class="\${cellClass}">\${inner}</td>\`;
                });
                bodyHTML += '</tr>';
            });
            tableBody.innerHTML = bodyHTML;
            
            // Update Counts
            document.getElementById('row-count').textContent = data.length;
            document.getElementById('col-count').textContent = headers.length;
            document.getElementById('row-count-header').textContent = data.length;
            document.getElementById('col-count-header').textContent = headers.length;
            document.getElementById('filtered-count').textContent = data.length;
            document.getElementById('total-count').textContent = processedData.length;
            
            renderPagination(data.length, rowsPerPage);
            lucide.createIcons();
        }

        // --- INTERACTION (Ported) ---
        function renderPagination(totalRows, rowsPerPage) {
            if (settings.rowsPerPage === 'all') {
                document.getElementById('pagination-container').innerHTML = '';
                return;
            }
            
            const totalPages = Math.ceil(totalRows / rowsPerPage);
            const container = document.getElementById('pagination-container');
            
            let html = '';
            if (currentPage > 1) {
                html += \`<button onclick="changePage(\${currentPage - 1})" class="px-3 py-1 bg-white border rounded-lg hover:bg-gray-50 transition text-sm shadow-sm" style="border-color: var(--border-color); color: var(--text-primary);">Previous</button>\`;
            }
            
            for (let i = 1; i <= totalPages; i++) {
                if (i === 1 || i === totalPages || (i >= currentPage - 1 && i <= currentPage + 1)) {
                    const activeClass = i === currentPage ? 'text-white' : 'bg-white hover:bg-gray-50';
                    const activeStyle = i === currentPage ? \`style="background: var(--accent-primary); border-color: var(--accent-primary);"\` : \`style="border-color: var(--border-color); color: var(--text-primary);"\`;
                    html += \`<button onclick="changePage(\${i})" class="px-3 py-1 border rounded-lg transition text-sm \${activeClass} shadow-sm" \${activeStyle}>\${i}</button>\`;
                } else if (i === currentPage - 2 || i === currentPage + 2) {
                    html += '<span class="px-2" style="color: var(--text-secondary);">...</span>';
                }
            }
            
            if (currentPage < totalPages) {
                html += \`<button onclick="changePage(\${currentPage + 1})" class="px-3 py-1 bg-white border rounded-lg hover:bg-gray-50 transition text-sm shadow-sm" style="border-color: var(--border-color); color: var(--text-primary);">Next</button>\`;
            }
            
            container.innerHTML = html;
        }
        function changePage(p) { currentPage = p; renderTable(filteredData || processedData, dataSchema); }
        
        function sortByColumn(col) {
            const data = filteredData || processedData;
            if (currentSort.column === col) currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
            else { currentSort.column = col; currentSort.direction = 'asc'; }
            
            data.sort((a,b) => {
                 const vA = parseValue(a[col]); const vB = parseValue(b[col]);
                 if(!isNaN(vA) && !isNaN(vB) && vA!==0 && vB!==0) return currentSort.direction==='asc'?vA-vB:vB-vA;
                 return currentSort.direction==='asc'?String(a[col]).localeCompare(String(b[col])):String(b[col]).localeCompare(String(a[col]));
            });
            renderTable(data, dataSchema);
        }
        
        function searchTable() {
            const input = document.getElementById('search-box').value.toLowerCase();
            const controls = document.getElementById('active-filter-controls');
            
            if(!input) { 
                filteredData = null; 
                if(controls) controls.style.display = 'none';
                renderTable(processedData, dataSchema); 
                return; 
            }
            
            filteredData = processedData.filter(row => Object.values(row).some(v => String(v).toLowerCase().includes(input)));
            currentPage = 1;
            
            if(controls) controls.style.display = 'flex';
            renderTable(filteredData, dataSchema);
        }
        function clearFilters() {
            document.getElementById('search-box').value = ''; 
            filteredData = null; 
            const controls = document.getElementById('active-filter-controls');
            if(controls) controls.style.display = 'none';
            renderTable(processedData, dataSchema);
        }
        
        // --- ADD ENTRY MODAL ---
        function openAddEntryModal() {
            const modal = document.getElementById('add-entry-modal');
            const fieldsContainer = document.getElementById('entry-fields');
            const headers = Object.keys(dataSchema);
            let fieldsHTML = '';
            headers.forEach(header => {
                fieldsHTML += \`<div><label class="block text-sm font-semibold mb-2">\${header}</label><input type="text" name="\${header}" class="w-full px-4 py-2 border-2 rounded-lg" style="border-color: var(--border-color); background: var(--bg-primary);"></div>\`;
            });
            fieldsContainer.innerHTML = fieldsHTML;
            modal.classList.add('active');
        }
        function closeAddEntryModal() { document.getElementById('add-entry-modal').classList.remove('active'); }
        function showSuccess(msg) {
             const t = document.getElementById('toast-container');
             const d = document.createElement('div');
             d.className = 'toast'; d.textContent = msg;
             t.appendChild(d);
             setTimeout(()=>d.remove(), 3000);
        }

        // --- EXPORT ---
        function exportCSV() {
             const data = filteredData || processedData;
             const headers = Object.keys(data[0]);
             let csv = headers.join(',') + '\\n';
             data.forEach(row => {
                 csv += headers.map(h => {
                     let val = row[h]; if(val===null||val===undefined) return '""';
                     val = String(val); if(val.includes(',') || val.includes('"')) return '"'+val.replace(/"/g,'""')+'"';
                     return val;
                 }).join(',') + '\\n';
             });
             const blob = new Blob([csv], {type:'text/csv'});
             const url = URL.createObjectURL(blob);
             const a = document.createElement('a'); a.href=url; a.download='export.csv'; a.click();
        }

        function renderTabs() {
            const container = document.getElementById('sheet-tabs');
            if (!container || !allSheets || allSheets.length <= 1) { 
                if(container) container.style.display = 'none'; 
                return; 
            }
            
            // Initial Active Index
            if (typeof window.currentSheetIndex === 'undefined') window.currentSheetIndex = 0;
            
            let html = '';
            allSheets.forEach((sheet, index) => {
                const isActive = index === window.currentSheetIndex;
                // Styles
                const activeClass = isActive 
                    ? 'bg-black text-white shadow-md transform scale-105' 
                    : 'bg-white text-gray-700 hover:bg-gray-100 border border-gray-200';
                
                html += \`<button onclick="switchSheet(\${index})" class="px-4 py-2 rounded-lg text-sm font-semibold transition whitespace-nowrap \${activeClass}">\${sheet.name}</button>\`;
            });
            
            container.innerHTML = html;
            container.style.display = 'flex';
        }

        function switchSheet(index) {
            window.currentSheetIndex = index;
            processedData = allSheets[index].data;
            dataSchema = allSheets[index].schema;
            
            // Reset state
            filteredData = null;
            currentPage = 1;
            
            // Re-render
            renderTable(processedData, dataSchema);
            renderKPIs(generateKPIs(processedData, dataSchema));
            renderHeaderButtons();
            renderTabs(); // Update active tab
            updateLastUpdated();
            
            // Clear search
            const searchBox = document.getElementById('search-box');
            if(searchBox) searchBox.value = '';
        }

        // --- HEADER BUTTONS ---
        function renderHeaderButtons() {
            const container = document.getElementById('header-buttons');
            if(!container) return;
            let buttonsHTML = '';
            
            if (settings.showExport) {
                buttonsHTML += \`
                    <button onclick="exportCSV()" class="bg-black hover:bg-gray-800 text-white px-5 py-2.5 rounded-lg text-sm font-medium transition flex items-center gap-2 shadow-sm hover:shadow-md">
                        <i data-lucide="download" class="w-4 h-4"></i> CSV
                    </button>
                    <button onclick="exportJSON()" class="bg-black hover:bg-gray-800 text-white px-5 py-2.5 rounded-lg text-sm font-medium transition flex items-center gap-2 shadow-sm hover:shadow-md">
                        <i data-lucide="file-json" class="w-4 h-4"></i> JSON
                    </button>
                \`;
            }
            
            if (settings.showPrint) {
                buttonsHTML += \`
                    <button onclick="printDashboard()" class="bg-gray-700 hover:bg-gray-800 text-white px-5 py-2.5 rounded-lg text-sm font-medium transition flex items-center gap-2 shadow-sm hover:shadow-md">
                        <i data-lucide="printer" class="w-4 h-4"></i> Print
                    </button>
                \`;
            }
            container.innerHTML = buttonsHTML;
            if(window.lucide) window.lucide.createIcons();
        }

        function exportJSON() {
            if (!processedData) return;
            const json = JSON.stringify(processedData, null, 2);
            downloadFile(json, \`\${settings.dashboardName.replace(/\\s+/g, '_')}_export.json\`, 'application/json');
        }

        function printDashboard() {
            window.print();
        }
        
        function downloadFile(content, filename, type) {
            const blob = new Blob([content], { type: type });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            a.click();
            window.URL.revokeObjectURL(url);
        }

        // --- CELL MODAL ---
        function showCellModal(content) {
            let modal = document.getElementById('cell-detail-modal');
            if (!modal) {
                const div = document.createElement('div');
                div.id = 'cell-detail-modal';
                div.className = 'fixed inset-0 z-50 hidden flex items-center justify-center';
                div.innerHTML = \`
                    <div class="absolute inset-0 bg-black/50 backdrop-blur-sm transition-opacity" onclick="closeCellModal()"></div>
                    <div class="relative bg-white rounded-xl shadow-2xl max-w-2xl w-full mx-4 overflow-hidden transform transition-all scale-95 opacity-0" id="cell-modal-content">
                        <div class="p-6 max-h-[70vh] overflow-y-auto custom-scrollbar">
                            <p class="text-gray-800 text-lg leading-relaxed whitespace-pre-wrap font-medium" id="cell-modal-text"></p>
                        </div>
                        <div class="bg-gray-50 px-6 py-4 flex justify-end border-t border-gray-100">
                            <button onclick="closeCellModal()" class="px-4 py-2 bg-gray-800 text-white text-sm font-medium rounded-lg hover:bg-gray-900 transition">Close</button>
                        </div>
                    </div>
                \`;
                document.body.appendChild(div);
                modal = div;
            }
            
            document.getElementById('cell-modal-text').textContent = content;
            modal.classList.remove('hidden');
            
            // Animation
            const contentDiv = document.getElementById('cell-modal-content');
            requestAnimationFrame(() => {
                contentDiv.classList.remove('scale-95', 'opacity-0');
                contentDiv.classList.add('scale-100', 'opacity-100');
            });
        }

        function closeCellModal() {
            const modal = document.getElementById('cell-detail-modal');
            if (!modal) return;
            
            const contentDiv = document.getElementById('cell-modal-content');
            contentDiv.classList.remove('scale-100', 'opacity-100');
            contentDiv.classList.add('scale-95', 'opacity-0');
            
            setTimeout(() => {
                modal.classList.add('hidden');
            }, 200);
        }
        `;
    }
};
