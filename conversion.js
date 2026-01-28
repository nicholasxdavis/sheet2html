// Sheet Conversion Logic (Flash & Recommended Models)

// Main logic to render the spreadsheet data to HTML (Recommended Model)
function renderSheetToHtml(sheet) {
  // sheet is now the individual sheet object, not the whole response
  const data = sheet.data ? sheet.data[0] : null; 
  if (!data || !data.rowData) return '<p>No data found in sheet.</p>';

  const rowData = data.rowData;
  const colMetadata = data.columnMetadata || [];
  const rowMetadata = data.rowMetadata || [];
  const merges = sheet.merges || [];
  
  // 1. Pre-process merges
  const mergeMap = new Map();
  merges.forEach(merge => {
      const startRow = merge.startRowIndex;
      const endRow = merge.endRowIndex;
      const startCol = merge.startColumnIndex;
      const endCol = merge.endColumnIndex;
      
      mergeMap.set(`${startRow},${startCol}`, {
          rowspan: endRow - startRow,
          colspan: endCol - startCol
      });
      
      for (let r = startRow; r < endRow; r++) {
          for (let c = startCol; c < endCol; c++) {
              if (r === startRow && c === startCol) continue;
              mergeMap.set(`${r},${c}`, 'skip');
          }
      }
  });

  // Start Table
  let html = '<table style="border-collapse: collapse; font-family: Arial, sans-serif; table-layout: fixed;">';
  
  // 2. ColGroup for Column Widths
  if (colMetadata.length > 0) {
      html += '<colgroup>';
      colMetadata.forEach(col => {
          const width = col.pixelSize || 100; // Default 100px if undefined
          html += `<col style="width: ${width}px;">`;
      });
      html += '</colgroup>';
  }

  html += '<tbody>';

  rowData.forEach((row, rowIndex) => {
    // 3. Row Height
    const rowMeta = rowMetadata[rowIndex];
    const height = rowMeta && rowMeta.pixelSize ? `${rowMeta.pixelSize}px` : 'auto';
    html += `<tr style="height: ${height};">`;
    
    // Some rows might be empty or missing values
    if (row.values) {
        row.values.forEach((cell, colIndex) => {
            const mergeStatus = mergeMap.get(`${rowIndex},${colIndex}`);
            
            if (mergeStatus === 'skip') {
                return; 
            }
            
            let style = 'padding: 2px 3px; overflow: hidden;'; // Base style
            let content = (cell.formattedValue || '');
            let tagName = 'td'; // Could infer 'th' but Sheets doesn't strictly define header rows.
            
            // Apply Styles from effectiveFormat
            if (cell.effectiveFormat) {
                const fmt = cell.effectiveFormat;
                
                // Background
                if (fmt.backgroundColor) {
                    const bg = colorToRgb(fmt.backgroundColor);
                    if (bg && bg !== 'rgb(255, 255, 255)') { 
                        style += `background-color: ${bg};`;
                    }
                }
                
                // Font
                if (fmt.textFormat) {
                    const tf = fmt.textFormat;
                    if (tf.bold) style += 'font-weight: bold;';
                    if (tf.italic) style += 'font-style: italic;';
                    if (tf.strikethrough) style += 'text-decoration: line-through;';
                    if (tf.underline) style += 'text-decoration: underline;';
                    if (tf.foregroundColor) {
                        const fg = colorToRgb(tf.foregroundColor);
                        if (fg && fg !== 'rgb(0, 0, 0)') { 
                            style += `color: ${fg};`;
                        }
                    }
                    if (tf.fontSize) {
                        style += `font-size: ${tf.fontSize}pt;`;
                    }
                    if (tf.fontFamily) {
                        style += `font-family: '${tf.fontFamily}', sans-serif;`;
                    }
                }
                
                // Alignment
                if (fmt.horizontalAlignment) {
                    const align = fmt.horizontalAlignment.toLowerCase();
                    style += `text-align: ${align};`;
                }
                
                if (fmt.verticalAlignment) {
                    const valign = fmt.verticalAlignment.toLowerCase();
                    const vMap = { 'top': 'top', 'middle': 'middle', 'bottom': 'bottom' };
                    style += `vertical-align: ${vMap[valign] || 'bottom'};`;
                }
                
                // Borders
                if (fmt.borders) {
                    if (fmt.borders.top) style += `border-top: ${mapBorderStyle(fmt.borders.top)};`;
                    if (fmt.borders.bottom) style += `border-bottom: ${mapBorderStyle(fmt.borders.bottom)};`;
                    if (fmt.borders.left) style += `border-left: ${mapBorderStyle(fmt.borders.left)};`;
                    if (fmt.borders.right) style += `border-right: ${mapBorderStyle(fmt.borders.right)};`;
                } else {
                    style += `border: 1px solid #e0e0e0;`; 
                }

                // Wrap Strategy
                if (fmt.wrapStrategy) {
                    switch (fmt.wrapStrategy) {
                        case 'OVERFLOW_CELL': style += 'white-space: nowrap;'; break;
                        case 'CLIP': style += 'white-space: nowrap; overflow: hidden;'; break;
                        case 'WRAP': style += 'white-space: normal; word-wrap: break-word;'; break;
                    }
                } else {
                     style += 'white-space: nowrap;'; 
                }
            } else {
                 // No effective format? Default fallback
                 style += `border: 1px solid #e0e0e0; white-space: nowrap;`;
            }
            
            // Attributes
            let attrs = `style="${style}"`;
            if (mergeStatus && typeof mergeStatus === 'object') {
                if (mergeStatus.rowspan > 1) attrs += ` rowspan="${mergeStatus.rowspan}"`;
                if (mergeStatus.colspan > 1) attrs += ` colspan="${mergeStatus.colspan}"`;
            }
            
            // Content processing
            content = content.replace(/\n/g, '<br>');
            
            // Hyperlinks
            if (cell.hyperlink) {
                content = `<a href="${cell.hyperlink}" target="_blank" style="color: inherit; text-decoration: underline;">${content}</a>`;
            }
            
            html += `<${tagName} ${attrs}>${content}</${tagName}>`;
        });
    }
    
    html += '</tr>';
  });

  html += '</tbody></table>';
  return html;
}

// --- RECOMMENDED MODEL (Auth + Full API) ---
async function runRecommendedConversion(spreadsheetId, targetGid = null, preFetchedData = null) {
  // UI Loading State
  convertBtn.innerHTML = '<span class="loading"></span>';
  convertBtn.disabled = true;
  outputPreview.classList.remove("active");
  outputPreview.innerHTML = `
      <div class="preview-placeholder">
          <div class="preview-placeholder-icon">
             <span class="loading" style="width: 48px; height: 48px;"></span>
          </div>
          <p>Fetching spreadsheet data (Recommended)...</p>
      </div>
  `;

  try {
    let result;
    if (preFetchedData) {
        console.log("Using background fetched data (Recommended)");
        result = preFetchedData;
    } else {
        const response = await gapi.client.sheets.spreadsheets.get({
            spreadsheetId: spreadsheetId,
            includeGridData: true,
            ranges: [], 
        });
        result = response.result;
    }
    
    if (!result.sheets || result.sheets.length === 0) {
        throw new Error("No sheets found in this spreadsheet.");
    }

    // Store raw data for JSON export
    state.lastRawData = {
        mode: 'recommended',
        data: result.sheets
    };

    // Build Tabbed Interface
    let tabsHtml = '<div class="modal-tabs">';
    let contentHtml = '<div class="sheet-views-container">';
    
    let activeIndex = 0;
    
    // Find target index if gid provided
    if (targetGid) {
        const found = result.sheets.findIndex(s => s.properties.sheetId == targetGid);
        if (found !== -1) activeIndex = found;
    }

    result.sheets.forEach((sheet, index) => {
        const sheetTitle = sheet.properties.title;
        const isActive = index === activeIndex;
        
        // Tab Button
        // We use onclick attribute for simplicity in injected HTML, calling a global function 
        tabsHtml += `
            <button class="sheet-tab ${isActive ? 'active' : ''}" 
                    onclick="switchTab(this, 'sheet-view-${index}')">
                ${sheetTitle}
            </button>
        `;

        // Sheet Content
        contentHtml += `
            <div id="sheet-view-${index}" class="sheet-view ${isActive ? 'active' : ''}">
                 <h3 style="margin-top: 0; border-bottom: 2px solid var(--accent-mint); display: inline-block; padding-bottom: 5px;">${sheetTitle}</h3>
                 ${renderSheetToHtml(sheet)}
            </div>
        `;
    });
    
    tabsHtml += '</div>'; // Close tabs container
    contentHtml += '</div>'; // Close views container

    // Combine: Tabs + Content (Header handled by UI)
    const finalHtml = tabsHtml + contentHtml;
    
    // Display Result in Modal AND Persistent Preview
    updatePreviewState(finalHtml);
    openModal(finalHtml);
    
    // Reset Button
    convertBtn.innerHTML = `
        <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="16 18 22 12 16 6"/>
            <polyline points="8 6 2 12 8 18"/>
        </svg>
        <span>Convert to HTML</span>
    `;
    convertBtn.disabled = false;

  } catch (err) {
    console.error("API Error:", err);
    convertBtn.innerHTML = '<span>Convert to HTML</span>';
    convertBtn.disabled = false;
    outputPreview.innerHTML = `
        <div style="text-align: center; color: #ff6b6b; padding: 2rem;">
            <p style="font-weight: 600;">Error fetching data</p>
            <p style="font-size: 0.9rem;">${err.result?.error?.message || err.message || JSON.stringify(err)}</p>
            <p style="font-size: 0.8rem; margin-top: 1rem; color: var(--text-secondary);">If this is a private sheet, ensure you are signed in.</p>
        </div>
    `;
  }
}

// --- FLASH MODEL (Smart Mode: API Key -> Legacy GViz) ---
async function runFlashConversion(spreadsheetId, targetGid = null, preFetchedData = null) {
    // UI Loading State
    convertBtn.innerHTML = '<span class="loading"></span>';
    convertBtn.disabled = true;
    outputPreview.classList.remove("active");
    outputPreview.innerHTML = `
        <div class="preview-placeholder">
            <div class="preview-placeholder-icon">
               <span class="loading" style="width: 48px; height: 48px;"></span>
            </div>
            <p>Fetching spreadsheet data (Flash)...</p>
        </div>
    `;

    // Check pre-fetched data
    if (preFetchedData) {
         if (preFetchedData.mode === 'recommended') {
             console.log("Using background fetched data (Flash - High Fidelity)");
             const sheets = preFetchedData.data;
             
             // Build Tabbed Interface
            let tabsHtml = '<div class="modal-tabs">';
            let contentHtml = '<div class="sheet-views-container">';
            
            let activeIndex = 0;
            if (targetGid) {
                const found = sheets.findIndex(s => s.properties.sheetId == targetGid);
                if (found !== -1) activeIndex = found;
            }

            sheets.forEach((sheet, index) => {
                const sheetTitle = sheet.properties.title;
                const isActive = index === activeIndex;
                
                tabsHtml += `
                    <button class="sheet-tab ${isActive ? 'active' : ''}" 
                            onclick="switchTab(this, 'sheet-view-${index}')">
                        ${sheetTitle}
                    </button>
                `;

                contentHtml += `
                    <div id="sheet-view-${index}" class="sheet-view ${isActive ? 'active' : ''}">
                         <h3 style="margin-top: 0; border-bottom: 2px solid var(--accent-mint); display: inline-block; padding-bottom: 5px;">${sheetTitle}</h3>
                         ${renderSheetToHtml(sheet)}
                    </div>
                `;
            });
            
            tabsHtml += '</div>';
            contentHtml += '</div>';

            const finalHtml = tabsHtml + contentHtml;
            updatePreviewState(finalHtml);
            openModal(finalHtml);
            
            convertBtn.innerHTML = `
                <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <polyline points="16 18 22 12 16 6"/>
                    <polyline points="8 6 2 12 8 18"/>
                </svg>
                <span>Convert to HTML</span>
            `;
            convertBtn.disabled = false;
            return;

         } else if (preFetchedData.mode === 'flash') {
             // GViz data
             console.log("Using background fetched data (Flash - GViz)");
             const data = preFetchedData.data;
             
            const numCols = data.getNumberOfColumns();
            const numRows = data.getNumberOfRows();
            
            let html = '<table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">';
            html += '<tbody>';
            
            // Headers
            let hasHeaders = false;
            for (let c = 0; c < numCols; c++) {
                if(data.getColumnLabel(c)) { hasHeaders = true; break; }
            }

            if (hasHeaders) {
                html += '<tr>';
                for (let c = 0; c < numCols; c++) {
                     let style = 'border: 1px solid #ccc; padding: 8px; background: #f3f3f3; font-weight: bold; text-align: left;';
                     const type = data.getColumnType(c); 
                     if (type === 'number') style += 'text-align: right;';
                     html += `<th style="${style}">${data.getColumnLabel(c) || ''}</th>`;
                }
                html += '</tr>';
            }

            // Data Rows
            for (let r = 0; r < numRows; r++) {
                html += '<tr>';
                for (let c = 0; c < numCols; c++) {
                    const type = data.getColumnType(c);
                    let style = 'border: 1px solid #e0e0e0; padding: 6px;';
                    if (type === 'number') style += 'text-align: right;';
                    if (type === 'boolean') style += 'text-align: center;';
                    
                    const content = data.getFormattedValue(r, c) || ''; 
                    html += `<td style="${style}">${content}</td>`;
                }
                html += '</tr>';
            }
            html += '</tbody></table>';

            const finalHtml = `
                <div class="sheet-container">
                    <h3 style="font-family: inherit; margin-bottom: 1rem; padding-bottom: 0.5rem; border-bottom: 2px solid var(--accent-mint);">Flash Output</h3>
                     <div style="font-size: 0.8rem; color: #666; margin-bottom: 1rem; display: flex; align-items: center; gap: 0.5rem;">
                        <span style="background: #fff3e0; color: #e65100; padding: 2px 6px; border-radius: 4px; font-weight: bold; font-size: 0.7rem;">FLASH MODEL</span>
                        <span>Optimized for speed. For full design fidelity, use Recommended model.</span>
                     </div>
                    ${html}
                </div>
            `;
            // Wait, previous variable name conflict in template string? No, just typo in my thought.
            // Actually, I should use `html` inside the string.
             
             // Correcting logic...
         }
    }

    // Attempt 1: High Fidelity via API Key
    try {
        console.log("Attempting Flash conversion via API Key...");
        
        // Note: gapi.client.init was called with apiKey, so this should work for public sheets
        const response = await gapi.client.sheets.spreadsheets.get({
            spreadsheetId: spreadsheetId,
            includeGridData: true,
            ranges: [], 
        });

        const result = response.result;
        
        if (!result.sheets || result.sheets.length === 0) {
            throw new Error("No sheets found in this spreadsheet.");
        }

        // Store raw data for JSON export
        state.lastRawData = {
            mode: 'recommended',
            data: result.sheets
        };

        // Build Tabbed Interface (Shared Logic with Recommended)
        let tabsHtml = '<div class="modal-tabs">';
        let contentHtml = '<div class="sheet-views-container">';
        
        let activeIndex = 0;
        
        // Find target index if gid provided
        if (targetGid) {
            const found = result.sheets.findIndex(s => s.properties.sheetId == targetGid);
            if (found !== -1) activeIndex = found;
        }

        result.sheets.forEach((sheet, index) => {
            const sheetTitle = sheet.properties.title;
            const isActive = index === activeIndex;
            
            tabsHtml += `
                <button class="sheet-tab ${isActive ? 'active' : ''}" 
                        onclick="switchTab(this, 'sheet-view-${index}')">
                    ${sheetTitle}
                </button>
            `;

            contentHtml += `
                <div id="sheet-view-${index}" class="sheet-view ${isActive ? 'active' : ''}">
                     <h3 style="margin-top: 0; border-bottom: 2px solid var(--accent-mint); display: inline-block; padding-bottom: 5px;">${sheetTitle}</h3>
                     ${renderSheetToHtml(sheet)}
                </div>
            `;
        });
        
        tabsHtml += '</div>';
        contentHtml += '</div>';

        const finalHtml = tabsHtml + contentHtml;
        
        // Display Result
        updatePreviewState(finalHtml);
        openModal(finalHtml);
        
        // Reset Button
        convertBtn.innerHTML = `
            <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <polyline points="16 18 22 12 16 6"/>
                <polyline points="8 6 2 12 8 18"/>
            </svg>
            <span>Convert to HTML</span>
        `;
        convertBtn.disabled = false;

    } catch (err) {
        console.warn("Flash API Key method failed. usage of fallback (GViz). Error:", err);
        // Fallback to Legacy GViz
        runLegacyFlashConversion(spreadsheetId, targetGid);
    }
}

// Legacy Flash Conversion (GViz Fallback)
function runLegacyFlashConversion(spreadsheetId, gid) {
    console.log("Running Legacy Flash Conversion (GViz)...");
    
    // UI Loading Update (if needed, though previous one persists)
    
    // Initialize the Google Visualization Query
    google.charts.load('current', {packages: ['table']});
    google.charts.setOnLoadCallback(() => {
        // Append GID if provided
        const queryUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?headers=1${gid ? '&gid=' + gid : ''}`;
        const query = new google.visualization.Query(queryUrl);
        query.send((response) => handleFlashResponse(response, spreadsheetId));
    });
}

function handleFlashResponse(response, spreadsheetId) {
    if (response.isError()) {
        console.error("Flash Error:", response.getMessage());
        convertBtn.innerHTML = '<span>Convert to HTML</span>';
        convertBtn.disabled = false;
        alert(`Flash Conversion Failed: ${response.getMessage()}\n\nMake sure the sheet is Public.`);
        return;
    }

    const data = response.getDataTable();
    
    // Store raw data for JSON export
    state.lastRawData = {
        mode: 'flash',
        data: data
    };
    
    const numCols = data.getNumberOfColumns();
    const numRows = data.getNumberOfRows();
    
    let html = '<table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">';
    html += '<tbody>';
    
    // Headers (Use Column Labels)
    let hasHeaders = false;
    for (let c = 0; c < numCols; c++) {
        if(data.getColumnLabel(c)) { hasHeaders = true; break; }
    }

    if (hasHeaders) {
        html += '<tr>';
        for (let c = 0; c < numCols; c++) {
             // In Flash model, we assume generic styling since we don't have formatting
             let style = 'border: 1px solid #ccc; padding: 8px; background: #f3f3f3; font-weight: bold; text-align: left;';
             const type = data.getColumnType(c); // 'string', 'number', 'boolean', 'date', 'datetime', 'timeofday'
             if (type === 'number') style += 'text-align: right;';
             
             html += `<th style="${style}">${data.getColumnLabel(c) || ''}</th>`;
        }
        html += '</tr>';
    }

    // Data Rows
    for (let r = 0; r < numRows; r++) {
        html += '<tr>';
        for (let c = 0; c < numCols; c++) {
            const type = data.getColumnType(c);
            let style = 'border: 1px solid #e0e0e0; padding: 6px;';
            if (type === 'number') style += 'text-align: right;';
            if (type === 'boolean') style += 'text-align: center;';
            
            // getFormattedValue returns the string representation formatted by the sheet
            const content = data.getFormattedValue(r, c) || ''; 
            html += `<td style="${style}">${content}</td>`;
        }
        html += '</tr>';
    }
    
    html += '</tbody></table>';

    const finalHtml = `
        <div class="sheet-container">
            <h3 style="font-family: inherit; margin-bottom: 1rem; padding-bottom: 0.5rem; border-bottom: 2px solid var(--accent-mint);">Flash Output</h3>
             <div style="font-size: 0.8rem; color: #666; margin-bottom: 1rem; display: flex; align-items: center; gap: 0.5rem;">
                <span style="background: #fff3e0; color: #e65100; padding: 2px 6px; border-radius: 4px; font-weight: bold; font-size: 0.7rem;">FLASH MODEL</span>
                <span>Optimized for speed. For full design fidelity, use Recommended model.</span>
             </div>
            ${html}
        </div>
    `;

    // Display Result in Modal AND Persistent Preview
    updatePreviewState(finalHtml);
    openModal(finalHtml);
    
    // Reset Button
    convertBtn.innerHTML = `
        <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="16 18 22 12 16 6"/>
            <polyline points="8 6 2 12 8 18"/>
        </svg>
        <span>Convert to HTML</span>
    `;
    convertBtn.disabled = false;
}

// Global Tab Switcher
window.switchTab = function(tabParams, viewId) { 
    // Deactivate all tabs
    const parentContainer = tabParams.closest('.modal-tabs');
    parentContainer.querySelectorAll('.sheet-tab').forEach(t => t.classList.remove('active'));
    // Activate clicked tab
    tabParams.classList.add('active');
    
    // Hide all views in the modal
    const modalContent = tabParams.closest('.modal-content') || document;
    modalContent.querySelectorAll('.sheet-view').forEach(v => v.classList.remove('active'));
    
    // Show target view
    const view = document.getElementById(viewId);
    if(view) view.classList.add('active');
};
