// Helper functions and utilities

// Helper to extract Spreadsheet ID and GID from URL
function extractSheetInfo(url) {
  const matches = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  const gidMatch = url.match(/[#&]gid=([0-9]+)/);
  return {
      id: matches ? matches[1] : null,
      gid: gidMatch ? gidMatch[1] : null
  };
}

// Helper to convert API color fraction to RGB
function colorToRgb(color) {
  if (!color) return null;
  const r = Math.floor((color.red || 0) * 255);
  const g = Math.floor((color.green || 0) * 255);
  const b = Math.floor((color.blue || 0) * 255);
  // Optional: Handle alpha if present
  const a = color.alpha !== undefined ? color.alpha : 1;
  if (a < 1) return `rgba(${r}, ${g}, ${b}, ${a})`;
  return `rgb(${r}, ${g}, ${b})`;
}

// Helper to map Sheets API border style to CSS
function mapBorderStyle(border) {
    if (!border || !border.style || border.style === 'NONE') return 'none';
    const color = colorToRgb(border.color) || '#000';
    
    let width = '1px';
    let style = 'solid';

    switch (border.style) {
        case 'DOTTED': style = 'dotted'; break;
        case 'DASHED': style = 'dashed'; break;
        case 'SOLID_MEDIUM': width = '2px'; break;
        case 'SOLID_THICK': width = '3px'; break;
        case 'DOUBLE': style = 'double'; width = '3px'; break; 
        default: break; // SOLID
    }
    
    return `${width} ${style} ${color}`;
}

// Helper: Download string as file
function downloadFile(content, filename, type) {
    const blob = new Blob([content], { type: type });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// Helper: Normalize Rows to JSON Array (Headers = Keys)
function normalizeRowsToJson(rows, sourceType) {
    let jsonArray = [];
    
    // Heuristic Helper: Score a row's suitability as a header
    // Higher score = better header candidate
    const scoreHeaderRow = (rowValues) => {
        if (!rowValues || rowValues.length === 0) return -1;
        let validStrings = 0;
        let emptyOrDash = 0;
        const uniqueValues = new Set();
        
        rowValues.forEach(val => {
            const s = String(val).trim();
            if (!s || s === '-' || s === '0' || s === 'null') {
                emptyOrDash++;
            } else {
                validStrings++;
                uniqueValues.add(s);
            }
        });
        
        // Penalties/Bonuses
        if (uniqueValues.size !== validStrings) return -1; // Duplicate headers checks bad
        if (validStrings === 0) return -1;
        
        return validStrings - (emptyOrDash * 0.5); 
    };

    if (sourceType === 'recommended') {
        if (!rows || rows.length === 0) return [];
        
        // Find Header Row (Scan first 10 rows)
        let bestHeaderIndex = 0;
        let maxScore = -999;
        
        const limit = Math.min(rows.length, 10);
        for (let i = 0; i < limit; i++) {
            const values = rows[i]?.values?.map(c => c.formattedValue) || [];
            const score = scoreHeaderRow(values);
            if (score > maxScore) {
                maxScore = score;
                bestHeaderIndex = i;
            }
        }
        
        if (maxScore <= 0) bestHeaderIndex = 0; // Fallback to 0 if nothing looks good

        // Extract Headers
        const headers = rows[bestHeaderIndex]?.values?.map((cell, index) => {
            return cell.formattedValue || `Column_${index}`;
        }) || [];
        
        // Process Data Rows (Start AFTER header row)
        for (let i = bestHeaderIndex + 1; i < rows.length; i++) {
            const row = rows[i];
            let obj = {};
            if(row.values) {
                row.values.forEach((cell, colIndex) => {
                    // Use header if mapped, else Column_X
                    const key = headers[colIndex] || `Column_${colIndex}`;
                    // Skip if key is empty or just "Column_X" AND value is empty? No, keep it.
                    // Actually, if key is duplicate (handled by map??), uniqueness handled above.
                    
                    // Sanitize Key to ensure no duplicates in object (though unlikely with unique check)
                    obj[key] = cell.formattedValue || ""; // Change null/undefined to empty string for cleaner JSON
                });
            }
            jsonArray.push(obj);
        }
    } else if (sourceType === 'flash') {
        const numCols = rows.getNumberOfColumns();
        const numRows = rows.getNumberOfRows();
        
        // Headers (GViz usually handles headers internally, but if we need to scan raw data...)
        // GViz separates labels from data, so checking Row 0 is usually not needed.
        // But if the user passed range A1:Z, headers ARE Labels.
        
        let headers = [];
        for (let c = 0; c < numCols; c++) {
            headers.push(rows.getColumnLabel(c) || `Column_${c}`);
        }
        
        // Rows
        for (let r = 0; r < numRows; r++) {
            let obj = {};
            for (let c = 0; c < numCols; c++) {
                 const key = headers[c];
                 obj[key] = rows.getFormattedValue(r, c) || "";
            }
            jsonArray.push(obj);
        }
    }
    
    return jsonArray;
}

// Helper: Convert JSON (Multi-sheet or array) to CSV
function jsonToCsv(data) {
    let allRows = [];
    let headers = new Set();

    // Flatten multi-sheet structure if necessary
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
