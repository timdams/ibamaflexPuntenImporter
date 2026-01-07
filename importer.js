/**
 * Grade Importer Main Logic (v2 with Column Selection)
 */
(function () {
    console.log('Initializing Grade Importer UI...');

    // Styles for the floating UI
    const style = document.createElement('style');
    style.textContent = `
        #gi-overlay {
            position: fixed;
            top: 20px;
            right: 20px;
            width: 320px;
            background: white;
            border: 1px solid #ccc;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            z-index: 10000;
            padding: 20px;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
            border-radius: 8px;
            font-size: 14px;
        }
        #gi-header {
            font-weight: bold;
            font-size: 16px;
            margin-bottom: 15px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid #eee;
            padding-bottom: 10px;
        }
        #gi-close {
            cursor: pointer;
            color: #999;
            font-size: 20px;
        }
        #gi-close:hover { color: #333; }
        .gi-row { margin-bottom: 10px; }
        .gi-label { display: block; margin-bottom: 4px; font-weight: 500; color: #555; }
        .gi-select { width: 100%; padding: 6px; border: 1px solid #ddd; border-radius: 4px; }
        .gi-btn {
            background: #2563eb;
            color: white;
            border: none;
            padding: 10px;
            border-radius: 4px;
            cursor: pointer;
            width: 100%;
            margin-top: 15px;
            font-weight: bold;
        }
        .gi-btn:hover { background: #1d4ed8; }
        .gi-btn:disabled { background: #ccc; cursor: not-allowed; }
        #gi-status {
            margin-top: 15px;
            font-size: 13px;
            color: #666;
            line-height: 1.4;
            background: #f9fafb;
            padding: 10px;
            border-radius: 4px;
        }
    `;
    document.head.appendChild(style);

    // Create UI Elements
    const overlay = document.createElement('div');
    overlay.id = 'gi-overlay';

    overlay.innerHTML = `
        <div id="gi-header">
            <span>Import Grades</span>
            <span id="gi-close">&times;</span>
        </div>
        
        <div class="gi-row">
            <input type="file" id="gi-file-input" accept=".xlsx, .xls" style="width: 100%" />
        </div>

        <div id="gi-mapping" style="display:none;">
            <div class="gi-row">
                <label class="gi-label">Student Name Column</label>
                <select id="gi-col-name" class="gi-select"></select>
            </div>
            <div class="gi-row">
                <label class="gi-label">Grade Column</label>
                <select id="gi-col-grade" class="gi-select"></select>
            </div>
        </div>

        <button id="gi-import-btn" class="gi-btn" disabled>Import Grades</button>
        <div id="gi-status">Select an Excel file to start.</div>
    `;

    document.body.appendChild(overlay);

    // References
    const closeBtn = overlay.querySelector('#gi-close');
    const fileInput = overlay.querySelector('#gi-file-input');
    const mappingDiv = overlay.querySelector('#gi-mapping');
    const nameSelect = overlay.querySelector('#gi-col-name');
    const gradeSelect = overlay.querySelector('#gi-col-grade');
    const importBtn = overlay.querySelector('#gi-import-btn');
    const statusDiv = overlay.querySelector('#gi-status');

    closeBtn.onclick = () => overlay.remove();

    let workbook = null;
    let jsonData = null;

    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        mappingDiv.style.display = 'none';
        importBtn.disabled = true;
        statusDiv.textContent = 'Reading file...';

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            try {
                workbook = XLSX.read(data, { type: 'array' });

                // Parse first sheet
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                // Get header row
                jsonData = XLSX.utils.sheet_to_json(sheet);

                if (jsonData.length === 0) {
                    statusDiv.textContent = 'File is empty.';
                    return;
                }

                // Get Headers
                const headers = Object.keys(jsonData[0]);

                // Populate Selects
                populateSelect(nameSelect, headers, ['naam', 'student', 'name']);
                populateSelect(gradeSelect, headers, ['punt', 'score', 'cijfer', 'grade', 'result']);

                mappingDiv.style.display = 'block';
                importBtn.disabled = false;
                statusDiv.textContent = 'File loaded. Please confirm columns.';

            } catch (err) {
                console.error(err);
                statusDiv.textContent = 'Error parsing Excel file. Ensure it is a valid .xlsx file.';
            }
        };
        reader.readAsArrayBuffer(file);
    });

    function populateSelect(select, options, heuristics) {
        select.innerHTML = '';
        let selectedIndex = 0;
        options.forEach((opt, index) => {
            const el = document.createElement('option');
            el.value = opt;
            el.textContent = opt;
            select.appendChild(el);

            if (heuristics.some(h => opt.toLowerCase().includes(h))) {
                selectedIndex = index;
            }
        });
        select.selectedIndex = selectedIndex;
    }

    importBtn.addEventListener('click', () => {
        if (!jsonData) return;

        const nameKey = nameSelect.value;
        const scoreKey = gradeSelect.value;

        processImport(nameKey, scoreKey);
    });

    function processImport(nameKey, scoreKey) {
        statusDiv.textContent = 'Processing...';

        let matches = 0;
        let notFound = 0;

        // Find table rows
        const rows = document.querySelectorAll('#ctl00_ctl00_cphGeneral_cphMain_rgPuntenP_ctl00 tbody tr');

        rows.forEach(row => {
            // Check headers index to be sure? Assuming standard iBaMaFlex layout
            // Usually Name is index 2 (if index 0 is checkbox/hidden) or index 1. 
            // Based on demo: <td>Code</td><td>Name</td>...
            const nameCell = row.cells[1];
            if (!nameCell) return;

            const domName = nameCell.textContent.trim().toLowerCase();

            // Clean DOM name: remove trailing content in brackets (e.g. "Name [B]")
            const cleanDomName = domName.replace(/\s*\[.*?\]$/, '').trim().toLowerCase();

            const record = jsonData.find(d => {
                const excelName = String(d[nameKey]).trim().toLowerCase();
                // Strict match on cleaned name
                return cleanDomName === excelName;
            });

            const scoreInput = row.querySelector('input[name$="txtScore"]');

            // Reset styles
            scoreInput.style.backgroundColor = '';
            nameCell.style.backgroundColor = '';

            if (record && scoreInput) {
                const scoreValue = record[scoreKey];

                if (scoreValue !== undefined) {
                    const score = String(scoreValue).trim();

                    if (score === '') {
                        // Skip empty grades
                        // keep yellow? or white? User said "skippen". 
                        // Let's leave it alone (white) to match "skip".
                        // But maybe yellow is useful to know match occurred? 
                        // "die student gewoon skippen" -> Do nothing.
                        return;
                    }

                    if (score.toUpperCase() === 'A') {
                        // Handle Absent: Simulate '+' keypress
                        scoreInput.focus();
                        scoreInput.dispatchEvent(new KeyboardEvent('keydown', { key: '+', code: 'NumpadAdd', keyCode: 107, which: 107, bubbles: true }));
                        scoreInput.dispatchEvent(new KeyboardEvent('keypress', { key: '+', keyCode: 43, which: 43, bubbles: true }));
                        scoreInput.style.backgroundColor = '#cfe2ff'; // Blue-ish
                        matches++;
                    } else {
                        // Simulate full interaction sequence including click
                        if (scoreInput.onclick) scoreInput.onclick();
                        scoreInput.focus();
                        if (scoreInput.onfocus) scoreInput.onfocus(); // Double tap just in case

                        scoreInput.value = score;

                        // Signal changes
                        if (scoreInput.onkeydown) scoreInput.onkeydown({ keyCode: 13 }); // Enter key simulation might help?
                        if (scoreInput.onchange) scoreInput.onchange();

                        scoreInput.blur();
                        if (scoreInput.onblur) scoreInput.onblur();

                        // --- SAVED STATE HACK ---
                        // Force the page to recognize this row as "dirty" by adding the ID to the hidden field.
                        try {
                            // Find hidden field and grid if we haven't already (declared outside loop in full version, but here inside for safety in this block patch)
                            var hiddenDirty = document.getElementById('ctl00_ctl00_cphGeneral_cphMain_txtGewijzigdPagina');
                            var grid = window.$find ? window.$find('ctl00_ctl00_cphGeneral_cphMain_rgPuntenP') : null;

                            if (hiddenDirty && grid) {
                                // Extract index from row ID: "ctl00_ctl00_cphGeneral_cphMain_rgPuntenP_ctl00__15" -> 15
                                var parts = row.id.split('__');
                                if (parts.length > 1) {
                                    var idx = parseInt(parts[parts.length - 1], 10);
                                    // Get data item from Telerik Grid
                                    var item = grid.get_masterTableView().get_dataItems()[idx];
                                    if (item) {
                                        var id = item.getDataKeyValue("p_examen");
                                        // Append ID if not present
                                        if (id && hiddenDirty.value.indexOf(id + ";") === -1) {
                                            hiddenDirty.value += id + ";";
                                            console.log("Forced dirty state for ID: " + id);
                                        }
                                    }
                                }
                            }
                        } catch (e) {
                            console.error("Error forcing dirty state:", e);
                        }
                        // ------------------------

                        // REMOVED manual styling so we can see if the native logic (red/white) works
                        // scoreInput.style.backgroundColor = '#d4edda'; 
                        matches++;
                    }
                } else {
                    // undefined score in excel (empty column)
                    // nameCell.style.backgroundColor = '#fff3cd'; 
                }
            } else {
                notFound++;
            }
        });

        statusDiv.innerHTML = `<strong>Done!</strong><br>Matched: ${matches} student(s)<br>Not found: ${notFound} student(s)`;
    }

})();
