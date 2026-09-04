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
            max-height: 90vh;
            overflow-y: auto;
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
        .gi-help { font-size: 12px; color: #888; line-height: 1.35; margin-top: 4px; }
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
        .gi-btn-secondary {
            background: #6c757d;
            margin-top: 10px;
        }
        .gi-btn-secondary:hover { background: #5a6268; }
        #gi-status {
            margin-top: 15px;
            font-size: 13px;
            color: #666;
            line-height: 1.4;
            background: #f9fafb;
            padding: 10px;
            border-radius: 4px;
        }
        #gi-log-container {
            display: none;
            margin-top: 15px;
            background: #fff;
            border: 1px solid #eee;
            border-radius: 4px;
            padding: 10px;
            font-size: 12px;
            max-height: 200px;
            overflow-y: auto;
        }
        .gi-log-item {
            padding: 4px 0;
            border-bottom: 1px solid #f0f0f0;
        }
        .gi-log-item:last-child { border-bottom: none; }
        .gi-log-warn { color: #d97706; }
        .gi-log-info { color: #2563eb; }
        #gi-disclaimer {
            margin-bottom: 15px;
            background: #fef2f2;
            border: 1px solid #fecaca;
            border-left: 4px solid #dc2626;
            border-radius: 4px;
            padding: 8px 10px;
            font-size: 11px;
            line-height: 1.4;
            color: #7f1d1d;
        }
        #gi-disclaimer strong { color: #991b1b; }
        #gi-update {
            margin-bottom: 15px;
            background: #fffbeb;
            border: 1px solid #fde68a;
            border-left: 4px solid #d97706;
            border-radius: 4px;
            padding: 8px 10px;
            font-size: 11px;
            line-height: 1.4;
            color: #78350f;
        }
        #gi-update a { color: #92400e; font-weight: bold; }
    `;
    document.head.appendChild(style);

    // Create UI Elements
    const overlay = document.createElement('div');
    overlay.id = 'gi-overlay';

    overlay.innerHTML = `
        <div id="gi-header">
            <span>Punten importeren</span>
            <span id="gi-close">&times;</span>
        </div>

        <div id="gi-update" style="display:none;"></div>

        <div id="gi-disclaimer">
            <strong>⚠ Geen officiële AP-tool.</strong> Niet ontwikkeld door AP Hogeschool, maar door Tim Dams.
            De lector blijft zelf eindverantwoordelijk voor het correct invoeren en controleren van de cijfers —
            controleer altijd de punten vóór je opslaat.
        </div>

        <div class="gi-row">
            <input type="file" id="gi-file-input" accept=".xlsx, .xls" style="width: 100%" />
        </div>

        <div class="gi-row">
            <label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-weight:normal;">
                <input type="checkbox" id="gi-header-check" checked> 
                <span>Eerste rij bevat kolomtitels</span>
            </label>
            <div class="gi-help" style="margin-left:26px;">De eerste rij van je Excel zijn titels zoals "Naam" of "Punt", geen echte student. Vink dit uit als rij 1 al meteen een student is.</div>
        </div>

        <div class="gi-row">
            <label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-weight:normal;">
                <input type="checkbox" id="gi-empty-absent-check">
                <span>Lege cel in Excel = Afwezig</span>
            </label>
            <div class="gi-help" style="margin-left:26px;">Studenten zonder ingevuld cijfer worden op "Afwezig" gezet. (Studenten die niet in je Excel staan, blijven altijd leeg.)</div>
        </div>

        <div class="gi-row">
            <label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-weight:normal;">
                <input type="checkbox" id="gi-allow-decimal-check">
                <span>Kommapunten toestaan</span>
            </label>
            <div class="gi-help" style="margin-left:26px;">Sta cijfers met decimalen toe (bv. 12,5). Staat dit uit, dan worden decimale cijfers <strong>niet</strong> ingevuld en in de log gemeld. Een punt (12.5) wordt automatisch omgezet naar een komma (12,5).</div>
        </div>

        <div class="gi-row">
            <label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-weight:normal;">
                <input type="checkbox" id="gi-split-name-check">
                <span>Naam staat in 2 aparte kolommen</span>
            </label>
            <div class="gi-help" style="margin-left:26px;">Aanvinken als voornaam en achternaam in losse kolommen staan (bv. kolom A = achternaam, kolom B = voornaam).</div>
        </div>

        <div id="gi-mapping" style="display:none;">
            <div class="gi-row" id="gi-name-row">
                <label class="gi-label">Kolom met de naam</label>
                <select id="gi-col-name" class="gi-select"></select>
            </div>
            <div class="gi-row" id="gi-lastname-row" style="display:none;">
                <label class="gi-label">Kolom met de achternaam</label>
                <select id="gi-col-lastname" class="gi-select"></select>
            </div>
            <div class="gi-row" id="gi-firstname-row" style="display:none;">
                <label class="gi-label">Kolom met de voornaam</label>
                <select id="gi-col-firstname" class="gi-select"></select>
            </div>
            <div class="gi-row">
                <label class="gi-label">Kolom met het cijfer</label>
                <select id="gi-col-grade" class="gi-select"></select>
            </div>
        </div>

        <button id="gi-import-btn" class="gi-btn" disabled>Punten importeren</button>
        <button id="gi-log-btn" class="gi-btn gi-btn-secondary" style="display:none;">Toon log</button>
        
        <div id="gi-status">Kies een Excel-bestand om te starten.</div>
        <div id="gi-log-container"></div>
    `;

    document.body.appendChild(overlay);

    // References
    const closeBtn = overlay.querySelector('#gi-close');
    const fileInput = overlay.querySelector('#gi-file-input');
    const headerCheck = overlay.querySelector('#gi-header-check');
    const emptyAbsentCheck = overlay.querySelector('#gi-empty-absent-check');
    const allowDecimalCheck = overlay.querySelector('#gi-allow-decimal-check');
    const splitNameCheck = overlay.querySelector('#gi-split-name-check');
    const mappingDiv = overlay.querySelector('#gi-mapping');
    const nameRow = overlay.querySelector('#gi-name-row');
    const lastnameRow = overlay.querySelector('#gi-lastname-row');
    const firstnameRow = overlay.querySelector('#gi-firstname-row');
    const nameSelect = overlay.querySelector('#gi-col-name');
    const lastnameSelect = overlay.querySelector('#gi-col-lastname');
    const firstnameSelect = overlay.querySelector('#gi-col-firstname');
    const gradeSelect = overlay.querySelector('#gi-col-grade');
    const importBtn = overlay.querySelector('#gi-import-btn');
    const logBtn = overlay.querySelector('#gi-log-btn');
    const statusDiv = overlay.querySelector('#gi-status');
    const logContainer = overlay.querySelector('#gi-log-container');

    closeBtn.onclick = () => overlay.remove();

    // Versiecheck: waarschuwt wanneer er op GitHub een nieuwere importer.js staat.
    // GI_VERSION wordt bij het deployen vervangen door de commit-hash van importer.js
    // (zie .github/workflows/deploy.yml). Lokaal blijft de placeholder staan; dan
    // slaan we de check over.
    const GI_VERSION = '__GI_VERSION__';
    const GI_BASE_URL = 'https://timdams.github.io/ibamaflexPuntenImporter/';
    const updateDiv = overlay.querySelector('#gi-update');

    function checkForUpdate() {
        if (GI_VERSION.indexOf('__') === 0) return;
        fetch(GI_BASE_URL + 'version.json?t=' + Date.now(), { cache: 'no-store' })
            .then(r => r.ok ? r.json() : null)
            .then(info => {
                if (!info || !info.version || info.version === 'dev') return;
                if (info.version === GI_VERSION) return;
                const datum = info.date ? ' (' + info.date + ')' : '';
                updateDiv.innerHTML = '<strong>⚠ Nieuwe versie beschikbaar' + datum + '.</strong> ' +
                    'Jouw bladwijzer bevat een oudere versie van de tool. ' +
                    '<a href="' + GI_BASE_URL + 'install.html" target="_blank" rel="noopener">Sleep de knop opnieuw</a> ' +
                    'vanaf de installatiepagina om te updaten. Je kan gerust eerst deze import afwerken.';
                updateDiv.style.display = 'block';
            })
            .catch(() => { /* offline of geblokkeerd: gewoon geen melding tonen */ });
    }

    checkForUpdate();

    let workbook = null;
    let jsonData = null;
    let currentFile = null;
    let importLog = [];

    fileInput.addEventListener('change', (e) => {
        currentFile = e.target.files[0];
        processFile();
    });

    headerCheck.addEventListener('change', () => {
        if (currentFile) processFile();
    });

    splitNameCheck.addEventListener('change', () => {
        const split = splitNameCheck.checked;
        nameRow.style.display = split ? 'none' : '';
        lastnameRow.style.display = split ? '' : 'none';
        firstnameRow.style.display = split ? '' : 'none';
    });

    logBtn.addEventListener('click', () => {
        if (logContainer.style.display === 'none') {
            logContainer.style.display = 'block';
            logBtn.textContent = 'Verberg log';
            renderLog();
        } else {
            logContainer.style.display = 'none';
            logBtn.textContent = 'Toon log';
        }
    });

    function renderLog() {
        if (importLog.length === 0) {
            logContainer.innerHTML = '<div class="gi-log-item">Niets te melden.</div>';
            return;
        }
        logContainer.innerHTML = importLog.map(item => {
            const cls = item.type === 'warn' ? 'gi-log-warn' : 'gi-log-info';
            return `<div class="gi-log-item ${cls}">${item.msg}</div>`;
        }).join('');
    }

    function processFile() {
        if (!currentFile) return;

        mappingDiv.style.display = 'none';
        importBtn.disabled = true;
        logBtn.style.display = 'none';
        logContainer.style.display = 'none';
        statusDiv.textContent = 'Bestand inlezen...';

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            try {
                workbook = XLSX.read(data, { type: 'array' });

                // Parse first sheet
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];

                // Get header option
                const useHeaders = headerCheck.checked;
                const opts = useHeaders ? {} : { header: "A" };

                jsonData = XLSX.utils.sheet_to_json(sheet, opts);

                if (jsonData.length === 0) {
                    statusDiv.textContent = 'Het bestand is leeg.';
                    return;
                }

                // Get Headers
                const headers = Object.keys(jsonData[0]);

                // Populate Selects
                populateSelect(nameSelect, headers, ['naam', 'student', 'name']);
                // For split mode: default to first column = lastname, second = firstname
                populateSelect(lastnameSelect, headers, ['achternaam', 'familienaam', 'lastname', 'last name', 'surname']);
                populateSelect(firstnameSelect, headers, ['voornaam', 'firstname', 'first name', 'given name', 'prenom']);
                if (headers.length >= 2) {
                    if (!hasHeuristicMatch(headers, ['achternaam', 'familienaam', 'lastname', 'last name', 'surname'])) {
                        lastnameSelect.selectedIndex = 0;
                    }
                    if (!hasHeuristicMatch(headers, ['voornaam', 'firstname', 'first name', 'given name', 'prenom'])) {
                        firstnameSelect.selectedIndex = 1;
                    }
                }
                populateSelect(gradeSelect, headers, ['punt', 'score', 'cijfer', 'grade', 'result']);

                mappingDiv.style.display = 'block';
                importBtn.disabled = false;
                statusDiv.textContent = 'Bestand geladen. Controleer hieronder de kolommen.';

            } catch (err) {
                console.error(err);
                statusDiv.textContent = 'Kon het Excel-bestand niet lezen. Zorg dat het een geldig .xlsx-bestand is.';
            }
        };
        reader.readAsArrayBuffer(currentFile);
    }

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

    function hasHeuristicMatch(options, heuristics) {
        return options.some(opt => heuristics.some(h => String(opt).toLowerCase().includes(h)));
    }

    // iBaMaFlex plakt achter de naam soms extra codes: {8} <J> [B].
    // Die horen niet bij de naam en worden weggeknipt voor het vergelijken.
    function stripNameMarkers(s) {
        return String(s || '')
            .replace(/[{[<][^{}[\]<>]*[}\]>]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function normalizeName(s) {
        return stripNameMarkers(s)
            .toLowerCase()
            .replace(/\s+/g, ' ')
            .trim();
    }

    // De bijnaam staat tussen haakjes: "Intzidis Alki (Alkiviadis)".
    // We proberen zowel mét als zonder die bijnaam te matchen.
    function nameVariants(s) {
        const base = stripNameMarkers(s);
        const variants = new Set([
            normalizeName(base),
            normalizeName(base.replace(/\([^)]*\)/g, ' ')),
            normalizeName(base.replace(/[()]/g, ' ')),
        ]);
        return [...variants].filter(Boolean);
    }

    function tokenize(s) {
        return normalizeName(s).split(' ').filter(Boolean).sort().join(' ');
    }

    importBtn.addEventListener('click', () => {
        if (!jsonData) return;

        const scoreKey = gradeSelect.value;
        const emptyAsAbsent = emptyAbsentCheck.checked;
        const allowDecimal = allowDecimalCheck.checked;
        const splitName = splitNameCheck.checked;

        const nameConfig = splitName
            ? { split: true, lastKey: lastnameSelect.value, firstKey: firstnameSelect.value }
            : { split: false, nameKey: nameSelect.value };

        processImport(nameConfig, scoreKey, emptyAsAbsent, allowDecimal);
    });

    function getExcelName(record, nameConfig) {
        if (nameConfig.split) {
            const last = String(record[nameConfig.lastKey] || '').trim();
            const first = String(record[nameConfig.firstKey] || '').trim();
            return (last + ' ' + first).trim();
        }
        return String(record[nameConfig.nameKey] || '').trim();
    }

    function matchesDomName(domVariants, record, nameConfig) {
        const domSet = new Set(domVariants);
        const domTokens = new Set(domVariants.map(tokenize));
        const hits = (s) => nameVariants(s).some(v => domSet.has(v));

        if (nameConfig.split) {
            const last = normalizeName(record[nameConfig.lastKey]);
            const first = normalizeName(record[nameConfig.firstKey]);
            if (!last && !first) return false;
            const candidates = [
                `${last} ${first}`.trim(),
                `${first} ${last}`.trim(),
                `${last}, ${first}`.trim(),
            ];
            if (candidates.some(hits)) return true;
            // Fall back to token-set match (order-independent)
            return domTokens.has(tokenize(`${last} ${first}`));
        }
        return hits(record[nameConfig.nameKey]);
    }

    function processImport(nameConfig, scoreKey, emptyAsAbsent, allowDecimal) {
        statusDiv.textContent = 'Bezig met verwerken...';
        importLog = []; // Reset log
        logBtn.style.display = 'none';
        logContainer.style.display = 'none';

        let matches = 0;
        let notFound = 0;
        let skippedDecimal = 0;

        // Find table rows
        const rows = document.querySelectorAll('#ctl00_ctl00_cphGeneral_cphMain_rgPuntenP_ctl00 tbody tr');

        // Capture data from the DOM first for efficient matching logic? 
        // Or just iterate rows as is. Iterate rows is fine.

        // We also want to know which Excel students were NOT found in the DOM (optional, but good for log)
        // Let's create a Set of found names from Excel to track reverse.
        // Actually, the request said "Student X not found in Excel". That means iterating DOM rows and checking Excel.
        // Wait, "deze studenten niet in de excel gevonden" -> "This student (from DOM?) not found in Excel"?
        // Usually import reports "Row X from Excel not found in DOM".
        // BUT phrasing "deze studenten niet in de excel gevonden" implies we look at the class list (DOM) and see who is missing in the file.
        // I will log both if possible, or stick to the requested direction.
        // Let's assume the user means: "I have a list of students in the app, and I want to know who was NOT in the Excel file".

        const excelNamesSet = new Set(); // To track which excel records were used

        rows.forEach(row => {
            const nameCell = row.cells[1];
            if (!nameCell) return;

            const domName = nameCell.textContent.trim();
            const domVariants = nameVariants(domName);
            const originalDomName = domVariants[0] || normalizeName(domName); // For display

            const record = jsonData.find(d => matchesDomName(domVariants, d, nameConfig));

            const scoreInput = row.querySelector('input[name$="txtScore"]');

            // Reset styles
            scoreInput.style.backgroundColor = '';
            nameCell.style.backgroundColor = '';

            if (record && scoreInput) {
                // Found match
                const excelNameKey = normalizeName(getExcelName(record, nameConfig));
                excelNamesSet.add(excelNameKey);

                // FIXED: Treat undefined as empty string. 
                const scoreValue = record[scoreKey];
                const rawScore = (scoreValue === undefined || scoreValue === null) ? '' : String(scoreValue).trim();

                let effectiveScore = rawScore;

                if (rawScore === '') {
                    if (emptyAsAbsent) {
                        effectiveScore = 'A';
                    } else {
                        // Regular behavior: skip empty grades
                        // Log that we skipped?
                        // importLog.push({ type: 'info', msg: `Skipped ${domName} (Empty grade)` });
                        return;
                    }
                }

                // Shared function to force dirty state
                const forceDirtyState = () => {
                    try {
                        var hiddenDirty = document.getElementById('ctl00_ctl00_cphGeneral_cphMain_txtGewijzigdPagina');
                        var grid = window.$find ? window.$find('ctl00_ctl00_cphGeneral_cphMain_rgPuntenP') : null;

                        if (hiddenDirty && grid) {
                            var parts = row.id.split('__');
                            if (parts.length > 1) {
                                var idx = parseInt(parts[parts.length - 1], 10);
                                var item = grid.get_masterTableView().get_dataItems()[idx];
                                if (item) {
                                    var id = item.getDataKeyValue("p_examen");
                                    if (id && hiddenDirty.value.indexOf(id + ";") === -1) {
                                        hiddenDirty.value += id + ";";
                                    }
                                }
                            }
                        }
                    } catch (e) {
                        console.error("Error forcing dirty state:", e);
                    }
                };

                if (effectiveScore.toUpperCase() === 'A') {
                    // Handle Absent
                    if (scoreInput.onclick) scoreInput.onclick();
                    scoreInput.focus();

                    const keyEvents = [
                        new KeyboardEvent('keydown', { key: '+', code: 'NumpadAdd', keyCode: 107, which: 107, bubbles: true }),
                        new KeyboardEvent('keypress', { key: '+', charCode: 43, keyCode: 43, which: 43, bubbles: true }),
                        new KeyboardEvent('keyup', { key: '+', code: 'NumpadAdd', keyCode: 107, which: 107, bubbles: true }),
                        new InputEvent('input', { data: '+', inputType: 'insertText', bubbles: true })
                    ];

                    keyEvents.forEach(evt => scoreInput.dispatchEvent(evt));

                    scoreInput.style.backgroundColor = '#cfe2ff'; // Blue-ish
                    forceDirtyState();
                    matches++;
                    importLog.push({ type: 'info', msg: `Op AFWEZIG gezet: ${domName}` });

                } else {
                    // Standard score — detecteer kommapunten (decimale cijfers)
                    const numeric = parseFloat(effectiveScore.replace(',', '.'));
                    const isDecimal = !isNaN(numeric) && !Number.isInteger(numeric);

                    if (isDecimal && !allowDecimal) {
                        // Optie staat uit: niet invullen, wel melden
                        skippedDecimal++;
                        importLog.push({ type: 'warn', msg: `Kommapunt overgeslagen: ${domName} (${rawScore}). Vink "Kommapunten toestaan" aan om dit cijfer wel over te nemen.` });
                        return;
                    }

                    // Punt → komma normaliseren (iBaMaFlex verwacht een komma)
                    const valueToSet = isDecimal ? effectiveScore.replace('.', ',') : effectiveScore;

                    if (scoreInput.onclick) scoreInput.onclick();
                    scoreInput.focus();

                    scoreInput.value = valueToSet;

                    if (scoreInput.onkeydown) scoreInput.onkeydown({ keyCode: 13 });
                    if (scoreInput.onchange) scoreInput.onchange();
                    if (scoreInput.onblur) scoreInput.onblur();

                    forceDirtyState();
                    matches++;
                    // Optional: Log success? Too spammy for large classes.
                }

            } else {
                notFound++;
                // Log that this student was not found in Excel
                importLog.push({ type: 'warn', msg: `Niet gevonden in Excel: ${domName}` });
            }
        });

        let summaryHtml = `<strong>Klaar!</strong><br>Ingevuld: ${matches}<br>Niet gevonden: ${notFound}`;
        if (skippedDecimal > 0) {
            summaryHtml += `<br>Kommapunten overgeslagen: ${skippedDecimal}`;
        }
        statusDiv.innerHTML = summaryHtml;

        // Show Log Button
        logBtn.style.display = 'block';
        if (importLog.length > 0) {
            logBtn.textContent = `Toon log (${importLog.length})`;
        } else {
            logBtn.textContent = "Toon log";
        }
    }

})();
