/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this script to only the current document. This is a
 * security measure to assure users that the script isn't accessing
 * other files in their Google Drive.
 */

/**
 * Creates a custom menu in the spreadsheet UI when the document is opened.
 * This function is a special "simple trigger" that Google Apps Script
 * automatically runs when the spreadsheet is opened.
 */
function onOpen() {
    // Get the Google Sheets user interface object.
    const ui = SpreadsheetApp.getUi();

    // Create a new menu named 'Randomizer'.
    ui.createMenu('Randomizer')
        // --- Onboarding & Setup ---
        .addItem('ℹ️ Help & Tutorial', 'showTutorialSidebar') // Feature 10
        .addItem('🧙 Roster Setup Wizard', 'showRosterSetupWizard') // Feature 11
        .addItem('🗺️ Generate Map Template', 'generateMapTemplate') // Feature 9
        .addItem('🎲 Generate Demo Class', 'generateDemoRoster') // Feature 10
        .addSeparator()

        // --- Configuration ---
        .addItem('Configure Tables (Groups)', 'configureTables')
        .addItem('Configure Data Balancing', 'configureBalancing') // Feature 7
        .addItem('Set Room Constraints', 'configureCapacityConstraints')
        .addItem('Select Absent Students', 'showAbsenceSelector')
        .addSubMenu(ui.createMenu('Layout Manager')
            .addItem('Save Current Layout', 'saveCurrentLayout')
            .addItem('Load Saved Layout', 'loadSavedLayout')
            .addItem('Manage Saved Layouts', 'manageSavedLayouts'))
        .addSeparator()

        // --- Student Rules ---
        .addItem('Configure Preferential Seating', 'configurePreferentialSeating')
        .addItem('Configure Student Separations', 'configureSeparatedStudents')
        .addItem('Configure Student Buddies', 'configureStudentBuddies') // New Feature
        .addSeparator()

        // --- Run Randomizer ---
        .addItem('▶️ Randomly Assign Students', 'randomizeStudents')
        .addItem('🥂 Run Social Mixer (Max Variety)', 'randomizeSocialMixer') // Feature 8
        .addItem('Preview Randomization', 'previewRandomization')
        .addSeparator()

        // --- Tools ---
        .addItem('View Assignment History', 'viewAssignmentHistory')
        .addItem('Manage & Clear Settings', 'manageSettings')
        .addToUi();

    // --- FEATURE 10: First Run Check ---
    const props = PropertiesService.getDocumentProperties();
    if (!props.getProperty('firstRunComplete')) {
        showTutorialSidebar();
        props.setProperty('firstRunComplete', 'true');
    }
}

/**
 * Stores the number of tables for each room.
 * This configuration is saved to the document's properties
 * and used by the randomizeStudents function.
 */
// ============================================================================
// ============================================================================
// CORE TRIGGERS & MENUS
// ============================================================================

/**
 * Wrapper for the Social Mixer mode.
 */
function randomizeSocialMixer() {
    randomizeStudents({ socialMixer: true });
}

/**
 * Creates a custom menu in the spreadsheet UI when the document is opened.
// ============================================================================

/**
 * Opens the Table & Capacity Configuration Dialog.
 */
function configureTables() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rosters");
    if (!sheet) return SpreadsheetApp.getUi().alert("Error: 'Rosters' sheet not found.");

    // Get Room List
    const sectionData = getSectionsAndRooms(sheet);
    const rooms = [...new Set(Object.values(sectionData).map(s => s.room))];

    // Get Current Configs
    const properties = PropertiesService.getDocumentProperties();
    const tableConfig = JSON.parse(properties.getProperty('tableConfig') || '{}');
    const capacityConstraints = JSON.parse(properties.getProperty('capacityConstraints') || '{}');

    // Build data object
    const data = {
        rooms: rooms,
        tableConfig: tableConfig,
        capacityConstraints: capacityConstraints
    };

    const html = buildTableConfigHtml(data);
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(600).setHeight(500), 'Configure Tables & Capacity');
}

/**
 * Server-side handler to save Table Config
 */
function saveTableConfig(formObject) {
    const properties = PropertiesService.getDocumentProperties();
    const tableConfig = {};
    const capacityConstraints = {};

    // formObject will look like { "tables_H201": "6", "min_H201": "4", ... }
    for (const key in formObject) {
        if (key.startsWith('tables_')) {
            const room = key.substring(7);
            tableConfig[room] = parseInt(formObject[key], 10) || 0;
        } else if (key.startsWith('min_')) {
            const room = key.substring(4);
            const min = parseInt(formObject[key], 10) || 0;
            if (min > 0) capacityConstraints[room] = { min: min };
        }
    }

    properties.setProperty('tableConfig', JSON.stringify(tableConfig));
    properties.setProperty('capacityConstraints', JSON.stringify(capacityConstraints));
}

/**
 * Opens the Preferential Seating Config Dialog.
 */
function configurePreferentialSeating() {
    showConfigDialog('preferentialSeating', 'Preferential Seating');
}

/**
 * Opens the Student Separations Config Dialog.
 */
function configureSeparatedStudents() {
    showConfigDialog('separatedStudents', 'Student Separations');
}

/**
 * Opens the Student Buddies Config Dialog.
 */
function configureStudentBuddies() {
    showConfigDialog('studentBuddies', 'Student Buddies');
}

/**
 * Generic function to open the unified config dialog for different types.
 */
function showConfigDialog(configType, title) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rosters");
    if (!sheet) return SpreadsheetApp.getUi().alert("Error: 'Rosters' sheet not found.");

    const sectionData = getSectionsAndRooms(sheet);
    const properties = PropertiesService.getDocumentProperties();
    const currentConfig = JSON.parse(properties.getProperty(configType) || '{}');
    const tableConfig = JSON.parse(properties.getProperty('tableConfig') || '{}');

    const data = {
        type: configType,
        sectionData: sectionData, // { "Section": { room: "H1", students: [{name:A}...] } }
        currentConfig: currentConfig, // { "Section": { "Student": [1] } } or { "Section": [["A","B"]] }
        tableConfig: tableConfig
    };

    const html = buildGenericConfigHtml(data, title);
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(900).setHeight(700), title);
}

/**
 * Server-side handler to save Generic Config
 */
function saveGenericConfig(type, configJson) {
    // configJson is the whole object { "Section": ... }
    PropertiesService.getDocumentProperties().setProperty(type, configJson);
}

/**
 * Open Balancing Configuration Dialog (Feature 7)
 */
function configureBalancing() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rosters");
    if (!sheet) return SpreadsheetApp.getUi().alert("Error: 'Rosters' sheet not found.");

    const sectionData = getSectionsAndRooms(sheet);

    // Aggregate unique attribute keys from all students
    const allAttributes = new Set();
    Object.values(sectionData).forEach(section => {
        section.students.forEach(student => {
            if (student.attributes) {
                Object.keys(student.attributes).forEach(k => allAttributes.add(k));
            }
        });
    });

    // If no attributes found (legacy sheet with only Name/Gender implied), warn user or just show Gender
    if (allAttributes.size === 0) {
        allAttributes.add('Gender');
    }

    const currentConfig = JSON.parse(PropertiesService.getDocumentProperties().getProperty('balancingConfig') || '{}');

    const data = {
        sections: Object.keys(sectionData),
        attributes: Array.from(allAttributes),
        currentConfig: currentConfig // { "SectionName": "AttributeName" }
    };

    const html = buildBalancingConfigHtml(data);
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(500).setHeight(600), 'Configure Data Balancing');
}

function saveBalancingConfig(config) {
    PropertiesService.getDocumentProperties().setProperty('balancingConfig', JSON.stringify(config));
}

// ============================================================================
// HTML BUILDERS (Embedded for simplicity)
// ============================================================================

// ============================================================================
// HTML BUILDERS
// ============================================================================

function buildBalancingConfigHtml(data) {
    return `
    <style>
        body { font-family: 'Segoe UI', sans-serif; padding: 20px; background: #fcfcfc; }
        .desc { color: #5f6368; margin-bottom: 20px; font-size: 13px; line-height: 1.5; }
        .row { display: flex; align-items: center; justify-content: space-between; padding: 10px; background: white; border: 1px solid #dadce0; border-radius: 6px; margin-bottom: 10px; }
        .section-name { font-weight: 600; color: #3c4043; }
        select { padding: 6px; border-radius: 4px; border: 1px solid #dadce0; font-size: 14px; min-width: 150px; }
        .save-btn { width: 100%; background: #1a73e8; color: white; padding: 12px; border: none; border-radius: 4px; cursor: pointer; font-weight: 600; margin-top: 20px; }
        .save-btn:hover { background: #1557b0; }
        .clear-btn { color: #ea4335; font-size: 12px; cursor: pointer; margin-left: 10px; text-decoration: underline; }
    </style>
    <h3>Data-Driven Balancing</h3>
    <div class="desc">
        Select a column to balance for each section. The randomizer will attempt to distribute values of that column (e.g. "High", "Low", "M", "F") evenly across tables.
        <br><br>
        <i>Note: If "None" is selected, standard randomization is used.</i>
    </div>
    
    <div id="form">
        ${data.sections.map(sec => `
            <div class="row">
                <div class="section-name">${sec}</div>
                <div>
                    <select id="sel_${sec}">
                        <option value="">-- None --</option>
                        ${data.attributes.map(attr =>
        `<option value="${attr}" ${data.currentConfig[sec] === attr ? 'selected' : ''}>${attr}</option>`
    ).join('')}
                    </select>
                </div>
            </div>
        `).join('')}
    </div>

    <button class="save-btn" onclick="save()">Save Configuration</button>

    <script>
        function save() {
            const config = {};
            const sections = ${JSON.stringify(data.sections)};
            sections.forEach(sec => {
                const val = document.getElementById('sel_' + sec).value;
                if (val) config[sec] = val;
            });

            const btn = document.querySelector('.save-btn');
            btn.textContent = 'Saving...';
            btn.disabled = true;

            google.script.run
                .withSuccessHandler(() => google.script.host.close())
                .saveBalancingConfig(config);
        }
    </script>
    `;
}

function buildTableConfigHtml(data) {
    return `
    <style>
        body { font-family: 'Segoe UI', sans-serif; padding: 20px; color: #333; }
        .room-row { display: flex; align-items: center; margin-bottom: 15px; background: #f8f9fa; padding: 10px; border-radius: 6px; }
        .room-name { flex: 1; font-weight: bold; font-size: 1.1em; }
        .input-group { display: flex; flex-direction: column; margin-left: 15px; }
        label { font-size: 0.85em; color: #666; margin-bottom: 3px; }
        input { padding: 5px; border: 1px solid #ccc; border-radius: 4px; width: 80px; }
        button { background: #1a73e8; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; float: right; margin-top: 10px; }
        button:hover { background: #1557b0; }
        .desc { margin-bottom: 20px; color: #555; font-size: 0.9em; }
    </style>
    <h3>Detailed Table Configuration</h3>
    <div class="desc">Set the number of table groups for each room. Optionally, set a Minimum Group Size to force clumping (e.g. fewer tables will be used if class size is small).</div>
    <form id="configForm">
        ${data.rooms.map(room => `
            <div class="room-row">
                <div class="room-name">${room}</div>
                <div class="input-group">
                    <label># Tables</label>
                    <input type="number" name="tables_${room}" value="${data.tableConfig[room] || ''}" min="1" required>
                </div>
                <div class="input-group">
                    <label>Min Size (Optional)</label>
                    <input type="number" name="min_${room}" value="${(data.capacityConstraints[room] && data.capacityConstraints[room].min) || 0}" min="0">
                </div>
            </div>
        `).join('')}
        <button type="button" onclick="save()">Save Settings</button>
    </form>
    <script>
        function save() {
            const form = document.getElementById('configForm');
            const formData = {};
            new FormData(form).forEach((value, key) => formData[key] = value);
            
            google.script.run
                .withSuccessHandler(() => google.script.host.close())
                .saveTableConfig(formData);
        }
    </script>
    `;
}

function buildGenericConfigHtml(data, title) {
    const isGroups = (data.type === 'separatedStudents' || data.type === 'studentBuddies');
    const isBuddies = (data.type === 'studentBuddies');
    const safeData = JSON.stringify(data);

    return `
    <style>
        body { font-family: 'Segoe UI', sans-serif; display: flex; flex-direction: column; height: 95vh; margin: 0; background: #fcfcfc; }
        .header { padding: 15px; background: #fff; border-bottom: 1px solid #ddd; flex-shrink: 0; }
        
        .main { flex: 1; display: flex; gap: 20px; padding: 20px; overflow: hidden; }
        
        /* Two Column Layout */
        .col { flex: 1; display: flex; flex-direction: column; background: white; border: 1px solid #dadce0; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        
        .col-header { padding: 12px 15px; background: #f1f3f4; border-bottom: 1px solid #dadce0; font-weight: 600; color: #3c4043; display: flex; justify-content: space-between; align-items: center; }
        .help-icon { color: #5f6368; font-size: 12px; font-weight: normal; }

        .scroll-area { flex: 1; overflow-y: auto; padding: 5px; }

        /* Left Col: Student List */
        .student-list-item { padding: 8px 12px; cursor: pointer; border-radius: 4px; margin-bottom: 2px; color: #3c4043; }
        .student-list-item:hover { background: #f1f3f4; }
        .student-list-item.selected { background: #e8f0fe; color: #1967d2; font-weight: 500; }

        /* Right Col: Action Area (Top) & Rules (Bottom) */
        .action-container { padding: 15px; border-bottom: 4px solid #f1f3f4; background: #fff; }
        
        .rule-card { background: white; border: 1px solid #e0e0e0; border-radius: 6px; padding: 8px 12px; margin-bottom: 8px; display: flex; justify-content: space-between; align-items: center; font-size: 13px; }
        .rule-content { flex: 1; color: #3c4043; }
        .delete-btn { color: #5f6368; cursor: pointer; padding: 4px 8px; font-weight: bold; border-radius: 4px; }
        .delete-btn:hover { background: #fce8e6; color: #d93025; }

        select.section-select { width: 100%; padding: 8px; font-size: 14px; border-radius: 4px; border: 1px solid #dadce0; }
        button.action-btn { width: 100%; background: #1a73e8; color: white; padding: 10px; border: none; border-radius: 4px; cursor: pointer; font-weight: 500; font-size: 14px; }
        button.action-btn:disabled { background: #dadce0; color: #80868b; cursor: not-allowed; }
        
        input[type="text"] { width: 100%; padding: 8px; box-sizing: border-box; border: 1px solid #dadce0; border-radius: 4px; margin-bottom: 10px; }
        .select-freq { width: 100%; padding: 8px; margin-bottom: 10px; border: 1px solid #dadce0; border-radius: 4px; font-size: 13px; }
    </style>
    
    <div class="header">
        <label style="font-weight:bold; color:#5f6368; display:block; margin-bottom:5px;">Select Section:</label>
        <select id="sectionSelect" class="section-select" onchange="loadSection()"></select>
    </div>

    <div class="main">
        <!-- LEFT COLUMN: STUDENTS -->
        <div class="col">
            <div class="col-header">
                1. Select Student(s)
                <span class="help-icon">${isGroups ? 'Pick multiple' : 'Pick one'}</span>
            </div>
            <div id="studentList" class="scroll-area"></div>
        </div>

        <!-- RIGHT COLUMN: ACTIONS & RULES -->
        <div class="col">
            <div class="col-header">2. Assign & Manage</div>
            
            <!-- Fixed Action Area at Top -->
            <div id="actionArea" class="action-container"></div>
            
            <!-- Scrollable Rules List -->
            <div class="col-header" style="background:#fff; border-top:1px solid #dadce0; border-bottom:1px solid #e0e0e0; padding: 8px 15px; font-size:12px; text-transform:uppercase; color:#5f6368;">Existing Rules</div>
            <div id="rulesList" class="scroll-area" style="background:#f8f9fa;"></div>
        </div>
    </div>
    
    <div style="padding:15px; text-align:right; border-top:1px solid #ddd; background: #fff;">
        <button class="action-btn" style="width:auto; background:#34a853; padding: 8px 24px;" onclick="saveAll()">Save & Close</button>
    </div>

    <script>
        const DATA = ${safeData};
        const CONFIG = DATA.currentConfig;
        const SECTIONS = DATA.sectionData;
        const IS_BUDDIES = ${isBuddies};
        
        // Initializer
        const select = document.getElementById('sectionSelect');
        Object.keys(SECTIONS).forEach(s => {
            const opt = document.createElement('option');
            opt.value = s;
            opt.textContent = s;
            select.appendChild(opt);
        });

        let currentSection = select.value;
        let selectedStudents = new Set();

        function loadSection() {
            currentSection = select.value;
            selectedStudents.clear();
            renderStudents();
            renderRules();
            renderActionArea();
        }

        function toggleStudent(name) {
            if (selectedStudents.has(name)) {
                selectedStudents.delete(name);
            } else {
                // If Preferential, only allow 1 selection
                if (!${isGroups}) selectedStudents.clear();
                selectedStudents.add(name);
            }
            renderStudents();
            renderActionArea();
        }

        function renderStudents() {
            const div = document.getElementById('studentList');
            div.innerHTML = '';
            const students = SECTIONS[currentSection].students;
            
            students.forEach(s => {
                const el = document.createElement('div');
                el.className = 'student-list-item' + (selectedStudents.has(s.name) ? ' selected' : '');
                el.textContent = s.name;
                el.onclick = () => toggleStudent(s.name);
                div.appendChild(el);
            });
        }

        function renderActionArea() {
            const div = document.getElementById('actionArea');
            div.innerHTML = '';
            
            if (${isGroups}) {
                // Frequency Selector (Buddies Only)
                if (IS_BUDDIES) {
                    const freqSelect = document.createElement('select');
                    freqSelect.id = 'freqSelect';
                    freqSelect.className = 'select-freq';
                    freqSelect.innerHTML = 
                        '<option value="1.0">Always Together (100%)</option>' +
                        '<option value="0.75">Often (75%)</option>' +
                        '<option value="0.5">Sometimes (50%)</option>' +
                        '<option value="0.25">Rarely (25%)</option>';
                    div.appendChild(freqSelect);
                }

                const btn = document.createElement('button');
                btn.className = 'action-btn';
                btn.textContent = 'Create Group';
                btn.disabled = selectedStudents.size < 2;
                btn.onclick = addGroupRule;
                div.appendChild(btn);
            } else {
                // Preferential
                if (selectedStudents.size === 1) {
                    const name = Array.from(selectedStudents)[0];
                    const numTables = DATA.tableConfig[SECTIONS[currentSection].room] || 0;
                    
                    div.innerHTML = 
                        '<div style="margin-bottom:5px;">Table(s) for <b>' + name + '</b>:</div>' +
                        '<input type="text" id="prefInput" placeholder="e.g. 1, 3" style="width:100%; padding:8px; box-sizing:border-box; margin-bottom:10px;">' +
                        '<button class="action-btn" onclick="addPrefRule(\\'' + name + '\\')">Set Preference</button>';
                } else {
                    div.innerHTML = '<div style="color:#888;">Select a student...</div>';
                }
            }
        }

        function addGroupRule() {
            if (!CONFIG[currentSection]) CONFIG[currentSection] = [];

            const names = Array.from(selectedStudents);

            if (IS_BUDDIES) {
                 const chance = parseFloat(document.getElementById('freqSelect').value);
                 CONFIG[currentSection].push({ names: names, chance: chance });
            } else {
                 CONFIG[currentSection].push(names);
            }

            selectedStudents.clear();
            renderStudents();
            renderRules();
            renderActionArea();
        }

        function addPrefRule(name) {
            const val = document.getElementById('prefInput').value;
            const tables = val.split(',').map(n => parseInt(n.trim())).filter(n => !isNaN(n) && n > 0);
            
            if (tables.length === 0) return alert('Invalid table numbers');
            
            if (!CONFIG[currentSection]) CONFIG[currentSection] = {};
            CONFIG[currentSection][name] = tables;
            
            selectedStudents.clear();
            renderStudents();
            renderRules();
            renderActionArea();
        }

        function deleteRule(idxOrKey) {
            if (${isGroups}) {
                CONFIG[currentSection].splice(idxOrKey, 1);
            } else {
                delete CONFIG[currentSection][idxOrKey];
            }
            renderRules();
        }

        function renderRules() {
            const div = document.getElementById('rulesList');
            div.innerHTML = '';
            const rules = CONFIG[currentSection];

            if (!rules || (Array.isArray(rules) && rules.length === 0) || (typeof rules === 'object' && Object.keys(rules).length === 0)) {
                div.innerHTML = '<div style="color:#999; text-align:center; padding:20px;">No rules set for this section.</div>';
                return;
            }

            if (${isGroups}) {
                // Array of items
                rules.forEach((group, idx) => {
                    const card = document.createElement('div');
                    card.className = 'rule-card';
                    
                    let content = '';
                    if (IS_BUDDIES && group.chance !== undefined) {
                        const pct = (group.chance * 100) + '%';
                        content = group.names.join(' + ') + ' <span style="color:#1a73e8; font-weight:bold; margin-left:5px;">(' + pct + ')</span>';
                    } else if (Array.isArray(group)) {
                         // Backwards compat or Separated Students
                         content = group.join(' + ');
                    } else if (typeof group === 'object' && group.names) {
                         // Handle incomplete object case if any
                         content = group.names.join(' + ');
                    }

                    card.innerHTML = 
                        '<div class="rule-content">' + content + '</div>' +
                        '<div class="delete-btn" onclick="deleteRule(' + idx + ')">✕</div>';
                    div.appendChild(card);
                });
            } else {
                // Object Key-Value (Preferential)
                Object.keys(rules).forEach(name => {
                    const tables = rules[name];
                    const card = document.createElement('div');
                    card.className = 'rule-card';
                    card.innerHTML = 
                        '<div class="rule-content"><b>' + name + '</b> ⮕ Table(s) ' + tables.join(', ') + '</div>' +
                        '<div class="delete-btn" onclick="deleteRule(\\'' + name + '\\')">✕</div>';
                    div.appendChild(card);
                });
            }
        }

        function saveAll() {
            // Clean up empty sections
            Object.keys(CONFIG).forEach(k => {
                if (Array.isArray(CONFIG[k]) && CONFIG[k].length === 0) delete CONFIG[k];
                if (!Array.isArray(CONFIG[k]) && Object.keys(CONFIG[k]).length === 0) delete CONFIG[k];
            });

            google.script.run
                .withSuccessHandler(() => google.script.host.close())
                .saveGenericConfig('${data.type}', JSON.stringify(CONFIG));
        }

                // Init
                loadSection();
            </script>
    `;
}

// ============================================================================
// ABSENT STUDENT FUNCTIONS
// ============================================================================

/**
 * Opens a sidebar for selecting absent students.
 */
function showAbsenceSelector() {
    const html = HtmlService.createHtmlOutput(buildAbsenceSidebarHtml())
        .setTitle('Select Absent Students')
        .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Builds the HTML for the absence sidebar.
 * @returns {string} HTML string.
 */
function buildAbsenceSidebarHtml() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("Rosters");
    if (!sheet) return "Error: 'Rosters' sheet not found.";

    const sectionData = getSectionsAndRooms(sheet);

    // Get currently saved absences to pre-check boxes
    const properties = PropertiesService.getDocumentProperties();
    const savedAbsences = JSON.parse(properties.getProperty('absentStudents') || '[]');

    let html = `
        < style >
        body { font - family: Arial, sans - serif; padding: 10px; }
        .section - header { font - weight: bold; margin - top: 15px; margin - bottom: 5px; color: #4285f4; border - bottom: 1px solid #ddd; }
        .student - item { margin: 5px 0; }
        label { cursor: pointer; }
        .buttons { margin - top: 20px; position: sticky; bottom: 0; background: white; padding: 10px 0; border - top: 1px solid #eee; }
        button { width: 100 %; padding: 10px; background: #ea4335; color: white; border: none; border - radius: 4px; cursor: pointer; font - size: 14px; margin - bottom: 5px; }
    button.save { background: #34a853; }
    button.clear { background: #f1f3f4; color: black; border: 1px solid #ddd; }
    </style >
    <form id="absenceForm">
    `;

    for (const sectionName in sectionData) {
        html += `<div class="section-header">${sectionName}</div>`;
        const students = sectionData[sectionName].students.sort((a, b) => a.name.localeCompare(b.name));

        students.forEach(s => {
            const isChecked = savedAbsences.includes(s.name) ? 'checked' : '';
            html += `
            <div class="student-item">
                <label>
                    <input type="checkbox" name="absent" value="${s.name}" ${isChecked}>
                    ${s.name}
                </label>
            </div>`;
        });
    }

    html += `
        <div class="buttons">
            <button type="button" class="save" onclick="saveAbsences()">Save & Close</button>
            <button type="button" class="clear" onclick="clearAbsences()">Clear All Absences</button>
        </div>
    </form>
    <script>
        function saveAbsences() {
            const checkboxes = document.querySelectorAll('input[name="absent"]:checked');
            const absentNames = Array.from(checkboxes).map(cb => cb.value);
            google.script.run.withSuccessHandler(() => google.script.host.close()).saveAbsenceList(absentNames);
        }
        function clearAbsences() {
             const checkboxes = document.querySelectorAll('input[name="absent"]');
             checkboxes.forEach(cb => cb.checked = false);
             google.script.run.saveAbsenceList([]); // Save empty list but keep sidebar open
        }
    </script>
    `;

    return html;
}

/**
 * Saves the list of absent students to DocumentProperties.
 * @param {Array<string>} absentNames - List of names.
 */
function saveAbsenceList(absentNames) {
    const properties = PropertiesService.getDocumentProperties();
    properties.setProperty('absentStudents', JSON.stringify(absentNames));
}

// ============================================================================
// ASSIGNMENT HISTORY FUNCTIONS (Sticky Table Avoidance)
// ============================================================================


/**
 * Gets the assignment history for all sections.
 * History structure: { "Section 1": { "StudentName": [3, 1, 5], ... }, ... }
 * Arrays store the last N table indices (0-based) the student was assigned to.
 * @returns {Object} The assignment history object.
 */
function getAssignmentHistory() {
    const properties = PropertiesService.getDocumentProperties();
    return JSON.parse(properties.getProperty('assignmentHistory') || '{}');
}

/**
 * Saves assignment history after a successful randomization.
 * Keeps only the last HISTORY_DEPTH assignments per student.
 * @param {Object} allAssignments - Map of section names to their table assignments.
 */
function saveAssignmentHistory(allAssignments) {
    const HISTORY_DEPTH = 5; // Number of past assignments to remember
    const properties = PropertiesService.getDocumentProperties();
    const history = getAssignmentHistory();

    for (const sectionName in allAssignments) {
        if (!history[sectionName]) history[sectionName] = {};
        const sectionHistory = history[sectionName];
        const tableAssignments = allAssignments[sectionName];

        // Loop through each table
        tableAssignments.forEach((table, tableIndex) => {
            // Loop through each student at this table
            table.forEach(student => {
                if (!sectionHistory[student.name]) sectionHistory[student.name] = [];
                // Add the current table index to the front of the history
                sectionHistory[student.name].unshift(tableIndex);
                // Keep only the last HISTORY_DEPTH entries
                if (sectionHistory[student.name].length > HISTORY_DEPTH) {
                    sectionHistory[student.name].pop();
                }
            });
        });
    }

    properties.setProperty('assignmentHistory', JSON.stringify(history));

    // --- FEATURE 8: Social Mixer History (Pair Tracking) ---
    const pairHistory = getPairHistory();

    for (const sectionName in allAssignments) {
        if (!pairHistory[sectionName]) pairHistory[sectionName] = {};
        const sectionPairs = pairHistory[sectionName];
        const tableAssignments = allAssignments[sectionName];

        // Loop through each table to identify pairs
        tableAssignments.forEach(table => {
            // Compare every student with every other student at the table
            for (let i = 0; i < table.length; i++) {
                const s1 = table[i].name;
                if (!sectionPairs[s1]) sectionPairs[s1] = [];

                for (let j = i + 1; j < table.length; j++) {
                    const s2 = table[j].name;
                    if (!sectionPairs[s2]) sectionPairs[s2] = [];

                    // Record pairing if not already recorded (Set-like behavior via array check)
                    // We simply append. Duplicates indicate multiple interactions (which might be useful info later).
                    // For now, let's just make it a unique set to save space? 
                    // Verify if we want "sat together ONCE" or "sat together 5 times". 
                    // Social Mixer usually demands "New People". So simple existence check is enough.
                    if (!sectionPairs[s1].includes(s2)) sectionPairs[s1].push(s2);
                    if (!sectionPairs[s2].includes(s1)) sectionPairs[s2].push(s1);
                }
            }
        });
    }
    properties.setProperty('pairHistory', JSON.stringify(pairHistory));
}

/**
 * Gets the history of student pairings.
 * Structure: { "Section A": { "StudentName": ["Partner1", "Partner2"], ... } }
 */
function getPairHistory() {
    const properties = PropertiesService.getDocumentProperties();
    return JSON.parse(properties.getProperty('pairHistory') || '{}');
}

/**
 * Calculates a penalty score for placing a student at a given table.
 * Higher penalty = student has been at this table recently.
 * @param {string} studentName - The student's name.
 * @param {number} tableIndex - The 0-based table index.
 * @param {Object} sectionHistory - History for this section.
 * @returns {number} Penalty score (0 = no history, higher = recent repeat).
 */
function getHistoryPenalty(studentName, tableIndex, sectionHistory) {
    if (!sectionHistory || !sectionHistory[studentName]) return 0;

    const history = sectionHistory[studentName];
    let penalty = 0;

    // Check each historical assignment, with more recent ones weighted higher
    history.forEach((pastTable, recency) => {
        if (pastTable === tableIndex) {
            // More recent = higher penalty (recency 0 = most recent)
            penalty += 10 / (recency + 1); // e.g., 10, 5, 3.3, 2.5, 2
        }
    });

    return penalty;
}

/**
 * Displays the current assignment history for all sections.
 */
function viewAssignmentHistory() {
    const ui = SpreadsheetApp.getUi();
    const history = getAssignmentHistory();

    if (Object.keys(history).length === 0) {
        ui.alert('Assignment History', 'No assignment history recorded yet. Run the randomizer at least once.', ui.ButtonSet.OK);
        return;
    }

    let display = '';
    for (const section in history) {
        display += `\n === ${section} ===\n`;
        const students = Object.keys(history[section]).sort();
        students.forEach(student => {
            const tables = history[section][student].map(t => `T${t + 1} `).join(', ');
            display += `${student}: ${tables} \n`;
        });
    }

    ui.alert('Assignment History (Most Recent First)', display, ui.ButtonSet.OK);
}

// ============================================================================
// FAILURE REPORTING FUNCTIONS
// ============================================================================

/**
 * Creates a detailed failure report for constraint violations.
 * @param {Array} separationFailures - Array of {studentA, studentB, tableIndex, reason}
 * @param {Array} genderFailures - Array of {tableIndex, maleCount, femaleCount}
 * @param {Array} placementFailures - Array of student names that couldn't be placed
 * @returns {string} Human-readable failure report.
 */
function formatFailureReport(separationFailures, genderFailures, placementFailures) {
    let report = '';

    if (placementFailures.length > 0) {
        report += '⚠️ PLACEMENT FAILURES:\n';
        placementFailures.forEach(f => {
            report += `• ${f.studentName}: ${f.reason} \n`;
        });
        report += '\n';
    }

    if (separationFailures.length > 0) {
        report += '⚠️ SEPARATION VIOLATIONS:\n';
        separationFailures.forEach(f => {
            report += `• ${f.studentA} and ${f.studentB} are at Table ${f.tableIndex + 1} \n`;
            if (f.reason) report += `  Reason: ${f.reason} \n`;
        });
        report += '\n';
    }

    if (genderFailures.length > 0) {
        report += '⚠️ GENDER BALANCE ISSUES:\n';
        genderFailures.forEach(f => {
            report += `• Table ${f.tableIndex + 1}: ${f.maleCount} M / ${f.femaleCount} F\n`;
        });
        report += '\n';
    }

    return report;
}

/**
 * Analyzes why a separation violation couldn't be fixed.
 * @param {Object} studentA - Student object
 * @param {Object} studentB - Student object
 * @param {number} tableIndex - The table they're both at
 * @param {Object} prefsForSection - Preferential seating config
 * @param {number} numTables - Total number of tables
 * @returns {string} Explanation of why the violation persists.
 */
function analyzeSeparationFailure(studentA, studentB, tableIndex, prefsForSection, numTables) {
    const aPrefs = prefsForSection[studentA.name] || null;
    const bPrefs = prefsForSection[studentB.name] || null;

    // Both have prefs at same table
    if (aPrefs && bPrefs) {
        const aHasThis = aPrefs.includes(tableIndex + 1);
        const bHasThis = bPrefs.includes(tableIndex + 1);
        if (aHasThis && bHasThis) {
            return `Both have preferential seating at Table ${tableIndex + 1} `;
        }
    }

    // One is locked to this table
    if (aPrefs && aPrefs.length === 1 && aPrefs[0] === tableIndex + 1) {
        return `${studentA.name} is locked to Table ${tableIndex + 1} `;
    }
    if (bPrefs && bPrefs.length === 1 && bPrefs[0] === tableIndex + 1) {
        return `${studentB.name} is locked to Table ${tableIndex + 1} `;
    }

    return 'No valid swap found after maximum attempts';
}

// ============================================================================
// PREVIEW MODE FUNCTIONS
// ============================================================================

/**
 * Preview randomization - shows proposed arrangement before committing.
 */
function previewRandomization() {
    const ui = SpreadsheetApp.getUi();

    // Generate the randomization without writing
    const result = generateRandomization();

    if (!result.success) {
        ui.alert('Error', result.error, ui.ButtonSet.OK);
        return;
    }

    // Build preview HTML
    const html = buildPreviewHtml(result);
    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(1000).setHeight(800), 'Preview Assignments');
}

/**
 * Confirms the randomization.
 * NOW accepts the assignments assignments data from the client side!
 * This enables "What You See Is What You Get" (including manual DnD moves).
 */
function confirmRandomization(data) {
    // data should contain { allAssignments, sectionMetadata, failureReport }
    if (!data || !data.allAssignments) {
        const result = generateRandomization();
        data = result;
    }

    if (data.allAssignments) {
        writeRandomizationToSheet(data.allAssignments, data.sectionMetadata, data.failureReport);
        saveAssignmentHistory(data.allAssignments);

        // Save this result as the "Last Run" so it can be saved as a Golden Layout if desired
        PropertiesService.getDocumentProperties().setProperty('lastRunResult', JSON.stringify(data));

        return true; // Signal success to client
    }
    return false;
}

/**
 * Builds the HTML for the preview dialog with Drag-and-Drop support.
 */
function buildPreviewHtml(result) {
    // Embed the initial state as a JSON string
    const initialState = JSON.stringify(result);

    let html = `
        < style >
        body { font - family: 'Segoe UI', Tahoma, Geneva, Verdana, sans - serif; background - color: #f4f6f8; margin: 0; padding: 20px; }
        .info - box { background: #e8f0fe; color: #1a73e8; padding: 10px; margin - bottom: 10px; border - radius: 4px; font - size: 13px; border: 1px solid #d2e3fc; }
        .controls { display: flex; justify - content: flex - end; gap: 10px; margin - bottom: 20px; position: sticky; top: 0; background: #f4f6f8; padding: 10px 0; z - index: 100; border - bottom: 1px solid #ddd; }
        button { padding: 10px 20px; font - size: 14px; border: none; border - radius: 4px; cursor: pointer; transition: background 0.2s; }
        .accept { background - color: #34a853; color: white; font - weight: bold; }
        .accept:hover { background - color: #2d8e47; }
        .reroll { background - color: #4285f4; color: white; }
        .reroll:hover { background - color: #357abd; }
        .cancel { background - color: #ea4335; color: white; }
        .cancel:hover { background - color: #d32f2f; }
        .action { background - color: #9334e6; color: white; }
        .action:hover { background - color: #7c22c7; }
        
        .section - container { background: white; border - radius: 8px; box - shadow: 0 1px 3px rgba(0, 0, 0, 0.1); margin - bottom: 20px; padding: 15px; }
        .section - title { font - size: 18px; font - weight: bold; color: #202124; margin - bottom: 5px; border - bottom: 2px solid #e8eaed; padding - bottom: 5px; }
        .section - room { font - size: 14px; color: #5f6368; margin - bottom: 15px; }
        
        .tables - grid { display: grid; grid - template - columns: repeat(auto - fill, minmax(140px, 1fr)); gap: 10px; }
        .table - card { background: #f8f9fa; border: 1px solid #dadce0; border - radius: 8px; padding: 10px; min - height: 100px; display: flex; flex - direction: column; }
        .table - header { font - weight: bold; text - align: center; color: #5f6368; margin - bottom: 8px; font - size: 12px; text - transform: uppercase; }
        
        .student - chip {
        background: white; border: 1px solid #dadce0; border - radius: 16px; padding: 4px 10px; margin - bottom: 4px; font - size: 13px; color: #3c4043; cursor: grab; user - select: none; box - shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
        display: flex; justify - content: space - between; align - items: center;
    }
        .student - chip:hover { box - shadow: 0 2px 4px rgba(0, 0, 0, 0.1); border - color: #bdc1c6; }
        .student - chip:active { cursor: grabbing; box - shadow: 0 4px 8px rgba(0, 0, 0, 0.2); }
        .student - chip.dragging { opacity: 0.5; border: 1px dashed #4285f4; }
        
        .table - card.drag - over { background: #e8f0fe; border: 2px dashed #4285f4; }
        
        .warning - box { background - color: #fce8e6; border - left: 5px solid #ea4335; padding: 10px; margin - bottom: 20px; border - radius: 4px; color: #c5221f; font - size: 13px; white - space: pre - wrap; }
    </style >

    <div class="info-box">
        <b>💡 Tips:</b> Drag & Drop students to swap them. Use buttons to adjust the layout.
    </div>

    <div class="controls">
        <button class="accept" title="Finalize and write to sheet" onclick="confirmAssignments()">✓ Accept & Write</button>
        <button class="reroll" title="Discard and generate new randomization" onclick="google.script.run.withSuccessHandler(closeDialog).previewRandomization()">↻ Re-roll</button>
        <button class="action" title="Shift all groups to the next table (1→2, Last→1)" onclick="rotateTables()">⏩ Rotate</button>
        <button class="action" title="Randomize which table each group sits at" onclick="shuffleGroups()">🔀 Shuffle</button>
        <button class="cancel" title="Close without saving" onclick="google.script.host.close()">✕ Cancel</button>
    </div>

    <div id="content"></div>

    <script>
        // Use the passed-in state to render the initial view
        let currentState = ${initialState};

        function render() {
            const container = document.getElementById('content');
            container.innerHTML = ''; // Clear current

            if (currentState.failureReport) {
                const warn = document.createElement('div');
                warn.className = 'warning-box';
                warn.textContent = "WARNINGS:\\n" + currentState.failureReport;
                container.appendChild(warn);
            }

            for (const sectionName in currentState.allAssignments) {
                const sectionDiv = document.createElement('div');
                sectionDiv.className = 'section-container';

                const title = document.createElement('div');
                title.className = 'section-title';
                title.textContent = sectionName;
                sectionDiv.appendChild(title);

                const room = document.createElement('div');
                room.className = 'section-room';
                room.textContent = currentState.sectionMetadata[sectionName];
                sectionDiv.appendChild(room);

                const grid = document.createElement('div');
                grid.className = 'tables-grid';

                currentState.allAssignments[sectionName].forEach((tableStudents, tableIndex) => {
                    const card = document.createElement('div');
                    card.className = 'table-card';
                    card.dataset.section = sectionName;
                    card.dataset.tableIndex = tableIndex;

                    // Header
                    const hdr = document.createElement('div');
                    hdr.className = 'table-header';
                    hdr.textContent = 'Table ' + (tableIndex + 1);
                    card.appendChild(hdr);

                    // Drop Zone Handlers
                    card.addEventListener('dragover', handleDragOver);
                    card.addEventListener('dragleave', handleDragLeave);
                    card.addEventListener('drop', handleDrop);

                    // Student Chips
                    tableStudents.forEach((student, studentIndex) => {
                        const chip = document.createElement('div');
                        chip.className = 'student-chip';
                        chip.draggable = true;
                        chip.textContent = student.name;
                        
                        // Store IDs for drag logic
                        chip.dataset.section = sectionName;
                        chip.dataset.tableIndex = tableIndex;
                        chip.dataset.studentIndex = studentIndex;
                        chip.dataset.studentName = student.name;

                        chip.addEventListener('dragstart', handleDragStart);
                        chip.addEventListener('dragend', handleDragEnd);
                        
                        card.appendChild(chip);
                    });

                    grid.appendChild(card);
                });

                sectionDiv.appendChild(grid);
                container.appendChild(sectionDiv);
            }
        }

        // --- Drag and Drop Logic ---
        let dragSource = null; // Stores {section, tableIndex, studentIndex}

        function handleDragStart(e) {
            this.classList.add('dragging');
            dragSource = {
                section: this.dataset.section,
                table: parseInt(this.dataset.tableIndex),
                studentIdx: parseInt(this.dataset.studentIndex),
                name: this.dataset.studentName,
                obj: currentState.allAssignments[this.dataset.section][this.dataset.tableIndex][this.dataset.studentIndex]
            };
            e.dataTransfer.effectAllowed = 'move';
        }

        function handleDragEnd(e) {
            this.classList.remove('dragging');
            document.querySelectorAll('.table-card').forEach(c => c.classList.remove('drag-over'));
            dragSource = null;
        }

        function handleDragOver(e) {
            // Allow drop only if same section
            if (dragSource && dragSource.section === this.dataset.section) {
                e.preventDefault();
                this.classList.add('drag-over');
                e.dataTransfer.dropEffect = 'move';
            }
        }

        function handleDragLeave(e) {
            this.classList.remove('drag-over');
        }

        function handleDrop(e) {
            e.preventDefault();
            this.classList.remove('drag-over');

            const targetSection = this.dataset.section;
            const targetTableIdx = parseInt(this.dataset.tableIndex);

            if (!dragSource || dragSource.section !== targetSection) return;

            // Update State
            // 1. Remove from old
            const oldTable = currentState.allAssignments[dragSource.section][dragSource.table];
            oldTable.splice(dragSource.studentIdx, 1);

            // 2. Add to new
            const newTable = currentState.allAssignments[targetSection][targetTableIdx];
            newTable.push(dragSource.obj);

            // Re-render
            render();
        }

        function confirmAssignments() {
            // Disable button to prevent double click
            document.querySelector('.accept').disabled = true;
            document.querySelector('.accept').textContent = 'Saving...';
            
            google.script.run
                .withSuccessHandler(closeDialog)
                .withFailureHandler(err => {
                    alert('Error: ' + err);
                    document.querySelector('.accept').disabled = false;
                })
                .confirmRandomization(currentState);
        }

        function closeDialog() {
            google.script.host.close();
        }

        // Initialize
        render();

        // --- Table Rotation Features ---
        function rotateTables() {
            for (const section in currentState.allAssignments) {
                const tables = currentState.allAssignments[section];
                if (tables.length > 1) {
                    // Move last table to first position (Shift 1 -> 2, 2 -> 3...)
                    tables.unshift(tables.pop());
                }
            }
            render();
        }

        function shuffleGroups() {
            for (const section in currentState.allAssignments) {
                let tables = currentState.allAssignments[section];
                // Fisher-Yates Shuffle of the table groups
                for (let i = tables.length - 1; i > 0; i--) {
                    const j = Math.floor(Math.random() * (i + 1));
                    [tables[i], tables[j]] = [tables[j], tables[i]];
                }
            }
            render();
        }
    </script>
    `;

    return html;
}

// ============================================================================
// CORE RANDOMIZATION FUNCTIONS
// ============================================================================

/**
 * The MAIN function. Runs the entire randomization and outputs to the "Tables" sheet.
 * It applies all rules in a specific order of priority:
 * 1. Preferential Seating (must-have)
 * 2. Student Separation (must-have)
 * 3. Gender Balancing (nice-to-have, will try its best)
 * 
 * Now acts as a wrapper that calls the generation and writing functions.
 * @param {Object} options - Optional settings (e.g. { socialMixer: true })
 */
function randomizeStudents(options = {}) {
    // Generate the randomization data
    const result = generateRandomization(options);

    if (result.success) {
        // Write to sheet
        writeRandomizationToSheet(result.allAssignments, result.sectionMetadata, result.failureReport);

        // --- FEATURE 9: Visual Map ---
        // If a visual map template exists, update it.
        updateVisualMap(result.allAssignments);

        // Save history
        saveAssignmentHistory(result.allAssignments);

        // Save as Last Run for Golden Layouts
        PropertiesService.getDocumentProperties().setProperty('lastRunResult', JSON.stringify(result));

        // Show warnings if any, but "success" means we generated a valid arrangement
        if (result.failureReport) {
            SpreadsheetApp.getUi().alert('Randomization Complete with Warnings', result.failureReport, SpreadsheetApp.getUi().ButtonSet.OK);
        }
    } else {
        SpreadsheetApp.getUi().alert('Error', result.error, SpreadsheetApp.getUi().ButtonSet.OK);
    }
}

/**
 * Core logic generator. Runs the algorithm but does not write to the sheet.
 * Returns { success: boolean, allAssignments: Object, error?: string, failureReport?: string }
 */
function generateRandomization(options = {}) {
    const properties = PropertiesService.getDocumentProperties();
    const storedConfig = properties.getProperty('tableConfig');
    if (!storedConfig) {
        return { success: false, error: 'Please set up your table numbers first.' };
    }

    const tableConfig = JSON.parse(storedConfig);
    const constraintConfig = JSON.parse(properties.getProperty('capacityConstraints') || '{}'); // NEW
    const preferentialConfig = JSON.parse(properties.getProperty('preferentialSeating') || '{}');
    const separatedConfig = JSON.parse(properties.getProperty('separatedStudents') || '{}');
    const buddiesConfig = JSON.parse(properties.getProperty('studentBuddies') || '{}');
    const balancingConfig = JSON.parse(properties.getProperty('balancingConfig') || '{}'); // NEW Feature 7
    const assignmentHistory = getAssignmentHistory();
    const pairHistory = getPairHistory(); // NEW Feature 8
    const absentStudents = JSON.parse(properties.getProperty('absentStudents') || '[]'); // NEW

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("Rosters");
    if (!sheet) return { success: false, error: "'Rosters' sheet not found." };

    const sectionData = getSectionsAndRooms(sheet);
    const allAssignments = {};
    const sectionMetadata = {}; // Maps Section Name -> Room Name

    // Arrays to collect failures for the report
    const separationFailures = [];
    const genderFailures = [];
    const placementFailures = [];

    for (const sectionName in sectionData) {
        const { room: roomName, students: allStudentsInSection } = sectionData[sectionName];

        // --- FEATURE 1: Filter Absent Students ---
        // We filter them out completely for this run.
        const activeStudents = allStudentsInSection.filter(s => !absentStudents.includes(s.name));

        if (activeStudents.length === 0) continue;

        sectionMetadata[sectionName] = roomName;
        let numTables = tableConfig[roomName];
        if (!numTables) {
            return { success: false, error: `No table config for room "${roomName}".Skipping section "${sectionName}".` };
        }

        // --- FEATURE 2: Capacity Constraints (Min Students per Table) ---
        // If there's a minimum, we might need to use fewer tables to promote clumping.
        if (constraintConfig[roomName] && constraintConfig[roomName].min > 0) {
            const minPerTable = constraintConfig[roomName].min;
            const maxTablesPossible = Math.floor(activeStudents.length / minPerTable);

            // If we have 20 students and min 4, maxTables is 5.
            // If config says 6 tables, we MUST reduce to 5 to satisfy min constraint.
            // If we have 5 students and min 4, maxTables is 1.

            if (maxTablesPossible < numTables) {
                // If maxTablesPossible is 0 (e.g. 3 students, min 4), we fallback to 1 table.
                numTables = Math.max(1, maxTablesPossible);
            }
        }

        const prefsForSection = preferentialConfig[sectionName] || {};
        const sepGroupsForSection = separatedConfig[sectionName] || [];

        // --- FEATURE: Student Buddies Frequency ---
        // Normalize the config to a simple array of "Active" buddy groups for this run.
        // Legacy: Array<String>. New: { names: [], chance: 0.x }
        const rawBuddies = buddiesConfig[sectionName] || [];
        const buddiesForSection = [];

        rawBuddies.forEach(group => {
            if (Array.isArray(group)) {
                // Legacy or 100% group
                buddiesForSection.push(group);
            } else if (group.names && group.chance !== undefined) {
                // Check probability
                if (Math.random() <= group.chance) {
                    buddiesForSection.push(group.names);
                }
            } else if (group.names) {
                // Fallback
                buddiesForSection.push(group.names);
            }
        });

        const historyForSection = assignmentHistory[sectionName] || {};
        const pairsForSection = pairHistory[sectionName] || {}; // { StudentA: [B, C] }
        const balancingAttribute = balancingConfig[sectionName];

        // --- FEATURE 7: Data Balancing Prep ---
        // Calculate max allowed students per attribute value (e.g. "High": 2, "Low": 3)
        const maxPerTableByAttribute = {};
        if (balancingAttribute) {
            const counts = {};
            activeStudents.forEach(s => {
                const val = (s.attributes && s.attributes[balancingAttribute]) ? s.attributes[balancingAttribute] : "N/A";
                counts[val] = (counts[val] || 0) + 1;
            });
            Object.keys(counts).forEach(val => {
                // We allow ceiling (e.g. 10 students / 4 tables = 2.5 -> Max 3)
                maxPerTableByAttribute[val] = Math.ceil(counts[val] / numTables);
            });
        }

        const tableAssignments = Array(numTables).fill().map(() => []);
        let studentsToRandomize = [...activeStudents];

        // --- STEP 1: Preferential Seating ---
        activeStudents.forEach(student => {
            if (prefsForSection[student.name]) {
                const preferredTables = prefsForSection[student.name];
                const targetTableIndex = preferredTables[Math.floor(Math.random() * preferredTables.length)] - 1;
                if (targetTableIndex >= 0 && targetTableIndex < numTables) {
                    tableAssignments[targetTableIndex].push(student);
                    studentsToRandomize = studentsToRandomize.filter(s => s !== student);
                }
            }
        });

        // --- STEP 1b: Target Capacities ---
        const baseSize = Math.floor(activeStudents.length / numTables);
        const remainder = activeStudents.length % numTables;
        const targetCapacities = Array(numTables).fill(0).map((_, i) => baseSize + (i < remainder ? 1 : 0));

        // --- STEP 2: Distribute Remaining ---
        const shuffledStudents = shuffleArray(studentsToRandomize);

        let studentIdx = 0;
        while (studentIdx < shuffledStudents.length) {
            const student = shuffledStudents[studentIdx];
            let placed = false;

            // Helper to score a candidate table
            // Lower score is better.
            const scoreTable = (tableIdx) => {
                let score = 0;
                // Penalty 1: Separation Violation (High Penalty)
                if (tableAssignments[tableIdx].some(seated => areSeparated(student.name, seated.name, sepGroupsForSection))) {
                    score += 1000;
                }
                // Penalty 2: History (Sticky Table Avoidance)
                score += getHistoryPenalty(student.name, tableIdx, historyForSection);

                // Bonus: Buddy Attraction (Negative Score)
                // Bonus: Buddy Attraction (Negative Score)
                if (tableAssignments[tableIdx].some(seated => areBuddies(student.name, seated.name, buddiesForSection))) {
                    score -= 50;
                }

                if (tableAssignments[tableIdx].some(seated => areBuddies(student.name, seated.name, buddiesForSection))) {
                    score -= 50;
                }

                // --- FEATURE 8: Social Mixer ---
                // If in Social Mixer mode, heavily prioritize tables where the student knows NO ONE.
                if (options.socialMixer) {
                    const studentPairs = pairsForSection[student.name] || [];
                    const seatedStudents = tableAssignments[tableIdx];

                    // Check if they have met ANYONE at this table
                    const knownCount = seatedStudents.filter(s => studentPairs.includes(s.name)).length;

                    if (knownCount > 0) {
                        score += (knownCount * 200); // 200pts per known person -> Strong avoidance
                    } else {
                        score -= 100; // Bonus for completely new group
                    }
                }

                // --- FEATURE 7: Balancing Penalty ---
                if (balancingAttribute) {
                    const val = (student.attributes && student.attributes[balancingAttribute]) ? student.attributes[balancingAttribute] : "N/A";
                    const currentCount = tableAssignments[tableIdx].filter(s => {
                        const sVal = (s.attributes && s.attributes[balancingAttribute]) ? s.attributes[balancingAttribute] : "N/A";
                        return sVal === val;
                    }).length;

                    if (currentCount >= maxPerTableByAttribute[val]) {
                        score += 250; // Significant penalty for exceeding balanced distribution
                    }
                }

                return score;
            };

            // Get all valid tables (capacity check only)
            const validTables = [];
            for (let i = 0; i < numTables; i++) {
                if (tableAssignments[i].length < targetCapacities[i]) {
                    validTables.push(i);
                }
            }

            if (validTables.length > 0) {
                // Sort tables by lowest score (best fit)
                // We add a tiny random decimal to break ties randomly
                validTables.sort((a, b) => (scoreTable(a) + Math.random()) - (scoreTable(b) + Math.random()));

                // Pick the best one
                const bestTableIdx = validTables[0];
                tableAssignments[bestTableIdx].push(student);
                placed = true;
            }

            if (!placed) {
                placementFailures.push({ studentName: student.name, reason: "Capacity reached/logic error" });
            }
            studentIdx++;
        }

        // --- STEP 3: Fix Separation Violations ---
        fixSeparationViolations(tableAssignments, sepGroupsForSection, prefsForSection);

        // Check for remaining separation violations
        const remainingViolations = findViolations(tableAssignments, sepGroupsForSection);
        remainingViolations.forEach(v => {
            const reason = analyzeSeparationFailure(v.studentA, v.studentB, v.tableIndex, prefsForSection, numTables);
            separationFailures.push({
                studentA: v.studentA.name,
                studentB: v.studentB.name,
                tableIndex: v.tableIndex,
                reason: reason
            });
        });

        // --- STEP 4: Fix Gender Balance ---
        fixGenderBalance(tableAssignments, sepGroupsForSection, prefsForSection);

        // Check for gender issues
        tableAssignments.forEach((t, idx) => {
            if (getGenderTableScore(t) >= 20) {
                let m = 0, f = 0;
                t.forEach(s => s.gender === 'M' ? m++ : (s.gender === 'F' ? f++ : null));
                genderFailures.push({ tableIndex: idx, maleCount: m, femaleCount: f });
            }
        });

        // --- STEP 5: Final Shuffle ---
        for (let i = 0; i < tableAssignments.length; i++) {
            tableAssignments[i] = shuffleArray(tableAssignments[i]);
        }

        allAssignments[sectionName] = tableAssignments;
    }

    const failureReport = formatFailureReport(separationFailures, genderFailures, placementFailures);

    return {
        success: true,
        allAssignments: allAssignments,
        sectionMetadata: sectionMetadata,
        failureReport: failureReport
    };
}

/**
 * Writes the given assignments to the "Tables" sheet.
 */
function writeRandomizationToSheet(allAssignments, sectionMetadata, failureReport) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let outputSheet = spreadsheet.getSheetByName("Tables");
    if (!outputSheet) {
        outputSheet = spreadsheet.insertSheet("Tables");
    } else {
        outputSheet.clear();
    }

    let rowIndex = 1;
    let maxColumnsUsed = 0;
    const colors = [
        '#FFB6C1', '#ADD8E6', '#90EE90', '#FFD700', '#FF8C69', '#DDA0DD', '#F08080',
        '#E0FFFF', '#FAFAD2', '#D3D3D3', '#FFA07A', '#20B2AA', '#87CEFA', '#778899'
    ];
    const fontSize = 14;

    for (const sectionName in allAssignments) {
        const tableAssignments = allAssignments[sectionName];
        const roomName = sectionMetadata[sectionName];
        const numTables = tableAssignments.length;
        const halfTables = Math.ceil(numTables / 2);
        maxColumnsUsed = Math.max(maxColumnsUsed, halfTables);

        const outputHeader = `${sectionName} (${roomName})`;
        outputSheet.getRange(rowIndex, 1).setValue(outputHeader)
            .setFontWeight("bold")
            .setFontSize(18)
            .setFontFamily("Roboto")
            .setHorizontalAlignment("center");
        if (halfTables > 1) outputSheet.getRange(rowIndex, 1, 1, halfTables).mergeAcross();
        rowIndex++;

        const maxRowsPerHalf = Math.max(0, ...tableAssignments.map(t => t.length));

        // Front tables
        for (let t = 0; t < halfTables; t++) {
            outputSheet.getRange(rowIndex, t + 1).setValue(`Table ${t + 1} `)
                .setFontWeight("bold")
                .setFontSize(fontSize)
                .setFontFamily("Roboto")
                .setHorizontalAlignment("center")
                .setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
            if (tableAssignments[t]) {
                const numStudents = tableAssignments[t].length;
                const tableColor = colors[t % colors.length];

                for (let j = 0; j < numStudents; j++) {
                    outputSheet.getRange(rowIndex + 1 + j, t + 1)
                        .setValue(tableAssignments[t][j].name)
                        .setBackground(tableColor)
                        .setFontSize(fontSize)
                        .setFontFamily("Roboto")
                        .setHorizontalAlignment("center")
                        .setVerticalAlignment("middle");
                }
                // Apply borders to the whole block
                // 1. Inner borders = Background Color (effectively invisible)
                // 2. Outer borders = Black (frame)
                const block = outputSheet.getRange(rowIndex + 1, t + 1, numStudents, 1);
                block.setBorder(null, null, null, null, true, true, tableColor, SpreadsheetApp.BorderStyle.SOLID);
                block.setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
            }
        }
        rowIndex += maxRowsPerHalf + 2;

        // Back tables
        for (let t = halfTables; t < numTables; t++) {
            outputSheet.getRange(rowIndex, (t - halfTables) + 1).setValue(`Table ${t + 1} `)
                .setFontWeight("bold")
                .setFontSize(fontSize)
                .setFontFamily("Roboto")
                .setHorizontalAlignment("center")
                .setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
            if (tableAssignments[t]) {
                const numStudents = tableAssignments[t].length;
                const tableColor = colors[t % colors.length];

                for (let j = 0; j < numStudents; j++) {
                    outputSheet.getRange(rowIndex + 1 + j, (t - halfTables) + 1)
                        .setValue(tableAssignments[t][j].name)
                        .setBackground(tableColor)
                        .setFontSize(fontSize)
                        .setFontFamily("Roboto")
                        .setHorizontalAlignment("center")
                        .setVerticalAlignment("middle");
                }
                // Apply borders to the whole block
                const block = outputSheet.getRange(rowIndex + 1, (t - halfTables) + 1, numStudents, 1);
                block.setBorder(null, null, null, null, true, true, tableColor, SpreadsheetApp.BorderStyle.SOLID);
                block.setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
            }
        }
        rowIndex += maxRowsPerHalf + 2;
    }

    // Add warnings if present (also written to sheet for record)
    if (failureReport) {
        outputSheet.getRange(rowIndex, 1).setValue("WARNINGS:").setFontWeight("bold").setFontColor("red");
        outputSheet.getRange(rowIndex + 1, 1).setValue(failureReport).setWrap(true);
        rowIndex += 3;
    }

    outputSheet.getRange(rowIndex, 1).setValue("Last randomized: " + new Date().toLocaleString()).setFontSize(12).setFontStyle("italic");
    rowIndex++;

    // Clean up the sheet by deleting unused rows/columns.
    const totalRows = outputSheet.getMaxRows();
    const totalColumns = outputSheet.getMaxColumns();
    if (rowIndex < totalRows) outputSheet.deleteRows(rowIndex + 1, totalRows - rowIndex);
    if (maxColumnsUsed < totalColumns) outputSheet.deleteColumns(maxColumnsUsed + 1, totalColumns - maxColumnsUsed);
}

// ============================================================================
// LAYOUT MANAGER FUNCTIONS (Golden Layouts)
// ============================================================================

/**
 * Saves the current "Last Assignment" (from history logic) as a named layout.
 * Note: This saves the MOST RECENT result generated/written.
 */
function saveCurrentLayout() {
    const ui = SpreadsheetApp.getUi();
    const properties = PropertiesService.getDocumentProperties();

    // We can allow the user to save the *last written* assignment from history?
    // Or we have to capture the state. 
    // Problem: `allAssignments` isn't fully persisted globally.
    // Solution: We rely on `assignmentHistory` which tracks per-student history, 
    // BUT that's designed for "sticky table avoidance", not "snapshotting a whole class".
    // Better: We should have saved the *last run* to a temporary property `lastRunResult`.
    // Let's modify `confirmRandomization` to save `lastRunResult` first.
    // For now, I will assume we can't save unless we just ran it, OR 
    // we parse the SHEET content? Parsing sheet is risky (formatting).

    // Alternative: We just save what's in `lastRunResult` property.
    // I need to add this property update to `confirmRandomization`.
    const lastRunJson = properties.getProperty('lastRunResult');
    if (!lastRunJson) {
        ui.alert('No recent randomization found to save. Please run a randomization first.');
        return;
    }

    const response = ui.prompt('Save Layout', 'Enter a name for this layout (e.g., "Lab Partners", "Exam Seating"):', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return;

    const layoutName = response.getResponseText().trim();
    if (!layoutName) {
        ui.alert('Invalid name.');
        return;
    }

    const savedLayouts = JSON.parse(properties.getProperty('savedLayouts') || '{}');
    if (savedLayouts[layoutName]) {
        const confirm = ui.alert('Overwrite?', `Layout "${layoutName}" already exists.Overwrite ? `, ui.ButtonSet.YES_NO);
        if (confirm !== ui.Button.YES) return;
    }

    // Save the data
    savedLayouts[layoutName] = JSON.parse(lastRunJson); // This is { allAssignments, sectionMetadata }
    properties.setProperty('savedLayouts', JSON.stringify(savedLayouts));

    ui.alert('Success', `Layout "${layoutName}" saved.`, ui.ButtonSet.OK);
}

/**
 * Loads a saved layout into the Preview Window.
 */
function loadSavedLayout() {
    const ui = SpreadsheetApp.getUi();
    const properties = PropertiesService.getDocumentProperties();
    const savedLayouts = JSON.parse(properties.getProperty('savedLayouts') || '{}');
    const layoutNames = Object.keys(savedLayouts);

    if (layoutNames.length === 0) {
        ui.alert('No saved layouts found.');
        return;
    }

    // Since we can't easily make a dropdown in a simple Alert, using HTML sidebar/dialog is better.
    // But for simplicity, let's use a prompt or simple list if small.
    // Let's use a small HTML dialog to pick.
    const html = `
        < style > body{ font - family: sans - serif; padding: 10px; } .item{ padding: 5px; border - bottom: 1px solid #eee; cursor: pointer; color:#00f; } .item:hover{ background: #eee; }</style >
    <h3>Select Layout</h3>
    <div id="list"></div>
    <script>
        const names = ${JSON.stringify(layoutNames)};
        const div = document.getElementById('list');
        names.forEach(n => {
            const el = document.createElement('div');
            el.className = 'item';
            el.textContent = n;
            el.onclick = () => google.script.run.withSuccessHandler(() => google.script.host.close()).loadLayoutByName(n);
            div.appendChild(el);
        });
    </script>
    `;
    ui.showModalDialog(HtmlService.createHtmlOutput(html).setHeight(300).setWidth(300), 'Load Layout');
}

/**
 * Helper to actually load the layout by name.
 */
function loadLayoutByName(layoutName) {
    const properties = PropertiesService.getDocumentProperties();
    const savedLayouts = JSON.parse(properties.getProperty('savedLayouts') || '{}');
    const layoutData = savedLayouts[layoutName];

    if (!layoutData) return; // Error

    // SYNC ROSTER logic
    // We need to check if the saved layout matches current roster.
    // 1. Get current Roster
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("Rosters");
    const sectionData = getSectionsAndRooms(sheet); // { SectionName: { room, students: [{name}] } }

    const validatedAssignments = {};
    let failureReport = "Loaded from layout: " + layoutName + "\n";
    let hasIssues = false;

    for (const sectionName in sectionData) {
        if (!layoutData.allAssignments[sectionName]) {
            // Section exists now but not in saved layout? Initialize empty/random?
            // For now, ignore or warn.
            continue;
        }

        const currentStudents = sectionData[sectionName].students.map(s => s.name);
        // Deep copy the saved table structure
        const savedTables = layoutData.allAssignments[sectionName]; // [[{name:A}, {name:B}], ...]

        // Reconstruct tables filtering out dropped students
        const newTables = savedTables.map(t => {
            return t.filter(s => currentStudents.includes(s.name)); // Only keep currently enrolled
        });

        // Find students who are enrolled but NOT in the saved layout (New students)
        const flatSaved = new Set(savedTables.flat().map(s => s.name));
        const newStudents = sectionData[sectionName].students.filter(s => !flatSaved.has(s.name));

        if (newStudents.length > 0) {
            hasIssues = true;
            failureReport += `[${sectionName}] ${newStudents.length} new student(s) added to Table 1: ${newStudents.map(s => s.name).join(', ')} \n`;
            // Append to Table 1 (or overflow)
            if (newTables.length > 0) {
                newTables[0].push(...newStudents);
            } else {
                newTables.push(newStudents);
            }
        }

        validatedAssignments[sectionName] = newTables;
    }

    const result = {
        success: true,
        allAssignments: validatedAssignments,
        sectionMetadata: layoutData.sectionMetadata,
        failureReport: hasIssues ? failureReport : null
    };

    // Open Preview
    const html = buildPreviewHtml(result);
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(1000).setHeight(800), 'Preview Saved Layout');
}

/**
 * Manage (Delete) Saved Layouts
 */
function manageSavedLayouts() {
    const properties = PropertiesService.getDocumentProperties();
    const savedLayouts = JSON.parse(properties.getProperty('savedLayouts') || '{}');
    const layoutNames = Object.keys(savedLayouts);

    if (layoutNames.length === 0) {
        SpreadsheetApp.getUi().alert('No saved layouts to manage.');
        return;
    }

    // Simple prompt to clear all or logic to delete specific is hard with basic UI.
    // Let's just provide "Clear ALL" for now or use the HTML dialog approach again?
    // HTML dialog for delete is better.
    const html = `
        < style > body{ font - family: sans - serif; padding: 10px; } .item{ padding: 5px; border - bottom: 1px solid #eee; display: flex; justify - content: space - between; align - items: center; } .del{ color: red; cursor: pointer; font - weight: bold; } </style >
    <h3>Manage Layouts</h3>
    <div id="list"></div>
    <script>
        const names = ${JSON.stringify(layoutNames)};
        function render() {
            const div = document.getElementById('list');
            div.innerHTML = '';
            names.forEach(n => {
                const el = document.createElement('div');
                el.className = 'item';
                el.innerHTML = '<span>' + n + '</span> <span class="del" onclick="deleteLayout(\\''+n+'\\')">✕</span>';
                div.appendChild(el);
            });
            if (names.length === 0) div.innerHTML = 'No layouts.';
        }
        function deleteLayout(name) {
            if(confirm('Delete "' + name + '"?')) {
                google.script.run.withSuccessHandler(() => {
                    const idx = names.indexOf(name);
                    if (idx > -1) names.splice(idx, 1);
                    render();
                }).deleteSavedLayout(name);
            }
        }
        render();
    </script>
    `;
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setHeight(400).setWidth(400), 'Manage Layouts');
}

function deleteSavedLayout(name) {
    const properties = PropertiesService.getDocumentProperties();
    const savedLayouts = JSON.parse(properties.getProperty('savedLayouts') || '{}');
    delete savedLayouts[name];
    properties.setProperty('savedLayouts', JSON.stringify(savedLayouts));
}

/**
 * Reads the "Rosters" sheet and builds the core data structure.
 *
 * --- ASSUMPTION ---
 * This function assumes a specific layout in "Rosters":
 * Row 1: Section Names (e.g., "AP Bio", blank, "Chem", blank)
 * Row 2: Room Names   (e.g., "H201", blank, "H306", blank)
 * Row 3+: Students. Col A = Name, Col B = Gender. Col C = Name, Col D = Gender.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The "Rosters" sheet object.
 * @returns {Object} A map of section data.
 */
function getSectionsAndRooms(sheet) {
    // Get all data from the sheet at once for efficiency.
    const data = sheet.getDataRange().getValues();
    if (data.length <= 2) return {}; // Not enough data

    const sections = data[0]; // Row 1: Section Names
    const rooms = data[1];    // Row 2: Room Names
    const sectionData = {};   // { "Section Name": { room: "H1", students: [...] } }

    // Loop through columns to find "Section Blocks"
    for (let col = 0; col < sections.length; col++) {
        const sectionName = String(sections[col]).trim();
        if (!sectionName) continue;

        const roomName = String(rooms[col]).trim() || "Unknown Room";

        // Find the width of this block (look for next section or end of sheet)
        let nextCol = col + 1;
        while (nextCol < sections.length && !String(sections[nextCol]).trim()) {
            nextCol++;
        }
        const blockWidth = nextCol - col; // Number of columns for this section

        // Detect Headers (Row 3 / Index 2)
        // If the first cell matches "name" or "student", we treat it as a header row.
        let startRow = 2;
        let attributeNames = [];

        const firstCell = (data.length > 2) ? String(data[2][col]).trim().toLowerCase() : "";
        const hasHeaders = (firstCell === 'name' || firstCell === 'student' || firstCell === 'student name');

        if (hasHeaders) {
            startRow = 3; // Data starts at Row 4
            for (let offset = 0; offset < blockWidth; offset++) {
                attributeNames.push(String(data[2][col + offset]).trim() || `Attr_${offset} `);
            }
        } else {
            // Default Headers for legacy support
            attributeNames.push('Name');
            if (blockWidth > 1) attributeNames.push('Gender'); // Assume Col 2 is Gender
            for (let offset = 2; offset < blockWidth; offset++) {
                attributeNames.push(`Column ${offset + 1} `);
            }
        }

        const studentsInSection = [];

        // Read Student Rows
        for (let row = startRow; row < data.length; row++) {
            const studentName = String(data[row][col]).trim();
            if (!studentName) continue;

            // Build Student Object
            const student = {
                name: studentName,
                gender: '', // Default, will overwrite if "Gender" attribute exists
                attributes: {}
            };

            // Read Attributes
            for (let offset = 0; offset < blockWidth; offset++) {
                // If this is the Name column (offset 0), skip unless we want strictly attributes
                // Actually, let's store everything in attributes for completeness?
                // But specifically pull out 'Gender' for backward compatibility.
                const val = String(data[row][col + offset]).trim();
                const attrName = attributeNames[offset];

                student.attributes[attrName] = val;

                if (attrName.toLowerCase() === 'gender') {
                    student.gender = val.toUpperCase();
                }
            }
            // Fallback: If no explicit gender column but width >= 2 and legacy mode, use Col 2
            if (!hasHeaders && blockWidth >= 2 && !student.gender) {
                student.gender = String(data[row][col + 1]).trim().toUpperCase();
            }

            studentsInSection.push(student);
        }

        // Merge logic for split sections (e.g. AP Bio in Col A and Col E)
        if (sectionData[sectionName]) {
            sectionData[sectionName].students.push(...studentsInSection);
        } else {
            sectionData[sectionName] = {
                room: roomName,
                students: studentsInSection
            };
        }

        // Advance loop to skip the columns we just processed
        col = nextCol - 1;
    }

    return sectionData;
}

// ============================================================================
// VISUAL CLASSROOM MAP (Feature 9)
// ============================================================================

/**
 * Generates a starter "Visual Map" sheet with placeholders.
 */
function generateMapTemplate() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let mapSheet = spreadsheet.getSheetByName("Visual Map");
    if (!mapSheet) {
        mapSheet = spreadsheet.insertSheet("Visual Map");
    } else {
        SpreadsheetApp.getUi().alert('Visual Map sheet already exists.');
        return;
    }

    const sheet = spreadsheet.getSheetByName("Rosters");
    if (!sheet) return;
    const sectionData = getSectionsAndRooms(sheet);

    let currentRow = 2;
    // For each section, print instructions and placeholders
    for (const sectionName in sectionData) {
        const roomName = sectionData[sectionName].room;
        const properties = PropertiesService.getDocumentProperties();
        const storedConfig = properties.getProperty('tableConfig');
        const config = JSON.parse(storedConfig || '{}');
        const numTables = config[roomName] || 5; // Default to 5 if unknown

        mapSheet.getRange(currentRow, 1).setValue(`-- - ${sectionName} (${roomName})--- `).setFontWeight('bold');
        currentRow += 2;

        for (let i = 1; i <= numTables; i++) {
            // Placeholders: {{SectionName|TableIndex}}
            // We use a pipe or unique separator
            mapSheet.getRange(currentRow, 1).setValue(`{ {${sectionName}| ${i} } } `).setBackground('#FFF2CC').setBorder(true, true, true, true, null, null);
            currentRow += 6; // Leave space for students
        }
        currentRow += 4;
    }

    SpreadsheetApp.getUi().alert('Visual Map Template created. Rearrange the yellow boxes anywhere you like! The script will find them and list students below them.');
}


/**
 * Updates the Visual Map sheet by finding placeholders and listing students.
 */
function updateVisualMap(allAssignments) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const mapSheet = spreadsheet.getSheetByName("Visual Map");
    if (!mapSheet) return; // Feature inactive if sheet missing

    // 1. Clear previous student lists (but keep placeholders)
    // Basic approach: We assume students are written BELOW the placeholder.
    // We can't just clear everything.
    // Optimally, we find placeholders, then clear the 10 cells below them.

    const dataRange = mapSheet.getDataRange();
    const values = dataRange.getValues();
    const clears = [];

    // 2. Find Placeholders
    for (let r = 0; r < values.length; r++) {
        for (let c = 0; c < values[r].length; c++) {
            const cellVal = String(values[r][c]);
            const match = cellVal.match(/\{\{([^|]+)\|(\d+)\}\}/); // Matches {{Section|1}}

            if (match) {
                const sectionName = match[1];
                const tableIndex = parseInt(match[2]) - 1; // 0-based

                // Clear 10 rows below
                // We use a helper to batch clears? GAS is slow with individual calls.
                // We'll just define a range efficiently.
                mapSheet.getRange(r + 2, c + 1, 10, 1).clearContent();

                // Write students
                if (allAssignments[sectionName] && allAssignments[sectionName][tableIndex]) {
                    const students = allAssignments[sectionName][tableIndex].map(s => [s.name]);
                    if (students.length > 0) {
                        mapSheet.getRange(r + 2, c + 1, students.length, 1).setValues(students);
                    }
                }
            }
        }
    }
}



/**
 * Helper: Attempts to resolve student separation violations by swapping students.
 * This function will *not* move students with preferential seating.
 *
 * @param {Array<Array<Object>>} tableAssignments
 * @param {Array<Array<string>>} separatedGroups
 * @param {Object} prefsForSection
 * @returns {boolean} True if all violations were resolved, otherwise false.
 */
function fixSeparationViolations(tableAssignments, separatedGroups, prefsForSection) {
    if (separatedGroups.length === 0) return true; // No rules to fix.

    // This is a safety valve to prevent an infinite loop if the rules
    // are impossible to solve.
    const MAX_ATTEMPTS = 200;

    for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt++) {
        // First, find all current violations.
        let violations = findViolations(tableAssignments, separatedGroups);
        if (violations.length === 0) return true; // Success! No violations left.

        // Pick one random violation to try and fix.
        const violation = violations[Math.floor(Math.random() * violations.length)];
        // 'studentB' is the student OBJECT we need to move.
        const { tableIndex, studentB } = violation;
        let swapped = false;

        // Try to swap studentB with *any other student* in *any other table*.
        for (let otherTableIdx = 0; otherTableIdx < tableAssignments.length; otherTableIdx++) {
            if (otherTableIdx === tableIndex) continue; // Don't swap with same table

            for (let otherStudentIdx = 0; otherStudentIdx < tableAssignments[otherTableIdx].length; otherStudentIdx++) {
                // 'studentToSwap' is the other student OBJECT.
                const studentToSwap = tableAssignments[otherTableIdx][otherStudentIdx];

                // This is the "gatekeeper" check.
                // It asks, "Is it safe to swap Student B and StudentToSwap?"
                // "Will this swap break any *other* separation rules or preference rules?"
                if (isSwapValid(studentB, studentToSwap, tableAssignments[tableIndex], tableAssignments[otherTableIdx], separatedGroups, prefsForSection, tableIndex, otherTableIdx)) {
                    // If the swap is "valid", perform it.
                    // Swap the objects in the arrays.
                    tableAssignments[tableIndex][tableAssignments[tableIndex].indexOf(studentB)] = studentToSwap;
                    tableAssignments[otherTableIdx][otherStudentIdx] = studentB;
                    swapped = true;
                    break; // Succeeded, so stop looping through 'otherStudentIdx'
                }
            }
            if (swapped) break; // Succeeded, so stop looping through 'otherTableIdx'
        }
    }

    // After MAX_ATTEMPTS, we check one last time.
    // If violations still exist, we "quit" and return false.
    return findViolations(tableAssignments, separatedGroups).length === 0;
}

/**
 * --- NEW Helper Function ---
 * Attempts to improve the total "Gender Score" of the room.
 * It uses a randomized "Hill Climbing" algorithm to find the best configuration.
 *
 * @param {Array<Array<Object>>} tableAssignments - Array of tables with student OBJECTS.
 * @param {Array<Array<string>>} separatedGroups - Groups of student NAMES.
 * @param {Object} prefsForSection - Map of preferential student NAMES.
 * @returns {boolean} True if no "Bad" (Male Bias) tables remain, false otherwise.
 */
function fixGenderBalance(tableAssignments, separatedGroups, prefsForSection) {
    const MAX_ATTEMPTS = 5000; // Increased attempts to ensure convergence with stricter rules

    // Identify tables that have students
    const populatedTables = [];
    tableAssignments.forEach((table, index) => {
        if (table.length > 0) populatedTables.push({ index, table });
    });

    // If there's 1 or fewer tables, we can't swap.
    if (populatedTables.length <= 1) return true;

    for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt++) {
        // 1. Pick two random distinct tables
        const t1_idx = Math.floor(Math.random() * populatedTables.length);
        let t2_idx = Math.floor(Math.random() * populatedTables.length);
        while (t1_idx === t2_idx) t2_idx = Math.floor(Math.random() * populatedTables.length);

        const tableA_wrapper = populatedTables[t1_idx];
        const tableB_wrapper = populatedTables[t2_idx];
        const tableA = tableA_wrapper.table;
        const tableB = tableB_wrapper.table;

        // 2. Pick a random student from each
        const sA_idx = Math.floor(Math.random() * tableA.length);
        const sB_idx = Math.floor(Math.random() * tableB.length);
        const studentA = tableA[sA_idx];
        const studentB = tableB[sB_idx];

        // Optimization: Calculate the combined score of these two tables BEFORE the swap
        const scoreBefore = getGenderTableScore(tableA) + getGenderTableScore(tableB);

        // 3. Check if swap is valid (Separation/Preference rules)
        if (isSwapValid(studentA, studentB, tableA, tableB, separatedGroups, prefsForSection, tableA_wrapper.index, tableB_wrapper.index)) {

            // 4. Perform Swap
            tableA[sA_idx] = studentB;
            tableB[sB_idx] = studentA;

            // 5. Calculate score AFTER the swap
            const scoreAfter = getGenderTableScore(tableA) + getGenderTableScore(tableB);

            // 6. DECISION:
            // If the new score is Lower (better) or Equal, keep the swap.
            // Allowing Equal swaps helps shuffle the board to find new solutions.
            if (scoreAfter <= scoreBefore) {
                // Keep swap
            } else {
                // Revert swap (Result was worse)
                tableA[sA_idx] = studentA;
                tableB[sB_idx] = studentB;
            }
        }
    }

    // Final Check: Do we have any "Bad" (Score 20) tables left?
    // If we have tables with score 20 (Male Bias), return false to trigger warning.
    const badTables = tableAssignments.filter(t => getGenderTableScore(t) >= 20);
    return badTables.length === 0;
}


/**
 * --- NEW Helper Function ---
 * Finds tables that are considered "Bad" configurations (Male Bias).
 * Used for generating warnings at the end of the script.
 * @param {Array<Array<Object>>} tableAssignments
 * @returns {Array<Object>} Array of table arrays that failed the check.
 */
function findGenderViolations(tableAssignments) {
    const violations = [];
    tableAssignments.forEach((table) => {
        // We only report a violation if the score is >= 20 (Male Bias).
        // Scores of 0 (Balanced), 1 (Monogender), and 2 (Female Bias) are considered success.
        if (getGenderTableScore(table) >= 20) {
            violations.push(table);
        }
    });
    return violations;
}

/**
 * --- NEW Helper Function ---
 * Calculates a "penalty score" for a table based on gender balance preferences.
 * Lower is better.
 *
 * Hierarchy of Needs:
 * 0  = Perfect Balance (Strict 50/50).
 * 1  = All One Gender (Privileged over unbalanced mixed tables)
 * 2  = Female Bias (More F than M)
 * 20 = Male Bias (More M than F - this is the least desirable)
 *
 * @param {Array<Object>} table - A single table array of student objects.
 * @returns {number} The penalty score.
 */
function getGenderTableScore(table) {
    if (table.length === 0) return 0; // Empty is neutral

    let maleCount = 0;
    let femaleCount = 0;

    table.forEach(student => {
        if (student.gender === 'M') maleCount++;
        else if (student.gender === 'F') femaleCount++;
    });

    // PRIORITY 1: Strict 50/50 Split
    if (maleCount === femaleCount) return 0;

    // PRIORITY 2: All One Gender (Privileged status)
    if (maleCount === 0 || femaleCount === 0) return 1;

    // PRIORITY 3: More Females than Males
    if (femaleCount > maleCount) return 2;

    // PRIORITY 4: More Males than Females (Mixed) - Least desirable
    // This includes scenarios like 2M/1F which previously passed as "balanced"
    // because the difference was only 1. Now they are strictly penalized.
    return 20;
}


/**
 * Helper: Finds all separation violations in the current table assignments.
 *
 * @param {Array<Array<Object>>} tableAssignments
 * @param {Array<Array<string>>} separatedGroups
 * @returns {Array<Object>} A list of violation objects.
 */
function findViolations(tableAssignments, separatedGroups) {
    const violations = [];

    // Create a "lookup map" for quick checking.
    const separationMap = new Map();
    separatedGroups.forEach(group => {
        for (const student of group) {
            if (!separationMap.has(student)) separationMap.set(student, new Set());
            group.forEach(s => { if (s !== student) separationMap.get(student).add(s); });
        }
    });

    // Now, check every table.
    tableAssignments.forEach((table, tableIndex) => {
        // Check every student against every *other* student *in the same table*.
        for (let i = 0; i < table.length; i++) {
            for (let j = i + 1; j < table.length; j++) {
                const studentA = table[i]; // This is an object
                const studentB = table[j]; // This is an object

                // Check by name: Is studentA's name on studentB's "do not sit with" list?
                if (separationMap.has(studentA.name) && separationMap.get(studentA.name).has(studentB.name)) {
                    // If yes, this is a violation.
                    violations.push({ tableIndex, studentA, studentB }); // Push the objects
                }
            }
        }
    });
    return violations;
}

/**
 * Helper: This is the "Master Gatekeeper" function.
 * It checks if swapping two students would create any new separation
 * or preference violations. It does NOT check gender.
 *
 * @param {Object} studentToMove
 * @param {Object} studentToTakePlace
 * @param {Array<Object>} sourceTable
 * @param {Array<Object>} destTable
 * @param {Array<Array<string>>} separatedGroups
 * @param {Object} prefsForSection
 * @param {number} sourceTableIndex
 * @param {number} destTableIndex
 * @returns {boolean} True if the swap is "safe", otherwise false.
 */
function isSwapValid(studentToMove, studentToTakePlace, sourceTable, destTable, separatedGroups, prefsForSection, sourceTableIndex, destTableIndex) {

    // --- Check Separation Rules ---

    // Check 1: Will 'studentToMove' conflict with anyone in its new table ('destTable')?
    for (const student of destTable) {
        if (student === studentToTakePlace) continue; // Don't check against the one leaving
        // Check by name
        if (areSeparated(studentToMove.name, student.name, separatedGroups)) return false;
    }

    // Check 2: Will 'studentToTakePlace' conflict with anyone in its new table ('sourceTable')?
    for (const student of sourceTable) {
        if (student === studentToMove) continue; // Don't check against the one leaving
        // Check by name
        if (areSeparated(studentToTakePlace.name, student.name, separatedGroups)) return false;
    }

    // --- Check Preferential Seating Rules ---

    // Check 3: Is 'studentToMove' preferential? If so, is its new table ('destTable')
    // one of its allowed tables?
    if (prefsForSection && prefsForSection[studentToMove.name]) {
        // The table index is 0-based, but preferences are 1-based (e.g., "Table 1").
        if (!prefsForSection[studentToMove.name].includes(destTableIndex + 1)) {
            return false; // Invalid: This swap moves them out of their preferred table.
        }
    }

    // Check 4: Is 'studentToTakePlace' preferential? If so, is its new table
    // ('sourceTable') one of its allowed tables?
    if (prefsForSection && prefsForSection[studentToTakePlace.name]) {
        if (!prefsForSection[studentToTakePlace.name].includes(sourceTableIndex + 1)) {
            return false; // Invalid: This swap moves them out of their preferred table.
        }
    }

    // If all 4 checks passed, the swap is safe.
    return true;
}

/**
 * Helper: Checks if two students are in the same separation group.
 *
 * @param {string} studentA - Name of student A
 * @param {string} studentB - Name of student B
 * @param {Array<Array<string>>} separatedGroups
 * @returns {boolean} True if they must be separated.
 */
function areSeparated(studentA, studentB, separatedGroups) {
    for (const group of separatedGroups) {
        // If both students are found in the *same* group array...
        if (group.includes(studentA) && group.includes(studentB)) {
            return true; // ...they must be separated.
        }
    }
    return false;
}

/**
 * Helper: Checks if two students are in the same buddy group.
 * Reuses logic similar to areSeparated but for attraction.
 *
 * @param {string} studentA - Name of student A
 * @param {string} studentB - Name of student B
 * @param {Array<Array<string>>} buddyGroups
 * @returns {boolean} True if they are buddies.
 */
function areBuddies(studentA, studentB, buddyGroups) {
    for (const group of buddyGroups) {
        if (group.includes(studentA) && group.includes(studentB)) {
            return true;
        }
    }
    return false;
}

/**
 * Randomizes array elements using the Fisher-Yates algorithm.
 * This is a standard, efficient way to shuffle an array.
 * @param {Array} array The array to shuffle.
 * @returns {Array} The shuffled array.
 */
function shuffleArray(array) {
    if (!array || array.length === 0) return [];
    // Create a copy so we don't change the original array.
    const newArray = [...array];
    // Loop backwards from the end of the array.
    for (let i = newArray.length - 1; i > 0; i--) {
        // Pick a random index 'j' from 0 up to 'i'.
        const j = Math.floor(Math.random() * (i + 1));
        // Swap the elements at 'i' and 'j'.
        [newArray[i], newArray[j]] = [newArray[j], newArray[i]];
    }
    return newArray;
}

/**
 * Removes unused rows and columns from the sheet to clean up the output.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to clean.
 * @param {number} usedRows The number of rows that contain data.
 * @param {number} usedColumns The number of columns that contain data.
 */
function removeUnusedCells(sheet, usedRows, usedColumns) {
    const totalRows = sheet.getMaxRows();
    const totalColumns = sheet.getMaxColumns();

    // If we used 50 rows and the sheet has 1000, delete rows 51-1000.
    if (usedRows <= totalRows) {
        sheet.deleteRows(usedRows, totalRows - usedRows + 1);
    }

    // If we used 3 columns and the sheet has 26, delete columns 4-26.
    if (usedColumns < totalColumns) {
        sheet.deleteColumns(usedColumns + 1, totalColumns - usedColumns);
    }
}

/**
 * Allows the user to clear settings for individual functions or all at once.
 */
function manageSettings() {
    const ui = SpreadsheetApp.getUi();
    const properties = PropertiesService.getDocumentProperties();

    // 'promptText' uses backticks (`) to create a multi - line string.
    const promptText = `
Enter the number for the setting you wish to clear:

1. Table Configuration
2. Preferential Seating Rules
3. Student Separation Rules
4. Student Buddies Rules
5. --- CLEAR ALL SETTINGS ---
  `;

    const response = ui.prompt('Manage & Clear Settings', promptText, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) {
        return;
    }

    const choice = response.getResponseText().trim();

    // A 'switch' statement is a clean alternative to many 'if/else if' blocks.
    switch (choice) {
        case '1':
            confirmAndClear('tableConfig', 'Table Configuration');
            break;
        case '2':
            confirmAndClear('preferentialSeating', 'Preferential Seating Rules');
            break;
        case '3':
            confirmAndClear('separatedStudents', 'Student Separation Rules');
            break;
        case '4':
            confirmAndClear('studentBuddies', 'Student Buddies Rules');
            break;
        case '5':
            // Clearing all needs a special, more stern confirmation.
            const confirmAll = ui.alert(
                'Confirm Reset All',
                'Are you sure you want to clear ALL saved settings? This action cannot be undone.',
                ui.ButtonSet.YES_NO
            );
            if (confirmAll === ui.Button.YES) {
                // Delete all four properties.
                properties.deleteProperty('tableConfig');
                properties.deleteProperty('preferentialSeating');
                properties.deleteProperty('separatedStudents');
                properties.deleteProperty('studentBuddies');
                ui.alert('Success', 'All saved randomizer settings have been cleared.', ui.ButtonSet.OK);
            }
            break;
        default:
            ui.alert('Invalid Selection', 'Please enter a number from 1 to 5.', ui.ButtonSet.OK);
            break;
    }
}



/**
 * Helper function for manageSettings to confirm and clear a single property.
 * This avoids repeating the same confirmation logic 3 times.
 *
 * @param {string} propertyName The key of the property to delete (e.g., 'tableConfig').
 * @param {string} displayName The user-friendly name (e.g., 'Table Configuration').
 */
function confirmAndClear(propertyName, displayName) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
        `Confirm Clear: ${displayName}`,
        `Are you sure you want to clear all saved ${displayName}? This cannot be undone.`,
        ui.ButtonSet.YES_NO
    );

    // If the user clicks "YES":
    if (response == ui.Button.YES) {
        try {
            // Delete the property from the document.
            PropertiesService.getDocumentProperties().deleteProperty(propertyName);
            ui.alert('Success', `All saved ${displayName} have been cleared.`, ui.ButtonSet.OK);
        } catch (e) {
            ui.alert('Error', `An error occurred while trying to clear ${displayName}. Please try again.`, ui.ButtonSet.OK);
        }
    }
}

/**
 * Helper function to prompt the user to select a section.
 * Used by configurePreferentialSeating and configureSeparatedStudents.
 *
 * @param {GoogleAppsScript.Base.Ui} ui The UI object.
 * @param {Object} sectionData The section data object.
 * @returns {string|null} The selected section name, or null if cancelled/invalid.
 */
function promptForSection(ui, sectionData) {
    const sectionNames = Object.keys(sectionData);
    if (sectionNames.length === 0) {
        ui.alert("No sections with students found in 'Rosters' sheet.");
        return null;
    }

    // Create a numbered list for the user to select from.
    const numberedSections = sectionNames.map((name, index) => `${index + 1}. ${name}`).join('\n');
    const sectionPromptText = 'Please enter the NUMBER for the section you wish to configure:\n\n' + numberedSections;

    // Show the prompt to the user.
    const sectionResponse = ui.prompt('Select a Section', sectionPromptText, ui.ButtonSet.OK_CANCEL);
    if (sectionResponse.getSelectedButton() !== ui.Button.OK) return null;

    const selection = sectionResponse.getResponseText().trim();
    // Convert the user's text ("1") to a 0-based array index (0).
    const selectedIndex = parseInt(selection, 10) - 1;

    // Validate the user's selection.
    if (isNaN(selectedIndex) || selectedIndex < 0 || selectedIndex >= sectionNames.length) {
        ui.alert('Invalid selection. Please run the function again and enter a valid number from the list.');
        return null;
    }
    const selectedSectionName = sectionNames[selectedIndex];
    const selectedSection = sectionData[selectedSectionName];
    if (!selectedSection) {
        ui.alert(`Section "${selectedSectionName}" not found. Please check the spelling and try again.`);
        return null;
    }

    return selectedSectionName;
}

// ============================================================================
// ONBOARDING SUITE (Feature 10)
// ============================================================================

/**
 * Opens the Help & Tutorial Sidebar.
 */
function showTutorialSidebar() {
    const html = HtmlService.createHtmlOutput(buildTutorialHtml())
        .setTitle('Randomizer Guide')
        .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Generates the "Demo Roster" with valid sample data.
 */
function generateDemoRoster() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName("Demo Roster");
    if (sheet) {
        SpreadsheetApp.getUi().alert('A "Demo Roster" sheet already exists.');
        return;
    }

    sheet = spreadsheet.insertSheet("Demo Roster");

    // Headers
    // Row 1: Section Name
    // Row 2: Room Name
    // Row 3: Headers (Name, Gender, Department, Role)

    const headers = [
        ["SciFi Cadets", "", "", "", "Fantasy Heroes", ""],
        ["Deck 12", "", "", "", "Great Hall", ""],
        ["Name", "Gender", "Dept", "", "Name", "Role"]
    ];

    sheet.getRange(1, 1, 3, 6).setValues(headers).setFontWeight("bold");
    sheet.getRange(1, 1).setBackground("#cfe2f3"); // Color for Section
    sheet.getRange(1, 5).setBackground("#fce5cd");

    // Sample Data (SciFi)
    const sciFiStudents = [
        ["Ellen Ripley", "F", "Command"],
        ["Luke Skywalker", "M", "Pilot"],
        ["Sarah Connor", "F", "Combat"],
        ["Jean-Luc Picard", "M", "Command"],
        ["Dana Scully", "F", "Science"],
        ["Marty McFly", "M", "Civilian"],
        ["Leia Organa", "F", "Command"],
        ["Neo", "M", "Combat"],
        ["Trinity", "F", "Combat"],
        ["Spock", "M", "Science"],
        ["Uhura", "F", "Comms"],
        ["Han Solo", "M", "Pilot"],
        ["Rey", "F", "Jedi"],
        ["Kylo Ren", "M", "Sith"],
        ["Gamora", "F", "Combat"]
    ];

    // Sample Data (Fantasy)
    const fantasyStudents = [
        ["Frodo Baggins", "Hobbit"],
        ["Hermione Granger", "Wizard"],
        ["Jon Snow", "Human"],
        ["Daenerys Targaryen", "Human"],
        ["Gandalf", "Wizard"],
        ["Arya Stark", "Human"],
        ["Harry Potter", "Wizard"],
        ["Katniss Everdeen", "Human"],
        ["Legolas", "Elf"],
        ["Gimli", "Dwarf"]
    ];

    // Write SciFi Data
    const sciFiRows = sciFiStudents.map(s => [s[0], s[1], s[2], ""]);
    sheet.getRange(4, 1, sciFiRows.length, 4).setValues(sciFiRows);

    // Write Fantasy Data
    const fantasyRows = fantasyStudents.map(s => ["", "", "", "", s[0], s[1]]);
    sheet.getRange(4, 1, fantasyRows.length, 6).setValues(fantasyRows);

    // Auto-resize
    sheet.autoResizeColumns(1, 6);

    // Set Default Config for these rooms so demo runs immediately
    const props = PropertiesService.getDocumentProperties();
    const tableConfig = JSON.parse(props.getProperty('tableConfig') || '{}');
    tableConfig["Deck 12"] = 4; // 15 students / 4 = ~4 per table
    tableConfig["Great Hall"] = 3; // 10 students / 3 = ~3 per table
    props.setProperty('tableConfig', JSON.stringify(tableConfig));

    SpreadsheetApp.getUi().alert('Demo Roster Created! Navigate to the "Demo Roster" sheet to try it out.');
}

function buildTutorialHtml() {
    return `
    <style>
        body { font-family: 'Segoe UI', Roboto, sans-serif; font-size: 14px; padding: 15px; color: #333; line-height: 1.5; }
        h3 { margin-top: 20px; color: #1a73e8; display: flex; align-items: center; gap: 8px; border-bottom: 1px solid #eee; padding-bottom: 5px; }
        h3:first-of-type { margin-top: 0; }
        .card { background: #f8f9fa; border: 1px solid #ddd; border-radius: 8px; padding: 15px; margin-bottom: 15px; }
        .step { display: flex; gap: 10px; margin-bottom: 8px; align-items: flex-start; }
        .num { background: #1a73e8; color: white; border-radius: 50%; width: 20px; height: 20px; display: flex; align-items: center; justify-content: center; font-size: 12px; flex-shrink: 0; margin-top: 2px; }
        button { background: #1a73e8; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; width: 100%; font-weight: 500; margin-top: 5px; }
        button.secondary { background: white; border: 1px solid #dadce0; color: #1a73e8; }
        button:hover { opacity: 0.9; }
        details { margin-bottom: 10px; border-bottom: 1px solid #eee; padding-bottom: 10px; }
        summary { font-weight: 600; cursor: pointer; color: #444; outline: none; }
        .section-title { font-weight: 700; color: #5f6368; margin-bottom: 5px; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 15px; }
        .menu-item { margin-bottom: 8px; }
        .menu-name { font-weight: 600; color: #202124; }
        .menu-desc { font-size: 13px; color: #5f6368; margin-top: 2px; }
        code { background: #f1f3f4; padding: 2px 4px; border-radius: 4px; font-family: monospace; font-size: 12px; }
    </style>
    
    <h3>👋 Randomizer Guide</h3>
    <p>Create perfect seating arrangements with powerful rules and controls.</p>

    <div class="card">
        <div style="font-weight:bold; margin-bottom:10px;">🚀 Quick Start</div>
        <div class="step"><div class="num">1</div><div><b>Setup Roster</b>: Use the <i>Roster Setup Wizard</i> to create your sheet.</div></div>
        <div class="step"><div class="num">2</div><div><b>Configure</b>: Set <i>Table Counts</i> for your rooms.</div></div>
        <div class="step"><div class="num">3</div><div><b>Run</b>: Click <i>Randomly Assign Students</i>.</div></div>
        
        <button class="secondary" onclick="google.script.run.generateDemoRoster()">🎲 Generate Demo Class</button>
    </div>

    <h3>📖 Menu Reference</h3>
    
    <div class="section-title">Onboarding & Setup</div>
    <div class="menu-item">
        <div class="menu-name">🧙 Roster Setup Wizard</div>
        <div class="menu-desc">Creates a perfectly formatted "Rosters" sheet for you. Just enter your class details and it does the rest.</div>
    </div>
    <div class="menu-item">
        <div class="menu-name">🗺️ Generate Map Template</div>
        <div class="menu-desc">Creates a "Visual Map" sheet. Drag the boxes to match your room layout. The script will fill them with names automatically.</div>
    </div>

    <div class="section-title">Configuration</div>
    <div class="menu-item">
        <div class="menu-name">Configure Tables (Groups)</div>
        <div class="menu-desc"><b>Required.</b> Set how many tables are in each room (e.g. 6 tables).</div>
    </div>
    <div class="menu-item">
        <div class="menu-name">Configure Data Balancing</div>
        <div class="menu-desc">
            Ensures groups are balanced by a specific attribute. <br>
            <b>Gender Logic:</b> The script prioritizes 50/50 splits. If that's not possible, it prefers single-gender tables over "unbalanced" mixed tables (e.g., 3 Boys, 1 Girl).
        </div>
    </div>
    <div class="menu-item">
        <div class="menu-name">Set Room Constraints</div>
        <div class="menu-desc">Set the maximum capacity of specific tables (e.g. Table 1 only has 3 chairs).</div>
    </div>
    <div class="menu-item">
        <div class="menu-name">Select Absent Students</div>
        <div class="menu-desc">Temporarily exclude students from the next randomization without deleting them.</div>
    </div>
    
    <div class="section-title">Student Rules</div>
    <div class="menu-item">
        <div class="menu-name">Preferential Seating</div>
        <div class="menu-desc">Force specific students to be always at Table 1 (e.g. for vision issues).</div>
    </div>
    <div class="menu-item">
        <div class="menu-name">Student Separations</div>
        <div class="menu-desc">Ensure specific students are <b>never</b> placed at the same table.</div>
    </div>
    <div class="menu-item">
        <div class="menu-name">Student Buddies</div>
        <div class="menu-desc">Ensure specific students are placed together. You can set the <b>frequency</b> (e.g. Always, Sometimes) for each group.</div>
    </div>
    
    <div class="section-title">Layout & History</div>
    <div class="menu-item">
        <div class="menu-name">Layout Manager</div>
        <div class="menu-desc">Save your current chart and reload it later. Great for keeping a good setup for a while.</div>
    </div>
    <div class="menu-item">
        <div class="menu-name">View Assignment History</div>
        <div class="menu-desc">See where students sat recently to ensure they are rotating properly.</div>
    </div>

    <div class="section-title">Execution</div>
    <div class="menu-item">
        <div class="menu-name">▶️ Randomly Assign</div>
        <div class="menu-desc">Standard mode. Follows all rules above.</div>
    </div>
    <div class="menu-item">
        <div class="menu-name">🥂 Social Mixer</div>
        <div class="menu-desc"><b>Party Mode!</b> Prioritizes placing students with people they haven't sat with before. Great for day 1 or community building.</div>
    </div>

    <div class="section-title">Tools</div>
    <div class="menu-item">
        <div class="menu-name">Manage & Clear Settings</div>
        <div class="menu-desc">Reset specific configurations or wipe everything to start fresh.</div>
    </div>

    <div style="margin-top:20px; font-size:12px; color:#666; text-align:center; border-top: 1px solid #eee; padding-top: 15px;">
        <p style="margin-bottom:5px;">Developed by <a href="https://knuffke.com/support" target="_blank" style="color:#333; text-decoration:none;"><b>David Knuffke</b></a></p>
        <p style="font-size:10px; margin-top:5px;">Made available under a <a href="http://creativecommons.org/licenses/by-nc-sa/4.0/" target="_blank">CC BY-NC-SA 4.0 License</a>.</p>
        <a href="#" onclick="google.script.host.close()">Close Guide</a>
    </div>
    `;
}

// ============================================================================
// ROSTER SETUP WIZARD (Feature 11)
// ============================================================================

/**
 * Shows the dialog to configure and create a new Roster sheet.
 */
function showRosterSetupWizard() {
    const html = `
    <html>
      <head>
        <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">
        <style>
          .branding-below { bottom: 56px; top: 0; }
          .row { margin-bottom: 10px; }
          label { font-weight: bold; display: block; margin-bottom: 4px; }
          input[type="text"], input[type="number"] { width: 100%; box-sizing: border-box; }
          .section-group { border: 1px solid #ccc; padding: 10px; margin-bottom: 10px; border-radius: 4px; background: #f9f9f9; }
          .col-chip { display: inline-block; background: #e8eaed; padding: 4px 8px; border-radius: 12px; margin-right: 4px; font-size: 11px; }
          .hint { color: #666; font-size: 11px; margin-top: 2px; }
        </style>
      </head>
      <body>
        <form id="wizardForm">
          <div class="row">
            <label>How many classes (sections) do you teach?</label>
            <input type="number" id="numSections" min="1" max="10" value="1" onchange="updateSectionInputs()">
          </div>
          
          <div id="sectionsContainer"></div>
          
          <div class="row">
            <label>Extra Columns (Attributes)</label>
             <div class="hint">Check attributes to include for every student:</div>
             <div style="margin-top:5px;">
               <input type="checkbox" name="trackGender" id="trackGender" checked> <label for="trackGender" style="display:inline; font-weight:normal;">Gender</label>
             </div>
             <div>
               <input type="checkbox" name="trackGrade" id="trackGrade"> <label for="trackGrade" style="display:inline; font-weight:normal;">Grade Level</label>
             </div>
             <div>
               <input type="checkbox" name="trackDesignator" id="trackDesignator"> <label for="trackDesignator" style="display:inline; font-weight:normal;">Designator (e.g. Skill, Role)</label>
             </div>
             <div style="margin-top:5px;">
               <input type="text" id="customCols" placeholder="Other (comma separated, e.g. House)">
             </div>
          </div>

          <div class="bottom">
             <button class="action" onclick="submitForm()">Create Roster Sheet</button>
             <button onclick="google.script.host.close()">Cancel</button>
          </div>
        </form>

        <script>
          function updateSectionInputs() {
             const count = document.getElementById('numSections').value;
             const container = document.getElementById('sectionsContainer');
             container.innerHTML = '';
             
             for (let i = 0; i < count; i++) {
                const div = document.createElement('div');
                div.className = 'section-group';
                div.innerHTML = \`
                  <label>Section \${i+1} Name</label>
                  <input type="text" class="secName" placeholder="e.g. Bio 101 Period 2" value="Period \${i+1}">
                  <div style="margin-top:5px;">
                    <label>Room Name</label>
                    <input type="text" class="roomName" placeholder="e.g. Room 304" value="Science Lab">
                  </div>
                \`;
                container.appendChild(div);
             }
          }
          
          // Init
          updateSectionInputs();

          function submitForm() {
             const secNames = Array.from(document.querySelectorAll('.secName')).map(i => i.value);
             const roomNames = Array.from(document.querySelectorAll('.roomName')).map(i => i.value);
             const customCols = document.getElementById('customCols').value;
             
             // Gather checkbox values
             const extras = [];
             if (document.getElementById('trackGender').checked) extras.push("Gender");
             if (document.getElementById('trackGrade').checked) extras.push("Grade");
             if (document.getElementById('trackDesignator').checked) extras.push("Designator");
             
             if (customCols) {
                customCols.split(',').forEach(c => extras.push(c.trim()));
             }

             google.script.run
               .withSuccessHandler(() => google.script.host.close())
               .createCustomRosterSheet({
                  sectionNames: secNames,
                  roomNames: roomNames,
                  extraColumns: extras
               });
          }
        </script>
      </body>
    </html>
    `;

    SpreadsheetApp.getUi().showModelessDialog(HtmlService.createHtmlOutput(html).setHeight(500).setWidth(400), 'Roster Setup Wizard');
}

/**
 * Logic to generate the Roster sheet from Wizard data.
 */
function createCustomRosterSheet(data) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetName = "Rosters";
    if (ss.getSheetByName(sheetName)) {
        sheetName = "Rosters (New)";
    }

    const sheet = ss.insertSheet(sheetName);
    const sections = data.sectionNames;
    const rooms = data.roomNames;
    const extraCols = data.extraColumns; // Array of strings e.g. ["Gender", "Grade"]

    // Each section needs: Name Column + Extra Columns + 1 Buffer Column
    // e.g. if Extras = ["Gender"], section width = 1 (Name) + 1 (Gender) = 2 columns.
    // If we want a blank buffer column between sections? Yes.

    const columnsPerSection = 1 + extraCols.length; // Name + Attributes
    // Buffer column is handled by skipping an index

    // Let's build the header array.
    // Row 1: Section Names
    // Row 2: Room Names
    // Row 3: Column Headers

    // We need to calculate total columns needed.
    // (columnsPerSection + 1 buffer) * numSections

    let currentColumn = 1;

    for (let i = 0; i < sections.length; i++) {
        const secName = sections[i];
        const roomName = rooms[i];

        // Row 1: Section Name (Merged across width)
        sheet.getRange(1, currentColumn, 1, columnsPerSection).merge().setValue(secName)
            .setBackground("#cfe2f3").setFontWeight("bold").setHorizontalAlignment("center");

        // Row 2: Room Name (Merged across width)
        sheet.getRange(2, currentColumn, 1, columnsPerSection).merge().setValue(roomName)
            .setBackground("#efefef").setFontStyle("italic").setHorizontalAlignment("center");

        // Row 3: Headers
        // First col is always "Name"
        sheet.getRange(3, currentColumn).setValue("Name").setFontWeight("bold");

        // Subsequent cols are Extra Attributes
        for (let j = 0; j < extraCols.length; j++) {
            sheet.getRange(3, currentColumn + 1 + j).setValue(extraCols[j]).setFontWeight("bold");
        }

        // Add borders to the section block
        sheet.getRange(3, currentColumn, 100, columnsPerSection).setBorder(null, true, null, true, null, null);

        // Move cursor: Width + 1 buffer column
        currentColumn += columnsPerSection + 1;
    }

    // Frozen rows
    sheet.setFrozenRows(3);

    // Auto resize
    sheet.autoResizeColumns(1, currentColumn);

    SpreadsheetApp.getUi().alert(`Created "${sheetName}"! You can now paste your student lists.`);
}