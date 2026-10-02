const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const DATA_CSV = path.join(__dirname, 'data.csv');
const DATA_XLSX = path.join(__dirname, 'data.xlsx');
const USERS_FILE = path.join(__dirname, 'users.json');
const STATUS_FILE = path.join(__dirname, 'status.json'); 
const TOGGLE_FILE = path.join(__dirname, 'toggle.json'); 

// Displayed columns
const DISPLAY_COLUMNS = [
  'NO','AJL/DMI','DEFECT/TASK','SPARES','DFP','REMARKS',
  'ROBBING DECLARATION','RECEIVING AIRCRAFT','9M-LDJ Compatibility',
  'MATPLAN UPDATE','SPARE EDD','OPTION','BOOK'
];

// --- Utility ---
function getDataFile() {
  if (fs.existsSync(DATA_XLSX)) return DATA_XLSX;
  if (fs.existsSync(DATA_CSV)) return DATA_CSV;
  return null;
}

function loadData() {
  const file = getDataFile();
  if (!file) return { rows: [], file: null };
  try {
    const wb = XLSX.readFile(file);
    const ws = wb.Sheets[wb.SheetNames[0]];
    let json = XLSX.utils.sheet_to_json(ws, { defval: '' });

    const rows = json.map((r, i) => ({ __id: i, ...r }));
    return { rows, file };
  } catch (err) {
    console.error('Failed to read data file:', err);
    return { rows: [], file: null };
  }
}

function saveData(json, file) {
  const newWs = XLSX.utils.json_to_sheet(json);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, newWs, "Sheet1");
  XLSX.writeFile(wb, file);
}

function unique(arr) {
  return [...new Set(arr.filter(v => v !== '' && v != null))].sort();
}

// --- Status persistence ---
function loadStatus() {
  if (!fs.existsSync(STATUS_FILE)) return {};
  try {
    return JSON.parse(fs.readFileSync(STATUS_FILE, 'utf8'));
  } catch {
    return {};
  }
}

function saveStatus(statuses) {
  fs.writeFileSync(STATUS_FILE, JSON.stringify(statuses, null, 2), 'utf8');
}

// --- Toggle persistence ---
function loadToggles() {
  if (!fs.existsSync(TOGGLE_FILE)) return {};
  try {
    return JSON.parse(fs.readFileSync(TOGGLE_FILE, 'utf8'));
  } catch {
    return {};
  }
}

function saveToggles(toggles) {
  fs.writeFileSync(TOGGLE_FILE, JSON.stringify(toggles, null, 2), 'utf8');
}

// --- Helpers ---
function summarizeStatuses(rows, statuses) {
  const filtered = rows.filter(r => String(r['Aircraft']).trim() === '9M-LNR');
  const groups = {};
  filtered.forEach(r => {
    const ajl = r['AJL/DMI'] || '';
    if (!groups[ajl]) groups[ajl] = { ajl, status: statuses[ajl] || 'OPEN' };
  });

  let open = 0, closed = 0;
  Object.values(groups).forEach(g => {
    if (g.status === 'CLOSED') closed++;
    else open++;
  });
  return { open, closed };
}

// --- API Endpoints ---
app.get('/api/health', (_, res) => res.json({ ok: true }));

app.get('/api/meta', (req, res) => {
  const { rows } = loadData();
  const aircrafts = unique(rows.map(r => String(r['Aircraft'] || '').trim()));
  const systems = unique(rows.map(r => String(r['System'] || '').trim()));
  res.json({ aircrafts, systems });
});

app.post('/api/login', (req, res) => {
  const { id, password } = req.body || {};
  try {
    if (!fs.existsSync(USERS_FILE)) {
      return res.status(500).json({ success: false, message: 'users.json missing' });
    }
    const users = JSON.parse(fs.readFileSync(USERS_FILE, 'utf8'));
    const ok = users.some(u => u.id === id && u.password === password);
    if (ok) return res.json({ success: true });
    return res.status(401).json({ success: false, message: 'Invalid credentials' });
  } catch (err) {
    console.error('Login error:', err);
    return res.status(500).json({ success: false, message: 'Server error' });
  }
});

// --- Search ---
app.post('/api/search', (req, res) => {
  const { aircraft, system } = req.body || {};
  const { rows } = loadData();
  const statuses = loadStatus();
  const toggles = loadToggles();

  let results = rows;
  if (aircraft) {
    results = results.filter(r => String(r['Aircraft'] || '').trim() === String(aircraft).trim());
  }
  if (system) {
    results = results.filter(r => String(r['System'] || '').trim() === String(system).trim());
  }

  let lastNo="", lastAjl="", lastDefect="", lastBook="";
  const out = results.map(r => {
    const currentNo = r['NO'] || lastNo;
    const currentAjl = r['AJL/DMI'] || lastAjl;
    const currentDefect = r['DEFECT/TASK'] || lastDefect;
    const currentBook = r['BOOK'] || lastBook;

    lastNo = currentNo;
    lastAjl = currentAjl;
    lastDefect = currentDefect;
    lastBook = currentBook;

    const o = { __id: r.__id };
    DISPLAY_COLUMNS.forEach(c => {
      if (c === 'NO') o[c] = currentNo;
      else if (c === 'AJL/DMI') o[c] = currentAjl;
      else if (c === 'DEFECT/TASK') o[c] = currentDefect;
      else if (c === 'BOOK') o[c] = currentBook;
      else o[c] = r[c] ?? '';
    });
    o['System'] = r['System'] ?? '';
    o['Aircraft'] = r['Aircraft'] ?? '';
    o['Status'] = statuses[currentAjl] || 'OPEN';
    o['Toggle'] = toggles[r.__id] || 'RED'; 
    return o;
  });

  res.json({ columns: DISPLAY_COLUMNS, count: out.length, rows: out });
});

// --- Update AJL status ---
app.post('/api/update-status', (req, res) => {
  const { ajl, status } = req.body || {};
  if (!ajl) return res.status(400).json({ success: false, message: 'Missing AJL' });

  const statuses = loadStatus();
  statuses[ajl] = status;
  saveStatus(statuses);

  const { rows } = loadData();
  const summary = summarizeStatuses(rows, statuses);
  res.json({ success: true, summary });
});

// --- Status summary for pie chart ---
app.get('/api/status-summary', (req, res) => {
  const { rows } = loadData();
  const statuses = loadStatus();
  const summary = summarizeStatuses(rows, statuses);
  res.json(summary);
});

// --- Update a row ---
app.post('/api/update-row', (req, res) => {
  const { id, updates } = req.body || {};
  if (id == null || !updates) {
    return res.status(400).json({ success: false, message: 'Missing id or updates' });
  }

  const file = getDataFile();
  if (!file) return res.status(500).json({ success: false, message: 'No data file found' });

  try {
    const wb = XLSX.readFile(file);
    const ws = wb.Sheets[wb.SheetNames[0]];
    let json = XLSX.utils.sheet_to_json(ws, { defval: '' });

    if (!json[id]) return res.status(404).json({ success: false, message: 'Row not found' });

    Object.assign(json[id], updates);
    saveData(json, file);

    res.json({ success: true });
  } catch (err) {
    console.error('Update row failed:', err);
    res.status(500).json({ success: false, message: 'Failed to update row' });
  }
});

// --- Add new row ---
app.post('/api/add-row', (req, res) => {
  const { ajl, no, defect, book, updates, aircraft, system } = req.body || {};
  if (!ajl || !no) {
    return res.status(400).json({ success: false, message: 'Missing AJL or NO' });
  }

  const file = getDataFile();
  if (!file) return res.status(500).json({ success: false, message: 'No data file found' });

  try {
    const wb = XLSX.readFile(file);
    const ws = wb.Sheets[wb.SheetNames[0]];
    let json = XLSX.utils.sheet_to_json(ws, { defval: '' });

    // inherit Aircraft/System if not provided
    let ajlAircraft = aircraft || "";
    let ajlSystem = system || "";
    for (let i = json.length - 1; i >= 0; i--) {
      if (String(json[i]['AJL/DMI'] || '').trim() === String(ajl).trim()) {
        ajlAircraft = ajlAircraft || json[i]['Aircraft'] || "";
        ajlSystem = ajlSystem || json[i]['System'] || "";
        break;
      }
    }

    const newRow = {
      "NO": no,
      "AJL/DMI": ajl,
      "DEFECT/TASK": defect,
      "BOOK": book,
      "Aircraft": ajlAircraft,
      "System": ajlSystem,
      ...updates
    };

    // insert at bottom of AJL group
    let insertIndex = json.length;
    for (let i = json.length - 1; i >= 0; i--) {
      if (String(json[i]['AJL/DMI'] || '').trim() === String(ajl).trim()) {
        insertIndex = i + 1;
        break;
      }
    }
    json.splice(insertIndex, 0, newRow);

    saveData(json, file);

    const statuses = loadStatus();
    const toggles = loadToggles();

    toggles[insertIndex] = 'RED';
    saveToggles(toggles);

    const rows = json.map((r, i) => {
      const o = { __id: i };
      DISPLAY_COLUMNS.forEach(c => o[c] = r[c] ?? '');
      o['System'] = r['System'] ?? '';
      o['Aircraft'] = r['Aircraft'] ?? '';
      o['Status'] = statuses[r['AJL/DMI']] || 'OPEN';
      o['Toggle'] = toggles[i] || 'RED';
      return o;
    });

    let results = rows;
    if (aircraft) results = results.filter(r => String(r['Aircraft']).trim() === String(aircraft).trim());
    if (system) results = results.filter(r => String(r['System']).trim() === String(system).trim());

    res.json({ success: true, columns: DISPLAY_COLUMNS, count: results.length, rows: results });
  } catch (err) {
    console.error('Add row failed:', err);
    res.status(500).json({ success: false, message: 'Failed to add row' });
  }
});

// --- Remove row ---
app.post('/api/remove-row', (req, res) => {
  const { id } = req.body || {};
  if (id == null) {
    return res.status(400).json({ success: false, message: 'Missing row id' });
  }

  const file = getDataFile();
  if (!file) return res.status(500).json({ success: false, message: 'No data file found' });

  try {
    const wb = XLSX.readFile(file);
    const ws = wb.Sheets[wb.SheetNames[0]];
    let json = XLSX.utils.sheet_to_json(ws, { defval: '' });

    if (!json[id]) {
      return res.status(404).json({ success: false, message: 'Row not found' });
    }

    json.splice(id, 1);
    saveData(json, file);

    res.json({ success: true });
  } catch (err) {
    console.error('Remove row failed:', err);
    res.status(500).json({ success: false, message: 'Failed to remove row' });
  }
});

// --- Toggle update ---
app.post('/api/update-toggle', (req, res) => {
  const { id, state } = req.body || {};
  if (id == null || !state) {
    return res.status(400).json({ success: false, message: 'Missing id or state' });
  }

  const toggles = loadToggles();
  toggles[id] = state;
  saveToggles(toggles);

  res.json({ success: true });
});

app.listen(PORT, () => {
  console.log(`Rob Tracker server running on http://localhost:${PORT}`);
});
