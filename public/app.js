let isProcessing = false;
const STORAGE_KEY = 'ai-sheets-data';

// ---- INIT ----
function loadData() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) return JSON.parse(saved);
  } catch (e) { console.warn('Could not load saved data'); }
  return null;
}

function saveToLocal() {
  try {
    const data = luckysheet.getSheetData();
    const sheets = luckysheet.getAllSheets();
    const payload = { sheets, timestamp: Date.now() };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    addChat('Saved!', 'success');
  } catch (e) {
    addChat('Save failed: ' + e.message, 'error');
  }
}

function autoSave() {
  try {
    const sheets = luckysheet.getAllSheets();
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ sheets, timestamp: Date.now() }));
  } catch (e) {}
}

// ---- INIT LUCKYSHEET ----
const saved = loadData();
const sheetData = saved && saved.sheets && saved.sheets.length ? saved.sheets : [
  { name: 'Sheet1', celldata: [], row: 100, column: 26, config: {} }
];

luckysheet.create({
  container: 'luckysheet',
  showinfobar: false,
  data: sheetData,
  allowEdit: true,
  showtoolbar: true,
  showstatisticBar: true,
  sheetFormulaBar: true,
  enableAddRow: true,
  enableAddCol: true,
  hook: {
    cellUpdated() { autoSave(); },
    sheetUpdated() { autoSave(); }
  }
});

// ---- CHAT ----
const chatArea = document.getElementById('chatArea');
const welcome = document.getElementById('welcome');

function addChat(text, type) {
  if (welcome) welcome.style.display = 'none';
  const div = document.createElement('div');
  div.className = `msg ${type}`;
  div.textContent = text;
  chatArea.appendChild(div);
  chatArea.scrollTop = chatArea.scrollHeight;
  return div;
}

function showLoading() {
  const div = document.createElement('div');
  div.className = 'msg ai';
  div.id = 'loadingMsg';
  div.innerHTML = '<span class="loading-dot"></span> <span class="loading-dot"></span> <span class="loading-dot"></span> Thinking...';
  chatArea.appendChild(div);
  chatArea.scrollTop = chatArea.scrollHeight;
}

function removeLoading() {
  const el = document.getElementById('loadingMsg');
  if (el) el.remove();
}

// ---- INPUT ----
const aiInput = document.getElementById('aiInput');
const sendBtn = document.getElementById('sendBtn');

aiInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !isProcessing) sendCommand();
});

function quickSend(btn) {
  aiInput.value = btn.textContent;
  sendCommand();
}

async function sendCommand() {
  const message = aiInput.value.trim();
  if (!message || isProcessing) return;

  isProcessing = true;
  aiInput.value = '';
  sendBtn.disabled = true;
  sendBtn.style.opacity = '0.5';

  addChat(message, 'user');
  showLoading();

  try {
    const res = await fetch('/api/ai', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ message, sheetData: getSheetContext() })
    });

    const data = await res.json();
    removeLoading();

    if (!data.success) {
      addChat(data.error || 'Something went wrong', 'error');
    } else {
      const count = data.actions.length;
      addChat(`Executing ${count} action${count > 1 ? 's' : ''}...`, 'ai');
      const results = executeActions(data.actions);
      const failed = results.filter(r => r.status === 'error');
      if (failed.length === 0) {
        addChat('Done!', 'success');
      } else {
        addChat(`${count - failed.length}/${count} succeeded, ${failed.length} failed`, 'error');
      }
    }
  } catch (err) {
    removeLoading();
    addChat('Error: ' + err.message, 'error');
  }

  isProcessing = false;
  sendBtn.disabled = false;
  sendBtn.style.opacity = '1';
}

function getSheetContext() {
  try {
    const sheet = luckysheet.getSheet();
    const data = luckysheet.getSheetData();
    let maxUsedRow = 0, maxUsedCol = 0;
    if (data) {
      for (let r = 0; r < data.length; r++) {
        for (let c = 0; c < (data[r] || []).length; c++) {
          if (data[r][c] && data[r][c].v != null && data[r][c].v !== '') {
            maxUsedRow = Math.max(maxUsedRow, r);
            maxUsedCol = Math.max(maxUsedCol, c);
          }
        }
      }
    }
    // Read header row
    const headers = [];
    if (data && data[0]) {
      for (let c = 0; c <= maxUsedCol; c++) {
        const cell = data[0][c];
        headers.push(cell && cell.v != null ? String(cell.v) : '');
      }
    }
    return { sheetName: sheet.name, rows: maxUsedRow + 1, cols: maxUsedCol + 1, headers };
  } catch { return null; }
}

// ---- ACTIONS ----
const COLOR_MAP = {
  red: '#ff4444', blue: '#3366cc', green: '#22aa55', yellow: '#ffcc00',
  orange: '#ff8800', purple: '#8833cc', pink: '#ff66aa', black: '#000000',
  white: '#ffffff', gray: '#666666', grey: '#666666',
  lightblue: '#87ceeb', lightgreen: '#90ee90', lightyellow: '#ffffe0',
  lightgray: '#cccccc', lightgrey: '#cccccc', cyan: '#00cccc',
  darkblue: '#1a2744', darkgreen: '#1a4a2a', darkgray: '#333333', darkgrey: '#333333',
  navy: '#1a1a5c', teal: '#006666', maroon: '#800000', olive: '#808000',
  coral: '#ff7f50', salmon: '#fa8072', gold: '#ffd700', silver: '#c0c0c0',
  lavender: '#e6e6fa', ivory: '#fffff0', beige: '#f5f5dc'
};

function resolveColor(c) {
  if (!c) return undefined;
  return COLOR_MAP[c.toLowerCase()] || c;
}

function colToLetter(col) {
  let letter = '', c = col;
  while (c >= 0) {
    letter = String.fromCharCode(65 + (c % 26)) + letter;
    c = Math.floor(c / 26) - 1;
  }
  return letter;
}

function executeActions(actions) {
  const results = [];
  for (const a of actions) {
    try {
      executeAction(a);
      results.push({ action: a.action, status: 'ok' });
    } catch (err) {
      console.error('Action failed:', a, err);
      results.push({ action: a.action, status: 'error', error: err.message });
      addChat(`Failed: ${a.action} - ${err.message}`, 'error');
    }
  }
  try { luckysheet.refresh(); } catch(e) {}
  autoSave();
  return results;
}

function ensureCell(row, col) {
  let cell = luckysheet.getCellValue(row, col);
  if (cell === null || cell === undefined) {
    luckysheet.setCellValue(row, col, '');
  }
}

function executeAction(a) {
  const data = luckysheet.getSheetData();
  const maxCol = data && data[0] ? data[0].length - 1 : 25;
  const maxRow = data ? data.length - 1 : 49;

  const endCol = a.endCol === -1 ? maxCol : a.endCol;
  const endRow = a.endRow === -1 ? maxRow : a.endRow;
  const addColIndex = a.index === -1 ? (maxCol + 1) : a.index;

  switch (a.action) {
    case 'setCell':
      luckysheet.setCellValue(a.row, a.col, a.value);
      break;

    case 'setCells':
      for (let r = 0; r < a.values.length; r++)
        for (let c = 0; c < a.values[r].length; c++)
          luckysheet.setCellValue(a.startRow + r, a.startCol + c, a.values[r][c]);
      break;

    case 'addColumn':
      luckysheet.insertColumn(addColIndex);
      if (a.name) luckysheet.setCellValue(0, addColIndex, a.name);
      break;

    case 'addRow':
      luckysheet.insertRow(a.index);
      if (a.values) for (let c = 0; c < a.values.length; c++) luckysheet.setCellValue(a.index, c, a.values[c]);
      break;

    case 'deleteColumn':
      luckysheet.deleteColumn(a.index, a.index);
      break;

    case 'deleteRow':
      luckysheet.deleteRow(a.index, a.index);
      break;

    case 'setCellStyle':
      applyStyle(a.row, a.col, a.row, a.col, a.style);
      break;

    case 'setRangeStyle':
      applyStyle(a.startRow, a.startCol, endRow, endCol, a.style);
      break;

    case 'mergeCells':
      luckysheet.mergeCells(
        `${colToLetter(a.startCol)}${a.startRow + 1}:${colToLetter(endCol)}${endRow + 1}`, 'merge'
      );
      break;

    case 'setColumnWidth':
    case 'setRowHeight': {
      const cfg = luckysheet.getConfig();
      if (a.action === 'setColumnWidth') {
        cfg.columnlen = cfg.columnlen || {};
        cfg.columnlen[a.col] = a.width;
      } else {
        cfg.rowlen = cfg.rowlen || {};
        cfg.rowlen[a.row] = a.height;
      }
      luckysheet.setConfig(cfg);
      break;
    }

    case 'setFormula':
      luckysheet.setCellValue(a.row, a.col, a.formula);
      break;

    case 'freezeRow':
    case 'freezeCol':
      addChat('Use toolbar: View > Freeze', 'ai');
      break;

    case 'addSheet':
      luckysheet.addSheet({ name: a.name || 'Sheet' + (luckysheet.getAllSheets().length + 1) });
      break;

    case 'renameSheet': {
      const sheets = luckysheet.getAllSheets();
      const target = sheets.find(s => s.name === a.oldName);
      if (target) luckysheet.setSheetName(target.index, a.newName);
      break;
    }

    case 'clearRange':
      for (let r = a.startRow; r <= endRow; r++)
        for (let c = a.startCol; c <= endCol; c++)
          luckysheet.setCellValue(r, c, '');
      break;

    case 'setTextAlign': {
      const ht = a.align === 'center' ? 0 : a.align === 'right' ? 2 : 1;
      const rE = a.rowEnd || a.row, cE = a.colEnd || a.col;
      for (let r = a.row; r <= rE; r++)
        for (let c = a.col; c <= cE; c++)
          luckysheet.setCellFormat(r, c, 'ht', ht);
      break;
    }

    case 'setBorder': {
      const bd = { borderType: 'border-all', style: 1, color: '#000000' };
      for (let r = a.startRow; r <= endRow; r++)
        for (let c = a.startCol; c <= endCol; c++)
          luckysheet.setCellFormat(r, c, 'bd', bd);
      break;
    }

    case 'fillRandom': {
      const min = a.min || 100, max = a.max || 9999, count = a.count || 10, type = a.type || 'numbers';
      const fn = ['James','Mary','John','Patricia','Robert','Jennifer','Michael','Linda','David','Elizabeth','William','Barbara','Richard','Susan','Joseph','Jessica','Thomas','Sarah','Charles','Karen','Emma','Oliver','Ava','Liam','Sophia','Noah','Isabella','Ethan','Mia','Lucas'];
      const ln = ['Smith','Johnson','Williams','Brown','Jones','Garcia','Miller','Davis','Rodriguez','Martinez','Anderson','Taylor','Thomas','Moore','Jackson','Martin','Lee','Thompson','White','Harris'];
      for (let i = 0; i < count; i++) {
        let val;
        switch (type) {
          case 'amounts': val = parseFloat((Math.random() * (max - min) + min).toFixed(2)); break;
          case 'names': val = fn[Math.floor(Math.random()*fn.length)] + ' ' + ln[Math.floor(Math.random()*ln.length)]; break;
          case 'emails': val = fn[Math.floor(Math.random()*fn.length)].toLowerCase()+'.'+ln[Math.floor(Math.random()*ln.length)].toLowerCase()+'@gmail.com'; break;
          default: val = Math.floor(Math.random()*(max-min+1))+min;
        }
        luckysheet.setCellValue(a.startRow + i, a.col, val);
      }
      break;
    }

    default:
      console.warn('Unknown action:', a.action);
  }
}

function applyStyle(startRow, startCol, endRow, endCol, s) {
  const bg = s.bg ? resolveColor(s.bg) : null;
  const fc = s.color ? resolveColor(s.color) : null;
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      ensureCell(r, c);
      if (bg) luckysheet.setCellFormat(r, c, 'bg', bg);
      if (fc) luckysheet.setCellFormat(r, c, 'fc', fc);
      if (s.bold) luckysheet.setCellFormat(r, c, 'bl', 1);
      if (s.italic) luckysheet.setCellFormat(r, c, 'it', 1);
      if (s.fontSize) luckysheet.setCellFormat(r, c, 'fs', s.fontSize);
      if (s.underline) luckysheet.setCellFormat(r, c, 'ul', { type: 1, color: fc || '#000000' });
      if (s.strike) luckysheet.setCellFormat(r, c, 'cl', 1);
      if (s.fontFamily) luckysheet.setCellFormat(r, c, 'ff', s.fontFamily);
      if (s.align) luckysheet.setCellFormat(r, c, 'ht', s.align === 'center' ? 0 : s.align === 'right' ? 2 : 1);
    }
  }
}

// ---- DOWNLOAD XLSX ----
function downloadXLSX() {
  try {
    const sheets = luckysheet.getAllSheets();
    const wb = XLSX.utils.book_new();

    for (const sheet of sheets) {
      const data = luckysheet.getSheetData({ order: sheet.order || 0 }) || [];
      const aoa = [];
      for (const row of data) {
        const rowData = [];
        for (const cell of (row || [])) {
          rowData.push(cell && cell.v != null ? cell.v : '');
        }
        if (rowData.some(v => v !== '')) aoa.push(rowData);
      }
      if (aoa.length === 0) aoa.push(['']);
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      XLSX.utils.book_append_sheet(wb, ws, sheet.name || 'Sheet');
    }

    XLSX.writeFile(wb, 'ai-sheets-export.xlsx');
    addChat('Downloaded as ai-sheets-export.xlsx', 'success');
  } catch (e) {
    addChat('Download failed: ' + e.message, 'error');
  }
}

// ---- CLEAR ----
function clearAll() {
  if (!confirm('Clear all data?')) return;
  localStorage.removeItem(STORAGE_KEY);
  location.reload();
}
