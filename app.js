// TPMS File Generator - Application Logic

// State
let processedFiles = [];
let deferredPrompt = null;

// Tab Navigation
const tabConvert = document.getElementById('tabConvert');
const tabGenerate = document.getElementById('tabGenerate');
const viewConvert = document.getElementById('viewConvert');
const viewGenerate = document.getElementById('viewGenerate');

// Generator Form Elements
const genForm = document.getElementById('genForm');
const genSerial = document.getElementById('genSerial');
const genDate = document.getElementById('genDate');
const genTrips = document.getElementById('genTrips');
const serialChips = document.getElementById('serialChips');
const shiftCards = document.querySelectorAll('.shift-card');
const genResultsCard = document.getElementById('genResultsCard');
const genStatSN = document.getElementById('genStatSN');
const genStatShift = document.getElementById('genStatShift');
const genStatTrips = document.getElementById('genStatTrips');
const genStatTonnage = document.getElementById('genStatTonnage');
const genThead = document.getElementById('genThead');
const genTbody = document.getElementById('genTbody');
const btnGenDownloadRaw = document.getElementById('btnGenDownloadRaw') || document.getElementById('btnGenDownloadXlsx');
const btnGenDownloadRawLabel = document.getElementById('btnGenDownloadRawLabel');
const btnGenDownloadXlsx = btnGenDownloadRaw;
const btnGenDownloadCsv = document.getElementById('btnGenDownloadCsv');
const btnGenDownloadZip = document.getElementById('btnGenDownloadZip');
const genTypeBeml = document.getElementById('genTypeBeml');
const genTypeCat = document.getElementById('genTypeCat');
const genSerialLabel = document.getElementById('genSerialLabel');
const genSubtitle = document.getElementById('genSubtitle');

// Generator State
let selectedShift = '1';
let currentGenerated = null;
let genDumperType = 'BEML';
const BEML_GEN_SERIALS = ['60610', '60611', '60643', '60644', '60645', '60646', '60647', '60648', '60904', '60905', '60906', '60907', '60926'];
const CAT_GEN_SERIALS = ['PRB00914', 'PRB00937', 'PRB00968', 'PRB01106', 'PRB01107', 'PRB01108', 'PRB01165', 'PRB01166', 'PRB01167', 'PRB01169', 'PRB01172', 'PRB01173', 'PRB01174'];

// Dumper Type State (BEML vs CAT for conversion view)
let currentDumperType = 'BEML';

// DOM Elements
const typeCardBeml = document.getElementById('typeCardBeml');
const typeCardCat = document.getElementById('typeCardCat');
const monthCard = document.getElementById('monthCard');
const monthSelect = document.getElementById('monthSelect');
const customMonthWrapper = document.getElementById('customMonthWrapper');
const customMonthInput = document.getElementById('customMonthInput');
const quickMonthChips = document.getElementById('quickMonthChips');
const dropZone = document.getElementById('dropZone');
const dropTitle = document.getElementById('dropTitle');
const dropDesc = document.getElementById('dropDesc');
const fileInput = document.getElementById('fileInput');
const folderInput = document.getElementById('folderInput');
const btnSelectFiles = document.getElementById('btnSelectFiles');
const btnSelectFolder = document.getElementById('btnSelectFolder');
const resultsCard = document.getElementById('resultsCard');
const fileList = document.getElementById('fileList');
const btnDownloadConsolidated = document.getElementById('btnDownloadConsolidated');
const btnDownloadAll = document.getElementById('btnDownloadAll');
const btnClearAll = document.getElementById('btnClearAll');
const btnInstall = document.getElementById('btnInstall');

// Stats
const statFiles = document.getElementById('statFiles');
const statTrips = document.getElementById('statTrips');
const statTonnage = document.getElementById('statTonnage');
const statAvgPayload = document.getElementById('statAvgPayload');

// Modal Elements
const previewModal = document.getElementById('previewModal');
const modalTitle = document.getElementById('modalTitle');
const modalMeta = document.getElementById('modalMeta');
const previewThead = document.getElementById('previewThead');
const previewTbody = document.getElementById('previewTbody');
const btnCloseModal = document.getElementById('btnCloseModal');
const toast = document.getElementById('toast');

// BEML 18-Column Headers
const BEML_HEADERS = [
  'TRUCK NO',
  'T_Date',
  'Time',
  'PayLoad',
  'Loading Time',
  'Hauling Time',
  'Stop Dump Time',
  'Return Time',
  'Return Stop Time',
  'Total Time',
  'Haul Travel Distance',
  'Return Travel Distance',
  'Total Distance',
  'MaxHaul Speed',
  'AvgHaul Speed',
  'MaxReturn Speed',
  'AvgReturn Speed',
  'Uploaded By'
];
const CSV_HEADERS = BEML_HEADERS;

// CAT 16-Column Headers
const CAT_HEADERS = [
  'TIMEDATE',
  'PAYLOAD SMH',
  'SERIAL NUMBER',
  'PAYLOADTONS',
  'TRAVEL EMPTY TIME HHMMSS',
  'TRAVEL EMPTY DISTANCEMI',
  'STOPPED EMPTY TIME HHMMSS',
  'LOAD TIME HHMMSS',
  'STOPPED LOADED TIME HHMMSS',
  'LOADED TRAVEL TIME HHMMSS',
  'LOADED TRAVEL DISTANCEMI',
  'CYCLE TIME HHMMSS',
  'CYCLE DISTANCEMI',
  'LOADER PASSES',
  'FUEL USED GAL',
  'UPLOADED BY'
];

// Helper to unpad hour in CAT MineStar time values: 00:29:55 -> 0:29:55
function transformCatTime(val) {
  if (!val) return '';
  const parts = String(val).trim().split(':');
  if (parts.length === 3) {
    const h = parseInt(parts[0], 10);
    return `${isNaN(h) ? parts[0] : h}:${parts[1]}:${parts[2]}`;
  }
  return String(val).trim();
}

// Switch Dumper Type between BEML and CAT
function setDumperType(type) {
  currentDumperType = type;
  if (type === 'BEML') {
    if (typeCardBeml) typeCardBeml.classList.add('active', 'beml');
    if (typeCardCat) typeCardCat.classList.remove('active', 'cat');
    if (dropTitle) dropTitle.textContent = 'Drag & Drop BEML .xls or .xlsx Files or Folders Here';
    if (dropDesc) dropDesc.textContent = 'Accepts BEML TPMS reports -> outputs 18-column <serial>_<month>.csv';
    showToast('Switched to BEML Dumper mode');
  } else {
    if (typeCardCat) typeCardCat.classList.add('active', 'cat');
    if (typeCardBeml) typeCardBeml.classList.remove('active', 'beml');
    if (dropTitle) dropTitle.textContent = 'Drag & Drop CAT MineStar .csv, .xlsx, or .xls Files Here';
    if (dropDesc) dropDesc.textContent = 'Accepts CAT MineStar reports -> outputs 16-column <Equipment>.csv';
    showToast('Switched to CAT Dumper mode');
  }
  refreshFilenames();
}

// Initialize
document.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
  registerServiceWorker();
});

// Month helper
function getSelectedMonth() {
  const selected = monthSelect.value;
  if (selected === 'custom') {
    const customVal = customMonthInput.value.trim().toLowerCase();
    return customVal || 'output';
  }
  return selected.toLowerCase();
}

function updateMonthSelection(monthVal) {
  if (monthVal !== 'custom') {
    monthSelect.value = monthVal;
    customMonthWrapper.classList.add('hidden');
    // Update active chip
    document.querySelectorAll('.chip').forEach(chip => {
      chip.classList.toggle('active', chip.dataset.month === monthVal);
    });
  } else {
    monthSelect.value = 'custom';
    customMonthWrapper.classList.remove('hidden');
    document.querySelectorAll('.chip').forEach(chip => chip.classList.remove('active'));
    customMonthInput.focus();
  }
  refreshFilenames();
}

function refreshFilenames() {
  const month = getSelectedMonth();
  processedFiles.forEach((item, index) => {
    if (item.dumperType === 'CAT') {
      item.outputFilename = `${item.serialNo}.csv`;
    } else {
      item.outputFilename = `${item.serialNo}_${month}.csv`;
    }
    const cardNameEl = document.getElementById(`filename-${index}`);
    if (cardNameEl) {
      cardNameEl.textContent = item.outputFilename;
    }
  });
  if (btnDownloadConsolidated) {
    btnDownloadConsolidated.title = `Download ${getConsolidatedFilename()}`;
  }
}

// Consolidation Helpers
function getCapitalizedMonth() {
  const m = getSelectedMonth();
  if (!m) return 'Output';
  return m.charAt(0).toUpperCase() + m.slice(1);
}

function getDistinctSerials() {
  const serials = [];
  processedFiles.forEach(f => {
    if (f.status === 'success' && f.serialNo && !serials.includes(f.serialNo)) {
      serials.push(f.serialNo);
    }
  });
  return serials;
}

function getConsolidatedFilename() {
  const capMonth = getCapitalizedMonth();
  const serials = getDistinctSerials();
  const valid = processedFiles.filter(f => f.status === 'success');
  const allCat = valid.length > 0 && valid.every(f => f.dumperType === 'CAT');
  if (serials.length === 0) {
    return allCat ? `CAT_consolidated.csv` : `${capMonth}_consolidated.csv`;
  }
  return `${capMonth}_${serials.join('_')}.csv`;
}

function getConsolidatedCsvContent() {
  const valid = processedFiles.filter(f => f.status === 'success');
  const allCat = valid.length > 0 && valid.every(f => f.dumperType === 'CAT');
  const headers = allCat ? CAT_HEADERS : BEML_HEADERS;
  let content = headers.join(',') + '\n';
  valid.forEach(file => {
    if (file.parsedRows) {
      file.parsedRows.forEach(row => {
        content += row.join(',') + '\n';
      });
    }
  });
  return content;
}

function downloadConsolidatedCsv() {
  const validItems = processedFiles.filter(f => f.status === 'success');
  if (validItems.length === 0) {
    showToast('No converted CSV data available.');
    return;
  }
  const content = getConsolidatedCsvContent();
  const filename = getConsolidatedFilename();
  const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
  triggerDownload(blob, filename);
  showToast(`Downloaded ${filename}`);
}

function setupEventListeners() {
  // Dumper Type Switcher (BEML vs CAT)
  if (typeCardBeml) {
    typeCardBeml.addEventListener('click', () => setDumperType('BEML'));
  }
  if (typeCardCat) {
    typeCardCat.addEventListener('click', () => setDumperType('CAT'));
  }

  // Month change
  monthSelect.addEventListener('change', (e) => {
    updateMonthSelection(e.target.value);
  });

  customMonthInput.addEventListener('input', () => {
    refreshFilenames();
  });

  // Month chips
  quickMonthChips.addEventListener('click', (e) => {
    const chip = e.target.closest('.chip');
    if (chip && chip.dataset.month) {
      updateMonthSelection(chip.dataset.month);
    }
  });

  // Buttons for file selection
  btnSelectFiles.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.click();
  });

  btnSelectFolder.addEventListener('click', (e) => {
    e.stopPropagation();
    folderInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    handleFileList(e.target.files);
    fileInput.value = '';
  });

  folderInput.addEventListener('change', (e) => {
    handleFileList(e.target.files);
    folderInput.value = '';
  });

  // Drop zone events
  dropZone.addEventListener('click', () => {
    fileInput.click();
  });

  ['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add('dragover');
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove('dragover');
    }, false);
  });

  dropZone.addEventListener('drop', async (e) => {
    const dt = e.dataTransfer;
    if (dt.items) {
      const files = await getAllFilesFromDataTransfer(dt.items);
      handleFileList(files);
    } else if (dt.files) {
      handleFileList(dt.files);
    }
  });

  // Actions
  btnClearAll.addEventListener('click', () => {
    processedFiles = [];
    renderQueue();
    showToast('Cleared all files');
  });

  if (btnDownloadConsolidated) {
    btnDownloadConsolidated.addEventListener('click', downloadConsolidatedCsv);
  }
  btnDownloadAll.addEventListener('click', downloadAllAsZip);

  // Modal close
  btnCloseModal.addEventListener('click', () => {
    previewModal.classList.remove('open');
  });

  previewModal.addEventListener('click', (e) => {
    if (e.target === previewModal) {
      previewModal.classList.remove('open');
    }
  });

  // Tab Switching
  tabConvert.addEventListener('click', () => {
    tabConvert.classList.add('active');
    tabGenerate.classList.remove('active');
    viewConvert.classList.remove('hidden');
    viewGenerate.classList.add('hidden');
  });

  tabGenerate.addEventListener('click', () => {
    tabGenerate.classList.add('active');
    tabConvert.classList.remove('active');
    viewGenerate.classList.remove('hidden');
    viewConvert.classList.add('hidden');
    if (!genDate.value) {
      genDate.value = new Date().toISOString().split('T')[0];
    }
  });

  // Generator Dumper Type Selector (BEML vs CAT)
  if (genTypeBeml) {
    genTypeBeml.addEventListener('click', () => setGenDumperType('BEML'));
  }
  if (genTypeCat) {
    genTypeCat.addEventListener('click', () => setGenDumperType('CAT'));
  }

  // Generator Serial Chips
  serialChips.addEventListener('click', (e) => {
    const chip = e.target.closest('.chip');
    if (chip && chip.dataset.sn) {
      genSerial.value = chip.dataset.sn;
      serialChips.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
    }
  });

  // Shift Card Selection
  shiftCards.forEach(card => {
    card.addEventListener('click', () => {
      shiftCards.forEach(c => c.classList.remove('active'));
      card.classList.add('active');
      selectedShift = card.dataset.shift;
    });
  });

  // Generator Form Submit
  genForm.addEventListener('submit', (e) => {
    e.preventDefault();
    handleGenerateReport();
  });

  // Generator Downloads
  if (btnGenDownloadRaw) {
    btnGenDownloadRaw.addEventListener('click', () => {
      if (currentGenerated && currentGenerated.rawBlob) {
        triggerDownload(currentGenerated.rawBlob, currentGenerated.rawFilename);
        showToast(`Downloaded ${currentGenerated.rawFilename}`);
      }
    });
  }

  btnGenDownloadCsv.addEventListener('click', () => {
    if (currentGenerated && currentGenerated.csvBlob) {
      triggerDownload(currentGenerated.csvBlob, currentGenerated.csvFilename);
      showToast(`Downloaded ${currentGenerated.csvFilename}`);
    }
  });

  btnGenDownloadZip.addEventListener('click', downloadGeneratedZip);

  // Set default generator date
  if (genDate && !genDate.value) {
    genDate.value = new Date().toISOString().split('T')[0];
  }

  // PWA install prompt
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    btnInstall.style.display = 'inline-flex';
  });

  btnInstall.addEventListener('click', async () => {
    if (deferredPrompt) {
      deferredPrompt.prompt();
      const { outcome } = await deferredPrompt.userChoice;
      if (outcome === 'accepted') {
        btnInstall.style.display = 'none';
      }
      deferredPrompt = null;
    }
  });
}

// Recursive directory scanning for drag-and-drop
async function getAllFilesFromDataTransfer(dataTransferItems) {
  const files = [];
  const entries = [];
  for (let i = 0; i < dataTransferItems.length; i++) {
    const item = dataTransferItems[i];
    if (item.kind === 'file') {
      const entry = item.webkitGetAsEntry ? item.webkitGetAsEntry() : null;
      if (entry) {
        entries.push(entry);
      } else {
        const file = item.getAsFile();
        if (file) files.push(file);
      }
    }
  }

  for (const entry of entries) {
    await scanEntry(entry, files);
  }
  return files;
}

async function scanEntry(entry, fileListOut) {
  if (entry.isFile) {
    const file = await new Promise((resolve) => entry.file(resolve));
    const ext = file.name.toLowerCase();
    if (ext.endsWith('.xls') || ext.endsWith('.xlsx') || ext.endsWith('.csv')) {
      fileListOut.push(file);
    }
  } else if (entry.isDirectory) {
    const dirReader = entry.createReader();
    const entries = await readAllDirectoryEntries(dirReader);
    for (const subEntry of entries) {
      await scanEntry(subEntry, fileListOut);
    }
  }
}

async function readAllDirectoryEntries(dirReader) {
  const entries = [];
  let readEntries = await new Promise((resolve) => dirReader.readEntries(resolve));
  while (readEntries.length > 0) {
    entries.push(...readEntries);
    readEntries = await new Promise((resolve) => dirReader.readEntries(resolve));
  }
  return entries;
}

// Handle file processing queue
async function handleFileList(rawFiles) {
  const validFiles = Array.from(rawFiles).filter(f => {
    const ext = f.name.toLowerCase();
    return ext.endsWith('.xls') || ext.endsWith('.xlsx') || ext.endsWith('.csv');
  });

  if (validFiles.length === 0) {
    showToast('No .xls, .xlsx, or .csv files found.');
    return;
  }

  showToast(`Processing ${validFiles.length} file(s)...`);
  resultsCard.classList.remove('hidden');

  for (const file of validFiles) {
    await processUploadedFile(file);
  }

  renderQueue();
  showToast(`Finished processing ${validFiles.length} file(s)!`);
}

// Parse individual uploaded file (BEML or CAT)
function processUploadedFile(file) {
  return new Promise((resolve) => {
    const isCsv = file.name.toLowerCase().endsWith('.csv');
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        let workbook;
        let rawRows;

        if (isCsv) {
          // Read CSV with raw text and without date auto-parsing so timestamps stay untouched
          const csvText = e.target.result;
          workbook = XLSX.read(csvText, { type: 'string', raw: true, cellDates: false });
          const worksheet = workbook.Sheets[workbook.SheetNames[0]];
          rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });
        } else {
          // Read Excel workbook (.xls / .xlsx)
          const data = new Uint8Array(e.target.result);
          workbook = XLSX.read(data, { type: 'array', cellDates: true });
          let targetSheetName = workbook.SheetNames.find(s => s.toLowerCase() === 'sheet2') || workbook.SheetNames[0];
          const worksheet = workbook.Sheets[targetSheetName];
          rawRows = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            raw: false,
            dateNF: 'dd/mm/yyyy'
          });
        }

        // Determine if file is CAT or BEML
        let isCat = (currentDumperType === 'CAT');
        // Check content markers to auto-detect if user dropped CAT files while BEML is selected, or vice versa
        for (let r = 0; r < Math.min(10, rawRows.length); r++) {
          const rowStr = (rawRows[r] || []).join(' ').toLowerCase();
          if (rowStr.includes('minestar') || rowStr.includes('equipment :')) {
            isCat = true;
            break;
          }
        }
        if (file.name.toLowerCase().startsWith('truckpayload_')) {
          isCat = true;
        }

        let fileRecord;
        if (isCat) {
          fileRecord = parseCatRecord(file, rawRows);
        } else {
          fileRecord = parseBemlRecord(file, rawRows);
        }

        const existingIdx = processedFiles.findIndex(f => f.originalName === file.name);
        if (existingIdx >= 0) {
          processedFiles[existingIdx] = fileRecord;
        } else {
          processedFiles.push(fileRecord);
        }

        resolve();
      } catch (err) {
        console.error('Error parsing file:', file.name, err);
        processedFiles.push({
          originalName: file.name,
          outputFilename: `${file.name.split('.')[0]}_error.csv`,
          serialNo: 'N/A',
          dumperType: currentDumperType,
          headers: currentDumperType === 'CAT' ? CAT_HEADERS : BEML_HEADERS,
          rowCount: 0,
          totalPayload: 0,
          csvContent: '',
          parsedRows: [],
          status: 'error',
          error: err.message
        });
        resolve();
      }
    };

    reader.onerror = () => {
      processedFiles.push({
        originalName: file.name,
        outputFilename: `${file.name.split('.')[0]}_error.csv`,
        serialNo: 'N/A',
        dumperType: currentDumperType,
        headers: currentDumperType === 'CAT' ? CAT_HEADERS : BEML_HEADERS,
        rowCount: 0,
        totalPayload: 0,
        csvContent: '',
        parsedRows: [],
        status: 'error',
        error: 'Failed to read file.'
      });
      resolve();
    };

    if (isCsv) {
      reader.readAsText(file, 'utf-8');
    } else {
      reader.readAsArrayBuffer(file);
    }
  });
}

// Parse CAT MineStar Data Record
function parseCatRecord(file, rawRows) {
  // 1. Extract Serial Number
  let serialNo = '';
  for (let r = 0; r < Math.min(10, rawRows.length); r++) {
    const row = rawRows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] || '');
      if (cell.includes('Equipment :')) {
        serialNo = cell.split('Equipment :')[1].trim();
        break;
      }
    }
    if (serialNo) break;
  }
  if (!serialNo) {
    const m = file.name.match(/TruckPayload_([A-Za-z0-9]+)_/i);
    if (m) serialNo = m[1];
  }
  if (!serialNo) {
    serialNo = file.name.split('.')[0].replace(/[^a-zA-Z0-9]/g, '') || 'CAT_TRUCK';
  }

  // 2. Locate Data Rows (Header starts with 'Date and Time')
  let startRow = -1;
  for (let r = 0; r < Math.min(15, rawRows.length); r++) {
    const row = rawRows[r] || [];
    if (row[0] && String(row[0]).trim().toLowerCase() === 'date and time') {
      startRow = r + 1;
      break;
    }
  }

  const parsedRows = [];
  let totalPayload = 0;

  if (startRow !== -1) {
    for (let r = startRow; r < rawRows.length; r++) {
      const row = rawRows[r];
      if (!row || !row[0] || String(row[0]).trim() === '') continue;

      const pTons = parseFloat(row[2]) || 0;
      totalPayload += pTons;

      // 16 CAT Columns matching Test 2 output
      const formattedRow = [
        String(row[0] || '').trim(),
        String(row[1] || '').trim(),
        serialNo,
        String(row[2] || '').trim(),
        transformCatTime(row[3]),
        String(row[4] || '').trim(),
        transformCatTime(row[5]),
        transformCatTime(row[6]),
        transformCatTime(row[7]),
        transformCatTime(row[8]),
        String(row[9] || '').trim(),
        transformCatTime(row[10]),
        String(row[11] || '').trim(),
        String(row[12] || '').trim(),
        String(row[15] || '').trim(),
        '' // Uploaded By
      ];
      parsedRows.push(formattedRow);
    }
  }

  let csvContent = CAT_HEADERS.join(',') + '\n';
  for (const pRow of parsedRows) {
    csvContent += pRow.join(',') + '\n';
  }

  const outputFilename = `${serialNo}.csv`;

  return {
    originalName: file.name,
    outputFilename,
    serialNo,
    dumperType: 'CAT',
    headers: CAT_HEADERS,
    rowCount: parsedRows.length,
    totalPayload: totalPayload.toFixed(2),
    csvContent,
    parsedRows,
    status: parsedRows.length > 0 ? 'success' : 'warning',
    error: parsedRows.length > 0 ? null : 'No CAT cycle data rows detected.'
  };
}

// Parse BEML Data Record
function parseBemlRecord(file, rawRows) {
  // 1. Extract Serial Number
  let serialNo = '';
  for (let r = 0; r < Math.min(10, rawRows.length); r++) {
    const row = rawRows[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (String(row[c] || '').toLowerCase().includes('serial number')) {
        const val = row[c + 2] || row[c + 1] || '';
        if (val) {
          serialNo = String(val).replace(/\.0+$/, '').trim();
          break;
        }
      }
    }
    if (serialNo) break;
  }

  // Fallback to filename prefix
  if (!serialNo) {
    const namePart = file.name.split('_')[0];
    serialNo = namePart.replace(/[^0-9a-zA-Z]/g, '').trim() || 'unknown';
  }

  // 2. Locate Data Rows
  let startRow = -1;
  for (let r = 0; r < Math.min(25, rawRows.length); r++) {
    if (rawRows[r] && String(rawRows[r][0] || '').trim() === '1') {
      startRow = r;
      break;
    }
  }

  const parsedRows = [];
  let totalPayload = 0;

  if (startRow !== -1) {
    for (let r = startRow; r < rawRows.length; r++) {
      const row = rawRows[r];
      if (!row || row[0] === undefined || row[0] === null || String(row[0]).trim() === '') {
        break;
      }
      if (isNaN(parseInt(row[0]))) break;

      const payloadVal = parseFloat(row[3]) || 0;
      totalPayload += payloadVal;

      const formattedRow = [
        serialNo,
        row[1] || '',
        row[2] || '',
        row[3] || '',
        row[4] || '',
        row[5] || '',
        row[6] || '',
        row[7] || '',
        row[8] || '',
        row[9] || '',
        row[10] || '',
        row[11] || '',
        row[12] || '',
        row[13] || '',
        row[14] || '',
        row[15] || '',
        row[16] || '',
        '' // Uploaded By
      ];
      parsedRows.push(formattedRow);
    }
  }

  const month = getSelectedMonth();
  const outputFilename = `${serialNo}_${month}.csv`;
  let csvContent = BEML_HEADERS.join(',') + '\n';
  for (const pRow of parsedRows) {
    csvContent += pRow.join(',') + '\n';
  }

  return {
    originalName: file.name,
    outputFilename,
    serialNo,
    dumperType: 'BEML',
    headers: BEML_HEADERS,
    rowCount: parsedRows.length,
    totalPayload: totalPayload.toFixed(2),
    csvContent,
    parsedRows,
    status: parsedRows.length > 0 ? 'success' : 'warning',
    error: parsedRows.length > 0 ? null : 'No BEML cycle data rows detected.'
  };
}

// Render the queue cards & update stats
function renderQueue() {
  if (processedFiles.length === 0) {
    resultsCard.classList.add('hidden');
    fileList.innerHTML = '';
    if (btnDownloadConsolidated) {
      btnDownloadConsolidated.classList.add('hidden');
    }
    return;
  }

  resultsCard.classList.remove('hidden');

  let totalTrips = 0;
  let totalTons = 0;
  let validFilesCount = 0;

  fileList.innerHTML = '';

  processedFiles.forEach((file, index) => {
    if (file.status === 'success') {
      validFilesCount++;
      totalTrips += file.rowCount;
      totalTons += parseFloat(file.totalPayload) || 0;
    }

    const card = document.createElement('div');
    card.className = 'file-card';

    const statusBadge = file.status === 'success'
      ? `<span class="badge badge-success">${file.rowCount} Cycles</span>`
      : file.status === 'warning'
      ? `<span class="badge badge-warning">0 Cycles</span>`
      : `<span class="badge badge-error">Error</span>`;

    card.innerHTML = `
      <div class="file-info">
        <div class="file-icon">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
        </div>
        <div class="file-details">
          <div class="file-name" id="filename-${index}">${file.outputFilename}</div>
          <div class="file-meta">
            <span class="badge ${file.dumperType === 'CAT' ? 'badge-cat' : 'badge-beml'}">${file.dumperType || 'BEML'}</span>
            <span class="badge badge-sn">SN: ${file.serialNo}</span>
            <span>Source: ${file.originalName}</span>
            ${statusBadge}
          </div>
        </div>
      </div>
      <div class="file-actions">
        ${file.status === 'success' ? `
          <button type="button" class="btn btn-secondary btn-sm" onclick="previewFile(${index})">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
            Preview
          </button>
          <button type="button" class="btn btn-primary btn-sm" onclick="downloadSingleCsv(${index})">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            Download
          </button>
        ` : `
          <span style="font-size: 0.75rem; color: #f87171;">${file.error || 'Failed'}</span>
        `}
      </div>
    `;

    fileList.appendChild(card);
  });

  // Update Stats
  statFiles.textContent = validFilesCount;
  statTrips.textContent = totalTrips.toLocaleString();
  statTonnage.textContent = totalTons.toLocaleString(undefined, { maximumFractionDigits: 1 });
  statAvgPayload.textContent = totalTrips > 0 ? (totalTons / totalTrips).toFixed(1) : '0';

  // Toggle Consolidated CSV Button
  if (btnDownloadConsolidated) {
    if (validFilesCount > 1) {
      btnDownloadConsolidated.classList.remove('hidden');
      btnDownloadConsolidated.title = `Download ${getConsolidatedFilename()}`;
    } else {
      btnDownloadConsolidated.classList.add('hidden');
    }
  }
}

// Download single CSV
window.downloadSingleCsv = function(index) {
  const item = processedFiles[index];
  if (!item || !item.csvContent) return;

  const blob = new Blob([item.csvContent], { type: 'text/csv;charset=utf-8;' });
  triggerDownload(blob, item.outputFilename);
  showToast(`Downloaded ${item.outputFilename}`);
};

// Preview Modal
window.previewFile = function(index) {
  const item = processedFiles[index];
  if (!item || !item.parsedRows) return;

  const headers = item.headers || (item.dumperType === 'CAT' ? CAT_HEADERS : BEML_HEADERS);
  modalTitle.textContent = `Preview: ${item.outputFilename} (${item.dumperType || 'BEML'})`;
  modalMeta.textContent = `Dumper: ${item.dumperType || 'BEML'} ${item.serialNo} | Total Cycles: ${item.rowCount} | Total Payload: ${item.totalPayload} Tons (Showing all ${item.rowCount} rows)`;

  // Build header
  previewThead.innerHTML = `<tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr>`;

  // Build rows (all rows)
  const rowsToShow = item.parsedRows;
  previewTbody.innerHTML = rowsToShow.map(row => 
    `<tr>${row.map(c => `<td>${c}</td>`).join('')}</tr>`
  ).join('');

  previewModal.classList.add('open');
};

// Download all processed CSVs in a ZIP
async function downloadAllAsZip() {
  const validItems = processedFiles.filter(f => f.status === 'success');
  if (validItems.length === 0) {
    showToast('No converted CSVs available to zip.');
    return;
  }

  btnDownloadAll.disabled = true;
  btnDownloadAll.innerHTML = `<span class="spinner"></span> Creating ZIP...`;

  try {
    const zip = new JSZip();
    validItems.forEach(item => {
      zip.file(item.outputFilename, item.csvContent);
    });

    // If multiple dumpers, include the consolidated CSV in the ZIP bundle
    if (validItems.length > 1) {
      zip.file(getConsolidatedFilename(), getConsolidatedCsvContent());
    }

    const month = getSelectedMonth();
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const zipFilename = `TPMS_CSV_${month}.zip`;

    triggerDownload(zipBlob, zipFilename);
    showToast(`Downloaded ${zipFilename} with ${validItems.length} CSV files!`);
  } catch (err) {
    console.error('Error generating zip:', err);
    showToast('Error generating zip file.');
  } finally {
    btnDownloadAll.disabled = false;
    btnDownloadAll.innerHTML = `
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
      Download All (ZIP)
    `;
  }
}

// Dumper Type Selection in Generator
function setGenDumperType(type) {
  genDumperType = type;
  if (type === 'BEML') {
    if (genTypeBeml) genTypeBeml.classList.add('active');
    if (genTypeCat) genTypeCat.classList.remove('active');
    if (genSerialLabel) genSerialLabel.textContent = 'Dumper Serial Number';
    if (genSubtitle) genSubtitle.textContent = 'Create realistic BEML TPMS payload workbooks (.xlsx) and matching CSVs for any dumper, date, shift, and trip count.';
    genSerial.placeholder = 'e.g. 60611';
    if (!BEML_GEN_SERIALS.includes(genSerial.value)) {
      genSerial.value = '60611';
    }
    renderGenSerialChips(BEML_GEN_SERIALS, genSerial.value);
    if (btnGenDownloadRawLabel) btnGenDownloadRawLabel.textContent = 'Download .XLSX';
  } else {
    if (genTypeCat) genTypeCat.classList.add('active');
    if (genTypeBeml) genTypeBeml.classList.remove('active');
    if (genSerialLabel) genSerialLabel.textContent = 'Equipment / Serial Number';
    if (genSubtitle) genSubtitle.textContent = 'Create realistic Caterpillar MineStar reports (.csv) and matching 16-column CSVs for any dumper, date, shift, and trip count.';
    genSerial.placeholder = 'e.g. PRB00914';
    if (!CAT_GEN_SERIALS.includes(genSerial.value)) {
      genSerial.value = 'PRB00914';
    }
    renderGenSerialChips(CAT_GEN_SERIALS, genSerial.value);
    if (btnGenDownloadRawLabel) btnGenDownloadRawLabel.textContent = 'Download MineStar Raw (.CSV)';
  }
}

function renderGenSerialChips(serials, activeSerial) {
  if (!serialChips) return;
  serialChips.innerHTML = serials.map(sn => `
    <button type="button" class="chip ${sn === activeSerial ? 'active' : ''}" data-sn="${sn}">${sn}</button>
  `).join('');
}

// Synthetic Report Generation Engine
function handleGenerateReport() {
  const isCat = (genDumperType === 'CAT');
  const defaultSerial = isCat ? 'PRB00914' : '60611';
  const serialNo = (genSerial.value || defaultSerial).trim();
  const dateVal = genDate.value;
  if (!dateVal) {
    showToast('Please select a date.');
    return;
  }

  const [yearStr, monthStr, dayStr] = dateVal.split('-');
  const year = parseInt(yearStr, 10);
  const monthIdx = parseInt(monthStr, 10) - 1;
  const day = parseInt(dayStr, 10);
  const dateObj = new Date(year, monthIdx, day);

  const monthNames = [
    'january', 'february', 'march', 'april', 'may', 'june',
    'july', 'august', 'september', 'october', 'november', 'december'
  ];
  const monthName = monthNames[monthIdx];

  // Shift config:
  // 1st Shift: 06:00-14:00 -> Production: 07:00 to 13:00
  // 2nd Shift: 14:00-22:00 -> Production: 15:00 to 21:00
  // 3rd Shift: 22:00-06:00 -> Production: 23:00 to 05:00 (next day)
  let shiftLabel = '06:00 to 14:00';
  let shiftName = '1st_Shift';
  let prodStart = new Date(year, monthIdx, day, 7, Math.floor(Math.random() * 4), Math.floor(Math.random() * 60));
  let prodEnd = new Date(year, monthIdx, day, 13, 0, 0);

  if (selectedShift === '2') {
    shiftLabel = '14:00 to 22:00';
    shiftName = '2nd_Shift';
    prodStart = new Date(year, monthIdx, day, 15, Math.floor(Math.random() * 4), Math.floor(Math.random() * 60));
    prodEnd = new Date(year, monthIdx, day, 21, 0, 0);
  } else if (selectedShift === '3') {
    shiftLabel = '22:00 to 06:00';
    shiftName = '3rd_Shift';
    prodStart = new Date(year, monthIdx, day, 23, Math.floor(Math.random() * 4), Math.floor(Math.random() * 60));
    prodEnd = new Date(year, monthIdx, day + 1, 5, 0, 0);
  }

  let numTrips = parseInt(genTrips.value, 10);
  if (isNaN(numTrips) || numTrips <= 0) {
    numTrips = Math.floor(Math.random() * 4) + 11; // 11 to 14 trips
  }

  const mealSec = Math.floor(Math.random() * 11 + 25) * 60; // 25-35 min meal break

  if (isCat) {
    // -------------------------------------------------------------
    // CAT MINESTAR GENERATOR PIPELINE
    // -------------------------------------------------------------
    let currTime = new Date(prodStart.getTime());
    let smhDelta = 11500.00 + Math.floor(Math.random() * 2000) + Math.round(Math.random() * 100) / 100;
    const catRecords = [];
    let totalTonnage = 0;

    const fmtHHMMSS = (sec) => {
      const h = Math.floor(sec / 3600);
      const m = Math.floor((sec % 3600) / 60);
      const s = sec % 60;
      return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
    };

    const fmtHMMSS = (sec) => {
      const h = Math.floor(sec / 3600);
      const m = Math.floor((sec % 3600) / 60);
      const s = sec % 60;
      return `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
    };

    for (let i = 1; i <= numTrips; i++) {
      // Payload Gaussian around 45.0 tons
      const randNorm = (Math.random() + Math.random() + Math.random() + Math.random() - 2) * 4.5;
      let rawPayload = 45.5 + randNorm;
      rawPayload = Math.max(33.0, Math.min(62.0, rawPayload));
      const payload = Math.round(rawPayload * 100) / 100;

      // Distances (miles)
      const emptyDist = Math.round((0.70 + Math.random() * 0.35) * 100) / 100;
      const loadDist = Math.round((0.75 + Math.random() * 0.35) * 100) / 100;
      const totDist = Math.round((emptyDist + loadDist) * 100) / 100;

      // Durations (seconds)
      const emptyTravelSec = Math.floor(Math.random() * 120) + 210;
      const emptyStopSec = Math.random() < 0.65 ? Math.floor(Math.random() * 10) + 1 : Math.floor(Math.random() * 60) + 20;
      const loadSec = Math.floor(Math.random() * 55) + 68;
      const loadStopSec = Math.floor(Math.random() * 25) + 25;
      const loadTravelSec = Math.floor(Math.random() * 120) + 220;

      const totCycleSec = emptyTravelSec + emptyStopSec + loadSec + loadStopSec + loadTravelSec;

      // Check cutoff
      const cycleFinish = new Date(currTime.getTime() + totCycleSec * 1000);
      if (cycleFinish > prodEnd) {
        break;
      }

      totalTonnage += payload;

      const dateStr = `${String(currTime.getDate()).padStart(2, '0')}/${String(currTime.getMonth() + 1).padStart(2, '0')}/${currTime.getFullYear()}`;
      const timeStr = `${String(currTime.getHours()).padStart(2, '0')}:${String(currTime.getMinutes()).padStart(2, '0')}:${String(currTime.getSeconds()).padStart(2, '0')}`;
      const timedate = `${dateStr} ${timeStr}`;

      const passCount = payload < 35.0 ? 2 : payload > 52.0 ? 4 : (Math.random() < 0.5 ? 3 : 4);
      const fuel = Math.round((1.9 + (totCycleSec / 600) * 0.8 + totDist * 0.35 + (Math.random() - 0.5) * 0.2) * 100) / 100;

      const currentSmh = smhDelta.toFixed(2);
      smhDelta += (totCycleSec / 3600.0);

      const loadLoc = `21.736${Math.floor(Math.random() * 400 + 200)},83.816${Math.floor(Math.random() * 300 + 100)},155.0`;
      const dumpLoc = `21.731${Math.floor(Math.random() * 400 + 500)},83.823${Math.floor(Math.random() * 300 + 600)},199.0`;

      catRecords.push({
        timedate,
        smhDelta: currentSmh,
        payload: payload.toFixed(2),
        emptyTravelHH: fmtHHMMSS(emptyTravelSec),
        emptyStopHH: fmtHHMMSS(emptyStopSec),
        loadTimeHH: fmtHHMMSS(loadSec),
        loadStopHH: fmtHHMMSS(loadStopSec),
        loadTravelHH: fmtHHMMSS(loadTravelSec),
        cycleTimeHH: fmtHHMMSS(totCycleSec),
        emptyTravelH: fmtHMMSS(emptyTravelSec),
        emptyStopH: fmtHMMSS(emptyStopSec),
        loadTimeH: fmtHMMSS(loadSec),
        loadStopH: fmtHMMSS(loadStopSec),
        loadTravelH: fmtHMMSS(loadTravelSec),
        cycleTimeH: fmtHMMSS(totCycleSec),
        emptyDist: emptyDist.toFixed(2),
        loadDist: loadDist.toFixed(2),
        totDist: totDist.toFixed(2),
        passCount,
        shiftCount: i,
        fuel: fuel.toFixed(2),
        loadLoc,
        dumpLoc
      });

      // Advance time
      if (i === Math.floor(numTrips / 2) && numTrips >= 7) {
        currTime = new Date(cycleFinish.getTime() + mealSec * 1000);
      } else {
        const tinyGap = Math.random() < 0.85 ? 0 : Math.floor(Math.random() * 8) + 2;
        currTime = new Date(cycleFinish.getTime() + tinyGap * 1000);
      }
    }

    // 1. Build Raw MineStar CSV
    const now = new Date();
    const daysArr = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const monthsArr = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const dayName = daysArr[now.getDay()];
    const monName = monthsArr[now.getMonth()];
    const dateRangeStr = `${String(day).padStart(2, '0')}/${String(monthIdx + 1).padStart(2, '0')}/${year}`;
    const repGenDateStr = `${dayName} ${monName} ${String(now.getDate()).padStart(2, '0')} ${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')} GMT+0530 (India Standard Time)`;

    const rawLines = [
      ',,,,,MineStar Health Technician Toolbox',
      `,,,,,Truck Payload Detail Data(Date Range :${dateRangeStr} 00:00 AM-${dateRangeStr} 11:59 PM)`,
      `,,,,,Report Generated Date :${repGenDateStr}`,
      `,,,,,Equipment :${serialNo}`,
      '',
      ',,,,, ',
      'Date and Time,SMH Delta,Payload Weight (ton),Empty Travel Time  (hh:mm:ss),Empty Travel Distance (mile),Empty Stop Time (hh:mm:ss),Load Time (hh:mm:ss),Load Stop Time (hh:mm:ss),Load Travel Time (hh:mm:ss),Load Travel Distance (mile),Total Cycle Time (hh:mm:ss),Total Distance Travelled (mile),Loader Pass Count,Shift Count,Operator ID,Fuel Burned (gal),Load Location,Dump Location'
    ];

    catRecords.forEach(r => {
      rawLines.push(`${r.timedate},${r.smhDelta},${r.payload},${r.emptyTravelHH},${r.emptyDist},${r.emptyStopHH},${r.loadTimeHH},${r.loadStopHH},${r.loadTravelHH},${r.loadDist},${r.cycleTimeHH},${r.totDist},${r.passCount},${r.shiftCount},9.0,${r.fuel},"${r.loadLoc}","${r.dumpLoc}"`);
    });
    const rawMineStarText = rawLines.join('\n') + '\n';
    const rawBlob = new Blob([rawMineStarText], { type: 'text/csv;charset=utf-8;' });

    // 2. Build Standard 16-Column Converted CSV
    const csvLines = [CAT_HEADERS.join(',')];
    catRecords.forEach(r => {
      csvLines.push(`${r.timedate},${r.smhDelta},${serialNo},${r.payload},${r.emptyTravelH},${r.emptyDist},${r.emptyStopH},${r.loadTimeH},${r.loadStopH},${r.loadTravelH},${r.loadDist},${r.cycleTimeH},${r.totDist},${r.passCount},${r.fuel},`);
    });
    const csvText = csvLines.join('\n') + '\n';
    const csvBlob = new Blob([csvText], { type: 'text/csv;charset=utf-8;' });

    // Filenames
    const rawFilename = `TruckPayload_${serialNo}_${dayName} ${String(now.getDate()).padStart(2, '0')} ${monName} ${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')} ${String(now.getMinutes()).padStart(2, '0')} ${String(now.getSeconds()).padStart(2, '0')}.csv`;
    const csvFilename = `${serialNo}.csv`;
    const zipFilename = `TPMS_CAT_${serialNo}_${shiftName}_${dateVal}.zip`;

    currentGenerated = {
      dumperType: 'CAT',
      serialNo,
      shiftLabel,
      shiftName,
      dateStr: dateVal,
      numTrips: catRecords.length,
      totalTonnage: totalTonnage.toFixed(2),
      records: catRecords,
      rawBlob,
      rawFilename,
      csvBlob,
      csvFilename,
      zipFilename,
      xlsxBlob: rawBlob,
      xlsxFilename: rawFilename
    };

    // Render Stats
    genStatSN.textContent = serialNo;
    genStatShift.textContent = `${selectedShift === '1' ? '1st' : selectedShift === '2' ? '2nd' : '3rd'} Shift`;
    genStatTrips.textContent = catRecords.length;
    genStatTonnage.textContent = `${totalTonnage.toFixed(1)} T`;

    // Render Table Preview (16 cols)
    genThead.innerHTML = `<tr>${CAT_HEADERS.map(h => `<th>${h}</th>`).join('')}</tr>`;
    genTbody.innerHTML = catRecords.map(r => `
      <tr>
        <td>${r.timedate}</td>
        <td>${r.smhDelta}</td>
        <td>${serialNo}</td>
        <td>${r.payload}</td>
        <td>${r.emptyTravelH}</td>
        <td>${r.emptyDist}</td>
        <td>${r.emptyStopH}</td>
        <td>${r.loadTimeH}</td>
        <td>${r.loadStopH}</td>
        <td>${r.loadTravelH}</td>
        <td>${r.loadDist}</td>
        <td>${r.cycleTimeH}</td>
        <td>${r.totDist}</td>
        <td>${r.passCount}</td>
        <td>${r.fuel}</td>
        <td></td>
      </tr>
    `).join('');

    if (btnGenDownloadRawLabel) {
      btnGenDownloadRawLabel.textContent = 'Download MineStar Raw (.CSV)';
    }

    genResultsCard.classList.remove('hidden');
    showToast(`Generated ${catRecords.length} cycles for CAT Dumper ${serialNo}!`);
    return;
  }

  // -------------------------------------------------------------
  // BEML GENERATOR PIPELINE (Existing verified logic)
  // -------------------------------------------------------------
  let currTime = new Date(prodStart.getTime());
  const records = [];
  let totalTonnage = 0;

  for (let i = 1; i <= numTrips; i++) {
    // Payload Gaussian ~53.0 tons
    const randNorm = (Math.random() + Math.random() + Math.random() + Math.random() - 2) * 4.0;
    let rawPayload = 53.0 + randNorm;
    rawPayload = Math.max(42.0, Math.min(64.0, rawPayload));
    const payload = Math.round(rawPayload * 2) / 2; // Step 0.50

    // Distances
    const haulDist = Math.round((1.60 + Math.random() * 0.25) * 100) / 100;
    const retDist = Math.round((haulDist + (Math.random() - 0.5) * 0.08) * 100) / 100;
    const totDist = Math.round((haulDist + retDist) * 100) / 100;

    // Speeds
    const avgHaulSpd = Math.round((14.2 + Math.random() * 4.0) * 100) / 100;
    const maxHaulSpd = Math.round((avgHaulSpd * (1.5 + Math.random() * 0.3)) * 100) / 100;

    const avgRetSpd = Math.round((15.2 + Math.random() * 4.6) * 100) / 100;
    const maxRetSpd = Math.round((avgRetSpd * (1.6 + Math.random() * 0.3)) * 100) / 100;

    // Times in seconds
    const loadSec = Math.floor(Math.random() * 125) + 85; // 1:25 to 3:30
    const haulSec = Math.round((haulDist / avgHaulSpd) * 3600);
    const dumpSec = Math.floor(Math.random() * 23) + 15; // 15 to 38 sec
    const retSec = Math.round((retDist / avgRetSpd) * 3600);
    const retStopSec = Math.random() < 0.75 ? 0 : Math.floor(Math.random() * 80) + 20;

    const totSec = loadSec + haulSec + dumpSec + retSec + retStopSec;

    // Check cutoff
    const cycleFinish = new Date(currTime.getTime() + totSec * 1000);
    if (cycleFinish > prodEnd) {
      break;
    }

    totalTonnage += payload;

    const fmtMS = (s) => `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`;
    const fmt000S = (s) => `${String(Math.floor(s / 60)).padStart(3, '0')}:${String(s % 60).padStart(2, '0')}`;

    const dateStr = `${String(currTime.getDate()).padStart(2, '0')}/${String(currTime.getMonth() + 1).padStart(2, '0')}/${currTime.getFullYear()}`;
    const timeStr = `${String(currTime.getHours()).padStart(2, '0')}:${String(currTime.getMinutes()).padStart(2, '0')}:${String(currTime.getSeconds()).padStart(2, '0')}`;

    records.push({
      slNo: i,
      truckNo: serialNo,
      date: dateStr,
      time: timeStr,
      payload: payload.toFixed(2),
      loadTime: fmtMS(loadSec),
      haulTime: fmtMS(haulSec),
      dumpTime: fmt000S(dumpSec),
      retTime: fmtMS(retSec),
      retStopTime: fmt000S(retStopSec),
      totTime: fmt000S(totSec),
      haulDist: haulDist.toFixed(2),
      retDist: retDist.toFixed(2),
      totDist: totDist.toFixed(2),
      maxHaulSpd: maxHaulSpd.toFixed(2),
      avgHaulSpd: avgHaulSpd.toFixed(2),
      maxRetSpd: maxRetSpd.toFixed(2),
      avgRetSpd: avgRetSpd.toFixed(2)
    });

    // Advance time
    if (i === Math.floor(numTrips / 2) && numTrips >= 7) {
      currTime = new Date(cycleFinish.getTime() + mealSec * 1000);
    } else {
      const tinyGap = Math.random() < 0.85 ? 0 : Math.floor(Math.random() * 11) + 5;
      currTime = new Date(cycleFinish.getTime() + tinyGap * 1000);
    }
  }

  // 1. Build CSV content
  let csvText = CSV_HEADERS.join(',') + '\n';
  for (const r of records) {
    const row = [
      r.truckNo,
      r.date,
      r.time,
      r.payload,
      r.loadTime,
      r.haulTime,
      r.dumpTime,
      r.retTime,
      r.retStopTime,
      r.totTime,
      r.haulDist,
      r.retDist,
      r.totDist,
      r.maxHaulSpd,
      r.avgHaulSpd,
      r.maxRetSpd,
      r.avgRetSpd,
      '' // Uploaded By
    ];
    csvText += row.join(',') + '\n';
  }
  const csvBlob = new Blob([csvText], { type: 'text/csv;charset=utf-8;' });

  // 2. Build Excel Workbook (.xlsx)
  const monthCap = dateObj.toLocaleString('en-US', { month: 'long' });
  const reportDateStr = `${monthCap} ${String(day).padStart(2, '0')},${year}`;
  const periodStr = `${String(day).padStart(2, '0')}/${String(monthIdx + 1).padStart(2, '0')}/${year} to ${String(day).padStart(2, '0')}/${String(monthIdx + 1).padStart(2, '0')}/${year}`;

  const aoaSheet2 = [
    [],
    ['BHARAT EARTH MOVERS LTD, INDIA'],
    ['Payload Report'],
    ['', 'Report Date', reportDateStr, '', '', '', '', '', '', '', '', '', 'Serial Number', '', parseInt(serialNo) || serialNo],
    ['', 'Period', periodStr, '', '', '', '', '', '', '', '', '', 'Model Number', '', 'BH60M   '],
    ['', 'Shift Time', shiftLabel, '', '', '', '', '', '', '', '', '', 'Customer', '', 'MCL LAKHANPUR       '],
    [],
    [
      'Sl. No.', 'Date', 'Time', 'PayLoad', 'Loading Time', 'Hauling Time', 'Stop & Dump Time',
      'Return Time', 'Return Stop Time', 'Total Time', 'Haul Travel Distance', 'Return Travel Distance',
      'Total Distance', 'Max.Haul Speed', 'Avg.Haul Speed', 'Max.Return Speed', 'Avg.Return Speed'
    ],
    [
      '', '', '', 'Tons', 'mm:ss', 'mm:ss', 'mm:ss', 'mm:ss', 'mm:ss', 'mm:ss',
      'km', 'km', 'km', 'kmph', 'kmph', 'kmph', 'kmph'
    ],
    []
  ];

  for (const r of records) {
    aoaSheet2.push([
      r.slNo,
      r.date,
      r.time,
      parseFloat(r.payload),
      r.loadTime,
      r.haulTime,
      r.dumpTime,
      r.retTime,
      r.retStopTime,
      r.totTime,
      parseFloat(r.haulDist),
      parseFloat(r.retDist),
      parseFloat(r.totDist),
      parseFloat(r.maxHaulSpd),
      parseFloat(r.avgHaulSpd),
      parseFloat(r.maxRetSpd),
      parseFloat(r.avgRetSpd)
    ]);
  }

  const wb = XLSX.utils.book_new();
  const ws2 = XLSX.utils.aoa_to_sheet(aoaSheet2);
  const ws1 = XLSX.utils.aoa_to_sheet([['BHARAT EARTH MOVERS LTD, INDIA'], ['Payload Report']]);
  XLSX.utils.book_append_sheet(wb, ws2, 'Sheet2');
  XLSX.utils.book_append_sheet(wb, ws1, 'Sheet1');

  const xlsxArray = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const xlsxBlob = new Blob([xlsxArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

  // Filenames
  const now = new Date();
  const timestampStr = `${String(day).padStart(2, '0')}_${String(monthIdx + 1).padStart(2, '0')}_${year}_${String(now.getHours()).padStart(2, '0')}_${String(now.getMinutes()).padStart(2, '0')}_${String(now.getSeconds()).padStart(2, '0')}`;
  const xlsxFilename = `${serialNo}_${timestampStr}.xlsx`;
  const csvFilename = `${serialNo}_${monthName}.csv`;
  const zipFilename = `TPMS_${serialNo}_${shiftName}_${dateVal}.zip`;

  currentGenerated = {
    dumperType: 'BEML',
    serialNo,
    shiftLabel,
    shiftName,
    dateStr: dateVal,
    numTrips: records.length,
    totalTonnage: totalTonnage.toFixed(2),
    records,
    csvText,
    csvBlob,
    rawBlob: xlsxBlob,
    rawFilename: xlsxFilename,
    xlsxBlob,
    xlsxFilename,
    csvFilename,
    zipFilename
  };

  // Render Stats
  genStatSN.textContent = serialNo;
  genStatShift.textContent = `${selectedShift === '1' ? '1st' : selectedShift === '2' ? '2nd' : '3rd'} Shift`;
  genStatTrips.textContent = records.length;
  genStatTonnage.textContent = `${totalTonnage.toFixed(1)} T`;

  // Render Table Preview (All rows)
  genThead.innerHTML = `<tr>${CSV_HEADERS.map(h => `<th>${h}</th>`).join('')}</tr>`;
  genTbody.innerHTML = records.map(r => `
    <tr>
      <td>${r.truckNo}</td>
      <td>${r.date}</td>
      <td>${r.time}</td>
      <td>${r.payload}</td>
      <td>${r.loadTime}</td>
      <td>${r.haulTime}</td>
      <td>${r.dumpTime}</td>
      <td>${r.retTime}</td>
      <td>${r.retStopTime}</td>
      <td>${r.totTime}</td>
      <td>${r.haulDist}</td>
      <td>${r.retDist}</td>
      <td>${r.totDist}</td>
      <td>${r.maxHaulSpd}</td>
      <td>${r.avgHaulSpd}</td>
      <td>${r.maxRetSpd}</td>
      <td>${r.avgRetSpd}</td>
      <td></td>
    </tr>
  `).join('');

  if (btnGenDownloadRawLabel) {
    btnGenDownloadRawLabel.textContent = 'Download .XLSX';
  }

  genResultsCard.classList.remove('hidden');
  showToast(`Generated ${records.length} cycles for Dumper ${serialNo}!`);
}

// Download Generated ZIP
async function downloadGeneratedZip() {
  if (!currentGenerated) return;

  btnGenDownloadZip.disabled = true;
  btnGenDownloadZip.innerHTML = `<span class="spinner"></span> Zipping...`;

  try {
    const zip = new JSZip();
    zip.file(currentGenerated.rawFilename, currentGenerated.rawBlob);
    zip.file(currentGenerated.csvFilename, currentGenerated.csvBlob);

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const zipFilename = currentGenerated.zipFilename || `TPMS_${currentGenerated.serialNo}_${currentGenerated.shiftName}_${currentGenerated.dateStr}.zip`;

    triggerDownload(zipBlob, zipFilename);
    showToast(`Downloaded ${zipFilename}`);
  } catch (err) {
    console.error('Error generating zip:', err);
    showToast('Failed to create ZIP.');
  } finally {
    btnGenDownloadZip.disabled = false;
    btnGenDownloadZip.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
      Download Both (ZIP)
    `;
  }
}

// Download Helper
function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// Toast helper
let toastTimer = null;
function showToast(message) {
  if (toastTimer) clearTimeout(toastTimer);
  toast.textContent = message;
  toast.classList.add('show');
  toastTimer = setTimeout(() => {
    toast.classList.remove('show');
  }, 3500);
}

// Service Worker Registration
function registerServiceWorker() {
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('./sw.js')
        .then((reg) => console.log('ServiceWorker registered:', reg.scope))
        .catch((err) => console.log('ServiceWorker registration failed:', err));
    });
  }
}
