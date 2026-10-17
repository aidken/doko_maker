import './style.css';
import 'tabulator-tables/dist/css/tabulator_modern.min.css';
import { TabulatorFull as Tabulator } from 'tabulator-tables';
import { createIcons, icons } from 'lucide';
import * as XLSX from 'xlsx';
import { ExcelHandler } from './excel-handler.js';

// Expose XLSX globally for Tabulator's internal downloader compatibility
window.XLSX = XLSX;

// Application State
const state = {
  currentWorkbook: null,
  currentFileName: '',
  currentFileSize: 0,
  activeSheetName: '',
  tabulatorTable: null,
  totalRows: 0,
  totalCols: 0
};

// DOM Elements
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('file-input');
const browseBtn = document.getElementById('browse-btn');
const loadSampleBtn = document.getElementById('load-sample-btn');
const headerSampleBtn = document.getElementById('header-sample-btn');

const uploadSection = document.getElementById('upload-section');
const editorSection = document.getElementById('editor-section');

const fileNameEl = document.getElementById('meta-filename');
const fileSizeEl = document.getElementById('meta-filesize');
const rowCountEl = document.getElementById('meta-rowcount');
const colCountEl = document.getElementById('meta-colcount');
const sheetTabsContainer = document.getElementById('sheet-tabs-container');

const searchInput = document.getElementById('search-input');
const clearSearchBtn = document.getElementById('clear-search-btn');
const addRowBtn = document.getElementById('add-row-btn');
const deleteRowBtn = document.getElementById('delete-row-btn');
const undoBtn = document.getElementById('undo-btn');
const redoBtn = document.getElementById('redo-btn');
const pageSizeSelect = document.getElementById('page-size-select');
const exportExcelBtn = document.getElementById('export-excel-btn');
const exportCsvBtn = document.getElementById('export-csv-btn');
const changeFileBtn = document.getElementById('change-file-btn');

/**
 * Initialize application
 */
function initApp() {
  refreshIcons();
  setupUploadEvents();
  setupToolbarEvents();
}

/**
 * Render or refresh Lucide icons
 */
function refreshIcons() {
  createIcons({ icons });
}

/**
 * Configure drag-and-drop & file selection events
 */
function setupUploadEvents() {
  browseBtn.addEventListener('click', () => fileInput.click());

  fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handleFile(file);
    fileInput.value = '';
  });

  ['dragenter', 'dragover'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('drag-active');
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('drag-active');
    }, false);
  });

  dropzone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    const file = dt.files[0];
    if (file) handleFile(file);
  });

  [loadSampleBtn, headerSampleBtn].forEach(btn => {
    if (btn) {
      btn.addEventListener('click', () => loadSampleData());
    }
  });

  changeFileBtn.addEventListener('click', () => {
    fileInput.click();
  });
}

/**
 * Process uploaded file
 */
async function handleFile(file) {
  const validExtensions = ['.xlsx', '.xls', '.csv'];
  const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

  if (!validExtensions.includes(ext)) {
    showToast(`Invalid file type "${ext}". Please upload an Excel (.xlsx, .xls) or .csv file.`, 'error');
    return;
  }

  showToast(`Parsing ${file.name}...`, 'info');

  try {
    const parsed = await ExcelHandler.readWorkbookFromFile(file);
    state.currentWorkbook = parsed.workbook;
    state.currentFileName = parsed.fileName;
    state.currentFileSize = parsed.fileSize;

    loadWorkbookIntoUI(parsed.sheetNames[0]);
    showToast(`Successfully loaded ${file.name}`, 'success');
  } catch (error) {
    console.error(error);
    showToast(`Error parsing file: ${error.message}`, 'error');
  }
}

/**
 * Load built-in sample transportation report data
 */
function loadSampleData() {
  const sample = ExcelHandler.generateSampleWorkbook();
  state.currentWorkbook = sample.workbook;
  state.currentFileName = sample.fileName;
  state.currentFileSize = sample.fileSize;

  loadWorkbookIntoUI(sample.sheetNames[0]);
  showToast('Loaded sample transportation report', 'success');
}

/**
 * Populate UI with current workbook sheets and active sheet data
 */
function loadWorkbookIntoUI(sheetToSelect) {
  if (!state.currentWorkbook || !state.currentWorkbook.SheetNames.length) return;

  const sheetNames = state.currentWorkbook.SheetNames;
  state.activeSheetName = sheetToSelect || sheetNames[0];

  fileNameEl.textContent = state.currentFileName;
  fileSizeEl.textContent = ExcelHandler.formatBytes(state.currentFileSize);

  renderSheetTabs(sheetNames, state.activeSheetName);
  loadSheet(state.activeSheetName);

  uploadSection.classList.add('hidden');
  editorSection.classList.remove('hidden');

  refreshIcons();
}

/**
 * Render multi-sheet tab navigation
 */
function renderSheetTabs(sheetNames, activeSheet) {
  sheetTabsContainer.innerHTML = '';

  sheetNames.forEach(name => {
    const tabBtn = document.createElement('button');
    const isActive = name === activeSheet;

    tabBtn.className = `px-3.5 py-1.5 text-xs md:text-sm font-medium rounded-lg transition-colors flex items-center gap-1.5 ${
      isActive
        ? 'bg-blue-600 text-white shadow-sm'
        : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
    }`;

    tabBtn.innerHTML = `
      <i data-lucide="file-spreadsheet" class="w-3.5 h-3.5"></i>
      <span>${name}</span>
    `;

    tabBtn.addEventListener('click', () => {
      if (state.activeSheetName !== name) {
        state.activeSheetName = name;
        renderSheetTabs(sheetNames, name);
        loadSheet(name);
      }
    });

    sheetTabsContainer.appendChild(tabBtn);
  });

  refreshIcons();
}

/**
 * Load specific sheet into Tabulator
 */
function loadSheet(sheetName) {
  const sheet = state.currentWorkbook.Sheets[sheetName];
  if (!sheet) {
    showToast(`Sheet "${sheetName}" not found.`, 'error');
    return;
  }

  const { columns, data, totalRows, totalCols } = ExcelHandler.parseSheetData(sheet);

  state.totalRows = totalRows;
  state.totalCols = totalCols;

  rowCountEl.textContent = `${totalRows} rows`;
  colCountEl.textContent = `${totalCols} cols`;

  if (state.tabulatorTable) {
    state.tabulatorTable.destroy();
  }

  state.tabulatorTable = new Tabulator('#excel-grid', {
    data: data,
    columns: columns,
    layout: "fitDataFill",
    responsiveLayout: false,
    history: true,
    pagination: "local",
    paginationSize: parseInt(pageSizeSelect.value) || 25,
    paginationSizeSelector: [10, 25, 50, 100, true],
    paginationCounter: "rows",
    placeholder: "No records found in this sheet.",
    movableColumns: true,
    renderVertical: "virtual",
    height: "560px",
    rowSelection: true,
    dataChanged: function() {
      updateGridMetaCounts();
    }
  });

  searchInput.value = '';
}

/**
 * Update row and column counters in metadata bar
 */
function updateGridMetaCounts() {
  if (!state.tabulatorTable) return;
  const count = state.tabulatorTable.getDataCount();
  rowCountEl.textContent = `${count} rows`;
}

/**
 * Set up grid actions (filtering, editing, row deletion, exports)
 */
function setupToolbarEvents() {
  let debounceTimeout;
  searchInput.addEventListener('input', (e) => {
    clearTimeout(debounceTimeout);
    debounceTimeout = setTimeout(() => {
      const term = e.target.value.trim().toLowerCase();
      if (!state.tabulatorTable) return;

      if (!term) {
        state.tabulatorTable.clearFilter();
        return;
      }

      state.tabulatorTable.setFilter((data) => {
        return Object.values(data).some(val => {
          if (val === null || val === undefined) return false;
          return String(val).toLowerCase().includes(term);
        });
      });
    }, 200);
  });

  clearSearchBtn.addEventListener('click', () => {
    searchInput.value = '';
    if (state.tabulatorTable) {
      state.tabulatorTable.clearFilter();
    }
  });

  undoBtn.addEventListener('click', () => {
    if (state.tabulatorTable) {
      state.tabulatorTable.undo();
      showToast('Action undone', 'info');
    }
  });

  redoBtn.addEventListener('click', () => {
    if (state.tabulatorTable) {
      state.tabulatorTable.redo();
      showToast('Action redone', 'info');
    }
  });

  addRowBtn.addEventListener('click', () => {
    if (!state.tabulatorTable) return;
    const newId = (state.tabulatorTable.getData().length || 0) + 1;
    state.tabulatorTable.addRow({ _id: newId }, true);
    showToast('New row added at top', 'info');
  });

  deleteRowBtn.addEventListener('click', () => {
    if (!state.tabulatorTable) return;
    const selectedRows = state.tabulatorTable.getSelectedRows();
    if (!selectedRows.length) {
      showToast('Select one or more rows to delete.', 'info');
      return;
    }

    const count = selectedRows.length;
    selectedRows.forEach(row => row.delete());
    showToast(`Deleted ${count} selected row${count > 1 ? 's' : ''}`, 'info');
  });

  pageSizeSelect.addEventListener('change', (e) => {
    if (!state.tabulatorTable) return;
    const val = e.target.value;
    if (val === 'all') {
      state.tabulatorTable.setPageSize(true);
    } else {
      state.tabulatorTable.setPageSize(parseInt(val, 10));
    }
  });

  exportExcelBtn.addEventListener('click', () => {
    if (!state.tabulatorTable) {
      showToast('No table data to export.', 'error');
      return;
    }

    const baseName = state.currentFileName
      ? state.currentFileName.replace(/\.[^/.]+$/, "")
      : 'doko_report';
    const exportName = `${baseName}_export.xlsx`;

    ExcelHandler.exportTableToExcel(
      state.tabulatorTable,
      exportName,
      state.activeSheetName || 'Sheet1'
    );
    showToast(`Exported "${exportName}"`, 'success');
  });

  exportCsvBtn.addEventListener('click', () => {
    if (!state.tabulatorTable) {
      showToast('No table data to export.', 'error');
      return;
    }

    const baseName = state.currentFileName
      ? state.currentFileName.replace(/\.[^/.]+$/, "")
      : 'doko_report';
    state.tabulatorTable.download("csv", `${baseName}_export.csv`);
    showToast(`Exported CSV file`, 'success');
  });
}

/**
 * Toast notification system
 */
function showToast(message, type = 'info') {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const toast = document.createElement('div');
  const bgColors = {
    info: 'bg-slate-800 text-white border-slate-700',
    success: 'bg-emerald-600 text-white border-emerald-500',
    error: 'bg-rose-600 text-white border-rose-500'
  };

  const iconsMap = {
    info: 'info',
    success: 'check-circle-2',
    error: 'alert-triangle'
  };

  toast.className = `toast px-4 py-3 rounded-lg shadow-lg border text-sm flex items-center gap-2.5 ${bgColors[type] || bgColors.info}`;
  toast.innerHTML = `
    <i data-lucide="${iconsMap[type] || 'info'}" class="w-4 h-4 flex-shrink-0"></i>
    <span>${message}</span>
  `;

  container.appendChild(toast);
  refreshIcons();

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(8px)';
    setTimeout(() => toast.remove(), 250);
  }, 3500);
}

// Start app
document.addEventListener('DOMContentLoaded', initApp);

