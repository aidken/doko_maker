import './style.css';
import 'tabulator-tables/dist/css/tabulator_modern.min.css';
import { TabulatorFull as Tabulator } from 'tabulator-tables';
import { createIcons, icons } from 'lucide';
import * as XLSX from 'xlsx';
import { ExcelHandler } from './excel-handler.js';

// Expose XLSX globally for Tabulator compatibility
window.XLSX = XLSX;

// Application State
const state = {
  tabulatorTable: null,
  currentMeta: null,
  activeFilter: {
    search: '',
    status: '',
    onlyChanges: false
  }
};

// DOM Elements
const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('file-input');
const browseBtn = document.getElementById('browse-btn');
const loadSampleBtn = document.getElementById('load-sample-btn');
const headerSampleBtn = document.getElementById('header-sample-btn');
const uploadNewBtn = document.getElementById('upload-new-btn');

const uploadSection = document.getElementById('upload-section');
const editorSection = document.getElementById('editor-section');
const progressContainer = document.getElementById('progress-container');
const progressBar = document.getElementById('progress-bar');
const progressText = document.getElementById('progress-text');
const progressPct = document.getElementById('progress-pct');

const metricTotalOrders = document.getElementById('metric-total-orders');
const metricOrdersChanged = document.getElementById('metric-orders-changed');
const metricFilesCount = document.getElementById('metric-files-count');
const metricDateRange = document.getElementById('metric-date-range');

const searchInput = document.getElementById('search-input');
const clearSearchBtn = document.getElementById('clear-search-btn');
const statusFilterSelect = document.getElementById('status-filter-select');
const onlyChangesToggle = document.getElementById('only-changes-toggle');
const pageSizeSelect = document.getElementById('page-size-select');
const exportExcelBtn = document.getElementById('export-excel-btn');
const exportCsvBtn = document.getElementById('export-csv-btn');

/**
 * Initialize application
 */
function initApp() {
  refreshIcons();
  setupUploadEvents();
  setupToolbarEvents();
}

/**
 * Refresh Lucide SVG icons in the DOM
 */
function refreshIcons() {
  createIcons({ icons });
}

/**
 * Set up drag & drop and file selection events
 */
function setupUploadEvents() {
  browseBtn.addEventListener('click', () => fileInput.click());

  fileInput.addEventListener('change', (e) => {
    const files = Array.from(e.target.files);
    if (files.length > 0) processFiles(files);
    fileInput.value = '';
  });

  ['dragenter', 'dragover'].forEach(name => {
    dropzone.addEventListener(name, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('drag-active');
    }, false);
  });

  ['dragleave', 'drop'].forEach(name => {
    dropzone.addEventListener(name, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('drag-active');
    }, false);
  });

  dropzone.addEventListener('drop', (e) => {
    const files = Array.from(e.dataTransfer.files).filter(f =>
      f.name.endsWith('.xlsx') || f.name.endsWith('.xls')
    );
    if (files.length > 0) {
      processFiles(files);
    } else {
      showToast('Please drop valid Excel (.xlsx / .xls) files.', 'error');
    }
  });

  // Sample data buttons
  [loadSampleBtn, headerSampleBtn].forEach(btn => {
    if (btn) {
      btn.addEventListener('click', () => loadSampleDataset());
    }
  });

  // Return to upload view
  uploadNewBtn.addEventListener('click', () => {
    editorSection.classList.add('hidden');
    uploadSection.classList.remove('hidden');
    progressContainer.classList.add('hidden');
  });
}

/**
 * Updates UI progress bar during file parsing
 */
function updateProgress(message, percent) {
  progressContainer.classList.remove('hidden');
  progressText.textContent = message;
  progressPct.textContent = `${percent}%`;
  progressBar.style.width = `${percent}%`;
}

/**
 * Ingest and process a list of snapshot Excel files
 */
async function processFiles(files) {
  try {
    updateProgress(`Reading ${files.length} file${files.length > 1 ? 's' : ''}...`, 10);

    const result = await ExcelHandler.processMultipleReports(files, (msg, pct) => {
      updateProgress(msg, pct);
    });

    state.currentMeta = result.meta;

    // Populate metrics
    metricTotalOrders.textContent = result.meta.totalOrders.toLocaleString();
    metricOrdersChanged.textContent = result.meta.ordersWithChanges.toLocaleString();
    metricFilesCount.textContent = `${result.meta.filesCount} snapshot${result.meta.filesCount > 1 ? 's' : ''}`;
    metricDateRange.textContent = result.meta.dateRange || 'N/A';

    // Initialize or re-create Tabulator table
    renderTable(result.columns, result.data);

    // Switch views
    setTimeout(() => {
      progressContainer.classList.add('hidden');
      uploadSection.classList.add('hidden');
      editorSection.classList.remove('hidden');
      refreshIcons();
      showToast(`Successfully processed ${result.meta.totalOrders} orders across ${result.meta.filesCount} snapshots!`, 'success');
    }, 300);

  } catch (error) {
    console.error(error);
    progressContainer.classList.add('hidden');
    showToast(`Error processing files: ${error.message}`, 'error');
  }
}

/**
 * Load built-in synthetic multi-day snapshot sample dataset
 */
function loadSampleDataset() {
  const sampleFiles = ExcelHandler.generateSampleSnapshots();
  processFiles(sampleFiles);
}

/**
 * Renders Tabulator table with parsed columns and data
 */
function renderTable(columns, data) {
  if (state.tabulatorTable) {
    state.tabulatorTable.destroy();
  }

  state.tabulatorTable = new Tabulator('#excel-grid', {
    data: data,
    columns: columns,
    layout: "fitDataFill",
    responsiveLayout: false,
    pagination: "local",
    paginationSize: parseInt(pageSizeSelect.value) || 25,
    paginationSizeSelector: [10, 25, 50, 100, true],
    paginationCounter: "rows",
    placeholder: "No orders match the current filter criteria.",
    movableColumns: true,
    renderVertical: "virtual",
    height: "580px",
    rowSelection: true
  });

  // Reset filters in UI
  searchInput.value = '';
  statusFilterSelect.value = '';
  onlyChangesToggle.checked = false;
  state.activeFilter = { search: '', status: '', onlyChanges: false };
}

/**
 * Applies multi-condition filtering (search query, status, and changes toggle)
 */
function applyFilters() {
  if (!state.tabulatorTable) return;

  const { search, status, onlyChanges } = state.activeFilter;

  state.tabulatorTable.setFilter((data) => {
    // 1. Changes filter
    if (onlyChanges && !data.hasChanges) {
      return false;
    }

    // 2. Status filter
    if (status && (!data.containerStatus || !data.containerStatus.toLowerCase().includes(status.toLowerCase()))) {
      return false;
    }

    // 3. Global search
    if (search) {
      const matchFound = Object.entries(data).some(([k, val]) => {
        if (k === '_id' || k === 'hasChanges') return false;
        if (val === null || val === undefined) return false;
        return String(val).toLowerCase().includes(search);
      });
      if (!matchFound) return false;
    }

    return true;
  });
}

/**
 * Configure search, filters, and export buttons
 */
function setupToolbarEvents() {
  // Global search input
  let debounceTimeout;
  searchInput.addEventListener('input', (e) => {
    clearTimeout(debounceTimeout);
    debounceTimeout = setTimeout(() => {
      state.activeFilter.search = e.target.value.trim().toLowerCase();
      applyFilters();
    }, 200);
  });

  clearSearchBtn.addEventListener('click', () => {
    searchInput.value = '';
    state.activeFilter.search = '';
    applyFilters();
  });

  // Status dropdown filter
  statusFilterSelect.addEventListener('change', (e) => {
    state.activeFilter.status = e.target.value;
    applyFilters();
  });

  // "Only Orders with Changes" toggle
  onlyChangesToggle.addEventListener('change', (e) => {
    state.activeFilter.onlyChanges = e.target.checked;
    applyFilters();
  });

  // Rows per page
  pageSizeSelect.addEventListener('change', (e) => {
    if (!state.tabulatorTable) return;
    const val = e.target.value;
    if (val === 'all') {
      state.tabulatorTable.setPageSize(true);
    } else {
      state.tabulatorTable.setPageSize(parseInt(val, 10));
    }
  });

  // Export to Excel (.xlsx)
  exportExcelBtn.addEventListener('click', () => {
    if (!state.tabulatorTable) {
      showToast('No report data available to export.', 'error');
      return;
    }

    const dateSuffix = state.currentMeta?.snapshotDates?.length
      ? `_${state.currentMeta.snapshotDates[state.currentMeta.snapshotDates.length - 1]}`
      : '';
    const fileName = `Order_Summary_Report${dateSuffix}.xlsx`;

    ExcelHandler.exportSummaryToExcel(state.tabulatorTable, fileName);
    showToast(`Exported "${fileName}"`, 'success');
  });

  // Export to CSV
  exportCsvBtn.addEventListener('click', () => {
    if (!state.tabulatorTable) {
      showToast('No report data available to export.', 'error');
      return;
    }

    const fileName = `Order_Summary_Report.csv`;
    state.tabulatorTable.download("csv", fileName);
    showToast(`Exported CSV file`, 'success');
  });
}

/**
 * Floating toast notification helper
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

// Start app on DOMContentLoaded
document.addEventListener('DOMContentLoaded', initApp);
