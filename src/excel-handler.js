import * as XLSX from 'xlsx';

/**
 * Normalizes dates from various Excel formats (strings like "YYYY/MM/DD",
 * serial numbers, or Date instances) into "YYYY-MM-DD".
 */
export function normalizeDate(val) {
  if (val === undefined || val === null || val === '') return '';

  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  if (typeof val === 'number') {
    // Excel serial number (days since 1899-12-30)
    const date = new Date(Math.round((val - 25569) * 86400 * 1000));
    if (!isNaN(date.getTime())) {
      const y = date.getUTCFullYear();
      const m = String(date.getUTCMonth() + 1).padStart(2, '0');
      const d = String(date.getUTCDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
    return String(val);
  }

  const str = String(val).trim();
  if (!str) return '';

  // Handle YYYY/MM/DD or YYYY.MM.DD or YYYY-MM-DD
  const match = str.match(/^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})/);
  if (match) {
    const y = match[1];
    const m = match[2].padStart(2, '0');
    const d = match[3].padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  return str;
}

/**
 * Extracts 10-digit order numbers starting with '45' from text
 */
export function extractOrderNumbers(val) {
  if (!val) return [];
  const str = String(val);
  const matches = str.match(/\b(45\d{8})\b/g);
  return matches ? Array.from(new Set(matches)) : [];
}

/**
 * Extracts date from filename formatted like "Inbound_Report_JP_YYYY-MM-DD.xlsx"
 */
export function extractSnapshotDate(filename, fallbackDate) {
  const match = filename.match(/\b(\d{4}-\d{2}-\d{2})\b/);
  if (match) return match[1];
  return fallbackDate || 'Unknown Date';
}

/**
 * Core ExcelHandler for Inbound Reports
 */
export const ExcelHandler = {
  /**
   * Reads a single file as ArrayBuffer and parses with SheetJS
   */
  async readWorkbookFromFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, {
            type: 'array',
            cellDates: true,
            dateNF: 'yyyy-mm-dd',
            raw: false
          });

          const snapshotDate = extractSnapshotDate(
            file.name,
            normalizeDate(new Date(file.lastModified || Date.now()))
          );

          resolve({
            file,
            fileName: file.name,
            fileSize: file.size,
            snapshotDate,
            workbook
          });
        } catch (error) {
          reject(new Error(`Failed to parse "${file.name}": ${error.message}`));
        }
      };

      reader.onerror = () => reject(new Error(`Failed to read "${file.name}"`));
      reader.readAsArrayBuffer(file);
    });
  },

  /**
   * Processes a list of uploaded daily reports in chronological order
   * and generates the unified Order Summary with change history.
   *
   * @param {File[]} files
   * @param {(statusText: string, progressPct: number) => void} [onProgress]
   * @returns {Promise<{ columns: Array, data: Array, meta: Object }>}
   */
  async processMultipleReports(files, onProgress = () => {}) {
    if (!files || files.length === 0) {
      throw new Error('No files provided to process.');
    }

    onProgress('Reading files...', 10);

    // Read all workbooks
    const workbooks = [];
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      onProgress(`Reading ${file.name} (${i + 1}/${files.length})...`, 10 + Math.floor((i / files.length) * 40));
      const wbData = await this.readWorkbookFromFile(file);
      workbooks.push(wbData);
    }

    // Sort chronologically by snapshot date
    workbooks.sort((a, b) => a.snapshotDate.localeCompare(b.snapshotDate));

    onProgress('Aggregating orders and tracking changes...', 60);

    // Track orders across snapshots
    const orderMap = new Map();
    const snapshotDates = workbooks.map(wb => wb.snapshotDate);

    // Key fields to track for change history
    const trackedFields = [
      { key: 'containerStatus', label: 'Latest Container Status' },
      { key: 'containerNumber', label: 'Container Number' },
      { key: 'lastLegETA', label: 'Last Leg ETA' },
      { key: 'lastLegATA', label: 'Last Leg ATA' },
      { key: 'firstLegETD', label: 'First Leg ETD' },
      { key: 'firstLegATD', label: 'First Leg ATD' },
      { key: 'firstLegVessel', label: 'First Leg Vessel' },
      { key: 'secondLegVessel', label: 'Second Leg Vessel' },
      { key: 'thirdLegVessel', label: 'Third Leg Vessel' },
      { key: 'route', label: 'Route' }
    ];

    for (let w = 0; w < workbooks.length; w++) {
      const { fileName, snapshotDate, workbook } = workbooks[w];
      const sheetName = workbook.SheetNames.includes('Export')
        ? 'Export'
        : workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      if (!worksheet) continue;

      const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
      if (!rawRows || rawRows.length < 2) continue;

      const headers = rawRows[0].map(h => String(h).trim());
      const getIdx = (name) => headers.indexOf(name);

      // Find ref no column by header substring, fallback to index 7 (Column H)
      let idxH = headers.findIndex(h => h.toLowerCase().includes('ref no'));
      if (idxH === -1) idxH = 7;

      const idxO = getIdx('Port of Origin Name');
      const idxD = getIdx('Carrier Name');
      const idxE = getIdx('Ocean Carrier Name');
      const idxF = getIdx('MBL No');
      const idxAS = getIdx('Route');
      const idxX = getIdx('Container Number');
      const idxAP = getIdx('Latest Container Status');
      const idxAW = getIdx('First Leg ETD');
      const idxAX = getIdx('First Leg ATD');
      const idxAT = getIdx('First Leg Vessel');
      const idxBA = getIdx('Second Leg Vessel');
      const idxBH = getIdx('Third Leg Vessel');
      const idxAG = getIdx('Last Leg ETA');
      const idxAH = getIdx('Last Leg ATA');
      const idxJ = getIdx('PO Number (Allocated to Container)');
      const idxI = getIdx('PO Number (Booking Stage)');

      for (let r = 1; r < rawRows.length; r++) {
        const row = rawRows[r];

        // First check allocated POs (Col J), fallback to booking stage POs (Col I)
        let pos = extractOrderNumbers(row[idxJ]);
        if (!pos.length) {
          pos = extractOrderNumbers(row[idxI]);
        }
        if (!pos.length) continue;

        const record = {
          forwarderRef: String(row[idxH] || '').trim(),
          originPort: String(row[idxO] || '').trim(),
          carrierName: String(row[idxD] || '').trim(),
          oceanCarrierName: String(row[idxE] || '').trim(),
          mblNo: String(row[idxF] || '').trim(),
          route: String(row[idxAS] || '').trim(),
          containerNumber: String(row[idxX] || '').trim(),
          containerStatus: String(row[idxAP] || '').trim(),
          firstLegETD: normalizeDate(row[idxAW]),
          firstLegATD: normalizeDate(row[idxAX]),
          firstLegVessel: String(row[idxAT] || '').trim(),
          secondLegVessel: String(row[idxBA] || '').trim(),
          thirdLegVessel: String(row[idxBH] || '').trim(),
          lastLegETA: normalizeDate(row[idxAG]),
          lastLegATA: normalizeDate(row[idxAH])
        };

        pos.forEach(po => {
          if (!orderMap.has(po)) {
            orderMap.set(po, {
              orderNumber: po,
              firstSeenDate: snapshotDate,
              lastSeenDate: snapshotDate,
              latest: { ...record },
              historyEvents: []
            });
          } else {
            const entry = orderMap.get(po);
            const prev = entry.latest;
            const changes = [];

            trackedFields.forEach(({ key, label }) => {
              const oldVal = prev[key] || '';
              const newVal = record[key] || '';
              if (oldVal !== newVal) {
                changes.push(`${label} ${newVal || '(empty)'}`);
              }
            });

            if (changes.length > 0) {
              entry.historyEvents.push(`${snapshotDate}: ${changes.join(', ')}`);
            }

            // Update latest known state
            entry.latest = { ...record };
            entry.lastSeenDate = snapshotDate;
          }
        });
      }
    }

    onProgress('Building summary rows...', 90);

    // Format output rows matching Columns A through Q
    const summaryData = [];
    let ordersWithChanges = 0;

    for (const [po, entry] of orderMap.entries()) {
      const { latest, historyEvents } = entry;
      const historyStr = historyEvents.length > 0
        ? historyEvents.join('. ') + '.'
        : 'No changes recorded';

      if (historyEvents.length > 0) {
        ordersWithChanges++;
      }

      summaryData.push({
        _id: po,
        orderNumber: po,                           // Col A
        forwarderRef: latest.forwarderRef,         // Col B
        originPort: latest.originPort,             // Col C
        carrierName: latest.carrierName,           // Col D
        oceanCarrierName: latest.oceanCarrierName, // Col E
        mblNo: latest.mblNo,                       // Col F
        route: latest.route,                       // Col G
        containerNumber: latest.containerNumber,   // Col H
        containerStatus: latest.containerStatus,   // Col I
        firstLegETD: latest.firstLegETD,           // Col J
        firstLegATD: latest.firstLegATD,           // Col K
        firstLegVessel: latest.firstLegVessel,     // Col L
        secondLegVessel: latest.secondLegVessel,   // Col M
        thirdLegVessel: latest.thirdLegVessel,     // Col N
        lastLegETA: latest.lastLegETA,             // Col O
        lastLegATA: latest.lastLegATA,             // Col P
        changeHistory: historyStr,                 // Col Q
        hasChanges: historyEvents.length > 0
      });
    }

    // Default sort by Order Number ascending
    summaryData.sort((a, b) => a.orderNumber.localeCompare(b.orderNumber));

    // Tabulator column definitions matching columns A-Q
    const columns = [
      {
        formatter: "rowSelection",
        titleFormatter: "rowSelection",
        hozAlign: "center",
        headerSort: false,
        width: 40,
        frozen: true,
        cellClick: (e, cell) => cell.getRow().toggleSelect()
      },
      {
        title: "Order Number",
        field: "orderNumber",
        frozen: true,
        width: 130,
        headerFilter: "input",
        headerFilterPlaceholder: "Filter PO...",
        formatter: (cell) => `<span class="font-mono font-semibold text-blue-700">${cell.getValue()}</span>`
      },
      {
        title: "Forwarder Ref No",
        field: "forwarderRef",
        width: 140,
        headerFilter: "input",
        headerFilterPlaceholder: "Filter Ref...",
        formatter: (cell) => `<span class="font-mono text-slate-700">${cell.getValue() || '-'}</span>`
      },
      {
        title: "Port of Origin Name",
        field: "originPort",
        width: 160,
        headerFilter: "input",
        headerFilterPlaceholder: "Filter Port..."
      },
      {
        title: "Carrier Name",
        field: "carrierName",
        width: 130,
        headerFilter: "input"
      },
      {
        title: "Ocean Carrier Name",
        field: "oceanCarrierName",
        width: 160,
        headerFilter: "input"
      },
      {
        title: "MBL No",
        field: "mblNo",
        width: 140,
        headerFilter: "input",
        formatter: (cell) => `<span class="font-mono text-xs">${cell.getValue() || '-'}</span>`
      },
      {
        title: "Route",
        field: "route",
        width: 150,
        headerFilter: "input"
      },
      {
        title: "Container Number",
        field: "containerNumber",
        width: 150,
        headerFilter: "input",
        headerFilterPlaceholder: "Filter Cont...",
        formatter: (cell) => {
          const val = cell.getValue();
          return val
            ? `<span class="font-mono font-medium text-emerald-800 bg-emerald-50 px-1.5 py-0.5 rounded">${val}</span>`
            : `<span class="text-slate-400 italic">Not Assigned</span>`;
        }
      },
      {
        title: "Latest Container Status",
        field: "containerStatus",
        width: 190,
        headerFilter: "input",
        headerFilterPlaceholder: "Filter Status...",
        formatter: (cell) => {
          const val = cell.getValue() || 'Unknown';
          let color = 'bg-slate-100 text-slate-700 border-slate-200';
          if (val.includes('Delivered') || val.includes('Empty Container Returned')) {
            color = 'bg-emerald-50 text-emerald-700 border-emerald-200';
          } else if (val.includes('Departed') || val.includes('Discharged')) {
            color = 'bg-blue-50 text-blue-700 border-blue-200';
          } else if (val.includes('Booked') || val.includes('Picked Up') || val.includes('Gated-In')) {
            color = 'bg-amber-50 text-amber-700 border-amber-200';
          } else if (val.includes('Booking Received')) {
            color = 'bg-purple-50 text-purple-700 border-purple-200';
          }
          return `<span class="inline-block text-xs font-medium px-2 py-0.5 rounded-full border ${color}">${val}</span>`;
        }
      },
      {
        title: "First Leg ETD",
        field: "firstLegETD",
        width: 120,
        sorter: "date",
        hozAlign: "center",
        headerFilter: "input"
      },
      {
        title: "First Leg ATD",
        field: "firstLegATD",
        width: 120,
        sorter: "date",
        hozAlign: "center",
        headerFilter: "input"
      },
      {
        title: "First Leg Vessel",
        field: "firstLegVessel",
        width: 160,
        headerFilter: "input"
      },
      {
        title: "Second Leg Vessel",
        field: "secondLegVessel",
        width: 160,
        headerFilter: "input"
      },
      {
        title: "Third Leg Vessel",
        field: "thirdLegVessel",
        width: 160,
        headerFilter: "input"
      },
      {
        title: "Last Leg ETA",
        field: "lastLegETA",
        width: 120,
        sorter: "date",
        hozAlign: "center",
        headerFilter: "input",
        formatter: (cell) => {
          const val = cell.getValue();
          return val ? `<span class="font-semibold text-slate-800">${val}</span>` : '-';
        }
      },
      {
        title: "Last Leg ATA",
        field: "lastLegATA",
        width: 120,
        sorter: "date",
        hozAlign: "center",
        headerFilter: "input"
      },
      {
        title: "Change History",
        field: "changeHistory",
        minWidth: 320,
        formatter: (cell) => {
          const val = cell.getValue();
          if (!val || val === 'No changes recorded') {
            return `<span class="text-slate-400 text-xs italic">No changes recorded</span>`;
          }
          return `<div class="text-xs text-slate-700 leading-snug py-1">${val}</div>`;
        }
      }
    ];

    onProgress('Done', 100);

    return {
      columns,
      data: summaryData,
      meta: {
        totalOrders: summaryData.length,
        ordersWithChanges,
        filesCount: workbooks.length,
        snapshotDates: snapshotDates,
        dateRange: `${snapshotDates[0]} to ${snapshotDates[snapshotDates.length - 1]}`
      }
    };
  },

  /**
   * Export the summary table data directly to an Excel (.xlsx) file
   * mapping column keys to human-readable headers (Columns A to Q).
   */
  exportSummaryToExcel(tabulatorTable, fileName = "Order_Summary_Report.xlsx") {
    if (!tabulatorTable) return;

    const rawData = tabulatorTable.getData();
    if (!rawData || rawData.length === 0) {
      alert("No data available to export.");
      return;
    }

    const columnOrder = [
      { field: 'orderNumber', title: 'Order Number' },
      { field: 'forwarderRef', title: 'Forwarder Ref No' },
      { field: 'originPort', title: 'Port of Origin Name' },
      { field: 'carrierName', title: 'Carrier Name' },
      { field: 'oceanCarrierName', title: 'Ocean Carrier Name' },
      { field: 'mblNo', title: 'MBL No' },
      { field: 'route', title: 'Route' },
      { field: 'containerNumber', title: 'Container Number' },
      { field: 'containerStatus', title: 'Latest Container Status' },
      { field: 'firstLegETD', title: 'First Leg ETD' },
      { field: 'firstLegATD', title: 'First Leg ATD' },
      { field: 'firstLegVessel', title: 'First Leg Vessel' },
      { field: 'secondLegVessel', title: 'Second Leg Vessel' },
      { field: 'thirdLegVessel', title: 'Third Leg Vessel' },
      { field: 'lastLegETA', title: 'Last Leg ETA' },
      { field: 'lastLegATA', title: 'Last Leg ATA' },
      { field: 'changeHistory', title: 'Change History' }
    ];

    const cleanData = rawData.map(row => {
      const item = {};
      columnOrder.forEach(col => {
        item[col.title] = row[col.field] !== undefined ? row[col.field] : '';
      });
      return item;
    });

    const worksheet = XLSX.utils.json_to_sheet(cleanData);

    // Set auto column widths for Excel
    worksheet['!cols'] = columnOrder.map(col => {
      let maxLen = col.title.length;
      cleanData.forEach(row => {
        const valStr = String(row[col.title] || '');
        if (valStr.length > maxLen) maxLen = Math.min(valStr.length, 60);
      });
      return { wch: Math.max(maxLen + 3, 12) };
    });

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Order Summary");

    let targetFileName = fileName.trim();
    if (!targetFileName.toLowerCase().endsWith('.xlsx')) {
      targetFileName += '.xlsx';
    }

    XLSX.writeFile(workbook, targetFileName);
  },

  /**
   * Generates a multi-day synthetic snapshot test dataset representing
   * daily snapshots for orders with status & ETA transitions.
   */
  generateSampleSnapshots() {
    const dates = ['2026-10-10', '2026-10-12', '2026-10-15', '2026-10-17'];
    const dummyHeaders = [
      "Lane ID", "Customer Business Unit", "POL-POD Lane Pair", "Carrier Name", "Ocean Carrier Name",
      "MBL No", "Ocean Carrier Ref No", "Forwarder Ref No", "PO Number (Booking Stage)",
      "PO Number (Allocated to Container)", "Customer Booking Received Date",
      "Carrier Booking Request Confirmation Date", "Cargo Ready Date", "Port of Origin Code",
      "Port of Origin Name", "Country of Origin Code", "Country of Origin Name",
      "Port of Destination Code", "Port of Destination Name", "Country of Destination Code",
      "Country of Destination Name", "Consignor Name", "Consignee Name", "Container Number",
      "MDLZ Equipment Type", "Planned Export Full Container Pickup Date",
      "Actual Export Full Container Pickup Date", "Vessel", "Baseline ETD", "1st Leg ETD",
      "1st Leg ATD", "Baseline ETA", "Last Leg ETA", "Last Leg ATA",
      "Carrier Booking Confirmation Upload Date", "Packing List Upload Date",
      "Commercial Invoice Upload Date", "Export Customs Release Upload Date",
      "Final BL Upload Date", "Import Customs Release Upload Date", "On-Time Transit Status",
      "Latest Container Status", "Exception Case", "Job No", "Route", "First Leg Vessel",
      "First Leg POL", "First Leg POD", "First Leg ETD", "First Leg ATD", "First Leg ETA",
      "First Leg ATA", "Second Leg Vessel", "Second Leg POL", "Second Leg POD",
      "Second Leg ETD", "Second Leg ATD", "Second Leg ETA", "Second Leg ATA",
      "Third Leg Vessel", "Third Leg POL", "Third Leg POD", "Third Leg ETD",
      "Third Leg ATD", "Third Leg ETA", "Third Leg ATA", "BL No"
    ];

    const files = dates.map((dateStr, dIndex) => {
      const rows = [dummyHeaders];

      // Order 1: ETA changes from 10/25 to 10/28 to 11/02
      const eta1 = dIndex === 0 ? '2026/10/25' : dIndex === 1 ? '2026/10/28' : '2026/11/02';
      const status1 = dIndex < 2 ? 'Departed' : 'Discharged from Vessel';
      const vessel1 = dIndex < 2 ? 'MAERSK DHAKA' : 'ONE HONOLULU';
      const r1 = new Array(dummyHeaders.length).fill('');
      r1[3] = 'Global Freight Logistics'; r1[4] = 'Maersk Line'; r1[5] = 'MEDU1029384';
      r1[7] = 'VEN-328901'; r1[8] = '4505127151'; r1[9] = '4505127151';
      r1[14] = 'Shanghai Port'; r1[23] = 'MSKU9281720'; r1[32] = eta1;
      r1[41] = status1; r1[44] = 'SHA - TYO'; r1[45] = vessel1;
      r1[48] = '2026/10/05'; r1[49] = '2026/10/06';
      rows.push(r1);

      // Order 2: Multiple POs allocated into one container, status progresses
      const status2 = dIndex === 0 ? 'Container Booked' : dIndex === 1 ? 'Picked Up to CY' : 'Departed';
      const container2 = dIndex === 0 ? '' : 'ONEU6213566';
      const r2 = new Array(dummyHeaders.length).fill('');
      r2[3] = 'Global Freight Logistics'; r2[4] = 'ONE Network'; r2[5] = 'ONEY9823411';
      r2[7] = 'VEN-328902'; r2[8] = '4505127158, 4505127167'; r2[9] = '4505127158, 4505127167';
      r2[14] = 'Busan Port'; r2[23] = container2; r2[32] = '2026/10/20';
      r2[41] = status2; r2[44] = 'PUS - YOK'; r2[45] = 'ONE MILLAU';
      r2[48] = '2026/10/12';
      rows.push(r2);

      // Order 3: Delivered order
      const status3 = dIndex >= 2 ? 'Delivered' : 'Arrived';
      const ata3 = dIndex >= 2 ? '2026/10/14' : '';
      const r3 = new Array(dummyHeaders.length).fill('');
      r3[3] = 'Global Freight Logistics'; r3[4] = 'CMA CGM'; r3[5] = 'CMAU3819201';
      r3[7] = 'VEN-328850'; r3[8] = '4505117511'; r3[9] = '4505117511';
      r3[14] = 'Singapore'; r3[23] = 'CMAU8192012'; r3[32] = '2026/10/12'; r3[33] = ata3;
      r3[41] = status3; r3[44] = 'SIN - KOB'; r3[45] = 'CMA CGM CHENNAI';
      r3[48] = '2026/09/28'; r3[49] = '2026/09/29';
      rows.push(r3);

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, 'Export');
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

      return new File([wbout], `Inbound_Report_JP_${dateStr}.xlsx`, {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
    });

    return files;
  }
};
