import * as XLSX from 'xlsx';

/**
 * ExcelHandler - Handles reading, parsing, transforming, and exporting Excel workbooks
 * using SheetJS (XLSX).
 */
export const ExcelHandler = {
  /**
   * Reads a File object and parses it into a SheetJS workbook.
   * @param {File} file 
   * @returns {Promise<{workbook: any, sheetNames: string[], fileName: string, fileSize: number}>}
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

          resolve({
            workbook,
            sheetNames: workbook.SheetNames,
            fileName: file.name,
            fileSize: file.size
          });
        } catch (error) {
          reject(new Error(`Failed to parse Excel file: ${error.message}`));
        }
      };

      reader.onerror = () => {
        reject(new Error('Failed to read file from disk.'));
      };

      reader.readAsArrayBuffer(file);
    });
  },

  /**
   * Converts a SheetJS worksheet into Tabulator column definitions and row data.
   * @param {any} worksheet 
   * @returns {{columns: Array, data: Array, totalRows: number, totalCols: number}}
   */
  parseSheetData(worksheet) {
    if (!worksheet || !worksheet['!ref']) {
      return { columns: [], data: [], totalRows: 0, totalCols: 0 };
    }

    // Convert sheet to 2D array of rows
    const rawRows = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: '',
      blankrows: false
    });

    if (!rawRows || rawRows.length === 0) {
      return { columns: [], data: [], totalRows: 0, totalCols: 0 };
    }

    // Identify header row
    const rawHeader = rawRows[0] || [];
    const rawBody = rawRows.slice(1);

    // Sanitize column names and handle blank/duplicate headers
    const usedFieldNames = new Map();
    const columns = [
      {
        formatter: "rowSelection",
        titleFormatter: "rowSelection",
        hozAlign: "center",
        headerSort: false,
        width: 45,
        frozen: true,
        cellClick: function(e, cell) {
          cell.getRow().toggleSelect();
        }
      },
      {
        formatter: "rownum",
        title: "#",
        hozAlign: "center",
        width: 55,
        headerSort: false,
        frozen: true
      }
    ];

    // Determine max columns in any row
    let maxCols = rawHeader.length;
    for (let i = 0; i < Math.min(rawBody.length, 100); i++) {
      if (rawBody[i] && rawBody[i].length > maxCols) {
        maxCols = rawBody[i].length;
      }
    }

    const fieldKeys = [];

    for (let c = 0; c < maxCols; c++) {
      let headerText = rawHeader[c] !== undefined && rawHeader[c] !== null && String(rawHeader[c]).trim() !== ''
        ? String(rawHeader[c]).trim()
        : `Column ${c + 1}`;

      // Ensure unique field name
      let fieldName = headerText;
      let count = usedFieldNames.get(headerText) || 0;
      if (count > 0) {
        fieldName = `${headerText}_${count + 1}`;
      }
      usedFieldNames.set(headerText, count + 1);
      fieldKeys.push({ field: fieldName, title: headerText });

      // Detect sample data types
      let isNumeric = true;
      let sampleCount = 0;

      for (let r = 0; r < Math.min(rawBody.length, 30); r++) {
        const val = rawBody[r] ? rawBody[r][c] : undefined;
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          sampleCount++;
          if (isNaN(Number(String(val).replace(/,/g, '')))) {
            isNumeric = false;
            break;
          }
        }
      }

      columns.push({
        title: headerText,
        field: fieldName,
        editor: "input",
        sorter: (sampleCount > 0 && isNumeric) ? "number" : "string",
        hozAlign: (sampleCount > 0 && isNumeric) ? "right" : "left",
        headerFilter: "input",
        headerFilterPlaceholder: "Filter...",
        resizable: true,
        minWidth: 120
      });
    }

    // Build row objects
    const data = rawBody.map((row, rowIndex) => {
      const rowObj = { _id: rowIndex + 1 };
      for (let c = 0; c < fieldKeys.length; c++) {
        const key = fieldKeys[c].field;
        const val = row ? row[c] : '';
        rowObj[key] = (val !== undefined && val !== null) ? val : '';
      }
      return rowObj;
    });

    return {
      columns,
      data,
      totalRows: data.length,
      totalCols: fieldKeys.length
    };
  },

  /**
   * Generates sample transportation expense report data for quick demo.
   * @returns {{fileName: string, sheetNames: string[], workbook: any}}
   */
  generateSampleWorkbook() {
    const sampleExpenses = [
      {
        "Date": "2026-10-01",
        "Employee": "Kenji Sato",
        "Mode": "Train",
        "Departure": "Tokyo Station",
        "Destination": "Shin-Yokohama",
        "Route": "JR Tokaido Line",
        "Cost (JPY)": 570,
        "Purpose": "Client meeting at ACME Corp",
        "Receipt Attached": "Yes",
        "Status": "Approved"
      },
      {
        "Date": "2026-10-02",
        "Employee": "Yuki Tanaka",
        "Mode": "Subway",
        "Departure": "Shinjuku",
        "Destination": "Roppongi",
        "Route": "Toei Oedo Line",
        "Cost (JPY)": 220,
        "Purpose": "Partner conference",
        "Receipt Attached": "No",
        "Status": "Approved"
      },
      {
        "Date": "2026-10-05",
        "Employee": "Alex Rivera",
        "Mode": "Taxi",
        "Departure": "Shibuya",
        "Destination": "Haneda Airport",
        "Route": "Metropolitan Expressway",
        "Cost (JPY)": 8400,
        "Purpose": "Late night urgent equipment delivery",
        "Receipt Attached": "Yes",
        "Status": "Pending Review"
      },
      {
        "Date": "2026-10-08",
        "Employee": "Kenji Sato",
        "Mode": "Shinkansen",
        "Departure": "Tokyo",
        "Destination": "Nagoya",
        "Route": "Nozomi 215",
        "Cost (JPY)": 11300,
        "Purpose": "Branch quarterly audit",
        "Receipt Attached": "Yes",
        "Status": "Approved"
      },
      {
        "Date": "2026-10-10",
        "Employee": "Mai Takahashi",
        "Mode": "Bus",
        "Departure": "Kyoto Station",
        "Destination": "Arashiyama",
        "Route": "City Bus 28",
        "Cost (JPY)": 230,
        "Purpose": "Regional onsite survey",
        "Receipt Attached": "Yes",
        "Status": "Approved"
      },
      {
        "Date": "2026-10-12",
        "Employee": "Alex Rivera",
        "Mode": "Train",
        "Departure": "Shinagawa",
        "Destination": "Yokohama",
        "Route": "Keikyu Main Line",
        "Cost (JPY)": 320,
        "Purpose": "Technical support onsite",
        "Receipt Attached": "Yes",
        "Status": "Approved"
      },
      {
        "Date": "2026-10-15",
        "Employee": "Yuki Tanaka",
        "Mode": "Train",
        "Departure": "Ikebukuro",
        "Destination": "Omiya",
        "Route": "JR Saikyo Line",
        "Cost (JPY)": 400,
        "Purpose": "Vendor demo showcase",
        "Receipt Attached": "Yes",
        "Status": "Pending Review"
      }
    ];

    const sampleSummary = [
      { "Category": "Train & Subway", "Total Trips": 5, "Total Cost (JPY)": 12740 },
      { "Category": "Taxi", "Total Trips": 1, "Total Cost (JPY)": 8400 },
      { "Category": "Bus", "Total Trips": 1, "Total Cost (JPY)": 230 },
      { "Category": "Grand Total", "Total Trips": 7, "Total Cost (JPY)": 21370 }
    ];

    const wb = XLSX.utils.book_new();
    const wsExpenses = XLSX.utils.json_to_sheet(sampleExpenses);
    const wsSummary = XLSX.utils.json_to_sheet(sampleSummary);

    XLSX.utils.book_append_sheet(wb, wsExpenses, "Transportation Expenses");
    XLSX.utils.book_append_sheet(wb, wsSummary, "Summary Report");

    return {
      fileName: "sample_transportation_report.xlsx",
      fileSize: 4096,
      sheetNames: wb.SheetNames,
      workbook: wb
    };
  },

  /**
   * Export current Tabulator data to an Excel (.xlsx) file.
   * Strips internal metadata columns like row selection and row number.
   * @param {any} tabulatorTable 
   * @param {string} fileName 
   * @param {string} sheetName 
   */
  exportTableToExcel(tabulatorTable, fileName = "report_export.xlsx", sheetName = "Sheet1") {
    if (!tabulatorTable) return;

    // Get current data (includes modifications)
    const rawData = tabulatorTable.getData();
    if (!rawData || rawData.length === 0) {
      alert("No data available to export.");
      return;
    }

    // Get column fields to export (exclude row selection and rownum)
    const cols = tabulatorTable.getColumns();
    const exportableCols = cols.filter(col => {
      const def = col.getDefinition();
      return def.field && !def.formatter?.includes('row');
    });

    // Format clean records
    const cleanData = rawData.map(row => {
      const item = {};
      exportableCols.forEach(col => {
        const def = col.getDefinition();
        const fieldName = def.field;
        const title = def.title || fieldName;
        item[title] = row[fieldName] !== undefined ? row[fieldName] : '';
      });
      return item;
    });

    const worksheet = XLSX.utils.json_to_sheet(cleanData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName.substring(0, 31));

    // Ensure valid filename extension
    let targetFileName = fileName.trim();
    if (!targetFileName.toLowerCase().endsWith('.xlsx')) {
      targetFileName += '.xlsx';
    }

    XLSX.writeFile(workbook, targetFileName);
  },

  /**
   * Format file sizes into human-readable strings (e.g. 1.2 MB).
   * @param {number} bytes 
   * @returns {string}
   */
  formatBytes(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }
};

