# doko_maker 🚅📦📊

An in-browser Single Page Application (SPA) designed to parse daily inbound ocean container transportation reports (Hellmann Inbound Reports), extract orders, display them in an interactive data grid, track schedule/vessel/ETA changes over time, and export unified summary reports to Excel (.xlsx) or CSV.

Built with **Vite**, **Tailwind CSS**, **Tabulator**, and **SheetJS (xlsx)**, configured for automated CI/CD deployment to **GitHub Pages** via **GitHub Actions**.

---

## ✨ Features

- 📁 **Zero-Server Processing (100% Client-Side)**: All processing runs locally in the browser via SheetJS. Sensitive order, container, and freight data never leaves your computer.
- 📅 **Multi-Snapshot Time-Series Ingestion**: Upload 10, 30, or 50+ daily reports (`Inbound_Report_JP_YYYY-MM-DD.xlsx`). The app sorts them chronologically and traces each order across the entire timeline.
- 🔍 **Order Extraction & Allocation**:
  - Automatically identifies 10-digit order numbers (`450xxxxxxx`) from container allocation (Column J) or booking stages (Column I).
  - Handles multiple orders grouped under a single Hellmann Reference or container.
- 🕒 **Chronological Change History (Column Q)**:
  - Detects and logs changes between snapshot dates for:
    - `Latest Container Status`
    - `Container Number`
    - `Last Leg ETA` and `Last Leg ATA`
    - `First Leg ETD` and `First Leg ATD`
    - `First Leg Vessel`, `Second Leg Vessel`, `Third Leg Vessel`
    - `Route`
  - Formats a clear log for each order, for example:  
    `2026-10-13: Container Number UETU7214088, Latest Container Status Container Booked. 2026-10-14: Latest Container Status Picked Up to CY.`
- 📊 **Interactive Data Grid (Tabulator)**:
  - Instant live filter/search across all columns
  - Filter by container status or toggle "Only Orders with Changes"
  - Column sorting, resizing, reordering, and row selection
- 📥 **Export to Excel & CSV**:
  - Download the clean, formatted 17-column summary report as `.xlsx` or `.csv`.
- 🧪 **Built-in Sample Data**: Test immediately with 4-day synthetic snapshot data directly in the browser.

---

## 📋 Summary Report Structure (Columns A – Q)

| Output Col | Output Field | Origin Col & Header |
| :--- | :--- | :--- |
| **A** | **Order Number** | Col J (`PO Number (Allocated)`) / Col I (`PO Number (Booking)`) |
| **B** | **Hellmann Ref No** | Col H (`Hellmann Ref No`) |
| **C** | **Port of Origin Name** | Col O (`Port of Origin Name`) |
| **D** | **Carrier Name** | Col D (`Carrier Name`) |
| **E** | **Ocean Carrier Name** | Col E (`Ocean Carrier Name`) |
| **F** | **MBL No** | Col F (`MBL No`) |
| **G** | **Route** | Col AS (`Route`) |
| **H** | **Container Number** | Col X (`Container Number`) |
| **I** | **Latest Container Status** | Col AP (`Latest Container Status`) |
| **J** | **First Leg ETD** | Col AW (`First Leg ETD`) *(standardized as `YYYY-MM-DD`)* |
| **K** | **First Leg ATD** | Col AX (`First Leg ATD`) *(standardized as `YYYY-MM-DD`)* |
| **L** | **First Leg Vessel** | Col AT (`First Leg Vessel`) |
| **M** | **Second Leg Vessel** | Col BA (`Second Leg Vessel`) |
| **N** | **Third Leg Vessel** | Col BH (`Third Leg Vessel`) |
| **O** | **Last Leg ETA** | Col AG (`Last Leg ETA`) *(standardized as `YYYY-MM-DD`)* |
| **P** | **Last Leg ATA** | Col AH (`Last Leg ATA`) *(standardized as `YYYY-MM-DD`)* |
| **Q** | **Change History** | Log of changes across snapshot files |

---

## 🚀 GitHub Pages Setup

This repository includes an automated GitHub Actions workflow ([`.github/workflows/deploy.yml`](.github/workflows/deploy.yml)).

### To activate GitHub Pages:

1. Push your code to GitHub:
   ```bash
   git add .
   git commit -m "Implement Hellmann Inbound report parsing and change history"
   git push origin main
   ```
2. In your GitHub repository:
   - Go to **Settings** ➔ **Pages**.
   - Under **Build and deployment > Source**, choose **GitHub Actions**.
3. Every time you push to the `main` branch, GitHub Actions will automatically install dependencies, build the Vite app, and publish it to:
   ```
   https://<username>.github.io/doko_maker/
   ```

---

## 💻 Local Development

### 1. Install Dependencies
```bash
npm install
```

### 2. Start Vite Dev Server
```bash
npm run dev
```
Open `http://localhost:5173` in your browser.

### 3. Build for Production
```bash
npm run build
```

---

## 📂 Project Structure

```
doko_maker/
├── .github/
│   └── workflows/
│       └── deploy.yml   # GitHub Actions CI/CD workflow
├── src/
│   ├── excel-handler.js # Hellmann report parser, date normalizer & export
│   ├── main.js          # App lifecycle, Tabulator grid, filters & metrics
│   └── style.css        # Tailwind directives & Tabulator styling
├── index.html           # SPA entry point
├── package.json         # Dependencies & scripts
├── postcss.config.js    # PostCSS configuration
├── tailwind.config.js   # Tailwind CSS configuration
├── vite.config.js       # Vite bundler configuration (relative base for Pages)
├── .gitignore
├── LICENSE
└── README.md
```

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
