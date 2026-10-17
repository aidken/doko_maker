# doko_maker 🚅📊

An in-browser Single Page Application (SPA) designed to parse Excel transportation and expense reports, display them in an interactive spreadsheet-like data grid, allow real-time inline editing, and export reports back to Excel (.xlsx) or CSV.

Built with **Vite**, **Tailwind CSS**, **Tabulator**, and **SheetJS (xlsx)**, configured for automated CI/CD deployment to **GitHub Pages** via **GitHub Actions**.

---

## ✨ Features

- 📁 **Zero-Server Processing (100% Client-Side)**: Files are processed directly in the browser with SheetJS. Sensitive financial or transportation data never leaves your device.
- ⚡ **Bundled with Vite**: Blazing-fast development and optimized production builds.
- 🤖 **GitHub Actions CI/CD**: Automatic building and deployment to GitHub Pages upon pushing to `main`.
- 📤 **Drag-and-Drop Ingestion**: Upload `.xlsx`, `.xls`, or `.csv` files effortlessly.
- 📑 **Multi-Sheet Navigation**: Workbooks with multiple worksheets automatically display tabs for quick sheet switching.
- 📊 **Interactive Data Grid (Tabulator)**:
  - Double-click any cell to edit inline
  - Instant live filter/search across all columns
  - Sort by any column
  - Drag-and-drop column reordering and resizing
  - Row selection and bulk row deletion
  - Add new rows directly into the sheet
  - Multi-step Undo & Redo history
- 📥 **Export to Excel & CSV**:
  - Export current data back into `.xlsx` or `.csv` with one click
- 🧪 **Built-in Sample Data**: Click "Sample Data" or "Load Sample Report" to instantly test features with a realistic transportation expense dataset.

---

## 🚀 GitHub Pages Setup

This repository includes an automated GitHub Actions workflow (`.github/workflows/deploy.yml`).

### To activate GitHub Pages:

1. Push your code to GitHub:
   ```bash
   git add .
   git commit -m "Configure Vite and GitHub Actions deployment"
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
Open the provided local URL (typically `http://localhost:5173`) in your browser.

### 3. Build for Production
```bash
npm run build
```
The compiled output will be generated in the `dist/` directory.

### 4. Preview Production Build
```bash
npm run preview
```

---

## 📂 Project Structure

```
doko_maker/
├── .github/
│   └── workflows/
│       └── deploy.yml   # GitHub Actions CI/CD workflow
├── src/
│   ├── excel-handler.js # SheetJS parsing, data transformation & export
│   ├── main.js          # Tabulator grid initialization & app event logic
│   └── style.css        # Tailwind directives & Tabulator custom styling
├── index.html           # SPA entry point
├── package.json         # Project dependencies and build scripts
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
