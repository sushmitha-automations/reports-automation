# Workforce & Onboarding Automation Platform

A high-performance Python analytics and reporting platform built with **Polars**, **OpenPyXL**, and **Streamlit** to generate executive-grade **Attendance Adherence** and **Onboarding Compliance** workbooks.

---

## 🚀 Key Features

- **Blazing Fast Polars Engine**: Processes 85,000+ records in seconds with zero Pandas dependencies.
- **Strict Report Fidelity**: Produces 11-sheet Attendance and 9-sheet Onboarding Excel reports with exact color palettes, fonts (`Calibri 11pt` / `Aptos Narrow 11pt`), live formulas (`COUNTIFS`, `SUM`, `XLOOKUP`), and text-formatting (`@`) for Aadhaar and Mobile numbers.
- **3-Way Identity Resolution Waterfall**: Resolves employee identity across Aadhaar $\rightarrow$ Mobile $\rightarrow$ Numeric ID.
- **Streamlit Web Interface**: Multi-file drag-and-drop upload, automatic base-file validator, real-time progress callbacks, and persistent multi-file downloads.
- **Zero-Dependency Launcher**: One-click bootstrapping with `runme.bat`.

---

## 📁 Repository Structure

```text
├── app.py                  # Streamlit web application & UI
├── orchestrator.py        # Pipeline coordinator & execution workflow
├── loaders.py             # High-speed Polars ingestion routines
├── adherence_engine.py    # 3-Way identity matching & Chroma matrix evaluation
├── onboarding_engine.py   # KYC Aadhaar compliance & ageing calculation
├── excel_builder.py       # OpenPyXL workbook formatting engine with live formulas
├── config.py              # Facility POD mappings, color palettes, and styling rules
├── requirements.txt       # Dependencies (polars, pyarrow, openpyxl, streamlit)
├── runme.bat              # Self-bootstrapping Windows launcher script
├── .streamlit/
│   └── config.toml        # Streamlit server and 4GB upload configuration
└── .gitignore             # Excludes raw data CSVs, Excel files, and cache
```

---

## 🛠️ Local Setup & Execution

### Option A: One-Click Windows Launcher (Recommended)
Simply double-click:
```bat
runme.bat
```
The script will automatically detect Python, install required dependencies, and launch the web interface in your default browser.

### Option B: Manual Command Line Execution
1. Ensure Python 3.10+ is installed.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the Streamlit web application:
   ```bash
   streamlit run app.py
   ```

---

## ☁️ Deployment & Hosting Options

### 1. Streamlit Community Cloud (Free & Instant)
1. Push this repository to **GitHub**.
2. Visit [share.streamlit.io](https://share.streamlit.io).
3. Connect your GitHub account and select this repository.
4. Set Main file path to: `app.py`.
5. Click **Deploy**. Your app will have a live public or private URL!

### 2. Internal Company Server / Cloud VM (AWS / Azure / GCP / Docker)
Run the application with:
```bash
streamlit run app.py --server.port 8501 --server.address 0.0.0.0
```
