# Agent Recovery Specification: Institutional Tracker

## Purpose
This document is specifically formulated for AI Agent ingestion. Reading this document should provide the agent with all necessary context, constraints, and structural blueprints required to fully recreate the `institutional-tracker` project from scratch, ensuring identical functionality, WAF evasion mechanisms, API contracts, and UI/UX design.

---

## 1. Project Architecture & Tech Stack
*   **Language**: Python 3.10+
*   **Web Framework**: Flask
*   **Data Processing**: Pandas
*   **Excel Generation**: openpyxl
*   **Network Requests**: Requests (Crucial for bypassing TPEX WAF/403 errors, replacing native `urllib`)
*   **Frontend**: HTML5, Vanilla JavaScript (ES8+ async/await), CSS3, Bootstrap 5 (Dark theme)
*   **Deployment**: Render (`gunicorn`, defined via `render.yaml`), GitHub Actions (Scheduled CRON execution)

## 2. Core Components & Logic

### A. Data Crawler & Analyzer (`analyze.py`)
*   **Role**: Standalone CLI script and internal module to fetch stock institutional data from TWSE (上市) and TPEX (上櫃), process it, and generate an Excel report.
*   **Input**: Arguments passed via `sys.argv[1]` as target date (format: `YYYYMMDD`). If none provided, defaults to current date.
*   **WAF Bypass Implementation (CRITICAL)**:
    *   TWSE endpoints: `https://www.twse.com.tw/fund/T86`, `https://www.twse.com.tw/exchangeReport/MI_INDEX`.
    *   TPEX endpoints: `https://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php`, `https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php`. TPEX requires Taiwan Republican calendar format (e.g., 2026 = 115).
    *   **Bypass Rule**: Must use `requests.Session()` with explicit headers including `User-Agent` (Windows Chrome pattern), `Accept`, `Connection`, and **`Referer: https://www.tpex.org.tw/`**. Failure to do so will result in 403 Forbidden.
*   **Data Filtration & Metrics**:
    *   Exclude ETFs and Warrants: Drop rows where Stock ID length > 4 or does not start with a digit.
    *   Calculate `VWAP` (Volume Weighted Average Price) = Total Transaction Value / Total Volume. Fallback to Close Price if missing.
    *   Compute `foreign_val` (Foreign Institutional estimated value) and `it_val` (Investment Trust estimated value) by multiplying net buy/sell shares with `VWAP`.
*   **Report Generation (`openpyxl`)**:
    *   Output filename: `market_analysis_YYYYMMDD.xlsx`.
    *   Contains two sheets: "上市" and "上櫃".
    *   Data is sorted dynamically by absolute valuation (`abs(foreign_val)`) in descending order.
    *   Stock IDs must be cast to `int` before writing to cells to prevent "Number stored as text" Excel warnings.
    *   **Styling**: Highlighting logic based on institutional cooperation/opposition (Red hues for same-direction, Green hues for opposite-direction based on share volume dominance).

### B. Web Backend (`app.py`)
*   **Role**: Flask application to serve the frontend and provide API endpoints to the generated reports.
*   **Security (Basic Authentication)**:
    *   Reads environment variables `USE_AUTH` (boolean string), `AUTH_USER`, and `AUTH_PASS`.
    *   If `USE_AUTH` is true, an `@app.before_request` hook validates the `Authorization` header. If missing or incorrect, it returns a 401 response with `WWW-Authenticate: Basic realm="VIP Login Required"`.
*   **Endpoints**:
    *   `GET /`: Serves `templates/index.html`.
    *   `GET /get_available_dates`: Scans local directory for `market_analysis_*.xlsx` files, extracts dates, and returns them as a sorted JSON array (newest first).
    *   `GET /get_report/<date>`: Uses `pandas` to read the specific Excel file (both sheets), formatting the layout (skipping the main date header, aligning headers, and mapping rows into nested JSON arrays).
    *   `GET /download/<date>`: Triggers file download via `send_file`.
    *   `POST /trigger_analysis`: Accepts JSON payload `{ "date": "YYYY-MM-DD" }`. Executes `analyze.py` via `subprocess.run()`. Crucially, it captures `stdout` and `stderr` and returns them in the payload as `debug_log`.

### C. Web Frontend (`templates/index.html`)
*   **Visual Identity**: Professional Trading Terminal concept. Uses `data-bs-theme="dark"`.
*   **DOM Structure**:
    *   **Header**: Flex-wrap enabled for mobile responsiveness using Bootstrap's grid system (`row`, `col-12`, `col-md-auto`). Contains Date Selector, Limit Selector (Top 35 vs All), Sort Rule (Valuation vs Shares), Compact Mode toggle ("👁️ 精簡模式"), Download button, Date Picker for manual fetching, Manual Fetch button ("⚡ 取得台股資料"), and Refresh button.
    *   **Color Legend**: A visual guide explaining the meaning of the four highlighting hues (Deep Red, Light Red, Deep Green, Light Green).
    *   **Tabs**: Bootstrap Nav-pills toggling "上市" (TWSE) and "上櫃" (TPEX) container views.
    *   **Data Containers**: Holds the respective tables, loading spinners, and progress bars. The table structure enforces 7 columns per block (Rank, Ticker, Name, Price, VWAP, Shares, Value).
    *   **Footer**: Contains GitHub Repository redirect (`institutional-tracker`) and Debug Log toggle.
*   **State & Interactivity**:
    *   **Auto-Fetch Logic**: On load, if the expected latest valid trading date (excluding weekends, and shifting to yesterday if current time < 15:00 Taipei time) is missing from available reports, automatically trigger the Fetch Data logic.
    *   **Loading UX**: Simulated progress bar (`setInterval`) displayed while waiting for `fetch('/trigger_analysis')`.
    *   **Compact Mode (`toggleCompactMode`)**: Toggles a `compact-mode` CSS class on `body` to hide price, VWAP, shares, and valuation columns for a streamlined view, automatically adjusting main header `colspans` from 7 to 3 to prevent layout breaking.
    *   **Table Replication**: Frontend must reconstruct the highlighting logic native to the backend (Red/Green hues).
    *   **Hover/Focus Sync**: Mousing over a stock name highlights the entire associated data blocks (same stock ticker across Foreign and IT sections) in striking yellow (`#ffe600` on `#3d3511`).
    *   **Click-to-Scroll**: Clicking a stock name smoothly scrolls the viewport to its counterpart on the opposing grid and triggers a 2-second CSS animation flash (Yellow/Red border) for rapid visual correlation.
    *   **Debug Console**: A hidden `<pre>` tag that pops up containing the `debug_log` returned by `/trigger_analysis` API if an explicit fetch is requested.

## 3. Configuration & CI/CD
*   **Dependencies (`requirements.txt`)**: `flask`, `pandas`, `openpyxl`, `requests`, `gunicorn`.
*   **Render Deployment (`render.yaml`)**:
    *   Uses Python 3.10.0 runtime.
    *   `buildCommand`: `pip install -r requirements.txt && pip install gunicorn`.
    *   `startCommand`: `gunicorn app:app -b 0.0.0.0:$PORT`.
*   **GitHub Actions (`.github/workflows/daily_analysis.yml`)**: Unattended CRON trigger (e.g., `0 8 * * 1-5` for weekdays 16:00 UTC+8) checking out code, running `python analyze.py`, and uploading the `xlsx` as a workflow artifact.

## 4. Execution Directives for Agents
If an agent is instructed to reconstruct this platform:
1.  Initialize the project directory with `.gitignore` (ignore `__pycache__`, `*.xlsx`, `.env`).
2.  Implement `analyze.py` strictly enforcing the TPEX `requests` logic + explicit Headers to prevent fatal initial 403 blocks.
3.  Implement `app.py` ensuring the `/trigger_analysis` route accurately captures and echoes standard outputs for frontend debugging.
4.  Implement `templates/index.html` ensuring structural sync algorithms are in place for the color highlighting mechanics, as Pandas drops color formatting when casting to JSON.
5.  Validate deployment files (`render.yaml`, `requirements.txt`).
