# Institutional Tracker: 台灣股市法人資金動向高頻量化追蹤系統

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Flask](https://img.shields.io/badge/flask-app-green.svg)](#)
[![Deployment](https://img.shields.io/badge/deployment-render-purple.svg)](#)

👉 **Live Demo**: [https://institutional-tracker.onrender.com](https://institutional-tracker.onrender.com)

---

## 摘要 (Abstract)

本著述提出 *Institutional Tracker*，這是一個專門設計用來監控與分析台灣證券交易所 (TWSE) 與證券櫃檯買賣中心 (TPEx) 內，機構投資人（外資與投信）淨買賣超活動之量化視覺化系統。本系統透過提出平行資料獲取架構 (Parallelized Data Acquisition)，解決了資料端點破碎化以及延遲問題。此外，本系統導入了成交量加權平均價 (VWAP) 作為法人資金流動的估算指標，以減輕單純依賴交易股數所造成的數值失真。為了確保即時的學術級市場分析，系統更實作了強健的反應用程式防火牆 (Anti-WAF) 機制以及高度響應式的交易終端使用者介面 (Trading Terminal UI)。

## 1. 簡介與動機 (Introduction)

在台灣的權益證券市場中，大型機構參與者（特別是外國機構投資者 FINI 與證券投資信託公司 ITC）的行為，是短期價格發現機能與中期價格動能的重要領先指標。然而，在不同的市場板塊 (TWSE 與 TPEx) 之間取得精準且同步的資料集合，面臨著巨大的工程挑戰，其中包括相異的 API 結構、嚴格的限流防火牆 (WAFs) 以及非交易日所產生的資料異常。*Institutional Tracker* 的開發正是為克服這些障礙，為研究人員與量化分析師提供自動化、不中斷的機構籌碼追蹤數據管線。

## 2. 研究方法與演算法 (Methodology)

### 2.1 平行資料採集架構 (Parallelized Data Acquisition)
為了將執行延遲降至最低並克服循序 I/O 存取上的通訊頸瓶，本系統採用了多執行緒模型 (`concurrent.futures.ThreadPoolExecutor`)。系統會同步發送針對 TWSE 與 TPEx 端點的資料請求。在各個市場範疇內，擷取法人交易紀錄 (T86) 與每日收盤行情 (MI_INDEX) 的子任務會進一步被平行化，將總體抓取耗時減少約 60%。

### 2.2 加權均價估值模型 (Volume-Weighted Average Price Valuation)
原始法人買賣資料的一個重大限制在於其計量單位為「股數」。為了能準確量化機構資金淨部位的財務規模，本系統會計算每檔股票當日的成交量加權平均價 (VWAP)：
$$  VWAP = \frac{\sum (Trade Value_i)}{\sum (Trade Volume_i)} $$
接著將法人的淨買賣股數乘上 VWAP，以推估真實的資金流動（單位：新台幣億元），藉此過濾掉大交易量低位階股票所帶來的數字雜訊，突顯真正的資金重兵所在。

### 2.3 智慧休市驗證與回溯機制 (Intelligent Market Status Validation)
有鑑於台灣股市常有突發性休市或連假（如：颱風假、國定假日），本追蹤系統實作了一套輕量級的前置驗證啟發式演算法 (`validate_trading_day`)，該演算法透過請求負載極小的市場概況 API 來執行。當使用者指定或系統排程遇到休市日，系統會自動啟動往回搜尋的遞迴機制（最多回溯到 $t-10$ 日），以保證能成功取得最近一個有效交易日的市況與籌碼，且不會對運算單元造成無謂的重試負擔或伺服器超時。

## 3. 系統架構 (System Architecture & Implementation)

本架構乃針對高可用性與極低資源消耗所設計，確保其能在無伺服器架構或限制性容器環境（如 Render 雲端免費方案）中穩定運算。

*   **資料擷取與清洗 (Data Extraction & Cleaning)**: 底層依賴 `pandas` 與 `requests` 套件。建立連線時可選擇性略過 TLS 驗證以規避部分容器環境的憑證錯誤，同時注入深度偽裝的 `User-Agent` 與 `Referer` 標頭 (Headers) 來穿透 TPEx 的 403 Forbidden 存取阻擋。
*   **運算引擎 (Computational Engine)**: 內建資料濾淨演算法，透過字元長度等特徵動態排除 ETF 及權證標的，確保純粹的普通股成分分析。
*   **網頁動態伺服器 (Web Server)**: 採用 WSGI 標準的 `Flask` 實例，並以 `gunicorn` 作為中介優化。工作進程與線程受到嚴格控制（Workers=$1$, Threads=$2$）並賦予較長的等待週期（$T=120s$），以此避免在高負載矩陣運算時發生伺服器記憶體溢出 (OOM) 或產生 502 Bad Gateway 錯誤。
*   **安全防護層 (Security Shell)**: 利用環境變數 (`USE_AUTH` 等) 進行抽離式組態設定，藉由簡易 HTTP 基本認證模式，實現系統的私有化保護與存取控制。

## 4. 部署與可靠度控制 (Deployment & Reliability)

本應用程式原生支援基於 Docker 的部署方式，並提供 `render.yaml` 用於平台即服務 (PaaS) 的一鍵設定。

為了對抗 PaaS 的冷啟動機制（如：閒置 15 分鐘後自動縮容降載），開發規範中強制要求整合外部健康檢查。追蹤系統提供無身分驗證的 `/health` 監聽端點。強烈建議建置方於外部監測伺服器（例如：[UptimeRobot](https://uptimerobot.com/)）上設定 5 分鐘為間隔（$f = \frac{1}{300} Hz$）的 `GET /health` 循環推播作業，確保伺服器永不僅入休眠狀態，維持毫秒級的 UI 初始載入反應。

## 5. 執行說明 (Execution)

### 5.1 本地執行與分析 (Local Execution CLI)
若需手動產生特定日期（例如：2026 年 2 月 24 日）的 Excel 分析報表：
```bash
pip install -r requirements.txt
python analyze.py 20260224
```

### 5.2 啟動終端視覺介面 (Terminal UI Execution)
於本地環境啟動 Flask 追蹤伺服器：
```bash
python app.py
```
啟動後，請將瀏覽器導航至 `http://127.0.0.1:5000` 即可進入多板塊視覺化交易分析終端。

## 6. AI 代理與開發須知 (Development Notes for AI Agents)
對於後續欲修改此儲存庫程式的 AI 代碼代理人 (AI Coding Agents)，必須嚴格遵守 `agent_recover.md` 規格內記載之反向工程模型與系統架構安全邊界，以維持反 WAF 功能及自動排程的邏輯完整性。

## 7. 結論與授權宣告 (License & Disclaimer)
本專案為開放原始碼軟體，宗旨為促進學術量化研究與程式化的市場數據探索。作者不保證衍生自本系統之任何交易決策的獲利可行性，使用者須自行承擔相關之一切金融操作風險。
