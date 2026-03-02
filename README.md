# Institutional Tracker: A High-Frequency Quantitative Assessment System for Institutional Fund Flows in Taiwan Equity Markets

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Flask](https://img.shields.io/badge/flask-app-green.svg)](#)
[![Deployment](https://img.shields.io/badge/deployment-render-purple.svg)](#)

👉 **Live Demo**: [https://institutional-tracker.onrender.com](https://institutional-tracker.onrender.com)

---

## Abstract

This paper presents the *Institutional Tracker*, a specialized quantitative visualization system designed to monitor and analyze the net buy/sell activities of institutional investors (Foreign Institutional Investors [FINI] and Investment Trust Companies [ITC]) within the Taiwan Stock Exchange (TWSE) and Taipei Exchange (TPEx). The system addresses the challenges of fragmented data endpoints and latency by proposing a parallelized data acquisition architecture. Furthermore, it incorporates Volume-Weighted Average Price (VWAP) as an estimator for institutional cash flows, mitigating the distortions caused by reliance solely on trade volume. A robust, anti-WAF (Web Application Firewall) mechanism and a highly responsive Trading Terminal UI are implemented to facilitate real-time, academic-grade market analysis.

## 1. Introduction

In the Taiwan equity market, the behavior of major institutional players—specifically FINIs and ITCs—serves as a critical leading indicator for short-term price discovery and mid-term momentum. However, acquiring precise, synchronized data across different market boards (TWSE and TPEx) poses substantial engineering challenges due to varying API structures, strict rate-limiting WAFs, and non-trading day anomalies. The *Institutional Tracker* was developed to overcome these barriers, providing researchers and quantitative analysts with an automated, uninterrupted pipeline for institutional footprint tracking.

## 2. Methodology

### 2.1 Parallelized Data Acquisition (平行資料採集架構)
To minimize execution latency and overcome sequential I/O bottlenecks, the system employs multithreading (`concurrent.futures.ThreadPoolExecutor`). Requests to TWSE and TPEx endpoints are dispatched simultaneously. Within each market scope, the retrieval of institutional trade records (T86) and closing price summaries (MI_INDEX) are further parallelized, reducing total fetch time by approximately 60%.

### 2.2 Volume-Weighted Average Price (VWAP) Valuation (加權均價估值模型)
A significant limitation of raw institutional data is its expression in "number of shares." To accurately quantify the financial magnitude of institutional positions, the system calculates the daily VWAP for each equity:
$$  VWAP = \frac{\sum (Trade Value_i)}{\sum (Trade Volume_i)} $$
Institutional net volume is subsequently multiplied by the VWAP to estimate the true capital flow (expressed in hundreds of millions, NTD), separating significant financial investments from large-volume, low-price penny stock anomalies.

### 2.3 Intelligent Market Status Validation (智慧休市驗證與回溯機制)
The Taiwan market frequently experiences unscheduled closures (e.g., typhoon days, sequential national holidays). The tracker implements a lightweight pre-validation heuristic (`validate_trading_day`) using minimal-payload market summary APIs. If an explicit or implicit target date returns a closed status, the system autonomously initiates a recursive backward search (up to $t-10$ days) to guarantee the retrieval of the most recent valid trading session without triggering heavy computational overhead or server timeouts.

## 3. System Architecture & Implementation

The architecture is designed for high availability and low resource consumption, operating within serverless or limited-tier container environments (e.g., Render Free Tier).

*   **Data Extraction & Cleaning**: `pandas` and `requests`. Network connections bypass TLS verification selectively to mitigate local container certificate errors, while deep-forged `User-Agent` and `Referer` headers prevent TPEx 403 Forbidden responses.
*   **Computational Engine**: Filtering algorithms exclude ETFs and warrants (identified via alphanumeric length heuristics), ensuring purity in the equity dataset.
*   **Web Server**: A WSGI `Flask` instance optimized with `gunicorn`. Workers are constrained ($N=1$, Threads=$2$) and provided extended timeouts ($T=120s$) to prevent OOM (Out-of-Memory) crashes and 502 Bad Gateway errors during high-load matrix calculations.
*   **Security Shell**: Environment variable injection (`USE_AUTH`, etc.) securely encapsulates the terminal behind Basic Authentication for private deployment scenarios.

## 4. Deployment & Reliability

The application natively supports Docker-based and Platform-as-a-Service (PaaS) deployments via `render.yaml`. 

To counter PaaS cold-start mechanics (down-scaling after 15 minutes of inactivity), the repository mandates an external health-check routine. The `/health` endpoint is kept unauthenticated. It is strongly recommended to utilize an external ping service (e.g., [UptimeRobot](https://uptimerobot.com/)) configured with a 5-minute interval ($f = \frac{1}{300} Hz$) targeting `GET /health` to ensure continuous runtime availability and sub-second UI responsiveness.

## 5. Execution

### 5.1 Local Execution (CLI)
To manually generate an Excel analytical report for a specific date (e.g., February 24, 2026):
```bash
pip install -r requirements.txt
python analyze.py 20260224
```

### 5.2 Terminal UI Execution
To boot the Flask server locally:
```bash
python app.py
```
Navigate to `http://127.0.0.1:5000` to access the Visual Trading Terminal.

## 6. Development Notes (AI Agents)
For subsequent AI Coding Agents contributing to this repository, adherence to the reverse-engineering models and system boundaries outlined in the `agent_recover.md` specification is mandatory to preserve the anti-WAF integrity.

## 7. License & Disclaimer
This project is open-source and intended solely for academic research and programmatic exploration. The authors provide no guarantees regarding the profitability of trading strategies derived from this system. Users assume all associated financial risks.
