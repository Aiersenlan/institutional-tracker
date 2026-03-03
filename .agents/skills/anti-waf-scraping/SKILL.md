---
name: anti-waf-scraping
description: 針對台灣證券市場 (TWSE/TPEX) 的抗 WAF 爬蟲技巧與數據獲取規範。
---

# Anti-WAF Web Scraping Skill (台股抗封鎖爬蟲技能)

本技能模組化了繞過台灣證券交易所 (TWSE) 與櫃買中心 (TPEX) 網路防火牆 (WAF) 的核心邏輯，適用於自動化金融數據採集。

## 核心規則與限制

### 1. HTTP Headers 偽裝 (必備)
為了避免被辨識為機器人，必須模擬真實瀏覽器行為：
- **User-Agent**: 使用現代 Windows Chrome 格式。
- **Referer**: 存取 TPEX (櫃買中心) 時，Header 必須包含 `'Referer': 'https://www.tpex.org.tw/'`，否則會直接回傳 403 Forbidden。
- **Accept**: 建議設定為 `application/json, text/javascript, */*; q=0.01`。

### 2. 日期格式轉換
- **TWSE**: 使用西元格式 `YYYYMMDD` (範例: `20240301`)。
- **TPEX**: 使用民國紀年格式 `YYY/MM/DD` (範例: `113/03/01`)。

### 3. Session 管理與 SSL
- 建議使用 `requests.Session()` 維護連線。
- 在某些容器環境 (如 Docker/Render) 中，建議設定 `verify=False` 以避免 SSL 憑證驗證失敗。

## 程式碼範例

### 初始化 Session
```python
import requests

def create_scraping_session():
    session = requests.Session()
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
        'Referer': 'https://www.tpex.org.tw/'
    })
    return session
```

### 日期轉換邏輯
```python
def get_tw_date_formats(target_date):
    """
    input: datetime object
    output: (twse_str, tpex_str)
    """
    twse_str = target_date.strftime('%Y%m%d')
    roc_year = target_date.year - 1911
    tpex_str = f"{roc_year:03d}/{target_date.strftime('%m/%d')}"
    return twse_str, tpex_str
```

## 疑難排解 (Troubleshooting)
- **出現 403 Forbidden**: 檢查 `Referer` 是否遺漏，或 IP 是否已被暫時封鎖 (建議增加延遲)。
- **連線逾時**: 建議 `timeout` 設定為 15 秒以上。
- **資料空白**: 檢查當天是否為休市日 (週末或國定假日)。
