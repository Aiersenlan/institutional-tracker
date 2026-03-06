import os
import sys
import io

# Force UTF-8 encoding for stdout/stderr to fix Chinese character corruption in Windows console
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

from flask import Flask, render_template, request, jsonify, send_file, Response
import pandas as pd

app = Flask(__name__)

# Basic Authentication Configuration
# 設定環境變數啟動「VIP 專屬包廂」保護機制
USE_AUTH = os.environ.get('USE_AUTH', 'false').lower() == 'true'
AUTH_USER = os.environ.get('AUTH_USER', 'admin')
AUTH_PASS = os.environ.get('AUTH_PASS', 'stock1234')

def check_auth(username, password):
    return username == AUTH_USER and password == AUTH_PASS

def authenticate():
    return Response(
    '您必須輸入正確的VIP通行證與密碼才能觀看目前的私有盤後報告。\n'
    'Login Required to access Institutional Tracker.', 401,
    {'WWW-Authenticate': 'Basic realm="VIP Login Required"'})

@app.before_request
def require_login():
    # Allow health check to pass without auth to help keep-alive services
    if request.path == '/health':
        return
    # Only enforce if USE_AUTH is true
    if USE_AUTH:
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()

@app.route('/health')
def health():
    print(">>> Health check ping received! Keeping server awake. <<<")
    return "OK", 200

# Ensure the template directory exists
os.makedirs('templates', exist_ok=True)
os.makedirs('static', exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get_available_dates')
def get_available_dates():
    daily_dates = []
    weekly_dates = []
    for filename in os.listdir('.'):
        if not filename.endswith('.xlsx'):
            continue
        if filename.startswith('market_analysis_week_'):
            date_str = filename.replace('market_analysis_week_', '').replace('.xlsx', '')
            weekly_dates.append(date_str)
        elif filename.startswith('market_analysis_'):
            date_str = filename.replace('market_analysis_', '').replace('.xlsx', '')
            if len(date_str) == 8:
                daily_dates.append(date_str)
    
    daily_dates.sort(reverse=True)
    weekly_dates.sort(reverse=True)
    return jsonify({
        'daily': daily_dates,
        'weekly': weekly_dates
    })

@app.route('/get_report/<date_str>')
def get_report(date_str):
    period = request.args.get('period', 'day')
    if period == 'week':
        filename = f'market_analysis_week_{date_str}.xlsx'
    else:
        filename = f'market_analysis_{date_str}.xlsx'
        
    if not os.path.exists(filename):
        return jsonify({'error': f'Report not found ({filename})'}), 404
    
    try:
        # Read all sheets
        xls = pd.ExcelFile(filename)
        result = {}
        for sheet_name in xls.sheet_names:
            print(f"Reading sheet: {sheet_name}")
            # First, read without header to get the raw data layout
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            
            # The structure is known: 
            # Row 0: Date
            # Row 1: Main headers (外資買超, etc)
            # Row 2: Sub headers
            # Row 3 onwards: Data
            
            # Extract header info
            main_headers = df.iloc[1].fillna('').tolist()
            sub_headers = df.iloc[2].fillna('').tolist()
            
            # Extract data
            data_rows = []
            for i in range(3, len(df)):
                row_data = df.iloc[i].fillna('').tolist()
                # Check if row is completely empty
                if not all(str(item).strip() == '' for item in row_data):
                    data_rows.append(row_data)
            
            result[sheet_name] = {
                'main_headers': main_headers,
                'sub_headers': sub_headers,
                'data': data_rows
            }
            
        return jsonify(result)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/download/<date_str>')
def download(date_str):
    period = request.args.get('period', 'day')
    if period == 'week':
        filename = f'market_analysis_week_{date_str}.xlsx'
    else:
        filename = f'market_analysis_{date_str}.xlsx'
    if os.path.exists(filename):
        return send_file(filename, as_attachment=True)
    return 'File not found', 404

@app.route('/trigger_analysis', methods=['POST'])
def trigger_analysis():
    # Execute analyze.py to generate new data
    try:
        import subprocess
        
        target_date = None
        period = 'day'
        if request.is_json:
            data = request.json or {}
            target_date = data.get('date') # e.g. YYYY-MM-DD
            period = data.get('period', 'day')
            
        cmd = ["python", "analyze.py"]
        if target_date:
            target_date_str = target_date.replace("-", "")
            cmd.append(target_date_str)
            print(f"Manual trigger activated for date {target_date_str} ({period})...")
        else:
            print(f"Manual trigger activated for today ({period})...")
            
        if period == 'week':
            cmd.extend(["--period", "week"])
            
        process = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
        
        debug_output = process.stdout + "\n" + process.stderr
        
        if process.returncode != 0:
            return jsonify({
                "status": "error", 
                "message": "分析失敗: 執行過程發生錯誤或沒有資料。",
                "debug_log": debug_output
            }), 500
            
        return jsonify({
            "status": "success", 
            "message": "分析完成，新的報表已產出！",
            "debug_log": debug_output
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"status": "error", "message": f"系統錯誤: {str(e)}", "debug_log": traceback.format_exc()}), 500

if __name__ == '__main__':
    print("啟動網頁伺服器: http://localhost:5000")
    app.run(debug=True, port=5000)
