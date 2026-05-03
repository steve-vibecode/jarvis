from flask import Flask, jsonify, request
from flask_cors import CORS
from datetime import datetime
import subprocess
import os
import re
import anthropic
import requests
from anthropic import Anthropic

app = Flask(__name__)
CORS(app)

SPREADSHEET_ID = '15K66bJwqfS4lh1c_oagH759XR25uu101rx5EBasWAqw'

import pandas as pd

SUPABASE_URL = "https://gslybwjzsfprkjgximqe.supabase.co"
SUPABASE_KEY = "xxx"

@app.route('/get-portfolio')
def get_portfolio():
    try:
        headers = {
            "apikey": SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}"
        }

        # Fetch all transactions
        tx_res = requests.get(
            f"{SUPABASE_URL}/rest/v1/transactions?select=*&order=date.asc",
            headers=headers
        )
        transactions = tx_res.json()

        # Fetch all dividends
        dv_res = requests.get(
            f"{SUPABASE_URL}/rest/v1/dividends?select=*",
            headers=headers
        )
        dividends = dv_res.json()

        # Group dividends by ticker
        div_by_ticker = {}
        for d in dividends:
            t = d['ticker']
            if t not in div_by_ticker:
                div_by_ticker[t] = 0.0
            div_by_ticker[t] += float(d['amount'] or 0)

        # Calculate holdings using average cost method
        holdings = {}
        for tx in transactions:
            ticker  = tx['ticker']
            market  = tx['market']
            action  = (tx['action'] or '').upper()
            units   = float(tx['units'] or 0)
            price   = float(tx['price'] or 0)
            tx_type = (tx.get('type') or 'growth').lower()

            if ticker not in holdings:
                holdings[ticker] = {
                    'ticker':     ticker,
                    'market':     market,
                    'type':       tx_type,
                    'units':      0.0,
                    'cost_basis': 0.0,
                    'realised':   0.0,
                }

            h = holdings[ticker]

            if action == 'BUY':
                h['cost_basis'] += units * price
                h['units']      += units

            elif action == 'SELL':
                if h['units'] > 0:
                    avg_cost = h['cost_basis'] / h['units']
                    realised = (price - avg_cost) * units
                    h['realised']   += realised
                    h['cost_basis'] -= avg_cost * units
                    h['units']      -= units

        # Build result with P/L calculation
        result = []
        for ticker, h in holdings.items():
            if h['units'] < 0.0001:
                continue

            avg_cost    = h['cost_basis'] / h['units'] if h['units'] > 0 else 0
            divs        = div_by_ticker.get(ticker, 0.0)
            is_dividend = h['type'] == 'dividend'

            # These will be filled with live price from frontend
            result.append({
                'ticker':    ticker,
                'market':    h['market'],
                'type':      h['type'],
                'units':     round(h['units'], 5),
                'avg_cost':  round(avg_cost, 4),
                'cost_basis':round(h['cost_basis'], 2),
                'realised':  round(h['realised'], 2),
                'dividends': round(divs, 2),
            })

        return jsonify({"holdings": result})

    except Exception as e:
        return jsonify({"error": str(e)})

@app.route('/jarvis-query', methods=['POST'])
def jarvis_query():
    try:
        data = request.json
        client = Anthropic(api_key="xxx")

        response = client.messages.create(
            model="claude-sonnet-4-5-20250929",
            max_tokens=300,
            system=data.get('system',''),
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=data['messages']
        )

        reply = " ".join(
            block.text for block in response.content
            if hasattr(block, 'text')
        ).strip()

        return jsonify({"reply": reply})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ============================================================
# APP LAUNCHER
# ============================================================
APP_MAP = {
    'chrome':        r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    'google chrome': r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    'excel':         r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE',
    'word':          r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE',
    'notepad':       'notepad.exe',
    'calculator':    'calc.exe',
    'file explorer': 'explorer.exe',
    'explorer':      'explorer.exe',
    'task manager':  'taskmgr.exe',
    'spotify':       r'C:\Users\steveloo95\AppData\Roaming\Spotify\Spotify.exe',
    'telegram':      r'C:\Users\steveloo95\AppData\Roaming\Telegram Desktop\Telegram.exe',
    'vscode':        r'C:\Users\steveloo95\AppData\Local\Programs\Microsoft VS Code\Code.exe',
    'vs code':       r'C:\Users\steveloo95\AppData\Local\Programs\Microsoft VS Code\Code.exe',
    'my playlist': 'https://www.youtube.com/watch?v=-Q3h0KhWykI&list=PLvw1x9noKvHvJPjz8upXixfK8F80JQIqz',
}

STORE_APPS = {
    'whatsapp':        'whatsapp:',
    'microsoft store': 'ms-windows-store:',
    'settings':        'ms-settings:',
    'photos':          'ms-photos:',
    'email':           'https://mail.google.com',
}

@app.route('/open-app', methods=['POST'])
def open_app():
    try:
        data = request.json
        app_name = data.get('app', '').lower().strip()

        # Store apps
        if app_name in STORE_APPS:
            os.startfile(STORE_APPS[app_name])
            return jsonify({"success": True, "opened": app_name})

        path = APP_MAP.get(app_name)
        if not path:
            return jsonify({"error": f"Unknown app: {app_name}"})

        # Handle URLs
        if path.startswith('http://') or path.startswith('https://'):
            chrome = r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
            subprocess.Popen([chrome, path])
            return jsonify({"success": True, "opened": app_name})

        # Regular exe
        system_apps = ['notepad.exe', 'calc.exe', 'explorer.exe', 'taskmgr.exe']
        if path not in system_apps and not os.path.exists(path):
            return jsonify({"error": f"App not found at: {path}"})

        subprocess.Popen([path], shell=True)
        return jsonify({"success": True, "opened": app_name})
    except Exception as e:
        return jsonify({"error": str(e)})


# ============================================================
# GET COMMITMENTS — from Google Sheets (published CSV)
# ============================================================
@app.route('/get-commitment')
def get_commitment():
    try:
        SHEETS_CSV_URL = f'https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/pub?output=csv&single=true&gid=1165077234'
        df = pd.read_csv(SHEETS_CSV_URL, header=None)
        df = df.iloc[4:200].reset_index(drop=True)

        commitments    = []
        savings        = []
        calendar_tasks = []
        running_comm   = 0
        running_sav    = 0

        MONTH_MAP = {
            'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
            'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12,
            'january':1,'february':2,'march':3,'april':4,'june':6,
            'july':7,'august':8,'september':9,'october':10,'november':11,'december':12
        }

        def clean_amount(val):
            if pd.isna(val): return None
            s = str(val).strip()
            if s in ('', '-', 'nan'): return None
            s = s.replace('RM','').replace('$','').replace('₱','').replace(',','').strip()
            try: return float(s)
            except: return None

        def extract_date(label):
            # Match "May 1 Investment" or "January 15 Rental"
            m = re.match(r'^([a-zA-Z]+)\s+(\d+)\s*(.*)', label.strip())
            if m:
                mo_str = m.group(1).lower()
                day    = int(m.group(2))
                rest   = m.group(3).strip()
                month  = MONTH_MAP.get(mo_str)
                if month and 1 <= day <= 31:
                    return day, month, rest
            return None, None, label

        now = datetime.now()

        for _, row in df.iterrows():
            label_a = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ''
            val_b   = clean_amount(row.iloc[1]) if len(row) > 1 else None
            label_d = str(row.iloc[3]).strip() if len(row) > 3 and not pd.isna(row.iloc[3]) else ''
            val_e   = clean_amount(row.iloc[4]) if len(row) > 4 else None

            la = label_a.lower()
            ld = label_d.lower()

            if la in ('', 'nan', 'commitment', 'commitments', 'description', 'item'):
                label_a = ''; val_b = None
            if ld in ('', 'nan', 'savings', 'description', 'item'):
                label_d = ''; val_e = None

            if val_b is not None and la not in ('total', 'nan', ''):
                running_comm += val_b
            if val_e is not None and ld not in ('total', 'nan', ''):
                running_sav += val_e

            is_comm_total = 'total' in la
            is_sav_total  = 'total' in ld

            commitments.append({
                "label":      label_a if la != 'nan' else '',
                "amount":     round(running_comm, 2) if is_comm_total else val_b,
                "amount_myr": round(running_comm, 2) if is_comm_total else val_b,
                "currency":   "MYR",
                "is_total":   is_comm_total
            })
            if is_comm_total: running_comm = 0

            savings.append({
                "label":      label_d if ld != 'nan' else '',
                "amount":     round(running_sav, 2) if is_sav_total else val_e,
                "amount_myr": round(running_sav, 2) if is_sav_total else val_e,
                "currency":   "MYR",
                "is_total":   is_sav_total
            })
            if is_sav_total: running_sav = 0

            # Parse calendar task from commitment label
            if val_b is not None and label_a and not is_comm_total:
                day, month, clean_label = extract_date(label_a)
                if day and month:
                    calendar_tasks.append({
                        "day":   day,
                        "month": month,
                        "year":  now.year,
                        "text":  f"PAY: {clean_label} RM{val_b:,.2f}"
                    })

        return jsonify({
            "commitments":    commitments,
            "savings":        savings,
            "calendar_tasks": calendar_tasks
        })

    except Exception as e:
        return jsonify({"error": str(e)})

# ============================================================
# WHATSAPP
# ============================================================
whatsapp_opened_before = False

@app.route('/whatsapp-message', methods=['POST'])
def whatsapp_message():
    global whatsapp_opened_before
    try:
        import urllib.parse, time, threading
        data    = request.json
        phone   = data.get('phone','').replace('+','').replace(' ','').replace('-','')
        message = data.get('message','')
        encoded = urllib.parse.quote(message)
        url     = f'https://web.whatsapp.com/send?phone={phone}&text={encoded}'
        chrome  = r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
        is_first = not whatsapp_opened_before
        whatsapp_opened_before = True

        def auto_send():
            try:
                import pyautogui, pyperclip
                if is_first:
                    subprocess.Popen([chrome, url])
                    time.sleep(2)
                    pyautogui.hotkey('alt','tab')
                    time.sleep(7)
                else:
                    pyautogui.hotkey('alt','tab')
                    time.sleep(0.5)
                    pyautogui.hotkey('ctrl','l')
                    time.sleep(0.3)
                    pyperclip.copy(url)
                    pyautogui.hotkey('ctrl','v')
                    time.sleep(0.2)
                    pyautogui.hotkey('enter')
                    time.sleep(6)
                pyautogui.hotkey('enter')
                time.sleep(0.5)
                pyautogui.hotkey('enter')
                print(f"[WHATSAPP] Sent to {phone}")
            except Exception as e:
                print(f"Auto-send error: {e}")

        threading.Thread(target=auto_send, daemon=True).start()
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)})


if __name__ == '__main__':
    app.run(port=5000, debug=True)
