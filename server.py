#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
台股即時查詢 - 本地 Proxy Server (優化版)
使用 Python 內建模組，不需安裝任何套件
"""
import json, os, sys, re, urllib.request, urllib.parse, urllib.error
from http.server import HTTPServer, BaseHTTPRequestHandler
from socketserver import ThreadingMixIn

class ThreadingHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeout

PORT = int(os.environ.get("PORT", 3838))
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

MIS_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36',
    'Referer': 'https://mis.twse.com.tw/stock/index.jsp',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'zh-TW,zh;q=0.9',
    'Connection': 'keep-alive',
}
TWSE_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Referer': 'https://www.twse.com.tw/',
    'Accept': 'application/json',
}

def fetch_url(url, headers, timeout=8):
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        raw = r.read()
        charset = r.headers.get_content_charset() or 'utf-8'
        return raw.decode(charset, errors='replace')

def pf(v):
    if not v or v == '-': return None
    try: return float(v)
    except: return None

def mis_to_dict(item, exchange):
    vol = pf(item.get('v'))
    return {
        'code': item.get('c',''), 'name': item.get('n',''),
        'exchange': exchange,
        'price': pf(item.get('z')), 'prev': pf(item.get('y')),
        'open': pf(item.get('o')), 'high': pf(item.get('h')),
        'low': pf(item.get('l')),
        'volume': vol * 1000 if vol else None,
        'week52High': None, 'week52Low': None, 'source': 'realtime',
    }

def query_mis(code, market):
    """TWSE MIS 盤中即時"""
    url = f'https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch={market}_{code}.tw&json=1&delay=0&_={int(datetime.now().timestamp()*1000)}'
    body = fetch_url(url, MIS_HEADERS, timeout=6)
    data = json.loads(body)
    items = data.get('msgArray', [])
    if not items: return None
    item = items[0]
    z = item.get('z', '')
    # 收盤後 z='-', 但 y(昨收) o h l 仍有值 → 顯示昨日收盤資料
    name = item.get('n', '')
    ex = 'TWSE 上市' if market == 'tse' else 'TPEx 上櫃'
    if not name: return None
    # 盤中現價
    if z and z != '-':
        return mis_to_dict(item, ex)
    # 盤後：z='-' 時用 y 當今日收盤，pz 當昨日收盤
    today_close = pf(item.get('y'))   # y = 最新收盤（盤後即今收）
    yest_close  = pf(item.get('pz'))  # pz = 昨日收盤
    if today_close:
        vol = pf(item.get('v'))
        return {
            'code': item.get('c', code), 'name': name,
            'exchange': ex + '（收盤）',
            'price': today_close,
            'prev': yest_close,
            'open': pf(item.get('o')), 'high': pf(item.get('h')),
            'low': pf(item.get('l')), 'volume': vol * 1000 if vol else None,
            'week52High': None, 'week52Low': None, 'source': 'mis_afterhours',
        }
    return None

def query_twse_day(code):
    """TWSE 日收盤資料（本月）"""
    today = datetime.now()
    # 用本月第一天查，可拿整月資料
    date_str = today.strftime('%Y%m') + '01'
    url = f'https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date={date_str}&stockNo={code}'
    body = fetch_url(url, TWSE_HEADERS, timeout=10)
    data = json.loads(body)
    if data.get('stat') != 'OK' or not data.get('data'):
        return None
    rows = data['data']
    last = rows[-1]
    def n(s):
        try: return float(s.replace(',',''))
        except: return None
    prev = n(rows[-2][6]) if len(rows) >= 2 else None
    title = data.get('title', '')
    m = re.search(r'(\d{4})\s+(.+)', title)
    name = m.group(2).strip() if m else code
    return {
        'code': code, 'name': name, 'exchange': 'TWSE（收盤）',
        'price': n(last[6]), 'prev': prev,
        'open': n(last[3]), 'high': n(last[4]), 'low': n(last[5]),
        'volume': n(last[1]),
        'week52High': None, 'week52Low': None, 'source': 'daily',
    }

def query_tpex_day(code):
    """TPEx 上櫃日收盤"""
    today = datetime.now()
    # TPEx 用民國年
    roc_year = today.year - 1911
    date_str = f'{roc_year}/{today.month:02d}'
    url = f'https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes?date={date_str}&stockNo={code}'
    body = fetch_url(url, TWSE_HEADERS, timeout=10)
    data = json.loads(body)
    if not data: return None
    # 找最新一筆
    rows = [r for r in data if r.get('SecuritiesCompanyCode') == code]
    if not rows: return None
    last = rows[-1]
    def n(s):
        try: return float(str(s).replace(',',''))
        except: return None
    return {
        'code': code, 'name': last.get('CompanyName', code),
        'exchange': 'TPEx 上櫃（收盤）',
        'price': n(last.get('Close')), 'prev': n(last.get('PreviousClose')),
        'open': n(last.get('Open')), 'high': n(last.get('High')),
        'low': n(last.get('Low')), 'volume': n(last.get('TradeVolume')),
        'week52High': None, 'week52Low': None, 'source': 'daily',
    }

def get_stock(code):
    # 1. 先試 TWSE MIS 上市（盤中 + 盤後昨收）
    try:
        r = query_mis(code, 'tse')
        if r and r.get('name'): return r
    except Exception as e:
        print(f'  tse miss: {e}')

    # 2. 試 TPEx MIS 上櫃
    try:
        r = query_mis(code, 'otc')
        if r and r.get('name'): return r
    except Exception as e:
        print(f'  otc miss: {e}')

    # 3. 試 TWSE 日收盤
    try:
        r = query_twse_day(code)
        if r and r.get('price'): return r
    except Exception as e:
        print(f'  twse_day miss: {e}')

    # 4. 試 TPEx OpenAPI 日收盤
    try:
        r = query_tpex_day(code)
        if r and r.get('price'): return r
    except Exception as e:
        print(f'  tpex_day miss: {e}')

    return None

# --- 新增：Yahoo Finance 代理 (支援美股、全球ETF及加密貨幣 fallback) ---
def query_yahoo(asset_symbol, market):
    suffixes = ['', '.SW', '.F'] if market == 'US' else ['.TW', '.TWO']
    if asset_symbol.upper().endswith('-USD'):
        suffixes = ['']
    for suffix in suffixes:
        symbol = f"{asset_symbol}{suffix}"
        url = f"https://query1.finance.yahoo.com/v8/finance/chart/{urllib.parse.quote(symbol)}"
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        try:
            body = fetch_url(url, headers, timeout=8)
            data = json.loads(body)
            if data.get('chart', {}).get('error'): continue
            price = data['chart']['result'][0]['meta'].get('regularMarketPrice')
            if price:
                return {'symbol': asset_symbol, 'price': price}
        except:
            continue
    return {'symbol': asset_symbol, 'price': 0}

# --- 新增：CoinGecko 代理 (加密貨幣) ---
def query_crypto(ids):
    url = f"https://api.coingecko.com/api/v3/simple/price?ids={urllib.parse.quote(ids)}&vs_currencies=usd"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    try:
        body = fetch_url(url, headers, timeout=10)
        return json.loads(body)
    except Exception as e:
        return {'error': str(e)}

# --- 新增：匯率代理 (USD/TWD) ---
def query_exchange_rate():
    url = "https://open.er-api.com/v6/latest/USD"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    try:
        body = fetch_url(url, headers, timeout=10)
        return json.loads(body)
    except Exception as e:
        return {'error': str(e)}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print(f'  [{datetime.now().strftime("%H:%M:%S")}] {args[0]} {args[1]}')

    def send_cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def do_OPTIONS(self):
        self.send_response(204); self.send_cors(); self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)

        if parsed.path in ('/', '/index.html'):
            fp = os.path.join(SCRIPT_DIR, 'index.html')
            try:
                with open(fp, 'rb') as f: content = f.read()
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.send_cors(); self.end_headers()
                self.wfile.write(content)
            except:
                self.send_response(404); self.end_headers()
            return

        if parsed.path.startswith('/assets/'):
            fp = os.path.join(SCRIPT_DIR, parsed.path.lstrip('/'))
            try:
                with open(fp, 'rb') as f: content = f.read()
                self.send_response(200)
                if fp.endswith('.css'):
                    self.send_header('Content-Type', 'text/css; charset=utf-8')
                elif fp.endswith('.js'):
                    self.send_header('Content-Type', 'application/javascript; charset=utf-8')
                self.send_cors(); self.end_headers()
                self.wfile.write(content)
            except:
                self.send_response(404); self.end_headers()
            return


        if parsed.path == '/api/stock':
            codes = params.get('code', [])
            if not codes or not codes[0].strip():
                return self._json(400, {'error': '請提供股票代碼'})
            code = codes[0].strip().upper()
            print(f'  查詢台股: {code}')
            try:
                result = get_stock(code)
                if result:
                    self._json(200, result)
                else:
                    self._json(404, {'error': f'查無 {code}，請確認台股代碼'})
            except Exception as e:
                self._json(500, {'error': str(e)})
            return
            
        if parsed.path == '/api/yahoo':
            symbols = params.get('symbol', [])
            markets = params.get('market', ['US'])
            if not symbols: return self._json(400, {'error': 'Missing symbol'})
            print(f'  查詢 Yahoo: {symbols[0]} ({markets[0]})')
            res = query_yahoo(symbols[0].strip().upper(), markets[0].strip().upper())
            return self._json(200, res)
            
        if parsed.path == '/api/crypto':
            ids = params.get('ids', [])
            if not ids: return self._json(400, {'error': 'Missing ids'})
            print(f'  查詢 Crypto: {ids[0]}')
            res = query_crypto(ids[0].strip())
            return self._json(200 if 'error' not in res else 500, res)
            
        if parsed.path == '/api/exchange':
            print('  查詢匯率')
            res = query_exchange_rate()
            return self._json(200 if 'error' not in res else 500, res)

        self.send_response(404); self.end_headers()

    def _json(self, status, data):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.send_cors(); self.end_headers()
        self.wfile.write(body)


if __name__ == '__main__':
    server = ThreadingHTTPServer(('0.0.0.0', PORT), Handler)
    print()
    print('  ╔══════════════════════════════════════╗')
    print(f'  ║  台股查詢伺服器  http://localhost:{PORT}  ║')
    print('  ╚══════════════════════════════════════╝')
    print()
    print(f'  → 請在瀏覽器開啟 http://localhost:{PORT}')
    print('  → 按 Ctrl+C 停止\n')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n  已停止。')
        sys.exit(0)
