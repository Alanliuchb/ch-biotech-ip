#!/usr/bin/env python3
"""正瀚生技 智財與登記管理 v3 - 自動更新腳本

以 v2（專利、產品登記可正常運作版）為基礎，僅強化商標資料讀取。
"""
import json, datetime, urllib.request, ssl, io, sys, os

GOOGLE_SERVICE_ACCOUNT_ENV = 'GOOGLE_SERVICE_ACCOUNT_JSON'
_GOOGLE_CREDENTIALS = None

URLS = {
    'trademark':    'https://docs.google.com/spreadsheets/d/13iX5d_tig149MicN-WvTENS_kTiQjkNb/export?format=xlsx',
    'patent':       'https://docs.google.com/spreadsheets/d/1Uj_PV344NkDiY2n8YCs_HyjpnQYn87Ca/export?format=xlsx',
    'registration': 'https://docs.google.com/spreadsheets/d/1llnfbjcPST6Wa0p6psIxUfnjlGicZEi9/export?format=xlsx',
}
SHEET_NAMES = {
    'trademark':    '商標進度管理表',
    'patent':       '專利進度管理表',
    'registration': '產品登記管理表',
}
# 產品登記：已從 Excel 刪除的肥料登記欄位，不顯示在案件明細
REG_HIDE = {'crops', 'n-p-k', 'organic matter', 'raw materials', 'others'}

_TW = datetime.timezone(datetime.timedelta(hours=8))
_NOW_TW = datetime.datetime.now(datetime.timezone.utc).astimezone(_TW)
TODAY = _NOW_TW.date()
TODAY_STR = TODAY.strftime('%Y/%m/%d')
NOW_STR = _NOW_TW.strftime('%Y/%m/%d %H:%M')


def download_excel(name, url):
    global _GOOGLE_CREDENTIALS
    ctx = ssl.create_default_context()
    service_account_json = os.environ.get(GOOGLE_SERVICE_ACCOUNT_ENV, '').strip()
    if not service_account_json:
        raise RuntimeError(
            f'找不到 GitHub Secret：{GOOGLE_SERVICE_ACCOUNT_ENV}。'
            '請確認已在 GitHub Repository 的 Actions Secrets 建立。'
        )
    try:
        from google.auth.transport.requests import Request as GoogleAuthRequest
        from google.oauth2.service_account import Credentials
        if _GOOGLE_CREDENTIALS is None:
            info = json.loads(service_account_json)
            _GOOGLE_CREDENTIALS = Credentials.from_service_account_info(
                info,
                scopes=['https://www.googleapis.com/auth/drive.readonly']
            )
        if not _GOOGLE_CREDENTIALS.valid:
            _GOOGLE_CREDENTIALS.refresh(GoogleAuthRequest())
    except Exception as e:
        raise RuntimeError(f'Google Service Account 驗證失敗：{e}') from e

    # 使用 OAuth Bearer Token 讀取私人 Google Sheets，不依賴公開分享連結。
    req = urllib.request.Request(url, headers={
        'User-Agent': 'CH-Biotech-IP-Updater/4.0',
        'Authorization': f'Bearer {_GOOGLE_CREDENTIALS.token}',
        'Cache-Control': 'no-cache',
        'Pragma': 'no-cache',
        'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream',
    })
    try:
        with urllib.request.urlopen(req, context=ctx, timeout=30) as r:
            data = r.read()
            content_type = (r.headers.get('Content-Type') or '').lower()
        if not data.startswith(b'PK\x03\x04'):
            preview = data[:200].decode('utf-8', errors='replace').replace('\n', ' ')
            raise RuntimeError(f'{name} 下載結果不是 XLSX（Content-Type: {content_type}；回應開頭: {preview}）')
        print(f'  OK {name}: {len(data):,} bytes')
        return data
    except Exception as e:
        print(f'  ERROR {name}: {e}', file=sys.stderr)
        sys.exit(1)


def format_cell_value(value):
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime('%Y-%m-%d')
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip() if value is not None else ''
    return '' if text.startswith('#') else text


def read_excel_rows(data, header_row=1):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    if len(all_rows) < header_row + 1:
        return []
    headers = [str(h).strip() if h is not None else '' for h in all_rows[header_row]]
    records = []
    for row in all_rows[header_row + 1:]:
        if all(v is None for v in row):
            continue
        rec = {}
        for h, v in zip(headers, row):
            if h:
                rec[h] = format_cell_value(v)
        records.append(rec)
    return records


def read_trademark_rows(data):
    """讀取商標表：自動辨識工作表與標題列，避免資料列位置改變後讀空。"""
    import openpyxl

    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    expected = {
        '商標', '商標案件', '商標名稱', '國別', '申請號', '申請案號',
        '進度狀況', '進度狀態', '狀態/進度說明', '商標分類', '申請日期',
        '使用期間', '使用期間-到期', '使用期間-到期日', '使用期間_到期日',
        '註冊編號', '註冊號', '證書號 (進度)', '證書號(進度)',
    }

    # 優先使用指定工作表；若匯出檔不存在該名稱，退回逐張表搜尋。
    sheets = []
    preferred = SHEET_NAMES.get('trademark')
    if preferred and preferred in wb.sheetnames:
        sheets.append(wb[preferred])
    sheets.extend(ws for ws in wb.worksheets if ws not in sheets)

    best = None
    for ws in sheets:
        rows = list(ws.iter_rows(values_only=True))
        for header_row, row in enumerate(rows[:8]):
            values = {str(v).strip() for v in row if v is not None}
            score = len(values & expected)
            if score >= 2 and (best is None or score > best[0]):
                best = (score, ws, header_row, rows)

    if best is None:
        print('  WARNING 商標：找不到明確標題列，退回第一張工作表第 1 列', file=sys.stderr)
        return read_excel_rows(data, header_row=0)

    score, ws, header_row, rows = best
    headers = [str(h).strip() if h is not None else '' for h in rows[header_row]]
    records = []
    for row in rows[header_row + 1:]:
        if all(v is None for v in row):
            continue
        record = {}
        for header, value in zip(headers, row):
            if header:
                record[header] = format_cell_value(value)
        if any(record.values()):
            records.append(record)

    print(f'  商標讀取：工作表「{ws.title}」、標題列第 {header_row + 1} 列、辨識欄位 {score} 個、資料 {len(records)} 筆')
    return records


def calc_tm_deadline(date_str):
    """商標：到期前 6 個月顯示警示"""
    v = (date_str or '').strip()
    if not v:
        return '', '待補期限'
    if v.upper() == 'N/A':
        return 'N/A', 'N/A'
    s = v.replace('/', '-').split(' ')[0]
    try:
        d = datetime.date.fromisoformat(s)
        delta = (d - TODAY).days
        ds = d.strftime('%Y-%m-%d')
        if delta < 0:      return ds, '期限已過'
        elif delta <= 180: return ds, '即將到期'
        else:              return ds, '正常'
    except Exception:
        return v, '日期異常'


def calc_deadline(date_str):
    """一般期限計算（多門檻）"""
    v = (date_str or '').strip()
    if not v:
        return '', '待補期限'
    if v.upper() == 'N/A':
        return 'N/A', 'N/A'
    s = v.replace('/', '-').split(' ')[0]
    try:
        d = datetime.date.fromisoformat(s)
        delta = (d - TODAY).days
        ds = d.strftime('%Y-%m-%d')
        if delta < 0:      return ds, '期限已過'
        elif delta <= 30:  return ds, '即將到期(30天)'
        elif delta <= 90:  return ds, '即將到期(90天)'
        elif delta <= 180: return ds, '即將到期(180天)'
        elif delta <= 365: return ds, '即將到期(365天)'
        else:              return ds, '正常'
    except Exception:
        return v, '日期異常'


def trademark_status(r):
    """4 分類：註冊案 / 申請案 / 核駁案 / 放棄案"""
    st = (r.get('進度狀況') or r.get('進度狀態') or '').strip()
    if st in ('已取證', '已取得', '註冊案'):
        return '註冊案'
    if st in ('申請案', '核駁案', '放棄案'):
        return st
    cert = (r.get('證書號 (進度)') or r.get('證書號(進度)') or r.get('註冊號') or '').strip()
    if not cert:
        return '申請案'
    if any(k in cert for k in ('放棄', '失效')):
        return '放棄案'
    if '核駁' in cert:
        return '核駁案'
    if cert.startswith('【') or '訴願' in cert or '申復' in cert:
        return '申請案'
    return '註冊案'


def patent_status(state):
    v = state.strip().replace('\n', '')
    if v.startswith('專利通過') or v.startswith('領證'):
        return '已取得'
    if any(k in v for k in ('放棄', '撤回', '失效', '結案')):
        return '已結案'
    return '申請中'


def process_trademark(records):
    for r in records:
        r['_status'] = trademark_status(r)
        # 到期日：優先用新分欄，備用舊合併欄
        end_raw = (r.get('使用期間-到期') or r.get('使用期間-到期日') or r.get('使用期間_到期日') or '').strip()
        if not end_raw:
            period = (r.get('使用期間') or '').strip()
            if period and '-' in period:
                end_raw = period.split('-')[-1].strip()
        end_date, dl = calc_tm_deadline(end_raw)
        r['_end_date'] = end_date
        r['_deadline_status'] = dl
        # 起始日
        start_raw = (r.get('使用期間-起始') or r.get('使用期間-起始日') or r.get('使用期間_起始日') or '').strip()
        if not start_raw:
            period = (r.get('使用期間') or '').strip()
            if period and '-' in period:
                start_raw = period.split('-')[0].strip()
        r['_start_date'] = start_raw
    return records


def process_patent(records):
    for r in records:
        r['_status'] = patent_status(r.get('目前狀態', ''))
        end_raw = (r.get('證書到期日') or '').strip()
        end_date, dl = calc_deadline(end_raw)
        r['_end_date'] = end_date
        r['_deadline_status'] = dl
    return records


def process_registration(records):
    sm = {'維持證書': '已取得', '已取得': '已取得',
          '登記中': '辦理中', '辦理中': '辦理中', '申請中': '辦理中'}
    for r in records:
        raw = r.get('進度', '')
        r['_status'] = sm.get(raw, raw or '辦理中')
        exp = (r.get('證書有效期限') or '').strip()
        end_date, dl = calc_deadline(exp)
        r['_end_date'] = end_date
        r['_deadline_status'] = dl
    return records


# ─── HTML ────────────────────────────────────────────────────────────────────

def build_html(trademark, patent, registration):
    data = {'trademark': trademark, 'patent': patent, 'registration': registration}
    data_js = json.dumps(data, ensure_ascii=False).replace('</', '<' + chr(92) + '/')
    tm_c = len(trademark)
    pt_c = len(patent)
    rg_c = len(registration)
    reg_hide_js = json.dumps(list(REG_HIDE))

    css = r'''
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Noto Sans TC",sans-serif;background:#f0f2f5;color:#1a1a2e;display:flex;height:100vh;overflow:hidden}
#sidebar{width:220px;min-width:220px;background:#0f1629;color:#c8d0e0;display:flex;flex-direction:column;height:100vh;overflow-y:auto;flex-shrink:0}
#main{min-width:0}
.brand{padding:24px 20px 16px;border-bottom:1px solid #1e2a4a}
.brand-logo{font-size:22px;font-weight:800;color:#fff;letter-spacing:2px}
.brand-sub{font-size:10px;color:#6b7a99;margin-top:3px;letter-spacing:1px}
.brand-title{font-size:12px;color:#8899bb;margin-top:10px;line-height:1.4}
.nav{padding:16px 0;flex:1}
.nav-section{padding:6px 20px 4px;font-size:10px;color:#4a5568;letter-spacing:1.5px;text-transform:uppercase;font-weight:600}
.nav-item{display:flex;align-items:center;justify-content:space-between;padding:10px 20px;cursor:pointer;border-left:3px solid transparent;transition:all .15s;font-size:13.5px;color:#9aabbf}
.nav-item:hover{background:#1a2540;color:#e0e8f5;border-left-color:#3b5bdb}
.nav-item.active{background:#1a2a50;color:#fff;border-left-color:#4c6ef5;font-weight:500}
.nav-badge{font-size:11px;background:#1e3060;color:#7c9be8;padding:2px 7px;border-radius:10px;font-weight:600}
.nav-badge.alert{background:#3d1515;color:#f87171}
.sidebar-footer{padding:16px 20px;border-top:1px solid #1e2a4a;font-size:11px;color:#4a5568;line-height:1.5}
#main{flex:1;min-width:0;overflow-y:auto;display:flex;flex-direction:column}
.topbar{background:#fff;border-bottom:1px solid #e0e6ef;padding:14px 28px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:10;box-shadow:0 1px 4px rgba(0,0,0,.05)}
.topbar-left{font-size:13px;color:#6b7a99}
.topbar-left span{color:#1a1a2e;font-weight:500}
.btn-outline{padding:7px 16px;border-radius:6px;border:1px solid #d0d7e3;background:#fff;color:#4a5568;cursor:pointer;font-size:13px;font-weight:500}
.btn-outline:hover{background:#f5f7fb;border-color:#4c6ef5;color:#4c6ef5}
.btn-primary{padding:7px 16px;border-radius:6px;border:none;background:#4c6ef5;color:#fff;cursor:pointer;font-size:13px;font-weight:500}
.btn-primary:hover{background:#3b5bdb}
.content{padding:28px;flex:1}
.ov-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;margin-bottom:24px}
.card{background:#fff;border-radius:10px;padding:18px 20px;box-shadow:0 1px 4px rgba(0,0,0,.06);border:1px solid #e8edf5}
.card-label{font-size:11px;color:#8899bb;font-weight:600;letter-spacing:.5px;margin-bottom:6px}
.card-value{font-size:26px;font-weight:700;color:#1a1a2e;line-height:1}
.card-sub{font-size:12px;color:#8899bb;margin-top:6px}
.card.ac{border-color:#fca5a5;background:#fff8f8}.card.ac .card-label{color:#dc2626}.card.ac .card-value{color:#dc2626}
.card.wc{border-color:#fcd34d;background:#fffdf0}.card.wc .card-label{color:#b45309}.card.wc .card-value{color:#b45309}
.section-title{font-size:15px;font-weight:600;color:#1a1a2e;margin-bottom:12px;display:flex;align-items:center;gap:8px}
.section-title::after{content:"";flex:1;height:1px;background:#e0e6ef}
.sync-bar{background:#eff3ff;border:1px solid #c5d2f6;border-radius:8px;padding:9px 14px;font-size:12px;color:#4c6ef5;margin-bottom:14px;display:flex;align-items:center;gap:6px}
.fbar{background:#fff;border-radius:10px;border:1px solid #e8edf5;padding:12px 16px;margin-bottom:12px;display:flex;flex-wrap:wrap;gap:8px;align-items:center;box-shadow:0 1px 3px rgba(0,0,0,.04)}
.fbar input,.fbar select{padding:7px 10px;border:1px solid #d0d7e3;border-radius:6px;font-size:13px;color:#1a1a2e;background:#fff;outline:none}
.fbar input:focus,.fbar select:focus{border-color:#4c6ef5}
.fbar input{width:190px}
.frs{font-size:12px;color:#4c6ef5;cursor:pointer;padding:4px 8px;border-radius:4px}
.frs:hover{background:#eff3ff}
.rcount{font-size:12px;color:#8899bb;margin-left:auto}
.twrap{background:#fff;border-radius:10px;border:1px solid #e8edf5;overflow-x:auto;overflow-y:hidden;box-shadow:0 1px 4px rgba(0,0,0,.06);margin-bottom:20px;-webkit-overflow-scrolling:touch}
table{width:100%;min-width:880px;border-collapse:collapse;font-size:13px}
thead th{background:#f7f9fd;padding:10px 13px;text-align:left;font-size:11px;font-weight:600;color:#6b7a99;letter-spacing:.5px;border-bottom:1px solid #e0e6ef;white-space:nowrap;cursor:pointer;user-select:none}
thead th:hover{color:#4c6ef5}
tbody tr{border-bottom:1px solid #f4f6fb;transition:background .1s;cursor:pointer}
tbody tr:hover{background:#f7f9fd}
tbody tr:last-child{border-bottom:none}
td{padding:9px 13px;vertical-align:middle;white-space:nowrap}
.twrap .tm-name{white-space:normal}
.cn{font-weight:500;color:#1a1a2e;max-width:none;overflow:visible;text-overflow:clip;white-space:normal;word-break:break-word}
.tm-name{font-weight:500;color:#1a1a2e;white-space:normal;overflow:visible;text-overflow:clip;word-break:break-word;line-height:1.35}
.cs{font-size:11px;color:#8899bb;margin-top:2px}
.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11.5px;font-weight:500;white-space:nowrap}
.b-註冊案{background:#dcfce7;color:#166534}
.b-申請案{background:#dbeafe;color:#1d4ed8}
.b-審查中{background:#dbeafe;color:#1d4ed8}
.b-核駁案{background:#fff7ed;color:#c2410c}
.b-放棄案{background:#f3f4f6;color:#6b7280}
.b-已取得{background:#dcfce7;color:#166534}
.b-申請中{background:#dbeafe;color:#1d4ed8}
.b-已結案{background:#f3f4f6;color:#6b7280}
.b-辦理中{background:#dbeafe;color:#1d4ed8}
.b-T{background:#f0f4ff;color:#3b5bdb}
.dl-正常{background:#f0fdf4;color:#16a34a}
.dl-即將到期{background:#fff7ed;color:#c2410c;font-weight:600}
.dl-期限已過{background:#fef2f2;color:#dc2626;font-weight:600}
.dl-30{background:#fef2f2;color:#dc2626}
.dl-90{background:#fff7ed;color:#c2410c}
.dl-180{background:#fff7ed;color:#d97706}
.dl-365{background:#fefce8;color:#a16207}
.dl-待補期限{background:#f3f4f6;color:#9ca3af}
.dl-日期異常{background:#fdf4ff;color:#7c3aed}
.dl-NA{background:#f3f4f6;color:#9ca3af}
.pgbar{display:flex;align-items:center;justify-content:flex-end;gap:6px;padding:12px 16px;border-top:1px solid #f0f4fa}
.pgb{width:30px;height:30px;border-radius:6px;border:1px solid #d0d7e3;background:#fff;cursor:pointer;font-size:13px;display:flex;align-items:center;justify-content:center;transition:all .15s}
.pgb:hover:not([disabled]){background:#eff3ff;border-color:#4c6ef5;color:#4c6ef5}
.pgb.apg{background:#4c6ef5;color:#fff;border-color:#4c6ef5}
.pgb[disabled]{opacity:.4;cursor:not-allowed}
.pgi{font-size:12px;color:#8899bb;margin:0 4px}
.plist{background:#fff;border-radius:10px;border:1px solid #e8edf5;overflow-x:auto;overflow-y:hidden;margin-bottom:20px;box-shadow:0 1px 4px rgba(0,0,0,.06);-webkit-overflow-scrolling:touch}
.pr{display:grid;grid-template-columns:minmax(260px,1fr) 140px 130px 150px;min-width:680px;border-bottom:1px solid #f0f4fa;font-size:13px;cursor:pointer;transition:background .1s}
.pr:hover:not(.prh){background:#f7f9fd}.pr:last-child{border-bottom:none}
.prh{background:#f7f9fd;font-size:11px;font-weight:600;color:#6b7a99;cursor:default}
.pr>div{padding:10px 14px}
.mo{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:100;display:flex;align-items:center;justify-content:center;opacity:0;pointer-events:none;transition:opacity .2s}
.mo.open{opacity:1;pointer-events:auto}
.modal{background:#fff;border-radius:12px;width:700px;max-width:95vw;max-height:88vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.2)}
.mh{padding:16px 22px;border-bottom:1px solid #e8edf5;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;background:#fff;z-index:1}
.mh h3{font-size:15px;font-weight:600}
.mclose{width:28px;height:28px;border:none;background:none;cursor:pointer;font-size:17px;color:#6b7a99;border-radius:6px}
.mclose:hover{background:#f3f4f6}
.mbody{padding:18px 22px}
.dg{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.di label{font-size:11px;color:#8899bb;font-weight:600;display:block;margin-bottom:2px}
.di .dv{font-size:13px;color:#1a1a2e}
.di.full{grid-column:1/-1}
.lock-screen{min-height:100vh;display:flex;align-items:center;justify-content:center;background:#f3f6fb;padding:20px}
.lock-card{width:min(390px,100%);background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:30px;box-shadow:0 8px 30px rgba(24,43,77,.10);text-align:center}
.lock-logo{font-size:28px;font-weight:800;letter-spacing:2px;color:#17213a;margin-bottom:4px}
.lock-sub{font-size:13px;color:#6b7a99;margin-bottom:24px}
.lock-card input{width:100%;box-sizing:border-box;border:1px solid #cfd8e8;border-radius:8px;padding:11px 12px;font-size:14px;outline:none;margin-bottom:10px}
.lock-card input:focus{border-color:#4263eb;box-shadow:0 0 0 3px rgba(66,99,235,.12)}
.lock-card button{width:100%;border:0;border-radius:8px;padding:11px;background:#315bdc;color:#fff;font-size:14px;cursor:pointer}
.lock-error{height:18px;color:#d9485f;font-size:12px;margin-top:10px}
body.locked>aside,body.locked>main,body.locked>.mo{display:none!important}
@media (max-width:900px){
  body{display:block;height:auto;min-height:100vh;overflow:auto}
  #sidebar{width:100%;min-width:0;height:auto;overflow:visible;position:relative}
  .brand{padding:14px 18px 10px;display:flex;align-items:center;gap:10px}
  .brand-logo{font-size:20px}.brand-sub{margin-top:2px}.brand-title{margin:0 0 0 4px}
  .nav{padding:6px 8px;display:flex;align-items:center;gap:4px;overflow-x:auto;flex:none}
  .nav-section{display:none}
  .nav-item{flex:0 0 auto;padding:9px 11px;border-left:0;border-bottom:3px solid transparent;font-size:12px;white-space:nowrap}
  .nav-item:hover,.nav-item.active{border-left-color:transparent;border-bottom-color:#4c6ef5}
  .sidebar-footer{display:none}
  #main{min-height:calc(100vh - 92px);overflow:visible}
  .content{padding:20px}
  .twrap{overflow-x:auto;-webkit-overflow-scrolling:touch}
  .twrap table{min-width:720px}
  .plist{overflow-x:auto;-webkit-overflow-scrolling:touch}
  .pr{min-width:560px}
}
@media (max-width:600px){
  .brand{padding:12px 14px 9px}.brand-title{font-size:11px}
  .nav{padding:5px 6px}.nav-item{padding:8px 9px;font-size:11.5px}.nav-badge{font-size:10px;padding:1px 5px;margin-left:5px}
  .topbar{padding:11px 14px}.topbar-left{font-size:12px}.btn-outline,.btn-primary{padding:7px 11px;font-size:12px}
  .content{padding:12px}.ov-grid{grid-template-columns:1fr;gap:10px;margin-bottom:18px}
  .card{padding:14px 16px}.card-value{font-size:24px}
  .fbar{padding:10px;gap:7px}.fbar input{width:100%;flex:1 1 100%}.fbar select{flex:1;min-width:0}.frs{padding:5px 6px}.rcount{width:100%;margin-left:0;text-align:right}
  .twrap table{min-width:700px;font-size:12px}.twrap td{padding:8px 10px}.twrap thead th{padding:9px 10px}
  .section-title{font-size:14px}.pr{min-width:530px;font-size:12px}.pr>div{padding:9px 10px}
  .modal{width:100%;max-width:100vw;max-height:100vh;border-radius:0}.mh{padding:13px 16px}.mbody{padding:15px 16px}.dg{grid-template-columns:1fr;gap:9px}
  .lock-card{padding:26px 22px}
}
.sc{background:#fff;border-radius:10px;border:1px solid #e8edf5;padding:20px;margin-bottom:14px;box-shadow:0 1px 4px rgba(0,0,0,.06)}
.sc h3{font-size:14px;font-weight:600;margin-bottom:10px}
.ibox{background:#eff3ff;border:1px solid #c5d2f6;border-radius:8px;padding:12px 14px;font-size:13px;color:#3b5bdb;line-height:1.5}
.empty{text-align:center;padding:36px 20px;color:#9ca3af;font-size:14px}
.chk-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:6px;margin:10px 0}
.chk-item{display:flex;align-items:center;gap:6px;font-size:13px;cursor:pointer;padding:4px 6px;border-radius:4px}
.chk-item:hover{background:#f7f9fd}
.chk-item input{cursor:pointer}
.exp-section{margin-bottom:20px}
.exp-section h4{font-size:13px;font-weight:600;margin-bottom:8px;color:#1a1a2e}
.mode2-tbl{width:100%;border-collapse:collapse;font-size:12px;margin-top:8px}
.mode2-tbl th,.mode2-tbl td{border:1px solid #e0e6ef;padding:6px 10px;text-align:center}
.mode2-tbl th{background:#f7f9fd;font-weight:600}
.mode2-tbl .rc{text-align:left;font-weight:500}
'''

    sync_section = f'''<div class="sc"><h3>⇄ 資料同步狀態</h3><div class="dg" style="margin-top:10px">
<div class="di"><label>最後同步時間</label><div class="dv">{NOW_STR}</div></div>
<div class="di"><label>來源</label><div class="dv">Google Sheets 自動同步</div></div>
<div class="di"><label>商標案件</label><div class="dv">{tm_c} 筆</div></div>
<div class="di"><label>專利案件</label><div class="dv">{pt_c} 筆</div></div>
<div class="di"><label>產品登記</label><div class="dv">{rg_c} 筆</div></div>
</div></div>'''

    html = f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>正瀚生技｜智財與登記管理</title>
<style>{css}</style>
</head>
<body class="locked">
<div id="app-lock" class="lock-screen">
  <div class="lock-card">
    <div class="lock-logo">CH BIOTECH</div>
    <div class="lock-sub">智財與登記管理系統</div>
    <form onsubmit="unlockApp(event)">
      <input id="app-password" type="password" autocomplete="current-password" placeholder="請輸入密碼" autofocus>
      <button type="submit">進入系統</button>
      <div id="lock-error" class="lock-error"></div>
    </form>
  </div>
</div>
<aside id="sidebar">
  <div class="brand">
    <div class="brand-logo">CH</div>
    <div class="brand-sub">BIOTECH</div>
    <div class="brand-title">正瀚生技<br>智財與登記管理</div>
  </div>
  <nav class="nav">
    <div class="nav-section">管理工作台</div>
    <div class="nav-item active" onclick="showPage('overview',this)"><span style="flex:1">▦ 主管總覽</span></div>
    <div class="nav-item" onclick="showPage('trademark',this)"><span style="flex:1">® 商標管理</span><span class="nav-badge">{tm_c}</span></div>
    <div class="nav-item" onclick="showPage('patent',this)"><span style="flex:1">◇ 專利管理</span><span class="nav-badge">{pt_c}</span></div>
    <div class="nav-item" onclick="showPage('registration',this)"><span style="flex:1">▤ 產品登記</span><span class="nav-badge">{rg_c}</span></div>
    <div class="nav-item" onclick="showPage('alerts',this)"><span style="flex:1">◷ 期限提醒</span><span class="nav-badge alert" id="nba">—</span></div>
    <div class="nav-section" style="margin-top:10px">設定</div>
    <div class="nav-item" onclick="showPage('sync',this)"><span style="flex:1">⇄ 資料與同步</span></div>
  </nav>
  <div class="sidebar-footer">同步：{NOW_STR}<br>© 正瀚生技 CH BIOTECH</div>
</aside>
<main id="main">
  <div class="topbar">
    <div class="topbar-left" id="bc">管理中心 ／ <span>主管總覽</span></div>
    <div id="topbar-actions"></div>
  </div>
  <div class="content" id="content"></div>
</main>

<!-- 案件明細 Modal -->
<div class="mo" id="mo" onclick="closeMo(event)">
  <div class="modal">
    <div class="mh"><h3 id="mt">案件明細</h3><button class="mclose" onclick="closeMo()">✕</button></div>
    <div class="mbody" id="mb"></div>
  </div>
</div>

<!-- 匯出 Modal -->
<div class="mo" id="expMo" onclick="closeExpMo(event)">
  <div class="modal" style="max-width:740px">
    <div class="mh"><h3 id="expTitle">匯出資料</h3><button class="mclose" onclick="closeExpMo()">✕</button></div>
    <div class="mbody">
      <div class="exp-section">
        <h4>模式 1：自選欄位匯出</h4>
        <div class="chk-grid" id="expColList"></div>
        <button class="btn-primary" style="margin-top:8px" onclick="doMode1Export('csv')">↓ 匯出 CSV</button>
        <button class="btn-outline" style="margin-top:8px;margin-left:6px" onclick="doMode1Export('pdf')">匯出 PDF</button>
      </div>
      <hr style="border:none;border-top:1px solid #e8edf5;margin:16px 0">
      <div class="exp-section" id="mode2Section">
        <h4>模式 2：各國登記類別彙總表（自行取得 vs 協助客戶）</h4>
        <button class="btn-outline" onclick="doMode2Preview()">預覽彙總表</button>
        <button class="btn-primary" style="margin-left:8px" onclick="doMode2Export('csv')">↓ 匯出 CSV</button>
        <button class="btn-outline" style="margin-left:6px" onclick="doMode2Export('pdf')">匯出 PDF</button>
        <div id="mode2Preview" style="margin-top:14px;overflow-x:auto"></div>
      </div>
    </div>
  </div>
</div>

<script type="application/json" id="raw-data">{data_js}</script>
<script>
const PASS_HASH = 'c5e8aa9dba7646c26a2241b8fd9d257834b96c6aed1722fdf40cbdd18072195c';
async function unlockApp(e) {{
  e.preventDefault();
  const input = document.getElementById('app-password');
  const msg = document.getElementById('lock-error');
  const bytes = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(input.value));
  const hash = Array.from(new Uint8Array(bytes)).map(b=>b.toString(16).padStart(2,'0')).join('');
  if (hash === PASS_HASH) {{
    document.body.classList.remove('locked');
    document.getElementById('app-lock').remove();
  }} else {{
    msg.textContent = '密碼錯誤，請重新輸入';
    input.value = '';
    input.focus();
  }}
}}
const RAW = JSON.parse(document.getElementById('raw-data').textContent);
const REG_HIDE = new Set({reg_hide_js});
const NOW_STR = '{NOW_STR}';
const SHEET_NAMES = {{
  trademark: '{SHEET_NAMES['trademark']}',
  patent:    '{SHEET_NAMES['patent']}',
  registration: '{SHEET_NAMES['registration']}'
}};

// ── Helpers ──────────────────────────────────────────────────────────────
function esc(s) {{ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }}
function dlCls(s) {{
  const m = {{'正常':'正常','期限已過':'期限已過','即將到期':'即將到期',
    '即將到期(30天)':'30','即將到期(90天)':'90','即將到期(180天)':'180','即將到期(365天)':'365',
    '待補期限':'待補期限','日期異常':'日期異常','N/A':'NA'}};
  return 'dl-' + (m[s] || '待補期限');
}}
function badge(t, c) {{ return `<span class="badge ${{c}}">${{esc(t)}}</span>`; }}
function dlBadge(s) {{ return badge(s, dlCls(s)); }}

// ── State ─────────────────────────────────────────────────────────────────
let pg = 'overview', cur = 1, pp = 30;
let flt = {{q:'', country:'all', tmSt:'all', ptSt:'all', ptType:'all', rgSt:'all', rgType:'all'}};
let srt = {{col:null, asc:true}};

function setF(k, v) {{ flt[k] = v; cur = 1; render(); }}
function resetF() {{
  if (pg==='trademark') Object.assign(flt, {{q:'',country:'all',tmSt:'all'}});
  else if (pg==='patent') Object.assign(flt, {{q:'',country:'all',ptSt:'all',ptType:'all'}});
  else if (pg==='registration') Object.assign(flt, {{q:'',country:'all',rgSt:'all',rgType:'all'}});
  cur = 1; render();
}}
function sortBy(c) {{ srt.col===c ? srt.asc=!srt.asc : (srt.col=c, srt.asc=true); render(); }}
function goP(n) {{ cur=n; render(); document.getElementById('main').scrollTo(0,0); }}

// ── Page routing ──────────────────────────────────────────────────────────
const PAGE_TITLES = {{overview:'主管總覽',trademark:'商標管理',patent:'專利管理',
  registration:'產品登記',alerts:'期限提醒',sync:'資料與同步'}};

function showPage(n, el) {{
  pg = n; cur = 1;
  Object.assign(flt, {{q:'',country:'all',tmSt:'all',ptSt:'all',ptType:'all',rgSt:'all',rgType:'all'}});
  srt = {{col:null, asc:true}};
  document.querySelectorAll('.nav-item').forEach(x => x.classList.remove('active'));
  if (el) el.classList.add('active');
  document.getElementById('bc').innerHTML = '管理中心 ／ <span>' + PAGE_TITLES[n] + '</span>';
  render();
}}

function render() {{
  const el = document.getElementById('content');
  const actions = document.getElementById('topbar-actions');
  actions.innerHTML = '';
  if (pg === 'overview')      el.innerHTML = renderOv();
  else if (pg === 'trademark')  {{ el.innerHTML = renderTrademark(); actions.innerHTML = `<button class="btn-outline" onclick="openExpMo('trademark')">↓ 匯出</button>`; }}
  else if (pg === 'patent')     el.innerHTML = renderPatent();
  else if (pg === 'registration') {{
    el.innerHTML = renderRegistration();
    actions.innerHTML = `<button class="btn-outline" onclick="openExpMo('registration')">↓ 匯出</button>`;
  }}
  else if (pg === 'alerts')    el.innerHTML = renderAlerts();
  else if (pg === 'sync')      el.innerHTML = `{sync_section}`;
}}

// ── Pager ─────────────────────────────────────────────────────────────────
function mkPager(total, pages) {{
  if (pages <= 1) return '';
  let btns = '', prev = -1;
  for (let i = 1; i <= pages; i++) {{
    if (i===1 || i===pages || Math.abs(i-cur)<=2) {{
      if (prev !== -1 && i-prev > 1) btns += '<span class="pgi">…</span>';
      btns += `<button class="pgb${{i===cur?' apg':''}}" onclick="goP(${{i}})">${{i}}</button>`;
      prev = i;
    }}
  }}
  return `<div class="pgbar">
    <button class="pgb" onclick="goP(${{cur-1}})" ${{cur===1?'disabled':''}}>‹</button>
    ${{btns}}
    <button class="pgb" onclick="goP(${{cur+1}})" ${{cur===pages?'disabled':''}}>›</button>
    <span class="pgi">${{(cur-1)*pp+1}}–${{Math.min(cur*pp,total)}}/${{total}}</span>
  </div>`;
}}

// ── Sort ──────────────────────────────────────────────────────────────────
function applySort(d) {{
  if (!srt.col) return d;
  return [...d].sort((a,b) => {{
    const va = a[srt.col]||'', vb = b[srt.col]||'';
    return srt.asc ? (va<vb?-1:va>vb?1:0) : (va>vb?-1:va<vb?1:0);
  }});
}}

// ── Overview ──────────────────────────────────────────────────────────────
function buildAll() {{
  const a = [];
  RAW.trademark.forEach(r => a.push({{
    _type:'商標', _name: r['商標']||r['商標案件']||r['商標名稱']||'—',
    _country: r['國別']||'—', _status: r._status,
    _deadline: r._end_date||'', _dl: r._deadline_status,
    _appNo: r['申請號']||r['申請案號']||'', _sub: r['商標分類']||'', _raw: r
  }}));
  RAW.patent.forEach(r => a.push({{
    _type:'專利', _name: r['專利名稱(中文)']||'—',
    _country: r['國別']||'—', _status: r._status,
    _deadline: r._end_date||'', _dl: r._deadline_status,
    _appNo: r['申請案號']||'', _sub: r['專利類別']||'', _raw: r
  }}));
  RAW.registration.forEach(r => a.push({{
    _type:'產品登記', _name: r['登記產品名']||'—',
    _country: r['國別']||'—', _status: r._status,
    _deadline: r._end_date||'', _dl: r._deadline_status,
    _appNo: '', _sub: r['登記類別']||r['登記公司']||'', _raw: r
  }}));
  return a;
}}
const ALL = buildAll();

// Alert count
const ALERT_DL = new Set(['期限已過','即將到期','即將到期(30天)','即將到期(90天)','即將到期(180天)','即將到期(365天)']);
document.getElementById('nba').textContent = ALL.filter(r => ALERT_DL.has(r._dl)).length;

function renderOv() {{
  const tm = RAW.trademark, pt = RAW.patent, rg = RAW.registration;
  const tmGet = tm.filter(r=>r._status==='註冊案' || String(r['狀態/進度說明']||'').includes('已取證') || String(r['狀態/進度說明']||'').includes('已取得')).length;
  const tmApply = tm.filter(r=>{{ const s=String(r['狀態/進度說明']||''); return s.includes('核駁案') || s.includes('審查中'); }}).length;
  const ptGet = pt.filter(r=>r._status==='已取得').length;
  const rgGet = rg.filter(r=>r._status==='已取得').length;
  const tmOver = tm.filter(r=>r._deadline_status==='期限已過').length;
  const tmSoon = tm.filter(r=>r._deadline_status==='即將到期').length;
  const rgSoon = rg.filter(r=>ALERT_DL.has(r._deadline_status)).length;
  const alertN = ALL.filter(r=>ALERT_DL.has(r._dl)).length;
  const pri = ALL.filter(r=>ALERT_DL.has(r._dl))
    .sort((a,b)=>(['期限已過','即將到期','即將到期(30天)','即將到期(90天)','即將到期(180天)','即將到期(365天)'].indexOf(a._dl)||9)
                -(['期限已過','即將到期','即將到期(30天)','即將到期(90天)','即將到期(180天)','即將到期(365天)'].indexOf(b._dl)||9));

  const priRows = pri.length===0
    ? '<div class="plist"><div class="empty">目前無需立即關注的案件 ✓</div></div>'
    : '<div class="plist"><div class="pr prh"><div>案件名稱</div><div>類型・國別</div><div>期限</div><div>提醒</div></div>'
      + pri.map(r=>`<div class="pr" onclick='openMo(${{JSON.stringify(JSON.stringify(r))}})'">
          <div><div class="cn">${{esc(r._name)}}</div><div class="cs">${{esc(r._sub)}}</div></div>
          <div style="font-size:12px">${{esc(r._type+'·'+r._country)}}</div>
          <div style="font-size:12px;color:#6b7a99">${{esc(r._deadline||'—')}}</div>
          <div>${{dlBadge(r._dl)}}</div>
        </div>`).join('')+'</div>';

  return `<div class="ov-grid">
    <div class="card"><div class="card-label">® 商標</div><div class="card-value">${{tm.length}}</div><div class="card-sub">已取證 ${{tmGet}} ／ 申請中 ${{tmApply}}</div></div>
    <div class="card"><div class="card-label">◇ 專利</div><div class="card-value">${{pt.length}}</div><div class="card-sub">已取得 ${{ptGet}} ／ 申請中 ${{pt.length-ptGet}}</div></div>
    <div class="card"><div class="card-label">▤ 產品登記</div><div class="card-value">${{rg.length}}</div><div class="card-sub">已取得 ${{rgGet}} ／ 辦理中 ${{rg.length-rgGet}}</div></div>
  </div>
  <div class="section-title">優先關注事項</div>
  ${{priRows}}`;
}}

// ── Trademark ─────────────────────────────────────────────────────────────
function renderTrademark() {{
  let d = [...RAW.trademark];
  if (flt.q) {{
    const q = flt.q.toLowerCase();
    d = d.filter(r =>
      (r['商標']||r['商標案件']||r['商標名稱']||'').toLowerCase().includes(q) ||
      (r['申請號']||r['申請案號']||'').toLowerCase().includes(q) ||
      (r['國別']||'').toLowerCase().includes(q));
  }}
  if (flt.country!=='all') d = d.filter(r=>r['國別']===flt.country);
  if (flt.tmSt!=='all') d = d.filter(r=>(r['狀態/進度說明']||r._status)===flt.tmSt);
  d = applySort(d);

  const total = d.length, pages = Math.ceil(total/pp)||1;
  if (cur>pages) cur=pages;
  const rows = d.slice((cur-1)*pp, cur*pp);
  const countries = [...new Set(RAW.trademark.map(r=>r['國別']).filter(Boolean))].sort();
  const tmStatuses = [...new Set(RAW.trademark.map(r=>r['狀態/進度說明']||r._status).filter(Boolean))].sort();

  const fbar = `<div class="fbar">
    <input type="text" placeholder="搜尋商標名稱、申請案號…" oninput="setF('q',this.value)" value="${{esc(flt.q)}}">
    <select onchange="setF('country',this.value)">
      <option value="all">所有國別</option>
      ${{countries.map(c=>`<option value="${{esc(c)}}"${{flt.country===c?' selected':''}}>${{esc(c)}}</option>`).join('')}}
    </select>
    <select onchange="setF('tmSt',this.value)">
      <option value="all">所有進度</option>
      ${{tmStatuses.map(s=>`<option value="${{s}}"${{flt.tmSt===s?' selected':''}}>${{s}}</option>`).join('')}}
    </select>
    <span class="frs" onclick="resetF()">重設</span>
    <span class="rcount">共 ${{total}} 筆${{total!==RAW.trademark.length?' (全 '+RAW.trademark.length+')':''}}</span>
  </div>`;

  const tbody = rows.length===0
    ? `<tr><td colspan="7"><div class="empty">無符合條件的案件</div></td></tr>`
    : rows.map(r => {{
        const name = r['商標']||r['商標案件']||r['商標名稱']||'—';
        const imgUrl = r['商標圖案']||r['商標圖示']||r['圖片URL']||r['圖片']||'';
        const imgTag = imgUrl && !String(imgUrl).startsWith('#') ? `<img src="${{esc(imgUrl)}}" onerror="this.style.display='none'" style="width:28px;height:28px;object-fit:contain;vertical-align:middle;margin-right:6px;border-radius:4px">` : '';
        const appNo = r['申請號']||r['申請案號']||'—';
        const regNo = r['註冊編號']||r['註冊號']||r['證書號 (進度)']||r['證書號(進度)']||'—';
        const cls = r['申請類別']||r['類別']||'';
        return `<tr onclick='openMo(${{JSON.stringify(JSON.stringify(r))}})'">
          <td><div style="display:flex;align-items:flex-start">${{imgTag}}<div><div class="tm-name">${{esc(name)}}</div></div></div></td>
          <td style="font-size:12px">${{esc(r['國別']||'—')}}</td>
          <td style="font-size:12px">${{esc(cls)}}</td>
          <td>${{badge(r['狀態/進度說明']||r._status,'b-'+(r['狀態/進度說明']||r._status))}}</td>
          <td style="font-size:12px;color:#4a5568">${{esc(appNo)}}</td>
          <td style="font-size:12px;color:#4a5568">${{esc(regNo)}}</td>
          <td style="font-size:12px;color:#6b7a99">${{esc(r._end_date||'-')}}</td>
        </tr>`;
      }}).join('');

  return fbar + `<div class="twrap"><table>
    <thead><tr>
      <th onclick="sortBy('_name')">商標名</th>
      <th onclick="sortBy('_country')">國別</th>
      <th>申請類別</th>
      <th onclick="sortBy('_status')">狀態/進度說明</th>
      <th>申請案號</th>
      <th>註冊號</th>
      <th onclick="sortBy('_end_date')">使用期限（到期日）</th>
    </tr></thead>
    <tbody>${{tbody}}</tbody>
  </table>${{mkPager(total,pages)}}</div>`;
}}

// ── Patent ────────────────────────────────────────────────────────────────
function renderPatent() {{
  let d = [...RAW.patent];
  if (flt.q) {{
    const q = flt.q.toLowerCase();
    d = d.filter(r =>
      (r['專利名稱(中文)']||'').toLowerCase().includes(q) ||
      (r['申請案號']||'').toLowerCase().includes(q) ||
      (r['專利編號']||'').toLowerCase().includes(q) ||
      (r['國別']||'').toLowerCase().includes(q));
  }}
  if (flt.country!=='all') d = d.filter(r=>r['國別']===flt.country);
  if (flt.ptSt!=='all') d = d.filter(r=>r._status===flt.ptSt);
  if (flt.ptType!=='all') d = d.filter(r=>(r['專利類別']||'')=== flt.ptType);
  d = applySort(d);

  const total = d.length, pages = Math.ceil(total/pp)||1;
  if (cur>pages) cur=pages;
  const rows = d.slice((cur-1)*pp, cur*pp);
  const countries = [...new Set(RAW.patent.map(r=>r['國別']).filter(Boolean))].sort();
  const ptTypes = [...new Set(RAW.patent.map(r=>r['專利類別']).filter(Boolean))].sort();

  const fbar = `<div class="fbar">
    <input type="text" placeholder="搜尋專利名稱、申請案號、專利編號…" oninput="setF('q',this.value)" value="${{esc(flt.q)}}">
    <select onchange="setF('country',this.value)">
      <option value="all">所有國別</option>
      ${{countries.map(c=>`<option value="${{esc(c)}}"${{flt.country===c?' selected':''}}>${{esc(c)}}</option>`).join('')}}
    </select>
    <select onchange="setF('ptType',this.value)">
      <option value="all">所有類別</option>
      ${{ptTypes.map(t=>`<option value="${{esc(t)}}"${{flt.ptType===t?' selected':''}}>${{esc(t)}}</option>`).join('')}}
    </select>
    <select onchange="setF('ptSt',this.value)">
      <option value="all">所有狀態</option>
      ${{['已取得','申請中','已結案'].map(s=>`<option value="${{s}}"${{flt.ptSt===s?' selected':''}}>${{s}}</option>`).join('')}}
    </select>
    <span class="frs" onclick="resetF()">重設</span>
    <span class="rcount">共 ${{total}} 筆${{total!==RAW.patent.length?' (全 '+RAW.patent.length+')':''}}</span>
  </div>`;

  const tbody = rows.length===0
    ? `<tr><td colspan="6"><div class="empty">無符合條件的案件</div></td></tr>`
    : rows.map(r => `<tr onclick='openMo(${{JSON.stringify(JSON.stringify(r))}})'">
        <td><div style="font-weight:500;color:#1a1a2e">${{esc(r['專利名稱(中文)']||'—')}}</div><div class="cs">${{esc(r['專利類別']||'')}}</div></td>
        <td style="font-size:12px">${{esc(r['國別']||'—')}}</td>
        <td style="font-size:12px">${{esc(r['申請案號']||'—')}}</td>
        <td style="font-size:12px">${{esc(r['專利編號']||'—')}}</td>
        <td>${{badge(r._status,'b-'+r._status)}}</td>
        <td style="font-size:12px;color:#6b7a99">${{esc(r._end_date||'—')}}</td>
      </tr>`).join('');

  return fbar + `<div class="twrap"><table>
    <thead><tr>
      <th onclick="sortBy('_name')">專利名稱（中文）</th>
      <th onclick="sortBy('_country')">國別</th>
      <th>申請案號</th>
      <th>專利編號</th>
      <th onclick="sortBy('_status')">目前狀態</th>
      <th onclick="sortBy('_end_date')">證書到期日</th>
    </tr></thead>
    <tbody>${{tbody}}</tbody>
  </table>${{mkPager(total,pages)}}</div>`;
}}

// ── Registration ──────────────────────────────────────────────────────────
function renderRegistration() {{
  let d = [...RAW.registration];
  if (flt.q) {{
    const q = flt.q.toLowerCase();
    d = d.filter(r =>
      (r['登記產品名']||'').toLowerCase().includes(q) ||
      (r['國別']||'').toLowerCase().includes(q) ||
      (r['證書/License ID']||'').toLowerCase().includes(q) ||
      (r['登記公司']||'').toLowerCase().includes(q));
  }}
  if (flt.country!=='all') d = d.filter(r=>r['國別']===flt.country);
  if (flt.rgSt!=='all') d = d.filter(r=>r._status===flt.rgSt);
  if (flt.rgType!=='all') d = d.filter(r=>(r['登記類別']||'')=== flt.rgType);
  d = applySort(d);

  const total = d.length, pages = Math.ceil(total/pp)||1;
  if (cur>pages) cur=pages;
  const rows = d.slice((cur-1)*pp, cur*pp);
  const countries = [...new Set(RAW.registration.map(r=>r['國別']).filter(Boolean))].sort();
  const rgTypes = [...new Set(RAW.registration.map(r=>r['登記類別']).filter(Boolean))].sort();
  const statuses = [...new Set(RAW.registration.map(r=>r._status).filter(Boolean))].sort();

  const fbar = `<div class="fbar">
    <input type="text" placeholder="搜尋產品名稱、國別、證書號…" oninput="setF('q',this.value)" value="${{esc(flt.q)}}">
    <select onchange="setF('country',this.value)">
      <option value="all">所有國別</option>
      ${{countries.map(c=>`<option value="${{esc(c)}}"${{flt.country===c?' selected':''}}>${{esc(c)}}</option>`).join('')}}
    </select>
    ${{rgTypes.length?`<select onchange="setF('rgType',this.value)">
      <option value="all">所有登記類別</option>
      ${{rgTypes.map(t=>`<option value="${{esc(t)}}"${{flt.rgType===t?' selected':''}}>${{esc(t)}}</option>`).join('')}}
    </select>`:''}}<select onchange="setF('rgSt',this.value)">
      <option value="all">所有狀態</option>
      ${{statuses.map(s=>`<option value="${{esc(s)}}"${{flt.rgSt===s?' selected':''}}>${{s}}</option>`).join('')}}
    </select>
    <span class="frs" onclick="resetF()">重設</span>
    <span class="rcount">共 ${{total}} 筆${{total!==RAW.registration.length?' (全 '+RAW.registration.length+')':''}}</span>
  </div>`;

  const tbody = rows.length===0
    ? `<tr><td colspan="6"><div class="empty">無符合條件的案件</div></td></tr>`
    : rows.map(r => `<tr onclick='openMo(${{JSON.stringify(JSON.stringify(r))}})'">
        <td><div class="cn">${{esc(r['登記產品名']||'—')}}</div></td>
        <td style="font-size:12px">${{esc(r['登記類別']||'—')}}</td>
        <td style="font-size:12px">${{esc(r['國別']||'—')}}</td>
        <td style="font-size:12px;max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${{esc(r['登記公司']||'—')}}</td>
        <td>${{badge(r._status,'b-'+r._status)}}</td>
        <td style="font-size:12px;color:#6b7a99">${{esc(r._end_date||'—')}}</td>
      </tr>`).join('');

  return fbar + `<div class="twrap"><table>
    <thead><tr>
      <th onclick="sortBy('_name')">登記產品名</th>
      <th onclick="sortBy('_sub')">登記類別</th>
      <th onclick="sortBy('_country')">國別</th>
      <th>登記公司</th>
      <th onclick="sortBy('_status')">狀態</th>
      <th onclick="sortBy('_end_date')">有效期限</th>
    </tr></thead>
    <tbody>${{tbody}}</tbody>
  </table>${{mkPager(total,pages)}}</div>`;
}}

// ── Alerts ────────────────────────────────────────────────────────────────
function renderAlerts() {{
  const RANKS = ['期限已過','即將到期','即將到期(30天)','即將到期(90天)','即將到期(180天)','即將到期(365天)'];
  const LABELS = {{'期限已過':'⚠ 期限已過','即將到期':'🟠 6個月內到期','即將到期(30天)':'🔴 30天內到期',
    '即將到期(90天)':'🟠 90天內到期','即將到期(180天)':'🟡 180天內到期','即將到期(365天)':'🟡 365天內到期'}};
  let html = RANKS.map(dl => {{
    const items = ALL.filter(r=>r._dl===dl);
    if (!items.length) return '';
    return `<div style="margin-bottom:20px"><div class="section-title">${{LABELS[dl]}} (${{items.length}})</div>
      <div class="plist"><div class="pr prh"><div>案件名稱</div><div>類型・國別</div><div>期限日期</div><div>提醒</div></div>
      ${{items.map(r=>`<div class="pr" onclick='openMo(${{JSON.stringify(JSON.stringify(r))}})'">
        <div><div class="cn">${{esc(r._name)}}</div></div>
        <div style="font-size:12px">${{esc(r._type+'·'+r._country)}}</div>
        <div style="font-size:12px;color:#6b7a99">${{esc(r._deadline||'—')}}</div>
        <div>${{dlBadge(r._dl)}}</div>
      </div>`).join('')}}
      </div></div>`;
  }}).join('');
  return html || '<div class="empty" style="padding:60px">目前無需關注的到期案件 ✓</div>';
}}

// ── Modal ─────────────────────────────────────────────────────────────────
function openMo(s) {{
  let r; try {{ r = JSON.parse(s); }} catch {{ return; }}
  const raw = r._raw || r;
  const type = r._type || '';
  document.getElementById('mt').textContent = type + ' 案件明細';
  const SKIP = new Set(['_status','_end_date','_deadline_status','_start_date','_type','_name','_country','_deadline','_dl','_appNo','_sub','_raw','_cert']);
  const FULL_COLS = new Set(['專利名稱(中文)','專利名稱(英文)','Raw Materials','備註','說明','目前狀態']);

  let topHtml = `<div class="dg" style="margin-bottom:14px">`;
  if (type==='商標') {{
    const tmDisplayStatus = raw['狀態/進度說明'] || raw['進度狀況'] || raw['進度狀態'] || r._status || '—';
    topHtml += `<div class="di"><label>類型</label><div class="dv">${{badge(type,'b-T')}}</div></div>
    <div class="di"><label>狀態</label><div class="dv">${{badge(tmDisplayStatus,'b-'+(r._status||''))}}</div></div>
    <div class="di"><label>使用起始日</label><div class="dv">${{esc(raw._start_date||'—')}}</div></div>
    <div class="di"><label>使用到期日</label><div class="dv">${{esc(raw._end_date||'—')}} ${{dlBadge(raw._deadline_status)}}</div></div>`;
  }} else {{
    topHtml += `<div class="di"><label>狀態</label><div class="dv">${{badge(r._status||'—','b-'+(r._status||''))}}</div></div>`;
  }}
  topHtml += '</div><hr style="border:none;border-top:1px solid #f0f4fa;margin:4px 0 12px"><div class="dg">';

  const hideSet = type==='產品登記' ? REG_HIDE : new Set();
  let items;
  if (type==='商標') {{
    const tmCols = [
      ['商標案件', raw['商標案件']||raw['商標']||raw['商標名稱']],
      ['商標圖案', raw['商標圖案']||raw['商標圖示']||raw['圖片URL']||raw['圖片']],
      ['商標類型', raw['商標分類']],
      ['國別', raw['國別']],
      ['類別', raw['申請類別']||raw['類別']],
      ['申請案號', raw['申請案號']||raw['申請號']],
      ['申請日期', raw['申請日期']],
      ['狀態/進度說明', raw['狀態/進度說明']||raw['進度狀況']||raw['進度狀態']],
      ['註冊號', raw['註冊編號']||raw['註冊號']||raw['證書號 (進度)']||raw['證書號(進度)']],
      ['商標標示', raw['商標標示']],
      ['使用期間-取得日期', raw['使用期間-起始']||raw['_start_date']],
      ['使用期間-到期日期', raw['使用期間-到期']||raw['_end_date']],
    ];
    items = tmCols.map(([k,v]) => `<div class="di${{k==='商標圖案'?' full':''}}"><label>${{k}}</label><div class="dv">${{esc((v && !String(v).startsWith('#'))?v:'—')}}</div></div>`).join('');
  }} else {{
    items = Object.entries(raw).filter(([k,v]) => !SKIP.has(k) && !hideSet.has(k.toLowerCase()))
      .map(([k,v]) => `<div class="di${{FULL_COLS.has(k)?' full':''}}"><label>${{esc(k)}}</label><div class="dv">${{esc(v||'—')}}</div></div>`)
      .join('');
  }}
  document.getElementById('mb').innerHTML = topHtml + items + '</div>';
  document.getElementById('mo').classList.add('open');
}}
function closeMo(e) {{ if (!e||e.target===document.getElementById('mo')) document.getElementById('mo').classList.remove('open'); }}

// ── Export ────────────────────────────────────────────────────────────────
let expType = 'registration';
function openExpMo(kind) {{
  expType = kind || 'registration';
  if (expType === 'trademark') {{
    const keys = ['商標分類','商標案件','商標圖案','國別','申請案號','申請日期','申請類別','狀態/進度說明','註冊編號','商標標示','使用期間-起始','使用期間-到期'];
    document.getElementById('expTitle').textContent = '匯出商標資料';
    document.querySelector('#mode2Section h4').textContent = '模式 2：商標狀態／國別數量彙總';
    document.getElementById('expColList').innerHTML = keys.map(k => `<label class="chk-item"><input type="checkbox" data-key="${{esc(k)}}" checked> ${{esc(k)}}</label>`).join('');
    document.getElementById('mode2Preview').innerHTML = renderTrademarkSummary();
    document.getElementById('expMo').classList.add('open');
    return;
  }}
  document.getElementById('expTitle').textContent = '匯出產品登記資料';
  // Build column checkboxes from registration data
  const allKeys = new Set();
  RAW.registration.forEach(r => Object.keys(r).forEach(k => {{ if (!k.startsWith('_') && !REG_HIDE.has(k.toLowerCase())) allKeys.add(k); }}));
  const suggested = ['登記產品名','登記類別','國別','登記公司','證書/License ID','取得日期','進度','證書有效期限'];
  const ordered = [...suggested.filter(k=>allKeys.has(k)), ...[...allKeys].filter(k=>!suggested.includes(k))];
  const defaultOn = new Set(['登記產品名','登記類別','國別','登記公司','狀態','進度','證書有效期限']);

  document.getElementById('expColList').innerHTML = ordered.map(k =>
    `<label class="chk-item"><input type="checkbox" data-key="${{esc(k)}}" ${{defaultOn.has(k)?'checked':''}}> ${{esc(k)}}</label>`
  ).join('');
  document.getElementById('mode2Preview').innerHTML = '';
  document.getElementById('expMo').classList.add('open');
}}
function closeExpMo(e) {{ if (!e||e.target===document.getElementById('expMo')) document.getElementById('expMo').classList.remove('open'); }}

function exportRows() {{
  if (expType==='trademark') {{
    let d=[...RAW.trademark],q=(flt.q||'').toLowerCase();
    if(q)d=d.filter(r=>[r['商標案件'],r['申請案號'],r['國別'],r['註冊編號']].some(v=>String(v||'').toLowerCase().includes(q)));
    if(flt.country!=='all')d=d.filter(r=>r['國別']===flt.country);
    if(flt.tmSt!=='all')d=d.filter(r=>(r['狀態/進度說明']||r._status)===flt.tmSt);
    return d;
  }}
  let d=[...RAW.registration],q=(flt.q||'').toLowerCase();
  if(q)d=d.filter(r=>[r['登記產品名'],r['國別'],r['證書/License ID'],r['登記公司']].some(v=>String(v||'').toLowerCase().includes(q)));
  if(flt.country!=='all')d=d.filter(r=>r['國別']===flt.country);
  if(flt.rgSt!=='all')d=d.filter(r=>r._status===flt.rgSt);
  if(flt.rgType!=='all')d=d.filter(r=>r['登記類別']===flt.rgType);
  return d;
}}
function csvCell(v) {{ return '"'+String(v??'').replace(/"/g,'""')+'"'; }}
function exportPDF(title,cols,rows) {{
  const body=rows.map(r=>'<tr>'+cols.map(c=>'<td>'+esc(r[c]||'—')+'</td>').join('')+'</tr>').join('');
  const w=window.open('','_blank');if(!w){{alert('請允許瀏覽器開啟彈出視窗後再匯出 PDF');return;}}
  w.document.write('<!doctype html><html><head><meta charset="utf-8"><title>'+esc(title)+'</title><style>body{{font-family:Arial,"Noto Sans TC",sans-serif;padding:24px;color:#222}}table{{border-collapse:collapse;width:100%;font-size:11px}}th,td{{border:1px solid #bbb;padding:5px;text-align:left;vertical-align:top}}th{{background:#eef2f7}}</style></head><body><h1>'+esc(title)+'</h1><p>匯出時間：'+esc(NOW_STR)+'；共 '+rows.length+' 筆</p><table><thead><tr>'+cols.map(c=>'<th>'+esc(c)+'</th>').join('')+'</tr></thead><tbody>'+body+'</tbody></table></body></html>');
  w.document.close();w.focus();setTimeout(()=>w.print(),250);
}}

function doMode1Export(fmt) {{
  const cols = [...document.querySelectorAll('#expColList input:checked')].map(el => el.dataset.key);
  const rows = exportRows();
  if (!cols.length) {{ alert('請至少勾選一個欄位'); return; }}
  const prefix = expType==='trademark' ? '正瀚_商標' : '正瀚_產品登記';
  if (fmt==='pdf') {{ exportPDF(prefix+'明細',cols,rows); return; }}
  const csv = '﻿' + [cols.join(','),
    ...rows.map(r => cols.map(c=>'"'+(r[c]||'').replace(/"/g,'""')+'"').join(','))
  ].join('\\n');
  dlCSV(csv, prefix+'_{TODAY_STR}'.replace(/[/]/g,'') + '.csv');
}}

function doMode2Preview() {{
  document.getElementById('mode2Preview').innerHTML = expType==='trademark' ? renderTrademarkSummary() : renderMode2HTML(buildMode2Table());
}}

function doMode2Export(fmt) {{
  if (expType==='trademark') {{
    const tbl=buildTrademarkSummary(), cols=['國別','已取證','申請中','放棄案','合計'];
    if(fmt==='pdf'){{exportPDF('商標狀態／國別彙總',cols,tbl.exportRows);return;}}
    const csv='﻿'+[cols.map(csvCell).join(','),...tbl.exportRows.map(r=>cols.map(c=>csvCell(r[c])).join(','))].join('\\n');
    dlCSV(csv,'正瀚_商標狀態彙總_{TODAY_STR}'.replace(/[/]/g,'')+'.csv');
    return;
  }}
  const tbl = buildMode2Table();
  const {{countries, types, data}} = tbl;
  const header = ['國別/登記類別', ...types, '合計'];
  const rows = countries.map(c => {{
    const row = [c];
    let tot = 0;
    types.forEach(t => {{
      const cell = (data[c]&&data[c][t]) || {{self:0, help:0}};
      row.push(`自行${{cell.self}}/協助${{cell.help}}`);
      tot += cell.self + cell.help;
    }});
    row.push(tot);
    return row;
  }});
  const totalRow=['總計'];
  types.forEach(t=>{{
    let self=0,help=0;
    countries.forEach(c=>{{const x=(data[c]&&data[c][t])||{{self:0,help:0}};self+=x.self;help+=x.help;}});
    totalRow.push('自行'+self+'/協助'+help);
  }});
  totalRow.push(countries.reduce((n,c)=>n+types.reduce((m,t)=>{{const x=(data[c]&&data[c][t])||{{self:0,help:0}};return m+x.self+x.help;}},0),0));
  rows.push(totalRow);
  if(fmt==='pdf'){{const pdfRows=rows.map(row=>Object.fromEntries(header.map((h,i)=>[h,row[i]])));exportPDF('產品登記各國類別彙總',header,pdfRows);return;}}
  const csv = '﻿' + [header.join(','), ...rows.map(r=>r.map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(','))].join('\\n');
  dlCSV(csv, '正瀚_產品登記彙總_{TODAY_STR}'.replace(/[/]/g,'') + '.csv');
}}

function buildTrademarkSummary() {{
  const labels=['已取證','申請中','放棄案'], data={{}};
  const normalize=s=>{{s=String(s||'');if(s.includes('已取證')||s.includes('已取得'))return '已取證';if(s.includes('放棄')||s.includes('失效'))return '放棄案';return '申請中';}};
  RAW.trademark.forEach(r=>{{const c=r['國別']||'未填寫',t=normalize(r['狀態/進度說明']||r._status);if(!data[c])data[c]={{'已取證':0,'申請中':0,'放棄案':0}};data[c][t]++;}});
  const countries=Object.keys(data).sort(), exportRows=countries.map(c=>{{const r={{國別:c}};labels.forEach(t=>r[t]=data[c][t]);r['合計']=labels.reduce((n,t)=>n+r[t],0);return r;}});
  const total={{國別:'總計'}};labels.forEach(t=>total[t]=exportRows.reduce((n,r)=>n+r[t],0));total['合計']=labels.reduce((n,t)=>n+total[t],0);exportRows.push(total);
  return {{exportRows,exportCols:['國別',...labels,'合計']}};
}}
function renderTrademarkSummary() {{
  const tbl=buildTrademarkSummary(), rows=tbl.exportRows.map(r=>'<tr><td class="rc">'+esc(r['國別'])+'</td><td>'+r['已取證']+'</td><td>'+r['申請中']+'</td><td>'+r['放棄案']+'</td><td><strong>'+r['合計']+'</strong></td></tr>').join('');
  return '<table class="mode2-tbl"><thead><tr><th class="rc">國別</th><th>已取證</th><th>申請中</th><th>放棄案</th><th>合計</th></tr></thead><tbody>'+rows+'</tbody></table><div style="font-size:11px;color:#8899bb;margin-top:6px">申請中包含審查中與核駁案</div>';
}}

function buildMode2Table() {{
  const CH = 'CH Biotech R&D Co., Ltd';
  const countries = [...new Set(RAW.registration.map(r=>r['國別']).filter(Boolean))].sort();
  const types = [...new Set(RAW.registration.map(r=>r['登記類別']).filter(Boolean))].sort();
  const data = {{}};
  RAW.registration.forEach(r => {{
    const c = r['國別']||'', t = r['登記類別']||'', co = r['登記公司']||'';
    if (!c||!t) return;
    if (!data[c]) data[c] = {{}};
    if (!data[c][t]) data[c][t] = {{self:0, help:0}};
    if (co === CH) data[c][t].self++; else data[c][t].help++;
  }});
  return {{countries, types, data}};
}}

function renderMode2HTML({{countries, types, data}}) {{
  const hdr = `<tr><th class="rc">國別 ＼ 登記類別</th>${{types.map(t=>`<th>${{esc(t)}}</th>`).join('')}}<th>合計</th></tr>`;
  const rows = countries.map(c => {{
    let tot = 0;
    const cells = types.map(t => {{
      const cell = (data[c]&&data[c][t])||{{self:0,help:0}};
      tot += cell.self+cell.help;
      if (!cell.self && !cell.help) return '<td>—</td>';
      return `<td>自行 ${{cell.self}}<br>協助 ${{cell.help}}</td>`;
    }}).join('');
    return `<tr><td class="rc">${{esc(c)}}</td>${{cells}}<td><strong>${{tot}}</strong></td></tr>`;
  }}).join('');
  const totals=types.map(t=>{{let self=0,help=0;countries.forEach(c=>{{const x=(data[c]&&data[c][t])||{{self:0,help:0}};self+=x.self;help+=x.help;}});return '<td><strong>自行'+self+'/協助'+help+'</strong></td>';}}).join('');
  const grand=countries.reduce((n,c)=>n+types.reduce((m,t)=>{{const x=(data[c]&&data[c][t])||{{self:0,help:0}};return m+x.self+x.help;}},0),0);
  return `<table class="mode2-tbl"><thead>${{hdr}}</thead><tbody>${{rows}}<tr><td class="rc"><strong>總計</strong></td>${{totals}}<td><strong>${{grand}}</strong></td></tr></tbody></table>
    <div style="font-size:11px;color:#8899bb;margin-top:6px">自行 = CH Biotech R&amp;D Co., Ltd；協助 = 非 CH Biotech（協助客戶取得）</div>`;
}}

function dlCSV(csv, name) {{
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], {{type:'text/csv;charset=utf-8'}}));
  a.download = name; a.click();
}}

// ── Init ──────────────────────────────────────────────────────────────────
render();
</script>
</body>
</html>'''
    return html.replace('</head>', '<style>' + HUB_CSS + '</style></head>', 1).replace('render();\n</script>', HUB_JS + '\n</script>')

HUB_CSS = r'''
:root{--green:#17715e;--ink:#17362f;--line:#dce7e3}
body{background:#f2f6f5;color:var(--ink)}
#sidebar{background:linear-gradient(180deg,#0a342c,#0c4035);width:236px;min-width:236px}
.brand,.sidebar-footer{border-color:#ffffff20}.brand-title{color:#b5d4ca}.brand-sub{color:#9ecabb}
.nav-item{border:0;border-radius:10px;margin:4px 10px;color:#c3d9d2;padding:12px}
.nav-item.active,.nav-item:hover{background:#5bddbb25;color:white;box-shadow:inset 3px 0 #5bd8ba}
.nav-badge{background:#ffffff20;color:#d5f8ed}.nav-section{color:#9abeb1}
.topbar{gap:14px;flex-wrap:wrap;padding:14px 24px}.content{width:100%;max-width:1580px;margin:auto;padding:24px}
.topbar-left{display:none}.panel,.sc,.fbar,.twrap{background:white;border:1px solid var(--line);border-radius:16px;box-shadow:0 5px 22px #123d3008}
button{font:inherit;cursor:pointer}.btn-primary{background:var(--green)}.btn-outline{color:var(--green);border-color:var(--line)}
.hub-head{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:20px}.hub-head h1{font-size:25px}.hub-head small{color:#768b82;font-size:11px}
.hub-filters{padding:14px;display:flex;flex-wrap:wrap;gap:10px;margin-bottom:20px}
.hub-filter{position:relative}.hub-filter summary{cursor:pointer;padding:8px 12px;border:1px solid var(--line);border-radius:9px;list-style:none}
.hub-options{position:absolute;top:44px;left:0;width:280px;max-height:330px;overflow:auto;padding:12px;background:white;border:1px solid var(--line);border-radius:12px;z-index:30;box-shadow:0 10px 30px #123d3020}
.hub-options label{display:flex;align-items:start;gap:8px;padding:7px;overflow-wrap:anywhere}.hub-options button{margin:0 8px 7px 0}
.hub-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px;margin-bottom:24px}.panel{min-width:0;padding:18px}.panel h2{font-size:17px;margin-bottom:16px}
.chart-mode{float:right;font-size:12px;border:1px solid var(--line);padding:5px;border-radius:6px;background:white}
.donut{display:block;margin:0 auto 16px;width:180px;height:180px}.donut path{cursor:pointer;stroke:white;stroke-width:1}.donut path:hover{opacity:.7}.donut text{fill:var(--ink);font-size:22px;font-weight:700}
.legend-row{display:grid;grid-template-columns:10px minmax(0,1fr) auto;gap:8px;align-items:center;background:none;border:0;border-bottom:1px solid #edf3ef;padding:9px 0;width:100%;text-align:left;color:var(--ink);font-size:13px}.legend-row i{width:8px;height:8px;border-radius:50%}.legend-row span{overflow-wrap:anywhere}.legend-row small{white-space:nowrap;color:#687e75}
.bar-track{height:8px;background:#edf3ef;border-radius:8px;overflow:hidden;margin:0 0 6px}.bar-track i{display:block;height:100%}
.twrap{max-height:66vh;overflow:auto}.twrap table{min-width:880px}.twrap th{position:sticky;top:0;z-index:2;background:#f5faf7;color:#607b6f;font-size:12px}.twrap td{white-space:nowrap;line-height:1.6;font-size:13px}.twrap td:first-child{min-width:220px;max-width:420px;white-space:normal;overflow-wrap:anywhere}.twrap tr{cursor:pointer}.hub-alerts{max-height:none}.hub-alerts table{min-width:780px}
.hub-search{position:relative;flex:1;max-width:620px}.hub-search input{width:100%;padding:11px 14px;border:1px solid var(--line);background:#f7faf9;border-radius:11px;font:inherit}.search-pop{position:absolute;top:48px;left:0;right:0;background:white;box-shadow:0 12px 35px #163d3025;border:1px solid var(--line);border-radius:12px;max-height:65vh;overflow:auto;z-index:40}.search-pop:empty{display:none}.search-pop button{display:block;width:100%;background:white;border:0;border-bottom:1px solid var(--line);text-align:left;padding:12px;color:var(--ink)}.search-pop small{display:block;color:#74877e;margin-top:4px}
#mo{justify-content:flex-end}#mo .modal{height:100vh;max-height:100dvh;width:560px;max-width:95vw;border-radius:18px 0 0 18px;transform:translateX(100%);transition:transform .2s}#mo.open .modal{transform:translateX(0)}.drawer-section{margin-bottom:24px}.drawer-section h3{font-size:15px;margin-bottom:12px}.dg{grid-template-columns:repeat(2,minmax(0,1fr))}.di{min-width:0;background:#f5faf7;border:1px solid var(--line);padding:10px;border-radius:9px}.di .dv{overflow-wrap:anywhere;white-space:pre-wrap;color:var(--ink)}.drawer-section summary{cursor:pointer;font-weight:600;margin-bottom:12px}
.lock-screen{width:100%;background:#f2f6f5}.lock-card button{background:var(--green)}.pgbar{flex-wrap:wrap}.muted{color:#71877e;font-size:12px;margin-bottom:12px}
@media(max-width:1500px){.hub-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
.custom-axis{display:flex;align-items:center;gap:16px;flex-wrap:wrap;margin:12px 0 20px}.custom-axis select{padding:8px;border:1px solid #d6e6e1;border-radius:7px}.hub-grid .chart-mode{float:none;margin-bottom:10px}.hub-grid .legend-row span{overflow-wrap:anywhere}
@media(max-width:700px){.hub-grid{grid-template-columns:1fr}.panel{padding:18px}.content{padding:18px}}
@media(max-width:900px){#sidebar{width:100%;min-width:0}.brand{display:block;padding:12px 18px}.brand-logo{font-size:18px}.brand-sub{display:none}.brand-title{margin:4px 0}.nav-item{margin:0;min-height:44px}.topbar{padding:10px 14px}.hub-head{align-items:flex-start}}
@media(max-width:600px){.content{padding:12px}.hub-head{display:block}.hub-head h1{font-size:22px;margin-bottom:6px}.hub-filters{gap:6px}.hub-options{position:fixed;left:12px;right:12px;top:220px;width:auto}.hub-search input{font-size:16px}.dg{grid-template-columns:1fr}#mo .modal{max-width:100vw;border-radius:0}.hub-alerts td{font-size:12px}}
'''

HUB_JS = r'''
// Intelligence Hub: all views derive from the same immutable source records.
const TYPE_NAMES={trademark:'商標',patent:'專利',registration:'產品登記'};
const countryAliases={'taiwan':'台灣','tw':'台灣','台湾':'台灣','usa':'美國','us':'美國','united states':'美國','united states of america':'美國','japan':'日本','jp':'日本','china':'中國','cn':'中國','brazil':'巴西','br':'巴西','eu':'歐盟','european union':'歐盟','europe':'歐洲','india':'印度','indonesia':'印尼','vietnam':'越南','thailand':'泰國','malaysia':'馬來西亞','canada':'加拿大','australia':'澳洲','korea':'韓國','south korea':'韓國','mexico':'墨西哥','argentina':'阿根廷','chile':'智利','peru':'秘魯','colombia':'哥倫比亞','south africa':'南非','new zealand':'紐西蘭','philippines':'菲律賓','singapore':'新加坡','germany':'德國','france':'法國','spain':'西班牙','italy':'義大利','uk':'英國','united kingdom':'英國','bulgaria':'保加利亞','russia':'俄羅斯','pct':'PCT（國際申請）'};
function zhCountry(v){v=String(v||'未填寫').trim();if(/[\u3400-\u9fff]/.test(v))return v.replace(/\s*[（(][A-Za-z\s]+[)）]/g,'').replace('台湾','台灣').replace('臺灣','台灣').replace('澳大利亞','澳洲');return countryAliases[v.toLowerCase()]||v;}
function rawStatus(r,t){const keys=t==='trademark'?['狀態/進度說明','進度狀況','進度狀態']:t==='patent'?['目前狀態','狀態','進度']:['進度','狀態'];return String(keys.map(k=>r[k]).find(v=>v&&String(v).trim())||'未填寫').trim();}
const HUB=Object.entries(RAW).flatMap(([t,rows])=>rows.map((r,i)=>({id:t+'-'+i,t,raw:r,c:zhCountry(r['國別']),s:rawStatus(r,t),name:r['商標案件']||r['商標']||r['商標名稱']||r['專利名稱(中文)']||r['登記產品名']||'未命名',date:r._end_date||'',numbers:['申請案號','申請號','註冊編號','註冊號','專利編號','證書/License ID'].map(k=>r[k]||'').join(' ')})));
const byId=new Map(HUB.map(r=>[r.id,r]));
const CHOICES={country:[...new Set(HUB.map(r=>r.c))].sort((a,b)=>a.localeCompare(b,'zh-Hant')),type:Object.keys(TYPE_NAMES),status:['已完成','進行中','需處理','已放棄']};
let selection=Object.fromEntries(Object.entries(CHOICES).map(([k,v])=>[k,new Set(v)]));
let chartModes={trademark:'donut',patent:'donut',registration:'donut'},drill=null,globalQuery='',lastFocus=null;
esc=s=>String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
function selected(r){return selection.country.has(r.c)&&selection.type.has(r.t)&&selection.status.has(managementStatus(r));}
function rank(r){if(managementStatus(r)==='已完成')return 0;if(managementStatus(r)==='進行中')return 1;if(/核駁/.test(r.s))return 2;if(/結案/.test(r.s))return 3;if(/放棄|失效|撤回/.test(r.s))return 4;return 5;}
function dateValue(r){return /^\d{4}-\d{2}-\d{2}$/.test(r.date)?Date.parse(r.date):Infinity;}
function defaultOrder(a,b){return rank(a)-rank(b)||(dateValue(a)-dateValue(b)||0);}
function attention(r){const d=dateValue(r);if(Number.isFinite(d)){const days=Math.ceil((d-Date.parse(NOW_STR.slice(0,10).replaceAll('/','-')))/86400000);if(days<0)return [0,'期限已過'];for(const [i,n]of [30,90,180,365].entries())if(days<=n)return [i+1,n+' 天內到期'];}if(/審查|申請中|申請案|辦理中/.test(r.s))return [5,'進度追蹤'];if(/核駁/.test(r.s))return [6,'核駁案'];if(!r.date)return [7,'待補期限'];if(r.raw._deadline_status==='日期異常')return [8,'日期異常'];return null;}
const PALETTE=['#178269','#3976c5','#d39a28','#c95762','#8d70aa','#63a5a0','#aa8262'];
function color(s){return PALETTE[CHOICES.status.indexOf(s)%PALETTE.length]||PALETTE[0];}

function managementStatus(r){
  if(/放棄|失效|撤回|結案/.test(r.s))return '已放棄';
  if(/核駁|駁回|補正|異常/.test(r.s))return '需處理';
  if(/領證中/.test(r.s))return '進行中';
  if(/已取證|已取得|註冊案|專利通過|維持證書|已完成/.test(r.s))return '已完成';
  if(/審查|審核|申請|辦理|登記中|領證/.test(r.s))return '進行中';
  return '需處理';
}
function patentStatus(r){
  if(/放棄|失效|撤回|結案/.test(r.s))return '放棄';
  if(/專利通過|已取得/.test(r.s))return '專利通過';
  if(/領證/.test(r.s))return '領證中';
  if(/審核|審查|申請/.test(r.s))return '審核中';
  return null;
}
function chartStatus(r){return r.t==='patent'?patentStatus(r):r.s;}
function visibleChoices(k){
  if(TYPE_NAMES[pg]&&k==='country')return CHOICES.country.filter(v=>HUB.some(r=>r.t===pg&&r.c===v));
  return CHOICES[k];
}
function filterHTML(){
  const keys=TYPE_NAMES[pg]?['country','status']:['country','type','status'];
  return '<div class="panel hub-filters">'+keys.map(k=>{
    const vs=visibleChoices(k);
    return `<details class="hub-filter"><summary>${{country:'國家／地區',type:'資產類型',status:'狀態'}[k]} · ${vs.filter(v=>selection[k].has(v)).length}/${vs.length} ▾</summary><div class="hub-options"><button class="btn-outline" onclick="selectGroup('${k}',true)">全選</button><button class="btn-outline" onclick="selectGroup('${k}',false)">取消全選</button>${vs.map(v=>`<label><input type="checkbox" ${selection[k].has(v)?'checked':''} onchange="changeSelection('${k}',${CHOICES[k].indexOf(v)},this.checked)">${esc(k==='type'?TYPE_NAMES[v]:v)}</label>`).join('')}</div></details>`;
  }).join('')+'<button class="btn-outline" onclick="resetHub()">重設全部</button></div>';
}
function changeSelection(k,i,on){on?selection[k].add(CHOICES[k][i]):selection[k].delete(CHOICES[k][i]);cur=1;refreshHub();}
function selectGroup(k,on){visibleChoices(k).forEach(v=>on?selection[k].add(v):selection[k].delete(v));cur=1;refreshHub();}
function resetHub(){selection=Object.fromEntries(Object.entries(CHOICES).map(([k,v])=>[k,new Set(v)]));drill=null;render();}
function refreshHub(){const open=[...document.querySelectorAll('.hub-filter')].map(x=>x.open);render();document.querySelectorAll('.hub-filter').forEach((x,i)=>x.open=open[i]);}
function drillTo(t,s){drill={t,s};showPage(t);}

let customDimension='status';
chartModes.custom='donut';
function chart(t){
  const custom=t==='custom',title=custom?'自訂分析圖表':TYPE_NAMES[t];
  const rows=HUB.filter(r=>custom?selected(r):r.t===t);
  const label=r=>custom?(customDimension==='country'?r.c:customDimension==='type'?TYPE_NAMES[r.t]:managementStatus(r)):chartStatus(r);
  const counts=new Map();
  if(t==='patent')['審核中','領證中','放棄','專利通過'].forEach(v=>counts.set(v,0));
  rows.forEach(r=>{const key=label(r);if(key!==null)counts.set(key,(counts.get(key)||0)+1);});
  const entries=[...counts],total=entries.reduce((sum,[,n])=>sum+n,0);
  const action=i=>custom?`openCustomGroup(${i})`:`drillTo('${t}',${JSON.stringify(entries[i][0])})`;
  let offset=0;
  const paths=entries.map(([s,n],i)=>{
    if(!n)return '';
    const start=offset;offset+=n/total;
    const point=a=>[90+72*Math.cos(a*2*Math.PI-Math.PI/2),90+72*Math.sin(a*2*Math.PI-Math.PI/2)];
    const a=point(start),b=point(offset===1?offset-.000001:offset);
    return `<path d="M90 90 L${a} A72 72 0 ${n/total>.5?1:0} 1 ${b} Z" fill="${color(s)}" onclick="${esc(action(i))}"><title>${esc(s)}：${n} 件</title></path>`;
  }).join('');
  return `<article class="panel"><h2>${title}</h2><select aria-label="${title}圖表形式" class="chart-mode" onchange="chartModes['${t}']=this.value;render()"><option value="donut" ${chartModes[t]==='donut'?'selected':''}>甜甜圈圖</option><option value="bar" ${chartModes[t]==='bar'?'selected':''}>橫向長條圖</option></select><div class="muted">${total} 件 · ${custom?'依自訂條件':'全部案件'}${rows.length>total?'（另有 '+(rows.length-total)+' 件狀態未對應，請至管理頁檢查）':''}</div>${t==='patent'?'<small class="muted">放棄含失效、撤回；申請中併入審核中</small>':''}${!total?'<div class="empty">沒有符合條件的案件</div>':chartModes[t]==='donut'?`<svg class="donut" viewBox="0 0 180 180" role="img" aria-label="${title}分布">${paths}<circle cx="90" cy="90" r="51" fill="white"/><text x="90" y="98" text-anchor="middle">${total}</text></svg>`:''}${entries.map(([s,n],i)=>`<button class="legend-row" onclick="${esc(action(i))}"><i style="background:${color(s)}"></i><span>${esc(s)}</span><small>${n} 件 · ${total?(100*n/total).toFixed(1):'0.0'}%</small></button>${chartModes[t]==='bar'?`<div class="bar-track"><i style="background:${color(s)};width:${total?100*n/total:0}%"></i></div>`:''}`).join('')}</article>`;
}
function openCustomGroup(i){
  const rows=HUB.filter(selected),get=r=>customDimension==='country'?r.c:customDimension==='type'?TYPE_NAMES[r.t]:managementStatus(r);
  const key=[...new Set(rows.map(get))][i];if(key===undefined)return;
  document.getElementById('custom-results').innerHTML='<h2 class="section-title">'+esc(key)+' · 案件明細</h2>'+listTable(rows.filter(r=>get(r)===key).sort(defaultOrder));
  document.getElementById('custom-results').scrollIntoView({behavior:'smooth',block:'start'});
}
function upcomingRows(){
  const today=new Date(NOW_STR.slice(0,10).replaceAll('/','-')+'T00:00:00Z');
  const end=new Date(today),day=end.getUTCDate();end.setUTCDate(1);end.setUTCMonth(end.getUTCMonth()+6);
  const last=new Date(Date.UTC(end.getUTCFullYear(),end.getUTCMonth()+1,0)).getUTCDate();end.setUTCDate(Math.min(day,last));
  return HUB.filter(r=>dateValue(r)>=+today&&dateValue(r)<=+end&&managementStatus(r)!=='已放棄').sort((a,b)=>dateValue(a)-dateValue(b));
}
function listTable(rows){return `<div class="twrap hub-alerts"><table><thead><tr>${['案件名稱','資產類型','國家','狀態','到期日','提醒'].map(s=>'<th>'+s+'</th>').join('')}</tr></thead><tbody>${rows.map(r=>`<tr tabindex="0" onclick="openHub('${r.id}')" onkeydown="if(event.key==='Enter')openHub('${r.id}')"><td>${esc(r.name)}</td><td>${TYPE_NAMES[r.t]}</td><td>${esc(r.c)}</td><td>${esc(r.s)}</td><td>${esc(r.date||'—')}</td><td>${esc(attention(r)?.[1]||'正常')}</td></tr>`).join('')||'<tr><td colspan="6">沒有符合條件的案件</td></tr>'}</tbody></table></div>`;}
const COLS={trademark:[['商標名',r=>r.name],['國別',r=>r.c],['申請類別',r=>r.raw['申請類別']||r.raw['類別']],['狀態/進度說明',r=>r.s],['申請案號',r=>r.raw['申請案號']||r.raw['申請號']],['註冊號',r=>r.raw['註冊編號']||r.raw['註冊號']||r.raw['證書號 (進度)']],['使用期限（到期日）',r=>r.date]],patent:[['專利名稱（中文）',r=>r.name],['國別',r=>r.c],['申請案號',r=>r.raw['申請案號']],['專利編號',r=>r.raw['專利編號']],['目前狀態',r=>r.s],['證書到期日',r=>r.date]],registration:[['登記產品名',r=>r.name],['國別',r=>r.c],['登記類別',r=>r.raw['登記類別']],['登記公司',r=>r.raw['登記公司']],['狀態',r=>r.s],['有效期限',r=>r.date]]};
function management(t){let rows=HUB.filter(r=>r.t===t&&selected(r)&&(!drill||drill.t!==t||chartStatus(r)===drill.s));const q=flt.q.toLowerCase();if(q)rows=rows.filter(r=>[r.name,r.c,r.s,r.numbers].join(' ').toLowerCase().includes(q));rows.sort(defaultOrder);if(srt.col!==null){const get=COLS[t][Number(srt.col)]?.[1];if(get)rows.sort((a,b)=>String(get(a)||'').localeCompare(String(get(b)||''),'zh-Hant')*(srt.asc?1:-1));}const pages=Math.ceil(rows.length/pp)||1;cur=Math.min(cur,pages);const visible=rows.slice((cur-1)*pp,cur*pp);return filterHTML()+`<div class="fbar"><input aria-label="搜尋目前案件" value="${esc(flt.q)}" placeholder="搜尋名稱、國家、狀態、案號" oninput="hubLocalSearch(this)"><span class="rcount">共 ${rows.length} 筆</span>${drill?'<button class="btn-outline" onclick="drill=null;render()">清除圖表選取</button>':''}</div><div class="twrap"><table><thead><tr>${COLS[t].map(([name],i)=>`<th onclick="sortBy('${i}')">${name} ↕</th>`).join('')}</tr></thead><tbody>${visible.map(r=>`<tr tabindex="0" onclick="openHub('${r.id}')" onkeydown="if(event.key==='Enter')openHub('${r.id}')">${COLS[t].map(([,get],i)=>'<td>'+esc(get(r)||'—')+(i===0&&t==='patent'?'<div class="cs">'+esc(r.raw['專利類別']||'')+'</div>':'')+'</td>').join('')}</tr>`).join('')||'<tr><td colspan="7">沒有符合條件的案件</td></tr>'}</tbody></table></div>${mkPager(rows.length,pages)}`;}
function hubLocalSearch(input){const pos=input.selectionStart;flt.q=input.value;cur=1;render();const next=document.querySelector('[aria-label="搜尋目前案件"]');next.focus();next.setSelectionRange(pos,pos);}
const legacyRender=render;
render=function(){const content=document.getElementById('content');document.getElementById('topbar-actions').innerHTML='';const header=`<div class="hub-head"><h1>${PAGE_TITLES[pg]||'搜尋結果'}</h1>${pg==='overview'?`<small>資料更新：${NOW_STR}（台灣時間） · Google Sheets</small>`:''}</div>`;let body='';if(pg==='overview'){const alertRows=HUB.filter(r=>attention(r)).sort((a,b)=>attention(a)[0]-attention(b)[0]||(dateValue(a)-dateValue(b)||0));body='<h2 class="section-title">自訂分析圖表</h2>'+filterHTML()+`<div class="custom-axis"><label>分析維度 <select aria-label="分析維度" onchange="customDimension=this.value;render()">${['status','country','type'].map(k=>`<option value="${k}" ${customDimension===k?'selected':''}>${{status:'狀態',country:'國家／地區',type:'資產類型'}[k]}</option>`).join('')}</select></label><small class="muted">條件即時更新右方自訂圖，前三張呈現全部案件。</small></div>`+'<div class="hub-grid">'+[...Object.keys(TYPE_NAMES),'custom'].map(chart).join('')+'</div><div id="custom-results"></div><h2 class="section-title">需要注意的案件 · '+alertRows.length+' 件</h2>'+listTable(alertRows);}else if(TYPE_NAMES[pg]){body=management(pg);if(pg!=='patent')document.getElementById('topbar-actions').innerHTML=`<button class="btn-outline" onclick="openExpMo('${pg}')">匯出</button>`;}else if(pg==='alerts'){body='<p class="muted">今天起六個月內到期的有效案件 · '+upcomingRows().length+' 件（不含已放棄、失效及撤回）</p>'+listTable(upcomingRows());}else if(pg==='search'){const rows=searchRows(globalQuery);body='<p class="muted">搜尋全部資產，共 '+rows.length+' 件</p>'+listTable(rows);}else{legacyRender();return;}content.innerHTML=header+body;};
showPage=function(n){pg=n;cur=1;flt.q='';srt={col:null,asc:true};if(drill&&drill.t!==n)drill=null;document.querySelectorAll('.nav-item').forEach(x=>x.classList.toggle('active',(x.getAttribute('onclick')||'').includes("'"+n+"'")));render();};
function openHub(id){const r=byId.get(id);if(!r)return;lastFocus=document.activeElement;const field=(k,v)=>`<div class="di"><label>${esc(k)}</label><div class="dv">${esc(v||'—')}</div></div>`;const dates=Object.entries(r.raw).filter(([k])=>!k.startsWith('_')&&/日期|期限|使用期間/.test(k));document.getElementById('mt').textContent=r.name;document.getElementById('mb').innerHTML='<section class="drawer-section"><h3>案件摘要</h3><div class="dg">'+field('類型',TYPE_NAMES[r.t])+field('國家',r.c)+field('原始狀態',r.s)+field('案號',r.numbers.trim())+'</div></section><section class="drawer-section"><h3>重要日期</h3><div class="dg">'+field('到期日',r.date)+field('提醒',attention(r)?.[1]||'正常')+dates.map(([k,v])=>field(k,v)).join('')+'</div></section><details class="drawer-section"><summary>完整原始資料</summary><div class="dg">'+Object.entries(r.raw).filter(([k])=>!k.startsWith('_')).map(([k,v])=>field(k,v)).join('')+'</div></details>';document.getElementById('mo').classList.add('open');document.querySelector('#mo .mclose').focus();}
const oldClose=closeMo;closeMo=function(e){oldClose(e);if(!document.getElementById('mo').classList.contains('open'))lastFocus?.focus();};
function searchRows(q){q=q.trim().toLowerCase();return q?HUB.filter(r=>[r.name,r.c,r.s,r.numbers].join(' ').toLowerCase().includes(q)):[];}
function searchHub(input){globalQuery=input.value;const rows=searchRows(globalQuery);document.getElementById('hub-results').innerHTML=!globalQuery.trim()?'':rows.slice(0,8).map(r=>`<button onclick="document.getElementById('hub-results').innerHTML='';openHub('${r.id}')">${esc(r.name)}<small>${esc(TYPE_NAMES[r.t]+' · '+r.c+' · '+r.s+' · '+r.numbers.trim())}</small></button>`).join('')+(rows.length?`<button onclick="showPage('search');document.getElementById('hub-results').innerHTML=''">查看全部搜尋結果（${rows.length}）</button>`:'<button disabled>沒有符合的案件</button>');}
document.querySelector('.brand').innerHTML='<div class="brand-logo">CH BIOTECH</div><div class="brand-title">正瀚生技<br>CH IP Intelligence Hub<br>智財與登記管理</div>';
document.querySelector('.topbar').insertAdjacentHTML('afterbegin','<div class="hub-search"><input id="hub-search" aria-label="全域搜尋" placeholder="搜尋商標、專利、登記、國家或案號…" oninput="searchHub(this)"><div class="search-pop" id="hub-results"></div></div>');
document.addEventListener('click',e=>{if(!e.target.closest('.hub-search'))document.getElementById('hub-results').innerHTML='';if(!e.target.closest('.hub-filter'))document.querySelectorAll('.hub-filter[open]').forEach(menu=>menu.open=false);});
document.addEventListener('keydown',e=>{if(e.key==='Escape'){closeMo();document.getElementById('hub-results').innerHTML='';}if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='k'){e.preventDefault();document.getElementById('hub-search').focus();}});
// Export receives the same filters as the visible management page.
const legacyExportRows=exportRows;exportRows=function(){return legacyExportRows().filter(raw=>HUB.some(r=>r.raw===raw&&selected(r)&&(!drill||r.t!==drill.t||chartStatus(r)===drill.s)));};
const baseAttention=attention;attention=function(r){const a=baseAttention(r);if(a&&a[0]<5)return a;if(/審核|登記中/.test(r.s))return [5,'進度追蹤'];return a;};
const baseColor=color;color=function(s){if(/需處理|核駁/.test(s))return '#c95762';if(s==='已完成')return '#178269';if(s==='進行中'||s==='領證中')return '#3976c5';if(/已取證|已取得|專利通過|領證|維持證書/.test(s))return '#178269';if(/審查|審核|申請|辦理|登記中/.test(s))return '#3976c5';if(/放棄|失效|撤回|結案/.test(s))return '#88968f';return baseColor(s);};
let categoryFilter='all';
const baseSelected=selected;selected=function(r){const key=pg==='patent'?'專利類別':pg==='registration'?'登記類別':null;return baseSelected(r)&&(!key||categoryFilter==='all'||r.raw[key]===categoryFilter);};
const viewSelections=new Map();
const baseShowPage=showPage;showPage=function(n){viewSelections.set(pg,selection);selection=viewSelections.get(n)||Object.fromEntries(Object.entries(CHOICES).map(([k,v])=>[k,new Set(v)]));if(drill&&drill.t===n)selection=Object.fromEntries(Object.entries(CHOICES).map(([k,v])=>[k,new Set(v)]));categoryFilter='all';baseShowPage(n);};
const baseResetHub=resetHub;resetHub=function(){categoryFilter='all';flt.q='';cur=1;baseResetHub();};
const baseHubRender=render;render=function(){baseHubRender();document.querySelectorAll('.hub-grid .panel').forEach(panel=>{const buttons=[...panel.querySelectorAll('.legend-row')];if(!buttons.length)return;const box=document.createElement('div');box.style.cssText='max-height:290px;overflow:auto;padding-right:4px';panel.insertBefore(box,buttons[0]);buttons.forEach(b=>{const bar=b.nextElementSibling?.classList.contains('bar-track')?b.nextElementSibling:null;box.append(b);if(bar)box.append(bar);});});const key=pg==='patent'?'專利類別':pg==='registration'?'登記類別':null;if(key){const vs=[...new Set(HUB.filter(r=>r.t===pg).map(r=>r.raw[key]).filter(Boolean))];document.querySelector('.fbar')?.insertAdjacentHTML('beforeend',`<select aria-label="${key}" onchange="categoryFilter=this.value;cur=1;render()"><option value="all">所有${key}</option>${vs.map(v=>`<option value="${esc(v)}" ${categoryFilter===v?'selected':''}>${esc(v)}</option>`).join('')}</select>`);}};
document.getElementById('nba').textContent=upcomingRows().length;
render();
'''


# ─── Main ─────────────────────────────────────────────────────────────────
print('Downloading from Google Sheets...')
tm_data  = download_excel('商標', URLS['trademark'])
pt_data  = download_excel('專利', URLS['patent'])
rg_data  = download_excel('登記', URLS['registration'])

print('Processing...')
trademark    = process_trademark(read_trademark_rows(tm_data))
patent       = process_patent(read_excel_rows(pt_data))
registration = process_registration(read_excel_rows(rg_data))
print(f'  商標:{len(trademark)} 專利:{len(patent)} 登記:{len(registration)}')

print('Building HTML...')
html = build_html(trademark, patent, registration)

script_dir = os.path.dirname(os.path.abspath(__file__))
repo_dir = os.path.dirname(script_dir) if os.path.basename(script_dir) == 'scripts' else script_dir
out = os.path.join(repo_dir, 'index.html')
with open(out, 'w', encoding='utf-8') as f:
    f.write(html)
print(f'Done: {len(html):,} bytes → {out}')
