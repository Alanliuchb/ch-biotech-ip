#!/usr/bin/env python3
"""正瀚生技 智財與登記管理 v2 - 自動更新腳本"""
import json, datetime, urllib.request, ssl, io, sys, os

URLS = {
    'trademark':    'https://docs.google.com/spreadsheets/d/1MfsBuMHVZDFd_MV1v0LTnUOsYqUGkhVn/export?format=xlsx',
    'patent':       'https://docs.google.com/spreadsheets/d/1Uj_PV344NkDiY2n8YCs_HyjpnQYn87Ca/export?format=xlsx',
    'registration': 'https://docs.google.com/spreadsheets/d/1llnfbjcPST6Wa0p6psIxUfnjlGicZEi9/export?format=xlsx',
}
SHEET_NAMES = {
    'trademark':    '商標進度管理表',
    'patent':       '專利進度管理表',
    'registration': '產品登記管理表',
}
REG_HIDE = {'crops', 'n-p-k', 'organic matter', 'raw materials', 'others'}

TODAY = datetime.date.today()
TODAY_STR = TODAY.strftime('%Y/%m/%d')
NOW_STR = datetime.datetime.now().strftime('%Y/%m/%d %H:%M')


def download_excel(name, url):
    ctx = ssl.create_default_context()
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (GitHub-Actions)'})
    try:
        with urllib.request.urlopen(req, context=ctx, timeout=30) as r:
            data = r.read()
        print(f'  OK {name}: {len(data):,} bytes')
        return data
    except Exception as e:
        print(f'  ERROR {name}: {e}', file=sys.stderr)
        sys.exit(1)


def read_excel_rows(data):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    if len(all_rows) < 2:
        return []
    headers = [str(h).strip() if h is not None else '' for h in all_rows[1]]
    records = []
    for row in all_rows[2:]:
        if all(v is None for v in row):
            continue
        rec = {}
        for h, v in zip(headers, row):
            if h:
                rec[h] = str(v).strip() if v is not None else ''
        records.append(rec)
    return records


def calc_tm_deadline(date_str):
    v = (date_str or '').strip()
    if not v: return '', '待補期限'
    if v.upper() == 'N/A': return 'N/A', 'N/A'
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
    v = (date_str or '').strip()
    if not v: return '', '待補期限'
    if v.upper() == 'N/A': return 'N/A', 'N/A'
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
    st = r.get('進度狀態', '').strip()
    if st in ('註冊案', '申請案', '核駁案', '放棄案'):
        return st
    cert = (r.get('證書號 (進度)') or r.get('證書號(進度)') or r.get('註冊號') or '').strip()
    if not cert: return '申請案'
    if any(k in cert for k in ('放棄', '失效')): return '放棄案'
    if '核駁' in cert: return '核駁案'
    if cert.startswith('【') or '訴願' in cert or '申復' in cert: return '申請案'
    return '註冊案'


def patent_status(state):
    v = state.strip().replace('\n', '')
    if v.startswith('專利通過') or v.startswith('領證'): return '已取得'
    if any(k in v for k in ('放棄', '撤回', '失效', '結案')): return '已結案'
    return '申請中'


def process_trademark(records):
    for r in records:
        r['_status'] = trademark_status(r)
        end_raw = (r.get('使用期間-到期日') or r.get('使用期間_到期日') or '').strip()
        if not end_raw:
            period = (r.get('使用期間') or '').strip()
            if period and '-' in period:
                end_raw = period.split('-')[-1].strip()
        end_date, dl = calc_tm_deadline(end_raw)
        r['_end_date'] = end_date
        r['_deadline_status'] = dl
        start_raw = (r.get('使用期間-起始日') or r.get('使用期間_起始日') or '').strip()
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


def build_html(trademark, patent, registration):
    data = {'trademark': trademark, 'patent': patent, 'registration': registration}
    data_js = json.dumps(data, ensure_ascii=False).replace('</', '<\\/')
    tm_c = len(trademark)
    pt_c = len(patent)
    rg_c = len(registration)
    reg_hide_js = json.dumps(list(REG_HIDE))

    css = r'''
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Noto Sans TC",sans-serif;background:#f0f2f5;color:#1a1a2e;display:flex;height:100vh;overflow:hidden}
#sidebar{width:220px;min-width:220px;background:#0f1629;color:#c8d0e0;display:flex;flex-direction:column;height:100vh;overflow-y:auto;flex-shrink:0}
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
#main{flex:1;overflow-y:auto;display:flex;flex-direction:column}
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
.twrap{background:#fff;border-radius:10px;border:1px solid #e8edf5;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.06);margin-bottom:20px}
table{width:100%;border-collapse:collapse;font-size:13px}
thead th{background:#f7f9fd;padding:10px 13px;text-align:left;font-size:11px;font-weight:600;color:#6b7a99;letter-spacing:.5px;border-bottom:1px solid #e0e6ef;white-space:nowrap;cursor:pointer;user-select:none}
thead th:hover{color:#4c6ef5}
tbody tr{border-bottom:1px solid #f4f6fb;transition:background .1s;cursor:pointer}
tbody tr:hover{background:#f7f9fd}
tbody tr:last-child{border-bottom:none}
td{padding:9px 13px;vertical-align:middle}
.cn{font-weight:500;color:#1a1a2e;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.cs{font-size:11px;color:#8899bb;margin-top:2px}
.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11.5px;font-weight:500;white-space:nowrap}
.b-註冊案{background:#dcfce7;color:#166534}
.b-申請案{background:#dbeafe;color:#1d4ed8}
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
.plist{background:#fff;border-radius:10px;border:1px solid #e8edf5;overflow:hidden;margin-bottom:20px;box-shadow:0 1px 4px rgba(0,0,0,.06)}
.pr{display:grid;grid-template-columns:1fr 110px 130px 110px;border-bottom:1px solid #f0f4fa;font-size:13px;cursor:pointer;transition:background .1s}
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
<div class="di"><label>自動更新</label><div class="dv">✅ 上班時間每小時（週一至週五 8:00–17:00）</div></div>
</div></div>'''

    html = f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>正瀚生技｜智財與登記管理</title>
<style>{css}</style>
</head>
<body>
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
<div class="mo" id="mo" onclick="closeMo(event)">
  <div class="modal">
    <div class="mh"><h3 id="mt">案件明細</h3><button class="mclose" onclick="closeMo()">✕</button></div>
    <div class="mbody" id="mb"></div>
  </div>
</div>
<div class="mo" id="expMo" onclick="closeExpMo(event)">
  <div class="modal" style="max-width:740px">
    <div class="mh"><h3>匯出產品登記資料</h3><button class="mclose" onclick="closeExpMo()">✕</button></div>
    <div class="mbody">
      <div class="exp-section">
        <h4>模式 1：自選欄位匯出</h4>
        <div class="chk-grid" id="expColList"></div>
        <button class="btn-primary" style="margin-top:8px" onclick="doMode1Export()">↓ 匯出所選欄位 CSV</button>
      </div>
      <hr style="border:none;border-top:1px solid #e8edf5;margin:16px 0">
      <div class="exp-section">
        <h4>模式 2：各國登記類別彙總表（自行取得 vs 協助客戶）</h4>
        <button class="btn-outline" onclick="doMode2Preview()">預覽彙總表</button>
        <button class="btn-primary" style="margin-left:8px" onclick="doMode2Export()">↓ 匯出彙總表 CSV</button>
        <div id="mode2Preview" style="margin-top:14px;overflow-x:auto"></div>
      </div>
    </div>
  </div>
</div>
<script type="application/json" id="raw-data">{data_js}</script>
<script>
const RAW = JSON.parse(document.getElementById('raw-data').textContent);
const REG_HIDE = new Set({reg_hide_js});
const NOW_STR = '{NOW_STR}';
const TODAY_STR = '{TODAY_STR}';
const SHEET_NAMES = {{trademark:'{SHEET_NAMES['trademark']}',patent:'{SHEET_NAMES['patent']}',registration:'{SHEET_NAMES['registration']}'}};

function esc(s) {{ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }}
function dlCls(s) {{
  const m = {{'正常':'正常','期限已過':'期限已過','即將到期':'即將到期',
    '即將到期(30天)':'30','即將到期(90天)':'90','即將到期(180天)':'180','即將到期(365天)':'365',
    '待補期限':'待補期限','日期異常':'日期異常','N/A':'NA'}};
  return 'dl-'+(m[s]||'待補期限');
}}
function badge(t,c){{return`<span class="badge ${{c}}">${{esc(t)}}</span>`;}}
function dlBadge(s){{return badge(s,dlCls(s));}}

let pg='overview',cur=1,pp=30;
let flt={{q:'',country:'all',tmSt:'all',ptSt:'all',ptType:'all',rgSt:'all',rgType:'all'}};
let srt={{col:null,asc:true}};

function setF(k,v){{flt[k]=v;cur=1;render();}}
function resetF(){{
  if(pg==='trademark')Object.assign(flt,{{q:'',country:'all',tmSt:'all'}});
  else if(pg==='patent')Object.assign(flt,{{q:'',country:'all',ptSt:'all',ptType:'all'}});
  else if(pg==='registration')Object.assign(flt,{{q:'',country:'all',rgSt:'all',rgType:'all'}});
  cur=1;render();
}}
function sortBy(c){{srt.col===c?srt.asc=!srt.asc:(srt.col=c,srt.asc=true);render();}}
function goP(n){{cur=n;render();document.getElementById('main').scrollTo(0,0);}}
const PAGE_TITLES={{overview:'主管總覽',trademark:'商標管理',patent:'專利管理',registration:'產品登記',alerts:'期限提醒',sync:'資料與同步'}};
function showPage(n,el){{
  pg=n;cur=1;
  Object.assign(flt,{{q:'',country:'all',tmSt:'all',ptSt:'all',ptType:'all',rgSt:'all',rgType:'all'}});
  srt={{col:null,asc:true}};
  document.querySelectorAll('.nav-item').forEach(x=>x.classList.remove('active'));
  if(el)el.classList.add('active');
  document.getElementById('bc').innerHTML='管理中心 ／ <span>'+PAGE_TITLES[n]+'</span>';
  render();
}}
function render(){{
  const el=document.getElementById('content');
  const ac=document.getElementById('topbar-actions');
  ac.innerHTML='';
  if(pg==='overview')el.innerHTML=renderOv();
  else if(pg==='trademark')el.innerHTML=renderTrademark();
  else if(pg==='patent')el.innerHTML=renderPatent();
  else if(pg==='registration'){{el.innerHTML=renderRegistration();ac.innerHTML='<button class="btn-outline" onclick="openExpMo()">↓ 匯出</button>';}}
  else if(pg==='alerts')el.innerHTML=renderAlerts();
  else if(pg==='sync')el.innerHTML=`{sync_section}`;
}}
function mkPager(total,pages){{
  if(pages<=1)return'';
  let btns='',prev=-1;
  for(let i=1;i<=pages;i++){{
    if(i===1||i===pages||Math.abs(i-cur)<=2){{
      if(prev!==-1&&i-prev>1)btns+='<span class="pgi">…</span>';
      btns+=`<button class="pgb${{i===cur?' apg':''}}" onclick="goP(${{i}})">${{i}}</button>`;
      prev=i;
    }}
  }}
  return`<div class="pgbar"><button class="pgb" onclick="goP(${{cur-1}})" ${{cur===1?'disabled':''}}>‹</button>${{btns}}<button class="pgb" onclick="goP(${{cur+1}})" ${{cur===pages?'disabled':''}}>›</button><span class="pgi">${{(cur-1)*pp+1}}–${{Math.min(cur*pp,total)}}/${{total}}</span></div>`;
}}
function applySort(d){{
  if(!srt.col)return d;
  return[...d].sort((a,b)=>{{const va=a[srt.col]||'',vb=b[srt.col]||'';return srt.asc?(va<vb?-1:va>vb?1:0):(va>vb?-1:va<vb?1:0);}});
}}
function buildAll(){{
  const a=[];
  RAW.trademark.forEach(r=>a.push({{_type:'商標',_name:r['商標案件']||r['商標名稱']||'—',_country:r['國別']||'—',_status:r._status,_deadline:r._end_date||'',_dl:r._deadline_status,_sub:r['商標分類']||'',_raw:r}}));
  RAW.patent.forEach(r=>a.push({{_type:'專利',_name:r['專利名稱(中文)']||'—',_country:r['國別']||'—',_status:r._status,_deadline:r._end_date||'',_dl:r._deadline_status,_sub:r['專利類別']||'',_raw:r}}));
  RAW.registration.forEach(r=>a.push({{_type:'產品登記',_name:r['登記產品名']||'—',_country:r['國別']||'—',_status:r._status,_deadline:r._end_date||'',_dl:r._deadline_status,_sub:r['登記類別']||'',_raw:r}}));
  return a;
}}
const ALL=buildAll();
const ALERT_DL=new Set(['期限已過','即將到期','即將到期(30天)','即將到期(90天)','即將到期(180天)','即將到期(365天)']);
document.getElementById('nba').textContent=ALL.filter(r=>ALERT_DL.has(r._dl)).length;

function renderOv(){{
  const tm=RAW.trademark,pt=RAW.patent,rg=RAW.registration;
  const tmReg=tm.filter(r=>r._status==='註冊案').length;
  const ptGet=pt.filter(r=>r._status==='已取得').length;
  const rgGet=rg.filter(r=>r._status==='已取得').length;
  const tmOver=tm.filter(r=>r._deadline_status==='期限已過').length;
  const alertN=ALL.filter(r=>ALERT_DL.has(r._dl)).length;
  const pri=ALL.filter(r=>ALERT_DL.has(r._dl)).sort((a,b)=>{{
    const rk=['期限已過','即將到期','即將到期(30天)','即將到期(90天)','即將到期(180天)','即將到期(365天)'];
    return(rk.indexOf(a._dl)||9)-(rk.indexOf(b._dl)||9);
  }}).slice(0,8);
  const priRows=pri.length===0
    ?'<div class="plist"><div class="empty">目前無需立即關注的案件 ✓</div></div>'
    :'<div class="plist"><div class="pr prh"><div>案件名稱</div><div>類型・國別</div><div>期限日期</div><div>提醒</div></div>'
      +pri.map(r=>`<div class="pr" onclick='openMo(${{JSON.stringify(JSON.stringify(r))}})'><div><div class="cn">${{esc(r._name)}}</div><div class="cs">${{esc(r._sub)}}</div></div><div style="font-size:12px">${{esc(r._type+'·'+r._country)}}</div><div style="font-size:12px;color:#6b7a99">${{esc(r._deadline||'—')}}</div><div>${{dlBadge(r._dl)}}</div></div>`).join('')+'</div>';
  return`<div class="ov-grid">
    <div class="card"><div class="card-label">® 商標</div><div class="card-value">${{tm.length}}</div><div class="card-sub">註冊案 ${{tmReg}} ／ 其他 ${{tm.length-tmReg}}</div></div>
    <div class="card"><div class="card-label">◇ 專利</div><div class="card-value">${{pt.length}}</div><div class="card-sub">已取得 ${{ptGet}} ／ 申請中 ${{pt.length-ptGet}}</div></div>
    <div class="card"><div class="card-label">▤ 產品登記</div><div class="card-value">${{rg.length}}</div><div class="card-sub">已取得 ${{rgGet}} ／ 辦理中 ${{rg.length-rgGet}}</div></div>
    <div class="card${{tmOver>0?' ac':''}}"><div class="card-label">${{tmOver>0?'⚠ ':''}}商標期限已過</div><div class="card-value">${{tmOver}}</div><div class="card-sub">需確認是否延展</div></div>
    <div class="card${{alertN>0?' wc':''}}"><div class="card-label">${{alertN>0?'⏰ ':''}}近期到期提醒</div><div class="card-value">${{alertN}}</div><div class="card-sub">需關注的案件數</div></div>
  </div><div class="section-title">優先關注事項</div>${{priRows}}`;
}}

function renderTrademark(){{
  let d=[...RAW.trademark];
  if(flt.q){{const q=flt.q.toLowerCase();d=d.filter(r=>(r['商標案件']||r['商標名稱']||'').toLowerCase().includes(q)||(r['申請案號']||'').toLowerCase().includes(q)||(r['國別']||'').toLowerCase().includes(q));}}
  if(flt.country!=='all')d=d.filter(r=>r['國別']===flt.country);
  if(flt.tmSt!=='all')d=d.filter(r=>r._status===flt.tmSt);
  d=applySort(d);
  const total=d.length,pages=Math.ceil(total/pp)||1;
  if(cur>pages)cur=pages;
  const rows=d.slice((cur-1)*pp,cur*pp);
  const countries=[...new Set(RAW.trademark.map(r=>r['國別']).filter(Boolean))].sort();
  const syncBar=`<div class="sync-bar">⇄ 同步時間：${{NOW_STR}}　·　來源：${{SHEET_NAMES.trademark}}</div>`;
  const fbar=`<div class="fbar">
    <input type="text" placeholder="搜尋商標名稱、申請案號…" oninput="setF('q',this.value)" value="${{esc(flt.q)}}">
    <select onchange="setF('country',this.value)"><option value="all">所有國別</option>${{countries.map(c=>`<option value="${{esc(c)}}"${{flt.country===c?' selected':''}}>${{esc(c)}}</option>`).join('')}}</select>
    <select onchange="setF('tmSt',this.value)"><option value="all">所有進度</option>${{['註冊案','申請案','核駁案','放棄案'].map(s=>`<option value="${{s}}"${{flt.tmSt===s?' selected':''}}>${{s}}</option>`).join('')}}</select>
    <span class="frs" onclick="resetF()">重設</span>
    <span class="rcount">共 ${{total}} 筆${{total!==RAW.trademark.length?' (全 '+RAW.trademark.length+')':''}}</span>
  </div>`;
  const tbody=rows.length===0?`<tr><td colspan="7"><div class="empty">無符合條件的案件</div></td></tr>`:rows.map(r=>{{
    const name=r['商標案件']||r['商標名稱']||'—';
    const imgUrl=r['商標圖示']||r['圖片URL']||r['圖片']||'';
    const imgTag=imgUrl?`<img src="${{esc(imgUrl)}}" style="width:28px;height:28px;object-fit:contain;vertical-align:middle;margin-right:6px;border-radius:4px">`:'';
    const appNo=r['申請案號']||'—';
    const regNo=r['證書號 (進度)']||r['證書號(進度)']||r['註冊號']||'—';
    return`<tr onclick='openMo(${{JSON.stringify(JSON.stringify(r))}})'><td><div style="display:flex;align-items:center">${{imgTag}}<div><div class="cn">${{esc(name)}}</div><div class="cs">${{esc(r['申請類別']||'')}}</div></div></div></td><td style="font-size:12px">${{esc(r['國別']||'—')}}</td><td style="font-size:12px">${{esc(r['商標分類']||r['類別']||'')}}</td><td>${{badge(r._status,'b-'+r._status)}}</td><td style="font-size:12px;color:#4a5568">${{esc(appNo)}}</td><td style="font-size:12px;color:#4a5568">${{esc(regNo)}}</td><td>${{r._end_date?`<div style="font-size:11px;color:#6b7a99;margin-bottom:2px">${{esc(r._end_date)}}</div>`:''}}${{dlBadge(r._deadline_status)}}</td></tr>`;
  }}).join('');
  return syncBar+fbar+`<div class="twrap"><table><thead><tr><th onclick="sortBy('_name')">商標案件</th><th onclick="sortBy('_country')">國別</th><th>類別</th><th onclick="sortBy('_status')">進度狀態</th><th>申請案號</th><th>證書號／進度</th><th onclick="sortBy('_end_date')">使用期限</th></tr></thead><tbody>${{tbody}}</tbody></table>${{mkPager(total,pages)}}</div>`;
}}

function renderPatent(){{
  let d=[...RAW.patent];
  if(flt.q){{const q=flt.q.toLowerCase();d=d.filter(r=>(r['專利名稱(中文)']||'').toLowerCase().includes(q)||(r['申請案號']||'').toLowerCase().includes(q)||(r['專利編號']||'').toLowerCase().includes(q)||(r['國別']||'').toLowerCase().includes(q));}}
  if(flt.country!=='all')d=d.filter(r=>r['國別']===flt.country);
  if(flt.ptSt!=='all')d=d.filter(r=>r._status===flt.ptSt);
  if(flt.ptType!=='all')d=d.filter(r=>(r['專利類別']||'')===flt.ptType);
  d=applySort(d);
  const total=d.length,pages=Math.ceil(total/pp)||1;
  if(cur>pages)cur=pages;
  const rows=d.slice((cur-1)*pp,cur*pp);
  const countries=[...new Set(RAW.patent.map(r=>r['國別']).filter(Boolean))].sort();
  const ptTypes=[...new Set(RAW.patent.map(r=>r['專利類別']).filter(Boolean))].sort();
  const fbar=`<div class="fbar">
    <input type="text" placeholder="搜尋專利名稱、申請案號、專利編號…" oninput="setF('q',this.value)" value="${{esc(flt.q)}}">
    <select onchange="setF('country',this.value)"><option value="all">所有國別</option>${{countries.map(c=>`<option value="${{esc(c)}}"${{flt.country===c?' selected':''}}>${{esc(c)}}</option>`).join('')}}</select>
    ${{ptTypes.length?`<select onchange="setF('ptType',this.value)"><option value="all">所有類別</option>${{ptTypes.map(t=>`<option value="${{esc(t)}}"${{flt.ptType===t?' selected':''}}>${{esc(t)}}</option>`).join('')}}</select>`:''}}
    <select onchange="setF('ptSt',this.value)"><option value="all">所有狀態</option>${{['已取得','申請中','已結案'].map(s=>`<option value="${{s}}"${{flt.ptSt===s?' selected':''}}>${{s}}</option>`).join('')}}</select>
    <span class="frs" onclick="resetF()">重設</span>
    <span class="rcount">共 ${{total}} 筆${{total!==RAW.patent.length?' (全 '+RAW.patent.length+')':''}}</span>
  </div>`;
  const tbody=rows.length===0?`<tr><td colspan="7"><div class="empty">無符合條件的案件</div></td></tr>`:rows.map(r=>`<tr onclick='openMo(${{JSON.stringify(JSON.stringify(r))}})'><td><div class="cn">${{esc(r['專利名稱(中文)']||'—')}}</div><div class="cs">${{esc(r['專利類別']||'')}}</div></td><td style="font-size:12px">${{esc(r['國別']||'—')}}</td><td style="font-size:12px">${{esc(r['申請案號']||'—')}}</td><td style="font-size:12px">${{esc(r['專利編號']||'—')}}</td><td>${{badge(r._status,'b-'+r._status)}}</td><td style="font-size:12px;color:#6b7a99">${{esc(r._end_date||'—')}}</td><td>${{dlBadge(r._deadline_status)}}</td></tr>`).join('');
  return fbar+`<div class="twrap"><table><thead><tr><th onclick="sortBy('_name')">專利名稱（中文）</th><th onclick="sortBy('_country')">國別</th><th>申請案號</th><th>專利編號</th><th onclick="sortBy('_status')">目前狀態</th><th onclick="sortBy('_end_date')">證書到期日</th><th onclick="sortBy('_dl')">期限提示</th></tr></thead><tbody>${{tbody}}</tbody></table>${{mkPager(total,pages)}}</div>`;
}}

function renderRegistration(){{
  let d=[...RAW.registration];
  if(flt.q){{const q=flt.q.toLowerCase();d=d.filter(r=>(r['登記產品名']||'').toLowerCase().includes(q)||(r['國別']||'').toLowerCase().includes(q)||(r['證書/License ID']||'').toLowerCase().includes(q)||(r['登記公司']||'').toLowerCase().includes(q));}}
  if(flt.country!=='all')d=d.filter(r=>r['國別']===flt.country);
  if(flt.rgSt!=='all')d=d.filter(r=>r._status===flt.rgSt);
  if(flt.rgType!=='all')d=d.filter(r=>(r['登記類別']||'')===flt.rgType);
  d=applySort(d);
  const total=d.length,pages=Math.ceil(total/pp)||1;
  if(cur>pages)cur=pages;
  const rows=d.slice((cur-1)*pp,cur*pp);
  const countries=[...new Set(RAW.registration.map(r=>r['國別']).filter(Boolean))].sort();
  const rgTypes=[...new Set(RAW.registration.map(r=>r['登記類別']).filter(Boolean))].sort();
  const statuses=[...new Set(RAW.registration.map(r=>r._status).filter(Boolean))].sort();
  const fbar=`<div class="fbar">
    <input type="text" placeholder="搜尋產品名稱、國別、證書號…" oninput="setF('q',this.value)" value="${{esc(flt.q)}}">
    <select onchange="setF('country',this.value)"><option value="all">所有國別</option>${{countries.map(c=>`<option value="${{esc(c)}}"${{flt.country===c?' selected':''}}>${{esc(c)}}</option>`).join('')}}</select>
    ${{rgTypes.length?`<select onchange="setF('rgType',this.value)"><option value="all">所有登記類別</option>${{rgTypes.map(t=>`<option value="${{esc(t)}}"${{flt.rgType===t?' selected':''}}>${{esc(t)}}</option>`).join('')}}</select>`:''}}
    <select onchange="setF('rgSt',this.value)"><option value="all">所有狀態</option>${{statuses.map(s=>`<option value="${{esc(s)}}"${{flt.rgSt===s?' selected':''}}>${{s}}</option>`).join('')}}</select>
    <span class="frs" onclick="resetF()">重設</span>
    <span class="rcount">共 ${{total}} 筆${{total!==RAW.registration.length?' (全 '+RAW.registration.length+')':''}}</span>
  </div>`;
  const tbody=rows.length===0?`<tr><td colspan="7"><div class="empty">無符合條件的案件</div></td></tr>`:rows.map(r=>`<tr onclick='openMo(${{JSON.stringify(JSON.stringify(r))}})'><td><div class="cn">${{esc(r['登記產品名']||'—')}}</div></td><td style="font-size:12px">${{esc(r['登記類別']||'—')}}</td><td style="font-size:12px">${{esc(r['國別']||'—')}}</td><td style="font-size:12px;max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${{esc(r['登記公司']||'—')}}</td><td>${{badge(r._status,'b-'+r._status)}}</td><td style="font-size:12px;color:#6b7a99">${{esc(r._end_date||'—')}}</td><td>${{dlBadge(r._deadline_status)}}</td></tr>`).join('');
  return fbar+`<div class="twrap"><table><thead><tr><th onclick="sortBy('_name')">登記產品名</th><th onclick="sortBy('_sub')">登記類別</th><th onclick="sortBy('_country')">國別</th><th>登記公司</th><th onclick="sortBy('_status')">狀態</th><th onclick="sortBy('_end_date')">有效期限</th><th onclick="sortBy('_dl')">期限提示</th></tr></thead><tbody>${{tbody}}</tbody></table>${{mkPager(total,pages)}}</div>`;
}}

function renderAlerts(){{
  const RANKS=['期限已過','即將到期','即將到期(30天)','即將到期(90天)','即將到期(180天)','即將到期(365天)'];
  const LABELS={{'期限已過':'⚠ 期限已過','即將到期':'🟠 6個月內到期','即將到期(30天)':'🔴 30天內到期','即將到期(90天)':'🟠 90天內到期','即將到期(180天)':'🟡 180天內到期','即將到期(365天)':'🟡 365天內到期'}};
  let html=RANKS.map(dl=>{{
    const items=ALL.filter(r=>r._dl===dl);
    if(!items.length)return'';
    return`<div style="margin-bottom:20px"><div class="section-title">${{LABELS[dl]}} (${{items.length}})</div><div class="plist"><div class="pr prh"><div>案件名稱</div><div>類型・國別</div><div>期限日期</div><div>提醒</div></div>${{items.map(r=>`<div class="pr" onclick='openMo(${{JSON.stringify(JSON.stringify(r))}})'><div><div class="cn">${{esc(r._name)}}</div></div><div style="font-size:12px">${{esc(r._type+'·'+r._country)}}</div><div style="font-size:12px;color:#6b7a99">${{esc(r._deadline||'—')}}</div><div>${{dlBadge(r._dl)}}</div></div>`).join('')}}</div></div>`;
  }}).join('');
  const na=ALL.filter(r=>r._dl==='N/A');
  if(na.length)html+=`<div style="margin-bottom:20px"><div class="section-title">無到期日 (N/A) (${{na.length}})</div><div class="ibox">共 ${{na.length}} 件案件登記無期限（N/A），無需展延。</div></div>`;
  const missing=ALL.filter(r=>r._dl==='待補期限');
  if(missing.length)html+=`<div style="margin-bottom:20px"><div class="section-title">待補期限 (${{missing.length}})</div><div class="ibox">共 ${{missing.length}} 件案件尚無到期日記錄。</div></div>`;
  return html||'<div class="empty" style="padding:60px">目前無需關注的到期案件 ✓</div>';
}}

function openMo(s){{
  let r;try{{r=JSON.parse(s);}}catch{{return;}}
  const raw=r._raw||r;
  const type=r._type||'';
  document.getElementById('mt').textContent=type+' 案件明細';
  const SKIP=new Set(['_status','_end_date','_deadline_status','_start_date','_type','_name','_country','_deadline','_dl','_sub','_raw']);
  const FULL=new Set(['專利名稱(中文)','專利名稱(英文)','備註','說明','目前狀態']);
  const hideSet=type==='產品登記'?REG_HIDE:new Set();
  let topHtml=`<div class="dg" style="margin-bottom:14px"><div class="di"><label>類型</label><div class="dv">${{badge(type,'b-T')}}</div></div><div class="di"><label>狀態</label><div class="dv">${{badge(r._status||'—','b-'+(r._status||''))}}</div></div>`;
  if(type==='商標'){{
    topHtml+=`<div class="di"><label>使用起始日</label><div class="dv">${{esc(raw._start_date||'—')}}</div></div><div class="di"><label>使用到期日</label><div class="dv">${{esc(raw._end_date||'—')}} ${{dlBadge(raw._deadline_status)}}</div></div>`;
  }}else{{
    topHtml+=`<div class="di"><label>期限日期</label><div class="dv">${{esc(r._deadline||'—')}}</div></div><div class="di"><label>期限提示</label><div class="dv">${{dlBadge(r._dl)}}</div></div>`;
  }}
  topHtml+='</div><hr style="border:none;border-top:1px solid #f0f4fa;margin:4px 0 12px"><div class="dg">';
  const items=Object.entries(raw).filter(([k])=>!SKIP.has(k)&&!hideSet.has(k.toLowerCase())).map(([k,v])=>`<div class="di${{FULL.has(k)?' full':''}}"><label>${{esc(k)}}</label><div class="dv">${{esc(v||'—')}}</div></div>`).join('');
  document.getElementById('mb').innerHTML=topHtml+items+'</div>';
  document.getElementById('mo').classList.add('open');
}}
function closeMo(e){{if(!e||e.target===document.getElementById('mo'))document.getElementById('mo').classList.remove('open');}}

function openExpMo(){{
  const allKeys=new Set();
  RAW.registration.forEach(r=>Object.keys(r).forEach(k=>{{if(!k.startsWith('_')&&!REG_HIDE.has(k.toLowerCase()))allKeys.add(k);}}));
  const suggested=['登記產品名','登記類別','國別','登記公司','證書/License ID','取得日期','進度','證書有效期限'];
  const ordered=[...suggested.filter(k=>allKeys.has(k)),...[...allKeys].filter(k=>!suggested.includes(k))];
  const defaultOn=new Set(['登記產品名','登記類別','國別','登記公司','進度','證書有效期限']);
  document.getElementById('expColList').innerHTML=ordered.map(k=>`<label class="chk-item"><input type="checkbox" id="exp_${{k}}" ${{defaultOn.has(k)?'checked':''}}> ${{esc(k)}}</label>`).join('');
  document.getElementById('mode2Preview').innerHTML='';
  document.getElementById('expMo').classList.add('open');
}}
function closeExpMo(e){{if(!e||e.target===document.getElementById('expMo'))document.getElementById('expMo').classList.remove('open');}}
function doMode1Export(){{
  const cols=[...document.querySelectorAll('#expColList input:checked')].map(el=>el.id.replace('exp_',''));
  if(!cols.length){{alert('請至少勾選一個欄位');return;}}
  const csv='﻿'+[cols.join(','),...RAW.registration.map(r=>cols.map(c=>'"'+(r[c]||'').replace(/"/g,'""')+'"').join(','))].join('\\n');
  dlCSV(csv,'正瀚_產品登記_{TODAY_STR}'.replace(/\//g,'')+'.csv');
}}
function doMode2Preview(){{
  document.getElementById('mode2Preview').innerHTML=renderMode2HTML(buildMode2());
}}
function doMode2Export(){{
  const{{countries,types,data}}=buildMode2();
  const hdr=['國別/登記類別',...types,'合計'];
  const rows=countries.map(c=>{{
    const row=[c];let tot=0;
    types.forEach(t=>{{const cell=(data[c]&&data[c][t])||{{self:0,help:0}};row.push(`自行${{cell.self}}/協助${{cell.help}}`);tot+=cell.self+cell.help;}});
    row.push(tot);return row;
  }});
  const csv='﻿'+[hdr.join(','),...rows.map(r=>r.map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(','))].join('\\n');
  dlCSV(csv,'正瀚_產品登記彙總_{TODAY_STR}'.replace(/\//g,'')+'.csv');
}}
function buildMode2(){{
  const CH='CH Biotech R&D Co., Ltd';
  const countries=[...new Set(RAW.registration.map(r=>r['國別']).filter(Boolean))].sort();
  const types=[...new Set(RAW.registration.map(r=>r['登記類別']).filter(Boolean))].sort();
  const data={{}};
  RAW.registration.forEach(r=>{{
    const c=r['國別']||'',t=r['登記類別']||'',co=r['登記公司']||'';
    if(!c||!t)return;
    if(!data[c])data[c]={{}};
    if(!data[c][t])data[c][t]={{self:0,help:0}};
    if(co===CH)data[c][t].self++;else data[c][t].help++;
  }});
  return{{countries,types,data}};
}}
function renderMode2HTML({{countries,types,data}}){{
  const hdr=`<tr><th class="rc">國別 ＼ 登記類別</th>${{types.map(t=>`<th>${{esc(t)}}</th>`).join('')}}<th>合計</th></tr>`;
  const rows=countries.map(c=>{{
    let tot=0;
    const cells=types.map(t=>{{const cell=(data[c]&&data[c][t])||{{self:0,help:0}};tot+=cell.self+cell.help;if(!cell.self&&!cell.help)return'<td>—</td>';return`<td>自行 ${{cell.self}}<br>協助 ${{cell.help}}</td>`;}}).join('');
    return`<tr><td class="rc">${{esc(c)}}</td>${{cells}}<td><strong>${{tot}}</strong></td></tr>`;
  }}).join('');
  return`<table class="mode2-tbl"><thead>${{hdr}}</thead><tbody>${{rows}}</tbody></table><div style="font-size:11px;color:#8899bb;margin-top:6px">自行 = CH Biotech R&amp;D Co., Ltd；協助 = 非 CH Biotech（協助客戶取得）</div>`;
}}
function dlCSV(csv,name){{const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{{type:'text/csv;charset=utf-8'}}));a.download=name;a.click();}}

render();
</script>
</body>
</html>'''
    return html


print('Downloading from Google Sheets...')
tm_data  = download_excel('商標', URLS['trademark'])
pt_data  = download_excel('專利', URLS['patent'])
rg_data  = download_excel('登記', URLS['registration'])

print('Processing...')
trademark    = process_trademark(read_excel_rows(tm_data))
patent       = process_patent(read_excel_rows(pt_data))
registration = process_registration(read_excel_rows(rg_data))
print(f'  商標:{len(trademark)} 專利:{len(patent)} 登記:{len(registration)}')

print('Building HTML...')
html = build_html(trademark, patent, registration)

out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'index.html')
with open(out, 'w', encoding='utf-8') as f:
    f.write(html)
print(f'Done: {len(html):,} bytes → {out}')
