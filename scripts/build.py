#!/usr/bin/env python3
import json, datetime, urllib.request, ssl, io, sys, os

URLS = {
    'trademark':    'https://docs.google.com/spreadsheets/d/1MfsBuMHVZDFd_MV1v0LTnUOsYqUGkhVn/export?format=xlsx',
    'patent':       'https://docs.google.com/spreadsheets/d/1Uj_PV344NkDiY2n8YCs_HyjpnQYn87Ca/export?format=xlsx',
    'registration': 'https://docs.google.com/spreadsheets/d/1llnfbjcPST6Wa0p6psIxUfnjlGicZEi9/export?format=xlsx',
}

TODAY = datetime.date.today()
TODAY_STR = TODAY.strftime('%Y/%m/%d')

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

def calc_deadline(date_str):
    if not date_str:
        return '', '待補期限'
    s = date_str.strip().replace('/', '-').split(' ')[0]
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
        return date_str, '日期異常'

def trademark_status(cert):
    v = cert.strip()
    if not v: return '申請中'
    if any(k in v for k in ['放棄', '失效', '結案']): return '已結案'
    if v.startswith('【') or '訴願' in v: return '申請中'
    return '已取得'

def patent_status(state):
    v = state.strip().replace('\n', '')
    if v.startswith('專利通過') or v.startswith('領證'): return '已取得'
    if any(k in v for k in ['放棄', '撤回', '失效', '結案']): return '已結案'
    return '申請中'

def process_trademark(records):
    for r in records:
        cert = r.get('證書號 (進度)', '')
        r['_status'] = trademark_status(cert)
        period = r.get('使用期間', '')
        end_raw = ''
        if period and '-' in period:
            end_raw = period.split('-')[-1].strip()
        end_date, dl = calc_deadline(end_raw)
        r['_end_date'] = end_date
        r['_deadline_status'] = dl
    return records

def process_patent(records):
    for r in records:
        r['_status'] = patent_status(r.get('目前狀態', ''))
        r['_end_date'] = ''
        r['_deadline_status'] = '待補期限'
    return records

def process_registration(records):
    sm = {'維持證書': '已取得', '已取得': '已取得',
          '登記中': '辦理中', '辦理中': '辦理中', '申請中': '辦理中'}
    for r in records:
        raw = r.get('進度', '')
        r['_status'] = sm.get(raw, raw or '辦理中')
        end_date, dl = calc_deadline(r.get('證書有效期限', ''))
        r['_end_date'] = end_date
        r['_deadline_status'] = dl
    return records

def build_html(trademark, patent, registration):
    data = {'trademark': trademark, 'patent': patent, 'registration': registration}
    data_js = json.dumps(data, ensure_ascii=False).replace('</', '<\\/')
    tm_count = len(trademark)
    pt_count = len(patent)
    rg_count = len(registration)

    css = '''*{box-sizing:border-box;margin:0;padding:0}
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
.content{padding:28px;flex:1}
.overview-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px;margin-bottom:28px}
.card{background:#fff;border-radius:10px;padding:20px;box-shadow:0 1px 4px rgba(0,0,0,.06);border:1px solid #e8edf5}
.card-label{font-size:11.5px;color:#8899bb;font-weight:600;letter-spacing:.5px;margin-bottom:6px}
.card-value{font-size:28px;font-weight:700;color:#1a1a2e;line-height:1}
.card-sub{font-size:12px;color:#8899bb;margin-top:6px}
.card.ac{border-color:#fca5a5;background:#fff8f8}.card.ac .card-label{color:#dc2626}.card.ac .card-value{color:#dc2626}
.card.wc{border-color:#fcd34d;background:#fffdf0}.card.wc .card-label{color:#b45309}.card.wc .card-value{color:#b45309}
.section-title{font-size:15px;font-weight:600;color:#1a1a2e;margin-bottom:14px;display:flex;align-items:center;gap:8px}
.section-title::after{content:"";flex:1;height:1px;background:#e0e6ef}
.plist{background:#fff;border-radius:10px;border:1px solid #e8edf5;overflow:hidden;margin-bottom:28px;box-shadow:0 1px 4px rgba(0,0,0,.06)}
.pr{display:grid;grid-template-columns:1fr 130px 140px 120px;border-bottom:1px solid #f0f4fa;font-size:13px;cursor:pointer;transition:background .1s}
.pr:hover:not(.prh){background:#f7f9fd}.pr:last-child{border-bottom:none}
.prh{background:#f7f9fd;font-size:11px;font-weight:600;color:#6b7a99;cursor:default}
.pr>div{padding:11px 16px}
.fbar{background:#fff;border-radius:10px;border:1px solid #e8edf5;padding:14px 18px;margin-bottom:14px;display:flex;flex-wrap:wrap;gap:8px;align-items:center;box-shadow:0 1px 3px rgba(0,0,0,.04)}
.fbar input,.fbar select{padding:7px 11px;border:1px solid #d0d7e3;border-radius:6px;font-size:13px;color:#1a1a2e;background:#fff;outline:none}
.fbar input:focus,.fbar select:focus{border-color:#4c6ef5}
.fbar input{width:200px}
.frs{font-size:12px;color:#4c6ef5;cursor:pointer;padding:4px 8px;border-radius:4px}
.frs:hover{background:#eff3ff}
.rcount{font-size:12px;color:#8899bb;margin-left:auto}
.twrap{background:#fff;border-radius:10px;border:1px solid #e8edf5;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.06)}
table{width:100%;border-collapse:collapse;font-size:13px}
thead th{background:#f7f9fd;padding:11px 14px;text-align:left;font-size:11px;font-weight:600;color:#6b7a99;letter-spacing:.5px;border-bottom:1px solid #e0e6ef;white-space:nowrap;cursor:pointer;user-select:none}
thead th:hover{color:#4c6ef5}
tbody tr{border-bottom:1px solid #f4f6fb;transition:background .1s;cursor:pointer}
tbody tr:hover{background:#f7f9fd}
tbody tr:last-child{border-bottom:none}
td{padding:10px 14px;vertical-align:middle}
.cn{font-weight:500;color:#1a1a2e;max-width:230px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.cs{font-size:11px;color:#8899bb;margin-top:2px}
.badge{display:inline-block;padding:2px 9px;border-radius:10px;font-size:11.5px;font-weight:500;white-space:nowrap}
.bT{background:#f0f4ff;color:#3b5bdb}
.b-已取得{background:#dcfce7;color:#166534}
.b-辦理中,.b-申請中{background:#dbeafe;color:#1d4ed8}
.b-已結案{background:#f3f4f6;color:#6b7280}
.dl-正常{background:#f0fdf4;color:#16a34a}
.dl-期限已過{background:#fef2f2;color:#dc2626;font-weight:600}
.dl-30{background:#fef2f2;color:#dc2626}
.dl-90{background:#fff7ed;color:#c2410c}
.dl-180{background:#fff7ed;color:#d97706}
.dl-365{background:#fefce8;color:#a16207}
.dl-待補期限{background:#f3f4f6;color:#9ca3af}
.dl-日期異常{background:#fdf4ff;color:#7c3aed}
.pgbar{display:flex;align-items:center;justify-content:flex-end;gap:6px;padding:14px 18px;border-top:1px solid #f0f4fa}
.pgb{width:32px;height:32px;border-radius:6px;border:1px solid #d0d7e3;background:#fff;cursor:pointer;font-size:13px;display:flex;align-items:center;justify-content:center;transition:all .15s}
.pgb:hover:not([disabled]){background:#eff3ff;border-color:#4c6ef5;color:#4c6ef5}
.pgb.apg{background:#4c6ef5;color:#fff;border-color:#4c6ef5}
.pgb[disabled]{opacity:.4;cursor:not-allowed}
.pgi{font-size:12px;color:#8899bb;margin:0 6px}
.mo{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:100;display:flex;align-items:center;justify-content:center;opacity:0;pointer-events:none;transition:opacity .2s}
.mo.open{opacity:1;pointer-events:auto}
.modal{background:#fff;border-radius:12px;width:680px;max-width:95vw;max-height:85vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.2)}
.mh{padding:18px 24px;border-bottom:1px solid #e8edf5;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;background:#fff;z-index:1}
.mh h3{font-size:16px;font-weight:600}
.mclose{width:30px;height:30px;border:none;background:none;cursor:pointer;font-size:18px;color:#6b7a99;border-radius:6px}
.mclose:hover{background:#f3f4f6}
.mbody{padding:20px 24px}
.dg{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.di label{font-size:11px;color:#8899bb;font-weight:600;display:block;margin-bottom:3px}
.di .dv{font-size:13.5px;color:#1a1a2e}
.di.full{grid-column:1/-1}
.sc{background:#fff;border-radius:10px;border:1px solid #e8edf5;padding:22px;margin-bottom:16px;box-shadow:0 1px 4px rgba(0,0,0,.06)}
.sc h3{font-size:15px;font-weight:600;margin-bottom:12px}
.ibox{background:#eff3ff;border:1px solid #c5d2f6;border-radius:8px;padding:14px 16px;font-size:13px;color:#3b5bdb;line-height:1.5}
.empty{text-align:center;padding:40px 20px;color:#9ca3af;font-size:14px}'''

    sync_body = f'''<div class="sc"><h3>⇄ 資料同步狀態</h3><div class="dg" style="margin-top:12px">
      <div class="di"><label>資料基準日</label><div class="dv">{TODAY_STR}</div></div>
      <div class="di"><label>來源</label><div class="dv">Google Sheets 自動同步</div></div>
      <div class="di"><label>商標案件</label><div class="dv">{tm_count} 筆</div></div>
      <div class="di"><label>專利案件</label><div class="dv">{pt_count} 筆</div></div>
      <div class="di"><label>產品登記</label><div class="dv">{rg_count} 筆</div></div>
      <div class="di"><label>自動更新</label><div class="dv">✅ 上班時間每小時更新</div></div>
    </div></div>'''

    sidebar_nav = f'''® 商標管理</span><span class="nav-badge">{tm_count}</span></div>
    <div class="nav-item" onclick="showPage('patent',this)"><span style="flex:1">◇ 專利管理</span><span class="nav-badge">{pt_count}</span></div>
    <div class="nav-item" onclick="showPage('registration',this)"><span style="flex:1">▤ 產品登記</span><span class="nav-badge">{rg_count}'''

    html = f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>正瀚生技｜智財與登記管理</title>
<style>
{css}
</style>
</head>
<body>
<aside id="sidebar">
  <div class="brand"><div class="brand-logo">CH</div><div class="brand-sub">BIOTECH</div><div class="brand-title">正瀚生技<br>智財與登記管理</div></div>
  <nav class="nav">
    <div class="nav-section">管理工作台</div>
    <div class="nav-item active" onclick="showPage('overview',this)"><span style="flex:1">▦ 主管總覽</span></div>
    <div class="nav-item" onclick="showPage('trademark',this)"><span style="flex:1">{sidebar_nav}</span></div>
    <div class="nav-item" onclick="showPage('alerts',this)"><span style="flex:1">◷ 期限提醒</span><span class="nav-badge alert" id="nba">—</span></div>
    <div class="nav-section" style="margin-top:12px">設定</div>
    <div class="nav-item" onclick="showPage('sync',this)"><span style="flex:1">⇄ 資料與同步</span></div>
  </nav>
  <div class="sidebar-footer">資料基準：{TODAY_STR}<br>© 正瀚生技 CH BIOTECH</div>
</aside>
<main id="main">
  <div class="topbar">
    <div class="topbar-left" id="bc">管理中心 ／ <span>主管總覽</span></div>
    <div><button class="btn-outline" onclick="exportCSV()">↓ 匯出 CSV</button></div>
  </div>
  <div class="content" id="content"></div>
</main>
<div class="mo" id="mo" onclick="closeMo(event)">
  <div class="modal"><div class="mh"><h3 id="mt">案件明細</h3><button class="mclose" onclick="closeMo()">✕</button></div><div class="mbody" id="mb"></div></div>
</div>
<script type="application/json" id="raw-data">{data_js}</script>
<script>
const RAW=JSON.parse(document.getElementById('raw-data').textContent);
const SYNC_HTML=`{sync_body}`;
function esc(s){{return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}}
function dlCls(s){{const m={{'正常':'正常','期限已過':'期限已過','即將到期(30天)':'30','即將到期(90天)':'90','即將到期(180天)':'180','即將到期(365天)':'365','待補期限':'待補期限','日期異常':'日期異常'}};return'dl-'+(m[s]||'待補期限')}}
function stCls(s){{return'b-'+s}}
function badge(t,c){{return`<span class="badge ${{c}}">${{esc(t)}}</span>`}}
function buildAll(){{
  const a=[];
  RAW.trademark.forEach(r=>a.push({{_type:'商標',_name:r['商標案件']||'—',_country:r['國別']||'—',_status:r._status,_deadline:r._end_date||'',_dl:r._deadline_status,_cert:r['證書號 (進度)']||'—',_sub:(r['商標分類']||'')+'・'+(r['申請類別']||''),_raw:r}}));
  RAW.patent.forEach(r=>a.push({{_type:'專利',_name:r['專利名稱縮寫/分類']||r['專利名稱(中文)']||'—',_country:r['國別']||'—',_status:r._status,_deadline:'',_dl:'待補期限',_cert:r['專利編號']||'—',_sub:r['專利類別']||'',_raw:r}}));
  RAW.registration.forEach(r=>a.push({{_type:'產品登記',_name:r['登記產品名']||'—',_country:r['國別']||'—',_status:r._status,_deadline:r._end_date||'',_dl:r._deadline_status,_cert:r['證書/License ID']||'—',_sub:r['登記公司']||'',_raw:r}}));
  return a;
}}
const ALL=buildAll();
let pg='overview',flt={{q:'',type:'all',country:'all',status:'all',dl:'all'}},srt={{col:null,asc:true}},cur=1,pp=30;
function showPage(n,el){{pg=n;cur=1;flt={{q:'',type:'all',country:'all',status:'all',dl:'all'}};document.querySelectorAll('.nav-item').forEach(x=>x.classList.remove('active'));if(el)el.classList.add('active');const T={{overview:'主管總覽',trademark:'商標管理',patent:'專利管理',registration:'產品登記',alerts:'期限提醒',sync:'資料與同步'}};document.getElementById('bc').innerHTML='管理中心 ／ <span>'+T[n]+'</span>';render()}}
function render(){{const el=document.getElementById('content');if(pg==='overview')el.innerHTML=renderOv();else if(pg==='sync')el.innerHTML=SYNC_HTML;else if(pg==='alerts')el.innerHTML=renderAlerts();else el.innerHTML=renderTbl(pg)}}
function getPool(t){{if(t==='all')return ALL;const m={{trademark:'商標',patent:'專利',registration:'產品登記'}};return ALL.filter(r=>r._type===m[t])}}
function applyF(pool){{let d=[...pool];if(flt.q){{const q=flt.q.toLowerCase();d=d.filter(r=>r._name.toLowerCase().includes(q)||r._country.toLowerCase().includes(q)||(r._cert&&r._cert.toLowerCase().includes(q)))}}if(flt.type!=='all')d=d.filter(r=>r._type===flt.type);if(flt.country!=='all')d=d.filter(r=>r._country===flt.country);if(flt.status!=='all')d=d.filter(r=>r._status===flt.status);if(flt.dl!=='all')d=d.filter(r=>r._dl===flt.dl);if(srt.col)d.sort((a,b)=>{{const va=a[srt.col]||'',vb=b[srt.col]||'';return srt.asc?(va<vb?-1:va>vb?1:0):(va>vb?-1:va<vb?1:0)}});return d}}
function renderOv(){{const tm=RAW.trademark,pt=RAW.patent,rg=RAW.registration;const tmGet=tm.filter(r=>r._status==='已取得').length;const ptGet=pt.filter(r=>r._status==='已取得').length;const rgGet=rg.filter(r=>r._status==='已取得').length;const tmOver=tm.filter(r=>r._deadline_status==='期限已過').length;const tmSoon=tm.filter(r=>r._deadline_status==='即將到期(365天)').length;const rgSoon=rg.filter(r=>r._deadline_status==='即將到期(365天)').length;const alertN=ALL.filter(r=>['期限已過','即將到期(30天)','即將到期(90天)','即將到期(365天)'].includes(r._dl)).length;document.getElementById('nba').textContent=alertN;const pri=ALL.filter(r=>['期限已過','即將到期(30天)','即將到期(90天)','即將到期(365天)'].includes(r._dl)).sort((a,b)=>{{const rk={{'期限已過':0,'即將到期(30天)':1,'即將到期(90天)':2,'即將到期(365天)':3}};return(rk[a._dl]||9)-(rk[b._dl]||9)}}).slice(0,8);return`<div class="overview-grid"><div class="card"><div class="card-label">® 商標案件</div><div class="card-value">${{tm.length}}</div><div class="card-sub">已取得 ${{tmGet}} ／ 申請中 ${{tm.length-tmGet}}</div></div><div class="card"><div class="card-label">◇ 專利案件</div><div class="card-value">${{pt.length}}</div><div class="card-sub">已取得 ${{ptGet}} ／ 辦理中 ${{pt.length-ptGet-pt.filter(r=>r._status==='已結案').length}}</div></div><div class="card"><div class="card-label">▤ 產品登記</div><div class="card-value">${{rg.length}}</div><div class="card-sub">已取得 ${{rgGet}} ／ 辦理中 ${{rg.length-rgGet}}</div></div><div class="card${{tmOver>0?' ac':''}}"><div class="card-label">${{tmOver>0?'⚠ ':''}}商標期限已過</div><div class="card-value">${{tmOver}}</div><div class="card-sub">需確認是否已延展</div></div><div class="card${{(tmSoon+rgSoon)>0?' wc':''}}"><div class="card-label">${{(tmSoon+rgSoon)>0?'⏰ ':''}}365天內到期</div><div class="card-value">${{tmSoon+rgSoon}}</div><div class="card-sub">商標 ${{tmSoon}} ・ 登記 ${{rgSoon}}</div></div></div><div class="section-title">優先關注事項</div>${{pri.length===0?'<div class="plist"><div class="empty">目前無需立即關注的案件</div></div>':'<div class="plist"><div class="pr prh"><div>案件名稱</div><div>類型・國別</div><div>期限提示</div><div>關注日期</div></div>'+pri.map(r=>`<div class="pr" onclick='openMo(${{JSON.stringify(JSON.stringify(r))}})'><div><div class="cn">${{esc(r._name)}}</div></div><div style="font-size:12px">${{esc(r._type+' · '+r._country)}}</div><div>${{badge(r._dl,dlCls(r._dl))}}</div><div style="font-size:12px;color:#6b7a99">${{esc(r._deadline||'—')}}</div></div>`).join('')+'</div>'}}<div style="margin:24px 0 14px"><div class="section-title">所有案件</div></div>${{renderTblContent('all')}}`}}
function renderTbl(t){{return renderTblContent(t)}}
function renderTblContent(t){{const pool=getPool(t);const fdata=applyF(pool);const total=fdata.length,pages=Math.ceil(total/pp)||1;if(cur>pages)cur=pages;const rows=fdata.slice((cur-1)*pp,cur*pp);const countries=[...new Set(pool.map(r=>r._country))].sort();const statuses=[...new Set(pool.map(r=>r._status))].sort();const dls=[...new Set(pool.map(r=>r._dl))].filter(Boolean).sort();const isAll=t==='all';const fb=`<div class="fbar"><input type="text" placeholder="搜尋案件名稱、案號…" value="${{esc(flt.q)}}" oninput="setF('q',this.value)">${{isAll?`<select onchange="setF('type',this.value)"><option value="all">所有類型</option><option value="商標">商標</option><option value="專利">專利</option><option value="產品登記">產品登記</option></select>`:''}}<select onchange="setF('country',this.value)"><option value="all">所有國別</option>${{countries.map(c=>`<option value="${{esc(c)}}"${{flt.country===c?' selected':''}}>${{esc(c)}}</option>`).join('')}}</select><select onchange="setF('status',this.value)"><option value="all">所有狀態</option>${{statuses.map(s=>`<option value="${{esc(s)}}"${{flt.status===s?' selected':''}}>${{esc(s)}}</option>`).join('')}}</select><select onchange="setF('dl',this.value)"><option value="all">所有期限</option>${{dls.map(s=>`<option value="${{esc(s)}}"${{flt.dl===s?' selected':''}}>${{esc(s)}}</option>`).join('')}}</select><span class="frs" onclick="resetF()">重設</span><span class="rcount">共 ${{total}} 筆${{total!==pool.length?' (全部'+pool.length+'筆)':''}}</span></div>`;const tbody=rows.length===0?`<tr><td colspan="${{isAll?6:5}}"><div class="empty">無符合條件的案件</div></td></tr>`:rows.map(r=>`<tr onclick='openMo(${{JSON.stringify(JSON.stringify(r))}})'><td><div class="cn">${{esc(r._name)}}</div><div class="cs">${{esc(r._sub||'')}}</div></td>${{isAll?`<td>${{badge(r._type,'bT')}}</td>`:''}}<td>${{esc(r._country)}}</td><td>${{badge(r._status,stCls(r._status))}}</td><td style="font-size:12px;color:#6b7a99">${{esc(r._deadline||'—')}}</td><td>${{badge(r._dl,dlCls(r._dl))}}</td></tr>`).join('');let pager='';if(pages>1){{let btns='',range=[],prev=-1;for(let i=1;i<=pages;i++){{if(i===1||i===pages||Math.abs(i-cur)<=2)range.push(i)}}for(const p of range){{if(prev!==-1&&p-prev>1)btns+=`<span class="pgi">…</span>`;btns+=`<button class="pgb${{p===cur?' apg':''}}" onclick="goP(${{p}})">${{p}}</button>`;prev=p}}pager=`<div class="pgbar"><button class="pgb" onclick="goP(${{cur-1}})" ${{cur===1?'disabled':''}}>‹</button>${{btns}}<button class="pgb" onclick="goP(${{cur+1}})" ${{cur===pages?'disabled':''}}>›</button><span class="pgi">${{(cur-1)*pp+1}}–${{Math.min(cur*pp,total)}}/${{total}}</span></div>`}}return fb+`<div class="twrap"><table><thead><tr><th onclick="sortBy('_name')">案件／產品名稱</th>${{isAll?'<th>類型</th>':''}}<th onclick="sortBy('_country')">國別</th><th onclick="sortBy('_status')">狀態</th><th onclick="sortBy('_deadline')">關注日期</th><th onclick="sortBy('_dl')">期限提示</th></tr></thead><tbody>${{tbody}}</tbody></table>${{pager}}</div>`}}
function renderAlerts(){{const RANKS=['期限已過','即將到期(30天)','即將到期(90天)','即將到期(180天)','即將到期(365天)'];const labels={{'期限已過':'⚠ 期限已過','即將到期(30天)':'🔴 30天內到期','即將到期(90天)':'🟠 90天內到期','即將到期(180天)':'🟡 180天內到期','即將到期(365天)':'🟡 365天內到期'}};let html=RANKS.map(dl=>{{const items=ALL.filter(r=>r._dl===dl);if(!items.length)return'';return`<div style="margin-bottom:24px"><div class="section-title">${{labels[dl]}} (${{items.length}})</div><div class="plist"><div class="pr prh"><div>案件名稱</div><div>類型・國別</div><div>期限提示</div><div>關注日期</div></div>${{items.map(r=>`<div class="pr" onclick='openMo(${{JSON.stringify(JSON.stringify(r))}})'><div><div class="cn">${{esc(r._name)}}</div></div><div style="font-size:12px">${{esc(r._type+' · '+r._country)}}</div><div>${{badge(r._dl,dlCls(r._dl))}}</div><div style="font-size:12px;color:#6b7a99">${{esc(r._deadline||'—')}}</div></div>`).join('')}}</div></div>`}}).join('');const missing=ALL.filter(r=>r._dl==='待補期限');html+=`<div style="margin-bottom:24px"><div class="section-title">待補期限 (${{missing.length}})</div><div class="ibox">共 ${{missing.length}} 件案件尚無到期日記錄。專利案件年費期限未於原表列示，標示「待補期限」，不自行推算。</div></div>`;return html}}
function openMo(s){{let r;try{{r=JSON.parse(s)}}catch{{return}}const raw=r._raw||r;document.getElementById('mt').textContent=r._type+' 案件明細';const top=`<div class="dg" style="margin-bottom:14px"><div class="di"><label>類型</label><div class="dv">${{badge(r._type,'bT')}}</div></div><div class="di"><label>狀態</label><div class="dv">${{badge(r._status,stCls(r._status))}}</div></div><div class="di"><label>關注日期</label><div class="dv">${{esc(r._deadline||'—')}}</div></div><div class="di"><label>期限提示</label><div class="dv">${{badge(r._dl,dlCls(r._dl))}}</div></div></div><hr style="border:none;border-top:1px solid #f0f4fa;margin:4px 0"><div class="dg" style="margin-top:12px">`;const items=Object.entries(raw).filter(([k])=>!k.startsWith('_')).map(([k,v])=>`<div class="di${{['專利名稱(中文)','專利名稱(英文)','Raw Materials'].includes(k)?' full':''}}"><label>${{esc(k)}}</label><div class="dv">${{esc(v||'—')}}</div></div>`).join('');document.getElementById('mb').innerHTML=top+items+'</div>';document.getElementById('mo').classList.add('open')}}
function closeMo(e){{if(!e||e.target===document.getElementById('mo'))document.getElementById('mo').classList.remove('open')}}
function setF(k,v){{flt[k]=v;cur=1;render()}}
function resetF(){{flt={{q:'',type:'all',country:'all',status:'all',dl:'all'}};cur=1;render()}}
function sortBy(c){{if(srt.col===c)srt.asc=!srt.asc;else{{srt.col=c;srt.asc=true}}render()}}
function goP(n){{cur=n;render();document.getElementById('main').scrollTo(0,0)}}
function exportCSV(){{const pool=pg==='overview'?ALL:getPool(pg);const d=applyF(pool);const cols=['_type','_name','_country','_status','_deadline','_dl','_cert'];const lbl={{'_type':'類型','_name':'案件名稱','_country':'國別','_status':'狀態','_deadline':'關注日期','_dl':'期限提示','_cert':'證書號/狀態'}};const csv='﻿'+[cols.map(c=>lbl[c]).join(','),...d.map(r=>cols.map(c=>'"'+String(r[c]||'').replace(/"/g,'""')+'"').join(','))].join('\n');const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{{type:'text/csv;charset=utf-8'}}));a.download='正瀚智財登記_'+new Date().toISOString().slice(0,10)+'.csv';a.click()}}
document.getElementById('nba').textContent=ALL.filter(r=>['期限已過','即將到期(30天)','即將到期(90天)','即將到期(365天)'].includes(r._dl)).length;
render();
</script>
</body>
</html>'''
    return html

print('Downloading from Google Sheets...')
tm_data  = download_excel('商標進度', URLS['trademark'])
pt_data  = download_excel('專利進度', URLS['patent'])
rg_data  = download_excel('登記進度', URLS['registration'])

print('Processing data...')
trademark    = process_trademark(read_excel_rows(tm_data))
patent       = process_patent(read_excel_rows(pt_data))
registration = process_registration(read_excel_rows(rg_data))
print(f'  商標: {len(trademark)}, 專利: {len(patent)}, 登記: {len(registration)}')

print('Generating index.html...')
html = build_html(trademark, patent, registration)

out_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'index.html')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)
print(f'Done! {len(html):,} bytes -> {out_path}')
