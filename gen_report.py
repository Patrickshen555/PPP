import pandas as pd
import numpy as np
import json

raw_path = r'C:/Users/patrickshen/Desktop/工作/CRM/20260522_07_COC.xlsx'
xls_path = r'C:/Users/patrickshen/Desktop/工作/CRM/COC_客户画像评分结果_v1.xlsx'

# ========== 1. 计算淡旺季标记（排除进口）==========
raw = pd.read_excel(raw_path)
raw['LANE']   = raw['LANE'].astype(str).str.strip()
raw['POL_CD'] = raw['POL_CD'].astype(str).str.strip()
raw['REV_YM'] = pd.to_datetime(raw['YM'], format='%Y%m', errors='coerce')
raw = raw.dropna(subset=['REV_YM'])

# 进出口分类
def is_cn(c): return str(c).startswith('CN')
raw['IS_IMPORT'] = raw.apply(lambda r: not is_cn(r['POL_CD']) and is_cn(r['POD_CD']), axis=1)
# 排除进口记录做淡旺季判断
raw_noimp = raw[~raw['IS_IMPORT']].copy()
raw_noimp['CM_GP']  = raw_noimp['CM']
raw_noimp['TEU_GP'] = raw_noimp['TEU']

lane_month_rate = (
    raw_noimp.groupby(['LANE','REV_YM'])
    .agg(MONTH_CM=('CM_GP','sum'), MONTH_TEU=('TEU_GP','sum'))
    .reset_index()
)
lane_month_rate['MONTH_CM_PER_TEU'] = lane_month_rate['MONTH_CM'] / lane_month_rate['MONTH_TEU'].replace(0, np.nan)
lane_month_rate = lane_month_rate.drop(columns=['MONTH_CM','MONTH_TEU'])
lane_annual = (
    raw_noimp.groupby('LANE').agg(TOTAL_CM=('CM_GP','sum'), TOTAL_TEU=('TEU_GP','sum'))
    .reset_index()
)
lane_annual['ANNUAL_AVG'] = lane_annual['TOTAL_CM'] / lane_annual['TOTAL_TEU']
lane_month_rate = lane_month_rate.merge(lane_annual[['LANE','ANNUAL_AVG']], on='LANE', how='left')
lane_month_rate['IS_PEAK'] = (lane_month_rate['MONTH_CM_PER_TEU'] >= lane_month_rate['ANNUAL_AVG']).astype(int)

# 提取每条航线的淡旺季月份
# REV_YM是Timestamp，转换为int格式的YYYYMM
raw['YM_INT'] = raw['REV_YM'].dt.year * 100 + raw['REV_YM'].dt.month
lane_month_rate['YM_INT'] = lane_month_rate['REV_YM'].dt.year * 100 + lane_month_rate['REV_YM'].dt.month
all_months = sorted(raw['YM_INT'].unique().tolist())
lane_season = {}
for lane, grp in lane_month_rate.groupby('LANE'):
    peak_months = grp[grp['IS_PEAK']==1]['YM_INT'].tolist()
    off_months = grp[grp['IS_PEAK']==0]['YM_INT'].tolist()
    lane_season[lane] = {
        'peak': [int(m) for m in peak_months],
        'off': [int(m) for m in off_months]
    }

def prep(df, keep_cols):
    df = df.fillna('')
    return df[keep_cols].values.tolist()

xls = pd.ExcelFile(xls_path)

df1 = pd.read_excel(xls, 'L1_Overall')
l1 = prep(df1, ['Rank','Code','Name','TEU','CM','CM_NI','TEU_Score','CM_Score','OffSeason_Score','Total_Score'])
df2 = pd.read_excel(xls, 'L2_ByLane')
l2 = prep(df2, ['Lane','Rank','Code','Name','TEU','CM','CM_NI','TEU_Score','CM_Score','OffSeason_Score','Total_Score'])
df3 = pd.read_excel(xls, 'L3_ByPort')
l3 = prep(df3, ['Port','Rank','Code','Name','TEU','CM','CM_NI','TEU_Score','CM_Score','OffSeason_Score','Total_Score'])
df4 = pd.read_excel(xls, 'L4_ByPortLane')
l4 = prep(df4, ['Port','Lane','Rank','Code','Name','TEU','CM','CM_NI','TEU_Score','CM_Score','OffSeason_Score','Total_Score'])

lanes = sorted(df2['Lane'].dropna().unique().tolist())
ports = sorted(df3['Port'].dropna().unique().tolist())

def s(v):
    if pd.isna(v) or v == '':
        return ''
    return str(v)

def rows_to_json(rows):
    out = []
    for r in rows:
        out.append([s(x) for x in r])
    return json.dumps(out, ensure_ascii=False)

data_js = f"const DATA = {{\n"
data_js += f'  l1: {rows_to_json(l1)},\n'
data_js += f'  l2: {rows_to_json(l2)},\n'
data_js += f'  l3: {rows_to_json(l3)},\n'
data_js += f'  l4: {rows_to_json(l4)},\n'
data_js += f'  lanes: {json.dumps(lanes, ensure_ascii=False)},\n'
data_js += f'  ports: {json.dumps(ports, ensure_ascii=False)},\n'
data_js += f'  allMonths: {json.dumps([int(m) for m in all_months], ensure_ascii=False)},\n'
data_js += f'  laneSeason: {json.dumps(lane_season, ensure_ascii=False)}\n'
data_js += "};"

html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>客户画像评分 - 2026-05-01 更新 (v5+++含进口排除)</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family: 'Segoe UI','Microsoft YaHei',sans-serif; background:#f5f6fa; color:#2c3e50; }}
.header {{ background:linear-gradient(135deg,#2c3e50,#3498db); color:#fff; padding:24px 32px; }}
.header h1 {{ font-size:22px; font-weight:600; }}
.header p {{ font-size:13px; opacity:0.8; margin-top:4px; }}
.controls {{ background:#fff; padding:16px 32px; border-bottom:1px solid #e0e0e0; display:flex; gap:16px; align-items:center; flex-wrap:wrap; }}
.controls label {{ font-size:13px; font-weight:600; color:#555; }}
.controls select, .controls input {{
  padding:6px 12px; border:1px solid #ddd; border-radius:6px; font-size:13px;
  background:#fff; color:#333; outline:none;
}}
.controls select:focus, .controls input:focus {{ border-color:#3498db; }}
.stat-bar {{ background:#fff; padding:12px 32px; border-bottom:1px solid #e0e0e0; display:flex; gap:24px; flex-wrap:wrap; }}
.stat-item {{ font-size:13px; color:#666; }}
.stat-item span {{ font-weight:700; color:#2c3e50; }}
.table-wrap {{ padding:16px 32px; overflow-x:auto; }}
table {{ width:100%; border-collapse:collapse; font-size:13px; }}
thead th {{
  background:#2c3e50; color:#fff; padding:10px 12px; text-align:left;
  position:sticky; top:0; cursor:pointer; user-select:none; white-space:nowrap;
}}
thead th:hover {{ background:#34495e; }}
thead th .sort-arrow {{ margin-left:4px; font-size:11px; opacity:0.6; }}
thead th.sorted .sort-arrow {{ opacity:1; color:#f1c40f; }}
tbody tr {{ background:#fff; transition:background 0.15s; }}
tbody tr:hover {{ background:#eaf2f8; }}
tbody tr:nth-child(even) {{ background:#fafbfc; }}
tbody tr:nth-child(even):hover {{ background:#eaf2f8; }}
td {{ padding:8px 12px; border-bottom:1px solid #eee; white-space:nowrap; }}
td.num {{ text-align:right; font-variant-numeric:tabular-nums; }}
td.code {{ font-family:'Consolas',monospace; font-size:12px; }}
.rank-1 {{ color:#e74c3c; font-weight:700; }}
.rank-top5 {{ color:#e67e22; font-weight:600; }}
.rank-top20 {{ color:#f39c12; }}
.score-bar {{ display:inline-block; height:6px; border-radius:3px; margin-left:6px; vertical-align:middle; }}
.tag {{
  display:inline-block; padding:2px 8px; border-radius:10px; font-size:11px;
  font-weight:600; color:#fff; margin-right:4px;
}}
.tag-l1 {{ background:#3498db; }}
.tag-l2 {{ background:#e67e22; }}
.tag-l3 {{ background:#27ae60; }}
.tag-l4 {{ background:#8e44ad; }}
.pager {{ padding:12px 32px; display:flex; align-items:center; gap:12px; font-size:13px; }}
.pager button {{
  padding:6px 14px; border:1px solid #ddd; border-radius:6px; background:#fff;
  cursor:pointer; font-size:13px;
}}
.pager button:hover {{ background:#f0f0f0; }}
.pager button:disabled {{ opacity:0.4; cursor:default; }}
.month-tag {{
  display:inline-block; padding:3px 8px; border-radius:6px; font-size:12px;
  font-weight:600; line-height:1.4;
}}
.month-peak {{ background:#fff3e0; color:#e67e22; border:1px solid #f0c27f; }}
.month-off {{ background:#e8f5e9; color:#27ae60; border:1px solid #a5d6a7; }}
.month-label {{ font-size:11px; color:#999; margin-right:4px; }}
</style>
</head>
<body>

<div class="header">
  <h1>客户画像评分 - 2026-05-01 更新 (v5+++)</h1>
  <p>分位数分段映射法 | 三维度综合评价 | 四个独立视图 | 含进口排除机制</p>
</div>

<div class="controls">
  <label>视图</label>
  <select id="viewSelect" onchange="onViewChange()">
    <option value="l1">视图1：总体评分（{len(l1)}条）</option>
    <option value="l2">视图2：航线评分（{len(l2)}条）</option>
    <option value="l3">视图3：口岸评分（{len(l3)}条）</option>
    <option value="l4">视图4：口岸x航线评分（{len(l4)}条）</option>
  </select>

  <label id="filterLabel" style="display:none">筛选</label>
  <select id="filterSelect" onchange="applyFilter()" style="display:none">
    <option value="">全部</option>
  </select>
  <label id="filterLabel2" style="display:none">航线</label>
  <select id="filterSelect2" onchange="applyFilter()" style="display:none">
    <option value="">全部航线</option>
  </select>

  <label>搜索</label>
  <input id="searchInput" type="text" placeholder="客户代码或名称..." oninput="applyFilter()" style="width:200px">

  <label>排序</label>
  <select id="sortSelect" onchange="onSortChange()">
    <option value="-1">默认排名</option>
  </select>
  <button id="sortDirBtn" onclick="toggleSortDir()" style="padding:6px 12px;border:1px solid #ddd;border-radius:6px;background:#fff;cursor:pointer;font-size:13px;">↓ 降序</button>
</div>

<div class="stat-bar" id="statBar"></div>
<div id="seasonBar" style="background:#fff;padding:10px 32px;border-bottom:1px solid #e0e0e0;display:none;">
  <span style="font-size:13px;font-weight:600;color:#555;">淡旺季标记</span>
  <span id="seasonLane" style="font-size:13px;font-weight:700;color:#2c3e50;margin-left:8px;"></span>
  <div id="seasonMonths" style="margin-top:6px;display:flex;gap:4px;flex-wrap:wrap;"></div>
</div>

<div class="table-wrap">
  <table>
    <thead id="thead"></thead>
    <tbody id="tbody"></tbody>
  </table>
</div>

<div class="pager" id="pager"></div>

<script>
{data_js}

const PAGE_SIZE = 100;
let currentView = 'l1';
let filteredData = [];
let sortCol = -1;
let sortDir = -1;
let currentPage = 1;
const colMeta = {{
  l1: [
    {{key:'Rank',label:'排名',type:'num'}},
    {{key:'Code',label:'客户代码',type:'str'}},
    {{key:'Name',label:'客户名称',type:'str'}},
    {{key:'TEU',label:'TEU',type:'num'}},
    {{key:'CM',label:'CM',type:'num'}},
    {{key:'CM_NI',label:'CM_NI',type:'num'}},
    {{key:'TEU_Score',label:'TEU得分',type:'num'}},
    {{key:'CM_Score',label:'CM得分',type:'num'}},
    {{key:'OffSeason_Score',label:'CM_OFF得分',type:'num'}},
    {{key:'Total_Score',label:'综合总分',type:'num'}}
  ],
  l2: [
    {{key:'Lane',label:'航线',type:'str'}},
    {{key:'Rank',label:'排名',type:'num'}},
    {{key:'Code',label:'客户代码',type:'str'}},
    {{key:'Name',label:'客户名称',type:'str'}},
    {{key:'TEU',label:'TEU',type:'num'}},
    {{key:'CM',label:'CM',type:'num'}},
    {{key:'CM_NI',label:'CM_NI',type:'num'}},
    {{key:'TEU_Score',label:'TEU得分',type:'num'}},
    {{key:'CM_Score',label:'CM得分',type:'num'}},
    {{key:'OffSeason_Score',label:'CM_OFF得分',type:'num'}},
    {{key:'Total_Score',label:'综合总分',type:'num'}}
  ],
  l3: [
    {{key:'Port',label:'口岸',type:'str'}},
    {{key:'Rank',label:'排名',type:'num'}},
    {{key:'Code',label:'客户代码',type:'str'}},
    {{key:'Name',label:'客户名称',type:'str'}},
    {{key:'TEU',label:'TEU',type:'num'}},
    {{key:'CM',label:'CM',type:'num'}},
    {{key:'CM_NI',label:'CM_NI',type:'num'}},
    {{key:'TEU_Score',label:'TEU得分',type:'num'}},
    {{key:'CM_Score',label:'CM得分',type:'num'}},
    {{key:'OffSeason_Score',label:'CM_OFF得分',type:'num'}},
    {{key:'Total_Score',label:'综合总分',type:'num'}}
  ],
  l4: [
    {{key:'Port',label:'口岸',type:'str'}},
    {{key:'Lane',label:'航线',type:'str'}},
    {{key:'Rank',label:'排名',type:'num'}},
    {{key:'Code',label:'客户代码',type:'str'}},
    {{key:'Name',label:'客户名称',type:'str'}},
    {{key:'TEU',label:'TEU',type:'num'}},
    {{key:'CM',label:'CM',type:'num'}},
    {{key:'CM_NI',label:'CM_NI',type:'num'}},
    {{key:'TEU_Score',label:'TEU得分',type:'num'}},
    {{key:'CM_Score',label:'CM得分',type:'num'}},
    {{key:'OffSeason_Score',label:'CM_OFF得分',type:'num'}},
    {{key:'Total_Score',label:'综合总分',type:'num'}}
  ]
}};

function onViewChange() {{
  currentView = document.getElementById('viewSelect').value;
  sortCol = -1; sortDir = -1; currentPage = 1;

  // 更新排序下拉
  const ss = document.getElementById('sortSelect');
  const meta = colMeta[currentView];
  ss.innerHTML = '<option value="-1">默认排名</option>' +
    meta.map((c, i) => '<option value="'+i+'">' + c.label + '</option>').join('');
  ss.value = '-1';
  document.getElementById('sortDirBtn').textContent = '↓ 降序';

  const fs = document.getElementById('filterSelect');
  const fs2 = document.getElementById('filterSelect2');
  const fl = document.getElementById('filterLabel');
  const fl2 = document.getElementById('filterLabel2');
  fs2.style.display = 'none'; fl2.style.display = 'none';

  if (currentView === 'l2') {{
    fl.style.display = ''; fl.textContent = '航线'; fs.style.display = '';
    fs.innerHTML = '<option value="">全部航线</option>' + DATA.lanes.map(l => '<option value="'+l+'">'+l+'</option>').join('');
  }} else if (currentView === 'l3') {{
    fl.style.display = ''; fl.textContent = '口岸'; fs.style.display = '';
    fs.innerHTML = '<option value="">全部口岸</option>' + DATA.ports.map(p => '<option value="'+p+'">'+p+'</option>').join('');
  }} else if (currentView === 'l4') {{
    fl.style.display = ''; fl.textContent = '口岸'; fs.style.display = '';
    fs.innerHTML = '<option value="">全部口岸</option>' + DATA.ports.map(p => '<option value="'+p+'">'+p+'</option>').join('');
    fl2.style.display = ''; fs2.style.display = '';
    fs2.innerHTML = '<option value="">全部航线</option>' + DATA.lanes.map(l => '<option value="'+l+'">'+l+'</option>').join('');
  }} else {{
    fl.style.display = 'none'; fs.style.display = 'none';
  }}

  applyFilter();
  updateSeasonBar();
}}

function updateSeasonBar() {{
  const bar = document.getElementById('seasonBar');
  const fv = document.getElementById('filterSelect').value;
  const showSeason = currentView === 'l2' || currentView === 'l4';

  if (!showSeason) {{ bar.style.display = 'none'; return; }}

  let laneName = '';
  if (currentView === 'l2') laneName = fv;
  if (currentView === 'l4') laneName = document.getElementById('filterSelect2').value;

  if (!laneName || !DATA.laneSeason[laneName]) {{
    bar.style.display = 'none';
    return;
  }}

  const season = DATA.laneSeason[laneName];
  bar.style.display = '';
  document.getElementById('seasonLane').textContent = laneName + ' 航线';

  const monthsDiv = document.getElementById('seasonMonths');
  monthsDiv.innerHTML = DATA.allMonths.map(m => {{
    const isPeak = season.peak.includes(m);
    const isOff = season.off.includes(m);
    const label = String(m).slice(4);
    if (isPeak) return '<span class="month-tag month-peak">' + label + ' 旺</span>';
    if (isOff) return '<span class="month-tag month-off">' + label + ' 淡</span>';
    return '<span class="month-tag" style="background:#f5f5f5;color:#bbb;border:1px solid #ddd;">' + label + '</span>';
  }}).join('');
}}

function applyFilter() {{
  const view = DATA[currentView];
  const search = document.getElementById('searchInput').value.toLowerCase().trim();
  const fv = document.getElementById('filterSelect').value;
  const fv2 = document.getElementById('filterSelect2').value;

  filteredData = view.filter(r => {{
    if (currentView === 'l2' && fv && r[0] !== fv) return false;
    if (currentView === 'l3' && fv && r[0] !== fv) return false;
    if (currentView === 'l4') {{
      if (fv && r[0] !== fv) return false;
      if (fv2 && r[1] !== fv2) return false;
    }}
    if (search) {{
      const code = String(r[colMeta[currentView].findIndex(c=>c.key==='Code') || 1]).toLowerCase();
      const name = String(r[colMeta[currentView].findIndex(c=>c.key==='Name') || 2]).toLowerCase();
      if (!code.includes(search) && !name.includes(search)) return false;
    }}
    return true;
  }});

  if (sortCol >= 0) doSort();
  else currentPage = 1;
  render();
  updateSeasonBar();
}}

function doSort() {{
  const meta = colMeta[currentView];
  filteredData.sort((a, b) => {{
    let va = a[sortCol], vb = b[sortCol];
    if (meta[sortCol].type === 'num') {{
      va = parseFloat(va) || 0;
      vb = parseFloat(vb) || 0;
    }} else {{
      va = String(va); vb = String(vb);
    }}
    if (va < vb) return -1 * sortDir;
    if (va > vb) return 1 * sortDir;
    return 0;
  }});
}}

function onSort(colIdx) {{
  if (sortCol === colIdx) sortDir *= -1;
  else {{ sortCol = colIdx; sortDir = -1; }}
  document.getElementById('sortSelect').value = String(colIdx);
  document.getElementById('sortDirBtn').textContent = sortDir === -1 ? '↓ 降序' : '↑ 升序';
  doSort();
  currentPage = 1;
  render();
}}

function onSortChange() {{
  sortCol = parseInt(document.getElementById('sortSelect').value);
  if (sortCol < 0) {{ sortCol = -1; }}
  else {{ sortDir = -1; }}
  document.getElementById('sortDirBtn').textContent = sortDir === -1 ? '↓ 降序' : '↑ 升序';
  if (sortCol >= 0) doSort();
  currentPage = 1;
  render();
}}

function toggleSortDir() {{
  if (sortCol < 0) return;
  sortDir *= -1;
  document.getElementById('sortDirBtn').textContent = sortDir === -1 ? '↓ 降序' : '↑ 升序';
  doSort();
  currentPage = 1;
  render();
}}

function render() {{
  const meta = colMeta[currentView];
  // thead
  const thead = document.getElementById('thead');
  thead.innerHTML = '<tr>' + meta.map((c, i) => {{
    const sorted = sortCol === i;
    const arrow = sorted ? (sortDir === 1 ? ' &#9650;' : ' &#9660;') : '';
    return '<th class="' + (sorted?'sorted':'') + '" onclick="onSort('+i+')">' + c.label + '<span class="sort-arrow">' + arrow + '</span></th>';
  }}).join('') + '</tr>';

  // stats
  const statBar = document.getElementById('statBar');
  const viewLabels = {{l1:'总体',l2:'航线',l3:'口岸',l4:'口岸x航线'}};
  if (filteredData.length > 0) {{
    const totalIdx = meta.length - 1;
    const teuIdx = meta.findIndex(c=>c.key==='TEU_Score');
    const cmIdx = meta.findIndex(c=>c.key==='CM_Score');
    const offIdx = meta.findIndex(c=>c.key==='OffSeason_Score');
    const totals = filteredData.map(r => parseFloat(r[totalIdx]) || 0);
    const avgTotal = (totals.reduce((a,b)=>a+b,0) / totals.length).toFixed(1);
    const medTotal = totals.sort((a,b)=>a-b)[Math.floor(totals.length/2)].toFixed(1);
    statBar.innerHTML =
      '<span class="stat-item">视图：<span class="tag tag-'+currentView.slice(-1)+'">' + viewLabels[currentView] + '</span></span>' +
      '<span class="stat-item">记录数：<span>' + filteredData.length + '</span></span>' +
      '<span class="stat-item">综合总分均值：<span>' + avgTotal + '</span></span>' +
      '<span class="stat-item">综合总分中位数：<span>' + medTotal + '</span></span>';
  }} else {{
    statBar.innerHTML = '<span class="stat-item">无匹配记录</span>';
  }}

  // tbody - paginated
  const tbody = document.getElementById('tbody');
  const totalPages = Math.max(1, Math.ceil(filteredData.length / PAGE_SIZE));
  if (currentPage > totalPages) currentPage = totalPages;
  const start = (currentPage - 1) * PAGE_SIZE;
  const page = filteredData.slice(start, start + PAGE_SIZE);

  tbody.innerHTML = page.map(r => {{
    return '<tr>' + meta.map((c, i) => {{
      let val = r[i];
      let cls = '';
      if (c.type === 'num' && val !== '') {{
        cls = 'num';
        val = parseFloat(val);
        if (isNaN(val)) val = '';
        else val = val.toLocaleString('zh-CN', {{maximumFractionDigits:1}});
      }}
      if (c.key === 'Rank' && val !== '') {{
        const rank = parseInt(val);
        if (rank === 1) cls += ' rank-1';
        else if (rank <= 5) cls += ' rank-top5';
        else if (rank <= 20) cls += ' rank-top20';
      }}
      if (c.key === 'Code') cls = 'code';
      if (c.key === 'Total_Score' && val !== '') {{
        const score = parseFloat(val.replace(/,/g,''));
        if (!isNaN(score)) {{
          const barW = Math.min(Math.max(score / 8, 0), 200);
          const barColor = score > 200 ? '#e74c3c' : score > 100 ? '#e67e22' : score > 50 ? '#27ae60' : '#95a5a6';
          val = val + '<span class="score-bar" style="width:'+barW+'px;background:'+barColor+'"></span>';
        }}
      }}
      if (c.key === 'TEU_Score' || c.key === 'CM_Score' || c.key === 'OffSeason_Score') {{
        const score = parseFloat(String(val).replace(/,/g,''));
        if (!isNaN(score)) {{
          const barW = Math.min(Math.max(score / 8, 0), 100);
          const barColor = score > 200 ? '#e74c3c' : score > 100 ? '#e67e22' : score > 50 ? '#27ae60' : '#95a5a6';
          val = val + '<span class="score-bar" style="width:'+barW+'px;background:'+barColor+'"></span>';
        }}
      }}
      return '<td class="'+cls+'">' + val + '</td>';
    }}).join('') + '</tr>';
  }}).join('');

  // pager
  const pager = document.getElementById('pager');
  if (totalPages <= 1) {{
    pager.innerHTML = '<span>共 ' + filteredData.length + ' 条</span>';
  }} else {{
    pager.innerHTML =
      '<span>共 ' + filteredData.length + ' 条，第 ' + currentPage + '/' + totalPages + ' 页</span>' +
      '<button onclick="goPage(1)" '+(currentPage<=1?'disabled':'')+'>首页</button>' +
      '<button onclick="goPage('+(currentPage-1)+')" '+(currentPage<=1?'disabled':'')+'>上一页</button>' +
      '<button onclick="goPage('+(currentPage+1)+')" '+(currentPage>=totalPages?'disabled':'')+'>下一页</button>' +
      '<button onclick="goPage('+totalPages+')" '+(currentPage>=totalPages?'disabled':'')+'>末页</button>';
  }}
}}

function goPage(p) {{
  const totalPages = Math.max(1, Math.ceil(filteredData.length / PAGE_SIZE));
  currentPage = Math.max(1, Math.min(p, totalPages));
  render();
  document.querySelector('.table-wrap').scrollTop = 0;
}}

onViewChange();
</script>
</body>
</html>'''

with open(r'c:/Users/patrickshen/WorkBuddy/20260517005736/客户画像评分_20260501.html', 'w', encoding='utf-8') as f:
    f.write(html)

import os
size = os.path.getsize(r'c:/Users/patrickshen/WorkBuddy/20260517005736/客户画像评分_20260501.html')
print(f'HTML generated: {size/1024/1024:.1f} MB')
