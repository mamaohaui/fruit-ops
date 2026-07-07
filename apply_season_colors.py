"""
Modify the dashboard HTML to implement season-based gradient coloring.
Changes:
1. Add month selector UI in header
2. Add getFruitMaturityInMonth function
3. Modify getRegionAggregation to include maturity score
4. Update renderMap visualMap with green->yellow->red->purple
5. Update tooltip with production stage
6. Update panel with season info
"""
import re

with open('中国水果产区数据大屏.html', 'r', encoding='utf-8') as f:
    html = f.read()

# === Change 1: Add month selector CSS ===
css_insert = '''
/* Month selector */
.month-selector{display:flex;align-items:center;gap:8px}
.month-btn{width:28px;height:28px;border:1px solid var(--border-glow);border-radius:4px;background:var(--bg-card);color:var(--accent2);cursor:pointer;font-size:14px;transition:all .2s;display:flex;align-items:center;justify-content:center}
.month-btn:hover{background:rgba(14,165,233,0.15);border-color:var(--accent)}
.month-label{font-size:13px;color:var(--accent2);font-weight:600;min-width:60px;text-align:center}
.season-legend{display:flex;gap:4px;align-items:center;margin-left:12px}
.season-dot{width:10px;height:10px;border-radius:50%;display:inline-block}
.season-text{font-size:10px;color:var(--text-dim)}'''

# Insert before the map-area CSS
html = html.replace('/* ===== Main Content ===== */', css_insert + '\n/* ===== Main Content ===== */')

# === Change 2: Add month selector HTML in header ===
old_header_actions = '''<div class="header-actions">
	    <button class="btn" onclick="drillUp()" title="返回上级">⬆ 返回上级</button>
	    <button class="btn" onclick="resetToChina()" title="复位到全国视图">🏠 全国总览</button>
	    <button class="btn" id="btnRefresh" onclick="refreshData()" title="重新读取Excel">🔄 刷新数据</button>
	    <input type="file" id="fileInput" accept=".xlsx,.xls" style="display:none" onchange="onFileSelected(event)">
	  </div>'''

new_header_actions = '''<div class="header-actions">
	    <div class="season-legend">
	      <span class="season-dot" style="background:#22c55e"></span><span class="season-text">初产期</span>
	      <span class="season-dot" style="background:#eab308"></span><span class="season-text">成长期</span>
	      <span class="season-dot" style="background:#ef4444"></span><span class="season-text">盛产期</span>
	      <span class="season-dot" style="background:#a855f7"></span><span class="season-text">尾果期</span>
	    </div>
	    <div class="month-selector">
	      <button class="month-btn" onclick="changeMonth(-1)" title="上个月">◀</button>
	      <span class="month-label" id="monthLabel">6月</span>
	      <button class="month-btn" onclick="changeMonth(1)" title="下个月">▶</button>
	    </div>
	    <button class="btn" onclick="drillUp()" title="返回上级">⬆ 返回上级</button>
	    <button class="btn" onclick="resetToChina()" title="复位到全国视图">🏠 全国总览</button>
	    <button class="btn" id="btnRefresh" onclick="refreshData()" title="重新读取Excel">🔄 刷新数据</button>
	    <input type="file" id="fileInput" accept=".xlsx,.xls" style="display:none" onchange="onFileSelected(event)">
	  </div>'''

html = html.replace(old_header_actions, new_header_actions)

# === Change 3: Add getFruitMaturityInMonth function ===
# Insert after parseExcelData function, before GeoJSON loading section
old_section = '/* ===== GeoJSON Loading & Map Rendering ===== */'
new_func = '''/* ===== Maturity Calculation ===== */
function getFruitMaturityInMonth(fruit, month) {
  var curve = state.curveDb[fruit.curveType];
  if (!curve) return 0;
  return curve[month - 1] || 0; // month is 1-indexed, curve array is 0-indexed
}

function getRegionMaturity(regionName, level) {
  var fruits = getFruitsForRegion(regionName, level);
  if (fruits.length === 0) return 0;
  var total = 0;
  fruits.forEach(function(f) {
    total += getFruitMaturityInMonth(f, state.currentMonth);
  });
  return Math.round(total / fruits.length);
}

function getMaturityStage(score) {
  if (score >= 85) return { name: '尾果期', color: '#a855f7', icon: '🍂' };
  if (score >= 60) return { name: '盛产期', color: '#ef4444', icon: '🔥' };
  if (score >= 35) return { name: '成长期', color: '#eab308', icon: '🌞' };
  if (score >= 10) return { name: '初产期', color: '#22c55e', icon: '🌱' };
  return { name: '休眠期', color: '#64748b', icon: '💤' };
}

function getMaturityColor(score) {
  // Green (0-30) -> Yellow (30-50) -> Red (50-80) -> Purple (80-100)
  if (score >= 85) return '#9333ea'; // deep purple - late
  if (score >= 70) return '#7c3aed'; // purple-red - late peak
  if (score >= 55) return '#dc2626'; // red - peak
  if (score >= 40) return '#ea580c'; // orange-red - approaching peak
  if (score >= 25) return '#eab308'; // yellow - growing
  if (score >= 10) return '#65a30d'; // lime green - early
  return '#16a34a'; // green - starting
}

function changeMonth(delta) {
  state.currentMonth += delta;
  if (state.currentMonth < 1) state.currentMonth = 12;
  if (state.currentMonth > 12) state.currentMonth = 1;
  document.getElementById('monthLabel').textContent = state.currentMonth + '月';
  // Re-render current map
  if (state.drillStack.length === 0) {
    loadAndRenderChina();
  } else {
    var current = state.drillStack[state.drillStack.length - 1];
    renderMap(String(current.adcode), current.name);
  }
}

/* ===== GeoJSON Loading & Map Rendering ===== */'''

html = html.replace(old_section, new_func + '\n' + old_section)

# === Change 4: Update renderMap to use maturity-based coloring ===
old_render = '''function renderMap(mapKey, mapName) {
	  const regions = getRegionAggregation(mapName);
	  const maxCount = Math.max(1, ...regions.map(r => r.count));

	  const option = {
	    backgroundColor: 'transparent',
	    tooltip: {
	      trigger: 'item',
	      backgroundColor: 'rgba(10,18,50,0.92)',
	      borderColor: 'rgba(0,150,255,0.3)',
	      textStyle: { color: '#e2e8f0', fontSize: 13 },
	      formatter: function(p) {
	        var r = regions.find(function(x) { return x.name === p.name; });
	        var count = r ? r.count : 0;
	        var names = r && r.topFruits ? r.topFruits.slice(0,5).map(function(f){return f.name;}).join('<br/>  ') : '';
	        var html = '<strong style="font-size:14px">' + p.name + '</strong><br/>水果品种: <strong>' + count + ' 种</strong>';
	        if (names) html += '<br/><br/>代表产品:<br/>  ' + names;
	        return html;
	      }
	    },
	    visualMap: {
	      show: true,
	      min: 0,
	      max: maxCount,
	      left: 10,
	      bottom: 10,
	      text: ['多', '少'],
	      textStyle: { color: '#94a3b8' },
	      inRange: { color: ['#1a3a5c', '#1a6db5', '#2098e0', '#38bdf8', '#7dd3fc', '#e0f2fe'] },
	    },
	    series: [{
	      type: 'map',
	      map: mapKey,
	      roam: true,
	      zoom: 1,
	      scaleLimit: { min: 0.8, max: 20 },
	      label: { show: true, color: '#94a3b8', fontSize: 10 },
	      emphasis: {
	        label: { show: true, color: '#fff', fontSize: 13, fontWeight: 'bold' },
	        itemStyle: {
	          areaColor: 'rgba(14,165,233,0.3)',
	          borderColor: '#22d3ee',
	          borderWidth: 2,
	        },
	      },
	      itemStyle: {
	        areaColor: '#0d1f3c',
	        borderColor: 'rgba(0,150,255,0.35)',
	        borderWidth: 1,
	      },
	      data: regions,
	    }],
	  };'''

new_render = '''function renderMap(mapKey, mapName) {
	  const regions = getRegionAggregation(mapName);
	  const maxMaturity = Math.max(1, ...regions.map(function(r) { return r.maturity; }));

	  // Build maturity label with month info
	  var monthNames = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
	  var currentMonthLabel = monthNames[state.currentMonth - 1];

	  const option = {
	    backgroundColor: 'transparent',
	    tooltip: {
	      trigger: 'item',
	      backgroundColor: 'rgba(10,18,50,0.92)',
	      borderColor: 'rgba(0,150,255,0.3)',
	      textStyle: { color: '#e2e8f0', fontSize: 13 },
	      formatter: function(p) {
	        var r = regions.find(function(x) { return x.name === p.name; });
	        var count = r ? r.count : 0;
	        var maturity = r ? r.maturity : 0;
	        var stage = getMaturityStage(maturity);
	        var names = r && r.topFruits ? r.topFruits.slice(0,5).map(function(f){return f.name;}).join('<br/>  ') : '';
	        var html = '<strong style="font-size:14px">' + p.name + '</strong>' +
	          '<br/>当前月份: <strong>' + currentMonthLabel + '</strong>' +
	          '<br/>水果品种: <strong>' + count + ' 种</strong>' +
	          '<br/>产季状态: <span style="color:' + stage.color + '">' + stage.icon + ' ' + stage.name + '</span>' +
	          '<br/>成熟指数: <strong>' + maturity + '%</strong>';
	        if (names) html += '<br/><br/>代表产品:<br/>  ' + names;
	        return html;
	      }
	    },
	    visualMap: {
	      show: true,
	      min: 0,
	      max: Math.max(maxMaturity, 30),
	      left: 10,
	      bottom: 10,
	      text: [currentMonthLabel + '盛产', currentMonthLabel + '淡季'],
	      textStyle: { color: '#94a3b8' },
	      inRange: {
	        color: ['#16a34a', '#65a30d', '#84cc16', '#eab308', '#f59e0b', '#ea580c', '#dc2626', '#b91c1c', '#a855f7', '#9333ea']
	      },
	      calculable: false,
	    },
	    series: [{
	      type: 'map',
	      map: mapKey,
	      roam: true,
	      zoom: 1,
	      scaleLimit: { min: 0.8, max: 20 },
	      label: { show: true, color: '#94a3b8', fontSize: 10 },
	      emphasis: {
	        label: { show: true, color: '#fff', fontSize: 13, fontWeight: 'bold' },
	        itemStyle: {
	          areaColor: 'rgba(14,165,233,0.3)',
	          borderColor: '#22d3ee',
	          borderWidth: 2,
	        },
	      },
	      itemStyle: {
	        areaColor: '#0d1f3c',
	        borderColor: 'rgba(0,150,255,0.35)',
	        borderWidth: 1,
	      },
	      data: regions.map(function(r) {
	        return { name: r.name, value: r.maturity, count: r.count, topFruits: r.topFruits, maturity: r.maturity };
	      }),
	    }],
	  };'''

html = html.replace(old_render, new_render)

# === Change 5: Update getRegionAggregation to include maturity ===
# Find the line that pushes results
old_result_push = "result.push({ name: name, value: fruits.length, count: fruits.length, topFruits: topFruits });"
new_result_push = "      // Compute aggregate maturity for this region\n" + \
    "      var maturityScore = 0;\n" + \
    "      if (fruits.length > 0) {\n" + \
    "        var totalM = 0;\n" + \
    "        fruits.forEach(function(f) {\n" + \
    "          totalM += getFruitMaturityInMonth(f, state.currentMonth);\n" + \
    "        });\n" + \
    "        maturityScore = Math.round(totalM / fruits.length);\n" + \
    "      }\n" + \
    "      result.push({ name: name, value: maturityScore, count: fruits.length, topFruits: topFruits, maturity: maturityScore });"

html = html.replace(old_result_push, new_result_push)

# === Change 6: Update scatter series data format ===
# The scatter data should work with the new map data format

# === Change 7: Update debugStatus to show month info ===
old_debug = "ds.textContent = '| 级别:' + (state.drillStack.length === 0 ? '全国' : state.drillStack[state.drillStack.length - 1].level) + ' | 区域数:' + regions.length + ' | 已就绪:' + state.dataReady;"
new_debug = "ds.textContent = '| 月份:' + state.currentMonth + '月 | 级别:' + (state.drillStack.length === 0 ? '全国' : state.drillStack[state.drillStack.length - 1].level) + ' | 区域数:' + regions.length + ' | 水果:' + state.fruitDb.length + '种' + ' | 曲线:' + Object.keys(state.curveDb).length + '条';"

html = html.replace(old_debug, new_debug)

# === Change 8: Update initial debugStatus display ===
old_init_debug = "if (ds) { ds.textContent = '| 数据:' + (state.dataReady ? 'OK(' + state.fruitDb.length + '种水果)' : '未加载') + ' | 就绪'; }"
new_init_debug = "if (ds) { ds.textContent = '| 月份:' + state.currentMonth + '月 | 数据:' + (state.dataReady ? 'OK(' + state.fruitDb.length + '种水果)' : '未加载') + ' | 就绪'; }"

html = html.replace(old_init_debug, new_init_debug)

# === Write result ===
with open('中国水果产区数据大屏.html', 'w', encoding='utf-8') as f:
    f.write(html)

# Validate braces
open_count = html.count('{')
close_count = html.count('}')
print(f'Braces: {open_count} open / {close_count} close')
print(f'Balanced: {open_count == close_count}')
print(f'File size: {len(html)} bytes')
print('Done - dashboard updated with season-based coloring')
