#!/usr/bin/env python3
"""
Flask backend for Chinese Fruit Dashboard: HTML-to-Excel data sync.
Serves REST API for CRUD operations on fruits/regions/curves.
Each write operation saves the Excel file and regenerates the HTML embedded data.

Usage: python server.py
Then open http://localhost:5000/ in browser.
"""

import os
import sys
import json
import re
import subprocess
import openpyxl
from datetime import datetime
from flask import Flask, request, jsonify, send_file
from zhdate import ZhDate
import hashlib
import urllib.request
import urllib.parse
import json as json_lib

# 高德地图 Web API 配置
AMAP_KEY = "d60264473fe914fa0cf21d1a22a4e206"
AMAP_SECRET = "ae7d546c1e8f25d8365b2c601fd3c96a"

app = Flask(__name__)

# CORS — allow requests from any origin (for file:// and local dev)
@app.after_request
def add_cors(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, PUT, DELETE, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return response

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'fruit_data.xlsx')
HTML_PATH = os.path.join(BASE_DIR, '中国水果产区数据大屏.html')

# ============================================================
# Column indices (1-based) — 17-column schema with lunar
# ============================================================
# Sheet1: 水果产品库
(F_ID, F_NAME, F_CAT, F_PROV, F_CITY, F_DIST, F_TOWN) = range(1, 8)
(F_SEA_START, F_SEA_END, F_PEAK_START, F_PEAK_END) = range(8, 12)
(F_LUNAR_SS, F_LUNAR_SE, F_LUNAR_PS, F_LUNAR_PE) = range(12, 16)
(F_CURVE, F_DESC) = (16, 17)

# Sheet2: 产区映射表
(R_ADCODE, R_PROV, R_CITY, R_DIST, R_TOWN) = range(1, 6)
(R_LNG, R_LAT, R_FIDS, R_LEVEL) = range(6, 10)

# Sheet3: 成熟度曲线
C_NAME = 1  # values for months 1-12 in cols 2-13


# ============================================================
# Helpers: read/write Excel
# ============================================================
def read_excel():
    """Read all data from Excel, return dict of {fruits, regions, curves}."""
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

    # Sheet1: 水果产品库
    ws_f = wb['水果产品库']
    fruits = []
    for r in range(2, ws_f.max_row + 1):
        fid = ws_f.cell(r, F_ID).value
        if not fid or not str(fid).strip():
            continue
        fruit = {
            'id': str(fid).strip(),
            'name': str(ws_f.cell(r, F_NAME).value or '').strip(),
            'category': str(ws_f.cell(r, F_CAT).value or '').strip(),
            'province': str(ws_f.cell(r, F_PROV).value or '').strip(),
            'city': str(ws_f.cell(r, F_CITY).value or '').strip(),
            'district': str(ws_f.cell(r, F_DIST).value or '').strip(),
            'town': str(ws_f.cell(r, F_TOWN).value or '').strip(),
            'seasonStart': int(ws_f.cell(r, F_SEA_START).value or 1),
            'seasonEnd': int(ws_f.cell(r, F_SEA_END).value or 12),
            'peakStart': int(ws_f.cell(r, F_PEAK_START).value or 1),
            'peakEnd': int(ws_f.cell(r, F_PEAK_END).value or 12),
            'lunarSeasonStart': int(ws_f.cell(r, F_LUNAR_SS).value or 0),
            'lunarSeasonEnd': int(ws_f.cell(r, F_LUNAR_SE).value or 0),
            'lunarPeakStart': int(ws_f.cell(r, F_LUNAR_PS).value or 0),
            'lunarPeakEnd': int(ws_f.cell(r, F_LUNAR_PE).value or 0),
            'curveType': str(ws_f.cell(r, F_CURVE).value or '').strip(),
            'desc': str(ws_f.cell(r, F_DESC).value or '').strip(),
        }
        fruits.append(fruit)

    # Sheet2: 产区映射表
    ws_r = wb['产区映射表']
    regions = []
    for r in range(2, ws_r.max_row + 1):
        adcode = ws_r.cell(r, R_ADCODE).value
        if not adcode or not str(adcode).strip():
            continue
        region = {
            'adcode': str(adcode).strip(),
            'province': str(ws_r.cell(r, R_PROV).value or '').strip(),
            'city': str(ws_r.cell(r, R_CITY).value or '').strip(),
            'district': str(ws_r.cell(r, R_DIST).value or '').strip(),
            'town': str(ws_r.cell(r, R_TOWN).value or '').strip(),
            'lng': float(ws_r.cell(r, R_LNG).value or 0),
            'lat': float(ws_r.cell(r, R_LAT).value or 0),
            'fruitIds': str(ws_r.cell(r, R_FIDS).value or '').strip(),
            'level': str(ws_r.cell(r, R_LEVEL).value or '').strip(),
        }
        regions.append(region)

    # Sheet3: 成熟度曲线
    ws_c = wb['成熟度曲线']
    curves = []
    for r in range(2, ws_c.max_row + 1):
        name = ws_c.cell(r, C_NAME).value
        if not name or not str(name).strip():
            continue
        values = []
        for c in range(2, 14):
            values.append(int(ws_c.cell(r, c).value or 0))
        curves.append({'name': str(name).strip(), 'values': values})

    wb.close()
    return {'fruits': fruits, 'regions': regions, 'curves': curves}


# ── Helper: solar month (1-12) → lunar month (1-12) via zhdate ──
REFERENCE_YEAR = 2025

def _solar_to_lunar_month(solar_month):
    """Convert 1-based solar month to 1-based lunar month using a reference year."""
    try:
        d = ZhDate.from_datetime(datetime(REFERENCE_YEAR, max(1, min(12, solar_month)), 15))
        return d.lunar_month
    except Exception:
        return 0


def _auto_fill_lunar(body):
    """Auto-calculate lunar month fields from solar month fields."""
    for solar_key, lunar_key in [
        ('seasonStart', 'lunarSeasonStart'),
        ('seasonEnd', 'lunarSeasonEnd'),
        ('peakStart', 'lunarPeakStart'),
        ('peakEnd', 'lunarPeakEnd'),
    ]:
        solar_val = body.get(solar_key)
        if solar_val is not None and body.get(lunar_key) is None:
            try:
                body[lunar_key] = _solar_to_lunar_month(int(solar_val))
            except (ValueError, TypeError):
                body[lunar_key] = 0


def geocode(province="", city="", district="", town=""):
    """调用高德地理编码 API，返回 (lng, lat)，失败返回 (0, 0)。"""
    parts = [p for p in [province, city, district, town] if p and str(p).strip()]
    address = "".join(parts)
    if not address.strip():
        return 0, 0
    try:
        params = f"address={urllib.parse.quote(address)}&output=JSON"
        sig_raw = f"/v3/geocode/geo?{params}&key={AMAP_KEY}{AMAP_SECRET}"
        sig = hashlib.md5(sig_raw.encode()).hexdigest()
        url = f"https://restapi.amap.com/v3/geocode/geo?{params}&key={AMAP_KEY}&sig={sig}"
        req = urllib.request.Request(url)
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json_lib.loads(resp.read().decode("utf-8"))
        if data.get("status") == "1" and data.get("geocodes"):
            loc = data["geocodes"][0].get("location", "0,0")
            parts = loc.split(",")
            if len(parts) == 2:
                return float(parts[0]), float(parts[1])
    except Exception as e:
        print(f"[geocode] 查询失败: {address}, 错误: {e}")
    return 0, 0


def write_fruit_row(ws, row_num, fruit_data):
    """Write a fruit dict to a specific row in the worksheet."""
    ws.cell(row_num, F_ID).value = fruit_data.get('id', '')
    ws.cell(row_num, F_NAME).value = fruit_data.get('name', '')
    ws.cell(row_num, F_CAT).value = fruit_data.get('category', '')
    ws.cell(row_num, F_PROV).value = fruit_data.get('province', '')
    ws.cell(row_num, F_CITY).value = fruit_data.get('city', '')
    ws.cell(row_num, F_DIST).value = fruit_data.get('district', '')
    ws.cell(row_num, F_TOWN).value = fruit_data.get('town', '')
    ws.cell(row_num, F_SEA_START).value = int(fruit_data.get('seasonStart', 1))
    ws.cell(row_num, F_SEA_END).value = int(fruit_data.get('seasonEnd', 12))
    ws.cell(row_num, F_PEAK_START).value = int(fruit_data.get('peakStart', 1))
    ws.cell(row_num, F_PEAK_END).value = int(fruit_data.get('peakEnd', 12))
    ws.cell(row_num, F_LUNAR_SS).value = int(fruit_data.get('lunarSeasonStart', 0))
    ws.cell(row_num, F_LUNAR_SE).value = int(fruit_data.get('lunarSeasonEnd', 0))
    ws.cell(row_num, F_LUNAR_PS).value = int(fruit_data.get('lunarPeakStart', 0))
    ws.cell(row_num, F_LUNAR_PE).value = int(fruit_data.get('lunarPeakEnd', 0))
    ws.cell(row_num, F_CURVE).value = fruit_data.get('curveType', '')
    ws.cell(row_num, F_DESC).value = fruit_data.get('desc', '')


def save_and_regenerate():
    """Save Excel, then regenerate HTML embedded data."""
    # Regenerate embedded JSON in HTML
    data = read_excel()
    fruits_json = []
    for f in data['fruits']:
        fruits_json.append([
            f['id'], f['name'], f['category'], f['province'], f['city'],
            f['district'], f['town'],
            str(f['seasonStart']), str(f['seasonEnd']),
            str(f['peakStart']), str(f['peakEnd']),
            str(f['lunarSeasonStart']), str(f['lunarSeasonEnd']),
            str(f['lunarPeakStart']), str(f['lunarPeakEnd']),
            f['curveType'], f['desc'],
        ])
    regions_json = []
    for r in data['regions']:
        regions_json.append([
            r['adcode'], r['province'], r['city'], r['district'], r['town'],
            str(r['lng']), str(r['lat']), r['fruitIds'], r['level'],
        ])
    curves_json = []
    for c in data['curves']:
        curves_json.append([c['name']] + [str(v) for v in c['values']])

    embedded = json.dumps(
        {'fruits': fruits_json, 'regions': regions_json, 'curves': curves_json},
        ensure_ascii=False
    )

    with open(HTML_PATH, 'r', encoding='utf-8') as f:
        html = f.read()
    pattern = r'var EMBEDDED_DATA = \{.*?\};'
    replacement = 'var EMBEDDED_DATA = ' + embedded + ';'
    html_new = re.sub(pattern, replacement, html, count=1, flags=re.DOTALL)
    if html_new != html:
        with open(HTML_PATH, 'w', encoding='utf-8') as f:
            f.write(html_new)
        print(f'HTML embedded data updated: {len(fruits_json)} fruits')
    return data


def generate_next_fruit_id(wb):
    """Generate the next fruit ID (e.g., F202 -> F203)."""
    ws = wb['水果产品库']
    max_num = 0
    for r in range(2, ws.max_row + 1):
        fid = str(ws.cell(r, F_ID).value or '').strip()
        if fid.startswith('F'):
            try:
                num = int(fid[1:])
                if num > max_num:
                    max_num = num
            except ValueError:
                pass
    return f'F{max_num + 1:03d}'


# ============================================================
# API Endpoints
# ============================================================

@app.route('/')
def index():
    """Serve the dashboard HTML."""
    return send_file(HTML_PATH)


@app.route('/api/health')
def health():
    data = read_excel()
    return jsonify({
        'status': 'ok',
        'fruit_count': len(data['fruits']),
        'region_count': len(data['regions']),
        'curve_count': len(data['curves']),
    })


# ---- Fruits ----

@app.route('/api/fruits', methods=['GET'])
def get_fruits():
    data = read_excel()
    return jsonify(data['fruits'])


@app.route('/api/fruits/<fruit_id>', methods=['GET'])
def get_fruit(fruit_id):
    data = read_excel()
    for f in data['fruits']:
        if f['id'] == fruit_id:
            return jsonify(f)
    return jsonify({'error': f'Fruit {fruit_id} not found'}), 404


@app.route('/api/fruits', methods=['POST'])
def create_fruit():
    body = request.get_json()
    if not body:
        return jsonify({'error': 'Request body required'}), 400

    # Validate required fields
    required = ['name', 'category', 'province', 'city']
    for field in required:
        if not body.get(field, '').strip():
            return jsonify({'error': f'Missing required field: {field}'}), 400

    # Fallback for optional fields
    if not body.get('district', '').strip():
        body['district'] = body.get('city', '')

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb['水果产品库']
    ws_r = wb['产区映射表']

    # Auto-calculate lunar fields from solar fields
    _auto_fill_lunar(body)

    # Generate new ID
    new_id = generate_next_fruit_id(wb)
    body['id'] = new_id

    # Append fruit row
    new_row = ws.max_row + 1
    write_fruit_row(ws, new_row, body)

    # Also add a region mapping row for this fruit
    new_region_row = ws_r.max_row + 1
    ws_r.cell(new_region_row, R_ADCODE).value = str(body.get('adcode', '000000000')).strip() or '000000000'
    ws_r.cell(new_region_row, R_PROV).value = body.get('province', '')
    ws_r.cell(new_region_row, R_CITY).value = body.get('city', '')
    ws_r.cell(new_region_row, R_DIST).value = body.get('district', '')
    ws_r.cell(new_region_row, R_TOWN).value = body.get('town', '')
    try: lng = float(body.get('lng', 0))
    except: lng = 0
    try: lat = float(body.get('lat', 0))
    except: lat = 0
    # 若坐标为 0，调用高德 API 自动补全
    if lng == 0 and lat == 0:
        lng, lat = geocode(body.get('province', ''), body.get('city', ''), body.get('district', ''), body.get('town', ''))
    ws_r.cell(new_region_row, R_LNG).value = lng
    ws_r.cell(new_region_row, R_LAT).value = lat
    ws_r.cell(new_region_row, R_FIDS).value = new_id
    level_val = str(body.get('level', '一般产区')).strip()
    ws_r.cell(new_region_row, R_LEVEL).value = level_val if level_val else '一般产区'

    wb.save(EXCEL_PATH)
    wb.close()

    data = save_and_regenerate()
    created = next((f for f in data['fruits'] if f['id'] == new_id), None)
    return jsonify(created), 201


@app.route('/api/fruits/<fruit_id>', methods=['PUT'])
def update_fruit(fruit_id):
    body = request.get_json()
    if not body:
        return jsonify({'error': 'Request body required'}), 400

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb['水果产品库']

    # Find the row
    target_row = None
    for r in range(2, ws.max_row + 1):
        fid = str(ws.cell(r, F_ID).value or '').strip()
        if fid == fruit_id:
            target_row = r
            break

    if target_row is None:
        wb.close()
        return jsonify({'error': f'Fruit {fruit_id} not found'}), 404

    # 先读取当前行所有数据，只更新 body 中存在的字段（防止覆盖丢失数据）
    current = {
        'id': str(ws.cell(target_row, F_ID).value or '').strip(),
        'name': str(ws.cell(target_row, F_NAME).value or '').strip(),
        'category': str(ws.cell(target_row, F_CAT).value or '').strip(),
        'province': str(ws.cell(target_row, F_PROV).value or '').strip(),
        'city': str(ws.cell(target_row, F_CITY).value or '').strip(),
        'district': str(ws.cell(target_row, F_DIST).value or '').strip(),
        'town': str(ws.cell(target_row, F_TOWN).value or '').strip(),
        'seasonStart': int(ws.cell(target_row, F_SEA_START).value or 1),
        'seasonEnd': int(ws.cell(target_row, F_SEA_END).value or 12),
        'peakStart': int(ws.cell(target_row, F_PEAK_START).value or 1),
        'peakEnd': int(ws.cell(target_row, F_PEAK_END).value or 12),
        'lunarSeasonStart': int(ws.cell(target_row, F_LUNAR_SS).value or 0),
        'lunarSeasonEnd': int(ws.cell(target_row, F_LUNAR_SE).value or 0),
        'lunarPeakStart': int(ws.cell(target_row, F_LUNAR_PS).value or 0),
        'lunarPeakEnd': int(ws.cell(target_row, F_LUNAR_PE).value or 0),
        'curveType': str(ws.cell(target_row, F_CURVE).value or '').strip(),
        'desc': str(ws.cell(target_row, F_DESC).value or '').strip(),
    }
    # 合并：body 中有值的字段覆盖 current
    for key in current:
        if key in body and body[key] is not None:
            val = body[key]
            if key in ('seasonStart', 'seasonEnd', 'peakStart', 'peakEnd',
                       'lunarSeasonStart', 'lunarSeasonEnd', 'lunarPeakStart', 'lunarPeakEnd'):
                try: val = int(val)
                except: val = current[key]
            if isinstance(val, str):
                val = val.strip()
            current[key] = val

    # Auto-calculate lunar fields from solar fields
    _auto_fill_lunar(current)

    write_fruit_row(ws, target_row, current)

    # If location changed, update matching region mapping
    ws_r = wb['产区映射表']
    for r in range(2, ws_r.max_row + 1):
        fids = str(ws_r.cell(r, R_FIDS).value or '').strip()
        if fruit_id in fids.split(','):
            if 'province' in current: ws_r.cell(r, R_PROV).value = current.get('province', '')
            if 'city' in current: ws_r.cell(r, R_CITY).value = current.get('city', '')
            if 'district' in current: ws_r.cell(r, R_DIST).value = current.get('district', '')
            if 'town' in current: ws_r.cell(r, R_TOWN).value = current.get('town', '')
            # 若产区坐标为 0，调用高德 API 补全
            try: exist_lng = float(ws_r.cell(r, R_LNG).value or 0)
            except: exist_lng = 0
            try: exist_lat = float(ws_r.cell(r, R_LAT).value or 0)
            except: exist_lat = 0
            if exist_lng == 0 and exist_lat == 0:
                prov = current.get('province', '') or str(ws_r.cell(r, R_PROV).value or '')
                city = current.get('city', '') or str(ws_r.cell(r, R_CITY).value or '')
                dist = current.get('district', '') or str(ws_r.cell(r, R_DIST).value or '')
                town = current.get('town', '') or str(ws_r.cell(r, R_TOWN).value or '')
                new_lng, new_lat = geocode(prov, city, dist, town)
                ws_r.cell(r, R_LNG).value = new_lng
                ws_r.cell(r, R_LAT).value = new_lat
            break

    wb.save(EXCEL_PATH)
    wb.close()

    data = save_and_regenerate()
    updated = next((f for f in data['fruits'] if f['id'] == fruit_id), None)
    return jsonify(updated)


@app.route('/api/fruits/<fruit_id>', methods=['DELETE'])
def delete_fruit(fruit_id):
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb['水果产品库']

    # Find and delete fruit row
    target_row = None
    for r in range(2, ws.max_row + 1):
        fid = str(ws.cell(r, F_ID).value or '').strip()
        if fid == fruit_id:
            target_row = r
            break

    if target_row is None:
        wb.close()
        return jsonify({'error': f'Fruit {fruit_id} not found'}), 404

    ws.delete_rows(target_row)

    # Delete matching region mappings
    ws_r = wb['产区映射表']
    rows_to_delete = []
    for r in range(2, ws_r.max_row + 1):
        fids = str(ws_r.cell(r, R_FIDS).value or '').strip()
        if fruit_id in [x.strip() for x in fids.split(',')]:
            rows_to_delete.append(r)

    # Delete from bottom up to avoid index shifting
    for r in sorted(rows_to_delete, reverse=True):
        ws_r.delete_rows(r)

    wb.save(EXCEL_PATH)
    wb.close()

    data = save_and_regenerate()
    return jsonify({'success': True, 'deleted': fruit_id,
                    'fruit_count': len(data['fruits']),
                    'region_count': len(data['regions'])})


# ---- Regions ----

@app.route('/api/regions', methods=['GET'])
def get_regions():
    data = read_excel()
    return jsonify(data['regions'])


# ---- Curves ----

@app.route('/api/curves', methods=['GET'])
def get_curves():
    data = read_excel()
    return jsonify(data['curves'])


# ---- Excel Download ----

@app.route('/api/excel')
def download_excel():
    return send_file(EXCEL_PATH, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='fruit_data.xlsx')


# ---- Script execution & refresh ----

ALLOWED_SCRIPTS = {
    'generate_fruit_data.py': '生成初始水果数据',
    'update_fruit_data.py': '数据质量修复 + 新增水果',
    'enrich_fruit_data.py': '增加西瓜、菠萝、樱桃等品种',
}


@app.route('/api/run-script', methods=['POST'])
def run_script():
    """Execute one of the allowed data-generation Python scripts."""
    body = request.get_json() or {}
    script_name = body.get('script', '').strip()

    if script_name not in ALLOWED_SCRIPTS:
        return jsonify({'error': f'无效脚本。允许: {list(ALLOWED_SCRIPTS.keys())}'}), 400

    script_path = os.path.join(BASE_DIR, script_name)
    if not os.path.isfile(script_path):
        return jsonify({'error': f'脚本文件不存在: {script_name}'}), 404

    try:
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True, text=True, cwd=BASE_DIR,
            timeout=120,
        )
    except subprocess.TimeoutExpired:
        return jsonify({'error': '脚本执行超时（>120秒）'}), 500

    if result.returncode != 0:
        return jsonify({
            'success': False,
            'script': script_name,
            'error': result.stderr[-1000:] or '未知错误',
            'stdout': result.stdout[-500:],
        }), 500

    # After script runs, re-read Excel and regenerate HTML
    data = save_and_regenerate()
    return jsonify({
        'success': True,
        'script': script_name,
        'label': ALLOWED_SCRIPTS[script_name],
        'fruit_count': len(data['fruits']),
        'region_count': len(data['regions']),
        'stdout': result.stdout[-500:],
    })


@app.route('/api/refresh', methods=['POST'])
def refresh_data():
    """Re-read Excel file and regenerate HTML embedded data."""
    try:
        data = save_and_regenerate()
        return jsonify({
            'success': True,
            'fruit_count': len(data['fruits']),
            'region_count': len(data['regions']),
            'curve_count': len(data['curves']),
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ---- 批量同步：将浏览器端所有数据写入 Excel ----

@app.route('/api/sync-all', methods=['POST'])
def sync_all():
    """接收浏览器端全部水果/产区/曲线数据，完整写入 Excel 并重新生成 HTML。"""
    body = request.get_json()
    if not body:
        return jsonify({'error': 'Request body required'}), 400

    fruits_data = body.get('fruits', [])
    regions_data = body.get('regions', [])
    curves_data = body.get('curves', [])

    # 安全防护：拒绝空数据写入（防止意外清空 Excel）
    if not fruits_data and not regions_data and not curves_data:
        return jsonify({'error': '拒绝空数据写入：fruits/regions/curves 均为空，未修改 Excel'}), 400
    if len(fruits_data) < 10:
        return jsonify({'error': f'安全防护：水果数据仅 {len(fruits_data)} 条（<10），拒绝写入以防数据丢失'}), 400

    wb = openpyxl.load_workbook(EXCEL_PATH)

    # 重写 Sheet1: 水果产品库
    ws_f = wb['水果产品库']
    # 清除旧数据（保留标题行）
    for r in range(ws_f.max_row, 1, -1):
        ws_f.delete_rows(r)
    for i, f in enumerate(fruits_data):
        row_num = i + 2
        _auto_fill_lunar(f)
        write_fruit_row(ws_f, row_num, f)

    # 重写 Sheet2: 产区映射表
    ws_r = wb['产区映射表']
    for r in range(ws_r.max_row, 1, -1):
        ws_r.delete_rows(r)
    for i, r in enumerate(regions_data):
        row_num = i + 2
        ws_r.cell(row_num, R_ADCODE).value = str(r.get('adcode', '')).strip()
        ws_r.cell(row_num, R_PROV).value = str(r.get('province', '')).strip()
        ws_r.cell(row_num, R_CITY).value = str(r.get('city', '')).strip()
        ws_r.cell(row_num, R_DIST).value = str(r.get('district', '')).strip()
        ws_r.cell(row_num, R_TOWN).value = str(r.get('town', '')).strip()
        try: lng = float(r.get('lng', 0))
        except: lng = 0
        try: lat = float(r.get('lat', 0))
        except: lat = 0
        # 若坐标为 0，调用高德 API 自动补全
        if lng == 0 and lat == 0:
            lng, lat = geocode(
                str(r.get('province', '')), str(r.get('city', '')),
                str(r.get('district', '')), str(r.get('town', '')))
        ws_r.cell(row_num, R_LNG).value = lng
        ws_r.cell(row_num, R_LAT).value = lat
        ws_r.cell(row_num, R_FIDS).value = str(r.get('fruitIds', '')).strip()
        ws_r.cell(row_num, R_LEVEL).value = str(r.get('level', '一般产区')).strip() or '一般产区'

    # 重写 Sheet3: 成熟度曲线
    ws_c = wb['成熟度曲线']
    for r in range(ws_c.max_row, 1, -1):
        ws_c.delete_rows(r)
    for i, c in enumerate(curves_data):
        row_num = i + 2
        ws_c.cell(row_num, C_NAME).value = str(c.get('name', '')).strip()
        vals = c.get('values', [])
        for j, v in enumerate(vals):
            try: ws_c.cell(row_num, j + 2).value = int(v)
            except: ws_c.cell(row_num, j + 2).value = 0

    wb.save(EXCEL_PATH)
    wb.close()

    # 重新生成 HTML 嵌入式数据
    data = save_and_regenerate()
    return jsonify({
        'success': True,
        'fruit_count': len(data['fruits']),
        'region_count': len(data['regions']),
        'curve_count': len(data['curves']),
    })


# ---- Static files ----

@app.route('/fruit_data.xlsx')
def serve_excel():
    """Serve Excel file for the existing fetch('fruit_data.xlsx') pattern."""
    return send_file(EXCEL_PATH, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    print(f'Starting Fruit Dashboard API server...')
    print(f'Excel: {EXCEL_PATH}')
    print(f'HTML:  {HTML_PATH}')
    print(f'Open http://localhost:5000/ in browser')
    app.run(host='0.0.0.0', port=5000, debug=True)
