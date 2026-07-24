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
import time
import openpyxl
from datetime import datetime
from flask import Flask, request, jsonify, send_file

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
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

# PyInstaller 打包适配：frozen 时 exe 目录可写，_MEIPASS 存放只读资源
if getattr(sys, 'frozen', False):
    _APP_DIR = os.path.dirname(sys.executable)          # exe 所在目录（可写）
    _RES_DIR = sys._MEIPASS                              # 临时解压目录（只读）
else:
    _APP_DIR = os.path.dirname(os.path.abspath(__file__))
    _RES_DIR = _APP_DIR

EXCEL_PATH = os.path.join(_APP_DIR, 'fruit_data.xlsx')
HTML_PATH = os.path.join(_RES_DIR, '中国水果产区数据大屏.html')

# 首次运行时从资源目录复制 Excel 到可写目录
if getattr(sys, 'frozen', False) and not os.path.exists(EXCEL_PATH):
    import shutil
    _src = os.path.join(_RES_DIR, 'fruit_data.xlsx')
    if os.path.exists(_src):
        shutil.copy2(_src, EXCEL_PATH)
        print(f'[初始化] 已复制数据文件到 {EXCEL_PATH}')

# ── 内存缓存：减少重复 Excel 读取 ──
_cache = {'data': None, 'time': 0}
CACHE_TTL = 2  # 秒，足以覆盖页面加载时的连续 API 请求
_html_dirty = False  # 启动时即刻生成，无需延迟

# ============================================================
# Column indices (1-based) — 13-column schema
# ============================================================
# Sheet1: 水果产品库
(F_ID, F_NAME, F_CAT, F_PROV, F_CITY, F_DIST, F_TOWN) = range(1, 8)
(F_SEA_START, F_SEA_END, F_PEAK_START, F_PEAK_END) = range(8, 12)
(F_CURVE, F_DESC) = (12, 13)

# Sheet2: 产区映射表
(R_ADCODE, R_PROV, R_CITY, R_DIST, R_TOWN) = range(1, 6)
(R_LNG, R_LAT, R_FIDS, R_LEVEL) = range(6, 10)

# Sheet3: 成熟度曲线
C_NAME = 1  # values for months 1-12 in cols 2-13


# ============================================================
# Helpers: read/write Excel
# ============================================================
def read_excel():
    """Read all data from Excel with short cache. Returns dict of {fruits, regions, curves}."""
    now = time.time()
    if _cache['data'] is not None and (now - _cache['time']) < CACHE_TTL:
        return _cache['data']

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
            'seasonStart': float(ws_f.cell(r, F_SEA_START).value or 1),
            'seasonEnd': float(ws_f.cell(r, F_SEA_END).value or 12),
            'peakStart': float(ws_f.cell(r, F_PEAK_START).value or 1),
            'peakEnd': float(ws_f.cell(r, F_PEAK_END).value or 12),
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
    _cache['data'] = {'fruits': fruits, 'regions': regions, 'curves': curves}
    _cache['time'] = time.time()
    return _cache['data']


def _in_range(m, start, end):
    """判断月份 m 是否在 [start, end] 区间内，支持跨年（start > end）。"""
    if start <= end:
        return start <= m <= end
    return m >= start or m <= end


def _month_diff(start, end):
    """计算从 start 到 end 的月份跨度（含两端），支持跨年。"""
    if start <= end:
        return end - start + 1
    return (12 - start + 1) + end


def validate_season_consistency(body):
    """校验上市期/盛产期字段是否构成合理的生命周期（支持小数月份）。
    返回 (True, dict) 或 (False, error_message)。
    """
    try:
        ss = float(body.get('seasonStart', 1))
        se = float(body.get('seasonEnd', 12))
        ps = float(body.get('peakStart', 1))
        pe = float(body.get('peakEnd', 12))
    except (ValueError, TypeError):
        return False, '月份字段必须是数字'

    # 1. 范围校验（1.0 ~ 12.9）
    for name, val in [('seasonStart', ss), ('seasonEnd', se),
                       ('peakStart', ps), ('peakEnd', pe)]:
        if val < 1.0 or val > 12.9:
            return False, f'{name} 必须在 1.0-12.9 之间，当前值: {val}'

    # 2. 盛产期必须在上市期内
    if not _in_range(ps, ss, se):
        return False, f'盛产开始月({ps})不在上市期({ss}-{se})内'
    if not _in_range(pe, ss, se):
        return False, f'盛产结束月({pe})不在上市期({ss}-{se})内'

    # 3. 计算各阶段长度（整月计数）
    season_len = _month_diff(int(ss), int(se))
    peak_len = _month_diff(int(ps), int(pe))

    # 成熟期: ss → ps 之前（整月数）
    maturity_end = int(ps) - 1 if int(ps) > 1 else 12
    maturity_len = _month_diff(int(ss), maturity_end) if int(ps) != int(ss) else 0

    # 尾产期: pe 之后 → se（整月数）
    tail_start = int(pe) + 1 if int(pe) < 12 else 1
    tail_len = _month_diff(tail_start, int(se)) if int(pe) != int(se) else 0

    # 4. 规则校验
    if season_len < 2:
        return False, f'上市期总长度至少2个月，当前{season_len}个月'

    if tail_len == 0:
        return False, (
            f'尾产期不能为0。盛产结束月({pe})等于果品结束月({se})，'
            f'请将果品结束月延后至少1个月，或提前盛产结束月。'
        )

    if peak_len > season_len - 1:
        return False, (
            f'盛产期({peak_len}个月)过长，上市期共{season_len}个月。'
            f'至少需要留1个月给尾产期。'
        )

    return True, {
        'season_len': season_len,
        'maturity_len': maturity_len,
        'peak_len': peak_len,
        'tail_len': tail_len,
    }


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
    ws.cell(row_num, F_SEA_START).value = float(fruit_data.get('seasonStart', 1))
    ws.cell(row_num, F_SEA_END).value = float(fruit_data.get('seasonEnd', 12))
    ws.cell(row_num, F_PEAK_START).value = float(fruit_data.get('peakStart', 1))
    ws.cell(row_num, F_PEAK_END).value = float(fruit_data.get('peakEnd', 12))
    ws.cell(row_num, F_CURVE).value = fruit_data.get('curveType', '')
    ws.cell(row_num, F_DESC).value = fruit_data.get('desc', '')


def invalidate_cache():
    """Clear the read cache so next read_excel() hits disk."""
    _cache['data'] = None
    _cache['time'] = 0


def save_and_regenerate():
    """Invalidate cache, mark HTML dirty, return current data (defers HTML write)."""
    invalidate_cache()
    global _html_dirty
    _html_dirty = True
    return read_excel()


def regenerate_html():
    """Actually regenerate the embedded JSON in the HTML file."""
    data = read_excel()
    # Format month: 11.0 -> "11", 4.5 -> "4.5"
    def _fmt_month(v):
        if v == int(v):
            return str(int(v))
        return str(v)

    fruits_json = []
    for f in data['fruits']:
        fruits_json.append([
            f['id'], f['name'], f['category'], f['province'], f['city'],
            f['district'], f['town'],
            _fmt_month(f['seasonStart']), _fmt_month(f['seasonEnd']),
            _fmt_month(f['peakStart']), _fmt_month(f['peakEnd']),
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
    # 使用 lambda replacement 防止 re.sub 对 JSON 中的 \n \t 等序列进行二次转义
    pattern = r'var EMBEDDED_DATA = \{.*?\};'
    html_new = re.sub(
        pattern,
        lambda _: 'var EMBEDDED_DATA = ' + embedded + ';',
        html, count=1, flags=re.DOTALL
    )
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
    """Serve the dashboard HTML, regenerating embedded data if stale."""
    global _html_dirty
    if _html_dirty:
        regenerate_html()
        _html_dirty = False
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

    # 验证生命周期一致性
    valid, result = validate_season_consistency(body)
    if not valid:
        wb.close()
        return jsonify({'error': result}), 400

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
        'seasonStart': float(ws.cell(target_row, F_SEA_START).value or 1),
        'seasonEnd': float(ws.cell(target_row, F_SEA_END).value or 12),
        'peakStart': float(ws.cell(target_row, F_PEAK_START).value or 1),
        'peakEnd': float(ws.cell(target_row, F_PEAK_END).value or 12),
        'curveType': str(ws.cell(target_row, F_CURVE).value or '').strip(),
        'desc': str(ws.cell(target_row, F_DESC).value or '').strip(),
    }
    # 合并：body 中有值的字段覆盖 current
    for key in current:
        if key in body and body[key] is not None:
            val = body[key]
            if key in ('seasonStart', 'seasonEnd', 'peakStart', 'peakEnd'):
                try: val = float(val)
                except: val = current[key]
            if isinstance(val, str):
                val = val.strip()
            current[key] = val

    # 验证生命周期一致性
    valid, result = validate_season_consistency(current)
    if not valid:
        wb.close()
        return jsonify({'error': result}), 400

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


@app.route('/api/regions', methods=['POST'])
def create_region():
    """创建独立产区行（不强制关联水果）。"""
    body = request.get_json()
    if not body:
        return jsonify({'error': 'Request body required'}), 400

    province = (body.get('province', '') or '').strip()
    city = (body.get('city', '') or '').strip()
    if not province:
        return jsonify({'error': '缺少必填字段: province'}), 400
    if not city:
        return jsonify({'error': '缺少必填字段: city'}), 400

    district = (body.get('district', '') or '').strip()
    town = (body.get('town', '') or '').strip()
    # 若未提供 adcode，使用时间戳生成一个唯一标识
    adcode = (body.get('adcode', '') or '').strip()
    if not adcode:
        adcode = '9' + str(int(time.time() * 1000))[-8:]
    level = (body.get('level', '') or '').strip()
    if not level:
        level = '一般产区'

    try:
        lng = float(body.get('lng', 0))
    except (ValueError, TypeError):
        lng = 0
    try:
        lat = float(body.get('lat', 0))
    except (ValueError, TypeError):
        lat = 0
    if lng == 0 and lat == 0:
        lng, lat = geocode(province, city, district, town)

    fruit_ids = (body.get('fruitIds', '') or '').strip()

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws_r = wb['产区映射表']
    new_row = ws_r.max_row + 1
    ws_r.cell(new_row, R_ADCODE).value = adcode
    ws_r.cell(new_row, R_PROV).value = province
    ws_r.cell(new_row, R_CITY).value = city
    ws_r.cell(new_row, R_DIST).value = district
    ws_r.cell(new_row, R_TOWN).value = town
    ws_r.cell(new_row, R_LNG).value = lng
    ws_r.cell(new_row, R_LAT).value = lat
    ws_r.cell(new_row, R_FIDS).value = fruit_ids
    ws_r.cell(new_row, R_LEVEL).value = level
    wb.save(EXCEL_PATH)
    wb.close()

    data = save_and_regenerate()
    created = next((r for r in data['regions'] if r['adcode'] == adcode), None)
    return jsonify(created), 201


# ---- Curves ----

@app.route('/api/curves', methods=['GET'])
def get_curves():
    data = read_excel()
    return jsonify(data['curves'])


# ---- Markets ----

# 市场表头列索引（1-based，对应「农批市场」Sheet）
(M_ID, M_PROV, M_NAME, M_CITY, M_LEVEL, M_ADDR,
 M_COVERAGE, M_PRODUCTS, M_NOTES, M_LNG, M_LAT) = range(1, 12)


@app.route('/api/markets', methods=['GET'])
def get_markets():
    """返回全国农批市场数据（从 Excel「农批市场」Sheet 读取）。"""
    markets = []
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    if '农批市场' in wb.sheetnames:
        ws = wb['农批市场']
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[1]:  # 跳过无市场名称的空行
                continue
            markets.append({
                'id': str(row[M_ID-1] or '').strip(),
                'province': str(row[M_PROV-1] or '').strip(),
                'name': str(row[M_NAME-1] or '').strip(),
                'city': str(row[M_CITY-1] or '').strip(),
                'level': str(row[M_LEVEL-1] or '').strip(),
                'address': str(row[M_ADDR-1] or '').strip(),
                'coverage': str(row[M_COVERAGE-1] or '').strip(),
                'products': str(row[M_PRODUCTS-1] or '').strip(),
                'notes': str(row[M_NOTES-1] or '').strip(),
                'lng': float(row[M_LNG-1]) if row[M_LNG-1] else 0,
                'lat': float(row[M_LAT-1]) if row[M_LAT-1] else 0,
            })
    wb.close()
    return jsonify(markets)


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
    # 启动时预生成 HTML 嵌入数据（避免首次请求卡顿）
    try:
        regenerate_html()
        print(f'[就绪] HTML 数据已嵌入，访问 http://localhost:5000/')
    except Exception as e:
        print(f'[警告] HTML 预生成失败（将在首次请求时重试）: {e}')
    app.run(host='0.0.0.0', port=5000, debug=True, use_reloader=False)
