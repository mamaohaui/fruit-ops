#!/usr/bin/env python3
"""
阶段3：写入 Excel + 高德坐标补全。
读取 _verified_new.json，写入 fruit_data.xlsx 的 Sheet1 和 Sheet2，
调用高德 API 获取坐标，重新生成 HTML 嵌入式数据。
"""
import os
import sys
import json
import shutil
import time
import hashlib
import urllib.request
import urllib.parse
from datetime import datetime
import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'fruit_data.xlsx')
HTML_PATH = os.path.join(BASE_DIR, '中国水果产区数据大屏.html')
VERIFIED_PATH = os.path.join(BASE_DIR, '_verified_new.json')

AMAP_KEY = "d60264473fe914fa0cf21d1a22a4e206"
AMAP_SECRET = "ae7d546c1e8f25d8365b2c601fd3c96a"


def geocode(province="", city="", district="", town=""):
    """调用高德地理编码 API，返回 (lng, lat)，失败返回 (0, 0)。"""
    parts = [p for p in [province, city, district, town] if p and str(p).strip()]
    address = "".join(parts)
    if not address.strip():
        return 0, 0
    for attempt in range(3):
        try:
            params = f"address={urllib.parse.quote(address)}&output=JSON"
            sig_raw = f"/v3/geocode/geo?{params}&key={AMAP_KEY}{AMAP_SECRET}"
            sig = hashlib.md5(sig_raw.encode()).hexdigest()
            url = f"https://restapi.amap.com/v3/geocode/geo?{params}&key={AMAP_KEY}&sig={sig}"
            req = urllib.request.Request(url)
            with urllib.request.urlopen(req, timeout=5) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            if data.get("status") == "1" and data.get("geocodes"):
                loc = data["geocodes"][0].get("location", "0,0")
                parts_loc = loc.split(",")
                if len(parts_loc) == 2:
                    return float(parts_loc[0]), float(parts_loc[1])
        except Exception as e:
            print(f"  [geocode] 尝试 {attempt+1}/3 失败: {address}, 错误: {e}")
            time.sleep(0.5)
    print(f"  [geocode] 最终失败: {address}")
    return 0, 0


REFERENCE_YEAR = 2025

def solar_to_lunar_month(solar_month):
    """Convert 1-based solar month to 1-based lunar month using zhdate."""
    try:
        from zhdate import ZhDate
        d = ZhDate.from_datetime(datetime(REFERENCE_YEAR, max(1, min(12, int(solar_month))), 15))
        return d.lunar_month
    except Exception:
        return 0


# Sheet1 列索引（1-based）
(F_ID, F_NAME, F_CAT, F_PROV, F_CITY, F_DIST, F_TOWN) = range(1, 8)
(F_SEA_START, F_SEA_END, F_PEAK_START, F_PEAK_END) = range(8, 12)
(F_LUNAR_SS, F_LUNAR_SE, F_LUNAR_PS, F_LUNAR_PE) = range(12, 16)
(F_CURVE, F_DESC) = (16, 17)

# Sheet2 列索引（1-based）
(R_ADCODE, R_PROV, R_CITY, R_DIST, R_TOWN) = range(1, 6)
(R_LNG, R_LAT, R_FIDS, R_LEVEL) = range(6, 10)

FRUIT_CATEGORY_MAP = {
    '梨': '仁果类', '苹果': '仁果类', '山楂': '仁果类', '枇杷': '仁果类',
    '桃': '核果类', '李': '核果类', '杏': '核果类', '枣': '核果类',
    '樱桃': '核果类', '杨梅': '核果类', '橄榄': '核果类',
    '柑橘': '柑橘类', '橙': '柑橘类', '柚': '柑橘类', '柠檬': '柑橘类',
    '桔': '柑橘类', '柑': '柑橘类', '椪柑': '柑橘类', '蜜桔': '柑橘类', '蜜柑': '柑橘类', '沃柑': '柑橘类', '砂糖桔': '柑橘类', '脐橙': '柑橘类', '金桔': '柑橘类',
    '葡萄': '浆果类', '猕猴桃': '浆果类', '草莓': '浆果类',
    '蓝莓': '浆果类', '桑葚': '浆果类', '无花果': '浆果类',
    '石榴': '浆果类', '百香果': '浆果类', '柿': '浆果类', '月柿': '浆果类',
    '蓝靛果': '浆果类', '红树莓': '浆果类', '沙棘': '浆果类',
    '西瓜': '瓜果类', '甜瓜': '瓜果类', '哈密瓜': '瓜果类', '硒砂瓜': '瓜果类',
    '芒果': '热带水果', '荔枝': '热带水果', '龙眼': '热带水果',
    '香蕉': '热带水果', '菠萝': '热带水果', '火龙果': '热带水果',
    '番石榴': '热带水果', '榴莲': '热带水果', '山竹': '热带水果',
    '菠萝蜜': '热带水果', '木瓜': '热带水果', '椰子': '热带水果',
    '甘蔗': '热带水果', '莲雾': '热带水果', '释迦': '热带水果',
    '核桃': '核果类', '板栗': '核果类', '枸杞': '浆果类',
}

def guess_category(name):
    """根据品种名推断分类。"""
    for keyword, cat in FRUIT_CATEGORY_MAP.items():
        if keyword in name:
            return cat
    return '仁果类'


def backup_excel():
    """备份 Excel 文件。"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(BASE_DIR, f'fruit_data_backup_{timestamp}.xlsx')
    shutil.copy2(EXCEL_PATH, backup_path)
    print(f"已备份: {backup_path}")
    return backup_path


def generate_next_fruit_id(ws):
    """生成下一个水果 ID（F182 -> F183 -> ...）"""
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


def write_fruit_row(ws, row_num, fruit_data):
    """写入一行水果数据到 Sheet1。"""
    ws.cell(row_num, F_ID).value = fruit_data['id']
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
    ws.cell(row_num, F_LUNAR_SS).value = solar_to_lunar_month(fruit_data.get('seasonStart', 1))
    ws.cell(row_num, F_LUNAR_SE).value = solar_to_lunar_month(fruit_data.get('seasonEnd', 12))
    ws.cell(row_num, F_LUNAR_PS).value = solar_to_lunar_month(fruit_data.get('peakStart', 1))
    ws.cell(row_num, F_LUNAR_PE).value = solar_to_lunar_month(fruit_data.get('peakEnd', 12))
    ws.cell(row_num, F_CURVE).value = fruit_data.get('curveType', '')
    ws.cell(row_num, F_DESC).value = fruit_data.get('desc', '')


def main():
    backup_excel()

    with open(VERIFIED_PATH, 'r', encoding='utf-8') as f:
        verified_new = json.load(f)

    print(f"待写入记录: {len(verified_new)} 条")

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws_f = wb['水果产品库']
    ws_r = wb['产区映射表']

    write_report = {'new_fruits': [], 'new_regions': [], 'geocode_fails': []}

    for idx, record in enumerate(verified_new):
        # 推断类别
        if not record.get('category'):
            record['category'] = guess_category(record['name'])

        # 生成 ID
        new_id = generate_next_fruit_id(ws_f)
        record['id'] = new_id

        # 写入 Sheet1
        new_row = ws_f.max_row + 1
        write_fruit_row(ws_f, new_row, record)
        print(f"  [{idx+1}/{len(verified_new)}] Sheet1: {new_id} {record['name']} ({record['province']} {record['city']})")

        # 获取坐标
        lng, lat = geocode(
            record.get('province', ''),
            record.get('city', ''),
            record.get('district', ''),
            record.get('town', ''),
        )

        if lng == 0 and lat == 0 and record.get('town'):
            print(f"  降级坐标查询: {record['province']}{record['city']}{record.get('district','')}")
            lng, lat = geocode(
                record.get('province', ''),
                record.get('city', ''),
                record.get('district', ''),
                '',
            )

        if lng == 0 and lat == 0:
            write_report['geocode_fails'].append({
                'id': new_id,
                'name': record['name'],
                'location': f"{record.get('province','')}{record.get('city','')}{record.get('district','')}{record.get('town','')}"
            })

        # 生成 adcode
        adcode = '9' + str(int(time.time() * 1000))[-8:]
        time.sleep(0.01)

        # 写入 Sheet2
        new_region_row = ws_r.max_row + 1
        ws_r.cell(new_region_row, R_ADCODE).value = adcode
        ws_r.cell(new_region_row, R_PROV).value = record.get('province', '')
        ws_r.cell(new_region_row, R_CITY).value = record.get('city', '')
        ws_r.cell(new_region_row, R_DIST).value = record.get('district', '')
        ws_r.cell(new_region_row, R_TOWN).value = record.get('town', '')
        ws_r.cell(new_region_row, R_LNG).value = lng
        ws_r.cell(new_region_row, R_LAT).value = lat
        ws_r.cell(new_region_row, R_FIDS).value = new_id
        ws_r.cell(new_region_row, R_LEVEL).value = record.get('level', '一般产区')

        write_report['new_fruits'].append({'id': new_id, 'name': record['name']})
        write_report['new_regions'].append({'adcode': adcode, 'lng': lng, 'lat': lat})

        # 每 10 条暂停，避免高德 API 限流
        if (idx + 1) % 10 == 0:
            print(f"  已写入 {idx+1}/{len(verified_new)}，暂停1秒...")
            time.sleep(1)

    # 保存 Excel
    wb.save(EXCEL_PATH)
    wb.close()
    print(f"\nExcel 保存完成！")

    # 写入报告
    report_path = os.path.join(BASE_DIR, '_write_report.json')
    with open(report_path, 'w', encoding='utf-8') as f:
        json.dump(write_report, f, ensure_ascii=False, indent=2)

    print(f"新增水果: {len(write_report['new_fruits'])} 条")
    print(f"新增产区: {len(write_report['new_regions'])} 条")
    print(f"坐标失败: {len(write_report['geocode_fails'])} 条")
    if write_report['geocode_fails']:
        print("坐标失败的条目（需人工处理）:")
        for g in write_report['geocode_fails']:
            print(f"  - {g['id']} {g['name']} @ {g['location']}")


if __name__ == '__main__':
    main()
