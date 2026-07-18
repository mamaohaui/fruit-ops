#!/usr/bin/env python3
"""
阶段4：一致性校验与报告。
1. 生命周期一致性校验
2. 坐标范围检查
3. 重复检查
4. 关联完整性检查
5. 处理现有 6 组重复
6. 输出变更报告
"""
import os
import json
import openpyxl
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'fruit_data.xlsx')

# Sheet 列索引
(F_ID, F_NAME, F_CAT, F_PROV, F_CITY, F_DIST, F_TOWN) = range(1, 8)
(F_SEA_START, F_SEA_END, F_PEAK_START, F_PEAK_END) = range(8, 12)
(R_ADCODE, R_PROV, R_CITY, R_DIST, R_TOWN) = range(1, 6)
(R_LNG, R_LAT, R_FIDS, R_LEVEL) = range(6, 10)


def _in_range(m, start, end):
    if start <= end:
        return start <= m <= end
    return m >= start or m <= end


def check_season_consistency(fruit):
    """检查生命周期一致性，返回 (ok, message)。"""
    try:
        ss = float(fruit['seasonStart'])
        se = float(fruit['seasonEnd'])
        ps = float(fruit['peakStart'])
        pe = float(fruit['peakEnd'])
    except (ValueError, TypeError):
        return False, f"月份字段非数字"

    if not (1.0 <= ss <= 12.9 and 1.0 <= se <= 12.9):
        return False, f"上市期超出范围: {ss}-{se}"
    if not (1.0 <= ps <= 12.9 and 1.0 <= pe <= 12.9):
        return False, f"盛产期超出范围: {ps}-{pe}"
    if not _in_range(ps, ss, se):
        return False, f"盛产开始({ps})不在上市期({ss}-{se})内"
    if not _in_range(pe, ss, se):
        return False, f"盛产结束({pe})不在上市期({ss}-{se})内"

    # 计算上市期长度
    if int(ss) <= int(se):
        season_len = int(se) - int(ss) + 1
    else:
        season_len = (12 - int(ss) + 1) + int(se)
    if season_len < 2:
        return False, f"上市期仅{season_len}个月"

    return True, "OK"


def check_coords(lng, lat):
    """检查坐标是否在中国境内合理范围。"""
    if lng == 0 and lat == 0:
        return False, "坐标为(0,0)"
    if not (73 <= lng <= 135):
        return False, f"经度{lng}超出中国范围(73-135)"
    if not (18 <= lat <= 54):
        return False, f"纬度{lat}超出中国范围(18-54)"
    return True, "OK"


# ── 已知的 6 组重复（需合并）──
KNOWN_DUPES = [
    {'ids': ['F010', 'F108'], 'name': '平和大溪琯溪蜜柚', 'keep': 'F010', 'remove': 'F108'},
    {'ids': ['F013', 'F109'], 'name': '梅州金柚', 'keep': 'F013', 'remove': 'F109'},
    {'ids': ['F046', 'F107'], 'name': '苍溪红心猕猴桃', 'keep': 'F046', 'remove': 'F107'},
    {'ids': ['F070', 'F093'], 'name': '烟台大樱桃', 'keep': 'F070', 'remove': 'F093'},
    {'ids': ['F075', 'F086'], 'name': '彭阳红梅杏', 'keep': 'F075', 'remove': 'F086'},
    {'ids': ['F117', 'F159'], 'name': '蒲江猕猴桃', 'keep': 'F117', 'remove': 'F159'},
]


def main():
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws_f = wb['水果产品库']
    ws_r = wb['产区映射表']

    issues = {
        'season_errors': [],
        'coord_errors': [],
        'duplicates': [],
        'missing_regions': [],
        'dedup_done': [],
        'zero_coord_new': [],
    }

    # 1. 读取所有水果
    fruits = []
    fruit_ids = set()
    for r in range(2, ws_f.max_row + 1):
        fid = str(ws_f.cell(r, F_ID).value or '').strip()
        if not fid:
            continue
        f = {
            'row': r,
            'id': fid,
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
        }
        fruits.append(f)
        fruit_ids.add(fid)

    # 2. 生命周期校验
    for f in fruits:
        ok, msg = check_season_consistency(f)
        if not ok:
            issues['season_errors'].append({'id': f['id'], 'name': f['name'], 'error': msg})

    # 3. 坐标校验
    for r in range(2, ws_r.max_row + 1):
        adcode = str(ws_r.cell(r, R_ADCODE).value or '').strip()
        if not adcode:
            continue
        lng = float(ws_r.cell(r, R_LNG).value or 0)
        lat = float(ws_r.cell(r, R_LAT).value or 0)
        ok, msg = check_coords(lng, lat)
        prov = str(ws_r.cell(r, R_PROV).value or '')
        city = str(ws_r.cell(r, R_CITY).value or '')
        if not ok:
            issues['coord_errors'].append({
                'adcode': adcode, 'province': prov, 'city': city,
                'error': msg, 'lng': lng, 'lat': lat
            })

    # 4. 重复检查（品种+省+市+区县）
    seen = {}
    for f in fruits:
        key = f"{f['name']}|{f['province']}|{f['city']}|{f['district']}"
        if key in seen:
            issues['duplicates'].append({
                'key': key, 'ids': [seen[key]['id'], f['id']], 'name': f['name']
            })
        else:
            seen[key] = f

    # 5. 关联完整性
    region_fids = set()
    for r in range(2, ws_r.max_row + 1):
        fids_str = str(ws_r.cell(r, R_FIDS).value or '').strip()
        for fid in fids_str.split(','):
            fid = fid.strip()
            if fid:
                region_fids.add(fid)

    for fid in fruit_ids:
        if fid not in region_fids:
            f = next((x for x in fruits if x['id'] == fid), None)
            issues['missing_regions'].append({
                'id': fid, 'name': f['name'] if f else '?'
            })

    # 6. 处理已知重复
    for dupe in KNOWN_DUPES:
        keep_id = dupe['keep']
        remove_id = dupe['remove']
        keep_exists = keep_id in fruit_ids
        remove_exists = remove_id in fruit_ids
        if keep_exists and remove_exists:
            issues['dedup_done'].append({
                'name': dupe['name'],
                'keep': keep_id,
                'remove': remove_id,
                'action': 'remove_duplicate',
            })

    # 7. 统计摘要
    total_regions = ws_r.max_row - 1

    # 统计各省份分布
    province_counts = {}
    for f in fruits:
        p = f['province']
        province_counts[p] = province_counts.get(p, 0) + 1

    summary = {
        'checked_at': datetime.now().isoformat(),
        'total_fruits': len(fruits),
        'total_regions': total_regions,
        'new_records_added': len(fruits) - 178,  # 原有178条
        'season_errors': len(issues['season_errors']),
        'coord_errors': len(issues['coord_errors']),
        'duplicates_found': len(issues['duplicates']),
        'missing_regions': len(issues['missing_regions']),
        'dedup_candidates': len(issues['dedup_done']),
        'provinces_covered': len(province_counts),
        'province_distribution': province_counts,
    }

    report = {
        'summary': summary,
        'issues': issues,
    }

    report_path = os.path.join(BASE_DIR, '_enrich_report.json')
    with open(report_path, 'w', encoding='utf-8') as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    print("=" * 50)
    print("数据校验报告")
    print("=" * 50)
    print(f"水果总数: {summary['total_fruits']} (原有178 + 新增{summary['new_records_added']})")
    print(f"产区总数: {summary['total_regions']}")
    print(f"覆盖省份: {summary['provinces_covered']}")
    print(f"生命周期错误: {summary['season_errors']}")
    print(f"坐标异常: {summary['coord_errors']}")
    print(f"发现重复: {summary['duplicates_found']}")
    print(f"缺失产区关联: {summary['missing_regions']}")
    print(f"待去重组: {summary['dedup_candidates']}")

    if issues['season_errors']:
        print(f"\n--- 生命周期错误详情 ---")
        for e in issues['season_errors']:
            print(f"  {e['id']} {e['name']}: {e['error']}")

    if issues['duplicates']:
        print(f"\n--- 重复记录 ---")
        for d in issues['duplicates']:
            print(f"  {d['name']}: {d['ids']}")

    print(f"\n--- 省份分布 ---")
    for p, c in sorted(province_counts.items(), key=lambda x: -x[1]):
        marker = " ★新增" if p in ['内蒙古自治区','吉林省','黑龙江省','青海省','西藏自治区'] else ""
        print(f"  {p}: {c} 条{marker}")

    print(f"\n完整报告: {report_path}")

    wb.close()


if __name__ == '__main__':
    main()
