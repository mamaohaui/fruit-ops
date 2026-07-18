#!/usr/bin/env python3
"""
阶段2：清洗与交叉验证脚本。
读取 _collected_data.json 和 fruit_data.xlsx，
比对现有数据、交叉验证、去重，输出三类 JSON。
"""
import json
import os
import openpyxl
from collections import defaultdict

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'fruit_data.xlsx')
COLLECTED_PATH = os.path.join(BASE_DIR, '_collected_data.json')

def read_existing_fruits():
    """从 Excel 读取现有水果数据，返回列表。"""
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb['水果产品库']
    fruits = []
    for r in range(2, ws.max_row + 1):
        fid = str(ws.cell(r, 1).value or '').strip()
        if not fid:
            continue
        fruits.append({
            'id': fid,
            'name': str(ws.cell(r, 2).value or '').strip(),
            'category': str(ws.cell(r, 3).value or '').strip(),
            'province': str(ws.cell(r, 4).value or '').strip(),
            'city': str(ws.cell(r, 5).value or '').strip(),
            'district': str(ws.cell(r, 6).value or '').strip(),
            'town': str(ws.cell(r, 7).value or '').strip(),
            'seasonStart': float(ws.cell(r, 8).value or 1),
            'seasonEnd': float(ws.cell(r, 9).value or 12),
            'peakStart': float(ws.cell(r, 10).value or 1),
            'peakEnd': float(ws.cell(r, 11).value or 12),
            'desc': str(ws.cell(r, 17).value or '').strip(),
        })
    wb.close()
    return fruits

def make_key(name, province, city, district=''):
    """生成比对的标准化 key。"""
    n = name.strip().replace(' ', '')
    p = province.strip()
    c = city.strip()
    d = district.strip() if district else ''
    return f"{n}|{p}|{c}|{d}"

def main():
    existing = read_existing_fruits()
    print(f"现有水果记录: {len(existing)} 条")

    with open(COLLECTED_PATH, 'r', encoding='utf-8') as f:
        collected = json.load(f)
    records = collected.get('records', [])
    print(f"采集到的候选记录: {len(records)} 条")

    # 构建现有数据的索引
    existing_keys = {}       # key -> fruit dict
    existing_names = set()   # 品种名集合
    existing_provinces = defaultdict(set)  # 省份 -> {品种名}

    for ef in existing:
        for d in [ef['district'], '']:
            key = make_key(ef['name'], ef['province'], ef['city'], d)
            existing_keys[key] = ef
        existing_names.add(ef['name'])
        existing_provinces[ef['province']].add(ef['name'])

    verified_new = []
    unverified = []
    merge_conflicts = []
    skipped = []

    for r in records:
        name = r['name'].strip()
        prov = r['province'].strip()
        city = r['city'].strip()
        district = r.get('district', '').strip()
        town = r.get('town', '').strip()

        # 1. 完全匹配（品种+省+市+区县）→ 跳过
        exact_key = make_key(name, prov, city, district)
        city_key = make_key(name, prov, city, '')
        if exact_key in existing_keys:
            existing_f = existing_keys[exact_key]
            has_new_info = False
            if town and not existing_f.get('town'):
                has_new_info = True
            if r.get('desc') and len(r.get('desc', '')) > len(existing_f.get('desc', '')):
                has_new_info = True
            if has_new_info:
                merge_conflicts.append({
                    'type': 'enrich_existing',
                    'collected': r,
                    'existing': existing_f,
                    'action': 'manual_review',
                })
            else:
                skipped.append({'name': name, 'province': prov, 'city': city, 'reason': '完全匹配已存在'})
            continue

        # 2. 品种+省+市匹配，区县不同 → 补充
        if city_key in existing_keys:
            existing_f = existing_keys[city_key]
            merge_conflicts.append({
                'type': 'different_district',
                'collected': r,
                'existing': existing_f,
                'action': 'manual_review',
            })
            continue

        # 3. 品种存在于该省但不同城市 → 标记审核
        if prov in existing_provinces and name in existing_provinces[prov]:
            merge_conflicts.append({
                'type': 'same_province_different_city',
                'collected': r,
                'action': 'manual_review',
            })
            continue

        # 4. 品种名完全不存在 OR 品种存在于其他省份 → 新增
        sources = r.get('sources', [])
        if len(sources) >= 2:
            r['verified'] = True
            r['confidence'] = 'high'
            verified_new.append(r)
        elif len(sources) == 1:
            r['verified'] = False
            r['confidence'] = 'low'
            verified_new.append(r)
        else:
            unverified.append(r)

    # 新增数据内部去重
    seen = {}
    deduped_new = []
    dup_count = 0
    for r in verified_new:
        key = make_key(r['name'], r['province'], r['city'], r.get('district', ''))
        if key in seen:
            existing_entry = seen[key]
            if len(r.get('sources', [])) > len(existing_entry.get('sources', [])):
                seen[key] = r
            dup_count += 1
        else:
            seen[key] = r
            deduped_new.append(r)

    if dup_count > 0:
        print(f"新增数据内部去重: {dup_count} 组重复")

    print(f"\n===== 验证结果 =====")
    print(f"验证通过（待写入）: {len(deduped_new)} 条")
    print(f"来源不足（待确认）: {len(unverified)} 条")
    print(f"冲突待处理: {len(merge_conflicts)} 条")
    print(f"已存在跳过: {len(skipped)} 条")

    # 写入输出文件
    for path, data in [
        ('_verified_new.json', deduped_new),
        ('_unverified.json', unverified),
        ('_merge_conflicts.json', merge_conflicts),
    ]:
        full = os.path.join(BASE_DIR, path)
        with open(full, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"已写入: {path} ({len(data)} 条)")

    # 打印冲突详情
    if merge_conflicts:
        print(f"\n===== 冲突详情 =====")
        for mc in merge_conflicts:
            cn = mc['collected']['name']
            cp = mc['collected']['province']
            cc = mc['collected']['city']
            print(f"  [{mc['type']}] {cn} @ {cp} {cc}")

    # 打印跳过的
    if skipped:
        print(f"\n===== 已存在跳过（前10条）=====")
        for s in skipped[:10]:
            print(f"  {s['name']} @ {s['province']} {s['city']}")

if __name__ == '__main__':
    main()
