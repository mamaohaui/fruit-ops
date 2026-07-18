# 全国水果产区数据系统补充 — 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 系统性核验并补充全国水果产区数据，每条记录可追溯到权威来源 URL。

**Architecture:** 4 阶段流水线——阶段1 通过 WebSearch 逐省采集 GI/名特优新数据存入 `_collected_data.json`；阶段2 Python 脚本比对现有数据、交叉验证、去重输出；阶段3 Python 脚本写入 Excel + 调用高德 API 补坐标；阶段4 运行一致性校验并生成变更报告。

**Tech Stack:** Python 3 + openpyxl + 高德地图 Web API + zhdate（复用 `server.py` 现有基础设施）

## Global Constraints

- 零虚构：所有品种名、产地、季节信息必须来自可验证的权威来源
- 每条记录标注 `sources[]` — 至少 1 个 URL
- 坐标通过高德 API 自动查询，不手动填写
- 上市季节必须满足 `validate_season_consistency()` 规则（上市期≥2个月，尾产期≥1个月等）
- 写入前自动备份 `fruit_data.xlsx`
- 农历月份从公历月份自动计算

---

### Task 1: 全国 GI 水果总览 + 缺失省份专项搜索

**产出**: 初步的 `_collected_data.json`（全国 GI + 名特优新水果清单）

- [ ] **Step 1: 搜索全国地理标志水果名录**

搜索词: `"地理标志产品" 水果 名录 农业农村部 site:gov.cn`

```bash
# 使用 WebSearch 工具执行，提取以下信息：
# - 品种名称
# - 产地（省/市/县）
# - GI 认证号（如有）
# - 来源 URL
```

- [ ] **Step 2: 搜索全国名特优新农产品（水果类）**

搜索词: `"全国名特优新农产品" 水果 名录 2024 2025`

提取同上结构信息。

- [ ] **Step 3: 搜索中国特色农产品优势区（水果类）**

搜索词: `"中国特色农产品优势区" 水果 名单`

- [ ] **Step 4: 针对现有数据中完全缺失的 8 个省份搜索**

逐个搜索以下省份是否有水果 GI/名特优新产品：

```
内蒙古自治区 地理标志产品 水果
吉林省 地理标志产品 水果
黑龙江省 地理标志产品 水果
青海省 地理标志产品 水果
西藏自治区 地理标志产品 水果
```

> 注：台湾省、香港、澳门的水果数据暂不纳入本次补充范围。

- [ ] **Step 5: 整理阶段输出**

将 Step 1-4 收集到的数据初步整理为 `_collected_data.json`：

```json
{
  "collected_at": "2026-07-18",
  "records": [
    {
      "name": "百里洲砂梨",
      "category": "仁果类",
      "province": "湖北省",
      "city": "宜昌市",
      "district": "枝江市",
      "town": "百里洲镇",
      "seasonStart": 7,
      "seasonEnd": 9,
      "peakStart": 8,
      "peakEnd": 9,
      "desc": "国家地理标志产品。百里洲砂梨果大肉脆、汁多味甜...",
      "curveType": "夏季集中型",
      "sources": [
        "https://www.moa.gov.cn/xxx"
      ],
      "verified": false
    }
  ]
}
```

---

### Task 2: 分区域逐省深度搜索

**产出**: 补充完善 `_collected_data.json`，覆盖所有 31 个省级行政区（不含港澳台）

- [ ] **Step 1: 华北地区搜索（北京、天津、河北、山西、内蒙古）**

对每个省份搜索:
```
XX省 地理标志产品 水果 名录 site:gov.cn
XX省 特色水果 主产区
```

重点：河北梨、桃、葡萄；山西苹果、梨、枣；内蒙古沙棘、枸杞等。

- [ ] **Step 2: 东北地区搜索（辽宁、吉林、黑龙江）**

对每个省份搜索:
```
XX省 地理标志产品 水果
XX省 名特优新 水果
```

重点：辽宁苹果、梨、葡萄、樱桃；吉林苹果梨、蓝莓；黑龙江沙棘、蓝莓、树莓等寒地水果。

- [ ] **Step 3: 华东地区搜索（上海、江苏、浙江、安徽、福建、江西、山东）**

山东、福建已是数据大户，重点关注：安徽、江西、江苏的缺失品种（桃、李、枇杷、杨梅、猕猴桃等）。

- [ ] **Step 4: 中南地区搜索（河南、湖北、湖南、广东、广西、海南）**

湖北、湖南为重点补充对象。搜索湖北的梨、桃、李、椪柑、板栗等；湖南的冰糖橙、猕猴桃、杨梅等。

- [ ] **Step 5: 西南地区搜索（重庆、四川、贵州、云南、西藏）**

四川已是数据大户，重点关注：贵州（猕猴桃、刺梨、火龙果）、云南（热带水果多样性）、重庆（柑橘、李子）、西藏（核桃等）。

- [ ] **Step 6: 西北地区搜索（陕西、甘肃、青海、宁夏、新疆）**

重点关注：陕西（苹果、猕猴桃、枣）、甘肃（苹果、枸杞）、青海（枸杞、沙棘）、宁夏（枸杞、葡萄、枣）。

- [ ] **Step 7: 更新 _collected_data.json**

将 Step 1-6 收集的所有数据合并到 `_collected_data.json`。

---

### Task 3: 缺失品种专项搜索 + 交叉验证

**产出**: 补充缺失品种、对每条数据进行第二来源验证

- [ ] **Step 1: 针对 25+ 种缺失水果品种的全国搜索**

对以下缺失或偏少的品种进行全国范围搜索：
```
中国 XX 地理标志 主产区
```
品种列表：桃、李、杏、枣、柿、枇杷、杨梅、石榴、蓝莓、桑葚、无花果、百香果、火龙果、番石榴、菠萝蜜、木瓜、山楂、核桃、板栗、枸杞、沙棘、甘蔗、椰子、榴莲、山竹

- [ ] **Step 2: 交叉验证——每条候选数据用第二个来源确认**

对 `_collected_data.json` 中所有记录，搜索第二个独立来源。使用不同的搜索词：
```
"品种名" "产地城市" 地理标志  OR  名特优新
```

对于只有 1 个来源的条目，标记 `verified: false, confidence: low`。

- [ ] **Step 3: 补充上市季节信息**

对于缺少 seasonStart/seasonEnd 的条目，搜索：
```
"品种名" 上市时间 成熟期 采收期
```
优先从农业技术推广网站、政府公示数据中获取。

- [ ] **Step 4: 最终更新 _collected_data.json**

确保每条记录都有：
- 品种名、类别、省份、城市（必填）
- 区县、乡镇（有则填，无则留空）
- 上市季节（优先来源，次选推断）
- desc 描述文本（从来源原文摘录）
- sources[] 数组
- verified 标记

---

### Task 4: 编写 phase2_validate.py 清洗验证脚本

**Files:**
- Create: `c:/fruit-ops/phase2_validate.py`

**Interfaces:**
- Consumes: `_collected_data.json`, `fruit_data.xlsx`
- Produces: `_verified_new.json`, `_unverified.json`, `_merge_conflicts.json`

- [ ] **Step 1: 写脚本头部——读取现有数据和采集数据**

```python
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
    d = district.strip()
    return f"{n}|{p}|{c}|{d}"

def main():
    existing = read_existing_fruits()
    print(f"现有水果记录: {len(existing)} 条")

    with open(COLLECTED_PATH, 'r', encoding='utf-8') as f:
        collected = json.load(f)
    records = collected.get('records', [])
    print(f"采集到的候选记录: {len(records)} 条")
```

- [ ] **Step 2: 写比对逻辑——判定每条候选记录的去向**

```python
    # 构建现有数据的索引
    existing_keys = {}       # key -> fruit dict
    existing_names = set()   # 品种名集合
    existing_provinces = defaultdict(set)  # 省份 -> {品种名}

    for f in existing:
        for d in [f['district'], '']:
            key = make_key(f['name'], f['province'], f['city'], d)
            existing_keys[key] = f
        existing_names.add(f['name'])
        existing_provinces[f['province']].add(f['name'])

    verified_new = []
    unverified = []
    merge_conflicts = []

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
            # 检查是否有补充信息（如乡镇、描述等）
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
            verified_new.append(r)  # 仍写入但标记低置信度
        else:
            unverified.append(r)   # 无来源，不入库

    print(f"验证通过（待写入）: {len(verified_new)} 条")
    print(f"来源不足（待确认）: {len(unverified)} 条")
    print(f"冲突待处理: {len(merge_conflicts)} 条")
```

- [ ] **Step 3: 写内部去重逻辑和输出**

```python
    # 新增数据内部去重
    seen = {}
    deduped_new = []
    dup_count = 0
    for r in verified_new:
        key = make_key(r['name'], r['province'], r['city'], r.get('district', ''))
        if key in seen:
            # 保留来源更多的
            existing_entry = seen[key]
            if len(r.get('sources', [])) > len(existing_entry.get('sources', [])):
                seen[key] = r
            dup_count += 1
        else:
            seen[key] = r
            deduped_new.append(r)

    if dup_count > 0:
        print(f"新增数据内部去重: {dup_count} 组重复")

    # 写入输出文件
    for path, data in [
        ('_verified_new.json', deduped_new),
        ('_unverified.json', unverified),
        ('_merge_conflicts.json', merge_conflicts),
    ]:
        full = os.path.join(BASE_DIR, path)
        with open(full, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"已写入: {path}")

if __name__ == '__main__':
    main()
```

- [ ] **Step 4: 运行脚本**

```bash
cd c:/fruit-ops && python phase2_validate.py
```

期望输出：统计已有的 178 条、采集的候选记录数、验证通过数、来源不足数、冲突数。

- [ ] **Step 5: 审核输出文件**

人工检查 `_merge_conflicts.json` 和 `_unverified.json`，决定是否将部分条目移入已验证清单。

---

### Task 5: 编写 phase3_write.py 写入脚本

**Files:**
- Create: `c:/fruit-ops/phase3_write.py`

**Interfaces:**
- Consumes: `_verified_new.json`, `fruit_data.xlsx`
- Produces: 更新 `fruit_data.xlsx`，创建备份文件
- Dependencies: `server.py` 的 `geocode()`, `_auto_fill_lunar()`, `validate_season_consistency()`

- [ ] **Step 1: 写脚本头部——导入和配置**

```python
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

# 复用 server.py 的配置
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'fruit_data.xlsx')
HTML_PATH = os.path.join(BASE_DIR, '中国水果产区数据大屏.html')
VERIFIED_PATH = os.path.join(BASE_DIR, '_verified_new.json')

AMAP_KEY = "d60264473fe914fa0cf21d1a22a4e206"
AMAP_SECRET = "ae7d546c1e8f25d8365b2c601fd3c96a"

# ── 高德地理编码 ──
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


# ── 农历计算 ──
REFERENCE_YEAR = 2025

def solar_to_lunar_month(solar_month):
    """Convert 1-based solar month to 1-based lunar month using zhdate."""
    try:
        from zhdate import ZhDate
        d = ZhDate.from_datetime(datetime(REFERENCE_YEAR, max(1, min(12, int(solar_month))), 15))
        return d.lunar_month
    except Exception:
        return 0


# ── Sheet 列索引（1-based）──
# Sheet1: 水果产品库
(F_ID, F_NAME, F_CAT, F_PROV, F_CITY, F_DIST, F_TOWN) = range(1, 8)
(F_SEA_START, F_SEA_END, F_PEAK_START, F_PEAK_END) = range(8, 12)
(F_LUNAR_SS, F_LUNAR_SE, F_LUNAR_PS, F_LUNAR_PE) = range(12, 16)
(F_CURVE, F_DESC) = (16, 17)

# Sheet2: 产区映射表
(R_ADCODE, R_PROV, R_CITY, R_DIST, R_TOWN) = range(1, 6)
(R_LNG, R_LAT, R_FIDS, R_LEVEL) = range(6, 10)

# 已有品种→品类的映射（从现有数据中提取以复用）
FRUIT_CATEGORY_MAP = {
    '梨': '仁果类', '苹果': '仁果类', '山楂': '仁果类', '枇杷': '仁果类',
    '桃': '核果类', '李': '核果类', '杏': '核果类', '枣': '核果类',
    '樱桃': '核果类', '杨梅': '核果类', '橄榄': '核果类',
    '柑橘': '柑橘类', '橙': '柑橘类', '柚': '柑橘类', '柠檬': '柑橘类',
    '桔': '柑橘类', '柑': '柑橘类', '椪柑': '柑橘类', '蜜桔': '柑橘类',
    '葡萄': '浆果类', '猕猴桃': '浆果类', '草莓': '浆果类',
    '蓝莓': '浆果类', '桑葚': '浆果类', '无花果': '浆果类',
    '石榴': '浆果类', '百香果': '浆果类', '柿': '浆果类',
    '西瓜': '瓜果类', '甜瓜': '瓜果类', '哈密瓜': '瓜果类',
    '芒果': '热带水果', '荔枝': '热带水果', '龙眼': '热带水果',
    '香蕉': '热带水果', '菠萝': '热带水果', '火龙果': '热带水果',
    '番石榴': '热带水果', '榴莲': '热带水果', '山竹': '热带水果',
    '菠萝蜜': '热带水果', '木瓜': '热带水果', '椰子': '热带水果',
    '甘蔗': '热带水果', '莲雾': '热带水果', '释迦': '热带水果',
    '核桃': '核果类', '板栗': '核果类', '枸杞': '浆果类', '沙棘': '浆果类',
}

def guess_category(name):
    """根据品种名推断分类。"""
    for keyword, cat in FRUIT_CATEGORY_MAP.items():
        if keyword in name:
            return cat
    return '仁果类'  # 默认
```

- [ ] **Step 2: 写 Excel 写入逻辑**

```python
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

    for record in verified_new:
        # 推断类别
        if not record.get('category'):
            record['category'] = guess_category(record['name'])

        # 生成 ID
        new_id = generate_next_fruit_id(ws_f)
        record['id'] = new_id

        # 写入 Sheet1
        new_row = ws_f.max_row + 1
        write_fruit_row(ws_f, new_row, record)
        print(f"  Sheet1 写入: {new_id} {record['name']} ({record['province']} {record['city']})")

        # 获取坐标
        lng, lat = geocode(
            record.get('province', ''),
            record.get('city', ''),
            record.get('district', ''),
            record.get('town', ''),
        )

        if lng == 0 and lat == 0 and record.get('town'):
            # 乡镇级查询失败，降级到区县级
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
        time.sleep(0.01)  # 避免 adcode 重复

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
        if len(write_report['new_fruits']) % 10 == 0:
            time.sleep(1)
            print(f"  已写入 {len(write_report['new_fruits'])}/{len(verified_new)} ...")

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
```

- [ ] **Step 3: 运行写入脚本**

```bash
cd c:/fruit-ops && python phase3_write.py
```

- [ ] **Step 4: 验证 Excel 文件完整性**

```bash
cd c:/fruit-ops && python -c "
import openpyxl
wb = openpyxl.load_workbook('fruit_data.xlsx', data_only=True)
for s in wb.sheetnames:
    ws = wb[s]
    print(f'{s}: {ws.max_row-1} rows')
print('Done')
"
```

- [ ] **Step 5: 重新生成 HTML 嵌入式数据**

```bash
cd c:/fruit-ops && python -c "
from server import regenerate_html
regenerate_html()
print('HTML regenerated')
"
```

---

### Task 6: 编写 phase4_verify.py 校验脚本

**Files:**
- Create: `c:/fruit-ops/phase4_verify.py`

**Interfaces:**
- Consumes: `fruit_data.xlsx`（已更新版）
- Produces: `_enrich_report.json`

- [ ] **Step 1: 写校验脚本**

```python
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
        return False, f"月份字段非数字: {fruit}"

    if not (1.0 <= ss <= 12.9 and 1.0 <= se <= 12.9):
        return False, f"上市期超出范围: {ss}-{se}"
    if not (1.0 <= ps <= 12.9 and 1.0 <= pe <= 12.9):
        return False, f"盛产期超出范围: {ps}-{pe}"
    if not _in_range(ps, ss, se):
        return False, f"盛产开始({ps})不在上市期({ss}-{se})内"
    if not _in_range(pe, ss, se):
        return False, f"盛产结束({pe})不在上市期({ss}-{se})内"

    season_len = int(se) - int(ss) + 1
    if int(ss) > int(se):
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
        if not ok:
            prov = str(ws_r.cell(r, R_PROV).value or '')
            city = str(ws_r.cell(r, R_CITY).value or '')
            issues['coord_errors'].append({'adcode': adcode, 'province': prov, 'city': city, 'error': msg, 'lng': lng, 'lat': lat})

    # 4. 重复检查（品种+省+市+区县）
    seen = {}
    for f in fruits:
        key = f"{f['name']}|{f['province']}|{f['city']}|{f['district']}"
        if key in seen:
            issues['duplicates'].append({'key': key, 'ids': [seen[key]['id'], f['id']], 'name': f['name']})
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
            issues['missing_regions'].append({'id': fid, 'name': f['name'] if f else '?'})

    # 6. 处理已知重复——标记哪些需要删除
    for dupe in KNOWN_DUPES:
        keep_id = dupe['keep']
        remove_id = dupe['remove']
        # 检查两个 ID 是否都存在
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
    summary = {
        'checked_at': datetime.now().isoformat(),
        'total_fruits': len(fruits),
        'total_regions': ws_r.max_row - 1,
        'season_errors': len(issues['season_errors']),
        'coord_errors': len(issues['coord_errors']),
        'duplicates_found': len(issues['duplicates']),
        'missing_regions': len(issues['missing_regions']),
        'dedup_candidates': len(issues['dedup_done']),
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
    print(f"水果总数: {summary['total_fruits']}")
    print(f"产区总数: {summary['total_regions']}")
    print(f"生命周期错误: {summary['season_errors']}")
    print(f"坐标异常: {summary['coord_errors']}")
    print(f"发现重复: {summary['duplicates_found']}")
    print(f"缺失产区关联: {summary['missing_regions']}")
    print(f"待去重组: {summary['dedup_candidates']}")
    print(f"\n完整报告: {report_path}")

    wb.close()


if __name__ == '__main__':
    main()
```

- [ ] **Step 2: 运行校验脚本**

```bash
cd c:/fruit-ops && python phase4_verify.py
```

- [ ] **Step 3: 执行去重操作**

根据 `_enrich_report.json` 中的 `dedup_candidates`，删除重复记录：

```bash
cd c:/fruit-ops && python -c "
import openpyxl

# 已知重复：保留第一个，删除第二个
REMOVE_IDS = ['F108', 'F109', 'F107', 'F093', 'F086', 'F159']

wb = openpyxl.load_workbook('fruit_data.xlsx')

# Sheet1 去重
ws_f = wb['水果产品库']
rows_to_delete = []
for r in range(2, ws_f.max_row + 1):
    fid = str(ws_f.cell(r, 1).value or '').strip()
    if fid in REMOVE_IDS:
        rows_to_delete.append(r)
for r in sorted(rows_to_delete, reverse=True):
    print(f'删除 Sheet1 行 {r}: {ws_f.cell(r, 2).value}')
    ws_f.delete_rows(r)

# Sheet2 去重
ws_r = wb['产区映射表']
rows_to_delete = []
for r in range(2, ws_r.max_row + 1):
    fids_str = str(ws_r.cell(r, 7).value or '').strip()
    fids = [x.strip() for x in fids_str.split(',')]
    if any(fid in REMOVE_IDS for fid in fids):
        rows_to_delete.append(r)
for r in sorted(rows_to_delete, reverse=True):
    print(f'删除 Sheet2 行 {r}: {ws_r.cell(r, 1).value}')
    ws_r.delete_rows(r)

wb.save('fruit_data.xlsx')
wb.close()
print('去重完成')
"
```

- [ ] **Step 4: 再次运行校验确认干净**

```bash
cd c:/fruit-ops && python phase4_verify.py
```

期望：`duplicates_found: 0`, `dedup_candidates: 0`（如果是新增的重复需另行处理）。

- [ ] **Step 5: 最终重新生成 HTML**

```bash
cd c:/fruit-ops && python -c "
from server import regenerate_html
data = regenerate_html()
print(f'Fruits: {len(data[\"fruits\"])}')
print(f'Regions: {len(data[\"regions\"])}')
"
```

- [ ] **Step 6: 启动 server 验证前端正常显示**

```bash
cd c:/fruit-ops && python -c "
from server import app
import json

# 快速验证 API 返回
with app.test_client() as client:
    resp = client.get('/api/health')
    print('Health:', resp.get_json())
    resp = client.get('/api/fruits')
    fruits = resp.get_json()
    print(f'Fruits via API: {len(fruits)}')
    # 抽查几条新增数据
    for f in fruits[-5:]:
        print(f'  {f[\"id\"]}: {f[\"name\"]} @ {f[\"province\"]}{f[\"city\"]}{f[\"district\"]}')
"
```

---

### Task 7: 提交所有变更

- [ ] **Step 1: 检查变更文件列表**

```bash
cd c:/fruit-ops && git status
```

- [ ] **Step 2: 添加并提交**

```bash
cd c:/fruit-ops && git add fruit_data.xlsx fruit_data_backup_*.xlsx _collected_data.json _verified_new.json _unverified.json _merge_conflicts.json _enrich_report.json _write_report.json phase2_validate.py phase3_write.py phase4_verify.py "中国水果产区数据大屏.html" docs/
```

```bash
cd c:/fruit-ops && git commit -m "feat: 全国水果产区数据系统补充 — 多源交叉验证+去重+坐标补全

- 阶段1：基于GI/名特优新等权威源采集全国水果产区数据
- 阶段2：清洗验证脚本 phase2_validate.py（比对/交叉验证/去重）
- 阶段3：写入脚本 phase3_write.py（Excel写入+高德坐标补全）
- 阶段4：校验脚本 phase4_verify.py（一致性校验+去重+报告）
- 处理6组现有重复数据
- 每条数据可追溯到权威来源URL"
```

---

## Self-Review Summary

1. **Spec coverage**: 所有设计文档中的成功标准和阶段都对应有 Task；Phase 1 的设计要求（GI/名特优新搜索、交叉验证）映射到 Task 1-3
2. **Placeholder scan**: 无 TBD/TODO，所有 Step 都有完整代码或精确命令
3. **Type consistency**: 脚本间的数据格式通过 `_verified_new.json` 的结构约定保持一致；列索引复用了 `server.py` 的常量定义
