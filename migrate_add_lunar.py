#!/usr/bin/env python3
"""
One-time migration: Add 4 lunar calendar columns to fruit_data.xlsx
- Inserts columns after PeakEnd (col 11), before Curve (col 12->16)
- New cols: LunarSeasonStart(12), LunarSeasonEnd(13), LunarPeakStart(14), LunarPeakEnd(15)
- Curve shifts from 12 to 16, Desc shifts from 13 to 17
- Copies Gregorian season values as initial lunar values
- Regenerates HTML embedded data with 17-column format
"""

import openpyxl
import json
import os
import re

EXCEL_PATH = 'fruit_data.xlsx'
HTML_PATH = '中国水果产区数据大屏.html'

wb = openpyxl.load_workbook(EXCEL_PATH)
ws = wb['水果产品库']

# Current layout:
# 1:ID, 2:Name, 3:Category, 4:Province, 5:City, 6:District, 7:Town
# 8:SeasonStart, 9:SeasonEnd, 10:PeakStart, 11:PeakEnd
# 12:Curve, 13:Desc

# New layout:
# 1:ID, 2:Name, 3:Category, 4:Province, 5:City, 6:District, 7:Town
# 8:SeasonStart, 9:SeasonEnd, 10:PeakStart, 11:PeakEnd
# 12:LunarSeasonStart, 13:LunarSeasonEnd, 14:LunarPeakStart, 15:LunarPeakEnd
# 16:Curve, 17:Desc

# Step 1: Insert 4 new columns at position 12 (before existing col 12 = Curve)
# openpyxl insert_cols shifts existing columns right
ws.insert_cols(12, 4)

# Step 2: Set headers for new columns
headers = ['农历产季开始月', '农历产季结束月', '农历盛产开始月', '农历盛产结束月']
for i, h in enumerate(headers):
    ws.cell(1, 12 + i).value = h

# Step 3: For each fruit row, copy Gregorian values as initial lunar values
# Also add the D (desc) column header if missing
copied = 0
for r in range(2, ws.max_row + 1):
    fid = ws.cell(r, 1).value
    if not fid or not str(fid).strip():
        continue
    # Copy Gregorian season values to lunar
    for offset in range(4):
        src_col = 8 + offset  # SeasonStart(8), SeasonEnd(9), PeakStart(10), PeakEnd(11)
        dst_col = 12 + offset  # Lunar equivalents
        src_val = ws.cell(r, src_col).value
        if src_val is not None:
            ws.cell(r, dst_col).value = src_val
    copied += 1

print(f'Copied lunar values for {copied} fruits')

# Save
wb.save(EXCEL_PATH)
print(f'Saved {EXCEL_PATH} with 17-column schema')

# Now regenerate HTML embedded data
wb2 = openpyxl.load_workbook(EXCEL_PATH)
ws_f = wb2['水果产品库']
ws_r = wb2['产区映射表']
ws_c = wb2['成熟度曲线']

# Build fruits array (17 columns)
fruits = []
for r in range(2, ws_f.max_row + 1):
    fid = str(ws_f.cell(r, 1).value or '').strip()
    if not fid:
        continue
    row = [
        fid,
        str(ws_f.cell(r, 2).value or '').strip(),
        str(ws_f.cell(r, 3).value or '').strip(),
        str(ws_f.cell(r, 4).value or '').strip(),
        str(ws_f.cell(r, 5).value or '').strip(),
        str(ws_f.cell(r, 6).value or '').strip(),
        str(ws_f.cell(r, 7).value or '').strip(),
        str(ws_f.cell(r, 8).value or '').strip(),
        str(ws_f.cell(r, 9).value or '').strip(),
        str(ws_f.cell(r, 10).value or '').strip(),
        str(ws_f.cell(r, 11).value or '').strip(),
        str(ws_f.cell(r, 12).value or '').strip(),   # lunarSeasonStart
        str(ws_f.cell(r, 13).value or '').strip(),   # lunarSeasonEnd
        str(ws_f.cell(r, 14).value or '').strip(),   # lunarPeakStart
        str(ws_f.cell(r, 15).value or '').strip(),   # lunarPeakEnd
        str(ws_f.cell(r, 16).value or '').strip(),   # curve (was col 12)
        str(ws_f.cell(r, 17).value or '').strip(),   # desc (was col 13)
    ]
    fruits.append(row)

# Build regions array (unchanged, 9 columns)
regions = []
for r in range(2, ws_r.max_row + 1):
    adcode = str(ws_r.cell(r, 1).value or '').strip()
    if not adcode:
        continue
    row = [
        adcode,
        str(ws_r.cell(r, 2).value or '').strip(),
        str(ws_r.cell(r, 3).value or '').strip(),
        str(ws_r.cell(r, 4).value or '').strip(),
        str(ws_r.cell(r, 5).value or '').strip(),
        str(ws_r.cell(r, 6).value or '').strip(),
        str(ws_r.cell(r, 7).value or '').strip(),
        str(ws_r.cell(r, 8).value or '').strip(),
        str(ws_r.cell(r, 9).value or '').strip(),
    ]
    regions.append(row)

# Build curves array (unchanged, 13 columns)
curves = []
for r in range(2, ws_c.max_row + 1):
    name = str(ws_c.cell(r, 1).value or '').strip()
    if not name:
        continue
    vals = [name]
    for c in range(2, 14):
        vals.append(str(ws_c.cell(r, c).value or '0'))
    curves.append(vals)

embedded = json.dumps({"fruits": fruits, "regions": regions, "curves": curves}, ensure_ascii=False)

# Update HTML file
with open(HTML_PATH, 'r', encoding='utf-8') as f:
    html = f.read()

# Replace EMBEDDED_DATA content
pattern = r'var EMBEDDED_DATA = \{.*?\};'
replacement = 'var EMBEDDED_DATA = ' + embedded + ';'
html_new = re.sub(pattern, replacement, html, flags=re.DOTALL)

if html_new != html:
    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(html_new)
    print(f'Updated embedded data in {HTML_PATH}')
else:
    print('WARNING: Could not find EMBEDDED_DATA pattern in HTML')

# Verify
brace_count = html_new.count('{') - html_new.count('}')
print(f'Brace balance: {brace_count} (0 = balanced)')
print(f'Fruits: {len(fruits)}, Regions: {len(regions)}, Curves: {len(curves)}')

wb2.close()
wb.close()
print('Migration complete!')
