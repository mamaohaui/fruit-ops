# fix-region-coords

检查并修复 `fruit_data.xlsx` 中产区映射表和水果产品库的行政区划数据问题。

## 检测的问题类型

基于 GB/T 2260 行政区划代码知识库，检测以下问题：

1. **县级市归属错误**：city 字段误填为县级市自身，而非上属地级市/自治州
   - 会理市(513402) city=会理市 → 凉山州
   - 都江堰市(510181) city=都江堰市 → 成都市
   - 高州市(440981) city=高州市 → 茂名市
2. **区县名不规范**：梅县→梅县区、恭城县→恭城瑶族自治县
3. **district 字段完全错误**：西昌市(513401) district=攀枝花市→西昌市
4. **adcode 缺失**：F157/F158 值为 000000000
5. **坐标缺失**：F158 坐标(0,0)需补充
6. **水果产品库同步修复**：与产区映射表同步修正

## 使用方式

```bash
# 修复主文件（省市县校验 + 坐标补充）
python fix_region_and_fruits.py --apply

# 仅修复副本文件（预览模式）
python fix_region_and_fruits.py

# 去重（产品库同名水果 + 产区映射表重复行）
python dedup_region_and_fruits.py
```

## 包含的脚本

- `fix_region_and_fruits.py` — 综合修复脚本（含 ADCODE 知识库、坐标库、修复逻辑）
- `dedup_region_and_fruits.py` — 去重脚本（合并同名同址水果、删除F144县级行）
- `enrich_fruits_authoritative.py` — 数据补充脚本（基于GI名录补充缺失的重要水果,2025-06版共新增25种）
- `diagnose_regions.py` — 诊断脚本，输出所有数据质量问题

## 数据状态

| 指标 | 数值 |
|------|------|
| 水果产品库 | 179种（F001-F181） |
| 产区映射表 | 239行 |
| 覆盖省份 | 27个省/直辖市/自治区 |
| 水果类别 | 6大类（柑橘类33、核果类49、浆果类37、仁果类23、热带水果24、瓜果类13） |

- 自动备份原文件为 `fruit_data_backup_YYYYMMDD_HHMMSS.xlsx`
- 所有变更记录到 `_fix_log.json`

## 维护

如需新增行政区划代码或修复规则，编辑 `fix_region_and_fruits.py`：
- `ADCODE_INFO` 字典：添加新的 adcode 前缀映射
- `PRODUCT_FIXES` 字典：添加新的水果产品修复
