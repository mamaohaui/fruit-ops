---
name: fix-region-coords
description: 修复 fruit_data.xlsx 产区映射表的行政区代码、经纬度坐标和县级市归属问题
---

# fix-region-coords

修复中国水果产区数据 `fruit_data.xlsx` 中"产区映射表"sheet 的数据质量问题。

## 能力

| 修复项 | 说明 |
|--------|------|
| 县级市归属 city 字段 | 县级市的 city 从自身改为上属地级市（如会理市→凉山州） |
| 区县名称标准化 | 如"梅县"→"梅县区"，"恭城县"→"恭城瑶族自治县" |
| 行政区代码校验 | 基于内置 GB/T 2260 代码库校验并修复 adcode |
| 坐标缺失补充 | 基于内置坐标库补充缺失的经纬度 |

## 使用方式

```bash
# 预览变更（dry-run）
python skill_fix_region_coords.py

# 应用修复
python skill_fix_region_coords.py --apply
```

## 包含的脚本

- `skill_fix_region_coords.py` — 核心修复脚本（含坐标库、行政区代码库、修复逻辑）
- `diagnose_regions.py` — 诊断脚本，输出数据质量问题清单

## 安全机制

- 自动备份原文件为 `fruit_data_backup_YYYYMMDD.xlsx`
- dry-run 模式先预览变更
- 所有变更记录到 `_fix_log.json`

## 数据来源

- 坐标数据：基于国家标准地理信息，约 200+ 区县/城市的精确坐标
- 行政区划代码：基于 GB/T 2260 标准
- 县级市归属参考：民政部行政区划信息
