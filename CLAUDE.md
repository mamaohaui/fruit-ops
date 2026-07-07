# CLAUDE.md — 中国水果产区数据大屏

## 项目概述

中国水果产区数据大屏 —— 以ECharts交互式地图为核心，展示全国水果产区的数据可视化系统。

## 技术栈

| 层 | 技术 |
|---|---|
| 后端 API | Python 3 + Flask (`server.py`) |
| 数据存储 | openpyxl (Excel `.xlsx` 作为数据库) |
| 前端大屏 | ECharts + 原生 HTML/JS |
| 数据生成/ETL | Python 脚本（openpyxl） |

## 项目结构

```
C:\fruit-ops\
├── CLAUDE.md                       # 本文件
├── reasonix.toml                   # AI 权限配置
├── server.py                       # Flask REST API（CRUD on Excel）
├── fruit_data.xlsx                 # 数据存储（3个Sheet）
├── 中国水果产区数据大屏.html        # ECharts 交互式地图大屏
├── generate_fruit_data.py          # 从权威来源生成水果数据
├── update_fruit_data.py            # 数据质量修复 + 新增50种水果
├── enrich_fruit_data.py            # 增加西瓜、菠萝、樱桃等品种
├── apply_season_colors.py          # 添加基于季节的渐变配色
├── migrate_add_lunar.py            # 添加4列农历日期迁移脚本
├── _data_summary.json              # 省份/城市统计摘要
└── _sheet_names_check.txt          # Sheet名称校验记录
```

## 架构

```
Flask server.py  ◄─── HTTP REST API ───► 中国水果产区数据大屏.html
      │
      ▼
openpyxl ──► fruit_data.xlsx（3个Sheet）
      │
      ▼ (重新生成嵌入式 JSON)
  中国水果产区数据大屏.html (ECharts 地图)
```

## AI 行为准则

### 1. 先思考再编码
不假设。不隐藏疑惑。摆出权衡。
实施前明确假设，如有多种解释则提出，存在更简单方案时指出来。

### 2. 简单至上
最小化代码解决问题。不添加未要求特性，不为一次性代码做抽象。

### 3. 精确改动
只动必须要动的。不"改进"相邻代码，匹配现有风格，清理自己产生的孤儿。

### 4. 目标驱动执行
定义成功标准，循环直到验证。多步骤任务先写计划，每步附带验证方式。
