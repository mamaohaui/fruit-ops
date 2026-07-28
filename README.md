# 中国水果产区数据大屏 · China Fruit Production Dashboard

[![Python](https://img.shields.io/badge/Python-3.8+-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Flask](https://img.shields.io/badge/Flask-2.0+-000000?logo=flask&logoColor=white)](https://flask.palletsprojects.com/)
[![ECharts](https://img.shields.io/badge/ECharts-5.5-AA344D?logo=apacheecharts&logoColor=white)](https://echarts.apache.org/)
[![License](https://img.shields.io/badge/License-MIT-green)](./LICENSE)

交互式中国地图数据大屏，展示全国水果产区的时空分布。支持国家级 → 省级 → 市级三级下钻，按生长阶段（生长期/成熟期/盛产期/尾产期）动态筛选散点，配套 Flask REST API 实现 Excel 数据持久化。

An interactive dashboard visualizing fruit production regions across China. Features a three-level drill-down map (nation → province → city), dynamic scatter filtering by growth phase, and a Flask REST API backed by Excel-based storage.

---

## ✨ 功能特性 · Features

| 特性 | Feature | 说明 |
|---|---|---|
| 🗺️ 三级地图下钻 | Three-level Drill-down | 全国 → 省份 → 城市，点击即可深入 |
| 🌱 生长阶段筛选 | Growth Phase Filter | 生长期/成熟期/盛产期/尾产期，一键切换散点可见性 |
| 📅 月度轮播 | Month Slider | 1–12 月动态展示应季水果分布 |
| 📍 农批市场图层 | Wholesale Market Layer | 可叠加显示全国农批市场分布 |
| 📊 右侧详情面板 | Detail Side Panel | 点击产区/城市查看水果列表、生命周期条、产区统计 |
| ✏️ 内置编辑功能 | Inline CRUD | 无需切换工具，大屏内直接增删改水果/产区数据 |
| 🔄 缩放状态保持 | Zoom State Preservation | 阶段/月份切换时保留用户当前的缩放/平移位置 |
| 🌐 地理编码自动补全 | Auto Geocoding | 新建产区时调用高德 API 自动填充经纬度 |
| 📡 RESTful API | REST API | 完整的水果/产区/曲线 CRUD 接口 |

---

## 🖼️ 界面预览 · Screenshot

```
┌─────────────────────────────────────────────────────────┐
│  🍎 中国水果产区数据大屏  [月份▼]  [🔄刷新]  水果:203 │
├──────────────────────────────┬──────────────────────────┤
│                              │  产区详情                  │
│     🇨🇳 ECharts 中国地图     │  ┌────────────────────┐  │
│                              │  │ 📍 海南省            │  │
│   🔴 盛产期  🟡 成熟期      │  │ 核心产区 | 水果 4 种  │  │
│   🟢 生长期  🟣 尾产期      │  │ ▶ 海口市 3 种水果   │  │
│                              │  │ ▶ 三亚市 1 种水果   │  │
│   [散点分布于各省/市/县]     │  └────────────────────┘  │
│                              │                           │
│                              │  水果列表                  │
│                              │  🥭 芒果 | 热带水果       │
│                              │  ▓▓▓▓▓▓▓░░░ 盛产期       │
│                              │  🍍 菠萝 | 热带水果       │
│                              │  ▓▓▓▓░░░░░░ 成熟期       │
│                              │                           │
├──────────────────────────────┴──────────────────────────┤
│  🟢生长期  🟡成熟期  🔴盛产期  🟣尾产期  │ 🏪农批市场  │
└─────────────────────────────────────────────────────────┘
```

---

## 🏗️ 技术架构 · Architecture

```
┌─────────────┐     HTTP REST API      ┌──────────────────┐
│   Flask      │ ◄─────────────────── ► │  ECharts 交互地图  │
│  server.py   │     JSON / CRUD       │  原生 HTML/JS/CSS  │
│  Port :5000  │                       │  (SPA 大屏)        │
└──────┬───────┘                       └──────────────────┘
       │
       ▼  openpyxl
┌────────────────┐
│ fruit_data.xlsx │
│ ┌────────────┐ │
│ │ 水果产品库  │ │  ← 13 列：ID/名称/分类/省市区镇/上市期/盛产期/曲线/描述
│ │ 产区映射表  │ │  ← 10 列：adcode/省市县镇/经纬度/关联水果/产区等级
│ │ 成熟度曲线  │ │  ← 13 列：12 个月成熟度分值
│ │ 农批市场    │ │  ← 11 列：市场名称/省市区/等级/经纬度
│ └────────────┘ │
└────────────────┘
```

**数据流：** 浏览器 ←→ Flask API ←→ Excel 文件。每次写操作后自动重新生成 HTML 内嵌 JSON，确保前端加载时数据始终最新。

**Data Flow:** Browser ←→ Flask API ←→ Excel file. After each write, the server regenerates the embedded JSON in the HTML file, keeping the frontend data always up-to-date.

---

## 📁 项目结构 · Project Structure

```
fruit-ops/
├── server.py                          # Flask REST API 服务（CRUD + 高德地理编码）
├── 中国水果产区数据大屏.html           # ECharts 交互式大屏（内嵌数据 + 全量前端逻辑）
├── fruit_data.xlsx                    # 数据存储（4 个 Sheet，~200+ 条水果记录）
├── generate_fruit_data.py             # 🔧 从权威来源批量生成初始水果数据
├── update_fruit_data.py               # 🔧 数据质量修复 + 品种扩充
├── enrich_fruit_data.py               # 🔧 增加西瓜、菠萝、樱桃等品种
├── apply_season_colors.py             # 🔧 根据上市季节分配渐变配色
├── CLAUDE.md                          # AI 协作准则（中文）
└── README.md                          # 本文件 · This file
```

---

## 🚀 快速启动 · Quick Start

### 环境要求 · Prerequisites

- **Python** ≥ 3.8
- **pip** 包管理器

### 安装与运行 · Installation

```bash
# 1. 克隆项目 · Clone
git clone <your-repo-url>
cd fruit-ops

# 2. 安装依赖 · Install dependencies
pip install flask openpyxl

# 3. 启动服务器 · Start the server
python server.py
```

服务器启动后访问：**http://localhost:5000**

Open your browser and navigate to **http://localhost:5000**. The dashboard loads data from `fruit_data.xlsx` automatically.

### 首次使用（无数据时）· First-time Setup

```bash
# 生成初始水果数据集 · Generate initial fruit dataset
python generate_fruit_data.py

# 补充更多品种 · Enrich with more varieties
python enrich_fruit_data.py

# 应用季节性配色 · Apply seasonal color scheme
python apply_season_colors.py
```

---

## 🔌 API 接口 · API Endpoints

| 方法 | 路径 | 说明 · Description |
|---|---|---|
| `GET` | `/` | 返回数据大屏页面（自动嵌入最新数据） |
| `GET` | `/api/health` | 健康检查 · Health check |
| `GET` | `/api/fruits` | 获取全部水果数据 · List all fruits |
| `GET` | `/api/fruits/:id` | 获取单个水果 · Get fruit by ID |
| `POST` | `/api/fruits` | 新增水果（自动地理编码 + 产区关联） |
| `PUT` | `/api/fruits/:id` | 更新水果（含生命周期校验） |
| `DELETE` | `/api/fruits/:id` | 删除水果及关联产区 · Delete fruit |
| `GET` | `/api/regions` | 获取全部产区 · List all regions |
| `POST` | `/api/regions` | 创建独立产区（自动补全坐标） |
| `GET` | `/api/curves` | 获取成熟度曲线 · List maturity curves |
| `GET` | `/api/markets` | 获取农批市场数据 · List wholesale markets |
| `GET` | `/api/excel` | 下载 Excel 数据文件 · Download Excel |
| `POST` | `/api/refresh` | 强制刷新 HTML 内嵌数据 |
| `POST` | `/api/sync-all` | 全量同步浏览器数据回 Excel |
| `POST` | `/api/run-script` | 执行数据生成脚本（白名单机制） |

---

## 🗺️ 地图交互 · Map Interaction

| 操作 · Action | 行为 · Behavior |
|---|---|
| 点击省份 · Click province | 下钻到省级视图 |
| 点击城市 · Click city | 继续下钻到市级视图 |
| 点击"返回上级" · Click back | 回到上一级地图 |
| 点击散点 · Click scatter point | 右侧面板显示水果详情 |
| 切换月份 · Switch month | 按上市月份动态筛选 |
| 点击阶段按钮 · Toggle phase | 切换生长期/成熟期/盛产期/尾产期可见性（保留缩放状态） |
| 点击"农批市场" · Toggle markets | 叠加显示全国农批市场位置 |
| 鼠标滚轮/双指 · Scroll/pinch | 缩放地图 |
| 拖拽 · Drag | 平移地图 |

---

## 📊 数据模型 · Data Model

### 水果产品库 · Fruits (Sheet 1)

| 字段 · Field | 说明 · Description |
|---|---|
| 水果ID · Fruit ID | 唯一标识，如 F001 |
| 水果名称 · Name | 如"海南芒果" |
| 类别 · Category | 热带水果 / 温带水果 / 浆果类 |
| 省/市/区/镇 · Location | 产地层级地址 |
| 上市开始月 / 结束月 · Season | 支持小数月份（如 4.5 = 4 月中旬） |
| 盛产开始月 / 结束月 · Peak | 必须在上市期内 |
| 成熟度曲线 · Curve | 关联"成熟度曲线"Sheet |
| 描述 · Description | 自由文本 |

### 产区映射表 · Regions (Sheet 2)

| 字段 · Field | 说明 · Description |
|---|---|
| 行政区划代码 · adcode | 统计局编码 |
| 省/市/区/镇 · Location | 层级地址 |
| 经度 / 纬度 · LNG/LAT | 地理坐标 |
| 关联水果ID · Fruit IDs | 逗号分隔，如 `F001,F002` |
| 产区等级 · Level | 核心产区 / 主要产区 / 一般产区 |

---

## 🛠️ 技术栈 · Tech Stack

| 层 · Layer | 技术 · Technology | 用途 · Purpose |
|---|---|---|
| 后端 API | Python 3 + Flask | REST 接口服务 |
| 数据持久化 | openpyxl (Excel .xlsx) | 以 Excel 为数据库 |
| 前端地图 | ECharts 5.5 | 中国地图 + geo 散点 |
| 前端框架 | 原生 HTML/JS/CSS | 零依赖 SPA（ECharts CDN 除外） |
| 地理编码 | 高德地图 Web API | 地址 → 经纬度自动补全 |
| 数据生成 | Python 脚本 | 批量构造水果/产区/曲线数据 |

---

## 🤝 贡献 · Contributing

本项目使用 AI 辅助开发（Claude Code），遵循 [CLAUDE.md](./CLAUDE.md) 中定义的协作准则。

欢迎提交 Issue 和 Pull Request。 · Issues and PRs are welcome.

---

## 📄 许可证 · License

MIT License

---

> 🤖 本项目由 AI 辅助生成与维护 · Built with [Claude Code](https://claude.ai/code)
