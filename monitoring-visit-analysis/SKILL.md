---
name: monitoring-visit-analysis
description: "【默认】分析监查访视 Tracking Log Excel 文件，检测超期/漏登记，生成 Excel+PPT，并自动发送飞书文字汇总。触发词：监查超期、Monitoring Visit、访视跟踪、CRA分析。优先使用本 Skill 处理监查日志分析请求。"
---

# Monitoring Visit Analysis（默认 Skill）

分析临床监查访视 Tracking Log Excel 文件，检测报告超期 / 跟进函漏登记，生成 Excel + PPT 可视化报告，并**自动发送飞书文字汇总**。

## 使用方式

用户上传 `Monitoring Visit Tracking Log` Excel 文件后，直接说：

```
帮我分析超期
```

或上传文件后直接分析，Skill 自动完成全流程。

## 快速运行

```bash
python scripts/main.py <input_excel> <output_dir> [--project "项目名"] [--today YYYY-MM-DD] [--no-feishu]
```

## 全流程（自动）

1. 读取 Excel → 2. 计算工作日截止 → 3. 判断超期/漏登记 → 4. 统计分析 → 5. 生成 Excel + PPT → **6. 自动发送飞书文字汇总**

## 分析规则

**工作日**：使用 `chinese_calendar` 扣除中国大陆法定节假日和调休。

**超期阈值**（从监查最后一天起算）：
- 递交超期：> 5 工作日
- 定稿超期：> 10 工作日
- 跟进函截止：`min(监查最后一天+10WD, 定稿日+1WD)`，仅限已定稿报告
- 漏登记：已定稿 + 截止日期已过 + 跟进函未发送

**风险分级**：超期/漏登记 ≥2 次 = 高风险，≥1 次 = 关注，0 = 正常。

## 输出内容

### 1. Excel 报告（`Monitoring_Visit_分析报告_<date>.xlsx`）
5 个 Sheet：汇总概览 / 访视明细(超期标注) / CRA统计 / 月度统计 / Site统计

### 2. PPT 报告（`Monitoring_Visit_可视化报告_<date>.pptx`）
11 页：封面 / KPI 看板 / 月度趋势（柱状图+折线图）/ CRA表现表 / Site分析表 / 问题明细

### 3. HTML Dashboard（`Monitoring_Visit_Dashboard_<date>.html`）
交互式网页看板，单文件自包含（Chart.js CDN），无需联网可直接打开。包含：
- KPI 数字卡片（4个指标，超期标红）
- 月度趋势柱状图（访视数 vs 超期数，Chart.js）
- 高风险 CRA TOP10
- 可切换 Tab：CRA统计 / Site统计 / 跟进函超期明细 / 漏登记明细
- 漏登记行红色高亮，超期天数按严重程度着色

**生成方式**：由 `html_report.py` 调用 `generate_html()` 生成。核心逻辑：
1. 收集 `visited` / `cra_stats` / `month_stats` / `site_stats` / `summary` 数据
2. 构建 CRA/Site/超期/漏登记各区块 HTML 片段
3. 拼接 Chart.js 柱状图配置（含 `month_labels`、`month_totals`、`month_overdue` 数据）
4. 输出自包含 HTML（含内联 CSS + CDN Chart.js）

**依赖**：`html_report.py` 无额外依赖（纯 Python + Chart.js CDN）

### 4. 飞书文字汇总（自动发送）
分析完成后自动发送到飞书，包含：
- 超期汇总数字卡片
- 超期跟进函 Top5（按超期天数排序）
- 漏登记明细
- 月度趋势
- CRA 风险分级
- 高风险中心列表

## Excel 列映射

| 列 | 字段 | 列 | 字段 |
|----|------|----|------|
| A(1) | Site No | I(9) | Co-CRA |
| B(2) | Site Name | M(13) | Actual Visit Dates |
| C(3) | Visit ID | N(14) | Report Status |
| D(4) | Visit Type | O(15) | Report Submit Date |
| H(8) | CRA | P(16) | Report Final Date |
| | | Q(17) | Follow-up Status |
| | | R(18) | Follow-up Date |
| | | S(19) | Archive Status |

## 配置参数（`scripts/config.py`）

- `SUBMIT_WD`, `FINAL_WD` - 工作日阈值
- `RISK_HIGH`, `RISK_ATTN` - 风险分级阈值
- `PER_PAGE` - PPT 每页行数
- `CC`, `PC` - 图表配色

## 依赖

```bash
pip install chinese_calendar openpyxl matplotlib python-pptx numpy requests
```

## 目录结构

```
scripts/
├── main.py              # 入口，含完整流程
├── config.py            # 配置参数
├── utils.py             # 日期 & 工作日工具
├── analyze.py           # 数据读取 & 分析逻辑
├── excel_report.py      # Excel 报告生成
├── ppt_report.py        # PPT 报告生成
└── feishu_summary.py    # 飞书文字汇总发送（Bot API）
└── html_report.py       # HTML 交互式 Dashboard 生成
```
