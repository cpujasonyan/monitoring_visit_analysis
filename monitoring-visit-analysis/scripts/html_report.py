# -*- coding: utf-8 -*-
"""Generate interactive HTML dashboard for monitoring visit analysis."""

import os
import json
from collections import defaultdict


def risk_color(risk):
    if risk == "高风险":
        return "#dc3545"
    if risk == "关注":
        return "#fd7e14"
    return "#28a745"


def risk_badge(risk):
    color = risk_color(risk)
    return f'<span style="background:{color};color:white;padding:2px 8px;border-radius:12px;font-size:12px">{risk}</span>'


def generate_html(visited, cra_stats, month_stats, site_stats, summary, output_path, today_str):
    """Generate self-contained HTML dashboard."""

    # Build CRA data
    cra_site_count = defaultdict(int)
    for r in visited:
        cra_site_count[r.get("cra", "")] += 1

    cra_rows = []
    for name, c in cra_stats.items():
        fin_od = c.get("fin_od", 0)
        fu_od = c.get("fu_od", 0)
        missing = c.get("missing", 0)
        total_od = fin_od + fu_od + missing
        risk = c.get("risk", "正常")
        sites = cra_site_count.get(name, 0)
        cra_rows.append({
            "cra": name, "total": c.get("total", 0),
            "fin_od": fin_od, "fu_od": fu_od, "missing": missing,
            "total_od": total_od, "risk": risk, "sites": sites
        })
    cra_rows.sort(key=lambda x: -x["total_od"])

    # Build site data
    site_rows = []
    for sid, s in site_stats.items():
        fin_od = s.get("fin_od", 0)
        fu_od = s.get("fu_od", 0)
        missing = s.get("missing", 0)
        total_od = fin_od + fu_od + missing
        risk = s.get("risk", "正常")
        site_rows.append({
            "site": sid, "name": s.get("name", ""),
            "total": s.get("total", 0),
            "fin_od": fin_od, "fu_od": fu_od, "missing": missing,
            "total_od": total_od, "risk": risk
        })
    site_rows.sort(key=lambda x: -x["total_od"])

    # Month data
    month_items = sorted(month_stats.items())
    month_labels = [m for m, _ in month_items]
    month_totals = [month_stats[m].get("total", 0) for m in month_labels]
    month_overdue = [
        month_stats[m].get("fin_od", 0) + month_stats[m].get("fu_od", 0) + month_stats[m].get("missing", 0)
        for m in month_labels
    ]

    # Top5 overdue follow-up
    fu_od_rows = [r for r in visited if r.get("followup_overdue") == "是"]
    fu_od_rows.sort(key=lambda r: r.get("followup_wd", 0), reverse=True)
    top5 = fu_od_rows[:5]

    # Missing registration
    missing_rows = [r for r in visited if r.get("missing_reg")]

    # Summary numbers
    sub_od = summary.get("total_sub_od", 0)
    fin_od = summary.get("total_fin_od", 0)
    fu_od = summary.get("total_fu_od", 0)
    missing = summary.get("total_missing", 0)
    finalized = summary.get("total_finalized", 0)
    visited_n = summary.get("total_visited", 0)

    date_range = ""
    all_ends = [str(r["visit_end"]) for r in visited if r.get("visit_end")]
    if all_ends:
        date_range = f"{min(all_ends)} ~ {max(all_ends)}"

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CRA监查超期分析看板</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: -apple-system, "Microsoft YaHei", "PingFang SC", sans-serif; background: #f0f2f5; color: #1f2329; font-size: 14px; }}
  .header {{ background: linear-gradient(135deg, #1a3a5c, #0d5f8a); color: white; padding: 24px 32px; }}
  .header h1 {{ font-size: 22px; font-weight: 600; }}
  .header p {{ margin-top: 6px; font-size: 13px; opacity: 0.85; }}
  .kpi-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; padding: 20px 32px; }}
  .kpi-card {{ background: white; border-radius: 12px; padding: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }}
  .kpi-label {{ font-size: 12px; color: #646a73; margin-bottom: 8px; }}
  .kpi-value {{ font-size: 32px; font-weight: 700; }}
  .kpi-sub {{ font-size: 11px; color: #94979c; margin-top: 4px; }}
  .kpi-ok {{ color: #28a745; }}
  .kpi-warn {{ color: #fd7e14; }}
  .kpi-danger {{ color: #dc3545; }}
  .section {{ padding: 0 32px 24px; }}
  .section-title {{ font-size: 16px; font-weight: 600; margin-bottom: 16px; padding-top: 20px; display: flex; align-items: center; gap: 8px; }}
  .card {{ background: white; border-radius: 12px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); overflow: hidden; }}
  .chart-wrap {{ padding: 20px; height: 300px; }}
  .cols-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 24px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  th {{ background: #f5f6f7; padding: 10px 12px; text-align: left; font-weight: 600; color: #646a73; border-bottom: 1px solid #e8e9eb; }}
  td {{ padding: 9px 12px; border-bottom: 1px solid #f0f1f2; }}
  tr:hover td {{ background: #f8f9fa; }}
  .badge {{ display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; }}
  .badge-red {{ background: #fff1f0; color: #dc3545; }}
  .badge-orange {{ background: #fff7e6; color: #fd7e14; }}
  .badge-green {{ background: #f6ffed; color: #52c41a; }}
  .tag {{ background: #e6f0ff; color: #1677ff; padding: 2px 6px; border-radius: 4px; font-size: 11px; }}
  .tab-bar {{ display: flex; gap: 4px; padding: 12px 20px 0; border-bottom: 1px solid #e8e9eb; }}
  .tab {{ padding: 8px 16px; cursor: pointer; border-radius: 8px 8px 0 0; font-size: 13px; color: #646a73; border: none; background: none; }}
  .tab.active {{ background: white; color: #1677ff; font-weight: 600; border: 1px solid #e8e9eb; border-bottom: 2px solid white; margin-bottom: -1px; }}
  .tab-content {{ display: none; }}
  .tab-content.active {{ display: block; }}
  .summary-bar {{ display: flex; gap: 24px; padding: 16px 20px; background: #fafafa; border-radius: 8px; margin-bottom: 12px; }}
  .summary-item {{ font-size: 12px; color: #646a73; }}
  .summary-item strong {{ color: #1f2329; }}
  .footer {{ text-align: center; padding: 20px; color: #94979c; font-size: 12px; }}
</style>
</head>
<body>

<div class="header">
  <h1>📊 CRA监查超期分析看板</h1>
  <p>分析范围：{date_range} &nbsp;|&nbsp; 共 {visited_n} 条已访视 / {finalized} 条已定稿 &nbsp;|&nbsp; 生成日期：{today_str}</p>
</div>

<div class="kpi-grid">
  <div class="kpi-card">
    <div class="kpi-label">报告递交超期</div>
    <div class="kpi-value {'kpi-ok' if sub_od == 0 else 'kpi-warn'}">{sub_od}</div>
    <div class="kpi-sub">{'✅ 全部按时递交' if sub_od == 0 else '⚠️ 存在超期'}</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">报告定稿超期</div>
    <div class="kpi-value {'kpi-ok' if fin_od == 0 else 'kpi-danger'}">{fin_od}</div>
    <div class="kpi-sub">{'✅ 全部按时定稿' if fin_od == 0 else '⚠️ 需跟进'}</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">跟进函超期</div>
    <div class="kpi-value {'kpi-ok' if fu_od == 0 else 'kpi-danger'}">{fu_od}</div>
    <div class="kpi-sub">{'✅ 全部按时发送' if fu_od == 0 else '🔴 需重点跟进'}</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">跟进函漏登记</div>
    <div class="kpi-value {'kpi-ok' if missing == 0 else 'kpi-danger'}">{missing}</div>
    <div class="kpi-sub">{'✅ 无漏登记' if missing == 0 else '⚠️ 需立即处理'}</div>
  </div>
</div>

<div class="cols-2 section">
  <div class="card">
    <div class="section-title">📈 月度趋势</div>
    <div class="chart-wrap"><canvas id="monthChart"></canvas></div>
  </div>
  <div class="card">
    <div class="section-title">🚩 高风险 CRA TOP10</div>
    <div style="padding: 12px 16px 0">
"""

    # Top 10 CRA table
    for i, c in enumerate(cra_rows[:10]):
        bg = "#fff1f0" if c["risk"] == "高风险" else ("#fff7e6" if c["risk"] == "关注" else "#f6ffed")
        color = "#dc3545" if c["risk"] == "高风险" else ("#fd7e14" if c["risk"] == "关注" else "#52c41a")
        html += f"""      <div style="display:flex;align-items:center;padding:7px 0;border-bottom:1px solid #f0f1f2;gap:10px">
        <span style="font-size:13px;font-weight:700;color:#646a73;min-width:18px">{i+1}</span>
        <span style="flex:1;font-size:13px">{c['cra']}</span>
        <span style="background:{bg};color:{color};padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600">{c['total_od']}次</span>
        <span style="font-size:11px;color:#94979c">{c['sites']}Site</span>
      </div>
"""

    html += """    </div>
  </div>
</div>

<div class="section">
  <div class="card">
    <div class="tab-bar">
      <button class="tab active" onclick="showTab('tab-cra')">CRA统计</button>
      <button class="tab" onclick="showTab('tab-site')">Site统计</button>
      <button class="tab" onclick="showTab('tab-fu')">跟进函超期明细</button>
      <button class="tab" onclick="showTab('tab-missing')">漏登记明细</button>
    </div>

    <div id="tab-cra" class="tab-content active">
      <div class="summary-bar">
        <div class="summary-item">CRA总数：<strong>""" + str(len(cra_stats)) + """</strong></div>
        <div class="summary-item">高风险：<strong style="color:#dc3545">""" + str(sum(1 for c in cra_rows if c["risk"] == "高风险")) + """</strong></div>
        <div class="summary-item">关注：<strong style="color:#fd7e14">""" + str(sum(1 for c in cra_rows if c["risk"] == "关注")) + """</strong></div>
        <div class="summary-item">正常：<strong style="color:#28a745">""" + str(sum(1 for c in cra_rows if c["risk"] == "正常")) + """</strong></div>
      </div>
      <table>
        <thead>
          <tr>
            <th>CRA</th>
            <th>访视数</th>
            <th>定稿超期</th>
            <th>跟进函超期</th>
            <th>漏登记</th>
            <th>总超期</th>
            <th>风险</th>
          </tr>
        </thead>
        <tbody>
"""

    for c in cra_rows:
        badge_cls = "badge-red" if c["risk"] == "高风险" else ("badge-orange" if c["risk"] == "关注" else "badge-green")
        html += f"""          <tr>
            <td><strong>{c['cra']}</strong></td>
            <td>{c['total']}</td>
            <td>{c['fin_od']}</td>
            <td>{c['fu_od']}</td>
            <td>{c['missing']}</td>
            <td><strong style="color:{'#dc3545' if c['total_od']>0 else '#28a745'}">{c['total_od']}</strong></td>
            <td><span class="badge {badge_cls}">{c['risk']}</span></td>
          </tr>
"""

    html += """        </tbody>
      </table>
    </div>

    <div id="tab-site" class="tab-content">
      <div class="summary-bar">
        <div class="summary-item">Site总数：<strong>""" + str(len(site_stats)) + """</strong></div>
        <div class="summary-item">高风险Site：<strong style="color:#dc3545">""" + str(sum(1 for s in site_rows if s["risk"] == "高风险")) + """</strong></div>
      </div>
      <table>
        <thead>
          <tr>
            <th>Site编号</th>
            <th>Site名称</th>
            <th>访视数</th>
            <th>定稿超期</th>
            <th>跟进函超期</th>
            <th>漏登记</th>
            <th>总超期</th>
            <th>风险</th>
          </tr>
        </thead>
        <tbody>
"""

    for s in site_rows:
        badge_cls = "badge-red" if s["risk"] == "高风险" else ("badge-orange" if s["risk"] == "关注" else "badge-green")
        html += f"""          <tr>
            <td><strong>{s['site']}</strong></td>
            <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="{s['name']}">{s['name']}</td>
            <td>{s['total']}</td>
            <td>{s['fin_od']}</td>
            <td>{s['fu_od']}</td>
            <td>{s['missing']}</td>
            <td><strong style="color:{'#dc3545' if s['total_od']>0 else '#28a745'}">{s['total_od']}</strong></td>
            <td><span class="badge {badge_cls}">{s['risk']}</span></td>
          </tr>
"""

    html += """        </tbody>
      </table>
    </div>

    <div id="tab-fu" class="tab-content">
      <table>
        <thead>
          <tr>
            <th>CRA</th>
            <th>Site</th>
            <th>监查日期</th>
            <th>截止日期</th>
            <th>实际发送日</th>
            <th>超期天数</th>
          </tr>
        </thead>
        <tbody>
"""

    for r in fu_od_rows[:30]:
        wd = r.get("followup_wd", 0)
        wd_color = "#dc3545" if wd >= 5 else ("#fd7e14" if wd >= 3 else "#646a73")
        html += f"""          <tr>
            <td>{r.get('cra','')}</td>
            <td style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="{r.get('site_name','')}">{r.get('site_name','')}</td>
            <td>{r.get('visit_end','')}</td>
            <td>{r.get('followup_deadline','')}</td>
            <td>{r.get('followup_date','')}</td>
            <td><strong style="color:{wd_color}">{wd}天</strong></td>
          </tr>
"""

    html += """        </tbody>
      </table>
    </div>

    <div id="tab-missing" class="tab-content">
"""

    if missing_rows:
        html += """      <table>
        <thead>
          <tr>
            <th>CRA</th>
            <th>Site</th>
            <th>监查日期</th>
            <th>截止日期</th>
            <th>跟进函状态</th>
          </tr>
        </thead>
        <tbody>
"""
        for r in missing_rows:
            html += f"""          <tr style="background:#fff1f0">
            <td><strong style="color:#dc3545">{r.get('cra','')}</strong></td>
            <td style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="{r.get('site_name','')}">{r.get('site_name','')}</td>
            <td>{r.get('visit_end','')}</td>
            <td>{r.get('followup_deadline','')}</td>
            <td><span class="badge badge-red">漏登记</span></td>
          </tr>
"""
        html += """        </tbody>
      </table>
"""
    else:
        html += """      <div style="padding:40px;text-align:center;color:#28a745;font-size:16px">✅ 暂无漏登记记录</div>
"""

    html += """    </div>
  </div>
</div>

<div class="footer">
  由 monitoring-visit-analysis Skill 自动生成 &nbsp;|&nbsp; """ + today_str + """
</div>

<script>
function showTab(id) {
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(el => el.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  event.target.classList.add('active');
}

// Monthly chart
new Chart(document.getElementById('monthChart'), {
  type: 'bar',
  data: {
    labels: """ + json.dumps(month_labels) + """,
    datasets: [
      {
        label: '访视数',
        data: """ + json.dumps(month_totals) + """,
        backgroundColor: 'rgba(22, 119, 255, 0.7)',
        borderRadius: 4,
      },
      {
        label: '超期数',
        data: """ + json.dumps(month_overdue) + """,
        backgroundColor: 'rgba(220, 53, 69, 0.75)',
        borderRadius: 4,
      }
    ]
  },
  options: {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { position: 'top', labels: { font: { size: 12 }, boxWidth: 14 } },
      tooltip: { mode: 'index', intersect: false }
    },
    scales: {
      x: { grid: { display: false }, ticks: { font: { size: 11 } } },
      y: { beginAtZero: true, ticks: { stepSize: 1, font: { size: 11 } } }
    }
  }
});
</script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"HTML Dashboard: {output_path}")
