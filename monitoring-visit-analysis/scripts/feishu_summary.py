# -*- coding: utf-8 -*-
"""Send Feishu summary message via Feishu Bot API (Skill B integration)."""

import sys, os, json, requests
from collections import defaultdict

FEISHU_APP_ID = "cli_a92c5e338939dbd1"
FEISHU_APP_SECRET = "GBzQl8MhS89aNujABkQMKbUFJDhzLqDH"
FEISHU_USER_OPEN_ID = "ou_8a96d4aa91371488f194e3475a841112"
FEISHU_API = "https://open.feishu.cn/open-apis"


def get_app_token():
    """Get Feishu app access token."""
    resp = requests.post(
        f"{FEISHU_API}/auth/v3/tenant_access_token/internal",
        headers={"Content-Type": "application/json"},
        json={"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET},
        timeout=10,
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"Failed to get app token: {data}")
    return data["tenant_access_token"]


def send_feishu_message(token, open_id, msg):
    """Send text message to Feishu user by open_id."""
    resp = requests.post(
        f"{FEISHU_API}/im/v1/messages?receive_id_type=open_id",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={
            "receive_id": open_id,
            "msg_type": "text",
            "content": json.dumps({"text": msg}),
        },
        timeout=15,
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"Failed to send message: {data}")
    return data.get("data", {}).get("message_id")


def build_summary_message(summary, cra_stats, month_stats, visited, today_str="2026-04-19"):
    """Build the formatted markdown summary string."""

    # Count unique sites per CRA
    cra_site_count = defaultdict(int)
    for r in visited:
        cra_site_count[r.get("cra", "")] += 1

    # --- 超期跟进函 Top5 ---
    fu_od_rows = [r for r in visited if r.get("followup_overdue") == "是"]
    fu_od_rows.sort(key=lambda r: r.get("followup_wd", 0), reverse=True)
    top5 = fu_od_rows[:5]

    top5_lines = []
    for i, r in enumerate(top5, 1):
        wd = r.get("followup_wd", 0)
        top5_lines.append(
            f"| {i} | {r.get('cra','')} | {str(r.get('site_name',''))[:18]} | "
            f"{r.get('visit_end','')} | {wd}天 |"
        )

    # --- 漏登记 ---
    missing_rows = [r for r in visited if r.get("missing_reg")]
    missing_lines = []
    for r in missing_rows:
        missing_lines.append(
            f"| {r.get('cra','')} | {str(r.get('site_name',''))[:18]} | "
            f"{r.get('visit_end','')} | {r.get('followup_deadline','')} |"
        )

    # --- CRA 风险 ---
    cra_list = []
    for name, c in cra_stats.items():
        total_od = c.get("fin_od", 0) + c.get("fu_od", 0) + c.get("missing", 0)
        cra_list.append({
            "cra": name,
            "total_od": total_od,
            "fin_od": c.get("fin_od", 0),
            "fu_od": c.get("fu_od", 0),
            "missing": c.get("missing", 0),
            "risk": c.get("risk", "正常"),
            "sites": cra_site_count.get(name, 0),
        })
    cra_list.sort(key=lambda x: -x["total_od"])
    high_risk = [c for c in cra_list if c["risk"] in ("高风险",)]
    attention  = [c for c in cra_list if c["risk"] in ("关注",)]

    cra_lines = []
    for c in (high_risk + attention)[:10]:
        level = "🔴高风险" if c["risk"] == "高风险" else "🟡关注"
        cra_lines.append(f"| {c['cra']} | {level} | {c['total_od']} | {c['sites']}个Site |")

    # --- Month trend ---
    month_lines = []
    for mo, m in sorted(month_stats.items())[-12:]:
        total_v  = m.get("total", 0)
        fin_od   = m.get("fin_od", 0)
        fu_od    = m.get("fu_od", 0)
        miss     = m.get("missing", 0)
        total_od = fin_od + fu_od + miss
        trend = "✅" if total_od == 0 else ("⚠️峰值" if total_od >= 3 else "↗️")
        month_lines.append(f"| {mo} | {total_v} | {total_od} | {trend} |")

    # --- Key figures ---
    sub_od   = summary.get("total_sub_od", 0)
    fin_od   = summary.get("total_fin_od", 0)
    fu_od    = summary.get("total_fu_od", 0)
    missing  = summary.get("total_missing", 0)
    finalized = summary.get("total_finalized", 0)
    visited_n = summary.get("total_visited", 0)
    high_sites = summary.get("high_risk_sites", [])

    sub_icon  = "✅" if sub_od  == 0 else "⚠️"
    fin_icon  = "⚠️" if fin_od  > 0  else "✅"
    fu_icon   = "🔴" if fu_od   > 0  else "✅"
    miss_icon = "🔴" if missing > 0  else "✅"

    all_ends = [str(r["visit_end"]) for r in visited if r.get("visit_end")]
    date_range = f"{min(all_ends)} ~ {max(all_ends)}" if all_ends else today_str

    # Feishu only supports plain text (no markdown tables), so use plain text format
    lines = []
    lines.append("📊 CRA监查超期分析报告")
    lines.append(f"分析范围：{date_range}（共{finalized}条已定稿记录，{visited_n}条已访视）")
    lines.append("")
    lines.append("【超期汇总】")
    lines.append(f"  报告递交超期（>5工作日）: {sub_od}条 {sub_icon}")
    lines.append(f"  报告定稿超期（>10工作日）: {fin_od}条 {fin_icon}")
    lines.append(f"  跟进函已发但超期: {fu_od}条 {fu_icon}")
    lines.append(f"  跟进函漏登记: {missing}条 {miss_icon}")
    lines.append("")

    if top5_lines:
        lines.append("【超期跟进函 Top5（按超期天数降序）】")
        for l in top5_lines:
            lines.append("  " + l)
        lines.append("")

    if missing_lines:
        lines.append(f"【跟进函漏登记（{len(missing_rows)}条）】")
        for l in missing_lines:
            lines.append("  " + l)
        lines.append("")

    if month_lines:
        lines.append("【月度趋势】")
        for l in month_lines:
            lines.append("  " + l)
        lines.append("")

    if cra_lines:
        lines.append("【CRA风险分级】")
        for l in cra_lines:
            lines.append("  " + l)
        lines.append("")

    if high_sites:
        lines.append(f"【高风险中心（共{len(high_sites)}个）】")
        lines.append("  " + "  ".join(str(s) for s in high_sites[:15]))
        lines.append("")

    lines.append("📎 Excel + PPT 报告已生成，请查收上方文件")

    return "\n".join(lines)


def send_feishu_summary(summary, cra_stats, month_stats, visited, today_str="2026-04-19"):
    """Main entry point: fetch token and send message."""
    try:
        token = get_app_token()
        msg = build_summary_message(summary, cra_stats, month_stats, visited, today_str)
        message_id = send_feishu_message(token, FEISHU_USER_OPEN_ID, msg)
        print(f"[feishu_summary] Message sent OK: {message_id}")
        return message_id
    except Exception as e:
        print(f"[feishu_summary] Error: {e}")
        return None


if __name__ == "__main__":
    # Load data from JSON files written by main.py
    if len(sys.argv) < 2:
        print("Usage: python feishu_summary.py <summary_json> [visited_json]")
        sys.exit(1)

    summary_json = sys.argv[1]
    visited_json = sys.argv[2] if len(sys.argv) > 2 else None

    with open(summary_json, encoding="utf-8") as f:
        summary = json.load(f)

    visited = []
    if visited_json:
        with open(visited_json, encoding="utf-8") as f:
            visited = json.load(f)

    today_str = summary.get("today_str", "2026-04-19")

    send_feishu_summary(
        summary,
        summary.get("cra_stats", {}),
        summary.get("month_stats", {}),
        visited,
        today_str,
    )
