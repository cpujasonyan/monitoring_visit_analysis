# -*- coding: utf-8 -*-
"""Monitoring Visit Analysis - Entry Point.

Usage:
    python main.py <input_excel> <output_dir> [--project NAME] [--today YYYY-MM-DD] [--no-feishu]
"""

import sys, os, argparse, tempfile, json
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from analyze import read_excel, compute_overdue, compute_stats, compute_summary
from excel_report import generate_excel
from ppt_report import generate_ppt
from html_report import generate_html


def main():
    parser = argparse.ArgumentParser(description='Monitoring Visit Analysis Report Generator')
    parser.add_argument('input_file', help='Path to the Monitoring Visit Tracking Log Excel file')
    parser.add_argument('output_dir', help='Directory to save output reports')
    parser.add_argument('--work-dir', default=None, help='Directory for temporary chart images')
    parser.add_argument('--project', default='', help='Project name for report title')
    parser.add_argument('--today', default=None, help='Override analysis date (YYYY-MM-DD, default: today)')
    parser.add_argument('--no-feishu', action='store_true', help='Skip Feishu summary message')
    args = parser.parse_args()

    today = date.today() if not args.today else datetime.strptime(args.today, '%Y-%m-%d').date()
    today_str = today.strftime('%Y-%m-%d')
    date_str = today.strftime('%Y%m%d')
    work_dir = args.work_dir or tempfile.mkdtemp(prefix='mv_analysis_')
    os.makedirs(args.output_dir, exist_ok=True)
    os.makedirs(work_dir, exist_ok=True)

    print(f"读取源文件: {args.input_file}")
    rows_data = read_excel(args.input_file)
    print(f"共读取 {len(rows_data)} 条记录")

    print("执行分析...")
    visited = compute_overdue(rows_data, today)
    print(f"已完成访视: {len(visited)} 条")

    cra_stats, month_stats, site_stats = compute_stats(visited)
    summary = compute_summary(rows_data, visited, cra_stats, site_stats)
    print(f"递交超期: {summary['total_sub_od']}, 定稿超期: {summary['total_fin_od']}, "
          f"跟进函超期: {summary['total_fu_od']}, 漏登记: {summary['total_missing']}")

    excel_out = os.path.join(args.output_dir, f'Monitoring_Visit_分析报告_{date_str}.xlsx')
    print(f"生成 Excel: {excel_out}")
    generate_excel(rows_data, visited, cra_stats, month_stats, site_stats, summary, excel_out, today)

    ppt_out = os.path.join(args.output_dir, f'Monitoring_Visit_可视化报告_{date_str}.pptx')
    print(f"生成 PPT: {ppt_out}")
    generate_ppt(visited, cra_stats, month_stats, site_stats, summary, work_dir, ppt_out, today, args.project)

    # Export analysis results to JSON for Feishu summary
    summary_json = os.path.join(work_dir, 'analysis_summary.json')
    visited_json = os.path.join(work_dir, 'analysis_visited.json')

    # Convert date objects to strings for JSON serialization
    def make_serializable(obj):
        if isinstance(obj, list):
            return [{k: str(v) if hasattr(v, 'year') else v for k, v in r.items()} for r in obj]
        return obj

    with open(summary_json, 'w', encoding='utf-8') as f:
        json.dump({**summary, 'today_str': today_str}, f, ensure_ascii=False, indent=2, default=str)

    with open(visited_json, 'w', encoding='utf-8') as f:
        json.dump(make_serializable(visited), f, ensure_ascii=False, default=str)

    html_out = os.path.join(args.output_dir, f'Monitoring_Visit_Dashboard_{date_str}.html')
    print(f"生成 HTML: {html_out}")
    generate_html(visited, cra_stats, month_stats, site_stats, summary, html_out, today_str)

    print(f"完成!")
    print(f"Excel: {excel_out}")
    print(f"PPT:   {ppt_out}")
    print(f"HTML:  {html_out}")

    # Send Feishu summary (unless --no-feishu)
    if not args.no_feishu:
        try:
            from feishu_summary import send_feishu_summary
            send_feishu_summary(
                {**summary, 'today_str': today_str},
                cra_stats,
                month_stats,
                visited,
                today_str,
            )
        except Exception as e:
            print(f"[Feishu] 汇总发送失败: {e}")


if __name__ == '__main__':
    main()
