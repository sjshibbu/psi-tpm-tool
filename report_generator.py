"""
PSI TPM Weekly Status Report Generator
Reads a project tracker Excel file and generates HTML status reports per client.
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("Error: openpyxl not installed. Run: pip install openpyxl")
    sys.exit(1)


STATUS_COLOR = {
    "RED":   "#dc2626",
    "AMBER": "#d97706",
    "GREEN": "#16a34a",
}

STATUS_BG = {
    "RED":   "#fee2e2",
    "AMBER": "#fef3c7",
    "GREEN": "#dcfce7",
}

PROB_COLOR = {
    "H": "#dc2626",
    "M": "#d97706",
    "L": "#16a34a",
}


def read_summary(ws):
    projects = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        if not row[0]:
            continue
        projects.append({
            "project":   row[0],
            "client":    row[1],
            "status":    row[2],
            "milestone": row[3],
            "milestone_date": row[4],
            "problem":   row[5],
        })
    return projects


def read_risks(ws):
    risks = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        if not row[0]:
            continue
        risks.append({
            "project":     row[0],
            "description": row[1],
            "probability": row[2],
            "impact":      row[3],
            "owner":       row[4],
            "action":      row[5],
            "due_date":    row[6],
        })
    return risks


def build_report(projects, risks, week_label):
    sections = []

    for proj in projects:
        name    = proj["project"]
        client  = proj["client"]
        status  = proj["status"] or "GREEN"
        color   = STATUS_COLOR.get(status, "#64748b")
        bg      = STATUS_BG.get(status, "#f1f5f9")
        prob_risks = [r for r in risks if r["project"] == name]

        risk_rows = ""
        for r in prob_risks:
            pc = PROB_COLOR.get(r["probability"], "#64748b")
            ic = PROB_COLOR.get(r["impact"], "#64748b")
            risk_rows += f"""
            <tr>
              <td style="padding:8px 10px;border-bottom:1px solid #e2e8f0;font-size:13px;">{r["description"]}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #e2e8f0;text-align:center;">
                <span style="color:{pc};font-weight:700;">{r["probability"]}</span>
              </td>
              <td style="padding:8px 10px;border-bottom:1px solid #e2e8f0;text-align:center;">
                <span style="color:{ic};font-weight:700;">{r["impact"]}</span>
              </td>
              <td style="padding:8px 10px;border-bottom:1px solid #e2e8f0;font-size:13px;">{r["owner"]}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #e2e8f0;font-size:13px;">{r["action"]}</td>
              <td style="padding:8px 10px;border-bottom:1px solid #e2e8f0;font-size:13px;white-space:nowrap;">{r["due_date"]}</td>
            </tr>"""

        sections.append(f"""
    <div style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;margin-bottom:28px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.06);">
      <!-- Project header -->
      <div style="padding:18px 24px;border-bottom:1px solid #e2e8f0;display:flex;align-items:center;justify-content:space-between;">
        <div>
          <div style="font-size:18px;font-weight:700;color:#0f172a;">{name}</div>
          <div style="font-size:13px;color:#64748b;margin-top:2px;">Client: {client}</div>
        </div>
        <div style="background:{bg};color:{color};font-weight:700;font-size:13px;padding:6px 16px;border-radius:99px;border:2px solid {color};">
          {status}
        </div>
      </div>

      <!-- Milestone -->
      <div style="padding:14px 24px;background:#f8fafc;border-bottom:1px solid #e2e8f0;display:flex;gap:40px;">
        <div>
          <div style="font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.6px;font-weight:600;">Next Milestone</div>
          <div style="font-size:14px;font-weight:600;color:#1e293b;margin-top:3px;">{proj["milestone"]}</div>
        </div>
        <div>
          <div style="font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.6px;font-weight:600;">Date</div>
          <div style="font-size:14px;font-weight:600;color:#1e293b;margin-top:3px;">{proj["milestone_date"]}</div>
        </div>
      </div>

      <!-- Main problem -->
      <div style="padding:14px 24px;border-bottom:1px solid #e2e8f0;">
        <div style="font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.6px;font-weight:600;margin-bottom:6px;">Current Situation</div>
        <div style="font-size:13px;color:#334155;line-height:1.6;">{proj["problem"]}</div>
      </div>

      <!-- Risks table -->
      <div style="padding:14px 24px 20px;">
        <div style="font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.6px;font-weight:600;margin-bottom:10px;">
          Risks &amp; Issues ({len(prob_risks)})
        </div>
        <table style="width:100%;border-collapse:collapse;font-family:Arial,sans-serif;">
          <thead>
            <tr style="background:#f1f5f9;">
              <th style="padding:8px 10px;text-align:left;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #e2e8f0;">Description</th>
              <th style="padding:8px 10px;text-align:center;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #e2e8f0;">Prob</th>
              <th style="padding:8px 10px;text-align:center;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #e2e8f0;">Impact</th>
              <th style="padding:8px 10px;text-align:left;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #e2e8f0;">Owner</th>
              <th style="padding:8px 10px;text-align:left;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #e2e8f0;">Action</th>
              <th style="padding:8px 10px;text-align:left;font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.5px;border-bottom:2px solid #e2e8f0;">Due</th>
            </tr>
          </thead>
          <tbody>{risk_rows}</tbody>
        </table>
      </div>
    </div>""")

    red_count   = sum(1 for p in projects if p["status"] == "RED")
    amber_count = sum(1 for p in projects if p["status"] == "AMBER")
    green_count = sum(1 for p in projects if p["status"] == "GREEN")
    high_risks  = sum(1 for r in risks if r["probability"] == "H" and r["impact"] == "H")

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PSI Weekly Status Report — {week_label}</title>
<style>
  body {{ font-family: Arial, sans-serif; background: #f0f4f8; margin: 0; padding: 24px; color: #0f172a; }}
  .container {{ max-width: 960px; margin: 0 auto; }}
</style>
</head>
<body>
<div class="container">

  <!-- Header -->
  <div style="margin-bottom:28px;">
    <div style="font-size:24px;font-weight:700;color:#1F3864;">PSI — Weekly Status Report</div>
    <div style="font-size:14px;color:#64748b;margin-top:4px;">Week: {week_label}  ·  TPM: Shivam Jatav  ·  Generated: {datetime.now().strftime("%d %b %Y %H:%M")}</div>
  </div>

  <!-- Summary tiles -->
  <div style="display:flex;gap:16px;margin-bottom:28px;flex-wrap:wrap;">
    <div style="flex:1;min-width:140px;background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:16px 20px;text-align:center;">
      <div style="font-size:28px;font-weight:700;color:#1F3864;">{len(projects)}</div>
      <div style="font-size:12px;color:#64748b;text-transform:uppercase;letter-spacing:.6px;margin-top:2px;">Projects</div>
    </div>
    <div style="flex:1;min-width:140px;background:#fee2e2;border:1px solid #fca5a5;border-radius:10px;padding:16px 20px;text-align:center;">
      <div style="font-size:28px;font-weight:700;color:#dc2626;">{red_count}</div>
      <div style="font-size:12px;color:#dc2626;text-transform:uppercase;letter-spacing:.6px;margin-top:2px;">Red</div>
    </div>
    <div style="flex:1;min-width:140px;background:#fef3c7;border:1px solid #fcd34d;border-radius:10px;padding:16px 20px;text-align:center;">
      <div style="font-size:28px;font-weight:700;color:#d97706;">{amber_count}</div>
      <div style="font-size:12px;color:#d97706;text-transform:uppercase;letter-spacing:.6px;margin-top:2px;">Amber</div>
    </div>
    <div style="flex:1;min-width:140px;background:#dcfce7;border:1px solid #86efac;border-radius:10px;padding:16px 20px;text-align:center;">
      <div style="font-size:28px;font-weight:700;color:#16a34a;">{green_count}</div>
      <div style="font-size:12px;color:#16a34a;text-transform:uppercase;letter-spacing:.6px;margin-top:2px;">Green</div>
    </div>
    <div style="flex:1;min-width:140px;background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:16px 20px;text-align:center;">
      <div style="font-size:28px;font-weight:700;color:#dc2626;">{high_risks}</div>
      <div style="font-size:12px;color:#64748b;text-transform:uppercase;letter-spacing:.6px;margin-top:2px;">High Risks</div>
    </div>
  </div>

  <!-- Project sections -->
  {"".join(sections)}

  <div style="text-align:center;font-size:12px;color:#94a3b8;margin-top:16px;">
    Generated by PSI TPM Report Generator · github.com/shivamjatav/psi-tpm-tool
  </div>

</div>
</body>
</html>"""
    return html


def main():
    parser = argparse.ArgumentParser(
        description="PSI TPM Weekly Status Report Generator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python report_generator.py tracker.xlsx
  python report_generator.py tracker.xlsx --week "10 Nov 2026" --output reports/
        """
    )
    parser.add_argument("tracker", help="Path to the Excel tracker file (.xlsx)")
    parser.add_argument("--week", default=datetime.now().strftime("%d %b %Y"),
                        help="Week label for the report (default: today's date)")
    parser.add_argument("--output", default="output",
                        help="Output directory for HTML reports (default: output/)")

    args = parser.parse_args()

    tracker_path = Path(args.tracker)
    if not tracker_path.exists():
        print(f"Error: File not found: {tracker_path}")
        sys.exit(1)

    print(f"Reading tracker: {tracker_path}")

    try:
        wb = openpyxl.load_workbook(tracker_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)

    if "Summary" not in wb.sheetnames:
        print("Error: 'Summary' sheet not found in tracker.")
        sys.exit(1)
    if "Risks and Issues" not in wb.sheetnames:
        print("Error: 'Risks and Issues' sheet not found in tracker.")
        sys.exit(1)

    projects = read_summary(wb["Summary"])
    risks    = read_risks(wb["Risks and Issues"])

    print(f"Found {len(projects)} project(s), {len(risks)} risk(s)/issue(s)")

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    report_html = build_report(projects, risks, args.week)

    timestamp   = datetime.now().strftime("%Y%m%d_%H%M")
    output_file = output_dir / f"status_report_{timestamp}.html"
    output_file.write_text(report_html, encoding="utf-8")

    print(f"Report saved: {output_file}")
    print(f"Open in browser: file://{output_file.resolve()}")


if __name__ == "__main__":
    main()
