# PSI TPM Report Generator

A CLI tool that reads a project tracker Excel file and generates a formatted HTML weekly status report.

## Setup

```bash
pip install -r requirements.txt
```

## Usage

```bash
python report_generator.py sample_tracker.xlsx
```

With options:

```bash
python report_generator.py sample_tracker.xlsx --week "10 Nov 2026" --output reports/
```

The HTML report is saved to `output/status_report_<timestamp>.html`.

## Expected Excel Format

The tracker must have two sheets:

**Sheet: `Summary`**  
Data starts at row 4 with columns:
| Project | Client | Status | Next Milestone | Milestone Date | Main Problem |

Status values: `RED`, `AMBER`, or `GREEN`

**Sheet: `Risks and Issues`**  
Data starts at row 4 with columns:
| Project | Description | Probability | Impact | Owner | Action | Due Date |

Probability / Impact values: `H` (High), `M` (Medium), `L` (Low)

## Sample Output

The report includes:
- Summary tiles (total projects, RED / AMBER / GREEN counts, high-risk count)
- One card per project showing status badge, next milestone, current situation, and risks table

## AI Assistance

This tool was built with help from [Claude Code](https://claude.com/claude-code).

### Prompts used

1. *"Build a Python CLI tool that reads the Task 5 Excel tracker and generates a formatted HTML weekly status report per client. Save output to an `output/` directory."*

2. *"Add summary tile row at the top showing total projects, RED/AMBER/GREEN counts, and high-risk count."*

### AI error encountered and fixed

**Error:** Initial script used `min_row=2` to read data rows, but the tracker had a merged header spanning rows 1–3, causing the first data row to be skipped.

**Fix:** Changed `min_row` to `4` in both `read_summary()` and `read_risks()` to match the actual first data row.
