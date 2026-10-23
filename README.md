# Sales Ranking Extractor

**A spreadsheet-to-text reporting tool built for an operational need.**

I built this tool to summarize sales achieved from a ranking sheet that was updated periodically. My campaign team leader used the report directly to support logistics and management.

This repository is a cleaned, runnable Python desktop edition. **All sample teams, agents, figures, and notes are fictional.**

[Portfolio case study](https://doma990.github.io/projects/sales-ranking-extractor/) · [Sample workbook](samples/sample-ranking.xlsx) · [Generated report](samples/sample-report.txt)

## From ranking sheet to report

The extractor reads sales values, parses an additional count from the sales cells’ comments, groups agents by campaign, and writes totals and individual entries to a UTF-8 text file.

```text
sample-ranking.xlsx       Python + openpyxl        sample-report.txt
Sales cells + notes   →   Validate and aggregate → Team and overall totals
```

The fictional workbook has two campaign blocks, named Team A and Team B:

| Campaign | Agent | Sales | Note attached to the sales cell |
| --- | --- | ---: | --- |
| Team A | Agent 01 | 18 | 12 confirmed; 1 duplicate |
| Team A | Agent 02 | 7 | 3 confirmed |
| Team A | Agent 03 | 0 | No note |
| Team B | Agent 04 | 9 | 2 confirmed; confirmed |
| Team B | Agent 05 | 5 | No note |
| Team B | Agent 06 | - | Excluded |

**Result: 39 sales and 18 confirmed.** Sales are summed from cell values; confirmed counts are summed separately from cell notes. The workbook’s TOTAL formulas are for the reader; the program recomputes totals from agent rows.

### How comments become numbers

The XLSX file uses traditional Excel comments, now called **Notes** in Excel. In Excel, right-click B3, B4, or E3 and choose **Show/Hide Note**. Threaded comment conversations are not supported.

- `12 confirmed` contributes **12**, including all digits.
- A bare `confirmed` contributes **1**.
- `2 confirmed; confirmed` contributes **3**.
- Missing comments contribute **0**.
- Separate note segments with a semicolon, comma, newline, or `|`. Matching ignores case.
- Other labels, such as `1 duplicate`, stay in the report but do not add to the confirmed count.
- A segment containing “confirmed” in an unsupported form, such as `-2 confirmed` or `two confirmed`, raises an error instead of guessing.

“Confirmed” is a generic label chosen for this public demonstration.

## Run the sample

Requires Python 3.10 or newer. Download the repository ZIP and extract it, then open a terminal in its folder.

**Windows**
```powershell
py -m pip install -r requirements.txt
py ranking_report.py samples/sample-ranking.xlsx --sheet "October 2026" --output samples/my-report.txt
```

**macOS / Linux**
```sh
python3 -m pip install -r requirements.txt
python3 ranking_report.py samples/sample-ranking.xlsx --sheet "October 2026" --output samples/my-report.txt
```

The report appears at `samples/my-report.txt`. If your Python installation requires an isolated environment, use `python3 -m venv .venv` and activate it before installing.

For another compatible workbook:
```sh
python3 ranking_report.py ranking.xlsx --sheet "October 2026"
```
Without `--output`, `ranking-report.txt` is written beside the input workbook. Add `--campaign "Team A"` to select one campaign; repeat the option for several.

## Expected sheet layout

Each campaign uses two adjacent columns. Its name is in **row 1** above the agent column; **row 2** contains `Agent` and `Sales`. Agent entries start in **row 3**. Multiple campaign blocks can sit side by side, separated by blank columns.

- Use `--sheet` to select the exact worksheet. A workbook with one sheet can omit it; a workbook with several must specify it.
- Blank agent rows are skipped. A `TOTAL` agent label ends that campaign block.
- Sales must be nonnegative whole numbers. Zero is kept; `-` skips an agent. Missing or invalid sales raise an error with the cell address.
- Duplicate campaign headings or duplicate agent names within one campaign are rejected.
- Sales formulas require an existing cached result from Excel or another spreadsheet application. The tool does not calculate formulas.
- The workbook is opened read-only in purpose: it is never saved or changed by the extractor. The entire report is validated before the output is replaced.

## Original workflow and public edition

The tool grew out of spreadsheet-based reporting and was later distributed as a Windows executable, with the ranking workbook and generated report in the same folder. This repository documents the Python desktop implementation. It does not include an Apps Script integration or the old executable.

The public edition removes fixed workplace paths and campaign names, fixes worksheet selection and multi-digit comment counts, handles absent notes, and validates inputs. The original workplace workbook and report are excluded. Only the fictional workbook and report are tracked as sample data.

## Checks

```sh
python3 -m unittest discover -s tests -v
```

Regression tests cover worksheet selection, two-digit and repeated counts, missing notes, blank rows, zero and invalid sales, duplicate entries, preservation of an existing report on error, and the actual sample command’s output.

## Files

- `ranking_report.py` — extraction, report rendering, and command-line entry point.
- `requirements.txt` — the single external dependency.
- `samples/sample-ranking.xlsx` — fictional workbook with cell notes.
- `samples/sample-report.txt` — actual output produced from that workbook.
- `tests/test_ranking_report.py` — regression checks using fictional data.
