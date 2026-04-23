# Public Sample Inputs

These Excel files are public-safe demonstration inputs for the local finish coordination tools.

They are intentionally fictional. The files preserve the column names, sheet shapes, and data patterns the browser tools expect, but they do not contain real project room numbers, departments, comments, client data, or hospital/project records.

## Why This Exists

The original private workbooks are useful schema references, but they should not be used for portfolio screenshots or public demos. These files give the tools enough realistic structure to run while keeping the content safe to share.

## Files

| Tool | Upload these files | What the sample demonstrates |
|---|---|---|
| `finish-qaqc-auditor.html` | `qaqc-room-report.xlsx` and `qaqc-finish-item-list.xlsx` | Finish value validation, missing values, duplicate entries, and ceiling material/finish rule checks. |
| `finish-template-analyzer.html` | `template-analyzer-room-report.xlsx` | Grouping repeated room finish combinations by department or another anchor column. |
| `template-finder.html` | `template-finder-item-list.xlsx` | Interactive lookup of template codes from selected floor/base/wall/ceiling values. |
| `room-type-analyzer.html` | `room-type-analyzer-sdp-reference.xlsx` | SDP reference parsing, generated room type abbreviations, skipped blank references, and variant numbering. |
| `column-comparator.html` | `column-comparator-main.xlsx` and `column-comparator-rules.xlsx` | RFS-to-RDS comparison with match, mismatch, and no-rule examples. |

## Screenshot Guidance

- Use these public files for portfolio screenshots, not the private project files.
- Keep screenshots focused on the tool interface, the workflow, and the output logic.
- It is okay for screenshots to show the fictional sample names because they are not project records.
- If a screenshot shows a file name, prefer these public filenames.

## Regeneration

Run this from the repository root:

```bash
python3 scripts/generate_sample_data.py
```

The generator uses only the Python standard library, so it does not need `openpyxl`, `xlsx`, or LibreOffice.
