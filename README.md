# Finish Coordination Tools

A suite of browser-based, local-first tools for finish coordination and room data workflows. Every tool runs entirely in the browser: no installs, no backend, and no project data leaving the user's machine.

This repository is the editable source for the five public demos and their synthetic sample workbooks. The Parkin project folder holds the case study and supporting evidence.

## Repo Structure

- `tools/` - browser tools you can open directly
- `sample-data/` - public-safe Excel files for demos and screenshots
- `screenshots/` - public-safe evidence captures
- `scripts/` - sample-data generator

## Tools

| Tool | File |
|---|---|
| Finish QAQC Auditor | `tools/finish-qaqc-auditor.html` |
| Finish Template Analyzer | `tools/finish-template-analyzer.html` |
| Template Finder | `tools/template-finder.html` |
| Room Type Analyzer | `tools/room-type-analyzer.html` |
| Column Comparator | `tools/column-comparator.html` |

## Usage

1. Open any `.html` file directly in your browser.
2. Use the matching files in `sample-data/`.
3. Export the results locally.

## Public-Safe Rule

This repo is intended to contain only public-safe artifacts.

- Use the files in `sample-data/` for demos and screenshots.
- Do not replace them with real project workbooks.
- Do not add client data, project room records, or internal exports.

## Maintenance Notes

- Choose a license intentionally. Right now the repo is source-visible but not licensed for reuse.
- Add screenshots to the README later if you want the repo to read more like a product showcase.
- Regenerate the synthetic inputs with `python3 scripts/generate_sample_data.py`. Keep internal tools and real project data in their authorized workspace.
