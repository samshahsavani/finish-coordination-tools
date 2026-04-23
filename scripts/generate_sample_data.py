#!/usr/bin/env python3
"""
Generate public-safe sample Excel inputs for the finish coordination tools.

The real test files remain useful as private schema references. These outputs
keep the same column shapes and workflow triggers, but all room names,
departments, IDs, and finish combinations are fictional.
"""

from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "sample-data"


def col_name(index: int) -> str:
    name = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def cell_xml(row_idx: int, col_idx: int, value: object) -> str:
    ref = f"{col_name(col_idx)}{row_idx}"
    text = "" if value is None else str(value)
    return (
        f'<c r="{ref}" t="inlineStr"><is><t>{escape(text)}</t></is></c>'
    )


def sheet_xml(rows: list[list[object]]) -> str:
    body = []
    max_cols = max((len(row) for row in rows), default=0)
    for r_idx, row in enumerate(rows, start=1):
        padded = list(row) + [""] * (max_cols - len(row))
        cells = "".join(cell_xml(r_idx, c_idx, value) for c_idx, value in enumerate(padded))
        body.append(f'<row r="{r_idx}">{cells}</row>')

    dimension = f"A1:{col_name(max_cols - 1)}{len(rows)}" if rows and max_cols else "A1"
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="{dimension}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    {''.join(body)}
  </sheetData>
</worksheet>'''


def workbook_xml(sheet_names: list[str]) -> str:
    sheets = []
    for idx, name in enumerate(sheet_names, start=1):
        sheets.append(
            f'<sheet name="{escape(name)}" sheetId="{idx}" r:id="rId{idx}"/>'
        )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>{''.join(sheets)}</sheets>
</workbook>'''


def workbook_rels(sheet_count: int) -> str:
    rels = []
    for idx in range(1, sheet_count + 1):
        rels.append(
            f'<Relationship Id="rId{idx}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{idx}.xml"/>'
        )
    rels.append(
        f'<Relationship Id="rId{sheet_count + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {''.join(rels)}
</Relationships>'''


def content_types(sheet_count: int) -> str:
    overrides = [
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    for idx in range(1, sheet_count + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{idx}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  {''.join(overrides)}
</Types>'''


def styles_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''


def write_xlsx(path: Path, sheets: dict[str, list[list[object]]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    names = list(sheets.keys())
    with ZipFile(path, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types(len(names)))
        z.writestr("_rels/.rels", '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>''')
        z.writestr("xl/workbook.xml", workbook_xml(names))
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels(len(names)))
        z.writestr("xl/styles.xml", styles_xml())
        z.writestr("docProps/core.xml", '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Sam Shahsavani</dc:creator>
  <cp:lastModifiedBy>Sam Shahsavani</cp:lastModifiedBy>
</cp:coreProperties>''')
        z.writestr("docProps/app.xml", '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Public sample generator</Application>
</Properties>''')
        for idx, rows in enumerate(sheets.values(), start=1):
            z.writestr(f"xl/worksheets/sheet{idx}.xml", sheet_xml(rows))


FINISH_ITEM_ROWS = [
    ["Name", "BIM ID", "Reference"],
    ["Resilient sheet flooring - standard", "RSF-1-01", "RSF"],
    ["Epoxy flooring - wet room", "EPOXY-1-01", "EPOXY"],
    ["Sealed concrete", "SCON-1-01", "SCON"],
    ["Integral cove base", "IC", "IC"],
    ["Stainless steel base", "SST-B01", "SST"],
    ["Paint finish", "PT", "PT"],
    ["Epoxy paint", "E/PT", "E/PT"],
    ["Sheet wall protection", "SPW", "SPW"],
    ["Ceramic wall tile", "CWT", "CWT"],
    ["Luxury vinyl tile - sample", "LVT-1-01", "LVT"],
    ["Rubber base - sample", "RB-1", "RB"],
    ["Wall protection panel", "WPT", "WPT"],
    ["Fiberglass reinforced panel", "FRP", "FRP"],
    ["Gypsum board ceiling", "GB", "GB"],
    ["Moisture resistant gypsum board", "MRGB", "MRGB"],
    ["Acoustic ceiling tile", "ACT", "ACT"],
    ["Special acoustic ceiling tile", "ACT-S", "ACT-S"],
    ["Fiber cement ceiling module", "FCM", "FCM"],
    ["Exposed structure", "EXP", "EXP"],
]


ROOM_HEADERS = [
    "Room Function #",
    "Room Number",
    "Room Name",
    "Department (Name)",
    "Room Data Template: Code",
    "FLOOR",
    "BASE",
    "WALL",
    "CEILING MATERIAL",
    "CEILING FINISH",
]


QAQC_ROOM_ROWS = [
    ROOM_HEADERS,
    ["S.01.001", "P01-A-001", "Patient room, standard", "Sample Inpatient Unit", "S-TYP-001", "RSF-1-01", "IC", "PT, SPW", "GB", "PT"],
    ["S.01.002", "P01-A-002", "Patient washroom", "Sample Inpatient Unit", "S-TYP-002", "EPOXY-1-01", "IC", "CWT", "MRGB", "E/PT"],
    ["S.01.003", "P01-A-003", "Team workroom", "Sample Clinical Support", "S-TYP-003", "RSF-1-01", "IC", "PT", "ACT", ""],
    ["S.01.004", "P01-A-004", "Clean supply", "Sample Clinical Support", "S-TYP-004", "", "IC", "PT", "GB", "PT"],
    ["S.01.005", "P01-A-005", "Consult room", "Sample Ambulatory Clinic", "S-TYP-005", "LVT-UNKNOWN", "IC", "PT", "ACT", ""],
    ["S.01.006", "P01-A-006", "Medication room", "Sample Inpatient Unit", "S-TYP-006", "RSF-1-01", "IC", "PT, PT", "GB", "PT"],
    ["S.01.007", "P01-A-007", "Staff lounge", "Sample Staff Support", "S-TYP-007", "RSF-1-01", "IC", "PT", "GB", ""],
    ["S.01.008", "P01-A-008", "Office", "Sample Administration", "S-TYP-008", "RSF-1-01", "IC", "PT", "ACT", "PT"],
    ["S.01.009", "P01-A-009", "Mechanical access", "Sample Support", "S-TYP-009", "SCON-1-01", "SST-B01", "PT", "EXP", "PT"],
    ["S.01.010", "P01-A-010", "Equipment alcove", "Sample Support", "S-TYP-010", "SCON-1-01", "SST-B01", "PT", "EXP", "SCON-1-01"],
    ["S.01.011", "P01-A-011", "Interview room", "Sample Ambulatory Clinic", "S-TYP-011", "RSF-1-01", "IC", "PT", "", ""],
    ["S.01.012", "P01-A-012", "Procedure room", "Sample Ambulatory Clinic", "S-TYP-012", "EPOXY-1-01", "IC", "E/PT, SPW", "MRGB", "E/PT"],
    ["S.01.013", "P01-A-013", "Exam room", "Sample Ambulatory Clinic", "S-TYP-013", "LVT-1-01", "RB-1", "PT, WPT", "GB", "PT"],
    ["S.01.014", "P01-A-014", "Clean utility", "Sample Clinical Support", "S-TYP-014", "RSF-1-01", "IC", "FRP", "GB", "PT"],
    ["S.01.015", "P01-A-015", "Staff washroom", "Sample Staff Support", "S-TYP-015", "EPOXY-1-01", "IC", "CWT", "MRGB", "E/PT"],
    ["S.01.016", "P01-A-016", "Waiting alcove", "Sample Public Zone", "S-TYP-016", "LVT-1-01", "RB-1", "PT", "ACT-S", ""],
    ["S.01.017", "P01-A-017", "Storage room", "Sample Support", "S-TYP-017", "SCON-1-01", "SST-B01", "PT", "EXP", "SCON-1-01"],
    ["S.01.018", "P01-A-018", "Family consult", "Sample Public Zone", "S-TYP-018", "RSF-1-01", "IC", "PT, WPT", "ACT", ""],
    ["S.01.019", "P01-A-019", "Isolation ante room", "Sample Inpatient Unit", "S-TYP-019", "EPOXY-1-01", "IC", "E/PT", "MRGB", "E/PT"],
    ["S.01.020", "P01-A-020", "Shared equipment bay", "Sample Support", "S-TYP-020", "SCON-1-01", "SST-B01", "PT, FRP", "EXP", "PT"],
]


TEMPLATE_ANALYZER_ROWS = [
    ROOM_HEADERS,
    ["S.02.001", "P02-B-001", "Exam room", "Sample Ambulatory Clinic", "S-TYP-101", "RSF-1-01", "IC", "PT, SPW", "GB", "PT"],
    ["S.02.002", "P02-B-002", "Exam room", "Sample Ambulatory Clinic", "S-TYP-101", "RSF-1-01", "IC", "PT, SPW", "GB", "PT"],
    ["S.02.003", "P02-B-003", "Consult room", "Sample Ambulatory Clinic", "S-TYP-102", "RSF-1-01", "IC", "PT", "ACT", ""],
    ["S.02.004", "P02-B-004", "Office", "Sample Administration", "S-TYP-103", "RSF-1-01", "IC", "PT", "ACT", ""],
    ["S.02.005", "P02-B-005", "Medication room", "Sample Clinical Support", "S-TYP-104", "RSF-1-01", "IC", "E/PT", "GB", "E/PT"],
    ["S.02.006", "P02-B-006", "Clean supply", "Sample Clinical Support", "S-TYP-105", "RSF-1-01", "IC", "PT", "GB", "PT"],
    ["S.02.007", "P02-B-007", "Soiled utility", "Sample Clinical Support", "S-TYP-106", "EPOXY-1-01", "IC", "E/PT", "MRGB", "E/PT"],
    ["S.02.008", "P02-B-008", "Patient washroom", "Sample Inpatient Unit", "S-TYP-107", "EPOXY-1-01", "IC", "CWT", "MRGB", "E/PT"],
    ["S.02.009", "P02-B-009", "Patient room", "Sample Inpatient Unit", "S-TYP-108", "RSF-1-01", "IC", "PT, SPW", "GB", "PT"],
    ["S.02.010", "P02-B-010", "Patient room", "Sample Inpatient Unit", "S-TYP-108", "RSF-1-01", "IC", "PT, SPW", "GB", "PT"],
    ["S.02.011", "P02-B-011", "Staff lounge", "Sample Staff Support", "S-TYP-109", "RSF-1-01", "IC", "PT", "ACT", ""],
    ["S.02.012", "P02-B-012", "Equipment alcove", "Sample Support", "S-TYP-110", "SCON-1-01", "SST-B01", "PT", "EXP", "SCON-1-01"],
    ["S.02.013", "P02-B-013", "Exam room", "Sample Ambulatory Clinic", "S-TYP-101", "RSF-1-01", "IC", "PT, SPW", "GB", "PT"],
    ["S.02.014", "P02-B-014", "Exam room", "Sample Ambulatory Clinic", "S-TYP-111", "LVT-1-01", "RB-1", "PT, WPT", "GB", "PT"],
    ["S.02.015", "P02-B-015", "Team station", "Sample Clinical Support", "S-TYP-112", "LVT-1-01", "RB-1", "PT", "ACT-S", ""],
    ["S.02.016", "P02-B-016", "Team station", "Sample Clinical Support", "S-TYP-112", "LVT-1-01", "RB-1", "PT", "ACT-S", ""],
    ["S.02.017", "P02-B-017", "Clean utility", "Sample Clinical Support", "S-TYP-113", "RSF-1-01", "IC", "FRP", "GB", "PT"],
    ["S.02.018", "P02-B-018", "Waiting area", "Sample Public Zone", "S-TYP-114", "LVT-1-01", "RB-1", "PT", "ACT", ""],
    ["S.02.019", "P02-B-019", "Family consult", "Sample Public Zone", "S-TYP-115", "RSF-1-01", "IC", "PT, WPT", "ACT", ""],
    ["S.02.020", "P02-B-020", "Shared equipment bay", "Sample Support", "S-TYP-116", "SCON-1-01", "SST-B01", "PT, FRP", "EXP", "SCON-1-01"],
]


TEMPLATE_FINDER_ROWS = [
    ["Code", "Template Name", "Item List: Name", "Category: Name", "BIM ID", "Name", "Rooms"],
    ["F_TYP-001", "Clinical room - resilient floor", "Finishes", "Floor", "RSF-1-01", "Resilient sheet flooring", "3"],
    ["F_TYP-001", "Clinical room - resilient floor", "Finishes", "Base", "IC", "Integral cove base", "3"],
    ["F_TYP-001", "Clinical room - resilient floor", "Finishes", "Wall", "PT", "Paint finish", "3"],
    ["F_TYP-001", "Clinical room - resilient floor", "Finishes", "Wall", "SPW", "Sheet wall protection", "3"],
    ["F_TYP-001", "Clinical room - resilient floor", "Finishes", "Ceiling Material", "GB", "Gypsum board ceiling", "3"],
    ["F_TYP-001", "Clinical room - resilient floor", "Finishes", "Ceiling Finish", "PT", "Paint finish", "3"],
    ["F_TYP-002", "Wet room - epoxy and tile", "Finishes", "Floor", "EPOXY-1-01", "Epoxy flooring", "2"],
    ["F_TYP-002", "Wet room - epoxy and tile", "Finishes", "Base", "IC", "Integral cove base", "2"],
    ["F_TYP-002", "Wet room - epoxy and tile", "Finishes", "Wall", "CWT", "Ceramic wall tile", "2"],
    ["F_TYP-002", "Wet room - epoxy and tile", "Finishes", "Ceiling Material", "MRGB", "Moisture resistant gypsum board", "2"],
    ["F_TYP-002", "Wet room - epoxy and tile", "Finishes", "Ceiling Finish", "E/PT", "Epoxy paint", "2"],
    ["F_TYP-003", "Office - ACT ceiling", "Finishes", "Floor", "RSF-1-01", "Resilient sheet flooring", "4"],
    ["F_TYP-003", "Office - ACT ceiling", "Finishes", "Base", "IC", "Integral cove base", "4"],
    ["F_TYP-003", "Office - ACT ceiling", "Finishes", "Wall", "PT", "Paint finish", "4"],
    ["F_TYP-003", "Office - ACT ceiling", "Finishes", "Ceiling Material", "ACT", "Acoustic ceiling tile", "4"],
    ["F_TYP-004", "Support alcove - exposed ceiling", "Finishes", "Floor", "SCON-1-01", "Sealed concrete", "1"],
    ["F_TYP-004", "Support alcove - exposed ceiling", "Finishes", "Base", "SST-B01", "Stainless steel base", "1"],
    ["F_TYP-004", "Support alcove - exposed ceiling", "Finishes", "Wall", "PT", "Paint finish", "1"],
    ["F_TYP-004", "Support alcove - exposed ceiling", "Finishes", "Ceiling Material", "EXP", "Exposed structure", "1"],
    ["F_TYP-004", "Support alcove - exposed ceiling", "Finishes", "Ceiling Finish", "SCON-1-01", "Sealed concrete", "1"],
    ["F_TYP-005", "Public area - resilient tile", "Finishes", "Floor", "LVT-1-01", "Luxury vinyl tile", "5"],
    ["F_TYP-005", "Public area - resilient tile", "Finishes", "Base", "RB-1", "Rubber base", "5"],
    ["F_TYP-005", "Public area - resilient tile", "Finishes", "Wall", "PT", "Paint finish", "5"],
    ["F_TYP-005", "Public area - resilient tile", "Finishes", "Wall", "WPT", "Wall protection panel", "5"],
    ["F_TYP-005", "Public area - resilient tile", "Finishes", "Ceiling Material", "ACT-S", "Special acoustic ceiling tile", "5"],
    ["F_TYP-006", "Utility room - washable walls", "Finishes", "Floor", "RSF-1-01", "Resilient sheet flooring", "2"],
    ["F_TYP-006", "Utility room - washable walls", "Finishes", "Base", "IC", "Integral cove base", "2"],
    ["F_TYP-006", "Utility room - washable walls", "Finishes", "Wall", "FRP", "Fiberglass reinforced panel", "2"],
    ["F_TYP-006", "Utility room - washable walls", "Finishes", "Ceiling Material", "GB", "Gypsum board ceiling", "2"],
    ["F_TYP-006", "Utility room - washable walls", "Finishes", "Ceiling Finish", "PT", "Paint finish", "2"],
]


ROOM_TYPE_ROWS = [
    ["Level: Name", "Dept", "Room Number", "Room Name", "Program area", "Drawn area", "SDP Reference", "Comments"],
    ["LEVEL 01", "S01", "P01-A-001", "Exam room", "120", "118", "Exam Room, Standard 12/A-100", ""],
    ["LEVEL 01", "S01", "P01-A-002", "Exam room", "120", "121", "Exam Room, Standard 12/A-100", ""],
    ["LEVEL 01", "S01", "P01-A-003", "Exam room bariatric", "150", "148", "Exam Room, Bariatric 13/A-101", ""],
    ["LEVEL 01", "S02", "P01-B-001", "Office", "110", "108", "Office 110 SF 1/B-200", ""],
    ["LEVEL 01", "S02", "P01-B-002", "Office", "110", "112", "Office 110 SF 1/B-200", ""],
    ["LEVEL 01", "S02", "P01-B-003", "Staff washroom", "70", "72", "Staff WR, Universal 2/C-300", ""],
    ["LEVEL 02", "S03", "P02-C-001", "Clean supply", "80", "82", "Clean Supply Room 80 SF 3/C-301", ""],
    ["LEVEL 02", "S03", "P02-C-002", "Medication room", "100", "99", "Medication Room 4/D-400", ""],
    ["LEVEL 02", "S03", "P02-C-003", "Equipment alcove", "45", "47", "Alcove, Equipment 5/E-500", ""],
    ["LEVEL 02", "S04", "P02-D-001", "Waiting area", "250", "252", "Waiting Area 250 SF 6/F-600", ""],
    ["LEVEL 02", "S04", "P02-D-002", "Quiet room", "95", "93", "Quiet Room 7/G-700", ""],
    ["LEVEL 02", "S04", "P02-D-003", "Unassigned room", "80", "80", "", "Blank reference retained for skipped-row count"],
    ["LEVEL 03", "S05", "P03-E-001", "Team station", "180", "177", "Workstation, Team 8/H-800", ""],
    ["LEVEL 03", "S05", "P03-E-002", "Team station", "180", "182", "Workstation, Team 8/H-800", ""],
    ["LEVEL 03", "S05", "P03-E-003", "Family consult room", "110", "112", "Quiet/Family/Consultation Room 9/J-900", ""],
    ["LEVEL 03", "S06", "P03-F-001", "Nourishment room", "90", "88", "Nourish Room 10/K-1000", ""],
    ["LEVEL 03", "S06", "P03-F-002", "Storage room", "70", "71", "Storage Room 70 SF 11/L-1100", ""],
]


COMPARISON_MAIN_ROWS = [
    [
        "Level: Name",
        "Room Number",
        "Room Name",
        "Room Data Template: Reference #",
        "Room Data Template: Code",
        "Finishes: Status",
        "Floor finish RFS",
        "Floor Finish RDS",
        "Base RFS",
        "Base RDS",
        "Wall finish RFS",
        "Wall finish RDS",
        "Ceiling Material RFS",
        "Ceiling Material RDS",
        "Ceiling Finish RFS",
        "Ceiling Finish RDS",
        "Room Finish Comments",
    ],
    ["LEVEL 01", "P01-A-001", "Patient room", "1001", "S-TYP-001", "From sample template", "RSF-1-01", "RSF", "IC", "IC", "PT, SPW", "PT, SPW", "GB", "GB", "PT", "PT", "Expected match"],
    ["LEVEL 01", "P01-A-002", "Patient washroom", "1002", "S-TYP-002", "From sample template", "EPOXY-1-01", "EPOXY", "IC", "IC", "CWT", "CWT", "MRGB", "MRGB", "E/PT", "E/PT", "Expected match"],
    ["LEVEL 01", "P01-A-003", "Office", "1003", "S-TYP-003", "From sample template", "RSF-1-01", "RSF", "IC", "IC", "PT", "PT", "ACT", "ACT", "", "", "ACT ceiling has no finish"],
    ["LEVEL 01", "P01-A-004", "Consult room", "1004", "S-TYP-004", "From sample template", "RSF-1-01", "EPOXY", "IC", "IC", "PT", "PT", "ACT", "ACT", "", "", "Floor mismatch"],
    ["LEVEL 02", "P02-B-001", "Equipment alcove", "1005", "S-TYP-005", "From sample template", "SCON-1-01", "SCON", "SST-B01", "SST-B", "PT", "PT", "EXP", "EXP", "SCON-1-01", "SCON", "Expected match"],
    ["LEVEL 02", "P02-B-002", "Unknown finish room", "1006", "S-TYP-006", "Draft", "LVT-9", "LVT", "IC", "IC", "PT", "PT", "ACT", "ACT", "", "", "No rule example"],
    ["LEVEL 02", "P02-B-003", "Clean utility", "1007", "S-TYP-007", "From sample template", "RSF-1-01", "RSF", "IC", "IC", "FRP", "FRP", "GB", "GB", "PT", "PT", "Expected match"],
    ["LEVEL 02", "P02-B-004", "Waiting area", "1008", "S-TYP-008", "From sample template", "LVT-1-01", "LVT", "RB-1", "RB", "PT, WPT", "PT, WPT", "ACT-S", "ACT-S", "", "", "Expected match"],
    ["LEVEL 03", "P03-C-001", "Family consult", "1009", "S-TYP-009", "From sample template", "RSF-1-01", "RSF", "IC", "IC", "PT, WPT", "PT", "ACT", "ACT", "", "", "Wall finish partial mismatch"],
    ["LEVEL 03", "P03-C-002", "Nourishment room", "1010", "S-TYP-010", "Draft", "LVT-1-01", "LVT", "RB-1", "RB", "FRP", "FRP", "GB", "GB", "PT", "PT", "Expected match"],
    ["LEVEL 03", "P03-C-003", "Draft test room", "1011", "S-TYP-011", "Draft", "RSF-ALT", "RSF", "IC", "IC", "PT", "PT", "ACT", "ACT", "", "", "No rule example"],
]


COMPARISON_RULE_ROWS = [
    ["Source BIM ID", "Allowed Target"],
    ["RSF-1-01", "RSF"],
    ["EPOXY-1-01", "EPOXY"],
    ["SCON-1-01", "SCON"],
    ["IC", "IC"],
    ["SST-B01", "SST-B"],
    ["PT", "PT"],
    ["E/PT", "E/PT"],
    ["SPW", "SPW"],
    ["CWT", "CWT"],
    ["LVT-1-01", "LVT"],
    ["RB-1", "RB"],
    ["WPT", "WPT"],
    ["FRP", "FRP"],
    ["GB", "GB"],
    ["MRGB", "MRGB"],
    ["ACT", "ACT"],
    ["ACT-S", "ACT-S"],
    ["FCM", "FCM"],
    ["EXP", "EXP"],
]


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)

    files = {
        "qaqc-room-report.xlsx": {"Room Report": QAQC_ROOM_ROWS},
        "qaqc-finish-item-list.xlsx": {"Finish Items": FINISH_ITEM_ROWS},
        "template-analyzer-room-report.xlsx": {"Room Finish Data": TEMPLATE_ANALYZER_ROWS},
        "template-finder-item-list.xlsx": {"Template Items": TEMPLATE_FINDER_ROWS},
        "room-type-analyzer-sdp-reference.xlsx": {"SDP Reference": ROOM_TYPE_ROWS},
        "column-comparator-main.xlsx": {"Room Finish Comparison": COMPARISON_MAIN_ROWS},
        "column-comparator-rules.xlsx": {"Rules": COMPARISON_RULE_ROWS},
    }

    for filename, sheets in files.items():
        write_xlsx(OUT / filename, sheets)
        print(f"Wrote {OUT / filename}")


if __name__ == "__main__":
    main()
