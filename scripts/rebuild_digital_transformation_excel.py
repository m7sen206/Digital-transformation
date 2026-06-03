#!/usr/bin/env python3
"""Rebuild the operational Digital Transformation Excel workbook from Markdown sources.

The workbook is intentionally not tracked in git because it is a binary artifact.
This script uses only Python standard-library modules and writes a populated .xlsx
file to outputs/digital_transformation_6_month_plan.xlsx.
"""
from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape
import re

ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / "docs"
OUT = ROOT / "outputs" / "digital_transformation_6_month_plan.xlsx"

SHEET_HEADERS = {
    "Roadmap": ["الشهر", "الفترة", "الهدف الرئيسي", "المالك", "محاور العمل", "المخرجات المطلوبة", "الاجتماعات المهمة", "الاعتمادات المطلوبة", "المخاطر المحتملة", "مؤشرات نجاح الشهر"],
    "Weekly Plan": ["الأسبوع", "الفترة التقريبية", "الهدف", "المهام القابلة للمتابعة", "المسؤول", "الأقسام المشاركة", "المخرجات", "حالة الاعتماد المطلوبة", "ملاحظات"],
    "Risk Register": ["رقم الخطر", "الوصف", "السبب", "الأثر", "الاحتمالية", "الشدة", "درجة الخطورة", "خطة المعالجة", "المالك", "تاريخ/تكرار المراجعة"],
    "KPI Dashboard Plan": ["اللوحة", "الهدف", "الجمهور المستهدف", "المؤشرات الرئيسية القابلة للقياس", "مصدر البيانات", "تكرار التحديث", "المسؤول عن البيانات", "شكل التقرير المقترح", "معيار جودة/قبول"],
    "Automation Backlog": ["المبادرة", "الوصف", "الفائدة", "الصعوبة", "الأولوية", "الأدوات المقترحة", "البيانات المطلوبة", "المخاطر", "التوقيت"],
    "Meeting Cadence": ["الاجتماع/التقرير", "التكرار", "الحضور", "الهدف", "المخرج"],
}


def parse_markdown_tables(path: Path) -> list[tuple[list[str], list[list[str]]]]:
    """Return all simple pipe tables from a Markdown file."""
    lines = path.read_text(encoding="utf-8").splitlines()
    tables: list[tuple[list[str], list[list[str]]]] = []
    i = 0
    separator = re.compile(r"^\|[\s:\-\|]+\|$")
    while i < len(lines):
        line = lines[i].strip()
        if line.startswith("|") and line.endswith("|") and i + 1 < len(lines) and separator.match(lines[i + 1].strip()):
            raw_rows: list[list[str]] = []
            while i < len(lines) and lines[i].strip().startswith("|") and lines[i].strip().endswith("|"):
                cells = [cell.strip().replace("<br>", "\n") for cell in lines[i].strip().strip("|").split("|")]
                raw_rows.append(cells)
                i += 1
            if len(raw_rows) >= 2:
                tables.append((raw_rows[0], raw_rows[2:]))
            continue
        i += 1
    return tables


def table_with_headers(path: Path, headers: list[str]) -> list[list[str]]:
    for found_headers, rows in parse_markdown_tables(path):
        if found_headers == headers:
            return [found_headers] + rows
    raise RuntimeError(f"Required table was not found in {path}: {headers}")


def col_letter(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def table_ref(start_row: int, end_row: int, cols: int) -> str:
    return f"A{start_row}:{col_letter(cols)}{end_row}"


def safe_table_name(name: str, table_id: int) -> str:
    safe = re.sub(r"[^A-Za-z0-9_]", "_", name)
    safe = re.sub(r"_+", "_", safe).strip("_") or f"Table{table_id}"
    if safe[0].isdigit():
        safe = "T_" + safe
    return safe


def cell_style(row_index: int, row: list[str], table_ranges: list[tuple[int, int, int, str]]) -> int:
    if len(row) == 1 and row and row[0]:
        return 3  # section title
    for start, _end, _cols, _name in table_ranges:
        if row_index == start:
            return 1  # table header
    return 2  # body


def worksheet_xml(rows: list[list[str]], table_ranges: list[tuple[int, int, int, str]]) -> str:
    max_cols = max((len(row) for row in rows), default=1)
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheetViews><sheetView rightToLeft="1" workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>',
        '<sheetFormatPr defaultRowHeight="18"/>',
        '<cols>',
    ]
    for col in range(1, max_cols + 1):
        width = 14 if col <= 2 else 28
        if col >= 4:
            width = 36
        parts.append(f'<col min="{col}" max="{col}" width="{width}" customWidth="1"/>')
    parts.append('</cols><sheetData>')
    for row_index, row in enumerate(rows, 1):
        height = ' ht="34" customHeight="1"' if row_index == 1 or (len(row) == 1 and row and row[0]) else ' ht="54" customHeight="1"'
        parts.append(f'<row r="{row_index}"{height}>')
        for col_index, value in enumerate(row, 1):
            if value in (None, ""):
                continue
            ref = f"{col_letter(col_index)}{row_index}"
            style = cell_style(row_index, row, table_ranges)
            parts.append(f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{escape(str(value))}</t></is></c>')
        parts.append('</row>')
    parts.append('</sheetData>')
    if table_ranges:
        start, end, cols, _name = table_ranges[0]
        parts.append(f'<autoFilter ref="{table_ref(start, end, cols)}"/>')
        parts.append(f'<tableParts count="{len(table_ranges)}">')
        for rel_id in range(1, len(table_ranges) + 1):
            parts.append(f'<tablePart r:id="rId{rel_id}"/>')
        parts.append('</tableParts>')
    parts.append('</worksheet>')
    return "".join(parts)


def table_xml(table_id: int, display_name: str, start: int, end: int, cols: int, headers: list[str]) -> str:
    ref = table_ref(start, end, cols)
    name = safe_table_name(display_name, table_id)
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="{table_id}" '
        f'name="{name}" displayName="{name}" ref="{ref}" totalsRowShown="0">',
        f'<autoFilter ref="{ref}"/>',
        f'<tableColumns count="{cols}">',
    ]
    for col in range(cols):
        header = headers[col] if col < len(headers) else f"Column{col + 1}"
        parts.append(f'<tableColumn id="{col + 1}" name="{escape(str(header))}"/>')
    parts.append('</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>')
    return "".join(parts)


def build_data() -> list[tuple[str, list[list[str]], list[tuple[int, int, int, str]]]]:
    roadmap = table_with_headers(DOCS / "02-six-month-roadmap-jul-dec-2026.md", SHEET_HEADERS["Roadmap"])
    weekly = table_with_headers(DOCS / "03-weekly-action-plan.md", SHEET_HEADERS["Weekly Plan"])
    raci_headers, raci_rows = parse_markdown_tables(DOCS / "10-raci-matrix.md")[0]
    raci = [raci_headers] + raci_rows
    risks = table_with_headers(DOCS / "11-risk-register.md", SHEET_HEADERS["Risk Register"])
    kpis = table_with_headers(DOCS / "07-kpi-dashboard-plan.md", SHEET_HEADERS["KPI Dashboard Plan"])
    automations = table_with_headers(DOCS / "09-automation-and-ai-backlog.md", SHEET_HEADERS["Automation Backlog"])
    meetings = table_with_headers(DOCS / "12-meeting-cadence-and-reporting.md", SHEET_HEADERS["Meeting Cadence"])

    odoo_rows: list[list[str]] = []
    odoo_ranges: list[tuple[int, int, int, str]] = []
    odoo_titles = ["السجلات الإلزامية", "البوابات الرقابية Stage Gates", "نموذج التقرير الأسبوعي عن Odoo"]
    for title, (headers, rows) in zip(odoo_titles, parse_markdown_tables(DOCS / "04-odoo-project-control-plan.md")[:3]):
        start = len(odoo_rows) + 2
        odoo_rows.append([title])
        odoo_rows.append(headers)
        odoo_rows.extend(rows)
        end = len(odoo_rows)
        odoo_ranges.append((start, end, len(headers), title))
        odoo_rows.append([])

    sheets = [
        ("Roadmap", roadmap, [(1, len(roadmap), len(roadmap[0]), "RoadmapTable")]),
        ("Weekly Plan", weekly, [(1, len(weekly), len(weekly[0]), "WeeklyPlanTable")]),
        ("RACI", raci, [(1, len(raci), len(raci[0]), "RACITable")]),
        ("Risk Register", risks, [(1, len(risks), len(risks[0]), "RiskRegisterTable")]),
        ("KPI Dashboard Plan", kpis, [(1, len(kpis), len(kpis[0]), "KPIDashboardPlanTable")]),
        ("Automation Backlog", automations, [(1, len(automations), len(automations[0]), "AutomationBacklogTable")]),
        ("Meeting Cadence", meetings, [(1, len(meetings), len(meetings[0]), "MeetingCadenceTable")]),
        ("Odoo Control Plan", odoo_rows, odoo_ranges),
    ]

    minimum_rows = {
        "Roadmap": 7,
        "Weekly Plan": 27,
        "RACI": 11,
        "Risk Register": 26,
        "KPI Dashboard Plan": 11,
        "Automation Backlog": 11,
        "Meeting Cadence": 8,
        "Odoo Control Plan": 25,
    }
    for name, rows, _ranges in sheets:
        non_empty = sum(1 for row in rows if any(cell not in (None, "") for cell in row))
        if non_empty < minimum_rows[name]:
            raise RuntimeError(f"{name} has only {non_empty} non-empty rows")
    return sheets


STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Arial"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
    <font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF5B9BD5"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD9E2F3"/></left><right style="thin"><color rgb="FFD9E2F3"/></right><top style="thin"><color rgb="FFD9E2F3"/></top><bottom style="thin"><color rgb="FFD9E2F3"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="top" wrapText="1" readingOrder="2"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1" readingOrder="2"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="top" wrapText="1" readingOrder="2"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center" wrapText="1" readingOrder="2"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''


def write_workbook(sheets: list[tuple[str, list[list[str]], list[tuple[int, int, int, str]]]]) -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    content_overrides: list[str] = []
    sheet_xmls: list[tuple[int, str]] = []
    sheet_rels: list[tuple[int, str]] = []
    table_xmls: list[tuple[int, str]] = []
    table_id = 1

    for sheet_index, (sheet_name, rows, table_ranges) in enumerate(sheets, 1):
        sheet_xmls.append((sheet_index, worksheet_xml(rows, table_ranges)))
        rel_parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        for local_rel_id, (start, end, cols, table_name) in enumerate(table_ranges, 1):
            rel_parts.append(f'<Relationship Id="rId{local_rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table{table_id}.xml"/>')
            table_xmls.append((table_id, table_xml(table_id, table_name, start, end, cols, rows[start - 1])))
            content_overrides.append(f'<Override PartName="/xl/tables/table{table_id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>')
            table_id += 1
        rel_parts.append('</Relationships>')
        sheet_rels.append((sheet_index, "".join(rel_parts)))

    with ZipFile(OUT, "w", ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, len(sheets) + 1))
            + "".join(content_overrides)
            + "</Types>",
        )
        archive.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(sheets) + 1))
            + f'<Relationship Id="rId{len(sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            + "</Relationships>",
        )
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView rightToLeft="1"/></bookViews><sheets>'
            + "".join(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>' for i, (name, _rows, _ranges) in enumerate(sheets, 1))
            + "</sheets></workbook>",
        )
        archive.writestr("xl/styles.xml", STYLES_XML)
        for index, xml in sheet_xmls:
            archive.writestr(f"xl/worksheets/sheet{index}.xml", xml)
        for index, xml in sheet_rels:
            archive.writestr(f"xl/worksheets/_rels/sheet{index}.xml.rels", xml)
        for index, xml in table_xmls:
            archive.writestr(f"xl/tables/table{index}.xml", xml)


def main() -> None:
    sheets = build_data()
    write_workbook(sheets)
    print(f"Generated {OUT.relative_to(ROOT)}")
    for name, rows, _ranges in sheets:
        non_empty = sum(1 for row in rows if any(cell not in (None, "") for cell in row))
        print(f"{name}: {non_empty} non-empty rows")


if __name__ == "__main__":
    main()
