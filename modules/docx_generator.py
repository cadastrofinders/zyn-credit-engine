"""
ZYN Capital — Gerador de MAC (.docx)
Memorando de Análise de Crédito com identidade visual ZYN.

Dependência: python-docx >= 1.1
"""

from __future__ import annotations

import locale
import os
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple, Union

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor, Emu
from docx.table import _Cell

# ---------------------------------------------------------------------------
# Color Palette
# ---------------------------------------------------------------------------
NAVY_HEX = "223040"
DARK_SLATE_HEX = "3A4F63"
MUTED_GRAY_HEX = "8B9197"
LIGHT_LABEL_HEX = "A8B4C0"
LIGHT_GRAY_BG_HEX = "F2F4F6"
LIGHT_GRAY_BG2_HEX = "E8ECF0"
BORDER_GRAY_HEX = "D0D5DA"
GREEN_BG_HEX = "EAF4EE"
DARK_GREEN_HEX = "1E6B42"
YELLOW_BG_HEX = "FEF9E7"
DARK_GOLD_HEX = "7D6608"
RED_BG_HEX = "FDECEA"
DARK_RED_HEX = "922B21"
DARK_BLUE_HEX = "1A5276"
WHITE_HEX = "FFFFFF"

NAVY = RGBColor(34, 48, 64)
DARK_SLATE = RGBColor(58, 79, 99)
MUTED_GRAY = RGBColor(139, 145, 151)
LIGHT_LABEL = RGBColor(168, 180, 192)
DARK_GREEN = RGBColor(30, 107, 66)
DARK_GOLD = RGBColor(125, 102, 8)
DARK_RED = RGBColor(146, 43, 33)
DARK_BLUE = RGBColor(26, 82, 118)
WHITE = RGBColor(255, 255, 255)

# KPI strip column colors (background)
KPI_COLORS = [NAVY_HEX, DARK_SLATE_HEX, DARK_GREEN_HEX, DARK_GOLD_HEX, DARK_RED_HEX]

# ---------------------------------------------------------------------------
# Section titles
# ---------------------------------------------------------------------------
SECTION_TITLES = {
    "tomador": "1. TOMADOR",
    "patrimonio": "2. PATRIMÔNIO",
    "producao": "3. PRODUÇÃO E ATIVIDADE OPERACIONAL",
    "capital": "4. CAPITAL E ESTRUTURA DA OPERAÇÃO",
    "operacao": "5. OPERAÇÃO",
    "pagamento": "6. PAGAMENTO — CRONOGRAMA FINANCEIRO",
    "onus": "7. ÔNUS E GARANTIAS",
    "riscos": "8. RISCOS",
    "covenants": "9. COVENANTS E CONDICIONANTES",
}

# Alert type configs: (fill_hex, border_hex, title_color_hex)
ALERT_STYLES = {
    "warning": (YELLOW_BG_HEX, DARK_GOLD_HEX, DARK_GOLD_HEX),
    "info": (LIGHT_GRAY_BG2_HEX, DARK_BLUE_HEX, DARK_BLUE_HEX),
    "risk": (RED_BG_HEX, DARK_RED_HEX, DARK_RED_HEX),
    "positive": (GREEN_BG_HEX, DARK_GREEN_HEX, DARK_GREEN_HEX),
}

# Risk level badges: (text_hex, fill_hex)
RISK_LEVEL_STYLES = {
    "ALTO": (DARK_RED_HEX, RED_BG_HEX),
    "MÉDIO": (DARK_GOLD_HEX, YELLOW_BG_HEX),
    "BAIXO": (DARK_GREEN_HEX, GREEN_BG_HEX),
}

# Table total width in dxa
TABLE_WIDTH_DXA = 9746


# ═══════════════════════════════════════════════════════════════════════════
# Helper Functions
# ═══════════════════════════════════════════════════════════════════════════

def _hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    """Convert hex color string (with or without #) to RGB tuple."""
    h = hex_color.lstrip("#")
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _set_cell_shading(cell: _Cell, hex_color: str) -> None:
    """Apply solid fill to a table cell."""
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shading)


def _set_cell_borders(
    cell: _Cell,
    top: Optional[Tuple[str, int, str]] = None,
    bottom: Optional[Tuple[str, int, str]] = None,
    left: Optional[Tuple[str, int, str]] = None,
    right: Optional[Tuple[str, int, str]] = None,
) -> None:
    """Set individual cell borders. Each arg is (val, sz, color) or None."""
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for side, spec in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        if spec is not None:
            val, sz, color = spec
            el = OxmlElement(f"w:{side}")
            el.set(qn("w:val"), val)
            el.set(qn("w:sz"), str(sz))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), color)
            borders.append(el)
    tc_pr.append(borders)


def _set_cell_no_borders(cell: _Cell) -> None:
    """Remove all borders from a cell."""
    _set_cell_borders(
        cell,
        top=("none", 0, WHITE_HEX),
        bottom=("none", 0, WHITE_HEX),
        left=("none", 0, WHITE_HEX),
        right=("none", 0, WHITE_HEX),
    )


def _set_cell_margins(cell: _Cell, top: int, left: int, bottom: int, right: int) -> None:
    """Set cell margins in dxa (twentieths of a point)."""
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = OxmlElement("w:tcMar")
    for side, val in [("top", top), ("start", left), ("bottom", bottom), ("end", right)]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        margins.append(el)
    tc_pr.append(margins)


def _set_run_font(
    run,
    size_pt: float,
    color_hex: str,
    bold: bool = False,
    italic: bool = False,
) -> None:
    """Configure a run's font properties (no explicit font name — inherits theme)."""
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size_pt)
    r, g, b = _hex_to_rgb(color_hex)
    run.font.color.rgb = RGBColor(r, g, b)


def _set_table_width(table, width_dxa: int = TABLE_WIDTH_DXA) -> None:
    """Set table width explicitly via XML."""
    tbl = table._tbl
    tbl_pr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:w"), str(width_dxa))
    tw.set(qn("w:type"), "dxa")
    # Remove existing tblW
    for existing in tbl_pr.findall(qn("w:tblW")):
        tbl_pr.remove(existing)
    tbl_pr.append(tw)


def _remove_table_borders(table) -> None:
    """Remove all table-level borders."""
    tbl = table._tbl
    tbl_pr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
    borders = OxmlElement("w:tblBorders")
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), WHITE_HEX)
        borders.append(el)
    # Remove existing
    for existing in tbl_pr.findall(qn("w:tblBorders")):
        tbl_pr.remove(existing)
    tbl_pr.append(borders)


def _set_col_width(cell: _Cell, width_dxa: int) -> None:
    """Set column width on a cell."""
    tc_pr = cell._tc.get_or_add_tcPr()
    tw = OxmlElement("w:tcW")
    tw.set(qn("w:w"), str(width_dxa))
    tw.set(qn("w:type"), "dxa")
    for existing in tc_pr.findall(qn("w:tcW")):
        tc_pr.remove(existing)
    tc_pr.append(tw)


def _clear_cell(cell: _Cell) -> None:
    """Remove all default paragraphs from a cell, leaving it empty."""
    for p in cell.paragraphs:
        p_element = p._element
        p_element.getparent().remove(p_element)


def _add_para_to_cell(cell: _Cell, text: str = "", **run_kwargs) -> None:
    """Add a paragraph with a single run to a cell."""
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    if text:
        run = p.add_run(text)
        if run_kwargs:
            _set_run_font(run, **run_kwargs)
    return p


# ═══════════════════════════════════════════════════════════════════════════
# Formatting helpers
# ═══════════════════════════════════════════════════════════════════════════

def _format_currency(value) -> str:
    """Format a number as Brazilian currency: R$ 50.000.000"""
    if value is None:
        return "N/D"
    try:
        v = float(value)
        # Format with dots as thousand separators, no decimals
        formatted = f"{v:,.0f}".replace(",", ".")
        return f"R$ {formatted}"
    except (ValueError, TypeError):
        return str(value)


def _format_pct(value) -> str:
    """Format as percentage: 12,5%"""
    if value is None:
        return "N/D"
    try:
        v = float(value)
        formatted = f"{v:.1f}".replace(".", ",")
        return f"{formatted}%"
    except (ValueError, TypeError):
        return str(value)


def _safe(val, default: str = "N/D") -> str:
    """Return string representation or default if None/empty."""
    if val is None or val == "":
        return default
    return str(val)


# ═══════════════════════════════════════════════════════════════════════════
# Document Component Builders
# ═══════════════════════════════════════════════════════════════════════════

def _add_footer(doc) -> None:
    """Add footer: left 'ZYN CAPITAL © 2026 · Documento confidencial...' right 'Pág. X'."""
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False

    # Remove default paragraph if any
    for p in footer.paragraphs:
        p.clear()

    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

    # Add right-aligned tab stop at 9746 dxa = 6.496 inches = ~4.94 inches from left margin
    # 9746 dxa is the full table width; we set tab at page width
    tab_stops = p.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Emu(9746 * 635))  # dxa to EMU: 1 dxa = 635 EMU

    # Left text
    year = datetime.now().year
    left_text = f"ZYN CAPITAL  \u00a9  {year}  \u00b7  Documento confidencial \u2014 uso exclusivo interno"
    run_left = p.add_run(left_text)
    _set_run_font(run_left, size_pt=7, color_hex=MUTED_GRAY_HEX, bold=False)

    # Tab
    run_tab = p.add_run("\t")
    _set_run_font(run_tab, size_pt=7, color_hex=MUTED_GRAY_HEX)

    # "Pág. " text
    run_pag = p.add_run("Pág. ")
    _set_run_font(run_pag, size_pt=7, color_hex=MUTED_GRAY_HEX)

    # PAGE field
    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = " PAGE "

    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")

    # Create a run element for the field
    run_field = p.add_run()
    _set_run_font(run_field, size_pt=7, color_hex=MUTED_GRAY_HEX)
    run_field._r.append(fld_char_begin)

    run_instr = p.add_run()
    _set_run_font(run_instr, size_pt=7, color_hex=MUTED_GRAY_HEX)
    run_instr._r.append(instr_text)

    run_end = p.add_run()
    _set_run_font(run_end, size_pt=7, color_hex=MUTED_GRAY_HEX)
    run_end._r.append(fld_char_end)


def _add_kpi_strip(doc, kpis: Dict[str, Tuple[str, str]]) -> None:
    """
    Add a 5-column KPI strip table.
    kpis: dict with keys like 'principal', 'prazo', 'taxa_all_in', 'juros_ano1', 'total_juros'
          each value is (label, value_str)
    """
    labels = ["PRINCIPAL", "PRAZO", "TAXA ALL-IN", "JUROS ANO 1", "TOTAL JUROS"]
    keys = ["principal", "prazo", "taxa_all_in", "juros_ano1", "total_juros"]

    table = doc.add_table(rows=2, cols=5)
    _set_table_width(table)
    _remove_table_borders(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    col_width = TABLE_WIDTH_DXA // 5

    for col_idx in range(5):
        color_hex = KPI_COLORS[col_idx]
        key = keys[col_idx]
        value_str = kpis.get(key, ("", "N/D"))
        if isinstance(value_str, tuple):
            value_str = value_str[1]

        # Label row (row 0)
        cell_label = table.cell(0, col_idx)
        _set_cell_shading(cell_label, color_hex)
        _set_cell_no_borders(cell_label)
        _set_cell_margins(cell_label, top=60, left=80, bottom=0, right=60)
        _set_col_width(cell_label, col_width)
        cell_label.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        _clear_cell(cell_label)
        p_label = cell_label.add_paragraph()
        p_label.paragraph_format.space_before = Pt(0)
        p_label.paragraph_format.space_after = Pt(0)
        run_label = p_label.add_run(labels[col_idx])
        _set_run_font(run_label, size_pt=6.5, color_hex=LIGHT_LABEL_HEX, bold=False)

        # Value row (row 1)
        cell_val = table.cell(1, col_idx)
        _set_cell_shading(cell_val, color_hex)
        _set_cell_no_borders(cell_val)
        _set_cell_margins(cell_val, top=60, left=80, bottom=0, right=60)
        _set_col_width(cell_val, col_width)
        cell_val.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        _clear_cell(cell_val)
        p_val = cell_val.add_paragraph()
        p_val.paragraph_format.space_before = Pt(0)
        p_val.paragraph_format.space_after = Pt(0)
        run_val = p_val.add_run(_safe(value_str))
        _set_run_font(run_val, size_pt=6.5, color_hex=WHITE_HEX, bold=True)

    # Small spacer after KPI strip
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(2)
    spacer.paragraph_format.space_after = Pt(6)


def _add_section_heading(doc, title: str) -> None:
    """Full-width navy bar with bold white text — implemented as a single-cell table."""
    table = doc.add_table(rows=1, cols=1)
    _set_table_width(table)
    _remove_table_borders(table)

    cell = table.cell(0, 0)
    _set_cell_shading(cell, NAVY_HEX)
    _set_cell_margins(cell, top=140, left=200, bottom=140, right=200)
    _set_col_width(cell, TABLE_WIDTH_DXA)

    _clear_cell(cell)
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(title)
    _set_run_font(run, size_pt=10, color_hex=WHITE_HEX, bold=True)


def _add_subsection_heading(doc, title: str) -> None:
    """Light gray bar with navy text — implemented as a single-cell table."""
    table = doc.add_table(rows=1, cols=1)
    _set_table_width(table)
    _remove_table_borders(table)

    cell = table.cell(0, 0)
    _set_cell_shading(cell, LIGHT_GRAY_BG2_HEX)
    _set_cell_margins(cell, top=100, left=180, bottom=100, right=180)
    _set_col_width(cell, TABLE_WIDTH_DXA)

    _clear_cell(cell)
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(title)
    _set_run_font(run, size_pt=8.5, color_hex=NAVY_HEX, bold=True)


def _add_kv_table(doc, rows: List[Tuple[str, str]]) -> None:
    """
    Key-value 2-column table.
    Label column: 2600 dxa, bold 8pt dark slate.
    Value column: 7146 dxa, 8pt.
    Alternating label fill.
    """
    table = doc.add_table(rows=len(rows), cols=2)
    _set_table_width(table)
    _remove_table_borders(table)

    for i, (label, value) in enumerate(rows):
        fill = LIGHT_GRAY_BG_HEX if i % 2 == 0 else WHITE_HEX
        cell_label = table.cell(i, 0)
        cell_value = table.cell(i, 1)

        _set_col_width(cell_label, 2600)
        _set_col_width(cell_value, 7146)

        _set_cell_shading(cell_label, fill)
        _set_cell_shading(cell_value, WHITE_HEX)

        # Borders
        border_spec = ("single", 4, BORDER_GRAY_HEX)
        _set_cell_borders(cell_label, top=border_spec, bottom=border_spec, left=border_spec, right=border_spec)
        _set_cell_borders(cell_value, top=border_spec, bottom=border_spec, left=border_spec, right=border_spec)

        _set_cell_margins(cell_label, top=60, left=120, bottom=60, right=80)
        _set_cell_margins(cell_value, top=60, left=120, bottom=60, right=80)

        cell_label.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell_value.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        _clear_cell(cell_label)
        p_l = cell_label.add_paragraph()
        p_l.paragraph_format.space_before = Pt(0)
        p_l.paragraph_format.space_after = Pt(0)
        run_l = p_l.add_run(_safe(label))
        _set_run_font(run_l, size_pt=8, color_hex=DARK_SLATE_HEX, bold=True)

        _clear_cell(cell_value)
        p_v = cell_value.add_paragraph()
        p_v.paragraph_format.space_before = Pt(0)
        p_v.paragraph_format.space_after = Pt(0)
        run_v = p_v.add_run(_safe(value))
        _set_run_font(run_v, size_pt=8, color_hex=NAVY_HEX)


def _add_data_table(
    doc,
    headers: List[str],
    rows: List[List[str]],
    col_widths: Optional[List[int]] = None,
    numeric_cols: Optional[List[int]] = None,
) -> None:
    """
    Data table with navy header and alternating row fills.
    numeric_cols: indices of columns that should be bold navy.
    """
    if not headers:
        return

    n_cols = len(headers)
    if col_widths is None:
        col_widths = [TABLE_WIDTH_DXA // n_cols] * n_cols
    if numeric_cols is None:
        numeric_cols = []

    table = doc.add_table(rows=1 + len(rows), cols=n_cols)
    _set_table_width(table)
    _remove_table_borders(table)

    # Header row
    for col_idx, header_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        _set_cell_shading(cell, NAVY_HEX)
        _set_cell_margins(cell, top=60, left=80, bottom=60, right=80)
        _set_col_width(cell, col_widths[col_idx])
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        _clear_cell(cell)
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(header_text)
        _set_run_font(run, size_pt=7.5, color_hex=WHITE_HEX, bold=True)

    # Data rows
    for row_idx, row_data in enumerate(rows):
        fill = LIGHT_GRAY_BG_HEX if row_idx % 2 == 0 else WHITE_HEX
        for col_idx in range(n_cols):
            cell = table.cell(row_idx + 1, col_idx)
            _set_cell_shading(cell, fill)
            _set_cell_margins(cell, top=40, left=80, bottom=40, right=80)
            _set_col_width(cell, col_widths[col_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            text = _safe(row_data[col_idx] if col_idx < len(row_data) else "")

            _clear_cell(cell)
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(text)
            if col_idx in numeric_cols:
                _set_run_font(run, size_pt=8, color_hex=NAVY_HEX, bold=True)
            else:
                _set_run_font(run, size_pt=8, color_hex=NAVY_HEX)


def _add_alert_box(doc, title: str, body: str, alert_type: str = "warning") -> None:
    """
    Alert box — single-cell table with thick left border and solid fill.
    alert_type: 'warning', 'info', 'risk', 'positive'
    """
    fill_hex, border_hex, title_hex = ALERT_STYLES.get(alert_type, ALERT_STYLES["warning"])

    table = doc.add_table(rows=1, cols=1)
    _set_table_width(table)
    _remove_table_borders(table)

    cell = table.cell(0, 0)
    _set_cell_shading(cell, fill_hex)
    _set_cell_margins(cell, top=120, left=200, bottom=120, right=200)
    _set_col_width(cell, TABLE_WIDTH_DXA)

    # Thick left border only
    _set_cell_borders(
        cell,
        left=("single", 20, border_hex),
        top=("none", 0, fill_hex),
        bottom=("none", 0, fill_hex),
        right=("none", 0, fill_hex),
    )

    _clear_cell(cell)

    # Title paragraph
    p_title = cell.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    run_title = p_title.add_run(title)
    _set_run_font(run_title, size_pt=8.5, color_hex=title_hex, bold=True)

    # Body paragraph
    p_body = cell.add_paragraph()
    p_body.paragraph_format.space_before = Pt(0)
    p_body.paragraph_format.space_after = Pt(0)
    run_body = p_body.add_run(_safe(body))
    _set_run_font(run_body, size_pt=8, color_hex=NAVY_HEX)


def _add_risk_table(doc, risks: List[Dict[str, str]]) -> None:
    """
    Risk matrix table with columns: RISCO, NÍVEL, PROB., MITIGANTE.
    Level badges with colored text and cell fill.
    """
    headers = ["RISCO", "NÍVEL", "PROB.", "MITIGANTE"]
    col_widths = [2800, 1200, 900, 4846]

    n_rows = len(risks)
    table = doc.add_table(rows=1 + n_rows, cols=4)
    _set_table_width(table)
    _remove_table_borders(table)

    # Header
    for col_idx, h in enumerate(headers):
        cell = table.cell(0, col_idx)
        _set_cell_shading(cell, NAVY_HEX)
        _set_cell_margins(cell, top=60, left=80, bottom=60, right=80)
        _set_col_width(cell, col_widths[col_idx])
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        _clear_cell(cell)
        p = cell.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        _set_run_font(run, size_pt=7.5, color_hex=WHITE_HEX, bold=True)

    # Data rows
    for row_idx, risk in enumerate(risks):
        fill = LIGHT_GRAY_BG_HEX if row_idx % 2 == 0 else WHITE_HEX
        nivel = _safe(risk.get("nivel", risk.get("nível", "")), "N/D").upper()

        row_data = [
            _safe(risk.get("risco", risk.get("descricao", "")), "N/D"),
            nivel,
            _safe(risk.get("probabilidade", risk.get("prob", "")), "N/D"),
            _safe(risk.get("mitigante", "")),
        ]

        for col_idx in range(4):
            cell = table.cell(row_idx + 1, col_idx)
            _set_cell_margins(cell, top=40, left=80, bottom=40, right=80)
            _set_col_width(cell, col_widths[col_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            _clear_cell(cell)
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)

            text = row_data[col_idx]

            if col_idx == 1:
                # Level badge: colored text + colored cell fill
                text_hex, badge_fill = RISK_LEVEL_STYLES.get(nivel, (NAVY_HEX, fill))
                _set_cell_shading(cell, badge_fill)
                run = p.add_run(text)
                _set_run_font(run, size_pt=8, color_hex=text_hex, bold=True)
            else:
                _set_cell_shading(cell, fill)
                run = p.add_run(text)
                _set_run_font(run, size_pt=8, color_hex=NAVY_HEX)


# ═══════════════════════════════════════════════════════════════════════════
# Section Renderers
# ═══════════════════════════════════════════════════════════════════════════

def _render_section_tomador(doc, data: dict, parametros: dict) -> None:
    """Render section 1 — Tomador."""
    _add_section_heading(doc, SECTION_TITLES["tomador"])

    ficha = [
        ("Razão Social", _safe(parametros.get("tomador"))),
        ("CNPJ", _safe(parametros.get("cnpj"))),
        ("Tipo Operação", _safe(parametros.get("tipo_operacao"))),
        ("Instrumento", _safe(parametros.get("instrumento"))),
        ("Sócio Responsável ZYN", _safe(parametros.get("socio_responsavel"))),
    ]

    # Add extra fields from data if available
    extra_fields = data.get("dados", data.get("ficha", {}))
    if isinstance(extra_fields, dict):
        for k, v in extra_fields.items():
            if k not in ("analise", "rating_secao", "flags"):
                ficha.append((k, _safe(v)))

    _add_subsection_heading(doc, "Ficha do Tomador")
    _add_kv_table(doc, ficha)

    analise_text = data.get("analise", "")
    if analise_text:
        _add_alert_box(doc, "Análise", analise_text, "info")

    _render_flags(doc, data.get("flags", []))


def _render_section_patrimonio(doc, data: dict) -> None:
    """Render section 2 — Patrimônio."""
    _add_section_heading(doc, SECTION_TITLES["patrimonio"])

    analise_text = data.get("analise", "")
    if analise_text:
        _add_alert_box(doc, "Análise Patrimonial", analise_text, "info")

    # Tabular data if available
    tabela = data.get("tabela", data.get("dados", []))
    if isinstance(tabela, list) and tabela:
        if isinstance(tabela[0], dict):
            headers = list(tabela[0].keys())
            rows = [[_safe(item.get(h)) for h in headers] for item in tabela]
            _add_data_table(doc, headers, rows)
        elif isinstance(tabela[0], list):
            _add_data_table(doc, tabela[0], tabela[1:])

    _render_flags(doc, data.get("flags", []))


def _render_section_producao(doc, data: dict) -> None:
    """Render section 3 — Produção e Atividade Operacional."""
    _add_section_heading(doc, SECTION_TITLES["producao"])

    analise_text = data.get("analise", "")
    if analise_text:
        _add_alert_box(doc, "Análise Operacional", analise_text, "info")

    tabela = data.get("tabela", data.get("dados", []))
    if isinstance(tabela, list) and tabela:
        if isinstance(tabela[0], dict):
            headers = list(tabela[0].keys())
            rows = [[_safe(item.get(h)) for h in headers] for item in tabela]
            _add_data_table(doc, headers, rows)

    _render_flags(doc, data.get("flags", []))


def _render_section_capital(doc, data: dict) -> None:
    """Render section 4 — Capital e Estrutura."""
    _add_section_heading(doc, SECTION_TITLES["capital"])

    analise_text = data.get("analise", "")
    if analise_text:
        _add_alert_box(doc, "Análise de Capital", analise_text, "info")

    tabela = data.get("tabela", data.get("dados", []))
    if isinstance(tabela, list) and tabela:
        if isinstance(tabela[0], dict):
            headers = list(tabela[0].keys())
            rows = [[_safe(item.get(h)) for h in headers] for item in tabela]
            _add_data_table(doc, headers, rows)

    _render_flags(doc, data.get("flags", []))


def _render_section_operacao(doc, data: dict, parametros: dict) -> None:
    """Render section 5 — Operação."""
    _add_section_heading(doc, SECTION_TITLES["operacao"])

    ficha = [
        ("Volume", _format_currency(parametros.get("volume"))),
        ("Prazo", f"{_safe(parametros.get('prazo_meses'))} meses"),
        ("Taxa", _safe(parametros.get("taxa"))),
        ("Amortização", _safe(parametros.get("amortizacao"))),
        ("Garantias", ", ".join(parametros.get("garantias", [])) or "N/D"),
    ]

    _add_subsection_heading(doc, "Parâmetros da Operação")
    _add_kv_table(doc, ficha)

    analise_text = data.get("analise", "")
    if analise_text:
        _add_alert_box(doc, "Análise da Operação", analise_text, "info")

    _render_flags(doc, data.get("flags", []))


def _render_section_pagamento(doc, data: dict) -> None:
    """Render section 6 — Pagamento / Cronograma Financeiro."""
    _add_section_heading(doc, SECTION_TITLES["pagamento"])

    analise_text = data.get("analise", "")
    if analise_text:
        _add_alert_box(doc, "Análise de Pagamento", analise_text, "info")

    # Cronograma table
    cronograma = data.get("cronograma", data.get("tabela", data.get("dados", [])))
    if isinstance(cronograma, list) and cronograma:
        if isinstance(cronograma[0], dict):
            headers = list(cronograma[0].keys())
            rows = [[_safe(item.get(h)) for h in headers] for item in cronograma]
            _add_subsection_heading(doc, "Cronograma de Pagamentos")
            _add_data_table(doc, headers, rows)

    _render_flags(doc, data.get("flags", []))


def _render_section_onus(doc, data: dict) -> None:
    """Render section 7 — Ônus e Garantias."""
    _add_section_heading(doc, SECTION_TITLES["onus"])

    analise_text = data.get("analise", "")
    if analise_text:
        _add_alert_box(doc, "Análise de Garantias", analise_text, "info")

    tabela = data.get("tabela", data.get("dados", []))
    if isinstance(tabela, list) and tabela:
        if isinstance(tabela[0], dict):
            headers = list(tabela[0].keys())
            rows = [[_safe(item.get(h)) for h in headers] for item in tabela]
            _add_data_table(doc, headers, rows)

    _render_flags(doc, data.get("flags", []))


def _render_section_riscos(doc, data: dict) -> None:
    """Render section 8 — Riscos."""
    _add_section_heading(doc, SECTION_TITLES["riscos"])

    analise_text = data.get("analise", "")
    if analise_text:
        _add_alert_box(doc, "Análise de Riscos", analise_text, "risk")

    riscos = data.get("riscos", data.get("tabela", data.get("dados", [])))
    if isinstance(riscos, list) and riscos:
        _add_subsection_heading(doc, "Matriz de Riscos")
        _add_risk_table(doc, riscos)

    _render_flags(doc, data.get("flags", []))


def _render_section_covenants(doc, data: dict) -> None:
    """Render section 9 — Covenants e Condicionantes."""
    _add_section_heading(doc, SECTION_TITLES["covenants"])

    analise_text = data.get("analise", "")
    if analise_text:
        _add_alert_box(doc, "Covenants Sugeridos", analise_text, "warning")

    tabela = data.get("tabela", data.get("dados", []))
    if isinstance(tabela, list) and tabela:
        if isinstance(tabela[0], dict):
            headers = list(tabela[0].keys())
            rows = [[_safe(item.get(h)) for h in headers] for item in tabela]
            _add_data_table(doc, headers, rows)
        elif isinstance(tabela[0], list):
            _add_data_table(doc, tabela[0], tabela[1:])

    _render_flags(doc, data.get("flags", []))


def _render_flags(doc, flags: List[str]) -> None:
    """Render flag alerts if any."""
    if not flags:
        return
    for flag in flags:
        _add_alert_box(doc, "\u26a0 Flag", flag, "warning")


def _render_rating_final(doc, analise: dict) -> None:
    """Render the final rating box if available."""
    rating = analise.get("rating_final", {})
    if not rating:
        return

    nota = _safe(rating.get("nota", rating.get("rating", "")), "N/D")
    descricao = _safe(rating.get("descricao", rating.get("analise", "")), "")

    alert_type = "info"
    if nota in ("A", "A+", "Forte"):
        alert_type = "positive"
    elif nota in ("C", "D", "Crítico"):
        alert_type = "risk"
    elif nota in ("B-", "Atenção"):
        alert_type = "warning"

    _add_alert_box(doc, f"RATING FINAL: {nota}", descricao, alert_type)


# ═══════════════════════════════════════════════════════════════════════════
# Main Function
# ═══════════════════════════════════════════════════════════════════════════

def generate_mac(analise: dict, parametros: dict, output_path: str) -> str:
    """
    Generate a MAC (Memorando de Análise de Crédito) .docx file.

    Args:
        analise: Analysis JSON with section keys (tomador, patrimonio, producao,
                 capital, operacao, pagamento, onus, riscos, covenants,
                 cronograma, rating_final, kpis).
        parametros: Operation parameters dict.
        output_path: Path where the .docx will be saved.

    Returns:
        The output_path of the generated file.
    """
    doc = Document()

    # ── Page setup ──────────────────────────────────────────────────────
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.97)
    section.bottom_margin = Inches(0.69)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    # Empty header
    header = section.header
    header.is_linked_to_previous = False
    if header.paragraphs:
        header.paragraphs[0].clear()

    # Footer
    _add_footer(doc)

    # ── 1. Empty spacer paragraph ───────────────────────────────────────
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(8)
    spacer.paragraph_format.space_before = Pt(0)

    # ── 2. Title ────────────────────────────────────────────────────────
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_after = Pt(4)
    p_title.paragraph_format.space_before = Pt(0)
    run_title = p_title.add_run("MEMORANDO DE ANÁLISE DE CRÉDITO")
    _set_run_font(run_title, size_pt=17, color_hex=NAVY_HEX, bold=True)

    # ── 3. Subtitle ────────────────────────────────────────────────────
    tipo_captacao = _safe(parametros.get("tipo_captacao"), "Captação PJ")
    tomador = _safe(parametros.get("tomador"), "N/D")

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(2)
    p_sub.paragraph_format.space_before = Pt(0)
    run_sub = p_sub.add_run(f"{tipo_captacao} — {tomador}")
    _set_run_font(run_sub, size_pt=11, color_hex=DARK_SLATE_HEX)

    # ── 4. Meta-line ───────────────────────────────────────────────────
    detalhes = _safe(parametros.get("tipo_operacao"), "N/D")
    taxa = _safe(parametros.get("taxa"), "N/D")
    prazo = f"{_safe(parametros.get('prazo_meses'), 'N/D')}m"
    data_str = datetime.now().strftime("%d/%m/%Y")

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(10)
    p_meta.paragraph_format.space_before = Pt(0)

    meta_prefix = f"{detalhes} \u00b7 {taxa} \u00b7 {prazo} \u00b7 "
    run_meta = p_meta.add_run(meta_prefix)
    _set_run_font(run_meta, size_pt=8, color_hex=MUTED_GRAY_HEX)

    run_date = p_meta.add_run(data_str)
    _set_run_font(run_date, size_pt=8, color_hex=MUTED_GRAY_HEX, italic=True)

    # ── 5. KPI Strip ──────────────────────────────────────────────────
    kpis_data = analise.get("kpis", {})
    if not kpis_data:
        # Build KPIs from parametros if not in analise
        volume = parametros.get("volume")
        prazo_m = parametros.get("prazo_meses")
        kpis_data = {
            "principal": _format_currency(volume) if volume else "N/D",
            "prazo": f"{prazo_m} meses" if prazo_m else "N/D",
            "taxa_all_in": _safe(parametros.get("taxa")),
            "juros_ano1": "N/D",
            "total_juros": "N/D",
        }

    _add_kpi_strip(doc, kpis_data)

    # ── 6. Sections 1-9 ───────────────────────────────────────────────
    section_renderers = {
        "tomador": lambda d: _render_section_tomador(doc, d, parametros),
        "patrimonio": lambda d: _render_section_patrimonio(doc, d),
        "producao": lambda d: _render_section_producao(doc, d),
        "capital": lambda d: _render_section_capital(doc, d),
        "operacao": lambda d: _render_section_operacao(doc, d, parametros),
        "pagamento": lambda d: _render_section_pagamento(doc, d),
        "onus": lambda d: _render_section_onus(doc, d),
        "riscos": lambda d: _render_section_riscos(doc, d),
        "covenants": lambda d: _render_section_covenants(doc, d),
    }

    for key in SECTION_TITLES:
        section_data = analise.get(key, {})
        if isinstance(section_data, dict):
            renderer = section_renderers.get(key)
            if renderer:
                renderer(section_data)

    # ── Rating Final ───────────────────────────────────────────────────
    _render_rating_final(doc, analise)

    # ── 9. Closing paragraph ──────────────────────────────────────────
    p_close = doc.add_paragraph()
    p_close.paragraph_format.space_before = Pt(16)
    p_close.paragraph_format.space_after = Pt(4)
    year = datetime.now().year
    close_text = (
        f"São Paulo, {data_str} \u00b7 ZYN Capital \u2014 "
        f"Documento confidencial de uso exclusivo interno"
    )
    run_close = p_close.add_run(close_text)
    _set_run_font(run_close, size_pt=7, color_hex=MUTED_GRAY_HEX, italic=True)

    # ── Save ──────────────────────────────────────────────────────────
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc.save(output_path)
    return output_path
