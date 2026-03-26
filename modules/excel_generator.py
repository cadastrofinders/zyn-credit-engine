"""
ZYN Capital — Gerador de Planilha de Análise Técnica de Crédito
Gera modelo financeiro .xlsx a partir da análise de crédito MAC ZYN v3.

Suporta:
- Análise por empresa individual
- Consolidado de grupo econômico
- Indicadores financeiros e contábeis
- Simulador de viabilidade
- Resumo executivo com parecer
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
    numbers,
)
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# ZYN Color Palette
# ---------------------------------------------------------------------------
NAVY = "223040"
GREEN = "2E7D4F"
GRAY = "8B9197"
LIGHT_GRAY = "F2F4F6"
WHITE = "FFFFFF"
RED = "922B21"
YELLOW = "7D6608"
DARK_SLATE = "3A4F63"

# Fills
FILL_NAVY = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
FILL_GREEN = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
FILL_LIGHT = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
FILL_WHITE = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")

# Fonts
FONT_TITLE = Font(name="Calibri", size=14, bold=True, color=WHITE)
FONT_SECTION = Font(name="Calibri", size=11, bold=True, color=NAVY)
FONT_HEADER = Font(name="Calibri", size=10, bold=True, color=WHITE)
FONT_LABEL = Font(name="Calibri", size=10, color=DARK_SLATE)
FONT_VALUE = Font(name="Calibri", size=10, color=NAVY)
FONT_VALUE_BOLD = Font(name="Calibri", size=10, bold=True, color=NAVY)
FONT_SMALL = Font(name="Calibri", size=9, color=GRAY)
FONT_ALERT = Font(name="Calibri", size=10, bold=True, color=RED)
FONT_OK = Font(name="Calibri", size=10, bold=True, color=GREEN)

# Borders
THIN_BORDER = Border(
    left=Side(style="thin", color=GRAY),
    right=Side(style="thin", color=GRAY),
    top=Side(style="thin", color=GRAY),
    bottom=Side(style="thin", color=GRAY),
)

ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

BRL_FORMAT = '#,##0'
BRL_MM_FORMAT = '#,##0.0,,"MM"'
PCT_FORMAT = '0.0%'
MULT_FORMAT = '0.00"x"'


def _safe(val, default=0):
    """Safely convert to number."""
    if isinstance(val, (int, float)):
        return val
    if isinstance(val, str):
        try:
            return float(val.replace(",", ".").replace("%", "").strip())
        except (ValueError, AttributeError):
            pass
    return default


def _safe_str(val, default="—"):
    if val is None or val == "":
        return default
    return str(val)


def _rating_color(nota: str) -> str:
    return {"A": GREEN, "B": "2E7D4F", "C": YELLOW, "D": RED, "E": "7B241C"}.get(nota, GRAY)


def _set_col_widths(ws, widths: list[float]):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_header_row(ws, row: int, values: list[str], fill=FILL_NAVY, font=FONT_HEADER):
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = font
        cell.fill = fill
        cell.alignment = ALIGN_CENTER
        cell.border = THIN_BORDER


def _write_row(ws, row: int, values: list, fonts=None, formats=None):
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = fonts[col - 1] if fonts else FONT_VALUE
        cell.alignment = ALIGN_LEFT if col == 1 else ALIGN_RIGHT
        cell.border = THIN_BORDER
        if formats and col - 1 < len(formats) and formats[col - 1]:
            cell.number_format = formats[col - 1]
        if row % 2 == 0:
            cell.fill = FILL_LIGHT


def _section_title(ws, row: int, title: str, cols: int = 7):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row=row, column=1, value=title)
    cell.font = FONT_SECTION
    cell.fill = FILL_LIGHT
    cell.alignment = ALIGN_LEFT
    cell.border = THIN_BORDER
    return row + 1


# ---------------------------------------------------------------------------
# Sheet 1 — Resumo Executivo
# ---------------------------------------------------------------------------
def _build_resumo(wb: Workbook, analise: dict, op: dict):
    ws = wb.active
    ws.title = "Resumo Executivo"
    _set_col_widths(ws, [30, 20, 20, 20, 20, 20, 20])

    rating = analise.get("rating_final", {})
    kpis = analise.get("kpis", {})
    tomador_info = analise.get("tomador", {})

    # Title bar
    ws.merge_cells("A1:G1")
    cell = ws.cell(row=1, column=1, value="ZYN Capital — Análise Técnica de Crédito")
    cell.font = FONT_TITLE
    cell.fill = FILL_NAVY
    cell.alignment = ALIGN_CENTER

    ws.merge_cells("A2:G2")
    cell = ws.cell(row=2, column=1, value=f"Data: {datetime.now().strftime('%d/%m/%Y')} | CONFIDENCIAL")
    cell.font = Font(name="Calibri", size=9, color=WHITE)
    cell.fill = FILL_NAVY
    cell.alignment = ALIGN_CENTER

    # Tomador info
    r = 4
    r = _section_title(ws, r, "IDENTIFICAÇÃO DO TOMADOR")
    info_pairs = [
        ("Razão Social", _safe_str(tomador_info.get("razao_social", op.get("tomador")))),
        ("CNPJ", _safe_str(tomador_info.get("cnpj", op.get("cnpj")))),
        ("Grupo Econômico", _safe_str(tomador_info.get("grupo_economico"))),
        ("Tipo Operação", _safe_str(op.get("tipo_operacao"))),
        ("Volume", f"R$ {_safe(op.get('volume', 0)):,.0f}"),
        ("Prazo", f"{op.get('prazo_meses', '—')} meses"),
        ("Taxa", _safe_str(op.get("taxa"))),
        ("Garantias", ", ".join(op.get("garantias", [])) or "—"),
    ]
    for label, val in info_pairs:
        ws.cell(row=r, column=1, value=label).font = FONT_LABEL
        ws.cell(row=r, column=2, value=val).font = FONT_VALUE_BOLD
        ws.cell(row=r, column=1).border = THIN_BORDER
        ws.cell(row=r, column=2).border = THIN_BORDER
        r += 1

    r += 1
    r = _section_title(ws, r, "RATING & PARECER")
    nota = rating.get("nota", "—")
    nota_cell = ws.cell(row=r, column=1, value="Rating Final")
    nota_cell.font = FONT_LABEL
    nota_val = ws.cell(row=r, column=2, value=nota)
    nota_val.font = Font(name="Calibri", size=14, bold=True, color=_rating_color(nota))
    r += 1

    ws.cell(row=r, column=1, value="Parecer").font = FONT_LABEL
    ws.cell(row=r, column=2, value=_safe_str(rating.get("parecer"))).font = FONT_VALUE_BOLD
    r += 1

    ws.cell(row=r, column=1, value="Justificativa").font = FONT_LABEL
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.cell(row=r, column=2, value=_safe_str(rating.get("justificativa"))).font = FONT_VALUE
    ws.cell(row=r, column=2).alignment = ALIGN_LEFT
    r += 2

    # KPIs
    r = _section_title(ws, r, "INDICADORES-CHAVE (KPIs)")
    _write_header_row(ws, r, ["Indicador", "Valor", "Benchmark", "Status", "", "", ""])
    r += 1

    dscr = _safe(kpis.get("dscr"))
    ltv = _safe(kpis.get("ltv"))
    div_ebitda = _safe(kpis.get("divida_liquida_ebitda"))
    margem = _safe(kpis.get("margem_ebitda"))

    kpi_rows = [
        ("Receita Líquida", _safe(kpis.get("receita_liquida")), "—", "", BRL_FORMAT),
        ("EBITDA", _safe(kpis.get("ebitda")), "—", "", BRL_FORMAT),
        ("Margem EBITDA", margem if margem <= 1 else margem / 100, ">15%",
         "OK" if (margem if margem <= 1 else margem / 100) > 0.15 else "Atenção", PCT_FORMAT),
        ("Dív.Líq./EBITDA", div_ebitda, "<3.5x",
         "OK" if div_ebitda < 3.5 else ("Atenção" if div_ebitda < 5 else "Crítico"), MULT_FORMAT),
        ("LTV", ltv if ltv <= 1 else ltv / 100, "<80%",
         "OK" if (ltv if ltv <= 1 else ltv / 100) < 0.80 else ("Atenção" if (ltv if ltv <= 1 else ltv / 100) < 1.0 else "Crítico"), PCT_FORMAT),
        ("DSCR", dscr, ">1.2x",
         "OK" if dscr > 1.2 else ("Atenção" if dscr > 1.0 else "Crítico"), MULT_FORMAT),
    ]
    for label, val, bench, status, fmt in kpi_rows:
        font_status = FONT_OK if status == "OK" else (FONT_ALERT if status == "Crítico" else FONT_LABEL)
        _write_row(ws, r, [label, val, bench, status], [FONT_LABEL, FONT_VALUE_BOLD, FONT_SMALL, font_status], [None, fmt])
        r += 1

    r += 1

    # Ratings por seção
    r = _section_title(ws, r, "RATINGS POR SEÇÃO")
    secoes = ["tomador", "patrimonio", "producao", "capital", "operacao",
              "pagamento", "onus", "riscos", "covenants", "cronograma"]
    _write_header_row(ws, r, ["Seção", "Rating", "Análise", "", "", "", ""])
    r += 1
    for s in secoes:
        sec = analise.get(s, {})
        rating_sec = _safe_str(sec.get("rating_secao"))
        analise_text = _safe_str(sec.get("analise"))
        ws.cell(row=r, column=1, value=s.title()).font = FONT_LABEL
        ws.cell(row=r, column=2, value=rating_sec).font = FONT_VALUE_BOLD
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        ws.cell(row=r, column=3, value=analise_text).font = FONT_SMALL
        ws.cell(row=r, column=3).alignment = ALIGN_LEFT
        for c in range(1, 8):
            ws.cell(row=r, column=c).border = THIN_BORDER
        if r % 2 == 0:
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = FILL_LIGHT
        r += 1

    r += 1

    # Recomendações
    recs = rating.get("recomendacoes", [])
    if recs:
        r = _section_title(ws, r, "RECOMENDAÇÕES")
        for i, rec in enumerate(recs, 1):
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
            ws.cell(row=r, column=1, value=f"{i}. {rec}").font = FONT_VALUE
            ws.cell(row=r, column=1).alignment = ALIGN_LEFT
            r += 1

    # Footer
    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    cell = ws.cell(row=r, column=1, value="ZYN Capital — Crédito Estruturado & M&A | São Paulo")
    cell.font = Font(name="Calibri", size=9, color=GRAY)
    cell.alignment = ALIGN_CENTER


# ---------------------------------------------------------------------------
# Sheet 2 — Indicadores Financeiros
# ---------------------------------------------------------------------------
def _build_indicadores(wb: Workbook, analise: dict, op: dict):
    ws = wb.create_sheet("Indicadores Financeiros")
    _set_col_widths(ws, [35, 20, 15, 15, 15, 15, 15])

    kpis = analise.get("kpis", {})
    capital = analise.get("capital", {})
    indicadores = capital.get("indicadores", {})

    ws.merge_cells("A1:G1")
    cell = ws.cell(row=1, column=1, value="Indicadores Financeiros & Contábeis")
    cell.font = FONT_TITLE
    cell.fill = FILL_NAVY
    cell.alignment = ALIGN_CENTER

    r = 3
    r = _section_title(ws, r, "INDICADORES DE RESULTADO")
    _write_header_row(ws, r, ["Indicador", "Valor", "Referência", "Status", "", "", ""])
    r += 1

    receita = _safe(kpis.get("receita_liquida"))
    ebitda = _safe(kpis.get("ebitda"))
    margem = _safe(kpis.get("margem_ebitda"))
    if margem > 1:
        margem = margem / 100

    result_rows = [
        ("Receita Líquida", receita, BRL_FORMAT),
        ("EBITDA", ebitda, BRL_FORMAT),
        ("Margem EBITDA", margem, PCT_FORMAT),
    ]
    for label, val, fmt in result_rows:
        _write_row(ws, r, [label, val, "", ""], [FONT_LABEL, FONT_VALUE_BOLD, FONT_SMALL, FONT_SMALL], [None, fmt])
        r += 1

    r += 1
    r = _section_title(ws, r, "INDICADORES DE ENDIVIDAMENTO")
    _write_header_row(ws, r, ["Indicador", "Valor", "Limite Atenção", "Limite Crítico", "Status", "", ""])
    r += 1

    div_ebitda = _safe(indicadores.get("divida_liquida_ebitda", kpis.get("divida_liquida_ebitda")))
    div_pl = _safe(indicadores.get("divida_pl"))
    liq_corrente = _safe(indicadores.get("liquidez_corrente"))
    roe = _safe(indicadores.get("roe"))
    dscr = _safe(kpis.get("dscr"))
    ltv = _safe(kpis.get("ltv"))
    if ltv > 1:
        ltv = ltv / 100

    debt_rows = [
        ("Dív.Líq./EBITDA", div_ebitda, 3.5, 5.0, MULT_FORMAT),
        ("Dívida/PL", div_pl, 1.5, 3.0, MULT_FORMAT),
        ("Liquidez Corrente", liq_corrente, 1.0, 0.8, MULT_FORMAT),
        ("DSCR", dscr, 1.2, 1.0, MULT_FORMAT),
        ("LTV", ltv, 0.80, 1.0, PCT_FORMAT),
        ("ROE", roe if roe <= 1 else roe / 100, 0.10, 0.05, PCT_FORMAT),
    ]
    for label, val, atencao, critico, fmt in debt_rows:
        if label in ("Liquidez Corrente", "DSCR", "ROE"):
            status = "OK" if val > atencao else ("Atenção" if val > critico else "Crítico")
        else:
            status = "OK" if val < atencao else ("Atenção" if val < critico else "Crítico")
        font_status = FONT_OK if status == "OK" else (FONT_ALERT if status == "Crítico" else FONT_LABEL)
        _write_row(ws, r, [label, val, atencao, critico, status],
                   [FONT_LABEL, FONT_VALUE_BOLD, FONT_SMALL, FONT_SMALL, font_status],
                   [None, fmt, fmt, fmt])
        r += 1

    r += 1
    r = _section_title(ws, r, "ESTRUTURA DE CAPITAL")
    cap_text = _safe_str(capital.get("estrutura_capital"))
    endiv_text = _safe_str(capital.get("endividamento"))
    analise_text = _safe_str(capital.get("analise"))

    for label, val in [("Estrutura", cap_text), ("Endividamento", endiv_text), ("Análise", analise_text)]:
        ws.cell(row=r, column=1, value=label).font = FONT_LABEL
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.cell(row=r, column=2, value=val).font = FONT_VALUE
        ws.cell(row=r, column=2).alignment = ALIGN_LEFT
        r += 1

    # Flags de atenção
    r += 1
    r = _section_title(ws, r, "FLAGS DE ATENÇÃO")
    todas_flags = []
    for s in ["tomador", "patrimonio", "producao", "capital", "operacao", "pagamento", "onus", "riscos", "covenants", "cronograma"]:
        for flag in analise.get(s, {}).get("flags", []):
            todas_flags.append((s.title(), flag))

    if todas_flags:
        _write_header_row(ws, r, ["Seção", "Flag", "", "", "", "", ""])
        r += 1
        for secao, flag in todas_flags:
            ws.cell(row=r, column=1, value=secao).font = FONT_LABEL
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
            ws.cell(row=r, column=2, value=flag).font = FONT_ALERT
            ws.cell(row=r, column=2).alignment = ALIGN_LEFT
            for c in range(1, 8):
                ws.cell(row=r, column=c).border = THIN_BORDER
            r += 1
    else:
        ws.cell(row=r, column=1, value="Nenhuma flag de atenção identificada.").font = FONT_OK
        r += 1


# ---------------------------------------------------------------------------
# Sheet 3 — Análise de Crédito Detalhada
# ---------------------------------------------------------------------------
def _build_credito(wb: Workbook, analise: dict, op: dict):
    ws = wb.create_sheet("Análise de Crédito")
    _set_col_widths(ws, [25, 40, 15, 15, 15, 15, 15])

    ws.merge_cells("A1:G1")
    cell = ws.cell(row=1, column=1, value="Análise de Crédito Detalhada — MAC ZYN v3")
    cell.font = FONT_TITLE
    cell.fill = FILL_NAVY
    cell.alignment = ALIGN_CENTER

    r = 3
    secoes = [
        ("tomador", "Tomador"),
        ("patrimonio", "Patrimônio & Garantias"),
        ("producao", "Produção"),
        ("capital", "Estrutura de Capital"),
        ("operacao", "Operação"),
        ("pagamento", "Capacidade de Pagamento"),
        ("onus", "Ônus & Gravames"),
        ("riscos", "Riscos"),
        ("covenants", "Covenants"),
        ("cronograma", "Cronograma"),
    ]

    for key, titulo in secoes:
        sec = analise.get(key, {})
        r = _section_title(ws, r, titulo.upper())

        # Main fields
        for field, val in sec.items():
            if field in ("flags", "rating_secao", "analise"):
                continue
            if isinstance(val, list):
                # Table for lists (e.g., gravames, matriz_riscos, clausulas, etapas)
                if val and isinstance(val[0], dict):
                    headers = list(val[0].keys())
                    _write_header_row(ws, r, headers + [""] * (7 - len(headers)))
                    r += 1
                    for item in val:
                        row_vals = [_safe_str(item.get(h)) for h in headers]
                        row_vals += [""] * (7 - len(row_vals))
                        _write_row(ws, r, row_vals)
                        r += 1
                else:
                    ws.cell(row=r, column=1, value=field.replace("_", " ").title()).font = FONT_LABEL
                    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
                    ws.cell(row=r, column=2, value=", ".join(str(v) for v in val)).font = FONT_VALUE
                    r += 1
            elif isinstance(val, dict):
                for sub_key, sub_val in val.items():
                    ws.cell(row=r, column=1, value=sub_key.replace("_", " ").title()).font = FONT_LABEL
                    ws.cell(row=r, column=2, value=_safe_str(sub_val)).font = FONT_VALUE
                    ws.cell(row=r, column=1).border = THIN_BORDER
                    ws.cell(row=r, column=2).border = THIN_BORDER
                    r += 1
            else:
                ws.cell(row=r, column=1, value=field.replace("_", " ").title()).font = FONT_LABEL
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
                ws.cell(row=r, column=2, value=_safe_str(val)).font = FONT_VALUE
                ws.cell(row=r, column=2).alignment = ALIGN_LEFT
                ws.cell(row=r, column=1).border = THIN_BORDER
                ws.cell(row=r, column=2).border = THIN_BORDER
                r += 1

        # Rating + Análise da seção
        ws.cell(row=r, column=1, value="Rating Seção").font = FONT_LABEL
        ws.cell(row=r, column=2, value=_safe_str(sec.get("rating_secao"))).font = FONT_VALUE_BOLD
        r += 1
        ws.cell(row=r, column=1, value="Análise").font = FONT_LABEL
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.cell(row=r, column=2, value=_safe_str(sec.get("analise"))).font = FONT_SMALL
        ws.cell(row=r, column=2).alignment = ALIGN_LEFT
        r += 2


# ---------------------------------------------------------------------------
# Sheet 4 — Riscos & Viabilidade
# ---------------------------------------------------------------------------
def _build_riscos(wb: Workbook, analise: dict, op: dict):
    ws = wb.create_sheet("Riscos & Viabilidade")
    _set_col_widths(ws, [25, 20, 15, 15, 30, 15, 15])

    ws.merge_cells("A1:G1")
    cell = ws.cell(row=1, column=1, value="Matriz de Riscos & Análise de Viabilidade")
    cell.font = FONT_TITLE
    cell.fill = FILL_NAVY
    cell.alignment = ALIGN_CENTER

    riscos = analise.get("riscos", {})
    r = 3

    # Categorias de risco
    r = _section_title(ws, r, "CATEGORIAS DE RISCO")
    for cat in ["mercado", "credito", "operacional", "legal"]:
        val = _safe_str(riscos.get(cat))
        ws.cell(row=r, column=1, value=f"Risco {cat.title()}").font = FONT_LABEL
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.cell(row=r, column=2, value=val).font = FONT_VALUE
        ws.cell(row=r, column=2).alignment = ALIGN_LEFT
        for c in range(1, 8):
            ws.cell(row=r, column=c).border = THIN_BORDER
        r += 1

    r += 1

    # Matriz de riscos
    matriz = riscos.get("matriz_riscos", [])
    if matriz:
        r = _section_title(ws, r, "MATRIZ DE RISCOS")
        _write_header_row(ws, r, ["Risco", "Probabilidade", "Impacto", "Mitigante", "", "", ""])
        r += 1
        for item in matriz:
            ws.cell(row=r, column=1, value=_safe_str(item.get("risco"))).font = FONT_LABEL
            ws.cell(row=r, column=2, value=_safe_str(item.get("probabilidade"))).font = FONT_VALUE
            ws.cell(row=r, column=3, value=_safe_str(item.get("impacto"))).font = FONT_VALUE
            ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
            ws.cell(row=r, column=4, value=_safe_str(item.get("mitigante"))).font = FONT_SMALL
            ws.cell(row=r, column=4).alignment = ALIGN_LEFT
            for c in range(1, 8):
                ws.cell(row=r, column=c).border = THIN_BORDER
            if r % 2 == 0:
                for c in range(1, 8):
                    ws.cell(row=r, column=c).fill = FILL_LIGHT
            r += 1

    r += 1

    # Viabilidade simplificada — stress scenarios
    kpis = analise.get("kpis", {})
    dscr = _safe(kpis.get("dscr"))
    ltv = _safe(kpis.get("ltv"))
    if ltv > 1:
        ltv = ltv / 100

    r = _section_title(ws, r, "SIMULAÇÃO DE STRESS — VIABILIDADE")
    _write_header_row(ws, r, ["Cenário", "DSCR", "LTV", "Status", "", "", ""])
    r += 1

    scenarios = [
        ("Base", dscr, ltv),
        ("Stress Juros +300bps", dscr * 0.85, ltv * 1.05),
        ("Stress Receita -20%", dscr * 0.80, ltv * 1.10),
        ("Stress Combinado", dscr * 0.70, ltv * 1.15),
    ]
    for nome, dscr_s, ltv_s in scenarios:
        status = "Viável" if dscr_s > 1.0 and ltv_s < 1.0 else "Inviável"
        font_s = FONT_OK if status == "Viável" else FONT_ALERT
        _write_row(ws, r, [nome, dscr_s, ltv_s, status],
                   [FONT_LABEL, FONT_VALUE_BOLD, FONT_VALUE_BOLD, font_s],
                   [None, MULT_FORMAT, PCT_FORMAT])
        r += 1

    r += 1

    # Covenants
    covenants = analise.get("covenants", {})
    clausulas = covenants.get("clausulas", [])
    if clausulas:
        r = _section_title(ws, r, "COVENANTS")
        _write_header_row(ws, r, ["Covenant", "Limite", "Atual", "Status", "", "", ""])
        r += 1
        for cl in clausulas:
            status = _safe_str(cl.get("status"))
            font_s = FONT_OK if "cumpr" in status.lower() or "ok" in status.lower() else FONT_ALERT
            _write_row(ws, r,
                       [_safe_str(cl.get("covenant")), _safe_str(cl.get("limite")),
                        _safe_str(cl.get("atual")), status],
                       [FONT_LABEL, FONT_VALUE, FONT_VALUE_BOLD, font_s])
            r += 1


# ---------------------------------------------------------------------------
# Sheet 5 — Checklist & Lacunas
# ---------------------------------------------------------------------------
def _build_checklist(wb: Workbook, analise: dict):
    ws = wb.create_sheet("Checklist & Lacunas")
    _set_col_widths(ws, [40, 15, 40, 15, 15, 15, 15])

    ws.merge_cells("A1:G1")
    cell = ws.cell(row=1, column=1, value="Checklist de Documentação & Lacunas")
    cell.font = FONT_TITLE
    cell.fill = FILL_NAVY
    cell.alignment = ALIGN_CENTER

    checklist = analise.get("checklist_lacunas", {})
    r = 3

    total_p = checklist.get("total_pendencias", 0)
    total_c = checklist.get("total_criticas", 0)
    ws.cell(row=r, column=1, value=f"Total Pendências: {total_p}").font = FONT_VALUE_BOLD
    ws.cell(row=r, column=2, value=f"Críticas: {total_c}").font = FONT_ALERT if total_c > 0 else FONT_VALUE
    r += 2

    docs_f = checklist.get("documentos_faltantes", [])
    if docs_f:
        r = _section_title(ws, r, "DOCUMENTOS FALTANTES")
        _write_header_row(ws, r, ["Item", "Criticidade", "Motivo", "", "", "", ""])
        r += 1
        for d in docs_f:
            crit = _safe_str(d.get("criticidade"))
            font_c = FONT_ALERT if crit.lower() in ("alta", "crítica") else FONT_LABEL
            _write_row(ws, r, [_safe_str(d.get("item")), crit, _safe_str(d.get("motivo"))],
                       [FONT_VALUE, font_c, FONT_SMALL])
            r += 1

    r += 1
    info_p = checklist.get("informacoes_pendentes", [])
    if info_p:
        r = _section_title(ws, r, "INFORMAÇÕES PENDENTES")
        _write_header_row(ws, r, ["Item", "Criticidade", "Motivo", "", "", "", ""])
        r += 1
        for p in info_p:
            crit = _safe_str(p.get("criticidade"))
            font_c = FONT_ALERT if crit.lower() in ("alta", "crítica") else FONT_LABEL
            _write_row(ws, r, [_safe_str(p.get("item")), crit, _safe_str(p.get("motivo"))],
                       [FONT_VALUE, font_c, FONT_SMALL])
            r += 1


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
def generate_excel(
    analise: dict,
    parametros: dict,
    output_path: str,
    historico_analyses: list[dict] | None = None,
) -> str:
    """
    Gera planilha de análise técnica de crédito.

    Args:
        analise: resultado do analyze_credit()
        parametros: parametros_operacao
        output_path: caminho de saída .xlsx
        historico_analyses: lista de análises anteriores para cruzamento (opcional)

    Returns:
        caminho do arquivo gerado
    """
    wb = Workbook()

    _build_resumo(wb, analise, parametros)
    _build_indicadores(wb, analise, parametros)
    _build_credito(wb, analise, parametros)
    _build_riscos(wb, analise, parametros)
    _build_checklist(wb, analise)

    # If historico provided, add cross-reference sheet
    if historico_analyses:
        _build_cross_reference(wb, analise, parametros, historico_analyses)

    wb.save(output_path)
    return output_path


def _build_cross_reference(wb: Workbook, analise: dict, op: dict, historico: list[dict]):
    """Aba de cruzamento com análises anteriores — memória do motor."""
    ws = wb.create_sheet("Cruzamento Histórico")
    _set_col_widths(ws, [25, 15, 15, 15, 15, 15, 20])

    ws.merge_cells("A1:G1")
    cell = ws.cell(row=1, column=1, value="Cruzamento com Análises Anteriores")
    cell.font = FONT_TITLE
    cell.fill = FILL_NAVY
    cell.alignment = ALIGN_CENTER

    r = 3
    r = _section_title(ws, r, "COMPARATIVO DE OPERAÇÕES ANALISADAS")
    _write_header_row(ws, r, ["Tomador", "Rating", "DSCR", "LTV", "Dív/EBITDA", "Parecer", "Data"])
    r += 1

    # Current analysis first
    kpis = analise.get("kpis", {})
    rating = analise.get("rating_final", {})
    _write_row(ws, r, [
        op.get("tomador", "—"),
        rating.get("nota", "—"),
        _safe(kpis.get("dscr")),
        _safe(kpis.get("ltv")),
        _safe(kpis.get("divida_liquida_ebitda")),
        rating.get("parecer", "—"),
        datetime.now().strftime("%d/%m/%Y"),
    ], [FONT_VALUE_BOLD] * 7, [None, None, MULT_FORMAT, PCT_FORMAT, MULT_FORMAT])
    r += 1

    # Historical analyses
    for item in historico:
        h_op = item.get("operacao", {})
        h_analise = item.get("analise", {})
        h_kpis = h_analise.get("kpis", {})
        h_rating = h_analise.get("rating_final", {})
        _write_row(ws, r, [
            h_op.get("tomador", "—"),
            h_rating.get("nota", "—"),
            _safe(h_kpis.get("dscr")),
            _safe(h_kpis.get("ltv")),
            _safe(h_kpis.get("divida_liquida_ebitda")),
            h_rating.get("parecer", "—"),
            item.get("data_analise", "—"),
        ], formats=[None, None, MULT_FORMAT, PCT_FORMAT, MULT_FORMAT])
        r += 1


def generate_excel_grupo(
    analyses: list[tuple[dict, dict]],
    grupo_nome: str,
    output_path: str,
) -> str:
    """
    Gera planilha consolidada para grupo econômico com múltiplas empresas.

    Args:
        analyses: lista de (analise, parametros) por empresa
        grupo_nome: nome do grupo econômico
        output_path: caminho de saída .xlsx

    Returns:
        caminho do arquivo gerado
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Consolidado Grupo"
    _set_col_widths(ws, [30, 20, 15, 15, 15, 15, 20])

    ws.merge_cells("A1:G1")
    cell = ws.cell(row=1, column=1, value=f"ZYN Capital — Análise Consolidada: {grupo_nome}")
    cell.font = FONT_TITLE
    cell.fill = FILL_NAVY
    cell.alignment = ALIGN_CENTER

    ws.merge_cells("A2:G2")
    cell = ws.cell(row=2, column=1, value=f"Data: {datetime.now().strftime('%d/%m/%Y')} | {len(analyses)} empresa(s) | CONFIDENCIAL")
    cell.font = Font(name="Calibri", size=9, color=WHITE)
    cell.fill = FILL_NAVY
    cell.alignment = ALIGN_CENTER

    r = 4
    r = _section_title(ws, r, "RESUMO POR EMPRESA")
    _write_header_row(ws, r, ["Empresa", "Rating", "DSCR", "LTV", "Dív/EBITDA", "Parecer", "Volume"])
    r += 1

    total_receita = 0
    total_ebitda = 0
    total_volume = 0

    for analise, params in analyses:
        kpis = analise.get("kpis", {})
        rating = analise.get("rating_final", {})
        vol = _safe(params.get("volume"))
        total_volume += vol
        total_receita += _safe(kpis.get("receita_liquida"))
        total_ebitda += _safe(kpis.get("ebitda"))

        _write_row(ws, r, [
            params.get("tomador", "—"),
            rating.get("nota", "—"),
            _safe(kpis.get("dscr")),
            _safe(kpis.get("ltv")),
            _safe(kpis.get("divida_liquida_ebitda")),
            rating.get("parecer", "—"),
            vol,
        ], formats=[None, None, MULT_FORMAT, PCT_FORMAT, MULT_FORMAT, None, BRL_FORMAT])
        r += 1

    # Totals row
    r += 1
    ws.cell(row=r, column=1, value="TOTAL GRUPO").font = FONT_VALUE_BOLD
    ws.cell(row=r, column=7, value=total_volume).font = FONT_VALUE_BOLD
    ws.cell(row=r, column=7).number_format = BRL_FORMAT
    r += 2

    # Consolidated KPIs
    r = _section_title(ws, r, "KPIs CONSOLIDADOS")
    consolidated = [
        ("Receita Líquida Total", total_receita, BRL_FORMAT),
        ("EBITDA Total", total_ebitda, BRL_FORMAT),
        ("Margem EBITDA Grupo", total_ebitda / total_receita if total_receita else 0, PCT_FORMAT),
        ("Volume Total", total_volume, BRL_FORMAT),
    ]
    for label, val, fmt in consolidated:
        ws.cell(row=r, column=1, value=label).font = FONT_LABEL
        ws.cell(row=r, column=2, value=val).font = FONT_VALUE_BOLD
        ws.cell(row=r, column=2).number_format = fmt
        ws.cell(row=r, column=1).border = THIN_BORDER
        ws.cell(row=r, column=2).border = THIN_BORDER
        r += 1

    # Individual company sheets
    for analise, params in analyses:
        empresa = params.get("tomador", "Empresa")[:31]  # sheet name limit
        _build_resumo_empresa(wb, analise, params, empresa)

    wb.save(output_path)
    return output_path


def _build_resumo_empresa(wb: Workbook, analise: dict, op: dict, sheet_name: str):
    """Cria aba de resumo executivo para uma empresa do grupo."""
    ws = wb.create_sheet(sheet_name)
    _set_col_widths(ws, [30, 20, 20, 20, 20, 20, 20])

    rating = analise.get("rating_final", {})
    kpis = analise.get("kpis", {})

    ws.merge_cells("A1:G1")
    cell = ws.cell(row=1, column=1, value=f"Resumo Executivo — {op.get('tomador', '—')}")
    cell.font = FONT_TITLE
    cell.fill = FILL_NAVY
    cell.alignment = ALIGN_CENTER

    r = 3
    r = _section_title(ws, r, "RATING & PARECER")
    nota = rating.get("nota", "—")
    ws.cell(row=r, column=1, value="Rating").font = FONT_LABEL
    ws.cell(row=r, column=2, value=nota).font = Font(name="Calibri", size=14, bold=True, color=_rating_color(nota))
    r += 1
    ws.cell(row=r, column=1, value="Parecer").font = FONT_LABEL
    ws.cell(row=r, column=2, value=_safe_str(rating.get("parecer"))).font = FONT_VALUE_BOLD
    r += 1
    ws.cell(row=r, column=1, value="Justificativa").font = FONT_LABEL
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.cell(row=r, column=2, value=_safe_str(rating.get("justificativa"))).font = FONT_VALUE
    r += 2

    r = _section_title(ws, r, "INDICADORES")
    _write_header_row(ws, r, ["Indicador", "Valor", "", "", "", "", ""])
    r += 1
    margem = _safe(kpis.get("margem_ebitda"))
    if margem > 1:
        margem = margem / 100
    ltv = _safe(kpis.get("ltv"))
    if ltv > 1:
        ltv = ltv / 100

    for label, val, fmt in [
        ("Receita Líquida", _safe(kpis.get("receita_liquida")), BRL_FORMAT),
        ("EBITDA", _safe(kpis.get("ebitda")), BRL_FORMAT),
        ("Margem EBITDA", margem, PCT_FORMAT),
        ("Dív.Líq./EBITDA", _safe(kpis.get("divida_liquida_ebitda")), MULT_FORMAT),
        ("LTV", ltv, PCT_FORMAT),
        ("DSCR", _safe(kpis.get("dscr")), MULT_FORMAT),
    ]:
        _write_row(ws, r, [label, val], [FONT_LABEL, FONT_VALUE_BOLD], [None, fmt])
        r += 1
