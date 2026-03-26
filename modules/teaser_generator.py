"""
ZYN Capital — Gerador de Teaser (.pptx) — 4 slides

Gera apresentação de teaser para investidores no padrão visual ZYN,
a partir da análise de crédito (MAC) já concluída.
"""

from datetime import datetime
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Cm, Emu, Pt

# ---------------------------------------------------------------------------
# ZYN Palette
# ---------------------------------------------------------------------------
NAVY = RGBColor(0x22, 0x30, 0x40)
DARK_SLATE = RGBColor(0x3A, 0x4F, 0x63)
GREEN = RGBColor(0x2E, 0x7D, 0x4F)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0x8B, 0x91, 0x97)
LIGHT_BG = RGBColor(0xF2, 0xF4, 0xF6)
GOLD = RGBColor(0x7D, 0x66, 0x08)
RED = RGBColor(0x92, 0x2B, 0x21)

SLIDE_W = Cm(33.867)  # 16:9
SLIDE_H = Cm(19.05)

FONT_NAME = "Montserrat"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _fmt_brl(value) -> str:
    try:
        v = float(value)
    except (TypeError, ValueError):
        return str(value)
    if v >= 1_000_000:
        return f"R$ {v / 1_000_000:,.1f} MM"
    if v >= 1_000:
        return f"R$ {v / 1_000:,.0f} mil"
    return f"R$ {v:,.0f}"


def _add_textbox(slide, left, top, width, height, text, font_size=12,
                 bold=False, color=NAVY, alignment=PP_ALIGN.LEFT,
                 font_name=FONT_NAME):
    """Add a simple textbox to a slide."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = str(text)
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name
    p.alignment = alignment
    return txBox


def _add_rect(slide, left, top, width, height, fill_color):
    """Add a filled rectangle shape."""
    shape = slide.shapes.add_shape(1, left, top, width, height)  # MSO_SHAPE.RECTANGLE
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape


def _add_table_row(table, row_idx, values, font_size=10, bold=False, color=NAVY, bg=None):
    """Fill a row in a table."""
    for col_idx, val in enumerate(values):
        cell = table.cell(row_idx, col_idx)
        cell.text = str(val)
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.size = Pt(font_size)
            paragraph.font.bold = bold
            paragraph.font.color.rgb = color
            paragraph.font.name = FONT_NAME
        if bg:
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg


def _rating_color(nota: str) -> RGBColor:
    return {"A": GREEN, "B": GREEN, "C": GOLD, "D": RED, "E": RED}.get(nota, GRAY)


# ---------------------------------------------------------------------------
# Slide builders
# ---------------------------------------------------------------------------
def _slide_cover(prs: Presentation, op: dict, analise: dict):
    """Slide 1 — Capa com dados da operação."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # Navy background
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = NAVY

    # Green accent bar top
    _add_rect(slide, Cm(0), Cm(0), SLIDE_W, Cm(0.6), GREEN)

    # ZYN Capital
    _add_textbox(slide, Cm(3), Cm(2.5), Cm(28), Cm(1.5),
                 "ZYN CAPITAL", font_size=14, bold=True, color=GRAY)
    _add_textbox(slide, Cm(3), Cm(3.8), Cm(28), Cm(1),
                 "CREDITO ESTRUTURADO & M&A", font_size=9, color=GRAY)

    # Tomador
    tomador = op.get("tomador", "Operacao")
    _add_textbox(slide, Cm(3), Cm(6.5), Cm(28), Cm(2.5),
                 tomador.upper(), font_size=32, bold=True, color=WHITE)

    # Tipo + Volume
    tipo = op.get("tipo_operacao", "")
    volume = _fmt_brl(op.get("volume", 0))
    _add_textbox(slide, Cm(3), Cm(9.5), Cm(28), Cm(1.5),
                 f"{tipo}  |  {volume}", font_size=18, color=GREEN)

    # Rating badge
    rating = analise.get("rating_final", {})
    nota = rating.get("nota", "—")
    parecer = rating.get("parecer", "—")
    _add_textbox(slide, Cm(3), Cm(12), Cm(28), Cm(1),
                 f"Rating: {nota}  |  Parecer: {parecer}", font_size=13, color=WHITE)

    # Date + Confidencial
    data = datetime.now().strftime("%B %Y")
    _add_textbox(slide, Cm(3), Cm(16), Cm(14), Cm(1),
                 data, font_size=10, color=GRAY)
    _add_textbox(slide, Cm(17), Cm(16), Cm(14), Cm(1),
                 "CONFIDENCIAL", font_size=10, bold=True, color=GRAY,
                 alignment=PP_ALIGN.RIGHT)

    # Green accent bar bottom
    _add_rect(slide, Cm(0), Cm(18.45), SLIDE_W, Cm(0.6), GREEN)


def _slide_overview(prs: Presentation, op: dict, analise: dict):
    """Slide 2 — Visão geral do tomador e grupo econômico."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Header bar
    _add_rect(slide, Cm(0), Cm(0), SLIDE_W, Cm(2.2), NAVY)
    _add_textbox(slide, Cm(2), Cm(0.4), Cm(20), Cm(1.5),
                 "VISAO GERAL", font_size=20, bold=True, color=WHITE)

    tomador_data = analise.get("tomador", {})
    patrimonio = analise.get("patrimonio", {})

    # Left column — Tomador info
    y = Cm(3)
    _add_textbox(slide, Cm(2), y, Cm(14), Cm(0.8),
                 "TOMADOR", font_size=9, bold=True, color=GRAY)
    y += Cm(0.9)
    _add_textbox(slide, Cm(2), y, Cm(14), Cm(0.8),
                 tomador_data.get("razao_social", op.get("tomador", "—")),
                 font_size=14, bold=True, color=NAVY)
    y += Cm(1.2)
    _add_textbox(slide, Cm(2), y, Cm(14), Cm(0.6),
                 f"CNPJ: {tomador_data.get('cnpj', op.get('cnpj', '—'))}",
                 font_size=10, color=DARK_SLATE)
    y += Cm(1)

    grupo = tomador_data.get("grupo_economico", "—")
    if grupo and grupo != "—":
        _add_textbox(slide, Cm(2), y, Cm(14), Cm(0.8),
                     "GRUPO ECONOMICO", font_size=9, bold=True, color=GRAY)
        y += Cm(0.9)
        # Truncate if too long
        if len(grupo) > 300:
            grupo = grupo[:300] + "..."
        _add_textbox(slide, Cm(2), y, Cm(14), Cm(3),
                     grupo, font_size=9, color=DARK_SLATE)

    # Right column — Key info box
    _add_rect(slide, Cm(18), Cm(3), Cm(14), Cm(13), LIGHT_BG)

    info_items = [
        ("Tipo", op.get("tipo_operacao", "—")),
        ("Volume", _fmt_brl(op.get("volume", 0))),
        ("Prazo", f"{op.get('prazo_meses', '—')} meses"),
        ("Taxa", op.get("taxa", "—")),
        ("Amortizacao", op.get("amortizacao", "—")),
        ("Garantias", op.get("garantias_text", "—") or "—"),
        ("Instrumento", op.get("instrumento", "—")),
    ]

    iy = Cm(3.5)
    _add_textbox(slide, Cm(19), iy, Cm(12), Cm(0.8),
                 "ESTRUTURA DA OPERACAO", font_size=9, bold=True, color=GREEN)
    iy += Cm(1.2)

    for label, valor in info_items:
        _add_textbox(slide, Cm(19), iy, Cm(5), Cm(0.6),
                     label, font_size=8, bold=True, color=GRAY)
        val_str = str(valor)
        if len(val_str) > 60:
            val_str = val_str[:60] + "..."
        _add_textbox(slide, Cm(24), iy, Cm(8), Cm(0.6),
                     val_str, font_size=9, color=NAVY)
        iy += Cm(1.1)

    # Green bar bottom
    _add_rect(slide, Cm(0), Cm(18.45), SLIDE_W, Cm(0.6), GREEN)


def _slide_financials(prs: Presentation, analise: dict):
    """Slide 3 — Indicadores financeiros e KPIs."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Header bar
    _add_rect(slide, Cm(0), Cm(0), SLIDE_W, Cm(2.2), NAVY)
    _add_textbox(slide, Cm(2), Cm(0.4), Cm(20), Cm(1.5),
                 "INDICADORES FINANCEIROS", font_size=20, bold=True, color=WHITE)

    kpis = analise.get("kpis", {})
    capital = analise.get("capital", {})
    indicadores = capital.get("indicadores", {})

    # KPI cards — 3x2 grid
    kpi_items = [
        ("Receita Liquida", _fmt_brl(kpis.get("receita_liquida", 0))),
        ("EBITDA", _fmt_brl(kpis.get("ebitda", 0))),
        ("Margem EBITDA", f"{kpis.get('margem_ebitda', 0):.1f}%" if isinstance(kpis.get("margem_ebitda"), (int, float)) else "—"),
        ("Div. Liq./EBITDA", f"{kpis.get('divida_liquida_ebitda', 0):.2f}x"),
        ("DSCR", f"{kpis.get('dscr', 0):.2f}x"),
        ("LTV", f"{kpis.get('ltv', 0):.1%}" if isinstance(kpis.get("ltv"), (int, float)) and 0 < kpis.get("ltv", 0) <= 1 else f"{kpis.get('ltv', 0):.1f}%"),
    ]

    col_w = Cm(9.5)
    row_h = Cm(3.5)
    start_x = Cm(2)
    start_y = Cm(3.2)
    gap = Cm(0.8)

    for idx, (label, valor) in enumerate(kpi_items):
        col = idx % 3
        row = idx // 3
        x = start_x + col * (col_w + gap)
        y = start_y + row * (row_h + gap)

        _add_rect(slide, x, y, col_w, row_h, LIGHT_BG)
        _add_textbox(slide, x + Cm(0.8), y + Cm(0.5), col_w - Cm(1.6), Cm(0.7),
                     label.upper(), font_size=8, bold=True, color=GRAY)
        _add_textbox(slide, x + Cm(0.8), y + Cm(1.3), col_w - Cm(1.6), Cm(1.8),
                     valor, font_size=22, bold=True, color=NAVY)

    # Ratings por seção — compact table
    secoes = ["tomador", "patrimonio", "producao", "capital", "operacao",
              "pagamento", "onus", "riscos", "covenants", "cronograma"]

    y_table = Cm(11.5)
    _add_textbox(slide, Cm(2), y_table, Cm(20), Cm(0.8),
                 "RATINGS POR SECAO", font_size=9, bold=True, color=GRAY)
    y_table += Cm(1)

    for idx, secao in enumerate(secoes):
        dados_secao = analise.get(secao, {})
        rating_secao = dados_secao.get("rating_secao", "N/A")
        col = idx % 5
        row = idx // 5
        x = Cm(2) + col * Cm(6.2)
        y = y_table + row * Cm(1.5)

        cor = {"Forte": GREEN, "Adequado": DARK_SLATE, "Atenção": GOLD, "Crítico": RED}.get(rating_secao, GRAY)
        label = secao.replace("_", " ").title()
        _add_textbox(slide, x, y, Cm(3), Cm(0.6),
                     label, font_size=8, color=GRAY)
        _add_textbox(slide, x + Cm(3.2), y, Cm(2.8), Cm(0.6),
                     rating_secao, font_size=9, bold=True, color=cor)

    # Green bar bottom
    _add_rect(slide, Cm(0), Cm(18.45), SLIDE_W, Cm(0.6), GREEN)


def _slide_conclusion(prs: Presentation, op: dict, analise: dict):
    """Slide 4 — Parecer, riscos e recomendações."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Header bar
    _add_rect(slide, Cm(0), Cm(0), SLIDE_W, Cm(2.2), NAVY)
    _add_textbox(slide, Cm(2), Cm(0.4), Cm(20), Cm(1.5),
                 "PARECER E RISCOS", font_size=20, bold=True, color=WHITE)

    rating = analise.get("rating_final", {})
    nota = rating.get("nota", "—")
    parecer = rating.get("parecer", "—")

    # Rating + Parecer box
    cor_nota = _rating_color(nota)
    _add_rect(slide, Cm(2), Cm(3), Cm(7), Cm(3.5), LIGHT_BG)
    _add_textbox(slide, Cm(2.8), Cm(3.3), Cm(5), Cm(0.7),
                 "RATING FINAL", font_size=9, bold=True, color=GRAY)
    _add_textbox(slide, Cm(2.8), Cm(4), Cm(5), Cm(2),
                 nota, font_size=48, bold=True, color=cor_nota)

    _add_rect(slide, Cm(10), Cm(3), Cm(10), Cm(3.5), LIGHT_BG)
    _add_textbox(slide, Cm(10.8), Cm(3.3), Cm(8), Cm(0.7),
                 "PARECER", font_size=9, bold=True, color=GRAY)
    cor_parecer = GREEN if "Favorável" in parecer and "Ressalvas" not in parecer else (GOLD if "Ressalvas" in parecer else RED)
    _add_textbox(slide, Cm(10.8), Cm(4.2), Cm(8), Cm(1.5),
                 parecer, font_size=18, bold=True, color=cor_parecer)

    # Justificativa
    justificativa = rating.get("justificativa", "—")
    if len(justificativa) > 400:
        justificativa = justificativa[:400] + "..."
    _add_textbox(slide, Cm(2), Cm(7.2), Cm(18), Cm(0.7),
                 "JUSTIFICATIVA", font_size=9, bold=True, color=GRAY)
    _add_textbox(slide, Cm(2), Cm(8), Cm(18), Cm(3),
                 justificativa, font_size=9, color=DARK_SLATE)

    # Top risks (from matriz_riscos)
    riscos = analise.get("riscos", {})
    matriz = riscos.get("matriz_riscos", [])

    y_risk = Cm(11.5)
    _add_textbox(slide, Cm(2), y_risk, Cm(15), Cm(0.7),
                 "PRINCIPAIS RISCOS", font_size=9, bold=True, color=GRAY)
    y_risk += Cm(0.9)

    for risco in matriz[:4]:  # top 4 risks
        nome = risco.get("risco", "—")
        prob = risco.get("probabilidade", "—")
        impacto = risco.get("impacto", "—")
        mitigante = risco.get("mitigante", "—")

        if len(nome) > 50:
            nome = nome[:50] + "..."
        if len(mitigante) > 80:
            mitigante = mitigante[:80] + "..."

        cor_risco = RED if impacto == "Alto" else (GOLD if impacto == "Médio" else GREEN)
        _add_textbox(slide, Cm(2), y_risk, Cm(0.4), Cm(0.5),
                     "●", font_size=10, color=cor_risco)
        _add_textbox(slide, Cm(2.6), y_risk, Cm(17), Cm(0.5),
                     f"{nome} ({prob}/{impacto}) — {mitigante}",
                     font_size=8, color=DARK_SLATE)
        y_risk += Cm(1)

    # Recommendations
    recs = rating.get("recomendacoes", [])
    if recs:
        y_rec = Cm(16)
        _add_textbox(slide, Cm(22), Cm(3), Cm(10), Cm(0.7),
                     "RECOMENDACOES", font_size=9, bold=True, color=GRAY)
        y_r = Cm(3.9)
        for rec in recs[:6]:
            if len(rec) > 80:
                rec = rec[:80] + "..."
            _add_textbox(slide, Cm(22), y_r, Cm(10), Cm(1),
                         f"• {rec}", font_size=8, color=DARK_SLATE)
            y_r += Cm(1.1)

    # Footer
    _add_rect(slide, Cm(0), Cm(18.45), SLIDE_W, Cm(0.6), GREEN)
    _add_textbox(slide, Cm(2), Cm(17.5), Cm(15), Cm(0.6),
                 "ZYN Capital  |  Credito Estruturado & M&A  |  Confidencial",
                 font_size=7, color=GRAY)


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------
def generate_teaser(
    analise: dict[str, Any],
    parametros: dict[str, Any],
    output_path: str,
) -> str:
    """
    Gera Teaser ZYN de 4 slides (.pptx).

    Args:
        analise: Resultado da análise de crédito (MAC).
        parametros: Parâmetros da operação (tomador, volume, etc.).
        output_path: Caminho de saída do arquivo .pptx.

    Returns:
        Caminho do arquivo gerado.
    """
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    _slide_cover(prs, parametros, analise)
    _slide_overview(prs, parametros, analise)
    _slide_financials(prs, analise)
    _slide_conclusion(prs, parametros, analise)

    prs.save(output_path)
    return output_path
