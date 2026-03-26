"""
ZYN Capital — Gerador de Teaser (.pptx) — 4 slides
Padrão visual idêntico ao template Grand Laguna ZYN 2026.

Fundo branco #F3F4F6, watermark CONFIDENCIAL, barra vertical verde,
barra inferior verde, tipografia sem serifa, paleta ZYN.
"""

from datetime import datetime
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt

# ---------------------------------------------------------------------------
# ZYN Palette — extraída pixel a pixel do Teaser_Grand_Laguna_ZYN2026.pptx
# ---------------------------------------------------------------------------
SLIDE_BG = RGBColor(0xF3, 0xF4, 0xF6)       # fundo claro de todos os slides
TITLE_COLOR = RGBColor(0x3B, 0x4C, 0x5E)     # títulos e valores
SUBTITLE_COLOR = RGBColor(0x8B, 0x95, 0xA5)  # subtítulos de seção
BODY_COLOR = RGBColor(0x6B, 0x72, 0x80)       # texto corpo
LABEL_COLOR = RGBColor(0x9C, 0xA3, 0xAF)      # labels KPI
WATERMARK_COLOR = RGBColor(0xF3, 0xF4, 0xF6)  # = bg (aparece ao imprimir)
GREEN_BAR = RGBColor(0x2E, 0x7D, 0x4F)        # barras verdes e accent
NAVY = RGBColor(0x22, 0x30, 0x40)             # table headers
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ROW_ODD = RGBColor(0xF9, 0xFA, 0xFB)          # linhas alternadas tabela
ROW_EVEN = RGBColor(0xF2, 0xF3, 0xF5)
GOLD = RGBColor(0x7D, 0x66, 0x08)
RED = RGBColor(0x92, 0x2B, 0x21)

# Slide dimensions — standard 16:9 (same as reference)
SW = 12192000  # EMU
SH = 6858000

# Margens e medidas recorrentes (EMU) extraídas da referência
LEFT_MARGIN = 612648
GREEN_BAR_W = 73152          # barra vertical verde
GREEN_BAR_BOTTOM_H = 155448  # barra inferior
SECTION_LINE_Y_OFFSET = 0   # linhas de seção
DATE_X = 10250424
DATE_W = 1344168

MESES_PT = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL",
    5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
    9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO",
}


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


def _trunc(text: str, limit: int = 120) -> str:
    text = str(text)
    return text[:limit] + "..." if len(text) > limit else text


def _rating_color(nota: str) -> RGBColor:
    return {"A": GREEN_BAR, "B": GREEN_BAR, "C": GOLD, "D": RED, "E": RED}.get(nota, LABEL_COLOR)


def _add_rect(slide, left, top, width, height, fill_color):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape


def _txt(slide, left, top, width, height, text, size=12, bold=False,
         color=TITLE_COLOR, align=PP_ALIGN.LEFT):
    """Add textbox with consistent settings."""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = str(text)
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.alignment = align
    return box


def _hline(slide, left, top, width):
    """Horizontal thin line (connector-like via thin rectangle)."""
    shape = slide.shapes.add_shape(1, left, top, width, Emu(1))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0xE0, 0xE2, 0xE6)
    shape.line.fill.background()
    return shape


def _base_slide(prs, green_bar_h=None):
    """Create slide with standard ZYN elements: bg, watermark, green bar bottom, ZYN logo, date."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Full slide background
    _add_rect(slide, 0, 0, SW, SH, SLIDE_BG)

    # Watermark CONFIDENCIAL (large, faint — same as reference)
    _txt(slide, 1097280, 2651760, 10058400, 1645920,
         "CONFIDENCIAL", size=90, bold=True, color=WATERMARK_COLOR)

    # Green bar bottom
    _add_rect(slide, 0, SH - GREEN_BAR_BOTTOM_H, SW, GREEN_BAR_BOTTOM_H, GREEN_BAR)

    # ZYN logo text
    _txt(slide, LEFT_MARGIN, 457200, 1252728, 566928,
         "ZYN", size=45, bold=True, color=TITLE_COLOR)
    _txt(slide, LEFT_MARGIN + 36576, 1024128, 1197864, 301752,
         "CAPITAL", size=18, color=SUBTITLE_COLOR)

    # Date block (top right)
    now = datetime.now()
    mes = MESES_PT.get(now.month, now.strftime("%B").upper())
    _txt(slide, DATE_X, 667512, DATE_W, 155448,
         "Data", size=9, bold=True, color=LABEL_COLOR)
    _txt(slide, DATE_X, 859536, DATE_W, 265176,
         f"{mes} / {now.year}", size=13, bold=True, color=TITLE_COLOR)

    # Green accent bar (vertical, next to section title)
    if green_bar_h:
        _add_rect(slide, LEFT_MARGIN, 1691640, GREEN_BAR_W, green_bar_h, GREEN_BAR)

    return slide


# ---------------------------------------------------------------------------
# Slide 1 — Capa (Apresentação)
# ---------------------------------------------------------------------------
def _slide_cover(prs, op, analise):
    slide = _base_slide(prs, green_bar_h=1371600)

    tomador = op.get("tomador", "Operacao")
    tipo = op.get("tipo_operacao", "")
    volume = _fmt_brl(op.get("volume", 0))
    rating = analise.get("rating_final", {})
    nota = rating.get("nota", "")
    parecer = rating.get("parecer", "")

    # Section subtitle
    _txt(slide, LEFT_MARGIN + 274320, 1682496, 10707624, 265176,
         "Apresentacao", size=15, bold=True, color=SUBTITLE_COLOR)

    # Tomador name (big)
    _txt(slide, LEFT_MARGIN + 374904, 2039112, 9144000, 530352,
         tomador, size=36, bold=True, color=TITLE_COLOR)

    # Subtitle line
    subtitle = f"{tipo}  ·  Operacao de Credito Estruturado"
    _txt(slide, LEFT_MARGIN + 374904, 2606040, 9144000, 320040,
         subtitle, size=18, color=BODY_COLOR)

    # Divider
    _hline(slide, LEFT_MARGIN, 3154680, 10972800)

    # 4 KPI cards
    kpis = analise.get("kpis", {})
    cards = [
        ("Volume", volume, f"Rating: {nota}" if nota else ""),
        ("DSCR", f"{kpis.get('dscr', 0):.2f}x", "Debt Service Coverage"),
        ("LTV", f"{kpis.get('ltv', 0):.1%}" if isinstance(kpis.get('ltv'), (int, float)) and 0 < kpis.get('ltv', 0) <= 1 else f"{kpis.get('ltv', 0):.1f}%", "Loan to Value"),
        ("Parecer", parecer, f"Nota {nota}"),
    ]

    card_w = 2560320
    card_gap = card_w + 411480
    cards_y_top = 3337560

    for idx, (label, value, sub) in enumerate(cards):
        x = LEFT_MARGIN + idx * card_gap
        # Top line
        _hline(slide, x, cards_y_top, card_w)
        # Label
        _txt(slide, x, cards_y_top + 91440, card_w, 201168,
             label, size=9, bold=True, color=LABEL_COLOR)
        # Value
        _txt(slide, x, cards_y_top + 292608, card_w, 475488,
             value, size=26, bold=True, color=TITLE_COLOR)
        # Sub
        _txt(slide, x, cards_y_top + 768096, card_w, 201168,
             sub, size=10, color=BODY_COLOR)

    # Bottom divider
    _hline(slide, LEFT_MARGIN, 5029200, 10972800)

    # Bottom info line
    cnpj = op.get("cnpj", "")
    prazo = op.get("prazo_meses", "")
    taxa = op.get("taxa", "")
    info_parts = [f"Tomador: {tomador}"]
    if cnpj:
        info_parts.append(f"CNPJ {cnpj}")
    if prazo:
        info_parts.append(f"Prazo: {prazo} meses")
    if taxa:
        info_parts.append(f"Taxa: {taxa}")
    _txt(slide, LEFT_MARGIN, 5102352, 10972800, 219456,
         "  ·  ".join(info_parts), size=10, color=BODY_COLOR)


# ---------------------------------------------------------------------------
# Slide 2 — Visao Geral + Estrutura
# ---------------------------------------------------------------------------
def _slide_overview(prs, op, analise):
    slide = _base_slide(prs, green_bar_h=1005840)

    tomador_data = analise.get("tomador", {})

    # Section subtitle
    _txt(slide, LEFT_MARGIN + 274320, 1682496, 10707624, 265176,
         "Visao Corporativa", size=15, bold=True, color=SUBTITLE_COLOR)

    # Title
    grupo = tomador_data.get("grupo_economico", "")
    titulo = _trunc(grupo, 80) if grupo and grupo != "—" else op.get("tomador", "—")
    _txt(slide, LEFT_MARGIN + 374904, 2039112, 10058400, 530352,
         titulo, size=32, bold=True, color=TITLE_COLOR)

    # Divider
    _hline(slide, LEFT_MARGIN, 2633472, 11277295)

    # Left side: 3 KPI cards
    kpis = analise.get("kpis", {})
    left_cards = [
        ("Receita Liquida", _fmt_brl(kpis.get("receita_liquida", 0)),
         f"EBITDA: {_fmt_brl(kpis.get('ebitda', 0))}"),
        ("Margem EBITDA",
         f"{kpis.get('margem_ebitda', 0):.1f}%" if isinstance(kpis.get("margem_ebitda"), (int, float)) else "—",
         f"Div.Liq./EBITDA: {kpis.get('divida_liquida_ebitda', 0):.2f}x"),
        ("Patrimonio",
         tomador_data.get("rating_secao", "—"),
         _trunc(analise.get("patrimonio", {}).get("avaliacao", "—"), 60)),
    ]

    card_w = 3291840
    card_gap = card_w + 365760

    for idx, (label, value, sub) in enumerate(left_cards):
        x = LEFT_MARGIN + idx * card_gap
        y = 3127248
        _hline(slide, x, y, card_w)
        _txt(slide, x, y + 91440, card_w, 201168,
             label, size=9, bold=True, color=LABEL_COLOR)
        _txt(slide, x, y + 292608, card_w, 475488,
             value, size=26, bold=True, color=TITLE_COLOR)
        _txt(slide, x, y + 768096, card_w, 201168,
             sub, size=10, color=BODY_COLOR)

    # Divider below cards
    _hline(slide, LEFT_MARGIN, 4462272, 10972800)

    # Left: Perfil da Operacao (key-value pairs)
    _txt(slide, LEFT_MARGIN, 4572000, 5120640, 237744,
         "Estrutura da Operacao", size=11, bold=True, color=TITLE_COLOR)

    info_items = [
        ("Tipo", op.get("tipo_operacao", "—")),
        ("Volume", _fmt_brl(op.get("volume", 0))),
        ("Prazo", f"{op.get('prazo_meses', '—')} meses"),
        ("Taxa", op.get("taxa", "—")),
        ("Amortizacao", op.get("amortizacao", "—")),
        ("Garantias", _trunc(op.get("garantias_text", "—") or "—", 50)),
    ]

    y = 4828032
    for label, valor in info_items:
        _txt(slide, LEFT_MARGIN, y, 1920240, 237744,
             label, size=9.5, bold=True, color=BODY_COLOR)
        _txt(slide, LEFT_MARGIN + 1965960, y, 3200400, 237744,
             valor, size=9.5, color=BODY_COLOR)
        y += 256032

    # Right: Ratings por secao (table)
    table_x = 6217920
    table_w = 5367528
    _txt(slide, table_x, 4572000, table_w, 237744,
         "Ratings por Secao", size=11, bold=True, color=TITLE_COLOR)

    # Table header
    _add_rect(slide, table_x, 4828032, table_w, 237744, NAVY)
    _txt(slide, table_x + 45720, 4846320, 2743200, 201168,
         "Secao", size=8, bold=True, color=WHITE)
    _txt(slide, table_x + 2834640, 4846320, 1371600, 201168,
         "Rating", size=8, bold=True, color=WHITE)
    _txt(slide, table_x + 4206240, 4846320, 1143000, 201168,
         "Status", size=8, bold=True, color=WHITE)

    secoes = ["Tomador", "Patrimonio", "Producao", "Capital", "Operacao",
              "Pagamento", "Onus", "Riscos", "Covenants", "Cronograma"]

    row_y = 5084064
    for idx, secao in enumerate(secoes):
        key = secao.lower().replace("ô", "o")
        dados = analise.get(key, {})
        rating_s = dados.get("rating_secao", "N/A")
        n_flags = len(dados.get("flags", []))
        status = f"{n_flags} flag(s)" if n_flags > 0 else "OK"

        bg = ROW_ODD if idx % 2 == 0 else ROW_EVEN
        _add_rect(slide, table_x, row_y, table_w, 237744, bg)
        _txt(slide, table_x + 45720, row_y + 18288, 2743200, 201168,
             secao, size=8.5, bold=True, color=TITLE_COLOR)

        cor_r = {"Forte": GREEN_BAR, "Adequado": TITLE_COLOR, "Atenção": GOLD, "Crítico": RED}.get(rating_s, LABEL_COLOR)
        _txt(slide, table_x + 2834640, row_y + 18288, 1371600, 201168,
             rating_s, size=8.5, bold=True, color=cor_r)
        _txt(slide, table_x + 4206240, row_y + 18288, 1143000, 201168,
             status, size=8.5, color=BODY_COLOR)
        row_y += 246888


# ---------------------------------------------------------------------------
# Slide 3 — Indicadores Financeiros
# ---------------------------------------------------------------------------
def _slide_financials(prs, analise):
    slide = _base_slide(prs, green_bar_h=1005840)

    kpis = analise.get("kpis", {})
    capital = analise.get("capital", {})
    indicadores = capital.get("indicadores", {})

    _txt(slide, LEFT_MARGIN + 274320, 1682496, 10707624, 265176,
         "Analise Financeira", size=15, bold=True, color=SUBTITLE_COLOR)
    _txt(slide, LEFT_MARGIN + 374904, 2039112, 10058400, 530352,
         "Indicadores Financeiros & Capacidade de Pagamento",
         size=32, bold=True, color=TITLE_COLOR)

    _hline(slide, LEFT_MARGIN, 2633472, 11277295)

    # 4 KPI big cards
    big_cards = [
        ("Receita Liquida", _fmt_brl(kpis.get("receita_liquida", 0)),
         "Ultimo exercicio"),
        ("EBITDA", _fmt_brl(kpis.get("ebitda", 0)),
         f"Margem: {kpis.get('margem_ebitda', 0):.1f}%" if isinstance(kpis.get("margem_ebitda"), (int, float)) else "—"),
        ("DSCR", f"{kpis.get('dscr', 0):.2f}x",
         "Debt Service Coverage Ratio"),
        ("Div. Liq./EBITDA", f"{kpis.get('divida_liquida_ebitda', 0):.2f}x",
         "Alavancagem"),
    ]

    card_w = 2560320
    card_gap = card_w + 411480

    for idx, (label, value, sub) in enumerate(big_cards):
        x = LEFT_MARGIN + idx * card_gap
        y = 3127248
        _hline(slide, x, y, card_w)
        _txt(slide, x, y + 91440, card_w, 201168,
             label, size=9, bold=True, color=LABEL_COLOR)
        _txt(slide, x, y + 292608, card_w, 475488,
             value, size=26, bold=True, color=TITLE_COLOR)
        _txt(slide, x, y + 768096, card_w, 201168,
             sub, size=10, color=BODY_COLOR)

    _hline(slide, LEFT_MARGIN, 4462272, 10972800)

    # Bottom section: more indicators + LTV/payment
    _txt(slide, LEFT_MARGIN, 4572000, 5120640, 237744,
         "Indicadores de Capital", size=11, bold=True, color=TITLE_COLOR)

    ind_items = [
        ("Divida/PL", f"{indicadores.get('divida_pl', 0):.2f}x"),
        ("Liquidez Corrente", f"{indicadores.get('liquidez_corrente', 0):.2f}"),
        ("ROE", f"{indicadores.get('roe', 0):.1%}" if isinstance(indicadores.get("roe"), (int, float)) and 0 < abs(indicadores.get("roe", 0)) <= 1 else f"{indicadores.get('roe', 0):.1f}%"),
        ("LTV Efetivo", f"{kpis.get('ltv', 0):.1%}" if isinstance(kpis.get("ltv"), (int, float)) and 0 < kpis.get("ltv", 0) <= 1 else f"{kpis.get('ltv', 0):.1f}%"),
    ]

    y = 4828032
    for label, valor in ind_items:
        _txt(slide, LEFT_MARGIN, y, 1920240, 237744,
             label, size=9.5, bold=True, color=BODY_COLOR)
        _txt(slide, LEFT_MARGIN + 1965960, y, 2286000, 237744,
             valor, size=9.5, color=BODY_COLOR)
        y += 256032

    # Right: Payment analysis
    pagamento = analise.get("pagamento", {})
    _txt(slide, 6217920, 4572000, 5367528, 237744,
         "Capacidade de Pagamento", size=11, bold=True, color=TITLE_COLOR)

    pay_text = _trunc(pagamento.get("analise", "—"), 500)
    _txt(slide, 6217920, 4828032, 5367528, 1828800,
         pay_text, size=9, color=BODY_COLOR)


# ---------------------------------------------------------------------------
# Slide 4 — Parecer & Riscos
# ---------------------------------------------------------------------------
def _slide_conclusion(prs, op, analise):
    slide = _base_slide(prs, green_bar_h=1005840)

    rating = analise.get("rating_final", {})
    nota = rating.get("nota", "—")
    parecer = rating.get("parecer", "—")

    _txt(slide, LEFT_MARGIN + 274320, 1682496, 10707624, 265176,
         "Parecer & Riscos", size=15, bold=True, color=SUBTITLE_COLOR)
    _txt(slide, LEFT_MARGIN + 374904, 2039112, 10058400, 530352,
         f"Rating {nota}  ·  {parecer}",
         size=32, bold=True, color=_rating_color(nota))

    _hline(slide, LEFT_MARGIN, 2633472, 11277295)

    # 4 summary cards
    kpis = analise.get("kpis", {})
    summary_cards = [
        ("Rating Final", nota, parecer),
        ("DSCR", f"{kpis.get('dscr', 0):.2f}x",
         "Cobertura do servico da divida"),
        ("LTV", f"{kpis.get('ltv', 0):.1%}" if isinstance(kpis.get("ltv"), (int, float)) and 0 < kpis.get("ltv", 0) <= 1 else f"{kpis.get('ltv', 0):.1f}%",
         "Loan to Value ajustado"),
        ("Flags",
         str(sum(len(analise.get(s, {}).get("flags", [])) for s in ["tomador", "patrimonio", "producao", "capital", "operacao", "pagamento", "onus", "riscos", "covenants", "cronograma"])),
         "Pontos de atencao"),
    ]

    card_w = 2560320
    card_gap = card_w + 411480

    for idx, (label, value, sub) in enumerate(summary_cards):
        x = LEFT_MARGIN + idx * card_gap
        y = 3127248
        _hline(slide, x, y, card_w)
        _txt(slide, x, y + 91440, card_w, 201168,
             label, size=9, bold=True, color=LABEL_COLOR)
        color = _rating_color(nota) if idx == 0 else TITLE_COLOR
        _txt(slide, x, y + 292608, card_w, 475488,
             value, size=26, bold=True, color=color)
        _txt(slide, x, y + 768096, card_w, 201168,
             sub, size=10, color=BODY_COLOR)

    _hline(slide, LEFT_MARGIN, 4462272, 10972800)

    # Left: Justificativa
    _txt(slide, LEFT_MARGIN, 4572000, 5120640, 237744,
         "Justificativa", size=11, bold=True, color=TITLE_COLOR)

    justificativa = _trunc(rating.get("justificativa", "—"), 500)
    _txt(slide, LEFT_MARGIN, 4828032, 5120640, 1371600,
         justificativa, size=9, color=BODY_COLOR)

    # Right: Riscos table
    riscos = analise.get("riscos", {})
    matriz = riscos.get("matriz_riscos", [])

    table_x = 6217920
    table_w = 5367528
    _txt(slide, table_x, 4572000, table_w, 237744,
         "Principais Riscos", size=11, bold=True, color=TITLE_COLOR)

    # Table header
    _add_rect(slide, table_x, 4828032, table_w, 237744, NAVY)
    _txt(slide, table_x + 45720, 4846320, 2286000, 201168,
         "Risco", size=8, bold=True, color=WHITE)
    _txt(slide, table_x + 2377440, 4846320, 914400, 201168,
         "Prob.", size=8, bold=True, color=WHITE)
    _txt(slide, table_x + 3337560, 4846320, 914400, 201168,
         "Impacto", size=8, bold=True, color=WHITE)
    _txt(slide, table_x + 4297680, 4846320, 1051560, 201168,
         "Mitigante", size=8, bold=True, color=WHITE)

    row_y = 5084064
    for idx, risco in enumerate(matriz[:6]):
        nome = _trunc(risco.get("risco", "—"), 30)
        prob = risco.get("probabilidade", "—")
        impacto = risco.get("impacto", "—")
        mitigante = _trunc(risco.get("mitigante", "—"), 18)

        bg = ROW_ODD if idx % 2 == 0 else ROW_EVEN
        _add_rect(slide, table_x, row_y, table_w, 237744, bg)

        _txt(slide, table_x + 45720, row_y + 18288, 2286000, 201168,
             nome, size=8.5, bold=True, color=TITLE_COLOR)

        cor_prob = RED if prob == "Alta" else (GOLD if prob in ("Média", "Media") else GREEN_BAR)
        _txt(slide, table_x + 2377440, row_y + 18288, 914400, 201168,
             prob, size=8.5, bold=True, color=cor_prob)

        cor_imp = RED if impacto == "Alto" else (GOLD if impacto in ("Médio", "Medio") else GREEN_BAR)
        _txt(slide, table_x + 3337560, row_y + 18288, 914400, 201168,
             impacto, size=8.5, bold=True, color=cor_imp)

        _txt(slide, table_x + 4297680, row_y + 18288, 1051560, 201168,
             mitigante, size=8, color=BODY_COLOR)

        row_y += 246888

    # Footer — disclaimer
    _txt(slide, LEFT_MARGIN, SH - 548640, 10972800, 219456,
         "ZYN Capital  ·  Credito Estruturado & M&A  ·  Confidencial  ·  Este documento nao constitui oferta de valores mobiliarios",
         size=8, color=LABEL_COLOR)


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------
def generate_teaser(
    analise: dict[str, Any],
    parametros: dict[str, Any],
    output_path: str,
) -> str:
    """
    Gera Teaser ZYN de 4 slides (.pptx) no padrao visual Grand Laguna.

    Args:
        analise: Resultado da análise de crédito (MAC).
        parametros: Parâmetros da operação (tomador, volume, etc.).
        output_path: Caminho de saída do arquivo .pptx.

    Returns:
        Caminho do arquivo gerado.
    """
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH

    _slide_cover(prs, parametros, analise)
    _slide_overview(prs, parametros, analise)
    _slide_financials(prs, analise)
    _slide_conclusion(prs, parametros, analise)

    prs.save(output_path)
    return output_path
