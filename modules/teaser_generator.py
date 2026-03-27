"""
ZYN Capital — Gerador de Teaser Premium (.pptx) — 8-10 slides
Baseado na analise de 8 teasers reais ZYN (SNOW FIDC, Ivanoff SLB, Baroon55 CRI,
Grupo Roca, Goulart, Edificatto, Frigomarca, Fibra Cotton).

Padrao visual institucional:
- Fundo navy escuro (#1B2838) com paineis de conteudo navy medio (#223040)
- KPI cards com valores grandes em branco, labels em cinza
- Tabelas com header navy, linhas alternadas
- Barra lateral verde accent, footer institucional
- Watermark CONFIDENCIAL

Suporta todos os produtos: CRI, CRA, CPR-F, SLB, NC/CCB, FIDC, Fiagro, Debenture.
"""

from datetime import datetime
from typing import Any, Optional

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------------------
# Paleta ZYN — extraida dos 8 teasers reais
# ---------------------------------------------------------------------------
NAVY_DARK = RGBColor(0x1B, 0x28, 0x38)        # fundo principal dos slides
NAVY = RGBColor(0x22, 0x30, 0x40)             # paineis, headers de tabela
NAVY_LIGHT = RGBColor(0x2C, 0x3E, 0x50)       # paineis secundarios
TITLE_COLOR = RGBColor(0xFF, 0xFF, 0xFF)       # titulos sobre fundo escuro
TITLE_DARK = RGBColor(0x22, 0x30, 0x40)        # titulos sobre fundo claro
SUBTITLE_COLOR = RGBColor(0x8B, 0x91, 0x97)   # subtitulos e labels
BODY_COLOR = RGBColor(0x34, 0x40, 0x50)        # texto corpo (fundo claro)
BODY_LIGHT = RGBColor(0xD1, 0xD5, 0xDB)       # texto corpo (fundo escuro)
LABEL_COLOR = RGBColor(0x9C, 0xA3, 0xAF)      # labels de KPI
GREEN = RGBColor(0x2E, 0x7D, 0x4F)            # accent verde, badges positivos
GREEN_LIGHT = RGBColor(0x05, 0x96, 0x69)       # indicadores positivos
BLUE_ACCENT = RGBColor(0x60, 0xA5, 0xFA)      # destaques, badges
BLUE_LIGHT = RGBColor(0xBF, 0xDB, 0xFE)       # texto accent claro
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GOLD = RGBColor(0xEA, 0xB3, 0x08)             # atencao
RED = RGBColor(0xEF, 0x44, 0x44)              # critico
ROW_ODD = RGBColor(0xF9, 0xFA, 0xFB)          # linhas alternadas tabela (claro)
ROW_EVEN = RGBColor(0xF2, 0xF3, 0xF5)
ROW_DARK_ODD = RGBColor(0x1E, 0x2D, 0x3D)     # linhas alternadas (escuro)
ROW_DARK_EVEN = RGBColor(0x24, 0x34, 0x47)
WATERMARK_COLOR = RGBColor(0x1F, 0x2C, 0x3C)  # watermark quase invisivel
PANEL_BG = RGBColor(0xF3, 0xF4, 0xF6)         # paineis brancos
DARK_OVERLAY = RGBColor(0x0F, 0x17, 0x2A)     # overlay escuro para cover

# Slide dimensions — 16:9 widescreen (padrao dos teasers mais recentes)
SW = 12192000  # EMU
SH = 6858000

# Margens e medidas recorrentes (EMU)
LEFT_MARGIN = Emu(457200)       # ~0.5 inch
RIGHT_MARGIN = Emu(457200)
TOP_MARGIN = Emu(457200)
CONTENT_W = SW - 2 * LEFT_MARGIN.emu if hasattr(LEFT_MARGIN, 'emu') else SW - 2 * 457200
HEADER_H = Emu(800100)          # altura da barra de header
FOOTER_H = Emu(365760)          # altura do footer
GREEN_BAR_W = Emu(54864)        # barra vertical verde accent

MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Marco", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}


# ---------------------------------------------------------------------------
# Helpers — Formatacao BR
# ---------------------------------------------------------------------------
def _fmt_brl(value, suffix: str = "") -> str:
    """Formata valor em Reais no padrao brasileiro."""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return str(value) if value else "—"
    if abs(v) >= 1_000_000_000:
        formatted = f"R$ {v / 1_000_000_000:,.2f} Bi".replace(",", "X").replace(".", ",").replace("X", ".")
    elif abs(v) >= 1_000_000:
        formatted = f"R$ {v / 1_000_000:,.1f} MM".replace(",", "X").replace(".", ",").replace("X", ".")
    elif abs(v) >= 1_000:
        formatted = f"R$ {v / 1_000:,.0f} mil".replace(",", "X").replace(".", ",").replace("X", ".")
    else:
        formatted = f"R$ {v:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return formatted + (f" {suffix}" if suffix else "")


def _fmt_pct(value, decimals: int = 1) -> str:
    """Formata percentual no padrao brasileiro."""
    try:
        v = float(value)
        if 0 < abs(v) <= 1:
            v = v * 100
        return f"{v:,.{decimals}f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return str(value) if value else "—"


def _fmt_mult(value) -> str:
    """Formata multiplo (ex: 2,34x)."""
    try:
        v = float(value)
        return f"{v:,.2f}x".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return str(value) if value else "—"


def _trunc(text: str, limit: int = 120) -> str:
    text = str(text) if text else ""
    return text[:limit] + "..." if len(text) > limit else text


def _safe_get(d: dict, *keys, default="—"):
    """Navegacao segura em dicts aninhados."""
    current = d
    for k in keys:
        if isinstance(current, dict):
            current = current.get(k, default)
        else:
            return default
    return current if current not in (None, "", 0, 0.0) else default


def _rating_color(nota: str) -> RGBColor:
    nota_upper = str(nota).upper().strip()
    if nota_upper in ("A", "AA", "AAA", "A+", "A-"):
        return GREEN
    if nota_upper in ("B", "BB", "BBB", "B+", "B-"):
        return GREEN_LIGHT
    if nota_upper in ("C", "CC", "CCC", "C+", "C-"):
        return GOLD
    return RED


# ---------------------------------------------------------------------------
# Primitivos de desenho
# ---------------------------------------------------------------------------
def _add_rect(slide, left, top, width, height, fill_color, border=False):
    """Adiciona retangulo preenchido."""
    shape = slide.shapes.add_shape(1, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if not border:
        shape.line.fill.background()
    return shape


def _add_rounded_rect(slide, left, top, width, height, fill_color):
    """Adiciona retangulo com cantos arredondados."""
    # AutoShape type 5 = Rounded Rectangle
    shape = slide.shapes.add_shape(5, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape


def _txt(slide, left, top, width, height, text, size=12, bold=False,
         color=WHITE, align=PP_ALIGN.LEFT, font_name="Calibri",
         word_wrap=True, v_anchor=None):
    """Adiciona textbox com configuracoes consistentes."""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = word_wrap
    if v_anchor:
        tf.paragraphs[0].space_before = Pt(0)
        tf.paragraphs[0].space_after = Pt(0)
    p = tf.paragraphs[0]
    p.text = str(text) if text else ""
    run = p.runs[0] if p.runs else p.add_run()
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font_name
    p.alignment = align
    return box


def _txt_multi(slide, left, top, width, height, lines: list,
               default_size=10, default_color=BODY_LIGHT, font_name="Calibri"):
    """Textbox com multiplos paragrafos (cada item: dict ou string)."""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        if isinstance(line, dict):
            p.text = str(line.get("text", ""))
            run = p.runs[0] if p.runs else p.add_run()
            run.font.size = Pt(line.get("size", default_size))
            run.font.bold = line.get("bold", False)
            run.font.color.rgb = line.get("color", default_color)
            run.font.name = font_name
            if "align" in line:
                p.alignment = line["align"]
        else:
            p.text = str(line)
            run = p.runs[0] if p.runs else p.add_run()
            run.font.size = Pt(default_size)
            run.font.color.rgb = default_color
            run.font.name = font_name
    return box


def _hline(slide, left, top, width, color=None):
    """Linha horizontal fina."""
    fill = color or RGBColor(0x33, 0x41, 0x55)
    shape = slide.shapes.add_shape(1, left, top, width, Emu(12700))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape


def _kpi_card(slide, x, y, w, label, value, subtitle="",
              label_color=LABEL_COLOR, value_color=WHITE,
              subtitle_color=SUBTITLE_COLOR, value_size=27,
              label_size=9, subtitle_size=9, bg_color=None):
    """Card de KPI padrao ZYN: label pequeno, valor grande, subtitulo."""
    card_h = Emu(1143000)
    if bg_color:
        _add_rounded_rect(slide, x, y, w, card_h, bg_color)
    _txt(slide, x + Emu(91440), y + Emu(91440), w - Emu(182880), Emu(182880),
         str(label).upper(), size=label_size, bold=True, color=label_color)
    _txt(slide, x + Emu(91440), y + Emu(320040), w - Emu(182880), Emu(457200),
         str(value), size=value_size, bold=True, color=value_color)
    if subtitle:
        _txt(slide, x + Emu(91440), y + Emu(822960), w - Emu(182880), Emu(228600),
             str(subtitle), size=subtitle_size, color=subtitle_color)


def _data_now() -> str:
    """Retorna data formatada: 'Marco 2026'."""
    now = datetime.now()
    mes = MESES_PT.get(now.month, now.strftime("%B"))
    return f"{mes} {now.year}"


def _data_now_upper() -> str:
    """Retorna data formatada maiuscula: 'MARCO / 2026'."""
    now = datetime.now()
    mes = MESES_PT.get(now.month, now.strftime("%B")).upper()
    return f"{mes} / {now.year}"


# ---------------------------------------------------------------------------
# Base slide — elementos comuns a todos os slides (exceto cover)
# ---------------------------------------------------------------------------
def _base_slide(prs, title: str = "", subtitle: str = "",
                slide_number: int = 0, total_slides: int = 0,
                show_header=True, dark_bg=True):
    """Cria slide com elementos ZYN padrao."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

    # Background
    bg_color = NAVY_DARK if dark_bg else PANEL_BG
    _add_rect(slide, 0, 0, SW, SH, bg_color)

    if show_header:
        # Header bar
        _add_rect(slide, 0, 0, SW, Emu(731520), NAVY)

        # Green accent line under header
        _add_rect(slide, 0, Emu(731520), SW, Emu(27432), GREEN)

        # ZYN CAPITAL logo (header)
        _txt(slide, Emu(365760), Emu(182880), Emu(2286000), Emu(365760),
             "ZYN CAPITAL", size=13, bold=True, color=WHITE)

        # Slide title (header right area)
        if title:
            _txt(slide, Emu(365760), Emu(365760), Emu(8229600), Emu(320040),
                 title.upper(), size=13, bold=True, color=WHITE,
                 align=PP_ALIGN.LEFT)

        # Data no canto direito do header
        _txt(slide, SW - Emu(3200400), Emu(228600), Emu(2834640), Emu(320040),
             _data_now(), size=10, color=SUBTITLE_COLOR,
             align=PP_ALIGN.RIGHT)

    # Footer
    footer_y = SH - Emu(320040)
    _add_rect(slide, 0, footer_y, SW, Emu(320040), NAVY)

    footer_text = "ZYN Capital  |  CONFIDENCIAL  |  " + _data_now()
    _txt(slide, Emu(365760), footer_y + Emu(54864), Emu(6858000), Emu(228600),
         footer_text, size=8, color=SUBTITLE_COLOR)

    if slide_number and total_slides:
        _txt(slide, SW - Emu(1371600), footer_y + Emu(54864),
             Emu(1005840), Emu(228600),
             f"{slide_number} / {total_slides}", size=8,
             color=SUBTITLE_COLOR, align=PP_ALIGN.RIGHT)

    return slide


# ---------------------------------------------------------------------------
# Slide 1 — Capa Premium
# ---------------------------------------------------------------------------
def _slide_cover(prs, op: dict, analise: dict):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Background navy escuro
    _add_rect(slide, 0, 0, SW, SH, DARK_OVERLAY)

    # Barra lateral verde accent (esquerda)
    _add_rect(slide, 0, 0, Emu(54864), SH, GREEN)

    # Watermark
    _txt(slide, Emu(2743200), Emu(2514600), Emu(7315200), Emu(1828800),
         "CONFIDENCIAL", size=72, bold=True, color=WATERMARK_COLOR,
         align=PP_ALIGN.CENTER)

    # ZYN CAPITAL (top left)
    _txt(slide, Emu(457200), Emu(365760), Emu(3657600), Emu(457200),
         "ZYN  CAPITAL", size=28, bold=True, color=WHITE)

    # Badge CONFIDENCIAL
    _add_rounded_rect(slide, SW - Emu(2286000), Emu(365760),
                      Emu(1828800), Emu(320040), NAVY)
    _txt(slide, SW - Emu(2286000), Emu(392430), Emu(1828800), Emu(274320),
         "CONFIDENCIAL", size=9, bold=True, color=GOLD,
         align=PP_ALIGN.CENTER)

    # Tipo de operacao (badge)
    tipo = op.get("tipo_operacao", "Credito Estruturado")
    _txt(slide, Emu(457200), Emu(1371600), Emu(5486400), Emu(320040),
         tipo.upper(), size=10, bold=True, color=BLUE_ACCENT)

    # Nome do tomador (titulo grande)
    tomador = op.get("tomador", "Operacao")
    _txt(slide, Emu(457200), Emu(1737360), Emu(9601200), Emu(914400),
         tomador, size=52, bold=True, color=WHITE)

    # Subtitulo descritivo
    setor = op.get("setor", "")
    localidade = op.get("localidade", "")
    sub_parts = [s for s in [setor, localidade] if s]
    if sub_parts:
        _txt(slide, Emu(457200), Emu(2651760), Emu(9601200), Emu(365760),
             "  |  ".join(sub_parts), size=13, color=SUBTITLE_COLOR)

    # Linha divisoria
    _hline(slide, Emu(457200), Emu(3200400), Emu(11277600))

    # KPI cards na capa (4 cards)
    volume = _fmt_brl(op.get("volume", 0))
    taxa = op.get("taxa", "—")
    prazo = op.get("prazo_meses", "—")
    prazo_str = f"{prazo} Meses" if prazo and prazo != "—" else "—"

    # Determinar 4o KPI baseado no tipo de operacao
    tipo_lower = str(tipo).lower()
    if "fidc" in tipo_lower:
        kpi4_label = "Subordinacao"
        kpi4_value = op.get("subordinacao", _safe_get(analise, "estrutura", "subordinacao"))
        kpi4_sub = "Cota Subordinada"
    elif "slb" in tipo_lower or "leaseback" in tipo_lower:
        kpi4_label = "LTV (VM)"
        ltv_val = _safe_get(analise, "kpis", "ltv", default=0)
        kpi4_value = _fmt_pct(ltv_val) if ltv_val and ltv_val != "—" else "—"
        kpi4_sub = "Loan-to-Value"
    elif "cri" in tipo_lower:
        kpi4_label = "LTV"
        ltv_val = _safe_get(analise, "kpis", "ltv", default=0)
        kpi4_value = _fmt_pct(ltv_val) if ltv_val and ltv_val != "—" else "—"
        kpi4_sub = "s/ Garantias Totais"
    else:
        kpi4_label = "DSCR"
        dscr_val = _safe_get(analise, "kpis", "dscr", default=0)
        kpi4_value = _fmt_mult(dscr_val) if dscr_val and dscr_val != "—" else "—"
        kpi4_sub = "Cobertura do Servico"

    cards = [
        ("Volume", volume, "Montante Principal"),
        ("Remuneracao", str(taxa), "Taxa Indicativa"),
        ("Prazo", prazo_str, "Prazo Total"),
        (kpi4_label, str(kpi4_value), kpi4_sub),
    ]

    card_w = Emu(2651760)
    card_gap = Emu(182880)
    cards_x_start = Emu(457200)
    cards_y = Emu(3383280)

    for idx, (label, value, sub) in enumerate(cards):
        x = cards_x_start + idx * (card_w + card_gap)
        _kpi_card(slide, x, cards_y, card_w, label, value, sub,
                  bg_color=NAVY, value_size=26)

    # Resumo executivo (se disponivel)
    resumo = op.get("resumo_executivo", "")
    if not resumo:
        # Montar resumo basico
        parts = []
        if tipo:
            parts.append(f"Operacao de {tipo}")
        if op.get("finalidade"):
            parts.append(f"Finalidade: {op['finalidade']}")
        if op.get("garantias_text"):
            parts.append(f"Garantias: {_trunc(op['garantias_text'], 80)}")
        resumo = ". ".join(parts) + "." if parts else ""

    if resumo:
        _add_rounded_rect(slide, Emu(457200), Emu(4754880),
                          Emu(11277600), Emu(914400), NAVY)
        _txt(slide, Emu(640080), Emu(4800600), Emu(1371600), Emu(228600),
             "RESUMO EXECUTIVO", size=10, bold=True, color=SUBTITLE_COLOR)
        _txt(slide, Emu(640080), Emu(5029200), Emu(10972800), Emu(548640),
             _trunc(resumo, 400), size=9.5, color=BODY_LIGHT)

    # Footer da capa
    _add_rect(slide, 0, SH - Emu(365760), SW, Emu(365760), NAVY)
    _txt(slide, Emu(457200), SH - Emu(310896),
         Emu(5486400), Emu(274320),
         f"ZYN CAPITAL  (C)  {datetime.now().year}  |  Confidencial",
         size=6.5, color=SUBTITLE_COLOR)

    cnpj = op.get("cnpj", "")
    if cnpj:
        _txt(slide, SW - Emu(4572000), SH - Emu(310896),
             Emu(4114800), Emu(274320),
             f"CNPJ: {cnpj}", size=6.5, color=SUBTITLE_COLOR,
             align=PP_ALIGN.RIGHT)


# ---------------------------------------------------------------------------
# Slide 2 — Resumo Executivo
# ---------------------------------------------------------------------------
def _slide_resumo(prs, op: dict, analise: dict, sn: int, total: int):
    slide = _base_slide(prs, title="Resumo Executivo",
                        slide_number=sn, total_slides=total)
    y_start = Emu(914400)

    # Painel esquerdo: Termos da operacao
    panel_w = Emu(5486400)
    panel_h = Emu(4572000)
    _add_rounded_rect(slide, Emu(365760), y_start, panel_w, panel_h, NAVY)

    _txt(slide, Emu(548640), y_start + Emu(137160), panel_w - Emu(365760), Emu(320040),
         "TERMOS DA OPERACAO", size=12, bold=True, color=WHITE)

    # Linha verde accent
    _add_rect(slide, Emu(548640), y_start + Emu(457200),
              Emu(2743200), Emu(27432), GREEN)

    # Dados da operacao (key-value pairs)
    items = [
        ("Tomador", op.get("tomador", "—")),
        ("CNPJ", op.get("cnpj", "—")),
        ("Instrumento", op.get("tipo_operacao", "—")),
        ("Volume", _fmt_brl(op.get("volume", 0))),
        ("Remuneracao", op.get("taxa", "—")),
        ("Prazo", f"{op.get('prazo_meses', '—')} meses" if op.get("prazo_meses") else "—"),
        ("Amortizacao", op.get("amortizacao", "—")),
        ("Carencia", op.get("carencia", "—")),
        ("Garantias", _trunc(op.get("garantias_text", "—") or "—", 60)),
        ("Finalidade", _trunc(op.get("finalidade", "—") or "—", 60)),
        ("Perfil Investidor", op.get("perfil_investidor", "Qualificado (CVM 30)")),
    ]

    row_y = y_start + Emu(594360)
    for label, valor in items:
        if valor and valor != "—":
            _txt(slide, Emu(548640), row_y,
                 Emu(1828800), Emu(274320),
                 label, size=8.5, bold=True, color=SUBTITLE_COLOR)
            _txt(slide, Emu(2468880), row_y,
                 Emu(3200400), Emu(274320),
                 str(valor), size=8.5, color=WHITE)
            row_y += Emu(320040)

    # Painel direito: KPIs financeiros
    right_x = Emu(6126480)
    right_w = Emu(5669280)
    _add_rounded_rect(slide, right_x, y_start, right_w, panel_h, NAVY)

    _txt(slide, right_x + Emu(182880), y_start + Emu(137160),
         right_w - Emu(365760), Emu(320040),
         "INDICADORES-CHAVE", size=12, bold=True, color=WHITE)

    _add_rect(slide, right_x + Emu(182880), y_start + Emu(457200),
              Emu(2743200), Emu(27432), GREEN)

    # KPI cards (2x3 grid)
    kpis = analise.get("kpis", {})
    rating = analise.get("rating_final", {})

    kpi_items = [
        ("Receita Liquida", _fmt_brl(kpis.get("receita_liquida", 0))),
        ("EBITDA", _fmt_brl(kpis.get("ebitda", 0))),
        ("Margem EBITDA", _fmt_pct(kpis.get("margem_ebitda", 0))),
        ("DSCR", _fmt_mult(kpis.get("dscr", 0))),
        ("LTV", _fmt_pct(kpis.get("ltv", 0))),
        ("Rating", rating.get("nota", "—")),
    ]

    mini_w = Emu(2468880)
    mini_h = Emu(640080)
    col = 0
    row = 0
    kpi_y_start = y_start + Emu(594360)

    for label, value in kpi_items:
        kx = right_x + Emu(182880) + col * (mini_w + Emu(182880))
        ky = kpi_y_start + row * (mini_h + Emu(91440))

        _add_rounded_rect(slide, kx, ky, mini_w, mini_h, NAVY_LIGHT)
        _txt(slide, kx + Emu(91440), ky + Emu(54864),
             mini_w - Emu(182880), Emu(182880),
             label.upper(), size=7.5, bold=True, color=LABEL_COLOR)

        # Cor do valor baseada no contexto
        v_color = WHITE
        if label == "Rating":
            v_color = _rating_color(str(value))
        _txt(slide, kx + Emu(91440), ky + Emu(274320),
             mini_w - Emu(182880), Emu(320040),
             str(value), size=18, bold=True, color=v_color)

        col += 1
        if col >= 2:
            col = 0
            row += 1

    # Parecer do rating (abaixo dos KPIs)
    parecer = rating.get("parecer", "")
    if parecer:
        parecer_y = kpi_y_start + 3 * (mini_h + Emu(91440)) + Emu(91440)
        _add_rounded_rect(slide, right_x + Emu(182880), parecer_y,
                          right_w - Emu(365760), Emu(548640), GREEN)
        _txt(slide, right_x + Emu(365760), parecer_y + Emu(54864),
             right_w - Emu(731520), Emu(182880),
             "PARECER", size=8, bold=True, color=WHITE)
        _txt(slide, right_x + Emu(365760), parecer_y + Emu(274320),
             right_w - Emu(731520), Emu(228600),
             parecer, size=10, bold=True, color=WHITE)


# ---------------------------------------------------------------------------
# Slide 3 — Perfil do Tomador / Cedente
# ---------------------------------------------------------------------------
def _slide_tomador(prs, op: dict, analise: dict, sn: int, total: int):
    slide = _base_slide(prs, title="Perfil do Tomador",
                        slide_number=sn, total_slides=total)
    y_start = Emu(914400)

    tomador_data = analise.get("tomador", {})

    # Titulo com nome do tomador
    _txt(slide, Emu(365760), y_start, Emu(11277600), Emu(457200),
         op.get("tomador", "—"), size=22, bold=True, color=WHITE)

    tipo_lower = str(op.get("tipo_operacao", "")).lower()
    if "fidc" in tipo_lower:
        label_tomador = "Cedente / Originador"
    elif "slb" in tipo_lower:
        label_tomador = "Arrendatario"
    else:
        label_tomador = "Tomador"

    # Descricao
    descricao = (tomador_data.get("descricao", "") or
                 tomador_data.get("historico", "") or
                 op.get("descricao_tomador", ""))
    if descricao:
        _txt(slide, Emu(365760), y_start + Emu(457200),
             Emu(11277600), Emu(457200),
             _trunc(descricao, 250), size=10.5, color=SUBTITLE_COLOR)

    # Linha divisoria
    _hline(slide, Emu(365760), y_start + Emu(914400), Emu(11277600))

    # KPI cards do tomador (4 em linha)
    kpis = analise.get("kpis", {})
    receita = _fmt_brl(kpis.get("receita_bruta", kpis.get("receita_liquida", 0)))
    lucro_bruto = _fmt_brl(kpis.get("lucro_bruto", kpis.get("lucro_liquido", 0)))
    pl = _fmt_brl(kpis.get("patrimonio_liquido", 0))
    ebitda = _fmt_brl(kpis.get("ebitda", 0))

    # Crescimentos (se disponiveis)
    cresc_receita = kpis.get("crescimento_receita", "")
    cresc_receita_str = f"+{_fmt_pct(cresc_receita)} vs anterior" if cresc_receita else ""
    margem_str = f"Margem {_fmt_pct(kpis.get('margem_ebitda', 0))}" if kpis.get("margem_ebitda") else ""

    tomador_cards = [
        ("Receita Bruta", receita, cresc_receita_str, GREEN_LIGHT if cresc_receita_str else SUBTITLE_COLOR),
        ("EBITDA", ebitda, margem_str, GREEN_LIGHT if margem_str else SUBTITLE_COLOR),
        ("Lucro Bruto/Liquido", lucro_bruto, "", SUBTITLE_COLOR),
        ("Patrimonio Liquido", pl, "", SUBTITLE_COLOR),
    ]

    card_w = Emu(2651760)
    card_gap = Emu(182880)
    cards_y = y_start + Emu(1005840)

    for idx, (label, value, sub, sub_color) in enumerate(tomador_cards):
        x = Emu(365760) + idx * (card_w + card_gap)
        _add_rounded_rect(slide, x, cards_y, card_w, Emu(914400), NAVY)
        _txt(slide, x + Emu(137160), cards_y + Emu(91440),
             card_w - Emu(274320), Emu(182880),
             label.upper(), size=9, bold=True, color=LABEL_COLOR)
        _txt(slide, x + Emu(137160), cards_y + Emu(320040),
             card_w - Emu(274320), Emu(365760),
             str(value), size=16, bold=True, color=TITLE_DARK if value == "—" else WHITE)
        if sub:
            _txt(slide, x + Emu(137160), cards_y + Emu(685800),
                 card_w - Emu(274320), Emu(182880),
                 sub, size=8.5, bold=True, color=sub_color)

    # Painel inferior: Destaques e dados adicionais
    bottom_y = cards_y + Emu(1097280)

    # Esquerda: Dados societarios
    _add_rounded_rect(slide, Emu(365760), bottom_y,
                      Emu(5486400), Emu(2743200), NAVY)
    _txt(slide, Emu(548640), bottom_y + Emu(137160),
         Emu(5120640), Emu(274320),
         label_tomador.upper(), size=9, bold=True, color=SUBTITLE_COLOR)
    _add_rect(slide, Emu(548640), bottom_y + Emu(411480),
              Emu(1828800), Emu(18288), GREEN)

    soc_items = [
        ("Grupo / Razao Social", tomador_data.get("grupo_economico", op.get("tomador", "—"))),
        ("CNPJ", op.get("cnpj", "—")),
        ("Atividade", tomador_data.get("atividade_principal", op.get("setor", "—"))),
        ("Localizacao", op.get("localidade", "—")),
        ("Regime", tomador_data.get("regime", "—")),
        ("Fundacao", tomador_data.get("fundacao", "—")),
    ]

    soc_y = bottom_y + Emu(502920)
    for label, valor in soc_items:
        if valor and valor != "—":
            _txt(slide, Emu(548640), soc_y, Emu(2057400), Emu(228600),
                 label, size=8, bold=True, color=SUBTITLE_COLOR)
            _txt(slide, Emu(2651760), soc_y, Emu(3017520), Emu(228600),
                 str(valor), size=8, color=WHITE)
            soc_y += Emu(274320)

    # Direita: Destaques
    right_x = Emu(6126480)
    right_w = Emu(5669280)
    _add_rounded_rect(slide, right_x, bottom_y, right_w, Emu(2743200), NAVY)
    _txt(slide, right_x + Emu(182880), bottom_y + Emu(137160),
         right_w - Emu(365760), Emu(274320),
         "DESTAQUES", size=9, bold=True, color=SUBTITLE_COLOR)
    _add_rect(slide, right_x + Emu(182880), bottom_y + Emu(411480),
              Emu(1828800), Emu(18288), GREEN)

    # Coletar destaques de varias fontes
    destaques = op.get("destaques", [])
    if not destaques:
        # Gerar destaques a partir da analise
        if tomador_data.get("historico"):
            destaques.append(_trunc(tomador_data["historico"], 80))
        if kpis.get("receita_liquida"):
            destaques.append(f"Receita de {_fmt_brl(kpis['receita_liquida'])}")
        prod = analise.get("producao", {})
        if prod.get("analise"):
            destaques.append(_trunc(prod["analise"], 80))
        if kpis.get("dscr"):
            destaques.append(f"DSCR de {_fmt_mult(kpis['dscr'])} — capacidade de pagamento adequada")

    dest_y = bottom_y + Emu(502920)
    for dest in destaques[:6]:
        _txt(slide, right_x + Emu(182880), dest_y,
             right_w - Emu(365760), Emu(320040),
             f"  {_trunc(str(dest), 90)}", size=9, color=BODY_LIGHT)
        dest_y += Emu(320040)


# ---------------------------------------------------------------------------
# Slide 4 — Indicadores Financeiros
# ---------------------------------------------------------------------------
def _slide_financeiros(prs, analise: dict, sn: int, total: int):
    slide = _base_slide(prs, title="Indicadores Financeiros",
                        slide_number=sn, total_slides=total)
    y_start = Emu(914400)

    kpis = analise.get("kpis", {})
    capital = analise.get("capital", {})
    indicadores = capital.get("indicadores", {})

    # Titulo da secao
    _txt(slide, Emu(365760), y_start, Emu(11277600), Emu(365760),
         "Desempenho & Solidez", size=22, bold=True, color=WHITE)
    _txt(slide, Emu(365760), y_start + Emu(365760), Emu(11277600), Emu(228600),
         "Indicadores financeiros consolidados do ultimo exercicio", size=10,
         color=SUBTITLE_COLOR)

    _hline(slide, Emu(365760), y_start + Emu(640080), Emu(11277600))

    # 6 KPI cards (3 x 2)
    big_kpis = [
        ("Receita Liquida", _fmt_brl(kpis.get("receita_liquida", 0)), "Ultimo Exercicio"),
        ("EBITDA", _fmt_brl(kpis.get("ebitda", 0)),
         f"Margem: {_fmt_pct(kpis.get('margem_ebitda', 0))}" if kpis.get("margem_ebitda") else ""),
        ("Lucro Liquido", _fmt_brl(kpis.get("lucro_liquido", 0)),
         f"Margem: {_fmt_pct(kpis.get('margem_liquida', 0))}" if kpis.get("margem_liquida") else ""),
        ("DSCR", _fmt_mult(kpis.get("dscr", 0)), "Cobertura do Servico da Divida"),
        ("Div. Liq./EBITDA", _fmt_mult(kpis.get("divida_liquida_ebitda", 0)), "Alavancagem"),
        ("Patrimonio Liquido", _fmt_brl(kpis.get("patrimonio_liquido", 0)), "Base de Capital"),
    ]

    card_w = Emu(3566160)
    card_h = Emu(914400)
    card_gap_x = Emu(182880)
    card_gap_y = Emu(137160)
    cards_y = y_start + Emu(731520)

    for idx, (label, value, sub) in enumerate(big_kpis):
        col = idx % 3
        row = idx // 3
        x = Emu(365760) + col * (card_w + card_gap_x)
        y = cards_y + row * (card_h + card_gap_y)

        _add_rounded_rect(slide, x, y, card_w, card_h, NAVY)
        _txt(slide, x + Emu(137160), y + Emu(91440),
             card_w - Emu(274320), Emu(182880),
             label.upper(), size=9, bold=True, color=LABEL_COLOR)
        _txt(slide, x + Emu(137160), y + Emu(320040),
             card_w - Emu(274320), Emu(365760),
             str(value), size=22, bold=True, color=WHITE)
        if sub:
            _txt(slide, x + Emu(137160), y + Emu(685800),
                 card_w - Emu(274320), Emu(182880),
                 sub, size=8.5, color=SUBTITLE_COLOR)

    # Painel inferior: Indicadores adicionais + Capacidade de Pagamento
    bottom_y = cards_y + 2 * (card_h + card_gap_y) + Emu(182880)

    # Esquerda: Indicadores de Capital
    _add_rounded_rect(slide, Emu(365760), bottom_y,
                      Emu(5486400), Emu(2194560), NAVY)
    _txt(slide, Emu(548640), bottom_y + Emu(137160),
         Emu(5120640), Emu(274320),
         "INDICADORES DE CAPITAL", size=10, bold=True, color=WHITE)
    _add_rect(slide, Emu(548640), bottom_y + Emu(411480),
              Emu(2286000), Emu(18288), GREEN)

    ind_items = [
        ("Divida / PL", _fmt_mult(indicadores.get("divida_pl", kpis.get("divida_pl", 0)))),
        ("Liquidez Corrente", _fmt_mult(indicadores.get("liquidez_corrente", kpis.get("liquidez_corrente", 0)))),
        ("ROE", _fmt_pct(indicadores.get("roe", kpis.get("roe", 0)))),
        ("LTV Efetivo", _fmt_pct(kpis.get("ltv", 0))),
        ("Cobertura de Juros", _fmt_mult(indicadores.get("cobertura_juros", kpis.get("cobertura_juros", 0)))),
    ]

    ind_y = bottom_y + Emu(502920)
    for label, valor in ind_items:
        if valor and valor not in ("—", "0,00x", "0,0%"):
            _txt(slide, Emu(548640), ind_y, Emu(2286000), Emu(228600),
                 label, size=9, bold=True, color=SUBTITLE_COLOR)
            _txt(slide, Emu(2926080), ind_y, Emu(2560320), Emu(228600),
                 str(valor), size=9, bold=True, color=WHITE)
            ind_y += Emu(274320)

    # Direita: Capacidade de Pagamento
    right_x = Emu(6126480)
    right_w = Emu(5669280)
    _add_rounded_rect(slide, right_x, bottom_y, right_w, Emu(2194560), NAVY)
    _txt(slide, right_x + Emu(182880), bottom_y + Emu(137160),
         right_w - Emu(365760), Emu(274320),
         "CAPACIDADE DE PAGAMENTO", size=10, bold=True, color=WHITE)
    _add_rect(slide, right_x + Emu(182880), bottom_y + Emu(411480),
              Emu(2286000), Emu(18288), GREEN)

    pagamento = analise.get("pagamento", {})
    pay_text = pagamento.get("analise", "")
    if pay_text:
        _txt(slide, right_x + Emu(182880), bottom_y + Emu(502920),
             right_w - Emu(365760), Emu(1554480),
             _trunc(pay_text, 600), size=9, color=BODY_LIGHT)
    else:
        # Mostrar KPIs de capacidade
        cap_items = [
            f"DSCR: {_fmt_mult(kpis.get('dscr', 0))}",
            f"Comprometimento da Receita: {_fmt_pct(kpis.get('comprometimento_receita', 0))}",
            f"EBITDA / Servico Divida: {_fmt_mult(kpis.get('ebitda_servico_divida', 0))}",
        ]
        cap_y = bottom_y + Emu(502920)
        for item in cap_items:
            _txt(slide, right_x + Emu(182880), cap_y,
                 right_w - Emu(365760), Emu(274320),
                 item, size=9.5, color=BODY_LIGHT)
            cap_y += Emu(320040)


# ---------------------------------------------------------------------------
# Slide 5 — Estrutura da Operacao
# ---------------------------------------------------------------------------
def _slide_estrutura(prs, op: dict, analise: dict, sn: int, total: int):
    slide = _base_slide(prs, title="Estrutura da Operacao",
                        slide_number=sn, total_slides=total)
    y_start = Emu(914400)

    tipo = op.get("tipo_operacao", "Credito Estruturado")
    taxa = op.get("taxa", "—")

    _txt(slide, Emu(365760), y_start, Emu(11277600), Emu(365760),
         f"Estrutura & Termos  |  {tipo}", size=22, bold=True, color=WHITE)

    _hline(slide, Emu(365760), y_start + Emu(457200), Emu(11277600))

    # 6 KPI cards da estrutura (em linha)
    struct_kpis = [
        ("Modalidade", tipo, ""),
        ("Volume", _fmt_brl(op.get("volume", 0)), ""),
        ("Taxa All-In", str(taxa), ""),
        ("Prazo Total", f"{op.get('prazo_meses', '—')} meses", ""),
        ("Carencia", op.get("carencia", "—"), ""),
        ("Amortizacao", op.get("amortizacao", "—"), ""),
    ]

    card_w = Emu(1828800)
    card_h = Emu(914400)
    cards_y = y_start + Emu(548640)
    card_gap = Emu(91440)

    for idx, (label, value, sub) in enumerate(struct_kpis):
        if not value or value == "—":
            continue
        x = Emu(365760) + idx * (card_w + card_gap)
        _add_rounded_rect(slide, x, cards_y, card_w, card_h, NAVY)
        _txt(slide, x + Emu(91440), cards_y + Emu(91440),
             card_w - Emu(182880), Emu(182880),
             label.upper(), size=7, bold=True, color=LABEL_COLOR)
        _txt(slide, x + Emu(91440), cards_y + Emu(320040),
             card_w - Emu(182880), Emu(457200),
             _trunc(str(value), 25), size=11, bold=True, color=WHITE)

    # Painel: Cronograma / Fluxo (se disponivel)
    cronograma = analise.get("cronograma", {})
    fluxo = op.get("fluxo_caixa", []) or cronograma.get("fluxo", [])

    mid_y = cards_y + card_h + Emu(274320)

    if fluxo:
        # Tabela de fluxo de caixa
        _add_rounded_rect(slide, Emu(365760), mid_y,
                          Emu(11277600), Emu(3200400), NAVY)
        _txt(slide, Emu(548640), mid_y + Emu(137160),
             Emu(10972800), Emu(274320),
             "FLUXO DE CAIXA — DESEMBOLSO ANUAL", size=10, bold=True, color=WHITE)
        _add_rect(slide, Emu(548640), mid_y + Emu(411480),
                  Emu(2743200), Emu(18288), GREEN)

        # Header da tabela
        cols = ["Ano / Fase", "Saldo Devedor", "Juros", "Amortizacao", "Desembolso", "DSCR"]
        col_widths = [Emu(1828800), Emu(1828800), Emu(1828800),
                      Emu(1828800), Emu(1828800), Emu(1371600)]
        header_y = mid_y + Emu(502920)
        _add_rect(slide, Emu(548640), header_y, Emu(10515600), Emu(274320), GREEN)

        col_x = Emu(548640)
        for ci, col_name in enumerate(cols):
            _txt(slide, col_x + Emu(45720), header_y + Emu(36576),
                 col_widths[ci], Emu(228600),
                 col_name, size=7, bold=True, color=WHITE)
            col_x += col_widths[ci]

        # Linhas de dados
        row_y = header_y + Emu(274320)
        for ri, row_data in enumerate(fluxo[:8]):
            bg = ROW_DARK_ODD if ri % 2 == 0 else ROW_DARK_EVEN
            _add_rect(slide, Emu(548640), row_y, Emu(10515600), Emu(228600), bg)

            vals = [
                row_data.get("ano", "—"),
                row_data.get("saldo", "—"),
                row_data.get("juros", "—"),
                row_data.get("amortizacao", "—"),
                row_data.get("desembolso", "—"),
                row_data.get("dscr", "—"),
            ]

            col_x = Emu(548640)
            for ci, val in enumerate(vals):
                v_color = GREEN_LIGHT if ci == 5 and str(val) not in ("—", "") else BODY_LIGHT
                is_bold = ci == 0
                _txt(slide, col_x + Emu(45720), row_y + Emu(27432),
                     col_widths[ci], Emu(182880),
                     str(val), size=7, bold=is_bold, color=v_color)
                col_x += col_widths[ci]
            row_y += Emu(228600)
    else:
        # Sem fluxo: mostrar uso dos recursos e dados complementares
        _add_rounded_rect(slide, Emu(365760), mid_y,
                          Emu(5486400), Emu(3200400), NAVY)
        _txt(slide, Emu(548640), mid_y + Emu(137160),
             Emu(5120640), Emu(274320),
             "USO DOS RECURSOS", size=10, bold=True, color=WHITE)
        _add_rect(slide, Emu(548640), mid_y + Emu(411480),
                  Emu(2286000), Emu(18288), GREEN)

        finalidade = op.get("finalidade", "—")
        uso_recursos = op.get("uso_recursos", [])
        if uso_recursos:
            uso_y = mid_y + Emu(502920)
            for item in uso_recursos[:8]:
                if isinstance(item, dict):
                    _txt(slide, Emu(548640), uso_y, Emu(3200400), Emu(228600),
                         item.get("descricao", "—"), size=8.5, color=BODY_LIGHT)
                    _txt(slide, Emu(3840480), uso_y, Emu(1554480), Emu(228600),
                         _fmt_brl(item.get("valor", 0)), size=8.5, bold=True,
                         color=WHITE, align=PP_ALIGN.RIGHT)
                else:
                    _txt(slide, Emu(548640), uso_y, Emu(4937760), Emu(228600),
                         f"  {str(item)}", size=8.5, color=BODY_LIGHT)
                uso_y += Emu(274320)
        elif finalidade and finalidade != "—":
            _txt(slide, Emu(548640), mid_y + Emu(502920),
                 Emu(5120640), Emu(2286000),
                 finalidade, size=9.5, color=BODY_LIGHT)

        # Direita: Dados complementares da operacao
        right_x = Emu(6126480)
        _add_rounded_rect(slide, right_x, mid_y,
                          Emu(5669280), Emu(3200400), NAVY)
        _txt(slide, right_x + Emu(182880), mid_y + Emu(137160),
             Emu(5303520), Emu(274320),
             "DADOS COMPLEMENTARES", size=10, bold=True, color=WHITE)
        _add_rect(slide, right_x + Emu(182880), mid_y + Emu(411480),
                  Emu(2286000), Emu(18288), GREEN)

        comp_items = [
            ("Securitizadora", op.get("securitizadora", "")),
            ("Administrador", op.get("administrador", "")),
            ("Custodiante", op.get("custodiante", "")),
            ("Agente Fiduciario", op.get("agente_fiduciario", "")),
            ("Fundo de Reserva", op.get("fundo_reserva", "")),
            ("Covenants", _trunc(op.get("covenants_text", "") or "", 80)),
            ("Regime Liberacao", op.get("regime_liberacao", "")),
        ]

        comp_y = mid_y + Emu(502920)
        for label, valor in comp_items:
            if valor and valor != "—":
                _txt(slide, right_x + Emu(182880), comp_y,
                     Emu(2286000), Emu(228600),
                     label, size=8, bold=True, color=SUBTITLE_COLOR)
                _txt(slide, right_x + Emu(2560320), comp_y,
                     Emu(2651760), Emu(228600),
                     str(valor), size=8, color=WHITE)
                comp_y += Emu(274320)


# ---------------------------------------------------------------------------
# Slide 6 — Garantias & Colateral
# ---------------------------------------------------------------------------
def _slide_garantias(prs, op: dict, analise: dict, sn: int, total: int):
    slide = _base_slide(prs, title="Garantias & Colateral",
                        slide_number=sn, total_slides=total)
    y_start = Emu(914400)

    kpis = analise.get("kpis", {})
    patrimonio = analise.get("patrimonio", {})
    garantias_list = op.get("garantias", []) or patrimonio.get("garantias", [])

    # Titulo
    ltv = kpis.get("ltv", 0)
    cobertura = kpis.get("cobertura_total", 0)
    title_parts = ["Garantias & Solidez Patrimonial"]
    if cobertura:
        title_parts.append(f"Cobertura {_fmt_mult(cobertura)}")
    if ltv:
        title_parts.append(f"LTV {_fmt_pct(ltv)}")

    _txt(slide, Emu(365760), y_start, Emu(11277600), Emu(365760),
         "  |  ".join(title_parts), size=22, bold=True, color=WHITE)

    _hline(slide, Emu(365760), y_start + Emu(457200), Emu(11277600))

    # KPI cards de garantia (ate 6)
    vol = op.get("volume", 0)
    vm_total = patrimonio.get("valor_mercado", kpis.get("valor_mercado_garantias", 0))
    vlf_total = patrimonio.get("valor_liquidacao", kpis.get("valor_liquidacao_garantias", 0))

    gar_kpis = []
    if vm_total:
        gar_kpis.append(("VM Total", _fmt_brl(vm_total), "Valor de Mercado"))
    if vlf_total:
        gar_kpis.append(("VLF Total", _fmt_brl(vlf_total), "Liquidacao Forcada"))
    if ltv:
        gar_kpis.append(("LTV", _fmt_pct(ltv), "Loan-to-Value"))
    if cobertura:
        gar_kpis.append(("Cobertura", _fmt_mult(cobertura), "VM / Operacao"))

    # Preencher ate 4 com dados extras
    ltv_vlf = patrimonio.get("ltv_vlf", kpis.get("ltv_vlf", 0))
    cobertura_vlf = patrimonio.get("cobertura_vlf", kpis.get("cobertura_vlf", 0))
    if ltv_vlf and len(gar_kpis) < 5:
        gar_kpis.append(("LTV (VLF)", _fmt_pct(ltv_vlf), "s/ Liquidacao Forcada"))
    if cobertura_vlf and len(gar_kpis) < 6:
        gar_kpis.append(("Cob. VLF", _fmt_mult(cobertura_vlf), "VLF / Operacao"))

    if gar_kpis:
        card_w = Emu(11277600 // min(len(gar_kpis), 6) - 91440)
        cards_y = y_start + Emu(548640)
        for idx, (label, value, sub) in enumerate(gar_kpis[:6]):
            x = Emu(365760) + idx * (card_w + Emu(91440))
            _add_rounded_rect(slide, x, cards_y, card_w, Emu(731520), NAVY)
            _txt(slide, x + Emu(91440), cards_y + Emu(54864),
                 card_w - Emu(182880), Emu(182880),
                 label.upper(), size=7, bold=True, color=LABEL_COLOR)
            _txt(slide, x + Emu(91440), cards_y + Emu(274320),
                 card_w - Emu(182880), Emu(320040),
                 str(value), size=14, bold=True, color=WHITE)
            _txt(slide, x + Emu(91440), cards_y + Emu(548640),
                 card_w - Emu(182880), Emu(137160),
                 sub, size=7, color=SUBTITLE_COLOR)

    # Tabela de garantias (se disponivel)
    table_y = y_start + Emu(1371600)

    if garantias_list and isinstance(garantias_list, list) and len(garantias_list) > 0:
        _add_rounded_rect(slide, Emu(365760), table_y,
                          Emu(11277600), Emu(3657600), NAVY)
        _txt(slide, Emu(548640), table_y + Emu(137160),
             Emu(10972800), Emu(274320),
             "DETALHAMENTO DAS GARANTIAS", size=10, bold=True, color=WHITE)
        _add_rect(slide, Emu(548640), table_y + Emu(411480),
                  Emu(2743200), Emu(18288), GREEN)

        # Determinar colunas baseado nos dados
        has_area = any(g.get("area") for g in garantias_list if isinstance(g, dict))

        if has_area:
            cols = ["Garantia", "Area", "VM (R$)", "VLF (R$)", "Status"]
            col_ws = [Emu(3200400), Emu(1371600), Emu(2286000), Emu(2286000), Emu(1371600)]
        else:
            cols = ["Garantia", "Detalhe", "Valor (R$)", "Cobertura"]
            col_ws = [Emu(3200400), Emu(3200400), Emu(2560320), Emu(1554480)]

        # Header
        header_y = table_y + Emu(502920)
        _add_rect(slide, Emu(548640), header_y, Emu(10515600), Emu(274320), GREEN)
        col_x = Emu(548640)
        for ci, col_name in enumerate(cols):
            _txt(slide, col_x + Emu(45720), header_y + Emu(36576),
                 col_ws[ci], Emu(228600),
                 col_name, size=7, bold=True, color=WHITE)
            col_x += col_ws[ci]

        # Dados
        row_y = header_y + Emu(274320)
        for ri, gar in enumerate(garantias_list[:10]):
            bg = ROW_DARK_ODD if ri % 2 == 0 else ROW_DARK_EVEN
            _add_rect(slide, Emu(548640), row_y, Emu(10515600), Emu(228600), bg)

            if isinstance(gar, dict):
                if has_area:
                    vals = [
                        gar.get("descricao", gar.get("nome", "—")),
                        str(gar.get("area", "—")),
                        _fmt_brl(gar.get("valor_mercado", gar.get("vm", 0))),
                        _fmt_brl(gar.get("valor_liquidacao", gar.get("vlf", 0))),
                        gar.get("status", "—"),
                    ]
                else:
                    vals = [
                        gar.get("descricao", gar.get("tipo", "—")),
                        gar.get("detalhe", "—"),
                        _fmt_brl(gar.get("valor", 0)),
                        _fmt_pct(gar.get("cobertura", 0)),
                    ]
            else:
                vals = [str(gar)] + ["—"] * (len(cols) - 1)

            col_x = Emu(548640)
            for ci, val in enumerate(vals):
                is_bold = ci == 0
                v_color = GREEN_LIGHT if "ok" in str(val).lower() or "principal" in str(val).lower() else BODY_LIGHT
                if ci == 0:
                    v_color = WHITE
                _txt(slide, col_x + Emu(45720), row_y + Emu(27432),
                     col_ws[ci], Emu(182880),
                     str(val), size=7.5, bold=is_bold, color=v_color)
                col_x += col_ws[ci]
            row_y += Emu(228600)

        # Linha de total (se tiver VM e VLF)
        if vm_total or vlf_total:
            _add_rect(slide, Emu(548640), row_y, Emu(10515600), Emu(274320), GREEN)
            _txt(slide, Emu(594360), row_y + Emu(36576),
                 Emu(3200400), Emu(228600),
                 "TOTAL", size=7.5, bold=True, color=WHITE)
            if has_area and vm_total:
                _txt(slide, Emu(548640) + col_ws[0] + col_ws[1] + Emu(45720),
                     row_y + Emu(36576), Emu(2286000), Emu(228600),
                     _fmt_brl(vm_total), size=7.5, bold=True, color=WHITE)
            if has_area and vlf_total:
                _txt(slide, Emu(548640) + col_ws[0] + col_ws[1] + col_ws[2] + Emu(45720),
                     row_y + Emu(36576), Emu(2286000), Emu(228600),
                     _fmt_brl(vlf_total), size=7.5, bold=True, color=WHITE)
    else:
        # Sem lista detalhada: mostrar texto das garantias
        _add_rounded_rect(slide, Emu(365760), table_y,
                          Emu(11277600), Emu(3657600), NAVY)
        _txt(slide, Emu(548640), table_y + Emu(137160),
             Emu(10972800), Emu(274320),
             "GARANTIAS DA OPERACAO", size=10, bold=True, color=WHITE)
        _add_rect(slide, Emu(548640), table_y + Emu(411480),
                  Emu(2743200), Emu(18288), GREEN)

        gar_text = op.get("garantias_text", "")
        if gar_text:
            _txt(slide, Emu(548640), table_y + Emu(502920),
                 Emu(10972800), Emu(2743200),
                 _trunc(gar_text, 800), size=10, color=BODY_LIGHT)


# ---------------------------------------------------------------------------
# Slide 7 — Riscos & Mitigantes
# ---------------------------------------------------------------------------
def _slide_riscos(prs, analise: dict, sn: int, total: int):
    slide = _base_slide(prs, title="Riscos & Mitigantes",
                        slide_number=sn, total_slides=total)
    y_start = Emu(914400)

    riscos = analise.get("riscos", {})
    matriz = riscos.get("matriz_riscos", [])

    _txt(slide, Emu(365760), y_start, Emu(11277600), Emu(365760),
         "Analise de Riscos & Mitigantes", size=22, bold=True, color=WHITE)

    _hline(slide, Emu(365760), y_start + Emu(457200), Emu(11277600))

    if matriz:
        # Tabela de riscos
        _add_rounded_rect(slide, Emu(365760), y_start + Emu(548640),
                          Emu(11277600), Emu(4572000), NAVY)
        _txt(slide, Emu(548640), y_start + Emu(685800),
             Emu(10972800), Emu(274320),
             "MATRIZ DE RISCOS", size=10, bold=True, color=WHITE)
        _add_rect(slide, Emu(548640), y_start + Emu(960120),
                  Emu(2743200), Emu(18288), GREEN)

        cols = ["Risco", "Mitigante ZYN"]
        col_ws = [Emu(4572000), Emu(5943600)]

        header_y = y_start + Emu(1051560)
        _add_rect(slide, Emu(548640), header_y, Emu(10515600), Emu(320040), GREEN)
        col_x = Emu(548640)
        for ci, col_name in enumerate(cols):
            _txt(slide, col_x + Emu(91440), header_y + Emu(54864),
                 col_ws[ci], Emu(228600),
                 col_name, size=8, bold=True, color=WHITE)
            col_x += col_ws[ci]

        row_y = header_y + Emu(320040)
        for ri, risco in enumerate(matriz[:8]):
            bg = ROW_DARK_ODD if ri % 2 == 0 else ROW_DARK_EVEN
            row_h = Emu(365760)
            _add_rect(slide, Emu(548640), row_y, Emu(10515600), row_h, bg)

            nome = risco.get("risco", "—")
            mitigante = risco.get("mitigante", "—")

            _txt(slide, Emu(640080), row_y + Emu(54864),
                 col_ws[0] - Emu(91440), Emu(274320),
                 _trunc(nome, 60), size=8, bold=True, color=WHITE)
            _txt(slide, Emu(548640) + col_ws[0] + Emu(91440), row_y + Emu(54864),
                 col_ws[1] - Emu(91440), Emu(274320),
                 _trunc(mitigante, 80), size=8, color=BODY_LIGHT)

            row_y += row_h
    else:
        # Sem matriz: mostrar analise de risco geral
        _add_rounded_rect(slide, Emu(365760), y_start + Emu(548640),
                          Emu(11277600), Emu(4572000), NAVY)
        _txt(slide, Emu(548640), y_start + Emu(685800),
             Emu(10972800), Emu(274320),
             "ANALISE DE RISCO", size=10, bold=True, color=WHITE)
        _add_rect(slide, Emu(548640), y_start + Emu(960120),
                  Emu(2743200), Emu(18288), GREEN)

        risco_text = riscos.get("analise", riscos.get("parecer", ""))
        if risco_text:
            _txt(slide, Emu(548640), y_start + Emu(1051560),
                 Emu(10972800), Emu(3657600),
                 _trunc(risco_text, 1000), size=10, color=BODY_LIGHT)

        # Flags como pontos de atencao
        all_flags = []
        for secao in ["tomador", "patrimonio", "producao", "capital", "operacao",
                       "pagamento", "onus", "riscos", "covenants", "cronograma"]:
            dados = analise.get(secao, {})
            flags = dados.get("flags", [])
            for f in flags:
                if isinstance(f, dict):
                    all_flags.append(f.get("descricao", str(f)))
                else:
                    all_flags.append(str(f))

        if all_flags:
            flag_y = y_start + Emu(1371600 if risco_text else 1051560)
            _txt(slide, Emu(548640), flag_y, Emu(10972800), Emu(274320),
                 "PONTOS DE ATENCAO (FLAGS)", size=9, bold=True, color=GOLD)
            flag_y += Emu(320040)
            for flag in all_flags[:8]:
                _txt(slide, Emu(548640), flag_y, Emu(10972800), Emu(274320),
                     f"  {_trunc(flag, 100)}", size=9, color=BODY_LIGHT)
                flag_y += Emu(274320)


# ---------------------------------------------------------------------------
# Slide 8 — Parecer Final & Rating
# ---------------------------------------------------------------------------
def _slide_parecer(prs, op: dict, analise: dict, sn: int, total: int):
    slide = _base_slide(prs, title="Parecer & Rating",
                        slide_number=sn, total_slides=total)
    y_start = Emu(914400)

    rating = analise.get("rating_final", {})
    nota = rating.get("nota", "—")
    parecer = rating.get("parecer", "—")
    justificativa = rating.get("justificativa", "")
    kpis = analise.get("kpis", {})

    # Badge de rating grande
    _add_rounded_rect(slide, Emu(365760), y_start,
                      Emu(11277600), Emu(1371600), NAVY)

    # Nota grande
    _txt(slide, Emu(548640), y_start + Emu(137160),
         Emu(1371600), Emu(182880),
         "RATING FINAL", size=9, bold=True, color=LABEL_COLOR)
    _txt(slide, Emu(548640), y_start + Emu(365760),
         Emu(1371600), Emu(914400),
         str(nota), size=48, bold=True, color=_rating_color(nota))

    # Parecer ao lado do rating
    _txt(slide, Emu(2194560), y_start + Emu(137160),
         Emu(9144000), Emu(182880),
         "PARECER", size=9, bold=True, color=LABEL_COLOR)
    _txt(slide, Emu(2194560), y_start + Emu(365760),
         Emu(9144000), Emu(457200),
         str(parecer), size=24, bold=True, color=WHITE)

    if justificativa:
        _txt(slide, Emu(2194560), y_start + Emu(822960),
             Emu(9144000), Emu(457200),
             _trunc(justificativa, 200), size=9, color=BODY_LIGHT)

    # KPI cards resumo (4 em linha)
    n_flags = sum(len(analise.get(s, {}).get("flags", []))
                  for s in ["tomador", "patrimonio", "producao", "capital",
                            "operacao", "pagamento", "onus", "riscos",
                            "covenants", "cronograma"])

    summary_cards = [
        ("Rating Final", str(nota), parecer, _rating_color(nota)),
        ("DSCR", _fmt_mult(kpis.get("dscr", 0)), "Cobertura do Servico", WHITE),
        ("LTV", _fmt_pct(kpis.get("ltv", 0)), "Loan to Value", WHITE),
        ("Flags", str(n_flags), "Pontos de Atencao", GOLD if n_flags > 3 else GREEN_LIGHT),
    ]

    card_w = Emu(2651760)
    card_gap = Emu(182880)
    cards_y = y_start + Emu(1554480)

    for idx, (label, value, sub, v_color) in enumerate(summary_cards):
        x = Emu(365760) + idx * (card_w + card_gap)
        _add_rounded_rect(slide, x, cards_y, card_w, Emu(914400), NAVY)
        _txt(slide, x + Emu(137160), cards_y + Emu(91440),
             card_w - Emu(274320), Emu(182880),
             label.upper(), size=9, bold=True, color=LABEL_COLOR)
        _txt(slide, x + Emu(137160), cards_y + Emu(320040),
             card_w - Emu(274320), Emu(365760),
             value, size=22, bold=True, color=v_color)
        _txt(slide, x + Emu(137160), cards_y + Emu(685800),
             card_w - Emu(274320), Emu(182880),
             sub, size=8.5, color=SUBTITLE_COLOR)

    # Ratings por secao (tabela)
    table_y = cards_y + Emu(1097280)
    _add_rounded_rect(slide, Emu(365760), table_y,
                      Emu(11277600), Emu(2743200), NAVY)
    _txt(slide, Emu(548640), table_y + Emu(137160),
         Emu(10972800), Emu(274320),
         "RATINGS POR SECAO", size=10, bold=True, color=WHITE)
    _add_rect(slide, Emu(548640), table_y + Emu(411480),
              Emu(2743200), Emu(18288), GREEN)

    secoes = [
        ("Tomador", "tomador"), ("Patrimonio", "patrimonio"),
        ("Producao", "producao"), ("Capital", "capital"),
        ("Operacao", "operacao"), ("Pagamento", "pagamento"),
        ("Onus", "onus"), ("Riscos", "riscos"),
        ("Covenants", "covenants"), ("Cronograma", "cronograma"),
    ]

    # Layout 2 colunas para ratings
    header_y = table_y + Emu(502920)
    cols_per_row = 2
    col_w = Emu(5120640)
    items_per_col = 5

    for col_idx in range(cols_per_row):
        col_start_x = Emu(548640) + col_idx * col_w
        start_idx = col_idx * items_per_col
        end_idx = min(start_idx + items_per_col, len(secoes))

        # Mini header
        _add_rect(slide, col_start_x, header_y,
                  col_w - Emu(182880), Emu(228600), GREEN)
        _txt(slide, col_start_x + Emu(91440), header_y + Emu(27432),
             Emu(2286000), Emu(182880),
             "Secao", size=7, bold=True, color=WHITE)
        _txt(slide, col_start_x + Emu(2743200), header_y + Emu(27432),
             Emu(1371600), Emu(182880),
             "Rating", size=7, bold=True, color=WHITE)
        _txt(slide, col_start_x + Emu(4114800), header_y + Emu(27432),
             Emu(822960), Emu(182880),
             "Flags", size=7, bold=True, color=WHITE)

        row_y = header_y + Emu(228600)
        for si in range(start_idx, end_idx):
            secao_nome, secao_key = secoes[si]
            dados = analise.get(secao_key, {})
            rating_s = dados.get("rating_secao", "N/A")
            n_sec_flags = len(dados.get("flags", []))

            bg = ROW_DARK_ODD if (si - start_idx) % 2 == 0 else ROW_DARK_EVEN
            _add_rect(slide, col_start_x, row_y,
                      col_w - Emu(182880), Emu(228600), bg)
            _txt(slide, col_start_x + Emu(91440), row_y + Emu(27432),
                 Emu(2651760), Emu(182880),
                 secao_nome, size=8, bold=True, color=WHITE)

            # Cor do rating
            rating_colors = {"Forte": GREEN, "Adequado": GREEN_LIGHT,
                             "Atencao": GOLD, "Critico": RED}
            cor_r = rating_colors.get(rating_s, SUBTITLE_COLOR)
            _txt(slide, col_start_x + Emu(2743200), row_y + Emu(27432),
                 Emu(1371600), Emu(182880),
                 str(rating_s), size=8, bold=True, color=cor_r)
            _txt(slide, col_start_x + Emu(4114800), row_y + Emu(27432),
                 Emu(822960), Emu(182880),
                 f"{n_sec_flags}" if n_sec_flags > 0 else "OK",
                 size=8, color=GOLD if n_sec_flags > 0 else GREEN_LIGHT)
            row_y += Emu(228600)


# ---------------------------------------------------------------------------
# Slide 9 — Proximos Passos & Contato
# ---------------------------------------------------------------------------
def _slide_contato(prs, op: dict, sn: int, total: int):
    slide = _base_slide(prs, title="", show_header=False,
                        slide_number=sn, total_slides=total)

    # Background escuro com accent
    _add_rect(slide, 0, 0, SW, SH, DARK_OVERLAY)
    _add_rect(slide, 0, 0, Emu(54864), SH, GREEN)

    # ZYN CAPITAL (centro superior)
    _txt(slide, 0, Emu(548640), SW, Emu(548640),
         "ZYN  CAPITAL", size=28, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER)

    # Proximos Passos
    _txt(slide, 0, Emu(1371600), SW, Emu(365760),
         "PROXIMOS PASSOS", size=10, bold=True, color=SUBTITLE_COLOR,
         align=PP_ALIGN.CENTER)

    steps = [
        ("01", "Envio do Memorando Executivo completo"),
        ("02", "Call de due diligence com equipe ZYN Capital"),
        ("03", "Assinatura de NDA e acesso ao data room"),
        ("04", "Formalizacao da proposta de investimento"),
    ]

    step_w = Emu(2560320)
    step_gap = Emu(137160)
    step_total_w = len(steps) * step_w + (len(steps) - 1) * step_gap
    step_start_x = (SW - step_total_w) // 2
    step_y = Emu(1828800)

    for idx, (num, desc) in enumerate(steps):
        x = step_start_x + idx * (step_w + step_gap)
        _add_rounded_rect(slide, x, step_y, step_w, Emu(822960), NAVY)

        # Numero grande
        _txt(slide, x + Emu(91440), step_y + Emu(91440),
             Emu(548640), Emu(457200),
             num, size=24, bold=True, color=GREEN)
        # Descricao
        _txt(slide, x + Emu(91440), step_y + Emu(502920),
             step_w - Emu(182880), Emu(274320),
             desc, size=8.5, color=WHITE)

    # Linha divisoria
    _hline(slide, Emu(2743200), Emu(3017520), Emu(6705600),
           color=RGBColor(0x33, 0x41, 0x55))

    # Contato
    _txt(slide, 0, Emu(3200400), SW, Emu(365760),
         "PARA MAIORES INFORMACOES", size=10, bold=True,
         color=SUBTITLE_COLOR, align=PP_ALIGN.CENTER)

    # Nome do assessor
    assessor = op.get("assessor", "Danilo Salasar")
    cargo = op.get("cargo_assessor", "Socio | Credito Estruturado & M&A")
    email = op.get("email_assessor", "danilo@zyncapital.com.br")
    telefone = op.get("telefone_assessor", "65 9 9987-8781")

    _txt(slide, 0, Emu(3566160), SW, Emu(457200),
         assessor, size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    _txt(slide, 0, Emu(3931920), SW, Emu(274320),
         cargo, size=10, color=SUBTITLE_COLOR, align=PP_ALIGN.CENTER)

    # Email e telefone
    contact_y = Emu(4389120)
    _add_rounded_rect(slide, (SW - Emu(4572000)) // 2, contact_y,
                      Emu(4572000), Emu(548640), NAVY)
    _txt(slide, (SW - Emu(4572000)) // 2, contact_y + Emu(91440),
         Emu(4572000), Emu(228600),
         f"{email}  |  {telefone}", size=10, color=WHITE,
         align=PP_ALIGN.CENTER)
    _txt(slide, (SW - Emu(4572000)) // 2, contact_y + Emu(320040),
         Emu(4572000), Emu(182880),
         "Sao Paulo — SP  |  zyncapital.com.br", size=9,
         color=SUBTITLE_COLOR, align=PP_ALIGN.CENTER)

    # Disclaimer
    _txt(slide, Emu(548640), Emu(5303520), Emu(11097600), Emu(548640),
         "AVISO LEGAL: Este material e de uso exclusivo do destinatario e nao "
         "constitui oferta publica de valores mobiliarios nem solicitacao de "
         "investimento. As informacoes foram fornecidas pela Companhia e por "
         "fontes de mercado. A ZYN Capital nao garante a exatidao ou completude "
         "dos dados. Investimentos envolvem riscos, incluindo risco de credito, "
         "liquidez, mercado e operacional. Rentabilidade passada nao e garantia "
         "de resultados futuros.",
         size=6.5, color=RGBColor(0x64, 0x74, 0x8B))

    # Footer
    _add_rect(slide, 0, SH - Emu(365760), SW, Emu(365760), NAVY)
    _txt(slide, Emu(365760), SH - Emu(310896),
         Emu(6858000), Emu(274320),
         f"ZYN CAPITAL  (C)  {datetime.now().year}  |  Confidencial  |  Credito Estruturado & M&A",
         size=7, color=SUBTITLE_COLOR)


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------
def generate_teaser(
    analise: dict[str, Any],
    parametros: dict[str, Any],
    output_path: str,
) -> str:
    """
    Gera Teaser Premium ZYN de 8-9 slides (.pptx).

    Baseado na analise de 8 teasers reais (SNOW FIDC, Ivanoff SLB, Baroon55 CRI,
    Grupo Roca, Goulart, Edificatto, Frigomarca, Fibra Cotton).

    Slides:
        1. Capa Premium (tomador, KPIs, resumo)
        2. Resumo Executivo (termos + indicadores-chave)
        3. Perfil do Tomador (dados societarios + destaques)
        4. Indicadores Financeiros (KPIs + capacidade pagamento)
        5. Estrutura da Operacao (termos + fluxo/uso recursos)
        6. Garantias & Colateral (KPIs + tabela detalhada)
        7. Riscos & Mitigantes (matriz de riscos)
        8. Parecer & Rating (nota final + ratings por secao)
        9. Proximos Passos & Contato

    Args:
        analise: Resultado da analise de credito (MAC). Deve conter:
            - kpis: dict com receita_liquida, ebitda, margem_ebitda, dscr, ltv, etc.
            - rating_final: dict com nota, parecer, justificativa
            - tomador, patrimonio, producao, capital, operacao, pagamento,
              onus, riscos, covenants, cronograma: dicts das secoes da MAC
        parametros: Parametros da operacao. Campos suportados:
            - tomador (str): Nome do tomador/cedente
            - cnpj (str): CNPJ do tomador
            - tipo_operacao (str): CRI, CRA, CPR-F, SLB, NC/CCB, FIDC, Fiagro, Debenture
            - volume (float): Volume em R$
            - taxa (str): Taxa indicativa (ex: "CDI + 5,0% a.a.")
            - prazo_meses (int): Prazo total em meses
            - amortizacao (str): Tipo de amortizacao
            - carencia (str): Periodo de carencia
            - garantias_text (str): Descricao das garantias
            - garantias (list[dict]): Lista detalhada de garantias
            - finalidade (str): Finalidade dos recursos
            - setor (str): Setor de atuacao
            - localidade (str): Localizacao
            - resumo_executivo (str): Resumo para a capa
            - fluxo_caixa (list[dict]): Projecao de fluxo
            - destaques (list[str]): Destaques do tomador
            - assessor, email_assessor, telefone_assessor (str): Dados de contato
        output_path: Caminho de saida do arquivo .pptx.

    Returns:
        Caminho do arquivo gerado.
    """
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH

    # 5 slides principais
    total = 5

    _slide_cover(prs, parametros, analise)              # 1. Capa
    _slide_resumo(prs, parametros, analise, 2, total)    # 2. Resumo Executivo
    _slide_financeiros(prs, analise, 3, total)           # 3. Indicadores Financeiros
    _slide_estrutura(prs, parametros, analise, 4, total) # 4. Estrutura & Garantias
    _slide_parecer(prs, parametros, analise, 5, total)   # 5. Parecer & Contato

    prs.save(output_path)
    return output_path
