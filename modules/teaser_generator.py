"""
ZYN Capital — Gerador de Teaser (.pptx) baseado em template oficial.

Abre o template Teaser_Template_ZYN.pptx (5 slides, 10x5.62 in, 16:9) e
substitui os placeholders pelo conteudo real da operacao, preservando toda
a formatacao visual (fontes, cores, tamanhos, posicoes).

Template:
    Slide 1 — Cover
    Slide 2 — Resumo Executivo & Termos Indicativos (KPI cards + tabela)
    Slide 3 — Overview da Empresa (info + KPI cards)
    Slide 4 — Estrutura da Operacao & Indicadores Financeiros (flow + tabelas)
    Slide 5 — Garantias & Soundness (6 cards + disclaimer)
"""

from datetime import datetime
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.util import Inches

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
TEMPLATE_PATH = Path(__file__).parent.parent / "templates" / "Teaser_Template_ZYN.pptx"

MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Marco", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}


# ---------------------------------------------------------------------------
# Helpers — Formatacao BR (mantidos para compatibilidade)
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


def _safe_get(d, *keys, default="—"):
    """Navegacao segura em dicts aninhados. Retorna default se qualquer nivel nao for dict."""
    current = d
    for k in keys:
        if isinstance(current, dict):
            current = current.get(k, default)
        else:
            return default
    return current if current not in (None, "", 0, 0.0) else default


def _safe_dict(value, default=None):
    """Garante que o valor é um dict; retorna default={} se nao for."""
    return value if isinstance(value, dict) else (default or {})


def _current_date_pt() -> str:
    """Retorna 'Marco 2026' no formato portugues."""
    now = datetime.now()
    return f"{MESES_PT.get(now.month, '')} {now.year}"


# ---------------------------------------------------------------------------
# Core: substituicao de texto preservando formatacao
# ---------------------------------------------------------------------------
def _replace_text_in_runs(text_frame, old_text: str, new_text: str) -> bool:
    """
    Substitui texto em runs de um text_frame, preservando formatacao.
    Tenta primeiro run-a-run; se nao encontrar, tenta concatenar runs
    de cada paragrafo para lidar com texto dividido entre runs.
    """
    replaced = False
    for para in text_frame.paragraphs:
        # Tentativa 1: substituicao direta em cada run
        for run in para.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
                replaced = True
        # Tentativa 2: texto dividido entre runs adjacentes
        if not replaced:
            full_text = "".join(r.text for r in para.runs)
            if old_text in full_text:
                new_full = full_text.replace(old_text, new_text)
                # Coloca todo o texto no primeiro run, limpa os demais
                if para.runs:
                    para.runs[0].text = new_full
                    for run in para.runs[1:]:
                        run.text = ""
                    replaced = True
    return replaced


def _replace_in_shape(shape, old_text: str, new_text: str) -> bool:
    """Substitui texto em shape (text_frame), preservando formatacao."""
    if not shape.has_text_frame:
        return False
    return _replace_text_in_runs(shape.text_frame, old_text, new_text)


def _replace_in_table_cell(table, row: int, col: int, new_text: str):
    """Substitui todo o conteudo de uma celula de tabela, preservando formatacao do primeiro run."""
    try:
        cell = table.cell(row, col)
        if cell.text_frame.paragraphs and cell.text_frame.paragraphs[0].runs:
            cell.text_frame.paragraphs[0].runs[0].text = str(new_text)
            for run in cell.text_frame.paragraphs[0].runs[1:]:
                run.text = ""
        else:
            cell.text = str(new_text)
    except (IndexError, AttributeError):
        pass


def _replace_in_table(table, old_text: str, new_text: str) -> bool:
    """Busca e substitui texto em qualquer celula da tabela."""
    replaced = False
    for row_idx in range(len(table.rows)):
        for col_idx in range(len(table.columns)):
            cell = table.cell(row_idx, col_idx)
            try:
                if _replace_text_in_runs(cell.text_frame, old_text, new_text):
                    replaced = True
            except Exception:
                pass
    return replaced


def _replace_on_slide(slide, old_text: str, new_text: str) -> bool:
    """Substitui texto em todas as shapes (incluindo tabelas) de um slide."""
    replaced = False
    for shape in slide.shapes:
        if shape.has_text_frame:
            if _replace_in_shape(shape, old_text, new_text):
                replaced = True
        if shape.has_table:
            if _replace_in_table(shape.table, old_text, new_text):
                replaced = True
    return replaced


# ---------------------------------------------------------------------------
# Slide 1 — Cover
# ---------------------------------------------------------------------------
def _fill_cover(slide, analise: dict, parametros: dict):
    tipo = parametros.get("tipo_operacao", _safe_get(analise, "operacao", "instrumento", default="NC/CCB"))
    tomador = parametros.get("tomador", _safe_get(analise, "tomador", "nome", default=""))
    data_pt = _current_date_pt()

    _replace_on_slide(slide, "[TIPO DE INSTRUMENTO]", str(tipo).upper())
    _replace_on_slide(slide, "[Nome da Operação / Tomador]", str(tomador))
    # Atualiza data (template tem "Março 2026")
    _replace_on_slide(slide, "Março 2026", data_pt)
    _replace_on_slide(slide, "Marco 2026", data_pt)


# ---------------------------------------------------------------------------
# Slide 2 — Resumo Executivo & Termos Indicativos
# ---------------------------------------------------------------------------
def _fill_resumo(slide, analise: dict, parametros: dict):
    data_pt = _current_date_pt()
    _replace_on_slide(slide, "Março 2026", data_pt)
    _replace_on_slide(slide, "Marco 2026", data_pt)

    # --- KPI cards ---
    volume_raw = parametros.get("volume", _safe_get(analise, "operacao", "volume", default=0))
    taxa_raw = parametros.get("taxa", _safe_get(analise, "operacao", "taxa", default=""))
    prazo_raw = parametros.get("prazo_meses", _safe_get(analise, "operacao", "prazo", default=""))
    ltv_raw = _safe_get(analise, "kpis", "ltv", default="")

    volume_str = _fmt_brl(volume_raw) if volume_raw and volume_raw != "—" else "—"
    taxa_str = str(taxa_raw) if taxa_raw and taxa_raw != "—" else "—"
    prazo_str = f"{prazo_raw} meses" if prazo_raw and prazo_raw != "—" else "—"
    ltv_str = _fmt_pct(ltv_raw) if ltv_raw and ltv_raw != "—" else "—"

    _replace_on_slide(slide, "R$ [XX] MM", volume_str)
    _replace_on_slide(slide, "CDI + [X]%", taxa_str)
    _replace_on_slide(slide, "[XX] meses", prazo_str)
    _replace_on_slide(slide, "[XX]%", ltv_str)

    # --- Tabela de termos (11 rows x 2 cols, header em row 0) ---
    table = None
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            break
    if table is None:
        return

    tomador = parametros.get("tomador", _safe_get(analise, "tomador", "nome", default="—"))
    instrumento = parametros.get("tipo_operacao", _safe_get(analise, "operacao", "instrumento", default="—"))
    volume_full = _fmt_brl(volume_raw) if volume_raw and volume_raw != "—" else "—"
    taxa = str(taxa_raw) if taxa_raw and taxa_raw != "—" else "—"
    prazo = str(prazo_raw) + " meses" if prazo_raw and prazo_raw != "—" else "—"
    amort = parametros.get("amortizacao", _safe_get(analise, "operacao", "amortizacao", default="—"))
    carencia = parametros.get("carencia", _safe_get(analise, "operacao", "carencia", default="—"))
    garantias = parametros.get("garantias_text", _safe_get(analise, "operacao", "garantias", default="—"))
    finalidade = parametros.get("finalidade", _safe_get(analise, "operacao", "finalidade", default="—"))
    rating = _safe_get(analise, "rating_final", "nota", default="—")

    termos = [tomador, instrumento, volume_full, taxa, prazo, amort, carencia, garantias, finalidade, rating]
    for i, val in enumerate(termos):
        _replace_in_table_cell(table, i + 1, 1, str(val))


# ---------------------------------------------------------------------------
# Slide 3 — Overview da Empresa
# ---------------------------------------------------------------------------
def _fill_overview(slide, analise: dict, parametros: dict):
    data_pt = _current_date_pt()
    _replace_on_slide(slide, "Março 2026", data_pt)
    _replace_on_slide(slide, "Marco 2026", data_pt)

    # --- Left panel: company info ---
    nome = parametros.get("tomador", _safe_get(analise, "tomador", "nome", default="[NOME]"))
    fundacao = _safe_get(analise, "tomador", "fundacao", default="[Ano]")
    sede = parametros.get("localidade", _safe_get(analise, "tomador", "sede", default="[Cidade/UF]"))
    segmento = parametros.get("setor", _safe_get(analise, "tomador", "segmento", default="[Segmento]"))
    socios = _safe_get(analise, "tomador", "socios", default="[Sócios / Gestão]")
    descricao = _safe_get(analise, "tomador", "descricao", default="[Descrição da empresa]")

    _replace_on_slide(slide, "[NOME DO GRUPO / EMPRESA]", str(nome))
    # Fundacao, Sede, Segmento sao parte do text block no shape 9
    _replace_on_slide(slide, "[Ano]", str(fundacao))
    _replace_on_slide(slide, "[Cidade/UF]", str(sede))
    _replace_on_slide(slide, "[Agronegócio / Imobiliário / Industrial / etc.]", str(segmento))

    # Socios e descricao
    _replace_on_slide(slide, "[Nomes e cargos principais]", str(socios))
    _replace_on_slide(
        slide,
        "[Breve histórico da empresa, atividades principais, diferenciais competitivos, principais clientes/offtakers, e posicionamento de mercado. 3-4 linhas.]",
        str(descricao),
    )

    # --- Right panel: KPI cards ---
    receita = _safe_get(analise, "kpis", "receita_liquida", default=0)
    ebitda = _safe_get(analise, "kpis", "ebitda", default=0)
    colaboradores = _safe_get(analise, "tomador", "colaboradores", default="—")
    capacidade = _safe_get(analise, "tomador", "capacidade", default="—")
    unidades = _safe_get(analise, "tomador", "unidades", default="—")
    clientes = _safe_get(analise, "tomador", "principais_clientes", default="—")

    _replace_on_slide(slide, "R$ [XXX] MM", _fmt_brl(receita) if receita != "—" else "—")
    _replace_on_slide(slide, "R$ [XX] MM", _fmt_brl(ebitda) if ebitda != "—" else "—")
    _replace_on_slide(slide, "[XXX]", str(colaboradores))
    _replace_on_slide(slide, "[XX.XXX ha / m² / ton]", str(capacidade))
    _replace_on_slide(slide, "[X] unidades em [UFs]", str(unidades))
    _replace_on_slide(slide, "[Cliente A, B, C]", str(clientes))


# ---------------------------------------------------------------------------
# Slide 4 — Estrutura da Operacao & Indicadores Financeiros
# ---------------------------------------------------------------------------
def _fill_estrutura(slide, analise: dict, parametros: dict):
    data_pt = _current_date_pt()
    _replace_on_slide(slide, "Março 2026", data_pt)
    _replace_on_slide(slide, "Marco 2026", data_pt)

    # --- Flow boxes ---
    tomador = parametros.get("tomador", _safe_get(analise, "tomador", "nome", default="TOMADOR"))
    tipo = parametros.get("tipo_operacao", _safe_get(analise, "operacao", "instrumento", default=""))
    tipo_upper = str(tipo).upper()

    # Determine veiculo com base no tipo
    veiculos_map = {
        "CRA": "SECURITIZADORA", "CRI": "SECURITIZADORA",
        "FIDC": "FIDC", "FIAGRO": "FIAGRO",
        "DEBENTURE": "EMISSORA", "NC": "BANCO EMISSOR", "CCB": "BANCO EMISSOR",
        "NC/CCB": "BANCO EMISSOR", "SLB": "OPERADOR S&LB",
    }
    veiculo = veiculos_map.get(tipo_upper, "VEÍCULO")

    _replace_on_slide(slide, "[TOMADOR / CEDENTE]", str(tomador).upper())
    _replace_on_slide(slide, "[SECURITIZADORA / VEÍCULO]", veiculo)
    _replace_on_slide(slide, "[INVESTIDOR]", "INVESTIDOR")

    # --- Destinacao table (shape 19 in template: 4 rows x 3 cols) ---
    destinacao_table = None
    financeiros_table = None
    for shape in slide.shapes:
        if shape.has_table:
            tbl = shape.table
            ncols = len(tbl.columns)
            if ncols == 3:
                destinacao_table = tbl
            elif ncols == 4:
                financeiros_table = tbl

    # Destinacao: tentamos preencher com dados de parametros ou analise
    if destinacao_table:
        destinacao = parametros.get("destinacao", [])
        if not destinacao:
            finalidade = parametros.get("finalidade", _safe_get(analise, "operacao", "finalidade", default=""))
            volume_raw = parametros.get("volume", _safe_get(analise, "operacao", "volume", default=0))
            if finalidade and finalidade != "—":
                destinacao = [{"destino": str(finalidade), "valor": volume_raw, "pct": "100%"}]

        for i, item in enumerate(destinacao[:3]):  # max 3 data rows
            row_idx = i + 1
            if isinstance(item, str):
                _replace_in_table_cell(destinacao_table, row_idx, 0, item)
                continue
            if not isinstance(item, dict):
                continue
            _replace_in_table_cell(destinacao_table, row_idx, 0, str(item.get("destino", "—")))
            val = item.get("valor", "—")
            try:
                val_mm = f"{float(val) / 1_000_000:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except (TypeError, ValueError):
                val_mm = str(val)
            _replace_in_table_cell(destinacao_table, row_idx, 1, val_mm)
            _replace_in_table_cell(destinacao_table, row_idx, 2, str(item.get("pct", "—")))

        # Limpa linhas nao usadas
        for i in range(len(destinacao), 3):
            row_idx = i + 1
            _replace_in_table_cell(destinacao_table, row_idx, 0, "—")
            _replace_in_table_cell(destinacao_table, row_idx, 1, "—")
            _replace_in_table_cell(destinacao_table, row_idx, 2, "—")

    # --- Indicadores Financeiros table (9 rows x 4 cols) ---
    if financeiros_table:
        kpis = analise.get("kpis", {})
        if not isinstance(kpis, dict):
            kpis = {}
        historico = analise.get("historico_financeiro", {})
        if not isinstance(historico, dict):
            historico = {}
        now = datetime.now()
        anos = [str(now.year - 2), str(now.year - 1), str(now.year)]

        # Headers: anos
        for ci, ano in enumerate(anos):
            _replace_in_table_cell(financeiros_table, 0, ci + 1, ano)

        # Mapping: row -> kpi key, formatter
        row_map = [
            ("receita_liquida", _fmt_brl),
            ("ebitda", _fmt_brl),
            ("margem_ebitda", _fmt_pct),
            ("divida_liquida", _fmt_brl),
            ("div_liq_ebitda", _fmt_mult),
            ("dscr", _fmt_mult),
            ("ltv", _fmt_pct),
            ("pl", _fmt_brl),
        ]

        for row_idx, (key, fmt) in enumerate(row_map):
            for col_idx, ano in enumerate(anos):
                # Tenta historico primeiro, depois kpi geral (apenas para ano corrente)
                val = "—"
                if historico:
                    val_raw = _safe_get(historico, ano, key, default="—")
                    if val_raw != "—":
                        val = fmt(val_raw)
                if val == "—" and col_idx == 2:  # Ano corrente: tenta kpis direto
                    val_raw = kpis.get(key, "—")
                    if val_raw and val_raw != "—":
                        val = fmt(val_raw)
                _replace_in_table_cell(financeiros_table, row_idx + 1, col_idx + 1, val)


# ---------------------------------------------------------------------------
# Slide 5 — Garantias & Soundness
# ---------------------------------------------------------------------------
def _fill_garantias(slide, analise: dict, parametros: dict):
    data_pt = _current_date_pt()
    _replace_on_slide(slide, "Março 2026", data_pt)
    _replace_on_slide(slide, "Marco 2026", data_pt)

    garantias_list = parametros.get("garantias", analise.get("garantias", []))
    if not isinstance(garantias_list, list):
        garantias_list = []

    # Mapa de tipo de garantia -> placeholder no template
    placeholder_map = {
        "alienacao_fiduciaria": "[Descrição do imóvel/bem, matrícula, localização, valor de avaliação]",
        "alienação fiduciária": "[Descrição do imóvel/bem, matrícula, localização, valor de avaliação]",
        "alienacao fiduciaria": "[Descrição do imóvel/bem, matrícula, localização, valor de avaliação]",
        "cessao_fiduciaria": "[Recebíveis cedidos, fluxo, prazo, valor estimado do lastro]",
        "cessão fiduciária": "[Recebíveis cedidos, fluxo, prazo, valor estimado do lastro]",
        "cessao fiduciaria": "[Recebíveis cedidos, fluxo, prazo, valor estimado do lastro]",
        "aval": "[Avalistas PF/PJ, patrimônio declarado, vínculos com o tomador]",
        "fianca": "[Avalistas PF/PJ, patrimônio declarado, vínculos com o tomador]",
        "aval_fianca": "[Avalistas PF/PJ, patrimônio declarado, vínculos com o tomador]",
        "aval / fiança": "[Avalistas PF/PJ, patrimônio declarado, vínculos com o tomador]",
        "fundo_reserva": "[X] parcelas equivalentes — constituição [pré/pós] emissão",
        "fundo de reserva": "[X] parcelas equivalentes — constituição [pré/pós] emissão",
    }

    # Tenta fazer match por tipo
    for gar in garantias_list:
        if isinstance(gar, str):
            # Garantia veio como string simples
            _replace_on_slide(slide, list(placeholder_map.values())[0], gar)
            continue
        if not isinstance(gar, dict):
            continue
        tipo_raw = str(gar.get("tipo", "")).lower().strip()
        desc = str(gar.get("descricao", "—"))
        placeholder = placeholder_map.get(tipo_raw)
        if placeholder:
            _replace_on_slide(slide, placeholder, desc)

    # Razao de garantia
    ltv = _safe_get(analise, "kpis", "ltv", default="")
    if ltv and ltv != "—":
        try:
            razao = 1 / float(ltv) if 0 < float(ltv) <= 1 else float(ltv)
            razao_str = _fmt_mult(razao)
        except (TypeError, ValueError, ZeroDivisionError):
            razao_str = "—"
    else:
        razao_str = "—"
    _replace_on_slide(slide, "[X,Xx]x", razao_str)

    # Nota de soundness
    parecer = _safe_get(analise, "rating_final", "parecer", default="")
    justificativa = _safe_get(analise, "rating_final", "justificativa", default="")
    soundness_text = str(parecer) if parecer != "—" else str(justificativa) if justificativa != "—" else "—"
    _replace_on_slide(
        slide,
        "[Resumo da tese de crédito: por que os riscos estão mitigados. 2-3 linhas.]",
        soundness_text[:300],
    )


# ---------------------------------------------------------------------------
# Fallback: gera apresentacao minima se template nao existir
# ---------------------------------------------------------------------------
def _generate_fallback(analise: dict, parametros: dict, output_path: str) -> str:
    """Gera apresentacao minimalista quando o template nao esta disponivel."""
    prs = Presentation()
    prs.slide_width = 9144000  # 10in
    prs.slide_height = 5143500  # 5.625in

    layout = prs.slide_layouts[6]  # blank

    # Slide unico com dados basicos
    slide = prs.slides.add_slide(layout)
    from pptx.util import Pt
    from pptx.dml.color import RGBColor

    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(4.5))
    tf = txBox.text_frame
    tf.word_wrap = True

    tomador = parametros.get("tomador", _safe_get(analise, "tomador", "nome", default="N/A"))
    tipo = parametros.get("tipo_operacao", "N/A")
    volume = _fmt_brl(parametros.get("volume", 0))
    nota = _safe_get(analise, "rating_final", "nota", default="N/A")

    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "ZYN CAPITAL — TEASER"
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x22, 0x30, 0x40)

    for line in [
        f"\n\nTomador: {tomador}",
        f"Instrumento: {tipo}",
        f"Volume: {volume}",
        f"Rating: {nota}",
        f"\n{_current_date_pt()}",
        "\n\nCONFIDENCIAL — DISTRIBUIÇÃO RESTRITA",
        "\n⚠ Template oficial nao encontrado. Reinstale o template em templates/Teaser_Template_ZYN.pptx",
    ]:
        p2 = tf.add_paragraph()
        run2 = p2.add_run()
        run2.text = line
        run2.font.size = Pt(12)
        run2.font.color.rgb = RGBColor(0x34, 0x40, 0x50)

    prs.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
# Funcao principal
# ---------------------------------------------------------------------------
def generate_teaser(
    analise: dict[str, Any],
    parametros: dict[str, Any],
    output_path: str,
) -> str:
    """
    Gera Teaser ZYN (.pptx) a partir do template oficial.

    Abre o template de 5 slides, substitui os placeholders pelos dados reais
    da operacao, e salva no output_path. Preserva toda a formatacao visual.

    Args:
        analise: Resultado da analise de credito (MAC).
        parametros: Parametros da operacao (tomador, volume, taxa, etc.).
        output_path: Caminho de saida do arquivo .pptx.

    Returns:
        Caminho do arquivo gerado.
    """
    # Garante que analise e parametros sao dicts
    if not isinstance(analise, dict):
        analise = {}
    if not isinstance(parametros, dict):
        parametros = {}

    if not TEMPLATE_PATH.exists():
        return _generate_fallback(analise, parametros, output_path)

    prs = Presentation(str(TEMPLATE_PATH))

    slides = list(prs.slides)
    if len(slides) < 5:
        return _generate_fallback(analise, parametros, output_path)

    _fill_cover(slides[0], analise, parametros)
    _fill_resumo(slides[1], analise, parametros)
    _fill_overview(slides[2], analise, parametros)
    _fill_estrutura(slides[3], analise, parametros)
    _fill_garantias(slides[4], analise, parametros)

    prs.save(output_path)
    return output_path
