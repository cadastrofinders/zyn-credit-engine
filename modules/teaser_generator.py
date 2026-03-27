"""
ZYN Capital — Gerador de Teaser (.pptx) baseado em template oficial.
v2.1 — Deep extraction com suporte a formato direto e wrapper.

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
# Helpers — Deep extraction from analise dict
# ---------------------------------------------------------------------------
def _deep_get(analise: dict, *paths, default="—"):
    """
    Busca um valor em multiplos caminhos possiveis no dict de analise.
    O Opus retorna dados em analise.analise.X ou analise.X dependendo do contexto.
    Tenta todos os caminhos ate achar um valor valido.
    """
    for path in paths:
        val = _safe_get(analise, *path, default=None)
        if val not in (None, "", 0, 0.0, "—", []):
            return val
    return default


def _extract_company_data(analise: dict, parametros: dict) -> dict:
    """Extrai todos os dados da empresa de todas as secoes da analise.

    Suporta dois formatos:
    1. Dict direto do Opus: {kpis: {}, tomador: {}, rating_final: {}, ...}
    2. Dict salvo com wrapper: {analise: {kpis: {}, ...}, operacao: {...}}
    """
    # Detecta se tem camada wrapper ou se é direto
    if isinstance(analise.get("analise"), dict) and "kpis" in analise.get("analise", {}):
        # Formato wrapper: {analise: {...}, operacao: {...}}
        a = _safe_dict(analise.get("analise"))
    elif "kpis" in analise or "tomador" in analise or "rating_final" in analise:
        # Formato direto do Opus
        a = _safe_dict(analise)
    else:
        a = {}

    tom_a = _safe_dict(a.get("tomador"))
    tom_root = _safe_dict(analise.get("tomador"))
    op_root = _safe_dict(analise.get("operacao")) if isinstance(analise.get("operacao"), dict) else {}
    op_a = _safe_dict(a.get("operacao"))
    kpis_a = _safe_dict(a.get("kpis"))
    kpis_root = _safe_dict(analise.get("kpis"))
    cap = _safe_dict(a.get("capital"))
    cap_ind = _safe_dict(cap.get("indicadores"))
    cap_end = _safe_dict(cap.get("endividamento"))
    pat = _safe_dict(a.get("patrimonio"))
    pat_ativos = _safe_dict(pat.get("ativos_reais"))
    prod = _safe_dict(a.get("producao"))
    pag = _safe_dict(a.get("pagamento"))
    rating_a = _safe_dict(a.get("rating_final")) if isinstance(a.get("rating_final"), dict) else {}
    rating_root = _safe_dict(analise.get("rating_final")) if isinstance(analise.get("rating_final"), dict) else {}
    rating = rating_a if rating_a else rating_root

    return {
        # Company info
        "nome": parametros.get("tomador") or tom_a.get("razao_social") or tom_root.get("nome") or tom_root.get("razao_social") or "—",
        "cnpj": parametros.get("cnpj") or tom_a.get("cnpj") or tom_root.get("cnpj") or "—",
        "fundacao": tom_a.get("fundacao") or tom_root.get("fundacao") or _extract_year_from_text(tom_a.get("historico", "")) or _extract_year_from_text(prod.get("historico_produtivo", "")) or "—",
        "sede": parametros.get("localidade") or tom_a.get("sede") or tom_root.get("sede") or _extract_city_from_text(tom_a.get("historico", "")) or "—",
        "segmento": parametros.get("setor") or tom_a.get("setor") or tom_root.get("segmento") or tom_root.get("setor") or "—",
        "socios": tom_a.get("grupo_economico") or tom_a.get("socios") or tom_root.get("socios") or tom_root.get("grupo_economico") or "—",
        "descricao": tom_a.get("historico") or tom_root.get("descricao") or tom_root.get("historico") or prod.get("capacidade") or "—",
        "colaboradores": tom_a.get("colaboradores") or tom_root.get("colaboradores") or _extract_number_near_keyword((tom_a.get("historico", "") if isinstance(tom_a.get("historico"), str) else "") + " " + (prod.get("capacidade", "") if isinstance(prod.get("capacidade"), str) else ""), "colaborador") or "—",
        "capacidade": tom_a.get("capacidade") or tom_root.get("capacidade") or _extract_number_near_keyword(prod.get("capacidade", ""), "equipamento") or _extract_number_near_keyword(prod.get("capacidade", ""), "hectare") or _extract_number_near_keyword(prod.get("capacidade", ""), " ha") or "—",
        "unidades": tom_a.get("unidades") or tom_root.get("unidades") or _extract_number_near_keyword(prod.get("capacidade", ""), "unidade") or "—",
        "clientes": tom_a.get("principais_clientes") or tom_root.get("principais_clientes") or _extract_clients_from_text(prod.get("analise", "")) or "—",

        # Financials
        "receita": kpis_a.get("receita_liquida") or kpis_root.get("receita_liquida") or 0,
        "ebitda": kpis_a.get("ebitda") or kpis_root.get("ebitda") or 0,
        "margem_ebitda": kpis_a.get("margem_ebitda") or kpis_root.get("margem_ebitda") or 0,
        "divida_liquida": cap_end.get("divida_liquida_2025") or cap_end.get("divida_liquida") or kpis_a.get("divida_liquida") or 0,
        "div_liq_ebitda": cap_ind.get("divida_liquida_ebitda") or kpis_a.get("divida_liquida_ebitda") or kpis_root.get("divida_liquida_ebitda") or 0,
        "dscr": pag.get("dscr") or kpis_a.get("dscr") or kpis_root.get("dscr") or 0,
        "ltv": kpis_a.get("ltv") or kpis_root.get("ltv") or pat.get("ltv") or 0,
        "pl": pat_ativos.get("patrimonio_liquido_2025") or pat_ativos.get("patrimonio_liquido") or kpis_a.get("pl") or 0,
        "liquidez": cap_ind.get("liquidez_corrente") or 0,

        # Operation
        "instrumento": parametros.get("tipo_operacao") or op_root.get("tipo_operacao") or op_a.get("instrumento") or op_root.get("instrumento") or "—",
        "volume": parametros.get("volume") or op_root.get("volume") or op_a.get("volume") or 0,
        "taxa": parametros.get("taxa") or op_root.get("taxa") or op_a.get("taxa") or "—",
        "prazo": parametros.get("prazo_meses") or op_root.get("prazo_meses") or op_a.get("prazo") or "—",
        "amortizacao": parametros.get("amortizacao") or op_root.get("amortizacao") or "—",
        "carencia": parametros.get("carencia") or _extract_from_text((op_a.get("estrutura", "") if isinstance(op_a.get("estrutura"), str) else "") + " " + (op_a.get("prazo", "") if isinstance(op_a.get("prazo"), str) else ""), r"(?:[Cc]ar[eê]ncia\s*(?:de\s*)?)(\d+\s*meses?)") or "—",
        "garantias_text": parametros.get("garantias_text") or op_root.get("garantias_text") or "—",
        "finalidade": parametros.get("finalidade") or op_a.get("analise", "")[:200] if isinstance(op_a.get("analise"), str) else "—",

        # Guarantees (detailed)
        "garantias_detalhadas": pat.get("garantias_detalhadas") or [],

        # Rating
        "rating": rating.get("nota") or op_root.get("rating") or "—",
        "parecer": rating.get("parecer") or op_root.get("parecer") or "—",
        "justificativa": rating.get("justificativa") or "—",
    }


def _extract_from_text(text: str, pattern: str) -> str:
    """Extrai valor de texto usando regex. Retorna '' se nao encontrar."""
    import re
    if not isinstance(text, str):
        return ""
    match = re.search(pattern, text)
    return match.group(1) if match else ""


def _extract_year_from_text(text: str) -> str:
    """Extrai ano de fundacao de texto tipo 'Fundada em 1978' ou 'desde 1985'."""
    import re
    if not isinstance(text, str) or not text:
        return ""
    for pattern in [r'[Ff]undad[ao]\s+em\s+(\d{4})', r'desde\s+(\d{4})', r'criada?\s+em\s+(\d{4})', r'(\d{4})\s+em\s+\w']:
        m = re.search(pattern, text)
        if m:
            year = m.group(1)
            if 1900 <= int(year) <= 2030:
                return year
    return ""


def _extract_city_from_text(text: str) -> str:
    """Extrai cidade/UF de texto tipo 'em Chapeco-SC' ou 'sediada em Sao Paulo/SP'."""
    import re
    if not isinstance(text, str) or not text:
        return ""
    # Tenta padroes comuns: Cidade-UF, Cidade/UF, Cidade (UF)
    for pattern in [r'em\s+([A-Z][\w\u00C0-\u00FF]+(?:\s+[\w\u00C0-\u00FF]+)*\s*[-/]\s*[A-Z]{2})',
                    r'sediada?\s+em\s+([\w\u00C0-\u00FF]+(?:\s+[\w\u00C0-\u00FF]+)*)',
                    r'sede\s+(?:em\s+)?([\w\u00C0-\u00FF]+[-/][A-Z]{2})']:
        m = re.search(pattern, text)
        if m:
            return m.group(1).strip()
    return ""


def _extract_number_near_keyword(text: str, keyword: str) -> str:
    """Extrai numero proximo a uma keyword. Ex: '~3.000 equipamentos' -> '3.000'."""
    import re
    if not isinstance(text, str) or not text or not keyword:
        return ""
    # Busca: ~N.NNN keyword ou keyword N.NNN
    pattern = rf'~?(\d[\d.]*)\s*{keyword}'
    m = re.search(pattern, text, re.IGNORECASE)
    if m:
        return m.group(1)
    pattern = rf'{keyword}\w*\s+~?(\d[\d.]*)'
    m = re.search(pattern, text, re.IGNORECASE)
    if m:
        return m.group(1)
    return ""


def _extract_clients_from_text(text: str) -> str:
    """Extrai nomes de clientes de texto tipo '(Ambev, Heineken, JBS, Suzano)'."""
    import re
    if not isinstance(text, str) or not text:
        return ""
    # Busca lista entre parenteses com nomes proprios
    m = re.search(r'\(([A-Z][\w]+(?:,\s*[A-Z][\w]+){2,})\)', text)
    if m:
        return m.group(1)
    return ""


# ---------------------------------------------------------------------------
# Slide 1 — Cover
# ---------------------------------------------------------------------------
def _fill_cover(slide, analise: dict, parametros: dict):
    data = _extract_company_data(analise, parametros)
    data_pt = _current_date_pt()

    _replace_on_slide(slide, "[TIPO DE INSTRUMENTO]", str(data["instrumento"]).upper())
    _replace_on_slide(slide, "[Nome da Operação / Tomador]", str(data["nome"]))
    _replace_on_slide(slide, "Março 2026", data_pt)
    _replace_on_slide(slide, "Marco 2026", data_pt)


# ---------------------------------------------------------------------------
# Slide 2 — Resumo Executivo & Termos Indicativos
# ---------------------------------------------------------------------------
def _fill_resumo(slide, analise: dict, parametros: dict):
    data = _extract_company_data(analise, parametros)
    data_pt = _current_date_pt()
    _replace_on_slide(slide, "Março 2026", data_pt)
    _replace_on_slide(slide, "Marco 2026", data_pt)

    # --- KPI cards ---
    volume_str = _fmt_brl(data["volume"]) if data["volume"] else "—"
    taxa_str = str(data["taxa"]) if data["taxa"] != "—" else "—"
    prazo_str = f"{data['prazo']}" if data["prazo"] != "—" else "—"
    if prazo_str != "—" and "mes" not in str(prazo_str).lower():
        prazo_str = f"{prazo_str} meses"
    ltv_str = _fmt_pct(data["ltv"]) if data["ltv"] else "—"

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

    termos = [
        str(data["nome"]),
        str(data["instrumento"]),
        volume_str,
        taxa_str,
        prazo_str,
        str(data["amortizacao"]),
        str(data["carencia"]),
        str(data["garantias_text"]),
        str(data["finalidade"])[:200],
        str(data["rating"]),
    ]
    for i, val in enumerate(termos):
        _replace_in_table_cell(table, i + 1, 1, val)


# ---------------------------------------------------------------------------
# Slide 3 — Overview da Empresa
# ---------------------------------------------------------------------------
def _fill_overview(slide, analise: dict, parametros: dict):
    data = _extract_company_data(analise, parametros)
    data_pt = _current_date_pt()
    _replace_on_slide(slide, "Março 2026", data_pt)
    _replace_on_slide(slide, "Marco 2026", data_pt)

    # --- Left panel: company info ---
    _replace_on_slide(slide, "[NOME DO GRUPO / EMPRESA]", str(data["nome"]))
    _replace_on_slide(slide, "[Ano]", str(data["fundacao"]))
    _replace_on_slide(slide, "[Cidade/UF]", str(data["sede"]))
    _replace_on_slide(slide, "[Agronegócio / Imobiliário / Industrial / etc.]", str(data["segmento"]))
    _replace_on_slide(slide, "[Nomes e cargos principais]", str(data["socios"]))
    # Descricao: truncar para caber no box
    desc = str(data["descricao"])[:500]
    _replace_on_slide(
        slide,
        "[Breve histórico da empresa, atividades principais, diferenciais competitivos, principais clientes/offtakers, e posicionamento de mercado. 3-4 linhas.]",
        desc,
    )

    # --- Right panel: KPI cards ---
    _replace_on_slide(slide, "R$ [XXX] MM", _fmt_brl(data["receita"]) if data["receita"] else "—")
    _replace_on_slide(slide, "R$ [XX] MM", _fmt_brl(data["ebitda"]) if data["ebitda"] else "—")
    _replace_on_slide(slide, "[XXX]", str(data["colaboradores"]))
    _replace_on_slide(slide, "[XX.XXX ha / m² / ton]", str(data["capacidade"]))
    _replace_on_slide(slide, "[X] unidades em [UFs]", str(data["unidades"]))
    _replace_on_slide(slide, "[Cliente A, B, C]", str(data["clientes"]))


# ---------------------------------------------------------------------------
# Slide 4 — Estrutura da Operacao & Indicadores Financeiros
# ---------------------------------------------------------------------------
def _fill_estrutura(slide, analise: dict, parametros: dict):
    data = _extract_company_data(analise, parametros)
    data_pt = _current_date_pt()
    _replace_on_slide(slide, "Março 2026", data_pt)
    _replace_on_slide(slide, "Marco 2026", data_pt)

    # --- Flow boxes ---
    tomador = str(data["nome"])
    tipo_upper = str(data["instrumento"]).upper()

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
        a = _safe_dict(analise.get("analise"))
        kpis = _safe_dict(a.get("kpis")) if a else _safe_dict(analise.get("kpis"))
        cap = _safe_dict(a.get("capital"))
        cap_ind = _safe_dict(cap.get("indicadores"))
        cap_end = _safe_dict(cap.get("endividamento"))
        pat = _safe_dict(a.get("patrimonio"))
        pat_ativos = _safe_dict(pat.get("ativos_reais"))
        pag = _safe_dict(a.get("pagamento"))

        now = datetime.now()
        anos = [str(now.year - 2), str(now.year - 1), str(now.year)]

        # Headers: anos
        for ci, ano in enumerate(anos):
            _replace_in_table_cell(financeiros_table, 0, ci + 1, ano)

        # Build latest year data from all sources
        latest = {
            "receita_liquida": data["receita"],
            "ebitda": data["ebitda"],
            "margem_ebitda": data["margem_ebitda"],
            "divida_liquida": data["divida_liquida"],
            "div_liq_ebitda": data["div_liq_ebitda"],
            "dscr": data["dscr"],
            "ltv": data["ltv"],
            "pl": data["pl"],
        }

        # Try to get historical data
        historico = _safe_dict(analise.get("historico_financeiro"))

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
                val = "—"
                # Try historico first
                if historico:
                    val_raw = _safe_get(historico, ano, key, default=None)
                    if val_raw not in (None, "", 0, 0.0, "—"):
                        val = fmt(val_raw)
                # Last column: use latest data
                if val == "—" and col_idx == 2:
                    val_raw = latest.get(key, 0)
                    if val_raw and val_raw not in (0, 0.0, "—"):
                        val = fmt(val_raw)
                _replace_in_table_cell(financeiros_table, row_idx + 1, col_idx + 1, val)


# ---------------------------------------------------------------------------
# Slide 5 — Garantias & Soundness
# ---------------------------------------------------------------------------
def _fill_garantias(slide, analise: dict, parametros: dict):
    data = _extract_company_data(analise, parametros)
    data_pt = _current_date_pt()
    _replace_on_slide(slide, "Março 2026", data_pt)
    _replace_on_slide(slide, "Marco 2026", data_pt)

    # Busca garantias detalhadas de todas as fontes
    garantias_list = data.get("garantias_detalhadas", [])
    if not isinstance(garantias_list, list):
        garantias_list = []
    # Fallback para garantias simples
    if not garantias_list:
        garantias_list = parametros.get("garantias", analise.get("garantias", []))
        if not isinstance(garantias_list, list):
            garantias_list = []

    # Placeholders do template
    ph_alienacao = "[Descrição do imóvel/bem, matrícula, localização, valor de avaliação]"
    ph_cessao = "[Recebíveis cedidos, fluxo, prazo, valor estimado do lastro]"
    ph_aval = "[Avalistas PF/PJ, patrimônio declarado, vínculos com o tomador]"
    ph_fundo = "[X] parcelas equivalentes — constituição [pré/pós] emissão"

    # Classificacao por tipo
    alienacao_descs = []
    cessao_descs = []
    aval_descs = []
    fundo_descs = []

    for gar in garantias_list:
        if isinstance(gar, str):
            tipo_lower = gar.lower()
        elif isinstance(gar, dict):
            tipo_lower = str(gar.get("tipo_garantia", gar.get("tipo", ""))).lower()
        else:
            continue

        # Build description
        if isinstance(gar, dict):
            desc = gar.get("descricao", "")
            valor = gar.get("valor_estimado", 0)
            if valor and valor != 0:
                desc += f" (valor estimado: {_fmt_brl(valor)})"
        else:
            desc = gar

        if any(k in tipo_lower for k in ["alienação", "alienacao", "fiduciária de", "fiduciaria de", "imóv", "imov"]):
            if "cessão" not in tipo_lower and "cessao" not in tipo_lower and "recebív" not in tipo_lower:
                alienacao_descs.append(desc)
        if any(k in tipo_lower for k in ["cessão", "cessao", "recebív", "recebiv", "crédito", "credito"]):
            cessao_descs.append(desc)
        if any(k in tipo_lower for k in ["aval", "fiança", "fianca"]):
            aval_descs.append(desc)
        if any(k in tipo_lower for k in ["fundo", "reserva"]):
            fundo_descs.append(desc)

    if alienacao_descs:
        _replace_on_slide(slide, ph_alienacao, "; ".join(alienacao_descs)[:300])
    if cessao_descs:
        _replace_on_slide(slide, ph_cessao, "; ".join(cessao_descs)[:300])
    if aval_descs:
        _replace_on_slide(slide, ph_aval, "; ".join(aval_descs)[:300])
    if fundo_descs:
        _replace_on_slide(slide, ph_fundo, "; ".join(fundo_descs)[:300])

    # Razao de garantia
    ltv_val = data["ltv"]
    if ltv_val and ltv_val not in (0, 0.0, "—"):
        try:
            ltv_f = float(ltv_val)
            razao = 1 / ltv_f if 0 < ltv_f <= 1 else ltv_f
            razao_str = _fmt_mult(razao)
        except (TypeError, ValueError, ZeroDivisionError):
            razao_str = "—"
    else:
        razao_str = "—"
    _replace_on_slide(slide, "[X,Xx]x", razao_str)

    # Nota de soundness — usar justificativa completa do rating
    soundness = str(data["justificativa"])[:300] if data["justificativa"] != "—" else str(data["parecer"])
    _replace_on_slide(
        slide,
        "[Resumo da tese de crédito: por que os riscos estão mitigados. 2-3 linhas.]",
        soundness,
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
