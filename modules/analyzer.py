"""
ZYN Capital — Módulo de Análise de Crédito (MAC ZYN v3)
Usa Claude Sonnet com streaming para resposta rápida.
"""

import json
import logging
import os
import re
import time
from typing import Any

import anthropic

logger = logging.getLogger(__name__)

MODEL = "claude-sonnet-4-6"
MAX_TOKENS = 10000
MAX_RETRIES = 2
RETRY_WAIT = 10  # seconds
MAX_INPUT_CHARS = 30000
MAX_DOC_CHARS = 5000
TIMEOUT = 180  # seconds

# System prompt compacto — sem repetição, direto ao ponto
SYSTEM_PROMPT = """\
Analista sênior de crédito estruturado, ZYN Capital (SP). Produza MAC ZYN v3 em JSON.

Regras:
- Perspectiva de analista apresentando ao comitê, NÃO decisor
- Dedução 30% sobre garantias, cobertura mínima 130% LTV
- Stress tests obrigatórios (juros, câmbio, commodities)
- Ratings seção: Forte/Adequado/Atenção/Crítico
- Rating final: A/B/C/D/E | Parecer: Favorável/Favorável com Ressalvas/Desfavorável
- DSCR<1.2=atenção, <1.0=crítico | LTV>80%=atenção, >100%=crítico | Dív/EBITDA>3.5x=atenção, >5x=crítico
- Dados não disponíveis: "Não disponível" (texto) ou 0 (numérico) + flag
- Valores em R$. Responda SOMENTE JSON válido."""

# Template compacto — JSON schema inline sem comentários verbosos
ANALYSIS_PROMPT = """\
MAC ZYN v3 para: {tomador} | CNPJ: {cnpj} | {tipo_operacao} | R$ {volume:,.0f} | {prazo_meses}m | {taxa} | {amortizacao}
Garantias: {garantias} | Sócio: {socio_responsavel}

DADOS:
{dados_formatados}

Disponíveis: {docs_disponiveis} | Faltantes: {docs_indisponiveis}

IMPORTANTE: gere as chaves NESTA ORDEM EXATA no JSON (rating e kpis PRIMEIRO):

rating_final(nota,dscr_projetado,ltv_efetivo,parecer,justificativa,recomendacoes[]), \
kpis(receita_liquida,ebitda,margem_ebitda,divida_liquida_ebitda,ltv,dscr), \
tomador(razao_social,cnpj,grupo_economico,historico,rating_secao,flags), \
patrimonio(ativos_reais,avaliacao,ltv,analise,rating_secao,flags), \
producao(capacidade,historico_produtivo,analise,rating_secao,flags), \
capital(estrutura_capital,endividamento,indicadores{{divida_liquida_ebitda,divida_pl,liquidez_corrente,roe}},analise,rating_secao,flags), \
operacao(tipo,volume,prazo,taxa,estrutura,analise,rating_secao,flags), \
pagamento(fluxo_amortizacao,dscr,cobertura,analise,rating_secao,flags), \
onus(gravames[],certidoes,analise,rating_secao,flags), \
riscos(mercado,credito,operacional,legal,matriz_riscos[{{risco,probabilidade,impacto,mitigante}}],rating_secao,flags), \
covenants(clausulas[{{covenant,limite,atual,status}}],analise,rating_secao,flags), \
cronograma(etapas[{{etapa,prazo,responsavel}}],analise,rating_secao,flags), \
checklist_lacunas(documentos_faltantes[{{item,criticidade,motivo}}],informacoes_pendentes[{{item,criticidade,motivo}}],total_pendencias,total_criticas)

Seja conciso: 1-2 frases por analise de seção, max 3 flags por seção. JSON puro, sem markdown."""

ALL_DOC_TYPES = [
    "balanco", "dre", "fluxo_caixa", "balancete", "matricula",
    "contrato_social", "certidoes", "laudo_avaliacao", "relatorio_producao",
    "escritura", "ccir", "car", "ita", "procuracoes", "endividamento",
]


def _get_client() -> anthropic.Anthropic:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise EnvironmentError("ANTHROPIC_API_KEY não configurada.")
    return anthropic.Anthropic(api_key=api_key, timeout=TIMEOUT)


def _format_dados(dados: dict[str, Any]) -> str:
    if not dados:
        return "Nenhum dado disponível."
    blocos = []
    for doc_type, conteudo in dados.items():
        label = doc_type.replace("_", " ").title()
        if isinstance(conteudo, dict):
            texto = json.dumps(conteudo, ensure_ascii=False, separators=(",", ":"))
        else:
            texto = str(conteudo)
        if len(texto) > MAX_DOC_CHARS:
            texto = texto[:MAX_DOC_CHARS] + "...[truncado]"
        blocos.append(f"[{label}] {texto}")
    resultado = "\n".join(blocos)
    if len(resultado) > MAX_INPUT_CHARS:
        resultado = resultado[:MAX_INPUT_CHARS] + "\n...[truncado]"
    return resultado


def _list_docs(dados: dict[str, Any], available: bool) -> str:
    presentes = set(dados.keys()) if dados else set()
    tipos = sorted(presentes) if available else sorted(set(ALL_DOC_TYPES) - presentes)
    return ", ".join(t.replace("_", " ").title() for t in tipos) or "Nenhum"


def _repair_json(text: str) -> str:
    """Try to repair truncated JSON by closing open braces/brackets."""
    # Remove trailing comma before closing
    text = re.sub(r',\s*$', '', text.rstrip())
    # Count open/close braces and brackets
    open_braces = text.count('{') - text.count('}')
    open_brackets = text.count('[') - text.count(']')
    # Check if we're inside a string (odd number of unescaped quotes)
    in_string = (text.count('"') - text.count('\\"')) % 2 == 1
    if in_string:
        text += '"'
    # Close open brackets then braces
    text += ']' * max(0, open_brackets)
    text += '}' * max(0, open_braces)
    return text


def _parse_json(text: str) -> dict:
    text = text.strip()

    # Remove markdown code block wrapper
    match = re.search(r"```(?:json)?\s*\n?(.*?)(?:\n?\s*```|$)", text, re.DOTALL)
    if match:
        text = match.group(1).strip()

    # Extract from first { to last }
    first = text.find("{")
    last = text.rfind("}")
    if first != -1:
        if last != -1 and last > first:
            text = text[first:last + 1]
        else:
            # JSON was truncated — no closing brace
            text = text[first:]

    # Try direct parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Try repairing truncated JSON
    try:
        repaired = _repair_json(text)
        return json.loads(repaired)
    except json.JSONDecodeError:
        pass

    raise ValueError(f"JSON inválido. Primeiros 500 chars: {text[:500]}")


def analyze_credit(
    dados_extraidos: dict[str, Any],
    parametros_operacao: dict[str, Any],
    status_callback=None,
) -> dict:
    """Executa análise MAC ZYN v3 via Claude Sonnet com streaming."""
    client = _get_client()

    def _status(msg):
        if status_callback:
            status_callback(msg)
        logger.info(msg)

    garantias_str = ", ".join(parametros_operacao.get("garantias", []))
    dados_formatados = _format_dados(dados_extraidos)
    user_prompt = ANALYSIS_PROMPT.format(
        tomador=parametros_operacao.get("tomador", "N/I"),
        cnpj=parametros_operacao.get("cnpj", "N/I"),
        tipo_operacao=parametros_operacao.get("tipo_operacao", "N/I"),
        volume=parametros_operacao.get("volume", 0),
        prazo_meses=parametros_operacao.get("prazo_meses", 0),
        taxa=parametros_operacao.get("taxa", "N/I"),
        amortizacao=parametros_operacao.get("amortizacao", "N/I"),
        garantias=garantias_str or "N/I",
        socio_responsavel=parametros_operacao.get("socio_responsavel", "N/I"),
        dados_formatados=dados_formatados,
        docs_disponiveis=_list_docs(dados_extraidos, True),
        docs_indisponiveis=_list_docs(dados_extraidos, False),
    )

    total_chars = len(user_prompt) + len(SYSTEM_PROMPT)
    est_tokens = total_chars // 4
    _status(f"Enviando ~{est_tokens:,} tokens para {MODEL}...")

    # Streaming API call with retry
    response_text = ""
    for attempt in range(MAX_RETRIES):
        try:
            _status(f"Conectando... (tentativa {attempt + 1}/{MAX_RETRIES})")
            with client.messages.stream(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": user_prompt}],
            ) as stream:
                chunks = 0
                for text in stream.text_stream:
                    response_text += text
                    chunks += 1
                    if chunks % 20 == 0:
                        _status(f"Recebendo resposta... ({len(response_text):,} chars)")

            _status(f"Resposta completa: {len(response_text):,} chars. Processando...")
            break

        except anthropic.RateLimitError as e:
            if attempt == MAX_RETRIES - 1:
                raise
            _status(f"Rate limit. Aguardando {RETRY_WAIT}s...")
            time.sleep(RETRY_WAIT)
            response_text = ""

        except anthropic.APIStatusError as e:
            if e.status_code == 529 and attempt < MAX_RETRIES - 1:
                _status(f"API sobrecarregada. Aguardando {RETRY_WAIT}s...")
                time.sleep(RETRY_WAIT)
                response_text = ""
            else:
                raise

    analise = _parse_json(response_text)

    logger.info(
        "Análise concluída — Rating: %s | Parecer: %s",
        analise.get("rating_final", {}).get("nota", "N/A"),
        analise.get("rating_final", {}).get("parecer", "N/A"),
    )

    return analise


def get_analysis_summary(analise: dict) -> str:
    """Gera resumo legível da análise de crédito."""
    rating = analise.get("rating_final", {})
    kpis = analise.get("kpis", {})
    tomador = analise.get("tomador", {})
    capital = analise.get("capital", {})
    operacao = analise.get("operacao", {})

    nota = rating.get("nota", "N/A")
    parecer = rating.get("parecer", "N/A")
    justificativa = rating.get("justificativa", "—")
    recomendacoes = rating.get("recomendacoes", [])

    todas_flags = []
    secoes = ["tomador", "patrimonio", "producao", "capital", "operacao",
              "pagamento", "onus", "riscos", "covenants", "cronograma"]
    for s in secoes:
        for flag in analise.get(s, {}).get("flags", []):
            todas_flags.append(f"  [{s.title()}] {flag}")

    ratings_secao = [f"  {s.title()}: {analise.get(s, {}).get('rating_secao', 'N/A')}" for s in secoes]

    def _safe_pct(val, as_pct=False):
        if not isinstance(val, (int, float)):
            return "—"
        if as_pct and 0 < abs(val) <= 1:
            return f"{val:.1%}"
        return f"{val:.1f}%"

    linhas = [
        "=" * 60,
        "  MAC ZYN v3 — RESUMO EXECUTIVO",
        "=" * 60, "",
        f"Tomador:  {tomador.get('razao_social', 'N/A')}",
        f"CNPJ:     {tomador.get('cnpj', 'N/A')}",
        f"Operação: {operacao.get('tipo', 'N/A')} — R$ {operacao.get('volume', 0):,.0f}",
        f"Prazo:    {operacao.get('prazo', 'N/A')} | Taxa: {operacao.get('taxa', 'N/A')}",
        "", "-" * 60,
        f"  RATING: {nota}  |  PARECER: {parecer}",
        "-" * 60, "",
        "KPIs:",
        f"  Receita:    R$ {kpis.get('receita_liquida', 0):,.0f}",
        f"  EBITDA:     R$ {kpis.get('ebitda', 0):,.0f}",
        f"  Margem:     {_safe_pct(kpis.get('margem_ebitda', 0), True)}",
        f"  Dív/EBITDA: {kpis.get('divida_liquida_ebitda', 0):.2f}x",
        f"  LTV:        {_safe_pct(kpis.get('ltv', 0), True)}",
        f"  DSCR:       {kpis.get('dscr', 0):.2f}x",
        "", "Ratings:",
        *ratings_secao, "",
    ]

    if todas_flags:
        linhas.extend([f"Flags ({len(todas_flags)}):", *todas_flags, ""])

    linhas.extend([f"Justificativa: {justificativa}", ""])

    if recomendacoes:
        linhas.append("Recomendações:")
        for i, r in enumerate(recomendacoes, 1):
            linhas.append(f"  {i}. {r}")
        linhas.append("")

    checklist = analise.get("checklist_lacunas", {})
    docs_f = checklist.get("documentos_faltantes", [])
    info_p = checklist.get("informacoes_pendentes", [])
    if docs_f or info_p:
        linhas.extend(["-" * 60,
            f"  LACUNAS | Pendências: {checklist.get('total_pendencias', 0)} | Críticas: {checklist.get('total_criticas', 0)}",
            "-" * 60, ""])
        for d in docs_f:
            linhas.append(f"  [{d.get('criticidade', '?')}] {d.get('item', '?')}")
        for p in info_p:
            linhas.append(f"  [{p.get('criticidade', '?')}] {p.get('item', '?')}")
        linhas.append("")

    linhas.extend(["=" * 60, "  ZYN Capital — Crédito Estruturado & M&A", "=" * 60])
    return "\n".join(linhas)
