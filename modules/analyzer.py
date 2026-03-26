"""
ZYN Capital — Módulo de Análise de Crédito (MAC ZYN v3)

Utiliza Claude Opus API para realizar análise de crédito completa
baseada em dados financeiros extraídos de múltiplos documentos.
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
MAX_TOKENS = 8000
MAX_RETRIES = 3
RETRY_WAIT_SECONDS = [15, 30, 60]
MAX_INPUT_CHARS = 40000  # truncate extracted data to stay within token limits

SYSTEM_PROMPT = """\
Você é um analista sênior de crédito estruturado na ZYN Capital, boutique de crédito \
estruturado e M&A em São Paulo. Sua função é produzir o MAC (Memorando de Análise de \
Crédito) ZYN v3 completo, seguindo rigorosamente a estrutura de 10 seções obrigatórias.

POSTURA E TOM:
- Você escreve da perspectiva de um analista APRESENTANDO a oportunidade de crédito \
ao comitê, NÃO como quem toma a decisão final.
- Não use linguagem impositiva de aprovação (ex: "recomendo aprovação"). Em vez disso, \
apresente os fatos, riscos e mitigantes de forma objetiva para subsidiar a decisão.
- Sempre inclua uma visão geral do grupo econômico, cobrindo TODAS as atividades \
relevantes (inclusive atividade-fim como operações agrícolas, industriais etc.).

CONSERVADORISMO OBRIGATÓRIO:
- Aplique dedução mínima de 30% sobre TODAS as avaliações de garantias (especialmente \
operações rurais). Use valor de liquidação forçada.
- Exija cobertura mínima de 130% em operações estruturadas (LTV calculado sobre valor \
de venda forçada com dedução de 30%).
- Execute stress tests obrigatórios em cenários de alta volatilidade: taxa de juros, \
câmbio, preços de commodities, vacância imobiliária (conforme aplicável).
- Inclua margem de segurança em TODAS as projeções de fluxo de caixa.

ROBUSTEZ DA ANÁLISE:
- A análise DEVE cobrir TODAS as atividades do grupo econômico (não apenas a SPE ou \
veículo da operação).
- Apresente TODOS os indicadores relevantes — análises superficiais são inaceitáveis \
para decisão de crédito.
- Detalhe o endividamento por credor nomeado, com soma de vencimentos anuais.
- Análises superficiais ou incompletas são inaceitáveis.

CRITÉRIOS POR PRODUTO:
- Crédito Agrícola (CRA/CPR): Analise estrutura produtiva, aplique critérios de garantia \
(30% dedução + 130% cobertura), diferencie capacidade máxima de crédito vs. demanda \
solicitada.
- Viabilidade de Fundo (FIDC/Fiagro): Inclua diagnóstico (setor, governança, financeiro, \
dívida, reputação), lastro dos ativos, estrutura de cotas (sênior/mezanino/subordinada), \
taxas de referência (benchmark).
- Sale-Leaseback (SLB): Detalhe valor da operação, prazo total, carência, estrutura de \
pagamento, prêmio de recompra.
- Mercado de Capitais (CRA, Debêntures): Hierarquia de fontes (CVM/ANBIMA primeiro), \
dados de benchmark, insights práticos.

MODELO DE SCORING — considere obrigatoriamente:
- Capacidade de pagamento: geração de caixa vs. serviço da dívida
- Qualidade das garantias: LTV ajustado (130% sobre venda forçada)
- Histórico e reputação: score de compliance, marcadores restritivos
- Resiliência setorial: sensibilidade a choques macroeconômicos

Diretrizes gerais:
- Analise exclusivamente com base nos dados fornecidos. Não invente dados.
- Quando uma informação não estiver disponível, indique "Não disponível" e registre uma flag.
- Seja conservador nas avaliações — prefira cautela a otimismo.
- Ratings de seção: "Forte", "Adequado", "Atenção" ou "Crítico".
- Rating final (nota): A (excelente), B (bom), C (aceitável), D (elevado), E (proibitivo).
- Parecer: "Favorável", "Favorável com Ressalvas" ou "Desfavorável".
- DSCR < 1.2 é flag de atenção; < 1.0 é flag crítico.
- LTV > 80% (sobre valor de venda forçada com dedução de 30%) é flag de atenção; > 100% é flag crítico.
- Dívida Líquida/EBITDA > 3.5x é flag de atenção; > 5x é flag crítico.
- Flags são strings curtas descrevendo riscos ou pontos de atenção.
- Todos os valores monetários em R$ (reais).
- Ao final, gere SEMPRE um "Checklist de Lacunas" (Q&A de Lacunas) listando documentos \
faltantes e informações pendentes, com criticidade e motivo.
- Responda SOMENTE com JSON válido. Nenhum texto antes ou depois do JSON.
"""

ANALYSIS_PROMPT_TEMPLATE = """\
Produza o MAC ZYN v3 completo para a operação abaixo.

## Parâmetros da Operação
- Tomador: {tomador}
- CNPJ: {cnpj}
- Tipo: {tipo_operacao}
- Volume: R$ {volume:,.2f}
- Prazo: {prazo_meses} meses
- Taxa: {taxa}
- Amortização: {amortizacao}
- Garantias: {garantias}
- Sócio responsável: {socio_responsavel}

## Dados Extraídos dos Documentos
{dados_formatados}

## Documentos Disponíveis
{docs_disponiveis}

## Documentos NÃO Disponíveis
{docs_indisponiveis}

## Estrutura de Resposta (JSON)
Retorne EXCLUSIVAMENTE um JSON válido com a seguinte estrutura. Preencha todos os campos \
com base nos dados disponíveis. Para dados não disponíveis, use "Não disponível" em campos \
texto, 0 em campos numéricos, e registre a ausência como flag.

{{
    "tomador": {{
        "razao_social": "string",
        "cnpj": "string",
        "grupo_economico": "string — descreva o grupo econômico se identificável",
        "historico": "string — histórico relevante do tomador",
        "rating_secao": "Forte|Adequado|Atenção|Crítico",
        "flags": ["lista de flags/alertas"]
    }},
    "patrimonio": {{
        "ativos_reais": "string — descrição dos ativos reais identificados",
        "avaliacao": "string — avaliação dos ativos/garantias",
        "ltv": 0.0,
        "analise": "string — análise patrimonial detalhada",
        "rating_secao": "Forte|Adequado|Atenção|Crítico",
        "flags": []
    }},
    "producao": {{
        "capacidade": "string — capacidade produtiva se aplicável",
        "historico_produtivo": "string — histórico de produção",
        "analise": "string — análise da capacidade produtiva/operacional",
        "rating_secao": "Forte|Adequado|Atenção|Crítico",
        "flags": []
    }},
    "capital": {{
        "estrutura_capital": "string — composição do capital",
        "endividamento": "string — perfil de endividamento",
        "indicadores": {{
            "divida_liquida_ebitda": 0.0,
            "divida_pl": 0.0,
            "liquidez_corrente": 0.0,
            "roe": 0.0
        }},
        "analise": "string — análise detalhada da estrutura de capital",
        "rating_secao": "Forte|Adequado|Atenção|Crítico",
        "flags": []
    }},
    "operacao": {{
        "tipo": "string",
        "volume": 0,
        "prazo": "string",
        "taxa": "string",
        "estrutura": "string — descrição da estrutura proposta",
        "analise": "string — análise da adequação da operação",
        "rating_secao": "Forte|Adequado|Atenção|Crítico",
        "flags": []
    }},
    "pagamento": {{
        "fluxo_amortizacao": "string — descrição do fluxo de amortização",
        "dscr": 0.0,
        "cobertura": "string — análise de cobertura",
        "analise": "string — análise da capacidade de pagamento",
        "rating_secao": "Forte|Adequado|Atenção|Crítico",
        "flags": []
    }},
    "onus": {{
        "gravames": ["lista de gravames identificados"],
        "certidoes": "string — situação das certidões",
        "analise": "string — análise de ônus e restrições",
        "rating_secao": "Forte|Adequado|Atenção|Crítico",
        "flags": []
    }},
    "riscos": {{
        "mercado": "string — riscos de mercado",
        "credito": "string — riscos de crédito",
        "operacional": "string — riscos operacionais",
        "legal": "string — riscos legais/regulatórios",
        "matriz_riscos": [
            {{"risco": "string", "probabilidade": "Alta|Média|Baixa", "impacto": "Alto|Médio|Baixo", "mitigante": "string"}}
        ],
        "rating_secao": "Forte|Adequado|Atenção|Crítico",
        "flags": []
    }},
    "covenants": {{
        "clausulas": [
            {{"covenant": "string", "limite": "string", "atual": "string", "status": "Cumprido|Descumprido|N/A"}}
        ],
        "analise": "string — análise dos covenants sugeridos",
        "rating_secao": "Forte|Adequado|Atenção|Crítico",
        "flags": []
    }},
    "cronograma": {{
        "etapas": [
            {{"etapa": "string", "prazo": "string", "responsavel": "string"}}
        ],
        "analise": "string — análise do cronograma de execução",
        "rating_secao": "Forte|Adequado|Atenção|Crítico",
        "flags": []
    }},
    "rating_final": {{
        "nota": "A|B|C|D|E",
        "dscr_projetado": 0.0,
        "ltv_efetivo": 0.0,
        "parecer": "Favorável|Favorável com Ressalvas|Desfavorável",
        "justificativa": "string — justificativa completa do parecer, escrita como apresentação ao comitê (não como decisão final)",
        "recomendacoes": ["lista de recomendações e condições precedentes"]
    }},
    "kpis": {{
        "receita_liquida": 0,
        "ebitda": 0,
        "margem_ebitda": 0.0,
        "divida_liquida_ebitda": 0.0,
        "ltv": 0.0,
        "dscr": 0.0
    }},
    "checklist_lacunas": {{
        "documentos_faltantes": [
            {{"item": "string — nome do documento faltante", "criticidade": "ALTA|MÉDIA|BAIXA", "motivo": "string — por que esse documento é necessário"}}
        ],
        "informacoes_pendentes": [
            {{"item": "string — informação pendente", "criticidade": "ALTA|MÉDIA|BAIXA", "motivo": "string — por que essa informação é necessária"}}
        ],
        "total_pendencias": 0,
        "total_criticas": 0
    }}
}}

INSTRUÇÕES FINAIS:
- Analise com rigor e conservadorismo. Flags devem destacar TODOS os pontos de atenção encontrados.
- Não omita riscos. Aplique dedução de 30% nas garantias e exija cobertura de 130%.
- Inclua stress tests quando aplicável (juros, câmbio, commodities, vacância).
- Detalhe endividamento por credor nomeado com vencimentos anuais.
- O checklist_lacunas DEVE listar TODOS os documentos e informações ausentes ou insuficientes, \
com criticidade (ALTA/MÉDIA/BAIXA) e motivo. total_pendencias = soma de todos os itens. \
total_criticas = soma dos itens com criticidade ALTA.
- Escreva como analista apresentando a oportunidade, não como decisor.
- Responda SOMENTE com o JSON acima, sem texto adicional.
"""

ALL_DOC_TYPES = [
    "balanco",
    "dre",
    "fluxo_caixa",
    "balancete",
    "matricula",
    "contrato_social",
    "certidoes",
    "laudo_avaliacao",
    "relatorio_producao",
    "escritura",
    "ccir",
    "car",
    "ita",
    "procuracoes",
    "endividamento",
]


def _get_client() -> anthropic.Anthropic:
    """Retorna instância autenticada do client Anthropic."""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise EnvironmentError(
            "Variável de ambiente ANTHROPIC_API_KEY não configurada. "
            "Defina-a antes de executar a análise."
        )
    return anthropic.Anthropic(api_key=api_key)


def _format_dados_extraidos(dados_extraidos: dict[str, Any]) -> str:
    """Formata os dados extraídos em texto estruturado para o prompt.

    Trunca o resultado total para MAX_INPUT_CHARS para não estourar rate limits.
    """
    if not dados_extraidos:
        return "Nenhum dado extraído disponível."

    blocos: list[str] = []
    for doc_type, conteudo in dados_extraidos.items():
        label = doc_type.replace("_", " ").title()
        blocos.append(f"### {label}")
        if isinstance(conteudo, dict):
            texto = json.dumps(conteudo, ensure_ascii=False, indent=2)
        elif isinstance(conteudo, str):
            texto = conteudo
        else:
            texto = str(conteudo)
        # Truncate individual document data if too large
        if len(texto) > 8000:
            texto = texto[:8000] + "\n... [dados truncados por limite de tokens]"
        blocos.append(texto)
        blocos.append("")

    resultado = "\n".join(blocos)
    if len(resultado) > MAX_INPUT_CHARS:
        resultado = resultado[:MAX_INPUT_CHARS] + "\n... [dados truncados por limite de tokens]"
    return resultado


def _list_docs(dados_extraidos: dict[str, Any], available: bool) -> str:
    """Lista documentos disponíveis ou indisponíveis."""
    presentes = set(dados_extraidos.keys()) if dados_extraidos else set()
    if available:
        tipos = sorted(presentes) if presentes else []
    else:
        tipos = sorted(set(ALL_DOC_TYPES) - presentes)

    if not tipos:
        return "Nenhum" if not available else "Todos os documentos padrão"

    return ", ".join(t.replace("_", " ").title() for t in tipos)


def _parse_json_response(text: str) -> dict:
    """
    Extrai e parseia JSON da resposta do modelo.

    Tenta múltiplas estratégias:
    1. Parse direto do texto completo
    2. Extração de bloco ```json ... ```
    3. Busca pelo primeiro '{' e último '}'
    """
    text = text.strip()

    # Estratégia 1: parse direto
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Estratégia 2: bloco markdown json
    match = re.search(r"```(?:json)?\s*\n?(.*?)\n?\s*```", text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1).strip())
        except json.JSONDecodeError:
            pass

    # Estratégia 3: primeiro '{' ao último '}'
    first_brace = text.find("{")
    last_brace = text.rfind("}")
    if first_brace != -1 and last_brace != -1 and last_brace > first_brace:
        candidate = text[first_brace : last_brace + 1]
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            pass

    raise ValueError(
        f"Não foi possível extrair JSON válido da resposta do modelo. "
        f"Primeiros 500 caracteres: {text[:500]}"
    )


def analyze_credit(
    dados_extraidos: dict[str, Any],
    parametros_operacao: dict[str, Any],
    status_callback=None,
) -> dict:
    """
    Executa análise de crédito completa (MAC ZYN v3) via Claude Opus.

    Args:
        dados_extraidos: Dicionário com dados extraídos de documentos,
            chaveado por tipo de documento (balanco, dre, matricula, etc.).
        parametros_operacao: Parâmetros da operação contendo tomador, cnpj,
            tipo_operacao, volume, prazo_meses, taxa, amortizacao, garantias
            e socio_responsavel.

    Returns:
        Dicionário com a análise MAC completa em 10 seções + rating_final + kpis.

    Raises:
        EnvironmentError: Se ANTHROPIC_API_KEY não estiver configurada.
        ValueError: Se a resposta do modelo não contiver JSON válido.
        anthropic.APIError: Se houver erro na chamada à API.
    """
    client = _get_client()

    garantias_str = ", ".join(parametros_operacao.get("garantias", []))
    dados_formatados = _format_dados_extraidos(dados_extraidos)
    docs_disponiveis = _list_docs(dados_extraidos, available=True)
    docs_indisponiveis = _list_docs(dados_extraidos, available=False)

    user_prompt = ANALYSIS_PROMPT_TEMPLATE.format(
        tomador=parametros_operacao.get("tomador", "Não informado"),
        cnpj=parametros_operacao.get("cnpj", "Não informado"),
        tipo_operacao=parametros_operacao.get("tipo_operacao", "Não informado"),
        volume=parametros_operacao.get("volume", 0),
        prazo_meses=parametros_operacao.get("prazo_meses", 0),
        taxa=parametros_operacao.get("taxa", "Não informada"),
        amortizacao=parametros_operacao.get("amortizacao", "Não informada"),
        garantias=garantias_str or "Não informadas",
        socio_responsavel=parametros_operacao.get("socio_responsavel", "Não informado"),
        dados_formatados=dados_formatados,
        docs_disponiveis=docs_disponiveis,
        docs_indisponiveis=docs_indisponiveis,
    )

    logger.info(
        "Iniciando análise de crédito para %s (%s) — %s R$ %s | modelo=%s",
        parametros_operacao.get("tomador"),
        parametros_operacao.get("cnpj"),
        parametros_operacao.get("tipo_operacao"),
        f"{parametros_operacao.get('volume', 0):,.2f}",
        MODEL,
    )

    def _status(msg: str):
        if status_callback:
            status_callback(msg)
        logger.info(msg)

    prompt_chars = len(user_prompt) + len(SYSTEM_PROMPT)
    _status(f"Enviando {prompt_chars:,} caracteres para {MODEL}...")

    # API call with retry logic for rate limits (429)
    message = None
    for attempt in range(MAX_RETRIES):
        try:
            _status(f"Chamando API (tentativa {attempt + 1}/{MAX_RETRIES})...")
            message = client.messages.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": user_prompt}],
            )
            break
        except anthropic.RateLimitError as e:
            wait = RETRY_WAIT_SECONDS[attempt] if attempt < len(RETRY_WAIT_SECONDS) else 60
            _status(f"Rate limit atingido. Aguardando {wait}s antes de tentar novamente...")
            logger.warning(
                "Rate limit (429) na tentativa %d/%d. Aguardando %ds... (%s)",
                attempt + 1, MAX_RETRIES, wait, e,
            )
            if attempt == MAX_RETRIES - 1:
                raise
            time.sleep(wait)
        except anthropic.APIStatusError as e:
            if e.status_code == 529:  # overloaded
                wait = RETRY_WAIT_SECONDS[attempt] if attempt < len(RETRY_WAIT_SECONDS) else 60
                _status(f"API sobrecarregada. Aguardando {wait}s...")
                if attempt == MAX_RETRIES - 1:
                    raise
                time.sleep(wait)
            else:
                raise

    _status("Resposta recebida. Processando JSON...")

    response_text = message.content[0].text
    analise = _parse_json_response(response_text)

    logger.info(
        "Análise concluída — Rating: %s | Parecer: %s",
        analise.get("rating_final", {}).get("nota", "N/A"),
        analise.get("rating_final", {}).get("parecer", "N/A"),
    )

    return analise


def get_analysis_summary(analise: dict) -> str:
    """
    Gera resumo legível da análise de crédito.

    Args:
        analise: Dicionário com a análise MAC completa retornada por analyze_credit().

    Returns:
        Texto formatado com resumo executivo da análise.
    """
    rating = analise.get("rating_final", {})
    kpis = analise.get("kpis", {})
    tomador = analise.get("tomador", {})
    capital = analise.get("capital", {})
    operacao = analise.get("operacao", {})

    nota = rating.get("nota", "N/A")
    parecer = rating.get("parecer", "N/A")
    justificativa = rating.get("justificativa", "Sem justificativa disponível.")
    recomendacoes = rating.get("recomendacoes", [])

    # Coleta todas as flags de todas as seções
    todas_flags: list[str] = []
    secoes_com_flags = [
        "tomador", "patrimonio", "producao", "capital", "operacao",
        "pagamento", "onus", "riscos", "covenants", "cronograma",
    ]
    for secao in secoes_com_flags:
        dados_secao = analise.get(secao, {})
        flags = dados_secao.get("flags", [])
        secao_label = secao.replace("_", " ").title()
        for flag in flags:
            todas_flags.append(f"  [{secao_label}] {flag}")

    # Ratings por seção
    ratings_secao: list[str] = []
    for secao in secoes_com_flags:
        dados_secao = analise.get(secao, {})
        r = dados_secao.get("rating_secao", "N/A")
        secao_label = secao.replace("_", " ").title()
        ratings_secao.append(f"  {secao_label}: {r}")

    # Monta o resumo
    linhas = [
        "=" * 70,
        "  MAC ZYN v3 — RESUMO EXECUTIVO",
        "=" * 70,
        "",
        f"Tomador:    {tomador.get('razao_social', 'N/A')}",
        f"CNPJ:       {tomador.get('cnpj', 'N/A')}",
        f"Operação:   {operacao.get('tipo', 'N/A')} — R$ {operacao.get('volume', 0):,.2f}",
        f"Prazo:      {operacao.get('prazo', 'N/A')}",
        f"Taxa:       {operacao.get('taxa', 'N/A')}",
        "",
        "-" * 70,
        f"  RATING FINAL: {nota}    |    PARECER: {parecer}",
        "-" * 70,
        "",
        "KPIs:",
        f"  Receita Líquida:      R$ {kpis.get('receita_liquida', 0):,.2f}",
        f"  EBITDA:               R$ {kpis.get('ebitda', 0):,.2f}",
        f"  Margem EBITDA:        {kpis.get('margem_ebitda', 0):.1%}" if isinstance(kpis.get("margem_ebitda"), (int, float)) and kpis.get("margem_ebitda", 0) <= 1 else f"  Margem EBITDA:        {kpis.get('margem_ebitda', 0):.1f}%",
        f"  Dív. Líq./EBITDA:     {kpis.get('divida_liquida_ebitda', 0):.2f}x",
        f"  LTV:                  {kpis.get('ltv', 0):.1%}" if isinstance(kpis.get("ltv"), (int, float)) and kpis.get("ltv", 0) <= 1 else f"  LTV:                  {kpis.get('ltv', 0):.1f}%",
        f"  DSCR:                 {kpis.get('dscr', 0):.2f}x",
        "",
        "Indicadores de Capital:",
        f"  Dív. Líq./EBITDA:     {capital.get('indicadores', {}).get('divida_liquida_ebitda', 0):.2f}x",
        f"  Dívida/PL:            {capital.get('indicadores', {}).get('divida_pl', 0):.2f}x",
        f"  Liquidez Corrente:    {capital.get('indicadores', {}).get('liquidez_corrente', 0):.2f}",
        f"  ROE:                  {capital.get('indicadores', {}).get('roe', 0):.1%}" if isinstance(capital.get("indicadores", {}).get("roe"), (int, float)) and capital.get("indicadores", {}).get("roe", 0) <= 1 else f"  ROE:                  {capital.get('indicadores', {}).get('roe', 0):.1f}%",
        "",
        "Ratings por Seção:",
        *ratings_secao,
        "",
    ]

    if todas_flags:
        linhas.extend([
            f"Flags de Atenção ({len(todas_flags)}):",
            *todas_flags,
            "",
        ])
    else:
        linhas.extend(["Flags de Atenção: Nenhuma", ""])

    linhas.extend([
        "Justificativa:",
        f"  {justificativa}",
        "",
    ])

    if recomendacoes:
        linhas.append("Recomendações:")
        for i, rec in enumerate(recomendacoes, 1):
            linhas.append(f"  {i}. {rec}")
        linhas.append("")

    # Checklist de Lacunas (Q&A de Lacunas)
    checklist = analise.get("checklist_lacunas", {})
    docs_faltantes = checklist.get("documentos_faltantes", [])
    info_pendentes = checklist.get("informacoes_pendentes", [])
    total_pendencias = checklist.get("total_pendencias", len(docs_faltantes) + len(info_pendentes))
    total_criticas = checklist.get("total_criticas", 0)

    if docs_faltantes or info_pendentes:
        linhas.extend([
            "-" * 70,
            f"  CHECKLIST DE LACUNAS  |  Pendências: {total_pendencias}  |  Críticas: {total_criticas}",
            "-" * 70,
            "",
        ])

        if docs_faltantes:
            linhas.append("Documentos Faltantes:")
            for doc in docs_faltantes:
                crit = doc.get("criticidade", "N/A")
                item = doc.get("item", "N/A")
                motivo = doc.get("motivo", "")
                linhas.append(f"  [{crit}] {item}")
                if motivo:
                    linhas.append(f"         {motivo}")
            linhas.append("")

        if info_pendentes:
            linhas.append("Informações Pendentes:")
            for info in info_pendentes:
                crit = info.get("criticidade", "N/A")
                item = info.get("item", "N/A")
                motivo = info.get("motivo", "")
                linhas.append(f"  [{crit}] {item}")
                if motivo:
                    linhas.append(f"         {motivo}")
            linhas.append("")

    linhas.extend(["=" * 70, "  ZYN Capital — Crédito Estruturado & M&A", "=" * 70])

    return "\n".join(linhas)
