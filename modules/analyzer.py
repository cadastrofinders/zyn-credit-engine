"""
ZYN Capital — Módulo de Análise de Crédito (MAC ZYN v4)
Usa Claude Opus com streaming para análise profunda e assertiva.

v4 — Melhorias ZYN-específicas:
  - Rating scale alinhada ao Pipeline: AAA/AA/A/BBB/BB/B/C/D
  - Análise setorial calibrada (Agro, Imobiliário, Indústria, FIDC, Special Sits)
  - Benchmarks setoriais do mercado brasileiro
  - Taxonomia de garantias do Pipeline (Real, Fiduciária, Aval, Coobrigação, Alienação, Penhor)
  - Investor matching integrado
"""

import json
import logging
import os
import re
import time
from typing import Any

import anthropic

from modules.investor_matching import get_sector_benchmarks, _detect_sector, match_investors

logger = logging.getLogger(__name__)

MODEL = "claude-opus-4-6"
MAX_TOKENS = 16000
MAX_RETRIES = 2
RETRY_WAIT = 10  # seconds
MAX_INPUT_CHARS = 180000  # Opus suporta 200K — usar a maior parte
MAX_DOC_CHARS = 15000  # Permitir demonstrativos financeiros completos
TIMEOUT = 300  # seconds — Opus pode demorar mais para análises profundas

# Prioridade de documentos para análise de crédito
DOC_PRIORITY = {
    # CRITICAL — núcleo financeiro (processar primeiro, dados completos)
    "balanco": 1, "dre": 1, "demonstracoes_financeiras": 1, "balancete": 1,
    # HIGH — garantias e endividamento
    "matricula": 2, "laudo_avaliacao": 2, "endividamento": 2,
    # MEDIUM — legal, receita, SCR
    "contrato": 3, "contrato_social": 3, "certidao": 3, "certidoes": 3,
    "faturamento": 3, "scr": 3,
    # LOW — cadastro e complementares
    "cnpj": 4, "kyc_publico": 4, "planejamento": 4,
    "alteracao_contratual": 4, "ccir_car": 4, "ccir": 4, "car": 4,
    "outro": 4, "ita": 4, "procuracoes": 4, "escritura": 4,
    "relatorio_producao": 4, "fluxo_caixa": 4,
}

# System prompt otimizado para Opus — instruções sofisticadas de análise setorial
SYSTEM_PROMPT = """\
Analista sênior de crédito estruturado, ZYN Capital (SP). Produza MAC ZYN v4 em JSON.

Regras fundamentais:
- Perspectiva de analista apresentando ao comitê, NÃO decisor
- Dedução 30% sobre garantias, cobertura mínima 130% LTV
- Ratings seção: Forte/Adequado/Atenção/Crítico
- Rating final: AAA/AA/A/BBB/BB/B/C/D (escala alinhada ao Pipeline ZYN)
  AAA = risco mínimo, lastro/garantias excepcionais, empresa grau de investimento
  AA = risco muito baixo, garantias sólidas, financeiro robusto
  A = risco baixo, estrutura adequada, métricas dentro dos parâmetros
  BBB = risco moderado, aceitável com ressalvas, garantias suficientes
  BB = risco elevado, necessita reforço de garantias ou covenants restritivos
  B = risco alto, operação viável apenas com mitigantes significativos
  C = risco muito alto, não recomendável sem reestruturação
  D = inviável / default iminente
- Parecer: Favorável / Favorável com Ressalvas / Desfavorável
- DSCR<1.2=atenção, <1.0=crítico | LTV>80%=atenção, >100%=crítico | Dív/EBITDA>3.5x=atenção, >5x=crítico
- Dados não disponíveis: "Não disponível" (texto) ou 0 (numérico) + flag
- Valores em R$. Responda SOMENTE JSON válido.

{setor_instrucoes}

Cross-referencing obrigatório:
- Cruze PL do Balanço com lucro líquido acumulado da DRE — flagge divergências >5%.
- Cruze receita do DRE com faturamento declarado — flagge divergências >10%.
- Cruze endividamento declarado com passivos do Balanço — identifique dívidas omitidas.
- Cruze dados do SCR (Banco Central) com endividamento declarado — flagge discrepâncias.
- Verifique se garantias reais (matrícula/laudo) são compatíveis com valor da operação.
- Se houver dados KYC públicos, valide razão social, CNPJ, QSA, atividade econômica e capital social. Flagge divergências.
- IMPORTANTE: A equipe da ZYN Capital (assessora financeira) inclui: Danilo Salasar, Luiz Roberto Evangelista, Renato Barison. Se esses nomes aparecerem nos documentos, são representantes da assessora, NÃO sócios/signatários da empresa analisada. Não flagge como divergência de QSA.

Benchmarking setorial (use os benchmarks fornecidos abaixo):
{benchmarks_texto}

Qualidade e confiabilidade dos documentos:
- Classifique cada documento como: Auditado / Não-auditado / Gerencial / Autodeclarado.
- Flagge documentos desatualizados (>12 meses), incompletos ou com formatação suspeita.
- Atribua peso maior a dados auditados e menor a autodeclarados na análise.
- Identifique explicitamente quais conclusões dependem de dados de baixa confiabilidade.

Identificação de inconsistências:
- Liste TODAS as inconsistências encontradas entre documentos, com gravidade (Alta/Média/Baixa).
- Para cada inconsistência, indique os documentos conflitantes e os valores divergentes.
- Avalie se inconsistências sugerem manipulação, erro contábil ou defasagem temporal.

Garantias — classifique cada garantia na taxonomia: Real / Fiduciária / Aval / Coobrigação / Alienação / Penhor.
Para cada garantia, avalie: tipo, valor estimado, LTV individual, e adequação ao instrumento.

Stress tests obrigatórios (com premissas explícitas):
- Cenário base: premissas atuais.
- Cenário estresse 1 (juros): CDI +300bps — recalcule DSCR e capacidade de pagamento.
- Cenário estresse 2 (receita): queda de 20% na receita — recalcule DSCR e covenants.
- Cenário estresse 3 (combinado): CDI +200bps + queda 15% receita + desvalorização 20% garantias.
{stress_setor}
- Para cada cenário: indique se a operação permanece viável e quais covenants seriam rompidos."""

# Instruções setoriais específicas
SETOR_INSTRUCOES = {
    "agro": """Análise setorial — AGRONEGÓCIO:
- Avalie capacidade produtiva: área plantada, produtividade histórica (3 safras), mix de culturas.
- Considere preços de commodities (soja, milho, algodão, boi) e tendência de mercado.
- Avalie riscos climáticos: região, histórico de sinistros, cobertura de seguro agrícola/PROAGRO.
- Analise estrutura de comercialização: % pré-vendido, contratos com tradings, hedge.
- Para SLB: avalie valor de mercado do(s) imóvel(is) vs valor contratual, prazo do arrendamento.
- Custo de produção: insumos, arrendamento, mão de obra, logística.
- Sazonalidade do fluxo de caixa: safra vs entressafra.""",

    "imobiliario": """Análise setorial — IMOBILIÁRIO:
- Avalie o empreendimento: VSO (velocidade sobre oferta), VGV, percentual de obras concluído.
- Analise estoque: unidades lançadas vs vendidas vs em estoque, ticket médio.
- Considere localização: cidade, bairro, demanda regional, preço por m².
- Para Compra de Estoque: avalie desconto sobre VGV, velocidade de giro projetada.
- Perfil do comprador final: renda, financiamento bancário, inadimplência histórica.
- Risco de obra: construtora, cronograma, habite-se, custos de construção.
- Recebíveis imobiliários: aging, diluição, performance histórica.""",

    "industria": """Análise setorial — INDÚSTRIA/SERVIÇOS:
- Avalie utilização de capacidade instalada e perspectiva de demanda.
- Analise concentração de clientes: top 5 clientes como % da receita.
- Considere posição competitiva: market share, barreiras de entrada, vantagens.
- Para NC/CCB: avalie finalidade do crédito (capital de giro, investimento, refinanciamento).
- Ciclo de conversão de caixa: prazo médio de recebimento, pagamento, estocagem.
- Risco operacional: dependência de fornecedores, mão de obra especializada.""",

    "fidc": """Análise setorial — FIDC:
- Avalie qualidade da carteira: aging, PDD, inadimplência > 30/60/90 dias.
- Concentração de sacados: top 10 sacados como % do PL.
- Elegibilidade do lastro: critérios de cessão, rebate, diluição.
- Estrutura de subordinação: % cota subordinada, excesso de spread, triggers.
- Performance histórica: yield líquido, taxa de prepagamento, default rate.
- Risco do cedente: coobrigação, saúde financeira, diversificação.""",

    "special_sits": """Análise setorial — SPECIAL SITUATIONS:
- Avalie natureza do ativo: precatório, NPL, crédito judicial, equity.
- Timing de recebimento: prazo estimado, riscos de atraso, recursal.
- Desconto sobre valor de face e TIR implícita.
- Risco jurídico: complexidade do caso, jurisprudência, instância atual.
- Liquidez do ativo: mercado secundário, possibilidade de cessão.""",
}

# Stress tests adicionais por setor
STRESS_SETOR = {
    "agro": "- Cenário estresse 4 (safra): quebra de 30% na safra + queda 15% no preço da commodity — recalcule DSCR.",
    "imobiliario": "- Cenário estresse 4 (vendas): queda de 40% no VSO + aumento de 10% no custo de obra — recalcule viabilidade.",
    "industria": "- Cenário estresse 4 (demanda): queda de 25% no volume + perda do principal cliente — recalcule DSCR.",
    "fidc": "- Cenário estresse 4 (carteira): inadimplência sobe para 2x a média histórica + 30% de diluição — recalcule subordinação.",
    "special_sits": "- Cenário estresse 4 (prazo): atraso de 24 meses no recebimento — recalcule TIR e viabilidade.",
}

# Template compacto — JSON schema inline sem comentários verbosos
ANALYSIS_PROMPT = """\
MAC ZYN v4 para: {tomador} | CNPJ: {cnpj} | {tipo_operacao} | R$ {volume:,.0f} | {prazo_meses}m | {taxa} | {amortizacao}
Garantias: {garantias} | Sócio: {socio_responsavel} | Setor: {setor}

DADOS (ordenados por relevância — financeiros primeiro):
{dados_formatados}

Disponíveis: {docs_disponiveis} | Faltantes: {docs_indisponiveis}

IMPORTANTE: gere as chaves NESTA ORDEM EXATA no JSON (rating e kpis PRIMEIRO):
Use escala de rating: AAA/AA/A/BBB/BB/B/C/D (NÃO use A-E).

rating_final(nota,dscr_projetado,ltv_efetivo,parecer,justificativa,recomendacoes[]), \
kpis(receita_liquida,ebitda,margem_ebitda,divida_liquida_ebitda,ltv,dscr), \
tomador(razao_social,cnpj,grupo_economico,historico,setor,rating_secao,flags), \
patrimonio(ativos_reais,avaliacao,ltv,analise,garantias_detalhadas[{{tipo_garantia,descricao,valor_estimado,ltv_individual,classificacao}}],rating_secao,flags), \
producao(capacidade,historico_produtivo,analise,rating_secao,flags), \
capital(estrutura_capital,endividamento,indicadores{{divida_liquida_ebitda,divida_pl,liquidez_corrente,roe}},analise,rating_secao,flags), \
operacao(tipo,instrumento,volume,prazo,taxa,estrutura,analise,rating_secao,flags), \
pagamento(fluxo_amortizacao,dscr,cobertura,stress_tests{{cenario_base,estresse_juros,estresse_receita,estresse_combinado,estresse_setorial}},analise,rating_secao,flags), \
onus(gravames[],certidoes,analise,rating_secao,flags), \
riscos(mercado,credito,operacional,legal,setorial,matriz_riscos[{{risco,probabilidade,impacto,mitigante}}],rating_secao,flags), \
covenants(clausulas[{{covenant,limite,atual,status}}],analise,rating_secao,flags), \
cronograma(etapas[{{etapa,prazo,responsavel}}],analise,rating_secao,flags), \
cross_referencing(inconsistencias[{{documentos,campo,valores,gravidade,comentario}}],qualidade_docs[{{documento,classificacao,observacao}}]), \
benchmarking(setor,metricas_vs_setor[{{metrica,valor_empresa,media_setor,posicao}}],contexto_macro,riscos_setoriais[]), \
checklist_lacunas(documentos_faltantes[{{item,criticidade,motivo}}],informacoes_pendentes[{{item,criticidade,motivo}}],total_pendencias,total_criticas)

Seja analítico e preciso. Cruze dados entre documentos. JSON puro, sem markdown."""

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
    """Formata dados priorizando documentos críticos para análise de crédito."""
    if not dados:
        return "Nenhum dado disponível."

    # Ordenar por prioridade (CRITICAL=1 primeiro, LOW=4 por último)
    sorted_docs = sorted(
        dados.items(),
        key=lambda x: DOC_PRIORITY.get(x[0], 4),
    )

    blocos = []
    for doc_type, conteudo in sorted_docs:
        priority = DOC_PRIORITY.get(doc_type, 4)
        priority_labels = {1: "CRITICAL", 2: "HIGH", 3: "MEDIUM", 4: "LOW"}
        label = doc_type.replace("_", " ").title()
        tag = priority_labels.get(priority, "LOW")

        if isinstance(conteudo, dict):
            texto = json.dumps(conteudo, ensure_ascii=False, separators=(",", ":"))
        else:
            texto = str(conteudo)

        # Truncar documentos individuais apenas se excederem MAX_DOC_CHARS
        if len(texto) > MAX_DOC_CHARS:
            texto = texto[:MAX_DOC_CHARS] + "...[truncado]"

        blocos.append(f"[{tag}:{label}] {texto}")

    resultado = "\n".join(blocos)

    # Truncar total apenas se exceder MAX_INPUT_CHARS (improvável com 180K)
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


def _build_sector_prompt(tipo_operacao: str, garantias: list[str] | None = None) -> tuple[str, str, str, str]:
    """Build sector-specific prompt sections. Returns (setor, instrucoes, benchmarks_texto, stress_setor)."""
    setor = _detect_sector(tipo_operacao, garantias)
    instrucoes = SETOR_INSTRUCOES.get(setor, SETOR_INSTRUCOES["industria"])
    stress = STRESS_SETOR.get(setor, "")

    # Build benchmarks text
    bench = get_sector_benchmarks(setor)
    bench_lines = [
        f"Setor: {bench['setor']}",
        f"- Margem EBITDA média: {bench['margem_ebitda_media']:.0%}" if bench.get("margem_ebitda_media") else "- Margem EBITDA média: N/A (setor não-operacional)",
        f"- Dív/EBITDA média: {bench['divida_ebitda_media']:.1f}x" if bench.get("divida_ebitda_media") else "- Dív/EBITDA média: N/A",
        f"- DSCR médio: {bench['dscr_medio']:.1f}x" if bench.get("dscr_medio") else "- DSCR médio: N/A",
        f"- LTV máximo recomendado: {bench['ltv_maximo_recomendado']:.0%}" if bench.get("ltv_maximo_recomendado") else "",
        f"- Volume mediano (Pipeline ZYN): R$ {bench['volume_mediano']:,.0f}",
        f"- Prazo típico: {bench['prazo_tipico_meses']} meses",
        f"- Instrumentos comuns: {', '.join(bench.get('instrumentos_comuns', []))}",
        f"- Garantias típicas: {', '.join(bench.get('garantias_tipicas', []))}",
    ]
    if bench.get("riscos_especificos"):
        bench_lines.append("Riscos setoriais específicos:")
        for r in bench["riscos_especificos"]:
            bench_lines.append(f"  - {r}")

    benchmarks_texto = "\n".join(line for line in bench_lines if line)

    return setor, instrucoes, benchmarks_texto, stress


def analyze_credit(
    dados_extraidos: dict[str, Any],
    parametros_operacao: dict[str, Any],
    status_callback=None,
) -> dict:
    """Executa análise MAC ZYN v4 via Claude Opus com streaming."""
    client = _get_client()

    def _status(msg):
        if status_callback:
            status_callback(msg)
        logger.info(msg)

    garantias_list = parametros_operacao.get("garantias", [])
    garantias_str = ", ".join(garantias_list)
    tipo_operacao = parametros_operacao.get("tipo_operacao", "N/I")

    # Build sector-specific prompt
    setor, setor_instrucoes, benchmarks_texto, stress_setor = _build_sector_prompt(
        tipo_operacao, garantias_list
    )
    _status(f"Setor detectado: {setor.upper()} — aplicando benchmarks setoriais...")

    # Build system prompt with sector context
    system_prompt = SYSTEM_PROMPT.format(
        setor_instrucoes=setor_instrucoes,
        benchmarks_texto=benchmarks_texto,
        stress_setor=stress_setor,
    )

    dados_formatados = _format_dados(dados_extraidos)
    user_prompt = ANALYSIS_PROMPT.format(
        tomador=parametros_operacao.get("tomador", "N/I"),
        cnpj=parametros_operacao.get("cnpj", "N/I"),
        tipo_operacao=tipo_operacao,
        volume=parametros_operacao.get("volume", 0),
        prazo_meses=parametros_operacao.get("prazo_meses", 0),
        taxa=parametros_operacao.get("taxa", "N/I"),
        amortizacao=parametros_operacao.get("amortizacao", "N/I"),
        garantias=garantias_str or "N/I",
        socio_responsavel=parametros_operacao.get("socio_responsavel", "N/I"),
        setor=setor.title(),
        dados_formatados=dados_formatados,
        docs_disponiveis=_list_docs(dados_extraidos, True),
        docs_indisponiveis=_list_docs(dados_extraidos, False),
    )

    total_chars = len(user_prompt) + len(system_prompt)
    est_tokens = total_chars // 4
    _status(f"Enviando ~{est_tokens:,} tokens para {MODEL} (setor: {setor})...")

    # Streaming API call with retry
    response_text = ""
    for attempt in range(MAX_RETRIES):
        try:
            _status(f"Conectando... (tentativa {attempt + 1}/{MAX_RETRIES})")
            with client.messages.stream(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                system=system_prompt,
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

    # Inject investor matching
    _status("Gerando matching de investidores...")
    volume = parametros_operacao.get("volume", 0)
    investors = match_investors(
        tipo_operacao=tipo_operacao,
        volume=volume,
        setor=setor,
        garantias=garantias_list,
        rating=analise.get("rating_final", {}).get("nota"),
        prazo_anos=parametros_operacao.get("prazo_meses", 0) / 12 if parametros_operacao.get("prazo_meses") else None,
        top_n=30,
    )
    analise["investor_matching"] = {
        "setor_detectado": setor,
        "investidores_sugeridos": investors,
        "total_matches": len(investors),
    }

    # Inject sector benchmarks reference
    analise["_setor"] = setor

    logger.info(
        "Análise concluída — Rating: %s | Parecer: %s | Setor: %s | %d investidores sugeridos",
        analise.get("rating_final", {}).get("nota", "N/A"),
        analise.get("rating_final", {}).get("parecer", "N/A"),
        setor,
        len(investors),
    )

    return analise


INCREMENTAL_PROMPT = """\
ATUALIZAÇÃO INCREMENTAL da análise MAC ZYN v4 para: {tomador} | {tipo_operacao}

ANÁLISE ANTERIOR (JSON completo):
{analise_anterior_json}

DOCUMENTO COMPLEMENTAR RECEBIDO:
{novos_dados}

INSTRUÇÃO: Atualize a análise anterior incorporando as informações do documento complementar.
- Mantenha a MESMA estrutura JSON exata da análise anterior
- Atualize APENAS os campos impactados pelo novo documento
- Recalcule rating, KPIs, DSCR, LTV se os novos dados afetarem esses indicadores
- Adicione/atualize lacunas resolvidas pelo novo documento
- Atualize o checklist de DD se aplicável
- Mantenha investor_matching inalterado (será recalculado externamente)
- Use escala de rating: AAA/AA/A/BBB/BB/B/C/D
- Responda SOMENTE JSON válido, estrutura idêntica à análise anterior
"""


def analyze_incremental(
    analise_anterior: dict,
    novos_dados_extraidos: dict,
    parametros_operacao: dict,
    status_callback=None,
) -> dict:
    """Atualiza análise existente com dados de documentos complementares (incremental)."""
    client = _get_client()

    def _status(msg):
        if status_callback:
            status_callback(msg)
        logger.info(msg)

    _status("Análise incremental — atualizando com documento complementar...")

    # Remove investor_matching from previous analysis to save tokens
    analise_clean = {k: v for k, v in analise_anterior.items() if k not in ("investor_matching", "_setor")}
    analise_json = json.dumps(analise_clean, ensure_ascii=False)

    # Truncate if too large
    if len(analise_json) > 80000:
        analise_json = analise_json[:80000] + "\n... (truncado)"

    novos_formatados = _format_dados(novos_dados_extraidos)

    user_prompt = INCREMENTAL_PROMPT.format(
        tomador=parametros_operacao.get("tomador", "N/I"),
        tipo_operacao=parametros_operacao.get("tipo_operacao", "N/I"),
        analise_anterior_json=analise_json,
        novos_dados=novos_formatados,
    )

    total_chars = len(user_prompt)
    est_tokens = total_chars // 4
    _status(f"Incremental: ~{est_tokens:,} tokens para {MODEL}...")

    response_text = ""
    for attempt in range(MAX_RETRIES):
        try:
            _status(f"Conectando... (tentativa {attempt + 1}/{MAX_RETRIES})")
            with client.messages.stream(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                system="Você é analista de crédito sênior da ZYN Capital. Atualize a análise existente com os novos dados. Responda SOMENTE JSON válido.",
                messages=[{"role": "user", "content": user_prompt}],
            ) as stream:
                chunks = 0
                for text in stream.text_stream:
                    response_text += text
                    chunks += 1
                    if chunks % 20 == 0:
                        _status(f"Recebendo atualização... ({len(response_text):,} chars)")

            _status(f"Atualização completa: {len(response_text):,} chars. Processando...")
            break

        except anthropic.RateLimitError:
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

    # Re-inject investor matching
    _status("Atualizando matching de investidores...")
    tipo_operacao = parametros_operacao.get("tipo_operacao", "N/I")
    garantias_list = parametros_operacao.get("garantias", [])
    setor = _detect_sector(tipo_operacao, garantias_list)
    volume = parametros_operacao.get("volume", 0)
    investors = match_investors(
        tipo_operacao=tipo_operacao,
        volume=volume,
        setor=setor,
        garantias=garantias_list,
        rating=analise.get("rating_final", {}).get("nota"),
        prazo_anos=parametros_operacao.get("prazo_meses", 0) / 12 if parametros_operacao.get("prazo_meses") else None,
        top_n=30,
    )
    analise["investor_matching"] = {
        "setor_detectado": setor,
        "investidores_sugeridos": investors,
        "total_matches": len(investors),
    }
    analise["_setor"] = setor

    logger.info(
        "Análise incremental concluída — Rating: %s | Parecer: %s",
        analise.get("rating_final", {}).get("nota", "N/A"),
        analise.get("rating_final", {}).get("parecer", "N/A"),
    )

    return analise


def get_analysis_summary(analise: dict) -> str:
    """Gera resumo legível da análise de crédito (MAC ZYN v4)."""

    def _to_float(val, default=0):
        """Safely convert any value to float."""
        if isinstance(val, (int, float)):
            return float(val)
        if isinstance(val, str):
            # Remove R$, %, x, dots as thousand sep, replace comma with dot
            cleaned = val.replace("R$", "").replace("%", "").replace("x", "").replace("X", "").strip()
            cleaned = cleaned.replace(".", "").replace(",", ".")
            try:
                return float(cleaned)
            except ValueError:
                return default
        return default

    def _fmt_brl(val):
        v = _to_float(val)
        return f"R$ {v:,.0f}" if v else str(val) if val else "N/D"

    def _fmt_x(val):
        v = _to_float(val)
        return f"{v:.2f}x" if v else str(val) if val else "N/D"

    def _safe_pct(val, as_pct=False):
        v = _to_float(val)
        if v == 0 and not isinstance(val, (int, float)):
            return str(val) if val else "—"
        if as_pct and 0 < abs(v) <= 1:
            return f"{v:.1%}"
        return f"{v:.1f}%"

    rating = analise.get("rating_final", {})
    if not isinstance(rating, dict):
        rating = {}
    kpis = analise.get("kpis", {})
    if not isinstance(kpis, dict):
        kpis = {}
    tomador = analise.get("tomador", {})
    if not isinstance(tomador, dict):
        tomador = {}
    operacao = analise.get("operacao", {})
    if not isinstance(operacao, dict):
        operacao = {}
    investor_matching = analise.get("investor_matching", {})
    if not isinstance(investor_matching, dict):
        investor_matching = {}
    setor = analise.get("_setor", investor_matching.get("setor_detectado", "N/A"))

    nota = rating.get("nota", "N/A")
    parecer = rating.get("parecer", "N/A")
    justificativa = rating.get("justificativa", "—")
    recomendacoes = rating.get("recomendacoes", [])
    if not isinstance(recomendacoes, list):
        recomendacoes = []

    todas_flags = []
    secoes = ["tomador", "patrimonio", "producao", "capital", "operacao",
              "pagamento", "onus", "riscos", "covenants", "cronograma"]
    for s in secoes:
        sec_data = analise.get(s, {})
        if isinstance(sec_data, dict):
            flags = sec_data.get("flags", [])
            if isinstance(flags, list):
                for flag in flags:
                    todas_flags.append(f"  [{s.title()}] {flag}")

    ratings_secao = []
    for s in secoes:
        sec_data = analise.get(s, {})
        r_s = sec_data.get("rating_secao", "N/A") if isinstance(sec_data, dict) else "N/A"
        ratings_secao.append(f"  {s.title()}: {r_s}")

    vol = operacao.get("volume", 0)
    vol_str = _fmt_brl(vol)

    linhas = [
        "=" * 60,
        "  MAC ZYN v4 — RESUMO EXECUTIVO",
        "=" * 60, "",
        f"Tomador:  {tomador.get('razao_social', 'N/A')}",
        f"CNPJ:     {tomador.get('cnpj', 'N/A')}",
        f"Setor:    {setor.title() if isinstance(setor, str) else 'N/A'}",
        f"Operação: {operacao.get('tipo', 'N/A')} — {vol_str}",
        f"Prazo:    {operacao.get('prazo', 'N/A')} | Taxa: {operacao.get('taxa', 'N/A')}",
        "", "-" * 60,
        f"  RATING: {nota}  |  PARECER: {parecer}",
        "-" * 60, "",
        "KPIs:",
        f"  Receita:    {_fmt_brl(kpis.get('receita_liquida', 0))}",
        f"  EBITDA:     {_fmt_brl(kpis.get('ebitda', 0))}",
        f"  Margem:     {_safe_pct(kpis.get('margem_ebitda', 0), True)}",
        f"  Dív/EBITDA: {_fmt_x(kpis.get('divida_liquida_ebitda', 0))}",
        f"  LTV:        {_safe_pct(kpis.get('ltv', 0), True)}",
        f"  DSCR:       {_fmt_x(kpis.get('dscr', 0))}",
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
    if not isinstance(checklist, dict):
        checklist = {}
    docs_f = checklist.get("documentos_faltantes", [])
    info_p = checklist.get("informacoes_pendentes", [])
    if not isinstance(docs_f, list):
        docs_f = []
    if not isinstance(info_p, list):
        info_p = []
    if docs_f or info_p:
        linhas.extend(["-" * 60,
            f"  LACUNAS | Pendências: {checklist.get('total_pendencias', 0)} | Críticas: {checklist.get('total_criticas', 0)}",
            "-" * 60, ""])
        for d in docs_f:
            if isinstance(d, dict):
                linhas.append(f"  [{d.get('criticidade', '?')}] {d.get('item', '?')}")
        for p in info_p:
            if isinstance(p, dict):
                linhas.append(f"  [{p.get('criticidade', '?')}] {p.get('item', '?')}")
        linhas.append("")

    # Investor matching summary
    investidores = investor_matching.get("investidores_sugeridos", [])
    if isinstance(investidores, list) and investidores:
        linhas.extend(["-" * 60,
            f"  INVESTIDORES SUGERIDOS ({len(investidores)})",
            "-" * 60, ""])
        for inv in investidores[:5]:
            if isinstance(inv, dict):
                motivos = inv.get("motivos", [])
                motivos_str = " | ".join(motivos) if isinstance(motivos, list) else str(motivos)
                score = _to_float(inv.get("score", 0))
                linhas.append(f"  [{score:.0f} pts] {inv.get('nome', '?')} — {motivos_str}")
        linhas.append("")

    linhas.extend(["=" * 60, "  ZYN Capital — Crédito Estruturado & M&A", "=" * 60])
    return "\n".join(linhas)
