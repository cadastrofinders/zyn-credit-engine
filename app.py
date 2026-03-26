"""
ZYN Credit Engine — Plataforma de Análise de Crédito Estruturado
Streamlit App principal.
"""

import json
import os
import time
from datetime import datetime
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# Streamlit Cloud: inject st.secrets into os.environ so modules can read them
# ---------------------------------------------------------------------------
try:
    for key in ("ANTHROPIC_API_KEY", "NOTION_TOKEN"):
        if key in st.secrets and not os.environ.get(key):
            os.environ[key] = st.secrets[key]
except Exception:
    pass  # secrets not available (local dev)

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
UPLOADS_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "output"
UPLOADS_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# Module imports (lazy to allow app to load even if deps missing)
# ---------------------------------------------------------------------------
try:
    from modules.extractor import process_file
    from modules.analyzer import analyze_credit, get_analysis_summary
    from modules.docx_generator import generate_mac

    MODULES_AVAILABLE = True
except ImportError as _imp_err:
    MODULES_AVAILABLE = False
    _IMPORT_ERROR = str(_imp_err)

# ---------------------------------------------------------------------------
# API Key check
# ---------------------------------------------------------------------------
API_KEY_SET = bool(os.environ.get("ANTHROPIC_API_KEY"))

# ---------------------------------------------------------------------------
# DD Checklist Template — 12 módulos com itens realistas
# ---------------------------------------------------------------------------
DD_CHECKLIST_TEMPLATE: dict[str, list[str]] = {
    "1. Cadastral": [
        "Cartão CNPJ atualizado",
        "Inscrição Estadual / Municipal",
        "Comprovante de endereço da sede (< 90 dias)",
        "Procurações vigentes dos representantes legais",
        "Ficha cadastral completa (KYC)",
        "Quadro societário atualizado",
        "Declaração de grupo econômico",
    ],
    "2. Societário": [
        "Contrato Social consolidado ou Estatuto Social",
        "Última alteração contratual",
        "Ata de eleição da diretoria vigente",
        "Acordo de acionistas (se houver)",
        "Organograma societário do grupo",
        "Certidão simplificada da Junta Comercial",
    ],
    "3. Financeiro": [
        "Balanço Patrimonial (último exercício auditado)",
        "Balanço Patrimonial (penúltimo exercício)",
        "DRE (último exercício)",
        "DRE (penúltimo exercício)",
        "Balancete acumulado do exercício corrente",
        "Fluxo de Caixa projetado (12 meses)",
        "Relatório de auditoria independente",
        "Notas explicativas",
    ],
    "4. Faturamento": [
        "Relatório de faturamento mensal (últimos 12 meses)",
        "Principais clientes (top 10 por receita)",
        "Contratos com clientes relevantes",
        "Concentração de receita por cliente",
        "Histórico de inadimplência de clientes",
    ],
    "5. Endividamento": [
        "Relação completa de endividamento bancário",
        "SCR / Registrato atualizado",
        "Contratos de empréstimo vigentes",
        "Cronograma de amortização por credor",
        "Cartas de anuência de credores (se necessário)",
        "Posição de derivativos / hedge",
    ],
    "6. FIDC/Fundo": [
        "Regulamento do Fundo (se aplicável)",
        "Política de crédito e elegibilidade de lastro",
        "Relatório do administrador/custodiante",
        "Composição da carteira de recebíveis",
        "Histórico de PDD e inadimplência da carteira",
        "Rating do fundo (se houver)",
    ],
    "7. Crédito": [
        "Histórico de crédito do tomador (Serasa/SPC)",
        "Score de crédito (bureau)",
        "Análise setorial",
        "Relatório de visita técnica",
        "Parecer jurídico sobre a operação",
        "Matriz de risco e mitigantes",
    ],
    "8. Sócios PF": [
        "Documentos pessoais (RG/CPF) dos sócios",
        "Declaração de IR dos sócios (último exercício)",
        "Certidão negativa de débitos fiscais (PF)",
        "Pesquisa de protestos e ações judiciais (PF)",
        "Certidão de distribuição cível e criminal",
        "Declaração de bens e patrimônio",
    ],
    "9. Certidões": [
        "CND Federal (Receita Federal / PGFN)",
        "CND Estadual (SEFAZ)",
        "CND Municipal (ISS / IPTU)",
        "CND Trabalhista (TST)",
        "CND FGTS (CEF)",
        "Certidão de protestos",
        "Certidão de distribuição cível",
        "Certidão de distribuição criminal",
    ],
    "10. Garantias": [
        "Matrícula atualizada do imóvel (< 30 dias)",
        "Laudo de avaliação do imóvel",
        "CCIR quitado (imóvel rural)",
        "CAR — Cadastro Ambiental Rural",
        "Certidão de ônus reais",
        "Apólice de seguro do bem",
        "Contrato de penhor / alienação fiduciária",
        "ITR quitado (imóvel rural)",
    ],
    "11. Jurídico": [
        "Pesquisa de processos judiciais (PJ)",
        "Contingências relevantes provisionadas",
        "Pareceres jurídicos sobre contingências",
        "Contratos materiais vigentes",
        "Licenças e alvarás operacionais",
        "Compliance e LGPD",
    ],
    "12. Institucional": [
        "Apresentação institucional da empresa",
        "Relatório de sustentabilidade / ESG",
        "Política de governança corporativa",
        "Estrutura de compliance / PLD",
        "Referências bancárias e comerciais",
    ],
}

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------
CUSTOM_CSS = """
<style>
    /* ── Sidebar ─────────────────────────────────────────── */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1a2633 0%, #223040 40%, #2a3d52 100%);
    }
    section[data-testid="stSidebar"] * {
        color: #FFFFFF !important;
    }
    section[data-testid="stSidebar"] hr {
        border-color: rgba(255,255,255,0.12) !important;
    }
    /* Radio buttons as nav items */
    section[data-testid="stSidebar"] label[data-baseweb="radio"] {
        background: rgba(255,255,255,0.06) !important;
        border: 1px solid rgba(255,255,255,0.12) !important;
        border-radius: 8px !important;
        padding: 10px 14px !important;
        margin-bottom: 4px !important;
        transition: all 0.2s ease !important;
        cursor: pointer !important;
    }
    section[data-testid="stSidebar"] label[data-baseweb="radio"]:hover {
        background: rgba(255,255,255,0.14) !important;
        border-color: rgba(255,255,255,0.25) !important;
    }
    /* Selected radio */
    section[data-testid="stSidebar"] label[data-baseweb="radio"]:has(input:checked) {
        background: rgba(30,107,66,0.35) !important;
        border-color: #1E6B42 !important;
    }
    /* Hide radio circle */
    section[data-testid="stSidebar"] [data-testid="stRadio"] > div[role="radiogroup"] > label > div:first-child {
        display: none !important;
    }

    /* ── Page title ──────────────────────────────────────── */
    .main-title {
        font-size: 2.2rem;
        font-weight: 800;
        color: #223040;
        letter-spacing: -0.5px;
        margin-bottom: 0;
    }
    .main-subtitle {
        font-size: 1.1rem;
        color: #8B9197;
        margin-top: 0;
        margin-bottom: 8px;
    }

    /* ── Metric cards ────────────────────────────────────── */
    [data-testid="stMetric"] {
        background: #FFFFFF;
        border: 1px solid #E8ECF0;
        border-left: 4px solid #223040;
        padding: 16px 20px;
        border-radius: 10px;
        box-shadow: 0 1px 4px rgba(34,48,64,0.06);
        transition: box-shadow 0.2s ease;
    }
    [data-testid="stMetric"]:hover {
        box-shadow: 0 4px 12px rgba(34,48,64,0.10);
    }
    [data-testid="stMetric"] label {
        color: #8B9197 !important;
        font-size: 0.85rem !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: #223040 !important;
        font-weight: 700 !important;
    }

    /* ── KPI strip (Nova Análise > Análise tab) ──────────── */
    .kpi-card {
        background: #FFFFFF;
        border: 1px solid #E8ECF0;
        border-radius: 10px;
        padding: 16px 12px;
        text-align: center;
        box-shadow: 0 1px 4px rgba(34,48,64,0.06);
    }
    .kpi-label {
        font-size: 0.72rem;
        color: #8B9197;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        margin-bottom: 4px;
    }
    .kpi-value {
        font-size: 1.5rem;
        font-weight: 800;
        color: #223040;
    }

    /* ── Rating badges ───────────────────────────────────── */
    .rating-badge {
        display: inline-block;
        padding: 12px 32px;
        border-radius: 12px;
        font-size: 2.8rem;
        font-weight: 800;
        color: #FFFFFF;
        text-align: center;
        letter-spacing: 2px;
        box-shadow: 0 4px 14px rgba(0,0,0,0.15);
    }
    .rating-A { background: linear-gradient(135deg, #1E6B42, #2E9B62); }
    .rating-B { background: linear-gradient(135deg, #223040, #3A5570); }
    .rating-C { background: linear-gradient(135deg, #7D6608, #B8960E); }
    .rating-D { background: linear-gradient(135deg, #E65100, #FF7A22); }
    .rating-E { background: linear-gradient(135deg, #922B21, #C0392B); }

    /* ── Parecer colors ──────────────────────────────────── */
    .parecer-favoravel { color: #1E6B42; font-weight: 700; font-size: 1.1rem; }
    .parecer-ressalvas { color: #7D6608; font-weight: 700; font-size: 1.1rem; }
    .parecer-desfavoravel { color: #922B21; font-weight: 700; font-size: 1.1rem; }

    /* ── Section headers ─────────────────────────────────── */
    .section-header {
        background: linear-gradient(90deg, #223040, #2a3d52);
        color: #FFFFFF;
        padding: 10px 18px;
        border-radius: 8px;
        margin-top: 18px;
        font-weight: 600;
        letter-spacing: 0.3px;
    }

    /* ── Status badges (DD checklist) ────────────────────── */
    .status-ok {
        background-color: #EAF4EE;
        color: #1E6B42;
        padding: 3px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.8rem;
    }
    .status-pendente {
        background-color: #FDECEA;
        color: #922B21;
        padding: 3px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.8rem;
    }
    .status-desatualizado {
        background-color: #FEF9E7;
        color: #7D6608;
        padding: 3px 10px;
        border-radius: 6px;
        font-weight: 600;
        font-size: 0.8rem;
    }

    /* ── Tabs ─────────────────────────────────────────────── */
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 10px 20px;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #223040 !important;
        color: #FFFFFF !important;
    }

    /* ── File uploader ───────────────────────────────────── */
    [data-testid="stFileUploader"] {
        border: 2px dashed #D0D5DA;
        border-radius: 12px;
        padding: 8px;
        transition: border-color 0.2s ease;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #223040;
    }

    /* ── Expander (analysis sections) ────────────────────── */
    .streamlit-expanderHeader {
        background-color: #F2F4F6 !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        color: #223040 !important;
    }

    /* ── Progress bars ───────────────────────────────────── */
    .stProgress > div > div > div {
        background-color: #1E6B42 !important;
    }

    /* ── Info / Warning / Error boxes ────────────────────── */
    .stAlert {
        border-radius: 10px !important;
    }

    /* ── Footer ──────────────────────────────────────────── */
    .footer-text {
        color: rgba(255,255,255,0.5);
        font-size: 0.75rem;
        text-align: center;
        padding-top: 24px;
        letter-spacing: 0.5px;
    }

    /* ── General polish ──────────────────────────────────── */
    .block-container {
        padding-top: 2rem !important;
    }
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.2s ease;
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(34,48,64,0.15);
    }
    .stDownloadButton > button {
        background-color: #1E6B42 !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 8px;
        font-weight: 600;
        padding: 8px 24px;
    }
    .stDownloadButton > button:hover {
        background-color: #246E45 !important;
        box-shadow: 0 4px 12px rgba(30,125,79,0.25);
    }
</style>
"""

# ---------------------------------------------------------------------------
# Session State Initialization
# ---------------------------------------------------------------------------
DEFAULTS: dict = {
    "operacoes": [],
    "current_op": None,
    "uploaded_files": [],
    "extracted_data": {},
    "analysis": None,
    "step": 0,
    "dd_status": {},
}

for key, default in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ---------------------------------------------------------------------------
# Persistence — save/load extraction + analysis + op to disk so they survive
# reloads and redeploys on Streamlit Cloud
# ---------------------------------------------------------------------------
CACHE_FILE = OUTPUT_DIR / "session_cache.json"


def _save_cache():
    """Persist extracted_data, analysis, and current_op to disk."""
    try:
        payload = {
            "extracted_data": st.session_state.extracted_data or {},
            "analysis": st.session_state.analysis,
            "current_op": st.session_state.current_op,
        }
        CACHE_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2))
    except Exception:
        pass


def _load_cache():
    """Restore cached data if session_state is empty and cache file exists."""
    if st.session_state.extracted_data or not CACHE_FILE.exists():
        return
    try:
        payload = json.loads(CACHE_FILE.read_text())
        if payload.get("extracted_data"):
            st.session_state.extracted_data = payload["extracted_data"]
        if payload.get("analysis"):
            st.session_state.analysis = payload["analysis"]
        if payload.get("current_op"):
            st.session_state.current_op = payload["current_op"]
    except Exception:
        pass


_load_cache()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _fmt_brl(value: float) -> str:
    """Formata valor em R$ brasileiro."""
    if value >= 1_000_000:
        return f"R$ {value / 1_000_000:,.2f} MM"
    if value >= 1_000:
        return f"R$ {value / 1_000:,.1f} mil"
    return f"R$ {value:,.2f}"


def _confidence_color(conf: float) -> str:
    if conf >= 0.85:
        return "green"
    if conf >= 0.6:
        return "orange"
    return "red"


def _rating_class(nota: str) -> str:
    return f"rating-{nota}" if nota in "ABCDE" else "rating-C"


def _parecer_class(parecer: str) -> str:
    if "Desfavorável" in parecer:
        return "parecer-desfavoravel"
    if "Ressalvas" in parecer:
        return "parecer-ressalvas"
    return "parecer-favoravel"


def _count_ops_by_status(status: str) -> int:
    return sum(1 for op in st.session_state.operacoes if op.get("status") == status)


def _detected_doc_types() -> set[str]:
    """Retorna conjunto de tipos de documento detectados na extração."""
    tipos = set()
    for item in st.session_state.extracted_data.values():
        classificacao = item.get("classificacao", {})
        tipo = classificacao.get("tipo")
        if tipo:
            tipos.add(tipo)
    return tipos


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------
def page_dashboard():
    # ── Hero Section ──────────────────────────────────────────────────────
    st.markdown(
        """
        <div style="
            background: linear-gradient(135deg, #223040 0%, #2a3d52 60%, #1E6B42 100%);
            border-radius: 16px;
            padding: 40px 36px 32px 36px;
            margin-bottom: 24px;
        ">
            <h1 style="color:#FFFFFF; font-size:2.4rem; font-weight:800; margin:0; letter-spacing:-0.5px;">
                ZYN Credit Engine
            </h1>
            <p style="color:rgba(255,255,255,0.7); font-size:1.05rem; margin:6px 0 0 0;">
                Motor de Análise de Crédito Estruturado &nbsp;·&nbsp; Powered by Claude AI
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── KPI Cards ─────────────────────────────────────────────────────────
    total = len(st.session_state.operacoes)
    em_andamento = _count_ops_by_status("Em Andamento")
    concluidas = _count_ops_by_status("Concluída")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(
            f"""<div style="background:#FFFFFF; border:1px solid #E8ECF0; border-left:4px solid #223040;
            border-radius:10px; padding:20px; box-shadow:0 1px 4px rgba(34,48,64,0.06);">
            <p style="color:#8B9197; font-size:0.75rem; text-transform:uppercase; letter-spacing:0.8px; margin:0 0 6px 0;">Total Operações</p>
            <p style="color:#223040; font-size:2rem; font-weight:800; margin:0;">{total}</p>
            </div>""", unsafe_allow_html=True)
    with col2:
        st.markdown(
            f"""<div style="background:#FFFFFF; border:1px solid #E8ECF0; border-left:4px solid #7D6608;
            border-radius:10px; padding:20px; box-shadow:0 1px 4px rgba(34,48,64,0.06);">
            <p style="color:#8B9197; font-size:0.75rem; text-transform:uppercase; letter-spacing:0.8px; margin:0 0 6px 0;">Em Andamento</p>
            <p style="color:#7D6608; font-size:2rem; font-weight:800; margin:0;">{em_andamento}</p>
            </div>""", unsafe_allow_html=True)
    with col3:
        st.markdown(
            f"""<div style="background:#FFFFFF; border:1px solid #E8ECF0; border-left:4px solid #1E6B42;
            border-radius:10px; padding:20px; box-shadow:0 1px 4px rgba(34,48,64,0.06);">
            <p style="color:#8B9197; font-size:0.75rem; text-transform:uppercase; letter-spacing:0.8px; margin:0 0 6px 0;">Concluídas</p>
            <p style="color:#1E6B42; font-size:2rem; font-weight:800; margin:0;">{concluidas}</p>
            </div>""", unsafe_allow_html=True)
    with col4:
        taxa_sucesso = f"{(concluidas / total * 100):.0f}%" if total > 0 else "—"
        st.markdown(
            f"""<div style="background:#FFFFFF; border:1px solid #E8ECF0; border-left:4px solid #922B21;
            border-radius:10px; padding:20px; box-shadow:0 1px 4px rgba(34,48,64,0.06);">
            <p style="color:#8B9197; font-size:0.75rem; text-transform:uppercase; letter-spacing:0.8px; margin:0 0 6px 0;">Taxa de Conclusão</p>
            <p style="color:#223040; font-size:2rem; font-weight:800; margin:0;">{taxa_sucesso}</p>
            </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)

    # ── Operações Recentes ────────────────────────────────────────────────
    if st.session_state.operacoes:
        st.markdown(
            '<p style="font-size:1.1rem; font-weight:700; color:#223040; margin-bottom:8px;">Operações Recentes</p>',
            unsafe_allow_html=True,
        )
        rows = []
        for op in reversed(st.session_state.operacoes[-20:]):
            rows.append({
                "Tomador": op.get("tomador", ""),
                "CNPJ": op.get("cnpj", ""),
                "Tipo": op.get("tipo_operacao", ""),
                "Volume": _fmt_brl(op.get("volume", 0)),
                "Status": op.get("status", ""),
                "Rating": op.get("rating", "—"),
                "Data": op.get("data_criacao", ""),
            })
        st.dataframe(rows, use_container_width=True, hide_index=True)
    else:
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        st.markdown(
            """
            <div style="background:#F2F4F6; border-radius:12px; padding:40px; text-align:center;">
                <p style="font-size:1.6rem; margin:0 0 8px 0;">Nenhuma operação registrada</p>
                <p style="color:#8B9197; margin:0;">Clique em <b>Nova Análise</b> no menu lateral para iniciar sua primeira análise de crédito.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # ── Pipeline rápido ───────────────────────────────────────────────────
    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    st.markdown(
        """
        <div style="background:linear-gradient(90deg, #223040, #2a3d52); border-radius:12px; padding:24px 28px;">
            <p style="color:#FFFFFF; font-weight:700; font-size:1rem; margin:0 0 12px 0;">Como funciona</p>
            <div style="display:flex; gap:16px; flex-wrap:wrap;">
                <div style="flex:1; min-width:140px; background:rgba(255,255,255,0.08); border-radius:8px; padding:14px; text-align:center;">
                    <p style="color:#A8B4C0; font-size:0.7rem; text-transform:uppercase; letter-spacing:1px; margin:0 0 4px 0;">Passo 1</p>
                    <p style="color:#FFFFFF; font-weight:600; font-size:0.9rem; margin:0;">Upload de Docs</p>
                </div>
                <div style="flex:1; min-width:140px; background:rgba(255,255,255,0.08); border-radius:8px; padding:14px; text-align:center;">
                    <p style="color:#A8B4C0; font-size:0.7rem; text-transform:uppercase; letter-spacing:1px; margin:0 0 4px 0;">Passo 2</p>
                    <p style="color:#FFFFFF; font-weight:600; font-size:0.9rem; margin:0;">Extração via Sonnet</p>
                </div>
                <div style="flex:1; min-width:140px; background:rgba(255,255,255,0.08); border-radius:8px; padding:14px; text-align:center;">
                    <p style="color:#A8B4C0; font-size:0.7rem; text-transform:uppercase; letter-spacing:1px; margin:0 0 4px 0;">Passo 3</p>
                    <p style="color:#FFFFFF; font-weight:600; font-size:0.9rem; margin:0;">Análise via Opus</p>
                </div>
                <div style="flex:1; min-width:140px; background:rgba(255,255,255,0.08); border-radius:8px; padding:14px; text-align:center;">
                    <p style="color:#A8B4C0; font-size:0.7rem; text-transform:uppercase; letter-spacing:1px; margin:0 0 4px 0;">Passo 4</p>
                    <p style="color:#FFFFFF; font-weight:600; font-size:0.9rem; margin:0;">MAC .docx Pronto</p>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def page_nova_analise():
    st.markdown(
        """
        <div style="margin-bottom:20px;">
            <h1 style="color:#223040; font-size:1.8rem; font-weight:800; margin:0;">Nova Análise de Crédito</h1>
            <p style="color:#8B9197; font-size:0.95rem; margin:4px 0 0 0;">
                Preencha os dados da operação, faça upload dos documentos e gere o MAC automaticamente.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not API_KEY_SET:
        st.warning(
            "**ANTHROPIC_API_KEY** não configurada. "
            "A extração e análise via Claude não funcionarão. "
            "Configure os Secrets no Streamlit Cloud."
        )

    if not MODULES_AVAILABLE:
        st.error(f"Erro ao importar módulos: {_IMPORT_ERROR}")
        return

    tab_upload, tab_extracao, tab_analise, tab_mac = st.tabs([
        "1 - Upload & Dados",
        "2 - Extração",
        "3 - Análise",
        "4 - MAC .docx",
    ])

    # ------------------------------------------------------------------
    # TAB 1 — Upload & Dados
    # ------------------------------------------------------------------
    with tab_upload:
        st.markdown(
            '<p style="font-size:1.1rem; font-weight:700; color:#223040; margin-bottom:4px;">Parâmetros da Operação</p>',
            unsafe_allow_html=True,
        )
        st.caption("Preencha os campos abaixo com os dados do tomador e da operação.")

        with st.form("form_parametros", clear_on_submit=False):
            col_a, col_b = st.columns(2)

            with col_a:
                tomador = st.text_input("Tomador *", value=st.session_state.current_op.get("tomador", "") if st.session_state.current_op else "")
                cnpj = st.text_input("CNPJ", placeholder="00.000.000/0001-00", value=st.session_state.current_op.get("cnpj", "") if st.session_state.current_op else "")
                tipo_operacao = st.selectbox(
                    "Tipo de Operação",
                    ["CRA", "CRI", "FIDC", "Fiagro", "SLB", "NC/CCB", "CPR-F", "Debênture"],
                    index=0,
                )
                volume = st.number_input("Volume (R$)", min_value=0.0, format="%.2f", value=st.session_state.current_op.get("volume", 0.0) if st.session_state.current_op else 0.0)
                prazo_meses = st.number_input("Prazo (meses)", min_value=1, value=st.session_state.current_op.get("prazo_meses", 12) if st.session_state.current_op else 12)

            with col_b:
                taxa = st.text_input("Taxa", placeholder="CDI+4%", value=st.session_state.current_op.get("taxa", "") if st.session_state.current_op else "")
                amortizacao = st.selectbox("Amortização", ["SAC", "Price", "Bullet"])
                garantias_text = st.text_area("Garantias (uma por linha)", value=st.session_state.current_op.get("garantias_text", "") if st.session_state.current_op else "")
                tipo_captacao = st.selectbox("Tipo de Captação", ["Captação PJ", "Captação PF"])
                instrumento = st.text_input("Instrumento", placeholder="NC / CCB", value=st.session_state.current_op.get("instrumento", "") if st.session_state.current_op else "")

            socio_responsavel = st.selectbox(
                "Sócio Responsável",
                ["Danilo Salasar", "Renato Barison", "Luiz Roberto Evangelista"],
            )

            submitted = st.form_submit_button("Salvar e Prosseguir", use_container_width=True)

        if submitted:
            if not tomador.strip():
                st.error("O campo **Tomador** é obrigatório.")
            else:
                garantias_list = [g.strip() for g in garantias_text.strip().split("\n") if g.strip()]

                op = {
                    "tomador": tomador.strip(),
                    "cnpj": cnpj.strip(),
                    "tipo_operacao": tipo_operacao,
                    "volume": volume,
                    "prazo_meses": prazo_meses,
                    "taxa": taxa.strip(),
                    "amortizacao": amortizacao,
                    "garantias": garantias_list,
                    "garantias_text": garantias_text,
                    "tipo_captacao": tipo_captacao,
                    "instrumento": instrumento.strip(),
                    "socio_responsavel": socio_responsavel,
                    "status": "Em Andamento",
                    "rating": "—",
                    "data_criacao": datetime.now().strftime("%Y-%m-%d %H:%M"),
                }
                st.session_state.current_op = op
                _save_cache()
                st.success("Parâmetros salvos com sucesso.")

        st.markdown("---")
        st.markdown("### Upload de Documentos")

        uploaded = st.file_uploader(
            "Arraste ou selecione os documentos",
            accept_multiple_files=True,
            type=["pdf", "xlsx", "xls", "png", "jpg", "jpeg", "docx", "pptx"],
        )

        if uploaded:
            st.session_state.uploaded_files = []
            for f in uploaded:
                file_bytes = f.read()
                st.session_state.uploaded_files.append({
                    "name": f.name,
                    "size": len(file_bytes),
                    "bytes": file_bytes,
                })

            st.markdown("**Arquivos carregados:**")
            for uf in st.session_state.uploaded_files:
                size_kb = uf["size"] / 1024
                st.markdown(f"- **{uf['name']}** — {size_kb:,.1f} KB")

    # ------------------------------------------------------------------
    # TAB 2 — Extração
    # ------------------------------------------------------------------
    with tab_extracao:
        st.markdown("### Extração de Dados")

        if not st.session_state.uploaded_files:
            st.info("Nenhum arquivo carregado. Faça o upload na aba anterior.")
        else:
            st.markdown(f"**{len(st.session_state.uploaded_files)}** arquivo(s) disponível(is) para extração.")
            for uf in st.session_state.uploaded_files:
                st.markdown(f"- {uf['name']}")

            if st.button("Iniciar Extração", use_container_width=True, type="primary"):
                if not API_KEY_SET:
                    st.error("ANTHROPIC_API_KEY não configurada. Não é possível realizar a extração.")
                else:
                    progress_bar = st.progress(0, text="Iniciando extração...")
                    total_files = len(st.session_state.uploaded_files)
                    results: dict = {}
                    errors: list[str] = []

                    for idx, uf in enumerate(st.session_state.uploaded_files):
                        progress_bar.progress(
                            (idx) / total_files,
                            text=f"Processando {uf['name']}... ({idx + 1}/{total_files})",
                        )
                        try:
                            result = process_file(uf["bytes"], uf["name"])
                            results[uf["name"]] = result
                        except Exception as e:
                            errors.append(f"{uf['name']}: {e}")
                            results[uf["name"]] = {
                                "classificacao": {"tipo": "erro", "confianca": 0.0, "descricao": str(e)},
                                "dados": {"error": str(e)},
                            }

                    progress_bar.progress(1.0, text="Extração concluída.")
                    st.session_state.extracted_data = results
                    _save_cache()

                    if errors:
                        for err in errors:
                            st.error(f"❌ {err}")

            # Show extraction results
            if st.session_state.extracted_data:
                st.markdown("---")
                st.markdown("### Resultados da Extração")

                for fname, result in st.session_state.extracted_data.items():
                    classificacao = result.get("classificacao", {})
                    dados = result.get("dados", {})
                    tipo = classificacao.get("tipo", "desconhecido")
                    confianca = classificacao.get("confianca", 0.0)
                    descricao = classificacao.get("descricao", "")
                    has_error = "error" in dados

                    status_icon = "❌ Erro" if has_error else "✅ Extraído"
                    conf_color = _confidence_color(confianca)

                    with st.expander(f"{status_icon} — **{fname}**", expanded=not has_error):
                        col1, col2 = st.columns([1, 2])
                        with col1:
                            st.markdown(f"**Tipo detectado:** `{tipo}`")
                            st.markdown(
                                f"**Confiança:** :{conf_color}[{confianca:.0%}]"
                            )
                            if descricao:
                                st.caption(descricao)
                        with col2:
                            if has_error:
                                st.error(f"Erro: {dados['error']}")
                            else:
                                st.json(dados, expanded=False)

                # Checklist automático de tipos documentais
                st.markdown("---")
                st.markdown("### Checklist Automático de Documentos")

                tipos_esperados = {
                    "balanco": "Balanço Patrimonial",
                    "dre": "DRE",
                    "balancete": "Balancete",
                    "matricula": "Matrícula de Imóvel",
                    "contrato": "Contrato",
                    "certidao": "Certidão",
                    "ccir_car": "CCIR / CAR",
                }
                tipos_detectados = _detected_doc_types()

                col_ok, col_pend = st.columns(2)
                with col_ok:
                    st.markdown("**Documentos encontrados:**")
                    encontrados = [
                        label for tipo, label in tipos_esperados.items() if tipo in tipos_detectados
                    ]
                    if encontrados:
                        for item in encontrados:
                            st.markdown(f"- ✅ {item}")
                    else:
                        st.caption("Nenhum tipo padrão detectado.")

                with col_pend:
                    st.markdown("**Documentos não encontrados:**")
                    faltantes = [
                        label for tipo, label in tipos_esperados.items() if tipo not in tipos_detectados
                    ]
                    if faltantes:
                        for item in faltantes:
                            st.markdown(f"- ⚠️ {item}")
                    else:
                        st.markdown("Todos os tipos padrão foram detectados.")

    # ------------------------------------------------------------------
    # TAB 3 — Análise
    # ------------------------------------------------------------------
    with tab_analise:
        st.markdown("### Análise de Crédito")

        if not st.session_state.current_op:
            st.info("Preencha os parâmetros da operação na aba **Upload & Dados** antes de gerar a análise.")
        elif not st.session_state.extracted_data:
            st.info("Execute a extração de documentos na aba **Extração** antes de gerar a análise.")
        else:
            op = st.session_state.current_op
            st.markdown(f"**Tomador:** {op['tomador']}  |  **Tipo:** {op['tipo_operacao']}  |  **Volume:** {_fmt_brl(op['volume'])}")

            # Build dados_extraidos dict keyed by doc type for the analyzer
            dados_para_analise: dict = {}
            for fname, result in st.session_state.extracted_data.items():
                classificacao = result.get("classificacao", {})
                dados = result.get("dados", {})
                tipo = classificacao.get("tipo", "outro")
                if "error" not in dados:
                    dados_para_analise[tipo] = dados

            st.markdown(f"Dados disponíveis: **{len(dados_para_analise)}** tipo(s) documental(is).")

            if st.button("Gerar Análise de Crédito", use_container_width=True, type="primary"):
                if not API_KEY_SET:
                    st.error("ANTHROPIC_API_KEY não configurada. Não é possível realizar a análise.")
                else:
                    status_container = st.empty()
                    def _update_status(msg):
                        status_container.info(f"⏳ {msg}")
                    with st.spinner("Analisando com Claude Sonnet..."):
                        try:
                            analise = analyze_credit(dados_para_analise, op, status_callback=_update_status)
                            st.session_state.analysis = analise

                            # Update operation record
                            rating_final = analise.get("rating_final", {})
                            op["rating"] = rating_final.get("nota", "—")
                            op["parecer"] = rating_final.get("parecer", "—")
                            st.session_state.current_op = op
                            _save_cache()

                            status_container.empty()
                            st.success("Análise concluída com sucesso.")
                        except Exception as e:
                            status_container.empty()
                            st.error(f"Erro durante a análise: {e}")

            # Display analysis results
            if st.session_state.analysis:
                analise = st.session_state.analysis
                rating_final = analise.get("rating_final", {})
                kpis = analise.get("kpis", {})
                nota = rating_final.get("nota", "N/A")
                parecer = rating_final.get("parecer", "N/A")

                st.markdown("---")

                # Rating + Parecer header
                col_rating, col_parecer = st.columns([1, 2])
                with col_rating:
                    st.markdown(
                        f'<div class="rating-badge {_rating_class(nota)}">{nota}</div>',
                        unsafe_allow_html=True,
                    )
                    st.caption("Rating Final")
                with col_parecer:
                    st.markdown(
                        f'<span class="{_parecer_class(parecer)}" style="font-size:1.5rem;">{parecer}</span>',
                        unsafe_allow_html=True,
                    )

                # KPIs row
                st.markdown("---")
                st.markdown("### Indicadores-Chave")
                k1, k2, k3, k4, k5, k6 = st.columns(6)
                k1.metric("Receita Líq.", _fmt_brl(kpis.get("receita_liquida", 0)))
                k2.metric("EBITDA", _fmt_brl(kpis.get("ebitda", 0)))

                margem = kpis.get("margem_ebitda", 0)
                margem_str = f"{margem:.1%}" if isinstance(margem, (int, float)) and margem <= 1 else f"{margem:.1f}%"
                k3.metric("Margem EBITDA", margem_str)

                k4.metric("Dív.Líq./EBITDA", f"{kpis.get('divida_liquida_ebitda', 0):.2f}x")

                ltv = kpis.get("ltv", 0)
                ltv_str = f"{ltv:.1%}" if isinstance(ltv, (int, float)) and ltv <= 1 else f"{ltv:.1f}%"
                k5.metric("LTV", ltv_str)

                k6.metric("DSCR", f"{kpis.get('dscr', 0):.2f}x")

                # 10 analysis sections
                st.markdown("---")
                st.markdown("### Seções da Análise")

                secoes = [
                    ("tomador", "1. Tomador"),
                    ("patrimonio", "2. Patrimônio"),
                    ("producao", "3. Produção"),
                    ("capital", "4. Estrutura de Capital"),
                    ("operacao", "5. Operação"),
                    ("pagamento", "6. Capacidade de Pagamento"),
                    ("onus", "7. Ônus e Restrições"),
                    ("riscos", "8. Riscos"),
                    ("covenants", "9. Covenants"),
                    ("cronograma", "10. Cronograma"),
                ]

                for key, label in secoes:
                    secao_data = analise.get(key, {})
                    rating_secao = secao_data.get("rating_secao", "N/A")
                    flags = secao_data.get("flags", [])

                    with st.expander(f"**{label}** — _{rating_secao}_"):
                        # Show all string fields
                        for campo, valor in secao_data.items():
                            if campo in ("rating_secao", "flags"):
                                continue
                            if isinstance(valor, str) and valor:
                                st.markdown(f"**{campo.replace('_', ' ').title()}:**")
                                st.markdown(valor)
                            elif isinstance(valor, dict):
                                st.markdown(f"**{campo.replace('_', ' ').title()}:**")
                                st.json(valor, expanded=False)
                            elif isinstance(valor, list) and valor:
                                st.markdown(f"**{campo.replace('_', ' ').title()}:**")
                                if isinstance(valor[0], dict):
                                    st.dataframe(valor, use_container_width=True, hide_index=True)
                                else:
                                    for item in valor:
                                        st.markdown(f"- {item}")

                        if flags:
                            st.markdown("**Flags de Atenção:**")
                            for flag in flags:
                                st.warning(flag)

                # Checklist de Lacunas
                checklist = analise.get("checklist_lacunas", {})
                docs_faltantes = checklist.get("documentos_faltantes", [])
                info_pendentes = checklist.get("informacoes_pendentes", [])
                total_pend = checklist.get("total_pendencias", 0)
                total_crit = checklist.get("total_criticas", 0)

                if docs_faltantes or info_pendentes:
                    st.markdown("---")
                    st.markdown(f"### Checklist de Lacunas — {total_pend} pendência(s), {total_crit} crítica(s)")

                    if docs_faltantes:
                        st.markdown("#### Documentos Faltantes")
                        df_docs = []
                        for doc in docs_faltantes:
                            df_docs.append({
                                "Item": doc.get("item", ""),
                                "Criticidade": doc.get("criticidade", ""),
                                "Motivo": doc.get("motivo", ""),
                            })
                        st.dataframe(df_docs, use_container_width=True, hide_index=True)

                    if info_pendentes:
                        st.markdown("#### Informações Pendentes")
                        df_info = []
                        for info in info_pendentes:
                            df_info.append({
                                "Item": info.get("item", ""),
                                "Criticidade": info.get("criticidade", ""),
                                "Motivo": info.get("motivo", ""),
                            })
                        st.dataframe(df_info, use_container_width=True, hide_index=True)

    # ------------------------------------------------------------------
    # TAB 4 — MAC .docx
    # ------------------------------------------------------------------
    with tab_mac:
        st.markdown("### Geração do MAC (.docx)")

        if not st.session_state.analysis:
            st.info("Execute a análise de crédito na aba **Análise** antes de gerar o MAC.")
        else:
            analise = st.session_state.analysis
            op = st.session_state.current_op

            st.markdown("#### Resumo da Análise")
            summary = get_analysis_summary(analise)
            st.code(summary, language=None)

            if st.button("Gerar MAC .docx", use_container_width=True, type="primary"):
                try:
                    tomador_clean = (op.get("tomador", "operacao") or "operacao").replace(" ", "_").replace("/", "-")
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"MAC_{tomador_clean}_{timestamp}.docx"
                    output_path = str(OUTPUT_DIR / filename)

                    with st.spinner("Gerando documento MAC..."):
                        generated_path = generate_mac(analise, op, output_path)

                    st.success(f"MAC gerado com sucesso: **{filename}**")

                    with open(generated_path, "rb") as f:
                        docx_bytes = f.read()

                    st.download_button(
                        label="📥 Baixar MAC .docx",
                        data=docx_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )

                    # Update operation as concluded
                    op["status"] = "Concluída"
                    st.session_state.current_op = op

                    # Add to operations list if not already there
                    already_registered = any(
                        o.get("tomador") == op.get("tomador") and o.get("data_criacao") == op.get("data_criacao")
                        for o in st.session_state.operacoes
                    )
                    if not already_registered:
                        st.session_state.operacoes.append(op)

                except Exception as e:
                    st.error(f"Erro ao gerar o MAC: {e}")


def page_checklist_dd():
    st.markdown(
        """
        <div style="margin-bottom:20px;">
            <h1 style="color:#223040; font-size:1.8rem; font-weight:800; margin:0;">Checklist Due Diligence</h1>
            <p style="color:#8B9197; font-size:0.95rem; margin:4px 0 0 0;">
                12 módulos &nbsp;·&nbsp; Acompanhamento documental da operação
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Initialize DD status from session state
    if not st.session_state.dd_status:
        for modulo, itens in DD_CHECKLIST_TEMPLATE.items():
            st.session_state.dd_status[modulo] = {}
            for item in itens:
                st.session_state.dd_status[modulo][item] = "PENDENTE"

    # Auto-populate based on extracted data
    tipos_detectados = _detected_doc_types()
    tipo_to_module_items: dict[str, list[tuple[str, str]]] = {
        "balanco": [("3. Financeiro", "Balanço Patrimonial (último exercício auditado)")],
        "dre": [("3. Financeiro", "DRE (último exercício)")],
        "balancete": [("3. Financeiro", "Balancete acumulado do exercício corrente")],
        "matricula": [("10. Garantias", "Matrícula atualizada do imóvel (< 30 dias)")],
        "contrato": [("2. Societário", "Contrato Social consolidado ou Estatuto Social")],
        "certidao": [
            ("9. Certidões", "CND Federal (Receita Federal / PGFN)"),
        ],
        "ccir_car": [
            ("10. Garantias", "CCIR quitado (imóvel rural)"),
            ("10. Garantias", "CAR — Cadastro Ambiental Rural"),
        ],
    }
    for tipo in tipos_detectados:
        mappings = tipo_to_module_items.get(tipo, [])
        for modulo, item in mappings:
            if modulo in st.session_state.dd_status and item in st.session_state.dd_status[modulo]:
                st.session_state.dd_status[modulo][item] = "OK"

    # Dashboard summary
    total_items = 0
    total_ok = 0
    total_pendente = 0
    total_desatualizado = 0

    for modulo, itens in st.session_state.dd_status.items():
        for item, status in itens.items():
            total_items += 1
            if status == "OK":
                total_ok += 1
            elif status == "PENDENTE":
                total_pendente += 1
            elif status == "DESATUALIZADO":
                total_desatualizado += 1

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Itens", total_items)
    col2.metric("OK", total_ok)
    col3.metric("Pendente", total_pendente)
    col4.metric("Desatualizado", total_desatualizado)

    if total_items > 0:
        overall_pct = total_ok / total_items
        st.markdown("**Progresso Geral**")
        st.progress(overall_pct, text=f"{overall_pct:.0%} concluído ({total_ok}/{total_items})")

    st.markdown("---")

    # Current operation context
    if st.session_state.current_op:
        st.markdown(f"**Operação:** {st.session_state.current_op.get('tomador', 'N/A')} — {st.session_state.current_op.get('tipo_operacao', 'N/A')}")
        st.markdown("---")

    # Module-by-module display
    status_options = ["OK", "PENDENTE", "DESATUALIZADO"]
    status_icons = {"OK": "✅", "PENDENTE": "⏳", "DESATUALIZADO": "⚠️"}

    for modulo, itens in DD_CHECKLIST_TEMPLATE.items():
        modulo_status = st.session_state.dd_status.get(modulo, {})
        ok_count = sum(1 for s in modulo_status.values() if s == "OK")
        total_mod = len(itens)
        pct = ok_count / total_mod if total_mod > 0 else 0

        with st.expander(f"**{modulo}** — {ok_count}/{total_mod} ({pct:.0%})", expanded=False):
            st.progress(pct)

            for item in itens:
                current_status = st.session_state.dd_status.get(modulo, {}).get(item, "PENDENTE")
                col_item, col_status = st.columns([3, 1])

                with col_item:
                    icon = status_icons.get(current_status, "⏳")
                    st.markdown(f"{icon} {item}")

                with col_status:
                    new_status = st.selectbox(
                        "Status",
                        status_options,
                        index=status_options.index(current_status),
                        key=f"dd_{modulo}_{item}",
                        label_visibility="collapsed",
                    )
                    if new_status != current_status:
                        st.session_state.dd_status[modulo][item] = new_status
                        st.rerun()

    # Export
    st.markdown("---")
    if st.button("Exportar Checklist", use_container_width=True):
        linhas = [
            "CHECKLIST DUE DILIGENCE — ZYN Capital",
            f"Data: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        ]
        if st.session_state.current_op:
            linhas.append(f"Operação: {st.session_state.current_op.get('tomador', 'N/A')}")
        linhas.append("")
        linhas.append(f"Total: {total_items} | OK: {total_ok} | Pendente: {total_pendente} | Desatualizado: {total_desatualizado}")
        linhas.append(f"Progresso: {total_ok}/{total_items} ({total_ok/total_items*100:.0f}%)" if total_items > 0 else "Progresso: 0/0")
        linhas.append("=" * 60)

        for modulo, itens in DD_CHECKLIST_TEMPLATE.items():
            linhas.append("")
            linhas.append(f"--- {modulo} ---")
            for item in itens:
                status = st.session_state.dd_status.get(modulo, {}).get(item, "PENDENTE")
                icon = {"OK": "[OK]", "PENDENTE": "[PEND]", "DESATUALIZADO": "[DESAT]"}.get(status, "[PEND]")
                linhas.append(f"  {icon} {item}")

        export_text = "\n".join(linhas)
        st.download_button(
            label="📥 Baixar Checklist (.txt)",
            data=export_text.encode("utf-8"),
            file_name=f"checklist_dd_{datetime.now().strftime('%Y%m%d')}.txt",
            mime="text/plain",
            use_container_width=True,
        )
        st.success("Checklist gerado para download.")


# ---------------------------------------------------------------------------
# Main Layout
# ---------------------------------------------------------------------------
def main():
    st.set_page_config(
        page_title="ZYN Credit Engine",
        page_icon="🏦",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.markdown(
            """
            <div style="padding:8px 0 4px 0;">
                <p style="color:#FFFFFF; font-size:1.5rem; font-weight:800; letter-spacing:-0.3px; margin:0;">
                    ZYN Capital
                </p>
                <p style="color:rgba(255,255,255,0.45); font-size:0.78rem; margin:2px 0 0 0; letter-spacing:0.5px;">
                    CRÉDITO ESTRUTURADO & M&A
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("---")

        st.markdown(
            '<p style="color:rgba(255,255,255,0.35); font-size:0.65rem; text-transform:uppercase; '
            'letter-spacing:1.5px; margin:0 0 4px 4px;">Navegação</p>',
            unsafe_allow_html=True,
        )

        page = st.radio(
            "Navegação",
            ["Dashboard", "Nova Análise", "Checklist DD"],
            label_visibility="collapsed",
        )

        st.markdown("---")

        # Quick stats in sidebar
        total = len(st.session_state.operacoes)
        em_andamento = _count_ops_by_status("Em Andamento")
        st.markdown(
            f"""
            <div style="padding:4px 8px;">
                <p style="color:rgba(255,255,255,0.35); font-size:0.65rem; text-transform:uppercase;
                letter-spacing:1.5px; margin:0 0 8px 0;">Resumo</p>
                <div style="display:flex; justify-content:space-between; margin-bottom:6px;">
                    <span style="color:rgba(255,255,255,0.6); font-size:0.8rem;">Operações</span>
                    <span style="color:#FFFFFF; font-weight:700; font-size:0.8rem;">{total}</span>
                </div>
                <div style="display:flex; justify-content:space-between;">
                    <span style="color:rgba(255,255,255,0.6); font-size:0.8rem;">Em andamento</span>
                    <span style="color:#FFFFFF; font-weight:700; font-size:0.8rem;">{em_andamento}</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown("---")
        st.markdown(
            '<p class="footer-text">ZYN Credit Engine v1.0</p>',
            unsafe_allow_html=True,
        )

    # Page routing
    if page == "Dashboard":
        page_dashboard()
    elif page == "Nova Análise":
        page_nova_analise()
    elif page == "Checklist DD":
        page_checklist_dd()


if __name__ == "__main__":
    main()
