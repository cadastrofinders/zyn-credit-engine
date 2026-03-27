"""
ZYN Capital — Investor Matching Engine v2
Sugere investidores com base no perfil da operacao.

3 fontes de dados integradas:
1. Pipeline ZYN (19 parceiros diretos — Notion)
2. CVM (231 gestoras — investor_profiles.csv do Sales Intelligence)
3. Family Offices / Tesourarias / Seguradoras / Previdencia (19 — base manual)

Scoring methodology (alinhado ao Sales Intelligence):
- Instrumento match: 40%
- Volume/ticket: 25%
- Prazo: 15%
- Indexador: 10%
- Track record: 10%
"""

import csv
import json
import logging
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

# Caminhos dos dados do Sales Intelligence
# Try multiple paths for CVM data (local dev vs Streamlit Cloud)
_local_intel = Path(__file__).resolve().parent.parent.parent / "zyn-sales-intelligence" / "data"
_repo_intel = Path(__file__).resolve().parent.parent / "data"
SALES_INTEL_DIR = _local_intel if _local_intel.exists() else _repo_intel
CVM_PROFILES_PATH = SALES_INTEL_DIR / "investor_profiles.csv"
FO_PATH = SALES_INTEL_DIR / "family_offices.json"

# ---------------------------------------------------------------------------
# 1. Parceiros diretos ZYN (Pipeline Notion — campo "Analisando")
# ---------------------------------------------------------------------------
ZYN_PARTNERS: list[dict[str, Any]] = [
    {"nome": "GCB", "deals_pipeline": 12, "setores": ["agro", "imobiliario", "industria", "servicos", "fidc", "special_sits"], "instrumentos": ["CRI", "CRA", "CPR-F", "NC/CCB", "FIDC", "SLB", "Debenture"], "volume_min": 10_000_000, "volume_max": 200_000_000, "perfil": "Generalista, maior parceiro ZYN. Apetite amplo.", "fonte": "pipeline_zyn"},
    {"nome": "Inco", "deals_pipeline": 7, "setores": ["imobiliario"], "instrumentos": ["CRI"], "volume_min": 10_000_000, "volume_max": 100_000_000, "perfil": "Foco exclusivo em CRI/imobiliario.", "fonte": "pipeline_zyn"},
    {"nome": "Fegik", "deals_pipeline": 4, "setores": ["imobiliario", "industria"], "instrumentos": ["CRI"], "volume_min": 10_000_000, "volume_max": 80_000_000, "perfil": "CRI com foco em imobiliario e industria.", "fonte": "pipeline_zyn"},
    {"nome": "Vinci", "deals_pipeline": 4, "setores": ["agro"], "instrumentos": ["CRA", "CPR-F", "SLB"], "volume_min": 20_000_000, "volume_max": 200_000_000, "perfil": "Agro: SLB rural, CPR-F, CRA. Tickets maiores.", "fonte": "pipeline_zyn"},
    {"nome": "Jive", "deals_pipeline": 4, "setores": ["agro", "imobiliario", "industria", "special_sits"], "instrumentos": ["CRI", "CRA", "NC/CCB", "NPL"], "volume_min": 10_000_000, "volume_max": 150_000_000, "perfil": "Diversificado, aceita special situations e NPL.", "fonte": "pipeline_zyn"},
    {"nome": "exes", "deals_pipeline": 3, "setores": ["imobiliario"], "instrumentos": ["CRI"], "volume_min": 10_000_000, "volume_max": 80_000_000, "perfil": "CRI/imobiliario.", "fonte": "pipeline_zyn"},
    {"nome": "BS2", "deals_pipeline": 3, "setores": ["industria", "servicos"], "instrumentos": ["NC/CCB", "Debenture"], "volume_min": 10_000_000, "volume_max": 100_000_000, "perfil": "DCM/CCB, credito corporativo.", "fonte": "pipeline_zyn"},
    {"nome": "Bside", "deals_pipeline": 3, "setores": ["imobiliario", "industria"], "instrumentos": ["CRI"], "volume_min": 10_000_000, "volume_max": 80_000_000, "perfil": "CRI com apetite medio.", "fonte": "pipeline_zyn"},
    {"nome": "Galapagos", "deals_pipeline": 3, "setores": ["agro", "imobiliario"], "instrumentos": ["CRI", "CRA", "CPR-F"], "volume_min": 15_000_000, "volume_max": 100_000_000, "perfil": "CRI e Agro, perfil diversificado.", "fonte": "pipeline_zyn"},
    {"nome": "Pine", "deals_pipeline": 2, "setores": ["agro"], "instrumentos": ["CPR-F", "CRA"], "volume_min": 20_000_000, "volume_max": 100_000_000, "perfil": "Agro: CPR-F e CRA.", "fonte": "pipeline_zyn"},
    {"nome": "Fibra", "deals_pipeline": 2, "setores": ["agro"], "instrumentos": ["CPR-F", "CRA", "SLB"], "volume_min": 15_000_000, "volume_max": 80_000_000, "perfil": "Agro com foco em CPR-F.", "fonte": "pipeline_zyn"},
    {"nome": "Kijani", "deals_pipeline": 1, "setores": ["agro"], "instrumentos": ["CRA", "CPR-F", "Fiagro"], "volume_min": 10_000_000, "volume_max": 80_000_000, "perfil": "Agro/Fiagro, perfil ESG.", "fonte": "pipeline_zyn"},
    {"nome": "EXT Capital", "deals_pipeline": 2, "setores": ["fidc"], "instrumentos": ["FIDC", "Fiagro"], "volume_min": 20_000_000, "volume_max": 100_000_000, "perfil": "FIDC e Fiagro especializado.", "fonte": "pipeline_zyn"},
    {"nome": "IORQ / Luso Brasileiro", "deals_pipeline": 2, "setores": ["industria", "servicos"], "instrumentos": ["NC/CCB", "Debenture"], "volume_min": 10_000_000, "volume_max": 80_000_000, "perfil": "Credito bancario corporativo.", "fonte": "pipeline_zyn"},
    {"nome": "Solis", "deals_pipeline": 1, "setores": ["imobiliario", "industria"], "instrumentos": ["CRI", "CRA"], "volume_min": 10_000_000, "volume_max": 60_000_000, "perfil": "CRI e CRA, tickets menores.", "fonte": "pipeline_zyn"},
    {"nome": "Oliveira Trust", "deals_pipeline": 1, "setores": ["imobiliario", "fidc"], "instrumentos": ["CRI", "FIDC"], "volume_min": 15_000_000, "volume_max": 100_000_000, "perfil": "Securitizadora/trustee, CRI e FIDC.", "fonte": "pipeline_zyn"},
    {"nome": "Opea", "deals_pipeline": 1, "setores": ["imobiliario", "agro"], "instrumentos": ["CRI", "CRA"], "volume_min": 10_000_000, "volume_max": 80_000_000, "perfil": "Securitizadora CRI/CRA.", "fonte": "pipeline_zyn"},
    {"nome": "Virgo", "deals_pipeline": 1, "setores": ["imobiliario", "industria"], "instrumentos": ["CRI", "Debenture"], "volume_min": 15_000_000, "volume_max": 100_000_000, "perfil": "Plataforma de securitizacao.", "fonte": "pipeline_zyn"},
    {"nome": "True Securitizadora", "deals_pipeline": 1, "setores": ["imobiliario"], "instrumentos": ["CRI"], "volume_min": 10_000_000, "volume_max": 80_000_000, "perfil": "CRI imobiliario.", "fonte": "pipeline_zyn"},
]

# Manter referência para backward compatibility
INVESTOR_DATABASE = ZYN_PARTNERS


# ---------------------------------------------------------------------------
# 2. Loader: CVM investor profiles (do Sales Intelligence)
# ---------------------------------------------------------------------------
def _load_cvm_profiles() -> list[dict[str, Any]]:
    """Carrega perfis de gestoras CVM do investor_profiles.csv."""
    if not CVM_PROFILES_PATH.exists():
        logger.info("CVM profiles not found at %s — skipping CVM data.", CVM_PROFILES_PATH)
        return []

    investors = []
    try:
        with open(CVM_PROFILES_PATH, encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                vol_total = float(row.get("vol_total", 0) or 0)
                if vol_total < 1_000_000:  # Ignorar gestoras com < R$1MM em RF estruturada
                    continue

                inv = {
                    "nome": row.get("gestora", "").strip(),
                    "cnpj": row.get("cnpj_gestora", ""),
                    "n_fundos": int(float(row.get("n_fundos", 0) or 0)),
                    "pl_total": float(row.get("pl_total", 0) or 0),
                    "vol_total": vol_total,
                    "vol_NC": float(row.get("vol_NC", 0) or 0),
                    "vol_CRI": float(row.get("vol_CRI", 0) or 0),
                    "vol_CRA": float(row.get("vol_CRA", 0) or 0),
                    "vol_CPR-F": float(row.get("vol_CPR-F", 0) or 0),
                    "vol_DEBENTURE": float(row.get("vol_DEBENTURE", 0) or 0),
                    "n_ops_NC": int(float(row.get("n_ops_NC", 0) or 0)),
                    "n_ops_CRI": int(float(row.get("n_ops_CRI", 0) or 0)),
                    "n_ops_CRA": int(float(row.get("n_ops_CRA", 0) or 0)),
                    "n_ops_CPR-F": int(float(row.get("n_ops_CPR-F", 0) or 0)),
                    "n_ops_DEBENTURE": int(float(row.get("n_ops_DEBENTURE", 0) or 0)),
                    "tipo_preferido": row.get("tipo_preferido", ""),
                    "ticket_medio": float(row.get("ticket_medio", 0) or 0),
                    "ticket_mediano": float(row.get("ticket_mediano", 0) or 0),
                    "ticket_max": float(row.get("ticket_max", 0) or 0),
                    "prazo_medio_anos": float(row.get("prazo_medio_anos", 0) or 0) if row.get("prazo_medio_anos") else None,
                    "indexador_principal": row.get("indexador_principal", ""),
                    "classe_predominante": row.get("classe_predominante", ""),
                    "fonte": "cvm",
                }
                investors.append(inv)
    except Exception as e:
        logger.warning("Error loading CVM profiles: %s", e)

    logger.info("Loaded %d CVM investor profiles.", len(investors))
    return investors


def _load_family_offices() -> list[dict[str, Any]]:
    """Carrega base de Family Offices / Tesourarias do Sales Intelligence."""
    if not FO_PATH.exists():
        logger.info("Family offices file not found at %s — skipping.", FO_PATH)
        return []

    try:
        with open(FO_PATH, encoding="utf-8") as f:
            raw = json.load(f)
        investors = []
        for fo in raw:
            inv = {
                "nome": fo.get("nome", ""),
                "tipo_investidor": fo.get("tipo", ""),
                "instrumentos": fo.get("apetite", []),
                "ticket_min": fo.get("ticket_min", 0),
                "ticket_max": fo.get("ticket_max", 0),
                "indexador_pref": fo.get("indexador_pref", ""),
                "fonte": "family_office",
            }
            investors.append(inv)
        logger.info("Loaded %d family offices/tesourarias.", len(investors))
        return investors
    except Exception as e:
        logger.warning("Error loading family offices: %s", e)
        return []


# ---------------------------------------------------------------------------
# Mapeamento tipo_operacao → instrumento CVM
# ---------------------------------------------------------------------------
_TIPO_TO_CVM = {
    "CRI": "CRI",
    "CRA": "CRA",
    "CPR-F": "CPR-F",
    "Fiagro": "CRA",
    "SLB": "CRA",
    "NC/CCB": "NC",
    "Debenture": "DEBENTURE",
    "FIDC": None,  # FIDC não aparece no CDA
    "Compra de Estoque": "CRI",
    "Precatorios": None,
    "NPL": None,
    "SCP": None,
}

INSTRUMENTO_TO_SETORES: dict[str, list[str]] = {
    "CRI": ["imobiliario"],
    "CRA": ["agro"],
    "CPR-F": ["agro"],
    "Fiagro": ["agro"],
    "SLB": ["agro"],
    "FIDC": ["fidc"],
    "NC/CCB": ["industria", "servicos"],
    "Debenture": ["industria", "servicos"],
}


def _detect_sector(tipo_operacao: str, garantias: list[str] | None = None) -> str:
    """Detecta setor provavel com base no instrumento e garantias."""
    tipo_upper = tipo_operacao.upper().strip()
    if tipo_upper in ("CRI",):
        return "imobiliario"
    if tipo_upper in ("CRA", "CPR-F", "FIAGRO", "SLB"):
        return "agro"
    if tipo_upper in ("FIDC",):
        return "fidc"
    if tipo_upper in ("PRECATORIOS", "NPL", "SCP"):
        return "special_sits"

    if garantias:
        garantias_lower = " ".join(garantias).lower()
        if any(k in garantias_lower for k in ["imovel", "imóvel", "matricula", "matrícula", "terreno", "apartamento"]):
            return "imobiliario"
        if any(k in garantias_lower for k in ["safra", "rural", "fazenda", "grao", "grão", "soja", "milho", "boi"]):
            return "agro"
    return "industria"


# ---------------------------------------------------------------------------
# Scoring — aligned with Sales Intelligence methodology
# ---------------------------------------------------------------------------
def _score_cvm_investor(deal_tipo: str, deal_volume: float, deal_prazo_anos: float | None,
                        deal_indexador: str | None, inv: dict) -> dict[str, Any]:
    """Score a CVM investor against a deal. Returns score breakdown."""
    cvm_tipo = _TIPO_TO_CVM.get(deal_tipo)
    if not cvm_tipo:
        return {"score_total": 0}

    scores = {}

    # 1. Tipo match (40%)
    vol_col = f"vol_{cvm_tipo}"
    n_ops_col = f"n_ops_{cvm_tipo}"
    vol_in_type = inv.get(vol_col, 0)
    n_ops = inv.get(n_ops_col, 0)

    if vol_in_type > 0:
        scores["tipo"] = 1.0
    elif inv.get("tipo_preferido", "") == cvm_tipo:
        scores["tipo"] = 0.8
    else:
        scores["tipo"] = 0.2 if inv.get("vol_total", 0) > 0 else 0.0

    # 2. Volume/ticket match (25%)
    ticket_med = inv.get("ticket_medio", 0) or 0
    if deal_volume > 0 and ticket_med > 0:
        ratio = deal_volume / ticket_med
        if 0.3 <= ratio <= 3.0:
            scores["volume"] = 1.0
        elif 0.1 <= ratio <= 5.0:
            scores["volume"] = 0.6
        else:
            scores["volume"] = 0.2
    else:
        scores["volume"] = 0.5

    # 3. Prazo match (15%)
    inv_prazo = inv.get("prazo_medio_anos")
    if deal_prazo_anos and inv_prazo:
        diff = abs(deal_prazo_anos - inv_prazo)
        if diff <= 1:
            scores["prazo"] = 1.0
        elif diff <= 3:
            scores["prazo"] = 0.6
        else:
            scores["prazo"] = 0.2
    else:
        scores["prazo"] = 0.5

    # 4. Indexador match (10%)
    inv_idx = str(inv.get("indexador_principal", "")).upper()
    deal_idx = (deal_indexador or "").upper()
    if deal_idx and inv_idx:
        if deal_idx in inv_idx or inv_idx in deal_idx:
            scores["indexador"] = 1.0
        else:
            scores["indexador"] = 0.3
    else:
        scores["indexador"] = 0.5

    # 5. Track record (10%)
    if n_ops >= 10:
        scores["historico"] = 1.0
    elif n_ops >= 5:
        scores["historico"] = 0.8
    elif n_ops >= 1:
        scores["historico"] = 0.5
    else:
        scores["historico"] = 0.1

    # Weighted total
    weights = {"tipo": 0.40, "volume": 0.25, "prazo": 0.15, "indexador": 0.10, "historico": 0.10}
    total = sum(scores.get(k, 0) * weights[k] for k in weights)

    return {
        "score_total": round(total * 100, 1),  # Scale to 0-100
        "scores": scores,
    }


def _score_fo_investor(deal_tipo: str, deal_volume: float, inv: dict) -> dict[str, Any]:
    """Score a Family Office / Tesouraria investor."""
    cvm_tipo = _TIPO_TO_CVM.get(deal_tipo, deal_tipo)
    apetite = [a.upper() for a in inv.get("instrumentos", [])]

    scores = {}

    # Tipo match
    if cvm_tipo and cvm_tipo.upper() in apetite:
        scores["tipo"] = 1.0
    else:
        scores["tipo"] = 0.1

    # Volume match
    tmin = inv.get("ticket_min", 0)
    tmax = inv.get("ticket_max", float("inf"))
    if tmin <= deal_volume <= tmax:
        scores["volume"] = 1.0
    elif deal_volume > 0 and deal_volume >= tmin * 0.5:
        scores["volume"] = 0.5
    else:
        scores["volume"] = 0.2

    # Indexador match
    idx_pref = inv.get("indexador_pref", "")
    scores["indexador"] = 0.5  # no deal indexador info in basic match

    total = scores["tipo"] * 0.50 + scores["volume"] * 0.35 + scores["indexador"] * 0.15

    return {"score_total": round(total * 100, 1), "scores": scores}


def _score_zyn_partner(deal_tipo: str, deal_volume: float, setor: str, inv: dict) -> dict[str, Any]:
    """Score a ZYN Pipeline partner."""
    scores = {}
    tipo_norm = deal_tipo.strip()

    # Instrumento match
    if tipo_norm in inv.get("instrumentos", []):
        scores["tipo"] = 1.0
    else:
        tipo_family = {"CRI": "securitizacao", "CRA": "securitizacao", "CPR-F": "agro",
                       "Fiagro": "agro", "SLB": "agro", "FIDC": "fundo",
                       "NC/CCB": "credito", "Debenture": "credito"}
        inv_families = {tipo_family.get(i, "") for i in inv.get("instrumentos", [])}
        if tipo_family.get(tipo_norm, "x") in inv_families:
            scores["tipo"] = 0.6
        else:
            scores["tipo"] = 0.1

    # Setor match
    if setor in inv.get("setores", []):
        scores["setor"] = 1.0
    else:
        scores["setor"] = 0.2

    # Volume match
    vmin = inv.get("volume_min", 0)
    vmax = inv.get("volume_max", float("inf"))
    if vmin <= deal_volume <= vmax:
        scores["volume"] = 1.0
    elif deal_volume >= vmin * 0.5:
        scores["volume"] = 0.5
    else:
        scores["volume"] = 0.2

    # Track record (deals no pipeline ZYN)
    deals = inv.get("deals_pipeline", 0)
    if deals >= 5:
        scores["historico"] = 1.0
    elif deals >= 3:
        scores["historico"] = 0.8
    elif deals >= 1:
        scores["historico"] = 0.5
    else:
        scores["historico"] = 0.2

    total = scores["tipo"] * 0.35 + scores["setor"] * 0.25 + scores["volume"] * 0.20 + scores["historico"] * 0.20
    # Bonus for ZYN relationship (cap at 0.95 to allow differentiation)
    total = min(total * 1.10 + 0.05, 0.95)  # 10% bonus + 5pts floor for being a ZYN partner

    return {"score_total": round(total * 100, 1), "scores": scores}


# ---------------------------------------------------------------------------
# Main matching function
# ---------------------------------------------------------------------------
def match_investors(
    tipo_operacao: str,
    volume: float,
    setor: str | None = None,
    garantias: list[str] | None = None,
    rating: str | None = None,
    prazo_anos: float | None = None,
    indexador: str | None = None,
    top_n: int = 15,
    include_cvm: bool = True,
    include_fo: bool = True,
) -> list[dict[str, Any]]:
    """
    Retorna ranking unificado de investidores.
    Combina: ZYN Partners (Pipeline) + CVM (gestoras) + Family Offices.

    Returns list of dicts: {nome, score, fonte, motivos[], perfil/detalhes, ...}
    """
    if not setor:
        setor = _detect_sector(tipo_operacao, garantias)

    results: list[dict[str, Any]] = []

    # --- ZYN Partners ---
    for inv in ZYN_PARTNERS:
        sc = _score_zyn_partner(tipo_operacao, volume, setor, inv)
        if sc["score_total"] >= 15:
            motivos = []
            if sc["scores"].get("tipo", 0) >= 0.6:
                motivos.append(f"Opera {tipo_operacao}")
            if sc["scores"].get("setor", 0) >= 0.8:
                motivos.append(f"Atua em {setor}")
            if sc["scores"].get("volume", 0) >= 0.8:
                motivos.append(f"Volume dentro da faixa")
            if inv.get("deals_pipeline", 0) >= 3:
                motivos.append(f"{inv['deals_pipeline']} deals ZYN")

            results.append({
                "nome": inv["nome"],
                "score": sc["score_total"],
                "fonte": "Pipeline ZYN",
                "motivos": motivos,
                "perfil": inv.get("perfil", ""),
                "deals_pipeline": inv.get("deals_pipeline", 0),
                "instrumentos": inv.get("instrumentos", []),
                "setores": inv.get("setores", []),
                "tipo_investidor": "Parceiro ZYN",
            })

    # --- CVM Investors ---
    if include_cvm:
        cvm_investors = _load_cvm_profiles()
        for inv in cvm_investors:
            sc = _score_cvm_investor(tipo_operacao, volume, prazo_anos, indexador, inv)
            if sc["score_total"] >= 30:
                motivos = []
                cvm_tipo = _TIPO_TO_CVM.get(tipo_operacao, "")
                vol_key = f"vol_{cvm_tipo}"
                n_ops_key = f"n_ops_{cvm_tipo}"
                vol_in_type = inv.get(vol_key, 0)
                n_ops = inv.get(n_ops_key, 0)

                if vol_in_type > 0:
                    motivos.append(f"R${vol_in_type/1e6:.0f}MM em {cvm_tipo}")
                if n_ops >= 5:
                    motivos.append(f"{n_ops} operacoes {cvm_tipo}")
                if inv.get("pl_total", 0) > 1e9:
                    motivos.append(f"PL R${inv['pl_total']/1e9:.1f}B")
                if sc["scores"].get("volume", 0) >= 0.8:
                    motivos.append("Ticket compativel")

                nome_curto = inv["nome"]
                if len(nome_curto) > 50:
                    nome_curto = nome_curto[:47] + "..."

                results.append({
                    "nome": nome_curto,
                    "nome_completo": inv["nome"],
                    "score": sc["score_total"],
                    "fonte": "CVM",
                    "motivos": motivos,
                    "perfil": f"{inv.get('n_fundos',0)} fundos | PL R${inv.get('pl_total',0)/1e9:.1f}B | Pref: {inv.get('tipo_preferido','')}",
                    "n_fundos": inv.get("n_fundos", 0),
                    "pl_total": inv.get("pl_total", 0),
                    "vol_total_rf": inv.get("vol_total", 0),
                    "ticket_medio": inv.get("ticket_medio", 0),
                    "tipo_preferido": inv.get("tipo_preferido", ""),
                    "indexador_principal": inv.get("indexador_principal", ""),
                    "tipo_investidor": "Gestora CVM",
                })

    # --- Family Offices / Tesourarias ---
    if include_fo:
        fos = _load_family_offices()
        for inv in fos:
            sc = _score_fo_investor(tipo_operacao, volume, inv)
            if sc["score_total"] >= 25:
                motivos = []
                if sc["scores"].get("tipo", 0) >= 0.8:
                    cvm_tipo = _TIPO_TO_CVM.get(tipo_operacao, tipo_operacao)
                    motivos.append(f"Apetite para {cvm_tipo}")
                if sc["scores"].get("volume", 0) >= 0.8:
                    motivos.append(f"Ticket R${inv.get('ticket_min',0)/1e6:.0f}-{inv.get('ticket_max',0)/1e6:.0f}MM")
                if inv.get("indexador_pref"):
                    motivos.append(f"Pref: {inv['indexador_pref']}")

                results.append({
                    "nome": inv["nome"],
                    "score": sc["score_total"],
                    "fonte": inv.get("tipo_investidor", "Family Office"),
                    "motivos": motivos,
                    "perfil": f"{inv.get('tipo_investidor', '')} | Ticket R${inv.get('ticket_min',0)/1e6:.0f}-{inv.get('ticket_max',0)/1e6:.0f}MM",
                    "tipo_investidor": inv.get("tipo_investidor", ""),
                    "instrumentos": inv.get("instrumentos", []),
                })

    # Sort by score desc
    results.sort(key=lambda x: -x["score"])
    return results[:top_n]


# ---------------------------------------------------------------------------
# Sector benchmarks (unchanged from v1)
# ---------------------------------------------------------------------------
def get_sector_benchmarks(setor: str) -> dict[str, Any]:
    """Retorna benchmarks setoriais calibrados pelo mercado brasileiro."""
    benchmarks = {
        "agro": {
            "setor": "Agronegocio",
            "margem_ebitda_media": 0.25,
            "divida_ebitda_media": 2.5,
            "dscr_medio": 1.5,
            "roe_medio": 0.15,
            "ltv_maximo_recomendado": 0.65,
            "volume_mediano": 40_000_000,
            "prazo_tipico_meses": "24-60",
            "instrumentos_comuns": ["CPR-F", "CRA", "SLB", "Fiagro"],
            "garantias_tipicas": ["Alienacao Fiduciaria de Imovel Rural", "Penhor de Safra", "Aval dos Socios", "CPR como lastro"],
            "riscos_especificos": [
                "Sazonalidade de safra e precos de commodities",
                "Risco climatico (seca, geada, excesso de chuvas)",
                "Variacao cambial (insumos dolarizados)",
                "Concentracao de receita em poucos compradores (tradings)",
            ],
            "notas": "Volume tipico ZYN: R$20-200MM (mediana ~R$40MM). Vinci/GCB sao investidores-chave. BB Gestao e Itau Asset sao maiores compradores CVM.",
        },
        "imobiliario": {
            "setor": "Imobiliario",
            "margem_ebitda_media": 0.20,
            "divida_ebitda_media": 3.0,
            "dscr_medio": 1.3,
            "roe_medio": 0.12,
            "ltv_maximo_recomendado": 0.70,
            "volume_mediano": 35_000_000,
            "prazo_tipico_meses": "12-36",
            "instrumentos_comuns": ["CRI", "Compra de Estoque"],
            "garantias_tipicas": ["Alienacao Fiduciaria de Imovel", "Cessao Fiduciaria de Recebiveis", "Fianca/Aval", "Fundo de Reserva"],
            "riscos_especificos": [
                "Velocidade de vendas (VSO) abaixo do projetado",
                "Risco de obra / entrega / habite-se",
                "Concentracao geografica do empreendimento",
                "Inadimplencia dos compradores finais",
            ],
            "notas": "Volume tipico ZYN: R$14-71MM. Inco/Fegik sao parceiros-chave. Maua Capital, Kinea e Capitania sao top CRI no CVM.",
        },
        "industria": {
            "setor": "Industria e Servicos",
            "margem_ebitda_media": 0.15,
            "divida_ebitda_media": 2.8,
            "dscr_medio": 1.4,
            "roe_medio": 0.13,
            "ltv_maximo_recomendado": 0.60,
            "volume_mediano": 30_000_000,
            "prazo_tipico_meses": "12-48",
            "instrumentos_comuns": ["NC/CCB", "Debenture", "CRI"],
            "garantias_tipicas": ["Alienacao Fiduciaria de Imovel", "Alienacao Fiduciaria de Maquinas", "Aval dos Socios", "Cessao de Recebiveis"],
            "riscos_especificos": [
                "Utilizacao de capacidade instalada",
                "Concentracao de clientes",
                "Obsolescencia tecnologica",
                "Risco regulatorio setorial",
            ],
            "notas": "Volume tipico ZYN: R$15-150MM. BTG Asset (R$7.2B em NC), BB Gestao e Caixa DTVM sao maiores compradores NC no CVM.",
        },
        "fidc": {
            "setor": "FIDC / Securitizacao de Recebiveis",
            "margem_ebitda_media": None,
            "divida_ebitda_media": None,
            "dscr_medio": None,
            "roe_medio": 0.18,
            "ltv_maximo_recomendado": 0.80,
            "volume_mediano": 30_000_000,
            "prazo_tipico_meses": "12-36 (revolving)",
            "instrumentos_comuns": ["FIDC", "Cota FIDC"],
            "garantias_tipicas": ["Cessao Fiduciaria de Recebiveis", "Subordinacao (mezanino/equity)", "Excesso de Spread", "Coobrigacao do Cedente"],
            "riscos_especificos": [
                "Qualidade da carteira de recebiveis",
                "Concentracao de sacados",
                "Risco de diluicao",
                "Inadimplencia acima do PDD",
            ],
            "notas": "Volume tipico ZYN: R$30-35MM. EXT Capital e GCB sao parceiros. Mercado de FIDC não aparece diretamente no CDA/CVM.",
        },
        "special_sits": {
            "setor": "Special Situations",
            "margem_ebitda_media": None,
            "divida_ebitda_media": None,
            "dscr_medio": None,
            "roe_medio": None,
            "ltv_maximo_recomendado": 0.50,
            "volume_mediano": 15_000_000,
            "prazo_tipico_meses": "6-24",
            "instrumentos_comuns": ["NPL", "Precatorios", "SCP"],
            "garantias_tipicas": ["Cessao de Credito Judicial", "Garantia Real", "Aval"],
            "riscos_especificos": [
                "Risco juridico / timing de recebimento",
                "Iliquidez do ativo",
                "Desconto excessivo sobre valor de face",
                "Risco de contraparte judicial",
            ],
            "notas": "Operacoes de nicho. Jive e GCB sao investidores com apetite.",
        },
    }
    return benchmarks.get(setor, benchmarks["industria"])


# ---------------------------------------------------------------------------
# Product-specific DD checklists
# ---------------------------------------------------------------------------
PRODUCT_DD_EXTRAS: dict[str, dict[str, list[str]]] = {
    "CRI": {
        "Estrutura CRI": [
            "Termo de Securitizacao",
            "Instrumento de Emissao do CRI",
            "Contrato de Cessao Fiduciaria de Recebiveis Imobiliarios",
            "Parecer juridico sobre lastro imobiliario",
            "Rating da emissao (se houver)",
            "Demonstrativo de lastro vs emissao",
            "Contrato de Distribuicao",
        ],
    },
    "CRA": {
        "Estrutura CRA": [
            "Termo de Securitizacao do CRA",
            "CPR ou duplicata rural como lastro",
            "Contrato de compra e venda de produto agropecuario",
            "Parecer juridico sobre lastro agro",
            "Rating da emissao (se houver)",
            "Seguro agricola / PROAGRO",
            "Relatorio de safra / producao",
        ],
    },
    "CPR-F": {
        "Estrutura CPR-F": [
            "CPR Financeira registrada",
            "Contrato de compra garantida (offtaker)",
            "Seguro agricola",
            "Laudo de capacidade produtiva",
            "Historico de entrega / producao (3 safras)",
            "Contrato de armazenagem (se aplicavel)",
        ],
    },
    "FIDC": {
        "Estrutura FIDC": [
            "Regulamento do Fundo",
            "Politica de Credito e Elegibilidade",
            "Relatorio do Administrador",
            "Relatorio do Custodiante",
            "Composicao da carteira de recebiveis",
            "Historico de PDD e inadimplencia",
            "Rating das cotas (senior/mezanino)",
            "Demonstracoes financeiras do fundo",
            "Parecer juridico sobre cessao de recebiveis",
        ],
    },
    "Fiagro": {
        "Estrutura Fiagro": [
            "Regulamento do Fiagro",
            "Politica de investimento agro",
            "Lastro: CPR / CRA / recebiveis rurais",
            "Due diligence ambiental dos ativos",
            "Rating (se houver)",
            "Demonstracoes financeiras do fundo",
        ],
    },
    "SLB": {
        "Estrutura SLB": [
            "Contrato de Sale & Leaseback",
            "Laudo de avaliacao do(s) imovel(is) rural(is)",
            "Matricula(s) atualizada(s)",
            "CCIR e CAR quitados",
            "Contrato de arrendamento reverso",
            "Due diligence ambiental",
            "ITR quitado",
        ],
    },
    "NC/CCB": {
        "Estrutura NC/CCB": [
            "Nota Comercial ou CCB emitida",
            "Termo de Adesao / Escritura de Emissao",
            "Contrato de garantia (alienacao fiduciaria / penhor)",
            "Parecer juridico sobre a emissao",
            "Agenda de pagamentos",
        ],
    },
    "Debenture": {
        "Estrutura Debenture": [
            "Escritura de Emissao de Debentures",
            "Ata de assembleia aprovando a emissao",
            "Contrato de garantias acessorias",
            "Contrato de Agente Fiduciario",
            "Rating da emissao (se houver)",
            "Parecer juridico sobre a emissao",
        ],
    },
}

# ---------------------------------------------------------------------------
# Taxonomia de garantias (alinhada ao Pipeline Notion)
# ---------------------------------------------------------------------------
GUARANTEE_TYPES = {
    "Real": {
        "descricao": "Garantia real sobre bem movel ou imovel",
        "exemplos": ["Alienacao fiduciaria de imovel", "Hipoteca", "Alienacao fiduciaria de veiculos/maquinas"],
        "docs_necessarios": ["Matricula atualizada", "Laudo de avaliacao", "Certidao de onus reais"],
    },
    "Fiduciaria": {
        "descricao": "Cessao fiduciaria de direitos creditorios ou recebiveis",
        "exemplos": ["Cessao fiduciaria de recebiveis", "Cessao fiduciaria de alugueis", "Cessao de direitos sobre contratos"],
        "docs_necessarios": ["Contrato de cessao", "Demonstrativo de recebiveis cedidos", "Notificacao aos devedores"],
    },
    "Aval": {
        "descricao": "Garantia pessoal dos socios / controladores",
        "exemplos": ["Aval dos socios PF", "Aval de empresa do grupo"],
        "docs_necessarios": ["Declaracao de bens dos avalistas", "IR PF dos avalistas", "Certidoes negativas PF"],
    },
    "Coobrigacao": {
        "descricao": "Coobrigacao solidaria de terceiro ou empresa do grupo",
        "exemplos": ["Coobrigacao da holding", "Coobrigacao do cedente (FIDC)"],
        "docs_necessarios": ["Contrato de coobrigacao", "Demonstracoes financeiras do coobrigado"],
    },
    "Alienacao": {
        "descricao": "Alienacao fiduciaria especifica (imoveis, equipamentos, estoques)",
        "exemplos": ["Alienacao fiduciaria de imovel urbano", "Alienacao fiduciaria de imovel rural", "Alienacao fiduciaria de equipamentos"],
        "docs_necessarios": ["Matricula", "Laudo", "CCIR/CAR (rural)", "Contrato de alienacao registrado"],
    },
    "Penhor": {
        "descricao": "Penhor sobre bens moveis, safra, ou titulos",
        "exemplos": ["Penhor de safra", "Penhor de acoes", "Penhor de maquinario"],
        "docs_necessarios": ["Contrato de penhor registrado", "Inventario dos bens empenhados"],
    },
}
