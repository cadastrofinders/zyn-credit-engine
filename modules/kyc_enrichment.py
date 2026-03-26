"""
ZYN Capital — Enriquecimento KYC Automático
Busca dados públicos do CNPJ para complementar a análise de crédito.

Fontes:
- BrasilAPI (ReceitaWS) — dados cadastrais, QSA, situação
- Suframa, Simples Nacional — consulta auxiliar
"""

import json
import logging
import ssl
import time
from typing import Any
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

logger = logging.getLogger(__name__)

BRASIL_API_BASE = "https://brasilapi.com.br/api/cnpj/v1"
TIMEOUT = 15  # seconds


def _fetch_json(url: str) -> dict | None:
    """Fetch JSON from URL with basic error handling."""
    try:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        req = Request(url, headers={"User-Agent": "ZYN-CreditEngine/1.0"})
        with urlopen(req, timeout=TIMEOUT, context=ctx) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except (URLError, HTTPError, json.JSONDecodeError) as e:
        logger.warning("KYC fetch failed for %s: %s", url, e)
        return None


def _clean_cnpj(cnpj: str) -> str:
    """Remove formatting from CNPJ."""
    return "".join(c for c in cnpj if c.isdigit())


def fetch_cnpj_data(cnpj: str) -> dict[str, Any]:
    """
    Busca dados públicos do CNPJ via BrasilAPI.

    Returns dict with:
        - razao_social, nome_fantasia
        - situacao_cadastral, data_situacao
        - endereco completo
        - qsa (quadro societário)
        - cnae_fiscal, cnaes_secundarios
        - capital_social
        - porte, natureza_juridica
        - data_inicio_atividade
    """
    cnpj_limpo = _clean_cnpj(cnpj)
    if len(cnpj_limpo) != 14:
        return {"error": f"CNPJ inválido: {cnpj}"}

    data = _fetch_json(f"{BRASIL_API_BASE}/{cnpj_limpo}")
    if not data:
        return {"error": f"Não foi possível consultar CNPJ {cnpj}"}

    # Normalize QSA
    qsa = []
    for socio in data.get("qsa", []):
        qsa.append({
            "nome": socio.get("nome_socio", ""),
            "qualificacao": socio.get("qualificacao_socio", ""),
            "cpf_cnpj": socio.get("cnpj_cpf_do_socio", ""),
            "data_entrada": socio.get("data_entrada_sociedade", ""),
        })

    # CNAEs secundários
    cnaes_sec = []
    for cnae in data.get("cnaes_secundarios", []):
        if cnae.get("codigo"):
            cnaes_sec.append({
                "codigo": cnae.get("codigo"),
                "descricao": cnae.get("descricao", ""),
            })

    return {
        "cnpj": cnpj,
        "razao_social": data.get("razao_social", ""),
        "nome_fantasia": data.get("nome_fantasia", ""),
        "situacao_cadastral": data.get("descricao_situacao_cadastral", ""),
        "data_situacao_cadastral": data.get("data_situacao_cadastral", ""),
        "data_inicio_atividade": data.get("data_inicio_atividade", ""),
        "natureza_juridica": data.get("natureza_juridica", ""),
        "porte": data.get("porte", ""),
        "capital_social": data.get("capital_social", 0),
        "cnae_fiscal": {
            "codigo": data.get("cnae_fiscal", ""),
            "descricao": data.get("cnae_fiscal_descricao", ""),
        },
        "cnaes_secundarios": cnaes_sec[:5],
        "endereco": {
            "logradouro": data.get("logradouro", ""),
            "numero": data.get("numero", ""),
            "complemento": data.get("complemento", ""),
            "bairro": data.get("bairro", ""),
            "municipio": data.get("municipio", ""),
            "uf": data.get("uf", ""),
            "cep": data.get("cep", ""),
        },
        "qsa": qsa,
        "total_socios": len(qsa),
    }


def enrich_analysis_data(
    cnpj: str,
    dados_extraidos: dict[str, Any],
    status_callback=None,
) -> dict[str, Any]:
    """
    Enriquece os dados extraídos com informações públicas do CNPJ.
    Adiciona entrada 'kyc_publico' aos dados para a análise.

    Args:
        cnpj: CNPJ do tomador
        dados_extraidos: dados já extraídos dos documentos
        status_callback: função para reportar progresso

    Returns:
        dados_extraidos atualizado com dados KYC
    """
    def _status(msg):
        if status_callback:
            status_callback(msg)
        logger.info(msg)

    if not cnpj or cnpj == "N/I":
        _status("CNPJ não informado — pulando enriquecimento KYC.")
        return dados_extraidos

    _status(f"Consultando dados públicos do CNPJ {cnpj}...")
    kyc_data = fetch_cnpj_data(cnpj)

    if "error" in kyc_data:
        _status(f"KYC: {kyc_data['error']}")
        return dados_extraidos

    _status(f"KYC: {kyc_data.get('razao_social', '—')} | {kyc_data.get('situacao_cadastral', '—')} | "
            f"Capital Social: R$ {kyc_data.get('capital_social', 0):,.0f} | "
            f"{kyc_data.get('total_socios', 0)} sócio(s)")

    # Add to extraction data
    dados_extraidos["kyc_publico"] = {
        "classificacao": {"tipo": "kyc_publico", "confianca": 1.0},
        "dados": kyc_data,
    }

    return dados_extraidos
