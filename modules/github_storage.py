"""
ZYN Capital — Persistência via GitHub API
Salva e carrega análises diretamente no repositório GitHub,
garantindo que dados sobrevivam a redeploys do Streamlit Cloud.
"""

import base64
import json
import logging
import os
from typing import Any

import requests

logger = logging.getLogger(__name__)

REPO = os.environ.get("GITHUB_REPO", "cadastrofinders/zyn-credit-engine")
BRANCH = os.environ.get("GITHUB_BRANCH", "main")
DATA_DIR = "data/analyses"


def _get_token() -> str | None:
    """Retorna GitHub token dos secrets."""
    return os.environ.get("GH_PAT") or os.environ.get("GITHUB_TOKEN")


def _headers() -> dict:
    token = _get_token()
    if not token:
        return {}
    return {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
    }


def _api_url(path: str) -> str:
    return f"https://api.github.com/repos/{REPO}/contents/{path}"


def save_analysis(filename: str, payload: dict) -> bool:
    """Salva análise como JSON no repositório GitHub."""
    token = _get_token()
    if not token:
        logger.warning("GH_PAT não configurado — salvando apenas localmente")
        return False

    path = f"{DATA_DIR}/{filename}"
    content = json.dumps(payload, ensure_ascii=False, indent=2)
    encoded = base64.b64encode(content.encode("utf-8")).decode("utf-8")

    # Check if file already exists (need SHA for update)
    sha = None
    resp = requests.get(_api_url(path), headers=_headers(), params={"ref": BRANCH})
    if resp.status_code == 200:
        sha = resp.json().get("sha")

    data: dict[str, Any] = {
        "message": f"auto: salvar análise {filename}",
        "content": encoded,
        "branch": BRANCH,
    }
    if sha:
        data["sha"] = sha

    resp = requests.put(_api_url(path), headers=_headers(), json=data)
    if resp.status_code in (200, 201):
        logger.info("Análise salva no GitHub: %s", filename)
        return True
    else:
        logger.error("Erro ao salvar no GitHub: %s %s", resp.status_code, resp.text[:200])
        return False


def list_analyses() -> list[dict]:
    """Lista todas as análises salvas no repositório."""
    token = _get_token()
    if not token:
        return []

    resp = requests.get(_api_url(DATA_DIR), headers=_headers(), params={"ref": BRANCH})
    if resp.status_code != 200:
        logger.warning("Erro ao listar análises: %s", resp.status_code)
        return []

    files = resp.json()
    if not isinstance(files, list):
        return []

    items = []
    for f in files:
        if not f.get("name", "").endswith(".json"):
            continue
        items.append({
            "name": f["name"],
            "sha": f["sha"],
            "size": f.get("size", 0),
            "download_url": f.get("download_url"),
        })

    # Sort by name descending (filenames include timestamp)
    items.sort(key=lambda x: x["name"], reverse=True)
    return items


def load_analysis(filename: str) -> dict | None:
    """Carrega uma análise específica do repositório."""
    token = _get_token()
    if not token:
        return None

    path = f"{DATA_DIR}/{filename}"
    resp = requests.get(_api_url(path), headers=_headers(), params={"ref": BRANCH})
    if resp.status_code != 200:
        return None

    data = resp.json()
    content = base64.b64decode(data["content"]).decode("utf-8")
    result = json.loads(content)
    result["_filename"] = filename
    result["_sha"] = data["sha"]
    return result


def delete_analysis(filename: str) -> bool:
    """Remove uma análise do repositório."""
    token = _get_token()
    if not token:
        return False

    path = f"{DATA_DIR}/{filename}"
    # Need SHA to delete
    resp = requests.get(_api_url(path), headers=_headers(), params={"ref": BRANCH})
    if resp.status_code != 200:
        return False

    sha = resp.json().get("sha")
    data = {
        "message": f"auto: excluir análise {filename}",
        "sha": sha,
        "branch": BRANCH,
    }
    resp = requests.delete(_api_url(path), headers=_headers(), json=data)
    return resp.status_code == 200


def sync_local_to_github(history_dir) -> int:
    """Sincroniza análises locais para GitHub (migração inicial)."""
    from pathlib import Path
    history_path = Path(history_dir)
    if not history_path.exists():
        return 0

    count = 0
    for f in history_path.glob("*.json"):
        try:
            payload = json.loads(f.read_text())
            if save_analysis(f.name, payload):
                count += 1
        except Exception as e:
            logger.error("Erro ao sincronizar %s: %s", f.name, e)
    return count
