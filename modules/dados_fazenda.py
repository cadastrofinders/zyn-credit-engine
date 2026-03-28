"""
Dados Fazenda — Integração com API para análise de crédito agro
Consulta automatizada de CAR, NDVI, embargos e sobreposições ambientais

API: https://api.dadosfazenda.com.br
Auth: Supabase JWT (email/password)
"""

import requests
import logging
import time
from typing import Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone

logger = logging.getLogger(__name__)

BASE_URL = "https://api.dadosfazenda.com.br"
APP_URL = "https://app.dadosfazenda.com.br"
SUPABASE_URL = "https://touwirvweqpddxpblgip.supabase.co"
SUPABASE_API_KEY = (
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9"
    ".eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRvdXdpcnZ3ZXFwZGR4cGJsZ2lwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzE3MTQ4NDMsImV4cCI6MjA4NzI5MDg0M30"
    ".DmoStkZO-At10tBwkeYkihNYT4_4GMZSyTynmfI1o9E"
)


class DadosFazendaError(Exception):
    """Erro genérico da integração Dados Fazenda."""
    pass


class DadosFazendaAuthError(DadosFazendaError):
    """Falha de autenticação."""
    pass


class DadosFazendaClient:
    """Client para a API Dados Fazenda."""

    def __init__(self, email: str, password: str):
        self.email = email
        self.password = password
        self.token: Optional[str] = None
        self.token_expires: float = 0
        self.session = requests.Session()
        self.session.headers.update({"Content-Type": "application/json"})

    # ------------------------------------------------------------------
    # Auth
    # ------------------------------------------------------------------

    def _authenticate(self) -> bool:
        """Authenticate via Supabase and cache JWT token."""
        url = f"{SUPABASE_URL}/auth/v1/token"
        params = {"grant_type": "password"}
        payload = {"email": self.email, "password": self.password}
        headers = {
            "Content-Type": "application/json",
            "apikey": SUPABASE_API_KEY,
        }

        t0 = time.monotonic()
        try:
            resp = self.session.post(
                url, params=params, json=payload, headers=headers, timeout=15
            )
            elapsed = time.monotonic() - t0
            logger.info(f"[DadosFazenda] AUTH {resp.status_code} em {elapsed:.2f}s")

            if resp.status_code == 200:
                data = resp.json()
                self.token = data.get("access_token")
                expires_in = data.get("expires_in", 3600)
                # Renew 60s before actual expiry
                self.token_expires = time.time() + expires_in - 60
                return True

            logger.error(f"[DadosFazenda] Auth falhou: {resp.status_code} {resp.text}")
            raise DadosFazendaAuthError(
                f"Autenticação falhou ({resp.status_code}): {resp.text}"
            )
        except requests.RequestException as e:
            logger.error(f"[DadosFazenda] Auth erro de rede: {e}")
            raise DadosFazendaAuthError(f"Erro de rede na autenticação: {e}") from e

    def _ensure_auth(self) -> None:
        """Authenticate if token is missing or expired."""
        if not self.token or time.time() >= self.token_expires:
            self._authenticate()

    # ------------------------------------------------------------------
    # HTTP helper
    # ------------------------------------------------------------------

    def _request(
        self,
        method: str,
        path: str,
        max_retries: int = 3,
        **kwargs,
    ) -> dict:
        """Make authenticated request with auto-retry on 401 and rate-limit."""
        self._ensure_auth()

        url = f"{BASE_URL}{path}"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
            "apikey": SUPABASE_API_KEY,
        }

        for attempt in range(1, max_retries + 1):
            t0 = time.monotonic()
            try:
                resp = self.session.request(
                    method, url, headers=headers, timeout=30, **kwargs
                )
                elapsed = time.monotonic() - t0
                logger.info(
                    f"[DadosFazenda] {method} {path} → {resp.status_code} "
                    f"em {elapsed:.2f}s (tentativa {attempt})"
                )

                # Success
                if resp.status_code == 200:
                    return resp.json()

                # Empty result (204 / 404 without body)
                if resp.status_code in (204, 404):
                    return {}

                # Token expired — re-auth and retry
                if resp.status_code == 401 and attempt < max_retries:
                    logger.warning("[DadosFazenda] Token expirado, re-autenticando...")
                    self._authenticate()
                    headers["Authorization"] = f"Bearer {self.token}"
                    continue

                # Rate limit — wait and retry
                if resp.status_code == 429 and attempt < max_retries:
                    wait = float(resp.headers.get("Retry-After", 2))
                    logger.warning(
                        f"[DadosFazenda] Rate limit, aguardando {wait}s..."
                    )
                    time.sleep(wait)
                    continue

                # Other errors
                logger.error(
                    f"[DadosFazenda] Erro {resp.status_code}: {resp.text[:500]}"
                )
                return {}

            except requests.RequestException as e:
                elapsed = time.monotonic() - t0
                logger.error(
                    f"[DadosFazenda] {method} {path} erro de rede "
                    f"({elapsed:.2f}s, tentativa {attempt}): {e}"
                )
                if attempt == max_retries:
                    return {}
                time.sleep(1)

        return {}

    # ------------------------------------------------------------------
    # Busca aberta — qualquer propriedade do Brasil
    # ------------------------------------------------------------------

    def search_car_by_location(self, lat: float, lon: float, radius_km: float = 5) -> list:
        """Busca CARs na base nacional por coordenadas + raio.

        Returns:
            list of CAR codes encontrados na área.
        """
        data = self._request(
            "GET",
            "/car/search/radius",
            params={"lat": lat, "lon": lon, "radius": radius_km},
        )
        results = data.get("data", data.get("features", []))
        if isinstance(results, list):
            return results
        return []

    def geocode(self, query: str) -> dict:
        """Geocode de município/localização via API do app.

        Returns:
            dict com lat, lon, display_name.
        """
        self._ensure_auth()
        try:
            resp = self.session.get(
                f"{APP_URL}/api/geocode",
                params={"q": query},
                headers={"Authorization": f"Bearer {self.token}"},
                timeout=10,
            )
            if resp.status_code == 200:
                return resp.json()
        except Exception as e:
            logger.warning(f"[DadosFazenda] Geocode falhou: {e}")
        return {}

    def consulta_car_aberta(self, car_code: str) -> dict:
        """Consulta completa de um CAR qualquer (base nacional).

        Não precisa estar cadastrado na conta — usa endpoints de consulta aberta.

        Returns:
            dict com ndvi, embargos, sobreposicoes, alertas, score_ambiental.
        """
        car_code = car_code.strip().upper()
        logger.info(f"[DadosFazenda] Consulta aberta CAR: {car_code}")

        resultado = {"car_code": car_code, "alertas": []}
        endpoints = {}

        # Parallel calls
        with ThreadPoolExecutor(max_workers=4) as executor:
            endpoints["ndvi"] = executor.submit(self.get_ndvi, car_code)
            endpoints["embargos"] = executor.submit(self.get_embargos, car_code)
            # Sobreposições usam o mesmo CAR code (não apenas SIGEF)
            endpoints["quilombolas"] = executor.submit(
                self._request, "GET", f"/quilombolas/sigef/{car_code}"
            )
            endpoints["terras_indigenas"] = executor.submit(
                self._request, "GET", f"/terras-indigenas/sigef/{car_code}"
            )
            endpoints["assentamentos"] = executor.submit(
                self._request, "GET", f"/assentamentos/sigef/{car_code}"
            )
            endpoints["unidades_conservacao"] = executor.submit(
                self._request, "GET", f"/unidades-conservacao/sigef/{car_code}"
            )
            endpoints["incra"] = executor.submit(
                self._request, "GET", f"/incra/{car_code}/local"
            )

        alertas = []

        # NDVI
        try:
            ndvi_raw = endpoints["ndvi"].result(timeout=15)
            ndvi_data = ndvi_raw.get("data", ndvi_raw)
            ts = ndvi_data.get("timeseries", [])
            if ts:
                latest = ts[-1] if isinstance(ts[-1], dict) else {}
                resultado["ndvi"] = {
                    "ndvi_mean": latest.get("ndvi_mean", 0),
                    "ndvi_median": latest.get("ndvi_median", 0),
                    "ndvi_min": latest.get("ndvi_min", 0),
                    "ndvi_max": latest.get("ndvi_max", 0),
                    "data_cena": latest.get("scene_date", ""),
                    "pontos": len(ts),
                }
                # Trend: compare last 3 vs previous 3
                if len(ts) >= 6:
                    recent = sum(t.get("ndvi_mean", 0) for t in ts[-3:]) / 3
                    prev = sum(t.get("ndvi_mean", 0) for t in ts[-6:-3]) / 3
                    diff = recent - prev
                    if diff > 0.05:
                        resultado["ndvi"]["tendencia"] = "Melhora"
                    elif diff < -0.05:
                        resultado["ndvi"]["tendencia"] = "Degradação"
                        alertas.append("⚠ NDVI em degradação nos últimos meses")
                    else:
                        resultado["ndvi"]["tendencia"] = "Estável"
                else:
                    resultado["ndvi"]["tendencia"] = "Dados insuficientes"

                veg_pct = latest.get("distribution", {})
                resultado["ndvi"]["cobertura_vegetal_pct"] = (
                    veg_pct.get("dense_vegetation", 0)
                    + veg_pct.get("moderate_vegetation", 0)
                ) * 100 if isinstance(veg_pct, dict) else 0

                # ── Distribuição de áreas (%) ──
                if isinstance(veg_pct, dict):
                    resultado["ndvi"]["distribuicao"] = veg_pct
        except Exception as e:
            logger.warning(f"[DadosFazenda] NDVI falhou: {e}")
            resultado["ndvi"] = {}

        # Embargos
        _emb_data_cache = {}
        try:
            emb_raw = endpoints["embargos"].result(timeout=15)
            emb_data = emb_raw.get("data", emb_raw)
            _emb_data_cache = emb_data  # cache for area extraction
            tem_embargo = False
            detalhes = ""
            if isinstance(emb_data, dict):
                mb = emb_data.get("mapbiomas", {})
                ibama = emb_data.get("ibama", [])
                if ibama and isinstance(ibama, list) and len(ibama) > 0:
                    tem_embargo = True
                    detalhes = f"IBAMA: {len(ibama)} embargo(s)"
                if isinstance(mb, dict) and mb.get("has_alerts"):
                    tem_embargo = True
                    detalhes += " | MapBiomas: alertas ativos"
            resultado["embargos"] = {
                "tem_embargo": tem_embargo,
                "detalhes": detalhes or "Sem embargos",
            }
            if tem_embargo:
                alertas.append(f"🔴 Embargo ativo: {detalhes}")
            else:
                alertas.append("✅ Sem embargos IBAMA/MapBiomas")
        except Exception as e:
            logger.warning(f"[DadosFazenda] Embargos falhou: {e}")
            resultado["embargos"] = {}

        # ── Área da propriedade (ha) ──
        area_ha = 0
        try:
            if isinstance(_emb_data_cache, dict):
                mb_area = _emb_data_cache.get("mapbiomas", _emb_data_cache)
                if isinstance(mb_area, dict):
                    area_ha = float(mb_area.get("area_ha", 0))
        except Exception:
            pass

        # Calcular breakdown de áreas via distribuição NDVI
        distribuicao = resultado.get("ndvi", {}).get("distribuicao", {})
        area_breakdown = {
            "area_total_ha": round(area_ha, 2),
            "area_agricultavel_ha": 0,
            "area_pastagem_ha": 0,
            "area_consolidada_ha": 0,
            "area_solo_exposto_ha": 0,
            "area_agua_urbano_ha": 0,
        }
        if area_ha > 0 and distribuicao:
            pct_dense = distribuicao.get("dense_healthy_vegetation", 0) / 100
            pct_sparse = distribuicao.get("sparse_vegetation", 0) / 100
            pct_bare = distribuicao.get("bare_soil", 0) / 100
            pct_bare_urban = distribuicao.get("bare_soil_urban", 0) / 100
            pct_water = distribuicao.get("water_clouds_urban", 0) / 100

            area_breakdown["area_agricultavel_ha"] = round(pct_dense * area_ha, 2)
            area_breakdown["area_pastagem_ha"] = round(pct_sparse * area_ha, 2)
            area_breakdown["area_consolidada_ha"] = round(
                (pct_dense + pct_sparse) * area_ha, 2
            )
            area_breakdown["area_solo_exposto_ha"] = round(
                (pct_bare + pct_bare_urban) * area_ha, 2
            )
            area_breakdown["area_agua_urbano_ha"] = round(pct_water * area_ha, 2)

        resultado["areas"] = area_breakdown

        # Sobreposições
        sobrep = {}
        for key in ("quilombolas", "terras_indigenas", "assentamentos", "unidades_conservacao"):
            try:
                raw = endpoints[key].result(timeout=15)
                data = raw.get("data", raw)
                tem = False
                if isinstance(data, dict):
                    tem = data.get("has_overlap", data.get("tem_sobreposicao", False))
                elif isinstance(data, list) and len(data) > 0:
                    tem = True
                label = key.replace("_", " ").title()
                sobrep[key] = {"tem_sobreposicao": tem, "dados": data}
                if tem:
                    alertas.append(f"🔴 Sobreposição com {label} detectada")
                else:
                    alertas.append(f"✅ {label}: Livre")
            except Exception as e:
                logger.warning(f"[DadosFazenda] {key} falhou: {e}")
                sobrep[key] = {"tem_sobreposicao": False, "dados": {}}
        resultado["sobreposicoes"] = sobrep

        # INCRA
        try:
            incra_raw = endpoints["incra"].result(timeout=15)
            resultado["incra"] = incra_raw.get("data", incra_raw)
        except Exception:
            resultado["incra"] = {}

        # Score ambiental
        tem_embargo = resultado.get("embargos", {}).get("tem_embargo", False)
        tem_ti = sobrep.get("terras_indigenas", {}).get("tem_sobreposicao", False)
        tem_quilombo = sobrep.get("quilombolas", {}).get("tem_sobreposicao", False)
        ndvi_deg = resultado.get("ndvi", {}).get("tendencia") == "Degradação"

        if tem_embargo or tem_ti or tem_quilombo:
            resultado["score_ambiental"] = "Vermelho"
        elif ndvi_deg or sobrep.get("assentamentos", {}).get("tem_sobreposicao"):
            resultado["score_ambiental"] = "Amarelo"
        else:
            resultado["score_ambiental"] = "Verde"

        resultado["alertas"] = alertas
        return resultado

    # ------------------------------------------------------------------
    # Endpoints individuais (propriedades monitoradas)
    # ------------------------------------------------------------------

    def get_properties(self) -> list:
        """Lista todas as propriedades monitoradas.

        Returns:
            list of dict com id, car_code, farm_name, municipio, estado,
            area_ha, latitude, longitude, source_type.
        """
        data = self._request("GET", "/api/properties")
        if isinstance(data, list):
            return data
        return data.get("data", data.get("properties", []))

    def buscar_cars_por_documento(self, documento: str, tipo: str = "CPF") -> list:
        """Busca CARs vinculados a um CPF ou CNPJ via endpoint /api/cpf-car/{ni}.

        Args:
            documento: CPF ou CNPJ (com ou sem formatação).
            tipo: "CPF" ou "CNPJ" (informativo).

        Returns:
            list de códigos CAR encontrados.
        """
        import re
        digits = re.sub(r"[^\d]", "", documento)
        if not digits:
            return []

        logger.info(f"[DadosFazenda] Buscando CARs por {tipo}: {digits} via /api/cpf-car/{digits}")

        try:
            data = self._request("GET", f"/api/cpf-car/{digits}")
            # Resposta pode ser lista direta ou dict com campo "data"/"cars"
            if isinstance(data, list):
                cars = data
            elif isinstance(data, dict):
                cars = data.get("data", data.get("cars", data.get("results", [])))
                if not isinstance(cars, list):
                    # Pode ser um dict com car_code direto
                    car_code = data.get("car_code") or data.get("codigo_car")
                    if car_code:
                        return [car_code]
                    # Resposta vazia ou formato inesperado
                    logger.warning(f"[DadosFazenda] /api/cpf-car retornou formato inesperado: {list(data.keys())}")
                    return []
            else:
                return []

            # Extrair códigos CAR da lista
            result = []
            for item in cars:
                if isinstance(item, str):
                    result.append(item)
                elif isinstance(item, dict):
                    code = item.get("car_code") or item.get("codigo_car") or item.get("car") or item.get("code")
                    if code:
                        result.append(code)

            unique = list(dict.fromkeys(result))
            logger.info(f"[DadosFazenda] /api/cpf-car retornou {len(unique)} CAR(s)")
            return unique
        except DadosFazendaError as e:
            logger.warning(f"[DadosFazenda] /api/cpf-car falhou: {e}")
            return []
        except Exception as e:
            logger.warning(f"[DadosFazenda] /api/cpf-car erro inesperado: {e}")
            return []

        logger.warning(f"[DadosFazenda] Nenhuma fonte retornou CARs para {tipo} {digits}")
        return []

    def get_ndvi(self, car_code: str) -> dict:
        """Série temporal de NDVI (saúde vegetativa) para um código CAR.

        Returns:
            dict com timeseries de ndvi_mean, ndvi_median, ndvi_min,
            ndvi_max, distribution.
        """
        return self._request("GET", f"/api/ndvi/car/{car_code}")

    def get_embargos(self, car_code: str) -> dict:
        """Embargos IBAMA e MapBiomas para um código CAR.

        Returns:
            dict com source, mapbiomas: {...}, ibama: [...].
        """
        return self._request("GET", f"/api/embargo/car/{car_code}")

    def get_quilombolas(self, sigef_id: str) -> dict:
        """Sobreposição com áreas quilombolas."""
        return self._request("GET", f"/quilombolas/sigef/{sigef_id}")

    def get_terras_indigenas(self, sigef_id: str) -> dict:
        """Sobreposição com terras indígenas."""
        return self._request("GET", f"/terras-indigenas/sigef/{sigef_id}")

    def get_assentamentos(self, sigef_id: str) -> dict:
        """Sobreposição com assentamentos."""
        return self._request("GET", f"/assentamentos/sigef/{sigef_id}")

    def get_unidades_conservacao(self, sigef_id: str) -> dict:
        """Sobreposição com unidades de conservação."""
        return self._request("GET", f"/unidades-conservacao/sigef/{sigef_id}")

    def get_sobreposicoes(self, sigef_id: str) -> dict:
        """Consulta todas as sobreposições ambientais em paralelo.

        Returns:
            dict com quilombolas, terras_indigenas, assentamentos,
            unidades_conservacao.
        """
        resultados = {
            "quilombolas": {},
            "terras_indigenas": {},
            "assentamentos": {},
            "unidades_conservacao": {},
        }

        calls = {
            "quilombolas": (self.get_quilombolas, sigef_id),
            "terras_indigenas": (self.get_terras_indigenas, sigef_id),
            "assentamentos": (self.get_assentamentos, sigef_id),
            "unidades_conservacao": (self.get_unidades_conservacao, sigef_id),
        }

        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = {
                executor.submit(fn, arg): key for key, (fn, arg) in calls.items()
            }
            for future in as_completed(futures):
                key = futures[future]
                try:
                    resultados[key] = future.result()
                except Exception as e:
                    logger.error(f"[DadosFazenda] Erro em {key}: {e}")
                    resultados[key] = {"erro": str(e)}

        return resultados

    def get_incra(self, sigef_id: str) -> dict:
        """Dados de georreferenciamento INCRA."""
        return self._request("GET", f"/incra/{sigef_id}/local")

    # ------------------------------------------------------------------
    # Análise NDVI
    # ------------------------------------------------------------------

    @staticmethod
    def _analisar_ndvi(ndvi_data: dict) -> dict:
        """Calcula tendência NDVI e classifica saúde vegetativa.

        Compara média dos últimos 3 meses vs 3 meses anteriores.
        Retorna dict com ultimo_ndvi, tendencia, variacao_pct.
        """
        resultado = {
            "ultimo_ndvi": None,
            "tendencia": "Indeterminado",
            "variacao_pct": 0.0,
            "dados_disponiveis": False,
        }

        # Extrair série temporal — aceita vários formatos de resposta
        timeseries = (
            ndvi_data.get("timeseries")
            or ndvi_data.get("data")
            or ndvi_data.get("ndvi")
        )
        if not timeseries or not isinstance(timeseries, list):
            return resultado

        # Ordenar por data (campo date ou period)
        def _get_date(entry: dict) -> str:
            return entry.get("date") or entry.get("period") or ""

        try:
            sorted_ts = sorted(timeseries, key=_get_date)
        except (TypeError, KeyError):
            sorted_ts = timeseries

        if len(sorted_ts) < 2:
            if sorted_ts:
                resultado["ultimo_ndvi"] = sorted_ts[-1].get("ndvi_mean")
                resultado["dados_disponiveis"] = True
            return resultado

        resultado["dados_disponiveis"] = True
        resultado["ultimo_ndvi"] = sorted_ts[-1].get("ndvi_mean")

        # Últimos 3 períodos vs 3 anteriores
        recentes = sorted_ts[-3:]
        anteriores = sorted_ts[-6:-3] if len(sorted_ts) >= 6 else sorted_ts[:-3]

        if not anteriores:
            return resultado

        def _media_ndvi(entries: list) -> Optional[float]:
            vals = [
                e.get("ndvi_mean")
                for e in entries
                if e.get("ndvi_mean") is not None
            ]
            return sum(vals) / len(vals) if vals else None

        media_recente = _media_ndvi(recentes)
        media_anterior = _media_ndvi(anteriores)

        if media_recente is None or media_anterior is None or media_anterior == 0:
            return resultado

        variacao = ((media_recente - media_anterior) / abs(media_anterior)) * 100
        resultado["variacao_pct"] = round(variacao, 2)

        if variacao <= -15:
            resultado["tendencia"] = "Degradação"
        elif variacao >= 5:
            resultado["tendencia"] = "Melhora"
        else:
            resultado["tendencia"] = "Estável"

        return resultado

    # ------------------------------------------------------------------
    # Análise de embargos
    # ------------------------------------------------------------------

    @staticmethod
    def _analisar_embargos(embargo_data: dict) -> dict:
        """Analisa se há embargos ativos."""
        resultado = {
            "tem_embargo": False,
            "total_embargos": 0,
            "fontes": [],
            "detalhes": [],
        }

        if not embargo_data:
            return resultado

        # IBAMA
        ibama = embargo_data.get("ibama", [])
        if isinstance(ibama, list) and ibama:
            resultado["tem_embargo"] = True
            resultado["total_embargos"] += len(ibama)
            resultado["fontes"].append("IBAMA")
            for e in ibama:
                resultado["detalhes"].append({
                    "fonte": "IBAMA",
                    "numero": e.get("numero") or e.get("id"),
                    "data": e.get("data") or e.get("date"),
                    "motivo": e.get("motivo") or e.get("descricao") or "N/I",
                    "status": e.get("status") or "Ativo",
                })

        # MapBiomas
        mapbiomas = embargo_data.get("mapbiomas", {})
        if isinstance(mapbiomas, dict):
            alertas = mapbiomas.get("alertas") or mapbiomas.get("alerts") or []
            if isinstance(alertas, list) and alertas:
                resultado["tem_embargo"] = True
                resultado["total_embargos"] += len(alertas)
                if "MapBiomas" not in resultado["fontes"]:
                    resultado["fontes"].append("MapBiomas")
                for a in alertas:
                    resultado["detalhes"].append({
                        "fonte": "MapBiomas",
                        "numero": a.get("id"),
                        "data": a.get("date") or a.get("data"),
                        "motivo": a.get("type") or a.get("tipo") or "Desmatamento",
                        "status": a.get("status") or "Ativo",
                    })
            # Pode vir como flag direto
            elif mapbiomas.get("embargo") or mapbiomas.get("embargado"):
                resultado["tem_embargo"] = True
                resultado["total_embargos"] += 1
                resultado["fontes"].append("MapBiomas")

        return resultado

    # ------------------------------------------------------------------
    # Análise de sobreposições
    # ------------------------------------------------------------------

    @staticmethod
    def _analisar_sobreposicoes(sobreposicoes: dict) -> dict:
        """Classifica severidade das sobreposições encontradas."""
        resultado = {
            "tem_sobreposicao": False,
            "criticas": [],  # Terra indígena, quilombola
            "moderadas": [],  # Unidades de conservação
            "baixas": [],  # Assentamentos
        }

        def _has_overlap(data: dict) -> bool:
            """Verifica se o retorno indica sobreposição."""
            if not data or data.get("erro"):
                return False
            # Checagens comuns
            if data.get("sobrepoe") or data.get("overlap"):
                return True
            if data.get("features") and len(data["features"]) > 0:
                return True
            if data.get("data") and isinstance(data["data"], list) and len(data["data"]) > 0:
                return True
            if data.get("results") and isinstance(data["results"], list) and len(data["results"]) > 0:
                return True
            return False

        # Quilombolas — crítico
        quilombolas = sobreposicoes.get("quilombolas", {})
        if _has_overlap(quilombolas):
            resultado["tem_sobreposicao"] = True
            resultado["criticas"].append({
                "tipo": "Área Quilombola",
                "dados": quilombolas,
            })

        # Terras indígenas — crítico
        ti = sobreposicoes.get("terras_indigenas", {})
        if _has_overlap(ti):
            resultado["tem_sobreposicao"] = True
            resultado["criticas"].append({
                "tipo": "Terra Indígena",
                "dados": ti,
            })

        # Unidades de conservação — moderado
        uc = sobreposicoes.get("unidades_conservacao", {})
        if _has_overlap(uc):
            resultado["tem_sobreposicao"] = True
            resultado["moderadas"].append({
                "tipo": "Unidade de Conservação",
                "dados": uc,
            })

        # Assentamentos — baixo
        assent = sobreposicoes.get("assentamentos", {})
        if _has_overlap(assent):
            resultado["tem_sobreposicao"] = True
            resultado["baixas"].append({
                "tipo": "Assentamento",
                "dados": assent,
            })

        return resultado

    # ------------------------------------------------------------------
    # Score ambiental
    # ------------------------------------------------------------------

    @staticmethod
    def _calcular_score(
        analise_ndvi: dict,
        analise_embargos: dict,
        analise_sobreposicoes: dict,
    ) -> str:
        """Calcula score ambiental consolidado.

        Vermelho: embargo ativo OU sobreposição crítica OU NDVI em degradação severa
        Amarelo: NDVI degradando OU sobreposições moderadas/baixas
        Verde: sem restrições
        """
        # Vermelho
        if analise_embargos.get("tem_embargo"):
            return "Vermelho"
        if analise_sobreposicoes.get("criticas"):
            return "Vermelho"
        if (
            analise_ndvi.get("tendencia") == "Degradação"
            and analise_ndvi.get("variacao_pct", 0) <= -30
        ):
            return "Vermelho"

        # Amarelo
        if analise_ndvi.get("tendencia") == "Degradação":
            return "Amarelo"
        if analise_sobreposicoes.get("moderadas") or analise_sobreposicoes.get("baixas"):
            return "Amarelo"

        return "Verde"

    # ------------------------------------------------------------------
    # Geração de alertas
    # ------------------------------------------------------------------

    @staticmethod
    def _gerar_alertas(
        car_code: str,
        analise_ndvi: dict,
        analise_embargos: dict,
        analise_sobreposicoes: dict,
    ) -> list:
        """Gera lista de alertas em texto para o relatório."""
        alertas = []

        # Embargos
        if analise_embargos.get("tem_embargo"):
            fontes = ", ".join(analise_embargos.get("fontes", []))
            total = analise_embargos.get("total_embargos", 0)
            alertas.append(
                f"\u26a0 Embargo ativo ({fontes}) na propriedade {car_code} "
                f"({total} registro(s))"
            )

        # Sobreposições críticas
        for item in analise_sobreposicoes.get("criticas", []):
            alertas.append(
                f"\U0001f534 Sobreposição com {item['tipo']} detectada "
                f"na propriedade {car_code}"
            )

        # Sobreposições moderadas
        for item in analise_sobreposicoes.get("moderadas", []):
            alertas.append(
                f"\u26a0 Sobreposição com {item['tipo']} detectada "
                f"na propriedade {car_code}"
            )

        # Sobreposições baixas
        for item in analise_sobreposicoes.get("baixas", []):
            alertas.append(
                f"\u26a0 Proximidade com {item['tipo']} detectada "
                f"na propriedade {car_code}"
            )

        # NDVI
        if analise_ndvi.get("tendencia") == "Degradação":
            var = analise_ndvi.get("variacao_pct", 0)
            alertas.append(
                f"\u26a0 NDVI em degradação nos últimos 6 meses "
                f"({var:+.1f}%) na propriedade {car_code}"
            )

        # Sem restrições
        if not alertas:
            alertas.append(
                f"\u2705 Sem restrições ambientais detectadas "
                f"na propriedade {car_code}"
            )

        return alertas

    # ------------------------------------------------------------------
    # Consulta completa (propriedade individual)
    # ------------------------------------------------------------------

    def cruzar_grupo_sigef(self, car_codes_documentos: list) -> dict:
        """
        Cross-reference CAR codes found in client documents with
        properties registered in Dados Fazenda account.

        Returns:
        {
            "propriedades_cadastradas": int,
            "propriedades_documentos": int,
            "matches": [{"car_code": str, "farm_name": str, "area_ha": float, "status": "Encontrada"}],
            "nao_cadastradas": [{"car_code": str, "status": "Não encontrada no monitoramento"}],
            "extras_monitoramento": [{"car_code": str, "farm_name": str, "area_ha": float, "status": "Cadastrada mas não citada em documentos"}],
            "cobertura_pct": float,  # % of document CARs found in monitoring
            "alertas": [str]
        }
        """
        props = self.get_properties()
        prop_cars = {p.get("car_code", "").upper(): p for p in props}
        doc_cars = set(c.upper() for c in car_codes_documentos)

        matches = []
        nao_cadastradas = []
        extras = []
        alertas = []

        for car in doc_cars:
            if car in prop_cars:
                p = prop_cars[car]
                matches.append({
                    "car_code": car,
                    "farm_name": p.get("farm_name", ""),
                    "area_ha": p.get("area_ha", 0),
                    "status": "Encontrada"
                })
            else:
                nao_cadastradas.append({
                    "car_code": car,
                    "status": "Não encontrada no monitoramento"
                })
                alertas.append(f"⚠ CAR {car} citado em documentos mas NÃO está no monitoramento Dados Fazenda")

        for car, p in prop_cars.items():
            if car not in doc_cars:
                extras.append({
                    "car_code": car,
                    "farm_name": p.get("farm_name", ""),
                    "area_ha": p.get("area_ha", 0),
                    "status": "Cadastrada mas não citada em documentos"
                })

        cobertura = (len(matches) / len(doc_cars) * 100) if doc_cars else 0

        if cobertura < 50:
            alertas.append(f"🔴 Apenas {cobertura:.0f}% dos CARs dos documentos estão monitorados")
        elif cobertura < 100:
            alertas.append(f"⚠ {cobertura:.0f}% dos CARs monitorados — {len(nao_cadastradas)} faltante(s)")
        else:
            alertas.append("✅ 100% dos CARs dos documentos estão monitorados no Dados Fazenda")

        if extras:
            alertas.append(f"ℹ {len(extras)} propriedade(s) monitorada(s) não citada(s) nos documentos do cliente")

        return {
            "propriedades_cadastradas": len(prop_cars),
            "propriedades_documentos": len(doc_cars),
            "matches": matches,
            "nao_cadastradas": nao_cadastradas,
            "extras_monitoramento": extras,
            "cobertura_pct": cobertura,
            "alertas": alertas
        }

    def consulta_completa(self, car_code: str, sigef_id: str = None) -> dict:
        """Consulta completa de uma propriedade rural para análise de crédito.

        Args:
            car_code: Código CAR da propriedade.
            sigef_id: Código SIGEF (opcional, necessário para sobreposições e INCRA).

        Returns:
            dict consolidado com ndvi, embargos, sobreposições, incra,
            alertas e score_ambiental.
        """
        resultado = {
            "car_code": car_code,
            "sigef_id": sigef_id,
            "ndvi": {},
            "embargos": {},
            "sobreposicoes": {},
            "incra": {},
            "alertas": [],
            "score_ambiental": "Indeterminado",
            "consultado_em": datetime.now(timezone.utc).isoformat(),
        }

        # Parallel calls: NDVI + embargos (por CAR) e sobreposições + INCRA (por SIGEF)
        with ThreadPoolExecutor(max_workers=4) as executor:
            future_ndvi = executor.submit(self.get_ndvi, car_code)
            time.sleep(0.5)
            future_embargo = executor.submit(self.get_embargos, car_code)

            future_sobr = None
            future_incra = None
            if sigef_id:
                time.sleep(0.5)
                future_sobr = executor.submit(self.get_sobreposicoes, sigef_id)
                time.sleep(0.5)
                future_incra = executor.submit(self.get_incra, sigef_id)

            # Coletar resultados
            try:
                ndvi_raw = future_ndvi.result(timeout=60)
            except Exception as e:
                logger.error(f"[DadosFazenda] NDVI falhou para {car_code}: {e}")
                ndvi_raw = {}

            try:
                embargo_raw = future_embargo.result(timeout=60)
            except Exception as e:
                logger.error(f"[DadosFazenda] Embargo falhou para {car_code}: {e}")
                embargo_raw = {}

            sobreposicoes_raw = {}
            incra_raw = {}
            if future_sobr:
                try:
                    sobreposicoes_raw = future_sobr.result(timeout=60)
                except Exception as e:
                    logger.error(
                        f"[DadosFazenda] Sobreposições falharam para {sigef_id}: {e}"
                    )
            if future_incra:
                try:
                    incra_raw = future_incra.result(timeout=60)
                except Exception as e:
                    logger.error(
                        f"[DadosFazenda] INCRA falhou para {sigef_id}: {e}"
                    )

        # Análise
        analise_ndvi = self._analisar_ndvi(ndvi_raw)
        analise_embargos = self._analisar_embargos(embargo_raw)
        analise_sobreposicoes = self._analisar_sobreposicoes(sobreposicoes_raw)

        resultado["ndvi"] = {
            "raw": ndvi_raw,
            "analise": analise_ndvi,
        }
        resultado["embargos"] = {
            "raw": embargo_raw,
            "analise": analise_embargos,
        }
        resultado["sobreposicoes"] = {
            "raw": sobreposicoes_raw,
            "analise": analise_sobreposicoes,
        }
        resultado["incra"] = incra_raw

        # ── Área breakdown ──
        area_ha = 0
        try:
            emb_inner = embargo_raw.get("data", embargo_raw) if isinstance(embargo_raw, dict) else {}
            if isinstance(emb_inner, dict):
                mb_area = emb_inner.get("mapbiomas", emb_inner)
                if isinstance(mb_area, dict):
                    area_ha = float(mb_area.get("area_ha", 0))
        except Exception:
            pass

        distribuicao = {}
        try:
            nd = ndvi_raw.get("data", ndvi_raw) if isinstance(ndvi_raw, dict) else {}
            ts = nd.get("timeseries", []) if isinstance(nd, dict) else []
            if ts and isinstance(ts[-1], dict):
                distribuicao = ts[-1].get("distribution", {})
        except Exception:
            pass

        area_breakdown = {
            "area_total_ha": round(area_ha, 2),
            "area_agricultavel_ha": 0,
            "area_pastagem_ha": 0,
            "area_consolidada_ha": 0,
            "area_solo_exposto_ha": 0,
            "area_agua_urbano_ha": 0,
        }
        if area_ha > 0 and isinstance(distribuicao, dict) and distribuicao:
            pct_dense = distribuicao.get("dense_healthy_vegetation", 0) / 100
            pct_sparse = distribuicao.get("sparse_vegetation", 0) / 100
            pct_bare = distribuicao.get("bare_soil", 0) / 100
            pct_bare_urban = distribuicao.get("bare_soil_urban", 0) / 100
            pct_water = distribuicao.get("water_clouds_urban", 0) / 100

            area_breakdown["area_agricultavel_ha"] = round(pct_dense * area_ha, 2)
            area_breakdown["area_pastagem_ha"] = round(pct_sparse * area_ha, 2)
            area_breakdown["area_consolidada_ha"] = round((pct_dense + pct_sparse) * area_ha, 2)
            area_breakdown["area_solo_exposto_ha"] = round((pct_bare + pct_bare_urban) * area_ha, 2)
            area_breakdown["area_agua_urbano_ha"] = round(pct_water * area_ha, 2)

        resultado["areas"] = area_breakdown

        # Score e alertas
        resultado["score_ambiental"] = self._calcular_score(
            analise_ndvi, analise_embargos, analise_sobreposicoes
        )
        resultado["alertas"] = self._gerar_alertas(
            car_code, analise_ndvi, analise_embargos, analise_sobreposicoes
        )

        logger.info(
            f"[DadosFazenda] Consulta completa {car_code}: "
            f"score={resultado['score_ambiental']}, "
            f"{len(resultado['alertas'])} alerta(s)"
        )

        return resultado

    # ------------------------------------------------------------------
    # Consulta em grupo
    # ------------------------------------------------------------------

    def consulta_grupo(
        self, car_codes: list, sigef_ids: list = None
    ) -> dict:
        """Consulta múltiplas propriedades (grupo econômico agro).

        Args:
            car_codes: Lista de códigos CAR.
            sigef_ids: Lista de códigos SIGEF (mesma ordem de car_codes).
                       Pode ser None ou conter None para propriedades sem SIGEF.

        Returns:
            dict consolidado com total_propriedades, area_total_ha,
            propriedades, alertas_consolidados, score_ambiental_grupo, resumo.
        """
        if sigef_ids is None:
            sigef_ids = [None] * len(car_codes)

        # Pad sigef_ids se menor que car_codes
        while len(sigef_ids) < len(car_codes):
            sigef_ids.append(None)

        propriedades = []
        alertas_consolidados = []

        # Processar propriedades com delay entre elas
        for i, (car, sigef) in enumerate(zip(car_codes, sigef_ids)):
            if i > 0:
                time.sleep(0.5)  # Rate limit entre propriedades
            try:
                resultado = self.consulta_completa(car, sigef)
                propriedades.append(resultado)
                alertas_consolidados.extend(resultado.get("alertas", []))
            except Exception as e:
                logger.error(f"[DadosFazenda] Falha na consulta de {car}: {e}")
                propriedades.append({
                    "car_code": car,
                    "sigef_id": sigef,
                    "erro": str(e),
                    "score_ambiental": "Indeterminado",
                    "alertas": [f"\u26a0 Consulta falhou para {car}: {e}"],
                })
                alertas_consolidados.append(
                    f"\u26a0 Consulta falhou para {car}: {e}"
                )

        # Score consolidado do grupo
        scores = [p.get("score_ambiental", "Indeterminado") for p in propriedades]
        if "Vermelho" in scores:
            score_grupo = "Vermelho"
        elif "Amarelo" in scores:
            score_grupo = "Amarelo"
        elif all(s == "Verde" for s in scores):
            score_grupo = "Verde"
        else:
            score_grupo = "Amarelo"

        # Área total — prioriza areas.area_total_ha (NDVI+embargos), fallback INCRA
        area_total = 0.0
        for p in propriedades:
            area_prop = p.get("areas", {}).get("area_total_ha", 0) or 0
            if area_prop <= 0:
                incra = p.get("incra", {})
                area_prop = (
                    incra.get("area_ha")
                    or incra.get("area")
                    or incra.get("areaHa")
                    or 0
                )
            try:
                area_total += float(area_prop)
            except (ValueError, TypeError):
                pass

        # Resumo textual para MAC
        resumo = _gerar_resumo_grupo(
            len(propriedades), area_total, score_grupo, alertas_consolidados
        )

        return {
            "total_propriedades": len(propriedades),
            "area_total_ha": round(area_total, 2),
            "propriedades": propriedades,
            "alertas_consolidados": alertas_consolidados,
            "score_ambiental_grupo": score_grupo,
            "resumo": resumo,
            "consultado_em": datetime.now(timezone.utc).isoformat(),
        }


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------


def _gerar_resumo_grupo(
    total: int, area_ha: float, score: str, alertas: list
) -> str:
    """Gera texto-resumo do grupo para inclusão no MAC."""
    area_str = f"{area_ha:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")

    if score == "Verde":
        status = (
            "Todas as propriedades analisadas apresentam situação ambiental "
            "regular, sem embargos, sobreposições ou degradação vegetativa "
            "identificados."
        )
    elif score == "Amarelo":
        pontos = [a for a in alertas if "\u26a0" in a]
        status = (
            f"Foram identificados {len(pontos)} ponto(s) de atenção nas "
            f"propriedades analisadas que requerem análise complementar."
        )
    else:
        criticos = [a for a in alertas if "\U0001f534" in a or "Embargo" in a]
        status = (
            f"ATENÇÃO: Foram identificadas {len(criticos)} restrição(ões) "
            f"crítica(s) nas propriedades analisadas. Recomenda-se análise "
            f"aprofundada antes de prosseguir com a operação."
        )

    return (
        f"Análise ambiental de {total} propriedade(s) totalizando "
        f"{area_str} ha. Score ambiental do grupo: {score}. {status}"
    )


def get_client() -> Optional[DadosFazendaClient]:
    """Retorna client autenticado a partir dos secrets do Streamlit.

    Secrets necessários:
        - DADOS_FAZENDA_EMAIL
        - DADOS_FAZENDA_PASSWORD
    """
    try:
        import streamlit as st

        email = st.secrets.get("DADOS_FAZENDA_EMAIL")
        password = st.secrets.get("DADOS_FAZENDA_PASSWORD")
        if not email or not password:
            logger.warning(
                "[DadosFazenda] Credenciais não configuradas em st.secrets"
            )
            return None
        return DadosFazendaClient(email, password)
    except ImportError:
        logger.warning("[DadosFazenda] Streamlit não instalado")
        return None
    except Exception as e:
        logger.warning(f"[DadosFazenda] Não configurado: {e}")
        return None
