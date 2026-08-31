# -*- coding: utf-8 -*-
"""Colector TED v1.1: Search API pública de Tenders Electronic Daily.

Sintaxis verificada de la API v3 (POST https://api.ted.europa.eu/v3/notices/search):
  body: {"query": "...", "fields": [...], "limit": N, "scope": "ACTIVE",
         "paginationMode": "ITERATION", "checkQuerySyntax": false}
  query: expresiones tipo  buyer-country=ESP · PD>=YYYYMMDD · FT~"texto"
         · classification-cpv=45231300 · AND/OR · SORT BY publication-date DESC

Robustez: se prueba una escalera de consultas (config → CPV → texto libre)
hasta que una responda; se registra cuál funcionó y, si ninguna lo hace,
el error exacto que devolvió la API para poder diagnosticarlo.
"""
from datetime import timedelta

import requests

from .util import ahora_utc, nueva_deteccion

CABECERAS = {"User-Agent": "RadarAguaESPT/1.1", "Content-Type": "application/json"}

CAMPOS = ["publication-number", "notice-title", "buyer-name", "buyer-country",
          "publication-date", "classification-cpv", "total-value", "deadline"]

CPVS = ["45231300", "45232150", "45232400", "45232410", "45232420", "45232440",
        "71322000", "90400000"]

_QUERY_ROTA_ANTIGUA = "place-of-performance IN (ESP PRT)"  # la del 400 de la v1.0


def _consultas_candidatas(ajustes):
    desde = (ahora_utc() - timedelta(days=4)).strftime("%Y%m%d")
    pais = "(buyer-country=ESP OR buyer-country=PRT)"
    candidatas = []
    cfg = (ajustes["fuentes"].get("ted_query") or "").strip()
    if cfg and _QUERY_ROTA_ANTIGUA not in cfg:
        candidatas.append(("config", cfg))
    cpv = " OR ".join(f"classification-cpv={c}" for c in CPVS)
    candidatas.append(("CPV", f"({cpv}) AND {pais} AND PD>={desde} SORT BY publication-date DESC"))
    texto = ' OR '.join(f'FT~"{t}"' for t in (
        "rehabilitación de colectores", "renovación de redes", "manga continua",
        "sin zanja", "reabilitação de coletores", "renovação de condutas"))
    candidatas.append(("texto libre", f"({texto}) AND {pais} AND PD>={desde} SORT BY publication-date DESC"))
    return candidatas


def _campo(aviso, *nombres):
    """Los campos de TED pueden venir como str, número, lista o dict multilingüe."""
    for n in nombres:
        v = aviso.get(n)
        if v in (None, "", []):
            continue
        if isinstance(v, dict):  # {'spa': [...], 'eng': '...'}
            for pref in ("spa", "por", "cat", "eng"):
                if v.get(pref):
                    x = v[pref]
                    return x[0] if isinstance(x, list) else x
            primero = next(iter(v.values()))
            return primero[0] if isinstance(primero, list) else primero
        if isinstance(v, list):
            return v[0]
        return v
    return None


def recolectar(ajustes, registro):
    detecciones = []
    f = ajustes["fuentes"]
    endpoint = f.get("ted_endpoint", "https://api.ted.europa.eu/v3/notices/search")
    limite = int(f.get("ted_limit", 100))

    datos = None
    for nombre, consulta in _consultas_candidatas(ajustes):
        cuerpo = {"query": consulta, "fields": CAMPOS, "limit": limite,
                  "scope": "ACTIVE", "paginationMode": "ITERATION",
                  "checkQuerySyntax": False}
        try:
            resp = requests.post(endpoint, json=cuerpo, headers=CABECERAS, timeout=60)
            if resp.status_code != 200:
                registro.append(f"[TED] consulta '{nombre}' rechazada ({resp.status_code}): {resp.text[:180]}")
                continue
            datos = resp.json()
            if not isinstance(datos, dict):
                registro.append(f"[TED] consulta '{nombre}': respuesta inesperada ({type(datos).__name__})")
                datos = None
                continue
            registro.append(f"[TED] consulta '{nombre}' aceptada")
            break
        except Exception as e:  # noqa: BLE001
            registro.append(f"[TED] consulta '{nombre}' falló: {e}")
    if datos is None:
        registro.append("[TED] aviso: ninguna consulta funcionó; PLACSP sigue cubriendo España. Revisar sintaxis en docs.ted.europa.eu")
        return detecciones

    avisos = datos.get("notices") or datos.get("results") or []
    for a in avisos:
        if not isinstance(a, dict):
            continue
        num = _campo(a, "publication-number", "ND") or "sin-num"
        pais_cod = str(_campo(a, "buyer-country") or "")
        pais = "Portugal" if "PRT" in pais_cod else "España"
        importe = _campo(a, "total-value")
        try:
            importe = float(str(importe).replace(",", ".")) if importe is not None else None
        except ValueError:
            importe = None
        cpv_raw = a.get("classification-cpv")
        if isinstance(cpv_raw, list):
            cpvs = [str(c) for c in cpv_raw if c]
        elif cpv_raw:
            cpvs = [s.strip() for s in str(cpv_raw).split(";") if s.strip()]
        else:
            cpvs = []
        detecciones.append(nueva_deteccion(
            "TED", num,
            _campo(a, "notice-title", "TI"),
            _campo(a, "buyer-name", "AU"),
            f"https://ted.europa.eu/es/notice/-/detail/{num}",
            importe_eur=importe,
            plazo_presentacion=_campo(a, "deadline"),
            fecha_publicacion=_campo(a, "publication-date", "PD"),
            cpv=cpvs[:8],
            pais=pais,
        ))
    registro.append(f"[TED] {len(avisos)} avisos recibidos")
    return detecciones
