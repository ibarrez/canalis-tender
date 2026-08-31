# -*- coding: utf-8 -*-
"""Colector TED: Search API pública (sin autenticación) de Tenders Electronic Daily."""
import requests

from .util import nueva_deteccion

CABECERAS = {"User-Agent": "RadarAguaESPT/1.0", "Content-Type": "application/json"}


def _campo(aviso, *nombres):
    """Los campos de TED pueden venir como str, lista o dict multilingüe."""
    for n in nombres:
        v = aviso.get(n)
        if v in (None, "", []):
            continue
        if isinstance(v, dict):  # {'spa': [...], 'eng': [...]} o similar
            for pref in ("spa", "por", "eng"):
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
    cuerpo = {
        "query": f.get("ted_query", ""),
        "fields": f.get("ted_fields", ["publication-number", "notice-title", "buyer-name", "publication-date"]),
        "limit": int(f.get("ted_limit", 100)),
        "page": 1,
    }
    try:
        resp = requests.post(f["ted_endpoint"], json=cuerpo, headers=CABECERAS, timeout=60)
        resp.raise_for_status()
        datos = resp.json()
    except Exception as e:  # noqa: BLE001
        registro.append(f"[TED] aviso: la consulta falló ({e}). Revisa 'ted_query' en config/ajustes.json — la sintaxis de la Search API cambia ocasionalmente.")
        return detecciones
    avisos = datos.get("notices") or datos.get("results") or []
    for a in avisos:
        num = _campo(a, "publication-number", "ND") or "sin-num"
        titulo = _campo(a, "notice-title", "TI")
        comprador = _campo(a, "buyer-name", "AU")
        importe = _campo(a, "estimated-value-lot", "estimated-value")
        try:
            importe = float(str(importe).replace(",", ".")) if importe is not None else None
        except ValueError:
            importe = None
        lugar = _campo(a, "place-of-performance") or ""
        pais = "Portugal" if "PRT" in str(lugar) or str(lugar).startswith("PT") else "España"
        detecciones.append(nueva_deteccion(
            "TED", num, titulo, comprador,
            f"https://ted.europa.eu/es/notice/-/detail/{num}",
            importe_eur=importe,
            plazo_presentacion=_campo(a, "deadline-date", "deadline-receipt-tenders-date-time"),
            fecha_publicacion=_campo(a, "publication-date", "PD"),
            cpv=[c for c in (a.get("classification-cpv") or []) if c] if isinstance(a.get("classification-cpv"), list) else ([a["classification-cpv"]] if a.get("classification-cpv") else []),
            pais=pais,
        ))
    registro.append(f"[TED] {len(avisos)} avisos recibidos")
    return detecciones
