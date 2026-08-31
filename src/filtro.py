# -*- coding: utf-8 -*-
"""Filtrado, cruce con el censo de entidades y puntuación automática 0-100.

La puntuación automática NO sustituye al scoring comercial de 13 criterios
del Excel maestro: mide únicamente la relevancia de la señal capturada
(tecnología sin zanja, entidad del censo, CPV de agua, fase de redacción
de proyecto, contexto de redes e importe).
"""
import math

from .util import normalizar


def _contiene(texto_norm, palabras):
    encontradas = [p for p in palabras if p in texto_norm]
    return encontradas


def evaluar(det, ajustes, entidades):
    """Enriquece la detección con matching y score. Devuelve True si es relevante."""
    texto = " ".join(str(det.get(c, "")) for c in ("titulo", "organo", "texto_extra", "expediente"))
    tn = normalizar(texto)

    if _contiene(tn, [normalizar(p) for p in ajustes.get("palabras_excluir", [])]):
        return False

    sz = _contiene(tn, [normalizar(p) for p in ajustes["palabras_sin_zanja"]])
    ctx = _contiene(tn, [normalizar(p) for p in ajustes["palabras_contexto_redes"]])
    rp = _contiene(tn, [normalizar(p) for p in ajustes["palabras_redaccion_proyecto"]])

    cpvs = [str(c) for c in det.get("cpv", [])]
    cpv_ok = any(any(c.startswith(pref) for pref in ajustes["cpv_incluir"]) for c in cpvs)

    organo_n = normalizar(det.get("organo", ""))
    titulo_n = normalizar(det.get("titulo", ""))
    entidad_match = None
    for e in entidades:
        if any(a and (a in organo_n or a in titulo_n) for a in e["alias_norm"]):
            entidad_match = e
            break

    # Criterio de retención: algo tiene que conectarlo con redes de agua
    relevante = bool(sz) or bool(entidad_match and (ctx or cpv_ok or rp)) or (cpv_ok and ctx) or (ctx and rp)
    if not relevante:
        return False

    p = ajustes["pesos_score"]
    score = 0
    if sz:
        score += p["sin_zanja"]
    if entidad_match:
        score += p["entidad_censo"]
    if cpv_ok:
        score += p["cpv_agua"]
    if rp:
        score += p["redaccion_proyecto"]
    if ctx:
        score += p["contexto_redes"]
    importe = det.get("importe_eur")
    if importe:
        tope = float(ajustes.get("importe_max_puntos_eur", 5_000_000))
        score += round(p["importe"] * min(1.0, math.log10(max(importe, 1)) / math.log10(tope)), 1)

    det["score"] = min(100, round(score))
    det["senales"] = {
        "sin_zanja": sz[:5], "contexto_redes": ctx[:5], "redaccion_proyecto": rp[:3],
        "cpv_agua": cpv_ok,
    }
    if entidad_match:
        det["entidad_censo"] = entidad_match["entidad"]
        det["entidad_id"] = entidad_match["id"]
        det["pais"] = entidad_match["pais"]
        det["region"] = entidad_match["region"]
    det["es_redaccion_proyecto"] = bool(rp)
    return True
