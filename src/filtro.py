# -*- coding: utf-8 -*-
"""Filtrado, cruce con el censo y puntuación automática (v1.2).

v1.2 — precisión del matching:
- Las palabras clave se buscan como PALABRA COMPLETA, no como fragmento:
  «manga» ya no aparece dentro de «permanganato» ni «mangas neumáticas»,
  ni «mare» dentro de «Miramar», ni «atl» dentro de «atletismo».
- Las entidades del censo se cruzan contra el ÓRGANO de contratación;
  el título solo se usa para alias largos (≥6 caracteres, p. ej.
  «aguasvira»), nunca para siglas cortas propensas a colarse.
- Exclusiones fijas adicionales (festival manga, mangas neumáticas…),
  que se suman a las de config/ajustes.json.
"""
import math
import re

from .util import normalizar

EXCLUSIONES_FIJAS = [
    "festival manga", "manga neumatica", "mangas neumaticas",
    "manga por hombro", "manga de riego",
    "manga larga", "manga corta", "camiseta", "camisetas", "polo de manga",
    "vestuario laboral", "ropa de trabajo", "uniformes",
    "salon del manga", "feria del manga", "salon manga", "manga y anime",
    "comic", "anime", "cosplay", "cultura japonesa",
    "antimedusas", "socorrismo", "salvamento"
]


def _busca(texto_norm, palabra_norm):
    """Coincidencia de palabra/expresión completa (con límites de palabra)."""
    return re.search(r"(?<!\w)" + re.escape(palabra_norm) + r"(?!\w)", texto_norm) is not None


def _contiene(texto_norm, palabras):
    return [p for p in palabras if p and _busca(texto_norm, p)]


def evaluar(det, ajustes, entidades):
    texto = " ".join(str(det.get(c, "")) for c in ("titulo", "organo", "texto_extra", "expediente"))
    tn = normalizar(texto)

    excluir = [normalizar(p) for p in ajustes.get("palabras_excluir", [])] + EXCLUSIONES_FIJAS
    if _contiene(tn, excluir):
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
        for a in e["alias_norm"]:
            if not a:
                continue
            if _busca(organo_n, a) or (len(a) >= 6 and _busca(titulo_n, a)):
                entidad_match = e
                break
        if entidad_match:
            break

    # "manga"/"mangas" son palabras traicioneras (La Manga del Mar Menor, salones
    # del manga...): a secas solo cuentan como señal si el texto habla de redes,
    # tuberías o agua. Las variantes específicas (manga continua, CIPP, curada
    # in situ...) siguen valiendo por sí solas.
    SZ_DEBILES = {"manga", "mangas"}
    if sz and set(sz) <= SZ_DEBILES:
        apoyo = ctx or cpv_ok or _contiene(tn, ["colector", "colectores", "tuberia", "tuberias",
                                                "saneamiento", "alcantarillado", "abastecimiento",
                                                "conduccion", "conducciones", "coletor", "tubagem",
                                                "tubagens", "esgoto"])
        if not apoyo:
            sz = []

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
    det["senales"] = {"sin_zanja": sz[:5], "contexto_redes": ctx[:5],
                      "redaccion_proyecto": rp[:3], "cpv_agua": cpv_ok}
    if entidad_match:
        det["entidad_censo"] = entidad_match["entidad"]
        det["entidad_id"] = entidad_match["id"]
        det["pais"] = entidad_match["pais"]
        det["region"] = entidad_match["region"]
    else:
        for k in ("entidad_censo", "entidad_id"):
            det.pop(k, None)
    det["es_redaccion_proyecto"] = bool(rp)
    return True
