# -*- coding: utf-8 -*-
"""Colector BASE.gov.pt: anuncios de contratación pública de Portugal.

BASE no publica una API documentada; se usa el punto de acceso JSON que
alimenta su propio buscador web. Es el eslabón más frágil del radar: si
BASE cambia su web, este colector avisará en el registro y el resto del
sistema seguirá funcionando (TED cubre igualmente los contratos
portugueses grandes).
"""
import requests

from .util import nueva_deteccion

CABECERAS = {
    "User-Agent": "Mozilla/5.0 (compatible; RadarAguaESPT/1.0)",
    "X-Requested-With": "XMLHttpRequest",
}


def recolectar(ajustes, registro):
    detecciones = []
    f = ajustes["fuentes"]
    endpoint = f.get("base_endpoint")
    if not endpoint:
        return detecciones
    for termino in f.get("base_terminos", []):
        cuerpo = {
            "type": "search_anuncios",
            "version": "121.0",
            "query": f"texto={termino}",
            "sort": "-datapublicacao",
            "page": 0,
            "size": 25,
        }
        try:
            resp = requests.post(endpoint, data=cuerpo, headers=CABECERAS, timeout=60)
            resp.raise_for_status()
            datos = resp.json()
        except Exception as e:  # noqa: BLE001
            registro.append(f"[BASE] aviso: '{termino}' falló ({e}). Si persiste, BASE cambió su web; TED sigue cubriendo Portugal.")
            continue
        items = datos.get("items") or datos.get("resultados") or []
        for it in items:
            num = str(it.get("id") or it.get("nAnuncio") or it.get("numero") or "sin-id")
            titulo = it.get("objectoContrato") or it.get("descricao") or it.get("titulo")
            organo = it.get("entidade") or it.get("adjudicante") or it.get("nomeEntidade")
            importe = it.get("precoBase") or it.get("preco")
            try:
                importe = float(str(importe).replace(".", "").replace(",", ".")) if importe not in (None, "") else None
            except ValueError:
                importe = None
            detecciones.append(nueva_deteccion(
                "BASE", num, titulo, organo,
                f"https://www.base.gov.pt/Base4/pt/detalhe/?type=anuncios&id={num}",
                importe_eur=importe,
                fecha_publicacion=it.get("dataPublicacao") or it.get("data"),
                pais="Portugal",
                texto_extra=str(it.get("modeloAnuncio") or ""),
            ))
        registro.append(f"[BASE] '{termino}': {len(items)} anuncios")
    return detecciones
