# -*- coding: utf-8 -*-
"""Colector BASE.gov.pt v1.1: anuncios de contratación pública de Portugal.

BASE no publica una API documentada; se usa el punto de acceso JSON que
alimenta su buscador web. v1.1: tolera respuestas vacías (null), listas con
elementos nulos o cambios de estructura — cualquier anomalía se registra y
el radar continúa (TED cubre igualmente los contratos portugueses grandes).
"""
import requests

from .util import nueva_deteccion

CABECERAS = {
    "User-Agent": "Mozilla/5.0 (compatible; RadarAguaESPT/1.1)",
    "X-Requested-With": "XMLHttpRequest",
    "Accept": "application/json",
    "Referer": "https://www.base.gov.pt/Base4/pt/pesquisa/",
}


def recolectar(ajustes, registro):
    detecciones = []
    f = ajustes["fuentes"]
    endpoint = f.get("base_endpoint")
    if not endpoint:
        return detecciones
    vacias = 0
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
            if not isinstance(datos, dict):
                vacias += 1
                registro.append(f"[BASE] '{termino}': respuesta sin datos ({type(datos).__name__})")
                continue
            items = datos.get("items") or datos.get("resultados") or []
            if not isinstance(items, list):
                registro.append(f"[BASE] '{termino}': formato de items inesperado")
                continue
            n_ok = 0
            for it in items:
                if not isinstance(it, dict):
                    continue
                num = str(it.get("id") or it.get("nAnuncio") or it.get("numero") or "").strip()
                titulo = it.get("objectoContrato") or it.get("descricao") or it.get("titulo")
                if not num or not titulo:
                    continue
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
                n_ok += 1
            registro.append(f"[BASE] '{termino}': {n_ok} anuncios")
        except Exception as e:  # noqa: BLE001
            registro.append(f"[BASE] '{termino}' falló: {type(e).__name__}: {e}")
            continue
    if vacias and vacias == len(f.get("base_terminos", [])):
        registro.append("[BASE] aviso: todas las consultas volvieron vacías — BASE pudo cambiar su web; TED sigue cubriendo Portugal")
    return detecciones
