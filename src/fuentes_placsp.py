# -*- coding: utf-8 -*-
"""Colector PLACSP: lee la sindicación ATOM diaria (formato CODICE) y extrae licitaciones.

Cubre dos feeds: perfiles alojados en PLACSP y plataformas agregadas
(que incluyen Cataluña, Euskadi, Navarra y otras plataformas propias).
Los feeds se encadenan mediante <link rel="next">; se recorren hasta
`max_paginas_atom` páginas o hasta superar la fecha del último barrido.
"""
import xml.etree.ElementTree as ET

import requests

from .util import nueva_deteccion, extraer_importe

CABECERAS = {"User-Agent": "RadarAguaESPT/1.0 (herramienta interna de vigilancia de licitaciones)"}


def _local(tag):
    return tag.split("}")[-1] if "}" in tag else tag


def _buscar_texto(nodo, nombres):
    """Primer texto de cualquier descendiente cuyo nombre local esté en `nombres`."""
    objetivo = set(nombres)
    for el in nodo.iter():
        if _local(el.tag) in objetivo and el.text and el.text.strip():
            return el.text.strip()
    return None


def _todos_textos(nodo, nombre):
    return [el.text.strip() for el in nodo.iter() if _local(el.tag) == nombre and el.text and el.text.strip()]


def _procesar_entry(entry, feed_nombre):
    entry_id = _buscar_texto(entry, {"id"}) or ""
    titulo = None
    enlace = ""
    actualizado = None
    for hijo in entry:
        ln = _local(hijo.tag)
        if ln == "title" and hijo.text:
            titulo = hijo.text.strip()
        elif ln == "link":
            enlace = hijo.get("href", enlace)
        elif ln == "updated" and hijo.text:
            actualizado = hijo.text.strip()

    organo = _buscar_texto(entry, {"PartyName", "Name"})  # órgano de contratación (CODICE)
    estado = _buscar_texto(entry, {"ContractFolderStatusCode"})
    expediente = _buscar_texto(entry, {"ContractFolderID"})
    importe_txt = _buscar_texto(entry, {"EstimatedOverallContractAmount", "TotalAmount", "TaxExclusiveAmount"})
    importe = None
    if importe_txt:
        try:
            importe = float(importe_txt.replace(",", "."))
        except ValueError:
            importe = extraer_importe(importe_txt)
    cpvs = _todos_textos(entry, "ItemClassificationCode")
    plazo = _buscar_texto(entry, {"EndDate"})
    resumen = _buscar_texto(entry, {"summary"}) or ""

    return nueva_deteccion(
        "PLACSP", entry_id or (expediente or titulo or "sin-id"),
        titulo, organo, enlace,
        expediente=expediente, estado=estado, importe_eur=importe,
        cpv=sorted(set(cpvs))[:8], plazo_presentacion=plazo,
        actualizado=actualizado, subfuente=feed_nombre, texto_extra=resumen,
        pais="España",
    )


def recolectar(ajustes, registro):
    """Devuelve lista de detecciones en bruto (sin filtrar)."""
    detecciones = []
    fuentes = ajustes["fuentes"]
    max_pag = int(ajustes.get("max_paginas_atom", 20))
    feeds = [("perfiles PLACSP", fuentes.get("placsp_perfiles")),
             ("plataformas agregadas", fuentes.get("placsp_agregadas"))]
    for nombre, url in feeds:
        if not url:
            continue
        paginas = 0
        while url and paginas < max_pag:
            try:
                resp = requests.get(url, headers=CABECERAS, timeout=60)
                resp.raise_for_status()
                raiz = ET.fromstring(resp.content)
            except Exception as e:  # noqa: BLE001 — un feed caído no debe parar el radar
                registro.append(f"[PLACSP:{nombre}] aviso: no se pudo leer {url} → {e}")
                break
            n_entries = 0
            siguiente = None
            for hijo in raiz:
                ln = _local(hijo.tag)
                if ln == "entry":
                    try:
                        detecciones.append(_procesar_entry(hijo, nombre))
                        n_entries += 1
                    except Exception as e:  # noqa: BLE001
                        registro.append(f"[PLACSP:{nombre}] entry ilegible: {e}")
                elif ln == "link" and hijo.get("rel") == "next":
                    siguiente = hijo.get("href")
            registro.append(f"[PLACSP:{nombre}] página {paginas + 1}: {n_entries} entradas")
            url = siguiente
            paginas += 1
    return detecciones
