# -*- coding: utf-8 -*-
"""Carga histórica del radar desde los datos abiertos de PLACSP.

PLACSP publica cada mes un ZIP con todas las actualizaciones de licitaciones
(mismo formato CODICE que la sindicación diaria):
  perfiles  → .../sindicacion_643/licitacionesPerfilesContratanteCompleto3_AAAAMM.zip
  agregadas → .../sindicacion_1044/PlataformasAgregadasSinMenores_AAAAMM.zip

Este script descarga los meses pedidos, aplica el mismo filtro del radar y
funde el resultado en data/historico.jsonl con las mismas reglas que el barrido
diario (clave por expediente; la información más reciente manda y se conserva
la primera fecha de detección). La fecha de detección de cada registro es la
fecha real de publicación del anuncio, para que los históricos por año queden
bien archivados. No genera alertas: solo repuebla el archivo y regenera las
salidas (Excel Histórico, panel e informe).

Uso (GitHub Actions o local):
  python backfill_historico.py --anio 2025
  python backfill_historico.py --anio 2026 --desde-mes 1 --hasta-mes 8
  python backfill_historico.py --archivo prueba.zip   (para pruebas locales)
"""
import argparse
import io
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import requests

sys.path.insert(0, str(Path(__file__).resolve().parent))

from main import clave_canonica, fusionar  # noqa: E402
from src import filtro, salidas  # noqa: E402
from src.fuentes_placsp import CABECERAS, _local, _procesar_entry  # noqa: E402
from src.util import (ahora_utc, cargar_ajustes, cargar_entidades,  # noqa: E402
                      cargar_historico, guardar_historico)

URLS = {
    "perfiles": "https://contrataciondelestado.es/sindicacion/sindicacion_643/licitacionesPerfilesContratanteCompleto3_{aaaamm}.zip",
    "agregadas": "https://contrataciondelestado.es/sindicacion/sindicacion_1044/PlataformasAgregadasSinMenores_{aaaamm}.zip",
}


def _procesar_zip(ruta_zip, etiqueta, ajustes, entidades, acumulado, registro):
    """Lee todos los .atom del zip, filtra y funde en `acumulado`."""
    leidos = relevantes = 0
    with zipfile.ZipFile(ruta_zip) as z:
        atoms = [n for n in z.namelist() if n.lower().endswith(".atom")]
        for nombre in atoms:
            try:
                raiz = ET.fromstring(z.read(nombre))
            except ET.ParseError as e:
                registro.append(f"[{etiqueta}] atom ilegible {nombre}: {e}")
                continue
            for hijo in raiz:
                if _local(hijo.tag) != "entry":
                    continue
                leidos += 1
                try:
                    d = _procesar_entry(hijo, f"histórico {etiqueta}")
                except Exception:  # noqa: BLE001
                    continue
                if not filtro.evaluar(d, ajustes, entidades):
                    continue
                # la fecha de detección pasa a ser la fecha real de publicación
                fecha_pub = str(d.get("actualizado") or "")[:10]
                if len(fecha_pub) == 10:
                    d["fecha_deteccion"] = fecha_pub
                relevantes += 1
                k = clave_canonica(d)
                d["id"] = k
                acumulado[k] = fusionar(acumulado[k], d) if k in acumulado else d
    registro.append(f"[{etiqueta}] {len(atoms)} atoms, {leidos} entradas, {relevantes} relevantes")
    return leidos, relevantes


TED_API = "https://api.ted.europa.eu/v3/notices/search"


def _ted_texto(v):
    """Los campos de TED llegan multilingües (dict) o como lista; devuelve un texto."""
    if isinstance(v, dict):
        for idioma in ("spa", "por", "eng"):
            if v.get(idioma):
                return _ted_texto(v[idioma])
        v = next(iter(v.values()), "")
    if isinstance(v, list):
        return "; ".join(str(x) for x in v[:3])
    return str(v or "")


def _cargar_ted(anio, mes, ajustes, entidades, acumulado, registro):
    import calendar
    from src.util import nueva_deteccion
    d1 = f"{anio}{mes:02d}01"
    d2 = f"{anio}{mes:02d}{calendar.monthrange(anio, mes)[1]:02d}"
    cpvs = ajustes.get("ted_cpvs") or ajustes.get("cpv_incluir", [])[:6]
    cpv_q = " OR ".join(f"classification-cpv={c}" for c in cpvs)
    q = f"(buyer-country=ESP OR buyer-country=PRT) AND ({cpv_q}) AND publication-date>={d1} AND publication-date<={d2}"
    campos = ["publication-number", "notice-title", "buyer-name", "publication-date",
              "classification-cpv", "deadline-receipt-tenders-date-lot"]
    token, pagina, leidos, relevantes = None, 0, 0, 0
    while pagina < 40:
        cuerpo = {"query": q, "limit": 250, "scope": "ALL", "paginationMode": "ITERATION"}
        if campos:
            cuerpo["fields"] = campos
        if token:
            cuerpo["iterationNextToken"] = token
        try:
            r = requests.post(TED_API, json=cuerpo, timeout=120)
            if r.status_code == 400 and campos:
                campos = None  # algún nombre de campo cambió: reintenta con los de serie
                continue
            r.raise_for_status()
            datos = r.json()
        except Exception as e:  # noqa: BLE001
            registro.append(f"[TED {anio}-{mes:02d}] aviso: {e}")
            break
        avisos = datos.get("notices", [])
        for n in avisos:
            leidos += 1
            num = str(n.get("publication-number", "")).strip()
            titulo = _ted_texto(n.get("notice-title"))
            organo = _ted_texto(n.get("buyer-name"))
            if not num or not titulo:
                continue
            d = nueva_deteccion(
                "TED", num, titulo, organo,
                f"https://ted.europa.eu/es/notice/-/detail/{num}",
                cpv=[str(c) for c in (n.get("classification-cpv") or [])][:8]
                if isinstance(n.get("classification-cpv"), list) else [],
                plazo_presentacion=_ted_texto(n.get("deadline-receipt-tenders-date-lot"))[:10],
                pais="España/Portugal (TED)",
            )
            fecha_pub = _ted_texto(n.get("publication-date"))[:10]
            if len(fecha_pub) == 10:
                d["fecha_deteccion"] = fecha_pub
            if not filtro.evaluar(d, ajustes, entidades):
                continue
            relevantes += 1
            k = clave_canonica(d)
            d["id"] = k
            acumulado[k] = fusionar(acumulado[k], d) if k in acumulado else d
        token = datos.get("iterationNextToken")
        pagina += 1
        if not token or not avisos:
            break
    registro.append(f"[TED {anio}-{mes:02d}] {leidos} avisos, {relevantes} relevantes")


def _descargar(url, registro):
    try:
        resp = requests.get(url, headers=CABECERAS, timeout=600, stream=True)
        if resp.status_code == 404:
            registro.append(f"[descarga] no publicado (404): {url}")
            return None
        resp.raise_for_status()
    except Exception as e:  # noqa: BLE001
        registro.append(f"[descarga] aviso: {url} → {e}")
        return None
    tmp = tempfile.NamedTemporaryFile(suffix=".zip", delete=False)
    for trozo in resp.iter_content(chunk_size=1 << 20):
        tmp.write(trozo)
    tmp.close()
    registro.append(f"[descarga] ok: {url} ({Path(tmp.name).stat().st_size // (1 << 20)} MB)")
    return tmp.name


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--anio", type=int)
    ap.add_argument("--desde-mes", type=int, default=1)
    ap.add_argument("--hasta-mes", type=int, default=12)
    ap.add_argument("--fuente", choices=["todas", "perfiles", "agregadas", "ted"], default="todas")
    ap.add_argument("--archivo", help="zip local (modo prueba)")
    args = ap.parse_args()

    ajustes = cargar_ajustes()
    entidades = cargar_entidades()
    historico = cargar_historico()
    registro = []
    print(f"Histórico de partida: {len(historico)} expedientes")

    if args.archivo:
        _procesar_zip(args.archivo, "archivo local", ajustes, entidades, historico, registro)
    else:
        if not args.anio:
            ap.error("indica --anio o --archivo")
        hoy = ahora_utc()
        if args.fuente == "todas":
            fuentes = ["perfiles", "agregadas", "ted"]
        else:
            fuentes = [args.fuente]
        for mes in range(args.desde_mes, args.hasta_mes + 1):
            aaaamm = f"{args.anio}{mes:02d}"
            for fuente in fuentes:
                if fuente == "ted":
                    if (args.anio, mes) > (hoy.year, hoy.month):
                        continue
                    _cargar_ted(args.anio, mes, ajustes, entidades, historico, registro)
                else:
                    if (args.anio, mes) >= (hoy.year, hoy.month):
                        registro.append(f"[plan] {fuente} {aaaamm} aún sin zip mensual cerrado; lo cubre el barrido diario")
                        continue
                    url = URLS[fuente].format(aaaamm=aaaamm)
                    ruta = _descargar(url, registro)
                    if not ruta:
                        continue
                    _procesar_zip(ruta, f"{fuente} {aaaamm}", ajustes, entidades, historico, registro)
                    Path(ruta).unlink(missing_ok=True)
                guardar_historico(historico)  # guardado incremental: si algo cae, lo hecho queda
                print(f"{fuente} {aaaamm}: histórico ahora {len(historico)} expedientes")

    guardar_historico(historico)
    salidas.generar_excel(historico, [], ajustes)
    salidas.generar_dashboard(historico, {"ejecutado": ahora_utc().strftime("%Y-%m-%d %H:%M"),
                                          "registro": registro, "nuevas_alerta": 0}, ajustes)
    salidas.generar_informe(historico, ajustes)
    print("\n".join(registro))
    print(f"Histórico final: {len(historico)} expedientes. Excel Histórico regenerado.")


if __name__ == "__main__":
    main()
