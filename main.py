# -*- coding: utf-8 -*-
"""Radar de Licitaciones de Agua ES-PT — programa principal (v1.1).

Novedad v1.1: PLACSP republica cada actualización de un expediente como una
entrada nueva del feed, lo que generaba detecciones duplicadas. Esta versión
funde todas las versiones de un mismo expediente en una sola detección
estable, usando como clave el enlace del expediente (o, si falta, el órgano +
número de expediente). También migra automáticamente el histórico ya
guardado, fundiendo los duplicados acumulados en ejecuciones anteriores.

Uso:
    python main.py            → barrido real (PLACSP + TED + BASE)
    python main.py --demo     → prueba sin internet con datos de ejemplo
"""
import json
import sys

from src import filtro, salidas
from src.util import (RAIZ, ahora_utc, cargar_ajustes, cargar_entidades,
                      cargar_historico, guardar_estado, guardar_historico,
                      normalizar)


def clave_canonica(d):
    """Clave estable por expediente, independiente del evento de publicación."""
    fuente = d.get("fuente", "SRC")
    enlace = (d.get("enlace") or "").strip()
    if enlace:
        return f"{fuente}|{enlace}"
    organo = normalizar(d.get("organo") or "")[:60]
    expediente = normalizar(d.get("expediente") or "")
    if expediente:
        return f"{fuente}|{organo}|{expediente}"
    return f"{fuente}|{organo}|{normalizar(d.get('titulo') or '')[:100]}"


def fusionar(base, nueva):
    """Combina dos versiones del mismo expediente: la información más
    reciente manda, pero se conservan la fecha de primera detección,
    el mejor score y el contador de actualizaciones."""
    resultado = dict(base)
    resultado.update({k: v for k, v in nueva.items() if v not in (None, "", [])})
    fechas = [f for f in (base.get("fecha_deteccion"), nueva.get("fecha_deteccion")) if f]
    if fechas:
        resultado["fecha_deteccion"] = min(fechas)
    resultado["score"] = max(base.get("score", 0) or 0, nueva.get("score", 0) or 0)
    resultado["actualizaciones"] = (base.get("actualizaciones", 1) or 1) + 1
    return resultado


def refundir(detecciones):
    """dict clave_canonica -> detección única (fundiendo duplicados)."""
    unicas = {}
    for d in detecciones:
        k = clave_canonica(d)
        d["id"] = k
        unicas[k] = fusionar(unicas[k], d) if k in unicas else d
    return unicas


def main():
    demo = "--demo" in sys.argv
    ajustes = cargar_ajustes()
    entidades = cargar_entidades()
    registro = []

    # ── Migración: re-clavar el histórico existente y fundir sus duplicados ──
    historico_bruto = cargar_historico()
    historico = refundir(list(historico_bruto.values()))
    fundidas_migracion = len(historico_bruto) - len(historico)
    if fundidas_migracion > 0:
        registro.append(f"[MIGRACIÓN] {fundidas_migracion} duplicado(s) del histórico fundido(s) por expediente")

    # ── Recolección ──
    if demo:
        with open(RAIZ / "demo" / "muestras.json", encoding="utf-8") as f:
            brutas = json.load(f)
        registro.append(f"[DEMO] {len(brutas)} registros de muestra cargados (sin internet)")
    else:
        from src import fuentes_base, fuentes_placsp, fuentes_ted
        brutas = []
        for mod, nombre in ((fuentes_placsp, "PLACSP"), (fuentes_ted, "TED"), (fuentes_base, "BASE")):
            try:
                brutas += mod.recolectar(ajustes, registro)
            except Exception as e:  # noqa: BLE001 — ninguna fuente puede tumbar el radar
                registro.append(f"[{nombre}] aviso: fallo inesperado del colector → {e}")

    # ── Filtrado + fusión del lote de hoy ──
    relevantes = [d for d in brutas if filtro.evaluar(d, ajustes, entidades)]
    lote = refundir(relevantes)
    fundidas_lote = len(relevantes) - len(lote)

    nuevas = [d for k, d in lote.items() if k not in historico]
    for k, d in lote.items():
        if k in historico:
            d["fecha_deteccion"] = historico[k].get("fecha_deteccion", d["fecha_deteccion"])
            historico[k] = fusionar(historico[k], d)
        else:
            historico[k] = d

    # ── Salidas ──
    n_alertas = salidas.generar_alertas(nuevas, ajustes)
    estado = {
        "ejecutado": ahora_utc().strftime("%Y-%m-%d %H:%M"),
        "modo": "demo" if demo else "real",
        "brutas": len(brutas), "relevantes": len(relevantes),
        "fundidas_hoy": fundidas_lote, "nuevas": len(nuevas),
        "nuevas_alerta": n_alertas, "historico_total": len(historico),
        "registro": registro,
    }
    guardar_historico(historico)
    guardar_estado(estado)
    salidas.generar_excel(historico, nuevas, ajustes)
    salidas.generar_informe(historico, ajustes)
    salidas.generar_dashboard(historico, estado, ajustes)

    print(json.dumps({k: v for k, v in estado.items() if k != "registro"}, ensure_ascii=False, indent=2))
    for linea in registro:
        print(" ", linea)


if __name__ == "__main__":
    main()
