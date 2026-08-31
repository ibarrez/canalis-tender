# -*- coding: utf-8 -*-
"""Radar de Licitaciones de Agua ES-PT — programa principal.

Uso:
    python main.py            → barrido real (PLACSP + TED + BASE)
    python main.py --demo     → prueba sin internet con datos de ejemplo

Cada ejecución: recolecta → filtra y puntúa → deduplica contra el histórico →
regenera Excel, dashboard (docs/index.html), informe y alertas.
"""
import json
import sys

from src import filtro, salidas
from src.util import (RAIZ, ahora_utc, cargar_ajustes, cargar_entidades,
                      cargar_historico, guardar_estado, guardar_historico)


def main():
    demo = "--demo" in sys.argv
    ajustes = cargar_ajustes()
    entidades = cargar_entidades()
    historico = cargar_historico()
    registro = []

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

    relevantes = [d for d in brutas if filtro.evaluar(d, ajustes, entidades)]
    nuevas = [d for d in relevantes if d["id"] not in historico]
    for d in relevantes:  # las repetidas actualizan su versión (estado, plazos…)
        previo = historico.get(d["id"])
        if previo:
            d["fecha_deteccion"] = previo.get("fecha_deteccion", d["fecha_deteccion"])
        historico[d["id"]] = d

    n_alertas = salidas.generar_alertas(nuevas, ajustes)
    estado = {
        "ejecutado": ahora_utc().strftime("%Y-%m-%d %H:%M"),
        "modo": "demo" if demo else "real",
        "brutas": len(brutas), "relevantes": len(relevantes),
        "nuevas": len(nuevas), "nuevas_alerta": n_alertas,
        "historico_total": len(historico), "registro": registro,
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
