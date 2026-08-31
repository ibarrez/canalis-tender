# -*- coding: utf-8 -*-
"""Salidas del radar: Excel de detecciones, dashboard HTML (GitHub Pages),
informe de los últimos N días y fichero de alertas para el aviso por Issue."""
from datetime import datetime, timedelta

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .util import DATA, DOCS, ahora_utc

AZUL = "1F3864"


def _hoja(ws, titulo, columnas, anchos):
    ws.append([])
    for i, (h, w) in enumerate(zip(columnas, anchos), 1):
        c = ws.cell(1, i, h)
        c.font = Font(name="Arial", size=9, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=AZUL)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


def _fila_det(d):
    s = d.get("senales", {})
    return [
        d.get("fecha_deteccion"), d.get("score"), d.get("fuente"),
        d.get("pais", ""), d.get("entidad_censo") or d.get("organo", ""),
        d.get("titulo", ""), d.get("importe_eur"),
        "Sí" if d.get("es_redaccion_proyecto") else "",
        ", ".join(s.get("sin_zanja", [])),
        d.get("plazo_presentacion", ""), d.get("estado", ""),
        ", ".join(d.get("cpv", [])[:5]), d.get("enlace", ""),
    ]


COLS = ["Detectado", "Score", "Fuente", "País", "Entidad / órgano", "Objeto",
        "Importe (€)", "Redacción proyecto", "Señales sin zanja",
        "Plazo presentación", "Estado", "CPV", "Enlace"]
ANCHOS = [11, 7, 8, 9, 28, 55, 13, 10, 22, 14, 14, 18, 45]


def generar_excel(historico, nuevas, ajustes):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Alertas nuevas"
    _hoja(ws, ws.title, COLS, ANCHOS)
    umbral = ajustes.get("umbral_alerta", 60)
    for d in sorted(nuevas, key=lambda x: -x.get("score", 0)):
        if d.get("score", 0) >= umbral:
            ws.append(_fila_det(d))
    ws2 = wb.create_sheet("Todas las detecciones")
    _hoja(ws2, ws2.title, COLS, ANCHOS)
    for d in sorted(historico.values(), key=lambda x: (x.get("fecha_deteccion", ""), x.get("score", 0)), reverse=True):
        ws2.append(_fila_det(d))
    for hoja in (ws, ws2):
        for fila in hoja.iter_rows(min_row=2):
            for c in fila:
                c.font = Font(name="Arial", size=9)
                c.alignment = Alignment(wrap_text=True, vertical="top")
            if fila[0].row and hoja.cell(fila[0].row, 7).value:
                hoja.cell(fila[0].row, 7).number_format = '#,##0 "€"'
        hoja.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{hoja.max_row}"
    DATA.mkdir(exist_ok=True)
    wb.save(DATA / "Radar_Detecciones.xlsx")


def generar_informe(historico, ajustes):
    dias = int(ajustes.get("dias_ventana_informe", 7))
    corte = (ahora_utc() - timedelta(days=dias)).strftime("%Y-%m-%d")
    recientes = [d for d in historico.values() if d.get("fecha_deteccion", "") >= corte]
    top = sorted(recientes, key=lambda x: -x.get("score", 0))[:10]
    lineas = [f"# Informe del radar — {ahora_utc().strftime('%d/%m/%Y')}",
              f"\nDetecciones de los últimos {dias} días: **{len(recientes)}**. Top 10 por puntuación:\n"]
    if not top:
        lineas.append("_Sin detecciones relevantes en la ventana. Revisa data/estado.json por si alguna fuente falló._")
    for i, d in enumerate(top, 1):
        imp = f" · {d['importe_eur']:,.0f} €".replace(",", ".") if d.get("importe_eur") else ""
        rp = " · **REDACCIÓN DE PROYECTO**" if d.get("es_redaccion_proyecto") else ""
        sz = f" · sin zanja: {', '.join(d['senales']['sin_zanja'])}" if d.get("senales", {}).get("sin_zanja") else ""
        lineas.append(f"{i}. **[{d.get('score', 0)}]** {d.get('entidad_censo') or d.get('organo', '¿?')} — "
                      f"{d.get('titulo', '')[:180]}{imp}{rp}{sz} · [{d.get('fuente')}]({d.get('enlace', '')})")
    lineas.append("\n---\n_Radar automático: verifica siempre fechas e importes en la fuente oficial antes de actuar. "
                  "Traslada las oportunidades maduras al Excel maestro (ficha comercial + scoring de 13 criterios)._")
    contenido = "\n".join(lineas)
    (DATA / "informe.md").write_text(contenido, encoding="utf-8")
    return contenido


def generar_alertas(nuevas, ajustes):
    """Escribe data/alertas_nuevas.md solo si hay novedades ≥ umbral (dispara el Issue)."""
    ruta = DATA / "alertas_nuevas.md"
    if ruta.exists():
        ruta.unlink()
    umbral = ajustes.get("umbral_alerta", 60)
    fuertes = sorted([d for d in nuevas if d.get("score", 0) >= umbral], key=lambda x: -x["score"])
    if not fuertes:
        return 0
    lineas = [f"## {len(fuertes)} detección(es) nueva(s) con score ≥ {umbral}\n"]
    for d in fuertes[:25]:
        imp = f" · {d['importe_eur']:,.0f} €".replace(",", ".") if d.get("importe_eur") else ""
        lineas.append(f"- **[{d['score']}]** {d.get('entidad_censo') or d.get('organo', '¿?')} — "
                      f"{d.get('titulo', '')[:200]}{imp} → {d.get('enlace', '')}")
    lineas.append("\nRevisar, verificar en fuente oficial y pasar al Excel maestro si procede.")
    ruta.write_text("\n".join(lineas), encoding="utf-8")
    return len(fuertes)


def _fila_html(d):
    imp = f"{d['importe_eur']:,.0f} €".replace(",", ".") if d.get("importe_eur") else "—"
    rp = "📐" if d.get("es_redaccion_proyecto") else ""
    sz = "🟢" if d.get("senales", {}).get("sin_zanja") else ""
    ent = d.get("entidad_censo") or d.get("organo", "")
    return (f"<tr><td class='sc'>{d.get('score', 0)}</td><td>{d.get('fecha_deteccion', '')}</td>"
            f"<td>{d.get('fuente', '')}</td><td>{ent[:60]}</td>"
            f"<td><a href='{d.get('enlace', '#')}' target='_blank'>{(d.get('titulo') or '')[:150]}</a> {rp}{sz}</td>"
            f"<td class='imp'>{imp}</td></tr>")


def generar_dashboard(historico, estado, ajustes):
    DOCS.mkdir(exist_ok=True)
    todas = sorted(historico.values(), key=lambda x: (-x.get("score", 0), x.get("fecha_deteccion", "")))
    corte7 = (ahora_utc() - timedelta(days=7)).strftime("%Y-%m-%d")
    nuevas7 = [d for d in historico.values() if d.get("fecha_deteccion", "") >= corte7]
    redaccion = [d for d in todas if d.get("es_redaccion_proyecto")]
    filas_top = "\n".join(_fila_html(d) for d in todas[:50])
    filas_rp = "\n".join(_fila_html(d) for d in redaccion[:30]) or "<tr><td colspan='6'>Sin señales de redacción de proyecto todavía.</td></tr>"
    avisos = "".join(f"<li>{a}</li>" for a in estado.get("registro", []) if "aviso" in a)
    html = f"""<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Radar de Licitaciones de Agua ES-PT</title>
<style>
 body{{font-family:Arial,Helvetica,sans-serif;margin:0;background:#f5f7fa;color:#1a1a2e}}
 header{{background:#1F3864;color:#fff;padding:22px 28px}}
 header h1{{margin:0;font-size:21px}} header p{{margin:6px 0 0;font-size:13px;opacity:.85}}
 .kpis{{display:flex;gap:14px;padding:18px 28px;flex-wrap:wrap}}
 .kpi{{background:#fff;border-radius:10px;padding:14px 20px;box-shadow:0 1px 4px rgba(0,0,0,.08);min-width:150px}}
 .kpi b{{font-size:26px;color:#1F3864;display:block}} .kpi span{{font-size:12px;color:#555}}
 section{{margin:8px 28px 26px}} h2{{font-size:16px;color:#1F3864}}
 table{{width:100%;border-collapse:collapse;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.08);font-size:13px}}
 th{{background:#2E5AA8;color:#fff;text-align:left;padding:8px 10px;font-size:12px}}
 td{{padding:7px 10px;border-top:1px solid #eef1f6;vertical-align:top}}
 td.sc{{font-weight:bold;color:#1F3864;text-align:center}} td.imp{{white-space:nowrap;text-align:right}}
 a{{color:#2E5AA8;text-decoration:none}} a:hover{{text-decoration:underline}}
 .avisos{{background:#FFF2CC;border-radius:10px;padding:10px 16px;font-size:12px}}
 footer{{padding:14px 28px 30px;font-size:11px;color:#777}}
 .leyenda{{font-size:12px;color:#555;margin:6px 0 12px}}
</style></head><body>
<header><h1>💧 Radar de Licitaciones de Agua · España y Portugal</h1>
<p>Vigilancia automática de rehabilitación y renovación de redes (tecnologías sin zanja) · Última ejecución: {estado.get('ejecutado', '—')} UTC</p></header>
<div class="kpis">
 <div class="kpi"><b>{len(historico)}</b><span>detecciones acumuladas</span></div>
 <div class="kpi"><b>{len(nuevas7)}</b><span>en los últimos 7 días</span></div>
 <div class="kpi"><b>{len(redaccion)}</b><span>señales de redacción de proyecto</span></div>
 <div class="kpi"><b>{estado.get('nuevas_alerta', 0)}</b><span>alertas nuevas (score ≥ {ajustes.get('umbral_alerta', 60)})</span></div>
</div>
<section><h2>Top 50 por puntuación</h2>
<div class="leyenda">📐 = fase de redacción de proyecto (señal temprana) · 🟢 = menciona tecnología sin zanja · el score automático mide relevancia de la señal, no sustituye al análisis comercial</div>
<table><tr><th>Score</th><th>Detectado</th><th>Fuente</th><th>Entidad / órgano</th><th>Objeto</th><th>Importe</th></tr>
{filas_top}</table></section>
<section><h2>📐 Radar de redacciones de proyecto (la señal más temprana)</h2>
<table><tr><th>Score</th><th>Detectado</th><th>Fuente</th><th>Entidad / órgano</th><th>Objeto</th><th>Importe</th></tr>
{filas_rp}</table></section>
{f'<section><h2>Avisos de fuentes</h2><div class="avisos"><ul>{avisos}</ul></div></section>' if avisos else ''}
<footer>Generado automáticamente. Verifica siempre fechas e importes en la fuente oficial (PLACSP · TED · BASE.gov.pt) antes de cualquier acción comercial.
El Excel completo está en <code>data/Radar_Detecciones.xlsx</code> del repositorio; el informe semanal en <code>data/informe.md</code>.</footer>
</body></html>"""
    (DOCS / "index.html").write_text(html, encoding="utf-8")
