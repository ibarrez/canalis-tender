# -*- coding: utf-8 -*-
"""Utilidades comunes del Radar de Licitaciones de Agua ES-PT."""
import csv
import json
import re
import unicodedata
from datetime import datetime, timezone
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
CONFIG = RAIZ / "config"
DATA = RAIZ / "data"
DOCS = RAIZ / "docs"


def ahora_utc():
    return datetime.now(timezone.utc)


def normalizar(texto):
    """minúsculas + sin tildes, para comparaciones robustas."""
    if not texto:
        return ""
    t = unicodedata.normalize("NFD", str(texto).lower())
    return "".join(c for c in t if unicodedata.category(c) != "Mn")


def cargar_ajustes():
    with open(CONFIG / "ajustes.json", encoding="utf-8") as f:
        return json.load(f)


def cargar_entidades():
    entidades = []
    with open(CONFIG / "entidades.csv", encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            fila["alias_norm"] = [normalizar(a) for a in fila["aliases"].split("|") if a.strip()]
            entidades.append(fila)
    return entidades


def extraer_importe(texto):
    """Devuelve el mayor importe en EUR hallado en un texto, o None."""
    if not texto:
        return None
    candidatos = []
    for m in re.finditer(r"(\d{1,3}(?:[.\s]\d{3})+(?:,\d+)?|\d+(?:,\d+)?)(?=\s*(?:eur|€))", texto.lower()):
        bruto = m.group(1).replace(".", "").replace(" ", "").replace(",", ".")
        try:
            candidatos.append(float(bruto))
        except ValueError:
            pass
    return max(candidatos) if candidatos else None


def cargar_historico():
    """Histórico como dict id -> detección (última versión)."""
    ruta = DATA / "historico.jsonl"
    historico = {}
    if ruta.exists():
        with open(ruta, encoding="utf-8") as f:
            for linea in f:
                linea = linea.strip()
                if linea:
                    try:
                        d = json.loads(linea)
                        historico[d["id"]] = d
                    except json.JSONDecodeError:
                        continue
    return historico


def guardar_historico(historico):
    DATA.mkdir(exist_ok=True)
    with open(DATA / "historico.jsonl", "w", encoding="utf-8") as f:
        for d in sorted(historico.values(), key=lambda x: x.get("fecha_deteccion", ""), reverse=True):
            f.write(json.dumps(d, ensure_ascii=False) + "\n")


def guardar_estado(estado):
    DATA.mkdir(exist_ok=True)
    with open(DATA / "estado.json", "w", encoding="utf-8") as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)


def nueva_deteccion(fuente, id_externo, titulo, organo, enlace, **extra):
    d = {
        "id": f"{fuente}:{id_externo}",
        "fuente": fuente,
        "titulo": (titulo or "").strip(),
        "organo": (organo or "").strip(),
        "enlace": enlace or "",
        "fecha_deteccion": ahora_utc().strftime("%Y-%m-%d"),
    }
    d.update({k: v for k, v in extra.items() if v not in (None, "", [])})
    return d
