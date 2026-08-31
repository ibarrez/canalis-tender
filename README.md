# 💧 Radar de Licitaciones de Agua · España y Portugal

Sistema **automático y gratuito** de vigilancia de licitaciones de renovación y
rehabilitación de redes de agua (con foco en tecnologías sin zanja: CIPP, manga,
sliplining, close-fit, pipe bursting).

Cada mañana laborable, sin que nadie toque nada:

1. Lee la **sindicación oficial de PLACSP** (perfiles + plataformas agregadas, que
   incluyen Cataluña, Euskadi y Navarra), la **Search API de TED** y **BASE.gov.pt**.
2. Filtra por palabras clave ES/PT, códigos CPV de agua y el **censo de 84 entidades**
   gestoras (`config/entidades.csv`).
3. Puntúa cada detección de 0 a 100 y marca las señales tempranas de
   **redacción de proyecto** 📐.
4. Actualiza un **panel web** (GitHub Pages), un **Excel** (`data/Radar_Detecciones.xlsx`),
   un **informe de los últimos 7 días** (`data/informe.md`) y, si hay novedades con
   score ≥ 60, **abre un Issue** → GitHub te envía un **correo automático**.

---

## 🚀 Puesta en marcha (una sola vez, ~20 minutos, sin programar)

1. **Crea una cuenta gratuita en [github.com](https://github.com)** (si no tienes).
2. Arriba a la derecha, pulsa **“+” → New repository**. Nombre: `radar-agua`.
   Déjalo **Public** (necesario para que Pages y Actions sean 100% gratuitos) y pulsa
   **Create repository**.
3. En el repositorio recién creado: **“uploading an existing file”** (o Add file →
   Upload files). **Arrastra TODO el contenido de esta carpeta** (incluida la carpeta
   `.github` — si tu ordenador la oculta, sube el ZIP descomprimido completo).
   Pulsa **Commit changes**.
4. Activa el panel web: **Settings → Pages → Branch: `main` / carpeta `/docs` → Save**.
   En unos minutos tu panel estará en `https://TU-USUARIO.github.io/radar-agua/`.
5. Activa las alertas por correo: en la portada del repositorio pulsa
   **Watch → All activity**. Cada Issue de alerta te llegará al email.
6. Primera ejecución manual: pestaña **Actions → Radar diario → Run workflow**.
   (Si GitHub pregunta si habilitas los workflows, acepta.)
   En 2-3 minutos verás el resultado; a partir de ahí corre solo cada mañana laborable.

Eso es todo. Coste: **0 €** (dentro de los límites gratuitos de GitHub, que este
radar no llega ni a rozar).

## 🔧 Ajustes sin tocar código

- **Palabras clave, CPVs, umbral de alerta y pesos del score** → edita
  `config/ajustes.json` desde la web de GitHub (icono del lápiz ✏️).
- **Añadir o quitar entidades del censo** → edita `config/entidades.csv`
  (columna `aliases`: variantes del nombre separadas por `|`, sin tildes).
- **Hora de ejecución** → línea `cron` en `.github/workflows/radar.yml`.

## 📁 Qué genera cada día

| Archivo | Contenido |
|---|---|
| `docs/index.html` | Panel web con KPIs, Top 50 y radar de redacciones de proyecto |
| `data/Radar_Detecciones.xlsx` | Excel: pestaña de alertas nuevas + histórico completo |
| `data/informe.md` | Informe con el Top 10 de los últimos 7 días |
| `data/historico.jsonl` | Base de datos histórica (deduplicada por ID de expediente) |
| `data/estado.json` | Diagnóstico de la última ejecución (qué fuente falló, si alguna) |

## 🔍 Probarlo sin internet

```
python main.py --demo
```
Usa `demo/muestras.json` y genera todas las salidas con datos de ejemplo.

## ⚠️ Honestidad sobre los límites

- **Lo que cubre:** todo lo publicado oficialmente en PLACSP (incl. agregadas),
  TED y BASE — licitaciones, anuncios previos, consultas preliminares y
  redacciones de proyecto.
- **Lo que NO cubre (capa humana):** actas de plenos, planes directores en PDF,
  notas de prensa y contactos comerciales. Esas señales se siguen trabajando en el
  **Excel maestro** (ficha comercial + scoring de 13 criterios) con la rutina
  semanal.
- **Fragilidad conocida:** BASE.gov.pt no tiene API oficial; si cambia su web, el
  colector lo avisará en `data/estado.json` y en el panel, y TED seguirá cubriendo
  los contratos portugueses grandes. La sintaxis de la query de TED
  (`ted_query` en `config/ajustes.json`) puede requerir un ajuste puntual si la
  API evoluciona — el error exacto quedará registrado en `data/estado.json`.
- El **score automático (0-100)** mide la relevancia de la señal capturada; no
  sustituye al análisis comercial de 13 criterios del Excel maestro.
- **Verifica siempre** fechas e importes en la fuente oficial antes de actuar.

## 🧭 Flujo de trabajo recomendado

1. Te llega el correo del Issue de alerta (o miras el panel el lunes).
2. Abres el enlace oficial de cada detección y verificas datos.
3. Las oportunidades maduras pasan al **Excel maestro** (`Radar Licitaciones
   Agua ES-PT — Canalis`): ficha comercial, contactos y scoring de 13 criterios.
4. El informe semanal (`data/informe.md`) alimenta la reunión comercial.
