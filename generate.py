"""
Mundial 2026 - Generador de página de apuestas
Uso: python generate.py --api-key TU_API_KEY
     python generate.py --api-key TU_API_KEY --demo   (datos de prueba sin API)
"""

import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
import requests

# ─── Configuración ────────────────────────────────────────────────────────────

EXCEL_FILE = "apuestas.xlsx"
OUTPUT_FILE = "index.html"
API_BASE   = "https://api.football-data.org/v4"
WC_2026_ID = 2000          # ID del Mundial 2026 en football-data.org

# Puntos por posición exacta
POINTS = {1: 4, 2: 3, 3: 2, 4: 1}

# ─── Lectura del Excel ────────────────────────────────────────────────────────

def leer_apuestas(path: str) -> list[dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Apuestas"]
    apuestas = []
    for row in ws.iter_rows(min_row=4, max_row=13, values_only=True):
        nombre = row[0]
        if not nombre or str(nombre).strip() == "":
            continue
        pago_raw = str(row[5]).strip() if row[5] else ""
        apuestas.append({
            "nombre":   str(nombre).strip(),
            "picks":    {1: str(row[1]).strip(), 2: str(row[2]).strip(),
                         3: str(row[3]).strip(), 4: str(row[4]).strip()},
            "pago":     "✓" in pago_raw or pago_raw.lower() == "true",
            "obs":      str(row[6]).strip() if row[6] else "",
        })
    return apuestas

# ─── API football-data.org ────────────────────────────────────────────────────

def fetch_standings(api_key: str) -> dict:
    """Retorna equipos vivos y resultados finales si el torneo terminó."""
    headers = {"X-Auth-Token": api_key}

    # Estado del torneo
    r = requests.get(f"{API_BASE}/competitions/{WC_2026_ID}", headers=headers, timeout=10)
    r.raise_for_status()
    comp = r.json()
    stage = comp.get("currentSeason", {}).get("currentMatchday", {})

    # Equipos que siguen en competencia
    r2 = requests.get(f"{API_BASE}/competitions/{WC_2026_ID}/teams", headers=headers, timeout=10)
    r2.raise_for_status()
    todos = {t["name"] for t in r2.json().get("teams", [])}

    # Intentar obtener resultados finales (fase eliminatoria)
    r3 = requests.get(f"{API_BASE}/competitions/{WC_2026_ID}/matches",
                      headers=headers, params={"stage": "FINAL"}, timeout=10)
    r3.raise_for_status()
    matches = r3.json().get("matches", [])

    resultados_finales = {}   # {1: "Argentina", 2: "Francia", 3: "Brasil", 4: "España"}
    eliminados = set()        # equipos ya eliminados

    for m in matches:
        status = m.get("status")
        if status != "FINISHED":
            continue
        stage_name = m.get("stage", "")
        home = m["homeTeam"]["name"]
        away = m["awayTeam"]["name"]
        score = m["score"]["fullTime"]
        winner = home if score["home"] > score["away"] else away
        loser  = away if score["home"] > score["away"] else home

        if stage_name == "FINAL":
            resultados_finales[1] = winner
            resultados_finales[2] = loser
        elif stage_name == "THIRD_PLACE":
            resultados_finales[3] = winner
            resultados_finales[4] = loser
        else:
            eliminados.add(loser)

    # Equipos vivos = todos los del torneo menos los eliminados (sin resultados finales aún)
    vivos = todos - eliminados if not resultados_finales else set(resultados_finales.values())

    return {
        "vivos": vivos,
        "eliminados": eliminados,
        "resultados": resultados_finales,
        "fase_actual": comp.get("currentSeason", {}).get("winner", {}) and "Finalizado" or "En curso",
    }

# ─── Datos demo (sin API) ─────────────────────────────────────────────────────

def fetch_demo() -> dict:
    """Simula que estamos en cuartos de final con algunos equipos eliminados."""
    vivos = {"Argentina", "Francia", "Brasil", "España", "Alemania",
             "Portugal", "Inglaterra", "Marruecos"}
    eliminados = {"México", "Estados Unidos", "Japón", "Senegal",
                  "Polonia", "Suiza", "Corea del Sur", "Australia"}
    return {
        "vivos": vivos,
        "eliminados": eliminados,
        "resultados": {},   # torneo sin terminar
        "fase_actual": "Cuartos de Final (demo)",
    }

# ─── Cálculo de puntajes ──────────────────────────────────────────────────────

def calcular(apuestas: list[dict], standings: dict) -> list[dict]:
    resultados = standings["resultados"]
    vivos      = standings["vivos"]

    calculados = []
    for a in apuestas:
        pts_reales  = 0
        pts_maximos = 0
        detalle     = {}

        for pos, equipo in a["picks"].items():
            pts_pos = POINTS[pos]

            if resultados:
                # Torneo terminado: puntos exactos
                ganado = resultados.get(pos) == equipo
                pts_reales  += pts_pos if ganado else 0
                pts_maximos += pts_pos if ganado else 0
                detalle[pos] = {"equipo": equipo, "estado": "acierto" if ganado else "fallo",
                                "pts": pts_pos if ganado else 0}
            else:
                # Torneo en curso
                vivo = equipo in vivos or not vivos  # si vivos vacío, no sabemos
                if vivo:
                    pts_maximos += pts_pos
                    detalle[pos] = {"equipo": equipo, "estado": "vivo", "pts": pts_pos}
                else:
                    detalle[pos] = {"equipo": equipo, "estado": "eliminado", "pts": 0}

        calculados.append({**a, "pts_reales": pts_reales,
                           "pts_maximos": pts_maximos, "detalle": detalle})

    # Ordenar: torneo terminado → por pts_reales; en curso → por pts_maximos desc
    if resultados:
        calculados.sort(key=lambda x: (-x["pts_reales"], -x["pts_maximos"]))
    else:
        calculados.sort(key=lambda x: (-x["pts_maximos"], -x["pts_reales"]))

    return calculados

# ─── Generación HTML ──────────────────────────────────────────────────────────

ESTADO_ICON = {"vivo": "🟢", "eliminado": "🔴", "acierto": "✅", "fallo": "❌"}
ESTADO_LABEL = {"vivo": "Vivo", "eliminado": "Eliminado", "acierto": "¡Acertó!", "fallo": "Falló"}
POS_LABEL = {1: "🥇 1°", 2: "🥈 2°", 3: "🥉 3°", 4: "4°"}

def generar_html(apuestas_calc: list[dict], standings: dict, generado: str) -> str:
    total_participantes = sum(1 for a in apuestas_calc if a["pago"])
    pozo = total_participantes * 100_000
    premio_1 = pozo - 100_000
    fase = standings["fase_actual"]
    torneo_terminado = bool(standings["resultados"])

    # Tarjetas de participantes
    cards_html = ""
    for i, a in enumerate(apuestas_calc):
        if not a["pago"]:
            continue
        rank = i + 1
        rank_badge = ""
        if rank == 1:   rank_badge = '<span class="rank gold">👑 1°</span>'
        elif rank == 2: rank_badge = '<span class="rank silver">🥈 2°</span>'
        elif rank == 3: rank_badge = '<span class="rank bronze">🥉 3°</span>'
        else:            rank_badge = f'<span class="rank">{rank}°</span>'

        picks_html = ""
        for pos in [1, 2, 3, 4]:
            d = a["detalle"].get(pos, {})
            estado = d.get("estado", "vivo")
            css_estado = estado
            picks_html += f"""
            <div class="pick {css_estado}">
              <span class="pick-pos">{POS_LABEL[pos]}</span>
              <span class="pick-team">{d.get('equipo','—')}</span>
              <span class="pick-badge">{ESTADO_ICON[estado]} {ESTADO_LABEL[estado]}</span>
              <span class="pick-pts">+{d.get('pts', POINTS[pos])} pts</span>
            </div>"""

        puntaje_label = "Puntos actuales" if torneo_terminado else "Máx. alcanzable"
        cards_html += f"""
        <div class="card rank-{min(rank,4)}">
          <div class="card-header">
            {rank_badge}
            <h3>{a['nombre']}</h3>
            <div class="score-box">
              <div class="score-main">{a['pts_maximos']}<span>pts</span></div>
              <div class="score-label">{puntaje_label}</div>
            </div>
          </div>
          <div class="picks-grid">{picks_html}</div>
        </div>"""

    # Resultados oficiales
    resultados_html = ""
    if standings["resultados"]:
        resultados_html = '<div class="resultados-finales"><h2>🏆 Resultados Oficiales</h2><div class="res-grid">'
        for pos, equipo in standings["resultados"].items():
            resultados_html += f'<div class="res-item pos-{pos}"><span class="res-pos">{POS_LABEL[pos]}</span><span class="res-team">{equipo}</span></div>'
        resultados_html += "</div></div>"

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Mundial 2026 — Apuestas</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
  :root {{
    --bg:        #0a0a0f;
    --surface:   #13131c;
    --surface2:  #1c1c2a;
    --border:    rgba(255,255,255,0.07);
    --gold:      #f5c842;
    --silver:    #b0bec5;
    --bronze:    #cd7f32;
    --green:     #00e676;
    --red:       #ff5252;
    --accent:    #5c6bc0;
    --text:      #eef0f8;
    --muted:     #6b6f8a;
    --radius:    16px;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: 'DM Sans', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    background-image: radial-gradient(ellipse at 20% 0%, #1a1a3a 0%, transparent 60%),
                      radial-gradient(ellipse at 80% 100%, #0d2a1a 0%, transparent 60%);
  }}

  /* ── Header ── */
  header {{
    text-align: center;
    padding: 56px 24px 32px;
    position: relative;
  }}
  header::after {{
    content: '';
    position: absolute;
    bottom: 0; left: 50%;
    transform: translateX(-50%);
    width: 80px; height: 2px;
    background: linear-gradient(90deg, transparent, var(--gold), transparent);
  }}
  .title {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: clamp(42px, 8vw, 80px);
    letter-spacing: 4px;
    line-height: 1;
    background: linear-gradient(135deg, #fff 0%, var(--gold) 60%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
  }}
  .subtitle {{
    margin-top: 8px;
    color: var(--muted);
    font-size: 14px;
    letter-spacing: 3px;
    text-transform: uppercase;
  }}
  .fase-badge {{
    display: inline-block;
    margin-top: 16px;
    padding: 6px 18px;
    border-radius: 999px;
    background: rgba(92,107,192,0.15);
    border: 1px solid rgba(92,107,192,0.35);
    color: #9fa8da;
    font-size: 13px;
    font-weight: 500;
    letter-spacing: 1px;
  }}

  /* ── Stats bar ── */
  .stats {{
    display: flex;
    justify-content: center;
    gap: 16px;
    flex-wrap: wrap;
    padding: 32px 24px;
    max-width: 900px;
    margin: 0 auto;
  }}
  .stat {{
    flex: 1;
    min-width: 160px;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 20px 24px;
    text-align: center;
  }}
  .stat-val {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 36px;
    letter-spacing: 2px;
    color: var(--gold);
  }}
  .stat-label {{
    font-size: 12px;
    color: var(--muted);
    text-transform: uppercase;
    letter-spacing: 1.5px;
    margin-top: 4px;
  }}

  /* ── Resultados Oficiales ── */
  .resultados-finales {{
    max-width: 700px;
    margin: 0 auto 32px;
    padding: 0 24px;
  }}
  .resultados-finales h2 {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 28px;
    letter-spacing: 2px;
    margin-bottom: 16px;
    text-align: center;
    color: var(--gold);
  }}
  .res-grid {{
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 12px;
  }}
  .res-item {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 16px 20px;
    display: flex;
    justify-content: space-between;
    align-items: center;
  }}
  .res-item.pos-1 {{ border-color: rgba(245,200,66,0.4); }}
  .res-item.pos-2 {{ border-color: rgba(176,190,197,0.4); }}
  .res-pos {{ font-size: 18px; }}
  .res-team {{ font-weight: 600; color: var(--text); }}

  /* ── Cards ── */
  .cards {{
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(340px, 1fr));
    gap: 20px;
    max-width: 1100px;
    margin: 0 auto;
    padding: 0 24px 60px;
  }}
  .card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    overflow: hidden;
    transition: transform .2s, box-shadow .2s;
  }}
  .card:hover {{
    transform: translateY(-4px);
    box-shadow: 0 20px 40px rgba(0,0,0,0.4);
  }}
  .card.rank-1 {{ border-color: rgba(245,200,66,0.35); box-shadow: 0 0 30px rgba(245,200,66,0.08); }}
  .card.rank-2 {{ border-color: rgba(176,190,197,0.25); }}
  .card.rank-3 {{ border-color: rgba(205,127,50,0.25); }}

  .card-header {{
    padding: 20px 20px 16px;
    display: flex;
    align-items: center;
    gap: 12px;
    border-bottom: 1px solid var(--border);
    background: var(--surface2);
  }}
  .card-header h3 {{
    flex: 1;
    font-size: 18px;
    font-weight: 600;
  }}
  .rank {{
    font-size: 13px;
    font-weight: 600;
    padding: 4px 10px;
    border-radius: 999px;
    background: rgba(255,255,255,0.07);
    color: var(--muted);
    white-space: nowrap;
  }}
  .rank.gold   {{ background: rgba(245,200,66,0.15); color: var(--gold); }}
  .rank.silver {{ background: rgba(176,190,197,0.15); color: var(--silver); }}
  .rank.bronze {{ background: rgba(205,127,50,0.15); color: var(--bronze); }}

  .score-box {{ text-align: right; }}
  .score-main {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 32px;
    line-height: 1;
    color: var(--gold);
  }}
  .score-main span {{ font-size: 14px; margin-left: 2px; color: var(--muted); font-family: 'DM Sans', sans-serif; }}
  .score-label {{ font-size: 10px; color: var(--muted); text-transform: uppercase; letter-spacing: 1px; }}

  /* ── Picks ── */
  .picks-grid {{
    padding: 16px;
    display: flex;
    flex-direction: column;
    gap: 8px;
  }}
  .pick {{
    display: grid;
    grid-template-columns: 48px 1fr auto auto;
    align-items: center;
    gap: 8px;
    padding: 10px 12px;
    border-radius: 10px;
    background: rgba(255,255,255,0.03);
    border: 1px solid var(--border);
    transition: background .15s;
  }}
  .pick.vivo     {{ border-color: rgba(0,230,118,0.2); }}
  .pick.acierto  {{ border-color: rgba(0,230,118,0.35); background: rgba(0,230,118,0.05); }}
  .pick.eliminado {{ border-color: rgba(255,82,82,0.2); opacity: 0.6; }}
  .pick.fallo    {{ border-color: rgba(255,82,82,0.2); opacity: 0.5; }}

  .pick-pos  {{ font-size: 13px; color: var(--muted); }}
  .pick-team {{ font-weight: 500; font-size: 14px; }}
  .pick-badge {{ font-size: 11px; color: var(--muted); white-space: nowrap; }}
  .pick-pts  {{ font-size: 12px; font-weight: 600; color: var(--muted); white-space: nowrap; }}
  .pick.vivo .pick-pts    {{ color: var(--green); }}
  .pick.acierto .pick-pts {{ color: var(--green); }}

  /* ── Footer ── */
  footer {{
    text-align: center;
    padding: 24px;
    color: var(--muted);
    font-size: 12px;
    border-top: 1px solid var(--border);
  }}

  @media (max-width: 600px) {{
    .cards {{ grid-template-columns: 1fr; }}
    .res-grid {{ grid-template-columns: 1fr; }}
  }}
</style>
</head>
<body>

<header>
  <div class="title">MUNDIAL 2026</div>
  <div class="subtitle">Control de Apuestas</div>
  <div class="fase-badge">📍 {fase}</div>
</header>

<div class="stats">
  <div class="stat">
    <div class="stat-val">{total_participantes}</div>
    <div class="stat-label">Participantes</div>
  </div>
  <div class="stat">
    <div class="stat-val">${pozo:,.0f}</div>
    <div class="stat-label">Pozo total</div>
  </div>
  <div class="stat">
    <div class="stat-val">${premio_1:,.0f}</div>
    <div class="stat-label">Premio 1° lugar</div>
  </div>
  <div class="stat">
    <div class="stat-val">$100.000</div>
    <div class="stat-label">Premio 2° lugar</div>
  </div>
</div>

{resultados_html}

<div class="cards">
{cards_html}
</div>

<footer>
  Última actualización: {generado} · Generado con Python + football-data.org
</footer>

</body>
</html>"""

# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Generador Mundial 2026")
    parser.add_argument("--api-key", help="API key de football-data.org")
    parser.add_argument("--demo", action="store_true", help="Usar datos de prueba sin API")
    parser.add_argument("--excel", default=EXCEL_FILE, help="Ruta al Excel de apuestas")
    parser.add_argument("--output", default=OUTPUT_FILE, help="Archivo HTML de salida")
    args = parser.parse_args()

    if not args.demo and not args.api_key:
        print("❌  Debes pasar --api-key o usar --demo para pruebas.")
        sys.exit(1)

    print(f"📂  Leyendo apuestas desde {args.excel}...")
    apuestas = leer_apuestas(args.excel)
    print(f"    {len(apuestas)} participantes encontrados.")

    if args.demo:
        print("🎮  Modo demo activado (datos simulados).")
        standings = fetch_demo()
    else:
        print("🌐  Consultando API football-data.org...")
        try:
            standings = fetch_standings(args.api_key)
        except requests.HTTPError as e:
            print(f"❌  Error de API: {e}")
            sys.exit(1)

    print(f"    Fase actual: {standings['fase_actual']}")
    print(f"    Equipos vivos: {len(standings['vivos'])}")

    apuestas_calc = calcular(apuestas, standings)
    generado = datetime.now().strftime("%d/%m/%Y %H:%M")
    html = generar_html(apuestas_calc, standings, generado)

    Path(args.output).write_text(html, encoding="utf-8")
    print(f"✅  HTML generado: {args.output}")

if __name__ == "__main__":
    main()
