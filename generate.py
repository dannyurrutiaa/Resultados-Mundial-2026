"""
Mundial 2026 - Generador de página de apuestas
Uso: python generate.py --api-key TU_API_KEY
     python generate.py --demo   (datos de prueba sin API)
"""
 
import argparse
import sys
from datetime import datetime, date
from pathlib import Path
 
import openpyxl
import requests
 
# ─── Configuración ────────────────────────────────────────────────────────────
 
EXCEL_FILE   = "mundial_2026_apuestas.xlsx"
OUTPUT_FILE  = "index.html"
API_BASE     = "https://api.football-data.org/v4"
WC_2026_ID   = 2000
POINTS       = {1: 4, 2: 3, 3: 2, 4: 1}
PAISES_ES_EN = {
    "Chequia": "Czechia",
    "Ecuador": "Ecuador",
    "Suecia": "Sweden",
    "Uzbekistán": "Uzbekistan",
    "Portugal": "Portugal",
    "Catar": "Qatar",
    "Irán": "Iran",
    "Australia": "Australia",
    "Brasil": "Brazil",
    "Argelia": "Algeria",
    "Suiza": "Switzerland",
    "Colombia": "Colombia",
    "Ghana": "Ghana",
    "Túnez": "Tunisia",
    "Curazao": "Curaçao",
    "Alemania": "Germany",
    "Japón": "Japan",
    "Croacia": "Croatia",
    "Uruguay": "Uruguay",
    "México": "Mexico",
    "Senegal": "Senegal",
    "Turquía": "Turkey",
    "Egipto": "Egypt",
    "Argentina": "Argentina",
    "Marruecos": "Morocco",
    "Jordania": "Jordan",
    "España": "Spain",
    "Austria": "Austria",
    "Sudáfrica": "South Africa",
    "Haití": "Haiti",
    "Costa de Marfil": "Ivory Coast",
    "Escocia": "Scotland",
    "Estados Unidos": "United States",
    "Panamá": "Panama",
    "Bélgica": "Belgium",
    "Nueva Zelanda": "New Zealand",
    "Arabia Saudita": "Saudi Arabia",
    "Corea del Sur": "South Korea",
    "Cabo Verde": "Cape Verde Islands",
    "República Democrática del Congo": "Congo DR",
    "Holanda": "Netherlands",
    "Bosnia y Herzegovina": "Bosnia-Herzegovina",
    "Francia": "France",
    "Canadá": "Canada",
    "Noruega": "Norway",
    "Irak": "Iraq",
    "Paraguay": "Paraguay",
    "Inglaterra": "England"
}
PAISES_EN_ES = {
    "Czechia": "Chequia",
    "Ecuador": "Ecuador",
    "Sweden": "Suecia",
    "Uzbekistan": "Uzbekistán",
    "Portugal": "Portugal",
    "Qatar": "Catar",
    "Iran": "Irán",
    "Australia": "Australia",
    "Brazil": "Brasil",
    "Algeria": "Argelia",
    "Switzerland": "Suiza",
    "Colombia": "Colombia",
    "Ghana": "Ghana",
    "Tunisia": "Túnez",
    "Curaçao": "Curazao",
    "Germany": "Alemania",
    "Japan": "Japón",
    "Croatia": "Croacia",
    "Uruguay": "Uruguay",
    "Mexico": "México",
    "Senegal": "Senegal",
    "Turkey": "Turquía",
    "Egypt": "Egipto",
    "Argentina": "Argentina",
    "Morocco": "Marruecos",
    "Jordan": "Jordania",
    "Spain": "España",
    "Austria": "Austria",
    "South Africa": "Sudáfrica",
    "Haiti": "Haití",
    "Ivory Coast": "Costa de Marfil",
    "Scotland": "Escocia",
    "United States": "Estados Unidos",
    "Panama": "Panamá",
    "Belgium": "Bélgica",
    "New Zealand": "Nueva Zelanda",
    "Saudi Arabia": "Arabia Saudita",
    "South Korea": "Corea del Sur",
    "Cape Verde Islands": "Cabo Verde",
    "Congo DR": "República Democrática del Congo",
    "Netherlands": "Holanda",
    "Bosnia-Herzegovina": "Bosnia y Herzegovina",
    "France": "Francia",
    "Canada": "Canadá",
    "Norway": "Noruega",
    "Iraq": "Irak",
    "Paraguay": "Paraguay",
    "England": "Inglaterra"
}
 
# Fecha de inicio del Mundial 2026 (11 de junio de 2026)
WC_START_DATE = date(2026, 6, 11)
 
# Mapeo de stages de la API a nombres en español
STAGE_LABELS = {
    "GROUP_STAGE":        "Fase de Grupos",
    "LAST_16":            "Octavos de Final",
    "QUARTER_FINALS":     "Cuartos de Final",
    "SEMI_FINALS":        "Semifinales",
    "THIRD_PLACE":        "Disputa del 3° Puesto",
    "FINAL":              "Final",
}
 
# ─── Lectura del Excel ────────────────────────────────────────────────────────
 
def leer_apuestas(path: str) -> list[dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Apuestas"]
    apuestas = []
    for row in ws.iter_rows(min_row=4, max_row=30, values_only=True):
        nombre = row[0]
        if not nombre or str(nombre).strip() == "":
            continue
        pago_raw = str(row[5]).strip() if row[5] else ""
        if row[1] is None or row[2] is None or row[3] is None or row[4] is None:
            opcion_1 = None
            opcion_2 = None
            opcion_3 = None
            opcion_4 = None
        else:
          opcion_1 = PAISES_ES_EN[str(row[1]).strip()]
          opcion_2 = PAISES_ES_EN[str(row[2]).strip()]
          opcion_3 = PAISES_ES_EN[str(row[3]).strip()]
          opcion_4 = PAISES_ES_EN[str(row[4]).strip()]
        apuestas.append({
            "nombre": str(nombre).strip(),
            "picks":  {1: opcion_1, 2: opcion_2,
                       3: opcion_3, 4: opcion_4},
            "pago":   "✓" in pago_raw or pago_raw.lower() == "true",
            "obs":    str(row[6]).strip() if row[6] else "",
        })
    return apuestas
 
# ─── API football-data.org ────────────────────────────────────────────────────
 
def fetch_standings(api_key: str) -> dict:
    headers = {"X-Auth-Token": api_key}
 
    # Info general del torneo
    r = requests.get(f"{API_BASE}/competitions/{WC_2026_ID}", headers=headers, timeout=10)
    r.raise_for_status()
    comp = r.json()
    season = comp.get("currentSeason", {})
    winner = season.get("winner")
 
    # Todos los equipos del torneo
    r2 = requests.get(f"{API_BASE}/competitions/{WC_2026_ID}/teams", headers=headers, timeout=10)
    r2.raise_for_status()
    todos = {t["name"] for t in r2.json().get("teams", [])}
 
    # Si el torneo no ha comenzado: todos vivos, fase = "Próximamente"
    hoy = date.today()
    if hoy < WC_START_DATE or not todos:
        return {
            "vivos":      set(),   # vacío = todos vivos por defecto en calcular()
            "eliminados": set(),
            "resultados": {},
            "fase_actual": "⏳ Próximamente — 11 Jun 2026",
            "torneo_iniciado": False,
        }
 
    # Partidos jugados para detectar eliminados y resultados finales
    r3 = requests.get(f"{API_BASE}/competitions/{WC_2026_ID}/matches",
                      headers=headers, params={"status": "FINISHED"}, timeout=10)
    r3.raise_for_status()
    matches = r3.json().get("matches", [])
 
    resultados_finales = {}
    eliminados = set()
    stage_actual = "GROUP_STAGE"
 
    for m in matches:
        stage_name = m.get("stage", "")
        home  = m["homeTeam"]["name"]
        away  = m["awayTeam"]["name"]
        score = m["score"]["fullTime"]
        if score["home"] is None:
            continue
        winner_m = home if score["home"] > score["away"] else away
        loser_m  = away if score["home"] > score["away"] else home
        stage_actual = stage_name  # el último jugado es la fase más avanzada
 
        if stage_name == "FINAL":
            resultados_finales[1] = winner_m
            resultados_finales[2] = loser_m
        elif stage_name == "THIRD_PLACE":
            resultados_finales[3] = winner_m
            resultados_finales[4] = loser_m
        elif stage_name not in ("GROUP_STAGE",):
            # En fase eliminatoria el perdedor queda fuera
            eliminados.add(loser_m)
 
    # Fase actual legible
    if winner:
        fase_label = "🏆 Torneo Finalizado"
    else:
        fase_label = f"📍 {STAGE_LABELS.get(stage_actual, stage_actual)}"
 
    vivos = todos - eliminados if eliminados else todos
 
    return {
        "vivos":           vivos,
        "eliminados":      eliminados,
        "resultados":      resultados_finales,
        "fase_actual":     fase_label,
        "torneo_iniciado": True,
    }
 
# ─── Datos demo ───────────────────────────────────────────────────────────────
 
def fetch_demo() -> dict:
    vivos = {"Argentina", "France", "Brazil", "Spain", "Germany",
             "Portugal", "Morocco", "Netherlands"}
    eliminados = {"Mexico", "United States", "Japan", "Senegal",
                  "Poland", "Switzerland", "England", "South Korea", "Australia"}
    return {
        "vivos":           vivos,
        "eliminados":      eliminados,
        "resultados":      {},
        "fase_actual":     "📍 Cuartos de Final (demo)",
        "torneo_iniciado": True,
    }
 
# ─── Cálculo de puntajes ──────────────────────────────────────────────────────
 
def calcular(apuestas: list[dict], standings: dict) -> list[dict]:
    resultados       = standings["resultados"]
    vivos            = standings["vivos"]
    torneo_iniciado  = standings["torneo_iniciado"]
 
    calculados = []
    for a in apuestas:
        pts_reales  = 0
        pts_maximos = 0
        detalle     = {}
 
        for pos, equipo in a["picks"].items():
            pts_pos = POINTS[pos]
 
            if resultados:
                # Torneo terminado o resultados conocidos
                ganado = resultados.get(pos) == equipo
                pts_reales  += pts_pos if ganado else 0
                pts_maximos += pts_pos if ganado else 0
                detalle[pos] = {"equipo": equipo,
                                "estado": "acierto" if ganado else "fallo",
                                "pts":    pts_pos if ganado else 0}
            elif not torneo_iniciado or not vivos:
                # Torneo sin iniciar: todo el mundo tiene máximo posible, todos vivos
                pts_maximos += pts_pos
                detalle[pos] = {"equipo": equipo, "estado": "vivo", "pts": pts_pos}
            else:
                # Torneo en curso con info de equipos vivos
                vivo = equipo in vivos
                if vivo:
                    pts_maximos += pts_pos
                    detalle[pos] = {"equipo": equipo, "estado": "vivo", "pts": pts_pos}
                else:
                    detalle[pos] = {"equipo": equipo, "estado": "eliminado", "pts": 0}
 
        calculados.append({**a, "pts_reales": pts_reales,
                           "pts_maximos": pts_maximos, "detalle": detalle})
 
    if resultados:
        calculados.sort(key=lambda x: (-x["pts_reales"], -x["pts_maximos"]))
    else:
        calculados.sort(key=lambda x: (-x["pts_maximos"], -x["pts_reales"]))
 
    return calculados
 
# ─── Generación HTML ──────────────────────────────────────────────────────────
 
ESTADO_ICON  = {"vivo": "🟢", "eliminado": "🔴", "acierto": "✅", "fallo": "❌"}
ESTADO_LABEL = {"vivo": "Vivo", "eliminado": "Eliminado", "acierto": "¡Acertó!", "fallo": "Falló"}
POS_LABEL    = {1: "🥇 1°", 2: "🥈 2°", 3: "🥉 3°", 4: "4°"}
 
def generar_html(apuestas_calc: list[dict], standings: dict, generado: str) -> str:
    total_participantes = sum(1 for a in apuestas_calc if a["pago"])
    pozo      = total_participantes * 100_000
    premio_1  = max(pozo - 100_000, 0)
    fase      = standings["fase_actual"]
    torneo_terminado = bool(standings["resultados"])
 
    cards_html = ""
    rank_real = 0
    for a in apuestas_calc:
        if not a["pago"]:
            continue
        rank_real += 1
        rank = rank_real
 
        if rank == 1:   rank_badge = '<span class="rank gold">👑 1°</span>'
        elif rank == 2: rank_badge = '<span class="rank silver">🥈 2°</span>'
        elif rank == 3: rank_badge = '<span class="rank bronze">🥉 3°</span>'
        else:           rank_badge = f'<span class="rank">{rank}°</span>'
 
        picks_html = ""
        for pos in [1, 2, 3, 4]:
            d = a["detalle"].get(pos, {})
            estado = d.get("estado", "vivo")
            picks_html += f"""
            <div class="pick {estado}">
              <span class="pick-pos">{POS_LABEL[pos]}</span>
              <span class="pick-team">{PAISES_EN_ES[d.get('equipo', '—')]}</span>
              <span class="pick-badge">{ESTADO_ICON[estado]} {ESTADO_LABEL[estado]}</span>
              <span class="pick-pts">+{d.get('pts', POINTS[pos])} pts</span>
            </div>"""
 
        score_val   = a["pts_reales"] if torneo_terminado else a["pts_maximos"]
        score_label = "Puntos totales" if torneo_terminado else "Máx. alcanzable"
 
        cards_html += f"""
        <div class="card rank-{min(rank,4)}">
          <div class="card-header">
            {rank_badge}
            <h3>{a['nombre']}</h3>
            <div class="score-box">
              <div class="score-main">{score_val}<span>pts</span></div>
              <div class="score-label">{score_label}</div>
            </div>
          </div>
          <div class="picks-grid">{picks_html}</div>
        </div>"""
 
    resultados_html = ""
    if standings["resultados"]:
        resultados_html = '<div class="resultados-finales"><h2>🏆 Resultados Oficiales</h2><div class="res-grid">'
        for pos, equipo in standings["resultados"].items():
            resultados_html += f'<div class="res-item pos-{pos}"><span class="res-pos">{POS_LABEL[pos]}</span><span class="res-team">{equipo}</span></div>'
        resultados_html += "</div></div>"
 
    # Formato pesos chilenos
    def clp(n):
        return f"${n:,.0f}".replace(",", ".")
 
    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Mundial 2026 — Apuestas</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:ital,wght@0,300;0,400;0,600;1,400&display=swap" rel="stylesheet">
<style>
  :root {{
    /* Paleta oficial FIFA 2026: negro, blanco y dorado como base */
    --bg:        #f7f5f0;
    --surface:   #ffffff;
    --surface2:  #f0ede6;
    --border:    #e2ddd4;
    --gold:      #c9a84c;
    --gold-dark: #a8832a;
    --gold-light:#f0d98a;
    --silver:    #8a9ba8;
    --bronze:    #a0674a;
    --green:     #2d7a4f;
    --green-bg:  #eaf4ee;
    --red:       #c0392b;
    --red-bg:    #fdf0ee;
    --text:      #1a1a1a;
    --text2:     #4a4540;
    --muted:     #9a9590;
    --radius:    14px;
    /* Colores países sede */
    --usa-blue:  #1a3a6e;
    --mex-green: #1a6e3a;
    --can-red:   #c0392b;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
 
  body {{
    font-family: 'DM Sans', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
  }}
 
  /* ── Franja decorativa superior tricolor ── */
  .top-stripe {{
    height: 5px;
    background: linear-gradient(90deg,
      var(--can-red) 0%, var(--can-red) 33.3%,
      var(--gold)    33.3%, var(--gold) 66.6%,
      var(--usa-blue) 66.6%, var(--usa-blue) 100%);
  }}
 
  /* ── Header ── */
  header {{
    background: #1a1a1a;
    text-align: center;
    padding: 48px 24px 36px;
    position: relative;
    overflow: hidden;
  }}
  header::before {{
    content: '26 26 26 26 26 26 26 26 26 26 26 26';
    position: absolute;
    top: 50%; left: 50%;
    transform: translate(-50%, -50%);
    font-family: 'Bebas Neue', sans-serif;
    font-size: 120px;
    color: rgba(255,255,255,0.03);
    white-space: nowrap;
    letter-spacing: 20px;
    pointer-events: none;
  }}
  .title {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: clamp(48px, 9vw, 88px);
    letter-spacing: 6px;
    line-height: 1;
    color: #ffffff;
    position: relative;
  }}
  .title span {{
    color: var(--gold);
  }}
  .subtitle {{
    margin-top: 8px;
    color: #888;
    font-size: 13px;
    letter-spacing: 4px;
    text-transform: uppercase;
    position: relative;
  }}
  .fase-badge {{
    display: inline-flex;
    align-items: center;
    gap: 6px;
    margin-top: 18px;
    padding: 7px 20px;
    border-radius: 999px;
    background: rgba(201,168,76,0.12);
    border: 1px solid rgba(201,168,76,0.35);
    color: var(--gold-light);
    font-size: 13px;
    font-weight: 500;
    letter-spacing: 0.5px;
    position: relative;
  }}
 
  /* ── Stats bar ── */
  .stats {{
    display: flex;
    justify-content: center;
    gap: 1px;
    background: var(--border);
    border-bottom: 1px solid var(--border);
    max-width: 100%;
    overflow-x: auto;
  }}
  .stat {{
    flex: 1;
    min-width: 140px;
    background: var(--surface);
    padding: 20px 24px;
    text-align: center;
  }}
  .stat-val {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 32px;
    letter-spacing: 1px;
    color: var(--gold-dark);
    line-height: 1;
  }}
  .stat-label {{
    font-size: 11px;
    color: var(--muted);
    text-transform: uppercase;
    letter-spacing: 1.5px;
    margin-top: 4px;
  }}
 
  /* ── Resultados Oficiales ── */
  .resultados-finales {{
    max-width: 700px;
    margin: 40px auto 0;
    padding: 0 24px;
  }}
  .resultados-finales h2 {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 26px;
    letter-spacing: 2px;
    margin-bottom: 14px;
    text-align: center;
    color: var(--gold-dark);
  }}
  .res-grid {{
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 10px;
  }}
  .res-item {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 14px 18px;
    display: flex;
    justify-content: space-between;
    align-items: center;
  }}
  .res-item.pos-1 {{ border-left: 3px solid var(--gold); }}
  .res-item.pos-2 {{ border-left: 3px solid var(--silver); }}
  .res-item.pos-3 {{ border-left: 3px solid var(--bronze); }}
  .res-pos  {{ font-size: 16px; }}
  .res-team {{ font-weight: 600; color: var(--text); }}
 
  /* ── Cards ── */
  .cards-section {{
    max-width: 1100px;
    margin: 0 auto;
    padding: 40px 24px 60px;
  }}
  .section-title {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 22px;
    letter-spacing: 3px;
    color: var(--muted);
    text-transform: uppercase;
    margin-bottom: 20px;
    padding-bottom: 10px;
    border-bottom: 1px solid var(--border);
  }}
  .cards {{
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
    gap: 16px;
  }}
  .card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    overflow: hidden;
    transition: transform .2s, box-shadow .2s;
  }}
  .card:hover {{
    transform: translateY(-3px);
    box-shadow: 0 12px 32px rgba(0,0,0,0.1);
  }}
  .card.rank-1 {{
    border-color: var(--gold);
    box-shadow: 0 0 0 1px rgba(201,168,76,0.3);
  }}
  .card.rank-2 {{ border-color: var(--silver); }}
  .card.rank-3 {{ border-color: var(--bronze); }}
 
  .card-header {{
    padding: 16px 18px 14px;
    display: flex;
    align-items: center;
    gap: 10px;
    border-bottom: 1px solid var(--border);
    background: var(--surface2);
  }}
  .card-header h3 {{
    flex: 1;
    font-size: 17px;
    font-weight: 600;
    color: var(--text);
  }}
  .rank {{
    font-size: 12px;
    font-weight: 600;
    padding: 3px 10px;
    border-radius: 999px;
    background: var(--surface);
    border: 1px solid var(--border);
    color: var(--muted);
    white-space: nowrap;
  }}
  .rank.gold   {{ background: #fdf6e3; border-color: var(--gold); color: var(--gold-dark); }}
  .rank.silver {{ background: #f4f6f8; border-color: var(--silver); color: #607080; }}
  .rank.bronze {{ background: #fdf0e8; border-color: var(--bronze); color: var(--bronze); }}
 
  .score-box {{ text-align: right; }}
  .score-main {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 30px;
    line-height: 1;
    color: var(--gold-dark);
  }}
  .score-main span {{
    font-size: 13px;
    color: var(--muted);
    font-family: 'DM Sans', sans-serif;
    font-weight: 400;
    margin-left: 2px;
  }}
  .score-label {{
    font-size: 10px;
    color: var(--muted);
    text-transform: uppercase;
    letter-spacing: 0.8px;
  }}
 
  /* ── Picks ── */
  .picks-grid {{
    padding: 14px;
    display: flex;
    flex-direction: column;
    gap: 7px;
  }}
  .pick {{
    display: grid;
    grid-template-columns: 46px 1fr auto auto;
    align-items: center;
    gap: 8px;
    padding: 9px 12px;
    border-radius: 8px;
    border: 1px solid var(--border);
    background: var(--bg);
  }}
  .pick.vivo      {{ border-color: #b7dfc8; background: var(--green-bg); }}
  .pick.acierto   {{ border-color: #2d7a4f; background: var(--green-bg); }}
  .pick.eliminado {{ border-color: #f0c4be; background: var(--red-bg); opacity: 0.7; }}
  .pick.fallo     {{ border-color: #f0c4be; background: var(--red-bg); opacity: 0.6; }}
 
  .pick-pos   {{ font-size: 12px; color: var(--muted); }}
  .pick-team  {{ font-weight: 500; font-size: 14px; color: var(--text); }}
  .pick-badge {{ font-size: 11px; color: var(--muted); white-space: nowrap; }}
  .pick-pts   {{ font-size: 12px; font-weight: 600; color: var(--muted); white-space: nowrap; }}
  .pick.vivo .pick-pts    {{ color: var(--green); }}
  .pick.acierto .pick-pts {{ color: var(--green); }}
  .pick.eliminado .pick-pts {{ color: var(--red); }}
 
  /* ── Footer ── */
  footer {{
    text-align: center;
    padding: 20px 24px;
    color: var(--muted);
    font-size: 12px;
    border-top: 1px solid var(--border);
    background: var(--surface);
  }}
 
  @media (max-width: 600px) {{
    .cards       {{ grid-template-columns: 1fr; }}
    .res-grid    {{ grid-template-columns: 1fr; }}
    .stat        {{ min-width: 120px; padding: 16px; }}
    .stat-val    {{ font-size: 26px; }}
  }}
</style>
</head>
<body>
 
<div class="top-stripe"></div>
 
<header>
  <div class="title">MUNDIAL <span>2026</span></div>
  <div class="subtitle">Control de Apuestas</div>
  <div class="fase-badge">{fase}</div>
</header>
 
<div class="stats">
  <div class="stat">
    <div class="stat-val">{total_participantes}</div>
    <div class="stat-label">Participantes</div>
  </div>
  <div class="stat">
    <div class="stat-val">{clp(pozo)}</div>
    <div class="stat-label">Pozo total</div>
  </div>
  <div class="stat">
    <div class="stat-val">{clp(premio_1)}</div>
    <div class="stat-label">Premio 1° lugar</div>
  </div>
  <div class="stat">
    <div class="stat-val">$100.000</div>
    <div class="stat-label">Premio 2° lugar</div>
  </div>
</div>
 
{resultados_html}
 
<div class="cards-section">
  <div class="section-title">Clasificación</div>
  <div class="cards">
{cards_html}
  </div>
</div>
 
<footer>
  Última actualización: {generado} &nbsp;·&nbsp; football-data.org
</footer>
 
</body>
</html>"""
 
# ─── Main ─────────────────────────────────────────────────────────────────────
 
def main():
    parser = argparse.ArgumentParser(description="Generador Mundial 2026")
    parser.add_argument("--api-key", help="API key de football-data.org")
    parser.add_argument("--demo",    action="store_true", help="Datos de prueba sin API")
    parser.add_argument("--excel",   default=EXCEL_FILE)
    parser.add_argument("--output",  default=OUTPUT_FILE)
    args = parser.parse_args()
 
    if not args.demo and not args.api_key:
        print("❌  Debes pasar --api-key o usar --demo para pruebas.")
        sys.exit(1)
 
    print(f"📂  Leyendo apuestas desde {args.excel}...")
    apuestas = leer_apuestas(args.excel)
    print(f"    {len(apuestas)} participantes encontrados.")
 
    if args.demo:
        print("🎮  Modo demo activado.")
        standings = fetch_demo()
    else:
        print("🌐  Consultando API football-data.org...")
        try:
            standings = fetch_standings(args.api_key)
        except requests.HTTPError as e:
            print(f"❌  Error de API: {e}")
            sys.exit(1)
 
    print(f"    Fase actual: {standings['fase_actual']}")
 
    apuestas_calc = calcular(apuestas, standings)
    generado = datetime.now().strftime("%d/%m/%Y %H:%M")
    html = generar_html(apuestas_calc, standings, generado)
 
    Path(args.output).write_text(html, encoding="utf-8")
    print(f"✅  HTML generado: {args.output}")
 
if __name__ == "__main__":
    main()