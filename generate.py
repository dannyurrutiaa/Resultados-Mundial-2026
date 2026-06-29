"""
Mundial 2026 - Generador de página de apuestas
Uso: python generate.py --api-key TU_API_KEY
     python generate.py --demo        (torneo en curso, sin API)
     python generate.py --demo-final  (torneo terminado, sin API)
"""

import argparse
import sys
from datetime import datetime, date
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl
import requests

# ─── Configuración ────────────────────────────────────────────────────────────

EXCEL_FILE    = "mundial_2026_apuestas.xlsx"
OUTPUT_FILE   = "index.html"
API_BASE      = "https://api.football-data.org/v4"
WC_2026_ID    = 2000
POINTS        = {1: 4, 2: 3, 3: 2, 4: 1}
WC_START_DATE = date(2026, 6, 11)

STAGE_LABELS = {
    "GROUP_STAGE":    "Fase de Grupos",
    "LAST_32":        "Dieciseisavos de Final",
    "LAST_16":        "Octavos de Final",
    "QUARTER_FINALS": "Cuartos de Final",
    "SEMI_FINALS":    "Semifinales",
    "THIRD_PLACE":    "Disputa del 3° Puesto",
    "FINAL":          "Final",
}

# Menor número = llegó más lejos en el torneo
STAGE_ORDER = {
    "FINAL": 0, "THIRD_PLACE": 2, "SEMI_FINALS": 4,
    "QUARTER_FINALS": 8, "LAST_16": 16, "LAST_32": 24,
    "GROUP_STAGE": 32,
}

# ─── Traducciones Español ↔ Inglés ────────────────────────────────────────────

PAISES_ES_EN = {
    "Chequia": "Czechia", "Ecuador": "Ecuador", "Suecia": "Sweden",
    "Uzbekistán": "Uzbekistan", "Portugal": "Portugal", "Catar": "Qatar",
    "Irán": "Iran", "Australia": "Australia", "Brasil": "Brazil",
    "Argelia": "Algeria", "Suiza": "Switzerland", "Colombia": "Colombia",
    "Ghana": "Ghana", "Túnez": "Tunisia", "Curazao": "Curaçao",
    "Alemania": "Germany", "Japón": "Japan", "Croacia": "Croatia",
    "Uruguay": "Uruguay", "México": "Mexico", "Senegal": "Senegal",
    "Turquía": "Turkey", "Egipto": "Egypt", "Argentina": "Argentina",
    "Marruecos": "Morocco", "Jordania": "Jordan", "España": "Spain",
    "Austria": "Austria", "Sudáfrica": "South Africa", "Haití": "Haiti",
    "Costa de Marfil": "Ivory Coast", "Escocia": "Scotland",
    "Estados Unidos": "United States", "Panamá": "Panama",
    "Bélgica": "Belgium", "Nueva Zelanda": "New Zealand",
    "Arabia Saudita": "Saudi Arabia", "Corea del Sur": "South Korea",
    "Cabo Verde": "Cape Verde Islands",
    "República Democrática del Congo": "Congo DR",
    "Holanda": "Netherlands", "Bosnia y Herzegovina": "Bosnia-Herzegovina",
    "Francia": "France", "Canadá": "Canada", "Noruega": "Norway",
    "Irak": "Iraq", "Paraguay": "Paraguay", "Inglaterra": "England",
}

PAISES_EN_ES = {v: k for k, v in PAISES_ES_EN.items()}

def a_es(nombre_en: str) -> str:
    """Inglés → Español para mostrar en HTML. Si no está en el dict, devuelve el original."""
    return PAISES_EN_ES.get(nombre_en, nombre_en)

def a_en(nombre_es: str) -> str:
    """Español → Inglés para comparar con la API."""
    return PAISES_ES_EN.get(nombre_es, nombre_es)

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

        if any(row[i] is None for i in [1, 2, 3, 4]):
            picks_en = {1: None, 2: None, 3: None, 4: None}
        else:
            picks_en = {
                1: a_en(str(row[1]).strip()),
                2: a_en(str(row[2]).strip()),
                3: a_en(str(row[3]).strip()),
                4: a_en(str(row[4]).strip()),
            }

        apuestas.append({
            "nombre": str(nombre).strip(),
            "picks":  picks_en,       # internamente en inglés (para comparar con API)
            "pago":   "✓" in pago_raw or pago_raw.lower() == "true",
            "obs":    str(row[6]).strip() if row[6] else "",
        })
    return apuestas

# ─── API football-data.org ────────────────────────────────────────────────────

def fetch_standings(api_key: str) -> dict:
    headers = {"X-Auth-Token": api_key}

    r = requests.get(f"{API_BASE}/competitions/{WC_2026_ID}", headers=headers, timeout=10)
    r.raise_for_status()
    comp   = r.json()
    season = comp.get("currentSeason", {})
    winner = season.get("winner")

    r2 = requests.get(f"{API_BASE}/competitions/{WC_2026_ID}/teams", headers=headers, timeout=10)
    r2.raise_for_status()
    todos = {t["name"] for t in r2.json().get("teams", [])}

    hoy = date.today()
    if hoy < WC_START_DATE or not todos:
        return {
            "vivos": set(), "eliminados": set(), "resultados": {},
            "fase_actual": "⏳ Próximamente — 11 Jun 2026",
            "torneo_iniciado": False,
            "posicion_final": {}, "mejor_stage": {},
        }

    r3 = requests.get(f"{API_BASE}/competitions/{WC_2026_ID}/matches",
                      headers=headers, params={"status": "FINISHED"}, timeout=10)
    r3.raise_for_status()
    matches = r3.json().get("matches", [])

    resultados_finales = {}
    eliminados         = set()
    stage_actual       = "GROUP_STAGE"
    mejor_stage        = {}

    for m in matches:
        stage_name = m.get("stage", "")
        home  = m["homeTeam"]["name"]
        away  = m["awayTeam"]["name"]
        score = m["score"]["fullTime"]
        if score["home"] is None:
            continue

        winner_m = home if score["home"] > score["away"] else away
        loser_m  = away if score["home"] > score["away"] else home
        stage_actual = stage_name

        for equipo in [home, away]:
            prev = mejor_stage.get(equipo, "GROUP_STAGE")
            if STAGE_ORDER.get(stage_name, 99) < STAGE_ORDER.get(prev, 99):
                mejor_stage[equipo] = stage_name

        if stage_name == "FINAL":
            resultados_finales[1] = winner_m
            resultados_finales[2] = loser_m
        elif stage_name == "THIRD_PLACE":
            resultados_finales[3] = winner_m
            resultados_finales[4] = loser_m
        elif stage_name not in ("GROUP_STAGE",):
            eliminados.add(loser_m)

    posicion_final = {eq: pos for pos, eq in resultados_finales.items()}

    fase_label = "🏆 Torneo Finalizado" if winner else f"📍 {STAGE_LABELS.get(stage_actual, stage_actual)}"
    vivos = todos - eliminados if eliminados else todos

    return {
        "vivos": vivos, "eliminados": eliminados,
        "resultados": resultados_finales, "fase_actual": fase_label,
        "torneo_iniciado": True,
        "posicion_final": posicion_final, "mejor_stage": mejor_stage,
    }

# ─── Datos demo ───────────────────────────────────────────────────────────────

def fetch_demo() -> dict:
    vivos = {"Argentina", "France", "Brazil", "Spain", "Germany",
             "Portugal", "Morocco", "Netherlands"}
    eliminados = {"Mexico", "United States", "Japan", "Senegal",
                  "Poland", "Switzerland", "England", "South Korea", "Australia"}
    mejor_stage = {e: "GROUP_STAGE" for e in eliminados}
    mejor_stage.update({e: "QUARTER_FINALS" for e in vivos})
    return {
        "vivos": vivos, "eliminados": eliminados, "resultados": {},
        "fase_actual": "📍 Cuartos de Final (demo)",
        "torneo_iniciado": True, "posicion_final": {}, "mejor_stage": mejor_stage,
    }

def fetch_demo_terminado() -> dict:
    """
    Resultado final demo:
      1° Argentina, 2° France, 3° Brazil, 4° Spain

    Apuestas de ejemplo (del Excel):
      Carlos:  1°Arg  2°Fra  3°Bra  4°Esp → 4+3+2+1 = 10 pts  (ganó todo)
      María:   1°Bra  2°Arg  3°Fra  4°Ale → 0+0+0+0 =  0 pts
      Pedro:   1°Fra  2°Esp  3°Arg  4°Bra → 0+0+0+0 =  0 pts
      Sofía:   1°Esp  2°Bra  3°Ale  4°Fra → 0+0+0+0 =  0 pts
      Javier:  1°Ale  2°Arg  3°Esp  4°Bra → 0+0+0+0 =  0 pts

    Desempate entre 0-pts: quien apostó al equipo que llegó más arriba.
      María  apostó 1°→Brasil (llegó 3°)
      Pedro  apostó 1°→Francia (llegó 2°)  ← mejor que Brasil
      Sofía  apostó 1°→España (llegó 4°)
      Javier apostó 1°→Alemania (eliminado → pos 99)
    → Pedro 2°, María 3°, Sofía 4°, Javier 5°
    """
    resultados = {1: "Argentina", 2: "France", 3: "Brazil", 4: "Spain"}
    posicion_final = {v: k for k, v in resultados.items()}
    mejor_stage = {
        "Argentina": "FINAL", "France": "FINAL",
        "Brazil": "THIRD_PLACE", "Spain": "THIRD_PLACE",
        "Germany": "SEMI_FINALS", "Portugal": "SEMI_FINALS",
    }
    return {
        "vivos": set(resultados.values()), "eliminados": {"Germany", "Portugal"},
        "resultados": resultados, "fase_actual": "🏆 Torneo Finalizado (demo)",
        "torneo_iniciado": True,
        "posicion_final": posicion_final, "mejor_stage": mejor_stage,
    }

# ─── Cálculo de puntajes y desempate ─────────────────────────────────────────

def calcular(apuestas: list[dict], standings: dict) -> list[dict]:
    resultados      = standings["resultados"]
    vivos           = standings["vivos"]
    torneo_iniciado = standings["torneo_iniciado"]
    posicion_final  = standings["posicion_final"]
    mejor_stage     = standings["mejor_stage"]

    calculados = []
    for a in apuestas:
        pts_reales  = 0
        pts_maximos = 0
        detalle     = {}

        for pos, equipo in a["picks"].items():
            if equipo is None:
                detalle[pos] = {"equipo": "—", "estado": "vivo", "pts": 0}
                continue
            pts_pos = POINTS[pos]

            if resultados:
                ganado = resultados.get(pos) == equipo
                pts_reales  += pts_pos if ganado else 0
                pts_maximos += pts_pos if ganado else 0
                detalle[pos] = {
                    "equipo": equipo,
                    "estado": "acierto" if ganado else "fallo",
                    "pts":    pts_pos if ganado else 0,
                }
            elif not torneo_iniciado or not vivos:
                pts_maximos += pts_pos
                detalle[pos] = {"equipo": equipo, "estado": "vivo", "pts": pts_pos}
            else:
                vivo = equipo in vivos
                pts_maximos += pts_pos if vivo else 0
                detalle[pos] = {
                    "equipo": equipo,
                    "estado": "vivo" if vivo else "eliminado",
                    "pts":    pts_pos if vivo else 0,
                }

        calculados.append({
            **a,
            "pts_reales":  pts_reales,
            "pts_maximos": pts_maximos,
            "detalle":     detalle,
        })

    torneo_terminado = bool(resultados)

    def tiebreak_key(a):
        """
        Desempate: para cada posición apostada (1°→4°), ¿qué tan arriba llegó ese equipo?
        posicion_final[equipo] = 1 si fue campeón, 2 si subcampeón, etc.
        Si el equipo no llegó al top4 usamos su mejor_stage (más es peor).
        Si ni siquiera aparece, asignamos 99 (eliminado temprano).

        IMPORTANTE: la comparación es sobre la posición FINAL del equipo apostado,
        no sobre la posición en que fue apostado. Así, quien apostó a Francia como
        1° gana el desempate sobre quien apostó a Brasil como 1° si Francia terminó
        en 2° y Brasil en 3°, porque Francia llegó más arriba.
        """
        picks_ordenados = [a["picks"].get(pos) for pos in [1, 2, 3, 4]]
        result = []
        for eq in picks_ordenados:
            if eq is None:
                result.append(99)
            elif eq in posicion_final:
                result.append(posicion_final[eq])          # 1, 2, 3 o 4
            else:
                # No llegó al top4: usamos cuánto avanzó como desempate secundario
                stage = mejor_stage.get(eq, "GROUP_STAGE")
                result.append(10 + STAGE_ORDER.get(stage, 99))
        return tuple(result)

    def tiebreak_key_en_curso(a):
        picks_ordenados = [a["picks"].get(pos) for pos in [1, 2, 3, 4]]
        result = []
        for eq in picks_ordenados:
            if eq is None:
                result.append(99)
            elif eq not in vivos:
                # Solo penaliza equipos eliminados, usando su mejor stage
                stage = mejor_stage.get(eq, "GROUP_STAGE")
                result.append(STAGE_ORDER.get(stage, 99))
            else:
                # Equipo vivo: todos equivalentes, no desempatar por stage aún
                result.append(0)
        return tuple(result)

    if torneo_terminado:
        calculados.sort(key=lambda x: (-x["pts_reales"], tiebreak_key(x)))
    else:
        calculados.sort(key=lambda x: (-x["pts_maximos"], tiebreak_key_en_curso(x)))

    # Asignar rank solo entre quienes pagaron, respetando empates
    pagantes = [a for a in calculados if a["pago"]]
    for i, a in enumerate(calculados):
        a["rank"] = None  # no pagó

    for i, a in enumerate(pagantes):
        if torneo_terminado:
            a["_sort_key"] = (-a["pts_reales"],) + tiebreak_key(a)
        else:
            a["_sort_key"] = (-a["pts_maximos"],) + tiebreak_key_en_curso(a)

    rank = 1
    for i, a in enumerate(pagantes):
        if i == 0:
            a["rank"] = 1
        else:
            if a["_sort_key"] == pagantes[i-1]["_sort_key"]:
                a["rank"] = pagantes[i-1]["rank"]
            else:
                a["rank"] = i + 1

    return calculados

# ─── Generación HTML ──────────────────────────────────────────────────────────

ESTADO_ICON  = {"vivo": "🟢", "eliminado": "🔴", "acierto": "✅", "fallo": "❌"}
ESTADO_LABEL = {"vivo": "Vivo", "eliminado": "Eliminado", "acierto": "¡Acertó!", "fallo": "Falló"}
POS_LABEL    = {1: "🥇 1°", 2: "🥈 2°", 3: "🥉 3°", 4: "4°"}

def generar_html(apuestas_calc: list[dict], standings: dict, generado: str) -> str:
    pagantes         = [a for a in apuestas_calc if a["pago"]]
    total_part       = len(pagantes)
    pozo             = total_part * 100_000
    premio_1         = max(pozo - 200_000, 0)
    fase             = standings["fase_actual"]
    torneo_terminado = bool(standings["resultados"])

    def clp(n):
        return f"${n:,.0f}".replace(",", ".")

    # ── Tarjetas ──────────────────────────────────────────────────────────────
    cards_html = ""
    for a in pagantes:
        rank   = a["rank"]
        empate = sum(1 for x in pagantes if x["rank"] == rank) > 1

        if torneo_terminado and not empate:
            if rank == 1:   rank_badge = '<span class="rank gold">👑 1° Lugar</span>'
            elif rank == 2: rank_badge = '<span class="rank silver">🥈 2° Lugar</span>'
            else:           rank_badge = f'<span class="rank">{rank}° Lugar</span>'
        elif torneo_terminado and empate:
            if rank == 1:   rank_badge = '<span class="rank gold tie">🤝 Empate 1°</span>'
            elif rank == 2: rank_badge = '<span class="rank silver tie">🤝 Empate 2°</span>'
            else:           rank_badge = f'<span class="rank tie">🤝 Empate {rank}°</span>'
        else:
            if empate:
                rank_badge = f'<span class="rank tentative">≈ {rank}° (empate)</span>'
            else:
                rank_badge = f'<span class="rank tentative">{rank}°</span>'

        if torneo_terminado and rank <= 2 and not empate:
            card_class = f"card rank-{rank}"
        elif rank == 1 and not empate:
            card_class = "card rank-lead"
        else:
            card_class = "card"

        picks_html = ""
        for pos in [1, 2, 3, 4]:
            d      = a["detalle"].get(pos, {})
            estado = d.get("estado", "vivo")
            # Mostrar nombre en español
            nombre_es = a_es(d.get("equipo", "—"))
            picks_html += f"""
            <div class="pick {estado}">
              <span class="pick-pos">{POS_LABEL[pos]}</span>
              <span class="pick-team">{nombre_es}</span>
              <span class="pick-badge">{ESTADO_ICON[estado]} {ESTADO_LABEL[estado]}</span>
              <span class="pick-pts">+{d.get('pts', POINTS[pos])} pts</span>
            </div>"""

        score_val   = a["pts_reales"] if torneo_terminado else a["pts_maximos"]
        score_label = "Puntos totales" if torneo_terminado else "Máx. alcanzable"

        tiebreak_note = ""
        if torneo_terminado and empate:
            tiebreak_note = '<div class="tiebreak-note">⚖️ Empate — se comparte el premio</div>'

        cards_html += f"""
        <div class="{card_class}">
          <div class="card-header">
            {rank_badge}
            <h3>{a['nombre']}</h3>
            <div class="score-box">
              <div class="score-main">{score_val}<span>pts</span></div>
              <div class="score-label">{score_label}</div>
            </div>
          </div>
          <div class="picks-grid">{picks_html}</div>
          {tiebreak_note}
        </div>"""

    # ── Resultados oficiales (nombres en español) ─────────────────────────────
    resultados_html = ""
    if standings["resultados"]:
        resultados_html = '<div class="resultados-finales"><h2>🏆 Resultados Oficiales</h2><div class="res-grid">'
        for pos, equipo in standings["resultados"].items():
            resultados_html += (f'<div class="res-item pos-{pos}">'
                                f'<span class="res-pos">{POS_LABEL[pos]}</span>'
                                f'<span class="res-team">{a_es(equipo)}</span></div>')
        resultados_html += "</div></div>"

    # ── Banner de empate durante el torneo ────────────────────────────────────
    hay_empates = (not torneo_terminado and
                   any(sum(1 for x in pagantes if x["rank"] == a["rank"]) > 1
                       for a in pagantes))
    nota_desempate = ""
    if hay_empates:
        nota_desempate = """
        <div class="empate-banner">
          ⚖️ Hay participantes empatados en puntaje máximo. El desempate final se resolverá
          según qué equipos apostados hayan llegado más arriba en el torneo (1° &gt; 2° &gt; 3° &gt; 4°).
        </div>"""

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
    --bg:         #f7f5f0;
    --surface:    #ffffff;
    --surface2:   #f0ede6;
    --border:     #e2ddd4;
    --gold:       #c9a84c;
    --gold-dark:  #a8832a;
    --gold-light: #f0d98a;
    --silver:     #8a9ba8;
    --bronze:     #a0674a;
    --green:      #2d7a4f;
    --green-bg:   #eaf4ee;
    --red:        #c0392b;
    --red-bg:     #fdf0ee;
    --text:       #1a1a1a;
    --muted:      #9a9590;
    --radius:     14px;
    --can-red:    #c0392b;
    --usa-blue:   #1a3a6e;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'DM Sans', sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; }}
  .top-stripe {{
    height: 5px;
    background: linear-gradient(90deg,
      var(--can-red) 0%, var(--can-red) 33.3%,
      var(--gold) 33.3%, var(--gold) 66.6%,
      var(--usa-blue) 66.6%, var(--usa-blue) 100%);
  }}
  header {{
    background: #1a1a1a; text-align: center;
    padding: 48px 24px 36px; position: relative; overflow: hidden;
  }}
  header::before {{
    content: '26 26 26 26 26 26 26 26 26';
    position: absolute; top: 50%; left: 50%;
    transform: translate(-50%, -50%);
    font-family: 'Bebas Neue', sans-serif; font-size: 120px;
    color: rgba(255,255,255,0.03); white-space: nowrap;
    letter-spacing: 20px; pointer-events: none;
  }}
  .title {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: clamp(48px, 9vw, 88px); letter-spacing: 6px;
    line-height: 1; color: #fff; position: relative;
  }}
  .title span {{ color: var(--gold); }}
  .subtitle {{
    margin-top: 8px; color: #888; font-size: 13px;
    letter-spacing: 4px; text-transform: uppercase; position: relative;
  }}
  .fase-badge {{
    display: inline-flex; align-items: center; gap: 6px;
    margin-top: 18px; padding: 7px 20px; border-radius: 999px;
    background: rgba(201,168,76,0.12); border: 1px solid rgba(201,168,76,0.35);
    color: var(--gold-light); font-size: 13px; font-weight: 500; position: relative;
  }}
  .stats {{
    display: flex; justify-content: center; gap: 1px;
    background: var(--border); border-bottom: 1px solid var(--border); overflow-x: auto;
  }}
  .stat {{ flex: 1; min-width: 140px; background: var(--surface); padding: 20px 24px; text-align: center; }}
  .stat-val {{ font-family: 'Bebas Neue', sans-serif; font-size: 32px; letter-spacing: 1px; color: var(--gold-dark); line-height: 1; }}
  .stat-label {{ font-size: 11px; color: var(--muted); text-transform: uppercase; letter-spacing: 1.5px; margin-top: 4px; }}
  .resultados-finales {{ max-width: 700px; margin: 40px auto 0; padding: 0 24px; }}
  .resultados-finales h2 {{ font-family: 'Bebas Neue', sans-serif; font-size: 26px; letter-spacing: 2px; margin-bottom: 14px; text-align: center; color: var(--gold-dark); }}
  .res-grid {{ display: grid; grid-template-columns: repeat(2,1fr); gap: 10px; }}
  .res-item {{ background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 14px 18px; display: flex; justify-content: space-between; align-items: center; }}
  .res-item.pos-1 {{ border-left: 3px solid var(--gold); }}
  .res-item.pos-2 {{ border-left: 3px solid var(--silver); }}
  .res-item.pos-3 {{ border-left: 3px solid var(--bronze); }}
  .res-pos {{ font-size: 16px; }}
  .res-team {{ font-weight: 600; }}
  .cards-section {{ max-width: 1100px; margin: 0 auto; padding: 40px 24px 60px; }}
  .section-title {{ font-family: 'Bebas Neue', sans-serif; font-size: 22px; letter-spacing: 3px; color: var(--muted); margin-bottom: 20px; padding-bottom: 10px; border-bottom: 1px solid var(--border); }}
  .empate-banner {{ background: #fff8e1; border: 1px solid #f0c060; border-radius: 10px; padding: 12px 18px; font-size: 13px; color: #7a5c00; margin-bottom: 20px; line-height: 1.5; }}
  .cards {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(320px,1fr)); gap: 16px; }}
  .card {{ background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); overflow: hidden; transition: transform .2s, box-shadow .2s; }}
  .card:hover {{ transform: translateY(-3px); box-shadow: 0 12px 32px rgba(0,0,0,.1); }}
  .card.rank-1  {{ border-color: var(--gold); box-shadow: 0 0 0 1px rgba(201,168,76,.3); }}
  .card.rank-2  {{ border-color: var(--silver); }}
  .card.rank-3  {{ border-color: var(--bronze); }}
  .card.rank-lead {{ border-color: rgba(201,168,76,.4); }}
  .card-header {{ padding: 16px 18px 14px; display: flex; align-items: center; gap: 10px; border-bottom: 1px solid var(--border); background: var(--surface2); }}
  .card-header h3 {{ flex: 1; font-size: 17px; font-weight: 600; }}
  .rank {{ font-size: 12px; font-weight: 600; padding: 3px 10px; border-radius: 999px; background: var(--surface); border: 1px solid var(--border); color: var(--muted); white-space: nowrap; }}
  .rank.gold   {{ background: #fdf6e3; border-color: var(--gold);   color: var(--gold-dark); }}
  .rank.silver {{ background: #f4f6f8; border-color: var(--silver); color: #607080; }}
  .rank.bronze {{ background: #fdf0e8; border-color: var(--bronze); color: var(--bronze); }}
  .rank.tentative {{ background: #f5f5f5; border-color: #ccc; color: #888; font-style: italic; }}
  .score-box {{ text-align: right; }}
  .score-main {{ font-family: 'Bebas Neue', sans-serif; font-size: 30px; line-height: 1; color: var(--gold-dark); }}
  .score-main span {{ font-size: 13px; color: var(--muted); font-family: 'DM Sans',sans-serif; margin-left: 2px; }}
  .score-label {{ font-size: 10px; color: var(--muted); text-transform: uppercase; letter-spacing: .8px; }}
  .picks-grid {{ padding: 14px; display: flex; flex-direction: column; gap: 7px; }}
  .pick {{ display: grid; grid-template-columns: 46px 1fr auto auto; align-items: center; gap: 8px; padding: 9px 12px; border-radius: 8px; border: 1px solid var(--border); background: var(--bg); }}
  .pick.vivo      {{ border-color: #b7dfc8; background: var(--green-bg); }}
  .pick.acierto   {{ border-color: #2d7a4f; background: var(--green-bg); }}
  .pick.eliminado {{ border-color: #f0c4be; background: var(--red-bg); opacity: .7; }}
  .pick.fallo     {{ border-color: #f0c4be; background: var(--red-bg); opacity: .6; }}
  .pick-pos   {{ font-size: 12px; color: var(--muted); }}
  .pick-team  {{ font-weight: 500; font-size: 14px; }}
  .pick-badge {{ font-size: 11px; color: var(--muted); white-space: nowrap; }}
  .pick-pts   {{ font-size: 12px; font-weight: 600; color: var(--muted); white-space: nowrap; }}
  .pick.vivo .pick-pts    {{ color: var(--green); }}
  .pick.acierto .pick-pts {{ color: var(--green); }}
  .pick.eliminado .pick-pts {{ color: var(--red); }}
  .tiebreak-note {{ padding: 10px 16px; font-size: 12px; color: #7a5c00; background: #fff8e1; border-top: 1px solid #f0c060; }}
  footer {{ text-align: center; padding: 20px 24px; color: var(--muted); font-size: 12px; border-top: 1px solid var(--border); background: var(--surface); }}
  @media (max-width: 600px) {{
    .cards       {{ grid-template-columns: 1fr; }}
    .res-grid    {{ grid-template-columns: 1fr; }}
    .stats       {{ flex-wrap: wrap; }}
    .stat        {{ min-width: calc(50% - 1px); padding: 14px 12px; }}
    .stat-val    {{ font-size: 24px; }}
    .stat-label  {{ font-size: 10px; }}
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
  <div class="stat"><div class="stat-val">{total_part}</div><div class="stat-label">Participantes</div></div>
  <div class="stat"><div class="stat-val">{clp(pozo)}</div><div class="stat-label">Pozo total</div></div>
  <div class="stat"><div class="stat-val">{clp(premio_1)}</div><div class="stat-label">Premio 1° lugar</div></div>
  <div class="stat"><div class="stat-val">$200.000</div><div class="stat-label">Premio 2° lugar</div></div>
</div>
{resultados_html}
<div class="cards-section">
  <div class="section-title">Clasificación</div>
  {nota_desempate}
  <div class="cards">
{cards_html}
  </div>
</div>
<footer>
  Última actualización: {generado} &nbsp;·&nbsp; football-data.org
</footer>
</body>
</html>"""

# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Generador Mundial 2026")
    parser.add_argument("--api-key",    help="API key de football-data.org")
    parser.add_argument("--demo",       action="store_true", help="Demo torneo en curso")
    parser.add_argument("--demo-final", action="store_true", help="Demo torneo terminado")
    parser.add_argument("--excel",      default=EXCEL_FILE)
    parser.add_argument("--output",     default=OUTPUT_FILE)
    args = parser.parse_args()

    if not args.demo and not args.demo_final and not args.api_key:
        print("❌  Debes pasar --api-key, --demo o --demo-final.")
        sys.exit(1)

    print(f"📂  Leyendo apuestas desde {args.excel}...")
    apuestas = leer_apuestas(args.excel)
    print(f"    {len(apuestas)} participantes encontrados.")

    if args.demo_final:
        print("🏆  Modo demo-final.")
        standings = fetch_demo_terminado()
    elif args.demo:
        print("🎮  Modo demo.")
        standings = fetch_demo()
    else:
        print("🌐  Consultando API...")
        try:
            standings = fetch_standings(args.api_key)
        except requests.HTTPError as e:
            print(f"❌  Error de API: {e}")
            sys.exit(1)

    print(f"    Fase: {standings['fase_actual']}")
    apuestas_calc = calcular(apuestas, standings)
    generado      = generado = datetime.now(ZoneInfo("America/Santiago")).strftime("%d/%m/%Y %H:%M")
    html          = generar_html(apuestas_calc, standings, generado)
    Path(args.output).write_text(html, encoding="utf-8")
    print(f"✅  HTML generado: {args.output}")

if __name__ == "__main__":
    main()