# ⚽ Mundial 2026 — Control de Apuestas

Página web generada automáticamente con los puntajes en tiempo real de tu grupo de apuestas.

## ¿Cómo funciona?

```
apuestas.xlsx  →  generate.py  →  index.html  →  GitHub Pages
                      ↑
              football-data.org API
              (equipos vivos / resultados)
```

---

## 🚀 Configuración paso a paso

### 1. Obtener API key gratuita

1. Ve a **https://www.football-data.org/client/register**
2. Regístrate con tu email (es gratis)
3. Recibirás tu API key por correo — guárdala

### 2. Crear el repositorio en GitHub

1. Ve a **https://github.com/new**
2. Nombre del repo: `mundial-2026` (o el que quieras)
3. Márcalo como **Public** (necesario para GitHub Pages gratis)
4. Haz clic en **Create repository**

### 3. Subir los archivos

Desde tu terminal, en la carpeta del proyecto:

```bash
git init
git add .
git commit -m "Primer commit - Mundial 2026"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/mundial-2026.git
git push -u origin main
```

### 4. Activar GitHub Pages

1. En tu repo, ve a **Settings → Pages**
2. En *Source*, selecciona **Deploy from a branch**
3. Branch: **main** / Folder: **/ (root)**
4. Haz clic en **Save**
5. En ~2 minutos tu página estará en:
   `https://TU_USUARIO.github.io/mundial-2026/`

### 5. Guardar la API key como secreto

1. En tu repo ve a **Settings → Secrets and variables → Actions**
2. Haz clic en **New repository secret**
3. Name: `FOOTBALL_API_KEY`
4. Value: tu API key del paso 1
5. Haz clic en **Add secret**

### 6. Activar GitHub Actions

1. Ve a la pestaña **Actions** de tu repo
2. Acepta activar los workflows si te lo pide
3. Listo — la página se regenerará **automáticamente cada hora**

---

## 🖥️ Uso local (actualización manual)

```bash
# Instalar dependencias
pip install openpyxl requests

# Generar con datos reales
python generate.py --api-key TU_API_KEY

# Generar con datos de prueba (sin API)
python generate.py --demo

# Abrir en el navegador
open index.html   # macOS
start index.html  # Windows
```

---

## 📋 Estructura del Excel (`apuestas.xlsx`)

| Columna | Contenido |
|---------|-----------|
| A | Nombre del participante |
| B | Selección para el 1° lugar |
| C | Selección para el 2° lugar |
| D | Selección para el 3° lugar |
| E | Selección para el 4° lugar |
| F | ¿Pagó? (✓ Sí / ✗ No) |
| G | Observaciones |

> Los nombres de los equipos deben coincidir exactamente con los que usa football-data.org
> (ej: "Argentina", "France", "Brazil"). Ejecuta `--demo` primero para ver el formato.

---

## 🏆 Sistema de puntaje

| Posición apostada | Puntos si acierta |
|-------------------|-------------------|
| 1° lugar          | 4 puntos          |
| 2° lugar          | 3 puntos          |
| 3° lugar          | 2 puntos          |
| 4° lugar          | 1 punto           |

**Durante el torneo**: la página muestra el máximo de puntos que cada participante aún puede alcanzar según qué equipos siguen vivos.

**Premios**:
- 🥇 1° lugar: Pozo total − $100.000
- 🥈 2° lugar: $100.000 fijo
- Desempate: el participante cuyo equipo haya llegado más arriba gana

---

## 🔄 Actualización automática

El archivo `.github/workflows/update.yml` configura GitHub Actions para regenerar `index.html` **cada hora** consultando la API. También puedes dispararlo manualmente desde la pestaña *Actions* de tu repo.
