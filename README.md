# Panel de Flota

Prototipo del sistema de gestión de flota de camiones: control de camiones,
acoplados, choferes, neumáticos (con posiciones y estado de recapado),
rutas y viajes.

## Requisitos

- Node.js 18 o superior
- Un proyecto de Supabase con el esquema de `schema-supabase.sql` ya corrido
  (ver ese archivo para las instrucciones)

## Cómo correrlo localmente

```bash
npm install
cp .env.example .env
```

Completá `.env` con la URL y la anon key de tu proyecto de Supabase
(Project Settings → API), y después:

```bash
npm run dev
```

Esto levanta un servidor local (por defecto en `http://localhost:5173`).

## Cómo subirlo a GitHub

1. Creá un repositorio nuevo en GitHub (sin README, para evitar conflictos).
2. Desde esta carpeta, corré:

```bash
git init
git add .
git commit -m "Primer commit: prototipo panel de flota"
git branch -M main
git remote add origin https://github.com/<tu-usuario>/<nombre-repo>.git
git push -u origin main
```

## Estructura

```
panel-flota/
├── index.html
├── package.json
├── vite.config.js
├── tailwind.config.js
├── postcss.config.js
├── src/
│   ├── main.jsx
│   ├── App.jsx
│   ├── PanelFlota.jsx   ← componente principal del panel
│   └── index.css
```

## Estado del prototipo

Los datos son de ejemplo y viven en memoria del navegador (no hay backend
ni base de datos todavía). Pensado como base visual y funcional para
seguir iterando sobre el sistema real.
