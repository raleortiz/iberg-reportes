# Iberg Reportes - Sistema de Gestión de Informes de Ventas

## Ubicación del proyecto

**Directorio local:** `D:\PROYECTOS\PROYECTO IBERG`

## Tecnologías utilizadas

| Componente | Servicio | Cuenta de acceso |
|------------|----------|------------------|
| **Código fuente** | GitHub | Usuario: `raleortiz` — Repo: `iberg-reportes` |
| **Hosting (servidor)** | Render | Login con GitHub |
| **Base de datos** | Supabase | Login con GitHub — URL: `https://flzwjulbcbpbzqoyawtb.supabase.co` |

## Estructura del proyecto

```
PROYECTO IBERG/
├── ARCHIVOS/                  ← Excel originales (no se sube a GitHub)
│   ├── EXCEL ART.xlsx
│   └── EXCEL VENTAS.xlsx
├── public/
│   └── index.html             ← Frontend (HTML + JavaScript puro)
├── server.js                  ← Backend (Node.js + Express)
├── render.yaml                ← Configuración de deploy en Render
├── package.json               ← Dependencias del proyecto
├── .env                       ← Variables de entorno (no se sube a GitHub)
├── .gitignore                 ← Archivos excluidos del repositorio
└── README.md
```

## Variables de entorno (Render y .env)

```
SUPABASE_URL=https://flzwjulbcbpbzqoyawtb.supabase.co
SUPABASE_KEY=<secret_key_de_supabase_-_ver_en_render_o_archivo_.env>
```

## URLs del proyecto

- **App en producción:** https://iberg-reportes.onrender.com
- **GitHub:** https://github.com/raleortiz/iberg-reportes
- **Supabase Dashboard:** https://supabase.com/dashboard/project/flzwjulbcbpbzqoyawtb

## Base de datos (Supabase)

Tres tablas en el esquema `public`:

- **informes** — Encabezados de cada informe cargado (empresa, NIT, fechas, tipo ART/VENTAS)
- **detalle_art** — Productos del informe ART (código, nombre, cantidades, valores, %)
- **detalle_ventas** — Clientes del informe VENTAS (zona, cliente, cantidades, valores, %)

## Keep-Alive (UptimeRobot)

Se configuró UptimeRobot (https://uptimerobot.com) para hacer ping cada 5 minutos
a `https://iberg-reportes.onrender.com/health` y evitar que Render duerma el servicio.

## Cómo retomar el proyecto

1. Clonar el repo: `git clone https://github.com/raleortiz/iberg-reportes.git`
2. Instalar dependencias: `npm install`
3. Crear archivo `.env` con las variables de entorno (ver arriba)
4. Ejecutar localmente: `npm start` → abre http://localhost:3000
5. Para subir cambios: `git add . && git commit -m "mensaje" && git push`
6. Render deploya automáticamente al hacer push a `main`

## Funcionalidades

- Subir archivos Excel (ART y VENTAS) desde el navegador
- Parser automático que extrae productos, clientes, zonas y fechas
- Visualización de datos en tablas con filtros por zona
- Exportar informes a Excel organizado (tabla plana)
- Eliminar informes cargados
