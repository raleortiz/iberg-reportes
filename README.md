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

## Cómo ejecutar en modo local

### 1. Instalar Node.js

1. Ve a https://nodejs.org
2. Descarga la versión **LTS** (recomendada)
3. Ejecuta el instalador y dale "Next" en todo
4. Para verificar que se instaló, abre una terminal y ejecuta:
   ```
   node --version
   npm --version
   ```
   Si muestra los números de versión, está listo.

### 2. Clonar el proyecto

```
git clone https://github.com/raleortiz/iberg-reportes.git
cd iberg-reportes
```

### 3. Instalar dependencias

```
npm install
```

### 4. Crear archivo .env

Crea un archivo llamado `.env` en la raíz del proyecto con estas dos líneas:

```
SUPABASE_URL=https://flzwjulbcbpbzqoyawtb.supabase.co
SUPABASE_KEY=<tu_secret_key_de_supabase>
```

La secret key la encuentras en:
- Supabase Dashboard → Settings → API Keys → Secret keys

### 5. Ejecutar el servidor

```
node server.js
```

Abre el navegador en **http://localhost:3000**

### 6. Subir cambios a GitHub

```
git add .
git commit -m "descripcion del cambio"
git push
```

Render deploya automáticamente al hacer push a la rama `main`.

## Funcionalidades

- Subir archivos Excel (ART y VENTAS) desde el navegador
- Parser automático que extrae productos, clientes, zonas y fechas
- Visualización de datos en tablas con filtros por zona
- Exportar informes a Excel organizado (tabla plana)
- Eliminar informes cargados
