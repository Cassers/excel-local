# Hoja — Excel local

Editor de datos en el navegador, creado el 2026-09-25. Los archivos pequeños se procesan en memoria. El modo grande usa un servicio en tu equipo y datos temporales en disco. No se envían archivos a servicios externos. Descarga una copia para conservar los cambios.

## Abrir

Haz doble clic en `index.html` para abrir el editor directamente en tu navegador. Conserva los demás archivos y la carpeta `vendor` junto al HTML. No requiere servidor ni instalación.

Opcionalmente, puedes usar un servidor local:

```bash
cd /home/sergio_castro/Proyectos/excel-local
python3 serve.py
```

Visitar http://127.0.0.1:8766. El servidor solo expone esta carpeta en loopback.

## Archivos de 100–500 MB

```bash
cd /home/sergio_castro/Proyectos/excel-local
./abrir.sh
```

Abre http://127.0.0.1:8768/ y pulsa **Archivos grandes**. También se selecciona ese modo automáticamente al elegir un archivo mayor de 20 MB. Acepta XLSX y CSV; convierte XLS antiguos a XLSX. El lanzador prepara un entorno Python local e instala las dependencias la primera vez. Para cambiar el puerto: `HOJA_PORT=8770 ./abrir.sh`.

El modo grande lee y escribe por partes, almacena las celdas en SQLite y muestra hasta 100 filas × 24 columnas por página. Permite editar celdas, pegar hasta 10.000 celdas por operación, deshacer/rehacer, buscar, cambiar de hoja y descargar XLSX/CSV. Conserva la estructura de las hojas: la ordenación y las operaciones de añadir filas, columnas u hojas siguen disponibles en el editor pequeño.

No hay un límite artificial de MB en el modo grande. Sigue dependiendo del contenido, el espacio libre, la RAM y los límites de Excel (1.048.576 filas × 16.384 columnas, 32.767 caracteres por celda). El servicio comprueba el espacio disponible antes de importar. Un XLSX muy comprimido puede ocupar mucho más al descomprimirse.

Los temporales quedan en `.local-data/`, excluido de Git y del servidor público. **Cerrar libro** elimina su copia temporal; Ctrl+C cierra el servicio y limpia sus sesiones. Si se interrumpe el proceso a la fuerza, pueden quedar residuos en esa carpeta. La página ofrece progreso y cancelación; la fase final de compresión puede tardar en responder a la cancelación. No hay recuperación automática de cambios tras cerrar la página o el servicio.

Implementación y mediciones reproducibles: [LARGE_FILES.md](LARGE_FILES.md). Motor basado en los [modos de lectura/escritura progresiva de openpyxl](https://openpyxl.readthedocs.io/en/stable/optimized.html), con textos compartidos también almacenados en disco.

## Alcance

Importar XLSX, XLS y CSV; editar celdas; pegar tablas; buscar; ordenar; añadir filas, columnas y hojas; deshacer/rehacer; exportar una copia XLSX o la hoja activa como CSV. La interfaz muestra hasta 100 filas por página.

Es un editor de datos. No conserva gráficos, macros, formato visual avanzado ni combinaciones de celdas. Las fórmulas se muestran y guardan como fórmulas; se recalculan al abrir la copia en Excel. Las operaciones que moverían referencias de fórmulas se bloquean cuando hay fórmulas en el libro. El editor pequeño tiene un máximo de 20 MB, 256 columnas y 100.000 filas por hoja, y 500.000 celdas ocupadas por libro. El modo grande reemplaza esos límites con paginación y almacenamiento en disco.

Dependencia local: [SheetJS CE 0.20.3](https://docs.sheetjs.com/docs/getting-started/installation/standalone/), licencia Apache 2.0 en `vendor/LICENSE`. No necesita Node ni instalar paquetes para funcionar.

## Estado y verificación

- Dependencia descargada del CDN oficial y guardada localmente.
- Interfaz y operaciones implementadas. Pruebas del modelo: 7/7 pasan (lectura/escritura XLSX, tipos, fechas, fórmulas, ordenación, pegado, límites, CSV, fechas 1904 y rangos con nombre).
- Chromium real: crear, pegar, editar, deshacer/rehacer, ordenar, filtrar, añadir/renombrar hojas, importar y descargar XLSX, conservar fórmulas, navegar entre 225 filas y usar la vista móvil. Todo pasa, sin errores JavaScript.
- Se volvió a leer el archivo descargado para verificar los cambios y ambas hojas. Capturas en `.test-output/` (ignorado por Git).
- Correcciones durante la prueba: navegar a una celda elimina correctamente el filtro; doble clic permite renombrar la hoja; cambiar de celda mientras se edita conserva el nuevo foco.

Pruebas reproducibles:

```bash
node build.mjs --check
node --test data.test.mjs
NODE_PATH=/home/sergio_castro/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules node browser.test.cjs
```

La prueba de navegador abre `index.html` mediante `file://` por defecto. Con el servidor iniciado, también se prueba HTTP añadiendo `APP_URL=http://127.0.0.1:8766/` al comando. Usa el Playwright incluido en el entorno y Chromium del sistema. `CHROMIUM_PATH` permite especificar otro ejecutable. El runtime de la página no depende de ellos.

## Apertura directa (2026-09-25)

El navegador bloqueaba `app.mjs` desde `file://` por CORS. `index.html` ahora carga `app.bundle.js` como script clásico local, después de SheetJS. La política de seguridad del navegador se conserva. El enlace del logotipo vuelve a `index.html` en ambos modos.

El bundle se incluye en el repositorio: los usuarios no necesitan compilarlo. Al modificar `app.mjs` o `data.mjs`, ejecutar `node build.mjs` y guardar también el bundle generado. `node build.mjs --check` detecta un bundle desactualizado.

Verificación del cambio: `node build.mjs --check` pasa; las 7 pruebas del modelo pasan; la prueba completa de Chromium pasa tanto en `file://` como en HTTP (servidor temporal en el puerto 8767), sin errores de JavaScript, CORS ni carga de recursos. Se verificaron importación, edición y descarga de XLSX en ambos modos.

Reversión de este ajuste: restaurar `index.html`, `browser.test.cjs`, `.gitignore` y esta documentación; retirar `build.mjs` y `app.bundle.js`. Los módulos originales no cambian.

Reversión de la herramienta completa: eliminar únicamente su carpeta y detener su servidor, si se inició. No modifica otros proyectos ni servicios.
