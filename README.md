# Hoja — Excel local

Editor de datos en el navegador, creado el 2026-09-25. Los archivos se procesan en memoria; no se envían a servicios externos ni se guardan automáticamente.

## Abrir

```bash
cd /home/sergio_castro/Proyectos/Metis/tools/excel-local
python3 serve.py
```

Visitar http://127.0.0.1:8766. El servidor solo expone esta carpeta en loopback.

## Alcance

Importar XLSX, XLS y CSV; editar celdas; pegar tablas; buscar; ordenar; añadir filas, columnas y hojas; deshacer/rehacer; exportar una copia XLSX o la hoja activa como CSV. La interfaz muestra hasta 100 filas por página.

Es un editor de datos. No conserva gráficos, macros, formato visual avanzado ni combinaciones de celdas. Las fórmulas se muestran y guardan como fórmulas; se recalculan al abrir la copia en Excel. Las operaciones que moverían referencias de fórmulas se bloquean cuando hay fórmulas en el libro. Máximo 20 MB, 256 columnas y 100.000 filas por hoja, y 500.000 celdas ocupadas por libro.

Dependencia local: [SheetJS CE 0.20.3](https://docs.sheetjs.com/docs/getting-started/installation/standalone/), licencia Apache 2.0 en `vendor/LICENSE`. No necesita Node ni instalar paquetes para funcionar.

## Estado y verificación

- Dependencia descargada del CDN oficial y guardada localmente.
- Interfaz y operaciones implementadas. Pruebas del modelo: 7/7 pasan (lectura/escritura XLSX, tipos, fechas, fórmulas, ordenación, pegado, límites, CSV, fechas 1904 y rangos con nombre).
- Chromium real: crear, pegar, editar, deshacer/rehacer, ordenar, filtrar, añadir/renombrar hojas, importar y descargar XLSX, conservar fórmulas, navegar entre 225 filas y usar la vista móvil. Todo pasa, sin errores JavaScript.
- Se volvió a leer el archivo descargado para verificar los cambios y ambas hojas. Capturas en `.test-output/` (ignorado por Git).
- Correcciones durante la prueba: navegar a una celda elimina correctamente el filtro; doble clic permite renombrar la hoja; cambiar de celda mientras se edita conserva el nuevo foco.

Pruebas reproducibles (con el servidor iniciado):

```bash
node --test data.test.mjs
NODE_PATH=/home/sergio_castro/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules node browser.test.cjs
```

La prueba de navegador usa el Playwright incluido en el entorno y Chromium del sistema. `CHROMIUM_PATH` permite especificar otro ejecutable. El runtime de la página no depende de ellos.

Reversión: eliminar únicamente `tools/excel-local/` y detener su servidor. No modifica otros proyectos ni servicios.
