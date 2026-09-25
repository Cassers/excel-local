# Hoja — Excel local

Editor de datos en el navegador, creado el 2026-09-25. Los archivos se procesan en memoria; no se envían a servicios externos ni se guardan automáticamente.

## Abrir

Haz doble clic en `index.html` para abrir el editor directamente en tu navegador. Conserva los demás archivos y la carpeta `vendor` junto al HTML. No requiere servidor ni instalación.

Opcionalmente, puedes usar un servidor local:

```bash
cd /home/sergio_castro/Proyectos/excel-local
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
