# Archivos de 100–500 MB

Solicitud: ampliar la capacidad para libros de 100–500 MB conservando el editor en navegador y la apertura directa de archivos pequeños.

## Requisitos

- Al abrir XLSX/CSV grandes por HTTP local, transferir el File directamente a disco, sin leerlo entero en JavaScript.
- Importar en segundo plano, con progreso y cancelación; almacenar celdas y cadenas compartidas en SQLite.
- Consultar páginas acotadas de filas y columnas; editar, pegar, deshacer, buscar, cambiar de hoja y exportar una copia.
- Mantener el flujo `file://` existente. Para archivos grandes, explicar cómo iniciar el modo local. XLS antiguos grandes requieren conversión a XLSX.
- Los datos quedan en el equipo. El servicio escucha solo en loopback y rechaza solicitudes de otros orígenes. Los temporales no se publican por HTTP ni se incluyen en Git.
- Ningún tamaño se anuncia como ilimitado. Disco, memoria, dimensiones de Excel y tiempo siguen siendo límites reales.

## Diseño

`serve.py` sirve la interfaz y una API con token por ejecución. Un proceso local usa openpyxl read-only/write-only y SQLite. Las cadenas compartidas del XLSX se indexan en disco para evitar la lista completa en RAM de openpyxl. Un trabajo de importación/exportación a la vez; las consultas devuelven hasta 100 filas × 24 columnas. Los originales nunca se sobrescriben.

Se mantiene el editor sencillo para archivos de hasta 20 MB. El modo grande ofrece edición de datos, no ordenación ni cambios estructurales de hojas: no mueve referencias de fórmulas. Exporta valores, tipos, fórmulas, formatos numéricos y nombres definidos; se mantienen las limitaciones de gráficos, macros y formatos visuales avanzados.

## Tareas y evidencia

- [x] Almacenamiento e importación/exportación progresiva.
- [x] API local, cancelación y limpieza explícita.
- [x] Interfaz paginada y acceso al modo grande.
- [x] Pruebas de tipos, referencias, errores, aislamiento y recorridos reales en navegador.
- [x] Ensayo sintético de 100–500 MB con tiempo y memoria registrados.
- [x] Documentación y entrega de cambios en la rama `main`.

Reversión: restaurar el servidor y los cambios en la interfaz/build; retirar el módulo local, el cliente grande, lanzadores, pruebas y dependencias añadidos. Los archivos Excel originales no se modifican. `.local-data/` contiene trabajo temporal; descargar antes de cerrar o limpiar.

## Ensayo del motor (2026-09-25)

Python 3.14.7, openpyxl 3.1.5, equipo de Sergio. Proceso separado por archivo, medición `resource.ru_maxrss`.

| Archivo sintético | Tamaño real | Filas | Importación | Edición + exportación CSV | RSS máximo |
|---|---:|---:|---:|---:|---:|
| CSV | 100.045.473 bytes | 3.087 | 3,56 s | 0,81 s | 57,6 MiB |
| XLSX | 503.278.063 bytes | 15.433 | 16,17 s | 3,98 s | 688,9 MiB |

El XLSX usa ZIP sin compresión y textos repetidos largos para generar volumen reproducible. Un XLSX de 500 MB comprimido con millones de celdas puede requerir mucho más disco y tiempo. Estos resultados no garantizan cualquier libro de ese peso. Las bases temporales midieron 209 MB y 1.048 MB respectivamente. Los ensayos y sus salidas se generan en `.test-output/large/`, excluido de Git.

Repetición tras eliminar el sondeo de dimensiones: 18,61 s de importación, 3,94 s de edición/exportación y 729,4 MiB de máximo informado por `ru_maxrss`. Un perfil adicional leyendo `VmRSS` durante todo el flujo observó aproximadamente 58–59 MiB mientras se importaban las filas y 48 MiB durante la exportación CSV. Se conservan ambos resultados; el tiempo y el máximo del proceso varían entre ejecuciones.

Prueba real en Chromium, incluyendo transferencia HTTP local, edición, exportación CSV y lectura del archivo descargado: 100 MB en 3,50 s y 503 MB en 18,08 s; heap de JavaScript de 14,9 y 18,0 MiB respectivamente. El temporizador de la interfaz continuó ejecutándose durante las importaciones. Cancelación de otro XLSX de 503 MB: aprobada. Sin errores de JavaScript. Estos valores de heap no representan la memoria total de Chromium ni del servicio Python.

Validación final con el servidor actualizado: CSV 4,85 s / 14,9 MiB de heap; XLSX 23,08 s / 9,1 MiB de heap. Edición, descarga y cancelación vuelven a pasar. `.local-data/` quedó vacía al cerrar/cancelar los libros de prueba.

## Reproducir

Con `./abrir.sh` activo en el puerto 8768:

```bash
.venv/bin/python -m unittest -v large_store_test
NODE_PATH=/home/sergio_castro/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules node large.browser.test.cjs
.venv/bin/python large_benchmark.py generate --kind csv --mb 100
.venv/bin/python large_benchmark.py generate --kind xlsx --mb 500
.venv/bin/python large_benchmark.py run --kind csv --mb 100
.venv/bin/python large_benchmark.py run --kind xlsx --mb 500
NODE_PATH=/home/sergio_castro/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules node large-volume.browser.test.cjs
```

Pruebas pequeñas: 7 del modelo JS, 5 del almacenamiento Python, recorrido del editor pequeño tanto por archivo como por HTTP y recorrido del modo grande. Este último verifica referencias de fórmulas, páginas de filas/columnas, deshacer/rehacer, descarga, limpieza, rechazo de peticiones sin token/de otro origen y protección de `.git/` y `.local-data/`.
