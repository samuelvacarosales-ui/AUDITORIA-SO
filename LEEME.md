# Encuesta SO — cómo actualizar tu GitHub

Usa el MISMO repositorio que ya tienes. Son 10 archivos, todos sueltos.
(jszip.min.js es nuevo: vuelve a estar porque ahora la app exporta el Excel.)

1. En tu repositorio: Add file → Upload files.
2. Arrastra los 10 archivos de esta carpeta a la vez. Los que ya existen se sustituyen.
3. Commit changes.
4. Espera 1-2 minutos. En el iPhone abre la app con internet y ciérrala del todo;
   al volver a abrirla ya carga la versión nueva.
   Si sigue la antigua: borra el icono y vuelve a añadirlo desde Safari.

## Flujo de trabajo
1. Descargas de Repsol el Excel de la estación, SIN rellenar.
2. En la app: «Cargar Excel».
3. En cada pregunta marcas SÍ / NO / N/A / PTE y, si hace falta, Valor/Fecha y Observaciones.
4. Fotos: «📷 Fotos de la estación» (las eliges de una vez) y «+» en cada pregunta.
5. «Guardar Excel»: sale el MISMO archivo de Repsol, con solo esas tres columnas rellenadas.
6. «Vista previa y PDF» → «Guardar PDF».
7. Subes a la plataforma el Excel y el PDF.

## Novedades (versión Encuesta SO)

- Al cargar el Excel limpio, todas las preguntas pasan a **SÍ** (verde).
- Por defecto: pregunta 13 → N/A "RECAUDACIÓN CON TABLET"; 37 → N/A "NO DISPONE"; 53 → N/A "NO APLICA".
- Interruptor **"Esta estación NO dispone de servicios"**: preguntas 6, 8, 56, 65 y 71 → N/A + "NO DISPONE". Al apagarlo vuelven a SÍ (si no las has cambiado a mano).
- Para que el móvil coja la versión nueva, sube de nuevo index.html, app.js, core.js, sw.js, manifest.json y LEEME.md.
- Preguntas 14–17 (inventarios): escribes Ajuste físico (€) y Ventas (€); la app calcula Valor = ajuste × 100 / ventas (con % y signo) y pone en observaciones "AJUSTE FISICO: x € VENTAS: y €". Sin importes queda "%" y "AJUSTE FISICO: € VENTAS: €".
