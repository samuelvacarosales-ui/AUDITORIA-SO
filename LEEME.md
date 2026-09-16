# Auditor SO — instalación en el iPhone

Esta app funciona sin conexión una vez instalada, pero para poder
"Añadirla a pantalla de inicio" en iOS, Safari exige que se sirva
por HTTPS (no vale abrir el HTML suelto desde el propio teléfono).
La forma más sencilla y gratuita de conseguir eso es **GitHub Pages**.

## 1. Subir los archivos a GitHub Pages (una sola vez)

Todos los archivos de esta carpeta van **sueltos, todos al mismo
nivel** (no hay subcarpetas) — así es imposible que algo se quede a
medias al subirlos.

1. Crea una cuenta gratuita en https://github.com si no tienes.
2. Crea un repositorio nuevo, por ejemplo `auditor-so`. Puede ser privado.
3. Ve a **Add file → Upload files**.
4. Selecciona **todos** los archivos de esta carpeta a la vez (11 en
   total: `index.html`, `app.js`, `questions.js`, `manifest.json`,
   `sw.js`, `icon-192.png`, `icon-512.png`, `xlsx.full.min.js`,
   `jspdf.umd.min.js`, `jszip.min.js`, este mismo `LEEME.md`) y
   arrástralos todos juntos a la zona de subida.
5. Espera a que termine la barra de progreso (el archivo
   `xlsx.full.min.js` pesa casi 1MB, puede tardar un poco) y pulsa
   **Commit changes**.
6. Comprueba que ha funcionado: en la página principal del
   repositorio deberías ver el listado con los 11 archivos, todos
   sueltos, sin ninguna carpeta.
7. Ve a **Settings → Pages**, en "Source" elige la rama `main` y
   carpeta `/ (root)`, y guarda.
8. En un par de minutos tendrás tu URL:
   `https://tu-usuario.github.io/auditor-so/`

### Cómo comprobar que todo subió bien

Abre en Safari, una por una, estas tres URLs (cambiando por tu
usuario/repositorio):
- `https://tu-usuario.github.io/auditor-so/xlsx.full.min.js`
- `https://tu-usuario.github.io/auditor-so/jszip.min.js`
- `https://tu-usuario.github.io/auditor-so/jspdf.umd.min.js`

Si cada una te muestra una pantalla llena de código, están bien
subidas. Si alguna te da un error de página no encontrada, ese
archivo no llegó a subirse — vuelve al paso 3 y súbelo suelto de
nuevo (arrastrándolo junto a los demás, sobrescribirá el que falte).

## 2. Instalar en el iPhone

1. Abre esa URL con **Safari** (tiene que ser Safari, no Chrome).
2. Toca el icono de compartir (el cuadrado con la flecha hacia arriba).
3. Elige **"Añadir a pantalla de inicio"**.
4. Ya tienes el icono "Auditor SO" como una app más. Ábrela desde ahí
   (no desde Safari) para que funcione en modo app completo y offline.

Si ya la habías instalado antes con la versión que tenía la carpeta
`vendor`, bórrala de la pantalla de inicio y vuelve a añadirla desde
la URL, para que coja la versión nueva sin esa carpeta.

## 3. Uso sin conexión

La primera vez que abras la app necesitas conexión para que se
descarguen y queden guardados en el teléfono los archivos (una sola
vez). A partir de ahí puedes auditar sin cobertura sin ningún
problema: todo se guarda en el propio teléfono. Solo necesitarás
conexión otra vez al final, para descargar los dos archivos
exportados (Excel + PDF) y subirlos a la plataforma de Repsol.

## 4. Cuando actualicemos la app

Si en el futuro te paso archivos nuevos o corregidos, solo tienes
que repetir el paso "Add file → Upload files" arrastrando el archivo
actualizado (mismo nombre) — GitHub lo sustituye automáticamente.
