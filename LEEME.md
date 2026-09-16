# Auditor SO — instalación en el iPhone

Esta app funciona sin conexión una vez instalada, pero para poder
"Añadirla a pantalla de inicio" en iOS, Safari exige que se sirva
por HTTPS (no vale abrir el HTML suelto desde el propio teléfono).
La forma más sencilla y gratuita de conseguir eso es **GitHub Pages**.

## 1. Subir la carpeta a GitHub Pages (una sola vez)

1. Crea una cuenta gratuita en https://github.com si no tienes.
2. Crea un repositorio nuevo, por ejemplo `auditor-so`. Puede ser privado.
3. Sube **todos** los archivos de esta carpeta (`index.html`, `app.js`,
   `questions.js`, `manifest.json`, `sw.js`, `icon-192.png`,
   `icon-512.png` y la carpeta `vendor/` completa) a la raíz del
   repositorio — arrastrándolos desde la web de GitHub es suficiente,
   no hace falta usar la terminal.
4. Ve a **Settings → Pages**, en "Source" selecciona la rama
   `main` y carpeta `/ (root)`, y guarda.
5. En un par de minutos GitHub te dará una URL del tipo:
   `https://tu-usuario.github.io/auditor-so/`

## 2. Instalar en el iPhone

1. Abre esa URL con **Safari** (tiene que ser Safari, no Chrome).
2. Toca el icono de compartir (el cuadrado con la flecha hacia arriba).
3. Elige **"Añadir a pantalla de inicio"**.
4. Ya tienes el icono "Auditor SO" como una app más. Ábrela desde ahí
   (no desde Safari) para que funcione en modo app completo y offline.

## 3. Uso sin conexión

La primera vez que abras la app necesitas conexión para que se
descarguen y queden guardados en el teléfono los archivos (una sola
vez). A partir de ahí puedes auditar sin cobertura sin ningún
problema: todo se guarda en el propio teléfono. Solo necesitarás
conexión otra vez al final, para descargar los dos archivos
exportados (Excel + PDF) y subirlos a la plataforma de Repsol.

## 4. Si prefieres no usar GitHub

Cualquier otro hosting estático gratuito por HTTPS sirve igual
(Netlify, Vercel, Cloudflare Pages...); el procedimiento es el mismo:
subir la carpeta tal cual y abrir la URL resultante desde Safari.

---

**Importante:** cuando actualicemos la app (nuevas funciones o
ajustes), habrá que volver a subir los archivos actualizados al
mismo repositorio para que se refresque.
