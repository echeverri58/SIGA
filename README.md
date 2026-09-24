# SIGA · Formato Planilla de Asistencia

Aplicación web **100% estática** (HTML + CSS + JavaScript) para registrar la
asistencia de los participantes, capturar su **firma en pantalla** y descargar la
planilla oficial en **Excel**, sin alterar el formato original.

Todo el proceso ocurre en el navegador: **no hay servidor ni se suben datos a
internet**. La planilla se genera en el dispositivo y se descarga directamente.

## Características

- Formulario para los datos del participante: tipo de documento (CC, TI, CE, DNI),
  número, nombres, apellidos, dirección/dependencia/cargo, correo y teléfono.
- **Firma dibujada con el dedo** (o el mouse), con soporte táctil para celular.
- Lista de **varios participantes** por planilla (hasta 20, la capacidad de la tabla).
- Descarga la planilla en **Excel** o en **PDF** (elige el formato en el botón de abajo).
- En Excel:
  - Columna `No:` numerada automáticamente (1, 2, 3…).
  - La firma incrustada **dentro de la celda `FIRMA`**.
  - El **formato original intacto**: bordes, celdas combinadas, logo del SENA,
    casillas de selección, desplegables, orientación de impresión y pie de página.
- En PDF: se **convierte el propio Excel** a PDF con Microsoft Excel, por lo que
  queda **idéntico al Excel** (firma incluida dentro de la celda). Si no hay
  servidor/Excel disponible (por ejemplo en GitHub Pages), se usa un PDF
  aproximado generado en el navegador.
- Diseño **adaptable (responsive)** y pensado para usarse desde el celular.

## Archivos

| Archivo        | Descripción                                          |
|----------------|------------------------------------------------------|
| `index.html`   | Estructura de la aplicación.                         |
| `style.css`    | Estilos y diseño responsivo.                         |
| `app.js`       | Lógica del formulario, firma y generación (Excel/PDF). |
| `jszip.min.js` | Librería JSZip (lee y escribe el `.xlsx`).             |
| `jspdf.umd.min.js`              | Librería jsPDF (genera el PDF).          |
| `jspdf.plugin.autotable.min.js` | Tablas para jsPDF (planilla en PDF).     |
| `server.js`    | Mini-servidor local + conversión Excel→PDF (solo Node). |
| `convertir_server.ps1` | Convertidor Excel→PDF persistente (rápido).     |
| `convertir.ps1` | Conversión puntual de Excel a PDF (uso manual).        |
| `sena_logo.png` | Logo del SENA.                                        |
| `SUBIR_GITHUB.bat` | Sube los archivos al repositorio de GitHub.        |
| `Siga.xlsx`    | Plantilla oficial (solo lectura).                      |

> **PDF idéntico al Excel**: requiere ejecutar `node server.js` en una máquina
> con **Microsoft Excel instalado**. Al iniciar, el servidor abre Excel en segundo
> plano (`convertir_server.ps1`) y lo reutiliza, por lo que las conversiones son
> mucho más rápidas (la primera puede tardar unos segundos).

## Créditos

Aplicación creada por **John Alexander Echeverry Ocampo**
_Instructor del SENA · Politólogo · Analista de Datos_

- echeverri58@gmail.com
- jaecheverry@sena.edu.co

## Cómo alojarla en GitHub Pages

1. Crea un repositorio nuevo en GitHub.
2. Sube estos archivos a la rama principal (`main`):
   `index.html`, `style.css`, `app.js`, `jszip.min.js`, `Siga.xlsx` y `README.md`.
3. Ve a **Settings → Pages** y en *Source* elige **Deploy from a branch**,
   rama `main` y carpeta `/ (root)`. Guarda.
4. En unos segundos la app quedará publicada en:
   `https://<tu-usuario>.github.io/<nombre-del-repositorio>/`

> El archivo `Siga.xlsx` debe quedar en la **raíz** del repositorio (junto a
> `index.html`), porque la app lo lee desde esa ruta.

> ⚠️ **PDF en GitHub Pages:** GitHub solo sirve páginas web, no ejecuta programas.
> Para que el PDF sea **idéntico al Excel** hace falta un conversor (Excel o
> LibreOffice). Por eso en GitHub el PDF se dibuja con una librería del navegador
> y queda *parecido*, pero no idéntico. Para el PDF idéntico usa la app local
> (`node server.js`) o comparte la app con **`COMPARTIR.bat`** (abajo).

## Cómo probarla en tu computadora (sin Python)

Al ser estática, no funciona con doble clic (`file://`) porque el navegador no
permite leer el Excel localmente. Usa el mini-servidor incluido (solo requiere
Node.js, que ya tienes):

```bash
cd SIGA-Web
node server.js
```

Luego abre **http://localhost:8080** (o la dirección que muestra la consola).
Para usarla desde el celular en la misma red WiFi, abre la URL de red que imprime
la consola (por ejemplo `http://192.168.x.x:8080`).

Al iniciar verás `[Excel] listo para convertir a PDF`: desde ese momento el botón
**PDF** entrega el Excel convertido (idéntico). En cada conversión la consola
muestra cuánto tardó (`[PDF] convertido en X.X s`).

## Compartirla con otras personas (con PDF idéntico)

Ejecuta **`COMPARTIR.bat`**. Ese archivo:

1. Enciende el servidor local (que sí convierte con Excel).
2. Abre un túnel público temporal (Cloudflare Tunnel, gratis).
3. Te muestra un enlace tipo `https://algo-random.trycloudflare.com`.

Comparte ese enlace: quien lo abra usará exactamente la misma app, **con el PDF
idéntico al Excel**. Debes dejar la ventana abierta mientras la usan.

Opciones alternativas:

- **VS Code**: instala la extensión *Live Server* y pulsa "Go Live".
- **Node.js**: desde esta carpeta ejecuta `npx serve` o `npx http-server -p 8080`.

## Uso

1. Llena los datos del participante.
2. Dibuja la firma en el recuadro.
3. Pulsa **Agregar participante**.
4. Repite con cada persona.
5. En la barra de abajo elige **Excel** o **PDF** y pulsa **Descargar**.

## Notas

- La tabla admite **20 participantes por archivo** (filas 11 a 30 de la planilla).
  Para más asistentes, genera el archivo y crea otro.
- Los documentos y teléfonos se guardan como texto para conservar los ceros a la
  izquierda.
- La firma se exporta con fondo blanco para que se vea bien al imprimir.
