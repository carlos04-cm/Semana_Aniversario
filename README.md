# Semana de Aniversario Uniremington 2026

Proyecto académico de una página web local sobre la Semana de Aniversario de Uniremington, sede Sahagún, del 21 al 26 de septiembre de 2026.

La página presenta un resumen de cada actividad y fotografías del inicio y del final. Los archivos de Excel contienen la información detallada de los eventos.

## Objetivo

Integrar una interfaz desarrollada en HTML, CSS y JavaScript con archivos de Microsoft Excel para organizar y consultar la información de las actividades universitarias.

## Tecnologías

- **HTML:** estructura de la página y tarjetas de eventos.
- **CSS:** diseño, distribución y efectos visuales.
- **JavaScript:** vista rápida y solicitud de apertura de Microsoft Excel.
- **Excel (.xlsx):** información detallada y evidencias de los eventos.

## Archivos del proyecto

| Archivo o carpeta | Función |
| --- | --- |
| `index.html` | Página principal con las seis tarjetas. |
| `estilos.css` | Estilos de la página y ventanas de consulta. |
| `script.js` | Interacciones de las tarjetas y botones. |
| `eventos.xlsx` | Libro general con la programación y las hojas de eventos. |
| `eventos_apertura.xlsx` | Copia del libro guardada con Apertura como hoja activa. |
| `eventos_talentos.xlsx` | Copia del libro guardada con Talentos como hoja activa. |
| `eventos_juegos.xlsx` | Copia del libro guardada con Juegos como hoja activa. |
| `eventos_declamadores.xlsx` | Copia del libro guardada con Declamadores como hoja activa. |
| `eventos_distinciones.xlsx` | Copia del libro guardada con Distinciones como hoja activa. |
| `eventos_familia.xlsx` | Copia del libro guardada con Familia como hoja activa. |
| `imagenes/` | Logo y fotografías de inicio y final. |
| `README.md` | Instrucciones del proyecto. |

Todos los archivos de Excel contienen siete hojas: **Programación, Apertura, Talentos, Juegos, Declamadores, Distinciones y Familia**.

## Programación presentada en la página

| Fecha de 2026 | Evento | Hora | Participación |
| --- | --- | --- | --- |
| 21 de septiembre | Apertura Uniremington Fest | 2:00 p. m. | Opcional |
| 22 de septiembre | Talentos Uniremington | 2:00 p. m. | Obligatoria |
| 23 de septiembre | Juegos – Cultura General | 2:00 p. m. | Opcional |
| 24 de septiembre | Declamadores Invitados | 2:00 p. m. | Opcional |
| 25 de septiembre | Entrega de Distinciones | 2:00 p. m. | Obligatoria |
| 26 de septiembre | Día de la Familia Uniremington y Feria Empresarial | 8:00 a. m. | Opcional |

Lugar indicado: **Uniremington, sede Sahagún**.

## Cómo abrir el proyecto

1. Extraer el archivo ZIP completo.
2. Entrar en la carpeta `Semana_Aniversario`.
3. Abrir `index.html` en el navegador.
4. Pulsar **Explorar eventos** para ir a las tarjetas.
5. Usar **Vista Rápida** para consultar el resumen dentro de la página.
6. Pulsar una fotografía o **Consultar Excel** para solicitar la apertura del libro correspondiente en Microsoft Excel.

Conservar los archivos y la carpeta `imagenes` en sus ubicaciones originales. No abrir la página directamente dentro del ZIP.

## Cómo funciona la integración con Excel

El archivo `script.js` utiliza la función `abrirAplicacionExcel()` y el protocolo `ms-excel:ofe|u|` para solicitar la apertura de Microsoft Excel en el equipo.

Cada tarjeta apunta a una copia del libro con la hoja correspondiente seleccionada. Por ejemplo, Talentos utiliza `eventos_talentos.xlsx`, cuya hoja activa es **Talentos**.

La apertura automática requiere Microsoft Excel instalado y que el navegador y el sistema permitan utilizar ese protocolo. El mensaje «Abriendo en Microsoft Excel...» indica que se realizó la solicitud; no confirma que Excel haya abierto el archivo.

Si el navegador no abre Excel, se puede consultar el libro correspondiente directamente desde la carpeta del proyecto. La apertura mediante el botón debe comprobarse en el equipo donde se presentará la actividad.

## Vista rápida y visor incluido en el código

**Vista Rápida** toma el título, fecha, hora, lugar y descripción de la tarjeta y los muestra en una ventana dentro de la página. Su botón de Excel solicita la apertura de la aplicación de escritorio.

El proyecto también incluye código para un visor web con apariencia de hoja de cálculo. Sus datos están definidos manualmente en `DATOS_EXCEL`, dentro de `script.js`. Los botones actuales de las tarjetas están conectados a la aplicación Microsoft Excel, no a este visor interno.

La página no lee automáticamente el contenido de los archivos `.xlsx` ni sincroniza sus modificaciones. Los textos de HTML, los datos del visor y los libros de Excel se mantienen por separado.

## Nombres de las imágenes

Guardar los archivos dentro de `imagenes/` con estos nombres exactos:

| Uso | Inicio | Final |
| --- | --- | --- |
| Apertura | `apertura_inicio.png` | `apertura_final.jpg` |
| Talentos | `talentos_inicio.jpg` | `talentos_final.jpg` |
| Juegos | `juegos_inicio.jpg` | `juegos_final.jpg` |
| Declamadores | `declamadores_inicio.jpg` | `declamadores_final.jpg` |
| Distinciones | `distinciones_inicio.jpg` | `distinciones_final.jpg` |
| Familia | `familia_inicio.jpg` | `familia_final.jpg` |

Logo: `logo_uniremington.png`.

La extensión debe coincidir con la indicada en el HTML. Cambiar únicamente el nombre de una extensión no convierte el formato de la imagen.

## Cómo actualizar la información

- Editar `index.html` para cambiar los resúmenes, horarios y textos de las tarjetas.
- Editar `estilos.css` para cambiar colores, tamaños y distribución.
- Editar `script.js` para cambiar las interacciones o los datos del visor interno.
- Actualizar los libros de Excel cuando cambie la información detallada.
- Mantener las copias de Excel coherentes con el libro general y guardar cada copia con su hoja correspondiente activa.
- Reemplazar las fotografías conservando sus nombres, o actualizar las rutas del HTML si se cambian.

## Revisión antes de entregar

- Comprobar que cargan el logo y las doce fotografías.
- Revisar que las fotos de inicio y final corresponden a cada evento.
- Probar los seis botones de Vista Rápida.
- Probar las fotografías y los botones Consultar Excel en el equipo de presentación.
- Verificar que cada libro abre en la hoja del evento esperado.
- Revisar que Excel contiene los detalles y las evidencias solicitadas.
- Entregar la carpeta completa comprimida, incluyendo HTML, CSS, JavaScript, imágenes, libros de Excel y este README.

## Alcance

Esta versión es un sitio local estático. No utiliza servidor de aplicaciones, base de datos ni API externa. Su integración consiste en enlazar la presentación web con los libros de Excel mediante los controles de la página.
