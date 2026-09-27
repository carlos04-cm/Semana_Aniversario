// ==========================================================================
// SEMANA DE ANIVERSARIO UNIREMINGTON 2026 - SEDE SAHAGÚN
// SCRIPT PRINCIPAL: VISTA RÁPIDA Y VISOR DE EXCEL INTERACTIVO
// ==========================================================================

// Elementos de la Ventana de Vista Rápida
const ventanaVistaRapida = document.getElementById("vista-rapida");
let eventoActualHoja = "Apertura";

// Elementos del Visor de Excel
const visorExcel = document.getElementById("visor-excel");
const cerrarVisorBtn = document.getElementById("cerrar-visor-excel");
const visorCuadricula = document.getElementById("visor-cuadricula");
const visorTabs = document.querySelectorAll(".visor-tab");
const visorCeldaActiva = document.getElementById("visor-celda-activa");
const visorContenidoFormula = document.getElementById("visor-contenido-formula");
const visorConteoRegistros = document.getElementById("visor-conteo-registros");
const tabPrevBtn = document.getElementById("tab-prev");
const tabNextBtn = document.getElementById("tab-next");
const tabsContainer = document.getElementById("visor-lista-pestanas");

// Quita espacios adicionales y saltos de línea del texto
function textoLimpio(elemento) {
    if (!elemento) return "";
    return elemento.textContent.replace(/\s+/g, " ").trim();
}

// --------------------------------------------------------------------------
// DATOS DEL ARCHIVO EVENTOS.XLSX (EXTRACCIÓN EXACTA DE LAS 7 HOJAS)
// --------------------------------------------------------------------------
const DATOS_EXCEL = {
    programacion: [
        {
            id: "01",
            evento: "Apertura Uniremington Fest",
            fecha: "21 de septiembre de 2026",
            hora: "2:00 p. m.",
            lugar: "Uniremington, sede Sahagún",
            descripcion: "Gran inauguración oficial de la Semana de Aniversario. Incluye palabras rectorales, presentaciones de un grupo de baile, actos de bienvenida, actos protocolarios, música en vivo y la proyección de una película.",
            participacion: "Opcional",
            hoja: "Apertura"
        },
        {
            id: "02",
            evento: "Talentos Uniremington",
            fecha: "22 de septiembre de 2026",
            hora: "2:00 p. m.",
            lugar: "Uniremington, sede Sahagún",
            descripcion: "Espacio de expresión artística donde estudiantes y docentes demuestran su talento a través de la interpretación de instrumentos, presentaciones de un grupo musical, baile y declamación de poemas.",
            participacion: "Obligatoria",
            hoja: "Talentos"
        },
        {
            id: "03",
            evento: "Juegos – Cultura General",
            fecha: "23 de septiembre de 2026",
            hora: "2:00 p. m.",
            lugar: "Uniremington, sede Sahagún",
            descripcion: "Jornada lúdica e interactiva con juegos de preguntas de cultura general en Kahoot y desafíos de lógica. Un espacio para poner a prueba los conocimientos, compartir con los compañeros y fortalecer la participación y el aprendizaje de forma divertida.",
            participacion: "Opcional",
            hoja: "Juegos"
        },
        {
            id: "04",
            evento: "Declamadores Invitados",
            fecha: "24 de septiembre de 2026",
            hora: "2:00 p. m.",
            lugar: "Uniremington, sede Sahagún",
            descripcion: "Recital poético y literario con la participación de destacados declamadores regionales, lecturas dramáticas y fomento a la creación lírica estudiantil. Incluye bailes al ritmo de música regional y presentaciones de canto y baile a cargo de instituciones invitadas.",
            participacion: "Opcional",
            hoja: "Declamadores"
        },
        {
            id: "05",
            evento: "Entrega de Distinciones",
            fecha: "25 de septiembre de 2026",
            hora: "2:00 p. m.",
            lugar: "Uniremington, sede Sahagún",
            descripcion: "Ceremonia solemne de reconocimiento al esfuerzo y la excelencia académica de los estudiantes, así como a la dedicación y el mérito de los docentes. Un espacio para celebrar sus logros, destacar su compromiso y valorar sus aportes al crecimiento de Uniremington y al fortalecimiento de nuestra comunidad universitaria.",
            participacion: "Obligatoria",
            hoja: "Distinciones"
        },
        {
            id: "06",
            evento: "Día de la Familia Uniremington y Feria Empresarial",
            fecha: "26 de septiembre de 2026",
            hora: "8:00 a. m.",
            lugar: "Uniremington, sede Sahagún",
            descripcion: "Jornada de integración comunitaria que reúne a estudiantes, docentes y personal administrativo. Cuenta con una muestra de emprendimientos universitarios y la presentación de un grupo vallenato.",
            participacion: "Opcional",
            hoja: "Familia"
        }
    ],

    eventos: {
        Apertura: {
            id: "01",
            titulo: "Apertura Uniremington Fest",
            fecha: "21 de septiembre de 2026",
            hora: "2:00 p. m.",
            lugar: "Uniremington, sede Sahagún",
            participacion: "Opcional",
            responsable: "Comité de Bienestar Institucional Uniremington",
            descripcion: "Gran inauguración oficial de la Semana de Aniversario. Incluye palabras rectorales, presentaciones de un grupo de baile, actos de bienvenida, actos protocolarios, música en vivo y la proyección de una película.",
            observaciones: "Asistir con documento de identidad y carné universitario. Actividad abierta para toda la comunidad académica de la sede.",
            fotos: {
                inicio: "imagenes/apertura_inicio.png",
                final: "imagenes/apertura_final.jpg"
            }
        },
        Talentos: {
            id: "02",
            titulo: "Talentos Uniremington",
            fecha: "22 de septiembre de 2026",
            hora: "2:00 p. m.",
            lugar: "Uniremington, sede Sahagún",
            participacion: "Obligatoria",
            responsable: "Coordinación Académica y Bienestar Universitario",
            descripcion: "Espacio de expresión artística donde estudiantes y docentes demuestran su talento a través de la interpretación de instrumentos, presentaciones de un grupo musical, baile y declamación de poemas.",
            observaciones: "Asistencia de carácter obligatorio para estudiantes y docentes según asignación académica.",
            fotos: {
                inicio: "imagenes/talentos_inicio.jpg",
                final: "imagenes/talentos_final.jpg"
            }
        },
        Juegos: {
            id: "03",
            titulo: "Juegos – Cultura General",
            fecha: "23 de septiembre de 2026",
            hora: "2:00 p. m.",
            lugar: "Uniremington, sede Sahagún",
            participacion: "Opcional",
            responsable: "Comité de Actividades Lúdicas y Docentes",
            descripcion: "Jornada lúdica e interactiva con juegos de preguntas de cultura general en Kahoot y desafíos de lógica. Un espacio para poner a prueba los conocimientos, compartir con los compañeros y fortalecer la participación y el aprendizaje de forma divertida.",
            observaciones: "Traer dispositivo móvil con acceso a internet o datos para participar en la plataforma Kahoot.",
            fotos: {
                inicio: "imagenes/juegos_inicio.jpg",
                final: "imagenes/juegos_final.jpg"
            }
        },
        Declamadores: {
            id: "04",
            titulo: "Declamadores Invitados",
            fecha: "24 de septiembre de 2026",
            hora: "2:00 p. m.",
            lugar: "Uniremington, sede Sahagún",
            participacion: "Opcional",
            responsable: "Área de Humanidades y Difusión Cultural",
            descripcion: "Recital poético y literario con la participación de destacados declamadores regionales, lecturas dramáticas y fomento a la creación lírica estudiantil. Incluye bailes al ritmo de música regional y presentaciones de canto y baile a cargo de instituciones invitadas.",
            observaciones: "Ingreso libre para estudiantes, egresados y familiares invitados.",
            fotos: {
                inicio: "imagenes/declamadores_inicio.jpg",
                final: "imagenes/declamadores_final.jpg"
            }
        },
        Distinciones: {
            id: "05",
            titulo: "Entrega de Distinciones",
            fecha: "25 de septiembre de 2026",
            hora: "2:00 p. m.",
            lugar: "Uniremington, sede Sahagún",
            participacion: "Obligatoria",
            responsable: "Rectoría y Consejo Académico Sede Sahagún",
            descripcion: "Ceremonia solemne de reconocimiento al esfuerzo y la excelencia académica de los estudiantes, así como a la dedicación y el mérito de los docentes. Un espacio para celebrar sus logros, destacar su compromiso y valorar sus aportes al crecimiento de Uniremington y al fortalecimiento de nuestra comunidad universitaria.",
            observaciones: "Código de vestimenta formal. Confirmar asistencia con antelación en la secretaría académica.",
            fotos: {
                inicio: "imagenes/distinciones_inicio.jpg",
                final: "imagenes/distinciones_final.jpg"
            }
        },
        Familia: {
            id: "06",
            titulo: "Día de la Familia Uniremington y Feria Empresarial",
            fecha: "26 de septiembre de 2026",
            hora: "8:00 a. m.",
            lugar: "Uniremington, sede Sahagún",
            participacion: "Opcional",
            responsable: "Comité de Emprendimiento y Bienestar",
            descripcion: "Jornada de integración comunitaria que reúne a estudiantes, docentes y personal administrativo. Cuenta con una muestra de emprendimientos universitarios y la presentación de un grupo vallenato.",
            observaciones: "Venir con familiares. Disponibilidad de stands de comidas, artesanías e ideas de negocio.",
            fotos: {
                inicio: "imagenes/familia_inicio.jpg",
                final: "imagenes/familia_final.jpg"
            }
        }
    }
};

// --------------------------------------------------------------------------
// RENDERIZADO DE LAS HOJAS DE CÁLCULO EN EL VISOR
// --------------------------------------------------------------------------

// Renderiza la hoja de Programación General
function renderizarHojaProgramacion() {
    visorCeldaActiva.textContent = "A2";
    visorContenidoFormula.textContent = "Semana de Aniversario Uniremington 2026 · Sede Sahagún";
    visorConteoRegistros.textContent = "Total eventos: 6 · Hoja: Programación";

    const filasHTML = DATOS_EXCEL.programacion.map((item, index) => {
        const filaNum = index + 6;
        const clasePart = item.participacion === "Obligatoria" ? "badge-obligatoria" : "badge-opcional";
        return `
            <tr>
                <td class="col-num">${filaNum}</td>
                <td class="celda-centro">${item.id}</td>
                <td class="celda-destacada">
                    <button type="button" class="link-hoja-directa" data-hoja="${item.hoja}" title="Ver hoja completa de ${item.evento}">
                        ${item.evento}
                    </button>
                </td>
                <td class="celda-centro">${item.fecha}</td>
                <td class="celda-centro">${item.hora}</td>
                <td>${item.lugar}</td>
                <td class="celda-texto">${item.descripcion}</td>
                <td class="celda-centro"><span class="visor-badge ${clasePart}">${item.participacion}</span></td>
                <td class="celda-centro">
                    <button type="button" class="btn-ir-hoja" data-hoja="${item.hoja}">
                        Ir a hoja ${item.hoja} ➔
                    </button>
                </td>
            </tr>
        `;
    }).join("");

    return `
        <div class="visor-tabla-excel-wrapper">
            <table class="visor-tabla-excel" aria-label="Hoja Programación">
                <thead>
                    <tr class="fila-columnas">
                        <th class="col-esquina"></th>
                        <th>A</th>
                        <th>B</th>
                        <th>C</th>
                        <th>D</th>
                        <th>E</th>
                        <th>F</th>
                        <th>G</th>
                        <th>H</th>
                    </tr>
                </thead>
                <tbody>
                    <tr class="fila-vacia">
                        <td class="col-num">1</td>
                        <td colspan="8"></td>
                    </tr>
                    <tr class="fila-banner-titulo">
                        <td class="col-num">2</td>
                        <td colspan="8" class="celda-titulo-principal">
                            Semana de Aniversario Uniremington 2026
                        </td>
                    </tr>
                    <tr class="fila-banner-subtitulo">
                        <td class="col-num">3</td>
                        <td colspan="8" class="celda-subtitulo-principal">
                            Sede Sahagún | Del 21 al 26 de septiembre de 2026
                        </td>
                    </tr>
                    <tr class="fila-vacia">
                        <td class="col-num">4</td>
                        <td colspan="8"></td>
                    </tr>
                    <tr class="fila-encabezados-tabla">
                        <td class="col-num">5</td>
                        <th class="th-excel">ID</th>
                        <th class="th-excel">Evento</th>
                        <th class="th-excel">Fecha</th>
                        <th class="th-excel">Hora</th>
                        <th class="th-excel">Lugar</th>
                        <th class="th-excel">Descripción</th>
                        <th class="th-excel">Participación</th>
                        <th class="th-excel">Hoja</th>
                    </tr>
                    ${filasHTML}
                    <tr class="fila-vacia">
                        <td class="col-num">12</td>
                        <td colspan="8"></td>
                    </tr>
                </tbody>
            </table>
        </div>
    `;
}

// Renderiza una hoja de evento individual (Apertura, Talentos, Juegos, etc.)
function renderizarHojaEvento(nombreHoja) {
    const evento = DATOS_EXCEL.eventos[nombreHoja];
    if (!evento) {
        return renderizarHojaProgramacion();
    }

    visorCeldaActiva.textContent = "A2";
    visorContenidoFormula.textContent = `${evento.titulo} (ID: ${evento.id})`;
    visorConteoRegistros.textContent = `Evento: ${evento.titulo} · Hoja: ${nombreHoja}`;

    const clasePart = evento.participacion === "Obligatoria" ? "badge-obligatoria" : "badge-opcional";

    // Manejo de fotografías con respaldo visual amigable si la imagen no está cargada
    const renderFoto = (ruta, titulo) => `
        <div class="visor-celda-foto">
            <div class="visor-foto-marco">
                <img src="${ruta}" 
                     alt="${titulo}" 
                     loading="lazy"
                     onerror="this.onerror=null; this.parentElement.innerHTML='<div class=\\'foto-no-disponible\\'>📷 ${titulo}<br><small>(Fotografía registrada en eventos.xlsx)</small></div>';">
            </div>
            <span class="visor-foto-pie">${titulo}</span>
        </div>
    `;

    return `
        <div class="visor-tabla-excel-wrapper">
            <table class="visor-tabla-excel" aria-label="Hoja ${nombreHoja}">
                <thead>
                    <tr class="fila-columnas">
                        <th class="col-esquina"></th>
                        <th>A</th>
                        <th>B</th>
                        <th>C</th>
                        <th>D</th>
                        <th>E</th>
                        <th>F</th>
                        <th>G</th>
                        <th>H</th>
                    </tr>
                </thead>
                <tbody>
                    <tr class="fila-vacia">
                        <td class="col-num">1</td>
                        <td colspan="8"></td>
                    </tr>
                    <tr class="fila-banner-titulo">
                        <td class="col-num">2</td>
                        <td colspan="8" class="celda-titulo-principal">
                            ${evento.titulo}
                        </td>
                    </tr>
                    <tr class="fila-banner-subtitulo">
                        <td class="col-num">3</td>
                        <td colspan="8" class="celda-subtitulo-principal">
                            Corporación Universitaria Remington · Sede Sahagún
                        </td>
                    </tr>
                    <tr class="fila-vacia">
                        <td class="col-num">4</td>
                        <td colspan="8"></td>
                    </tr>
                    <tr>
                        <td class="col-num">5</td>
                        <th colspan="2" class="th-campo">ID del Evento:</th>
                        <td colspan="6" class="celda-valor celda-destacada">#${evento.id}</td>
                    </tr>
                    <tr>
                        <td class="col-num">6</td>
                        <th colspan="2" class="th-campo">Fecha:</th>
                        <td colspan="6" class="celda-valor">${evento.fecha}</td>
                    </tr>
                    <tr>
                        <td class="col-num">7</td>
                        <th colspan="2" class="th-campo">Hora de inicio:</th>
                        <td colspan="6" class="celda-valor">${evento.hora}</td>
                    </tr>
                    <tr>
                        <td class="col-num">8</td>
                        <th colspan="2" class="th-campo">Lugar / Espacio:</th>
                        <td colspan="6" class="celda-valor">${evento.lugar}</td>
                    </tr>
                    <tr>
                        <td class="col-num">9</td>
                        <th colspan="2" class="th-campo">Participación:</th>
                        <td colspan="6" class="celda-valor">
                            <span class="visor-badge ${clasePart}">${evento.participacion}</span>
                        </td>
                    </tr>
                    <tr>
                        <td class="col-num">10</td>
                        <th colspan="2" class="th-campo">Responsable:</th>
                        <td colspan="6" class="celda-valor">${evento.responsable}</td>
                    </tr>
                    <tr class="fila-vacia">
                        <td class="col-num">11</td>
                        <td colspan="8"></td>
                    </tr>
                    <tr class="fila-seccion">
                        <td class="col-num">12</td>
                        <th colspan="8" class="th-seccion">Descripción detallada del evento</th>
                    </tr>
                    <tr>
                        <td class="col-num">13</td>
                        <td colspan="8" class="celda-descripcion-bloque">
                            <p>${evento.descripcion}</p>
                        </td>
                    </tr>
                    <tr class="fila-vacia">
                        <td class="col-num">17</td>
                        <td colspan="8"></td>
                    </tr>
                    <tr class="fila-seccion">
                        <td class="col-num">18</td>
                        <th colspan="4" class="th-seccion celda-centro">Fotografía del inicio</th>
                        <th colspan="4" class="th-seccion celda-centro">Fotografía del final</th>
                    </tr>
                    <tr>
                        <td class="col-num">19</td>
                        <td colspan="4" class="celda-fotos-col">
                            ${renderFoto(evento.fotos.inicio, "Fotografía del inicio")}
                        </td>
                        <td colspan="4" class="celda-fotos-col">
                            ${renderFoto(evento.fotos.final, "Fotografía del final")}
                        </td>
                    </tr>
                    <tr class="fila-vacia">
                        <td class="col-num">34</td>
                        <td colspan="8"></td>
                    </tr>
                    <tr>
                        <td class="col-num">35</td>
                        <th colspan="2" class="th-campo">Observaciones adicionales:</th>
                        <td colspan="6" class="celda-valor">${evento.observaciones}</td>
                    </tr>
                </tbody>
            </table>
        </div>
    `;
}

// --------------------------------------------------------------------------
// CONTROL DEL VISOR DE EXCEL
// --------------------------------------------------------------------------

// Cambia la pestaña activa y renderiza la hoja correspondiente
function activarPestañaHoja(nombreHoja) {
    if (!nombreHoja) nombreHoja = "Programación";

    // Actualiza el aspecto de las pestañas
    visorTabs.forEach((tab) => {
        const esActiva = tab.dataset.hoja.toLowerCase() === nombreHoja.toLowerCase();
        tab.setAttribute("aria-selected", esActiva ? "true" : "false");
        tab.classList.toggle("tab-activa", esActiva);
        if (esActiva) {
            tab.scrollIntoView({ behavior: "smooth", inline: "center", block: "nearest" });
        }
    });

    // Renderiza el contenido
    if (nombreHoja.toLowerCase() === "programación" || nombreHoja.toLowerCase() === "programacion") {
        visorCuadricula.innerHTML = renderizarHojaProgramacion();
    } else {
        visorCuadricula.innerHTML = renderizarHojaEvento(nombreHoja);
    }

    // Scroll superior en la cuadrícula
    visorCuadricula.scrollTop = 0;
    visorCuadricula.scrollLeft = 0;

    // Asocia los botones internos de enlace entre hojas
    visorCuadricula.querySelectorAll("[data-hoja]").forEach((btn) => {
        btn.addEventListener("click", () => {
            activarPestañaHoja(btn.dataset.hoja);
        });
    });
}

// Abre el visor de Excel en la hoja indicada
function abrirVisorExcel(nombreHoja) {
    if (ventanaVistaRapida && ventanaVistaRapida.open) {
        ventanaVistaRapida.close();
    }

    activarPestañaHoja(nombreHoja || "Programación");

    if (visorExcel) {
        visorExcel.showModal();
        document.body.classList.add("modal-abierto");
    }
}

// Cierra el visor de Excel
function cerrarVisorExcel() {
    if (visorExcel && visorExcel.open) {
        visorExcel.close();
        document.body.classList.remove("modal-abierto");
    }
}

// --------------------------------------------------------------------------
// EVENT LISTENERS E INTERACCIÓN
// --------------------------------------------------------------------------

// Cierre del Visor de Excel
if (cerrarVisorBtn) {
    cerrarVisorBtn.addEventListener("click", cerrarVisorExcel);
}

// Cerrar al hacer clic en el backdrop fuera del modal
if (visorExcel) {
    visorExcel.addEventListener("click", (evento) => {
        const rect = visorExcel.getBoundingClientRect();
        const clicEnBackdrop =
            evento.clientX < rect.left ||
            evento.clientX > rect.right ||
            evento.clientY < rect.top ||
            evento.clientY > rect.bottom;
        if (clicEnBackdrop) {
            cerrarVisorExcel();
        }
    });
}

// Botones de las pestañas en el visor
visorTabs.forEach((tab) => {
    tab.addEventListener("click", () => {
        activarPestañaHoja(tab.dataset.hoja);
    });
});

// Flechas para desplazar la barra de pestañas
if (tabPrevBtn && tabsContainer) {
    tabPrevBtn.addEventListener("click", () => {
        tabsContainer.scrollBy({ left: -150, behavior: "smooth" });
    });
}

if (tabNextBtn && tabsContainer) {
    tabNextBtn.addEventListener("click", () => {
        tabsContainer.scrollBy({ left: 150, behavior: "smooth" });
    });
}

// --------------------------------------------------------------------------
// APERTURA DE LA APLICACIÓN MICROSOFT EXCEL
// --------------------------------------------------------------------------
function mostrarNotificacionApertura(nombreHoja) {
    let toast = document.getElementById("toast-excel");
    if (!toast) {
        toast = document.createElement("div");
        toast.id = "toast-excel";
        toast.className = "toast-excel-notificacion";
        document.body.appendChild(toast);
    }
    toast.innerHTML = `
        <div class="toast-contenido">
            <span class="toast-icono">📊</span>
            <div>
                <strong>Abriendo en Microsoft Excel...</strong>
                <p>Hoja activa: <em>${nombreHoja}</em></p>
                <span class="toast-tip">💡 Si Excel muestra aviso de seguridad, pulsa <strong>«Sí»</strong> para ver el archivo.</span>
            </div>
        </div>
    `;
    toast.classList.add("visible");
    setTimeout(() => {
        toast.classList.remove("visible");
    }, 6000);
}

function abrirAplicacionExcel(nombreHoja) {
    const hojaKey = (nombreHoja || "Apertura").toLowerCase().trim();
    const archivoExcel = `eventos_${hojaKey}.xlsx`;

    // Obtener la URL completa del archivo actual
    let base = window.location.href;
    const ultimoSlash = base.lastIndexOf("/");
    if (ultimoSlash !== -1) {
        base = base.substring(0, ultimoSlash + 1);
    }
    const urlArchivo = base + archivoExcel;
    const urlProtocolo = `ms-excel:ofe|u|${urlArchivo}`;

    // Mostrar aviso visual informativo
    mostrarNotificacionApertura(nombreHoja);

    // Abre la aplicación Microsoft Excel directamente en el equipo
    try {
        const link = document.createElement("a");
        link.href = urlProtocolo;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    } catch (e) {
        window.location.href = urlProtocolo;
    }
}

// --------------------------------------------------------------------------
// INTERACCIÓN CON LAS TARJETAS DE EVENTO (FOTOS Y BOTÓN CONSULTAR EXCEL)
// --------------------------------------------------------------------------

// 1. Clic en las fotos de los eventos (Inicio o Final) -> Abre Microsoft Excel en la hoja del evento
document.querySelectorAll(".enlace-imagen").forEach((enlaceFoto) => {
    enlaceFoto.addEventListener("click", (e) => {
        e.preventDefault();
        const hoja = enlaceFoto.dataset.hoja || enlaceFoto.closest(".tarjeta")?.dataset.hoja || "Apertura";
        abrirAplicacionExcel(hoja);
    });
});

// 2. Clic en el botón "📊 Consultar Excel" de las tarjetas -> Abre Microsoft Excel en la hoja del evento
document.querySelectorAll(".btn-excel").forEach((botonExcel) => {
    if (botonExcel.id === "vista-excel") return;

    botonExcel.addEventListener("click", (e) => {
        e.preventDefault();
        const tarjeta = botonExcel.closest(".tarjeta");
        const hoja = botonExcel.dataset.hoja || tarjeta?.dataset.hoja || "Apertura";
        abrirAplicacionExcel(hoja);
    });
});

// 3. Ventana de Vista Rápida
document.querySelectorAll(".btn-vista").forEach((boton) => {
    boton.addEventListener("click", () => {
        const tarjeta = boton.closest(".tarjeta");
        eventoActualHoja = tarjeta.dataset.hoja || "Apertura";

        // Nombre del evento
        document.getElementById("titulo-vista").textContent =
            textoLimpio(tarjeta.querySelector("h3"));

        // Identificador
        document.getElementById("vista-id").textContent =
            "#" + tarjeta.dataset.id;

        // Fecha
        document.getElementById("vista-fecha").textContent =
            textoLimpio(tarjeta.querySelector("time"));

        // Hora
        document.getElementById("vista-hora").textContent =
            textoLimpio(tarjeta.querySelector(".hora-evento"));

        // Lugar
        document.getElementById("vista-lugar").textContent =
            textoLimpio(tarjeta.querySelector(".lugar-evento"));

        // Descripción
        document.getElementById("vista-descripcion").textContent =
            textoLimpio(tarjeta.querySelector(".descripcion-evento"));

        // Abre la ventana de vista rápida
        ventanaVistaRapida.showModal();
    });
});

// 4. Botón dentro de la ventana de Vista Rápida -> Abre Microsoft Excel
const btnVistaExcel = document.getElementById("vista-excel");
if (btnVistaExcel) {
    btnVistaExcel.addEventListener("click", (e) => {
        e.preventDefault();
        if (ventanaVistaRapida && ventanaVistaRapida.open) {
            ventanaVistaRapida.close();
        }
        abrirAplicacionExcel(eventoActualHoja);
    });
}

// 5. Soporte para cerrar con tecla Escape
document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && visorExcel && visorExcel.open) {
        cerrarVisorExcel();
    }
});
