const SPREADSHEET_ID = "TU_SPREADSHEET_ID_AQUI";
const FOLDER_EVIDENCIAS_ID = "TU_CARPETA_EVIDENCIAS_ID_AQUI";
const FOLDER_PIEZAS_ID = "TU_CARPETA_PIEZAS_ID_AQUI";
const FOLDER_PDF_ID = "TU_CARPETA_PDF_ID_AQUI";

const ENCABEZADOS_REPORTES_CALIDAD = [
  "ID INCIDENCIA", "FECHA", "TICKET", "EQUIPO", "NS", "TECNICO", "CLIENTE", "DESCRIPCION",
  "EVIDENCIA", "CLASIFICACION", "ACCION CORRECTIVA (TECNICO)", "ACCION PREVENTIVA (TECNICO)", "¿REQUIERE ENVIO DE PIEZAS?",
  "¿CUAL PIEZA?", "NUMERO PARTE", "FOTO PIEZA", "DESTINO", "¿REQUIERE APOYO?", "INGENIERO QUE ATENDIO",
  "DESCRIPCION DEL APOYO BRINDADO", "AREA RESPONSABLE", "CAUSA RAIZ", "ACCION CORRECTIVA", "ACCION PREVENTIVA", "SUCURSAL"
];
const ENCABEZADOS_A_ELIMINAR_REPORTES_CALIDAD = ["SEMANA", "REPORTE", "FAMILIA"];
const PROPIEDAD_ULTIMO_ID_INCIDENCIA = "ULTIMO_ID_INCIDENCIA";

function normalizarEncabezado(encabezado) {
  return String(encabezado || "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function obtenerMapaColumnas(encabezados) {
  const columnas = {};
  encabezados.forEach((encabezado, indice) => {
    const llave = normalizarEncabezado(encabezado);
    if (!llave) return;
    if (columnas[llave]) throw new Error(`El encabezado "${encabezado}" está duplicado en Reportes_Calidad.`);
    columnas[llave] = indice + 1;
  });
  return columnas;
}

function obtenerColumnasReportesCalidad(sheet) {
  const encabezados = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const columnas = obtenerMapaColumnas(encabezados);
  const faltantes = ENCABEZADOS_REPORTES_CALIDAD.filter(encabezado => !columnas[normalizarEncabezado(encabezado)]);
  
  if (faltantes.length > 0) {
    throw new Error(`Faltan encabezados en Reportes_Calidad: ${faltantes.join(", ")}`);
  }
  return columnas;
}

function valorReporte(fila, columnas, encabezado) {
  return fila[columnas[normalizarEncabezado(encabezado)] - 1];
}

function asignarValorReporte(fila, columnas, encabezado, valor) {
  fila[columnas[normalizarEncabezado(encabezado)] - 1] = valor;
}

function generarIdIncidencia() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const propiedades = PropertiesService.getScriptProperties();
    const ultimoId = Number(propiedades.getProperty(PROPIEDAD_ULTIMO_ID_INCIDENCIA) || 999);
    const siguienteId = ultimoId + 1;
    propiedades.setProperty(PROPIEDAD_ULTIMO_ID_INCIDENCIA, String(siguienteId));
    return `INC-${siguienteId}`;
  } finally {
    lock.releaseLock();
  }
}

function migrarEstructuraReportesCalidad() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Reportes_Calidad");
  if (!sheet) throw new Error("No se encontró la hoja 'Reportes_Calidad'.");

  let encabezados = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  let columnas = obtenerMapaColumnas(encabezados);
  const idIncidencia = normalizarEncabezado("ID INCIDENCIA");

  if (!columnas[idIncidencia]) {
    const faltantes = ENCABEZADOS_A_ELIMINAR_REPORTES_CALIDAD.filter(encabezado => !columnas[normalizarEncabezado(encabezado)]);
    if (faltantes.length > 0) throw new Error(`No se puede migrar: faltan encabezados esperados: ${faltantes.join(", ")}`);

    sheet.insertColumnBefore(1);
    sheet.getRange(1, 1).setValue("ID INCIDENCIA");
    encabezados = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    columnas = obtenerMapaColumnas(encabezados);

    const columnasAEliminar = ENCABEZADOS_A_ELIMINAR_REPORTES_CALIDAD
      .map(encabezado => columnas[normalizarEncabezado(encabezado)])
      .sort((a, b) => b - a);
    columnasAEliminar.forEach(columna => sheet.deleteColumn(columna));
  }

  const columnasFinales = obtenerColumnasReportesCalidad(sheet);
  const cantidadRegistros = Math.max(sheet.getLastRow() - 1, 0);
  
  if (cantidadRegistros > 0) {
    const columnaId = columnasFinales[normalizarEncabezado("ID INCIDENCIA")];
    const idsActuales = sheet.getRange(2, columnaId, cantidadRegistros, 1).getValues();
    const filasSinId = idsActuales.map((fila, indice) => ({ fila: indice, valor: fila[0] })).filter(item => !item.valor);

    if (filasSinId.length > 0) {
      const lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try {
        const propiedades = PropertiesService.getScriptProperties();
        let ultimoId = Number(propiedades.getProperty(PROPIEDAD_ULTIMO_ID_INCIDENCIA) || 999);
        filasSinId.forEach(item => { idsActuales[item.fila][0] = `INC-${++ultimoId}`; });
        propiedades.setProperty(PROPIEDAD_ULTIMO_ID_INCIDENCIA, String(ultimoId));
        sheet.getRange(2, columnaId, cantidadRegistros, 1).setValues(idsActuales);
      } finally {
        lock.releaseLock();
      }
    }
  }
  return "Estructura migrada correctamente.";
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ Generar Reporte')
    .addItem('📊 Generar Reporte Histórico', 'lanzarPromptReporte')
    .addItem('🔄 Migrar estructura de Reportes_Calidad', 'migrarEstructuraReportesCalidad')
    .addToUi();
}

function doGet(e) {
  const idIncidencia = e.parameter.id || "";

  if (e.parameter.page === 'soporte') {
    const template = HtmlService.createTemplateFromFile('SoporteUI');
    template.ticketPrecargado = "";
    template.idIncidencia = idIncidencia;
    let cerrado = false;
    let ingenieroCierre = "";
    
    if (idIncidencia) {
      const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Reportes_Calidad");
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        const columnas = obtenerColumnasReportesCalidad(sheet);
        for (let i = data.length - 1; i >= 1; i--) {
          if (String(valorReporte(data[i], columnas, "ID INCIDENCIA")) === String(idIncidencia)) {
            template.ticketPrecargado = String(valorReporte(data[i], columnas, "TICKET") || "");
            const valorU = valorReporte(data[i], columnas, "INGENIERO QUE ATENDIO");
            if (valorU && String(valorU).trim() !== "") {
              cerrado = true;
              ingenieroCierre = String(valorU);
            }
            break;
          }
        }
      }
    }
    template.yaCerrado = cerrado;
    template.ingenieroPrevio = ingenieroCierre;
    return template.evaluate().setTitle('Cierre de Soporte').addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  }

  if (e.parameter.page === 'calidad') {
    const template = HtmlService.createTemplateFromFile('CalidadUI');
    template.ticketPrecargado = "";
    template.idIncidencia = idIncidencia;
    let cerrado = false;
    let areaPrevia = "";
    
    if (idIncidencia) {
      const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Reportes_Calidad");
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        const columnas = obtenerColumnasReportesCalidad(sheet);
        for (let i = data.length - 1; i >= 1; i--) {
          if (String(valorReporte(data[i], columnas, "ID INCIDENCIA")) === String(idIncidencia)) {
            template.ticketPrecargado = String(valorReporte(data[i], columnas, "TICKET") || "");
            const valorW = valorReporte(data[i], columnas, "AREA RESPONSABLE");
            if (valorW && String(valorW).trim() !== "") {
              cerrado = true;
              areaPrevia = String(valorW);
            }
            break;
          }
        }
      }
    }
    template.yaCerrado = cerrado;
    template.areaPrevia = areaPrevia;
    return template.evaluate().setTitle('Dictamen de Calidad').addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  }
  
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Reporte de Incidencias en Campo').addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function guardarCierreCalidad(idIncidencia, area, causa, correctiva, preventiva) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Reportes_Calidad");
    if (!sheet) return { exito: false, mensaje: "Error interno: Hoja no encontrada." };
    
    const data = sheet.getDataRange().getValues();
    const columnas = obtenerColumnasReportesCalidad(sheet);
    let filaExacta = 0;
    let filaData = null;
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(valorReporte(data[i], columnas, "ID INCIDENCIA")) === String(idIncidencia)) {
        filaExacta = i + 1;
        filaData = data[i]; 
        break;
      }
    }
    
    if (filaExacta === 0) return { exito: false, mensaje: "No se encontró el reporte exacto." };
    
    const columnaArea = columnas[normalizarEncabezado("AREA RESPONSABLE")];
    if (sheet.getRange(filaExacta, columnaArea).getValue() && String(sheet.getRange(filaExacta, columnaArea).getValue()).trim() !== "") {
      return { exito: false, mensaje: "🔒 Este dictamen ya fue llenado..." };
    }
    
    const areaLimpia = sanitizarTexto(area);
    const causaLimpia = sanitizarTexto(causa);
    
    sheet.getRange(filaExacta, columnaArea).setValue(areaLimpia);
    sheet.getRange(filaExacta, columnas[normalizarEncabezado("CAUSA RAIZ")]).setValue(causaLimpia);
    sheet.getRange(filaExacta, columnas[normalizarEncabezado("ACCION CORRECTIVA")]).setValue(sanitizarTexto(correctiva));
    sheet.getRange(filaExacta, columnas[normalizarEncabezado("ACCION PREVENTIVA")]).setValue(sanitizarTexto(preventiva));

    const cliente = valorReporte(filaData, columnas, "CLIENTE") || "N/A";
    const equipo = valorReporte(filaData, columnas, "EQUIPO") || "N/A";
    const falla = valorReporte(filaData, columnas, "DESCRIPCION") || "N/A";
    const ticket = valorReporte(filaData, columnas, "TICKET") || "N/A";

    const htmlCorreo = `<div style="font-family:Arial,sans-serif;color:#333;max-width:600px;margin:0 auto;border:1px solid #CCC;padding:20px;"><h2 style="color:#E60404;font-size:18px;margin-top:0;">Notificación de incidente de calidad</h2><p style="font-size:14px;">Hola,</p><p style="font-size:14px;">Te informamos que se ha registrado un incidente de calidad en el cual tu área podría participar durante el proceso de análisis.</p><p style="font-size:14px;">Es posible que el área de Calidad solicite tu apoyo para recopilar información, analizar la situación y, si aplica, establecer acciones que ayuden a evitar que el incidente vuelva a presentarse.</p><div style="background-color:#F9F9F9;border-left:4px solid #E60404;padding:15px;margin:20px 0;"><h3 style="margin-top:0;font-size:15px;color:#333;">Detalles del incidente:</h3><p style="margin:5px 0;font-size:13px;"><strong>Ticket:</strong> #${ticket}</p><p style="margin:5px 0;font-size:13px;"><strong>Cliente:</strong> ${cliente}</p><p style="margin:5px 0;font-size:13px;"><strong>Equipo / Modelo:</strong> ${equipo}</p><p style="margin:5px 0;font-size:13px;"><strong>Falla reportada:</strong> ${falla}</p></div><p style="font-size:14px;">Agradecemos tu disposición y apoyo.</p><p style="font-size:14px;margin-bottom:0;">Saludos,<br><strong>Equipo de Calidad</strong></p></div>`;

    const dicCorreos = {
      "LOGISTICA": "logistica@empresa.com",
      "ALMACEN": "almacen@empresa.com",
      "PRODUCCION": "produccion@empresa.com",
      "NPI/INGENIERIA": "ingenieria@empresa.com",
      "CALIDAD": "calidad@empresa.com",
      "STP": "stp@empresa.com"
    };

    let destinos = [];
    for (let key in dicCorreos) {
      if (areaLimpia.includes(key)) destinos.push(dicCorreos[key]);
    }
    
    let destinosUnicos = [...new Set(destinos)].join(",");
    
    if (destinosUnicos !== "") {
      try { MailApp.sendEmail({ to: destinosUnicos, subject: `Notificación de incidente de calidad - Ticket #${ticket}`, htmlBody: htmlCorreo, name: "Control de Calidad" }); } catch (e) {}
    }

    return { exito: true, mensaje: "Dictamen de calidad guardado correctamente." };
  } catch (error) { return { exito: false, mensaje: "Error: " + error.toString() }; }
}

function guardarCierreSoporte(idIncidencia, ingeniero, resolucion) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Reportes_Calidad");
    if (!sheet) return { exito: false, mensaje: "Error interno: Hoja no encontrada." };
    
    const data = sheet.getDataRange().getValues();
    const columnas = obtenerColumnasReportesCalidad(sheet);
    let filaExacta = 0;
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(valorReporte(data[i], columnas, "ID INCIDENCIA")) === String(idIncidencia)) {
        filaExacta = i + 1; 
        break;
      }
    }
    
    if (filaExacta === 0) return { exito: false, mensaje: "No se encontró el reporte." };
    
    const filaData = data[filaExacta - 1];
    const ticket = valorReporte(filaData, columnas, "TICKET") || "N/A";
    const columnaIngeniero = columnas[normalizarEncabezado("INGENIERO QUE ATENDIO")];
    const valorU = sheet.getRange(filaExacta, columnaIngeniero).getValue();
    
    if (valorU && String(valorU).trim() !== "") return { exito: false, mensaje: "🔒 Este soporte ya fue cerrado..." };
    
    sheet.getRange(filaExacta, columnaIngeniero).setValue(sanitizarTexto(ingeniero));
    sheet.getRange(filaExacta, columnas[normalizarEncabezado("DESCRIPCION DEL APOYO BRINDADO")]).setValue(sanitizarTexto(resolucion));

    const html = `<div style="font-family:Arial,sans-serif;color:#333;max-width:600px;margin:0 auto;border:1px solid #CCC;"><div style="background-color:#393D42;color:white;padding:12px 20px;"><h2 style="margin:0;font-size:16px;">SOPORTE ATENDIDO</h2></div><div style="padding:20px;"><p style="font-size:14px;">Ingeniería ha cerrado el soporte del ticket <strong>#${ticket}</strong>.</p><p style="font-size:13px;margin:0 0 5px;"><strong>Ingeniero:</strong> ${sanitizarTexto(ingeniero)}</p><p style="font-size:13px;margin:15px 0 5px;"><strong>Resolución brindada:</strong></p><p style="font-size:13px;margin:0;padding:10px;background:#F9F9F9;border-left:3px solid #E60404;">${sanitizarTexto(resolucion)}</p></div></div>`;
    
    try { MailApp.sendEmail({ to: "soporte_tecnico@empresa.com", subject: `SOPORTE ATENDIDO - Ticket #${ticket}`, htmlBody: html, name: "Ingeniería" }); } catch (e) {}
    
    return { exito: true, mensaje: "Resolución de soporte guardada correctamente." };
  } catch (error) { return { exito: false, mensaje: "Error: " + error.toString() }; }
}

function guardarReporteCalidad(formulario) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const reportSheet = ss.getSheetByName("Reportes_Calidad"); 
    if (!reportSheet) return { exito: false, mensaje: "Error interno: Hoja no encontrada." };
    const columnasReportes = obtenerColumnasReportesCalidad(reportSheet);
    
    const ticketLimpio = sanitizarTexto(formulario.ticket);
    const clienteLimpio = sanitizarTexto(formulario.cliente);
    const equipoLimpio = sanitizarTexto(formulario.equipo);
    const tecnicoRealLimpio = sanitizarTexto(formulario.tecnicoReal);
    const sucursalLim = asignarSucursal(tecnicoRealLimpio); 
    const noSerieLimpio = sanitizarTexto(formulario.noSerie);
    const clasificacionLim = sanitizarTexto(formulario.clasificacion);
    const descripcionLimpia = sanitizarTexto(formulario.descripcion);
    const correctivaTecLim = sanitizarTexto(formulario.accionCorrectivaTec);
    let preventivaTecLim = sanitizarTexto(formulario.accionPreventivaTec) || "N/A";
    
    const reqPiezasLim = sanitizarTexto(formulario.reqPiezas);
    const cualPiezaLim = sanitizarTexto(formulario.cualPieza);
    const numParteLim = sanitizarTexto(formulario.numParte); 
    const destinoLim = sanitizarTexto(formulario.destinoPieza); 
    const reqApoyoLim = sanitizarTexto(formulario.reqApoyo);
    
    const hoy = new Date();
    const idIncidencia = generarIdIncidencia();
    
    let listaUrls = [];
    const evidencias = Array.isArray(formulario.archivosEvidencia) ? formulario.archivosEvidencia.slice(0, 5) : [];
    
    if (evidencias.length > 0) {
      const folder = DriveApp.getFolderById(FOLDER_EVIDENCIAS_ID);
      evidencias.forEach((item, k) => {
        const bytes = Utilities.base64Decode(item.base64.split(",")[1]);
        const ext = item.nombre.split('.').pop();
        listaUrls.push(folder.createFile(Utilities.newBlob(bytes, item.base64.substring(5, item.base64.indexOf(";base64,")), `${idIncidencia}_F${k + 1}.${ext}`)).getUrl());
      });
    }
    
    let listaUrlsPiezas = [];
    let listaIdsPiezas = [];
    const evidenciasPiezas = Array.isArray(formulario.fotosPiezas) ? formulario.fotosPiezas.slice(0, 5) : [];
    
    if (evidenciasPiezas.length > 0 && formulario.reqPiezas === "SI") {
      const folderPiezas = DriveApp.getFolderById(FOLDER_PIEZAS_ID);
      evidenciasPiezas.forEach((item, k) => {
        const bytes = Utilities.base64Decode(item.base64.split(",")[1]);
        const ext = item.nombre.split('.').pop();
        const archivo = folderPiezas.createFile(Utilities.newBlob(bytes, item.base64.substring(5, item.base64.indexOf(";base64,")), `${idIncidencia}_REFACCION_${k + 1}.${ext}`));
        listaUrlsPiezas.push(archivo.getUrl());
        listaIdsPiezas.push(archivo.getId());
      });
    }

    const stringUrlsFinal = listaUrls.length > 0 ? listaUrls.join(", ") : "N/A";
    let urlPiezaFinal = listaUrlsPiezas.length > 0 ? listaUrlsPiezas.join(", ") : "N/A";
    const fechaSolo = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    let piezaFinal = cualPiezaLim, destinoFinal = destinoLim, numParteFinal = numParteLim; 
    
    if (reqPiezasLim === "NO") { 
      piezaFinal = "N/A"; 
      destinoFinal = "N/A"; 
      numParteFinal = "N/A"; 
      urlPiezaFinal = "N/A"; 
    }
    
    let ingenieroFinal = "N/A", resolucionFinal = "N/A";
    
    if (reqApoyoLim === "SI") { 
      ingenieroFinal = ""; 
      resolucionFinal = ""; 
    }

    const nuevaFila = Array(reportSheet.getLastColumn()).fill("");
    asignarValorReporte(nuevaFila, columnasReportes, "ID INCIDENCIA", idIncidencia);
    asignarValorReporte(nuevaFila, columnasReportes, "FECHA", fechaSolo);
    asignarValorReporte(nuevaFila, columnasReportes, "TICKET", ticketLimpio);
    asignarValorReporte(nuevaFila, columnasReportes, "EQUIPO", equipoLimpio);
    asignarValorReporte(nuevaFila, columnasReportes, "NS", noSerieLimpio);
    asignarValorReporte(nuevaFila, columnasReportes, "TECNICO", tecnicoRealLimpio);
    asignarValorReporte(nuevaFila, columnasReportes, "CLIENTE", clienteLimpio);
    asignarValorReporte(nuevaFila, columnasReportes, "DESCRIPCION", descripcionLimpia);
    asignarValorReporte(nuevaFila, columnasReportes, "EVIDENCIA", stringUrlsFinal);
    asignarValorReporte(nuevaFila, columnasReportes, "CLASIFICACION", clasificacionLim);
    asignarValorReporte(nuevaFila, columnasReportes, "ACCION CORRECTIVA (TECNICO)", correctivaTecLim);
    asignarValorReporte(nuevaFila, columnasReportes, "ACCION PREVENTIVA (TECNICO)", preventivaTecLim);
    asignarValorReporte(nuevaFila, columnasReportes, "¿REQUIERE ENVIO DE PIEZAS?", reqPiezasLim);
    asignarValorReporte(nuevaFila, columnasReportes, "¿CUAL PIEZA?", piezaFinal);
    asignarValorReporte(nuevaFila, columnasReportes, "NUMERO PARTE", numParteFinal);
    asignarValorReporte(nuevaFila, columnasReportes, "FOTO PIEZA", urlPiezaFinal);
    asignarValorReporte(nuevaFila, columnasReportes, "DESTINO", destinoFinal);
    asignarValorReporte(nuevaFila, columnasReportes, "¿REQUIERE APOYO?", reqApoyoLim);
    asignarValorReporte(nuevaFila, columnasReportes, "INGENIERO QUE ATENDIO", ingenieroFinal);
    asignarValorReporte(nuevaFila, columnasReportes, "DESCRIPCION DEL APOYO BRINDADO", resolucionFinal);
    asignarValorReporte(nuevaFila, columnasReportes, "AREA RESPONSABLE", "");
    asignarValorReporte(nuevaFila, columnasReportes, "CAUSA RAIZ", "");
    asignarValorReporte(nuevaFila, columnasReportes, "ACCION CORRECTIVA", "");
    asignarValorReporte(nuevaFila, columnasReportes, "ACCION PREVENTIVA", "");
    asignarValorReporte(nuevaFila, columnasReportes, "SUCURSAL", sucursalLim);
    
    reportSheet.getRange(reportSheet.getLastRow() + 1, 1, 1, nuevaFila.length).setValues([nuevaFila]);
    
    const registrosSheet = ss.getSheetByName("Registros Zoho");
    if (registrosSheet) {
      const celdaTicket = registrosSheet.getRange("A:A").createTextFinder(ticketLimpio).matchEntireCell(true).findNext();
      if (celdaTicket) {
        const filaIdx = celdaTicket.getRow();
        registrosSheet.getRange(filaIdx, 3).setValue(tecnicoRealLimpio);
        if (noSerieLimpio) registrosSheet.getRange(filaIdx, 6).setValue(noSerieLimpio);
      } else {
        registrosSheet.appendRow([ticketLimpio, clienteLimpio, tecnicoRealLimpio, equipoLimpio, "Instalación y Capacitación//CREADO MANUALMENTE", noSerieLimpio]);
      }
    }
    
    const correoCalidadDictamen = "calidad_aprobaciones@empresa.com";
    const correoCalidadLectura = "calidad_notificaciones@empresa.com";
    const correoGarantias = "garantias@empresa.com";
    const correoIngenieria = "soporte_tecnico@empresa.com";
    
    const urlApp = ScriptApp.getService().getUrl();
    const linkBaseHtml = `<div style="font-family:Arial,sans-serif;color:#333;max-width:600px;margin:0 auto;border:1px solid #CCC;">`;
    const headerHtml = `</div><div style="padding:20px;"><table style="width:100%;border-collapse:collapse;margin-top:15px;"><tr><td style="padding:8px;border:1px solid #DDD;font-weight:bold;width:35%;font-size:13px;">ID de incidencia:</td><td style="padding:8px;border:1px solid #DDD;font-size:13px;">${idIncidencia}</td></tr><tr><td style="padding:8px;border:1px solid #DDD;font-weight:bold;width:35%;font-size:13px;">Ticket:</td><td style="padding:8px;border:1px solid #DDD;font-size:13px;">#${ticketLimpio}</td></tr><tr><td style="padding:8px;border:1px solid #DDD;font-weight:bold;font-size:13px;">Cliente:</td><td style="padding:8px;border:1px solid #DDD;font-size:13px;">${clienteLimpio}</td></tr><tr><td style="padding:8px;border:1px solid #DDD;font-weight:bold;font-size:13px;">Equipo:</td><td style="padding:8px;border:1px solid #DDD;font-size:13px;">${equipoLimpio}</td></tr><tr><td style="padding:8px;border:1px solid #DDD;font-weight:bold;font-size:13px;">NS:</td><td style="padding:8px;border:1px solid #DDD;font-size:13px;">${noSerieLimpio}</td></tr><tr><td style="padding:8px;border:1px solid #DDD;font-weight:bold;font-size:13px;">Técnico:</td><td style="padding:8px;border:1px solid #DDD;font-size:13px;">${tecnicoRealLimpio}</td></tr></table>`;

    const cuerpoCalidadBase = `${linkBaseHtml}<div style="background-color:#E60404;color:white;padding:12px 20px;"><h2 style="margin:0;font-size:16px;">NUEVO REPORTE EN CAMPO</h2>${headerHtml}<h3 style="font-size:14px;margin-top:20px;border-bottom:1px solid #DDD;padding-bottom:5px;">Detalles:</h3><p style="font-size:13px;margin:0 0 5px;"><strong>Falla:</strong></p><p style="font-size:13px;margin:0;padding:10px;background:#F9F9F9;border-left:3px solid #CCC;">${descripcionLimpia}</p><p style="font-size:13px;margin:15px 0 5px;"><strong>Correctiva:</strong></p><p style="font-size:13px;margin:0;padding:10px;background:#F9F9F9;border-left:3px solid #CCC;">${correctivaTecLim}</p><p style="font-size:13px;margin:15px 0 5px;"><strong>Preventiva:</strong></p><p style="font-size:13px;margin:0;padding:10px;background:#F9F9F9;border-left:3px solid #CCC;">${preventivaTecLim}</p>`;

    const htmlCalidadDictamen = `${cuerpoCalidadBase}<div style="text-align:center;margin:30px 0 10px;"><a href="${urlApp}?page=calidad&id=${encodeURIComponent(idIncidencia)}" style="background:#393D42;color:white;padding:12px 24px;text-decoration:none;font-weight:bold;border-radius:4px;display:inline-block;">Emitir Dictamen</a></div></div></div>`;
    const htmlCalidadLectura = `${cuerpoCalidadBase}</div></div>`;

    try { MailApp.sendEmail({ to: correoCalidadDictamen, subject: `NUEVO REPORTE EN CAMPO (Dictamen Pendiente) - Ticket #${ticketLimpio}`, htmlBody: htmlCalidadDictamen, name: "Control de Calidad" }); } catch (e) {}
    
    if (correoCalidadLectura.trim() !== "") {
      try { MailApp.sendEmail({ to: correoCalidadLectura, subject: `NUEVO REPORTE EN CAMPO (Notificación) - Ticket #${ticketLimpio}`, htmlBody: htmlCalidadLectura, name: "Control de Calidad" }); } catch (e) {}
    }

    if (reqPiezasLim === "SI") {
      let fotosHtml = "";
      if (listaIdsPiezas.length > 0) {
        listaIdsPiezas.forEach((id, idx) => fotosHtml += `<div style="margin-top:15px;text-align:center;background:#fff;padding:15px;border:1px solid #ddd;border-radius:6px;"><p style="margin:0 0 10px;font-size:13px;color:#666;font-weight:bold;">📸 Evidencia ${idx+1}</p><a href="${listaUrlsPiezas[idx]}" target="_blank"><img src="https://drive.google.com/thumbnail?id=${id}&sz=w200" style="max-width:150px;border-radius:4px;display:block;margin:0 auto;"></a></div>`);
      }
      const htmlGarantias = `${linkBaseHtml}<div style="background-color:#333;color:white;padding:12px 20px;"><h2 style="margin:0;font-size:16px;">SOLICITUD DE REFACCIÓN</h2></div><div style="padding:20px;"><p style="font-size:14px;">Técnico <strong>${tecnicoRealLimpio}</strong> solicita pieza.</p><div style="background:#F9F9F9;border-left:4px solid #E60404;padding:15px;margin:15px 0;"><p style="margin:0 0 5px;font-size:13px;color:#666;">Pieza:</p><p style="margin:0 0 10px;font-size:16px;font-weight:bold;color:#E60404;">${cualPiezaLim}</p><p style="margin:0 0 5px;font-size:13px;color:#666;">No. Parte:</p><p style="margin:0 0 10px;font-size:16px;font-weight:bold;color:#333;">${numParteFinal}</p><p style="margin:0 0 5px;font-size:13px;color:#666;">Destino:</p><p style="margin:0 0 15px;font-size:16px;font-weight:bold;color:#333;">${destinoLim}</p>${fotosHtml}</div>${headerHtml}</div></div>`;
      try { MailApp.sendEmail({ to: correoGarantias, subject: `SOLICITUD DE REFACCIÓN - Ticket #${ticketLimpio}`, htmlBody: htmlGarantias, name: "Garantías" }); } catch (e) {}
    }

    if (reqApoyoLim === "SI") {
      const htmlIngenieria = `${linkBaseHtml}<div style="background-color:#E60404;color:white;padding:12px 20px;"><h2 style="margin:0;font-size:16px;">SOPORTE EN CAMPO REQUERIDO</h2>${headerHtml}<h3 style="font-size:14px;margin-top:20px;border-bottom:1px solid #DDD;padding-bottom:5px;">Contexto:</h3><p style="font-size:13px;margin:0 0 5px;"><strong>Descripción:</strong></p><p style="font-size:13px;margin:0;padding:10px;background:#F9F9F9;border-left:3px solid #CCC;">${descripcionLimpia}</p><div style="text-align:center;margin:30px 0 10px;"><a href="${urlApp}?page=soporte&id=${encodeURIComponent(idIncidencia)}" style="background:#393D42;color:white;padding:12px 24px;text-decoration:none;font-weight:bold;border-radius:4px;display:inline-block;">Atender Soporte</a></div></div></div>`;
      try { MailApp.sendEmail({ to: correoIngenieria, subject: `SOPORTE TÉCNICO - Ticket #${ticketLimpio}`, htmlBody: htmlIngenieria, name: "Ingeniería" }); } catch (e) {}
    }

    return { exito: true, mensaje: `Reporte realizado con éxito. ID: ${idIncidencia}` };
  } catch (error) { return { exito: false, mensaje: "Error: " + error.toString() }; }
}

function consultarTicketEnZoho(numTicket) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Registros Zoho"); 
    if (!sheet) return { encontrado: false };
    
    const celdaTicket = sheet.getRange("A:A").createTextFinder(sanitizarTexto(numTicket)).matchEntireCell(true).findNext();
    if (celdaTicket) {
      const filaData = sheet.getRange(celdaTicket.getRow(), 1, 1, 6).getValues()[0];
      return { 
        encontrado: true, 
        ticket: filaData[0], 
        cliente: filaData[1] || "", 
        tecnicoZoho: filaData[2] ? String(filaData[2]).trim() : "", 
        equipo: filaData[3] ? String(filaData[3]).toUpperCase() : "", 
        noSerie: filaData[5] ? String(filaData[5]).trim() : "" 
      };
    }
    return { encontrado: false };
  } catch (error) { return { encontrado: false }; }
}

function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    const numTicket = String(datos.ticket).trim();
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Registros Zoho");
    
    if (sheet.getRange("A:A").createTextFinder(numTicket).matchEntireCell(true).findNext()) {
      return ContentService.createTextOutput("Ignorado: Existe.");
    }
    
    sheet.appendRow([numTicket, sanitizarTexto(datos.cliente), sanitizarTexto(datos.tecnico), sanitizarTexto(datos.equipo), datos.asunto]);
    return ContentService.createTextOutput("Éxito: Guardado.");
  } catch(error) { return ContentService.createTextOutput("Error: " + error.toString()); }
}

function lanzarPromptReporte() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.prompt('Generar Reporte Histórico', 'Ingrese el número de ticket:', ui.ButtonSet.OK_CANCEL);
  
  if (respuesta.getSelectedButton() == ui.Button.OK) {
    const tId = respuesta.getResponseText().trim();
    if (!tId || isNaN(tId)) return ui.alert('⚠️ Validación', 'Número inválido.', ui.ButtonSet.OK);
    generarPdfDeTicketHistorico(tId);
  }
}

function generarPdfDeTicketHistorico(ticketId) {
  const ui = SpreadsheetApp.getUi();
  const docApp = DocumentApp;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reportes_Calidad");
  
  if (!sheet) return ui.alert("Error", "Hoja 'Reportes_Calidad' no encontrada.", ui.ButtonSet.OK);
  
  const columnasReportes = obtenerColumnasReportesCalidad(sheet);
  const ticketBuscado = String(ticketId).trim();
  const registrosEncontrados = sheet.getDataRange().getValues().filter(row => {
    const ticket = valorReporte(row, columnasReportes, "TICKET");
    return ticket != null && String(ticket).trim() === ticketBuscado;
  });
  
  if (registrosEncontrados.length === 0) return ui.alert("Aviso", "Sin registros para ticket #" + ticketId, ui.ButtonSet.OK);
  
  let cliente = valorReporte(registrosEncontrados[0], columnasReportes, "CLIENTE") || "N/A";
  let equipo = valorReporte(registrosEncontrados[0], columnasReportes, "EQUIPO") || "N/A";
  let noSerie = "N/A";
  const hojaZoho = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registros Zoho");
  
  if (hojaZoho) {
    const celdaTicket = hojaZoho.getRange("A:A").createTextFinder(ticketBuscado).matchEntireCell(true).findNext();
    if (celdaTicket) {
      const filaZoho = hojaZoho.getRange(celdaTicket.getRow(), 1, 1, 6).getValues()[0];
      if (filaZoho[1]) cliente = filaZoho[1];
      if (filaZoho[3]) equipo = filaZoho[3];
      if (filaZoho[5] && String(filaZoho[5]).trim() !== "") noSerie = String(filaZoho[5]).trim().replace(/\.0$/, "");
    }
  }

  const nombreDoc = "Reporte_" + ticketBuscado;
  const docTemp = docApp.create(nombreDoc);
  const body = docTemp.getBody().setMarginTop(40).setMarginBottom(40).setMarginLeft(40).setMarginRight(40);
  
  const tablaHeader = body.appendTable([["", ""]]);
  tablaHeader.setBorderWidth(0); 
  tablaHeader.getRow(0).getCell(0).setWidth(350).clear().appendParagraph("NOMBRE EMPRESA").setSpacingAfter(0).editAsText().setFontSize(26).setFontFamily("Arial").setBold(true).setForegroundColor("#E60404");
  tablaHeader.getRow(0).getCell(0).appendParagraph("REPORTE DE INCIDENCIAS").setSpacingBefore(0).editAsText().setFontSize(11).setFontFamily("Arial").setBold(true).setForegroundColor("#555555");
  
  const pLogo = tablaHeader.getRow(0).getCell(1).clear().appendParagraph("").setAlignment(docApp.HorizontalAlignment.RIGHT);
  try { 
    pLogo.appendInlineImage(UrlFetchApp.fetch("https://drive.google.com/thumbnail?id=1s8sO5yMzxBL5r6lAN78pIs65c8pG8gxV&sz=w1000").getBlob()).setWidth(160).setHeight(48); 
  } catch (e) { 
    pLogo.appendText("[Logotipo de la Empresa]").editAsText().setFontSize(10).setForegroundColor("#393D42").setBold(true); 
  }

  body.appendHorizontalRule();
  body.appendParagraph(`Control de Incidencias vinculadas al Ticket #${ticketBuscado} | Emisión: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy")}`).setSpacingBefore(6).setSpacingAfter(14).editAsText().setFontSize(9.5).setFontFamily("Arial").setItalic(true).setForegroundColor("#393D42");
  
  body.appendParagraph("DATOS GENERALES DEL EQUIPO").editAsText().setBold(true).setFontSize(10).setFontFamily("Arial").setForegroundColor("#393D42");
  const tablaFicha = body.appendTable([["Cliente:", cliente], ["Equipo:", equipo], ["NS:", noSerie]]);
  tablaFicha.setBorderColor("#D1D5DB").setBorderWidth(1);
  
  for (let r = 0; r < 3; r++) {
    tablaFicha.getRow(r).getCell(0).setAttributes({ [docApp.Attribute.BACKGROUND_COLOR]: "#F3F4F6", [docApp.Attribute.BOLD]: true, [docApp.Attribute.FONT_SIZE]: 9.5, [docApp.Attribute.FOREGROUND_COLOR]: "#393D42" }).setWidth(140);
    tablaFicha.getRow(r).getCell(1).editAsText().setFontSize(9.5).setFontFamily("Arial").setForegroundColor("#111827");
  }

  body.appendHorizontalRule();
  body.appendParagraph("\nREGISTRO DE INCIDENCIAS").editAsText().setBold(true).setFontSize(11).setFontFamily("Arial").setForegroundColor("#393D42");
  
  registrosEncontrados.forEach(row => {
    const fecha = valorReporte(row, columnasReportes, "FECHA");
    const txtFecha = (fecha && !isNaN(new Date(fecha).getTime())) ? Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), "dd/MM/yyyy") : "Sin fecha";
    const tablaBanner = body.appendTable([["Fecha: " + txtFecha, "👤 Técnico: " + valorReporte(row, columnasReportes, "TECNICO")]]);
    tablaBanner.setBorderWidth(0);
    
    const estiloBanner = { [docApp.Attribute.BACKGROUND_COLOR]: "#E60404", [docApp.Attribute.FOREGROUND_COLOR]: "#FFFFFF", [docApp.Attribute.BOLD]: true, [docApp.Attribute.FONT_SIZE]: 10 };
    tablaBanner.getRow(0).getCell(0).setAttributes(estiloBanner);
    tablaBanner.getRow(0).getCell(1).setAttributes(estiloBanner).getChild(0).asParagraph().setAlignment(docApp.HorizontalAlignment.RIGHT);
    
    let filasDetalle = [
      ["Falla:", valorReporte(row, columnasReportes, "DESCRIPCION") || "N/A"],
      ["Correctiva:", valorReporte(row, columnasReportes, "ACCION CORRECTIVA (TECNICO)") || "N/A"],
      ["Preventiva:", valorReporte(row, columnasReportes, "ACCION PREVENTIVA (TECNICO)") || "N/A"],
      ["¿Envío piezas?:", valorReporte(row, columnasReportes, "¿REQUIERE ENVIO DE PIEZAS?") || "N/A"]
    ];
    
    if (String(valorReporte(row, columnasReportes, "¿REQUIERE ENVIO DE PIEZAS?")).toUpperCase() === "SI") {
      filasDetalle.push(["Pieza:", valorReporte(row, columnasReportes, "¿CUAL PIEZA?") || "N/A"]);
      filasDetalle.push(["No. Parte:", valorReporte(row, columnasReportes, "NUMERO PARTE") || "N/A"]);
    }
    filasDetalle.push(["¿Apoyo?:", valorReporte(row, columnasReportes, "¿REQUIERE APOYO?") || "N/A"]);
    
    const tablaInfo = body.appendTable(filasDetalle);
    tablaInfo.setBorderColor("#D1D5DB").setBorderWidth(1);
    
    for (let i = 0; i < filasDetalle.length; i++) {
      tablaInfo.getRow(i).getCell(0).setAttributes({ [docApp.Attribute.BACKGROUND_COLOR]: "#F9FAFB", [docApp.Attribute.BOLD]: true, [docApp.Attribute.FOREGROUND_COLOR]: "#393D42", [docApp.Attribute.FONT_SIZE]: 9 }).setWidth(160);
      tablaInfo.getRow(i).getCell(1).editAsText().setFontSize(9).setFontFamily("Arial").setForegroundColor("#4B5563");
    }
    
    const evidencia = valorReporte(row, columnasReportes, "EVIDENCIA");
    if (evidencia && String(evidencia).trim() !== "N/A") {
      const urlsValidas = String(evidencia).split(",").filter(u => u.trim() !== "");
      if (urlsValidas.length > 0) {
        body.appendParagraph("Evidencias Fotográficas:").setSpacingBefore(10).setSpacingAfter(4).editAsText().setBold(true).setFontSize(9).setFontFamily("Arial").setForegroundColor("#393D42");
        const tablaFotos = body.appendTable([["", "", ""], ["", "", ""]]).setBorderWidth(0);
        
        urlsValidas.forEach((url, idx) => {
          if (idx >= 5) return;
          const filaFoto = Math.floor(idx / 3);
          const columnaFoto = idx % 3;
          const idMatch = url.match(/[-\w]{25,}/);
          if (idMatch) {
            try {
              const celdaFoto = tablaFotos.getRow(filaFoto).getCell(columnaFoto);
              const imgDoc = celdaFoto.getChild(0).asParagraph().setAlignment(docApp.HorizontalAlignment.LEFT).appendInlineImage(DriveApp.getFileById(idMatch[0]).getBlob());
              const originalW = imgDoc.getWidth() || 1; 
              const originalH = imgDoc.getHeight() || 1;
              imgDoc.setWidth(135).setHeight(Math.round(originalH * (135 / originalW)));
            } catch(e) { 
              tablaFotos.getRow(filaFoto).getCell(columnaFoto).setText("[Fallo]"); 
            }
          }
        });
      }
    }
    body.appendParagraph(" ").setSpacingAfter(15);
  });
  
  docTemp.saveAndClose();
  const pdfId = DriveApp.getFolderById(FOLDER_PDF_ID).createFile(docTemp.getAs('application/pdf').setName(nombreDoc + ".pdf")).getId();
  DriveApp.getFileById(docTemp.getId()).setTrashed(true);
  
  const plantillaHtml = HtmlService.createTemplateFromFile('DescargadorUI');
  plantillaHtml.url = "https://drive.google.com/uc?export=download&id=" + pdfId;
  ui.showModalDialog(plantillaHtml.evaluate().setWidth(250).setHeight(100), "Procesando");
}

function sanitizarTexto(texto) {
  if (!texto) return "";
  return String(texto).toUpperCase()
    .replace(/[ÁÀÄÂ]/g, "A").replace(/[ÉÈËÊ]/g, "E").replace(/[ÍÌÏÎ]/g, "I").replace(/[ÓÒÖÔ]/g, "O").replace(/[ÚÙÜÛ]/g, "U")
    .replace(/&nbsp;/g, " ").replace(/\s+/g, " ").trim();
}

function asignarSucursal(tecnico) {
  const zonas = {
      "TECNICO 1": "CDMX", "TECNICO 2": "CDMX", "TECNICO 3": "CDMX", "TECNICO 4": "CDMX", "TECNICO 5": "CDMX", 
      "TECNICO 6": "CDMX", "TECNICO 7": "CDMX", "TECNICO 8": "CDMX", "TECNICO 9": "CDMX", "TECNICO 10": "CDMX", 
      "TECNICO 11": "MTY", "TECNICO 12": "MTY", "TECNICO 13": "MTY", "TECNICO 14": "MTY", "TECNICO 15": "MTY",
      "TECNICO 16": "GTO", "TECNICO 17": "OCOTLAN",
      "TECNICO 18": "GDL", "TECNICO 19": "GDL", "TECNICO 20": "GDL", "TECNICO 21": "GDL", "TECNICO 22": "GDL", 
      "TECNICO 23": "GDL", "TECNICO 24": "GDL", "TECNICO 25": "GDL", "TECNICO 26": "GDL", "TECNICO 27": "GDL", 
      "TECNICO 28": "GDL", "TECNICO 29": "GDL", "TECNICO 30": "GDL", "TECNICO 31": "GDL"
  };
  return zonas[tecnico] || "NO ENCONTRADO";
}

function buscarMisReportesFront(ticket) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Reportes_Calidad");
    if (!sheet) return { exito: false, mensaje: "Error interno: Hoja no encontrada." };
    
    const data = sheet.getDataRange().getValues();
    const columnas = obtenerColumnasReportesCalidad(sheet);
    const ticketBuscado = String(ticket).trim();
    let encontrados = [];

    for (let i = data.length - 1; i >= 1; i--) {
      if (String(valorReporte(data[i], columnas, "TICKET")) === ticketBuscado) {
        encontrados.push({
          fecha: valorReporte(data[i], columnas, "FECHA") ? Utilities.formatDate(new Date(valorReporte(data[i], columnas, "FECHA")), Session.getScriptTimeZone(), "dd/MM/yyyy") : "Sin fecha",
          tecnico: valorReporte(data[i], columnas, "TECNICO") || "N/A",
          cliente: valorReporte(data[i], columnas, "CLIENTE") || "N/A",
          equipo: valorReporte(data[i], columnas, "EQUIPO") || "N/A",
          ns: valorReporte(data[i], columnas, "NS") || "N/A",
          clasificacion: valorReporte(data[i], columnas, "CLASIFICACION") || "N/A",
          falla: valorReporte(data[i], columnas, "DESCRIPCION") || "N/A",
          correctiva: valorReporte(data[i], columnas, "ACCION CORRECTIVA (TECNICO)") || "N/A",
          preventiva: valorReporte(data[i], columnas, "ACCION PREVENTIVA (TECNICO)") || "N/A",
          fotos: valorReporte(data[i], columnas, "EVIDENCIA") || "N/A",
          piezas: valorReporte(data[i], columnas, "¿REQUIERE ENVIO DE PIEZAS?") || "NO",
          cualPieza: valorReporte(data[i], columnas, "¿CUAL PIEZA?") || "N/A",
          numParte: valorReporte(data[i], columnas, "NUMERO PARTE") || "N/A",
          destino: valorReporte(data[i], columnas, "DESTINO") || "N/A",
          fotosPiezas: valorReporte(data[i], columnas, "FOTO PIEZA") || "N/A",
          apoyo: valorReporte(data[i], columnas, "¿REQUIERE APOYO?") || "NO"
        });
      }
    }
    
    if (encontrados.length === 0) return { exito: false, mensaje: "No se encontraron reportes previos para el ticket #" + ticketBuscado };
    return { exito: true, datos: encontrados };
  } catch (error) { return { exito: false, mensaje: "Error: " + error.toString() }; }
}