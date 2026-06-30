// ============================================================
//  APPS SCRIPT - MARMOLERÍA GRILLO  (versión 3)
//  3 archivos separados + encabezado en fila 2 del Maestro
// ============================================================

// ── IDs DE LOS TRES ARCHIVOS ─────────────────────────────────
const ID_PRINCIPAL = "17F_pea3QV28Eu9HB7ULtQKjG3TQHX81KGf-yTM4pvXw"; // Copia ORDENES MARMOLERIA GRILLO
const ID_TALLER    = "1MHez_wVyiMlHiJ9pQfGM90xyQ79pDqUv-z4dFrKbPtg"; // PLANILLA DE PRODUCCION Y TALLER
const ID_LOGISTICA = "1yVttqklEcEtHs_8TgDPP5LF5VHiY5_N8A8M7BPeoY_8"; // PLANILLA DE LOGISTICA Y ENTREGAS

// ── NOMBRES DE LAS PESTAÑAS ───────────────────────────────────
const HOJA_MAESTRO   = "MAESTRO DE ORDENES";
const HOJA_TALLER    = "Hoja 1";
const HOJA_LOGISTICA = "Hoja 1";

// ── COLUMNAS DE ESCRITURA EN TALLER (1-based) ────────────────
// A=OT B=CLIENTE C=AMBIENTE D=PRODUCTO
// E=CORTADORA F=INICIO G=FIN
// H=TRASFORERA I=INICIO J=FIN
// K=PULIDORA L=INICIO M=FIN
// N=TERMINACION O=INICIO P=FIN
// Q=ESTADO
// R=OBS_CORTADORA S=OBS_TRASFORERA T=OBS_PULIDORA U=OBS_TERMINACION
// V=OBS_GENERAL  W=UNIDADES_HECHAS  ← nueva
const COL_TALLER = {
  cortadora:   { operario: 5,  inicio: 6,  fin: 7,  obs: 18 },
  trasforera:  { operario: 8,  inicio: 9,  fin: 10, obs: 19 },
  pulidora:    { operario: 11, inicio: 12, fin: 13, obs: 20 },
  terminacion: { operario: 14, inicio: 15, fin: 16, obs: 21 },
  estado:      { valor: 17 },
  general:     { obs: 22, unidadesHechas: 23 },
};

// ── COLUMNAS DE ESCRITURA EN LOGÍSTICA (1-based) ─────────────
const COL_LOG = {
  fechaProg:     11,
  equipo:        12,
  hsEstimado:    13,
  estado:        14,
  observaciones: 15,
};

// ============================================================
//  PUNTO DE ENTRADA GET
// ============================================================
function doGet(e) {
  try {
    const p      = e.parameter;
    const accion = p.accion;
    let resultado;

    if      (accion === "getOrdenes")   resultado = getOrdenes();
    else if (accion === "getTaller")    resultado = getTaller();
    else if (accion === "getLogistica") resultado = getLogistica();
    else if (accion === "getOrdenUrl")  resultado = getOrdenUrl(p.nombre);

    else if (accion === "actualizarTaller") {
      resultado = actualizarTaller({
        ot:      p.ot,
        maquina: p.maquina,
        campo:   p.campo,
        valor:   p.valor,
        fila:    p.fila || null,
      });
    }

    else if (accion === "actualizarLogistica") {
      resultado = actualizarLogistica({
        ot:    p.ot,
        campo: p.campo,
        valor: p.valor,
      });
    }

    else {
      resultado = { error: "Acción no reconocida: " + accion };
    }

    var callback = p.callback;
    if (callback) {
      return ContentService
        .createTextOutput(callback + "(" + JSON.stringify(resultado) + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    var callback = e.parameter.callback;
    var errObj = JSON.stringify({ error: err.message });
    if (callback) {
      return ContentService
        .createTextOutput(callback + "(" + errObj + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(errObj)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
//  PUNTO DE ENTRADA POST
// ============================================================
function doPost(e) {
  try {
    const datos  = JSON.parse(e.postData.contents);
    const accion = datos.accion;
    let resultado;
    if      (accion === "actualizarTaller")    resultado = actualizarTaller(datos);
    else if (accion === "actualizarLogistica") resultado = actualizarLogistica(datos);
    else resultado = { error: "Acción POST no reconocida: " + accion };
    return respuestaJSON(resultado);
  } catch (err) {
    return respuestaJSON({ error: err.message });
  }
}

// ============================================================
//  GET - MAESTRO DE ÓRDENES
// ============================================================
function getOrdenes() {
  const ss   = SpreadsheetApp.openById(ID_PRINCIPAL);
  const hoja = ss.getSheetByName(HOJA_MAESTRO);
  if (!hoja) return { ok: false, error: "No se encontró la hoja: " + HOJA_MAESTRO };

  const datos   = hoja.getDataRange().getValues();
  const ordenes = [];

  for (let i = 2; i < datos.length; i++) {
    const f  = datos[i];
    const ot = f[0];
    if (!ot) continue;
    ordenes.push({
      ot:          String(ot),
      ingreso:     fFecha(f[1]),
      cuil:        String(f[2]  || ""),
      cliente:     f[3]  || "",
      domicilio:   f[4]  || "",
      tel:         String(f[5]  || ""),
      vendedor:    f[6]  || "",
      sku:         f[7]  || "",
      producto:    f[8]  || "",
      categoria:   f[9]  || "",
      unidad:      f[10] || "",
      ambiente:    f[11] || "",
      largo:       f[12] || 0,
      ancho:       f[13] || 0,
      m2:          f[14] || 0,
      moneda:      f[15] || "",
      fechaEntrega:fFecha(f[26]),
      tipoEntrega: f[31] || "",
      estadoVentas:f[32] || "",
      observacion: f[33] || "",
      estadoFinal: f[34] || "",
      orden:       f[35] || "",
      obsVendedor: f[36] || "",
      fila:        i + 1,
    });
  }
  return { ok: true, datos: ordenes };
}

// ============================================================
//  GET - TALLER enriquecido con datos del Maestro
// ============================================================
function getTaller() {
  const ssTaller   = SpreadsheetApp.openById(ID_TALLER);
  const hojaTaller = ssTaller.getSheetByName(HOJA_TALLER);
  if (!hojaTaller) return { ok: false, error: "No se encontró hoja de taller: " + HOJA_TALLER };

  const ssMaestro    = SpreadsheetApp.openById(ID_PRINCIPAL);
  const hojaMaestro  = ssMaestro.getSheetByName(HOJA_MAESTRO);
  const datosMaestro = hojaMaestro ? hojaMaestro.getDataRange().getValues() : [];

  var maestroPorOT = {};
  for (var m = 2; m < datosMaestro.length; m++) {
    var fm = datosMaestro[m];
    if (!fm[0]) continue;
    var otM = String(fm[0]);
    if (!maestroPorOT[otM]) maestroPorOT[otM] = [];
    maestroPorOT[otM].push({
      cliente:     fm[3]  || "",
      domicilio:   fm[4]  || "",
      tel:         String(fm[5] || ""),
      vendedor:    fm[6]  || "",
      unidad:      fm[10] || "",
      largo:       fm[12] || 0,
      ancho:       fm[13] || 0,
      m2:          fm[14] || 0,
      fechaEntrega:fFecha(fm[26]),
      tipoEntrega: fm[31] || "",
      estadoFinal: fm[34] || "",
      orden:       fm[35] || "",
      obsVendedor: fm[36] || "",
    });
  }

  var contadorOT = {};

  const datosTaller = hojaTaller.getDataRange().getValues();
  const registros   = [];

  for (let i = 1; i < datosTaller.length; i++) {
    const f  = datosTaller[i];
    const ot = f[0];
    if (!ot) continue;

    const otStr = String(ot);
    if (!contadorOT[otStr]) contadorOT[otStr] = 0;
    const posicion = contadorOT[otStr];
    contadorOT[otStr]++;

    const filasM = maestroPorOT[otStr] || [];
    const e = filasM[posicion] || filasM[0] || {};

    registros.push({
      ot:                otStr,
      cliente:           f[1]  || e.cliente    || "",
      ambiente:          f[2]  || "",
      producto:          f[3]  || "",
      cortadora:         f[4]  || "",
      cortadoraInicio:   fHora(f[5]),
      cortadoraFin:      fHora(f[6]),
      trasforera:        f[7]  || "",
      trasforeraInicio:  fHora(f[8]),
      trasforeraFin:     fHora(f[9]),
      pulidora:          f[10] || "",
      pulidoraInicio:    fHora(f[11]),
      pulidoraFin:       fHora(f[12]),
      terminacion:       f[13] || "",
      terminacionInicio: fHora(f[14]),
      terminacionFin:    fHora(f[15]),
      estado:            f[16] || "",
      obsCortadora:      f[17] || "",
      obsTrasforera:     f[18] || "",
      obsPulidora:       f[19] || "",
      obsTerminacion:    f[20] || "",
      obsGeneral:        f[21] || "",
      unidadesHechas:    f[22] || 0,
      // Datos del Maestro por posición
      tel:         e.tel         || "",
      vendedor:    e.vendedor    || "",
      unidad:      e.unidad      || "",
      largo:       e.largo       || 0,
      ancho:       e.ancho       || 0,
      m2:          e.m2          || 0,
      fechaEntrega:e.fechaEntrega|| "",
      tipoEntrega: e.tipoEntrega || "",
      estadoFinal: e.estadoFinal || "",
      orden:       e.orden       || "",
      domicilio:   e.domicilio   || "",
      obsVendedor: e.obsVendedor || "",
      fila:        i + 1,
    });
  }
  return { ok: true, datos: registros };
}

// ============================================================
//  GET - LOGÍSTICA (archivo separado)
// ============================================================
function getLogistica() {
  const ss   = SpreadsheetApp.openById(ID_LOGISTICA);
  const hoja = ss.getSheetByName(HOJA_LOGISTICA);
  if (!hoja) return { ok: false, error: "No se encontró hoja de logística: " + HOJA_LOGISTICA };

  const datos     = hoja.getDataRange().getValues();
  const registros = [];

  for (let i = 1; i < datos.length; i++) {
    const f  = datos[i];
    const ot = f[0];
    if (!ot) continue;
    registros.push({
      ot:           String(ot),
      cliente:      f[1]  || "",
      tel:          String(f[2] || ""),
      tipoEntrega:  f[3]  || "",
      domicilio:    f[4]  || "",
      vendedor:     f[5]  || "",
      producto:     f[6]  || "",
      ambiente:     f[7]  || "",
      largo:        f[8]  || 0,
      ancho:        f[9]  || 0,
      fechaProg:    fFecha(f[10]),
      equipo:       f[11] || "",
      hsEstimado:   f[12] || "",
      estado:       f[13] || "",
      observaciones:f[14] || "",
      fila:         i + 1,
    });
  }
  return { ok: true, datos: registros };
}

// ============================================================
//  POST - ACTUALIZAR TALLER
// ============================================================
function actualizarTaller(datos) {
  const ss   = SpreadsheetApp.openById(ID_TALLER);
  const hoja = ss.getSheetByName(HOJA_TALLER);
  if (!hoja) return { ok: false, error: "No se encontró hoja de taller" };

  let filaNum;
  if (datos.fila && !isNaN(parseInt(datos.fila))) {
    filaNum = parseInt(datos.fila);
  } else {
    const filas = hoja.getDataRange().getValues();
    filaNum = buscarFila(filas, datos.ot);
  }
  if (!filaNum || filaNum === -1) return { ok: false, error: "OT no encontrada en taller: " + datos.ot };

  let columna;
  if (datos.maquina === "estado") {
    columna = COL_TALLER.estado.valor;
  } else if (datos.maquina === "general") {
    columna = COL_TALLER.general[datos.campo];
    if (!columna) return { ok: false, error: "Campo general no reconocido: " + datos.campo };
  } else {
    const maq = COL_TALLER[datos.maquina];
    if (!maq)     return { ok: false, error: "Máquina no reconocida: " + datos.maquina };
    columna = maq[datos.campo];
    if (!columna) return { ok: false, error: "Campo no reconocido: " + datos.campo };
  }

  var rango = hoja.getRange(filaNum, columna);
  rango.clearDataValidations();
  rango.setValue(datos.valor);
  return { ok: true, mensaje: "Taller OT-" + datos.ot + " fila " + filaNum + " actualizada" };
}

// ============================================================
//  POST - ACTUALIZAR LOGÍSTICA
// ============================================================
function actualizarLogistica(datos) {
  const ss   = SpreadsheetApp.openById(ID_LOGISTICA);
  const hoja = ss.getSheetByName(HOJA_LOGISTICA);
  if (!hoja) return { ok: false, error: "No se encontró hoja de logística" };

  const filas   = hoja.getDataRange().getValues();
  const filaNum = buscarFila(filas, datos.ot);
  if (filaNum === -1) return { ok: false, error: "OT no encontrada en logística: " + datos.ot };

  const columna = COL_LOG[datos.campo];
  if (!columna) return { ok: false, error: "Campo de logística no reconocido: " + datos.campo };

  var rango = hoja.getRange(filaNum, columna);
  rango.clearDataValidations();
  rango.setValue(datos.valor);

  if (datos.campo === "estado" && datos.valor === "ENTREGADO/FINALIZADO") {
    try {
      var ssMaestro = SpreadsheetApp.openById(ID_PRINCIPAL);
      var hojaMaestro = ssMaestro.getSheetByName(HOJA_MAESTRO);
      if (hojaMaestro) {
        var filasMaestro = hojaMaestro.getDataRange().getValues();
        for (var i = 2; i < filasMaestro.length; i++) {
          if (String(filasMaestro[i][0]) === String(datos.ot)) {
            var filaRealMaestro = i + 1;
            var rangoMaestro = hojaMaestro.getRange(filaRealMaestro, 35);
            rangoMaestro.clearDataValidations();
            rangoMaestro.setValue("ENTREGADO/FINALIZADO");
            Logger.log("Maestro actualizado: OT " + datos.ot + " fila " + filaRealMaestro);
            break;
          }
        }
      }
    } catch(e) {
      Logger.log("Error actualizando maestro: " + e.message);
    }
  }

  return { ok: true, mensaje: "Logística OT-" + datos.ot + " actualizada" };
}

// ============================================================
//  BUSCAR FOTO DE ORDEN EN DRIVE POR NOMBRE
// ============================================================
var CARPETA_ORDENES = "1IJLrkjko-hseJyppNs-tsvUKCwjGp8v6";

function getOrdenUrl(nombre) {
  if (!nombre) return { ok: false, error: "Sin nombre" };
  try {
    var carpeta = DriveApp.getFolderById(CARPETA_ORDENES);
    var archivos = carpeta.getFilesByName(nombre);
    if (archivos.hasNext()) {
      var archivo = archivos.next();
      return { ok: true, url: "https://drive.google.com/file/d/" + archivo.getId() + "/view" };
    }
    var nombreBase = nombre.split(".")[0];
    var todos = carpeta.getFiles();
    while (todos.hasNext()) {
      var f = todos.next();
      var fn = f.getName();
      if (fn.indexOf(nombreBase) !== -1 || nombreBase.indexOf(fn.split(".")[0]) !== -1) {
        return { ok: true, url: "https://drive.google.com/file/d/" + f.getId() + "/view" };
      }
    }
    return { ok: false, error: "No encontrado en carpeta: " + nombre };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ============================================================
//  UTILIDADES
// ============================================================
function buscarFila(filas, ot) {
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][0]) === String(ot)) return i + 1;
  }
  return -1;
}

function fFecha(valor) {
  if (!valor) return "";
  if (valor instanceof Date)
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
  return String(valor);
}

function fHora(valor) {
  if (!valor) return "";
  if (valor instanceof Date)
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "HH:mm");
  return String(valor);
}

function respuestaJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  CÓMO APLICAR ESTOS CAMBIOS
// ============================================================
/*
  1. Abrí la PLANILLA DE PRODUCCION TALLER (planilla separada de taller)
  2. Agregá el encabezado "UNIDADES_HECHAS" en la columna W (columna 23, fila 1)
  3. Abrí Copia ORDENES MARMOLERIA GRILLO → Extensiones → Apps Script
  4. Borrá todo y pegá este código
  5. Guardá con Ctrl+S
  6. Implementar → Administrar las implementaciones
  7. Lápiz ✏️ → Nueva versión → Implementar
  ✅ La URL no cambia
*/
