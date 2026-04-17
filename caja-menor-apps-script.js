/**
 * ================================================================
 * SISTEMA DE CAJA MENOR - PUNTO MEDICAL
 * Google Apps Script
 * ================================================================
 */

const SPREADSHEET_ID = '1qmks4ElTbXzA6iIktibasX-vHQYtnyNnLfDQzrWCtb4';

const MESES = {
  1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL',
  5: 'MAYO', 6: 'JUNIO', 7: 'JULIO', 8: 'AGOSTO',
  9: 'SEPTIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'
};

// ==================== ENTRY POINTS ====================

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'consultarCajaMenor') {
      return consultarCajaMenor(e.parameter.mes, e.parameter.anio);
    }
    return jsonResponse({ success: false, message: 'Accion GET no valida: ' + action });
  } catch (error) {
    Logger.log('[doGet] Error: ' + error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ success: false, message: 'Request vacio o malformado' });
    }
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    Logger.log('[doPost] action=' + action);

    if (action === 'registrarCajaMenor') {
      return registrarCajaMenor(
        data.fecha, data.detalle, data.nit,
        data.proveedor, data.factura, data.valor, data.observaciones, data.esReembolso
      );
    }
    if (action === 'editarCajaMenor') {
      return editarCajaMenor(
        data.mes, data.anio, data.fila,
        data.detalle, data.valor, data.observaciones
      );
    }
    return jsonResponse({ success: false, message: 'Accion POST no valida: ' + action });
  } catch (error) {
    Logger.log('[doPost] Error: ' + error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

// ==================== REGISTRO ====================

function registrarCajaMenor(fecha, detalle, nit, proveedor, factura, valor, observaciones, esReembolso) {
  try {
    const { mes, anio } = parseFecha(fecha);
    Logger.log('[registrarCajaMenor] fecha=' + fecha + ' mes=' + mes + ' anio=' + anio + ' detalle=' + detalle + ' valor=' + valor);

    if (anio < 2025 || (anio === 2025 && mes < 11)) {
      return jsonResponse({ success: false, message: 'Solo se permite desde NOVIEMBRE 2025' });
    }

    if (!MESES[mes]) {
      return jsonResponse({ success: false, message: 'Mes invalido extraido de la fecha: ' + mes + ' (fecha raw: ' + fecha + ')' });
    }

    const nombreHoja = MESES[mes] + ' ' + anio;
    Logger.log('[registrarCajaMenor] Hoja destino resuelta: "' + nombreHoja + '"');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) {
      Logger.log('[registrarCajaMenor] Creando hoja: ' + nombreHoja);
      sheet = crearHojaCajaMenor(ss, nombreHoja);
    }

    // ── Buscar la primera fila libre desde la fila 10 en la columna C (col 3) ──
    const PRIMERA_FILA_DATOS = 10;
    const COL_FECHA = 3; // Columna C

    // Leer todos los valores de la columna C desde la fila 10 hasta el final de la hoja
    const ultimaFila = Math.max(sheet.getLastRow(), PRIMERA_FILA_DATOS - 1);
    let filaLibre = PRIMERA_FILA_DATOS;

    if (ultimaFila >= PRIMERA_FILA_DATOS) {
      const valoresColC = sheet.getRange(PRIMERA_FILA_DATOS, COL_FECHA, ultimaFila - PRIMERA_FILA_DATOS + 1, 1).getValues();
      filaLibre = PRIMERA_FILA_DATOS; // por defecto, si todo está vacío
      for (let i = 0; i < valoresColC.length; i++) {
        if (valoresColC[i][0] !== '' && valoresColC[i][0] !== null) {
          filaLibre = PRIMERA_FILA_DATOS + i + 1; // la siguiente a la última ocupada
        }
      }
    }

    Logger.log('[registrarCajaMenor] Escribiendo en fila: ' + filaLibre);

    // Escribir número correlativo en col B
    const numRegistro = filaLibre - PRIMERA_FILA_DATOS + 1;
    sheet.getRange(filaLibre, 2).setValue(numRegistro); // Col B: número

    // Escribir datos en columnas C a I
    sheet.getRange(filaLibre, COL_FECHA, 1, 7).setValues([[
      fecha,
      detalle || '',
      nit || '',
      proveedor || '',
      factura || '',
      Number(valor) || 0,
      observaciones || ''
    ]]);

    // Columna J: tipo de registro (GASTO / REEMBOLSO)
    var detalleUpper = String(detalle || '').trim().toUpperCase();
    var esReembolsoFinal = !!esReembolso || detalleUpper === 'REEMBOLSO';
    sheet.getRange(filaLibre, 10).setValue(esReembolsoFinal ? 'REEMBOLSO' : 'GASTO');

    // Formato de moneda en columna H (col 8)
    sheet.getRange(filaLibre, 8).setNumberFormat('$ #,##0.00');

    Logger.log('[registrarCajaMenor] Guardado exitosamente en fila ' + filaLibre);
    return jsonResponse({ success: true, message: 'Registro de Caja Menor guardado en fila ' + filaLibre });

  } catch (error) {
    Logger.log('[registrarCajaMenor] Error: ' + error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

// ==================== CONSULTA ====================

function consultarCajaMenor(mes, anio) {
  try {
    const nombreMes = resolverNombreMes(mes);
    if (!nombreMes) {
      return jsonResponse({ success: false, message: 'Mes invalido: ' + mes });
    }

    const nombreHoja = nombreMes + ' ' + anio;
    Logger.log('[consultarCajaMenor] Buscando hoja: ' + nombreHoja);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(nombreHoja);

    if (!sheet) {
      return jsonResponse({ success: false, message: 'No existe la hoja: ' + nombreHoja });
    }

    const datos = sheet.getDataRange().getValues();
    const registros = [];

    // Datos empiezan en fila 10 (índice 9), columnas C-I (índices 2-8)
    for (let i = 9; i < datos.length; i++) {
      if (!datos[i][2]) continue; // Saltar si col C (fecha) está vacía
      registros.push({
        fila: i + 1,           // Fila real en la hoja (1-indexed)
        numero: datos[i][1],  // Col B: número correlativo
        fecha: datos[i][2],  // Col C
        detalle: datos[i][3],  // Col D
        nit: datos[i][4],  // Col E
        proveedor: datos[i][5],  // Col F
        factura: datos[i][6],  // Col G
        valor: datos[i][7],  // Col H
        observaciones: datos[i][8],   // Col I
        tipoRegistro: (datos[i][9] || 'GASTO').toString().trim().toUpperCase() // Col J
      });
    }

    Logger.log('[consultarCajaMenor] Registros encontrados: ' + registros.length);

    // Total de gastos del mes: excluye reembolsos
    const totalCaja = registros
      .filter(function (r) { return (r.tipoRegistro || 'GASTO') !== 'REEMBOLSO'; })
      .reduce(function (s, r) { return s + (Number(r.valor) || 0); }, 0);

    Logger.log('[consultarCajaMenor] totalCaja (solo gastos)=' + totalCaja);
    return jsonResponse({ success: true, registros: registros, totalCaja: totalCaja });

  } catch (error) {
    Logger.log('[consultarCajaMenor] Error: ' + error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

// ==================== EDICIÓN ====================

function editarCajaMenor(mes, anio, fila, detalle, valor, observaciones) {
  try {
    const nombreMes = resolverNombreMes(mes);
    if (!nombreMes) return jsonResponse({ success: false, message: 'Mes invalido: ' + mes });

    const nombreHoja = nombreMes + ' ' + anio;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(nombreHoja);

    if (!sheet) return jsonResponse({ success: false, message: 'No existe la hoja: ' + nombreHoja });

    const filaNum = parseInt(fila);
    if (isNaN(filaNum) || filaNum < 10) {
      return jsonResponse({ success: false, message: 'Fila invalida: ' + fila });
    }

    sheet.getRange(filaNum, 4).setValue(detalle || '');        // Col D: detalle
    sheet.getRange(filaNum, 8).setValue(Number(valor) || 0);  // Col H: valor
    sheet.getRange(filaNum, 8).setNumberFormat('$ #,##0.00');
    sheet.getRange(filaNum, 9).setValue(observaciones || ''); // Col I: observaciones

    Logger.log('[editarCajaMenor] Fila ' + filaNum + ' actualizada en ' + nombreHoja);
    return jsonResponse({ success: true, message: 'Registro actualizado correctamente' });

  } catch (error) {
    Logger.log('[editarCajaMenor] Error: ' + error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

// ==================== AUXILIARES ====================

function resolverNombreMes(mes) {
  const mesUpper = mes.toString().toUpperCase().trim();
  if (Object.values(MESES).includes(mesUpper)) return mesUpper;
  const mesNum = parseInt(mes);
  if (!isNaN(mesNum) && MESES[mesNum]) return MESES[mesNum];
  return null;
}

/**
 * Parsea la fecha de forma robusta.
 *
 * Acepta los formatos más comunes que puede enviar el frontend o Google:
 *   • YYYY-MM-DD  (input type="date" estándar)
 *   • DD/MM/YYYY  (formato colombiano)
 *   • DD-MM-YYYY
 *   • Objeto Date de JavaScript (si Apps Script lo deserializa así)
 *
 * IMPORTANTE: NO usamos `new Date(fechaString)` porque el constructor de Date
 * en Apps Script/V8 interpreta "YYYY-MM-DD" como UTC medianoche, y según la
 * zona horaria del script (ej. America/Bogota, UTC-5) puede retroceder un día
 * o cambiar el mes. En su lugar, parseamos los componentes manualmente.
 */
function parseFecha(fecha) {
  // Si ya es un objeto Date (Apps Script a veces deserializa así)
  if (fecha instanceof Date) {
    return {
      anio: fecha.getFullYear(),
      mes: fecha.getMonth() + 1,
      dia: fecha.getDate()
    };
  }

  const str = fecha.toString().trim();
  Logger.log('[parseFecha] Parseando fecha raw: "' + str + '"');

  // Formato YYYY-MM-DD (el más común desde input type="date")
  const matchISO = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (matchISO) {
    return {
      anio: parseInt(matchISO[1]),
      mes: parseInt(matchISO[2]),
      dia: parseInt(matchISO[3])
    };
  }

  // Formato DD/MM/YYYY o DD-MM-YYYY (formato colombiano)
  const matchCOL = str.match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})/);
  if (matchCOL) {
    return {
      anio: parseInt(matchCOL[3]),
      mes: parseInt(matchCOL[2]),
      dia: parseInt(matchCOL[1])
    };
  }

  // Último recurso: intentar con Date pero usando zona horaria local del script
  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    // Usamos la zona horaria del spreadsheet para evitar desfase UTC
    const tz = SpreadsheetApp.openById(SPREADSHEET_ID).getSpreadsheetTimeZone();
    const partes = Utilities.formatDate(d, tz, 'yyyy-MM-dd').split('-');
    Logger.log('[parseFecha] Fallback Date usado, tz=' + tz + ' resultado=' + partes);
    return {
      anio: parseInt(partes[0]),
      mes: parseInt(partes[1]),
      dia: parseInt(partes[2])
    };
  }

  throw new Error('Formato de fecha no reconocido: "' + str + '"');
}

function crearHojaCajaMenor(ss, nombreHoja) {
  const sheet = ss.insertSheet(nombreHoja);

  // Encabezados en fila 9, columnas C-J
  const headers = [['FECHA', 'DETALLE', 'NIT', 'PROVEEDOR', 'N° FAC', 'VALOR', 'OBSERVACIONES', 'TIPO REGISTRO']];
  const anchos = [120, 180, 120, 180, 120, 120, 200, 140];

  // Encabezado de columna B
  sheet.getRange(9, 2).setValue('#').setFontWeight('bold').setBackground('#ffe599')
    .setBorder(true, true, true, true, true, true);
  sheet.setColumnWidth(2, 40);

  const headerRange = sheet.getRange(9, 3, 1, 8);
  headerRange.setValues(headers);
  headerRange.setFontWeight('bold').setBackground('#ffe599')
    .setBorder(true, true, true, true, true, true);

  anchos.forEach(function (ancho, i) { sheet.setColumnWidth(3 + i, ancho); });

  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== TEST MANUAL ====================
function testRegistrar() {
  const resultado = registrarCajaMenor(
    '2026-02-25', 'CAFETERIA', '123456',
    'Proveedor Test', 'FAC-001', 15000, 'Prueba manual'
  );
  Logger.log('Resultado: ' + resultado.getContent());
}