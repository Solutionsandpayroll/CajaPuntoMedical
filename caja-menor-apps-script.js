/**
 * ================================================================
 * SISTEMA DE CAJA MENOR - PUNTO MEDICAL
 * Google Apps Script — versión auditada y corregida
 * ================================================================
 */

const SPREADSHEET_ID = '1qmks4ElTbXzA6iIktibasX-vHQYtnyNnLfDQzrWCtb4';

const MESES = {
  1: 'ENERO',    2: 'FEBRERO',   3: 'MARZO',     4: 'ABRIL',
  5: 'MAYO',     6: 'JUNIO',     7: 'JULIO',      8: 'AGOSTO',
  9: 'SEPTIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'
};

// ==================== ENTRY POINTS ====================

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'consultarCajaMenor') {
      return consultarCajaMenor(e.parameter.mes, e.parameter.anio);
    }
    return jsonResponse({ success: false, message: 'Acción GET no válida' });
  } catch (error) {
    Logger.log('[doGet] Error: ' + error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

function doPost(e) {
  try {
    // Apps Script recibe text/plain y application/json igual en postData.contents
    const data   = JSON.parse(e.postData.contents);
    const action = data.action;
    if (action === 'registrarCajaMenor') {
      return registrarCajaMenor(
        data.fecha, data.detalle, data.nit,
        data.proveedor, data.factura, data.valor, data.observaciones
      );
    }
    return jsonResponse({ success: false, message: 'Acción POST no válida' });
  } catch (error) {
    Logger.log('[doPost] Error: ' + error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

// ==================== REGISTRO ====================

function registrarCajaMenor(fecha, detalle, nit, proveedor, factura, valor, observaciones) {
  try {
    const { mes, anio } = parseFecha(fecha);
    Logger.log('[registrarCajaMenor] fecha=%s mes=%s anio=%s detalle=%s valor=%s', fecha, mes, anio, detalle, valor);

    if (anio < 2025 || (anio === 2025 && mes < 11)) {
      Logger.log('[registrarCajaMenor] Fecha no permitida');
      return jsonResponse({ success: false, message: 'Solo se permite desde NOVIEMBRE 2025' });
    }

    const nombreHoja = `${MESES[mes]} ${anio}`;
    const ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
    let   sheet      = ss.getSheetByName(nombreHoja);
    if (!sheet) {
      Logger.log('[registrarCajaMenor] Creando hoja: ' + nombreHoja);
      sheet = crearHojaCajaMenor(ss, nombreHoja);
    }

    // Obtener la primera fila vacía de forma segura (sin bucle con getRange individual)
    const lastRow = sheet.getLastRow();
    const row     = Math.max(lastRow + 1, 10); // nunca antes de fila 10

    Logger.log('[registrarCajaMenor] Escribiendo en fila: ' + row);

    // Una sola llamada a la API en vez de 7 individuales
    sheet.getRange(row, 3, 1, 7).setValues([[
      fecha,
      detalle      || '',
      nit          || '',
      proveedor    || '',
      factura      || '',
      valor        || 0,
      observaciones || ''
    ]]);
    sheet.getRange(row, 8).setNumberFormat('$ #,##0.00');

    Logger.log('[registrarCajaMenor] Guardado exitosamente');
    return jsonResponse({ success: true, message: 'Registro de Caja Menor guardado' });

  } catch (error) {
    Logger.log('[registrarCajaMenor] Error: ' + error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

// ==================== CONSULTA ====================

function consultarCajaMenor(mes, anio) {
  try {
    // FIX: mes puede llegar como nombre ("NOVIEMBRE") o como número ("11")
    // Normalizamos siempre a nombre en español mayúsculas
    const nombreMes = resolverNombreMes(mes);
    if (!nombreMes) {
      return jsonResponse({ success: false, message: 'Mes inválido: ' + mes });
    }

    const nombreHoja = `${nombreMes} ${anio}`;
    Logger.log('[consultarCajaMenor] Buscando hoja: ' + nombreHoja);

    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(nombreHoja);

    if (!sheet) {
      Logger.log('[consultarCajaMenor] Hoja no encontrada: ' + nombreHoja);
      return jsonResponse({ success: false, message: `No existe la hoja para ${nombreHoja}` });
    }

    // Leer todos los datos de una vez (más eficiente)
    const datos     = sheet.getDataRange().getValues();
    const registros = [];

    // Fila 10 = índice 9 en el array (encabezados en fila 9 = índice 8)
    for (let i = 9; i < datos.length; i++) {
      if (!datos[i][2]) continue; // columna C vacía = fila vacía
      registros.push({
        fecha:         datos[i][2],
        detalle:       datos[i][3],
        nit:           datos[i][4],
        proveedor:     datos[i][5],
        factura:       datos[i][6],
        valor:         datos[i][7],
        observaciones: datos[i][8]
      });
    }

    Logger.log('[consultarCajaMenor] Registros encontrados: ' + registros.length);
    return jsonResponse({ success: true, registros });

  } catch (error) {
    Logger.log('[consultarCajaMenor] Error: ' + error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

// ==================== FUNCIONES AUXILIARES ====================

/**
 * Resuelve el nombre del mes ya sea que llegue como número ("11") o nombre ("NOVIEMBRE").
 * Retorna el nombre en mayúsculas, o null si es inválido.
 */
function resolverNombreMes(mes) {
  // Si ya es un nombre de mes válido en español
  const mesUpper = mes.toString().toUpperCase().trim();
  const esNombre = Object.values(MESES).includes(mesUpper);
  if (esNombre) return mesUpper;

  // Si es un número
  const mesNum = parseInt(mes);
  if (!isNaN(mesNum) && MESES[mesNum]) return MESES[mesNum];

  return null;
}

function parseFecha(fecha) {
  const p = fecha.split('-');
  return { anio: parseInt(p[0]), mes: parseInt(p[1]), dia: parseInt(p[2]) };
}

function crearHojaCajaMenor(ss, nombreHoja) {
  const sheet   = ss.insertSheet(nombreHoja);
  const headers = ['FECHA', 'DETALLE', 'NIT', 'PROVEEDOR', 'N° FAC', 'VALOR', 'OBSERVACIONES'];
  const anchos  = [120, 180, 120, 180, 120, 120, 200];

  // Encabezados en fila 9 (datos desde fila 10)
  sheet.getRange(9, 3, 1, headers.length).setValues([headers]);
  sheet.getRange(9, 3, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#ffe599')
    .setBorder(true, true, true, true, true, true);

  anchos.forEach((ancho, i) => sheet.setColumnWidth(3 + i, ancho));

  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}