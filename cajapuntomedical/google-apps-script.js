/**
 * ================================================================
 * SISTEMA DE CAJA - PUNTO MEDICAL
 * Google Apps Script — versión auditada y corregida
 * ================================================================
 */

// ==================== CONFIGURACIÓN ====================
const SPREADSHEET_ID = '1xoSAY47E2X9pAA7_hU6ZLzfHOaE2Br-kbeO4DXhuY1Q';
const PLANTILLA_GID   = '1606540802';

const MESES = {
  1: 'ENERO',    2: 'FEBRERO',   3: 'MARZO',     4: 'ABRIL',
  5: 'MAYO',     6: 'JUNIO',     7: 'JULIO',      8: 'AGOSTO',
  9: 'SEPTIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'
};

// Valor de caja acumulada antes de AGOSTO 2025 (base histórica)
const CAJA_BASE_HISTORICA = 3326232;

// ==================== ENTRY POINTS ====================

function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === 'consultarCaja') {
      return consultarCaja(e.parameter.mes, e.parameter.anio);
    }
    if (action === 'consultarCajaMenor') {
      return consultarCajaMenor(e.parameter.mes, e.parameter.anio);
    }
    return jsonResponse({ success: false, message: 'Acción GET no válida' });
  } catch (error) {
    Logger.log('[doGet] Error:', error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

function doPost(e) {
  try {
    const data    = JSON.parse(e.postData.contents);
    const action  = data.action;

    if (action === 'acopleCaja')        return acopleCaja(data.fecha, data.movimientos);
    if (action === 'registrarEntrada')  return registrarMovimiento(data.fecha, data.monto, data.concepto, 'entrada');
    if (action === 'registrarSalida')   return registrarMovimiento(data.fecha, data.monto, data.concepto, 'salida');
    if (action === 'exportarComprobante') return exportarComprobante(data.fecha, data.movimientos);
    if (action === 'registrarCajaMenor')  return registrarCajaMenor(data);

    return jsonResponse({ success: false, message: 'Acción POST no válida' });
  } catch (error) {
    Logger.log('[doPost] Error:', error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

// ==================== CAJA MENOR ====================

/**
 * Registra un movimiento de Caja Menor.
 * Payload esperado: { fecha, detalle, nit, proveedor, factura, valor, observaciones }
 */
function registrarCajaMenor(data) {
  try {
    const { mes, anio } = parseFecha(data.fecha);

    // Solo permitir desde NOVIEMBRE 2025
    if (anio < 2025 || (anio === 2025 && mes < 11)) {
      return jsonResponse({ success: false, message: 'Solo se permite desde NOVIEMBRE 2025' });
    }

    const nombreHoja = `${MESES[mes]} ${anio}`;
    const ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
    let   sheet      = ss.getSheetByName(nombreHoja);
    if (!sheet) sheet = crearHojaCajaMenor(ss, nombreHoja);

    const lastRow = Math.max(sheet.getLastRow(), 1);
    const fila    = lastRow + 1;

    sheet.getRange(fila, 1).setValue(data.fecha);
    sheet.getRange(fila, 2).setValue(data.detalle      || '');
    sheet.getRange(fila, 3).setValue(data.proveedor    || '');
    sheet.getRange(fila, 4).setValue(data.nit          || '');
    sheet.getRange(fila, 5).setValue(data.factura      || '');
    sheet.getRange(fila, 6).setValue(data.valor        || 0);
    sheet.getRange(fila, 6).setNumberFormat('$ #,##0.00');
    sheet.getRange(fila, 7).setValue(data.observaciones || '');

    return jsonResponse({ success: true, message: 'Registro de Caja Menor guardado' });
  } catch (error) {
    Logger.log('[registrarCajaMenor] Error:', error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

/**
 * Crea hoja de Caja Menor con encabezados completos (7 columnas).
 */
function crearHojaCajaMenor(ss, nombreHoja) {
  const sheet = ss.insertSheet(nombreHoja);
  const headers = ['FECHA', 'DETALLE', 'PROVEEDOR', 'NIT', 'FACTURA', 'VALOR', 'OBSERVACIONES'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#ffe599');
  sheet.setColumnWidths(1, headers.length, 140);
  return sheet;
}

/**
 * Consulta registros de Caja Menor para un mes/año.
 */
function consultarCajaMenor(mes, anio) {
  try {
    const nombreHoja = `${mes} ${anio}`;
    const ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet      = ss.getSheetByName(nombreHoja);

    if (!sheet) {
      return jsonResponse({ success: false, message: `No existe la hoja ${nombreHoja}` });
    }

    const datos     = sheet.getDataRange().getValues();
    const registros = [];
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0]) continue;
      registros.push({
        fecha:         datos[i][0],
        detalle:       datos[i][1],
        proveedor:     datos[i][2],
        nit:           datos[i][3],
        factura:       datos[i][4],
        valor:         datos[i][5],
        observaciones: datos[i][6]
      });
    }
    return jsonResponse({ success: true, registros });
  } catch (error) {
    Logger.log('[consultarCajaMenor] Error:', error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

// ==================== ACOPLE DE CAJA ====================

function acopleCaja(fecha, movimientos) {
  try {
    const { mes, anio } = parseFecha(fecha);
    const nombreHoja    = `${MESES[mes]} ${anio}`;
    const ss            = SpreadsheetApp.openById(SPREADSHEET_ID);

    let sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) sheet = crearHojaMes(ss, nombreHoja);

    const plantillaSheet = ss.getSheetByName('PLANTILLA');
    if (!plantillaSheet) throw new Error('No existe la hoja PLANTILLA');

    // Limpiar rango de trabajo en plantilla
    plantillaSheet.getRange('A4:D18').clearContent();
    plantillaSheet.getRange('B2').setValue(fecha);

    const entradas = movimientos.filter(m => m.tipo === 'entrada');
    const salidas  = movimientos.filter(m => m.tipo === 'salida');

    for (let i = 0; i < 15; i++) {
      plantillaSheet.getRange(4 + i, 1).setValue(entradas[i] ? entradas[i].concepto : '');
      plantillaSheet.getRange(4 + i, 2).setValue(entradas[i] ? entradas[i].monto    : '');
      plantillaSheet.getRange(4 + i, 3).setValue(salidas[i]  ? salidas[i].concepto  : '');
      plantillaSheet.getRange(4 + i, 4).setValue(salidas[i]  ? salidas[i].monto     : '');
    }

    const totalEntradas = entradas.reduce((s, m) => s + (parseFloat(m.monto) || 0), 0);
    const totalSalidas  = salidas.reduce((s, m)  => s + (parseFloat(m.monto) || 0), 0);
    plantillaSheet.getRange('B19').setValue(totalEntradas);
    plantillaSheet.getRange('D19').setValue(totalSalidas);
    plantillaSheet.getRange('B20').setValue(totalEntradas - totalSalidas);

    // Copiar bloque a la hoja del mes
    const sourceRange = plantillaSheet.getRange('A2:D20');
    let   ultimaFila  = encontrarUltimaFilaConDatos(sheet);
    let   filaInsertar = ultimaFila > 0 ? ultimaFila + 2 : 2;
    sourceRange.copyTo(sheet.getRange(filaInsertar, 1, 19, 4));

    return jsonResponse({ success: true, message: 'Acople de caja registrado correctamente' });
  } catch (error) {
    Logger.log('[acopleCaja] Error:', error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

// ==================== REGISTRO INDIVIDUAL (entrada/salida) ====================

/**
 * Función unificada que reemplaza registrarEntrada y registrarSalida.
 * tipo: 'entrada' | 'salida'
 */
function registrarMovimiento(fecha, monto, concepto, tipo) {
  try {
    const { mes, anio } = parseFecha(fecha);
    const nombreHoja    = `${MESES[mes]} ${anio}`;
    const ss            = SpreadsheetApp.openById(SPREADSHEET_ID);

    let sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) sheet = crearHojaMes(ss, nombreHoja);

    const datos   = sheet.getDataRange().getValues();
    let filaDia   = -1;

    for (let i = 2; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString().startsWith(fecha)) {
        filaDia = i + 1;
        break;
      }
    }

    let filaInsertar;
    if (filaDia !== -1) {
      filaInsertar = filaDia;
      while (
        filaInsertar <= sheet.getLastRow() &&
        (sheet.getRange(filaInsertar, 1).getValue() || sheet.getRange(filaInsertar, 3).getValue())
      ) {
        filaInsertar++;
      }
    } else {
      let ultimaFila  = encontrarUltimaFilaConDatos(sheet);
      filaInsertar    = ultimaFila + 3;
      sheet.getRange(filaInsertar, 1, 1, 4).merge();
      sheet.getRange(filaInsertar, 1).setValue(fecha).setFontWeight('bold').setBackground('#e0f4ff');
      filaInsertar++;
    }

    if (tipo === 'entrada') {
      sheet.getRange(filaInsertar, 1).setValue(concepto);
      sheet.getRange(filaInsertar, 2).setValue(monto).setNumberFormat('$ #,##0.00');
    } else {
      sheet.getRange(filaInsertar, 3).setValue(concepto);
      sheet.getRange(filaInsertar, 4).setValue(monto).setNumberFormat('$ #,##0.00');
    }

    actualizarTotales(sheet);

    return jsonResponse({ success: true, message: `${tipo === 'entrada' ? 'Entrada' : 'Salida'} registrada correctamente` });
  } catch (error) {
    Logger.log('[registrarMovimiento] Error:', error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

// ==================== CONSULTA DE CAJA ====================

function consultarCaja(mes, anio) {
  try {
    const nombreHoja = `${mes} ${anio}`;
    const ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet      = ss.getSheetByName(nombreHoja);

    if (!sheet) {
      return jsonResponse({ success: false, message: `No existe la hoja para ${nombreHoja}` });
    }

    const datos    = sheet.getDataRange().getValues();
    const anioNum  = parseInt(anio);
    let   mesNum   = parseInt(Object.keys(MESES).find(k => MESES[k] === mes));

    const esAntigua = (anioNum === 2025 && mesNum >= 8) || (anioNum === 2026 && mesNum <= 2);

    let totalEntradas = 0, totalSalidas = 0, saldo = 0;

    if (esAntigua) {
      let fe = 17, fs = 17, ft = 6;
      while (fe < datos.length && datos[fe][2] !== '' && datos[fe][2] != null) {
        const v = parseFloat(datos[fe][2]); if (!isNaN(v)) totalEntradas += v; fe += 15;
      }
      while (fs < datos.length && datos[fs][4] !== '' && datos[fs][4] != null) {
        const v = parseFloat(datos[fs][4]); if (!isNaN(v)) totalSalidas += v; fs += 15;
      }
      while (ft < datos.length && datos[ft][7] !== '' && datos[ft][7] != null) {
        const v = parseFloat(datos[ft][7]); if (!isNaN(v)) saldo += v; ft += 15;
      }
    } else {
      let fe = 18, fs = 18, ft = 19;
      while (fe < datos.length && datos[fe][1] !== '' && datos[fe][1] != null) {
        const v = parseFloat(datos[fe][1]); if (!isNaN(v)) totalEntradas += v; fe += 20;
      }
      while (fs < datos.length && datos[fs][3] !== '' && datos[fs][3] != null) {
        const v = parseFloat(datos[fs][3]); if (!isNaN(v)) totalSalidas += v; fs += 20;
      }
      while (ft < datos.length && datos[ft][1] !== '' && datos[ft][1] != null) {
        const v = parseFloat(datos[ft][1]); if (!isNaN(v)) saldo += v; ft += 20;
      }
    }

    // Caja total acumulada (iterativa, sin recursión)
    const cajaTotalMes = calcularCajaTotalAcumulada(mes, anio, ss) + saldo;

    // Verificar si I2 tiene valor — determina si aplica SOBRA/FALTA
    let i2Activa = false;
    try {
      const valorI2 = sheet.getRange('I2').getValue();
      i2Activa = valorI2 !== '' && valorI2 !== null && !isNaN(parseFloat(valorI2));
      Logger.log('[consultarCaja] I2 valor:', valorI2, '| i2Activa:', i2Activa);
    } catch (e) {
      Logger.log('[consultarCaja] Error leyendo I2:', e);
    }

    // SOBRA/FALTA desde celda L2 — solo se usa si I2 tiene valor
    let sobraFalta = 0;
    if (i2Activa) {
      try {
        const rangoL2   = sheet.getRange('L2');
        const valorL2   = rangoL2.getValue();
        const displayL2 = rangoL2.getDisplayValue();
        if (typeof valorL2 === 'number' && !isNaN(valorL2)) {
          sobraFalta = valorL2;
        } else {
          let clean = displayL2.toString().replace(/[^\d,.-]/g, '');
          if (clean.match(/,\d{2}$/)) clean = clean.replace(/\./g, '').replace(',', '.');
          const num = parseFloat(clean);
          sobraFalta = isNaN(num) ? 0 : num;
        }
        Logger.log('[consultarCaja] L2 (sobraFalta):', sobraFalta);
      } catch (e) {
        Logger.log('[consultarCaja] Error leyendo L2:', e);
      }
    }

    // CAJA FINAL:
    // - Si I2 tiene valor: Caja Total + Sobra/Falta
    // - Si I2 está vacía:  Caja Final = Caja Total (sobraFalta queda en 0, no afecta)
    const cajaFinal = cajaTotalMes + sobraFalta;

    // Redondear a 2 decimales para evitar acumulación de errores de punto flotante
    // Ejemplo: 0.1 + 0.2 = 0.30000000000000004 en JS — round2 lo corrige
    const round2 = v => Math.round(v * 100) / 100;

    Logger.log('[consultarCaja] Resultado:', { totalEntradas, totalSalidas, saldo, cajaTotalMes, sobraFalta, cajaFinal, i2Activa });

    return jsonResponse({
      success:       true,
      mes,
      anio,
      totalEntradas: round2(totalEntradas),
      totalSalidas:  round2(totalSalidas),
      saldo:         round2(saldo),
      cajaTotalMes:  round2(cajaTotalMes),
      sobraFalta:    round2(sobraFalta),
      cajaFinal:     round2(cajaFinal)
    });
  } catch (error) {
    Logger.log('[consultarCaja] Error:', error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

/**
 * Calcula la caja acumulada hasta el mes anterior — ITERATIVA (sin recursión).
 * Límite: máx 36 meses hacia atrás para evitar ejecuciones largas.
 */
function calcularCajaTotalAcumulada(mes, anio, ss) {
  const MAX_ITERACIONES = 36;
  let acumulado  = 0;
  let mesActual  = mes;
  let anioActual = anio;

  for (let i = 0; i < MAX_ITERACIONES; i++) {
    const anterior = getMesAnterior(mesActual, anioActual);

    // Condición de parada: llegamos a antes de AGOSTO 2025
    const anioAnt = parseInt(anterior.anio);
    const mesAnt  = parseInt(Object.keys(MESES).find(k => MESES[k] === anterior.mes));
    if (anioAnt < 2025 || (anioAnt === 2025 && mesAnt < 8)) {
      acumulado += CAJA_BASE_HISTORICA;
      break;
    }

    const sheetAnt = ss.getSheetByName(`${anterior.mes} ${anterior.anio}`);
    if (!sheetAnt) {
      acumulado += CAJA_BASE_HISTORICA;
      break;
    }

    const datosAnt   = sheetAnt.getDataRange().getValues();
    const esAntigua  = (anioAnt === 2025 && mesAnt >= 8) || (anioAnt === 2026 && mesAnt <= 2);
    let   saldoAnt   = 0;

    if (esAntigua) {
      let ft = 6;
      while (ft < datosAnt.length && datosAnt[ft][7] !== '' && datosAnt[ft][7] != null) {
        const v = parseFloat(datosAnt[ft][7]); if (!isNaN(v)) saldoAnt += v; ft += 15;
      }
    } else {
      let ft = 19;
      while (ft < datosAnt.length && datosAnt[ft][1] !== '' && datosAnt[ft][1] != null) {
        const v = parseFloat(datosAnt[ft][1]); if (!isNaN(v)) saldoAnt += v; ft += 20;
      }
    }

    acumulado += saldoAnt;
    mesActual  = anterior.mes;
    anioActual = anterior.anio;
  }

  return acumulado;
}

// ==================== EXPORTAR COMPROBANTE ====================

function exportarComprobante(fecha, movimientos) {
  try {
    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tempSheet = ss.getSheetByName('TMP_COMPROBANTE');
    if (!tempSheet) throw new Error('No existe hoja TMP_COMPROBANTE');

    const exportUrl = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export`
      + `?format=pdf&size=letter&portrait=true&fitw=true`
      + `&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`
      + `&gid=${tempSheet.getSheetId()}&range=A2:D20`;

    // Limpiar triggers anteriores de borrado antes de crear uno nuevo
    ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === 'borrarHojaTemporal')
      .forEach(t => ScriptApp.deleteTrigger(t));

    ScriptApp.newTrigger('borrarHojaTemporal').timeBased().after(2 * 60 * 1000).create();

    return jsonResponse({ success: true, url: exportUrl });
  } catch (error) {
    Logger.log('[exportarComprobante] Error:', error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

function borrarHojaTemporal() {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const temp = ss.getSheetByName('TMP_COMPROBANTE');
  if (temp) ss.deleteSheet(temp);

  // Auto-eliminar este trigger
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'borrarHojaTemporal')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// ==================== FUNCIONES AUXILIARES ====================

function parseFecha(fecha) {
  const p = fecha.split('-');
  return { anio: parseInt(p[0]), mes: parseInt(p[1]), dia: parseInt(p[2]) };
}

function getMesAnterior(mes, anio) {
  const mesNum = parseInt(Object.keys(MESES).find(k => MESES[k] === mes));
  if (mesNum === 1) return { mes: 'DICIEMBRE', anio: (parseInt(anio) - 1).toString() };
  return { mes: MESES[mesNum - 1], anio };
}

function encontrarUltimaFilaConDatos(sheet) {
  const datos = sheet.getDataRange().getValues();
  for (let i = datos.length - 1; i >= 0; i--) {
    if (datos[i].some(c => c !== '' && c !== null)) return i + 1;
  }
  return 0;
}

function actualizarTotales(sheet) {
  const datos = sheet.getDataRange().getValues();
  for (let i = 2; i < datos.length; i++) {
    if (datos[i][0] && datos[i][0].toString().toUpperCase() === 'TOTAL') {
      const fila = i + 1;
      sheet.getRange(fila, 2).setFormula(`=SUM(B3:B${fila - 1})`).setNumberFormat('$ #,##0.00');
      sheet.getRange(fila, 4).setFormula(`=SUM(D3:D${fila - 1})`).setNumberFormat('$ #,##0.00');
    }
  }
}

function crearHojaMes(ss, nombreHoja) {
  const sheet = ss.insertSheet(nombreHoja);
  sheet.getRange('A1:D1').merge();
  sheet.getRange('A1').setValue(`CAJA DE ${nombreHoja}`)
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#ffff00').setFontColor('#ff0000');

  const headers = ['CONCEPTO DE ENTRADA', 'ENTRADAS DE DINERO', 'CONCEPTO DE SALIDA', 'SALIDAS DE DINERO'];
  sheet.getRange('A2:D2').setValues([headers]).setFontWeight('bold').setBackground('#cfe2f3').setBorder(true,true,true,true,true,true);
  sheet.setColumnWidths(1, 4, 180);

  // Fila de TOTAL en fila 19 (consistente con parser "nueva estructura")
  sheet.getRange('A19').setValue('TOTAL').setFontWeight('bold').setBackground('#ffff00');
  sheet.getRange('B19').setFormula('=SUM(B3:B18)').setBackground('#ffff00').setNumberFormat('$ #,##0.00');
  sheet.getRange('C19').setValue('TOTAL').setFontWeight('bold').setBackground('#ffff00');
  sheet.getRange('D19').setFormula('=SUM(D3:D18)').setBackground('#ffff00').setNumberFormat('$ #,##0.00');

  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== TESTS MANUALES ====================

function testCrearHoja() {
  crearHojaMes(SpreadsheetApp.openById(SPREADSHEET_ID), 'TEST 2026');
  Logger.log('Hoja de prueba creada');
}

function testRegistrarEntrada() {
  registrarMovimiento('2026-02-15', 500000, 'Pago cliente TEST', 'entrada');
  Logger.log('Entrada de prueba registrada');
}

function testRegistrarSalida() {
  registrarMovimiento('2026-02-15', 100000, 'Compra suministros TEST', 'salida');
  Logger.log('Salida de prueba registrada');
}