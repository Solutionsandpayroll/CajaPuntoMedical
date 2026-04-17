/**
 * ================================================================
 * SISTEMA DE CAJA (ARQUEO) - PUNTO MEDICAL
 * Google Apps Script
 * ================================================================
 */

const SPREADSHEET_ID = '1xoSAY47E2X9pAA7_hU6ZLzfHOaE2Br-kbeO4DXhuY1Q';
const PLANTILLA_GID = '1606540802';
const CAJA_BASE_HISTORICA = 3326232;

const MESES = {
  1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL',
  5: 'MAYO', 6: 'JUNIO', 7: 'JULIO', 8: 'AGOSTO',
  9: 'SEPTIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'
};

// ==================== ENTRY POINTS ====================

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'consultarCaja') {
      return consultarCaja(e.parameter.mes, e.parameter.anio);
    }
    if (action === 'getPDFBase64') {
      return getPDFBase64();
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

    if (action === 'acopleCaja') return acopleCaja(data.fecha, data.movimientos);
    if (action === 'registrarEntrada') return registrarMovimiento(data.fecha, data.monto, data.concepto, 'entrada');
    if (action === 'registrarSalida') return registrarMovimiento(data.fecha, data.monto, data.concepto, 'salida', data.factura, data.nit);
    if (action === 'exportarComprobante') return exportarComprobante(data.fecha, data.movimientos);

    return jsonResponse({ success: false, message: 'Accion POST no valida: ' + action });
  } catch (error) {
    Logger.log('[doPost] Error: ' + error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

// ==================== ACOPLE DE CAJA ====================

function acopleCaja(fecha, movimientos) {
  try {
    const { mes, anio } = parseFecha(fecha);
    const nombreHoja = MESES[mes] + ' ' + anio;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    let sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) sheet = crearHojaMes(ss, nombreHoja);

    const plantillaSheet = ss.getSheetByName('PLANTILLA');
    if (!plantillaSheet) throw new Error('No existe la hoja PLANTILLA');

    plantillaSheet.getRange('A4:F18').clearContent();
    plantillaSheet.getRange('B2').setValue(fecha);

    const entradas = movimientos.filter(function (m) { return m.tipo === 'entrada'; });
    const salidas = movimientos.filter(function (m) { return m.tipo === 'salida'; });

    for (let i = 0; i < 15; i++) {
      plantillaSheet.getRange(4 + i, 1).setValue(entradas[i] ? entradas[i].concepto : '');
      plantillaSheet.getRange(4 + i, 2).setValue(entradas[i] ? entradas[i].monto : '');
      plantillaSheet.getRange(4 + i, 3).setValue(salidas[i] ? salidas[i].concepto : '');
      plantillaSheet.getRange(4 + i, 4).setValue(salidas[i] ? salidas[i].monto : '');
      plantillaSheet.getRange(4 + i, 5).setValue(salidas[i] ? (salidas[i].factura || '') : '');
      plantillaSheet.getRange(4 + i, 6).setValue(salidas[i] ? (salidas[i].nit || '') : '');
    }

    const totalEntradas = entradas.reduce(function (s, m) { return s + (parseFloat(m.monto) || 0); }, 0);
    const totalSalidas = salidas.reduce(function (s, m) { return s + (parseFloat(m.monto) || 0); }, 0);
    plantillaSheet.getRange('B19').setValue(totalEntradas);
    plantillaSheet.getRange('D19').setValue(totalSalidas);
    plantillaSheet.getRange('B20').setValue(totalEntradas - totalSalidas);

    const sourceRange = plantillaSheet.getRange('A2:F20');
    const ultimaFila = encontrarUltimaFilaConDatos(sheet);
    const filaInsertar = ultimaFila > 0 ? ultimaFila + 2 : 2;
    sourceRange.copyTo(sheet.getRange(filaInsertar, 1, 19, 6));

    Logger.log('[acopleCaja] Guardado en hoja: ' + nombreHoja + ' fila: ' + filaInsertar);
    return jsonResponse({ success: true, message: 'Acople de caja registrado correctamente' });
  } catch (error) {
    Logger.log('[acopleCaja] Error: ' + error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

// ==================== REGISTRO INDIVIDUAL ====================

function registrarMovimiento(fecha, monto, concepto, tipo, factura, nit) {
  try {
    const { mes, anio } = parseFecha(fecha);
    const nombreHoja = MESES[mes] + ' ' + anio;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    let sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) sheet = crearHojaMes(ss, nombreHoja);

    const datos = sheet.getDataRange().getValues();
    let filaDia = -1;

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
      ) { filaInsertar++; }
    } else {
      const ultimaFila = encontrarUltimaFilaConDatos(sheet);
      filaInsertar = ultimaFila + 3;
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
      sheet.getRange(filaInsertar, 5).setValue(factura || '');
      sheet.getRange(filaInsertar, 6).setValue(nit || '');
    }

    actualizarTotales(sheet);
    return jsonResponse({ success: true, message: (tipo === 'entrada' ? 'Entrada' : 'Salida') + ' registrada correctamente' });
  } catch (error) {
    Logger.log('[registrarMovimiento] Error: ' + error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

// ==================== CONSULTA DE CAJA ====================

function consultarCaja(mes, anio) {
  try {
    const nombreHoja = mes + ' ' + anio;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(nombreHoja);

    if (!sheet) {
      return jsonResponse({ success: false, message: 'No existe la hoja: ' + nombreHoja });
    }

    const datos = sheet.getDataRange().getValues();
    const anioNum = parseInt(anio);
    const mesNum = parseInt(Object.keys(MESES).find(function (k) { return MESES[k] === mes; }));
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
      let fe = 20, fs = 20, ft = 21;
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

    const cajaTotalMes = calcularCajaTotalAcumulada(mes, anio, ss) + saldo;

    // SOBRA/FALTA desde I2 / L2
    let sobraFalta = 0;
    try {
      const valorI2 = sheet.getRange('I2').getValue();
      const i2Activa = valorI2 !== '' && valorI2 !== null && !isNaN(parseFloat(valorI2));
      if (i2Activa) {
        const rangoL2 = sheet.getRange('L2');
        const valorL2 = rangoL2.getValue();
        const displayL2 = rangoL2.getDisplayValue();
        if (typeof valorL2 === 'number' && !isNaN(valorL2)) {
          sobraFalta = valorL2;
        } else {
          let clean = displayL2.toString().replace(/[^\d,.-]/g, '');
          if (clean.match(/,\d{2}$/)) clean = clean.replace(/\./g, '').replace(',', '.');
          const num = parseFloat(clean);
          sobraFalta = isNaN(num) ? 0 : num;
        }
      }
    } catch (e) { Logger.log('[consultarCaja] Error leyendo I2/L2: ' + e); }

    const cajaFinal = cajaTotalMes + sobraFalta;
    const registros = extraerRegistrosCaja(datos);
    const round2 = function (v) { return Math.round(v * 100) / 100; };

    Logger.log('[consultarCaja] totalEntradas=' + totalEntradas + ' totalSalidas=' + totalSalidas + ' saldo=' + saldo + ' cajaTotalMes=' + cajaTotalMes + ' sobraFalta=' + sobraFalta);

    return jsonResponse({
      success: true,
      mes: mes,
      anio: anio,
      totalEntradas: round2(totalEntradas),
      totalSalidas: round2(totalSalidas),
      saldo: round2(saldo),
      cajaTotalMes: round2(cajaTotalMes),
      sobraFalta: round2(sobraFalta),
      cajaFinal: round2(cajaFinal),
      registros: registros
    });
  } catch (error) {
    Logger.log('[consultarCaja] Error: ' + error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

function calcularCajaTotalAcumulada(mes, anio, ss) {
  const MAX_ITERACIONES = 36;
  let acumulado = 0;
  let mesActual = mes;
  let anioActual = anio;

  for (let i = 0; i < MAX_ITERACIONES; i++) {
    const anterior = getMesAnterior(mesActual, anioActual);
    const anioAnt = parseInt(anterior.anio);
    const mesAnt = parseInt(Object.keys(MESES).find(function (k) { return MESES[k] === anterior.mes; }));

    if (anioAnt < 2025 || (anioAnt === 2025 && mesAnt < 8)) {
      acumulado += CAJA_BASE_HISTORICA;
      break;
    }

    const sheetAnt = ss.getSheetByName(anterior.mes + ' ' + anterior.anio);
    if (!sheetAnt) { acumulado += CAJA_BASE_HISTORICA; break; }

    const datosAnt = sheetAnt.getDataRange().getValues();
    const esAntigua = (anioAnt === 2025 && mesAnt >= 8) || (anioAnt === 2026 && mesAnt <= 2);
    let saldoAnt = 0;

    if (esAntigua) {
      let ft = 6;
      while (ft < datosAnt.length && datosAnt[ft][7] !== '' && datosAnt[ft][7] != null) {
        const v = parseFloat(datosAnt[ft][7]); if (!isNaN(v)) saldoAnt += v; ft += 15;
      }
    } else {
      let ft = 21;
      while (ft < datosAnt.length && datosAnt[ft][1] !== '' && datosAnt[ft][1] != null) {
        const v = parseFloat(datosAnt[ft][1]); if (!isNaN(v)) saldoAnt += v; ft += 20;
      }
    }

    acumulado += saldoAnt;
    mesActual = anterior.mes;
    anioActual = anterior.anio;
  }
  return acumulado;
}

// ==================== PDF BASE64 (descarga sin auth) ====================

function getPDFBase64() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const plantillaSheet = ss.getSheetByName('PLANTILLA');
    if (!plantillaSheet) throw new Error('No existe la hoja PLANTILLA');

    const exportUrl = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/export'
      + '?format=pdf&size=letter&portrait=true&fitw=true'
      + '&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false'
      + '&gid=' + plantillaSheet.getSheetId() + '&range=A2:F20';

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { 'Authorization': 'Bearer ' + token }
    });

    const base64 = Utilities.base64Encode(response.getBlob().getBytes());
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: base64 }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log('[getPDFBase64] Error: ' + error);
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==================== EXPORTAR COMPROBANTE ====================

function exportarComprobante(fecha, movimientos) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tempSheet = ss.getSheetByName('TMP_COMPROBANTE');
    if (!tempSheet) throw new Error('No existe hoja TMP_COMPROBANTE');

    const exportUrl = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/export'
      + '?format=pdf&size=letter&portrait=true&fitw=true'
      + '&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false'
      + '&gid=' + tempSheet.getSheetId() + '&range=A2:F20';

    ScriptApp.getProjectTriggers()
      .filter(function (t) { return t.getHandlerFunction() === 'borrarHojaTemporal'; })
      .forEach(function (t) { ScriptApp.deleteTrigger(t); });

    ScriptApp.newTrigger('borrarHojaTemporal').timeBased().after(2 * 60 * 1000).create();

    return jsonResponse({ success: true, url: exportUrl });
  } catch (error) {
    Logger.log('[exportarComprobante] Error: ' + error);
    return jsonResponse({ success: false, message: error.toString() });
  }
}

function borrarHojaTemporal() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const temp = ss.getSheetByName('TMP_COMPROBANTE');
  if (temp) ss.deleteSheet(temp);
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'borrarHojaTemporal'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
}

// ==================== AUXILIARES ====================

function parseFecha(fecha) {
  const p = fecha.split('-');
  return { anio: parseInt(p[0]), mes: parseInt(p[1]), dia: parseInt(p[2]) };
}

function getMesAnterior(mes, anio) {
  const mesNum = parseInt(Object.keys(MESES).find(function (k) { return MESES[k] === mes; }));
  if (mesNum === 1) return { mes: 'DICIEMBRE', anio: (parseInt(anio) - 1).toString() };
  return { mes: MESES[mesNum - 1], anio: anio };
}

function encontrarUltimaFilaConDatos(sheet) {
  const datos = sheet.getDataRange().getValues();
  for (let i = datos.length - 1; i >= 0; i--) {
    if (datos[i].some(function (c) { return c !== '' && c !== null; })) return i + 1;
  }
  return 0;
}

function actualizarTotales(sheet) {
  const datos = sheet.getDataRange().getValues();
  for (let i = 2; i < datos.length; i++) {
    if (datos[i][0] && datos[i][0].toString().toUpperCase() === 'TOTAL') {
      const fila = i + 1;
      sheet.getRange(fila, 2).setFormula('=SUM(B3:B' + (fila - 1) + ')').setNumberFormat('$ #,##0.00');
      sheet.getRange(fila, 4).setFormula('=SUM(D3:D' + (fila - 1) + ')').setNumberFormat('$ #,##0.00');
    }
  }
}

function crearHojaMes(ss, nombreHoja) {
  const sheet = ss.insertSheet(nombreHoja);
  sheet.getRange('A1:D1').merge();
  sheet.getRange('A1').setValue('CAJA DE ' + nombreHoja)
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#ffff00').setFontColor('#ff0000');

  const headers = ['CONCEPTO DE ENTRADA', 'ENTRADAS DE DINERO', 'CONCEPTO DE SALIDA', 'SALIDAS DE DINERO'];
  sheet.getRange('A2:D2').setValues([headers]).setFontWeight('bold').setBackground('#cfe2f3').setBorder(true, true, true, true, true, true);
  sheet.setColumnWidths(1, 4, 180);

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

function testAcople() {
  const resultado = acopleCaja('2026-02-25', [
    { tipo: 'entrada', monto: 500000, concepto: 'Pago cliente TEST' },
    { tipo: 'salida', monto: 100000, concepto: 'Gasto TEST' }
  ]);
  Logger.log('Resultado: ' + resultado.getContent());
}

function testConsulta() {
  const resultado = consultarCaja('FEBRERO', '2026');
  Logger.log('Resultado: ' + resultado.getContent());
}

function extraerRegistrosCaja(datos) {
  var registros = [];
  var fechaActual = '';
  var contador = 0;

  for (var i = 0; i < datos.length; i++) {
    var row = datos[i] || [];

    // Detecta y arrastra la fecha del bloque actual.
    if (esFechaValida_(row[0])) fechaActual = normalizarFecha_(row[0]);
    if (esFechaValida_(row[1])) fechaActual = normalizarFecha_(row[1]);

    var conceptoEntrada = limpiarTexto_(row[0]);
    var montoEntrada = parseMonto_(row[1]);
    var conceptoSalida = limpiarTexto_(row[2]);
    var montoSalida = parseMonto_(row[3]);
    var factura = limpiarTexto_(row[4]);
    var nit = limpiarTexto_(row[5]);

    if (conceptoEntrada && montoEntrada > 0 && !esTextoSistema_(conceptoEntrada)) {
      contador++;
      registros.push({
        numero: contador,
        fecha: fechaActual || '',
        tipo: 'entrada',
        concepto: conceptoEntrada,
        monto: Math.round(montoEntrada * 100) / 100,
        factura: '',
        nit: ''
      });
    }

    if (conceptoSalida && montoSalida > 0 && !esTextoSistema_(conceptoSalida)) {
      contador++;
      registros.push({
        numero: contador,
        fecha: fechaActual || '',
        tipo: 'salida',
        concepto: conceptoSalida,
        monto: Math.round(montoSalida * 100) / 100,
        factura: factura,
        nit: nit
      });
    }
  }

  return registros;
}

function esFechaValida_(valor) {
  if (valor instanceof Date) return !isNaN(valor.getTime());
  if (valor === '' || valor === null || valor === undefined) return false;
  var s = String(valor).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return true;
  if (/^\d{2}\/\d{2}\/\d{4}/.test(s)) return true;
  if (/^\d{2}-\d{2}-\d{4}/.test(s)) return true;
  return false;
}

function normalizarFecha_(valor) {
  if (valor instanceof Date) {
    var y = valor.getFullYear();
    var m = ('0' + (valor.getMonth() + 1)).slice(-2);
    var d = ('0' + valor.getDate()).slice(-2);
    return d + '/' + m + '/' + y;
  }
  var s = String(valor || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    var p = s.split('T')[0].split('-');
    return p[2] + '/' + p[1] + '/' + p[0];
  }
  if (/^\d{2}[\/-]\d{2}[\/-]\d{4}/.test(s)) {
    return s.replace(/-/g, '/');
  }
  return s;
}

function limpiarTexto_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor).trim();
}

function parseMonto_(valor) {
  if (typeof valor === 'number') return isNaN(valor) ? 0 : valor;
  if (valor === '' || valor === null || valor === undefined) return 0;
  var s = String(valor).replace(/[^\d,.-]/g, '');
  if (/\,\d{2}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function esTextoSistema_(texto) {
  var t = String(texto || '').toUpperCase().trim();
  if (!t) return true;
  if (t === 'TOTAL') return true;
  if (t === 'FECHA DE CAJA') return true;
  if (t === 'TOTAL DEL DIA') return true;
  if (t === 'TOTAL DEL DÍA') return true;
  if (t === 'CONCEPTO DE ENTRADA') return true;
  if (t === 'CONCEPTO DE SALIDA') return true;
  if (t.indexOf('CAJA DE ') === 0) return true;
  return false;
}