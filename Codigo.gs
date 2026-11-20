// ==========================================
// Google Apps Script - San Juantxines Basauri
// Concurso de Putxeras (con Histórico P20xx)
// ==========================================

const SW = SpreadsheetApp.openById('1vBkln5NjLvkbeTUC0tpf4yobTiJtW_CYI05BPAEqZFg');
const sControlWeb = SW.getSheetByName('ControlWeb');
const sPutxerasInscritas = SW.getSheetByName('PutxerasInscritas');
const sPutxerasPoblacion = SW.getSheetByName('PutxerasPoblacion');
const sPutxeras = SW.getSheetByName('Putxeras');
const sVotacionB5 = SW.getSheetByName('VotacionB5');
const sVotacionG10 = SW.getSheetByName('VotacionG10');
const sVotacionG11 = SW.getSheetByName('VotacionG11');
const sClasificacion = SW.getSheetByName('Clasificacion');
const sBasauri = SW.getSheetByName('Basauri');
const sTransparencia = SW.getSheetByName('Transparencia');

/**
 * Lee una pestaña de clasificación (actual o histórica P2023, P2024, P2025...)
 * detectando columnas por cabecera si existen, o usando el orden estándar 0..8.
 */
function leerHojaClasificacion(sheet) {
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  function norm(txt) {
    return String(txt || '')
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim();
  }

  function extraerNumero(val) {
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
    var s = String(val || '').replace(',', '.').replace(/[^0-9.\-]/g, '');
    var n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  // Buscar fila de cabeceras (normalmente fila 0, pero por si hay título arriba revisamos hasta fila 2)
  var headerRowIdx = 0;
  for (var r = 0; r < Math.min(3, values.length); r++) {
    var rowStr = values[r].map(norm).join(' ');
    if (rowStr.indexOf('nombre') !== -1 || rowStr.indexOf('putxera') !== -1 || rowStr.indexOf('sabor') !== -1) {
      headerRowIdx = r;
      break;
    }
  }

  var headers = values[headerRowIdx].map(norm);

  function findCol(keywords, fallbackIdx) {
    // Primero coincidencia exacta
    for (var i = 0; i < headers.length; i++) {
      for (var k = 0; k < keywords.length; k++) {
        if (headers[i] === keywords[k]) return i;
      }
    }
    // Luego coincidencia parcial
    for (var i = 0; i < headers.length; i++) {
      for (var k = 0; k < keywords.length; k++) {
        if (headers[i].indexOf(keywords[k]) !== -1) return i;
      }
    }
    return fallbackIdx;
  }

  var idxPutxera = findCol(['putxera', 'nº', 'num', 'numero'], 0);
  var idxNombre = findCol(['nombre', 'cuadrilla', 'participante'], 1);
  var idxPoblacion = findCol(['poblacion', 'localidad', 'pueblo'], 2);
  var idxTotal = findCol(['total', 'puntos'], 3);
  var idxPresencia = findCol(['presencia'], 4);
  var idxPunto = findCol(['punto'], 5);
  var idxSabor = findCol(['sabor'], 6);
  var idxBasauri = findCol(['basauri', 'local'], -1);
  var idxGeneral = findCol(['general', 'puesto', 'pos'], 8);
  if (idxBasauri === -1 && headers.length >= 9 && headers.indexOf('nombre') !== -1) {
    // Si tiene 9 columnas o más y cabecera estándar, comprobamos si col 7 no es General
    if (idxGeneral !== 7) idxBasauri = 7;
  }

  var dataRows = values.slice(headerRowIdx + 1);
  var lista = [];
  dataRows.forEach(function (row) {
    var putxeraNum = extraerNumero(row[idxPutxera]);
    var nombre = String(row[idxNombre] || '').trim();
    if (putxeraNum <= 0 && !nombre) return;
    lista.push({
      Putxera: putxeraNum,
      Nombre: nombre,
      Poblacion: idxPoblacion >= 0 && idxPoblacion < row.length ? String(row[idxPoblacion] || '').trim() : '',
      Total: idxTotal >= 0 && idxTotal < row.length ? extraerNumero(row[idxTotal]) : 0,
      Presencia: idxPresencia >= 0 && idxPresencia < row.length ? extraerNumero(row[idxPresencia]) : 0,
      Punto: idxPunto >= 0 && idxPunto < row.length ? extraerNumero(row[idxPunto]) : 0,
      Sabor: idxSabor >= 0 && idxSabor < row.length ? extraerNumero(row[idxSabor]) : 0,
      Basauri: (idxBasauri >= 0 && idxBasauri !== idxGeneral && idxBasauri < row.length) ? String(row[idxBasauri] || '').trim() : '',
      General: idxGeneral >= 0 && idxGeneral < row.length ? extraerNumero(row[idxGeneral]) : 0
    });
  });

  return lista;
}

function doGet(e) {
  // 1. ControlWeb
  var dataControlWeb = sControlWeb.getDataRange().getValues();
  dataControlWeb.shift();
  var ControlWeb = [];
  dataControlWeb.forEach(function (row) {
    ControlWeb.push({ Estado: String(row[1] || "").trim(), Podio: String(row[2] || "").trim() });
  });

  // 2. PutxerasInscritas
  var dataPutxerasInscritas = sPutxerasInscritas.getDataRange().getValues();
  dataPutxerasInscritas.shift();
  var PutxerasInscritas = [];
  dataPutxerasInscritas.forEach(function (row) {
    PutxerasInscritas.push({ Putxeras: Number(row[0]) || 0 });
  });

  // 3. PutxerasPoblacion
  var dataPutxerasPoblacion = sPutxerasPoblacion.getDataRange().getValues();
  dataPutxerasPoblacion.shift();
  var PutxerasPoblaciones = [];
  dataPutxerasPoblacion.forEach(function (row) {
    PutxerasPoblaciones.push({ Poblacion: String(row[0] || ""), Putxeras: Number(row[1]) || 0 });
  });

  // 4. Putxeras
  var dataPutxeras = sPutxeras.getDataRange().getValues();
  dataPutxeras.shift();
  var Putxeras = [];
  dataPutxeras.forEach(function (row) {
    Putxeras.push({
      Putxera: Number(row[0]) || 0,
      Nombre: String(row[1] || ""),
      Provincia: String(row[2] || ""),
      Poblacion: String(row[3] || ""),
      Barrio: String(row[4] || "")
    });
  });

  // 5. VotacionB5
  var dataVotacionB5 = sVotacionB5.getDataRange().getValues();
  dataVotacionB5.shift();
  var VotacionB5es = [];
  dataVotacionB5.forEach(function (row) {
    VotacionB5es.push({
      Putxera: Number(row[0]) || 0,
      Nombre: String(row[1] || ""),
      Poblacion: String(row[2] || "")
    });
  });

  // 6. VotacionG10
  var dataVotacionG10 = sVotacionG10.getDataRange().getValues();
  dataVotacionG10.shift();
  var VotacionG10es = [];
  dataVotacionG10.forEach(function (row) {
    VotacionG10es.push({
      Putxera: Number(row[0]) || 0,
      Nombre: String(row[1] || ""),
      Poblacion: String(row[2] || "")
    });
  });

  // 7. VotacionG11
  var dataVotacionG11 = sVotacionG11.getDataRange().getValues();
  dataVotacionG11.shift();
  var VotacionG11es = [];
  dataVotacionG11.forEach(function (row) {
    VotacionG11es.push({
      Putxera: Number(row[0]) || 0,
      Nombre: String(row[1] || ""),
      Poblacion: String(row[2] || "")
    });
  });

  // 8. Clasificacion actual
  var Clasificaciones = leerHojaClasificacion(sClasificacion);

  // 9. Basauri actual
  var dataBasauri = sBasauri.getDataRange().getValues();
  dataBasauri.shift();
  var Basauris = [];
  dataBasauri.forEach(function (row) {
    Basauris.push({
      Putxera: Number(row[0]) || 0,
      Nombre: String(row[1] || ""),
      Poblacion: String(row[2] || ""),
      Total: Number(row[3]) || 0,
      Presencia: Number(row[4]) || 0,
      Punto: Number(row[5]) || 0,
      Sabor: Number(row[6]) || 0,
      Basauri: Number(row[7]) || 0
    });
  });

  // 10. Transparencia
  var dataTransparencia = sTransparencia.getDataRange().getValues();
  dataTransparencia.shift();
  var Transparencia = [];
  dataTransparencia.forEach(function (row) {
    Transparencia.push({
      Concepto: String(row[1] || ""),
      Tipo: String(row[2] || ""),
      Euros: Number(row[3]) || 0
    });
  });

  // 11. Histórico automático: busca todas las pestañas llamadas P2023, P2024, P2025, etc.
  var Historico = {};
  var DebugHojas = [];
  SW.getSheets().forEach(function (sheet) {
    var name = sheet.getName().trim();
    DebugHojas.push(name);
    if (/^P\d{4}$/i.test(name)) {
      var year = name.substring(1); // "2023", "2024", "2025"...
      Historico[year] = leerHojaClasificacion(sheet);
    }
  });

  var datosRespuesta = {
    ControlWeb: ControlWeb,
    PutxerasInscritas: PutxerasInscritas,
    PutxerasPoblaciones: PutxerasPoblaciones,
    Putxeras: Putxeras,
    VotacionB5es: VotacionB5es,
    VotacionG10es: VotacionG10es,
    VotacionG11es: VotacionG11es,
    Clasificaciones: Clasificaciones,
    Basauris: Basauris,
    Transparencia: Transparencia,
    Historico: Historico,
    DebugHojas: DebugHojas,
    timestamp: new Date().toISOString()
  };

  return ContentService
    .createTextOutput(JSON.stringify(datosRespuesta))
    .setMimeType(ContentService.MimeType.JSON);
}
