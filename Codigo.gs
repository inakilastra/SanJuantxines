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

  var headers = values[0].map(function (h) {
    return String(h || '')
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim();
  });

  function findCol(keywords, fallbackIdx) {
    for (var i = 0; i < headers.length; i++) {
      for (var k = 0; k < keywords.length; k++) {
        if (headers[i].indexOf(keywords[k]) !== -1) return i;
      }
    }
    return fallbackIdx;
  }

  var idxPutxera = findCol(['putxera', 'num', 'nº'], 0);
  var idxNombre = findCol(['nombre', 'cuadrilla'], 1);
  var idxPoblacion = findCol(['poblacion', 'localidad', 'pueblo'], 2);
  var idxTotal = findCol(['total', 'puntos'], 3);
  var idxPresencia = findCol(['presencia'], 4);
  var idxPunto = findCol(['punto'], 5);
  var idxSabor = findCol(['sabor'], 6);
  var idxBasauri = findCol(['basauri', 'local'], 7);
  var idxGeneral = findCol(['general', 'puesto', 'pos'], 8);

  values.shift();
  var lista = [];
  values.forEach(function (row) {
    var putxeraNum = Number(row[idxPutxera]) || 0;
    if (putxeraNum <= 0 && !row[idxNombre]) return;
    lista.push({
      Putxera: putxeraNum,
      Nombre: String(row[idxNombre] || '').trim(),
      Poblacion: String(row[idxPoblacion] || '').trim(),
      Total: Number(row[idxTotal]) || 0,
      Presencia: Number(row[idxPresencia]) || 0,
      Punto: Number(row[idxPunto]) || 0,
      Sabor: Number(row[idxSabor]) || 0,
      Basauri: idxBasauri < row.length ? String(row[idxBasauri] || '').trim() : '',
      General: Number(row[idxGeneral]) || 0
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
  SW.getSheets().forEach(function (sheet) {
    var name = sheet.getName().trim();
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
    timestamp: new Date().toISOString()
  };

  return ContentService
    .createTextOutput(JSON.stringify(datosRespuesta))
    .setMimeType(ContentService.MimeType.JSON);
}
