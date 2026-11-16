// ==========================================
// Google Apps Script - San Juantxines Basauri
// Concurso de Putxeras
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

function doGet(e) {
  // 1. Obtener datos de ControlWeb
  var dataControlWeb = sControlWeb.getDataRange().getValues();
  dataControlWeb.shift();
  var ControlWeb = [];
  dataControlWeb.forEach(row => {
    ControlWeb.push({ Estado: String(row[1] || "").trim(), Podio: String(row[2] || "").trim() });
  });

  // 2. Obtener datos de PutxerasInscritas
  var dataPutxerasInscritas = sPutxerasInscritas.getDataRange().getValues();
  dataPutxerasInscritas.shift();
  var PutxerasInscritas = [];
  dataPutxerasInscritas.forEach(row => {
    PutxerasInscritas.push({ Putxeras: Number(row[0]) || 0 });
  });

  // 3. Obtener datos de PutxerasPoblacion
  var dataPutxerasPoblacion = sPutxerasPoblacion.getDataRange().getValues();
  dataPutxerasPoblacion.shift();
  var PutxerasPoblaciones = [];
  dataPutxerasPoblacion.forEach(row => {
    PutxerasPoblaciones.push({ Poblacion: String(row[0] || ""), Putxeras: Number(row[1]) || 0 });
  });

  // 4. Obtener datos de Putxeras
  var dataPutxeras = sPutxeras.getDataRange().getValues();
  dataPutxeras.shift();
  var Putxeras = [];
  dataPutxeras.forEach(row => {
    Putxeras.push({
      Putxera: Number(row[0]) || 0,
      Nombre: String(row[1] || ""),
      Provincia: String(row[2] || ""),
      Poblacion: String(row[3] || ""),
      Barrio: String(row[4] || "")
    });
  });

  // 5. Obtener datos de VotacionB5
  var dataVotacionB5 = sVotacionB5.getDataRange().getValues();
  dataVotacionB5.shift();
  var VotacionB5es = [];
  dataVotacionB5.forEach(row => {
    VotacionB5es.push({
      Putxera: Number(row[0]) || 0,
      Nombre: String(row[1] || ""),
      Poblacion: String(row[2] || "")
    });
  });

  // 6. Obtener datos de VotacionG10
  var dataVotacionG10 = sVotacionG10.getDataRange().getValues();
  dataVotacionG10.shift();
  var VotacionG10es = [];
  dataVotacionG10.forEach(row => {
    VotacionG10es.push({
      Putxera: Number(row[0]) || 0,
      Nombre: String(row[1] || ""),
      Poblacion: String(row[2] || "")
    });
  });

  // 7. Obtener datos de VotacionG11
  var dataVotacionG11 = sVotacionG11.getDataRange().getValues();
  dataVotacionG11.shift();
  var VotacionG11es = [];
  dataVotacionG11.forEach(row => {
    VotacionG11es.push({
      Putxera: Number(row[0]) || 0,
      Nombre: String(row[1] || ""),
      Poblacion: String(row[2] || "")
    });
  });

  // 8. Obtener datos de Clasificacion
  var dataClasificacion = sClasificacion.getDataRange().getValues();
  dataClasificacion.shift();
  var Clasificaciones = [];
  dataClasificacion.forEach(row => {
    Clasificaciones.push({
      Putxera: Number(row[0]) || 0,
      Nombre: String(row[1] || ""),
      Poblacion: String(row[2] || ""),
      Total: Number(row[3]) || 0,
      Presencia: Number(row[4]) || 0,
      Punto: Number(row[5]) || 0,
      Sabor: Number(row[6]) || 0,
      Basauri: String(row[7] || ""),
      General: Number(row[8]) || 0
    });
  });

  // 9. Obtener datos de Basauri
  var dataBasauri = sBasauri.getDataRange().getValues();
  dataBasauri.shift();
  var Basauris = [];
  dataBasauri.forEach(row => {
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

  // 10. Obtener datos de Transparencia
  var dataTransparencia = sTransparencia.getDataRange().getValues();
  dataTransparencia.shift();
  var Transparencia = [];
  dataTransparencia.forEach(row => {
    Transparencia.push({
      Concepto: String(row[1] || ""),
      Tipo: String(row[2] || ""),
      Euros: Number(row[3]) || 0
    });
  });

  // ========================================================
  // RESPUESTA EN FORMATO JSON PARA GITHUB PAGES / WEB
  // Devuelve JSON si se llama con ?api=json o por defecto
  // ========================================================
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
    timestamp: new Date().toISOString()
  };

  // Si se pide como API o si no se usa como template HTML
  if (!e || !e.parameter || e.parameter.api === 'json' || e.parameter.format === 'json') {
    return ContentService
      .createTextOutput(JSON.stringify(datosRespuesta))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Compatibilidad hacia atrás si se sigue abriendo dentro de Apps Script
  try {
    var template = HtmlService.createTemplateFromFile('index');
    template.ControlWeb = ControlWeb;
    template.PutxerasPoblaciones = PutxerasPoblaciones;
    template.PutxerasInscritas = PutxerasInscritas;
    template.Putxeras = Putxeras;
    template.VotacionB5es = VotacionB5es;
    template.VotacionG10es = VotacionG10es;
    template.VotacionG11es = VotacionG11es;  
    template.Clasificaciones = Clasificaciones;
    template.Basauris = Basauris;
    template.Transparencia = Transparencia;
    return template.evaluate()
      .setTitle("Concurso de Putxeras · San Juantxines")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify(datosRespuesta))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

