//Global Variables
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

function doGet(e)
{
//Obtener los datos de la hoja ControlWeb
  var dataControlWeb = sControlWeb.getDataRange().getValues();
  dataControlWeb.shift();

  var ControlWeb    = [];
  var aControlWeb     = {
    Estado:        String,
    Podio:         String
  };

  // Relleno Array ControlWeb
  dataControlWeb.forEach( row =>{
    aControlWeb = {Estado: row[1], Podio: row[2]};
    ControlWeb.push(aControlWeb);
  });
  
//Obtener los datos de la hoja PutxerasInscritas
  var dataPutxerasInscritas = sPutxerasInscritas.getDataRange().getValues();
  dataPutxerasInscritas.shift();

  var PutxerasInscritas    = [];
  var aPutxerasInscrita     = {
    Putxeras:  Number
  };

  // Relleno Array PutxerasInscritas
  dataPutxerasInscritas.forEach( row =>{
    aPutxerasInscrita = {Putxeras: row[0]};
    PutxerasInscritas.push(aPutxerasInscrita);
  });

//Obtener los datos de la hoja PutxerasPoblacion
  var dataPutxerasPoblacion = sPutxerasPoblacion.getDataRange().getValues();
  dataPutxerasPoblacion.shift();

  var PutxerasPoblaciones    = [];
  var aPutxerasPoblacion     = {
    Poblacion:  String,
    Putxeras:  Number
  };

  // Relleno Array PutxerasPoblacion
  dataPutxerasPoblacion.forEach( row =>{
    aPutxerasPoblacion = {Poblacion: row[0], Putxeras: row[1]};
    PutxerasPoblaciones.push(aPutxerasPoblacion);
  });

//Obtener los datos de la hoja Putxeras
  var dataPutxeras = sPutxeras.getDataRange().getValues();
  dataPutxeras.shift();

  var Putxeras    = [];
  var aPutxeras     = {
    Putxera:   Number,
  	Nombre:    String,
    Provincia: String,
    Poblacion: String,
    Barrio:    String,
  };

  // Relleno Array Putxeras
  dataPutxeras.forEach( row =>{
    aPutxeras = {Putxera: row[0],Nombre: row[1],Provincia: row[2],Poblacion: row[3],Barrio: row[4],};
    Putxeras.push(aPutxeras);
  });

//Obtener los datos de la VotacionB5
  var dataVotacionB5 = sVotacionB5.getDataRange().getValues();
  dataVotacionB5.shift();

  var VotacionB5es    = [];
  var aVotacionB5     = {
    Putxera:  Number,
    Nombre:   String,
    Poblacion:  String    
  };

  // Relleno Array VotacionB5
  dataVotacionB5.forEach( row =>{
    aVotacionB5 = {Putxera: row[0], Nombre: row[1], Poblacion: row[2]};
    VotacionB5es.push(aVotacionB5);
  });

//Obtener los datos de la VotacionG10
  var dataVotacionG10 = sVotacionG10.getDataRange().getValues();
  dataVotacionG10.shift();

  var VotacionG10es    = [];
  var aVotacionG10     = {
    Putxera:  Number,
    Nombre:   String,
    Poblacion:  String    
  };

  // Relleno Array VotacionG10
  dataVotacionG10.forEach( row =>{
    aVotacionG10 = {Putxera: row[0], Nombre: row[1], Poblacion: row[2]};
    VotacionG10es.push(aVotacionG10);
  });

//Obtener los datos de la VotacionG11
  var dataVotacionG11 = sVotacionG11.getDataRange().getValues();
  dataVotacionG11.shift();

  var VotacionG11es    = [];
  var aVotacionG11     = {
    Putxera:  Number,
    Nombre:   String,
    Poblacion:  String    
  };

  // Relleno Array VotacionG11
  dataVotacionG11.forEach( row =>{
    aVotacionG11 = {Putxera: row[0], Nombre: row[1], Poblacion: row[2]};
    VotacionG11es.push(aVotacionG11);
  });

//Obtener los datos de Clasificacion
  var dataClasificacion = sClasificacion.getDataRange().getValues();
  dataClasificacion.shift();

  var Clasificaciones    = [];
  var aClasificacion     = {
    Putxera:  Number,
    Nombre:   String,
    Poblacion:   String,
    Total:  Number,
    Presencia:  Number,
    Punto:  Number,
    Sabor:  Number,
    Basauri:  String,
    General:  Number
  };

  // Relleno Array Clasificacion
  dataClasificacion.forEach( row =>{
    aClasificacion = {Putxera: row[0], Nombre: row[1], Poblacion: row[2], Total: row[3], Presencia: row[4], Punto: row[5], Sabor: row[6], Basauri: row[7], General: row[8]};
    Clasificaciones.push(aClasificacion);
  });

//Obtener los datos de la Basauri
  var dataBasauri = sBasauri.getDataRange().getValues();
  dataBasauri.shift();

  var Basauris    = [];
  var aBasauri     = {
    Putxera:  Number,
    Nombre:   String,
    Poblacion:   String,
    Total:  Number,
    Presencia:  Number,
    Punto:  Number,
    Sabor:  Number,
    Basauri:  Number
  };

  // Relleno Array Basauri
  dataBasauri.forEach( row =>{
    aBasauri = {Putxera: row[0], Nombre: row[1], Poblacion: row[2], Total: row[3], Presencia: row[4], Punto: row[5], Sabor: row[6], Basauri: row[7]};
    Basauris.push(aBasauri);
  });

//Obtener los datos de la hoja Transparencia
  var dataTransparencia = sTransparencia.getDataRange().getValues();
  dataTransparencia.shift();

  var Transparencia    = [];
  var aTransparencia     = {
    Concepto:  String,
    Tipo:      String,
    Euros:     Number
  };

  // Relleno Array Transparencia
  dataTransparencia.forEach( row =>{
    aTransparencia = {Concepto: row[1], Tipo: row[2], Euros: row[3]};
    Transparencia.push(aTransparencia);
  });
  

//Carga la pagina icluyendo los sitios
  var template = HtmlService.createTemplateFromFile('Index');
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
  var output = template.evaluate();
  return output; 
}

function include( filename ){
  return HtmlService.createHtmlOutputFromFile( filename ).getContent();
}
