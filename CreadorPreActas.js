// Migración de CreadorPreActas.bas a Google Apps Script
// Autor: migración automática
// Fecha: 2025-10-06

// =============================
// Constantes
// =============================

//----------------------------------------------------------------------------
// Constantes para PreActas, valores de fila y columna definidos dinámicamente
//----------------------------------------------------------------------------
// Se asume que el formato de acta tiene ciertas palabras clave para localizar filas y columnas
// Si no se encuentran, se usan valores por defecto
//----------------------------------------------------------------------------

// Nombre del acta base
const NOMBRE_ACTA = "CORTE DE OBRA";
const NOMBRE_ACTA0 = "FORMATO CORTE";
// Variables globales para filas y columnas de PreActa
let FILA_PREACTA_ITEM;
let COL_PREACTA_ITEM;
let COL_PREACTA_DESCRIPCION;
let COL_PREACTA_UNIDAD;
let COL_PREACTA_CANTIDAD;
let FILA_PREACTA_NUMERO;
let COL_PREACTA_NUMERO;
let FILA_PREACTA_FECHA;
let COL_PREACTA_FECHA_I;
let COL_PREACTA_FECHA_F;
let FILA_PREACTA_SUBCONTRA;

// Variables globales para filas y columnas de Acta
let FILA_PRESENTE_ACTA;
let COL_PRESENTE_ACTA;
let FILA_ACTA_SUBCONTRA;
let FILA_ACTA_CORTENo;
let FILA_ACTA_FECHA;
let NO_ACTA;

// Array global para almacenar las hojas de los archivos creados
let sheetAllHojas= [];

// Inicializar las variables globales al cargar el script 
getCeldasFormatoActa();
getCeldasFormatoPreActa();

// =============================
// Funciones
// =============================  


// Función para obtener las filas y columnas dinámicamente
function getCeldasFormatoPreActa() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_ACTA0);
  if (!sheet) {
    throw new Error("No se encontró la hoja 'FORMATO CORTE'.");
  }
  // busca la palabra "ITEM" en el formato de acta para definir la fila
  let cell = sheet.createTextFinder("Ítem:").findNext();
  FILA_PREACTA_ITEM = cell ? cell.getRow() : 10;  // Valor por defecto si no se encuentra
  // busca la columna "Ítem"
  cell = sheet.createTextFinder("Ítem:").findNext();
  COL_PREACTA_ITEM = cell ? cell.getColumn() + 1 : 3;  // Valor por defecto si no se encuentra
  COL_PREACTA_DESCRIPCION = COL_PREACTA_ITEM + 1;
  // busca la columna "Unidad"
  cell = sheet.createTextFinder("Unidad:").findNext();
  COL_PREACTA_UNIDAD = cell ? cell.getColumn() + 1 : 10;  // Valor por defecto si no se encuentra
  //busca la columna "Cantidad"
  cell = sheet.createTextFinder("Cantidad:").findNext();
  COL_PREACTA_CANTIDAD = cell ? cell.getColumn() + 1 : 12;  // Valor por defecto si no se encuentra

  // busca la palabra MEMORIA DE CALCULO para definir la fila del número de preacta
  cell = sheet.createTextFinder("MEMORIA DE CÁLCULO").findNext();
  FILA_PREACTA_NUMERO = cell ? cell.getRow() + 1 : 8;  // Valor por defecto si no se encuentra
  cell = sheet.createTextFinder("MEMORIA DE CÁLCULO").findNext();
  COL_PREACTA_NUMERO = cell ? cell.getColumn() : 6; // Valor por defecto si no se encuentra

  // busca la palabra PERIODO ACTA para definir la fila de la fecha
  cell = sheet.createTextFinder("PERIODO ACTA:").findNext();
  FILA_PREACTA_FECHA = cell ? cell.getRow(): 10;  // Valor por defecto si no se encuentra
  COL_PREACTA_FECHA_I = cell ? cell.getColumn() + 1 : 7;  // Valor por defecto si no se encuentra
  COL_PREACTA_FECHA_F = COL_PREACTA_FECHA_I + 4;  // Valor por defecto si no se encuentra

  // busca la palabra SUBCONTRATISTA para definir la fila de subcontratista
  cell = sheet.createTextFinder("SUBCONTRATISTA:").findNext();
  FILA_PREACTA_SUBCONTRA = cell ? cell.getRow() : 9; // Valor por defecto si no se encuentra
  COL_PREACTA_SUBCONTRA = cell ? cell.getColumn() + 1 : 7; // Valor por defecto si no se encuentra
}

// Función para obtener las filas y columnas dinámicamente para el acta
function getCeldasFormatoActa() {

  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_ACTA);
  if (!sheet) {
    throw new Error("No se encontró la hoja 'CORTE DE OBRA'.");
  }

  cell = sheet.createTextFinder("Presente acta:").findNext();
  FILA_PRESENTE_ACTA = cell ? cell.getRow() : 7; // Valor por defecto si no se encuentra
  COL_PRESENTE_ACTA = cell ? cell.getColumn() + 1 : 4; // Valor por defecto si no se encuentra

  // busca la palabra Subcontratista
  cell = sheet.createTextFinder("SUBCONTRATISTA:").findNext();
  FILA_ACTA_SUBCONTRA = cell ? cell.getRow() : 7; // Valor por defecto si no se encuentra
  FILA_ACTA_CORTENo = FILA_ACTA_SUBCONTRA + 1;
  FILA_ACTA_FECHA = FILA_ACTA_SUBCONTRA + 2 ;

  NO_ACTA = 2;

}


// optiene el id de la carpeta de la hoja activa
function obtenerIdCarpeta() {
  const archivo = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const carpetas = archivo.getParents();
  if (carpetas.hasNext()) {
    const carpeta = carpetas.next();
    return carpeta.getId();
  } else {
    throw new Error("El archivo no está en ninguna carpeta.");
  }
}

// crear un archivo con el nombre de la preacta si existe retorna que el archivo ya existe
function crearArchivoPreActa() { 
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // Validar que estamos en la hoja correcta
    if (!sheet) {
      throw new Error('No se pudo obtener la hoja activa');
    }
    
    const NoActa = sheet.getRange(FILA_PRESENTE_ACTA, COL_PRESENTE_ACTA + 1).getValue();
    const NombreHoja = sheet.getRange(FILA_PRESENTE_ACTA, COL_PRESENTE_ACTA).getValue() + NoActa;
    // obtiene el nombre del archivo 
    const nombreHojaBase = ss.getName() + ".";
    const nombrePreActa = nombreHojaBase + NombreHoja;
    
    const idCarpeta = obtenerIdCarpeta();
    if (!idCarpeta) {
      throw new Error('No se pudo obtener el ID de la carpeta');
    }
    
    const carpeta = DriveApp.getFolderById(idCarpeta);
    const archivos = carpeta.getFilesByName(nombrePreActa);
    
    // Verificar si el archivo ya existe
    if (archivos.hasNext()) {
      const archivoExistente = archivos.next();
      Logger.log(`Archivo ya existe: ${nombrePreActa}`);
      SpreadsheetApp.getUi().alert(`El archivo "${nombrePreActa}" ya existe.`);
      return archivoExistente.getId();
    }
    
    // Crear el archivo CORRECTAMENTE - usando SpreadsheetApp.create()
    const nuevoArchivo = SpreadsheetApp.create(nombrePreActa);
    const archivoDrive = DriveApp.getFileById(nuevoArchivo.getId());
    
    // Mover a la carpeta destino
    archivoDrive.moveTo(carpeta);
    
    Logger.log(`Nuevo archivo creado: ${nombrePreActa}`);
    return nuevoArchivo.getId();
    
  } catch (error) {
    Logger.log(`Error en crearArchivoPreActa: ${error.toString()}`);
    throw error; // Relanzar el error para manejo externo
  }
}

// crea un array con los hojas de los archivos con el nombre base del archivo principal
function getPreActasArchivos() {
  sheetAllHojas = []; // reinicia el array global
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHojaBase = ss.getName() + ".";
  const idCarpeta = obtenerIdCarpeta();
  if (!idCarpeta) {
    throw new Error('No se pudo obtener el ID de la carpeta');
  }
  const carpeta = DriveApp.getFolderById(idCarpeta);
  const archivos = carpeta.getFiles();
  
  while (archivos.hasNext()) {
    const archivo = archivos.next();
    const nombreArchivo = archivo.getName();
    if (nombreArchivo.startsWith(nombreHojaBase)) {
      const idArchivo = archivo.getId();
      const ssArchivo = SpreadsheetApp.openById(idArchivo);
      const hojas = ssArchivo.getSheets(); // retorna las hojas del archivo
      sheetAllHojas.push(...hojas); // agrega las hojas al array global
    }
  }
  // retorna el array de todas las hojas encontradas 
  return sheetAllHojas;
}

// buscar si exsten actas para un item determinado
function BuscarActas(baseNombre) {
  const sheets = getPreActasArchivos();
  for (let sheet of sheets) {
    if (sheet.getName().toUpperCase().startsWith(baseNombre.toUpperCase())) {
      return true;
    }
  }
  return false;
}

// obtener el número de la última acta creada para un item determinado
function UltimaActaDeItem(baseNombre) {
  const sheets = getPreActasArchivos();
  let maxNum = 0;
  for (let sheet of sheets) {
    if (sheet.getName().startsWith(baseNombre)) {
      let numStr = sheet.getName().replace(baseNombre, "");
      if (!isNaN(numStr)) {
        let n = parseInt(numStr, 10);
        if (n > maxNum) maxNum = n;
      }
    }
  }
  return maxNum;
}

// Convierte un número de columna a letra de columna (1 -> A, 27 -> AA)
function numToCol(n) {
  let s = "";
  while (n > 0) {
    let m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}




// Crear una nueva acta parcial en la hoja activa

function NuevaActaParcial() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const Acta = ss.getActiveSheet();
  // Buscar "ACUMULADO"
  let celda = Acta.createTextFinder("ACUMULADO").findNext();
  if (!celda) {
    SpreadsheetApp.getUi().alert("No se encontró la palabra 'ACUMULADO'.");
    return;
  }
  let ColAcumulado = celda.getColumn();
  let RowAcumulado = celda.getRow();
  // Insertar 2 columnas a la izquierda
  Acta.insertColumnsBefore(ColAcumulado, 2);
  // Buscar "VALOR TOTAL OBRA EJECUTADA"
  celda = Acta.createTextFinder("VALOR TOTAL OBRA EJECUTADA").findNext();
  if (!celda) {
    SpreadsheetApp.getUi().alert("No se encontró 'VALOR TOTAL OBRA EJECUTADA'.");
    return;
  }
  let filaTotal = celda.getRow() + 5;
  // copia la formula de la columna anterior de colacumado a la columna de colacumado + 1
  Acta.getRange(11, ColAcumulado - 1, filaTotal - 11, 1).copyTo(Acta.getRange(11, ColAcumulado + 1, filaTotal - 11, 1));
  // Copiar formato de la segunda columna anterior de colacumado -2 a la columna de colacumado
  Acta.getRange(11, ColAcumulado - 2, filaTotal - 11, 2).copyTo(Acta.getRange(11, ColAcumulado, filaTotal - 11, 2), { formatOnly: true });
  // copia la fecha de las dos columnas anteriores de colacumulado
  Acta.getRange(7, ColAcumulado-2,4,2).copyTo(Acta.getRange(7, ColAcumulado,4,2));
  // Aumentar número de acta
  let Nacta = Acta.getRange(7, 5).getValue() + 1;
  Acta.getRange(7, 5).setValue(Nacta);
  // Ajustar ancho de columna
  Acta.setColumnWidth(ColAcumulado + 1, 144);
  // Actualizar encabezado de corte
  Acta.getRange(RowAcumulado, ColAcumulado).setValue(Acta.getRange(FILA_ACTA_SUBCONTRA, 4).getValue() + Acta.getRange(FILA_ACTA_SUBCONTRA, 5).getValue());
  SpreadsheetApp.getUi().alert("Nueva acta parcial creada.");

}
