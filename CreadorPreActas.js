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
// Funciones principales
// =============================  

// funcion para crear una nueva preacta
function CrearPreActa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cell = sheet.getActiveCell();
  const cRow = cell.getRow();
  const cCol = cell.getColumn();
  // Suponiendo que "Acta" es la hoja activa
  const Acta = sheet;
  const TituloHoja = Acta.getRange(FILA_PRESENTE_ACTA, COL_PRESENTE_ACTA).getValue();
  const NoActa = Acta.getRange(FILA_PRESENTE_ACTA, COL_PRESENTE_ACTA + 1).getValue();
  const NombreHoja = sheet.getRange(cRow, 1).getValue() + " " + TituloHoja + NoActa;
  const BaseNombre = sheet.getRange(cRow, 1).getValue() + " " + TituloHoja;

  let idArchivo = BuscarArchivoPreActaActual(TituloHoja + NoActa);

  if (BuscarHoja(NombreHoja)) {
    SpreadsheetApp.getUi().alert("La Pre-acta " + NombreHoja + " ya existe");
  } else {
    if (BuscarActas(BaseNombre)) {
      GenerarSiguienteActa(NombreHoja, NoActa, false, cRow, cCol, idArchivo);
      SpreadsheetApp.getUi().alert("Se creó " + NombreHoja + ", ya existen preactas del item.");
    } else {
      CopiarHoja(NombreHoja, NoActa, true, cRow, cCol, idArchivo);
    }
  }
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
  let filaTotal = celda.getRow() + 6;
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

// =============================
// funciones de busqueda
// =============================

// buscar hoja en el array global de hojas si existe retorna distinto de null mediante el nombre de la hoja
function BuscarHoja(nombreHoja) {
  const sheets = getPreActasArchivos();
  for (let sheet of sheets) {
    if (sheet.getName() === nombreHoja) {
      return true;
    }
  }
  return false;
} 

// Busca el archivo con el mismo nombre base ejemplo MiProyecto de la preacta con numero de corte ejemplo "MiProyecto.Corte No.1"

function BuscarArchivoPreActaActual(CorteObra) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHojaBase = ss.getName() + ".";
  const idCarpeta = getIdCarpeta();
  if (!idCarpeta) {
    throw new Error('No se pudo obtener el ID de la carpeta');
  }
  const carpeta = DriveApp.getFolderById(idCarpeta);
  const archivos = carpeta.getFiles();
  while (archivos.hasNext()) {
    const archivo = archivos.next();
    const nombreArchivo = archivo.getName();
    if (nombreArchivo === nombreHojaBase + CorteObra) {
      return archivo.getId();     
    }
  }
  // crear el archivo si no existe y retorna el id del archivo creado
  CrearArchivoPreActa();
  return BuscarArchivoPreActaActual(CorteObra);
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
// =============================
// funciones de obtención de filas y columnas dinámicas
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
  cell = sheet.createTextFinder("MEMORIA DE CALCULO").findNext();
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

// =============================
// funciones de obtención de hojas y archivos
// =============================

// crea un array con los hojas de los archivos con el nombre base del archivo principal
function getPreActasArchivos() {
  sheetAllHojas = []; // reinicia el array global
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHojaBase = ss.getName() + ".";
  const idCarpeta = getIdCarpeta();
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

// obtiene una hoja por su nombre
function getSheetByName(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name);
}

// obtiene el id de la carpeta de la hoja activa
function getIdCarpeta() {
  const archivo = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const carpetas = archivo.getParents();
  if (carpetas.hasNext()) {
    const carpeta = carpetas.next();
    return carpeta.getId();
  } else {
    throw new Error("El archivo no está en ninguna carpeta.");
  }
}

// borra todas las imagenes de una hoja
function BorrarImagenes(destino) {
  const images = destino.getImages();
  images.forEach(img => img.remove());
}

// =============================
// funciones de creación y copia de hojas
// =============================

// crear un archivo con el nombre de la preacta si existe retorna que el archivo ya existe
function CrearArchivoPreActa() { 
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // Validar que estamos en la hoja correcta
    if (!sheet) {
      throw new Error('No se pudo obtener la hoja activa');
    }
    
    const NoActa = sheet.getRange(FILA_PRESENTE_ACTA, COL_PRESENTE_ACTA + 1).getValue();
    const NombreHoja = sheet.getRange(FILA_PRESENTE_ACTA, COL_PRESENTE_ACTA).getValue() + NoActa;
    // obtiene el nombre del archio 
    const nombreHojaBase = ss.getName() + ".";
    const nombrePreActa = nombreHojaBase + NombreHoja;
    
    const idCarpeta = getIdCarpeta();
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

// Genera la siguiete aca 
function GenerarSiguienteActa(NombreHoja, NoActa, PrimeraActa, cRow, cCol, idArchivo) {
  CopiarHoja(NombreHoja, NoActa, PrimeraActa, cRow, cCol, idArchivo);
}

// Copial la hoja en en el archivo del corte al archivo correspondiente de idArchivo
function CopiarHoja(NombreHoja, NoActa, PrimeraActa, cRow, cCol, idArchivo) {
  if (!idArchivo) {
    throw new Error('El idArchivo es inválido o no se encontró el archivo de destino.');
  }
  const hojas = getPreActasArchivos();
  let hojaOriginal;
  let BaseNombre = NombreHoja.substring(0, NombreHoja.length - ("" + NoActa).length);
  let UltActa = UltimaActaDeItem(BaseNombre);
  if (UltActa === 0) {
    hojaOriginal = getSheetByName(NOMBRE_ACTA0);
  } else {
    hojaOriginal = hojas.find(h => h.getName() === (BaseNombre + UltActa));
    if (!hojaOriginal) {
      throw new Error('No se encontró la hoja original para copiar.');
    }
  }
  // Copiar hoja en el archivo correspondiente
  const ssDestino = SpreadsheetApp.openById(idArchivo);
  const HojaFuente = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let hojaNueva = hojaOriginal.copyTo(ssDestino).setName(BaseNombre + NoActa);
  EscribirEncabezado(hojaNueva, HojaFuente, cRow, cCol);
  AjustarTotales(hojaNueva, hojaOriginal, BaseNombre, PrimeraActa, NoActa, cRow, cCol);
}

function EscribirEncabezado(destino, fuente, fRow, fCol) {

  // Obtener la hoja de formato de acta
  // OPTENER EL NOMBRE DE LA HOJA DESTINO
  let nombreSheetDestino = destino.getName(); 
  let hoja = getPreActasArchivos();
  // busca la hoja destino en el array de hojas
  let sheet = hoja.find(h => h.getName() === nombreSheetDestino);

  if (!sheet) {
    throw new Error("No se encontró la hoja");
  }

  // busca la palabra "ITEM" en el formato de acta para definir la fila
  cell = sheet.createTextFinder("Ítem:").findNext();
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
  FILA_PREACTA_NUMERO = cell ? cell.getRow() + 1 : 3;  // Valor por defecto si no se encuentra
  cell = sheet.createTextFinder("MEMORIA DE CALCULO").findNext();
  COL_PREACTA_NUMERO = cell ? cell.getColumn() : 6; // Valor por defecto si no se encuentra

  // busca la palabra PERIODO ACTA para definir la fila de la fecha
  cell = sheet.createTextFinder("PERIODO ACTA:").findNext();
  FILA_PREACTA_FECHA = cell ? cell.getRow(): 7;  // Valor por defecto si no se encuentra
  COL_PREACTA_FECHA_I = cell ? cell.getColumn() + 1 : 7;  // Valor por defecto si no se encuentra
  COL_PREACTA_FECHA_F = COL_PREACTA_FECHA_I + 4;  // Valor por defecto si no se encuentra

  // busca la palabra SUBCONTRATISTA para definir la fila de subcontratista
  cell = sheet.createTextFinder("SUBCONTRATISTA:").findNext();
  FILA_PREACTA_SUBCONTRA = cell ? cell.getRow() : 6; // Valor por defecto si no se encuentra
  COL_PREACTA_SUBCONTRA = cell ? cell.getColumn() + 1 : 7; // Valor por defecto si no se encuentra

  //----------------------------------------------
  // Escribir ítem, descripción, unidad y cantidad
  //----------------------------------------------
  
  //obtener url de la hoja fuente
  let urlHojaFuente = fuente.getParent().getUrl();
  let nombreHojaFuente = fuente.getName();
  const X = '"';

  // encabezado del corte de obra
  destino.getRange(2,2).setFormula("=IMPORTRANGE(" + X + urlHojaFuente + X + "," + X + "'" + nombreHojaFuente + "'!A1" + X + ")");
  destino.getRange(2,10).setFormula("=IMPORTRANGE(" + X + urlHojaFuente + X + "," + X + "'" + nombreHojaFuente + "'!E4" + X + ")");
  destino.getRange(4,4).setFormula("=IMPORTRANGE(" + X + urlHojaFuente + X + "," + X + "'" + nombreHojaFuente + "'!C1" + X + ")");

  // encabezado de ítem, descripción, unidad y cantidad
  destino.getRange(FILA_PREACTA_ITEM, COL_PREACTA_ITEM).setFormula("=IMPORTRANGE(" + X + urlHojaFuente + X + "," + X + "'" + nombreHojaFuente + "'!A" + fRow + X + ")");
  destino.getRange(FILA_PREACTA_ITEM, COL_PREACTA_DESCRIPCION).setFormula("=IMPORTRANGE(" + X + urlHojaFuente + X + "," + X + "'" + nombreHojaFuente + "'!B" + fRow + X + ")");
  destino.getRange(FILA_PREACTA_ITEM, COL_PREACTA_UNIDAD).setFormula("=IMPORTRANGE(" + X + urlHojaFuente + X + "," + X + "'" + nombreHojaFuente + "'!C" + fRow + X + ")");
  destino.getRange(FILA_PREACTA_ITEM, COL_PREACTA_CANTIDAD).setFormula("=IMPORTRANGE(" + X + urlHojaFuente + X + "," + X + "'" + nombreHojaFuente + "'!D" + fRow + X + ")");
  let colCelda = numToCol(fCol); // o calcula la letra según la columna
  let ColSigCelda = numToCol(fCol + 1); // Ejemplo para columna de fecha
  // Escribir número de acta, fechas y subcontratista
  destino.getRange(FILA_PREACTA_NUMERO, COL_PREACTA_NUMERO).setFormula("=IMPORTRANGE(" + X + urlHojaFuente + X + "," + X + "'" + nombreHojaFuente + "'!" + colCelda + FILA_ACTA_CORTENo + X + ")");
  destino.getRange(FILA_PREACTA_FECHA, COL_PREACTA_FECHA_I).setFormula("=IMPORTRANGE(" + X + urlHojaFuente + X + "," + X + "'" + nombreHojaFuente + "'!" + colCelda + FILA_ACTA_FECHA + X + ")");
  destino.getRange(FILA_PREACTA_FECHA, COL_PREACTA_FECHA_F).setFormula("=IMPORTRANGE(" + X + urlHojaFuente + X + "," + X + "'" + nombreHojaFuente + "'!" + ColSigCelda + FILA_ACTA_FECHA + X + ")");
  destino.getRange(FILA_PREACTA_SUBCONTRA, COL_PREACTA_SUBCONTRA).setFormula("=IMPORTRANGE(" + X + urlHojaFuente + X + "," + X + "'" + nombreHojaFuente + "'!" + colCelda + FILA_ACTA_SUBCONTRA + X + ")");
}

function AjustarTotales(Destino, Origen, BaseNombre, PrimeraActa, NoActa, cRow, cCol) {
  // obtener url de la hoja destino
  let UrlHojaDestino = Destino.getParent().getUrl();
  let UrlHojaOrigen = Origen.getParent().getUrl();
  const X = '"';

  // Buscar texto y ajustar fórmulas
  let celdaDestino = Destino.createTextFinder("Menos (-) Cantidad Pagada Actas Anteriores").findNext();
  let coordCol = celdaDestino.getColumn();
  let coordRow = celdaDestino.getRow();
  // insertar fórmula en la celda siguiente si en la celda destino no es FORMATO CORTE
  if (Destino.getName() !== 'FORMATO CORTE') {
    Destino.getRange(coordRow, coordCol + 1).setFormula("=IMPORTRANGE(" + X + UrlHojaOrigen + X + "," + X + "'" + Origen.getName() + "'!" + numToCol(coordCol + 1) + (coordRow + 1) + X + ")");
    // borra el contenido de las celdas que tienen insertado contenido, en este caso fotos
    let celdaDestinoPhotos = Destino.createTextFinder("Croquis  y/o  Record  Fotográfico").findNext();
    let coordColPhoto = celdaDestinoPhotos.getColumn();
    let coordRowphoto = celdaDestinoPhotos.getRow() + 1;
    Destino.getRange(coordRowphoto, coordColPhoto, 3, 10).clearContent();
  }
  if (!PrimeraActa) {
    // borrar imágenes
    BorrarImagenes(Destino);

    // Obtener la fila a partir de coordRow hacia arriba hasta encontrar una celda con datos
    let startRow = coordRow - 1;
    let celdainicio = Destino.getRange(startRow, coordCol);
    var celda = celdainicio.getNextDataCell(SpreadsheetApp.Direction.UP);
    startRow = celda.getRow();

    SpreadsheetApp.getUi().alert("startRow: " + startRow);
    // Obtener la ultima columna de la hoja, y la celda que contenga "Descripción:"
    let lastCol = Destino.getLastColumn();
    let filaDesp = Destino.createTextFinder("Descripción:").findNext();
    let coordColDesp = filaDesp.getColumn();
    let coordRowDesp = filaDesp.getRow() + 2;

    // Colorear el rango de las filas que contiene datos
    let rango = Destino.getRange(coordRowDesp, coordColDesp, startRow - coordRowDesp + 1, lastCol - coordColDesp + 1);
    rango.setFontColor("#4285F4"); // Color de letra azul

  }
  // Insertar fórmula en la celda activa
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let Acta = ss.getActiveSheet();
  Acta.getRange(cRow, cCol).setFormula("=IMPORTRANGE(" + X + UrlHojaDestino + X + "," + X + "'" + BaseNombre + NoActa + "'!" + numToCol(coordCol + 1) + (coordRow + 2) + X + ")");
  Acta.getRange(cRow, cCol).setNumberFormat("0.00");
}
