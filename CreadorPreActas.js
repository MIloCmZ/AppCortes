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


// Inicializar las variables globales al cargar el script 
getCeldasFormatoActa();
getCeldasFormatoPreActa();

// =============================
// Funciones principales
// =============================  

// funcion de actualizar formulas de preacta, actualiza las formulas celdas de la columna activa
function ActualizarFormulasPreActa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const col = sheet.getActiveCell().getColumn();
  const fila1 = sheet.createTextFinder("Presente acta:").findNext().getRow()+4;
  const fila2 = sheet.createTextFinder("VALOR TOTAL OBRA EJECUTADA").findNext().getRow()-6;
  const FILA_ACTA_SUBCONTRA = sheet.createTextFinder("SUBCONTRATISTA:").findNext().getRow();
  // Actualiza las fórmulas en la columna activa
  const rango = sheet.getRange(fila1, col, fila2 - fila1 + 1);
  const Acta = sheet;
  const TituloHoja = Acta.getRange(FILA_PRESENTE_ACTA, COL_PRESENTE_ACTA).getValue();
  const NoActa = Acta.getRange(FILA_PRESENTE_ACTA, COL_PRESENTE_ACTA + 1).getValue();

  let idArchivo = BuscarArchivoPreActaActual(TituloHoja + NoActa);
  const hojasPreActa = getPreActasArchivos();

  for (let i = 0; i < rango.getNumRows(); i++) {
    const cell = rango.getCell(i + 1, 1);
    // Verificar si la celda no está vacía
    if (cell.getValue() > 0) {
      const cRow = cell.getRow();
      const cCol = cell.getColumn();
      const Ftilulo = sheet.getRange(FILA_ACTA_SUBCONTRA + 1, cCol).getValue();
      const NombreHoja = sheet.getRange(cRow, 1).getValue() + " " + Ftilulo;
      const BaseNombre = sheet.getRange(cRow, 1).getValue() + " " + TituloHoja;
      const Destino = ObtenerHoja(hojasPreActa, NombreHoja);
      const AnteriorActa = AnteriorActaDeItem(hojasPreActa,BaseNombre);
      let hojaOriginal = getSheetByName(NOMBRE_ACTA0);
      if (AnteriorActa !== null) {
        hojaOriginal = hojasPreActa.find(h => h.getName() === (BaseNombre + AnteriorActa));
        if (!hojaOriginal) {
          throw new Error('No se encontró la hoja original para copiar.');
        }
      }
      EscribirEncabezado(Destino, sheet, cRow, cCol);
      AjustarTotalesformulas(Destino, hojaOriginal, BaseNombre, NoActa, cRow, cCol);
    }
  }
}

// funcion para crear una nueva preacta
function CrearPreActa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cell = sheet.getActiveCell();
  const cRow = cell.getRow();
  const cCol = cell.getColumn();
  const Acta = sheet;
  const TituloHoja = Acta.getRange(FILA_PRESENTE_ACTA, COL_PRESENTE_ACTA).getValue();
  const NoActa = Acta.getRange(FILA_PRESENTE_ACTA, COL_PRESENTE_ACTA + 1).getValue();
  const NombreHoja = sheet.getRange(cRow, 1).getValue() + " " + TituloHoja + NoActa;
  const BaseNombre = sheet.getRange(cRow, 1).getValue() + " " + TituloHoja;

  let idArchivo = BuscarArchivoPreActaActual(TituloHoja + NoActa);

  const hojasPreActa = getPreActasArchivos();

  if (BuscarHoja(hojasPreActa, NombreHoja)) {
    SpreadsheetApp.getUi().alert("La Pre-acta " + NombreHoja + " ya existe");
  } else {
    if (BuscarActas(hojasPreActa, BaseNombre)) {
      GenerarSiguienteActa(hojasPreActa, NombreHoja, NoActa, cRow, cCol, idArchivo);
      SpreadsheetApp.getUi().alert("Se creó " + NombreHoja + ", ya existen preactas del item.");
    } else {
      CopiarHoja(hojasPreActa, NombreHoja, NoActa, cRow, cCol, idArchivo);
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

//busca hoja en el array global de hojas si existe retorna la hoja mediante el nombre de la hoja
function ObtenerHoja(sheets, nombreHoja) {
  return sheets.find(h => h.getName() === nombreHoja) || null;
}

// buscar hoja en el array global de hojas si existe retorna distinto de null mediante el nombre de la hoja
function BuscarHoja(sheets, nombreHoja) {
  return hojaExiste = sheets.some(h => h.getName() === nombreHoja);
}

// buscar si exsten actas para un item determinado
function BuscarActas(hojasPreActa ,BaseNombre) {
 return actaExiste = hojasPreActa.some(h => h.getName().toUpperCase().startsWith(BaseNombre.toUpperCase()));
}

// Busca el archivo con el mismo nombre base ejemplo MiProyecto de la preacta con numero de corte ejemplo "MiProyecto.Corte No.1"
// incluyendo las actas las hojas del archivo principal
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

  // inicializa el array global
  sheetAllHojas = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHojaBase = ss.getName();
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
      const hojas = ssArchivo.getSheets();
      sheetAllHojas.push(...hojas);
    }
  }
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

// función para obtener url de hoja y ajustar totales
function getUrlehojas(hoja) {
  // obtener url de la hoja desde https://docs.google.com/spreadsheets/d/ hasta /edit
  // ejemplo "https://docs.google.com/spreadsheets/d/1zJXc1Pxe-yh1krWC9HhCLpE-J1j8ObkuU6-9OsHx68k/edit"
  //obtener esto : 1zJXc1Pxe-yh1krWC9HhCLpE-J1j8ObkuU6-9OsHx68k
  let urlHoja = hoja.getParent().getUrl();
  let url = urlHoja.replace('https://docs.google.com/spreadsheets/d/', '');
  url = url.replace('/edit', '');
  return url;
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
function UltimaActaDeItem(sheets, baseNombre) {
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

// obtener el número de la acta anterior a la última creada para un item determinado, aunque los números no sean consecutivos
function AnteriorActaDeItem(sheets, baseNombre) {
  let numeros = [];
  for (let sheet of sheets) {
    if (sheet.getName().startsWith(baseNombre)) {
      let numStr = sheet.getName().replace(baseNombre, "");
      if (!isNaN(numStr) && numStr !== "") {
        numeros.push(parseInt(numStr, 10));
      }
    }
  }
  numeros = numeros.filter(n => !isNaN(n));
  numeros.sort((a, b) => b - a); // ordenar de mayor a menor
  return numeros.length >= 2 ? numeros[1] : null; // Retorna el penúltimo o null si no existe
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
function GenerarSiguienteActa(hojasPreActa, NombreHoja, NoActa, cRow, cCol, idArchivo) {
  CopiarHoja(hojasPreActa, NombreHoja, NoActa, cRow, cCol, idArchivo);
}

// Copial la hoja en en el archivo del corte al archivo correspondiente de idArchivo
function CopiarHoja(hojasPreActa, NombreHoja, NoActa, cRow, cCol, idArchivo) {
  if (!idArchivo) {
    throw new Error('El idArchivo es inválido o no se encontró el archivo de destino.');
  }
  const hojas = hojasPreActa
  let hojaOriginal;
  let BaseNombre = NombreHoja.substring(0, NombreHoja.length - ("" + NoActa).length);
  let UltActa = UltimaActaDeItem(hojas, BaseNombre);
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
  AjustarTotales(hojaNueva, hojaOriginal, BaseNombre, NoActa, cRow, cCol);
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
  let urlHojaFuente = getUrlehojas(fuente);
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

// Ajustar totales y fórmulas en la hoja destino
function AjustarTotales(Destino, Origen, BaseNombre, NoActa, cRow, cCol) {
  // obtener url de la hoja destino
  let UrlHojaDestino = getUrlehojas(Destino);
  let UrlHojaOrigen = getUrlehojas(Origen);
  const X = '"';

  // Buscar texto y ajustar fórmulas
  let celdaDestino = Destino.createTextFinder("Menos (-) Cantidad Pagada Actas Anteriores").findNext();
  let coordColDestino = celdaDestino.getColumn();
  let coordRowDestino = celdaDestino.getRow();

  let CeldaOrigen = Origen.createTextFinder("Menos (-) Cantidad Pagada Actas Anteriores").findNext();
  let coordColOrigen = CeldaOrigen.getColumn();
  let coordRowOrigen = CeldaOrigen.getRow();
  // insertar fórmula en la celda siguiente si en la celda destino no es FORMATO CORTE
  if (Origen.getName() !== 'FORMATO CORTE') {
    Destino.getRange(coordRowDestino, coordColDestino + 1).setFormula("=IMPORTRANGE(" + X + UrlHojaOrigen + X + "," + X + "'" + Origen.getName() + "'!" + numToCol(coordColOrigen + 1) + (coordRowOrigen + 1) + X + ")");
    // borra el contenido de las celdas que tienen insertado contenido, en este caso fotos
    let celdaDestinoPhotos = Destino.createTextFinder("Croquis  y/o  Record  Fotográfico").findNext();
    let coordColPhoto = celdaDestinoPhotos.getColumn();
    let coordRowphoto = celdaDestinoPhotos.getRow() + 1;
    Destino.getRange(coordRowphoto, coordColPhoto, 3, 10).clearContent();
    // borrar imágenes
    BorrarImagenes(Destino);

    // Obtener la fila a partir de coordRow hacia arriba hasta encontrar una celda con datos
    let startRow = coordRowDestino - 1;
    let celdainicio = Destino.getRange(startRow, coordColDestino);
    var celda = celdainicio.getNextDataCell(SpreadsheetApp.Direction.UP);
    startRow = celda.getRow();

    // Obtener la ultima columna de la hoja, y la celda que contenga "Descripción:"
    let lastCol = Destino.getLastColumn();
    let filaDesp = Destino.createTextFinder("Descripción:").findNext();
    let coordColDesp = filaDesp.getColumn();
    let coordRowDesp = filaDesp.getRow() + 2;

    if (coordRowDesp < coordRowDestino) {
    // Colorear el rango de las filas que contiene datos
    let rango = Destino.getRange(coordRowDesp, coordColDesp, startRow - coordRowDesp + 1, lastCol - coordColDesp + 1);
    rango.setFontColor("#4285F4"); // Color de letra azul
    }
  }
  // Insertar fórmula en la celda activa
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let Acta = ss.getActiveSheet();
  Acta.getRange(cRow, cCol).setFormula("=IMPORTRANGE(" + X + UrlHojaDestino + X + "," + X + "'" + BaseNombre + NoActa + "'!" + numToCol(coordColDestino + 1) + (coordRowDestino + 2) + X + ")");
  Acta.getRange(cRow, cCol).setNumberFormat("0.00");
}

function AjustarTotalesformulas(Destino, Origen,  BaseNombre, NoActa, cRow, cCol) {
  // obtener url de la hoja destino
  let UrlHojaDestino = getUrlehojas(Destino);
  let UrlHojaOrigen = getUrlehojas(Origen);
    const X = '"';

  // Buscar texto y ajustar fórmulas
  let celdaDestino = Destino.createTextFinder("Menos (-) Cantidad Pagada Actas Anteriores").findNext();
  let coordColDestino = celdaDestino.getColumn();
  let coordRowDestino = celdaDestino.getRow();

  let CeldaOrigen = Origen.createTextFinder("Subtotal Cantidad Acumulada Presente Acta").findNext();
  let coordColOrigen = CeldaOrigen.getColumn();
  let coordRowOrigen = CeldaOrigen.getRow();
  // insertar fórmula en la celda siguiente si en la celda destino no es FORMATO CORTE
  if (Origen.getName() !== 'FORMATO CORTE') {
    Destino.getRange(coordRowDestino, coordColDestino + 1).setFormula("=IMPORTRANGE(" + X + UrlHojaOrigen + X + "," + X + "'" + Origen.getName() + "'!" + numToCol(coordColOrigen + 1) + (coordRowOrigen + 1) + X + ")");
  }
  // Insertar fórmula en la celda activa
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let Acta = ss.getActiveSheet();
  Acta.getRange(cRow, cCol).setFormula("=IMPORTRANGE(" + X + UrlHojaDestino + X + "," + X + "'" + BaseNombre + NoActa + "'!" + numToCol(coordColDestino + 1) + (coordRowDestino + 2) + X + ")");
  Acta.getRange(cRow, cCol).setNumberFormat("0.00");
}