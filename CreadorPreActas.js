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

//----------------------------
// Constantes para Actas
//----------------------------

// Obtener la hoja de formato de acta

let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_ACTA0);
if (!sheet) {
  throw new Error("No se encontró la hoja 'FORMATO CORTE'.");
}

// busca la palabra "ITEM" en el formato de acta para definir la fila
let cell = sheet.createTextFinder("Ítem:").findNext();
let FILA_PREACTA_ITEM = cell ? cell.getRow() : 10;  // Valor por defecto si no se encuentra
// busca la columna "Ítem"
cell = sheet.createTextFinder("Ítem:").findNext();
let COL_PREACTA_ITEM = cell ? cell.getColumn() + 1 : 3;  // Valor por defecto si no se encuentra
let COL_PREACTA_DESCRIPCION = COL_PREACTA_ITEM + 1;
// busca la columna "Unidad"
cell = sheet.createTextFinder("Unidad:").findNext();
let COL_PREACTA_UNIDAD = cell ? cell.getColumn() + 1 : 10;  // Valor por defecto si no se encuentra
//busca la columna "Cantidad"
cell = sheet.createTextFinder("Cantidad:").findNext();
let COL_PREACTA_CANTIDAD = cell ? cell.getColumn() + 1 : 12;  // Valor por defecto si no se encuentra

// busca la palabra MEMORIA DE CALCULO para definir la fila del número de preacta
cell = sheet.createTextFinder("MEMORIA DE CÁLCULO").findNext();
let FILA_PREACTA_NUMERO = cell ? cell.getRow() + 1 : 8;  // Valor por defecto si no se encuentra
cell = sheet.createTextFinder("MEMORIA DE CÁLCULO").findNext();
let COL_PREACTA_NUMERO = cell ? cell.getColumn() : 6; // Valor por defecto si no se encuentra

// busca la palabra PERIODO ACTA para definir la fila de la fecha
cell = sheet.createTextFinder("PERIODO ACTA:").findNext();
let FILA_PREACTA_FECHA = cell ? cell.getRow(): 10;  // Valor por defecto si no se encuentra
let COL_PREACTA_FECHA_I = cell ? cell.getColumn() + 1 : 7;  // Valor por defecto si no se encuentra
let COL_PREACTA_FECHA_F = COL_PREACTA_FECHA_I + 4;  // Valor por defecto si no se encuentra

// busca la palabra SUBCONTRATISTA para definir la fila de subcontratista
cell = sheet.createTextFinder("SUBCONTRATISTA:").findNext();
let FILA_PREACTA_SUBCONTRA = cell ? cell.getRow() : 9; // Valor por defecto si no se encuentra
let COL_PREACTA_SUBCONTRA = cell ? cell.getColumn() + 1 : 7; // Valor por defecto si no se encuentra

// Estas filas y columnas son fijas en el formato de acta son valores por defecto

// Busca la palabra Presente acta
// Obtener la hoja de formato de acta

sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_ACTA);
if (!sheet) {
  throw new Error("No se encontró la hoja 'CORTE DE OBRA'.");
}

cell = sheet.createTextFinder("Presente acta:").findNext();
const FILA_PRESENTE_ACTA = cell ? cell.getRow() : 7; // Valor por defecto si no se encuentra
const COL_PRESENTE_ACTA = cell ? cell.getColumn() + 1 : 4; // Valor por defecto si no se encuentra

// busca la palabra Subcontratista
cell = sheet.createTextFinder("SUBCONTRATISTA:").findNext();
const FILA_ACTA_SUBCONTRA = cell ? cell.getRow() : 7; // Valor por defecto si no se encuentra
const FILA_ACTA_CORTENo = FILA_ACTA_SUBCONTRA + 1;
const FILA_ACTA_FECHA = FILA_ACTA_SUBCONTRA + 2 ;

const NO_ACTA = 2;


// =============================
// Utilidades para hojas de cálculo
// =============================
function getSheetByName(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name);
}

function getAllSheets() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets();
}

// =============================
// Abrir Navegador de Hojas (simulación)
// =============================
function AbrirNavegadorHojas() {
  // En Apps Script no hay formularios como en VBA
  // Puedes crear un diálogo HTML si lo necesitas
  SpreadsheetApp.getUi().alert('Función NavegadorHojasCorte no migrada.');
}

// =============================
// Buscar si existe hoja
// =============================
function BuscarHoja(nombreHoja) {
  return getSheetByName(nombreHoja) !== null;
}

// =============================
// Buscar si existen actas para un ítem
// =============================
function BuscarActas(baseNombre) {
  const sheets = getAllSheets();
  for (let sheet of sheets) {
    if (sheet.getName().toUpperCase().startsWith(baseNombre.toUpperCase())) {
      return true;
    }
  }
  return false;
}

// =============================
// Buscar última acta de un ítem
// =============================
function UltimaActaDeItem(baseNombre) {
  const sheets = getAllSheets();
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

// =============================
// Crear nueva PreActa
// =============================
function CrearPreActa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cell = sheet.getActiveCell();
  const cRow = cell.getRow();
  const cCol = cell.getColumn();
  // Suponiendo que "Acta" es la hoja activa
  const Acta = sheet;
  const NoActa = Acta.getRange(FILA_PRESENTE_ACTA, COL_PRESENTE_ACTA + 1).getValue();
  const NombreHoja = sheet.getRange(cRow, 1).getValue() + " " + Acta.getRange(FILA_PRESENTE_ACTA, COL_PRESENTE_ACTA).getValue() + NoActa;
  const BaseNombre = sheet.getRange(cRow, 1).getValue() + " " + Acta.getRange(FILA_PRESENTE_ACTA, COL_PRESENTE_ACTA).getValue();

  if (BuscarHoja(NombreHoja)) {
    SpreadsheetApp.getUi().alert("La pre-acta " + NombreHoja + " ya existe");
  } else {
    if (BuscarActas(BaseNombre)) {
      SpreadsheetApp.getUi().alert("Se creó " + Acta.getName() + ", ya existen preactas del item.");
      GenerarSiguienteActa(NombreHoja, NoActa, false, cRow, cCol);
    } else {
      CopiarHoja(NombreHoja, NoActa, true, cRow, cCol);
    }
  }
}

// =============================
// Generar siguiente acta
// =============================
function GenerarSiguienteActa(NombreHoja, NoActa, PrimeraActa, cRow, cCol) {
  CopiarHoja(NombreHoja, NoActa, PrimeraActa, cRow, cCol);
}

// =============================
// Copiar hoja y preparar
// =============================
function CopiarHoja(NombreHoja, NoActa, PrimeraActa, cRow, cCol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hojaOriginal;
  let BaseNombre = NombreHoja.substring(0, NombreHoja.length - ("" + NoActa).length);
  let UltActa = UltimaActaDeItem(BaseNombre);
  if (UltActa === 0) {
    hojaOriginal = getSheetByName(NOMBRE_ACTA0);
  } else {
    hojaOriginal = getSheetByName(BaseNombre + UltActa);
  }
  // Copiar hoja
  let hojaNueva = hojaOriginal.copyTo(ss).setName(BaseNombre + NoActa);
  EscribirEncabezado(hojaNueva, ss.getActiveSheet(), cRow, cCol, NoActa);
  AjustarTotales(hojaNueva, hojaOriginal, BaseNombre, PrimeraActa, NoActa, cRow, cCol);
}



// =============================
// Escribir encabezado
// =============================
function EscribirEncabezado(destino, fuente, fRow, fCol,NoActa) {

// Obtener la hoja de formato de acta
// OPTENER EL NOMBRE DE LA HOJA DESTINO
sheet = destino.getName(); 

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
COL_PREACTA_CANTIDAD = cell ? cell.getColumn() + 1 : 10;  // Valor por defecto si no se encuentra

// busca la palabra MEMORIA DE CALCULO para definir la fila del número de preacta
cell = sheet.createTextFinder("MEMORIA DE CÁLCULO").findNext();
FILA_PREACTA_NUMERO = cell ? cell.getRow() + 1 : 2;  // Valor por defecto si no se encuentra
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

  //----------------------------------------------
  // Escribir ítem, descripción, unidad y cantidad
  //----------------------------------------------
  destino.getRange(FILA_PREACTA_ITEM, COL_PREACTA_ITEM).setFormula("='" + fuente.getName() + "'!A" + fRow);
  destino.getRange(FILA_PREACTA_ITEM, COL_PREACTA_DESCRIPCION).setFormula("='" + fuente.getName() + "'!B" + fRow);
  destino.getRange(FILA_PREACTA_ITEM, COL_PREACTA_UNIDAD).setFormula("='" + fuente.getName() + "'!C" + fRow);
  destino.getRange(FILA_PREACTA_ITEM, COL_PREACTA_CANTIDAD).setFormula("='" + fuente.getName() + "'!D" + fRow);

    let colCelda = numToCol(fCol); // o calcula la letra según la columna
    let ColSigCelda = numToCol(fCol + 1); // Ejemplo para columna de fecha

  // Escribir número de acta, fechas y subcontratista
  destino.getRange(FILA_PREACTA_NUMERO, COL_PREACTA_NUMERO).setFormula("='" + fuente.getName() + "'!" + colCelda + FILA_ACTA_CORTENo);
  destino.getRange(FILA_PREACTA_FECHA, COL_PREACTA_FECHA_I).setFormula("='" + fuente.getName() + "'!" + colCelda + FILA_ACTA_FECHA);
  destino.getRange(FILA_PREACTA_FECHA, COL_PREACTA_FECHA_F).setFormula("='" + fuente.getName() + "'!" + ColSigCelda+ FILA_ACTA_FECHA);
  destino.getRange(FILA_PREACTA_SUBCONTRA, COL_PREACTA_SUBCONTRA).setFormula("='" + fuente.getName() + "'!" + colCelda + FILA_ACTA_SUBCONTRA);
}

// Convierte número de columna a letra (A, B, ..., Z, AA, AB, ...)
function numToCol(n) {
  let s = "";
  while (n > 0) {
    let m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// =============================
// Ajustar totales y acumulados
// =============================
function AjustarTotales(destino, origen, BaseNombre, PrimeraActa, NoActa, cRow, cCol) {
  // Buscar texto y ajustar fórmulas
  let celdaDestino = destino.createTextFinder("Menos (-) Cantidad Pagada Actas Anteriores").findNext();
  let coordCol = celdaDestino.getColumn();
  let coordRow = celdaDestino.getRow();
  if (celdaDestino.name !== 'FORMATO CORTE') {
    destino.getRange(coordRow, coordCol + 1).setFormula("='" + origen.getName() + "'!" + numToCol(coordCol + 1) + (coordRow - 1));
  }
  if (!PrimeraActa) {
    // borrar imágenes
    BorrarImagenes(destino);
    // Colorear celdas
    let LastColumn = destino.getLastColumn();
    let lastcell = destino.getRange(coordRow,LastColumn);
    // Colorear las filas en el rango 17,2 a lastcell que contengan datos dejar para despues
  }
  // Insertar fórmula en la celda activa
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let Acta = ss.getActiveSheet();
  Acta.getRange(cRow, cCol).setFormula("='" + BaseNombre + NoActa + "'!" + numToCol(coordCol + 1) + (coordRow + 2) );
  Acta.getRange(cRow, cCol).setNumberFormat("0.00");
}

// =============================
// Mostrar todas las hojas
// =============================
function MostrarPreActas() {
  const sheets = getAllSheets();
  for (let sheet of sheets) {
    sheet.showSheet();
  }
}

// =============================
// Borrar imágenes (no soportado en Sheets)
// =============================
function BorrarImagenes(destino) {
  const images = destino.getImages();
  images.forEach(img => img.remove());
}

// =============================
// Nueva Acta Parcial
// =============================

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

// =============================
// Renombrar hojas quitando punto
// =============================
function RenombrarHojas_QuitarPunto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const Caracter = ".";
  const Cambio = "";
  for (let i = NO_ACTA - 1; i < sheets.length; i++) {
    let ws = sheets[i];
    let nombreActual = ws.getName();
    let parte1 = "", parte2 = "";
    if (nombreActual.indexOf(" ") > 0) {
      parte1 = nombreActual.split(" ")[0];
      parte2 = nombreActual.substring(parte1.length + 1);
    } else {
      parte1 = nombreActual;
      parte2 = "";
    }
    if (parte1.indexOf(Caracter) > 0) {
      parte1 = parte1.replace(Caracter, Cambio);
      let nuevoNombre = (parte1 + " " + parte2).trim();
      if (nuevoNombre !== ws.getName()) {
        try {
          ws.setName(nuevoNombre);
        } catch (e) {
          SpreadsheetApp.getUi().alert("No se pudo renombrar la hoja: " + ws.getName());
        }
      }
    }
  }
  SpreadsheetApp.getUi().alert("Renombrado completado.");
}

// =============================
// Fin de migración
// =============================
