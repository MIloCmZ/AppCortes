// Navegador de Hojas - Apps Script
// Migración desde VBA Excel

// constantes para ventana HTML
var FILA_DES_ACTA = 10;
var COL_DES_ACTA = 4; 
var COL_UNIDAD = 6;
var COL_TOTAL = 8;
// fin constantes

// Funciones para el Navegador de Hojas
function mostrarNavegadorHojas() {
  var html = HtmlService.createHtmlOutputFromFile('NavegadorHojasHTML')
    .setWidth(900)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Navegador de Hojas');
}
// Obtener datos de las hojas
function obtenerHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var hojas = [];
  sheets.forEach(function(sheet) {
    hojas.push({
      nombre: sheet.getName(),
      descripcion: sheet.getRange(FILA_DES_ACTA,COL_DES_ACTA).getValue(), // FILA_DES_ACTA, COL_DES_ACTA
      unidad: buscarTextoEnHoja(sheet, 'Unidad:'),
      total: buscarTextoEnHoja(sheet, 'Total Cantidad A Pagar presente Acta')
    });
  });
  return hojas;
}
// Buscar texto en hoja y devolver valor de la celda a la derecha
function buscarTextoEnHoja(sheet, texto) {
  var finder = sheet.createTextFinder(texto).findNext();
  if (finder) {
    return sheet.getRange(finder.getRow(), finder.getColumn() + 1).getValue();
  }
  return '';
}
// Navegar a hoja específica por nombre
function irAHoja(nombre) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nombre);
  if (sheet) {
    ss.setActiveSheet(sheet);
    return true;
  }
  return false;
}
// Mostrar u ocultar hojas por nombre
function mostrarOcultarHojas(nombres, mostrar) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  nombres.forEach(function(nombre) {
    var sheet = ss.getSheetByName(nombre);
    if (sheet) {
      if (mostrar) {
        sheet.showSheet();
      } else {
        sheet.hideSheet();
      }
    }
  });
}

// Copiar hojas seleccionadas a un nuevo libro y devolver URL del nuevo libro
// para que el usuario pueda guardarlo donde desee en Google Drive
function copiarHojas(nombres) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var idCarpetaDestino = '1jZ3L8iudC52V68S3PLoZmbdk3F98KJQH';
  var carpetaDestino = DriveApp.getFolderById(idCarpetaDestino);
  //optiene nombre del archivo original para crear nombre consecutivo
  var nombreBase = ss.getName();
  var nombreConsecutivo = obtenerNombreConsecutivo(nombreBase);
    // Crear nuevo libro en la carpeta destino
  var nuevoLibro = SpreadsheetApp.create(nombreConsecutivo);  
  nombres.forEach(function(nombre) {
    var sheet = ss.getSheetByName(nombre);
    if (sheet) {
      sheet.copyTo(nuevoLibro).setName(nombreConsecutivo);
    }
  });
  // mueve el nuevo libro a la carpeta destino
  var archivoNuevo = DriveApp.getFileById(nuevoLibro.getId());
  archivoNuevo.moveTo(carpetaDestino); 
  // renombrar el nuevo libro con nombre consecutivo
  return nuevoLibro.getUrl();
}

// Busca archivos en la carpeta destino y retorna un nombre consecutivo no repetido
function obtenerNombreConsecutivo(baseNombre) {
  var idCarpetaDestino = '1jZ3L8iudC52V68S3PLoZmbdk3F98KJQH';
  var carpetaDestino = DriveApp.getFolderById(idCarpetaDestino);
  var archivos = carpetaDestino.getFiles();
  var maxConsecutivo = 0;
  // Buscar el mayor número consecutivo en los archivos existentes 
  // que coincidan con el patrón "baseNombre X.xlsx", si no existe, iniciar en 0
  while (archivos.hasNext()) {
    var archivo = archivos.next();
    var nombreArchivo = archivo.getName();
    var regex = new RegExp('^' + baseNombre + ' (\\d+)\\.xlsx$');
    var match = nombreArchivo.match(regex);
    if (match && match[1]) {
      var numero = parseInt(match[1], 10);
      if (numero > maxConsecutivo) {
        maxConsecutivo = numero;
      }
    }
  }
  // Retornar el nuevo nombre con el siguiente número consecutivo
  return baseNombre + ' ' + (maxConsecutivo + 1) + '.xlsx';
}

// Fin Funciones para el Navegador de Hojas
// Fin Navegador de Hojas - Apps Script
// Migración desde VBA Excel