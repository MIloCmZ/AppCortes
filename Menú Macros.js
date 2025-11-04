// =============================
// Barra de Menú Personalizada
// =============================


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Funciones')
    .addItem('Crear hoja de la Pre-Acta', 'crearArchivoPreActa')
    .addToUi();
}