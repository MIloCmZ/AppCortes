// =============================
// Barra de Menú Personalizada
// =============================


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Funciones')
    .addItem('Crear Pre-Acta','CrearPreActa')
    .addItem('Crear Acta Parcial', 'NuevaActaParcial')
    .addToUi();
}