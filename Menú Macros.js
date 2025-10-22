// =============================
// Barra de Menú Personalizada
// =============================


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Funciones')
    .addItem('Crear Pre-Acta', 'CrearPreActa')
    .addItem('Crear Nueva Acta Parcial', 'NuevaActaParcial')
    .addItem('Abrir Navegador', 'mostrarNavegadorHojas')
    .addToUi();
}