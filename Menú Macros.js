// =============================
// Barra de Menú Personalizada
// =============================


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Funciones')
    .addItem('Crear Pre-Acta','CrearPreActa')
    .addItem('Actualizar Formulas Pre-Actas', 'ActualizarFormulasPreActa')
    .addItem('Crear Acta Parcial', 'NuevaActaParcial')
    .addToUi();
}