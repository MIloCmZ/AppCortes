const { execSync } = require('child_process');
const fs = require('fs');

const spreadsheetIds = [
  // hoja prueba
  //'1sa75xQjpMKRfXi-jQbokkKWrNjnK_ONLhGOAWKf224Xhqe0o-vMbOkhq'
  // IDs de las hojas de cálculo
  // Salon comunal Yuldaima
  '1yRN1cWL81nbGNIqlJ_BV8lEaVpmy7TOtXIMf9ej64dGtObqbeHTF9hqP',
  // Salon comunal Floresta
  '169VCvBFwJRLouWVAWYMgh87lOuJggfitZcGlM9QG_Heb4by1txDYkN6e',
  // Salon comunal Nazareth
  '1e8fRGqapp16MWoke2vNoleDfZ5d2xgoPQPHt4zCFHqxw6nL0QMVjoV_t',
  // Salon comunal Villa Clara
  '1hQPbNFadGD6OKDt7ja9gvWJQ3uSa_DlTsW8pwCvGlPKFNmvujCbr4jKy',
  // Salon comunal Santa Teresa
  '1SDj56SKZsRmSypFFSsGF9wY9uHheTxW2coBlSXpm0m69pWT5aVh5PDbD',
  // Salon comunal La Libertad
  '1BSd8o8EurIZXNSfhmvCDkGQn-4HuWXJ9ZVk7HN3msHwWF_8zIS8oExd6',
  // Salon comunal Boyaca
  '1yBKXNgr5nmy2h14pBRzNrkptx7YxP77TdY77WtqPTjL4DV3RXDuMoJEs',
  // Salon comunal Simon Bolivar
  '1sdnQRC4nmhjQRCO3Ad3yIyIZCSqX7dUmtfHvXMH9KLyQgMcpqdZahHya',
  // Salon comunal Murillo
  '1dGSKDqBjlsUtpNRLLPm0LQAchz_NlId93IFdGprusu9pSf_0l8m6Bkjr'
];

function syncToAllSheets() {
  spreadsheetIds.forEach((scriptId) => {
    // La variable scriptId ahora se usa dentro del bucle donde fue declarada
    // muestra el nombre del archivo que se está sincronizando
    
    console.log(`\nSincronizando con la hoja de cálculo ID: ${scriptId}`);

    // Paso 1: Genera el archivo .clasp.json temporal
    const claspConfig = {
      scriptId: scriptId,
      rootDir: '.', // La raíz de tu proyecto
    };
    fs.writeFileSync('.clasp.json', JSON.stringify(claspConfig, null, 2));

    // Paso 2: Ejecuta clasp push
    try {
      execSync('clasp push', { stdio: 'inherit' });
      console.log(`Éxito al sincronizar con ${scriptId}`);
    } catch (error) {
      console.error(`Error al sincronizar con ${scriptId}:`, error.message);
    }
  });

  // Paso 3: Limpia el archivo .clasp.json para evitar conflictos
  fs.unlinkSync('.clasp.json');
  console.log('\nFinalizada la sincronización con todas las hojas de cálculo.');
}

syncToAllSheets();

