const { execSync } = require('child_process');
const fs = require('fs');

// IDs de tus hojas de cálculo. Puedes obtenerlos de la URL de cada hoja.
const spreadsheetIds = [
  '1yRN1cWL81nbGNIqlJ_BV8lEaVpmy7TOtXIMf9ej64dGtObqbeHTF9hqP', // Yuldaima hoja de calculo
  '169VCvBFwJRLouWVAWYMgh87lOuJggfitZcGlM9QG_Heb4by1txDYkN6e', // Floresta hoja de calculo 
  '1e8fRGqapp16MWoke2vNoleDfZ5d2xgoPQPHt4zCFHqxw6nL0QMVjoV_t', // Nazareth hoja de calculo
  '1hQPbNFadGD6OKDt7ja9gvWJQ3uSa_DlTsW8pwCvGlPKFNmvujCbr4jKy'  // Villa Clara hoja de calculo

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

