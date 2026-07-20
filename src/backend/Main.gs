/**
 * Main.gs
 * Punto de entrada del Web App y utilidades para plantillas HTML.
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('frontend/index');
  return template.evaluate()
    .setTitle('Gestión de Asesores')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Permite incluir archivos HTML/CSS/JS dentro de una plantilla.
 * Uso dentro de un .html: <?!= include('css/styles'); ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
