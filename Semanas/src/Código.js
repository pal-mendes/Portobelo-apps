/** Code.gs (Google Apps Script) **/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Calculadora do Número da Semana')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

