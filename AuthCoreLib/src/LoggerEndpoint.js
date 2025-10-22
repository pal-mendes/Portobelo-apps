
// =========================
// File: LoggerEndpoint.gs - em AuthCoreLib (biblioteca)
// =========================

function getLoggerJs_() {
  // Lê o conteúdo cru do ficheiro logger_js.html (apenas JS, sem <script> tags)
  return HtmlService.createTemplateFromFile('logger_js').getRawContent();
}

function doGet(e) {
  const js = getLoggerJs_();
  return ContentService.createTextOutput(js)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
