//CSV.gs em "Movimentosfinanceiros"
function importSemicolonCSV() {
  Logger.log("importSemicolonCSV()");
  var nomeFicheiro = "Movimentos financeiros.csv";

  // Se houver vários com o mesmo nome, escolhe o mais recente:
  var files = DriveApp.getFilesByName(nomeFicheiro);
  var file = null, last = 0;
  while (files.hasNext()) {
    var f = files.next();
    var t = f.getLastUpdated().getTime();
    if (t > last) { last = t; file = f; }
  }
  if (!file) {
    Logger.log("Ficheiro não encontrado: " + nomeFicheiro);
    SpreadsheetApp.getActive().toast("Ficheiro não encontrado: " + nomeFicheiro, "Importar CSV", 5);
    return;
  } else {
    Logger.log("ficheiro encontrado");
  }

  // Lê e faz parse ao CSV (ponto-e-vírgula)
  var csvData = file.getBlob().getDataAsString("UTF-8");
  var rows = Utilities.parseCsv(csvData, ";");
  if (!rows || rows.length === 0) {
    Logger.log("CSV sem dados");
    SpreadsheetApp.getActive().toast("CSV sem dados.", "Importar CSV", 5);
    return;
  }

  //Logger.log("rows.shift");
  //rows.shift(); // CORREÇÃO: Remove a primeira linha (cabeçalhos) da matriz do CSV

  // Processa: troca Number<->Memo quando Number começa por "split"
  // Índices (0-based): Account(0) Date(1) Payment(2) Number(3) Payee(4) Memo(5) Amount(6) C(7) Category(8) Tags(9)
  var IDX_NUMBER = 3;
  var IDX_MEMO   = 5;
  var IDX_CATEGORY = 8;
  
  var processed = rows.map(function (r, i) {
    Logger.log("row: i=" + i + ", r[IDX_CATEGORY]=" + r[IDX_CATEGORY] + ", r[IDX_MEMO]=" + r[IDX_MEMO] + ", r[IDX_NUMBER]=" + r[IDX_NUMBER]);
    if (i === 0) return r; // cabeçalho tal como vem
    // limpa "Associação Portobelo:" da Category 
    r[IDX_CATEGORY] = (r[IDX_CATEGORY] || "").replace(/^Associação Portobelo:\s*/i, "").trim();
    var number = (r[IDX_NUMBER] || "").trim();
    if (/^split\s*/i.test(number)) {
      var memo = (r[IDX_MEMO] || "").trim();
      var newRow = r.slice();
      // Number passa a ser o que estava no Memo (ex.: telefone)
      newRow[IDX_NUMBER] = memo;
      // Memo passa a ser o texto do Number mas sem o prefixo "split"
      //newRow[IDX_MEMO] = number.replace(/^split\s*/i, "").trim();
      newRow[IDX_MEMO] = number.trim();
      return newRow;
    }
    //Logger.log("return r");
    return r;
  });

  // Garante retângulo (caso haja linhas com comprimentos diferentes)
  var maxCols = processed.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
  Logger.log("maxCols=" + maxCols);

  processed = processed.map(function (r) {
    if (r.length < maxCols) return r.concat(new Array(maxCols - r.length).fill(""));
    return r;
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Transações");
  if (!sheet) throw new Error('A aba "Transações" não existe.');

  // Limpa e escreve
  var lastRow = sheet.getLastRow();
  Logger.log("lastRow=" + lastRow);
  if (lastRow > 0) sheet.getRange(2, 1, lastRow, maxCols).clearContent();

  sheet.getRange(2, 1, processed.length, maxCols).setValues(processed);
 
   // (Opcional) se havia mais linhas antes, garante que A..J abaixo fica limpo
  if (processed.length < lastRow) {
    sheet.getRange(processed.length + 2, 1, lastRow - processed.length, maxCols).clearContent();
  }

  // Opcional: auto-ajuste de colunas
  try { sheet.autoResizeColumns(1, maxCols); } catch (_) {}

  Logger.log("Importação concluída");
  SpreadsheetApp.getActive().toast("Importação concluída (" + processed.length + " linhas).", "Importar CSV", 5);
}
