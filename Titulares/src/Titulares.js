
// =========================
// Titulares.gs
// Utilizado para fórmulas a usar no spreadsheet
// =========================


function getTitularesColumns_(sheet) {
  const headerRow = 6;
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};

  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  
  // Mapeamento dos nomes exatos das colunas para o seu número (1-based)
  const targets = ["€", "Adesão", "T0", "T1", "T2", "Fim", "Quota", "Jóia", "Saldo"];
  headers.forEach((label, index) => {
    if (targets.includes(label)) {
      map[label] = index + 1;
    }
  });
  return map;
}


/**
 * Gatilho simples que monitoriza edições no documento.
 */
function onEdit(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // 1. Edição na aba Titulares (Apenas a partir da linha 7)
  if (sheetName === "Titulares" && row >= 7) {
    const cols = getTitularesColumns_(sheet);
    const editedCol = e.range.getColumn();   

    // Verifica se a coluna editada é uma das que afeta os cálculos
    const influenceCols = [cols["€"], cols["Adesão"], cols["T0"], cols["T1"], cols["T2"], cols["Fim"]];
    if (influenceCols.includes(editedCol)) {
      updateTitularesRowValues_(sheet, row, cols);
    }
  } 
  
  // 2. Edição na aba Quotas (Intervalo A2:D3)
  else if (sheetName === "Quotas" && row >= 2 && row <= 3 && col <= 4) {
    updateAllTitularesValues();
  }
}

/**
 * Adiciona um menu manual para forçar a sincronização de todos os valores.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Associação Portobelo')
    .addItem('Recalcular Quotas/Saldos (Titulares)', 'updateAllTitularesValues')
    .addToUi();
}

/**
 * Determina o ano limite para o cálculo de quotas com base na coluna "Fim".
 */
function getAnoReferencia_(valorFim) {
  const anoAtual = new Date().getFullYear();
  if (!valorFim) return anoAtual;

  let anoFim;
  if (valorFim instanceof Date) {
    anoFim = valorFim.getFullYear();
  } else {
    anoFim = parseInt(valorFim);
  }
  
  // Se o ano de fim for válido, o cálculo deve parar nesse ano (até 31 de dezembro)
  return (!isNaN(anoFim) && anoFim > 0) ? Math.min(anoAtual, anoFim) : anoAtual;
}

/**
 * Atualiza os valores estáticos de uma única linha na aba Titulares.
 */
function updateTitularesRowValues_(sheet, row, cols) {
  if (!cols) cols = getTitularesColumns_(sheet);

  const ss = sheet.getParent();
  const quotasSheet = ss.getSheetByName("Quotas");
  const matrizQuotas = quotasSheet.getRange("A2:D3").getValues();
  
  // Captura os dados necessários usando o mapeamento
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const valorPago = rowData[cols["€"] - 1];
  const dataAdesao = rowData[cols["Adesão"] - 1];
  const t0 = rowData[cols["T0"] - 1];
  const t1 = rowData[cols["T1"] - 1];
  const t2 = rowData[cols["T2"] - 1];
  const dataFim = rowData[cols["Fim"] - 1];

  let anoReferencia = getAnoReferencia_(dataFim);

  const quota = CALCULAR_QUOTA(t0, t1, t2, anoReferencia, matrizQuotas);
  const joia = CALCULAR_JOIA(t0, t1, t2, dataAdesao, matrizQuotas);
  // Nota: A função CALCULAR_SALDO deve ser ajustada para aceitar o anoReferencia calculado
  const saldo = CALCULAR_SALDO(valorPago, t0, t1, t2, dataAdesao, anoReferencia, matrizQuotas);

  // Escrita nas colunas de destino
  sheet.getRange(row, cols["Quota"]).setValue(quota);
  sheet.getRange(row, cols["Jóia"]).setValue(joia);
  sheet.getRange(row, cols["Saldo"]).setValue(saldo);
}

/**
 * Atualiza todas as linhas da tabela tblTitulares de uma só vez (Batch Update).
 */
function updateAllTitularesValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Titulares");
  const cols = getTitularesColumns_(sheet);
  const matrizQuotas = ss.getSheetByName("Quotas").getRange("A2:D3").getValues();
  
  const startRow = 7;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn()).getValues();
  const anoCivil = new Date().getFullYear();
  
  const output = data.map(row => {
    const anoRef = getAnoReferencia_(row[cols["Fim"] - 1]);
    const q = CALCULAR_QUOTA(row[cols["T0"]-1], row[cols["T1"]-1], row[cols["T2"]-1], anoCivil, matrizQuotas);
    const j = CALCULAR_JOIA(row[cols["T0"]-1], row[cols["T1"]-1], row[cols["T2"]-1], row[cols["Adesão"]-1], matrizQuotas);
    const s = CALCULAR_SALDO(row[cols["€"]-1], row[cols["T0"]-1], row[cols["T1"]-1], row[cols["T2"]-1], row[cols["Adesão"]-1], anoRef, matrizQuotas);
    return [q, j, s];
  });

  // Escreve os resultados nas colunas corretas (Quota é a primeira das 3 consecutivas)
  sheet.getRange(startRow, cols["Quota"], output.length, 1).setValues(output.map(r => [r[0]]));
  sheet.getRange(startRow, cols["Jóia"], output.length, 1).setValues(output.map(r => [r[1]]));
  sheet.getRange(startRow, cols["Saldo"], output.length, 1).setValues(output.map(r => [r[2]]));
}


/**
 * Helper interno para extrair o preço correto de um determinado ano
 */
function getPrecosAno_(ano, matrizQuotas) {
  // matrizQuotas é o range Quotas!A2:D
  let t0 = 0, t1 = 0, t2 = 0;
  for (let i = 0; i < matrizQuotas.length; i++) {
    let row = matrizQuotas[i];
    if (!row[0]) continue; // linha vazia
    let rowAno = new Date(row[0]).getFullYear();
    if (rowAno <= ano) {
      t0 = Number(row[1]) || 0;
      t1 = Number(row[2]) || 0;
      t2 = Number(row[3]) || 0;
    }
  }
  return { t0, t1, t2 };
}

/**
 * Calcula a Quota do Ano.
 * Uso no Sheets: =CALCULAR_QUOTA(F7; G7; H7; YEAR(TODAY()); Quotas!$A$2:$D$3)
 *
 * @customfunction
 */
function CALCULAR_QUOTA(numT0, numT1, numT2, anoAtual, matrizQuotas) {
  numT0 = Number(numT0) || 0; numT1 = Number(numT1) || 0; numT2 = Number(numT2) || 0;
  if (!anoAtual) anoAtual = new Date().getFullYear();
  
  const p = getPrecosAno_(anoAtual, matrizQuotas);
  return (numT0 * p.t0) + (numT1 * p.t1) + (numT2 * p.t2);
}

/**
 * Calcula a Jóia de entrada.
 * Uso no Sheets: =CALCULAR_JOIA(F7; G7; H7; E7; Quotas!$A$2:$D$3)
 * 
 * 
 * =LET(
    join_raw; E7;
    adj; IF(join_raw<DATE(2025;3;26); DATE(2025;3;26); join_raw);
    y; YEAR(adj);
    m; MONTH(adj);
    qnum; INT((m-1)/3)+1;
    qstart; DATE(y; (qnum-1)*3+1; 1);
    full_quarters; 5 - qnum - N(adj>qstart);
    base; O7 * full_quarters / 4;
    IF(C7<1; 0; FLOOR(base; 0,05))
  )
 * 
 *
 * @customfunction
 */
function CALCULAR_JOIA(numT0, numT1, numT2, dataAdesaoRaw, matrizQuotas) {
  numT0 = Number(numT0) || 0; numT1 = Number(numT1) || 0; numT2 = Number(numT2) || 0;
  if (numT0 + numT1 + numT2 === 0) return 0;

  let dataAdesao = new Date(dataAdesaoRaw);
  if (isNaN(dataAdesao.getTime())) return 0;

  // Tudo o que é anterior a 26/03/2025 conta como tendo aderido nessa data
  const minDate = new Date(2025, 2, 26); 
  if (dataAdesao < minDate) dataAdesao = minDate;

  const ano = dataAdesao.getFullYear();
  const mes = dataAdesao.getMonth() + 1;
  const qnum = Math.floor((mes - 1) / 3) + 1;
  const qstart = new Date(ano, (qnum - 1) * 3, 1);
  const afterQstart = dataAdesao.getTime() > qstart.getTime() ? 1 : 0;
  const fullQuarters = 5 - qnum - afterQstart;

  const p = getPrecosAno_(ano, matrizQuotas);
  const cotaAnual = (numT0 * p.t0) + (numT1 * p.t1) + (numT2 * p.t2);
  
  const base = (cotaAnual * fullQuarters) / 4;
  return Math.floor(base / 0.05) * 0.05; // Arredondamento da Jóia aos 5 cêntimos
}

/**
 * Calcula a Quotização Total (Jóia + Quotas anuais completas até ao ano atual).
 * Uso no Sheets: =CALCULAR_QUOTIZACAO(F7; G7; H7; E7; YEAR(TODAY()); Quotas!$A$2:$D$20)
 *
 * @customfunction
 */
function CALCULAR_QUOTIZACAO(numT0, numT1, numT2, dataAdesaoRaw, anoLimite, matrizQuotas) {
  numT0 = Number(numT0) || 0; numT1 = Number(numT1) || 0; numT2 = Number(numT2) || 0;
  let total = CALCULAR_JOIA(numT0, numT1, numT2, dataAdesaoRaw, matrizQuotas);
  
  let dataAdesao = new Date(dataAdesaoRaw);
  if (isNaN(dataAdesao.getTime())) return total;

  const minDate = new Date(2025, 2, 26); 
  if (dataAdesao < minDate) dataAdesao = minDate;

  // O primeiro ano completo de quotas é o ano seguinte ao da adesão (que pagou jóia)
  let primeiroAnoCheio = dataAdesao.getFullYear() + 1;
  if (!anoLimite) anoLimite = new Date().getFullYear();

  for (let ano = primeiroAnoCheio; ano <= anoLimite; ano++) {
    let p = getPrecosAno_(ano, matrizQuotas);
    total += (numT0 * p.t0) + (numT1 * p.t1) + (numT2 * p.t2);
  }
  
  return total;
}

/**
 * Calcula o Saldo (Pagamentos - Quotização Total).
 * Uso no Sheets: =CALCULAR_SALDO(C7; F7; G7; H7; E7; YEAR(TODAY()); Quotas!$A$2:$D$3)
 *
=LET(
  join_raw; E7;
  join; IF(join_raw<DATE(2025;3;26); DATE(2025;3;26); join_raw);
  thisYear; YEAR(NOW());
  firstDue; YEAR(join)+1;
  yearsDue; MAX(0; thisYear-firstDue+1);
  if(C7<1;0;C7 - P7 - yearsDue*O7)
)
 * @customfunction
 */
function CALCULAR_SALDO(valorPago, numT0, numT1, numT2, dataAdesaoRaw, anoAtual, matrizQuotas) {
  valorPago = Number(valorPago) || 0;
  const devidas = CALCULAR_QUOTIZACAO(numT0, numT1, numT2, dataAdesaoRaw, anoAtual, matrizQuotas);
  return valorPago - devidas;
}
// ====================================================================



function dumpScriptProps(){
  const sp = PropertiesService.getScriptProperties();
  const cfg = authCfg_();
  return {
    keys: Object.keys(sp.getProperties()).sort(),
    CLIENT_ID_present: !!cfg.clientId,
    CLIENT_SECRET_present: !!cfg.clientSecret,
    sampleAuthUrl: AuthCoreLib.buildAuthUrlFor("diag-"+Date.now(), false, false, cfg),
  };
}

function pingDrive() {
  const it = DriveApp.getRootFolder().getFiles();
  Logger.log('has files? ' + it.hasNext());
}

function debugDriveAuth() { DriveApp.getRootFolder().getName(); }



function debugFetchAuth(){
  //UrlFetchApp.fetch('https://httpbin.org/get');
  UrlFetchApp.fetch('https://www.google.com/generate_204', {muteHttpExceptions:true});
  // ou
  // UrlFetchApp.fetch('https://www.googleapis.com/discovery/v1/apis?fields=discoveryVersion', {muteHttpExceptions:true});

}

function autorizarEmail() {
  MailApp.sendEmail(Session.getActiveUser().getEmail(), "Autorização", "O script já pode enviar e-mails!");
}

