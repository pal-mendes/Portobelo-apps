
// =========================
// Associados.gs
// =========================

/*************************
 * 
 * Utilizado para fórmulas a usar no spreadsheet, mas também para a web app
 * 
 ************************/



// ====================================================================
// FUNÇÕES PERSONALIZADAS PARA O GOOGLE SHEETS
// ====================================================================


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
    // Colunas de influência: F(6)=€ pagos, H(8)=Adesão, I(9)=T0, J(10)=T1, K(11)=T2
    if ([6, 8, 9, 10, 11].includes(col)) {
      updateTitularesRowValues_(sheet, row);
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
 * Atualiza os valores estáticos de uma única linha na aba Titulares.
 */
function updateTitularesRowValues_(sheet, row) {
  const ss = sheet.getParent();
  const quotasSheet = ss.getSheetByName("Quotas");
  const matrizQuotas = quotasSheet.getRange("A2:D3").getValues();
  
  // Lê os dados da linha (C até H)
  const data = sheet.getRange(row, 1, 1, 8).getValues()[0];
  const valorPago = data[6];     // F
  const dataAdesao = data[8];    // H
  const numT0 = data[9];         // I
  const numT1 = data[10];         // J
  const numT2 = data[11];         // K
  const anoAtual = new Date().getFullYear();

  // Reutiliza as funções de cálculo existentes
  const quota = CALCULAR_QUOTA(numT0, numT1, numT2, anoAtual, matrizQuotas);
  const joia = CALCULAR_JOIA(numT0, numT1, numT2, dataAdesao, matrizQuotas);
  const saldo = CALCULAR_SALDO(valorPago, numT0, numT1, numT2, dataAdesao, anoAtual, matrizQuotas);

  // Escreve valores nas colunas O(15)=Quota, P(16)=Jóia, Q(17)=Saldo
  sheet.getRange(row, 18, 1, 3).setValues([[quota, joia, saldo]]);
}

/**
 * Atualiza todas as linhas da tabela tblTitulares de uma só vez (Batch Update).
 */
function updateAllTitularesValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Titulares");
  const quotasSheet = ss.getSheetByName("Quotas");
  const matrizQuotas = quotasSheet.getRange("A2:D3").getValues();
  
  const startRow = 7;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const numRows = lastRow - startRow + 1;
  const rangeIn = sheet.getRange(startRow, 1, numRows, 8); // A até H
  const valuesIn = rangeIn.getValues();
  const anoAtual = new Date().getFullYear();
  
  const results = [];

  for (let i = 0; i < valuesIn.length; i++) {
    const r = valuesIn[i];
    const pago = r[5]; const adesao = r[7]; const t0 = r[8]; const t1 = r[9]; const t2 = r[10];
    
    const quota = CALCULAR_QUOTA(t0, t1, t2, anoAtual, matrizQuotas);
    const joia = CALCULAR_JOIA(t0, t1, t2, adesao, matrizQuotas);
    const saldo = CALCULAR_SALDO(pago, t0, t1, t2, adesao, anoAtual, matrizQuotas);
    
    results.push([quota, joia, saldo]);
  }

  // Escrita em massa para performance máxima
  sheet.getRange(startRow, 15, numRows, 3).setValues(results);
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
function CALCULAR_QUOTIZACAO(numT0, numT1, numT2, dataAdesaoRaw, anoAtual, matrizQuotas) {
  numT0 = Number(numT0) || 0; numT1 = Number(numT1) || 0; numT2 = Number(numT2) || 0;
  let total = CALCULAR_JOIA(numT0, numT1, numT2, dataAdesaoRaw, matrizQuotas);
  
  let dataAdesao = new Date(dataAdesaoRaw);
  if (isNaN(dataAdesao.getTime())) return total;

  const minDate = new Date(2025, 2, 26); 
  if (dataAdesao < minDate) dataAdesao = minDate;

  // O primeiro ano completo de quotas é o ano seguinte ao da adesão (que pagou jóia)
  let primeiroAnoCheio = dataAdesao.getFullYear() + 1;
  if (!anoAtual) anoAtual = new Date().getFullYear();

  for (let ano = primeiroAnoCheio; ano <= anoAtual; ano++) {
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

