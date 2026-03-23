
// =========================
// Titulares.gs - uso duplo: spreadsheet e web app de edição de titulares
// =========================


/******
 * 
 * Código Utilizado para fórmulas a usar no spreadsheet
 * 
 **************/

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


/*
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
*/

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





/******************************
 * 
 * web app para edição de informação dos associados
 * 
 ******************************/

// ====================================================================
// NOVAS FUNÇÕES PARA A WEB APP DE GESTÃO (ÓRGÃOS SOCIAIS)
// ====================================================================

const VERSION = "v1.0";


function buildAuthUrlFor(nonce, dbg, embed, clientUrl) { return AuthCoreLib.buildAuthUrlFor(nonce, dbg, embed, authCfg_(), clientUrl); }
function pollTicket(nonce)                  { return AuthCoreLib.pollTicket(nonce); }
function isTicketValid(ticket, dbg)              { return AuthCoreLib.isTicketValid(ticket, dbg); }

// modo de visualização. Trocar por outros, para permitir mais do que dois.
const MODO_PROC = true
const MODO_IPS = false

//Onde encontrar os documentos PDF
const FOLDER_PROCURACOES = "18LFaOcQ53pJ6FVQ6mhq6RO_PxQIK1fsx"; //https://drive.google.com/drive/folders/18LFaOcQ53pJ6FVQ6mhq6RO_PxQIK1fsx?usp=sharing
const FOLDER_IPS = "1TXL942FE_Z05gSCJ_f_1lj_nEPnD5DWv"; //https://drive.google.com/drive/folders/1TXL942FE_Z05gSCJ_f_1lj_nEPnD5DWv?usp=sharing

/////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////
//Identificação das tabelas Google Sheets

const SS_TITULARES_ID = "1YE16kNuiOjb1lf4pbQBIgDCPWlEkmlf5_-DDEZ1US3g";
const SS_IPS_ID = "1rsxlHYHrXfSdpgjfhA193xYkhi3r-RMK9cc04l7Sqq8"; // ID da folha IPS

// === CONFIG: ranges nomeados ou A1 (fallback) ===
const RANGES = {
  titulares: { name: "tblTitulares", sheet: "Titulares", a1: "A6:Z" },
  ips:       { name: "tblIPS",      sheet: "IPS",        a1: "A4:N" },
  transacoes:    { name: "tblTransacoes",   sheet: "Transações",     a1: "A2:L" },
};

// Cabeçalhos esperados
const COLS_TITULARES = {
  CT_MEMBROS: "Nome membros (e titulares representados)",
  CT_NUM: "Num.",
  CT_NIF: "NIF",
  CT_NOMEFISCAL: "Nome fiscal",
  CT_PAGO: "€",
  CT_SEMANAS: "Semanas",
  CT_ADESAO: "Adesão",
  CT_TEL: "Telemóvel",
  CT_EMAIL: "e-mail",
  CT_FIM: "Fim",
  CT_ESTADO: "Estado",
  CT_RGPD: "RGPD",
  CT_QUOTA: "Quota",
  CT_JOIA: "Jóia",
  CT_SALDO: "Saldo",
};

const COLS_IPS = {
  CI_NUM: "Num.",   //Esta é a chave, e a identificação do registo de titulares é feita por aqui.
  CI_SEMANAS: "Semanas",    // Deixa de ser necessário, porque se usa CI_NUM para saber quem é o titular
  CI_DATA: "Data IPS",
  CI_STATUS: "IPS",  //Estado da IPS
  CI_PRIMEIRO: "Primeiro titular",
  CI_OUTROS: "Outros titulares",
};

//  const { header: th, rows: titulares } = fetchTable_(SS_TITULARES_ID, RANGES.titulares);
//  const col = indexByHeader_(th); const H = COLS_TITULARES;
//  const linhas = titulares.filter(r => cellHasEmail_(r[col[H.CT_EMAIL]], emailLC));
//  linhas.forEach(r=>{
//    const nomeFiscal   = r[col[H.CT_NOMEFISCAL]] || "";

/////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////


// ===== Titulares helpers =====
function fetchTable_(ssId, cfg){
  const ss = SpreadsheetApp.openById(ssId);
  let range = null;
  if (cfg.name) { try { range = ss.getRangeByName(cfg.name); } catch(_){ } }
  if (!range)   { range = ss.getSheetByName(cfg.sheet).getRange(cfg.a1); }
  const values  = range.getDisplayValues();
  if (!values || !values.length) return { header: [], rows: [] };
  const header = values[0];
  const rows   = values.slice(1).filter(r => !isRowEmpty_(r));
  return { header, rows };
}
function indexByHeader_(header){ const m = {}; header.forEach((h,i)=> m[String(h).trim()] = i); return m; }
function isRowEmpty_(row){ return row.every(v => String(v).trim()===""); }

function doGet(e) {
  const DBG = !!(e && e.parameter && e.parameter.debug === '1');
  const ticket = (e && e.parameter && e.parameter.ticket) || "";
  //console.log("doGet(e): DBG=", DBG, ", ticket=", ticket);

  var opts = { ticket: ticket, debug: DBG, serverLog: ["Gestão Titulares"], wipe: false };

  if (e && e.parameter && e.parameter.action === 'logout') {
    opts.wipe = true; return AuthCoreLib.renderLoginPage(opts); 
  }

  if (e && e.parameter && e.parameter.code) return AuthCoreLib.finishAuth(e, authCfg_());

  if (ticket) {
    try {
      const sess = AuthCoreLib.requireSession(ticket);

      const editorCfg = getEditorConfig_(sess.email);
      if (!editorCfg) return HtmlService.createHtmlOutput("<h3>Acesso não autorizado para este editor.</h3>");
      return renderMainPage_(ticket, DBG);
    } catch(err) {
      opts.wipe = true;
      return AuthCoreLib.renderLoginPage(opts);
    }
  }
  return AuthCoreLib.renderLoginPage(opts);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function authCfg_() {
  const sp = PropertiesService.getScriptProperties();
  return { clientId: sp.getProperty("CLIENT_ID"), clientSecret: sp.getProperty("CLIENT_SECRET") };
}

function getEditorConfig_(email) {
  console.log("getEditorConfig_(" + email + ")"); 
  const sp = PropertiesService.getScriptProperties();
  //Não é preciso WHITELIST_CSV
  //const whitelist = (sp.getProperty("WHITELIST_CSV") || "").toLowerCase().split(",");
  //console.log("whitelist=", whitelist); 
  //if (!whitelist.includes(email.toLowerCase())) return null;

  // Lógica de mapeamento: email=A0001-A0049
  const mappingStr = sp.getProperty("EDITOR_MAPPING") || "";
  //console.log("mappingStr=", mappingStr); 
  const lines = mappingStr.split(";");
  for (let line of lines) {
    //console.log("line=", line); 
    const [mEmail, range] = line.split("=");
    if (mEmail && mEmail.trim().toLowerCase() === email.toLowerCase()) {
      let [start, end] = range.split("-");
      //console.log("start.length=" + start.length);
      //console.log("after start.length");
      if (start.length === 0) {
        //console.log("start.length é zero");
        start = "A0001";
        console.log("start forced to " + start);
      }
      //console.log("end.length=" + end.length);
      //console.log("after end.length");
      if (end.length === 0) {
        //console.log("end.length é zero");
        end = "Z9999";
        console.log("end forced to " + end);
      }      
      //console.log("fim start/end");
      console.log("start=" + start); // + ", " + start.trim());
      console.log("end=" + end); //+ ", " + end.trim());
      return { start: start.trim(), end: end.trim() };
    }
  }
  return null; //Não foi encontrado o e-mail do login
}

function renderMainPage_(ticket, DBG) {
  const t = HtmlService.createTemplateFromFile('Main');
  t.ticket = ticket;
  t.DEBUG = DBG ? '1' : '';
  return t.evaluate().setTitle("Gestão de Titulares").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// API chamada pelo Frontend
function getEditorInitialData(ticket) {
  const sess = AuthCoreLib.requireSession(ticket);
  const cfg = getEditorConfig_(sess.email);
  //console.log("getEditorInitialData: start=",cfg.start,", end=",cfg.end);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Titulares");
  const data = sheet.getRange("B7:B").getValues(); // Coluna B = Num.
  //console.log("data=", data);
  const filtered = data.flat().filter(num => num >= cfg.start && num <= cfg.end);

  //Fica para mais tarde usar nomes automáticos...
  //const { header: th, rows: titulares } = fetchTable_(SS_TITULARES_ID, RANGES.titulares);

  console.log("filtered=" + filtered);
  return { range: cfg, list: filtered };
}

function loadRecordData(ticket, numA) {
  console.log("loadRecordData(" + numA + ")");
  AuthCoreLib.requireSession(ticket);

  const { header: th, rows: titulares } = fetchTable_(SS_TITULARES_ID, RANGES.titulares);
  const col = indexByHeader_(th); const H = COLS_TITULARES;
  //  const linhas = titulares.filter(r => cellHasEmail_(r[col[H.CT_EMAIL]], emailLC));
  //  linhas.forEach(r=>{
  //    const nomeFiscal   = r[col[H.CT_NOMEFISCAL]] || "";
  //console.log("loadRecordData: col[H.CT_MEMBROS]=" + col[H.CT_MEMBROS]);
  //console.log("loadRecordData: col[H.CT_NUM]=" + col[H.CT_NUM]);
  //console.log("loadRecordData: col[H.CT_SALDO]=" + col[H.CT_SALDO]);
  const rowT = titulares.find(r => String(r[col[H.CT_NUM]]) === numA);
  console.log("rowT=" + rowT);
  console.log("rowT[col[H.CT_SALDO]]=" + rowT[col[H.CT_SALDO]]);
  if (!rowT) throw new Error("Registo não encontrado em Titulares.");

  const shI = SpreadsheetApp.openById(SS_IPS_ID).getSheetByName("IPS");
  const dataI = shI.getRange("A4:N").getValues();
  const rowI = dataI.find(r => String(r[0]) === numA); // Col CI_NUM
  console.log("rowI=" + rowI);

  // Formatar a data para o input HTML (yyyy-mm-dd)
  let ipsDataStr = "";
  if (rowI && rowI[9]) {
    const dataVal = rowI[9];
    ipsDataStr = (dataVal instanceof Date) 
      ? Utilities.formatDate(dataVal, "GMT", "yyyy-MM-dd")
      : dataVal;
    console.log("dataVal=" + dataVal + ", ipsDataStr=" + ipsDataStr);
  }
  
  return {
    nomeMembros: rowT[col[H.CT_MEMBROS]],
    nif: rowT[col[H.CT_NIF]],
    nomeFiscal: rowT[col[H.CT_NOMEFISCAL]],
    semanas: rowT[col[H.CT_SEMANAS]], // Coluna I
    telemoveis: rowT[col[H.CT_TEL]] || "---",
    emails: rowT[col[H.CT_EMAIL]] || "---",
    adesao: rowT[col[H.CT_ADESAO]] || "---",
    fim: rowT[col[H.CT_FIM]] || "---",
    estado: rowT[col[H.CT_ESTADO]] || "---",
    pago: rowT[col[H.CT_PAGO]] || "---",
    saldo: rowT[col[H.CT_SALDO]] || "---",
    ipsData: ipsDataStr, // Nova coluna J (índice 9)
    ipsEstado: rowI ? rowI[10] : "N/D",
    ipsComentario: rowI ? rowI[11] : "",
    primTitular: rowI ? rowI[12] : "",
    outrosTitulares: rowI ? rowI[13] : ""
  };
}


/**
 * Função auxiliar para pesquisar ficheiros em toda a árvore de subpastas.
 */
function getAllFilesRecursive_(folder, query) {
  let results = [];
  
  // Procura na pasta atual
  const files = folder.searchFiles(query);
  while (files.hasNext()) {
    results.push(files.next());
  }
  
  // Procura recursivamente em todas as subpastas
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    results = results.concat(getAllFilesRecursive_(sub, query));
  }
  
  return results;
}

/**
 * Localiza o PDF com base no modo (Procuração ou IPS) e critérios de prioridade.
 */
function findPdfPreview(ticket, semana, modo) {
  console.log("findPdfPreview(" + semana + ", " + modo + ")");
  AuthCoreLib.requireSession(ticket);
  
  const prefix = semana.replace("/", "-");
  console.log("findPdfPreview: prefix=" + prefix);

  // Mantemos a query larga para o Drive encontrar os candidatos
  let query = `title contains '${prefix}' and mimeType = 'application/pdf' and trashed = false`;
  let folderId = modo === MODO_PROC ? FOLDER_PROCURACOES : FOLDER_IPS;
  const rootFolder = DriveApp.getFolderById(folderId);

  const candidates = getAllFilesRecursive_(rootFolder, query);
  console.log("candidates.length=" + candidates.length);
  console.log("candidates=" + candidates);

  if (candidates.length > 0) {
    let strictPattern;
    if (modo === MODO_PROC) {
      console.log("strict for MODO_PROC");
      ///Localizar ficheiros das procurações dos titulares: 
      // "915-26-2025 Procuração_signed.pdf" ou
      // "114-33 2025 Procuração_Marta.pdf" ou
      // "210-31-2025.pdf"
      // RegEx: Começa com o prefixo, seguido de '-' ou ' ', seguido do ano, como '2025'
      // Exemplo: ^116-36[- ]202.
      strictPattern = new RegExp("^" + prefix + "[- ]20..", "i");
    } else {
      console.log("strict for MODO_IPS");
      //Localizar ficheiros das IPS:
      //"101-29 FT-IPS 2025.pdf"
      //"101-31 FT-IPS 2025-01-18.pdf"
      //"104-29 FT-IPS 2024x.pdf"
      //"319-27 CP-FT 2024x.pdf", "319-28 CP-FT 2024x.pdf", "319-27 CP-FT 2024x.pdf"  - e não há nenhuma FT-IPS nem FT-IPS
      //"603-30 IPS 2020x.pdf"
      // RegEx: Começa com o prefixo, e mais nada (nem era preciso o strictPattern neste caso)
      // Exemplo: ^116-36
      strictPattern = new RegExp("^" + prefix, "i");
    }

    // Filtramos os candidatos pelo padrão exato
    let matches = candidates.filter(f => strictPattern.test(f.getName()));
    console.log("matches.length=" + matches.length);
    console.log("matches=" + matches);

    if (matches.length > 0) {
      // Lógica de Seleção do "Best Match"
      matches.sort((a, b) => {
        console.log("sort: a=" + a.getName() + ", b=" + b.getName());
        if (modo === MODO_PROC) {
          // Prioridade ao mais recente (Data de modificação)
          return b.getLastUpdated() - a.getLastUpdated();
        } else {
          // Prioridade MODO_IPS:
          const nameA = a.getName();
          const nameB = b.getName();
          
          // 1. Extrair ano (procura 4 dígitos 202x)
          const yearA = (nameA.match(/202\d/) || [0])[0];
          const yearB = (nameB.match(/202\d/) || [0])[0];
          
          if (yearA !== yearB) return yearB - yearA; // Ano mais recente primeiro
          
          // 2. Se anos iguais, ordem alfabética (CP-FT < FT-IPS < IPS)
          // O localeCompare resolve isto naturalmente para os seus termos
          return nameA.localeCompare(nameB);
        }
      });

      const bestMatch = matches[0];
      console.log("Selecionado (" + (modo ? "PROC" : "IPS") + "): " + bestMatch.getName());
      const fileId = bestMatch.getId();
      
      return { 
        name: bestMatch.getName(), 
        id: bestMatch.getId(),
        
        //O navegador não permite copiar texto, mas é o único funcional.
        url: `https://drive.google.com/file/d/${fileId}/preview` 
        
        // Usando o motor do Docs Viewer que por vezes permite seleção de texto => erro "No preview available" porque não usa a autenticação Google
        //url: `https://docs.google.com/viewer?srcid=${fileId}&pid=explorer&efh=false&a=v&chrome=false&embedded=true`
        //url: `https://docs.google.com/viewer?srcid=${bestMatch.getId()}&pid=explorer&efh=false&a=v&chrome=false&embedded=true`
      };
    }
  }
  
  console.warn("Aviso: Nenhum ficheiro PDF encontrado para: " + prefix + " (mesmo em subpastas)");
  return null;
}

function saveRecordData(ticket, numA, payload) {
  console.log("saveRecordData(" + numA + "," + JSON.stringify(payload) + ")");
  AuthCoreLib.requireSession(ticket);
  
  // Update Titulares
  const shI = SpreadsheetApp.openById(SS_IPS_ID).getSheetByName("IPS");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shT = ss.getSheetByName("Titulares");
  // Procuramos o numA na coluna B (B7:B)
  const dataT = shT.getRange("B7:B").getValues().flat();
  //console.log("dataT=" + dataT);
  const idxT = dataT.indexOf(numA);
  if (idxT !== -1) {
  // Para atualizar A, C e D sem corromper o ID na coluna B,
    // escrevemos o bloco de 4 colunas (A a D), reinserindo o numA na coluna B.
    const rowNumberT = idxT + 7;

    // O ideal seria ler a linha toda, modificar as colunas desejadas, usando nomes automáticos, e depois escrever a linha toda
    shT.getRange(rowNumberT, 1, 1, 4).setValues([[
      payload.nomeMembros, // Coluna A (index 0 no load)
      numA,                // Coluna B (ID - mantemos o valor original)
      payload.nif,         // Coluna C (index 2 no load)
      payload.nomeFiscal   // Coluna D (index 3 no load)
    ]]);

    console.log("payload.estadp=" + payload.estado);    
    shT.getRange(rowNumberT, 17, 1, 1).setValues([[
      payload.estado //CT_ESTADO
    ]]);
  }

  // Update IPS (Colunas K, L, M, N)
  const dataI = shI.getRange("A4:A").getValues().flat();
  const idxI = dataI.indexOf(numA);
  if (idxI !== -1) {
    // No loadRecordData, o ipsComentario é rowI[11].
    // O index 11 num array que começa em A corresponde à 12ª coluna (Coluna L).
    const rowNumberI = idxI + 4;
    shI.getRange(rowNumberI, 10, 1, 5).setValues([[
      payload.ipsData,
      payload.ipsEstado,
      payload.ipsComentario,
      payload.primTitular,
      payload.outrosTitulares
    ]]);
  }
  
  return { ok: true };
}



