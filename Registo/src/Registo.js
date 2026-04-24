
// =========================
// Registo.gs - web app de Registo de Associados
// =========================

const VERSION = "v1.8";


function buildAuthUrlFor(nonce, dbg, embed, clientUrl) { return AuthCoreLib.buildAuthUrlFor(nonce, dbg, embed, authCfg_(), clientUrl); }
function pollTicket(nonce)                  { return AuthCoreLib.pollTicket(nonce); }
function isTicketValid(ticket, dbg)              { return AuthCoreLib.isTicketValid(ticket, dbg); }

// modo de visualização. Conforme id="selModo" em Main.html
const MODO_PROC = "PROC";
const MODO_IPS = "IPS";


//Onde encontrar os documentos PDF
const FOLDER_PROCURACOES = "18LFaOcQ53pJ6FVQ6mhq6RO_PxQIK1fsx"; //https://drive.google.com/drive/folders/18LFaOcQ53pJ6FVQ6mhq6RO_PxQIK1fsx?usp=sharing
const FOLDER_IPS = "1TXL942FE_Z05gSCJ_f_1lj_nEPnD5DWv"; //https://drive.google.com/drive/folders/1TXL942FE_Z05gSCJ_f_1lj_nEPnD5DWv?usp=sharing

/////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////
//Identificação das tabelas Google Sheets

var SS_TITULARES_ID = "1YE16kNuiOjb1lf4pbQBIgDCPWlEkmlf5_-DDEZ1US3g";
const SS_IPS_ID = "1rsxlHYHrXfSdpgjfhA193xYkhi3r-RMK9cc04l7Sqq8"; // ID da folha IPS
const SS_PROC_ID = "1j8c97XMoHvmmUhvVP3brsIq42DQzmkml4B1N7KeSHNg"; // ID da folha Procurações


// === CONFIG: ranges nomeados ou A1 (fallback) ===
const RANGES = {
  titulares: { name: "tblTitulares", sheet: "Titulares", a1: "A6:Z" },
  ips:       { name: "tblIPS",      sheet: "IPS",        a1: "A4:N" },
  procuracoes:    { name: "tblProc",   sheet: "Procurações",     a1: "A4:Q" },
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
  CT_MEMBROS: "Membros",
  CI_NIF: "NIF",
  CI_NOMEFISCAL: "Nome fiscal",
  CI_TEL: "Telemóvel",
  CI_EMAIL: "e-mail",
  CI_DATA: "Data IPS",
  CI_STATUS: "IPS",  //Estado da IPS
  CI_COMENT: "Comentário IPS", 
  CI_PRIMEIRO: "Primeiro titular",
  CI_OUTROS: "Outros titulares",
};

const COLS_PROCURACOES = {
  CP_NUM: "Num.",   //Esta é a chave, e a identificação do registo de titulares é feita por aqui.
  CP_PART: "Participação", 
  CP_DATA: "Data Proc.",
  CP_FORM: "Formato",  
  CP_CERT: "Certificação",
  CP_CC: "Cartão Cidadão",
  CP_DOCS: "Docs.",
  CP_ESTADO: "Estado",
  CP_COMENT: "Comentário",
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

function getTitularesColumns_(sheet) {
  const headerRow = 6;
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};

  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  
  // Mapeamento dos nomes exatos das colunas para o seu número (1-based)
  const targets = ["Nome membros (e titulares representados)", "Num.", "NIF", "Nome fiscal", "€", "Semanas", "Adesão", "T0", "T1", "T2", "Telemóvel", "e-mail", "Fim", "Estado", "RGPD", "Quota", "Jóia", "Saldo"];
  headers.forEach((label, index) => {
    if (targets.includes(label)) {
      map[label] = index + 1;
    }
  });
  return map;
}


function isRowEmpty_(row){ return row.every(v => String(v).trim()===""); }

function doGet(e) {
  const DBG = !!(e && e.parameter && e.parameter.debug === '1');
  const ticket = (e && e.parameter && e.parameter.ticket) || "";
  //console.log("doGet(e): DBG=", DBG, ", ticket=", ticket);

  var opts = { 
    ticket: ticket, 
    debug: DBG, 
    serverLog: ["Registo Associados"], 
    wipe: false,
    appTitle: "Registo de associados",
    appPermissions: "Só tem acesso quem for designado pela direção da associação.",
  };

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


/**
 * Gere o sistema de bloqueio (Locking) de registos.
 * Formato: email=A0130(2026-03-26 11:50);outro@email.com=A0007(...)
 */
function getLockStatus(ticket, numA) {
  const sess = AuthCoreLib.requireSession(ticket);
  const sp = PropertiesService.getScriptProperties();
  let lockStr = sp.getProperty("EDITOR_LOCK") || "";
  const now = new Date();
  const SIX_HOURS = 6 * 60 * 60 * 1000;

  let locks = lockStr.split(";").filter(l => l.includes("="));
  let updatedLocks = [];
  let currentLock = null;

  locks.forEach(l => {
    // Regex para extrair: email=registo(timestamp)
    const match = l.match(/^([^=]+)=([^\(]+)\(([^\)]+)\)$/);
    if (match) {
      const [_, user, registo, ts] = match;
      const lockDate = new Date(ts.replace(/-/g, "/")); // Fix para parsing de data
      
      // Manter apenas se tiver menos de 6 horas (Ponto 11)
      if (now - lockDate < SIX_HOURS) {
        updatedLocks.push(l);
        if (registo === numA) {
          currentLock = { user, time: ts };
        }
      }
    }
  });

  // Salvar limpeza automática se necessário
  if (updatedLocks.length !== locks.length) {
    sp.setProperty("EDITOR_LOCK", updatedLocks.join(";"));
  }

  return { 
    isLocked: !!currentLock, 
    lockedBy: currentLock ? currentLock.user : null,
    isMine: currentLock ? currentLock.user === sess.email : false
  };
}

function acquireLock(ticket, numA) {
  const sess = AuthCoreLib.requireSession(ticket);
  const status = getLockStatus(ticket, numA);
  
  if (status.isLocked && !status.isMine) return { ok: false, owner: status.lockedBy };

  const sp = PropertiesService.getScriptProperties();
  const timestamp = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd HH:mm");
  const newEntry = `${sess.email}=${numA}(${timestamp})`;
  
  let lockStr = sp.getProperty("EDITOR_LOCK") || "";
  let locks = lockStr.split(";").filter(l => l && !l.startsWith(sess.email + "=")); // Remove lock anterior do mesmo utilizador
  locks.push(newEntry);
  
  sp.setProperty("EDITOR_LOCK", locks.join(";"));
  return { ok: true };
}

function releaseLock(ticket, numA) {
  const sess = AuthCoreLib.requireSession(ticket);
  const sp = PropertiesService.getScriptProperties();
  let lockStr = sp.getProperty("EDITOR_LOCK") || "";
  if (!lockStr) return { ok: true };
  
  // Filtra para manter todos os locks EXCETO o deste utilizador para este registo específico
  let locks = lockStr.split(";").filter(l => {
    return l !== "" && !l.startsWith(sess.email + "=" + numA + "(");
  });

  sp.setProperty("EDITOR_LOCK", locks.join(";"));
  return { ok: true };
}

function renderMainPage_(ticket, DBG) {
  const t = HtmlService.createTemplateFromFile('Main');
  t.ticket = ticket;
  t.DEBUG = DBG ? '1' : '';
  return t.evaluate().setTitle("Registo de Associados").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Ponto 7: Função de Pesquisa no Servidor, que pode ou não vir restringida a um filtro de estado do associado ou IPS
function searchRecords(ticket, query, filtroEstado) {
  AuthCoreLib.requireSession(ticket);

  const fullQuery = query.trim();
  if (!fullQuery) return [];

  const ss = SpreadsheetApp.openById(SS_TITULARES_ID);
  const shT = ss.getSheetByName("Titulares");
  const headersT = shT.getRange(6, 1, 1, shT.getLastColumn()).getValues()[0];
  const colT = indexByHeader_(headersT);
  const HT = COLS_TITULARES;
  const dataT = shT.getRange(7, 1, shT.getLastRow() - 6, shT.getLastColumn()).getValues();

  // Determinar se o filtro é de Titulares ou IPS
  const estadosTitulares = ["0S-Incompleto", "0Y-NãoTitular", "0Z-Completo"]; //Os outros valores de estado são da tblIPS
  const isFiltroTit = estadosTitulares.includes(filtroEstado);

  let numsValidos = null;

  // Se o filtro for de IPS, precisamos de saber quais os Nums que cumprem o critério na tblIPS, e depois filtrarão a pesquisa em tblTitulares
  if (filtroEstado && !isFiltroTit) {
    const shI = SpreadsheetApp.openById(SS_IPS_ID).getSheetByName("IPS");
    const headersI = shI.getRange(4, 1, 1, shI.getLastColumn()).getValues()[0];
    const colI = indexByHeader_(headersI);
    const HI = COLS_IPS;
    const dataI = shI.getRange(4, 1, shI.getLastRow() - 3, shI.getLastColumn()).getValues();
    
    //Em numsValidos temos os números de registo de associados no estado de IPS definido pelo filtro
    numsValidos = new Set(
      dataI.filter(r => String(r[colI[HI.CI_STATUS]]) === filtroEstado)
           .map(r => String(r[colI[HI.CI_NUM]]))
    );
  }

  // 1. Obter a primeira parte para análise de padrão (LowerCase)
  const queryLower = fullQuery.toLowerCase();
  const firstPart = queryLower.split(' ')[0];

  let filterFn;

  // 2. Identificar o tipo de pesquisa por Regex
  if (firstPart.includes('@')) {
    // Ponto 4: Email
    console.log("pesquisa por email");
    const idx = colT[HT.CT_EMAIL];
    filterFn = r => String(r[idx]).toLowerCase().includes(firstPart);
  } else if (/^\d{9}$/.test(queryLower.replace(/\s+/g, ''))) {
    // Ponto 6: Telefone ou NIF (ignorando espaços)
    const idxTel = colT[HT.CT_TEL];
    const idxNif = colT[HT.CT_NIF];
    
    // String de 9 dígitos limpa de quaisquer espaços
    const numLimpo = queryLower.replace(/\s+/g, '');
    
    // Procura o número limpo removendo também os espaços dos valores guardados na folha
    filterFn = r => {
      const telNaFolha = String(r[idxTel]).replace(/\s+/g, '');
      const nifNaFolha = String(r[idxNif]).replace(/\s+/g, '');
      return telNaFolha.includes(numLimpo) || nifNaFolha.includes(numLimpo);
    };
  } else if (/^\d{3}(\/\d{2})?$/.test(firstPart)) {
    // Ponto 3: Semanas (Ex: 104 ou 104/30)
    console.log("pesquisa por apt/sem");
    const idx = colT[HT.CT_SEMANAS];
    filterFn = r => String(r[idx]).toLowerCase().includes(firstPart);
  } else if (/^[a-z]\d{4}$/.test(firstPart)) {
    // Ponto 5: Num (Ex: a1234)
    console.log("pesquisa por num.");
    const idx = colT[HT.CT_NUM];
    filterFn = r => String(r[idx]).toLowerCase().includes(firstPart);
  } else {
    // Ponto 7: Nome Membros ou Nome Fiscal (String completa)
    console.log("pesquisa por membros ou nome fiscal - todas as palavras");
    const idxMem = colT[HT.CT_MEMBROS];
    const idxFis = colT[HT.CT_NOMEFISCAL];

    // Divide a query em palavras individuais (ex: "Pedro Andrade Mendes" -> ["pedro", "andrade", "mendes"])
    const queryWords = queryLower.split(/\s+/).filter(w => w.length > 0);

    // Retorna true se TODAS as palavras existirem em Membros OU TODAS no Nome Fiscal
    filterFn = r => {
      const valMem = String(r[idxMem]).toLowerCase();
      const valFis = String(r[idxFis]).toLowerCase();
      return queryWords.every(w => valMem.includes(w)) || queryWords.every(w => valFis.includes(w));            
    };
  }
  
  //console.log("searchRecords(" + q + ")");
  //console.log("data.length=" + data.length);
  //console.log("data[[0]]=" + data[[0]]);
  //console.log("data[[1]]=" + data[[1]]);

  // Executar o filtro e mapear resultados
  return dataT.filter(r => {
    const num = String(r[colT[HT.CT_NUM]]); //O número de registo passa a ser necessário para o caso de estado IPS
    if (filtroEstado) { //Lidar com os filtros de estado
      if (isFiltroTit) {
        if (r[colT[HT.CT_ESTADO]] !== filtroEstado) return false; //Filtro de titulares: comparar o estado com o pesquisado
      } else {
        if (!numsValidos.has(num)) return false; //Filtro de IPS: verificar se é um dos números no estado desejado
      }
    }

    // 1. Executar a função de pesquisa para a linha atual
    const matchesSearch = filterFn(r);
    return matchesSearch; //Já deu return false se o estado não for o desejado.
  }).map(r => ({
    num: r[colT[HT.CT_NUM]],
    estado: r[colT[HT.CT_ESTADO]],
    membros: r[colT[HT.CT_MEMBROS]],
    semanas: r[colT[HT.CT_SEMANAS]],
    tel: r[colT[HT.CT_TEL]],
    email: r[colT[HT.CT_EMAIL]]
  })).slice(0, 50); // Limite de 50 resultados para rapidezH
}

// API chamada pelo Frontend
function getEditorInitialData(ticket) {
  const sess = AuthCoreLib.requireSession(ticket);
  const cfg = getEditorConfig_(sess.email);
  //console.log("getEditorInitialData: start=",cfg.start,", end=",cfg.end);

  // 1. Obter dados de Titulares
  const shT = SpreadsheetApp.openById(SS_TITULARES_ID).getSheetByName("Titulares");	
  const dataT = shT.getRange(7, 1, shT.getLastRow() - 6, shT.getLastColumn()).getValues();
  const th = shT.getRange(6, 1, 1, shT.getLastColumn()).getValues()[0];
  const colT = indexByHeader_(th);
  const HT = COLS_TITULARES;

  // 2. Obter dados de IPS
  const shI = SpreadsheetApp.openById(SS_IPS_ID).getSheetByName("IPS");
  const dataI = shI.getRange(5, 1, shI.getLastRow() - 4, shI.getLastColumn()).getValues();
  const ih = shI.getRange(4, 1, 1, shI.getLastColumn()).getValues()[0];
  const colI = indexByHeader_(ih);
  const HI = COLS_IPS;

  // 2. Obter dados de Proc
  const shP = SpreadsheetApp.openById(SS_PROC_ID).getSheetByName("Procurações");
  const dataP = shP.getRange(5, 1, shP.getLastRow() - 4, shP.getLastColumn()).getValues();
  const ph = shP.getRange(4, 1, 1, shP.getLastColumn()).getValues()[0];
  const colP = indexByHeader_(ph);
  const HP = COLS_PROCURACOES;

  // Criar um mapa de estados IPS indexado pelo Num
  const ipsMap = {};
  dataI.forEach(r => {
    const num = String(r[colI[HI.CI_NUM]]);
    if (num) ipsMap[num] = r[colI[HI.CI_STATUS]] || "?";
  });

  // Criar um mapa de estados Proc indexado pelo Num
  const procMap = {};
  dataP.forEach(r => {
    const num = String(r[colP[HP.CP_NUM]]);
    if (num) procMap[num] = r[colP[HP.CP_ESTADO]] || "?";
  });

  // Consolidar lista para o editor e ordenar
  const filtered = dataT
    .filter(r => {
      const num = String(r[colT[HT.CT_NUM]]);
      return num >= cfg.start && num <= cfg.end;
    })
    .map(r => {
      const num = String(r[colT[HT.CT_NUM]]);
      return {
        num: num,
        estTit: String(r[colT[HT.CT_ESTADO]]),
        estIps: ipsMap[num] || "?",
        estProc: procMap[num] || "?"
      };
    })
    // Adicionamos a ordenação alfabética pelo campo 'num'
    .sort((a, b) => a.num.localeCompare(b.num));
  
  console.log("filtered=" + filtered);
  return { range: cfg, list: filtered };
}

function loadRecordData(ticket, numA) {
  console.log("loadRecordData(" + numA + ")");
  AuthCoreLib.requireSession(ticket);

  const { header: th, rows: titulares } = fetchTable_(SS_TITULARES_ID, RANGES.titulares);
  const colT = indexByHeader_(th); const HT = COLS_TITULARES;
  const rowT = titulares.find(r => String(r[colT[HT.CT_NUM]]) === numA);
  console.log("rowT=" + rowT);
  console.log("rowT[colT[HT.CT_SALDO]]=" + rowT[colT[HT.CT_SALDO]]);
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
  
  const { header: ph, rows: procuracoes } = fetchTable_(SS_PROC_ID, RANGES.procuracoes);
  const colP = indexByHeader_(ph); const HP = COLS_PROCURACOES;
  const rowP = procuracoes.find(r => String(r[colP[HP.CP_NUM]]) === numA);
  console.log("rowP=" + rowP);

  if (rowP)  console.log("rowP[colP[HP.CP_ESTADO]]=" + rowP[colP[HP.CP_ESTADO]]);
  //if (!rowP) throw new Error("Registo não encontrado em Procurações.");

  // Formatar a data para o input HTML (yyyy-mm-dd)
  let procDataStr = "";
  if (rowP && rowP[colP[HP.CP_DATA]]) {
    const dataVal = rowP[colP[HP.CP_DATA]];
    procDataStr = (dataVal instanceof Date) 
      ? Utilities.formatDate(dataVal, "GMT", "yyyy-MM-dd")
      : dataVal;
    console.log("dataVal=" + dataVal + ", procDataStr=" + procDataStr);
  }

  return {
    nomeMembros: rowT[colT[HT.CT_MEMBROS]],
    nif: rowT[colT[HT.CT_NIF]],
    nomeFiscal: rowT[colT[HT.CT_NOMEFISCAL]],
    semanas: rowT[colT[HT.CT_SEMANAS]], // Coluna I
    telemoveis: rowT[colT[HT.CT_TEL]] || "---",
    emails: rowT[colT[HT.CT_EMAIL]] || "---",
    adesao: rowT[colT[HT.CT_ADESAO]] || "---",
    fim: rowT[colT[HT.CT_FIM]] || "---",
    estado: rowT[colT[HT.CT_ESTADO]] || "---",
    rgpd: rowT[colT[HT.CT_RGPD]] || "---",
    pago: rowT[colT[HT.CT_PAGO]] || "---",
    saldo: rowT[colT[HT.CT_SALDO]] || "---",
    ipsData: ipsDataStr, // Nova coluna J (índice 9)
    ipsEstado: rowI ? rowI[10] : "N/D",
    ipsComentario: rowI ? rowI[11] : "",
    primTitular: rowI ? rowI[12] : "",
    outrosTitulares: rowI ? rowI[13] : "",
    procPart: rowP ? rowP[colP[HP.CP_PART]] : "",
    procData: rowP ? rowP[colP[HP.CP_DATA]] : "",
    procForm: rowP ? rowP[colP[HP.CP_FORM]] : "",
    procCert: rowP ? rowP[colP[HP.CP_CERT]] : "",
    procCc: rowP ? rowP[colP[HP.CP_CC]] : "",
    procDocs: rowP ? rowP[colP[HP.CP_DOCS]] : "",
    procEstado: rowP ? rowP[colP[HP.CP_ESTADO]] : "",
    procComent: rowP ? rowP[colP[HP.CP_COMENT]] : ""
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
 * Lista todos os PDFs que começam com o prefixo da primeira semana
 */
function listPdfs(ticket, numA, semana, modo) {
  console.log("listPdfs(" + numA + ", "  + semana + ", " + modo + ")");
  AuthCoreLib.requireSession(ticket);
  
  const semDash = semana.replace("/", "-"); //Em 'semana' já só vem a primeira semana (104/30)
  console.log("listPdfs: semDash=" + semDash);
  if (!semDash || semDash === "---") return [];

  let folderId = (modo === "PROC") ? FOLDER_PROCURACOES : FOLDER_IPS;
  const rootFolder = DriveApp.getFolderById(folderId);

  // A query exata que pediu: procura pelo ID do associado OU pela semana
  let query = `(title contains '${numA}' or title contains '${semDash}') and mimeType = 'application/pdf' and trashed = false`;

  const candidates = getAllFilesRecursive_(rootFolder, query);
  console.log("candidates.length=" + candidates.length);
  console.log("candidates=" + candidates);

  //// Filtro estrito: apenas os que começam exatamente pelo prefixo (ignora case)
  //const strictPattern = new RegExp("^" + prefix, "i");
  
  if (candidates.length > 0) {
    const matches = candidates
      //.filter(f => strictPattern.test(f.getName()))
      .map(f => ({
        name: f.getName(),
        url: `https://drive.google.com/file/d/${f.getId()}/preview`
      }));

    // Ordenar alfabeticamente para facilitar a leitura no select
    return matches.sort((a, b) => a.name.localeCompare(b.name));
  }
  
  console.warn("Aviso: Nenhum ficheiro PDF encontrado para: " + numA + " ou " + semDash + " (mesmo em subpastas)");
  return [];
}

function saveRecordData(ticket, numA, payload) {
  console.log("saveRecordData(" + numA + "," + JSON.stringify(payload) + ")");
  AuthCoreLib.requireSession(ticket);
  
  // Update Titulares
  const ss = SpreadsheetApp.openById(SS_TITULARES_ID);
  const shT = ss.getSheetByName("Titulares");
  // Procuramos o numA na coluna B (B7:B)
  const dataT = shT.getRange("B:B").getValues().flat();
  //console.log("dataT=" + dataT);
  const idxT = dataT.indexOf(numA) + 1;
  if (idxT > 0) {
    const colsT = getTitularesColumns_(shT);
    // Para atualizar A, C e D sem corromper o ID na coluna B,
    // escrevemos o bloco de 4 colunas (A a D), reinserindo o numA na coluna B.

    // O ideal seria ler a linha toda, modificar as colunas desejadas, usando nomes automáticos, e depois escrever a linha toda
    shT.getRange(idxT, 1, 1, 4).setValues([[
      payload.nomeMembros, // Coluna A (index 0 no load)
      numA,                // Coluna B (ID - mantemos o valor original)
      payload.nif,         // Coluna C (index 2 no load)
      payload.nomeFiscal   // Coluna D (index 3 no load)
    ]]);

    console.log("payload.estado=" + payload.estado);    
    shT.getRange(idxT, 17, 1, 1).setValues([[payload.estado]]); //CT_ESTADO
  }


  // Update IPS (Colunas K, L, M, N)
  const shI = SpreadsheetApp.openById(SS_IPS_ID).getSheetByName("IPS");
  const dataI = shI.getRange("A:A").getValues().flat();
  const idxI = dataI.indexOf(numA) + 1;

  const ipsRowData = [
    payload.ipsData, 
    payload.ipsEstado, 
    payload.ipsComentario, 
    payload.primTitular, 
    payload.outrosTitulares
  ];

  if (idxI > 0) {
    // No loadRecordData, o ipsComentario é rowI[11].
    // O index 11 num array que começa em A corresponde à 12ª coluna (Coluna L).
    shI.getRange(idxI, 10, 1, 5).setValues([ipsRowData]);
  } else {
    // Ponto 9: Cria nova linha se o ID não existir na tblIPS
    const newRow = [numA, "", "", "", "", "", "", "", ""]; // Colunas A-I vazias
    shI.appendRow(newRow.concat(ipsRowData));
  }
  

    // Update Procuracoes
  const shP = SpreadsheetApp.openById(SS_PROC_ID).getSheetByName("Procurações");
  // Procuramos o numA na coluna A (A5:A)
  const dataP = shP.getRange("A:A").getValues().flat();
  //console.log("dataP=" + dataP);
  const idxP = dataP.indexOf(numA) + 1;
  if (idxP > 0) {
    //console.log("idxP=" + idxP);
    //const colsP = getProcuracoesColumns_(shP);
    // escrevemos o bloco de 8 colunas (J a Q)

    shP.getRange(idxP, 10, 1, 8).setValues([[
      payload.procPart, 
      payload.procData,         
      payload.procForm,
      payload.procCert,
      payload.procCc,
      payload.procDocs,
      payload.procEstado,
      payload.procComent
    ]]);
  }


  return { ok: true };
}



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