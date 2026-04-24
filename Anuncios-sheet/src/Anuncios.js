
// =========================
// File: Anuncios.gs - projeto independente do Google Sheets
// =========================


/***********************************************
*
* Google Apps Script: Web App para exibir a aba "Anúncios" com filtros e ordenação
*
*  Acesso restrito aos titulares contribuintes da associação, até porque não é permitido por lei alugar timeshares sem licença de alojamento local
*
*  ****** Apps Script ***********
*
*  "Execute as" = "Me" (para o web app ler/escrever no Sheets da associação)
*  "Who has access" = "Anyone" (ou "Anyone with Google Account").
*
*  O domínio titulares-portobelo.pt pertence à nossa associação, e é gerido com OVH.
*
*
*  ********** TESTES *********************
*  Abre a página dos anúncios no Sites.
*  Carrega F12 para abrir a consola.
*  No topo da consola há um menu de contexto (no teu screenshot aparece “top”).
*  Clica e escolhe o frame chamado userCodeAppPanel (é o iframe da web app).
*  Isto é importante: o localStorage a apagar está nesse frame, não em top.
*  Com o contexto certo selecionado, corre:
*
*  //Para consultar o ticket guardado
*  separador Application → Storage → Cookies → clica no domínio https://script.google.com e procura a cookie sessTicket
*  localStorage.getItem('sessTicket');
*  document.cookie
*  //Para apagar o ticket
*  localStorage.removeItem('sessTicket');
*  document.cookie = 'sessTicket=; Max-Age=0; Path=/; Secure; SameSite=Lax';
*  location.reload();
*
*  //Para validar o polling sem CORS
*  fetch('./?route=poll&nonce=teste&_=' + Date.now(), {mode:'same-origin'})
*    .then(r => r.text()).then(console.log).catch(console.error);
*
*  Se o cache ainda não tiver ticket para esse nonce, deves ver {"ok":false}. Depois do login, verás {"ok":true,"ticket":"..."}
*  e a página navega para a Main.
*
*  de quiseres “desligar” o sticky depois, apaga o flag indo a about:blank e na consola:
*  localStorage.removeItem('pbDebug')
*
**********************************************/


const VERSION = "v7.4";


const SS_ANUNCIOS_ID  = "1oacSvYMYrcJeaUV9XFLcylrAX2PJGodGjpeeHl96Smo";
const ANUNCIOS_SHEET = "Anúncios";
//const ANUNCIOS_HEADER_ROW = 4;


// === CONFIG: ranges nomeados ou A1 (fallback) ===
const RANGES = {
  titulares: { name: "tblTitulares", sheet: "Titulares", a1: "A6:AA" },
  anuncios:  { name: "tblAnuncios",  sheet: "Anúncios",   a1: "A4:K" },
};

// Cabeçalhos esperados na Folha Anúncios
const COLS_ANUNCIOS = {
  CA_DATA: "Anúncio",
  CA_TIPO: "Tipo",
  CA_TEL: "Telemóvel",
  CA_DESC: "Descrição",
  CA_ENT: "Entrada",
  CA_PRECO: "Preço",
  CA_APROV: "Aprovação",
  CA_EMAIL: "e-mail",
  CA_SALDO: "Saldo",
  CA_SEMS: "Semanas"
};

// Função utilitária para injetar ficheiros HTML noutros ficheiros HTML
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------- Auth wrappers chamados pelo Login.html (biblioteca) ----------
// Note: A AuthCoreLib agora tem defaults inteligentes
function gatesCfg_(){ return { canon: ScriptApp.getService().getUrl(), appName: "Anúncios" }; }

function authCfg_() {
  const sp = PropertiesService.getScriptProperties();
  return {
    //appTitle: "authcfg app title",
    //appPermissions: "authcfg app permissions",
    clientId:     sp.getProperty("CLIENT_ID"),
    clientSecret: sp.getProperty("CLIENT_SECRET"),
  };
}

function buildAuthUrlFor(nonce, dbg, embed, clientUrl) { return AuthCoreLib.buildAuthUrlFor(nonce, dbg, embed, authCfg_(), clientUrl); }
function pollTicket(nonce)                  { return AuthCoreLib.pollTicket(nonce); }
function isTicketValid(ticket, dbg)         { return AuthCoreLib.isTicketValid(ticket, dbg); }
function getProfileStats(ticket) { return AuthCoreLib.getProfileStats(ticket, authCfg_()); }

// Funções ponte para o RGPD HTML (agora delegadas à AuthCoreLib)
function listRgpdRowsFor(ticket) { return AuthCoreLib.hostListRgpdRowsFor(ticket, gatesCfg_()); }

// ===== DEBUG infra =====
function isDebug_(e){ return !!(e && e.parameter && (e.parameter.debug === '1' || e.parameter.Debug === '1')); }

function makeLogger_(DBG) {
  const start = new Date();
  const log = [];
  function L() {
    const ts = new Date();
    const t = new Date(ts - start).toISOString().substr(11, 8);
    log.push(t + " " + Array.prototype.map.call(arguments, String).join(" "));
  }
  L.dump = () => log.slice();
  return L;
}

function renderRgpdPage_(ticket, DBG, serverLogLines, blockOpts) {
  const opts = { ticket: ticket, debug: DBG, serverLog: serverLogLines, wipe: false };    
  if (blockOpts) {
     opts.blockMsg = blockOpts.msg;
     opts.blockBtnLabel = blockOpts.btnLabel;
     opts.blockBtnUrl = blockOpts.btnUrl;
  }
  opts.serverLog.unshift("RGPD");
  return AuthCoreLib.renderRgpdPage(opts);
}

// ===== Routing Principal =====
function doGet(e) {
  const DBG = isDebug_(e);
  const L = makeLogger_(DBG);
  console.log("doGet start Anúncios");

  // Passamos o canon forçado para o renderLoginPage da AuthCoreLib
  var optsProc = {
    ticket: "",
    debug: DBG, 
    serverLog: [], 
    wipe: false,
    appTitle: "Consulta de anúncios",
    appPermissions: "Os associados com quotas em dia e RGPD totalmente aceite podem ver e publicar todo o tipo de anúncios. Os outros utilizadores podem apenas ver os anúncios de 'Venda na Internet' e publicar anúncios de 'compra na Internet'.",
  };    

  const action = (e && e.parameter && e.parameter.action) || "";

  const canon = ScriptApp.getService().getUrl();
  L('e.parameter.action=' + e.parameter.action);
  L('e.parameter.ticket=' + e.parameter.ticket);
  L('e.parameter.__canon=' + e.parameter.__canon);
  L('e.parameter.code=' + e.parameter.code);
  L('e.parameter.state=' + e.parameter.state);
  L('e.parameter.logout=' + e.parameter.logout);
  L('e.parameter.debug=' + e.parameter.debug);
  L('e.parameter.wipe=' + e.parameter.wipe);

  // 1) OAuth callback
  if (e && e.parameter && e.parameter.code){ return AuthCoreLib.finishAuth(e, authCfg_()); }

  // 2) Utilitários / Resets
  if (action === "reset"){
    const next = canon + (DBG ? '?debug=1' : '');
    const html = `<!doctype html><html><head><meta charset="utf-8">
      <style>body{font-family:system-ui,sans-serif;padding:18px} .btn{display:inline-block;padding:10px 16px;background:#000;color:#fff;text-decoration:none;border-radius:6px;}</style>
      </head><body><h3>Cookies limpos.</h3><p><a class="btn" href="${next}" target="_top">Ir para o Início</a></p>
      <script>try{ localStorage.removeItem('sessTicket'); }catch(_){} try{ document.cookie = 'sessTicket=; Max-Age=0; Path=/; SameSite=Lax; Secure'; }catch(_){}
      var url = "${next}"; try{ if (window.top !== window.self) window.top.location.href = url; else location.replace(url); } catch(e){}
      </script></body></html>`;
    return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (action === "rgpd") { return renderRgpdPage_(e.parameter.ticket || '', DBG, L.dump()); }

  if (action === "postrgpd") {
      const ticket = e.parameter.ticket || '';
      const rowsQS = String(e.parameter.rows || '').trim();
      const rows   = rowsQS ? rowsQS.split(',').map(s => parseInt(s, 10)).filter(n => Number.isFinite(n)) : [];
      try { AuthCoreLib.hostSaveRgpdRowsFor(ticket, rows, gatesCfg_()); } catch(err) { L('RGPD save FAIL: ' + err); }
  
      // Deixamos o código continuar (NÃO fazemos return aqui).
      // Ao continuar para a validação abaixo, ele vai ler o perfil já atualizado
      // e aplicar a Regra do Saldo normalmente!
      // SEGURANÇA: Em vez de mostrar a grelha direta, recarrega a app principal 
      //return renderMainPage_(ticket, DBG, L.dump());
      // para obrigar o sistema a validar se o utilizador tem o Saldo Positivo!
      //const nextUrl = forceCanon + "?ticket=" + encodeURIComponent(ticket) + (DBG ? "&debug=1" : "");
      //return HtmlService.createHtmlOutput(`<script>window.top.location.href="${nextUrl}";</script>`).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

  if (action === "login") {
    optsProc.serverLog = L.dump(); 
    optsProc.serverLog.unshift("login");
    console.log("doGet ACTION LOGIN");
    //console.log("doGet action login: optsProc.appTitle=" + optsProc.appTitle );
    //console.log("doGet action login: optsProc.appPermissions=" + optsProc.appPermissions);
    return AuthCoreLib.renderLoginPage(optsProc);
  }
  
  if (action === "logout") { optsProc.serverLog = L.dump(); optsProc.serverLog.unshift("logout"); optsProc.wipe = true; return AuthCoreLib.renderLoginPage(optsProc); }

  // 3) Validar Ticket & Regras Estritas de Negócio
  const ticket = (e && e.parameter && e.parameter.ticket) || "";
  if (ticket){
    try { AuthCoreLib.requireSession(ticket); } 
    catch(err) {
      console.log("doGet error ticket");
      optsProc.serverLog = L.dump(); 
      optsProc.serverLog.unshift("wipe (Motivo: " + err.message + ")"); 
      optsProc.wipe = true; 
      return AuthCoreLib.renderLoginPage(optsProc); 
    }
	
	 // Carrega o Perfil antes de qualquer decisão de ecrã
	 
    //if (String(e.parameter.go || '') === 'main') { return renderMainPage_(ticket, DBG, L.dump()); }

    // Utiliza a biblioteca para ler a ficha do Associado!
    const profile = AuthCoreLib.getProfileStats(ticket, gatesCfg_());
    L(`[DEBUG PERFIL] Email: ${profile.email} | Linhas: ${profile.hasLines} | RGPD: ${profile.rgpdState} | Saldo: €${profile.saldo}`);

    // AVALIADOR DE ESTADO DO UTILIZADOR
    let userStatus = "";
    if (!profile.hasLines || profile.pago < 1) {
        userStatus = "Não é associado";
    } else if (profile.rgpdState !== 'total') {
        userStatus = "RGPD não aceite";
    } else if (profile.saldo < 0) {
        userStatus = "Saldo negativo";
    }
	
    // REGRA 0: Se não for Titular (0 Linhas), Bloqueia Absoluto Imediato
    if (!profile.hasLines || profile.pago < 1) {
      console.log("Bloqueio: Acesso Negado. Motivo: Não é titular (0 linhas) ou nunca pagou quotas");
      /*
      AuthCoreLib.logFailedAccess(ticket, "Não é titular", gatesCfg_()); 
      return renderBlocked_(
          "Acesso negado: O seu e-mail não está registado como titular.", 
          "Ir para Titulares Portobelo", 
          "https://titulares-portobelo.pt", 
          DBG, L.dump()
      );
      */
    }

    // Pedido manual para ver a página de RGPD
    if (String(e.parameter.go || '') === 'rgpd') { return renderRgpdPage_(ticket, DBG, L.dump()); }

    // Regra 1: RGPD DEVE ESTAR TOTALMENTE ACEITE
    if (profile.rgpdState !== 'total') {
      console.log("Bloqueio: RGPD está '" + profile.rgpdState + "'. Exigindo ida à página local de RGPD.");
      /*
      AuthCoreLib.logFailedAccess(ticket, "RGPD pendente ou parcial", gatesCfg_());
      return renderRgpdPage_(ticket, DBG, L.dump(), {
          msg: "Para aceder à secção de anúncios é obrigatório aceitar o RGPD para todas as suas frações/semanas. Pode fazê-lo agora mesmo usando o formulário abaixo."
          //btnLabel: "Ou vá para a Área do Associado",
          //btnUrl: "https://associados.titulares-portobelo.pt"
      });
      */
    }

  
    // Regra 2: Saldo global deve ser positivo/zero
    if (profile.saldo < 0) {
      const motivo = "Saldo negativo (€" + profile.saldo + ")";
	    console.log("Bloqueio: Acesso Negado. Motivo: " + motivo);
      /*
      AuthCoreLib.logFailedAccess(ticket, "Saldo negativo", gatesCfg_()); // <-- REGISTA NA FOLHA ACESSOS
      return renderBlocked_(
          "Acesso negado: Para aceder aos anúncios é necessário não ter quotas em dívida (o seu saldo global é de €" + profile.saldo + ").", 
          "Ir para a Área do Associado", 
          "https://associados.titulares-portobelo.pt", 
          DBG, L.dump()
      );
      */
    }
  
    // OK -> Main e Mostra a grelha de Anúncios
    console.log("Sucesso: Acesso aos Anúncios concedido.");
    return renderMainPage_(ticket, DBG, L.dump(), userStatus);
  }

  optsProc.serverLog = L.dump(); optsProc.serverLog.unshift("sem ticket");
  console.log("doGet no action");
  return AuthCoreLib.renderLoginPage(optsProc);
}

// O ecrã de bloqueio foi otimizado para mostrar botões dinâmicos e os logs na consola
function renderBlocked_(msg, btnLabel, btnUrl, DBG, serverLogLines) {
  const canon = ScriptApp.getService().getUrl();
  const safeLog = (serverLogLines || []).join('\\n').replace(/`/g, '\\`'); // Protege quebras de linha

  // Se passarmos texto e URL para o botão, criamos o HTML do botão
  let btnHtml = "";
  if (btnLabel && btnUrl) {
    btnHtml = `
      <p style="margin-top: 2.5em;">
          <a href="${btnUrl}" target="_top" style="padding:12px 24px; background:#111; color:#fff; text-decoration:none; border-radius:8px; font-weight:500; font-size:16px;">${btnLabel}</a>
      </p>`;
  }

  const html = `<!doctype html><html><head><meta charset="utf-8"><title>Acesso Negado</title>
  <script>
      const SERVER_LOG = \`${safeLog}\`;
      console.log("[SERVER LOG]\\n" + SERVER_LOG);
  </script>
  </head>
  <body style="font-family:sans-serif; padding: 2em; text-align: center; max-width: 600px; margin: 0 auto;">
      <h3 style="color: #dc2626; line-height: 1.4;">${msg}</h3>
      
      ${btnHtml}
      
      <p style="margin-top: 2em;">
          <a href="${canon}?action=logout${DBG ? '&debug=1' : ''}" target="_top" style="color:#666; text-decoration:underline; font-size:14px;">Sair / Terminar Sessão</a>
      </p>
  </body></html>`;

  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderMainPage_(ticket, DBG, serverLogLines, userStatus) {
  const t = HtmlService.createTemplateFromFile('Main');
  t.VERSION = VERSION;
  const visits = incrementVisitCounters(userStatus);
  t.count = visits.total;
  t.count30 = visits.last30; // Passa os últimos 30 dias para o HTML
  t.countGuests = visits.guests || 0; // <-- Nova variável
  t.ticket = ticket || "";  

  //t.CANON_URL = ScriptApp.getService().getUrl(); 
  const sp = PropertiesService.getScriptProperties();
  t.CANON_URL = ScriptApp.getService().getUrl();

  t.DEBUG = DBG ? '1' : '';
  t.SERVER_LOG = (serverLogLines || []).join('\n');
  t.PAGE_TAG = 'MAIN';

  // INJETA O ESTADO NO ECRÃ
  t.USER_STATUS = userStatus || "";
  
  const out = t.evaluate().setTitle('Anúncios Portobelo');
  out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return out;
}

// =============================
// Interação de Dados / Anúncios
// =============================
function validateApiAccess_(ticket) {
  AuthCoreLib.requireSession(ticket);

  // Já não lançamos erro de Saldo/RGPD aqui. O sistema agora aceita
  // utilizadores com limitações de leitura e escrita em vez de os expulsar.
  return true;

  const profile = AuthCoreLib.getProfileStats(ticket, gatesCfg_());
  if (profile.rgpdState !== 'total' || profile.saldo < 0) {
    throw new Error("Acesso negado às operações de Anúncios. Saldo/RGPD pendentes.");
  }
}

function incrementVisitCounters(userStatus) {
  const props = PropertiesService.getScriptProperties();
  
  // 1. Total Global
  const total = parseInt(props.getProperty('visitCount') || '0', 10) + 1;
  props.setProperty('visitCount', total.toString());

  // 2. Histórico dos últimos 30 dias (por dia)
  const today = new Date().toISOString().slice(0, 10); // YYYY-MM-DD
  let history = {};
  try { history = JSON.parse(props.getProperty('visitHistory') || '{}'); } catch(e){}
  history[today] = (history[today] || 0) + 1;

  // 3. Limpar dias velhos e somar o mês
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 30);
  const cutoffStr = cutoffDate.toISOString().slice(0, 10);
  
  let last30 = 0;
  for (const dateStr in history) {
    if (dateStr < cutoffStr) {
      delete history[dateStr]; // Apaga os registos com mais de 30 dias
    } else {
      last30 += history[dateStr];
    }
  }
  props.setProperty('visitHistory', JSON.stringify(history));

  let countGuests = parseInt(props.getProperty('VISITAS_GUESTS') || '0', 10);
  if (userStatus === "Não é associado") {
    countGuests++;
    props.setProperty('VISITAS_GUESTS', countGuests.toString());
  }

  // 3. NO RETURN EXISTENTE, adicione a propriedade guests:
  return { 
    total: total, 
    last30: last30, 
    guests: countGuests
  };
}

function getUltimaAtualizacao() {
  const sheet = SpreadsheetApp.openById(SS_ANUNCIOS_ID).getSheetByName(ANUNCIOS_SHEET);
  return sheet ? sheet.getRange('E1').getDisplayValue() : 'Indisponível';
}

/**
 * Regista uma linha na aba "Acessos" (tabela tblAcessos) a partir da coluna B:
 * A=Data, B=Nome, C=Email, D=€, E=Semanas
 */
function logDeniedAccess_(ss, name, email, dues, weeks) {
  const sht = ss.getSheetByName(ACESSOS_SHEET); if (!sht) return;

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const row = sht.getLastRow() + 1;
  const values = [[ts, String(name || ''), String(email || ''), Number(dues) || 0, Number(weeks) || 0]];
  sht.getRange(row, 1, 1, 5).setValues(values);
}

function findDuesPaidByEmail(email, name) {
  const ss = SpreadsheetApp.openById(SS_ANUNCIOS_ID);
  const sh = ss.getSheetByName(TITULARES_SHEET);
  if (!sh) return false;

  const emails = sh.getRange('L2:L').getValues().flat();
  const dues = sh.getRange('C2:C').getValues().flat();
  const weeks = sh.getRange('I2:I').getValues().flat();
  const target = (email || '').toLowerCase().trim();
  let c = 0, w = 0;

  for (let i = 0; i < emails.length; i++) {
    const list = (emails[i] || '').toString().split(';').map(function (x) { return x.trim().toLowerCase(); });
    if (list.indexOf(target) !== -1) {
      c = (Number(dues[i]) || 0);
      w = (Number(weeks[i]) || 0);
      if (c >= 1 && c >= w) return true;  // Pelo menos 1 euro por semana
    }
  }
  logDeniedAccess_(ss, name, email, c, w);
  return false;
}

function findRowByEmailAndPhone(email, telefoneDisplay) {
  // Verifica se, na mesma linha, o email (col L) contém o email e a coluna K contém o telefone indicado
  const ss = SpreadsheetApp.openById(SS_ANUNCIOS_ID);
  const sh = ss.getSheetByName(TITULARES_SHEET);
  if (!sh) return false;

  const phones = sh.getRange('K2:K').getValues().flat();
  const emails = sh.getRange('L2:L').getValues().flat();

  const targetEmail = (email || '').toLowerCase().trim();
  const telDigits = String(telefoneDisplay || '').replace(/\D/g, ''); // limpeza para digitos

  for (let i = 0; i < phones.length; i++) {
    const rowPhones = String(phones[i] || '');
    const rowEmails = String(emails[i] || '');
    const hasEmail = rowEmails.split(';').some(function (x) { return x.trim().toLowerCase() === targetEmail; });
    const hasPhone = rowPhones.split(';').some(function (p) {
      const d = String(p || '').replace(/\D/g, '');
      // aceitar fim coincidente (para permitir +351 vs 351 vs sem espaços)
      return d && telDigits && d.endsWith(telDigits);
    });
    if (hasEmail && hasPhone) return i + 2; // linha real na sheet
  }
  return 0;
}


function findDuesPaid(index, email) {
    if (index < 0)
        return false;
    const ss = SpreadsheetApp.openById(SS_ANUNCIOS_ID);
    const sheet = ss.getSheetByName(TITULARES_SHEET);
    const row = index + 2;
    const emailCell = sheet.getRange(row, 12).getValue(); // L
    const duesCell = sheet.getRange(row, 3).getValue(); // C
    const allEmails = emailCell.toString().split(';').map(e => e.trim().toLowerCase());
    return allEmails.includes((email || '').toLowerCase()) && (Number(duesCell) || 0) >= 1;
}


function getAnunciosDataSess(ticket) { 
  // 1. Forçamos o Logger a arrancar em modo de debug
  const DBG = true; //isDebug_(e);
  const L = makeLogger_(DBG);
  console.log("getAnunciosDataSess()");

  try {
    validateApiAccess_(ticket); // Bloqueia API a quem não cumpra os requisitos
    console.log("getAnunciosDataSess: Acesso da API validado com sucesso.");
    const profile = AuthCoreLib.getProfileStats(ticket, gatesCfg_());
    console.log("getAnunciosDataSess: Perfil lido. Email associado: " + profile.email);
	
	  // Define o estatuto do utilizador (Membro Completo vs Guest)
    const isEmDia = (profile.hasLines && profile.pago >= 1 && profile.saldo >= 0 && profile.rgpdState === 'total');
    console.log("isemDia=" + isEmDia);

    // Extrai todos os números de telemóvel do associado
    const myPhones = new Set();
    if (profile.cardsLinhas) {
      profile.cardsLinhas.forEach(c => {
        if (c.telefones) c.telefones.split(/[;,]/).forEach(t => {
          console.log("getAnunciosDataSess: telefone=" + t);
          const d = String(t).replace(/\D/g, '');
          console.log("getAnunciosDataSess: telefone transformado=" + d);
          // O SEU FILTRO: Ignora os números de controlo financeiro (ex: 351000...)
          if (d && !d.startsWith('351000') && !d.startsWith('000')) {
            console.log("getAnunciosDataSess: telefone adicionado");
            myPhones.add(d);
          }
        });
      });
    }
    console.log("getAnunciosDataSess: Telemóveis válidos encontrados: " + Array.from(myPhones).join(', '));
    const myPhonesArr = Array.from(myPhones);
    
    const ownsPhone = (phoneStr) => {
      const d = String(phoneStr).replace(/\D/g, '');
      return myPhonesArr.some(my => my.endsWith(d) || d.endsWith(my));
    };

    const ss = SpreadsheetApp.openById(SS_ANUNCIOS_ID);
    /*
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ANUNCIOS_SHEET);
    const numRows = Math.max(0, sheet.getLastRow() - ANUNCIOS_HEADER_ROW);
    console.log("getAnunciosDataSess: numRows=", numRows);
    */
    let range;
    try { range = ss.getRangeByName(RANGES.anuncios.name); } catch(e) {}
	  console.log("getAnunciosDataSess: getRangeByName"); 
    if (!range) {
      range = ss.getSheetByName(RANGES.anuncios.sheet).getRange(RANGES.anuncios.a1);
      console.log("getAnunciosDataSess: getSheetByName"); 
    }
  
    const values = range.getValues();
    if (values.length === 0) return { ads: [], myPhones: myPhonesArr, isEmDia: isEmDia };
    console.log("getAnunciosDataSess: values.length=" + values.length);

    const headers = values[0];
    const colIdx = {};
    for (let key in COLS_ANUNCIOS) colIdx[key] = headers.indexOf(COLS_ANUNCIOS[key]);
    
    // Fallbacks de segurança caso não encontre o nome da coluna
    if (colIdx.CA_DATA === -1) colIdx.CA_DATA = 0;
    if (colIdx.CA_TIPO === -1) colIdx.CA_TIPO = 1;
    if (colIdx.CA_TEL === -1) colIdx.CA_TEL = 2;
    if (colIdx.CA_DESC === -1) colIdx.CA_DESC = 3;
    if (colIdx.CA_ENT === -1) colIdx.CA_ENT = 4;
    if (colIdx.CA_PRECO === -1) colIdx.CA_PRECO = 5;
    if (colIdx.CA_APROV === -1) colIdx.CA_APROV = 6;
	
    let ads = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row[colIdx.CA_DATA]) continue; // Ignora linhas vazias
      
      let tipo = String(row[colIdx.CA_TIPO] || '').trim();
      let tel = String(row[colIdx.CA_TEL] || '').trim();
      let aprovacao = String(row[colIdx.CA_APROV] || '').trim();
      //console.log("Anuncio: tipo=" + tipo + ", tel=" + tel + ", aprovacao=" + aprovacao);
      
      // Nova lógica de propriedade: pertence-lhe se o telefone coincidir OU se o e-mail da Coluna H coincidir
      let isMyAd = ownsPhone(tel) || (String(row[colIdx.CA_EMAIL] || '').trim().toLowerCase() === profile.email.toLowerCase());
      
      // REGRAS DE VISUALIZAÇÃO DE MARKETPLACE
      if (!isEmDia) {
        // Guests: Só vêem "Vende na Internet" (aprovados) ou os SEUS PRÓPRIOS "Compra na Internet" (qualquer estado)
        if (tipo !== "Venda na Internet" && !(tipo === "Compra na Internet" && isMyAd)) continue;
        if (tipo === "Venda na Internet" && aprovacao !== "Sim") continue;
      } else {
        // Members em Dia: Veem tudo o que está aprovado, mais os seus próprios pendentes/rejeitados
        // Members em Dia: Veem tudo o que está aprovado, mais os seus próprios pendentes/rejeitados
        if (aprovacao !== "Sim" && !isMyAd) continue;
      }
      
      // Normalizar datas
      let dataFormatada = "";
      if (row[colIdx.CA_DATA] instanceof Date) {
        dataFormatada = Utilities.formatDate(row[colIdx.CA_DATA], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
	    // Fallback: se for um texto antigo como 20/10/2024, reorganiza
        let s = String(row[colIdx.CA_DATA] || '').trim();
        let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
        if (m) dataFormatada = `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
        else dataFormatada = s.split(' ')[0];
      }
      
      ads.push([
        dataFormatada,
        tipo,
        tel,
        String(row[colIdx.CA_DESC] || '').trim(),
        String(row[colIdx.CA_ENT] || '').trim(),
        String(row[colIdx.CA_PRECO] || '').trim(),
        aprovacao,
        isMyAd // NOVO: Índice 7 diz ao ecrã se o utilizador é dono do anúncio
      ]);
    }
    console.log("Tudo OK. A devolver o pacote ao frontend. Anúncios totais na folha: " + ads.length);
    
    // 2. Descarrega os logs para a plataforma do Google
    //console.log(L.dump().join('\n'));    
    // Agora devolve a tabela E os contactos aprovados!
    return { ads: ads, myPhones: myPhonesArr, isEmDia: isEmDia };

  } catch (error) {
    console.log("ERRO FATAL: " + error.message);
    // Em caso de erro, descarrega os logs também
    //console.log(L.dump().join('\n'));
    throw error; // Lança o erro para que o F12 o mostre como "Erro ao carregar dados"
  }
}


function apiGestaoAnuncios(ticket, payload) {
  // 1. Forçamos o Logger a arrancar em modo de debug
  //const DBG = true; //isDebug_(e);
  //const L = makeLogger_(DBG);
  console.log("apiGestaoAnuncios()");

  validateApiAccess_(ticket);   // 1. Acesso blindado (Se não tiver quotas ou RGPD, falha aqui e nem chega a executar o resto)
  console.log("Acesso da API validado com sucesso.");
  const profile = AuthCoreLib.getProfileStats(ticket, gatesCfg_());
  console.log("Perfil lido. Email associado: " + profile.email);
  
  const isEmDia = (profile.rgpdState === 'total' && profile.saldo >= 0 && profile.hasLines);
	
  const allEmails = new Set();
  const myPhones = new Set();
  
  // Recolhe e-mails para CC e Telefones para controlo de acesso
  if (profile.email) allEmails.add(profile.email.toLowerCase());
  if (profile.cardsLinhas) {
    profile.cardsLinhas.forEach(c => {
      if (c.emails) c.emails.split(/[;,]/).forEach(e => {
        console.log("email =" + String(e));
        const clean = String(e).trim().toLowerCase();
        if (clean) allEmails.add(clean);
      });
      if (c.telefones) c.telefones.split(/[;,]/).forEach(t => {
        console.log("telefone =" + String(t));
        const d = String(t).replace(/\D/g, '');
        if (d) myPhones.add(d);
      });
    });
  }
  console.log("Telemóveis válidos encontrados: " + Array.from(myPhones).join(', '));
  
  const ccEmails = Array.from(allEmails).join(',');
  console.log("emails válidos encontrados: " + ccEmails);

  const myPhonesArr = Array.from(myPhones);
  // Validação de Propriedade Híbrida (Telemóvel ou E-mail)
  const isMyAd = (phoneStr, emailStr) => {
    const d = String(phoneStr).replace(/\D/g, '');
    return myPhonesArr.some(my => my.endsWith(d) || d.endsWith(my)) || 
           (String(emailStr).trim().toLowerCase() === profile.email.toLowerCase());
  };

  console.log("Phone OK");

  const ss = SpreadsheetApp.openById(SS_ANUNCIOS_ID);
  //const sheet = ss.getSheetByName(ANUNCIOS_SHEET); // 2. Agora aponta DIRETAMENTE para a folha principal de Anúncios
  let range;
  try { range = ss.getRangeByName(RANGES.anuncios.name); } catch(e) {}
  const sheet = range ? range.getSheet() : ss.getSheetByName(RANGES.anuncios.sheet);
  
  const dataStartRow = range ? range.getRow() + 1 : 5; // A4 cabeçalhos, A5 dados
  const numAnunciosTotal = Math.max(0, sheet.getLastRow() - dataStartRow + 1);
  
  const allData = sheet.getDataRange().getDisplayValues();
  let myAdsCount = 0;
  for (let r = dataStartRow - 1; r < allData.length; r++) {
    if (isMyAd(allData[r][2], allData[r][7])) myAdsCount++;
  }
  
  // Função auxiliar para atualizar sempre a data em E1 com formato legível
  const atualizarDataE1 = () => {
    const agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    try { sheet.getRange('E1').setValue(agora); } catch(e){}
  };

  console.log("Conteúdo completo: " + JSON.stringify(payload));
  console.log(" payload.novo.length=" +  payload.novo.length);

  // AÇÃO: PUBLICAR / ATUALIZAR
  if (payload.action === 'add' || payload.action === 'update') {
    console.log("add");
    if (!payload.novo || !Array.isArray(payload.novo) || payload.novo.length !== 6) {
      throw new Error('Dados inválidos. O array de dados tem de ter 6 posições e tem ' + payload.novo.length);
    }
    const tipo = payload.novo[1];
    const preco = (payload.novo[5] || "").trim();
	    
    // Regra de Submissão Condicional
    if (!isEmDia && tipo !== 'Compra na Internet') throw new Error("Apenas pode publicar anúncios do tipo 'Compra na Internet'.");
    if (isEmDia && tipo === 'Compra na Internet') throw new Error("Associados com quotas em dia não usam o tipo 'Compra na Internet'.");
    if (tipo === "Venda na Internet" && !preco) throw new Error("O preço é obrigatório para anúncios de Venda na Internet.");
    
    // Limites de Segurança
    if (payload.action === 'add') {
      if (isEmDia) {
        if (myAdsCount >= 20) throw new Error("Limite pessoal atingido (Max. 20 anúncios publicados).");
      } else {
        if (numAnunciosTotal >= 500) throw new Error("Limite global da plataforma atingido. Não é possível publicar.");
        if (myAdsCount >= 2) throw new Error("Limite pessoal atingido (Max. 2 anúncios publicados para o seu escalão).");
      }
    }

    // Estado de Aprovação
    let aprovacao = (tipo.includes("Internet")) ? "Pendente" : "Sim";
    if (!isEmDia) aprovacao = "Pendente"; 
    
	  const dataFormatada = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
	
    // Força o Google Sheets a aceitar o número de telemóvel como texto, preservando os espaços!
    let telFormatado = "'" + payload.novo[2];
	
	  // Colunas A a H preparadas (confiamos na ordem rígida conforme requisito 8)
    const rowToWrite = [dataFormatada, tipo, telFormatado, payload.novo[3], payload.novo[4], preco, aprovacao, profile.email];
    if (payload.action === 'add') {
      sheet.appendRow(rowToWrite);
    } else {
      const orig = payload.original;
      let foundRow = -1;
      for (let r = allData.length - 1; r >= dataStartRow - 1; r--) {
        if (allData[r][1] === orig[1] && allData[r][2] === orig[2] && allData[r][3] === orig[3]) {
          if (!isMyAd(allData[r][2], allData[r][7])) throw new Error("Segurança: Anúncio não lhe pertence.");
          foundRow = r + 1; break;
        }
      }
      if (foundRow === -1) throw new Error("O anúncio original já não foi encontrado.");
      sheet.getRange(foundRow, 1, 1, 8).setValues([rowToWrite]);
	  }
    
    atualizarDataE1();
    SpreadsheetApp.flush();
    console.log("flush");
    try {
      MailApp.sendEmail({
        to: 'log-apps@titulares-portobelo.pt', cc: ccEmails, replyTo: 'geral@titulares-portobelo.pt',
        subject: 'Novo anúncio AUTOMÁTICO: ' + payload.novo[2],
        body: profile.email + ' publicou um anúncio diretamente na plataforma.\n\nTelemóvel: ' + payload.novo[2] + '\nTipo: ' + payload.novo[1] + '\nDescrição: ' + payload.novo[3]
      });
    } catch(e) {}
    return { ok: true };
  }
  
  
  /*
  // 2. ATUALIZAR EXISTENTE
  if (payload.action === 'update') {
    console.log("update");

    const data = sheet.getDataRange().getDisplayValues();
    const orig = payload.original;
    let foundRow = -1;
    
    // Procura de baixo para cima para encontrar a versão mais recente
    for (let r = data.length - 1; r >= ANUNCIOS_HEADER_ROW; r--) {
      const rowData = data[r];
      if (rowData[1] === orig[1] && rowData[2] === orig[2] && rowData[3] === orig[3]) {
        if (!ownsPhone(rowData[2])) throw new Error("Segurança: Este anúncio não pertence aos seus contactos.");
        foundRow = r + 1; break;
      }
    }
    
    if (foundRow === -1) throw new Error("O anúncio original já não foi encontrado.");
    
    // Força o formato de texto
    payload.novo[2] = "'" + payload.novo[2];
    const dataFormatada = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.getRange(foundRow, 1, 1, 5).setValues([[dataFormatada, ...payload.novo.slice(1)]]);
    atualizarDataE1();
    SpreadsheetApp.flush();
    console.log("flush");

    try {
      MailApp.sendEmail({
        to: 'log-apps@titulares-portobelo.pt', cc: ccEmails, replyTo: 'geral@titulares-portobelo.pt',
        subject: 'Anúncio ATUALIZADO pelo titular: ' + payload.novo[2],
        body: 'Um titular atualizou o seu anúncio.\n\nTelemóvel: ' + payload.novo[2] + '\nTipo: ' + payload.novo[1] + '\nNova Descrição: ' + payload.novo[3]
      });
    } catch(e) {}
    console.log("fim");
    return { ok: true };
  }
  */
  
  // 3. APAGAR
  if (payload.action === 'delete') {
    const toDelete = [];
    
    for (let r = allData.length - 1; r >= dataStartRow - 1; r--) {
      const rowData = allData[r];
      const matchIdx = payload.originais.findIndex(orig => rowData[1] === orig[1] && rowData[2] === orig[2] && rowData[3] === orig[3]);
      
      if (matchIdx !== -1) {
        if (!isMyAd(rowData[2], rowData[7])) throw new Error("Segurança falhou. Anúncio não lhe pertence.");
        // Guest pode apagar os seus (A validação de que só veem "Comprar" na sua lista já os protege, mas duplo check:)
        if (!isEmDia && rowData[1] !== "Compra na Internet") throw new Error("Apenas pode apagar anúncios 'Compra na Internet'.");
        
        toDelete.push(r + 1);
        payload.originais.splice(matchIdx, 1);
      }
    }
    
    // Apaga de baixo para cima para não alterar os índices das linhas de cima
    toDelete.sort((a, b) => b - a).forEach(rowIdx => sheet.deleteRow(rowIdx));
    atualizarDataE1();
    SpreadsheetApp.flush();
    
    try {
      MailApp.sendEmail({
        to: 'log-apps@titulares-portobelo.pt', cc: ccEmails, replyTo: 'geral@titulares-portobelo.pt',
        subject: 'Anúncio(s) APAGADO(S)',
        body: profile.email + ' removeu ' + toDelete.length + ' anúncio(s) da plataforma.'
      });
    } catch(e) {}
    return { ok: true };
  }
  throw new Error("Ação inválida");
}

function testFind() {
  const result = findDuesPaidByEmail("pal.mendes23@gmail2.com", "Pedro Mendes");
  Logger.log(result);
}