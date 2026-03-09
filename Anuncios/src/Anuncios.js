
// =========================
// File: Anuncios.gs
// =========================


/***********************************************
 * Este ficheiro serve dois propósitos:
 * 1 - Suportar fórmulas no spreadsheet, tal como Pedidos!G3=findTelefoneIndex(C3)
 * 2 - Suportar a web app dos anúncios
**********************************************/


/***********************************************
* Fórmulas em uso no spreadsheet
 **********************************************/

 const TITULARES_SHEET = "Titulares"; //Usado por uma fórmula na aba "Pedidos"

function findTelefoneIndex(numero) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const titulares = ss.getSheetByName(TITULARES_SHEET);
    const data = titulares.getRange('K2:K').getValues();
    const cleanedNumero = numero.toString().replace(/\D/g, '');
    //Logger.log('numero=' + numero + ', cleanedNumero=' + cleanedNumero);
    for (let i = 0; i < data.length; i++) {
        const linha = data[i][0];
        if (typeof linha === 'string') {
            const telefones = linha.split(';').map(t => t.replace(/\D/g, ''));
            for (const tel of telefones) {
                if (tel.endsWith(cleanedNumero)) return i; // índice base 0 do array (linha real = i + 2)
            }
        }
    }
    return 0; // Not found
}



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


const VERSION = "v6.0";

// Folhas pertencentes APENAS a esta App
const ANUNCIOS_SHEET = "Anúncios";
const ANUNCIOS_HEADER_ROW = 4;
const PEDIDOS_SHEET = "Pedidos";

//const ACESSOS_SHEET = "Acessos"; //Vai deixar de ser usada aqui, porque AuthCoreLib passa a gerir os acessos. Há que mover para tblTitulares.

// ---------- Auth wrappers chamados pelo Login.html (biblioteca) ----------
// Note: Como a AuthCoreLib agora tem defaults inteligentes, só precisamos do redirectUri
function gatesCfg_(){ return { canon: ScriptApp.getService().getUrl() }; }


function authCfg_() {
  const sp = PropertiesService.getScriptProperties();
  return {
    clientId:     sp.getProperty("CLIENT_ID"),
    clientSecret: sp.getProperty("CLIENT_SECRET"),
    //redirectUri:  sp.getProperty("REDIRECT_URI") || ScriptApp.getService().getUrl(),
  };
}

function buildAuthUrlFor(nonce, dbg, embed, clientUrl) { return AuthCoreLib.buildAuthUrlFor(nonce, dbg, embed, authCfg_(), clientUrl); }
function pollTicket(nonce)                  { return AuthCoreLib.pollTicket(nonce); }
function isTicketValid(ticket, dbg)         { return AuthCoreLib.isTicketValid(ticket, dbg); }

// Funções ponte para o RGPD HTML (agora delegadas à AuthCoreLib)
function listRgpdRowsFor(ticket) { return AuthCoreLib.hostListRgpdRowsFor(ticket, gatesCfg_()); }

// ===== DEBUG infra =====
function isDebug_(e){ return !!(e && e.parameter && (e.parameter.debug === '1' || e.parameter.Debug === '1')); }

function makeLogger_(DBG) {
  const start = new Date();
  const log = [];
  function L() {
    const t = new Date(new Date() - start).toISOString().substr(11, 8);
    log.push(t + " " + Array.prototype.map.call(arguments, String).join(" "));
  }
  L.dump = () => log.slice();
  return L;
}

function renderRgpdPage_(ticket, DBG, serverLogLines) {
  const opts = { ticket: ticket, debug: DBG, serverLog: serverLogLines, wipe: false };    
  opts.serverLog.unshift("RGPD");
  return AuthCoreLib.renderRgpdPage(opts);
}

// ===== Routing Principal =====
function doGet(e) {
  const DBG = isDebug_(e);
  const L = makeLogger_(DBG);
  L("doGet start Anúncios");

  var optsProc = { ticket: "", debug: DBG, serverLog: [], wipe: false };    
  const canon = ScriptApp.getService().getUrl();
  L('e.parameter.action=' + e.parameter.action);
  L('e.parameter.ticket=' + e.parameter.ticket);
  L('e.parameter.__canon=' + e.parameter.__canon);
  L('e.parameter.code=' + e.parameter.code);
  L('e.parameter.state=' + e.parameter.state);
  L('e.parameter.logout=' + e.parameter.logout);
  L('e.parameter.debug=' + e.parameter.debug);
  L('e.parameter.wipe=' + e.parameter.wipe);
  //L('here=' + ScriptApp.getService().getUrl());
  //L('canon=' + canonicalAppUrl_());
  //L('REDIRECT_URI=' + (REDIRECT_URI||''));
  //L('redirectUri_()=' + redirectUri_());


  const action = (e && e.parameter && e.parameter.action) || "";

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
    return renderMainPage_(ticket, DBG, L.dump());
  }

  if (action === "login") { optsProc.serverLog = L.dump(); optsProc.serverLog.unshift("login"); return AuthCoreLib.renderLoginPage(optsProc); }
  if (action === "logout") { optsProc.serverLog = L.dump(); optsProc.serverLog.unshift("logout"); optsProc.wipe = true; return AuthCoreLib.renderLoginPage(optsProc); }

  // 3) Validar Ticket & Regras Estritas
  const ticket = (e && e.parameter && e.parameter.ticket) || "";
  if (ticket){
    try { AuthCoreLib.requireSession(ticket); } 
    catch(err){ optsProc.serverLog = L.dump(); optsProc.serverLog.unshift("wipe (Motivo: " + err.message + ")"); optsProc.wipe = true; return AuthCoreLib.renderLoginPage(optsProc); }
    if (String(e.parameter.go || '') === 'rgpd') { return renderRgpdPage_(ticket, DBG, L.dump()); }
    if (String(e.parameter.go || '') === 'main') { return renderMainPage_(ticket, DBG, L.dump()); }

    // Utiliza a biblioteca para ler a ficha do Associado!
    const profile = AuthCoreLib.getProfileStats(ticket, gatesCfg_());

    // Regra 1: RGPD DEVE ESTAR TOTALMENTE ACEITE
    if (profile.rgpdState !== 'total') {
      L("RGPD pendente ou parcial. Bloquear.");
      return renderRgpdPage_(ticket, DBG, L.dump());
    }

    // Regra 2: Saldo global deve ser positivo/zero (e deve ser titular)
    if (!profile.hasLines || profile.saldo < 0) {
      L("Acesso Negado: Não tem linhas ou o Saldo é negativo: " + profile.saldo);
      const motivo = !profile.hasLines ? "Não é titular (0 linhas)" : "Saldo negativo (€" + profile.saldo + ")";
      AuthCoreLib.logFailedAccess(ticket, motivo, gatesCfg_()); // <-- REGISTA NA FOLHA ACESSOS
      return renderBlocked_("Acesso negado: Para aceder aos anúncios é necessário ser titular registado e não ter quotas em dívida.", DBG);
    }

    // OK -> Main
    return renderMainPage_(ticket, DBG, L.dump());
  }

  optsProc.serverLog = L.dump(); optsProc.serverLog.unshift("sem ticket");
  return AuthCoreLib.renderLoginPage(optsProc);
}

function renderBlocked_(msg, DBG) {
  const canon = ScriptApp.getService().getUrl();
  const html = `<!doctype html><html><head><meta charset="utf-8"><title>Acesso Negado</title></head>
  <body style="font-family:sans-serif; padding: 2em; text-align: center; max-width: 600px; margin: 0 auto;">
      <h3 style="color: #dc2626;">${msg}</h3>
      <p style="margin-top: 2em;"><a href="${canon}?action=logout${DBG ? '&debug=1' : ''}" target="_top" style="padding:10px 20px; background:#111; color:#fff; text-decoration:none; border-radius:5px;">Sair / Terminar Sessão</a></p>
  </body></html>`;
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderMainPage_(ticket, DBG, serverLogLines) {
  const t = HtmlService.createTemplateFromFile('Main');
  t.VERSION = VERSION;
  const visits = incrementVisitCounters();
  t.count = visits.total;
  t.count30 = visits.last30; // Passa os últimos 30 dias para o HTML
  t.ticket = ticket || "";  
  t.CANON_URL = ScriptApp.getService().getUrl(); 
  t.DEBUG = DBG ? '1' : '';
  t.SERVER_LOG = (serverLogLines || []).join('\n');
  t.PAGE_TAG = 'MAIN';
  const out = t.evaluate().setTitle('Anúncios Portobelo');
  out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return out;
}

// =============================
// Interação de Dados / Anúncios
// =============================
function validateApiAccess_(ticket) {
  AuthCoreLib.requireSession(ticket);
  const profile = AuthCoreLib.getProfileStats(ticket, gatesCfg_());
  if (profile.rgpdState !== 'total' || profile.saldo < 0) {
    throw new Error("Acesso negado às operações de Anúncios. Saldo/RGPD pendentes.");
  }
}

function incrementVisitCounters() {
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

  return { total: total, last30: last30 };
}

function getUltimaAtualizacao() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ANUNCIOS_SHEET);
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TITULARES_SHEET);
    const row = index + 2;
    const emailCell = sheet.getRange(row, 12).getValue(); // L
    const duesCell = sheet.getRange(row, 3).getValue(); // C
    const allEmails = emailCell.toString().split(';').map(e => e.trim().toLowerCase());
    return allEmails.includes((email || '').toLowerCase()) && (Number(duesCell) || 0) >= 1;
}

function getAnunciosDataSess(ticket) { 
  validateApiAccess_(ticket); // Bloqueia API a quem não cumpra os requisitos
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ANUNCIOS_SHEET);
  const numRows = sheet.getLastRow() - ANUNCIOS_HEADER_ROW;
  if (numRows <= 0) return [];
  return sheet.getRange(ANUNCIOS_HEADER_ROW + 1, 1, numRows, 5).getValues().map(row => row.map(cell => {
    if (typeof cell === 'string') return cell.replace(/[\r\n]+/g, ' ').trim();
    return (cell !== undefined && cell !== null) ? cell.toString() : '';
  }));
}

function submitAnuncioSess(ticket, anuncioParcial) { 
  validateApiAccess_(ticket);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PEDIDOS_SHEET) || ss.insertSheet(PEDIDOS_SHEET);

  if (!Array.isArray(anuncioParcial) || anuncioParcial.length !== 5) throw new Error('Dados inválidos');
  if (sheet.getLastRow() >= 50) throw new Error('Limite de 50 anúncios atingido.');

  sheet.appendRow([new Date(), ...anuncioParcial.slice(1)]);
  SpreadsheetApp.flush();

  const email = 'geral@titulares-portobelo.pt';
  MailApp.sendEmail(email, 'Novo anúncio submetido: ' + anuncioParcial[2], 'Telemóvel=' + anuncioParcial[2] + '\nTipo=' + anuncioParcial[1] + '\n' + anuncioParcial[3]);
  return true; 
}

function testFind() {
  const result = findDuesPaidByEmail("pal.mendes23@gmail2.com", "Pedro Mendes");
  Logger.log(result);
}