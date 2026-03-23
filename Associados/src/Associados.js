
// =========================
// Associados.gs (no projeto independente)
// =========================

/*************************
 * 
 * Utilizado só para a web app.
 * Foi separado do spreadsheet para se poder partilhar com os associados, mas nem isso resolveu
 *  o problema do Android acrescentar /u/1 aos URL da web app quando existe mais do que uma conta Google no telemóvel.
 *  Não há solução para esse problema, a não ser incorporar a página no Google Sites, mas isso traz outros problemas ainda piores
 *  porque os telemóveis tornam-se muito exigentes com a navegação entre páginas.
 * 
 ************************/


/*

  SSO entre apps: se quiseres que um login numa app valha na outra,
  copia para as Script Properties os MESMOS valores de
  SESSION_HMAC_SECRET_B64 e STATE_SECRET.


  action=reset: limpa só o que está no browser (localStorage/cookies) para aquele origin. É o que precisas para o caso do iframe.
  action=logout: faz o mesmo wipe, mas já vem com um redirect para ?action=login. É um “sair” clássico.
  exec?action=login&debug=1&noredirect=1
  exec?action=logout&debug=1
  exec?action=reset&debug=1
  exec?action=who   => endpoint de diagnóstico.


  - O download de FT-IPS passa por endpoints da própria app (não expõe Drive).
  - Mantemos action=ips (teste) e adicionamos action=ipszip (uso normal).
  
  ######### NÃO FAZER login quando se chama com reset!  #########
  Para testes, entra sempre por …/exec?action=login&debug=1 (ou reset) — evita o logout para não haver corridas com location.replace.
  …/exec?action=reset&debug=1 (limpa storage/cookies sem redirecionar) e depois vai para action=login
  
  Consola Google Cloud (OAuth 2.0 clients): não apaga os teus tickets — eles são stateless (HMAC) e vivem no browser do utilizador.

  Se queres invalidar todos os tickets do mundo de uma vez, roda a chave:
  muda a Script Property SESSION_HMAC_SECRET_B64 (ou apaga-a para o código gerar outra).
  Todos os tickets existentes deixam de validar e toda a gente tem de relogar.

  Para obrigar a fazer login:
    1 - No navegador, Abre DevTools (F12) → tab Console
    2 - No topo da consola, muda o frame (dropdown) para o iframe do script.googleusercontent.com / script.google.com
    3 - Executa: location.href = (window.__SERVER_VARS?.CANON_URL || window.__SERVER_VARS?.canonUrl || window.CANON_URL) + "?action=reset";

  Para mandar para o ChatGPT, filtrar a consola F12 do browser:
    Navigated|Login|cookie|ticket|userCodeAppPanel|/exec\?|action=|LOGIN|MAIN|RGPD|Rgpd


  Emails autorizados para testes => criar a propriedade de script:
  ALLOWLIST_CSV=pal.mendes23@gmail.com:xxx@yyy.zzz


  - A web app deve estar implantada como: Executar como "Eu (admin@…)" e
    Acesso: "Qualquer pessoa com o link".
*/

const VERSION = "v1.20";

// === CONFIG: IDs das 3 folhas ===
const SS_TITULARES_ID = "1YE16kNuiOjb1lf4pbQBIgDCPWlEkmlf5_-DDEZ1US3g";
const SS_ANUNCIOS_ID  = "1oacSvYMYrcJeaUV9XFLcylrAX2PJGodGjpeeHl96Smo";
const SS_IPS_ID       = "1rsxlHYHrXfSdpgjfhA193xYkhi3r-RMK9cc04l7Sqq8";
const IPS_FOLDER_ID   = "1TXL942FE_Z05gSCJ_f_1lj_nEPnD5DWv";

// === CONFIG: ranges nomeados ou A1 (fallback) ===
const RANGES = {
  titulares: { name: "tblTitulares", sheet: "Titulares", a1: "A6:Z" },
  anuncios:  { name: "tblAnuncios",         sheet: "Anúncios",   a1: "A4:H" },
  ips:       { name: "tblIPS",      sheet: "IPS",        a1: "A4:N" },
  quotas:    { name: "tblQuotas",   sheet: "Quotas",     a1: "A1:D" },
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

// Coluna Telefone dos anúncios para mostrar na Área do Associado quais os anúncios que publicou, e a que telemóveis estão associados.
const COLS_ANUNCIOS = {
  CA_EMAIL: "e-mail",   //Coluna nova, onde se deve escrever (na web app dos anúncios) o e-mail do associado que publicou o anúncio
  CA_TEL: "Telemóvel", 
};

// ---------- Auth wrappers chamados pelo Login.html (biblioteca) ----------
// Lê as Script Properties do host e constrói a cfg
function authCfg_() {
  const sp = PropertiesService.getScriptProperties();
  //const canon = ScriptApp.getService().getUrl().replace(/\/a\/[^/]+\/macros/, "/macros");
  const canon = ScriptApp.getService().getUrl();
  return {
    clientId:     sp.getProperty("CLIENT_ID"),
    clientSecret: sp.getProperty("CLIENT_SECRET"),
  };
}
function buildAuthUrlFor(nonce, dbg, embed, clientUrl) { return AuthCoreLib.buildAuthUrlFor(nonce, dbg, embed, authCfg_(), clientUrl); }
function pollTicket(nonce)                  { return AuthCoreLib.pollTicket(nonce); }
function isTicketValid(ticket, dbg)              { return AuthCoreLib.isTicketValid(ticket, dbg); }
function acceptRgpdForMe(ticket, decision){
  // decision: 'accept' | 'reject'
  return AuthCoreLib.acceptRgpdForMe(ticket, decision, gatesCfg_());
}

// Funções ponte para o RGPD HTML
function listRgpdRowsFor(ticket) { return AuthCoreLib.hostListRgpdRowsFor(ticket, gatesCfg_()); }


// ===== DEBUG infra =====

// Utilitário de logs (cai para Google Apps Script executions logs, se console não existir)  //log em Execuções Apps Script
function dbgLog() {
  var msg = Array.prototype.map.call(arguments, String).join(' ');

  try { Logger.log(msg); } catch(_) { try { console.log(msg); } catch(__) {} }
}

function isDebug_(e){
  const p = e && e.parameter ? e.parameter : {};
  if (!p || (!Object.prototype.hasOwnProperty.call(p, "debug") && !Object.prototype.hasOwnProperty.call(p, "Debug"))) {
    DBG = false;
  } else {
    //return v === "" || v === "1" || v === "true";
    DBG = true;
  }
  const v = String(p.debug || p.Debug || "").toLowerCase();
  console.log("isDebug_() => DBG=", DBG, "v = ", v);
  dbgLog("Associados: isDebug_() => DBG=", DBG, "v = ", v);
  return DBG;
}

function makeLogger_(DBG) {
  const start = new Date();
  const log = [];
  function L() {
    //if (!DBG) return;
    const ts = new Date();
    const t = new Date(ts - start).toISOString().substr(11, 8); // hh:mm:ss desde início
    log.push(t + " " + Array.prototype.map.call(arguments, String).join(" "));
  }
  L.dump = () => log.slice();
  return L;
}

function gatesCfg_(){
  //const canon = ScriptApp.getService().getUrl().replace(/\/a\/[^/]+\/macros/, '/macros');
  const canon = ScriptApp.getService().getUrl();
  return {
    ssTitularesId: SS_TITULARES_ID,
    ranges: { titulares: RANGES.titulares },
    cols:   { 
      email: COLS_TITULARES.CT_EMAIL, 
      rgpd: COLS_TITULARES.CT_RGPD, 
      pago: COLS_TITULARES.CT_PAGO,
      saldo: COLS_TITULARES.CT_SALDO
    },
    notify: { to: "log-apps@titulares-portobelo.pt", ccAllRows: true },
    canon: canon,
    appName: "Associados",
  };
}

function renderRgpdPage_(ticket, DBG, serverLogLines) {
  const opts = {
    ticket: ticket,
    debug: DBG,    
    serverLog: serverLogLines,
    wipe: false,
  };    
  //opts.serverLog = L.dump();
  opts.serverLog.unshift("RGPD");
  return AuthCoreLib.renderRgpdPage(opts);
}

function getRgpdStatus(ticket){
  return AuthCoreLib.getRgpdStatusFor(ticket, gatesCfg_());
}


// ===== Main.html (template da app) =====

function renderMainPage_(ticket, DBG, serverLogLines){
  const t = HtmlService.createTemplateFromFile("Main");
  //t.CANON_URL  = ScriptApp.getService().getUrl().replace(/\/a\/[^/]+\/macros/, "/macros");
  t.CANON_URL  = ScriptApp.getService().getUrl();
  t.DEBUG      = DBG ? "1" : ""; // '1' se quiseres forçar debug visual
  t.TICKET     = ticket || "";
  t.SERVER_LOG = (serverLogLines || []).join("\n");
  t.PAGE_TAG = 'MAIN';
  const out = t.evaluate();
  out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return out;
}

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
function cellHasEmail_(cell, emailLC){
  const all = String(cell||"").split(/[;,]/).map(s=>s.trim().toLowerCase()).filter(Boolean);
  return all.includes(emailLC);
}
function splitPhones_(cell){ return String(cell||"").split(/[;,]/).map(s=>s.replace(/[^\d]/g,"")).filter(Boolean); }
function getFirstPhone_(cell){ const a=splitPhones_(cell); return a.length?a[0]:""; }
function splitSemanas_(cell){ return String(cell||"").split(/[+,\s]+/).map(s=>s.trim()).filter(Boolean); }
function dedupe_(arr){ return Array.from(new Set(arr)); }
function parsePtNumber_(s){
  const clean = String(s||"").replace(/\s/g,"").replace(/\./g,"").replace(",",".").replace(/[^\d.\-]/g,"");
  const n = parseFloat(clean); return isNaN(n) ? 0 : n;
}

// ===== IPS helpers =====
function parseIpsDateScore_(name){
  // e.g.: "104-30 FT-IPS 2025-08-19.pdf", "204-34 FT-IPS 2024.pdf", "104-30 FT-IPS 2020x.pdf"
  const m = String(name||"").match(/FT-IPS\s+(\d{4})(?:-(\d{2}))?(?:-(\d{2}))?x?\.pdf$/i);
  if (!m) return 0;
  const y = +m[1], mo = m[2] ? +m[2] : 0, d = m[3] ? +m[3] : 0;
  return (y*10000 + mo*100 + d);
}

function findLatestIpsFileForWeek_(week) {
  dbgLog('[IPS] findLatest… week=', week, 'FOLDER_ID=', IPS_FOLDER_ID);
  var folder;
  try {
    folder = DriveApp.getFolderById(IPS_FOLDER_ID);
    dbgLog('[IPS] folder name=', folder.getName());
  } catch (e) {
    dbgLog('[IPS] getFolderById FAIL:', e && e.message, e && e.stack);
    throw e;
  }

  var prefix = String(week).replace('/', '-') + ' FT-IPS ';
  var q = "title contains '" + prefix.replace(/'/g, "\\'") + "' and mimeType = 'application/pdf'";
  dbgLog('[IPS] query=', q);

  var it;
  try {
    it = folder.searchFiles(q);
  } catch (e) {
    dbgLog('[IPS] searchFiles FAIL:', e && e.message, e && e.stack);
    throw e;
  }

  var best=null, bestScore=0, count=0;
  try {
    while (it.hasNext()) {
      var f = it.next(); count++;
      var name = f.getName();
      var score = parseIpsDateScore_(name);
      dbgLog('[IPS] hit:', name, 'score=', score);
      if (score >= bestScore){ bestScore=score; best=f; }
    }
    dbgLog('[IPS] total hits=', count, 'bestScore=', bestScore, 'best=', best && best.getName());
  } catch (e) {
    dbgLog('[IPS] iterate FAIL:', e && e.message, e && e.stack);
    throw e;
  }
  return best;
}

// Lê todas as linhas de Titulares com um certo email (match robusto)
function listWeeksForEmail_(email){
  const emailLC = String(email||"").trim().toLowerCase();
  const { header, rows } = fetchTable_(SS_TITULARES_ID, RANGES.titulares);
  const col = indexByHeader_(header);
  const iEmail = col[COLS_TITULARES.CT_EMAIL];
  const iSem   = col[COLS_TITULARES.CT_SEMANAS];
  if (iEmail==null || iSem==null) return [];
  const weeks = [];
  rows.forEach(r=>{
    if (cellHasEmail_(r[iEmail], emailLC)){
      splitSemanas_(r[iSem]).forEach(w => weeks.push(w));
    }
  });
  return dedupe_(weeks);
}
function isWeekOfEmail_(email, week){ return listWeeksForEmail_(email).includes(String(week||"").trim()); }

// ===== Acesso/allowlist/rgpd =====

function parseAllowlist_(){
  const csv = (PropertiesService.getScriptProperties().getProperty("ALLOWLIST_CSV") || "");
  return csv.split(/[,\s;]+/).map(s=>s.trim().toLowerCase()).filter(Boolean);
}

// E-mails da lista de validação
function isAllowedEmail_(email){
  console.log("isAllowedEmail_ =>", email);
  dbgLog("isAllowedEmail_ =>", email);
  const list = parseAllowlist_();
  console.log("isAllowedEmail_ => list.length", list.length);
  if (!list.length) return true;// ← limitação OFF quando property vazia
  return list.includes(String(email||"").toLowerCase());
}

function renderNotAllowed_(email, DBG){
  const sp    = PropertiesService.getScriptProperties();
  //const canon = ScriptApp.getService().getUrl().replace(/\/a\/[^/]+\/macros/, "/macros");
  const canon = ScriptApp.getService().getUrl();
  const dbgBlock = DBG ? (function(){
    const raw = sp.getProperty("ALLOWLIST_CSV") || "(vazio)";
    const parsed = parseAllowlist_().join(", ");
    return `<details style="margin-top:12px"><summary>DEBUG</summary>
      <pre>Email sessão: ${escapeHtml_(email)}
ALLOWLIST_CSV (raw): ${escapeHtml_(raw)}
ALLOWLIST parsed: ${escapeHtml_(parsed)}</pre></details>`;
  })() : "";
  const out = HtmlService.createHtmlOutput(
    '<meta charset="utf-8">' +
    "<h3>Acesso não autorizado</h3><p>Este endereço não está na lista da fase de validação.</p>" +
    `<p><a href="${canon}?action=logout${DBG ? "&debug=1" : ""}" target="_top">Terminar sessão</a></p>` +
    dbgBlock
  );
  return out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function isRgpdAccepted(ticket){
  const sess = AuthCoreLib.requireSession(ticket);
  const emailLC = String(sess.email||"").trim().toLowerCase();
  const { header, rows } = fetchTable_(SS_TITULARES_ID, RANGES.titulares);
  const col = indexByHeader_(header);
  const iEmail = col[COLS_TITULARES.CT_EMAIL];
  const iRGPD  = col[COLS_TITULARES.CT_RGPD];
  if (iEmail==null || iRGPD==null) return false;
  for (const r of rows){
    if (cellHasEmail_(r[iEmail], emailLC) && String(r[iRGPD]||"").trim().toLowerCase()==="sim")
      return true;
  }
  return false;
}


// ===== Página principal: view/API =====
function buildAssociadosView_(loginEmail){
  const dbg=[]; function D(){ try{ dbg.push([(new Date()).toISOString().slice(11,19), [].map.call(arguments,String).join(" ")].join(" ")); }catch(_){ } }
  //D("build view for ", loginEmail);
  console.log("build view for ", loginEmail);

  const { header: th, rows: titulares } = fetchTable_(SS_TITULARES_ID, RANGES.titulares);
  const col = indexByHeader_(th); const H = COLS_TITULARES;

  const emailLC = String(loginEmail).trim().toLowerCase();
  const linhas = titulares.filter(r => cellHasEmail_(r[col[H.CT_EMAIL]], emailLC));
  console.log("linhas do email:", linhas.length);
  
  const semanasTodas = dedupe_(linhas.flatMap(r => splitSemanas_(r[col[H.CT_SEMANAS]] || "")));
  console.log("semanasTodas:", semanasTodas);

  const { header: ih, rows: ipsRows } = fetchTable_(SS_IPS_ID, RANGES.ips);
  const icol = indexByHeader_(ih);
  const ipsBySemanas = {}; //Isto é para deixar de ser assim, e passar-e a usar numA
  ipsRows.forEach(r => {
    const key = String(r[icol[COLS_IPS.CI_SEMANAS]]||"").trim();
    if (key) ipsBySemanas[key] = r;
  });

  const cards=[]; const tot={joia:0, quota:0, pago:0, quotizacao:0, saldo:0};
  const allPhones=[]; const firstPhones=[]; const allNumA=[];

  linhas.forEach(r=>{
    const nomes   = r[col[H.CT_MEMBROS]] || "";
    const numA   = r[col[H.CT_NUM]] || "";
    const nif   = r[col[H.CT_NIF]] || "";
    const nomeFiscal   = r[col[H.CT_NOMEFISCAL]] || "";

    const pago    = r[col[H.CT_PAGO]] || "";
    const semanas = r[col[H.CT_SEMANAS]] || "";

    const iT0 = col["T0"] ?? 5, iT1 = col["T1"] ?? 6, iT2 = col["T2"] ?? 7;
    const t0 = parseInt(r[iT0],10)||0, t1 = parseInt(r[iT1],10)||0, t2 = parseInt(r[iT2],10)||0;

    const adesaoRaw = r[col[H.CT_ADESAO]] || "";
    const telef     = r[col[H.CT_TEL]]    || "";
    const emails    = r[col[H.CT_EMAIL]]  || "";
    const fimRaw    = r[col[H.CT_FIM]]  || "";
    const estado    = col[H.CT_ESTADO] >= 0 ? r[col[H.CT_ESTADO]] || "" : "";
    const quota     = r[col[H.CT_QUOTA]] || "";
    const joia      = r[col[H.CT_JOIA]]  || "";
    const saldo     = r[col[H.CT_SALDO]] || "";

    const rgpdOk    = String(r[col[H.CT_RGPD]]||"").trim().toLowerCase()==="sim";

    // Obter valores numéricos para efetuar a matemática
    const pagoNum  = parsePtNumber_(pago);
    const quotaNum = parsePtNumber_(quota);
    const joiaNum  = parsePtNumber_(joia);
    const saldoNum = parsePtNumber_(saldo);
    const saldoNeg = saldoNum < 0;

    // A Quotização Total é simplesmente tudo o que já se pagou menos o saldo (positivo ou negativo)
    const quotizacaoNum = pagoNum - saldoNum;

    // Atualizar Acumuladores Globais
    tot.pago       += pagoNum;
    tot.quota      += quotaNum;
    tot.joia       += joiaNum;
    tot.saldo      += saldoNum;
    //tot.quotizacao += quotizacaoNum;
    tot.quotizacao = (tot.quotizacao || 0) + quotizacaoNum;

    console.log("numA=", numA, ", pago=", pago, ", tot.pago=", tot.pago); //Só mostra a linha do registo.

    if (numA) allNumA.push(String(numA).trim().toUpperCase());
    console.log("allNumA=", JSON.stringify(allNumA));
    console.log("allNumA=", allNumA);

    allPhones.push(...splitPhones_(telef));
    const fp = getFirstPhone_(telef); if (fp) firstPhones.push(fp);

    const ip = ipsBySemanas[String(semanas).trim()]; //Substituir por IpsByNumA
    const dataIPS = ip ? (ip[icol[COLS_IPS.CI_DATA]]||"") : "";
    const statIPS = ip ? (ip[icol[COLS_IPS.CI_STATUS]]||"") : "";
    const primIPS = ip ? (ip[icol[COLS_IPS.CI_PRIMEIRO]]||"") : "";
    const outIPS  = ip ? (ip[icol[COLS_IPS.CI_OUTROS]]||"") : "";

    cards.push({
      numA, nif, nomeFiscal, semanas, t0, t1, t2, adesaoRaw, fimRaw, estado, dataIPS, statIPS, primIPS, outIPS,
      nomes, emails, telefones: telef, pago, quota, joia, saldo,
      quotizacao: quotizacaoNum,
      rgpdOk, saldoNeg
    });
  });

  const telefonesAssociados = dedupe_(allPhones);
  const primeirosTelefones  = dedupe_(firstPhones);
  const registosAssociados  = dedupe_(allNumA);
  console.log("registosAssociados=", registosAssociados);
  console.log("registosAssociados=", JSON.stringify(registosAssociados));
  const anunciosPorTelefone = countAnunciosByPhones_(telefonesAssociados); //OK assim.
  const transacoes          = fetchTransacoes_(registosAssociados);
  console.log("transacoes=", JSON.stringify(transacoes));
  console.log("transacoes=", transacoes);

  return {
    user: { email: loginEmail },
    cardsLinhas: cards,
    totais: tot,
    telefonesAssociados, primeirosTelefones, registosAssociados,
    anunciosPorTelefone, transacoes,
    semanasTodas,
  };
}

function apiGetAssociados(ticket){
  console.log("enter apiGetAssociados");
  const sess = AuthCoreLib.requireSession(ticket);
  if (!isAllowedEmail_(sess.email)) throw new Error("Acesso não autorizado");
  const view = buildAssociadosView_(sess.email);
  view.user = view.user || {};
  view.user.name    = sess.name    || "";
  view.user.picture = sess.picture || "";
  console.log("exit apiGetAssociados");
  return view;
}

function apiGetWeekIpsBase64(ticket, week){
  const sess = AuthCoreLib.requireSession(ticket);
  week = String(week||'').trim();
  if (!week) throw new Error('Falta week');
  if (!isWeekOfEmail_(sess.email, week)) throw new Error('Sem autorização');

  dbgLog('[IPS] apiGetWeekIpsBase64 start, week=', week, 'email=', sess.email);
  const f = findLatestIpsFileForWeek_(week); // <- sem L
  if (!f) { dbgLog('[IPS] no file found'); return { ok:false, reason:'not_found' }; }

  dbgLog('[IPS] file chosen:', f.getId(), f.getName());
  const blob = f.getBlob();
  const b64  = Utilities.base64Encode(blob.getBytes());
  return { ok:true, name: f.getName(), mime: blob.getContentType() || 'application/pdf', b64 };
}

// ===== Anúncios / Transações =====
function countAnunciosByPhones_(phones){
  if (!phones || !phones.length) return {};
  const phoneSet = new Set(phones);
  const { header: th, rows} = fetchTable_(SS_ANUNCIOS_ID, RANGES.anuncios);
  const col = indexByHeader_(th); const H = COLS_ANUNCIOS;
  const counts = {};
  rows.forEach(r=>{
    const raw  = r[col[H.CA_TEL]];
    const norm = String(raw||"").replace(/[^\d]/g,"");
    if (norm && phoneSet.has(norm)) counts[norm] = (counts[norm]||0) + 1;
  });
  return counts;
}

function fetchTransacoes_(registos){
   if (!registos || !registos.length) return [];
  const want = new Set((registos||[]).map(r => String(r).trim().toUpperCase()).filter(Boolean));
  console.log("fetchTransacoes_ -> want:", Array.from(want));
  
  if (want.size === 0) return [];
  const { header, rows } = fetchTable_(SS_TITULARES_ID, RANGES.transacoes);
  const col = indexByHeader_(header);
  const iDate = col["Date"]   ?? 1; // B
  const iNum  = col["Number"] ?? 3; // D
  const iAmt  = col["Amount"] ?? 6; // G

  const out=[];
  rows.forEach(r=>{
    const raw = String(r[iNum]||"").trim();
    //dbgLog('fetchTransacoes_: date=', r[iDate], ', ref=', r[iNum], ', amount=', r[iAmt]);
    if (!raw) return;
    
    const asRegisto = raw.toUpperCase();
    
    if (asRegisto && want.has(asRegisto)) {
      console.log("fetchTransacoes_ push");
      out.push({ date:r[iDate]||"", ref:raw, amount:r[iAmt]||"", raw:{row:r} });
    }
  });
  return out;
}

// ===== doGet router =====
function doGet(e){
  const DBG = isDebug_(e);
  console.log("doGet() starg => DBG=", DBG);
  const L = makeLogger_(DBG);
  L("doGet start");
  Logger.log(AuthCoreLib.libBuild()); //Registo de execuções Apps Script
  L(AuthCoreLib.libBuild());

  var optsProc = {
    ticket: "",
    debug: DBG,    
    serverLog: [],
    wipe: false,
  };    

  //const canon = ScriptApp.getService().getUrl().replace(/\/a\/[^/]+\/macros/, "/macros");
  const canon = ScriptApp.getService().getUrl();
  const action = (e && e.parameter && e.parameter.action) || "";
  L("action=", action || "(none)");
  //L("code=", (e && e.parameter && e.parameter.code) || "(none)");
  //if (String(e.parameter.go || '') === 'main') {
  L("go=", (e && e.parameter && e.parameter.go) || "(none)");
  //L("ticket: (e && e.parameter && e.parameter.ticket) || (none)=", (e && e.parameter && e.parameter.ticket) || "(none)");
  //L("ticket: e?.parameter?.ticket || (none)=", e?.parameter?.ticket || "(none)");
  L("ticket: e?.parameter?.ticket=", e?.parameter?.ticket);

  // 1) OAuth callback
  if (e && e.parameter && e.parameter.code){ L("route: finishAuth"); return AuthCoreLib.finishAuth(e, authCfg_()); }

  // 2) utilitários
  if (action === "ping") { L("route: ping"); return HtmlService.createHtmlOutput("OK"); }
  if (action === "who")  {
    L("route: who");
    try {
      const s = AuthCoreLib.requireSession(e.parameter.ticket||"");
      return HtmlService.createHtmlOutput("<pre>"+JSON.stringify(s,null,2)+"</pre>");
    } catch (err) {
      return HtmlService.createHtmlOutput("<pre>ERR: "+String(err)+"</pre>");
    }
  }
  if (action === 'diag') {
    const a = diagDescribeIpsFolder_();
    const b = diagListIpsPdfs_();
    return HtmlService.createHtmlOutput(
      '<pre>DESCRIBE FOLDER:\n' + JSON.stringify(a,null,2) +
      '\n\nLIST IPS PDFs:\n' + JSON.stringify(b,null,2) + '</pre>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (action === 'rgpdstatus') {
    const ticket = (e && e.parameter && e.parameter.ticket) || '';
    try {
      const st = AuthCoreLib.getRgpdStatusFor(ticket, gatesCfg_());
      return HtmlService.createHtmlOutput('<pre>'+JSON.stringify(st,null,2)+'</pre>')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (err) {
      return HtmlService.createHtmlOutput('<pre>ERR '+(err && err.message)+'</pre>')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }


  if (action === "reset"){
    const next = canon + (DBG ? '?debug=1' : '');
    const html =
    '<!doctype html><html><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<style>body{font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif;padding:18px} .btn{display:inline-block;padding:10px 16px;background:#000;color:#fff;text-decoration:none;border-radius:6px;}</style>' +
    "</head><body>" +
    '<h3>Cookies e armazenamento local limpos.</h3>' +
    '<p><a class="btn" href="' + next.replace(/"/g, "&quot;") + '" target="_top">Ir para o Início</a></p>' +
    "<script>(function(){\n" +
    "  try{ localStorage.removeItem('sessTicket'); }catch(_){}\n" +
    "  try{ document.cookie = 'sessTicket=; Max-Age=0; Path=/; SameSite=Lax; Secure'; }catch(_){}\n" +
    "  var url = " + JSON.stringify(next) + ";\n" +
    "  try { if (window.top !== window.self) window.top.location.href = url; else location.replace(url); }\n" +
    "  catch(e){ console.warn('Auto-redirect bloqueado na sandbox, clique no botão.'); }\n" +
    "})();</script>" +
    "</body></html>";
    const out = HtmlService.createHtmlOutput(html);
    return out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // --- RGPD standalone: usa a página local + logs
  if (action === 'rgpd') {
    L("action rgpd");
    const ticket = (e && e.parameter && e.parameter.ticket) || '';
    const isEmbed = String(e && e.parameter && e.parameter.embed || '') === '1';
    try {
      L("route: rgpd — via biblioteca");
      return renderRgpdPage_(ticket, DBG, L.dump());
    } catch (err) {
      L("RGPD render FAIL: " + (err && err.message));
      return HtmlService.createHtmlOutput('<pre>RGPD render FAIL: '
        + String(err && err.message || err) + '</pre>')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }

  // --- pós-RGPD: página “ponte” com conteúdo/trace e fallback em navegação robusta + logs
  if (action === "postrgpd") {
    dbgLog("action postrgpd início! DBG=", DBG); //log em Execuções Apps Script
    L("L => action postrgpd");

    const ticket = (e && e.parameter && e.parameter.ticket) || '';
	  
    // CSV de linhas aceites vindo do RGPD.html
    const rowsQS = String((e && e.parameter && e.parameter.rows) || '').trim();
    const rows   = rowsQS ? rowsQS.split(',').map(s => parseInt(s, 10)).filter(n => Number.isFinite(n)) : [];

    /*
    try {
      const touched = (setRgpdRowsFor(ticket, rows) | 0); // <-- grava no servidor
      L(`RGPD save: rows=[${rows.join(',')}] touched=${touched}`);
      //console.log(`RGPD save: rows=[${rows.join(',')}] touched=${touched}`);
    } catch (err) {
      L('RGPD save FAIL: ' + (err && err.message));
    }
    */

    try { 
      AuthCoreLib.hostSaveRgpdRowsFor(ticket, rows, gatesCfg_());
      L(`RGPD save: rows=[${rows.join(',')}]`);
    } catch(err) { L('RGPD save FAIL: ' + err); }

    return renderMainPage_(ticket, DBG, L.dump());
  }


  if (action === "ipsview") {
    try {
      const ticket = e.parameter.ticket || "";
      const week   = String(e.parameter.week || "").trim();
      const autoDl = String(e.parameter.dl || "") === "1";
      const sess   = AuthCoreLib.requireSession(ticket);

      if (!week) {
        return HtmlService.createHtmlOutput("<p>Falta parâmetro <code>week</code>.</p>");
      }
      if (!isWeekOfEmail_(sess.email, week)) {
        return HtmlService.createHtmlOutput("<h3>Sem autorização para esta semana.</h3>");
      }

      const f = findLatestIpsFileForWeek_(week);
      if (!f) {
        return HtmlService.createHtmlOutput("<h3>Sem FT-IPS disponível para " + week + ".</h3>");
      }

      const blob = f.getBlob(); // PDF
      const b64  = Utilities.base64Encode(blob.getBytes());
      const name = String(f.getName()).replace(/"/g,"");
      const mime = blob.getContentType() || "application/pdf";

      // Página minimal que cria um Blob URL e usa-o no <object> e no botão Download
      const html =
        '<!doctype html><html><head><meta charset="utf-8">' +
        // CSP com suporte a blob:
        "<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'self' blob:; img-src 'self' data: blob:; media-src 'self' blob:; frame-src 'self' blob:;\">" +
        '<title>FT-IPS — ' + name + '</title>' +
        '<style>html,body{height:100%;margin:0} .wrap{height:100%;display:flex;flex-direction:column} .bar{padding:8px;border-bottom:1px solid #e5e7eb;font-family:system-ui,sans-serif} .view{flex:1} .btn{display:inline-block;padding:6px 10px;border:1px solid #999;border-radius:8px;background:#fff;text-decoration:none}</style>' +
        '</head><body><div class="wrap">' +
        '<div class="bar"><a id="dl" class="btn">Download (PDF)</a>' +
        '<span style="margin-left:8px;color:#666;font:12px system-ui,sans-serif">Semana ' + week.replace(/</g,"&lt;").replace(/>/g,"&gt;") + '</span></div>' +
        '<div class="view"><object id="viewer" type="application/pdf" width="100%" height="100%"></object></div>' +
        '</div>' +
        '<script>' +
        '(function(){' +
        'function b64ToBlobUrl(b64, mime){var bin=atob(b64),len=bin.length,buf=new Uint8Array(len);for(var i=0;i<len;i++)buf[i]=bin.charCodeAt(i);var url=URL.createObjectURL(new Blob([buf],{type:mime||"application/pdf"}));return url;}' +
        'var b64=' + JSON.stringify(b64) + ';' +
        'var mime=' + JSON.stringify(mime) + ';' +
        'var name=' + JSON.stringify(name) + ';' +
        'var url=b64ToBlobUrl(b64,mime);' +
        'var obj=document.getElementById("viewer"); if(obj) obj.data=url;' +
        'var a=document.getElementById("dl"); if(a){ a.href=url; a.download=name; }' +
        (autoDl ? 'setTimeout(function(){ try{ if(a) a.click(); }catch(_){} }, 80);' : '') +
        '})();' +
        '</script>' +
        '</body>' +
      '</html>';

      const out = HtmlService.createHtmlOutput(html);
      out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      return out;
    } catch (err) {
      const msg = (err && err.message) ? err.message : String(err);
      return HtmlService.createHtmlOutput("<pre>Erro IPS VIEW: " + msg + "</pre>");
    }
  }

  if (action === "login") {
    L("route: login");
    //optsProc.serverLog = ["login"];
    optsProc.serverLog = L.dump();
    optsProc.serverLog.unshift("login");

    return AuthCoreLib.renderLoginPage(optsProc); 
  }

  if (action === "logout") {
    L("route: logout");
    optsProc.serverLog = L.dump();
    optsProc.serverLog.unshift("logout");
    optsProc.wipe = true; // Diz à biblioteca para destruir o ticket
    return AuthCoreLib.renderLoginPage(optsProc); 
  }

  // 3) ticket → Main (ou gates)
  const ticket = (e && e.parameter && e.parameter.ticket) || "";
  if (ticket){
    L("have ticket, validating");
    let sess;
    try{
      sess = AuthCoreLib.requireSession(ticket);
    } catch(err){
      L("invalid ticket → login(wipe) ERR="+(err && err.message));
      //optsProc.serverLog = ["wipe"];
      optsProc.serverLog = L.dump();
      optsProc.serverLog.unshift("wipe");
      optsProc.wipe = true;
      return AuthCoreLib.renderLoginPage(optsProc);
    }
    L("session ok for", sess.email);
    L("go=", (e && e.parameter && e.parameter.go) || "(none)");

    // 👇 se pedido explícito, render RGPD já aqui (mesmo que já esteja aceite)
    if (String(e.parameter.go || '') === 'rgpd') {
      L("go=rgpd → render RGPD (via ticket-branch)");
      return renderRgpdPage_(ticket, DBG, L.dump());
    }

    // 👇 override: voltar sempre ao Main, ignorando gates
    if (String(e.parameter.go || '') === 'main') {
      L("go=main → render Main (skip gates)");
      return renderMainPage_(ticket, DBG, L.dump());
    }


    // Quando se vem do POST-RGPD (guardar RGPD), queremos avançar para Main
    // mesmo que RGPD ainda não esteja "Sim" (p.ex. guardou "Não").
    // Mantemos, no entanto, enforceGates (allowlist, saldo, etc), mas desligamos o gate RGPD.

    let gatesCfg = gatesCfg_();
    let st;
    try {
      // LOGA estado RGPD ANTES de decidir
      st = AuthCoreLib.getRgpdStatusFor(ticket, gatesCfg);
    } catch(err){
      L("getRgpdStatusFor ERR="+(err && err.message));
      optsProc.serverLog = L.dump();
      optsProc.serverLog.unshift("RGPDSTATUS");
      return AuthCoreLib.renderLoginPage(optsProc);
    }

    L(`RGPD status: total=${st.total} sim=${st.sim} nao=${st.nao} state=${st.state}`);
    if (st.state === 'pendente') {
      L("→ RGPD pendente → render RGPD (LIB)");
      return renderRgpdPage_(ticket, DBG, L.dump());
    }

    L('check enforceGates');
    let gate;
    try {
      // enforceGates pode voltar a barrar RGPD; loga se acontecer:
      gate = AuthCoreLib.enforceGates(sess.email, ticket, DBG, gatesCfg, L);
    } catch(err){
      L("enforceGates ERR="+(err && err.message));
      optsProc.serverLog = L.dump();
      optsProc.serverLog.unshift("enforceGates");
      return AuthCoreLib.renderLoginPage(optsProc);
    }
    L("check gate");
    if (gate){ L("gate returned HTML (provavelmente RGPD ou allowlist)"); return gate; }

    // OK → Main
    try {
      L("render Main");
      return renderMainPage_(ticket, DBG, L.dump());
    } catch (err) {
      L("render Main FAIL: " + (err && err.message));
      return HtmlService.createHtmlOutput(
        '<pre>render Main FAIL: ' + String(err && err.message || err) + '</pre>'
      ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }


  }

  // sem ticket → login
  L("no ticket → render login");
  //optsProc.serverLog = ["sem ticket"];
  optsProc.serverLog = L.dump();
  optsProc.serverLog.unshift("sem ticket");
  return AuthCoreLib.renderLoginPage(optsProc);

}

function diagDescribeIpsFolder_() {
  const tok = ScriptApp.getOAuthToken();
  const url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(IPS_FOLDER_ID)
            + '?fields=id,name,mimeType,driveId,owners(displayName,emailAddress),parents'
            + '&supportsAllDrives=true';
  const resp = UrlFetchApp.fetch(url, { headers:{ Authorization:'Bearer '+tok }, muteHttpExceptions:true });
  return { status: resp.getResponseCode(), body: resp.getContentText() };
}

function diagListIpsPdfs_() {
  const tok = ScriptApp.getOAuthToken();
  const q = "'" + IPS_FOLDER_ID + "' in parents and mimeType='application/pdf' and trashed=false and name contains 'FT-IPS'";
  const url = 'https://www.googleapis.com/drive/v3/files'
            + '?q=' + encodeURIComponent(q)
            + '&fields=files(id,name,mimeType,modifiedTime,driveId,parents)'
            + '&pageSize=10'
            + '&supportsAllDrives=true'
            + '&includeItemsFromAllDrives=true';
  const resp = UrlFetchApp.fetch(url, { headers:{ Authorization:'Bearer '+tok }, muteHttpExceptions:true });
  return { status: resp.getResponseCode(), body: resp.getContentText() };
}

function diagFetchFileBytes_(fileId) {
  const tok = ScriptApp.getOAuthToken();
  const url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) + '?alt=media'
            + '&supportsAllDrives=true';
  const resp = UrlFetchApp.fetch(url, { headers:{ Authorization:'Bearer '+tok }, muteHttpExceptions:true });
  const st = resp.getResponseCode();
  if (st >= 200 && st < 300) {
    // atenção: para debug, não devolvas os bytes todos — mostra só o tamanho
    return { status: st, size: resp.getContent().length };
  }
  return { status: st, body: resp.getContentText() };
}



// ===== Debug utils =====
function debugWho(ticket){ return AuthCoreLib.requireSession(ticket); }
function debugAuthConfig(){
  const cfg = authCfg_();
  return {
    CLIENT_ID_present: !!cfg.clientId,
    CLIENT_SECRET_present: !!cfg.clientSecret,
    sampleAuthUrl: AuthCoreLib.buildAuthUrlFor("diag-"+Date.now(), false, false, cfg),
  };
}
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

