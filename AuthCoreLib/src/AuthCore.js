
// =========================
// File: AuthCore.gs - em AuthCoreLib (biblioteca)
//
// Script ID: 1cyjgfEtpjB7muIaSwmoLr_BD_BbgeEzHPA8M8Upni_LJkzT2hK6V1Tuh
// URL: https://script.google.com/macros/library/d/1cyjgfEtpjB7muIaSwmoLr_BD_BbgeEzHPA8M8Upni_LJkzT2hK6V1Tuh/146
// =========================


// API: Funções públicas sem underscore e com nomes limpos:
function renderLoginPage(opts) { return renderLoginPage_(opts); }
function buildAuthUrlFor(nonce, dbg, embed, cfg, clientUrl) { return buildAuthUrlFor_(nonce, dbg, embed, cfg, clientUrl); }
function beginAuth(e, cfg) { return beginAuth_(e, cfg); }
function finishAuth(e, cfg) { return finishAuth_(e, cfg); }
function isTicketValid(ticket, dbg) { return !!validateSessionToken_(ticket); }
function pollTicket(nonce) { return takeTicketForNonce_(nonce) || ""; }
function requireSession(ticket) {
  const p = getSession(ticket);
  if (!p) throw new Error("Sessão inválida/expirada");
  return JSON.parse(p);
}
function enforceGates(email, ticket, DBG, gatesCfg, extLogger) {
  return enforceGates_(email, ticket, DBG, gatesCfg, extLogger) || "";
}
function renderRgpdPage(opts) { return renderRgpdPage_(opts); }

function getProfileStats(ticket, cfg) { return getProfileStats_(ticket, cfg); }
function hostListRgpdRowsFor(ticket, cfg) { return hostListRgpdRowsFor_(ticket, cfg); }
function hostSaveRgpdRowsFor(ticket, acceptedRows, cfg) { return hostSaveRgpdRowsFor_(ticket, acceptedRows, cfg); }

// NOVO: Função para uso das web apps clientes para registar bloqueios personalizados
function logFailedAccess(ticket, reason, cfg) { return logFailedAccessPublic_(ticket, reason, cfg); }
function libBuild(){ return "AuthCoreLib build 2026-03-06 23:48 - development mode"; }


// ===== Config / Constantes comuns =====
const OAUTH_SCOPES = "openid email profile";
const STATE_MAX_AGE_MS = 15 * 60 * 1000; // 15 min
const NONCE_TTL_SEC = 180; // 3 min para nonce→ticket


// Defaults internos da biblioteca (só para a Associação Portobelo)
var LIB_SS_TITULARES_ID = "1YE16kNuiOjb1lf4pbQBIgDCPWlEkmlf5_-DDEZ1US3g";
var LIB_RANGES = { titulares: { name:"tblTitulares", sheet:"Titulares", a1:"A6:V" } };
var LIB_COLS   = { email:"e-mail", rgpd:"RGPD", pago:"€", saldo:"Saldo", semanas:"Semanas" };
var LIB_NOTIFY = { to:"secretario-direcao@titulares-portobelo.pt", ccAllRows:true };

function __defCfg(cfgParam){
  //if (g && g.ssTitularesId) return g; // host forneceu cfg
  //return { ssTitularesId: LIB_SS_TITULARES_ID, ranges: LIB_RANGES, cols: LIB_COLS, notify: LIB_NOTIFY };
  const out = cfgParam || {};
  out.ssTitularesId = out.ssTitularesId || LIB_SS_TITULARES_ID;
  out.ranges = out.ranges || LIB_RANGES;
  out.cols = out.cols || LIB_COLS;
  out.notify = out.notify || {};
  out.notify.to = out.notify.to || LIB_NOTIFY.to;
  return out;  
}


// ---------- Config resolver ----------
function resolveCfg_(cfg) {
  const sp = PropertiesService.getScriptProperties();
  // usa o que vier do host (cfg) ou, em último caso, as props da própria biblioteca
  const clientId = (cfg && cfg.clientId) || sp.getProperty("CLIENT_ID");
  const clientSecret = (cfg && cfg.clientSecret) || sp.getProperty("CLIENT_SECRET");
  const ruProp = (cfg && cfg.redirectUri) || sp.getProperty("REDIRECT_URI");
  const redirectUri = (ruProp && ruProp.trim()) || canonicalAppUrl_();
  return { clientId, clientSecret, redirectUri };
}

// ===== (Opcional) Debug helpers reutilizáveis =====
function isDebug_(e){
  const p = e && e.parameter ? e.parameter : {};
  let DBG = false;
  if (p && Object.prototype.hasOwnProperty.call(p, "debug")) DBG = true;
  return DBG;
}

function makeLogger_(DBG) {
  const start = new Date();
  const log = [];
  function L() {
    const t = new Date(new Date() - start).toISOString().substr(11, 8); // hh:mm:ss desde início
    log.push(t + " " + Array.prototype.map.call(arguments, String).join(" "));
  }
  L.dump = () => log.slice();
  return L;
}


// ===== Script properties helpers =====

function getScriptProp_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function setScriptProp_(k, v) { return PropertiesService.getScriptProperties().setProperty(k, v); }

// ===== URL canónico do deployment =====
function canonicalAppUrl_() {
  var url = ScriptApp.getService().getUrl();
  // Mantém o URL canónico real do deployment.
  // Em contas Workspace, isto inclui /a/<domínio>/ e deve ser preservado.
  //return url.replace(/\/a\/[^/]+\/macros/, "/macros"); // força .../macros/s/ID/exec
  return url;
}

// Helper para incluir ficheiros HTML (templates parciais)
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

// Preferir REDIRECT_URI definido; fallback para canónico
//function redirectUri_() {
  // Preferir a property se estiver correcta; senão, cair para a canónica deste deployment
//  return (REDIRECT_URI && REDIRECT_URI.trim()) || canonicalAppUrl_();
//}

// ===== OAuth URL =====
function toQueryString_(obj) {
  var parts = [];
  Object.keys(obj || {}).forEach(function (k) {
    const v = obj[k];
    if (v === undefined || v === null) return;
    parts.push(encodeURIComponent(k) + "=" + encodeURIComponent(String(v)));
  });
  return parts.join("&");
}

// ===== Ticket “stateless” com HMAC =====
function getSessSecretBytes_() {
  const p = PropertiesService.getScriptProperties();
  let b64 = p.getProperty("SESSION_HMAC_SECRET_B64");
  if (b64) b64 = b64.trim();
  if (!b64) {
    // Deriva 32 bytes “random enough” a partir de dois UUIDs (evita depender de RNG externo)
    const seed = Utilities.getUuid() + ":" + Utilities.getUuid() + ":" + Date.now();
    const bytes = Utilities.computeHmacSha256Signature(seed, seed); // 32 bytes
    b64 = Utilities.base64EncodeWebSafe(bytes);
    p.setProperty("SESSION_HMAC_SECRET_B64", b64);
    // opcional: apaga o legado se existir
    p.deleteProperty("SESSION_HMAC_SECRET");
  }
  return Utilities.base64DecodeWebSafe(b64);
}

function issueSessionToken_(email, days, extra){
  const exp = Date.now() + (days || 14) * 24 * 60 * 60 * 1000;
  const payloadObj = Object.assign(
    {
      email: email,
      exp: exp,
      v: 2,
      iat: Date.now(),
    },
    extra || {}
  );

  const payload = JSON.stringify(payloadObj);
  const pBytes = Utilities.newBlob(payload).getBytes();
  const pB64 = Utilities.base64EncodeWebSafe(pBytes);
  const pB64bytes = Utilities.newBlob(pB64).getBytes();
  const sigBytes = Utilities.computeHmacSha256Signature(pB64bytes, getSessSecretBytes_());
  const sig = Utilities.base64EncodeWebSafe(sigBytes);
  return pB64 + "." + sig;
}


function validateSessionToken_(tok) {
  if (!tok || tok.indexOf(".") < 0) throw new Error("Formato de token inválido");
  const parts = tok.split("."), pB64 = parts[0], sig = parts[1];

  const pB64bytes = Utilities.newBlob(pB64).getBytes(); // bytes da string base64
  const expSig = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(pB64bytes, getSessSecretBytes_()), // bytes,bytes
  );
  if (expSig !== sig) throw new Error("Assinatura HMAC inválida (As propriedades de script das apps são diferentes?)");

  const data = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(pB64)).getDataAsString());
  if (!data) throw new Error("Token vazio");
  if (!data.email) throw new Error("Token sem email");
  if (!data.exp) throw new Error("Token sem data de expiração");
  if (Date.now() > data.exp) throw new Error("Sessão expirada");
  return data;
}

// ---------- Helpers de sessão ----------
function getSession(ticket) {
  const d = validateSessionToken_(ticket);
  return d ? JSON.stringify(d) : null;
}

// ===== State / nonce (HMAC) =====
function getStateSecret_() {
  let s = getScriptProp_("STATE_SECRET");
  if (!s) {
    const rand = Utilities.getUuid();
    s = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(rand, rand));
    setScriptProp_("STATE_SECRET", s);
  }
  return s;
}

// Gera state assinado: base64url(payload) + '.' + base64url( HMAC_SHA256(payload, secret) )
// aceitar flags no state
function createStateToken_(dbg, embed, nonceOpt, redirectUri) {
  const payload = { ts: Date.now(), nonce: nonceOpt || Utilities.getUuid(), dbg: !!dbg, embed: !!embed, ru: redirectUri || "" };
  const payloadBytes = Utilities.newBlob(JSON.stringify(payload), "application/json").getBytes();
  const sigBytes = Utilities.computeHmacSha256Signature(payloadBytes, Utilities.base64DecodeWebSafe(getStateSecret_()));
  return Utilities.base64EncodeWebSafe(payloadBytes) + "." + Utilities.base64EncodeWebSafe(sigBytes);
}

function parseStateToken_(state) {
  if (!state || state.indexOf(".") < 1) return { ok: false };
  const [payloadB64, sigB64] = state.split(".", 2);
  try {
    const payloadBytes = Utilities.base64DecodeWebSafe(payloadB64);
    const secretBytes = Utilities.base64DecodeWebSafe(getStateSecret_());
    const expSigBytes = Utilities.computeHmacSha256Signature(payloadBytes, secretBytes);
    const gotSigBytes = Utilities.base64DecodeWebSafe(sigB64);
    if (expSigBytes.length !== gotSigBytes.length) return { ok: false };
    let diff = 0;
    for (let i = 0; i < expSigBytes.length; i++)
      diff |= expSigBytes[i] ^ gotSigBytes[i];
    if (diff !== 0) return { ok: false };
    const payload = JSON.parse(Utilities.newBlob(payloadBytes).getDataAsString());
    if (typeof payload.ts !== "number" || Date.now() - payload.ts > STATE_MAX_AGE_MS) return { ok: false };
    return { ok: true, payload: payload };
  } catch (e) { return { ok: false }; }
}

// nonce → ticket cache (para polling)
function putTicketForNonce_(nonce, ticket) { if (nonce && ticket) CacheService.getScriptCache().put("nonce:" + nonce, ticket, NONCE_TTL_SEC); }

function takeTicketForNonce_(nonce) {
  if (!nonce) return "";
  const cache = CacheService.getScriptCache();
  const k = "nonce:" + nonce;
  const t = cache.get(k) || "";
  if (t) cache.remove(k);
  return t;
}

// ===== OAuth client URL =====
/*
function getClientId_() {
  return getScriptProp_("CLIENT_ID");
}

function getClientSecret_() {
  return getScriptProp_("CLIENT_SECRET");
}
*/

// Nota: tornamos a construção do URL "interna" e expomos via buildAuthUrlFor()
function buildAuthUrlFor_(nonce, dbg, embed, cfg) {
  const L = makeLogger_(dbg);
  L('Entrada em buildAuthUrlFor_');

  const { clientId, redirectUri: fallbackRu } = resolveCfg_(cfg);
  if (!clientId)
    throw new Error("CLIENT_ID ausente nas Script Properties do projeto host.");
  L('clientId =' + clientId);

  // A MAGIA: Usa o URL do browser se fornecido, senão usa o padrão
  const finalRu = clientUrl || fallbackRu; 
  if (!finalRu) throw new Error("REDIRECT_URI ausente (e URL canónico indisponível).");
  L('redirectUri =' + redirectUri);

  const params = {
    client_id: clientId,
    redirect_uri: redirectUri,
    response_type: 'code',
    scope: OAUTH_SCOPES,
    access_type: 'offline',
    prompt: 'select_account consent',
    include_granted_scopes: 'true',
    state: createStateToken_(!!dbg, !!embed, nonce,finalRu),
  };

  return "https://accounts.google.com/o/oauth2/v2/auth?" + toQueryString_(params);
}

// ===== Render do Login (comum às apps) =====
function renderLoginPage_(opts) {

  const L = makeLogger_(opts.debug);
  //const L = makeLogger_(1); //parece que opts.debug não está a funcionar?
  L('function renderLoginPage_');
  const t = HtmlService.createTemplateFromFile("Login");
  t.CANON_URL = canonicalAppUrl_();
  //t.CLIENT_ID = getClientId_(); // opcional; o HTML atual nem usa
  t.AUTOSTART = "1"; // auto-inicia o popup
  t.DEBUG = opts.debug ? "1" : "";
  t.SERVER_LOG = 
      opts.serverLog && opts.serverLog.join
        ? opts.serverLog.join("\n")
        : String(opts.serverLog || "");
  // 🔒 Defesas contra ReferenceError no template
  //if (typeof t.SERVER_LOG === "undefined" || t.SERVER_LOG == null) t.SERVER_LOG = "";
  t.WIPE = opts.wipe ? "1" : "";  
  L('renderLoginPage_: opts.wipe = ' + opts.wipe + ', t.WIPE = ' + t.WIPE);
  t.ticket = "";   // ← mesmo que não uses, evita o erro quando existir <?= ticket ?>
  t.TICKET = "";   // ← alias, caso o HTML use TICKET
  t.SERVER_VARS = "";
  t.PAGE_TAG = 'LOGIN';
  try {
    return t.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    const msg =
      "Login.html evaluate() falhou\n" +
      String(err) +
      "\n--- SERVER LOG ---\n" +
      (opts.serverLog && opts.serverLog.join
        ? opts.serverLog.join("\n")
        : String(opts.serverLog || ""));
    return HtmlService.createHtmlOutput('<pre style="white-space:pre-wrap">' + msg + "</pre>",
    );
  }
}

// ===== Fluxo OAuth: begin + finish =====
function beginAuth_(e, cfg) {
  const nonce = (e && e.parameter && e.parameter.nonce) || Utilities.getUuid();
  const dbg = isDebug_(e);
  const embed = (e && e.parameter && e.parameter.embed === "1") || false;
  const url = buildAuthUrlFor_(nonce, dbg, embed, cfg);
  return HtmlService.createHtmlOutput(
    '<meta charset="utf-8"><script>location.replace(' +
      JSON.stringify(url) +
      ");</script>",
  );
}

function finishAuth_(e, cfg) {
  if (e && e.parameter && e.parameter.error) {
    return HtmlService.createHtmlOutput(
      "<pre>Erro de autenticação: " + String(e.parameter.error) + "</pre>",
    );
  }
  var state = (e && e.parameter && e.parameter.state) || "";
  var code = (e && e.parameter && e.parameter.code) || "";
  var parsed = parseStateToken_(state);
  if (!parsed.ok) return HtmlService.createHtmlOutput("<pre>State inválido.</pre>");
  var nonce = parsed.payload.nonce || "";
  //var dbg = !!parsed.payload.dbg;

  // usa cfg do host
  const { clientId, clientSecret, redirectUri:fallbackRu } = resolveCfg_(cfg);
  
  // A MAGIA PARTE 2: Recupera o URL exato (/dev ou /exec) de dentro do State encriptado
  const finalRu = (parsed.payload && parsed.payload.ru) ? parsed.payload.ru : fallbackRu;
  
  /*
  if (!clientId || !clientSecret) {
    return HtmlService.createHtmlOutput(
      "<pre>CLIENT_ID/CLIENT_SECRET ausentes no projeto host.</pre>",
    );
  }
  */
  // troca code→token
  var resp = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method: "post",
    payload: {
      code: code,
      client_id: clientId,
      client_secret: clientSecret,
      redirect_uri: finalRu,
      grant_type: "authorization_code",
    },
    muteHttpExceptions: true,
  });
  
  var status = resp.getResponseCode(),
    body = resp.getContentText();
  //if (status < 200 || status >= 300) {
  if (status >= 300) {
    return HtmlService.createHtmlOutput(
      "<pre>Falha a trocar o código por token (" +
        status +
        "):\n" +
        body +
        "</pre>",
    );
  }

  var tok = {};
  try { tok = JSON.parse(body);} catch (_) {}
  var idt = tok.id_token;
  if (!idt) return HtmlService.createHtmlOutput("<pre>Sem id_token devolvido pelo Google.</pre>");

  // valida id_token (aud/iss)
  var email = "", name = "", picture = "";
  try {
    var parts = idt.split(".");
    var payloadJson = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[1])).getDataAsString();
    var claims = JSON.parse(payloadJson);
    var audOk = claims.aud === clientId;
    var issOk =
      claims.iss === "https://accounts.google.com" ||
      claims.iss === "accounts.google.com";
    //if (!audOk || !issOk) throw new Error("iss/aud inválidos");
	  if (!audOk) throw new Error("aud inválido");
    email = String(claims.email || '');
    name = claims.name || '';
    picture = claims.picture || '';
  } catch (err) {
    return HtmlService.createHtmlOutput(
      "<pre>id_token inválido: " + String(err) + "</pre>",
    );
  }

  if (!email) return HtmlService.createHtmlOutput("<pre>id_token sem email.</pre>");

  var ticket = issueSessionToken_(email, 14, { name: name, picture: picture });
  if (nonce) putTicketForNonce_(nonce, ticket);

  //var canon = canonicalAppUrl_(); 
  //var canon = redirectUri; // mantém EXACTAMENTE o mesmo /a/<domínio>/... usado no OAuth
  var html = `
<meta charset="utf-8"><title>Autenticado</title>
<style>body{font-family:system-ui,sans-serif;padding:1rem}</style>
<div>A redirecionar…</div><script>
(function(){
  var t = ${JSON.stringify(ticket)};

  // 0) guardar ticket
  try{ localStorage.setItem('sessTicket', t); }catch(_){}
  try{ document.cookie = 'sessTicket='+encodeURIComponent(t)+'; Path=/; SameSite=Lax; Secure'; }catch(_){}

  // 1) postMessage com o ticket
  try{ if (window.opener) window.opener.postMessage({type:'portobelo_ticket', ticket:t}, '*'); }catch(_){}

  // 2) define cookie (mesma origem script.google.com)
  try{ document.cookie = 'sessTicket='+encodeURIComponent(t)+'; Path=/; SameSite=Lax; Secure'; }catch(_){}

  // 3) fecha o popup – quem navega é a janela principal via goWithTicket()
  setTimeout(function(){ try{ window.close(); }catch(_){} }, 200);
})();
</script>`;

  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// =========================
// AuthCoreLib — Gates (membership + RGPD)
// =========================

// Espera um objeto cfg vindo do host:
// {
//   ssTitularesId: "…",
//   ranges: { titulares:{name|sheet|a1}, quotas:{name|sheet|a1} },
//   cols: { email:"e-mail", pago:"€", rgpd:"RGPD" },
//   mailTo: "geral@titulares-portobelo.pt"
// }

function logFailedAccessToSheet_(email, name, saldo, semanas, reason, cfg) {
  try {
    const ss = SpreadsheetApp.openById(cfg.ssTitularesId);
    let sheet = ss.getSheetByName("Acessos");
    if (!sheet) {
      sheet = ss.insertSheet("Acessos");
      sheet.appendRow(["Data", "Nome", "Email", "€", "Semanas", "Motivo"]);
    }
    sheet.appendRow([new Date(), name || "", email || "", saldo !== undefined ? saldo : "", semanas || "", reason || ""]);
  } catch(e) { console.log("Erro a gravar acessos:", e); }
}

function extractFinancialInfo_(info, cfg) {
  let saldo = "", semanas = "";
  if (info.matches.length > 0) {
    const iSaldo = info.ix[cfg.cols.saldo || "Saldo"];
    const iSem = info.ix[cfg.cols.semanas || "Semanas"];
    let tot = 0; let semArr = [];
    for(const r of info.matches) {
      if (iSaldo != null) tot += parsePtNumber_AuthCore_(info.values[r][iSaldo]);
      if (iSem != null && info.values[r][iSem]) semArr.push(info.values[r][iSem]);
    }
    saldo = tot; semanas = semArr.join(" · ");
  }
  return { saldo, semanas };
}

function logFailedAccessPublic_(ticket, reason, cfg) {
  cfg = __defCfg(cfg);
  try {
    const sess = requireSession(ticket);
    const info = getTitularesRowsByEmail_AuthCore_(sess.email, cfg);
    const fin = extractFinancialInfo_(info, cfg);
    logFailedAccessToSheet_(sess.email, sess.name, fin.saldo, fin.semanas, reason, cfg);
  } catch(e) {
    logFailedAccessToSheet_("Desconhecido", "Ticket Inválido", "", "", reason, cfg);
  }
}
function enforceGates_(email, ticket, DBG, gatesCfg, extLogger){
  const L = extLogger || makeLogger_(DBG);
  L('function enforceGates_');

  gatesCfg = __defCfg(gatesCfg);
  let sess = { name: "", email: email };
  try { sess = requireSession(ticket); } catch(e){}
  // 1) allowlist
  if (!isAllowedEmail_AuthCore_(email)) {
    L('enforceGates_: email not registered');
	logFailedAccessToSheet_(email, sess.name, "", "", "Não registado na allowlist", gatesCfg);
    return renderNotAllowed_AuthCore_(email, DBG);
  }

  // 2) tem pelo menos uma linha e ≥1€ pago?
  const info = getTitularesRowsByEmail_AuthCore_(email, gatesCfg);
  if (!(info.matches.length && hasAnyRowMinPayment_AuthCore_(info, gatesCfg, 1))) {
    L("enforceGates_: no payments - lines = ",  info.matches.length);
    const fin = extractFinancialInfo_(info, gatesCfg);
    logFailedAccessToSheet_(email, sess.name, fin.saldo, fin.semanas, "S/ linhas ou quotas atraso", gatesCfg);
    return HtmlService.createHtmlOutput('<meta charset="utf-8"><h3>Acesso negado</h3><p>Email não registado ou quotas em atraso.</p>').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 3) RGPD (nenhuma linha pendente de aceitação?)
  const s = getRgpdStatusFor(ticket, gatesCfg);
  var optsProc = { ticket: ticket, debug: DBG, serverLog: L.dump(), wipe: false };
  optsProc.serverLog.unshift("enforceGates");
  
  if (s.state === 'pendente') {
    L("enforceGates_: pelo menos uma das ",  s.total, " linhs tem RGPD vazio");
    //return renderRgpdPage_(DBG, ticket, gatesCfg && gatesCfg.canon);
    //optsProc.ticket = ticket;
    //optsProc.serverLog = ["enforceGates"];
    //optsProc.serverLog = L.dump();
    //optsProc.serverLog.unshift("enforceGates");    
    return renderRgpdPage_(optsProc);
  }

  // OK → deixa o host seguir para o Main
  L("enforceGates_ tudo OK: ", s.total, " linhas, com RGPD aceite para ", s.sim, " e recusado para ",  s.nao);
  return null;
}

/*
function notifyRgpdDecision_AuthCore_(info, sess, accept, cfg){
  const to = (cfg.notify && cfg.notify.to) || Session.getActiveUser().getEmail();
  const subj = `RGPD: ${accept ? 'ACEITE' : 'REJEITADO'} — ${sess.name||''} <${sess.email}>`;

  const iNome    = info.ix['Nome membros (e titulares representados)'];
  const iSemanas = info.ix['Semanas'];

  const linhas = info.matches.map(r=>{
    const nome = (iNome!=null ? info.values[r][iNome] : '') || '';
    const sem  = (iSemanas!=null ? info.values[r][iSemanas] : '') || '';
    const rgpd = info.values[r][info.iRGPD] || '';
    return `- linha ${r+1}: ${nome} | Semanas: ${sem} | RGPD=${rgpd}`;
  }).join('\n');

  const body = [
    `Utilizador: ${sess.name||''} <${sess.email}>`,
    `Decisão: ${accept ? 'ACEITOU' : 'REJEITOU'}`,
    `Linhas afetadas: ${info.matches.length}`,
    ``,
    linhas
  ].join('\n');

  MailApp.sendEmail({ to, subject: subj, body, name: 'AT Portobelo' });
}
*/

/**
 * Chamada a partir do RGPD.html da biblioteca.
 * `gatesCfg` vem do host e deve conter pelo menos:
 *   { titularesId: "...", rangeNameOrA1: "...", notify: { to: "geral@...", ccAllRows: true } }
 */
/*
function acceptRgpdForMe(ticket, decision, gatesCfg){
  gatesCfg = __defCfg(gatesCfg);

  const sess   = requireSession(ticket);
  const accept = (decision === 'accept');

  // 1) ler uma vez e manter em memória
  const rowsInfo = getTitularesRowsByEmail_AuthCore_(sess.email, gatesCfg);

  // 2) escrever apenas se muda (Sim/Não) nas linhas desse email
  var changed = setRgpdForEmail_AuthCore_(rowsInfo, gatesCfg, accept ? 'Sim' : 'Não');

  // 3) notificar só quando houve alteração
  if (changed > 0) {
    try { notifyRgpdDecision_AuthCore_(rowsInfo, sess, accept, gatesCfg); } catch(_) {}
  }
  return { ok:true, changed: changed|0 };
}


// ---------- Helpers internos (AuthCoreLib) ----------

function isAllowedEmail_AuthCore_(email){
  const csv = (PropertiesService.getScriptProperties().getProperty('ALLOWLIST_CSV') || '');
  const list = csv.split(/[,\s;]+/).map(s=>s.trim().toLowerCase()).filter(Boolean);
  if (!list.length) return true; // allowlist desligada
  return list.includes(String(email||'').toLowerCase());
}

function renderNotAllowed_AuthCore_(email, DBG){
  const sp = PropertiesService.getScriptProperties();
  //const canon = ScriptApp.getService().getUrl().replace(/\/a\/[^/]+\/macros/, '/macros');
  const canon = ScriptApp.getService().getUrl();
  const dbgBlock = DBG ? (function(){
    const raw = sp.getProperty('ALLOWLIST_CSV') || '(vazio)';
    const parsed = raw.split(/[,\s;]+/).map(s=>s.trim()).filter(Boolean).join(', ');
    return `<details style="margin-top:12px"><summary>DEBUG</summary>
<pre>Email sessão: ${email}
ALLOWLIST_CSV (raw): ${raw}
ALLOWLIST parsed: ${parsed}</pre></details>`;
  })() : '';
  return HtmlService.createHtmlOutput(
    '<meta charset="utf-8"><h3>Acesso não autorizado</h3>'+
    '<p>Este endereço não está na lista da fase de validação.</p>'+
    `<p><a href="${canon}?action=login${DBG?'&debug=1':''}">Voltar</a></p>`+dbgBlock
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
*/

// ... Helpers de RGPD e Fetching

function fetchTable_AuthCore_(ssId, cfg){
  const ss = SpreadsheetApp.openById(ssId);
  let range = null;
  if (cfg && cfg.name) { try{ range = ss.getRangeByName(cfg.name); }catch(_){ } }
  if (!range) range = ss.getSheetByName(cfg.sheet).getRange(cfg.a1);
  const values = range.getDisplayValues();
  if (!values.length) return { header:[], rows:[] };
  return { header: values[0], rows: values.slice(1).filter(r => r.some(v => String(v).trim()!=='') ) };
}

function isAllowedEmail_AuthCore_(email){
  const csv = (PropertiesService.getScriptProperties().getProperty('ALLOWLIST_CSV') || '');
  const list = csv.split(/[,\s;]+/).map(s=>s.trim().toLowerCase()).filter(Boolean);
  if (!list.length) return true; 
  return list.includes(String(email||'').toLowerCase());
}

function renderNotAllowed_AuthCore_(email, DBG){
  const canon = ScriptApp.getService().getUrl();
  return HtmlService.createHtmlOutput('<meta charset="utf-8"><h3>Acesso não autorizado</h3><p>Este endereço não está na lista da fase de validação.</p><p><a href="'+canon+'?action=login'+(DBG?'&debug=1':'')+'">Voltar</a></p>').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/*
function indexByHeader_AuthCore_(header){
  const map={}; header.forEach((h,i)=> map[String(h).trim()]=i); return map;
}
function cellHasEmail_AuthCore_(cell, emailLC){
  return String(cell||'').split(/[;,]/).map(s=>s.trim().toLowerCase()).filter(Boolean).includes(emailLC);
}
*/



function parsePtNumber_AuthCore_(s){
  const clean = String(s||'').replace(/\s/g,'').replace(/\./g,'').replace(',','.').replace(/[^\d.\-]/g,'');
  const n = parseFloat(clean); return isNaN(n)?0:n;
}

function getTitularesRowsByEmail_AuthCore_(email, cfg){
  cfg = __defCfg(cfg);

  const ss = SpreadsheetApp.openById(cfg.ssTitularesId);
  let range=null; try{ if (cfg.ranges && cfg.ranges.titulares && cfg.ranges.titulares.name) range=ss.getRangeByName(cfg.ranges.titulares.name); }catch(_){}
  if (!range) range = ss.getSheetByName(cfg.ranges.titulares.sheet).getRange(cfg.ranges.titulares.a1);

  const values = range.getValues();
  const header = values.length ? values[0].map(v=>String(v).trim()) : [];
  const ix = {}; header.forEach((h,i)=> ix[h]=i);

  const iEmail = ix[cfg.cols.email];
  const iRGPD  = ix[cfg.cols.rgpd];
  if (iEmail==null || iRGPD==null) throw new Error('Config/colunas RGPD ou e-mail em falta');

  const emailLC = String(email||'').trim().toLowerCase();
  const matches = [];
  for (let r=1; r<values.length; r++){
    const cells = String(values[r][iEmail]||'')
      .split(/[;,]/).map(s=>s.trim().toLowerCase()).filter(Boolean);
    if (cells.includes(emailLC)) matches.push(r);
  }
  return { range, values, header, ix, iEmail, iRGPD, matches, emailLC };
}


function hasAnyRowMinPayment_AuthCore_(info, gatesCfg, minEUR){
  gatesCfg = __defCfg(gatesCfg);
  const iPago = info.ix[gatesCfg.cols.pago];     // ex.: "€"
  if (iPago == null) return false;
  const lim = typeof minEUR === 'number' ? minEUR : 1;

  for (const r of info.matches){
    const val = info.values[r][iPago];
    if (parsePtNumber_AuthCore_(val) >= lim) return true;
  }
  return false;
}

function rgpdStats_(email, cfg){
  const info = getTitularesRowsByEmail_AuthCore_(email, cfg);
  let total=0, sim=0, nao=0;
  for (const r of info.matches) {
    total++;
    const v = String(info.values[r][info.iRGPD]||'').trim().toLowerCase();
    if (v==='sim') sim++; if (v==='não') nao++;
  }
  return { total, sim, nao };
}

// Lê estatuto RGPD do email da sessão
// state:
//  - pendente - não há linhas, ou há alguma em que RGPD não esteja definido (Sim ou Não)
//  - parcial - todas as linhas têm o RGPD definido, mas algumas têm Não 
//  - total - todas as linhas têm o RGPD definido a Sim 
function getRgpdStatusFor(ticket, cfg){
  const sess = requireSession(ticket);
  const st = rgpdStats_(sess.email, cfg); // { total, sim, nao }
  var state = 'pendente';
  if (st.total>0 && st.sim===st.total) state = 'total';
  else if (st.total === (st.sim + st.nao)) state = 'parcial';
  return { email: sess.email, total: st.total, sim: st.sim, nao: st.nao, state };
}

/*
// true se TODAS as linhas estão "Sim"
function isRgpdAllAcceptedFor(ticket, cfg){
  const s = getRgpdStatusFor(ticket, cfg);
  return s.total>0 && s.sim===s.total;
}
*/

// usado pelo host em action=rgpd
//function renderRgpdPage_(DBG, ticket, canonOverride) {
function renderRgpdPage_(opts) {
  const L = makeLogger_(opts.debug);
  L('renderRgpdPage_');
  const t = HtmlService.createTemplateFromFile('RGPD');
  //t.CANON_URL  = canonOverride || canonicalAppUrl_(); // << usa o do HOST se vier
  t.CANON_URL  = canonicalAppUrl_(); // << usa o do HOST se vier
  t.ticket = opts.ticket || '';
  t.TICKET = opts.ticket || '';
  t.DEBUG  = opts.debug ? '1' : '';
  t.WIPE = opts.wipe ? "1" : "";
  t.SERVER_LOG = 
      opts.serverLog && opts.serverLog.join
        ? opts.serverLog.join("\n")
        : String(opts.serverLog || "");
  //if (typeof t.SERVER_LOG === "undefined" || t.SERVER_LOG == null) t.SERVER_LOG = "";
  t.SERVER_VARS = "";
  t.PAGE_TAG = 'RGPD';
  try {
    return t.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    const msg =
      "RGPD.html evaluate() falhou\n" +
      String(err) +
      "\n--- SERVER LOG ---\n" +
      (opts.serverLog && opts.serverLog.join
        ? opts.serverLog.join("\n")
        : String(opts.serverLog || ""));
    return HtmlService.createHtmlOutput(
      '<pre style="white-space:pre-wrap">' + msg + "</pre>",
    );
  }
}

function acceptRgpdForMe(ticket, decision, gatesCfg){
  gatesCfg = __defCfg(gatesCfg);
  const sess = requireSession(ticket), accept = (decision === 'accept');
  const rowsInfo = getTitularesRowsByEmail_AuthCore_(sess.email, gatesCfg);
  var changed = setRgpdForEmail_AuthCore_(rowsInfo, gatesCfg, accept ? 'Sim' : 'Não');
  return { ok:true, changed: changed|0 };
}

/*
function isRgpdAcceptedForEmail_AuthCore_(rowsInfo, gatesCfg){
  gatesCfg = __defCfg(gatesCfg);

  const iRGPD = rowsInfo.iRgpd;
  if (iRGPD == null) return true; // coluna ainda não criada → não bloquear
  if (!rowsInfo.rows.length) return false;
  return rowsInfo.rows.some(r => String(r[iRGPD]||'').trim().toLowerCase().startsWith('s'));
}
*/



// --- helpers privados na biblioteca ---

// helper interno (igual em espírito ao do host)
function rgpdStats_(email, cfg){
  const ss = SpreadsheetApp.openById(cfg.ssTitularesId);
  let range=null; try{ if (cfg.ranges && cfg.ranges.titulares && cfg.ranges.titulares.name) range = ss.getRangeByName(cfg.ranges.titulares.name); }catch(_){}
  if (!range) range = ss.getSheetByName(cfg.ranges.titulares.sheet).getRange(cfg.ranges.titulares.a1);
  const values = range.getValues(); if (!values.length) return {total:0, sim:0};
  const hdr=values[0].map(v=>String(v).trim()); const ix={};
  hdr.forEach((h,i)=> ix[h]=i);
  const iEmail = ix[cfg.cols.email], iRGPD=ix[cfg.cols.rgpd];
  if (iEmail==null || iRGPD==null) return {total:0, sim:0};
  const emailLC = String(email||'').trim().toLowerCase();
  let total=0, sim=0; nao=0;
  for (let r=1; r<values.length; r++){
    const cell = String(values[r][iEmail]||'');
    const emails = cell.split(/[;,]/).map(s=>s.trim().toLowerCase()).filter(Boolean);
    if (emails.includes(emailLC)){ total++; if (String(values[r][iRGPD]||'').trim().toLowerCase()==='sim') sim++; if (String(values[r][iRGPD]||'').trim().toLowerCase()==='não') nao++; }
  }
  return { total, sim, nao };
}

function setRgpdForEmail_(emailLC, accept, cfg){
  const ss = SpreadsheetApp.openById(cfg.titularesId);
  let range = null;
  if (cfg.rangeNameOrA1 && !/[!]/.test(cfg.rangeNameOrA1)) {
    try { range = ss.getRangeByName(cfg.rangeNameOrA1); } catch(_){}
  }
  if (!range) {
    // fallback: sheet + A1 se vierem separados
    const sheet = ss.getSheetByName(cfg.sheetName || 'Titulares');
    range = sheet.getRange(cfg.a1 || 'A6:V');
  }

  const values = range.getValues(); if (!values.length) return 0;
  const header = values[0].map(String);
  const idx = header.reduce((m, h, i)=>{ m[String(h).trim()] = i; return m; }, {});
  const iEmail = idx[cfg.colEmail || 'e-mail'];
  const iRGPD  = idx[cfg.colRgpd  || 'RGPD'];
  if (iEmail == null || iRGPD == null) throw new Error('Colunas "e-mail" ou "RGPD" não encontradas');

  let touched = 0;
  for (let r=1; r<values.length; r++){
    const cell = String(values[r][iEmail] || '').toLowerCase();
    const emails = cell.split(/[;,]/).map(s=>s.trim());
    if (emails.includes(emailLC)){
      values[r][iRGPD] = accept ? 'Sim' : 'Não';
      touched++;
    }
  }
  if (touched) range.setValues(values);
  return touched;
}

function setRgpdForEmail_AuthCore_(info, cfg, want){
  const sheet = info.range.getSheet(), r0 = info.range.getRow(), c0 = info.range.getColumn(), bodyRows = info.range.getNumRows() - 1;
  /*
  cfg = __defCfg(cfg);

  const range   = info.range;                 // A6:V ou equivalente
  const sheet   = range.getSheet();
  const nRows   = range.getNumRows();
  const r0      = range.getRow();
  const c0      = range.getColumn();
  const bodyRows = nRows - 1;
  */
  if (bodyRows <= 0) return 0;

  /*
  const iEmail  = info.iEmail;
  const iRGPD   = info.iRGPD;
  if (iEmail==null || iRGPD==null) throw new Error('Config/colunas RGPD ou e-mail em falta');

  const emailRange = sheet.getRange(r0+1, c0+iEmail, bodyRows, 1);
  const rgpdRange  = sheet.getRange(r0+1, c0+iRGPD,  bodyRows, 1);

  const emailVals  = emailRange.getDisplayValues();
  const rgpdVals   = rgpdRange.getValues();
  */
  const emailVals = sheet.getRange(r0+1, c0+info.iEmail, bodyRows, 1).getDisplayValues();
  const rgpdRange = sheet.getRange(r0+1, c0+info.iRGPD, bodyRows, 1), rgpdVals = rgpdRange.getValues();
  
  let changed = 0;
  for (let i=0; i<bodyRows; i++){
	/*
    const emails = String(emailVals[i][0]||'')
      .split(/[;,]/).map(s=>s.trim().toLowerCase()).filter(Boolean);
    if (emails.includes(info.emailLC)){
      const prev = String(rgpdVals[i][0]||'').trim();
      if (prev !== want){
        rgpdVals[i][0] = want;
        changed++;
      }
    }
	*/
	if (String(emailVals[i][0]||'').split(/[;,]/).map(s=>s.trim().toLowerCase()).includes(info.emailLC)){
      if (String(rgpdVals[i][0]||'').trim() !== want){ rgpdVals[i][0] = want; changed++; }
	}
  }
  if (changed) { rgpdRange.setValues(rgpdVals); SpreadsheetApp.flush(); }
  return changed;
}

/*
function dedupeEmails_AuthCore_(arr){
  const out=[]; const seen=new Set();
  (arr||[]).forEach(e=>{ const v=String(e||'').trim().toLowerCase(); if (v && !seen.has(v)){ seen.add(v); out.push(v);} });
  return out;
}


function renderRgpdPage_AuthCore_(DBG, email, ticket, gatesCfg){
  gatesCfg = __defCfg(gatesCfg);

  //L('function renderRgpdPage_Authcore_');
  const t = HtmlService.createTemplateFromFile('RGPD'); // o HTML acima
  t.CANON_URL = canonicalAppUrl_();           // a tua função já existente que dá o URL canónico
  t.TICKET = ticket || '';
  t.DEBUG = DBG ? '1' : '';
  t.PAGE_TAG = 'RGPD';
  //t.GATES_CFG_JSON = JSON.stringify(gatesCfg || {}); // passa config do HOST
  return t.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- Expansão para a App Anúncios ---
*/
function getProfileStats_(ticket, cfg) {
  const sess = requireSession(ticket);
  cfg = __defCfg(cfg);
  const st = getRgpdStatusFor(ticket, cfg);
  const info = getTitularesRowsByEmail_AuthCore_(sess.email, cfg);
  
  let totalSaldo = 0;
  const idxSaldo = info.ix[cfg.cols.saldo || "Saldo"];
  if (idxSaldo != null) {
    for (const r of info.matches) {
      totalSaldo += parsePtNumber_AuthCore_(info.values[r][idxSaldo]);
    }
  }
  return {
    email: sess.email,
    rgpdState: st.state,      // 'total', 'parcial', 'pendente'
    saldo: totalSaldo,        // Saldo somado
    hasLines: info.matches.length > 0
  };
}

function hostListRgpdRowsFor_(ticket, cfg) {
  //const sess = requireSession(ticket);
  //cfg = __defCfg(cfg);
  const info = getTitularesRowsByEmail_AuthCore_(sess.email, cfg);
  const out = [];
  const r0 = info.range.getRow();
  const iSem = info.ix[cfg.cols.semanas || "Semanas"];
  for (const r of info.matches) {
    out.push({
      row: r0 + r, // linha real na sheet
      semanas: iSem != null ? info.values[r][iSem] : "",
      rgpd: String(info.values[r][info.iRGPD]||"").trim().toLowerCase() === "sim"
    });
  }
  return out;
}

function hostSaveRgpdRowsFor_(ticket, acceptedRows, cfg) {
  //cfg = __defCfg(cfg);
  const sess = requireSession(ticket);
  const info = getTitularesRowsByEmail_AuthCore_(sess.email, cfg);
  const sheet = info.range.getSheet();
  const r0 = info.range.getRow();
  const c0 = info.range.getColumn();
  const bodyRows = info.range.getNumRows() - 1;
  if (bodyRows <= 0) return { ok: true, touched: 0 };
  
  const rgpdRange = sheet.getRange(r0 + 1, c0 + info.iRGPD, bodyRows, 1);
  const rgpdVals = rgpdRange.getValues();
  let touched = 0;
  
  //let acceptedSemanas = [];
  //const iSem = info.ix[cfg.cols.semanas || "Semanas"];
  
  for (const r of info.matches) {
    const sheetRow = r0 + r; 
    const wantAccept = acceptedRows.includes(sheetRow);
    const next = wantAccept ? "Sim" : "Não";
    if (String(rgpdVals[r-1][0]||"") !== next) {
      rgpdVals[r-1][0] = next;
      touched++;
    }
	/*
    if (wantAccept && iSem != null) {
      const rawSemanas = info.values[r][iSem] || "";
      const formatted = String(rawSemanas).split(/[+,\s]+/).filter(Boolean).join(" · ");
      if (formatted) acceptedSemanas.push(formatted);
    }
	*/
  }
  if (touched) {
    rgpdRange.setValues(rgpdVals);
    SpreadsheetApp.flush();
    /*
	try {
      const toEmail = cfg.notify.to || "secretario-direcao@titulares-portobelo.pt";
      let bodyText = "O associado " + sess.email + " atualizou o RGPD.\n\n";
      bodyText += acceptedSemanas.length > 0 
        ? acceptedSemanas.map(sem => "Aceite: " + sem).join("\n")
        : "Nenhuma linha aceite (Rejeitado).";
      MailApp.sendEmail(toEmail, "[App] Atualização de RGPD - " + sess.email, bodyText);
    } catch (e) {}
	*/
  }
  return { ok: true, touched };
}
