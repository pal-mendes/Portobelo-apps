
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
var LIB_NOTIFY = { to:"log-apps@titulares-portobelo.pt", ccAllRows:true };

function __defCfg(cfgParam){
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
  var url = ScriptApp.getService().getUrl().replace(/\/a\/[^\/]+\/macros\//, "/macros/");
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
  // Com DEBUG AGRESSIVO DO TICKET
  if (!tok || tok.indexOf(".") < 0) throw new Error("[DEBUG] Formato de token inválido: " + tok);
  const parts = tok.split("."), pB64 = parts[0], sig = parts[1];

  let data;
  try {
     data = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(pB64)).getDataAsString());
  } catch(e) {
     throw new Error("[DEBUG] Falha no parse do token JSON");
  }

  const secretBytes = getSessSecretBytes_();
  const expSig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(Utilities.newBlob(pB64).getBytes(), secretBytes));

  // A MENSAGEM CHAVE: Se falhar, diz-nos a versão exata do ticket e o segredo!
  if (expSig !== sig) {
     const sp = PropertiesService.getScriptProperties();
     const cs = sp.getProperty("CLIENT_SECRET") || "";
     const hmac = sp.getProperty("SESSION_HMAC_SECRET_B64") || "";
     const csSuffix = cs.length > 4 ? cs.slice(-4) : cs;
     const hmacSuffix = hmac.length > 4 ? hmac.slice(-4) : hmac;
     
     throw new Error(`[DEBUG] HMAC Inválida! Ticket gerado na versão v:${data.v}. CLIENT_SECRET termina em: ${csSuffix}. SESSION_HMAC_SECRET_B64 termina em: ${hmacSuffix}. (Provável redirecionamento para o /exec antigo)`);
  }

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

function buildAuthUrlFor_(nonce, dbg, embed, cfg, clientUrl) {
  // Nota: tornamos a construção do URL "interna" e expomos via buildAuthUrlFor()
  const L = makeLogger_(dbg);
  L('Entrada em buildAuthUrlFor_');

  const { clientId, redirectUri: fallbackRu } = resolveCfg_(cfg);
  if (!clientId) throw new Error("CLIENT_ID ausente.");
  L('clientId =' + clientId);

  // AGORA: A configuração explícita do servidor (fallbackRu) tem prioridade absoluta!
  const finalRu = clientUrl || fallbackRu; 
  if (!finalRu) throw new Error("REDIRECT_URI ausente.");
  L('clientUrl=' + clientUrl + ', fallbackRu =' + fallbackRu + ' finalRu =' + finalRu);

  const params = {
    client_id: clientId, redirect_uri: finalRu, response_type: 'code',
    scope: OAUTH_SCOPES, access_type: 'offline', prompt: 'select_account consent',
    include_granted_scopes: 'true', state: createStateToken_(!!dbg, !!embed, nonce, finalRu),
  };
  return "https://accounts.google.com/o/oauth2/v2/auth?" + toQueryString_(params);
}

// ===== Render do Login (comum às apps) =====
function renderLoginPage_(opts) {
  // FORÇAR O CANON URL NO HTML
  const L = makeLogger_(opts.debug);
  //const L = makeLogger_(1); //parece que opts.debug não está a funcionar?
  L('function renderLoginPage_');

  const t = HtmlService.createTemplateFromFile("Login");
  // Permite à App forçar o URL de testes
  t.CANON_URL = opts.canon || canonicalAppUrl_(); 
  //t.CLIENT_ID = getClientId_(); // opcional; o HTML atual nem usa
  t.AUTOSTART = "1"; t.DEBUG = opts.debug ? "1" : ""; // auto-inicia o popup
  t.DEBUG = opts.debug ? "1" : "";
  t.SERVER_LOG = opts.serverLog && opts.serverLog.join ? opts.serverLog.join("\n") : String(opts.serverLog || "");
  t.WIPE = opts.wipe ? "1" : "";  
  L('renderLoginPage_: opts.wipe = ' + opts.wipe + ', t.WIPE = ' + t.WIPE);
  t.ticket = ""; t.TICKET = ""; t.SERVER_VARS = ""; t.PAGE_TAG = 'LOGIN';

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
  // 400 Bad Request ou 401 Unauthorized
  // a boa prática padrão (Standard) é aceitar qualquer código da "família dos 200" (2xx) como sucesso.
  // Algumas APIs podem devolver 201 Created (Criado) ou 204 No Content (Sem conteúdo), que são sucessos
  //if (status !== 200) {
  //if (status >= 300) {
  if (status < 200 || status >= 300) {
    const errHtml = `<!doctype html><html><head><meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
      body { font-family: system-ui, sans-serif; padding: 30px 20px; text-align: center; color: #111; line-height: 1.5; max-width: 500px; margin: 0 auto; }
      .box { background: #fee2e2; border: 1px solid #f87171; border-radius: 12px; padding: 20px; margin-top: 20px; }
    </style>
    </head><body>
      <h2>Autenticação interrompida</h2>
      <div class="box">
        <p style="color: #991b1b; font-weight: 500; margin-top: 0;">O código de segurança expirou ou já foi utilizado.</p>
        <p style="color: #7f1d1d; font-size: 14px; margin-bottom: 0;">Isto acontece frequentemente se a página for recarregada (refresh) durante o processamento no telemóvel.</p>
      </div>
      <p style="margin-top: 30px;"><b>Por favor, feche esta janela/aba e inicie o processo de login novamente na Área do Associado.</b></p>
      <script>
        setTimeout(function(){ try { window.close(); } catch(e){} }, 4000);
      </script>
    </body></html>`;
    
    return HtmlService.createHtmlOutput(errHtml).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
  var appUrl = ${JSON.stringify(finalRu)}; // O URL da Web App que gerou o pedido

  // 1) guardar ticket e cookie na memória local do navegador
  try{ localStorage.setItem('sessTicket', t); }catch(_){}
  try{ document.cookie = 'sessTicket='+encodeURIComponent(t)+'; Path=/; SameSite=Lax; Secure'; }catch(_){}

  // 2) postMessage com o ticket para tentar avisar a janela mãe (se for popup)
  try{ if (window.opener) window.opener.postMessage({type:'portobelo_ticket', ticket:t}, '*'); }catch(_){}

  // 3) fecha o popup – quem navega é a janela principal via goWithTicket() - suportar iPhone
  setTimeout(function(){
    // Tenta fechar a janela primeiro
    try{ window.close(); }catch(_){}

    // Se o código chegou aqui (a janela não fechou), navega à força de volta para a App!
    window.location.replace(appUrl + "?ticket=" + encodeURIComponent(t));
  }, 600);
})();
</script>`;

  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// =========================
// AuthCoreLib — Gates (membership + RGPD)
// =========================

function logFailedAccessToSheet_(email, name, saldo, semanas, reason, cfg) {
  try {
    const ss = SpreadsheetApp.openById(cfg.ssTitularesId);
    let sheet = ss.getSheetByName("Acessos");
    if (!sheet) {
      sheet = ss.insertSheet("Acessos");
      sheet.appendRow(["Data", "Nome", "Email", "Saldo €", "Semanas", "Motivo", "Web App"]);
    }
    sheet.appendRow([new Date(), name || "", email || "", saldo !== undefined ? saldo : "", semanas || "", reason || "", cfg.appName || "Desconhecida"]);
  } catch(e) { console.log("Erro a gravar acessos:", e); }
}

function extractFinancialInfo_(info, cfg) {
  let pago = "", saldo = "", semanas = "";
  if (info.matches.length > 0) {
    const iPago = info.ix[cfg.cols.pago || "€"];
    const iSaldo = info.ix[cfg.cols.saldo || "Saldo"];
    const iSem = info.ix[cfg.cols.semanas || "Semanas"];
    let totP = 0; let totS = 0; let semArr = [];
    for(const r of info.matches) {
      if (iPago != null) totP += parsePtNumber_AuthCore_(info.values[r][iPago]);
      if (iSaldo != null) totS += parsePtNumber_AuthCore_(info.values[r][iSaldo]);
      if (iSem != null && info.values[r][iSem]) semArr.push(info.values[r][iSem]);
    }
    pago = totP; saldo = totS; semanas = semArr.join(" · ");
  }
  return { pago, saldo, semanas };
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
    logFailedAccessToSheet_(email, sess.name, fin.saldo, fin.semanas, "S/ linhas ou quotas atraso: pagamentos=" + fin.pago + "saldo=" + fin.saldo, gatesCfg);
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
  const canon = ScriptApp.getService().getUrl().replace(/\/a\/[^\/]+\/macros\//, "/macros/");
  return HtmlService.createHtmlOutput('<meta charset="utf-8"><h3>Acesso não autorizado</h3><p>Este endereço não está na lista da fase de validação.</p><p><a href="'+canon+'?action=login'+(DBG?'&debug=1':'')+'">Voltar</a></p>').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}



function parsePtNumber_AuthCore_(val){
  // Se o Google Sheets já nos der um número nativo, devolve-o intacto!
  if (typeof val === 'number') return val;
  
  // Caso contrário, limpa o texto formatado (ex: "1.000,50" -> 1000.5)
  const clean = String(val||'').replace(/\s/g,'').replace(/\./g,'').replace(',','.').replace(/[^\d.\-]/g,'');
  const n = parseFloat(clean); 
  return isNaN(n) ? 0 : n;
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
  console.log("hasAnyRowMinPayment_AuthCore_: iPago=" + iPago);
  if (iPago == null) return false;
  const lim = typeof minEUR === 'number' ? minEUR : 1;
  console.log("hasAnyRowMinPayment_AuthCore_: lim=" + lim);

  for (const r of info.matches){
    const val = info.values[r][iPago];
    console.log("hasAnyRowMinPayment_AuthCore_: pago=" + val + ", lim=" + lim);
    if (parsePtNumber_AuthCore_(val) >= lim) {
      console.log("hasAnyRowMinPayment_AuthCore_ OK");
      return true;
    }
  }
  console.log("hasAnyRowMinPayment_AuthCore_ NOK");
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
  t.CANON_URL  = canonicalAppUrl_(); // << usa o do HOST se vier
  t.ticket = opts.ticket || '';
  t.TICKET = opts.ticket || '';
  t.DEBUG  = opts.debug ? '1' : '';
  t.WIPE = opts.wipe ? "1" : "";

// NOVAS VARIÁVEIS DE ALERTA DE BLOQUEIO
  t.BLOCK_MSG = opts.blockMsg || "";
  t.BLOCK_BTN_LABEL = opts.blockBtnLabel || "";
  t.BLOCK_BTN_URL = opts.blockBtnUrl || "";

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
    return HtmlService.createHtmlOutput('<pre style="white-space:pre-wrap">' + msg + "</pre>");
  }
}

function acceptRgpdForMe(ticket, decision, gatesCfg){
  gatesCfg = __defCfg(gatesCfg);
  const sess = requireSession(ticket), accept = (decision === 'accept');
  const rowsInfo = getTitularesRowsByEmail_AuthCore_(sess.email, gatesCfg);
  var changed = setRgpdForEmail_AuthCore_(rowsInfo, gatesCfg, accept ? 'Sim' : 'Não');
  return { ok:true, changed: changed|0 };
}



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
  if (bodyRows <= 0) return 0;


  const emailVals = sheet.getRange(r0+1, c0+info.iEmail, bodyRows, 1).getDisplayValues();
  const rgpdRange = sheet.getRange(r0+1, c0+info.iRGPD, bodyRows, 1), rgpdVals = rgpdRange.getValues();
  
  let changed = 0;
  for (let i=0; i<bodyRows; i++){
    if (String(emailVals[i][0]||'').split(/[;,]/).map(s=>s.trim().toLowerCase()).includes(info.emailLC)){
        if (String(rgpdVals[i][0]||'').trim() !== want){ rgpdVals[i][0] = want; changed++; }
    }
  }
  if (changed) { rgpdRange.setValues(rgpdVals); SpreadsheetApp.flush(); }
  return changed;
}



function getProfileStats_(ticket, cfg) {
  //console.log("getProfileStats_()");
  const sess = requireSession(ticket);
  //console.log("getProfileStats_: email=", sess.email);
  cfg = __defCfg(cfg);
  const st = getRgpdStatusFor(ticket, cfg);
  const info = getTitularesRowsByEmail_AuthCore_(sess.email, cfg);
  
  let totalPago = 0;
  const idxPago = info.ix[cfg.cols.pago || "€"];
  //console.log("getProfileStats_: idxSaldo=", idxSaldo);
 
  let totalSaldo = 0;
  //console.log("getProfileStats_: totalSaldo=", totalSaldo);
  const idxSaldo = info.ix[cfg.cols.saldo || "Saldo"];
  //console.log("getProfileStats_: idxSaldo=", idxSaldo);

  // NOVO: Procurar a coluna dos telefones de forma flexível
  let idxTel = cfg.cols.telefones ? info.ix[cfg.cols.telefones] : null;
  if (idxTel == null) idxTel = info.ix["Telefones"];
  if (idxTel == null) idxTel = info.ix["Telemóvel"];
  if (idxTel == null) idxTel = info.ix["Telemóveis"];
  if (idxTel == null) idxTel = 10; // Fallback seguro para a Coluna K (índice 10)

  // NOVO: Guardar as linhas cruas (emails e telefones) para a App Anúncios usar
  const cardsLinhas = [];

  for (const r of info.matches) {
    if (idxSaldo != null) {
      console.log("getProfileStats_: Pago da linha=", info.values[r][idxPago]);
      totalPago += parsePtNumber_AuthCore_(info.values[r][idxPago]);
      console.log("getProfileStats_: totalPagoo=", totalPagoo);

      console.log("getProfileStats_: Saldo da linha=", info.values[r][idxSaldo]);
      totalSaldo += parsePtNumber_AuthCore_(info.values[r][idxSaldo]);
      console.log("getProfileStats_: totalSaldo=", totalSaldo);
    }
    
    cardsLinhas.push({
      emails: info.iEmail != null ? String(info.values[r][info.iEmail]) : "",
      telefones: idxTel != null ? String(info.values[r][idxTel]) : ""
    });
    //console.log("getProfileStats_: telefones=", String(info.values[r][idxTel]));
  }

  totalPago = Math.round(totalPago * 100) / 100;
  console.log("getProfileStats_: totalPago=", totalPago);

  //console.log("getProfileStats_: totalSaldo=", totalSaldo);
  // A MAGIA DA LIMPEZA: Arredonda o saldo total para 2 casas decimais (ex: 0.00)
  // Isto elimina os lixos de precisão decimal do JavaScript (ex: -1.776e-15)
  totalSaldo = Math.round(totalSaldo * 100) / 100;
  console.log("getProfileStats_: totalSaldo=", totalSaldo);

  return {
    email: sess.email,
    rgpdState: st.state,      // 'total', 'parcial', 'pendente'
    pago: totalPago,        // 
    saldo: totalSaldo,        // Saldo somado
    hasLines: info.matches.length > 0,
    cardsLinhas: cardsLinhas  // A magia volta a acontecer aqui!
  };
}

function hostListRgpdRowsFor_(ticket, cfg) {
  const sess = requireSession(ticket);
  cfg = __defCfg(cfg);
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
  // 1. Restauramos a configuração para ler o mail de destino
  cfg = __defCfg(cfg);
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
  
  // Vamos guardar quais semanas foram aceites e quais foram rejeitadas
  let acceptedSemanas = [];
  let rejectedSemanas = [];
  const iSem = info.ix[cfg.cols.semanas || "Semanas"];
  
  // NOVO: Coleção para guardar todos os e-mails únicos das linhas afetadas
  const allEmailsToCc = new Set();
  
  // Garantimos que o e-mail de quem fez o login entra sempre no CC
  if (sess.email) allEmailsToCc.add(String(sess.email).trim().toLowerCase());
  
  for (const r of info.matches) {
    const sheetRow = r0 + r; 
    const wantAccept = acceptedRows.includes(sheetRow);
    const next = wantAccept ? "Sim" : "Não";
  
	if (String(rgpdVals[r-1][0]||"") !== next) {
      rgpdVals[r-1][0] = next;
      touched++;
    }

    
    // NOVO: Extrair e limpar todos os e-mails desta linha (separados por ; ou ,)
    if (info.iEmail != null) {
       const rawEmails = String(info.values[r][info.iEmail] || "");
       rawEmails.split(/[;,]/).forEach(e => {
          const cleanEmail = e.trim().toLowerCase();
          if (cleanEmail) allEmailsToCc.add(cleanEmail);
       });
    }
	
    // Formatar as semanas para o E-mail
    if (iSem != null) {
      const rawSemanas = info.values[r][iSem] || "";
      const formatted = String(rawSemanas).split(/[+,\s]+/).filter(Boolean).join(" · ");
      if (formatted) {
        if (wantAccept) acceptedSemanas.push(formatted);
        else rejectedSemanas.push(formatted);
      } 
    }
  }
  if (touched) {
    rgpdRange.setValues(rgpdVals);
    SpreadsheetApp.flush();

  // 2. ENVIO DO E-MAIL CENTRALIZADO
  try {
      const toEmail = cfg.notify.to || "log-apps@titulares-portobelo.pt";
      let bodyText = "O associado " + (sess.name ? sess.name + " (" + sess.email + ")" : sess.email) + " atualizou o seu consentimento do RGPD da Associação.\n\n";
   
      if (acceptedSemanas.length > 0) {
        bodyText += "✅ Semanas ACEITES:\n" + acceptedSemanas.map(sem => "- " + sem).join("\n") + "\n\n";
      }
      if (rejectedSemanas.length > 0) {
        bodyText += "❌ Semanas REJEITADAS:\n" + rejectedSemanas.map(sem => "- " + sem).join("\n") + "\n\n";
      }
      bodyText += "Associação dos titulares de DRHP do Portobelo\n";
      bodyText += "https://titulares-portobelo.pt\n\n";
      bodyText += "Em caso de dúvidas, contactar geral@titulares-portobelo.pt\n";

      // Converte o Set de e-mails únicos numa string separada por vírgulas para o MailApp
      const ccEmails = Array.from(allEmailsToCc).join(",");
      MailApp.sendEmail({
        to: toEmail,
        cc: ccEmails, // Agora envia para todos os e-mails associados àquelas semanas!
        replyTo: "geral@titulares-portobelo.pt",
        subject: "[Associação Portobelo] Atualização de RGPD - " + sess.email,
        body: bodyText,
        name: "Associação dos titulares de DRHP do Portobelo"
      });
    } catch (e) {
      console.log("Erro a enviar email RGPD: " + e);
    }
  }
  return { ok: true, touched };
}
