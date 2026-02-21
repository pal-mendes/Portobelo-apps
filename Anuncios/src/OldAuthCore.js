// =========================
// File: AuthCore.gs
// =========================


/*
	SSO entre apps: se quiseres que um login numa app valha na outra, copia para as Script Properties do novo projeto os mesmos valores de SESSION_HMAC_SECRET_B64 e STATE_SECRET. Caso contrário, cada app terá sessões independentes.

*/



// ===== Config / Constantes comuns =====
const OAUTH_SCOPES = 'openid email profile';
const STATE_MAX_AGE_MS = 15 * 60 * 1000; // 15 min
const NONCE_TTL_SEC = 180; // 3 min para nonce→ticket

const REDIRECT_URI = PropertiesService.getScriptProperties().getProperty('REDIRECT_URI');

// ===== (Opcional) Debug helpers reutilizáveis =====
function isDebug_(e){ return e && e.parameter && e.parameter.debug === '1'; }

function makeLogger_(DBG){
  const start=new Date(), log=[];
  function L(){ if(!DBG) return; const dt=((new Date())-start); const hh=new Date(dt).toISOString().substr(11,8); log.push(hh+' '+[].map.call(arguments,String).join(' '));}
  L.dump = () => log.slice();
  return L;
}

// ===== Script properties helpers =====


function getScriptProp_(k) { return PropertiesService.getScriptProperties().getProperty(k); }

function setScriptProp_(k, v) { return PropertiesService.getScriptProperties().setProperty(k, v); }

// ===== URL canónico do deployment =====
function canonicalAppUrl_() {
  var url = ScriptApp.getService().getUrl();
  return url.replace(/\/a\/[^/]+\/macros/, '/macros'); // força .../macros/s/ID/exec
}

// Helper para incluir ficheiros HTML (templates parciais)
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// Preferir REDIRECT_URI definido; fallback para canónico
function redirectUri_() {
  // Preferir a property se estiver correcta; senão, cair para a canónica deste deployment
  return (REDIRECT_URI && REDIRECT_URI.trim()) || canonicalAppUrl_();
}

// ===== OAuth URL =====

function toQueryString_(obj) {
  var parts = [];
  Object.keys(obj||{}).forEach(function (k) {
    const v = obj[k];
    if (v === undefined || v === null) return;
    parts.push(encodeURIComponent(k) + '=' + encodeURIComponent(String(v)));
  });
  return parts.join('&');
}

// ===== Ticket “stateless” com HMAC =====
function getSessSecretBytes_() {
  const p = PropertiesService.getScriptProperties();
  let b64 = p.getProperty('SESSION_HMAC_SECRET_B64');
  if (!b64) {
    // Deriva 32 bytes “random enough” a partir de dois UUIDs (evita depender de RNG externo)
    const seed = Utilities.getUuid() + ':' + Utilities.getUuid() + ':' + Date.now();
    const bytes = Utilities.computeHmacSha256Signature(seed, seed); // 32 bytes
    b64 = Utilities.base64EncodeWebSafe(bytes);
    p.setProperty('SESSION_HMAC_SECRET_B64', b64);
    // opcional: apaga o legado se existir
    p.deleteProperty('SESSION_HMAC_SECRET');
  }
  return Utilities.base64DecodeWebSafe(b64);
}


function issueSessionToken_(email, days) {
  const exp = Date.now() + ((days || 14) * 24 * 60 * 60 * 1000);
  const payload = JSON.stringify({ email: email, exp: exp, v: 1, iat: Date.now() });
  
  const pBytes = Utilities.newBlob(payload).getBytes();              // bytes do payload
  const pB64 = Utilities.base64EncodeWebSafe(pBytes);               // string base64 do payload
  const pB64bytes = Utilities.newBlob(pB64).getBytes();              // bytes da string base64
  const sigBytes = Utilities.computeHmacSha256Signature(pB64bytes, getSessSecretBytes_()); // bytes,bytes
  const sig = Utilities.base64EncodeWebSafe(sigBytes);

  return pB64 + '.' + sig;
}

function validateSessionToken_(tok) {
  if (!tok || tok.indexOf('.') < 0) return null;
  const parts = tok.split('.'), pB64 = parts[0], sig = parts[1];

  const pB64bytes = Utilities.newBlob(pB64).getBytes();              // bytes da string base64
  const expSig = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(pB64bytes, getSessSecretBytes_()) // bytes,bytes
  );
  if (expSig !== sig) return null;

  const json = Utilities.newBlob(Utilities.base64DecodeWebSafe(pB64)).getDataAsString();
  const data = JSON.parse(json);
  if (!data || !data.email || !data.exp || Date.now() > data.exp) return null;
  return data;
}



// Mantém visível a API usada pelo front-end:
function isTicketValid(ticket) { return !!validateSessionToken_(ticket); }


// ---------- Helpers de sessão ----------
function getSession(ticket) {
  const d = validateSessionToken_(ticket);
  return d ? JSON.stringify(d) : null;
}

function requireSession(ticket) {
  const p = getSession(ticket);
  if (!p) throw new Error('Sessão inválida/expirada');
  return JSON.parse(p);
}


// ===== State / nonce (HMAC) =====
function getStateSecret_() {
  let s = getScriptProp_('STATE_SECRET');
  if (!s) {
    const rand = Utilities.getUuid();
    s = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(rand, rand));
    setScriptProp_('STATE_SECRET', s);
  }
  return s;
}


// Gera state assinado: base64url(payload) + '.' + base64url( HMAC_SHA256(payload, secret) )
// aceitar flags no state
function createStateToken_(dbg, embed, nonceOpt) {
  const payload = {
    ts: Date.now(),
    nonce: nonceOpt || Utilities.getUuid(),
    dbg: !!dbg,
    embed: !!embed
  };
  const payloadBytes = Utilities.newBlob(JSON.stringify(payload),'application/json').getBytes();
  const secretBytes  = Utilities.base64DecodeWebSafe(getStateSecret_());
  const sigBytes     = Utilities.computeHmacSha256Signature(payloadBytes, secretBytes);
  return Utilities.base64EncodeWebSafe(payloadBytes) + '.' + Utilities.base64EncodeWebSafe(sigBytes);
}


function parseStateToken_(state) {
  if (!state || state.indexOf('.') < 1) return { ok: false };
  const [payloadB64, sigB64] = state.split('.', 2);
  try {
    const payloadBytes = Utilities.base64DecodeWebSafe(payloadB64);
    const secretBytes = Utilities.base64DecodeWebSafe(getStateSecret_());
    const expSigBytes = Utilities.computeHmacSha256Signature(payloadBytes, secretBytes);
    const gotSigBytes = Utilities.base64DecodeWebSafe(sigB64);
    if (expSigBytes.length !== gotSigBytes.length) return { ok: false };
    let diff = 0; for (let i = 0; i < expSigBytes.length; i++) diff |= (expSigBytes[i] ^ gotSigBytes[i]);
    if (diff !== 0) return { ok: false };
    const payload = JSON.parse(Utilities.newBlob(payloadBytes).getDataAsString());
    if (typeof payload.ts !== 'number') return { ok: false };
    if (Date.now() - payload.ts > STATE_MAX_AGE_MS) return { ok: false };
    return { ok: true, payload: payload };
  } catch (e) { return { ok: false }; }
}


// nonce → ticket cache (para polling)
function putTicketForNonce_(nonce, ticket) {
  if (!nonce || !ticket) return;
  CacheService.getScriptCache().put('nonce:'+nonce, ticket, NONCE_TTL_SEC);
}

function takeTicketForNonce_(nonce) {
  if (!nonce) return '';
  const cache = CacheService.getScriptCache();
  const k = 'nonce:'+nonce;
  const t = cache.get(k) || '';
  if (t) cache.remove(k);
  return t;
}

// Poll: devolve o ticket (e consome-o) para um dado nonce
function pollTicket(nonce) {
  return takeTicketForNonce_(nonce) || '';
}

// ===== OAuth URL =====

function getClientId_() { return getScriptProp_('CLIENT_ID'); }

function getClientSecret_() { return getScriptProp_('CLIENT_SECRET'); }

function buildAuthUrlFor(nonce, dbg, embed) {
  const state = createStateToken_(!!dbg, !!embed, nonce);
  const params = {
    client_id: getClientId_(),
    redirect_uri: redirectUri_(),           // TEM de ser exatamente o registado no OAuth client
    response_type: 'code',
    scope: OAUTH_SCOPES,
    access_type: 'offline',
    prompt: 'consent',
    include_granted_scopes: 'true',
    state: state
  };
  return 'https://accounts.google.com/o/oauth2/v2/auth?' + toQueryString_(params);
}


// ===== Render do Login (comum às apps) =====
function renderLoginPage_(DBG, serverLog, wipe) {
  try {
    const t = HtmlService.createTemplateFromFile('Login');
    t.CANON_URL = canonicalAppUrl_();
    t.CLIENT_ID = getClientId_();
    t.DEBUG = DBG ? '1' : '';
    t.SERVER_LOG = (serverLog && serverLog.join) ? serverLog.join('\n') : String(serverLog || '');
    t.WIPE = wipe ? '1' : '';
    var out = t.evaluate();
    out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return out;
  } catch (err) {
    const msg = 'Login.html evaluate() falhou:\n' + String(err) + '\n--- SERVER LOG ---\n' +
      ((serverLog && serverLog.join) ? serverLog.join('\n') : String(serverLog || ''));
    return HtmlService.createHtmlOutput('<pre style="white-space:pre-wrap">' + msg + '</pre>');
  }
}

