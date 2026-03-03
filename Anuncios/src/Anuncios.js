
// =========================
// File: Anuncios.gs
// =========================

/*

 Google Apps Script: Web App para exibir a aba "Anúncios" com filtros e ordenação

  Acesso restrito aos titulares contribuintes da associação, até porque não é permitido por lei alugar timeshares sem licença de alojamento local

  ****** Apps Script ***********

  "Execute as" = "Me" (para o web app ler/escrever no Sheets da associação)
  "Who has access" = "Anyone" (ou "Anyone with Google Account").

  O domínio titulares-portobelo.pt pertence à nossa associação, e é gerido com OVH.


  ********** TESTES *********************
  Abre a página dos anúncios no Sites.
  Carrega F12 para abrir a consola.
  No topo da consola há um menu de contexto (no teu screenshot aparece “top”).
  Clica e escolhe o frame chamado userCodeAppPanel (é o iframe da web app).
  Isto é importante: o localStorage a apagar está nesse frame, não em top.
  Com o contexto certo selecionado, corre:

  //Para consultar o ticket guardado
  separador Application → Storage → Cookies → clica no domínio https://script.google.com e procura a cookie sessTicket
  localStorage.getItem('sessTicket');
  document.cookie
  //Para apagar o ticket
  localStorage.removeItem('sessTicket');
  document.cookie = 'sessTicket=; Max-Age=0; Path=/; Secure; SameSite=Lax';
  location.reload();

  //Para validar o polling sem CORS
  fetch('./?route=poll&nonce=teste&_=' + Date.now(), {mode:'same-origin'})
    .then(r => r.text()).then(console.log).catch(console.error);

  Se o cache ainda não tiver ticket para esse nonce, deves ver {"ok":false}. Depois do login, verás {"ok":true,"ticket":"..."}
  e a página navega para a Main.

  de quiseres “desligar” o sticky depois, apaga o flag indo a about:blank e na consola:
  localStorage.removeItem('pbDebug')

*/

const VERSION = "v5.0";
const ANUNCIOS_SHEET = "Anúncios";
const ANUNCIOS_HEADER_ROW = 4;
const PEDIDOS_SHEET = "Pedidos";
const TITULARES_SHEET = "Titulares";
const ACESSOS_SHEET = "Acessos";

const ACCESS_OK = 0;
const ACCESS_DENIED = 1;
const ACCESS_LOCKED = 9;

// OAuth Client ID do Google Identity Services  
const CLIENT_ID = PropertiesService.getScriptProperties().getProperty('CLIENT_ID');
const CLIENT_SECRET = PropertiesService.getScriptProperties().getProperty('CLIENT_SECRET');
//const REDIRECT_URI = PropertiesService.getScriptProperties().getProperty('REDIRECT_URI');
// Ex.: https://script.google.com/macros/s/AKfycbz.../exec


let LOGGING = false;



// ---- Sessão stateless com HMAC ----
function getSessSecret_() {
  var p = PropertiesService.getScriptProperties();
  var s = p.getProperty('SESSION_HMAC_SECRET');
  if (!s) { s = Utilities.base64EncodeWebSafe(Utilities.getUuid() + ':' + Date.now()); p.setProperty('SESSION_HMAC_SECRET', s); }
  return s;
}



function b64wBytes_(b) { return Utilities.base64DecodeWebSafe(b); }
function bytesB64w_(bytes) { return Utilities.base64EncodeWebSafe(bytes); }



// ===== DEBUG infra =====
function isDebug_(e) {
  return e && e.parameter && e.parameter.debug === '1';
}

function makeLogger_(DBG) {
  const start = new Date();
  const log = [];
  function L() {
    if (!DBG) return;
    const ts = new Date();
    const t = new Date(ts - start).toISOString().substr(11, 8); // hh:mm:ss desde início
    log.push(t + ' ' + Array.prototype.map.call(arguments, String).join(' '));
  }
  L.dump = () => log.slice();
  return L;
}

function needTopLevelLogin_(loginUrl) {
  var html = '<script>if(top!==self){top.location.href="' + loginUrl + '"}else{location.href="' + loginUrl + '"};</script>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}





// ---------- OAuth Authorization Code ----------
// Constrói o URL de autorização (fluxo authorization_code)
function buildAuthUrl_(DBG, EMBED) {
  var state = createStateToken_(DBG, EMBED); // <-- agora inclui embed
  var params = {
    client_id: getClientId_(),
    redirect_uri: redirectUri_(),
    response_type: 'code',
    scope: OAUTH_SCOPES,
    access_type: 'offline',
    prompt: 'consent',
    include_granted_scopes: 'true',
    state: state
  };
  return 'https://accounts.google.com/o/oauth2/v2/auth?' + toQueryString_(params);
}


function getAuthUrl() { return buildAuthUrl_(); }


function renderMainPage_(DBG, serverLog, ticket) {
  try {
    var t = HtmlService.createTemplateFromFile('Main');
    t.VERSION = VERSION;
    t.count = incrementVisitCounter();
    t.ticket = ticket;
    t.CANON_URL = canonicalAppUrl_(); // mantém /a/<dom>/macros/; sem /u/{n}
    t.DEBUG = DBG ? '1' : '';
    t.SERVER_LOG = (serverLog && serverLog.join) ? serverLog.join('\n') : String(serverLog || '');
    var out = t.evaluate().setTitle('Anúncios Portobelo');
    out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return out;
  } catch (err) {
    var msg = 'Main.html evaluate() falhou:\n' + String(err) + '\n--- SERVER LOG ---\n' +
      ((serverLog && serverLog.join) ? serverLog.join('\n') : String(serverLog || ''));
    return HtmlService.createHtmlOutput('<pre style="white-space:pre-wrap">' + msg + '</pre>');
  }
}



function doGet(e) {
  var DBG = isDebug_(e);
  var L = makeLogger_(DBG);
  L('doGet: start');
  L('e.parameter.action=' + e.parameter.action);
  L('e.parameter.ticket=' + e.parameter.ticket);
  L('e.parameter.__canon=' + e.parameter.__canon);
  L('e.parameter.code=' + e.parameter.code);
  L('e.parameter.state=' + e.parameter.state);
  L('e.parameter.logout=' + e.parameter.logout);
  L('e.parameter.debug=' + e.parameter.debug);
  L('e.parameter.wipe=' + e.parameter.wipe);
  L('here=' + ScriptApp.getService().getUrl());
  L('canon=' + canonicalAppUrl_());
  //L('REDIRECT_URI=' + (REDIRECT_URI||''));
  L('redirectUri_()=' + redirectUri_());


  const WIPE = e && e.parameter && e.parameter.wipe === '1';

  /*
  // (i) Rota de polling — devolve JSON { ok, ticket? }
  if (e && e.parameter && e.parameter.route === 'poll') {
    var nonce = e.parameter.nonce || '';
    var ticket = takeTicketForNonce_(nonce);
    var body = ticket ? { ok:true, ticket: ticket } : { ok:false };
    return ContentService.createTextOutput(JSON.stringify(body))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Cache-Control','no-store')
      .setHeader('Access-Control-Allow-Origin','*');
  }
  */

  if (e && e.parameter && e.parameter.logout === '1') {
    var back = canonicalAppUrl_() + '?wipe=1' + (isDebug_(e) ? '&debug=1' : '');
    var html = [
      '<!doctype html><meta charset="utf-8"><title>Logout</title>',
      '<script>',
      'try{ localStorage.removeItem("sessTicket"); }catch(_){ }',
      // remove cookie na origem script.google.com
      'try{ document.cookie="sessTicket=; Max-Age=0; Path=/; Secure; SameSite=Lax"; }catch(_){ }',
      // volta à página base da web app (com debug se vieste com debug)
      'try{ location.replace(', JSON.stringify(back), '); }catch(_){ location.href=', JSON.stringify(back), '; }',
      '</script>',
      '<p>Terminou a sessão. A voltar à página inicial…</p>'
    ].join('');
    return HtmlService.createHtmlOutput(html)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Ping utilitário
  if (e && e.parameter && e.parameter.ping === '1') {
    var now = new Date().toISOString();
    var body = [
      'PING @ ' + now,
      'here=' + ScriptApp.getService().getUrl(),
      'canon=' + canonicalAppUrl_(),
      'qs=' + (e.queryString || '')
    ].join('\n');
    return HtmlService.createHtmlOutput('<pre>' + body.replace(/[<>&]/g, s => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;' }[s])) + '</pre>')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }



  // 0) Ticket já válido → Main
  L('0) Ticket já válido?');

  try {
    var ticket0 = (e && e.parameter && e.parameter.ticket) || '';
    if (ticket0 && isTicketValid(ticket0)) {
      L('ticket ok → Main');
      return renderMainPage_(DBG, L.dump(), ticket0);
    } else {
      L('ticket' + ticket0 + 'nok');
    }
  } catch (err) {
    L('ticket check err: ' + err);
  }


  // 1) Canónico (normalizar /a/<dom>/macros → /macros) c/ guarda __canon=1
  L('1) Canónico');
  try { L('qs=' + (e && e.queryString)); } catch (_) { } // qs=debug=1
  var here = ScriptApp.getService().getUrl();
  L('here=' + here);  //here=https://script.google.com/a/titulares-portobelo.pt/macros/s/AKfycbzJzmhAdT2oMK9IOJKuNRIerrceumJkE_LnEJW2pXeo_znTrgaCBJM2UW658xM_R6GcQg/exec
  var canon = canonicalAppUrl_(); // idealmente .../macros/s/ID/exec 
  L('canon=' + canon);                       // canon=https://script.google.com/macros/s/AKfycbzJzmhAdT2oMK9IOJKuNRIerrceumJkE_LnEJW2pXeo_znTrgaCBJM2UW658xM_R6GcQg/exec
  var hereNorm = here.replace(/\/a\/[^/]+\/macros/, '/macros');
  L('hereNorm=' + hereNorm);               //hereNorm=https://script.google.com/macros/s/AKfycbzJzmhAdT2oMK9IOJKuNRIerrceumJkE_LnEJW2pXeo_znTrgaCBJM2UW658xM_R6GcQg/exec


  if (hereNorm !== canon && !(e && e.parameter && e.parameter.__canon === '1')) {
    L('redirecting to canon with meta+js');
    var qsParts = [];
    if (e && e.queryString) qsParts.push(e.queryString);
    qsParts.push('__canon=1');
    var qs = '?' + qsParts.join('&');
    var target = canon + qs;

    var htmlCanon = [
      '<!doctype html><meta charset="utf-8"><title>A redirecionar…</title>',
      (DBG ? '<pre style="white-space:pre-wrap;">' + L.dump().join('\n') + '</pre>' : ''),
      '<script>try{location.replace(', JSON.stringify(target), ');}catch(_){location.href=',
      JSON.stringify(target), ';}</script>',
      '<noscript><meta http-equiv="refresh" content="0;url=',
      target.replace(/"/g, '&quot;'),
      '"/></noscript>',
      '<p>Se não avançar automaticamente, <a href="', target.replace(/"/g, '&quot;'), '">clique aqui</a>.</p>'
    ].join('');
    return HtmlService.createHtmlOutput(htmlCanon)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 2) Rota explícita para construir URL de OAuth (botão "Entrar")
  L('2) Rota explícita?');
  if (e && e.parameter && e.parameter.action === 'auth') {
    L('route=auth matched → buildAuthUrl_');
    var FORCED_NONCE = e.parameter && e.parameter._nonce;
    var EMB = (e.parameter.embed === '1');
    var state = createStateToken_(DBG, EMB, FORCED_NONCE);

    var params = {
      client_id: getClientId_(),
      redirect_uri: redirectUri_(),
      response_type: 'code',
      scope: OAUTH_SCOPES,
      access_type: 'offline',
      prompt: 'consent',
      include_granted_scopes: 'true',
      state: state
    };
    var url = 'https://accounts.google.com/o/oauth2/v2/auth?' + toQueryString_(params);

    var html = [
      '<!doctype html><meta charset="utf-8"><title>Auth…</title>',
      // sem target/_top e sem rel=noopener para preservar window.opener
      '<a id="j" href="', url.replace(/"/g,'&quot;'), '">Entrar com conta Google</a>',
      '<script>try{document.getElementById("j").click()}catch(_){location.href=', JSON.stringify(url), ';}</script>'
    ].join('');
    return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  
    /*
    var EMBED = (e.parameter.embed === '1');
    var url = buildAuthUrl_(isDebug_(e), EMBED);
    var html = [
      '<!doctype html><meta charset="utf-8"><title>Auth…</title>',
      '<a id="j" href="', url.replace(/"/g,'&quot;'), '" target="_top" rel="noopener">auth</a>',
      '<script>try{document.getElementById("j").click()}catch(_){location.href=', JSON.stringify(url), ';}</script>'
    ].join('');
    return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    */
    /*
    try {
      var url = buildAuthUrl_(DBG);
      L('url=' + url);

      var htmlAuth;
      if (DBG) {
        // ===== Modo DEBUG: mostra logs e link visível =====
        L('AUTH URL: ' + url.replace(getClientId_(), 'CLIENT_ID'));

        htmlAuth = [
          '<!doctype html><meta charset="utf-8"><title>A redirecionar…</title>',
          '<pre style="white-space:pre-wrap;">', L.dump().join('\n'), '</pre>',
          '<a id="jump" href="', url.replace(/"/g, '&quot;'), '" target="_top" rel="noopener">Ir para autenticação</a>',
          '<script>(function(){try{document.getElementById("jump").click();}catch(e){location.href=',
          JSON.stringify(url), ';}})();</script>',
          '<noscript><meta http-equiv="refresh" content="0; url=', url.replace(/"/g, '&quot;'), '"/></noscript>'
        ].join('');
      } else {
        // ===== Produção: auto-redirect silencioso, sem texto na página =====
        htmlAuth = [
          '<!doctype html><html><head><meta charset="utf-8">',
          '<meta name="viewport" content="width=device-width,initial-scale=1">',
          '<style>html,body{margin:0;background:#fff}</style></head><body>',
          // link oculto com target=_top (evita tocar em window.top.*)
          '<a id="jump" href="', url.replace(/"/g, '&quot;'),
          '" target="_top" rel="noopener" style="position:fixed;left:-9999px;top:-9999px;width:1px;height:1px;opacity:0">_</a>',
          '<script>(function(){var a=document.getElementById("jump");',
          'try{if(a&&a.click)a.click();else location.href=', JSON.stringify(url), ';}',
          'catch(_){location.href=', JSON.stringify(url), ';}})();</script>',
          '<noscript><meta http-equiv="refresh" content="0; url=', url.replace(/"/g, '&quot;'), '"/></noscript>',
          '</body></html>'
        ].join('');
      }

      return HtmlService.createHtmlOutput(htmlAuth)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    } catch (err) {
      L('auth build ERROR: ' + err);
      return HtmlService.createHtmlOutput(
        '<pre>Falhou a construir o URL de OAuth:\n' + String(err) + '\n' + L.dump().join('\n') + '</pre>'
      ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    */
  }


  // 3) CALLBACK OAuth — TRATAR O CODE **ANTES** DO "no active user"
  L('3) CALLBACK OAuth?');
  var t = e && e.parameter && e.parameter.ticket;
  var code = e && e.parameter && e.parameter.code;
  var state = e && e.parameter && e.parameter.state;

  if (code) {
    L('oauth callback: code present');

    var st = parseStateToken_(state);
    if (!st.ok) {
      L('state invalid/expired');
      return HtmlService.createHtmlOutput('<pre>O token de estado é inválido/expirou.\n' + L.dump().join('\n') + '</pre>')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    var dbgFromState = !!(st.payload && st.payload.dbg);
    var embedFromState = !!(st.payload && st.payload.embed);
    var stateNonce = (st.payload && st.payload.nonce) || '';

    try {
      var tokResp = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
        method: 'post',
        contentType: 'application/x-www-form-urlencoded',
        payload: {
          code: code,
          client_id: getClientId_(),
          client_secret: getClientSecret_(),
          redirect_uri: redirectUri_(), // EXACTAMENTE o registado no OAuth client
          grant_type: 'authorization_code'
        },
        muteHttpExceptions: true
      });
      L('token resp code=' + tokResp.getResponseCode());
      if (tokResp.getResponseCode() !== 200) {
        var txt = tokResp.getContentText() || '';
        if (txt.indexOf('"invalid_grant"') !== -1) {
          var back = canonicalAppUrl_() + '?action=auth' + (DBG ? '&debug=1' : '');
          return HtmlService.createHtmlOutput('<script>location.replace(' + JSON.stringify(back) + ');</script>')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        }
        return HtmlService.createHtmlOutput('<pre>Falha ao obter tokens: ' + txt + '\n' + L.dump().join('\n') + '</pre>')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      var tok = JSON.parse(tokResp.getContentText());
      var idToken = tok.id_token;
      if (!idToken) {
        return HtmlService.createHtmlOutput('<pre>Sem id_token.\n' + L.dump().join('\n') + '</pre>')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      var v = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
      L('tokeninfo code=' + v.getResponseCode());
      if (v.getResponseCode() !== 200)
        return HtmlService.createHtmlOutput('<pre>ID token inválido.\n' + L.dump().join('\n') + '</pre>')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      var info = JSON.parse(v.getContentText());
      if (info.aud !== getClientId_())
        return HtmlService.createHtmlOutput('<pre>aud inválido.\n' + L.dump().join('\n') + '</pre>')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      if (info.iss !== 'https://accounts.google.com' && info.iss !== 'accounts.google.com')
        return HtmlService.createHtmlOutput('<pre>iss inválido.\n' + L.dump().join('\n') + '</pre>')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      if (info.email_verified !== 'true')
        return HtmlService.createHtmlOutput('<pre>email não verificado.\n' + L.dump().join('\n') + '</pre>')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      var email = info.email;
      var nomeT = info.name || info.fullName || [info.given_name, info.family_name].filter(Boolean).join(' ') || '';

      if (!findDuesPaidByEmail(email, nomeT)) {
        return HtmlService.createHtmlOutput('<pre>Acesso negado. Email não registado ou quotas em atraso.\n' + L.dump().join('\n') + '</pre>')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      // Sessão via ticket
      // var ticketNew = Utilities.getUuid();
      // CacheService.getScriptCache().put('sess:' + ticketNew, JSON.stringify({ email: email, ts: Date.now() }), SESSION_TTL_SEC);
      var ticketNew = issueSessionToken_(email, 14); // 14 dias de sessão
      /*
      var nonce = st.payload && st.payload.nonce;
      if (nonce) {
        putTicketForNonce_(nonce, ticketNew);
        L('stored ticket for nonce=' + nonce);
      }
      */
      putTicketForNonce_(stateNonce, ticketNew);
      L('stored ticket for stateNonce=' + stateNonce);

      var next = canonicalAppUrl_() + '?ticket=' + encodeURIComponent(ticketNew) + (dbgFromState ? '&debug=1' : '');
      L('redirecting back to app …');

      var logs = DBG ? ('<pre style="white-space:pre-wrap;">' + L.dump().join('\n') + '</pre>') : '';

      // … após obter ticketNew, dbgFromState/embedFromState/next …
      var htmlBack = [
        '<!doctype html><html><head><meta charset="utf-8"><title>Auth OK</title>',
        '<style>body{font:14px system-ui,sans-serif;padding:16px}#log{white-space:pre-wrap;background:#111;color:#0f0;padding:8px;margin-top:12px;font:12px/1.35 monospace}</style>',
        '</head><body>', (dbgFromState ? '<div id="log"></div>' : ''), '<script>(function(){',
        'function log(m){try{console.log("[AUTH]",m)}catch(_){ } var el=document.getElementById("log"); if(el){el.textContent+=(new Date()).toISOString().slice(11,19)+" "+m+"\\n";}}',
        'var T=', JSON.stringify(ticketNew), ';',
        'var NEXT=', JSON.stringify(next), ';',
        'try{ localStorage.setItem("sessTicket", T); log("set localStorage"); }catch(e){ log("ls ERR "+e); }',
        'try{ document.cookie="sessTicket="+encodeURIComponent(T)+"; Max-Age=', (14*24*3600), '; Path=/; Secure; SameSite=Lax"; log("cookie set"); }catch(e){ log("cookie ERR "+e); }',
        '// 1) avisar opener (janela que abriu o popup)',
        'try{ if (window.opener) { window.opener.postMessage({type:"portobelo_ticket", ticket:T}, "*"); log("postMessage → opener ok"); } else { log("no opener"); } }catch(e){ log("postMessage ERR "+e); }',
        '// 2) pequeno atraso para o opener reagir e depois FECHAR de forma teimosa',
        'setTimeout(function(){',
        '  log("try close()");',
        '  try{ window.close(); log("close() called"); }catch(e){ log("close() ERR "+e); }',
        '  // Fallbacks (alguns browsers só deixam fechar se reabrirmos a nós próprios)',
        '  try{ window.open("","_self"); window.close(); log("fallback self-close"); }catch(e){ log("fallback self-close ERR "+e); }',
        '  // Último recurso: navegar para a app (o popup fica com a app, mas o utilizador consegue fechá-lo)',
        '  try{ location.replace(NEXT); }catch(_){ location.href=NEXT; }',
        '}, 400);',
        '})();</script></body></html>'
      ].join('');

      return HtmlService.createHtmlOutput(htmlBack)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);


    } catch (err) {
      L('callback ERROR: ' + err);
      return HtmlService.createHtmlOutput('<pre>Erro no callback:\n' + String(err) + '\n' + L.dump().join('\n') + '</pre>')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }

  // 4) (só agora) Sem ticket → sem sessão → vai DIRETO ao Google (sem página intermédia)
  L('4) no ticket sem sessão?');
  var user = '';
  try { user = Session.getActiveUser().getEmail() || ''; } catch (_) { }
  // L('user=' + user); //é sempre admin@titulares-portobelo.pt, que é o owner da Apps Script
  if (!user) {
    L('!user → render Login (sem auto-redirect)');
    /*
    var url = buildAuthUrl_(DBG);
    L('url=' + url);
    var htmlHop = DBG ? [
      '<!doctype html><meta charset="utf-8"><title>Autenticação necessária</title>',
      '<pre style="white-space:pre-wrap;">', L.dump().join('\n'), '</pre>',
      '<a id="jump" href="', url.replace(/"/g, '&quot;'), '" target="_top" rel="noopener">Ir para autenticação</a>',
      '<script>(function(){try{document.getElementById("jump").click();}catch(e){location.href=',
      JSON.stringify(url), ';}})();</script>',
      '<noscript><meta http-equiv="refresh" content="0; url=', url.replace(/"/g, '&quot;'), '"/></noscript>'
    ].join('') : [
      '<!doctype html><html><head><meta charset="utf-8">',
      '<meta http-equiv="refresh" content="0; url=', url.replace(/"/g, '&quot;'), '"/>',
      '</head><body>',
      '<a id="jump" href="', url.replace(/"/g, '&quot;'),
      '" target="_top" rel="noopener" style="position:fixed;left:-9999px;top:-9999px;width:1px;height:1px;opacity:0">_</a>',
      '<script>(function(){var a=document.getElementById("jump");',
      'try{if(a&&a.click)a.click();else location.href=', JSON.stringify(url), ';}',
      'catch(_){location.href=', JSON.stringify(url), ';}})();</script>',
      '<noscript><meta http-equiv="refresh" content="0; url=', url.replace(/"/g, '&quot;'), '"/></noscript>',
      '</body></html>'
    ].join('');
    return HtmlService.createHtmlOutput(htmlHop)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    */
    return renderLoginPage_(DBG, L.dump(), WIPE);
  }

  // 5) Sem ticket mas com sessão → mostra Login (botão entra em ?action=auth)
  L('5) no ticket mas com sessão → render Login');
  return renderLoginPage_(DBG, L.dump(), WIPE);
}





// ---------- Login com GIS One Tap (ID Token) ----------
function loginWithIdToken(idToken) {
  if (!idToken) throw new Error('Sem credential.');

  const resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true }
  );
  if (resp.getResponseCode() !== 200) throw new Error('Token inválido.');
  const info = JSON.parse(resp.getContentText());

  if (info.aud !== getClientId_()) throw new Error('aud inválido.');
  if (info.iss !== 'https://accounts.google.com' && info.iss !== 'accounts.google.com') throw new Error('iss inválido.');
  if (info.email_verified !== 'true') throw new Error('email não verificado.');

  const email = info.email;
  const nomeT = info.name || info.fullName || [info.given_name, info.family_name].filter(Boolean).join(' ') || '';

  if (!findDuesPaidByEmail(email, nomeT)) throw new Error('Acesso negado. Email não registado ou quotas em atraso.');

  // const ticket = Utilities.getUuid();
  // CacheService.getScriptCache().put('sess:' + ticket, JSON.stringify({ email: email, ts: Date.now() }), SESSION_TTL_SEC);
  // 👉 passa a HMAC “stateless”
  const ticket = issueSessionToken_(email, 14);


  return { ticket: ticket, appUrl: canonicalAppUrl_() };
}


// ===== Utilitários, regras, endpoints protegidos =====

function doPost(e) {
  if (e.parameter && e.parameter.route === 'login') return handleGoogleLogin(e);
  return HtmlService.createHtmlOutput('Rota inválida.');
}

// Recebe credential (ID Token), valida e cria sessão
function handleGoogleLogin(e) {
  try {
    const idToken = e.parameter && e.parameter.credential;
    if (!idToken) return HtmlService.createHtmlOutput('Faltou credential.');

    const resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return HtmlService.createHtmlOutput('Token inválido.');
    const info = JSON.parse(resp.getContentText());

    if (info.aud !== getClientId_()) return HtmlService.createHtmlOutput('aud inválido.');
    if (info.iss !== 'https://accounts.google.com' && info.iss !== 'accounts.google.com') return HtmlService.createHtmlOutput('iss inválido.');
    if (info.email_verified !== 'true') return HtmlService.createHtmlOutput('email não verificado.');

    const email = info.email;

    // 👉 passa a HMAC “stateless”
    const ticket = issueSessionToken_(email, 14);
    const redirect = canonicalAppUrl_() + '?ticket=' + encodeURIComponent(ticket);

    // opcional: set cookie (útil para recuperar em nova aba)
    const html = [
      '<script>',
      'try{ localStorage.setItem("sessTicket", ', JSON.stringify(ticket), '); }catch(_){ }',
      'document.cookie="sessTicket=' + encodeURIComponent(ticket) + '; Max-Age=' + (14 * 24 * 3600) + '; Path=/; Secure; SameSite=Lax";',
      'window.top.location.href=', JSON.stringify(redirect), ';',
      '</script>'
    ].join('');
    return HtmlService.createHtmlOutput(html);

  } catch (err) {
    return HtmlService.createHtmlOutput('Erro no login.');
  }
}


// ---------- Endpoints protegidos ----------
function getAnunciosDataSess(ticket) { var sess = requireSession(ticket); return getAnunciosData(sess.email); }

function submitAnuncioSess(ticket, anuncioParcial) { var sess = requireSession(ticket); /* validações */ return submitAnuncio(anuncioParcial); }


// ---------- Regras de negócio / utilitários ----------
function incrementVisitCounter() {
  const props = PropertiesService.getScriptProperties();
  const count = parseInt(props.getProperty('visitCount') || '0', 10) + 1;
  props.setProperty('visitCount', count.toString());
  return count;
}

function getUltimaAtualizacao() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ANUNCIOS_SHEET);
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




function findTelefoneIndex(numero) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const titulares = ss.getSheetByName(TITULARES_SHEET);
  const data = titulares.getRange('K2:K').getValues();

  const cleanedNumero = numero.toString().replace(/\D/g, '');
  if (LOGGING) Logger.log('numero=' + numero + ', cleanedNumero=' + cleanedNumero);

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

function findDuesPaid(index, email) {
  if (index < 0) return false;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TITULARES_SHEET);

  const row = index + 2;
  const emailCell = sheet.getRange(row, 12).getValue(); // L
  const duesCell = sheet.getRange(row, 3).getValue();  // C

  const allEmails = emailCell.toString().split(';').map(e => e.trim().toLowerCase());
  return allEmails.includes((email || '').toLowerCase()) && (Number(duesCell) || 0) >= 1;
}

function getAnunciosData(emailAutenticado) {
  if (LOGGING) Logger.log('Função getAnunciosData chamada com sucesso.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ANUNCIOS_SHEET);
  if (!sheet) {
    const msg = "Erro interno: aba '" + ANUNCIOS_SHEET + "' não encontrada.";
    if (LOGGING) Logger.log(msg);
    throw new Error(msg);
  }

  const lastRow = sheet.getLastRow();
  const numRows = lastRow - ANUNCIOS_HEADER_ROW;
  if (numRows <= 0) return [];

  const raw = sheet.getRange(ANUNCIOS_HEADER_ROW + 1, 1, numRows, 5).getValues();
  const data = raw.map(row => row.map(function (cell) {
    if (typeof cell === 'string') return cell.replace(/[\r\n]+/g, ' ').trim();
    return (cell !== undefined && cell !== null) ? cell.toString() : '';
  }));
  return data;
}

function submitAnuncio(anuncioParcial) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PEDIDOS_SHEET) || ss.insertSheet(PEDIDOS_SHEET);

  if (!Array.isArray(anuncioParcial) || anuncioParcial.length !== 5) throw new Error('Dados inválidos');

  const ultimaLinha = sheet.getLastRow();
  const limiteLinhas = 50;
  if (ultimaLinha >= limiteLinhas) {
    throw new Error('Limite de ' + limiteLinhas + ' anúncios atingido. Não é possível adicionar mais anúncios.');
  }

  const dataAtual = new Date();
  const anuncioCompleto = [dataAtual, ...anuncioParcial.slice(1)];
  sheet.appendRow(anuncioCompleto);

  SpreadsheetApp.flush();

  const email = 'geral@titulares-portobelo.pt';
  const assunto = 'Novo anúncio do Portobelo submetido: ' + anuncioParcial[2];
  const corpo = 'Um novo anúncio foi submetido na lista do Portobelo:\n' +
    'Telemóvel=' + anuncioParcial[2] + ', Tipo=' + anuncioParcial[1] + '\n' + anuncioParcial[3];
  MailApp.sendEmail(email, assunto, corpo);

  return true;
}


function testFind() {
  const result = findDuesPaidByEmail("pal.mendes23@gmail2.com", "Pedro Mendes");
  Logger.log(result);
}
