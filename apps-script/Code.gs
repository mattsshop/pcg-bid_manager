/**
 * PCG Bid Manager — Google Apps Script back end
 * Version 3.0.0 — 2026-11-17
 *
 * Two files in the Apps Script project:
 *   Code.gs     (this file)  — saves data to a Google Sheet and sends email
 *   Index.html               — the whole app (UI)
 *
 * Changelog
 *   3.0.0  Password-protected JSON API (doPost) so the app page can be hosted on your own website
 *   2.4.0  (UI only) project dropdown, rename/add/remove projects, project filter
 *   2.3.0  Optional sending through Microsoft 365 (Graph API) as your Outlook address
 *   2.2.0  Formatted (HTML) emails from the app, logo embedded inline under the signature
 *   2.1.0  Budget table (project cost codes + amounts); packages store CSI cost codes
 *   2.0.0  Rebuilt as a standalone web app. No sheet menus, no sub-facing
 *          pages, no Drive folders. Bid board, sub directory, email invites
 *          and reminders sent from Gmail with Reply-To your PCG address.
 */

const APP_VERSION = '3.0.0';

const TABLES = {
  Subs:     ['id', 'company', 'trades', 'contact', 'email', 'phone', 'notes', 'active', 'created'],
  Packages: ['id', 'project', 'name', 'trades', 'scope', 'due', 'plansLink', 'status', 'notes', 'created', 'costCodes'],
  Invites:  ['id', 'packageId', 'subId', 'company', 'contact', 'email', 'status', 'sentAt', 'remindedAt', 'reminders', 'amount', 'receivedAt', 'notes'],
  Budget:   ['id', 'project', 'code', 'description', 'amount'],
  Settings: ['key', 'value']
};

/**
 * RUN THIS ONCE from the editor (it is first in the dropdown).
 * Grants the two permissions the app needs: Google Sheets and send email.
 */
function authorize() {
  const ss = db_();
  Object.keys(TABLES).forEach(sheet_);
  Logger.log('Ready. Data sheet: ' + ss.getUrl());
  Logger.log('Emails left today: ' + MailApp.getRemainingDailyQuota());
  Logger.log('Microsoft 365 redirect URI (for Settings → Send email from): ' + msRedirectUri_());
  Logger.log('Next: Deploy > Test deployments > copy the Web app URL and bookmark it. (v' + APP_VERSION + ')');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('PCG Bid Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ===================== WEBSITE API (password protected) =====================
 * The copy of the app hosted on your website talks to this script through
 * doPost. Every call except "login" needs a valid sign-in token, and the
 * password is checked here on Google's side.
 *
 * Set or change the password: Apps Script editor → Project Settings (gear) →
 * Script Properties → add or edit APP_PASSWORD. On the next sign-in it is
 * replaced with a scrambled (hashed) copy, and everyone is signed out.
 */

const API_FUNCTIONS = {
  apiLoad: apiLoad, apiSave: apiSave, apiDelete: apiDelete, apiSaveSettings: apiSaveSettings, apiSend: apiSend,
  apiMsStatus: apiMsStatus, apiMsSaveConfig: apiMsSaveConfig, apiMsDisconnect: apiMsDisconnect, apiMsTest: apiMsTest
};
const TOKEN_DAYS = 30;

function doPost(e) {
  let req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, error: 'Bad request.' }); }
  try {
    if (req.fn === 'login') return json_({ ok: true, result: login_(String(req.password || ''), !!req.remember) });
    if (!verifyToken_(req.token)) return json_({ ok: false, auth: true, error: 'Please sign in again.' });
    const fn = API_FUNCTIONS[req.fn];
    if (!fn) return json_({ ok: false, error: 'Unknown action: ' + req.fn });
    return json_({ ok: true, result: fn.apply(null, Array.isArray(req.args) ? req.args : []) });
  } catch (err) {
    return json_({ ok: false, error: err && err.message ? err.message : String(err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function login_(password, remember) {
  const props = PropertiesService.getScriptProperties();
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('LOGIN_FAILS') || 0);
  if (fails >= 10) throw new Error('Too many wrong passwords. Wait 15 minutes and try again.');

  // A newly typed APP_PASSWORD is hashed on first use; changing it signs everyone out
  const plain = props.getProperty('APP_PASSWORD');
  if (plain) {
    const salt = Utilities.getUuid();
    props.setProperty('APP_PASSWORD_SALT', salt);
    props.setProperty('APP_PASSWORD_HASH', hash_(salt + plain));
    props.setProperty('TOKEN_SECRET', Utilities.getUuid() + Utilities.getUuid());
    props.deleteProperty('APP_PASSWORD');
  }
  const salt = props.getProperty('APP_PASSWORD_SALT'), stored = props.getProperty('APP_PASSWORD_HASH');
  if (!stored) throw new Error('No password is set yet. In the Apps Script editor, open Project Settings → Script Properties and add APP_PASSWORD.');
  if (!password || !sameString_(hash_(salt + password), stored)) {
    cache.put('LOGIN_FAILS', String(fails + 1), 900);
    Utilities.sleep(1200);
    throw new Error('Wrong password.');
  }
  cache.remove('LOGIN_FAILS');
  const days = remember ? TOKEN_DAYS : 1;
  const body = Utilities.base64EncodeWebSafe(JSON.stringify({ exp: Date.now() + days * 864e5, n: Utilities.getUuid().slice(0, 8) }));
  return { token: body + '.' + sign_(body), days: days, version: APP_VERSION };
}

function verifyToken_(token) {
  const parts = String(token || '').split('.');
  if (parts.length !== 2 || !sameString_(sign_(parts[0]), parts[1])) return false;
  try {
    const data = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
    return Number(data.exp) > Date.now();
  } catch (e) { return false; }
}

function sign_(text) {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('TOKEN_SECRET');
  if (!secret) { secret = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('TOKEN_SECRET', secret); }
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(text, secret)).replace(/=+$/, '');
}

function hash_(text) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8));
}

function sameString_(a, b) {
  a = String(a); b = String(b);
  let diff = a.length ^ b.length;
  for (let i = 0; i < Math.max(a.length, b.length); i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

/* ============================== API ============================== */

function apiLoad() {
  return {
    version: APP_VERSION,
    subs: readAll_('Subs'),
    packages: readAll_('Packages'),
    invites: readAll_('Invites'),
    budget: readAll_('Budget'),
    settings: settings_(),
    quota: MailApp.getRemainingDailyQuota(),
    ms: safeMsStatus_(),
    dbUrl: db_().getUrl()
  };
}

function apiSave(table, records) {
  checkTable_(table);
  if (!Array.isArray(records) || !records.length) return true;
  withLock_(() => upsert_(table, records));
  return true;
}

function apiDelete(table, ids) {
  checkTable_(table);
  const set = new Set((ids || []).map(String));
  if (!set.size) return true;
  withLock_(() => {
    if (table === 'Packages') {
      const inv = readAll_('Invites').filter(i => set.has(String(i.packageId))).map(i => i.id);
      deleteIds_('Invites', new Set(inv));
    }
    deleteIds_(table, set);
  });
  return true;
}

function apiSaveSettings(obj) {
  const rows = Object.keys(obj || {}).map(k => ({ key: k, value: obj[k] == null ? '' : String(obj[k]) }));
  if (rows.length) withLock_(() => upsert_('Settings', rows));
  return true;
}

/**
 * messages: [{inviteId, to, subject, body}]  (already filled in by the app)
 * kind: 'invite' | 'reminder'
 */
function apiSend(messages, kind) {
  if (!Array.isArray(messages) || !messages.length) throw new Error('No recipients selected.');
  if (kind !== 'invite' && kind !== 'reminder') throw new Error('Unknown email type.');
  const s = settings_();
  const copy = isEmail_(s.copyTo) ? String(s.copyTo).trim() : '';
  const viaMs = s.sendVia === 'microsoft';
  if (viaMs) {
    msAccessToken_(); // fails fast, before anything is sent, if Outlook isn't connected
  } else {
    const need = messages.length * (copy ? 2 : 1);
    const quota = MailApp.getRemainingDailyQuota();
    if (quota < need) throw new Error('Only ' + quota + ' emails left today and this needs ' + need + '. Send fewer, or try again tomorrow.');
  }

  const logo = logoBlob_(s.logo);
  const pace = viaMs && messages.length > 25; // Exchange allows about 30 messages a minute

  return withLock_(() => {
    const byId = {};
    readAll_('Invites').forEach(i => { byId[i.id] = i; });
    const now = new Date().toISOString();
    const sent = [], failed = [], updated = [];

    messages.forEach(m => {
      const inv = byId[m.inviteId];
      if (!inv) { failed.push({ id: m.inviteId, error: 'Bidder not found' }); return; }
      const to = String(m.to || '').trim();
      if (!isEmail_(to)) { failed.push({ id: inv.id, error: inv.company + ': bad email "' + to + '"' }); return; }
      try {
        const opts = {
          to: to,
          subject: String(m.subject || '').slice(0, 250),
          body: String(m.body || ''),
          htmlBody: m.html ? String(m.html) : textToHtml_(m.body)
        };
        if (opts.htmlBody.indexOf('cid:pcglogo') >= 0) {
          if (logo) opts.inlineImages = { pcglogo: logo };
          else opts.htmlBody = opts.htmlBody.replace(/<img[^>]*cid:pcglogo[^>]*>/g, '');
        }
        if (viaMs) {
          msSendMail_(to, opts.subject, opts.htmlBody, copy, s.replyTo, opts.htmlBody.indexOf('cid:pcglogo') >= 0 ? logo : null);
          if (pace) Utilities.sleep(2100);
        } else {
          const name = s.senderName || s.companyName;
          if (name) opts.name = name;
          if (isEmail_(s.replyTo)) opts.replyTo = String(s.replyTo).trim();
          if (copy) opts.bcc = copy;
          MailApp.sendEmail(opts);
        }

        if (kind === 'invite') {
          if (!inv.status || inv.status === 'Not Sent') inv.status = 'Invited';
          inv.sentAt = now;
        } else {
          inv.remindedAt = now;
          inv.reminders = String((Number(inv.reminders) || 0) + 1);
        }
        updated.push(inv);
        sent.push(inv.id);
      } catch (e) {
        failed.push({ id: inv.id, error: inv.company + ': ' + e.message });
      }
    });

    if (updated.length) upsert_('Invites', updated);
    return { sent: sent, failed: failed, invites: updated, quota: MailApp.getRemainingDailyQuota() };
  });
}

/* ======================= MICROSOFT 365 (GRAPH) ======================= */
// Sends as your Outlook mailbox through Microsoft Graph. Setup steps are in Settings and the README.
// The client secret and sign-in tokens live in this script's private properties, not in the data sheet.

const MS_SCOPES = 'offline_access https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/User.Read';

function apiMsStatus() { return msStatus_(); }

function apiMsSaveConfig(cfg) {
  cfg = cfg || {};
  const tenant = String(cfg.tenant || '').trim();
  const clientId = String(cfg.clientId || '').trim();
  const secret = String(cfg.secret || '').trim();
  const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  if (!guid.test(tenant) && !/^[a-z0-9.-]+\.[a-z]{2,}$/i.test(tenant)) throw new Error('The Directory (tenant) ID should look like 1a2b3c4d-....');
  if (!guid.test(clientId)) throw new Error('The Application (client) ID should look like 1a2b3c4d-....');
  if (guid.test(secret)) throw new Error('That looks like the secret\'s ID. Copy the secret\'s Value column instead.');
  const p = PropertiesService.getScriptProperties();
  const old = msCfg_();
  if (old.tenant !== tenant || old.clientId !== clientId) PropertiesService.getUserProperties().deleteProperty('MS_TOKENS');
  p.setProperty('MS_TENANT', tenant);
  p.setProperty('MS_CLIENT_ID', clientId);
  if (secret) p.setProperty('MS_CLIENT_SECRET', secret);
  return msStatus_();
}

function apiMsDisconnect() {
  PropertiesService.getUserProperties().deleteProperty('MS_TOKENS');
  return msStatus_();
}

function apiMsTest() {
  const t = msTokens_();
  if (!t || !t.email) throw new Error('Outlook isn\'t connected yet.');
  const s = settings_();
  msSendMail_(t.email, 'PCG Bid Manager test email',
    textToHtml_('This test was sent from the PCG Bid Manager through your Microsoft 365 account.\n\nIf you are reading this in Outlook, sending works, and a copy is in your Sent Items.'),
    '', '', null);
  return t.email;
}

/** Microsoft sends you back here after you sign in (the "usercallback" redirect URI). */
function msAuthCallback(request) {
  const prm = (request && request.parameter) || {};
  if (prm.error) return msPage_('Microsoft sign-in didn\'t finish', msErr_({ error: prm.error, error_description: prm.error_description }), false);
  try {
    const tok = msTokenRequest_({ grant_type: 'authorization_code', code: prm.code, redirect_uri: msRedirectUri_(), scope: MS_SCOPES });
    const res = UrlFetchApp.fetch('https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName',
      { headers: { Authorization: 'Bearer ' + tok.access_token }, muteHttpExceptions: true });
    const me = JSON.parse(res.getContentText() || '{}');
    tok.email = me.mail || me.userPrincipalName || '';
    tok.name = me.displayName || '';
    msSaveTokens_(tok);
    return msPage_('Outlook connected', 'The bid manager will now send email as ' + tok.email + '. Close this tab and click "Refresh status" in the bid manager.', true);
  } catch (e) {
    return msPage_('Couldn\'t connect Outlook', e.message, false);
  }
}

function safeMsStatus_() {
  try { return msStatus_(); } catch (e) { return { error: e.message }; }
}

function msStatus_() {
  const c = msCfg_(), t = msTokens_();
  const st = {
    tenant: c.tenant, clientId: c.clientId, hasSecret: !!c.secret, redirectUri: msRedirectUri_(),
    configured: !!(c.tenant && c.clientId && c.secret),
    connected: !!(t && t.refresh_token), email: (t && t.email) || '', name: (t && t.name) || '', authUrl: ''
  };
  if (st.configured) {
    const state = ScriptApp.newStateToken().withMethod('msAuthCallback').withTimeout(3600).createToken();
    st.authUrl = 'https://login.microsoftonline.com/' + encodeURIComponent(c.tenant) + '/oauth2/v2.0/authorize?' + query_({
      client_id: c.clientId, response_type: 'code', redirect_uri: msRedirectUri_(), response_mode: 'query',
      scope: MS_SCOPES, state: state, prompt: 'select_account'
    });
  }
  return st;
}

function msCfg_() {
  const p = PropertiesService.getScriptProperties();
  return { tenant: p.getProperty('MS_TENANT') || '', clientId: p.getProperty('MS_CLIENT_ID') || '', secret: p.getProperty('MS_CLIENT_SECRET') || '' };
}
function msRedirectUri_() { return 'https://script.google.com/macros/d/' + ScriptApp.getScriptId() + '/usercallback'; }
function msTokens_() {
  const raw = PropertiesService.getUserProperties().getProperty('MS_TOKENS');
  try { return raw ? JSON.parse(raw) : null; } catch (e) { return null; }
}
function msSaveTokens_(t) {
  PropertiesService.getUserProperties().setProperty('MS_TOKENS', JSON.stringify({
    access_token: t.access_token, refresh_token: t.refresh_token, expires_at: t.expires_at, email: t.email || '', name: t.name || ''
  }));
}

function msTokenRequest_(params) {
  const c = msCfg_();
  if (!c.tenant || !c.clientId || !c.secret) throw new Error('Microsoft 365 settings are incomplete. Fill them in under Settings.');
  const res = UrlFetchApp.fetch('https://login.microsoftonline.com/' + encodeURIComponent(c.tenant) + '/oauth2/v2.0/token', {
    method: 'post', muteHttpExceptions: true,
    payload: Object.assign({ client_id: c.clientId, client_secret: c.secret }, params)
  });
  let j = {};
  try { j = JSON.parse(res.getContentText() || '{}'); } catch (e) { /* not JSON */ }
  if (res.getResponseCode() !== 200 || !j.access_token) throw new Error(msErr_(j));
  j.expires_at = Date.now() + (Number(j.expires_in) || 3600) * 1000;
  return j;
}

function msAccessToken_(force) {
  const t = msTokens_();
  if (!t || !t.refresh_token) throw new Error('Outlook isn\'t connected. Go to Settings → Send email from → Connect Outlook, or switch back to Google.');
  if (!force && t.access_token && t.expires_at > Date.now() + 120000) return t.access_token;
  const n = msTokenRequest_({ grant_type: 'refresh_token', refresh_token: t.refresh_token, scope: MS_SCOPES });
  const merged = { access_token: n.access_token, refresh_token: n.refresh_token || t.refresh_token, expires_at: n.expires_at, email: t.email, name: t.name };
  msSaveTokens_(merged);
  return merged.access_token;
}

function msSendMail_(to, subject, html, bcc, replyTo, logo) {
  const t = msTokens_() || {};
  const message = {
    subject: subject,
    body: { contentType: 'HTML', content: html },
    toRecipients: [{ emailAddress: { address: to } }]
  };
  if (isEmail_(bcc)) message.bccRecipients = [{ emailAddress: { address: String(bcc).trim() } }];
  if (isEmail_(replyTo) && String(replyTo).trim().toLowerCase() !== String(t.email || '').toLowerCase()) {
    message.replyTo = [{ emailAddress: { address: String(replyTo).trim() } }];
  }
  if (logo) {
    message.attachments = [{
      '@odata.type': '#microsoft.graph.fileAttachment', name: logo.getName(), contentType: logo.getContentType(),
      contentBytes: Utilities.base64Encode(logo.getBytes()), isInline: true, contentId: 'pcglogo'
    }];
  } else {
    message.body.content = html.replace(/<img[^>]*cid:pcglogo[^>]*>/g, '');
  }
  const payload = JSON.stringify({ message: message, saveToSentItems: true });
  let token = msAccessToken_();
  for (let attempt = 0; attempt < 3; attempt++) {
    const res = UrlFetchApp.fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
      method: 'post', contentType: 'application/json', payload: payload, muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + token }
    });
    const code = res.getResponseCode();
    if (code === 202 || code === 200) return;
    if (code === 401 && attempt === 0) { token = msAccessToken_(true); continue; }
    if (code === 429 || code === 503) {
      const wait = Number((res.getHeaders() || {})['Retry-After']) || 10;
      Utilities.sleep(Math.min(wait, 30) * 1000);
      continue;
    }
    let msg = 'Outlook error ' + code;
    try { const j = JSON.parse(res.getContentText()); if (j.error && j.error.message) msg += ': ' + j.error.message; } catch (e) { /* ignore */ }
    throw new Error(msg);
  }
  throw new Error('Outlook is busy (rate limited). Try again in a minute.');
}

function msErr_(j) {
  const d = String((j && (j.error_description || j.error)) || 'Unknown error').split(/\r?\n/)[0];
  if (/AADSTS7000222/.test(d)) return 'The client secret has expired. Create a new one in Microsoft Entra and paste it into Settings. (' + d + ')';
  if (/AADSTS7000215/.test(d)) return 'The client secret is wrong. Copy the secret\'s Value (not its ID) into Settings. (' + d + ')';
  if (/AADSTS50011/.test(d)) return 'The redirect URI in Microsoft Entra doesn\'t match. Copy it exactly from Settings. (' + d + ')';
  if (/AADSTS65001|AADSTS90094|consent/i.test(d)) return 'Your Microsoft 365 admin needs to approve this app: Entra → App registrations → PCG Bid Manager → API permissions → Grant admin consent. (' + d + ')';
  if (/invalid_grant|AADSTS70008|AADSTS700082|AADSTS50173/.test(d)) return 'The Outlook connection expired. Go to Settings and click Connect Outlook again. (' + d + ')';
  return d;
}

function msPage_(title, msg, ok) {
  return HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;max-width:520px;margin:60px auto;padding:24px;border:1px solid #e3e7ee;border-radius:12px">' +
    '<h2 style="margin-top:0;color:' + (ok ? '#1e7b34' : '#b42318') + '">' + (ok ? '✓ ' : '') + escHtml_(title) + '</h2>' +
    '<p style="font-size:15px;line-height:1.5">' + escHtml_(msg) + '</p></div>').setTitle(title);
}

function query_(o) { return Object.keys(o).map(k => encodeURIComponent(k) + '=' + encodeURIComponent(o[k])).join('&'); }
function escHtml_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
}

/* ============================ STORAGE ============================ */

function db_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('DB_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* recreate below */ }
  }
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { /* standalone script */ }
  if (!ss) ss = SpreadsheetApp.create('PCG Bid Manager – Data');
  props.setProperty('DB_ID', ss.getId());
  return ss;
}

function sheet_(name) {
  const ss = db_();
  let sh = ss.getSheetByName(name);
  const headers = TABLES[name];
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    const blank = ss.getSheetByName('Sheet1');
    if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  } else {
    // Add any columns introduced by a newer version
    const have = headers_(sh);
    const missing = headers.filter(h => have.indexOf(h) < 0);
    if (missing.length) {
      sh.getRange(1, have.length + 1, 1, missing.length).setValues([missing])
        .setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
    }
  }
  return sh;
}

function headers_(sh) {
  const n = sh.getLastColumn();
  return n ? sh.getRange(1, 1, 1, n).getValues()[0].map(h => String(h).trim()).filter(Boolean) : [];
}

function readAll_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const headers = headers_(sh);
  const tz = Session.getScriptTimeZone();
  return sh.getRange(2, 1, last - 1, headers.length).getValues()
    .filter(r => r[0] !== '' && r[0] != null)
    .map(r => {
      const o = {};
      headers.forEach((h, i) => {
        let v = r[i];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm");
        o[h] = v == null ? '' : String(v);
      });
      return o;
    });
}

function upsert_(name, records) {
  const sh = sheet_(name);
  const headers = headers_(sh);
  const last = sh.getLastRow();
  const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
  const appends = [];
  records.forEach(o => {
    const row = headers.map(h => (o[h] == null ? '' : String(o[h])));
    const i = ids.indexOf(String(o[headers[0]]));
    if (i >= 0) writeRows_(sh, i + 2, [row]);
    else { appends.push(row); ids.push(String(o[headers[0]])); }
  });
  if (appends.length) writeRows_(sh, last + 1, appends);
}

function writeRows_(sh, start, rows) {
  const end = start + rows.length - 1;
  if (end > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), end - sh.getMaxRows() + 100);
  sh.getRange(start, 1, rows.length, rows[0].length).setNumberFormat('@').setValues(rows);
}

function deleteIds_(name, idSet) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return;
  const col = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = col.length - 1; i >= 0; i--) {
    if (idSet.has(String(col[i][0]))) sh.deleteRow(i + 2);
  }
}

function settings_() {
  const o = {};
  readAll_('Settings').forEach(r => { o[r.key] = r.value; });
  return o;
}

function checkTable_(t) {
  if (['Subs', 'Packages', 'Invites', 'Budget'].indexOf(t) < 0) throw new Error('Unknown table: ' + t);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ============================ HELPERS ============================ */

function isEmail_(e) {
  return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(String(e || '').trim());
}

function logoBlob_(dataUrl) {
  const m = String(dataUrl || '').match(/^data:(image\/[\w.+-]+);base64,(.+)$/);
  if (!m) return null;
  const ext = m[1].split('/')[1].replace('jpeg', 'jpg').replace(/\+.*/, '');
  return Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'logo.' + ext);
}

function textToHtml_(text) {
  const esc = String(text || '').replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const linked = esc.replace(/https?:\/\/[^\s<]+/g, u => '<a href="' + u + '">' + u + '</a>');
  return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.5;color:#222">' +
    linked.replace(/\r?\n/g, '<br>') + '</div>';
}
