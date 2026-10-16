/***** EffeOre — registro attività staff Effestudio *****/
const TZ = 'Europe/Rome';
const SH = { staff: 'Staff', rates: 'Tariffe', log: 'Registro', cfg: 'Config', fixed: 'Voci fisse', extra: 'Voci extra', access: 'Accessi' };
const ACT_ORDER = ['Classi', 'Ore', 'PT', 'Coaching', 'Schede', 'Pulizia'];
const CLIENT_ACTS = ['PT', 'Coaching'];

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('EffeOre · Registro staff Effestudio')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ================= ingresso dati per la PWA ================= */
const API_FN = ['api_init', 'api_requestPin', 'api_login', 'api_me', 'api_logout', 'api_mine', 'api_add', 'api_delete', 'api_adminCheck', 'api_admin', 'api_adminDelete', 'api_setClosed', 'api_export', 'api_modulo', 'api_addExtra', 'api_deleteExtra'];

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (API_FN.indexOf(req.fn) < 0) throw new Error('Funzione non disponibile.');
    const args = Array.isArray(req.args) ? req.args : [];
    out = { ok: true, data: globalThis[req.fn].apply(null, args) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= accesso staff (email + PIN) ================= */

function api_requestPin(email) {
  email = String(email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Email non valida.');
  const st = staff_().find(s => s.active && s.email === email);
  if (!st) throw new Error('Questa email non è registrata. Chiedi ad Andrea di aggiungerla.');
  const cache = CacheService.getScriptCache(), k = 'pinreq_' + st.id;
  if (cache.get(k)) throw new Error('PIN già inviato da poco: controlla la posta (anche lo spam) o riprova tra 2 minuti.');
  const pin = String(Math.floor(100000 + Math.random() * 900000));
  setStaffCell_(st.id, 9, hash_(st.id + ':' + pin));
  GmailApp.sendEmail(email, 'EffeOre: il tuo PIN di accesso',
    'Ciao ' + st.name.split(' ')[0] + ',\n\nil tuo PIN per accedere a EffeOre è: ' + pin +
    '\n\nApri l\'app: https://andreadamo.github.io/effeore/\nInserisci la tua email e questo PIN. Su quel telefono resterai collegato.\n\nNon condividere il PIN con nessuno.\n\nEffestudio Training Center',
    { name: 'EffeOre · Effestudio' });
  cache.put(k, '1', 120);
  return { sent: true };
}

function api_login(email, pin) {
  email = String(email || '').trim().toLowerCase();
  pin = String(pin || '').replace(/\D/g, '');
  const st = staff_().find(s => s.active && s.email === email);
  const cache = CacheService.getScriptCache(), k = 'fail_' + email;
  const fails = Number(cache.get(k) || 0);
  if (fails >= 5) throw new Error('Troppi tentativi sbagliati. Riprova tra 15 minuti o chiedi un nuovo PIN.');
  if (!st || !st.pinHash || hash_(st.id + ':' + pin) !== st.pinHash) {
    cache.put(k, String(fails + 1), 900);
    throw new Error('Email o PIN non corretti.');
  }
  cache.remove(k);
  const token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  const sh = sheet_(SH.access), now = new Date().toISOString();
  sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setNumberFormat('@').setValues([[token, st.id, now, now]]);
  return { token: token, me: meOut_(st) };
}

function api_me(token) { return meOut_(who_(token)); }

function api_logout(token) {
  const r = tokenRow_(token);
  if (r > 0) sheet_(SH.access).deleteRow(r);
  return true;
}

function meOut_(st) { return { id: st.id, name: st.name, role: st.role, admin: adminIds_().indexOf(st.id) >= 0 }; }
function adminIds_() { return (cfg_('ID amministratori') || 'andrea').split(/[,;\s]+/).filter(String); }
function hash_(txt) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, cfg_('Segreto PIN (non toccare)') + '|' + txt)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function tokenRow_(token) {
  token = String(token || '');
  if (token.length < 20) return -1;
  const sh = sheet_(SH.access), n = sh.getLastRow();
  if (n < 2) return -1;
  const v = sh.getRange(2, 1, n - 1, 1).getDisplayValues();
  for (let i = 0; i < v.length; i++) if (v[i][0] === token) return i + 2;
  return -1;
}
function who_(token) {
  const r = tokenRow_(token);
  if (r < 0) throw new Error('Sessione scaduta: accedi di nuovo.');
  const id = sheet_(SH.access).getRange(r, 2).getDisplayValue();
  const st = staff_().find(s => s.id === id && s.active);
  if (!st) throw new Error('Sessione scaduta: accedi di nuovo.');
  return st;
}
function setStaffCell_(id, col, val) {
  const sh = sheet_(SH.staff), n = sh.getLastRow();
  const ids = sh.getRange(2, 1, n - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === id) { sh.getRange(i + 2, col).setNumberFormat('@').setValue(val); return; }
}

/* ================= API staff ================= */

function api_init() {
  return {
    rates: rates_(),
    sedi: sedi_(),
    today: today_(),
    closed: closed_()
  };
}

function api_mine(token, month) {
  checkMonth_(month);
  const staffId = who_(token).id;
  const entries = log_().filter(e => e.staffId === staffId && e.month === month)
    .sort((a, b) => (b.date + b.created).localeCompare(a.date + a.created));
  const clients = {};
  log_().forEach(x => { if (x.staffId === staffId && x.cliente) clients[x.cliente] = 1; });
  return { month: month, closed: closed_().indexOf(month) >= 0, entries: entries, clients: Object.keys(clients).sort((a, b) => a.localeCompare(b)) };
}

function api_add(token, e) {
  const st = who_(token);
  const rt = rates_().find(r => r.role === st.role && r.act === e.act);
  if (!rt) throw new Error('Attività non prevista per il tuo ruolo.');
  const date = String(e.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Data non valida.');
  if (date > today_()) throw new Error('Non puoi segnare attività nei giorni futuri.');
  const month = date.slice(0, 7);
  if (closed_().indexOf(month) >= 0) throw new Error('Il mese di ' + monthLabel_(month) + ' è chiuso. Chiedi in amministrazione.');
  let q = num_(e.qty);
  if (rt.unit === 'ora') {
    q = Math.round(q * 4) / 4;
    if (!(q > 0) || q > 16) throw new Error('Ore non valide: da 0,25 a 16 per registrazione.');
  } else {
    q = Math.round(q);
    if (q < 1 || q > 30) throw new Error('Quantità non valida: da 1 a 30 per registrazione.');
  }
  const sede = sedi_().indexOf(e.sede) >= 0 ? e.sede : '';
  const note = String(e.note || '').slice(0, 200);
  const cliente = String(e.cliente || '').replace(/\s+/g, ' ').trim().slice(0, 80);
  if (CLIENT_ACTS.indexOf(rt.act) >= 0 && !cliente) throw new Error(rt.act === 'PT' ? 'Scrivi il nome del cliente del PT.' : 'Scrivi nome e cognome del cliente del coaching online.');
  if (rt.act === 'Coaching' && log_().some(x => x.staffId === st.id && x.act === 'Coaching' && x.month === month && x.cliente.toLowerCase() === cliente.toLowerCase()))
    throw new Error('Coaching online di ' + cliente + ' già registrato per ' + monthLabel_(month) + '.');
  const amount = Math.round(q * rt.rate * 100) / 100;

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const now = new Date();
    const dup = log_().some(x => x.staffId === st.id && x.date === date && x.act === rt.act && x.qty === q && x.sede === sede && x.cliente === cliente &&
      (now - new Date(x.created)) < 120000);
    if (dup) throw new Error('Registrazione già salvata un attimo fa.');
    sheet_(SH.log).appendRow([Utilities.getUuid().slice(0, 8), date, month, st.id, st.name, st.role, rt.act, q, rt.unit, rt.rate, amount, sede, note, now.toISOString(), cliente]);
  } finally { lock.releaseLock(); }
  return api_mine(token, month);
}

function api_delete(token, id) {
  const staffId = who_(token).id;
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let month;
  try {
    const e = log_().find(x => x.id === id);
    if (!e || e.staffId !== staffId) throw new Error('Registrazione non trovata.');
    if (closed_().indexOf(e.month) >= 0) throw new Error('Il mese è chiuso: non si può più modificare.');
    month = e.month;
    deleteRow_(id);
  } finally { lock.releaseLock(); }
  return api_mine(token, month);
}

/* ================= API amministrazione ================= */

function api_adminCheck(pin) { return pinOk_(pin); }

function api_admin(pin, month) {
  auth_(pin); checkMonth_(month);
  const entries = log_().filter(e => e.month === month)
    .sort((a, b) => a.name.localeCompare(b.name) || a.date.localeCompare(b.date));
  const sum = summary_(entries, month);
  const pt = {};
  entries.filter(e => CLIENT_ACTS.indexOf(e.act) >= 0).forEach(e => {
    const k = e.staffId; pt[k] = pt[k] || { name: e.name, clients: {} };
    const c = (e.cliente || '(senza nome)') + (e.act === 'Coaching' ? ' · coaching online' : '');
    pt[k].clients[c] = (pt[k].clients[c] || 0) + e.qty;
  });
  return {
    month: month, closed: closed_().indexOf(month) >= 0, summary: sum, entries: entries,
    extras: extras_(month), fixed: fixedAll_(),
    staff: staff_().filter(s => s.active).map(s => ({ id: s.id, name: s.name, role: s.role })),
    pt: Object.keys(pt).map(k => ({ id: k, name: pt[k].name, clients: Object.keys(pt[k].clients).sort((a, b) => a.localeCompare(b)).map(c => ({ name: c, n: pt[k].clients[c] })) }))
      .sort((a, b) => a.name.localeCompare(b.name)),
    updated: Utilities.formatDate(new Date(), TZ, 'HH:mm')
  };
}

function api_addExtra(pin, x) {
  auth_(pin);
  const month = String(x && x.month || ''); checkMonth_(month);
  if (closed_().indexOf(month) >= 0) throw new Error('Il mese è chiuso: riaprilo prima di modificare.');
  const st = staff_().find(s => s.id === x.staffId);
  if (!st) throw new Error('Scegli la persona.');
  const desc = String(x.desc || '').trim().slice(0, 120);
  if (!desc) throw new Error('Scrivi la descrizione della voce.');
  const amount = round_(num_(x.amount));
  if (!amount) throw new Error('Importo non valido (usa il meno per le trattenute).');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_(SH.extra), r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, 7).setNumberFormats([['@', '@', '@', '@', '@', '#,##0.00', '@']])
      .setValues([[Utilities.getUuid().slice(0, 8), month, st.id, st.name, desc, amount, new Date().toISOString()]]);
  } finally { lock.releaseLock(); }
  return api_admin(pin, month);
}

function api_deleteExtra(pin, id) {
  auth_(pin);
  const sh = sheet_(SH.extra), n = sh.getLastRow();
  let month = '';
  if (n >= 2) {
    const v = sh.getRange(2, 1, n - 1, 2).getDisplayValues();
    for (let i = 0; i < v.length; i++) if (v[i][0] === id) {
      month = v[i][1];
      if (closed_().indexOf(month) >= 0) throw new Error('Il mese è chiuso: riaprilo prima di modificare.');
      sh.deleteRow(i + 2); break;
    }
  }
  if (!month) throw new Error('Voce non trovata.');
  return api_admin(pin, month);
}

function api_adminDelete(pin, id) {
  auth_(pin);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let month;
  try {
    const e = log_().find(x => x.id === id);
    if (!e) throw new Error('Registrazione non trovata.');
    if (closed_().indexOf(e.month) >= 0) throw new Error('Il mese è chiuso: riaprilo prima di modificare.');
    month = e.month;
    deleteRow_(id);
  } finally { lock.releaseLock(); }
  return api_admin(pin, month);
}

function api_setClosed(pin, month, close) {
  auth_(pin); checkMonth_(month);
  let list = closed_().filter(m => m !== month);
  if (close) list.push(month);
  list.sort();
  setCfg_('Mesi chiusi', list.join(', '));
  return api_admin(pin, month);
}

function api_export(pin, month) {
  auth_(pin); checkMonth_(month);
  const entries = log_().filter(e => e.month === month)
    .sort((a, b) => a.name.localeCompare(b.name) || a.date.localeCompare(b.date));
  const sum = summary_(entries, month);
  const label = monthLabel_(month);
  const out = SpreadsheetApp.create('Compensi staff ' + label + ' - Effestudio');
  out.setSpreadsheetLocale('it_IT');

  // Riepilogo
  const r = out.getSheets()[0]; r.setName('Riepilogo');
  const head = ['Nome', 'Ruolo', 'Ore classi', 'Ore', 'PT (n.)', 'Coaching online (n.)', 'Schede (n.)', 'Ore pulizia', 'Compenso attività €', 'Fisso mensile €', 'Extra / rettifiche €', 'Totale €'];
  const rows = sum.rows.map(p => [p.name, p.role, p.q.Classi || 0, p.q.Ore || 0, p.q.PT || 0, p.q.Coaching || 0, p.q.Schede || 0, p.q.Pulizia || 0, p.activity, p.fixed, p.extra, null]);
  r.getRange(1, 1).setValue('Compensi staff Effestudio Training Center SSD a RL — ' + label).setFontWeight('bold').setFontSize(13);
  r.getRange(3, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#14263B').setFontColor('#FFFFFF');
  if (rows.length) {
    r.getRange(4, 1, rows.length, head.length).setValues(rows);
    r.getRange(4, 12, rows.length, 1).setFormulas(rows.map((_, i) => ['=I' + (i + 4) + '+J' + (i + 4) + '+K' + (i + 4)]));
    const t = rows.length + 4;
    r.getRange(t, 1).setValue('Totale').setFontWeight('bold');
    ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'].forEach(c => r.getRange(c + t).setFormula('=SUM(' + c + '4:' + c + (t - 1) + ')').setFontWeight('bold'));
    r.getRange(4, 3, rows.length + 1, 6).setNumberFormat('0.##');
    r.getRange(4, 9, rows.length + 1, 4).setNumberFormat('#,##0.00 €');
  }
  r.getRange(rows.length + 6, 1).setValue('Tariffe nette: ' + rates_().map(x => x.role + ' ' + x.act + ' ' + x.rate + ' €/' + x.unit).join(' · ')).setFontColor('#5E6B7A');
  r.setFrozenRows(3); r.autoResizeColumns(1, head.length);

  // Dettaglio
  const d = out.insertSheet('Dettaglio');
  const dh = ['Data', 'Nome', 'Ruolo', 'Attività', 'Quantità', 'Unità', 'Tariffa €', 'Importo €', 'Sede', 'Cliente PT', 'Note'];
  d.getRange(1, 1, 1, dh.length).setValues([dh]).setFontWeight('bold').setBackground('#14263B').setFontColor('#FFFFFF');
  if (entries.length) {
    d.getRange(2, 1, entries.length, 1).setNumberFormat('@');
    d.getRange(2, 1, entries.length, dh.length).setValues(entries.map(e =>
      [e.date.split('-').reverse().join('/'), e.name, e.role, e.act, e.qty, e.unit, e.rate, e.amount, e.sede, e.cliente, e.note]));
    d.getRange(2, 7, entries.length, 2).setNumberFormat('#,##0.00 €');
  }
  d.setFrozenRows(1); d.autoResizeColumns(1, dh.length);

  const file = DriveApp.getFileById(out.getId());
  file.moveTo(exportFolder_());
  const base = 'https://docs.google.com/spreadsheets/d/' + out.getId();
  return {
    url: out.getUrl(),
    xlsx: base + '/export?format=xlsx',
    pdf: base + '/export?format=pdf&gid=' + r.getSheetId() + '&portrait=false&fitw=true&gridlines=false',
    text: message_(label, sum)
  };
}

/* ================= modulo per lo Studio Mattonai ================= */
const STUDIO_EMAIL = 'consulenzalavoro@studiomattonai.it';

function api_modulo(pin, month) {
  auth_(pin); checkMonth_(month);
  const entries = log_().filter(e => e.month === month);
  const sum = summary_(entries, month);
  const staff = staff_();
  const rows = sum.rows.filter(p => p.total > 0).map(p => {
    const s = staff.find(x => x.id === p.id) || {};
    return { name: s.formName || p.name, order: s.order || 999, total: p.total };
  }).sort((a, b) => a.order - b.order || a.name.localeCompare(b.name));
  if (!rows.length) throw new Error('Nessun compenso da inviare per ' + monthLabel_(month) + '.');

  const label = monthLabel_(month).toUpperCase();
  const html = moduloHtml_(label, rows);
  const pdf = Utilities.newBlob(html, 'text/html', 'modulo.html').getAs('application/pdf')
    .setName('Modulo_Compensi_Effestudio_' + label.replace(' ', '_') + '.pdf');
  const file = exportFolder_().createFile(pdf);

  const body = 'Buongiorno,\n\nin allegato la richiesta compensi dei collaboratori di Effestudio Training Center SSD a RL relativa a ' +
    monthLabel_(month) + '.\n\nGrazie,\nAndrea D\'Amore\nEffestudio Training Center SSD a RL';
  const draft = GmailApp.createDraft(STUDIO_EMAIL, 'Richiesta compensi collaboratori - ' + label + ' - Effestudio Training Center', body,
    { attachments: [file.getBlob()], name: 'Effestudio Training Center' });
  const me = Session.getEffectiveUser().getEmail();
  return {
    pdf: file.getUrl(),
    draft: 'https://mail.google.com/mail/?authuser=' + encodeURIComponent(me) + '#drafts?compose=' + draft.getMessage().getId(),
    total: rows.reduce((t, r) => round_(t + r.total), 0),
    count: rows.length
  };
}

function moduloHtml_(label, rows) {
  const eur = v => {
    const p = Math.abs(v).toFixed(2).split('.');
    return '€ ' + p[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + p[1];
  };
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const tr = rows.map(r => '<tr><td>' + esc(r.name) + '</td><td>' + eur(r.total) + '</td><td></td></tr>').join('');
  return '<html><head><meta charset="utf-8"><style>' +
    '@page{size:A4;margin:18mm 16mm}' +
    'body{font-family:Arial,Helvetica,sans-serif;font-size:10.5pt;color:#000}' +
    'h1{font-size:14pt;text-align:center;margin:0 0 26px}' +
    '.hd{display:table;width:100%;margin-bottom:26px}.hd>div{display:table-cell;vertical-align:top}' +
    '.co{line-height:1.6}.to{padding-left:30px}.to b{font-size:12.5pt}.to small{display:block;font-size:8.5pt;line-height:1.5}' +
    '.mese{margin-bottom:22px;line-height:1.6}' +
    'table{width:100%;border-collapse:collapse}td,th{border:1px solid #000;padding:6px 5px;text-align:left;font-size:9.5pt}' +
    'th{font-weight:normal;font-size:8.5pt}th.x{font-weight:bold;text-align:center;font-size:7.5pt}' +
    'td:nth-child(1){width:46%}td:nth-child(2){width:28%}' +
    '</style></head><body>' +
    '<h1>Richiesta Compensi Collaboratori</h1>' +
    '<div class="hd"><div class="co"><b>EFFESTUDIO TRAINING CENTER SSD A RL</b><br>Via Giosuè Borsi , 20<br>Cap 57023 CECINA (LI)<br>Cod.Fisc: 02053370496</div>' +
    '<div class="to">A:&nbsp; <b>Studio Mattonai</b><small>Inviare via fax al num. 0587/756991</small><small>Inviare via mail a ' + STUDIO_EMAIL + '</small></div></div>' +
    '<div class="mese">Vi chiediamo di predisporre le seguenti ricevute<br>relative al mese di &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>' + esc(label) + '</b></div>' +
    '<table><tr><th>COGNOME E NOME COLLABORATORE</th><th>COMPENSO MENSILE LORDO</th><th class="x">EXTRA ALTRE PALESTRE</th></tr>' + tr + '</table>' +
    '</body></html>';
}

/* ================= calcoli ================= */

function summary_(entries, month) {
  const staff = staff_(), fixed = fixedAll_(), extras = month ? extras_(month) : [];
  const by = {};
  const get = (id, name, role) => by[id] = by[id] || { id: id, name: name, role: role, q: {}, eur: {}, activity: 0, fixed: 0, extra: 0, fixedItems: [], extraItems: [] };
  entries.forEach(e => {
    const p = get(e.staffId, e.name, e.role);
    p.q[e.act] = round_((p.q[e.act] || 0) + e.qty);
    p.eur[e.act] = round_((p.eur[e.act] || 0) + e.amount);
    p.activity = round_(p.activity + e.amount);
  });
  fixed.filter(f => f.active).forEach(f => {
    const s = staff.find(x => x.id === f.staffId); if (!s || !s.active) return;
    const p = get(s.id, s.name, s.role); p.fixed = round_(p.fixed + f.amount); p.fixedItems.push({ desc: f.desc, amount: f.amount });
  });
  extras.forEach(x => {
    const s = staff.find(y => y.id === x.staffId) || { id: x.staffId, name: x.name, role: '' };
    const p = get(s.id, s.name, s.role); p.extra = round_(p.extra + x.amount); p.extraItems.push({ desc: x.desc, amount: x.amount });
  });
  const rows = Object.keys(by).map(k => {
    const p = by[k], s = staff.find(x => x.id === k);
    if (s) { p.role = s.role; p.name = s.name; }
    p.total = round_(p.activity + p.fixed + p.extra);
    return p;
  }).sort((a, b) => roleRank_(a.role) - roleRank_(b.role) || a.name.localeCompare(b.name));
  const tot = rows.reduce((t, p) => ({ activity: round_(t.activity + p.activity), fixed: round_(t.fixed + p.fixed), extra: round_(t.extra + p.extra), total: round_(t.total + p.total) }), { activity: 0, fixed: 0, extra: 0, total: 0 });
  return { rows: rows, tot: tot, acts: ACT_ORDER };
}

function fixedAll_() {
  let sh; try { sh = sheet_(SH.fixed); } catch (e) { return []; }
  const n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 5).getValues().filter(r => String(r[0]).trim()).map(r => ({
    staffId: String(r[0]).trim(), name: String(r[1]).trim(), desc: String(r[2]).trim(), amount: num_(r[3]),
    active: !(r[4] === false || /^(no|false|0)$/i.test(String(r[4]).trim()))
  }));
}

function extras_(month) {
  let sh; try { sh = sheet_(SH.extra); } catch (e) { return []; }
  const n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 7).getValues().filter(r => String(r[0]).trim()).map(r => ({
    id: String(r[0]), month: r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'yyyy-MM') : String(r[1]).slice(0, 7),
    staffId: String(r[2]), name: String(r[3]), desc: String(r[4]), amount: num_(r[5])
  })).filter(x => x.month === month);
}

function message_(label, sum) {
  const eur = v => {
    const p = Math.abs(v).toFixed(2).split('.');
    return (v < 0 ? '-' : '') + p[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + p[1] + ' €';
  };
  const unit = { Classi: 'h classi', Ore: 'h', PT: 'PT', Coaching: 'coaching online', Schede: 'schede', Pulizia: 'h pulizia' };
  const lines = sum.rows.map(p => {
    const det = ACT_ORDER.filter(a => p.q[a]).map(a => String(p.q[a]).replace('.', ',') + ' ' + unit[a]);
    if (p.fixed) det.push('fisso ' + eur(p.fixed));
    if (p.extra) det.push('extra ' + eur(p.extra));
    return '- ' + p.name + ': ' + eur(p.total) + (det.length ? ' (' + det.join(', ') + ')' : '');
  });
  return 'Buongiorno,\nvi invio i compensi dello staff di Effestudio Training Center SSD a RL relativi a ' + label + '.\n\n' +
    lines.join('\n') + '\n\nTotale: ' + eur(sum.tot.total) + '\n\nIn allegato il riepilogo con il dettaglio delle attività.\nGrazie,\nAndrea D\'Amore';
}

/* ================= foglio ================= */

function sheet_(n) {
  const s = SpreadsheetApp.getActive().getSheetByName(n);
  if (!s) throw new Error('Tab "' + n + '" non trovata nel foglio.');
  return s;
}

function data_(n) {
  const s = sheet_(n), r = s.getLastRow(), c = s.getLastColumn();
  if (r < 2) return [];
  return s.getRange(2, 1, r - 1, c).getValues();
}

function staff_() {
  return data_(SH.staff).filter(r => String(r[0]).trim() && String(r[1]).trim()).map(r => ({
    id: String(r[0]).trim(), name: String(r[1]).trim(), role: String(r[2]).trim(),
    active: !(r[3] === false || /^(no|false|0)$/i.test(String(r[3]).trim())),
    fixed: num_(r[4]),
    formName: String(r[5] || '').trim(),
    order: num_(r[6]) || 999,
    email: String(r[7] || '').trim().toLowerCase(),
    pinHash: String(r[8] || '').trim()
  }));
}

function rates_() {
  return data_(SH.rates).filter(r => String(r[0]).trim() && String(r[1]).trim()).map(r => ({
    role: String(r[0]).trim(), act: String(r[1]).trim(), unit: String(r[2]).trim().toLowerCase() || 'ora', rate: num_(r[3])
  }));
}

function log_() {
  return data_(SH.log).filter(r => String(r[0]).trim()).map(r => ({
    id: String(r[0]), date: dstr_(r[1]), month: dstr_(r[1]).slice(0, 7),
    staffId: String(r[3]), name: String(r[4]), role: String(r[5]), act: String(r[6]),
    qty: num_(r[7]), unit: String(r[8]), rate: num_(r[9]), amount: num_(r[10]),
    sede: String(r[11] || ''), note: String(r[12] || ''), created: r[13] instanceof Date ? r[13].toISOString() : String(r[13] || ''),
    cliente: String(r[14] || '').trim()
  }));
}

function deleteRow_(id) {
  const s = sheet_(SH.log), n = s.getLastRow();
  if (n < 2) return;
  const ids = s.getRange(2, 1, n - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) { s.deleteRow(i + 2); return; }
}

function cfgRow_(key) {
  const s = sheet_(SH.cfg), v = s.getDataRange().getDisplayValues();
  for (let i = 0; i < v.length; i++) if (v[i][0] === key) return { row: i + 1, value: String(v[i][1]).trim() };
  return null;
}
function cfg_(key) { const r = cfgRow_(key); return r ? r.value : ''; }
function setCfg_(key, val) {
  const s = sheet_(SH.cfg), r = cfgRow_(key);
  const row = r ? r.row : s.getLastRow() + 1;
  s.getRange(row, 1, 1, 2).setNumberFormat('@').setValues([[key, val]]);
}

function closed_() { return cfg_('Mesi chiusi').split(/[,;\s]+/).filter(m => /^\d{4}-\d{2}$/.test(m)); }
function sedi_() { const v = cfg_('Sedi'); return v ? v.split(',').map(x => x.trim()).filter(String) : ['Cecina Center', 'LAB', 'Pisa', 'Tennis', 'Rosignano']; }

function exportFolder_() {
  const id = cfg_('Cartella export ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('Export commercialista - EffeOre');
  setCfg_('Cartella export ID', f.getId());
  return f;
}

/* ================= utilità ================= */

function pinOk_(pin) { const p = cfg_('PIN amministrazione'); return !!p && String(pin || '').trim() === p; }
function auth_(pin) { if (!pinOk_(pin)) throw new Error('PIN amministrazione non valido.'); }
function checkMonth_(m) { if (!/^\d{4}-\d{2}$/.test(String(m))) throw new Error('Mese non valido.'); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function dstr_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '').trim(); }
function round_(v) { return Math.round(v * 100) / 100; }
function roleRank_(r) { const i = ['Operativo', 'New entry', 'Amministrativo'].indexOf(r); return i < 0 ? 9 : i; }
function num_(v) {
  if (typeof v === 'number') return v;
  let s = String(v || '').replace(/[€\s]/g, '');
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function monthLabel_(m) {
  const n = ['gennaio', 'febbraio', 'marzo', 'aprile', 'maggio', 'giugno', 'luglio', 'agosto', 'settembre', 'ottobre', 'novembre', 'dicembre'];
  return n[Number(m.slice(5, 7)) - 1] + ' ' + m.slice(0, 4);
}

/* Esegui una volta dall'editor per autorizzare e controllare il foglio */
function setup() {
  [SH.staff, SH.rates, SH.log, SH.cfg, SH.fixed, SH.extra, SH.access].forEach(sheet_);
  exportFolder_();
  Logger.log('Pronto. Staff attivi: ' + staff_().filter(s => s.active).length + ' · Tariffe: ' + rates_().length);
}

/* ---------- keep-warm: evita avvio lento dell'app ---------- */
function keepWarm() {
  try {
    UrlFetchApp.fetch(ScriptApp.getService().getUrl(), {
      method: 'post', contentType: 'text/plain', muteHttpExceptions: true,
      payload: JSON.stringify({ fn: 'api_init', args: [] })
    });
  } catch (e) {}
}

function setupWarm() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'keepWarm').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('keepWarm').timeBased().everyMinutes(5).create();
  GmailApp.getAliases();
  keepWarm();
  Logger.log('Keep-warm attivo ogni 5 minuti.');
}
