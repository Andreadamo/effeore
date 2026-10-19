const SC_SS_ID = '1_m6MwvASf7epZNC9JhM2H4AcYDiy7ZIJIT3t4by-XYc';
const SC_INV = 'Inventario';
const SC_PREN = 'Prenotazioni';

function scan_ss_() { return SpreadsheetApp.openById(SC_SS_ID); }

function scan_pinOk_(pin) {
  const p = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || '1234';
  if (String(pin) !== String(p)) throw new Error('PIN errato');
}

function scan_ping(pin) { scan_pinOk_(pin); return true; }

function scan_prezzo_(r) {
  const g = Number(r[6]) || 0, o = Number(r[14]) || 0;
  return (o > 0 && o < g) ? o : g;
}

function scan_inv_() {
  const sh = scan_ss_().getSheetByName(SC_INV);
  const v = sh.getDataRange().getValues();
  const bc = v[0].map(String).indexOf('Barcode');
  return { sh: sh, v: v, bc: bc };
}

function scan_attivo_(x) { return x !== false && String(x).toUpperCase() !== 'FALSE'; }

function scan_data_(d) {
  return d instanceof Date ? Utilities.formatDate(d, 'Europe/Rome', 'dd/MM HH:mm') : String(d);
}

function scan_cerca(pin, codice) {
  scan_pinOk_(pin);
  const c = String(codice || '').trim().toUpperCase();
  if (!c) throw new Error('Codice vuoto');
  const pv = scan_ss_().getSheetByName(SC_PREN).getDataRange().getValues();

  let idProd = null, rigaPren = null;
  for (let i = 1; i < pv.length; i++) {
    if (String(pv[i][0]).trim().toUpperCase() === c) { idProd = String(pv[i][5]).toUpperCase(); rigaPren = i + 1; break; }
  }

  const inv = scan_inv_();
  let r = null;
  for (let i = 1; i < inv.v.length; i++) {
    const id = String(inv.v[i][0]).trim().toUpperCase();
    const bc = inv.bc >= 0 ? String(inv.v[i][inv.bc]).trim().toUpperCase() : '';
    if (idProd ? id === idProd : (id === c || (bc && bc === c))) { r = inv.v[i]; break; }
  }
  if (!r) throw new Error('Codice non trovato: ' + c);

  const pren = [];
  for (let i = 1; i < pv.length; i++) {
    if (String(pv[i][5]) === String(r[0]) && String(pv[i][10]) === 'RISERVATO' && (!rigaPren || rigaPren === i + 1)) {
      pren.push({ riga: i + 1, codice: String(pv[i][0]), nome: String(pv[i][2]), qta: Number(pv[i][8]) || 1, totale: Number(pv[i][9]) || 0, data: scan_data_(pv[i][1]) });
    }
  }
  const stock = Number(r[7]) || 0, ris = Number(r[8]) || 0;
  return { id: String(r[0]), prodotto: String(r[1]), linea: String(r[2]), taglia: String(r[5]), prezzo: scan_prezzo_(r), stock: stock, disp: stock - ris, pren: pren };
}

function scan_aggiornaInv_(id, dStock, dRis, dVend) {
  const inv = scan_inv_();
  for (let i = 1; i < inv.v.length; i++) {
    if (String(inv.v[i][0]) === String(id)) {
      const rg = inv.sh.getRange(i + 1, 8, 1, 3);
      const c = rg.getValues()[0].map(function (x) { return Number(x) || 0; });
      if (c[0] + dStock < 0) throw new Error('Stock insufficiente per ' + id);
      rg.setValues([[c[0] + dStock, Math.max(0, c[1] + dRis), c[2] + dVend]]);
      return;
    }
  }
  throw new Error('Articolo non trovato in inventario: ' + id);
}

function scan_vendiPrenotazione(pin, riga, operatore, metodo) {
  scan_pinOk_(pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = scan_ss_().getSheetByName(SC_PREN);
    const r = sh.getRange(riga, 1, 1, 14).getValues()[0];
    if (String(r[10]) !== 'RISERVATO') throw new Error('Prenotazione non attiva (' + r[10] + ')');
    const q = Number(r[8]) || 1;
    scan_aggiornaInv_(r[5], -q, -q, q);
    const nota = (r[13] ? r[13] + ' | ' : '') + 'Metodo: ' + metodo;
    sh.getRange(riga, 11, 1, 4).setValues([['VENDUTO', new Date(), operatore || 'scanner', nota]]);
    SpreadsheetApp.flush();
    return 'VENDUTO: ' + r[6] + ' tg ' + r[7] + ' — ' + r[2] + ' — € ' + r[9] + ' (' + metodo + ')';
  } finally {
    lock.releaseLock();
  }
}

function scan_vendiBanco(pin, id, operatore, metodo) {
  scan_pinOk_(pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const inv = scan_inv_();
    let r = null;
    for (let i = 1; i < inv.v.length; i++) if (String(inv.v[i][0]) === String(id)) { r = inv.v[i]; break; }
    if (!r) throw new Error('Articolo non trovato: ' + id);
    if ((Number(r[7]) || 0) - (Number(r[8]) || 0) < 1) throw new Error('Nessun pezzo libero: tutti prenotati o esauriti');
    const p = scan_prezzo_(r);
    scan_aggiornaInv_(id, -1, 0, 1);
    const now = new Date();
    const cod = 'BC' + Utilities.formatDate(now, 'Europe/Rome', 'MMdd') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
    scan_ss_().getSheetByName(SC_PREN).appendRow([cod, now, 'Vendita banco', '', '', r[0], r[1], r[5], 1, p, 'VENDUTO', now, operatore || 'scanner', 'Metodo: ' + metodo]);
    SpreadsheetApp.flush();
    return 'VENDUTO (banco): ' + r[1] + ' tg ' + r[5] + ' — € ' + p + ' (' + metodo + ')';
  } finally {
    lock.releaseLock();
  }
}

function scan_etichette(pin) {
  scan_pinOk_(pin);
  return scan_inv_().v.slice(1)
    .filter(function (r) { return r[0] && scan_attivo_(r[12]) && Number(r[7]) > 0; })
    .map(function (r) { return { id: String(r[0]), prodotto: String(r[1]), taglia: String(r[5]), prezzo: scan_prezzo_(r), pz: Number(r[7]) }; });
}
