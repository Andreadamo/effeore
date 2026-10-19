/***** API JSON per le pagine merch su GitHub (andreadamo.github.io/effeore/merch) *****/
const MERCH_TZ = 'Europe/Rome';
const MERCH_API_FN = ['api_catalogo', 'api_riserva', 'api_login', 'api_prenotazioni', 'api_stato',
  'api_cerca', 'api_vendiBanco', 'api_etichette', 'api_dashboard'];

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (MERCH_API_FN.indexOf(req.fn) < 0) throw new Error('Funzione non disponibile.');
    const args = Array.isArray(req.args) ? req.args : [];
    out = { ok: true, data: globalThis[req.fn].apply(null, args) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function merchIso_(d) {
  return d instanceof Date ? Utilities.formatDate(d, MERCH_TZ, "yyyy-MM-dd'T'HH:mm") : '';
}

function merchMetodo_(note) {
  const m = String(note || '').match(/Metodo:\s*([^|]+)/);
  return m ? m[1].trim() : '';
}

/* ---------- clienti ---------- */
function api_catalogo() { return getCatalogo(); }
function api_riserva(p) { return riserva(p || {}); }

/* ---------- staff ---------- */
function api_login(pin) { checkPin_(pin); return true; }

function api_prenotazioni(pin) {
  checkPin_(pin);
  const rows = scan_ss_().getSheetByName(SC_PREN).getDataRange().getValues().slice(1).filter(function (r) { return r[0]; });
  const out = rows.map(function (r) {
    return {
      codice: String(r[0]), data: merchIso_(r[1]), nome: String(r[2]), tel: String(r[3]), email: String(r[4]),
      id: String(r[5]), prodotto: String(r[6]), taglia: String(r[7]), qta: Number(r[8]) || 0, totale: Number(r[9]) || 0,
      stato: String(r[10]), chiusura: merchIso_(r[11]), operatore: String(r[12] || ''), metodo: merchMetodo_(r[13])
    };
  }).reverse();
  return {
    scadenzaGg: SCADENZA_GG,
    aperte: out.filter(function (x) { return x.stato === 'RISERVATO'; }),
    chiuse: out.filter(function (x) { return x.stato !== 'RISERVATO'; }).slice(0, 60)
  };
}

function api_stato(pin, codice, stato, operatore, metodo) {
  aggiornaStato(pin, codice, stato, operatore || 'staff');
  if (stato === 'VENDUTO' && metodo) {
    const sh = scan_ss_().getSheetByName(SC_PREN);
    const v = sh.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][0]) === String(codice)) {
        sh.getRange(i + 1, 14).setValue((v[i][13] ? v[i][13] + ' | ' : '') + 'Metodo: ' + metodo);
        break;
      }
    }
  }
  return api_prenotazioni(pin);
}

function api_cerca(pin, codice) { return scan_cerca(pin, codice); }
function api_vendiBanco(pin, id, operatore, metodo) { return scan_vendiBanco(pin, id, operatore, metodo); }
function api_etichette(pin) { return scan_etichette(pin); }

/* ---------- dashboard di controllo ---------- */
function api_dashboard(pin) {
  checkPin_(pin);
  const ss = scan_ss_();
  const inv = ss.getSheetByName(SC_INV).getDataRange().getValues().slice(1).filter(function (r) { return r[0]; })
    .map(function (r) {
      const stock = Number(r[7]) || 0, ris = Number(r[8]) || 0;
      return {
        id: String(r[0]), prodotto: String(r[1]), linea: String(r[2]), genere: String(r[3]), categoria: String(r[4]),
        taglia: String(r[5]), prezzo: Number(r[6]) || 0, stock: stock, riservati: ris, venduti: Number(r[9]) || 0,
        disp: stock - ris, attivo: scan_attivo_(r[12]), promo: Number(r[14]) || 0
      };
    });
  const pren = ss.getSheetByName(SC_PREN).getDataRange().getValues().slice(1).filter(function (r) { return r[0]; })
    .map(function (r) {
      return {
        codice: String(r[0]), data: merchIso_(r[1]), nome: String(r[2]), id: String(r[5]), prodotto: String(r[6]),
        taglia: String(r[7]), qta: Number(r[8]) || 0, totale: Number(r[9]) || 0, stato: String(r[10]),
        chiusura: merchIso_(r[11]), operatore: String(r[12] || ''), metodo: merchMetodo_(r[13]),
        canale: /^BC/.test(String(r[0])) ? 'Banco' : 'Online'
      };
    });
  return { generato: merchIso_(new Date()), scadenzaGg: SCADENZA_GG, ricarico: RICARICO, inventario: inv, prenotazioni: pren };
}
