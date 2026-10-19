// Chiamate all'archivio merch: POST text/plain (nessun preflight), risposta {ok,data|error}
async function run(fn, ...args) {
  if (!navigator.onLine) throw new Error("Sei offline: collegati a internet e riprova.");
  let r;
  try {
    r = await fetch(window.MERCH_API, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ fn, args }), redirect: "follow", cache: "no-store" });
  } catch (e) { throw new Error("Connessione non riuscita. Controlla internet e riprova."); }
  let j;
  try { j = await r.json(); } catch (e) { throw new Error("Risposta non valida dal server (" + r.status + "). Riprova tra poco."); }
  if (!j.ok) throw new Error(j.error || "Errore sconosciuto.");
  return j.data;
}
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const eur = n => (Number(n) || 0).toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €";
function store(k, v) { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch (e) {} }
function load(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
const FOTO = { "AB114-NER": "img/AB114-NER.webp", "AB114-BIA": "img/AB114-BIA.webp", "AS104-NER": "img/AS104-NER.webp",
  "AK011-MAR": "img/AK011-MAR.webp", "AK104-NER": "img/AK104-NER.webp", "AS029-NER": "img/AS029-NER.webp" };
const fotoDi = g => g.foto || FOTO[String(g.id).split("-").slice(0, 2).join("-")] || "";
