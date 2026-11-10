// 🗂 8/11: τωρινοί ασθενείς → «Αρχείο Ασθενών» (βάση στο Παράθυρο).
// Καθαρή συνάρτηση (χωρίς δίκτυο): Παραδόχαρτο (oncall_patients) + λίστα ΝΠΣ (npslist) + στοιχεία ΝΠΣ (npsenrich)
// → γραμμές για POST /api/patients/bulk {surname, firstname, father, amka, am, dx}.
// Κανόνες:
//  • Ταίριασμα Παραδόχαρτο ↔ ΝΠΣ: ίδιος ΑΜ, αλλιώς ίδιες λέξεις ονόματος με τη σειρά (επιτρέπεται πρόθεμα ≥4 γραμμάτων, όπως η σύγκριση του Παραδόχαρτου).
//  • Όνομα: το πληρέστερο (ΝΠΣ συνήθως «ΕΠΩΝΥΜΟ ΟΝΟΜΑ»). ΑΜΚΑ/ΑΜ/διαγνώσεις: ένωση όλων των πηγών.
//  • ΣΤΕΛΝΕΤΑΙ ΜΟΝΟ ασθενής με έγκυρο ΑΜΚΑ ή ΑΜ: έτσι κάθε επανάληψη ενημερώνει τον ίδιο ασθενή και δεν φτιάχνει διπλούς.
//    Οι υπόλοιποι επιστρέφονται στο skipped (τους συμπληρώνει το npsenrich και μπαίνουν στην επόμενη φορά).
(function (root) {
  "use strict";
  const nz = (x) => String(x || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/[^A-ZΑ-Ω0-9 ]/g, " ").replace(/\s+/g, " ").trim();
  const words = (x) => nz(x).split(" ").filter(Boolean);
  function amkaOk(a) {
    a = String(a || "").replace(/\D/g, ""); if (!/^\d{11}$/.test(a)) return false;
    let s = 0; for (let i = 0; i < 11; i++) { let d = +a[10 - i]; if (i % 2) { d *= 2; if (d > 9) d -= 9; } s += d; } return s % 10 === 0;
  }
  const cleanAmka = (a) => (amkaOk(a) ? String(a).replace(/\D/g, "") : "");
  const cleanAm = (a) => { const d = String(a || "").replace(/\D/g, ""); return /^\d{8}$/.test(d) ? d : ""; };
  const CODE = /\b([A-Z]\d\d(?:\.\d{1,2})?)\b/g;
  function codesOf(v) {
    const out = [];
    const take = (s) => { const m = String(s || "").toUpperCase().match(CODE); if (m) out.push(...m); };
    if (Array.isArray(v)) v.forEach((x) => take(x && typeof x === "object" ? x.code : x)); else take(v);
    return out;
  }
  // ίδιο πρόσωπο κατά όνομα (όπως nbHere του Παραδόχαρτου)
  function sameName(a, b) {
    const u = words(a), t = words(b); if (!u.length || !t.length) return false;
    const [x, y] = u.length <= t.length ? [u, t] : [t, u];
    return x.every((w, i) => y[i] && (y[i] === w || (w.length >= 4 && y[i].startsWith(w)) || (y[i].length >= 4 && w.startsWith(y[i]))));
  }
  // πλήρες όνομα από το Παραδόχαρτο: «ΕΠΩΝΥΜΟ» + fname (μικρό όνομα από ΝΠΣ) ή «ΕΠΩΝΥΜΟ ΟΝΟΜΑ»
  function pdxFull(rec) {
    const n = String(rec.name || "").replace(/\s+/g, " ").trim();
    if (n && !n.includes(" ") && rec.fname) return n + " " + String(rec.fname).trim();
    return n;
  }
  function split(full) { const w = String(full || "").replace(/\s+/g, " ").trim().split(" "); return { surname: w[0] || "", firstname: w.slice(1).join(" ") }; }

  // 🎗 9/11 (χρήστης: «όταν φτιάχνεται διαφάνεια, το ιστορικό από το Παραδόχαρτο αυτόματα»): ιστορικό ασθενή → βάση (γεγονός «history»),
  // στη μορφή που χρειάζεται η διαφάνεια του ογκολογικού: θεράπων, Α/α, GFR, απεικονίσεις, αιτία εισόδου, προηγούμενα χειρουργεία.
  const DXA = [[/^I1[0-5]/, "ΑΥ"], [/^E1[0-4]/, "ΣΔ"], [/^J4[34]/, "ΧΑΠ"], [/^I48/, "ΚΜ"], [/^I2[0-5]/, "ΣΝ"], [/^N18/, "ΧΝΝ"], [/^E78/, "ΔΛΔ"], [/^I50/, "ΚΑ"], [/^E0[3-5]/, "Θυρεοειδής"], [/^I6[3-9]/, "ΑΕΕ"], [/^G20/, "Parkinson"], [/^J45/, "Άσθμα"], [/^K7[0-4]/, "Ηπατοπάθεια"]];
  const female = (first, sur) => { const f = nz(first).split(" ")[0] || "", s = nz(sur); if (f) return /(Α|Η|Ω|Ι|ΟΥ)$/.test(f) && !/(ΟΣ|ΗΣ|ΑΣ|ΩΝ|ΗΣ)$/.test(f); return /(ΟΥ|Η|Α)$/.test(s) && !/(ΟΣ|ΗΣ|ΑΣ)$/.test(s); };
  function egfr(cr, age, fem) { cr = parseFloat(String(cr || "").replace(",", ".")); age = +age; if (!(cr > 0.2 && cr < 20) || !(age > 17 && age < 110)) return null;
    const k = fem ? 0.7 : 0.9, a = fem ? -0.241 : -0.302; return Math.round(142 * Math.pow(Math.min(cr / k, 1), a) * Math.pow(Math.max(cr / k, 1), -1.2) * Math.pow(0.9938, age) * (fem ? 1.012 : 1)); }
  const IMG = [[/ct|αξον/i, "Ct"], [/mri|μαγν/i, "MRI"], [/pet/i, "PET"], [/u\/?s|υπερ[ηή]χ/i, "U/S"], [/cysto|κυστεοσκ/i, "Cysto"]];
  function histOf(rec, ther, full, amka) {
    const d = (rec.caseHist && rec.caseHist.d) || {}, o = {}, t = (x) => String(x == null ? "" : x).replace(/\s+/g, " ").trim();
    const tf = ther && rec.therapist ? t(ther[rec.therapist] || "") : ""; o.therapist = tf ? tf.split(" ").find((w) => w.length > 2) || tf : t(d.attending || rec.therapist || "");
    const sys = d.systems && typeof d.systems === "object" ? Object.values(d.systems).map(t).filter((x) => x && !/^[-–—]$|^όχι$|^ουδέν$/i.test(x)) : [];
    const cod = rec.comorb && Array.isArray(rec.comorb.codes) ? rec.comorb.codes.map((c) => String((c && c.code) || c || "").toUpperCase()) : [];
    const ab = []; for (const c of cod) { const m = DXA.find(([re]) => re.test(c)); if (m && !ab.includes(m[1])) ab.push(m[1]); }
    o.comorb = sys.length ? sys : ab;
    o.imaging = (Array.isArray(d.workup) ? d.workup : []).filter((w) => t(w && w.text)).map((w) => { const lb = t(w.label), m = IMG.find(([re]) => re.test(lb) || re.test(t(w.text).slice(0, 12))); return { kind: m ? m[1] : (lb.split(/[\/ ]/)[0] || "Απεικόνιση"), text: t(w.text).replace(/^\*+|\*+$/g, "").replace(/^(ct|mri|pet(-ct)?|u\/?s|cysto|αξονικ\S*|μαγνητικ\S*)\s*[:\-–]\s*/i, "") }; });
    o.notes = [t(d.admission_reason)].filter(Boolean);
    if (Array.isArray(d.past_surgeries)) o.past = d.past_surgeries.map(t).filter(Boolean).slice(0, 8);
    const nm = split(full), age = parseInt(d.age, 10) || (amka ? (() => { const y = +amka.slice(4, 6), now = new Date().getFullYear() % 100; const yy = (y > now ? 1900 : 2000) + y; return new Date().getFullYear() - yy; })() : 0);
    const g = egfr(rec.cr, age, female(nm.firstname, nm.surname)); if (g) o.gfr = g;
    o.src = "pdx"; o.at = new Date().toISOString().slice(0, 10);
    return (o.comorb.length || o.imaging.length || o.notes.length || o.therapist || o.gfr) ? o : null;
  }
  function buildCurrentRows(src) {
    const pdx = (src && src.pdx) || [], nps = (src && src.nps) || [], enr0 = (src && src.enr) || {};
    const enr = {}; Object.keys(enr0).forEach((k) => { enr[nz(k)] = enr0[k] || {}; });
    const enrOf = (name) => enr[nz(name)] || {};
    const people = [], used = new Set();
    // 1) Παραδόχαρτο
    for (const row of pdx) {
      const rec = (row && row.record) || {}; const full = pdxFull(rec); if (!full) continue;
      const p = { full, father: String(rec.father || "").trim(), amka: [cleanAmka(rec.amka)], am: [cleanAm(rec.am)], dx: [], src: "pdx", key: row.key };
      try { p.hist = histOf(rec, (src && src.ther) || {}, full, cleanAmka(rec.amka)); } catch (e) {}
      // 🔪 8/11: επέμβαση(εις) + ημερομηνία (podDate) → «αναμονή βιοψίας» στη βάση μέχρι να βγει ιστολογική
      { const d = /^\d{4}-\d{2}-\d{2}/.test(String(rec.podDate || "")) ? String(rec.podDate).slice(0, 10) : "";
        const nm = [rec.surgeryCustom || rec.surgery, rec.surgery2Custom || rec.surgery2].map((x) => String(x || "").trim()).filter((x) => x && !/^άλλ[οη]$/i.test(x));
        p.surg = d ? [...new Set(nm)].map((name) => ({ name, date: d })) : []; }
      p.dx.push(...codesOf(rec.dx));
      if (rec.comorb && rec.comorb.dcx !== false && Array.isArray(rec.comorb.codes)) p.dx.push(...codesOf(rec.comorb.codes));
      // ταίριασμα με τη λίστα ΝΠΣ
      let mi = -1; const myAm = cleanAm(rec.am);
      if (myAm) mi = nps.findIndex((r, i) => !used.has(i) && cleanAm(r.am) === myAm);
      if (mi < 0) mi = nps.findIndex((r, i) => !used.has(i) && sameName(r.name, full));
      if (mi >= 0) {
        used.add(mi); const r = nps[mi];
        if (words(r.name).length > words(full).length) p.full = String(r.name).replace(/\s+/g, " ").trim();
        p.amka.push(cleanAmka(r.amka)); p.am.push(cleanAm(r.am));
        const e = enrOf(r.name); p.amka.push(cleanAmka(e.amka)); p.am.push(cleanAm(e.am)); p.dx.push(...codesOf(e.dx));
      }
      const e2 = enrOf(full); p.amka.push(cleanAmka(e2.amka)); p.am.push(cleanAm(e2.am)); p.dx.push(...codesOf(e2.dx));
      people.push(p);
    }
    // 2) ασθενείς του ΝΠΣ που δεν είναι στο Παραδόχαρτο
    nps.forEach((r, i) => {
      if (used.has(i) || !String(r.name || "").trim()) return;
      const e = enrOf(r.name);
      people.push({ full: String(r.name).replace(/\s+/g, " ").trim(), father: "", amka: [cleanAmka(r.amka), cleanAmka(e.amka)], am: [cleanAm(r.am), cleanAm(e.am)], dx: codesOf(e.dx), src: "nps" });
    });
    // 3) τελικές γραμμές — ένας ανά ΑΜΚΑ/ΑΜ
    const rows = [], skipped = [], conflicts = [], seen = {};
    for (const p of people) {
      const amkas = [...new Set(p.amka.filter(Boolean))], ams = [...new Set(p.am.filter(Boolean))];
      if (amkas.length > 1 || ams.length > 1) { conflicts.push({ name: p.full, why: amkas.length > 1 ? "διαφορετικοί ΑΜΚΑ" : "διαφορετικοί ΑΜ" }); continue; }
      const amka = amkas[0] || "", am = ams[0] || "";
      if (!amka && !am) { skipped.push({ name: p.full, why: "χωρίς ΑΜΚΑ/ΑΜ" }); continue; }
      const id = amka ? "K" + amka : "M" + am;
      const nm = split(p.full); const dx = [...new Set(p.dx)];
      if (seen[id] !== undefined) {   // ίδιος άνθρωπος δύο φορές (π.χ. Παραδόχαρτο και ΝΠΣ με λίγο άλλο όνομα)
        const o = rows[seen[id]]; o.dx = [...new Set(o.dx.split(",").filter(Boolean).concat(dx))].join(",");
        if (!o.am && am) o.am = am; if (!o.amka && amka) o.amka = amka; if (!o.father && p.father) o.father = p.father;
        if (!o.firstname && nm.firstname) { o.surname = nm.surname; o.firstname = nm.firstname; }
        if (p.surg && p.surg.length) o.surg = (o.surg || []).concat(p.surg);
        if (p.hist && !o.hist) o.hist = p.hist;
        continue;
      }
      seen[id] = rows.length;
      rows.push(Object.assign({ surname: nm.surname, firstname: nm.firstname, father: p.father, amka, am, dx: dx.join(","), src: p.src }, p.surg && p.surg.length ? { surg: p.surg } : {}, p.hist ? { hist: p.hist } : {}));
    }
    return { rows, skipped, conflicts, n: { pdx: pdx.length, nps: nps.length } };
  }
  // γραμμές γεγονότων «history» για POST /api/events (ένα ανά ασθενή — ενημερώνεται σε κάθε πέρασμα)
  function historyEvents(rows) { return rows.filter((r) => r.hist).map((r) => ({ patient: { surname: r.surname, firstname: r.firstname, amka: r.amka || null, am: r.am || null }, kind: "history", date: r.hist.at, title: "Ιστορικό (Παραδόχαρτο)", body: r.hist, src: "pdx", ext_id: "pdxhist|" + (r.amka ? "K" + r.amka : "M" + r.am) })); }
  const api = { buildCurrentRows, historyEvents, histOf, egfr, sameName, codesOf, amkaOk, pdxFull };
  if (typeof module === "object" && module.exports) module.exports = api;
  root.pdxCur = api;
})(typeof self !== "undefined" ? self : this);
