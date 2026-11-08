// 🗓 8/11: Πρόγραμμα χειρουργείου — σύνθεση στο Παραδόχαρτο, ΑΚΡΙΒΩΣ στη μορφή των Excel της κλινικής.
// Μέρη: (α) ORP.build(prog) → .xlsx (ίδια στυλ με το «9.11.2026») + HTML εκτύπωσης, (β) σελίδα επεξεργασίας (window.pdxOrProg).
// Δεδομένα (έτοιμα σετ, κωδικοί, στυλ Excel): orprog-data.js (window.ORP_DATA). Το πρόχειρο αποθηκεύεται ΚΡΥΠΤΟΓΡΑΦΗΜΕΝΟ (κλειδί κλινικής)
// στο oncall_labs «orprog|<ημερομηνία>»· η «Υποβολή» στέλνει εργασία «orsubmit» στον Μεσαίο: Excel στον φάκελο δικτύου + εκτύπωση.
(function (root) {
  "use strict";
  const D = () => root.ORP_DATA || (typeof require === "function" ? require("./orprog-data.js") : {});
  const XE = (s) => String(s == null ? "" : s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const HE = XE;
  const pad = (n) => String(n).padStart(2, "0");
  const ROOM_ORDER = (r) => /Α$/.test(r) ? -1 : +r;   // απογευματινή τελευταία
  const roomLabel = (r) => /Α$/.test(r) ? "ΑΙΘΟΥΣΑ " + r.replace(/Α$/, "") + " (ΑΠΟΓΕΥΜΑ)" : "ΑΙΘΟΥΣΑ " + r;
  // ημερομηνία προγράμματος: αύριο· Παρασκευή → Δευτέρα (+3), Σάββατο → Δευτέρα (+2)
  function defDate(now) { const d = new Date(now || Date.now()); const wd = d.getDay(); d.setDate(d.getDate() + (wd === 5 ? 3 : wd === 6 ? 2 : 1)); return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
  const dShort = (iso) => { const [y, m, d] = String(iso).split("-"); return +d + "." + +m + "." + String(y).slice(2); };   // 9.11.26
  const MONTHS = ["ΙΑΝΟΥΑΡΙΟΣ", "ΦΕΒΡΟΥΑΡΙΟΣ", "ΜΑΡΤΙΟΣ", "ΑΠΡΙΛΙΟΣ", "ΜΑΙΟΣ", "ΙΟΥΝΙΟΣ", "ΙΟΥΛΙΟΣ", "ΑΥΓΟΥΣΤΟΣ", "ΣΕΠΤΕΜΒΡΙΟΣ", "ΟΚΤΩΒΡΙΟΣ", "ΝΟΕΜΒΡΙΟΣ", "ΔΕΚΕΜΒΡΙΟΣ"];
  const monthDir = (iso) => { const [y, m] = String(iso).split("-"); return MONTHS[+m - 1] + " " + String(y).slice(2); };   // «ΝΟΕΜΒΡΙΟΣ 26»
  const fileName = (iso, kind) => (kind && kind !== "ΠΡΟΓΡΑΜΜΑ" ? kind.replace(/ ΠΡΟΓΡΑΜΜΑ$/, "") + " " : "") + dShort(iso) + ".xlsx";
  // επέμβαση με πλευρά: η πλευρά μπαίνει ΠΡΙΝ από τον πρώτο κωδικό («ΟΥΡΗΤΗΡΟΛΙΘΟΤΡΙΨΙΑ ΑΡ 356.02.05»)
  function withSide(t, side) { t = String(t || "").trim(); if (!side) return t; const m = /\s\d{3}\.\d/.exec(t); return m ? t.slice(0, m.index) + " " + side + t.slice(m.index) : t + " " + side; }
  function procText(p) { return [withSide(p.t, p.side), p.t2 ? withSide(p.t2, p.side2) : ""].filter(Boolean).join(" + "); }
  // σειρά αιθουσών: πρωινές με φθίνοντα αριθμό, μετά η απογευματινή
  function rooms(prog) {
    const R = [...new Set([...(prog.rooms || []), ...(prog.pts || []).map((p) => p.room)].filter(Boolean))];
    R.sort((a, b) => ROOM_ORDER(b) - ROOM_ORDER(a) || (/Α$/.test(a) ? 1 : 0) - (/Α$/.test(b) ? 1 : 0));
    return R.map((r) => ({ r, label: roomLabel(r), pts: (prog.pts || []).filter((p) => p.room === r).sort((a, b) => (a.n || 0) - (b.n || 0)) }));
  }

  // ===== .xlsx (Office Open XML) με τα στυλ του προτύπου =====
  const S = { a1: 118, h2: 119, date: 120, resp: 121, hd: 37, hdw: 38, rmL: 111, rmM: 112, rmR: 113, endA: 59, endB: 53, endC: 52, endM: 51, endH: 58, dir: 42, foot: 54, small: 2,
    p1: [90, 114, 116, 55, 64, 56, 57, 69], p2: [91, 115, 117, 99, 100, 100, 101, 70] };
  const COLS = "ABCDEFGH";
  function cell(ref, s, v) { return '<c r="' + ref + '" s="' + s + '"' + (v === "" || v == null ? "/>" : ' t="inlineStr"><is><t xml:space="preserve">' + XE(v) + "</t></is></c>"); }
  // εμπλουτισμένο κείμενο: ΑΜ (8ψήφιο) και κωδικοί επεμβάσεων με έντονα, όπως στα αρχεία της κλινικής
  const BOLD = /\b\d{8}\b|\b\d{3}\.\d{1,2}(?:\.\d{1,2})*\b/g;
  const RPR = '<rPr><b/><sz val="11"/><color theme="1"/><rFont val="Arial"/><family val="2"/><charset val="161"/></rPr>';
  function rich(ref, s, v) {
    v = String(v == null ? "" : v); if (!v) return cell(ref, s); const runs = []; let k = 0; let m; BOLD.lastIndex = 0;
    while ((m = BOLD.exec(v))) { if (m.index > k) runs.push('<r><t xml:space="preserve">' + XE(v.slice(k, m.index)) + "</t></r>"); runs.push("<r>" + RPR + '<t xml:space="preserve">' + XE(m[0]) + "</t></r>"); k = m.index + m[0].length; }
    if (!runs.length) return cell(ref, s, v);
    if (k < v.length) runs.push('<r><t xml:space="preserve">' + XE(v.slice(k)) + "</t></r>");
    return '<c r="' + ref + '" s="' + s + '" t="inlineStr"><is>' + runs.join("") + "</is></c>";
  }
  const hb = (v) => HE(v).replace(BOLD, (x) => "<b>" + x + "</b>");
  function plain(ref, v) { return v === "" || v == null ? "" : '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + XE(v) + "</t></is></c>"; }
  function sheetXml(prog) {
    const d = D(), rows = [], merges = ["A1:H1", "A2:C3", "D2:D3", "E2:E3", "F2:H2", "F3:H3"];
    const side = {};   // δεξιά (K:N): τα έτοιμα σετ, όπως στα αρχεία της κλινικής
    { let r = 2; const groups = []; for (const x of [...(d.lib || []), ...(d.extra || [])]) { let g = groups.find((y) => y.g === x.g); if (!g) groups.push(g = { g: x.g, L: [] }); g.L.push(x); }
      for (const g of groups) { (side[r] = side[r] || []).push(plain("K" + r, g.g)); r += 2; for (const x of g.L) { (side[r] = side[r] || []).push(plain("K" + r, x.dx), plain("L" + r, x.an), plain("M" + r, x.bl), plain("N" + r, x.ca)); (side[r + 1] = side[r + 1] || []).push(plain("K" + (r + 1), x.t)); merges.push("K" + (r + 1) + ":N" + (r + 1)); r += 3; } r++; } }
    const row = (r, ht, cells) => rows.push('<row r="' + r + '"' + (ht ? ' ht="' + ht + '" customHeight="1"' : "") + ">" + cells.join("") + (side[r] || []).join("") + "</row>");
    row(1, 13.5, COLS.split("").map((c, i) => cell(c + 1, S.a1, i ? "" : d.a1 || 'Γ Ν Α  "Λ Α Ϊ  Κ Ο"')));
    row(2, 13.5, [cell("A2", S.h2, "ΟΥΡΟΛΟΓΙΚΗ ΚΛΙΝΙΚΗ"), cell("B2", S.h2), cell("C2", S.h2), cell("D2", S.h2, " " + (prog.kind || "ΠΡΟΓΡΑΜΜΑ")), cell("E2", S.date, dShort(prog.date)), cell("F2", S.resp, "ΥΠΕΥΘΥΝΟΣ ΠΡΟΓΡΑΜΜΑΤΟΣ:"), cell("G2", S.resp), cell("H2", S.resp)]);
    row(3, 15.75, [cell("A3", S.h2), cell("B3", S.h2), cell("C3", S.h2), cell("D3", S.h2), cell("E3", S.date), cell("F3", S.h2, prog.resp || ""), cell("G3", S.h2), cell("H3", S.h2)]);
    row(4, 15.75, ["ΘΑΛ", "ΗΜ. ΕΙΣ.", "ΟΝΟΜΑΤΕΠΩΝΥΜΟ", "ΕΠΕΜΒΑΣΗ", "ΑΝΑΙΣΘΗΣΙΑ", "ΑΙΜΑ", "c-ARM", "ΧΕΙΡΟΥΡΓΟΣ"].map((t, i) => cell(COLS[i] + 4, i === 1 || i === 7 ? S.hdw : S.hd, t)));
    let r = 5;
    for (const R of rooms(prog)) {
      row(r, 15.75, COLS.split("").map((c, i) => cell(c + r, i === 0 ? S.rmL : i === 7 ? S.rmR : S.rmM, i ? "" : R.label))); merges.push("A" + r + ":H" + r); r++;
      for (const p of R.pts) {
        const v1 = [p.bed, p.adm, p.nm, p.dx, p.an, p.bl, p.ca, p.surg], v2 = ["", "", "", procText(p), "", "", "", ""];
        row(r, 15, COLS.split("").map((c, i) => (i === 2 ? rich : cell)(c + r, S.p1[i], v1[i])));
        row(r + 1, 21, COLS.split("").map((c, i) => (i === 3 ? rich : cell)(c + (r + 1), S.p2[i], v2[i])));
        merges.push("A" + r + ":A" + (r + 1), "B" + r + ":B" + (r + 1), "C" + r + ":C" + (r + 1), "H" + r + ":H" + (r + 1), "D" + (r + 1) + ":G" + (r + 1));
        r += 2;
      }
    }
    row(r, 14.25, [cell("A" + r, S.endA), cell("B" + r, S.endB), cell("C" + r, S.endC), cell("D" + r, S.endM), cell("E" + r, S.endM), cell("F" + r, S.endM), cell("G" + r, S.endM), cell("H" + r, S.endH)]); r++;
    row(r, 15, [cell("A" + r, S.foot), cell("B" + r, S.small), cell("H" + r, S.dir, "Ο ΔΙΕΥΘΥΝΤΗΣ")]);
    const last = r;
    for (const k of Object.keys(side).map(Number).filter((k) => k > last).sort((a, b) => a - b)) rows.push('<row r="' + k + '">' + side[k].join("") + "</row>");
    const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      + '<sheetPr><pageSetUpPr fitToPage="0"/></sheetPr><dimension ref="A1:N' + Math.max(last, ...Object.keys(side).map(Number)) + '"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15.75"/>'
      + '<cols><col min="1" max="1" width="12.7109375" style="2" customWidth="1"/><col min="2" max="2" width="12.85546875" style="2" customWidth="1"/><col min="3" max="3" width="46.5703125" style="1" customWidth="1"/><col min="4" max="4" width="45.28515625" style="1" customWidth="1"/><col min="5" max="5" width="14.7109375" style="1" customWidth="1"/><col min="6" max="6" width="8.140625" style="1" customWidth="1"/><col min="7" max="7" width="8.85546875" style="1" customWidth="1"/><col min="8" max="8" width="27.42578125" style="1" customWidth="1"/><col min="10" max="10" width="7.7109375" customWidth="1"/><col min="11" max="11" width="41.7109375" customWidth="1"/><col min="12" max="12" width="13.85546875" customWidth="1"/><col min="14" max="14" width="9.28515625" customWidth="1"/></cols>'
      + "<sheetData>" + rows.join("") + "</sheetData>"
      + '<mergeCells count="' + merges.length + '">' + merges.map((m) => '<mergeCell ref="' + m + '"/>').join("") + "</mergeCells>"
      + '<pageMargins left="0.15748031496062992" right="0.11811023622047245" top="0.15748031496062992" bottom="0.15748031496062992" header="0.31496062992125984" footer="0"/><pageSetup paperSize="9" scale="80" orientation="landscape"/></worksheet>';
    return { xml, last };
  }
  // ZIP χωρίς συμπίεση (STORED) + CRC32 — αρκεί για .xlsx
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = (u) => { let c = 0xFFFFFFFF; for (let i = 0; i < u.length; i++) c = CRC[(c ^ u[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  function zip(files) {
    const enc = new TextEncoder(), parts = [], cen = []; let off = 0;
    const u16 = (v) => [v & 255, (v >>> 8) & 255], u32 = (v) => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255];
    for (const [name, data] of files) {
      const nb = enc.encode(name), db = typeof data === "string" ? enc.encode(data) : data, c = crc32(db);
      const head = [0x50, 0x4b, 3, 4, ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(c), ...u32(db.length), ...u32(db.length), ...u16(nb.length), ...u16(0)];
      parts.push(new Uint8Array(head), nb, db);
      cen.push(new Uint8Array([0x50, 0x4b, 1, 2, ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(c), ...u32(db.length), ...u32(db.length), ...u16(nb.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(off)]), nb);
      off += head.length + nb.length + db.length;
    }
    const csz = cen.reduce((a, u) => a + u.length, 0);
    const end = new Uint8Array([0x50, 0x4b, 5, 6, ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length), ...u32(csz), ...u32(off), ...u16(0)]);
    const all = [...parts, ...cen, end], out = new Uint8Array(all.reduce((a, u) => a + u.length, 0)); let k = 0; for (const u of all) { out.set(u, k); k += u.length; }
    return out;
  }
  function xlsx(prog) {
    const d = D(), sh = sheetXml(prog);
    const ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';
    const rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    const wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$H$' + sh.last + "</definedName></definedNames></workbook>";
    const wrels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
    return zip([["[Content_Types].xml", ct], ["_rels/.rels", rels], ["xl/workbook.xml", wb], ["xl/_rels/workbook.xml.rels", wrels], ["xl/styles.xml", d.styles], ["xl/theme/theme1.xml", d.theme], ["xl/worksheets/sheet1.xml", sh.xml]]);
  }
  // ===== HTML εκτύπωσης (A4 οριζόντιο, όπως η προεπισκόπηση του Excel στο 80%) =====
  function html(prog, opt) {
    opt = opt || {};
    const W = [12.71, 12.86, 46.57, 45.29, 14.71, 8.14, 8.86, 27.43], tw = W.reduce((a, b) => a + b, 0);
    let h = '<table class="orp"><colgroup>' + W.map((w) => '<col style="width:' + (w / tw * 100).toFixed(2) + '%">').join("") + "</colgroup>";
    h += '<tr><td colspan="8" class="b c">Γ Ν Α &nbsp;"Λ Α Ϊ &nbsp;Κ Ο"</td></tr>';
    h += '<tr><td colspan="3" rowspan="2" class="b c">ΟΥΡΟΛΟΓΙΚΗ ΚΛΙΝΙΚΗ</td><td rowspan="2" class="b c">' + HE(prog.kind || "ΠΡΟΓΡΑΜΜΑ") + '</td><td rowspan="2" class="b c">' + HE(dShort(prog.date)) + '</td><td colspan="3" class="l">ΥΠΕΥΘΥΝΟΣ ΠΡΟΓΡΑΜΜΑΤΟΣ:</td></tr>';
    h += '<tr><td colspan="3" class="b c">' + HE(prog.resp || "") + "</td></tr>";
    h += '<tr class="hd">' + ["ΘΑΛ", "ΗΜ. ΕΙΣ.", "ΟΝΟΜΑΤΕΠΩΝΥΜΟ", "ΕΠΕΜΒΑΣΗ", "ΑΝΑΙΣΘΗΣΙΑ", "ΑΙΜΑ", "c-ARM", "ΧΕΙΡΟΥΡΓΟΣ"].map((t) => "<td>" + t + "</td>").join("") + "</tr>";
    for (const R of rooms(prog)) {
      h += '<tr class="rm"><td colspan="8">' + HE(R.label) + "</td></tr>";
      for (const p of R.pts) {
        h += '<tr class="p1"><td rowspan="2" class="b">' + HE(p.bed) + '</td><td rowspan="2" class="b">' + HE(p.adm) + '</td><td rowspan="2">' + hb(p.nm) + "</td><td>" + HE(p.dx) + "</td><td>" + HE(p.an) + '</td><td class="b">' + HE(p.bl) + '</td><td class="b">' + HE(p.ca) + '</td><td rowspan="2" class="b">' + HE(p.surg) + "</td></tr>";
        h += '<tr class="p2"><td colspan="4">' + hb(procText(p)) + "</td></tr>";
      }
    }
    h += '<tr class="end"><td colspan="8"></td></tr><tr><td colspan="7" style="border:0"></td><td class="b c" style="border:0;font-size:14pt">Ο ΔΙΕΥΘΥΝΤΗΣ</td></tr></table>';
    const css = ".orp{border-collapse:collapse;width:100%;table-layout:fixed;font:11pt Arial,Helvetica,sans-serif;color:#000}.orp td{border:1px solid #000;padding:2px 4px;text-align:center;vertical-align:middle;word-wrap:break-word}"
      + ".orp .b{font-weight:700}.orp .l{text-align:left;border:0}.orp tr:nth-child(1) td,.orp tr:nth-child(2) td,.orp tr:nth-child(3) td{border:0;font-weight:700}.orp tr:nth-child(2) td.l{font-weight:400}"
      + ".orp .hd td{background:#C0C0C0;font-weight:700}.orp .rm td{background:#BFBFBF;font-weight:700;text-align:center}.orp .p2 td{font-size:11pt}.orp .end td{border:0;height:8px}";
    if (opt.bodyOnly) return "<style>" + css + "</style>" + h;
    return '<!doctype html><html><head><meta charset="utf-8"><title>' + HE((prog.kind || "ΠΡΟΓΡΑΜΜΑ") + " " + dShort(prog.date)) + "</title><style>@page{size:A4 landscape;margin:4mm 3mm}body{margin:0}" + css + "</style></head><body>" + h + "</body></html>";
  }
  const ORP = { defDate, dShort, monthDir, fileName, withSide, procText, rooms, roomLabel, sheetXml, xlsx, html, zip, crc32 };
  root.ORP = ORP;
  if (typeof module !== "undefined") module.exports = ORP;
})(typeof window !== "undefined" ? window : globalThis);

// ===================== Σελίδα «🗓 Πρόγραμμα χειρουργείου» (μόνο στο Παραδόχαρτο) =====================
(function () {
  "use strict";
  if (typeof window === "undefined" || typeof document === "undefined") return;
  const ORP = window.ORP, E = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const UP = (s) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase();
  const NZ = (s) => UP(s).replace(/[^A-ZΑ-Ω0-9]/g, "");
  const say = (t) => { try { if (typeof window.toast === "function") return window.toast(t); } catch (e) {} const d = document.createElement("div"); d.textContent = t; d.style.cssText = "position:fixed;left:50%;bottom:90px;transform:translateX(-50%);background:#1C1C1E;color:#fff;padding:10px 16px;border-radius:12px;z-index:99999;font:600 14px -apple-system,system-ui;max-width:86vw;text-align:center"; document.body.appendChild(d); setTimeout(() => d.remove(), 2600); };
  const SB = () => (typeof sb6 !== "undefined" ? sb6 : null);
  const CK = () => window.pdxCK;
  const BOT = "bmw51nismmo82h", PRINTER = "Lexmark MX530 Series XL";
  const ROOMS_ALL = ["10", "9", "8", "8Α"];
  const RESIDENTS = ["ΓΑΡΥΦΑΛΛΟΣ", "ΓΙΑΓΚΟΣ", "ΘΕΟΔΩΡΟΛΕΑΣ", "ΚΑΛΦΑΣ", "ΚΑΛΦΟΥΝΤΖΟΣ", "ΜΑΝΟΣ", "ΜΑΥΡΟΥΔΗΣ", "ΠΑΝΑΓΟΠΟΥΛΟΣ"];
  const ANS = ["ΓΕΝΙΚΗ", "ΤΟΠΙΚΗ"], BLOOD = ["", "300", "600", "900", "1200", "1500"], SIDES = ["", "ΑΡ", "ΔΕ", "ΑΜΦΩ"];
  const PK = () => { const d = window.ORP_DATA || {}; return [...(d.lib || []), ...(d.extra || [])]; };
  const FULL = () => window.THERAPIST_FULL || { "Αν": "Αναστασίου", "Αδ": "Αδαμάκης", "Στρ": "Στραβοδήμος", "Φρ": "Φραγκιάδης", "Μουλ": "Μουλαβασίλης", "Λεβ": "Λεβής", "Δουρ": "Δουρούμης" };
  const surgeonOf = (t) => { const f = FULL()[t]; return f ? "κ " + UP(f) : ""; };
  const SURGEONS = () => Object.values(FULL()).map((f) => "κ " + UP(f));
  // συντομογραφίες επεμβάσεων του Παραδόχαρτου → έτοιμο σετ (κείμενο που περιέχει)
  const MAP = { "TURP": "ΔΙΟΥΡΗΘΡΙΚΗ ΠΡΟΣΤΑΤΟΥ", "TURB": "ΔΙΟΥΡΗΘΡΙΚΗ ΚΥΣΤΗΣ", "Tur αυχ": "ΔΙΟΥΡΗΘΡΙΚΗ ΑΥΧΕΝΑ", "ΟΛΨ": "ΟΥΡΗΤΗΡΟΛΙΘΟΤΡΙΨΙΑ", "ΝΛΨ": "ΝΕΦΡΟΛΙΘΟΤΡΙΨΙΑ", "τοπ. PT": "ΤΟΠΟΘΕΤΗΣΗ PIG-TAIL", "αλλ. PT": "ΑΛΛΑΓΗ PIGTAIL", "τοπ. Allium": "ΤΟΠΟΘΕΤΗΣΗ ALLIUM", "Optilume": "ΔΙΑΣΤΟΛΗ ΜΕ OPTILUME", "Rezum": "REZUM", "Nx": "ΝΕΦΡΕΚΤΟΜΗ 355.04.04.00", "pNx": "ΜΕΡΙΚΗ ΝΕΦΡΕΚΤΟΜΗ", "Lap Nx": "LAP ΝΕΦΡΕΚΤΟΜΗ", "NUT": "ΝΕΦΡΟΟΥΡΗΤΗΡΕΚΤΟΜΗ ΜΕ 2 ΤΟΜΕΣ", "Lap NUT": "LAP ΝΕΦΡΟΟΥΡΗΤΗΡΕΚΤΟΜΗ", "RC Bricker": "BRICKER", "RC στομίες": "ΡΙΖΙΚΗ ΚΥΣΤΕΚΤΟΜΗ 357.06.02.00 / + ΣΤΟΜΙΕΣ", "RRP": "ΡΙΖΙΚΗ ΠΡΟΣΤΑΤΕΚΤΟΜΗ + ΛΕΜΦΑΔΕΝΕΚΤΟΜΗ", "Ορχεκτομή": "ΟΡΧΕΚΤΟΜΗ", "Υδροκήλη": "ΥΔΡΟΚΗΛΗΣ", "Κιρσοκήλη": "ΚΙΡΣΟΚΗΛΗΣ", "Περιτομή": "ΠΕΡΙΤΟΜΗ", "AUS": "ΤΕΧΝΗΤΟΥ ΣΦΙΓΚΤΗΡΑ", "Fusion": "FUSION", "DVIU": "ΟΥΡΗΘΡΟΤΟΜΗ", "cysto-b": "ΚΥΣΤΕΟΣΚΟΠΗΣΗ + ΒΙΟΨΙΕΣ" };
  const pkFor = (s) => { const k = MAP[s]; if (!k) return null; return PK().find((x) => x.t.indexOf(k) >= 0) || null; };
  const sideOf = (s) => ({ "Δεξιά": "ΔΕ", "Αριστερά": "ΑΡ", "άμφω": "ΑΜΦΩ" }[s] || "");
  // ασθενείς του Παραδόχαρτου (θάλαμοι + εκτός θαλάμου)
  function inpatients() {
    const L = []; try {
      const D = typeof data !== "undefined" ? data : {}, X = typeof extraPatients !== "undefined" ? extraPatients : {};
      for (const [k, d] of [...Object.entries(D), ...Object.entries(X)]) { if (!d || !String(d.name || "").trim()) continue; L.push({ k, d }); }
    } catch (e) {}
    const pre = (x) => x.d.pod === "Προς Χ/Ο" ? 0 : 1;
    return L.sort((a, b) => pre(a) - pre(b) || String(a.k).localeCompare(String(b.k), "el", { numeric: true }));
  }
  function fromPatient(k, d) {
    const info = (window.pdxNpsInfo && window.pdxNpsInfo(d)) || {};
    const full = UP(info.full || d.name || ""), am = info.am || d.am || "", fa = UP(info.father || d.father || "");
    const pk = pkFor(d.surgery) || null;
    const m = /^(\d+|352-Α)-(\d+)$/.exec(k);
    return { id: "p" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), key: k, bed: m ? m[1].replace("-Α", "Α") + "." + m[2] : "", adm: String(d.admitDate || "").replace(/^(\d{1,2})[\/.](\d{1,2}).*$/, (x, a, b) => a.padStart(2, "0") + "/" + b.padStart(2, "0")),
      nm: [full, fa ? "ΤΟΥ " + fa : "", am].filter(Boolean).join(" "), dx: pk ? pk.dx : UP(d.dx || ""), an: pk ? pk.an : "ΓΕΝΙΚΗ", bl: pk ? pk.bl : "", ca: pk ? pk.ca : "", t: pk ? pk.t : "", side: sideOf(d.side), t2: "", side2: "", surg: surgeonOf(d.therapist) };
  }
  // ---- αποθήκευση ----
  const keyOf = (date) => "orprog|" + date;
  async function loadCfg() { try { const r = await SB().from("oncall_labs").select("data").eq("name_key", "orprog|cfg").maybeSingle(); return (r.data && r.data.data) || {}; } catch (e) { return {}; } }
  async function saveCfg(c) { await SB().from("oncall_labs").upsert({ name_key: "orprog|cfg", name: "Πρόγραμμα χειρουργείου", pid: null, data: c, updated_at: new Date().toISOString() }, { onConflict: "name_key" }); }
  async function loadDraft(date) { const r = await SB().from("oncall_labs").select("data,updated_at").eq("name_key", keyOf(date)).maybeSingle(); if (r.error) throw new Error(r.error.message); const env = r.data && r.data.data; if (!env) return null; return await CK().decrypt(env.enc || env); }
  async function saveDraft(P) { P.at = Date.now(); P.by = String(window.authName || "").replace(/@paradoxarto\.app$/, ""); const enc = await CK().encrypt(P); const r = await SB().from("oncall_labs").upsert({ name_key: keyOf(P.date), name: "Πρόγραμμα χειρουργείου " + ORP.dShort(P.date), pid: null, data: { enc, st: P.status || "draft", n: (P.pts || []).length }, updated_at: new Date().toISOString() }, { onConflict: "name_key" }); if (r.error) throw new Error(r.error.message); }
  async function sealB64(u8) { const s = await CK().sealBytes(u8); let b = ""; for (let i = 0; i < s.length; i += 0x8000) b += String.fromCharCode.apply(null, s.subarray(i, i + 0x8000)); return btoa(b); }
  function printLocal(P) { const w = window.open("", "_blank"); if (!w) { say("Επιτρέψτε τα αναδυόμενα για εκτύπωση"); return; } w.document.write(ORP.html(P)); w.document.close(); setTimeout(() => { try { w.focus(); w.print(); } catch (e) {} }, 400); }

  window.pdxOrProg = async function (date0) {
    let P = { date: date0 || ORP.defDate(), kind: "ΠΡΟΓΡΑΜΜΑ", resp: "", rooms: ["10", "8"], pts: [], status: "draft" }, cfg = {}, dirty = false, busy = "", loaded = false;
    const can = ["resident", "attending", "admin"].indexOf(window.authRole) >= 0;
    const css = '<style>#orpPg .chip{display:inline-flex;align-items:center;min-height:36px;padding:6px 12px;margin:3px;border-radius:18px;border:1px solid #C7CDD4;background:#fff;font:600 14px -apple-system,system-ui;color:#1F2933}#orpPg .chip.on{background:#0A84FF;border-color:#0A84FF;color:#fff}'
      + '#orpPg .card{margin:8px 12px;padding:10px 12px;border-radius:12px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.08);font:14px -apple-system,system-ui}#orpPg .card b{font-weight:700}#orpPg .rm{margin:16px 16px 4px;font:700 13px -apple-system,system-ui;color:#6B7280;letter-spacing:.4px}'
      + '#orpPg .mini{border:0;background:#EEF2F6;border-radius:8px;min-width:38px;min-height:34px;font-size:15px;margin-left:4px}#orpPg .bar{position:sticky;bottom:0;display:flex;gap:8px;padding:10px 12px calc(10px + env(safe-area-inset-bottom));background:rgba(242,242,247,.96);backdrop-filter:blur(8px);border-top:1px solid #DDE1E6}'
      + '#orpPg .bar button{flex:1;min-height:46px;border:0;border-radius:12px;font:700 14px -apple-system,system-ui}#orpPg input,#orpPg select{font-size:16px;padding:8px 10px;border:1px solid #C7CDD4;border-radius:10px;background:#fff}'
      + 'html.pdx-night #orpPg .card{background:#1C1C1E;color:#fff}html.pdx-night #orpPg .chip{background:#2C2C2E;color:#fff;border-color:#3A3A3C}html.pdx-night #orpPg .chip.on{background:#0A84FF}html.pdx-night #orpPg .bar{background:rgba(28,28,30,.96);border-color:#38383A}html.pdx-night #orpPg input,html.pdx-night #orpPg select{background:#2C2C2E;color:#fff;border-color:#3A3A3C}html.pdx-night #orpPg .mini{background:#2C2C2E;color:#fff}</style>';
    const chip = (attr, v, on, t) => '<button class="chip' + (on ? " on" : "") + '" ' + attr + '="' + E(v) + '">' + (t == null ? E(v || "—") : t) + "</button>";
    const draw = () => {
      if (!loaded) return css + '<div id="orpPg"><div class="card">Φόρτωση…</div></div>';
      let h = css + '<div id="orpPg">';
      h += '<div class="card"><div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap"><input type="date" data-f="date" value="' + E(P.date) + '"><select data-f="kind">' + ["ΠΡΟΓΡΑΜΜΑ", "ΕΚΤΑΚΤΟ ΠΡΟΓΡΑΜΜΑ", "ΣΥΜΠΛΗΡΩΜΑΤΙΚΟ ΠΡΟΓΡΑΜΜΑ"].map((k) => "<option" + (P.kind === k ? " selected" : "") + ">" + k + "</option>").join("") + "</select>"
        + '<span style="color:#6B7280;font-size:13px">' + (P.status === "submitted" ? "✅ υποβλήθηκε" + (P.file ? " · " + E(P.file) : "") : dirty ? "● μη αποθηκευμένο" : P.at ? "💾 πρόχειρο " + new Date(P.at).toLocaleTimeString("el-GR", { hour: "2-digit", minute: "2-digit" }) + (P.by ? " · " + E(P.by) : "") : "νέο") + "</span></div>";
      // 8/11: οι 8 ειδικευόμενοι (από τα προγράμματα του φακέλου) σε λίστα επιλογής· «＋ άλλος…» μόνο για νέο όνομα
      const RES = [...new Set([...RESIDENTS, ...(cfg.residents || []), ...(P.resp ? [P.resp] : [])])];
      h += '<div style="margin-top:8px;font-weight:700">Υπεύθυνος προγράμματος</div><select data-f="resp" style="width:100%;margin-top:4px"><option value="">— διάλεξε ειδικευόμενο —</option>' + RES.map((n) => "<option" + (P.resp === n ? " selected" : "") + ">" + E(n) + "</option>").join("") + '<option value="__add">＋ άλλος…</option></select>';
      h += (P.pts.length && window.pdxXrQueue ? '<button class="chip" data-act="xr" style="margin-top:8px;background:#0F172A;color:#fff;border-color:#0F172A">🫁 Ακτινογραφίες θώρακα (' + P.pts.length + ') → Νίκος CT Viewer</button>' : "");
      h += '<div style="margin-top:8px;font-weight:700">Αίθουσες</div><div>' + ROOMS_ALL.map((r) => chip("data-room", r, P.rooms.indexOf(r) >= 0, E(ORP.roomLabel(r).replace("ΑΙΘΟΥΣΑ ", "")))).join("") + "</div></div>";
      for (const R of ORP.rooms(P)) {
        h += '<div class="rm">' + E(R.label) + " · " + R.pts.length + "</div>";
        R.pts.forEach((p, i) => {
          h += '<div class="card" data-open="' + E(p.id) + '"><div style="display:flex;align-items:flex-start;gap:6px"><div style="flex:1;min-width:0"><b>' + (i + 1) + ". " + E(p.nm || "—") + '</b><div style="color:#6B7280;font-size:13px">' + E([p.bed, p.adm && "εισ. " + p.adm].filter(Boolean).join(" · ")) + "</div>"
            + '<div style="margin-top:3px">' + E(p.dx || "") + '</div><div style="font-weight:600">' + E(ORP.procText(p) || "⚠ χωρίς επέμβαση") + '</div><div style="color:#6B7280;font-size:13px">' + E([p.an, p.bl && "αίμα " + p.bl, p.ca && "c-ARM", p.surg].filter(Boolean).join(" · ")) + "</div></div>"
            + '<span style="display:flex;flex-direction:column"><button class="mini" data-up="' + E(p.id) + '">▲</button><button class="mini" data-dn="' + E(p.id) + '" style="margin-top:4px">▼</button></span></div></div>';
        });
        h += '<div style="margin:2px 12px"><button class="chip" data-add="' + E(R.r) + '">＋ Ασθενής στην ' + E(R.label.replace("ΑΙΘΟΥΣΑ", "αίθουσα")) + "</button></div>";
      }
      if (!P.rooms.length) h += '<div class="card">Διαλέξτε αίθουσες.</div>';
      h += '<div class="card" style="overflow:auto;padding:6px"><div style="font-weight:700;margin:4px 6px">Προεπισκόπηση</div><div style="zoom:.55;min-width:1100px">' + ORP.html(P, { bodyOnly: true }) + "</div></div><div style=\"height:20px\"></div>";
      h += '<div class="bar"><button data-act="save" style="background:#EEF2F6;color:#0A84FF">💾 Πρόχειρο</button><button data-act="print" style="background:#EEF2F6;color:#0A84FF">🖨 Εκτύπωση</button><button data-act="submit" style="background:#34C759;color:#fff"' + (can ? "" : " disabled") + ">✅ Υποβολή</button></div></div>";
      return h;
    };
    const pg = window.iosPage("Πρόγραμμα χειρουργείου", "Ρυθμίσεις", draw);
    const redraw = () => { const y = pg.el.querySelector(".ios-body") && pg.el.querySelector(".ios-body").scrollTop; pg.redraw(); if (y) pg.el.querySelector(".ios-body").scrollTop = y; };
    const load = async (date) => {
      loaded = false; redraw();
      try { if (CK() && CK().ready) await CK().ready(); cfg = await loadCfg(); const d = await loadDraft(date); P = d || { date, kind: "ΠΡΟΓΡΑΜΜΑ", resp: "", rooms: ["10", "8"], pts: [], status: "draft" }; dirty = false; }
      catch (e) { say("⚠ " + (e.message || e)); P = { date, kind: "ΠΡΟΓΡΑΜΜΑ", resp: "", rooms: ["10", "8"], pts: [], status: "draft" }; }
      loaded = true; redraw();
    };
    const touch = () => { dirty = true; if (P.status === "submitted") P.status = "changed"; redraw(); };
    const renum = (room) => P.pts.filter((p) => p.room === room).sort((a, b) => (a.n || 0) - (b.n || 0)).forEach((p, i) => { p.n = i + 1; });
    pg.el.addEventListener("change", async (e) => { const f = e.target.dataset.f; if (!f) return;
      if (f === "resp" && e.target.value === "__add") { const n = UP(prompt("Επώνυμο ειδικευόμενου:") || "").trim(); if (n) { cfg.residents = [...new Set([...(cfg.residents || []), n])].sort((a, b) => a.localeCompare(b, "el")); try { await saveCfg(cfg); } catch (er) {} P.resp = n; } touch(); return; } if (f === "date") { if (dirty && !confirm("Αλλαγή ημερομηνίας χωρίς αποθήκευση;")) { e.target.value = P.date; return; } load(e.target.value); return; } P[f] = e.target.value; touch(); });
    pg.el.addEventListener("click", async (e) => {
      const b = e.target.closest("button,[data-open]"); if (!b) return; const ds = b.dataset;
      if (ds.room) { const i = P.rooms.indexOf(ds.room); if (i >= 0) { if (P.pts.some((p) => p.room === ds.room)) { say("Η αίθουσα έχει ασθενείς"); return; } P.rooms.splice(i, 1); } else P.rooms.push(ds.room); touch(); return; }
      if (ds.resp) { P.resp = P.resp === ds.resp ? "" : ds.resp; touch(); return; }
      if (ds.respAdd !== undefined) { const n = UP(prompt("Επώνυμο ειδικευόμενου (όπως θα γράφεται):") || "").trim(); if (!n) return; cfg.residents = [...new Set([...(cfg.residents || []), n])].sort((a, b) => a.localeCompare(b, "el")); try { await saveCfg(cfg); } catch (er) {} P.resp = n; touch(); return; }
      if (ds.up || ds.dn) { const p = P.pts.find((x) => x.id === (ds.up || ds.dn)); const L = P.pts.filter((x) => x.room === p.room).sort((a, b) => a.n - b.n); const i = L.indexOf(p), j = ds.up ? i - 1 : i + 1; if (j < 0 || j >= L.length) return; const t = L[j].n; L[j].n = p.n; p.n = t; touch(); return; }
      if (ds.add) { pick(ds.add); return; }
      if (ds.open) { edit(P.pts.find((x) => x.id === ds.open)); return; }
      if (ds.act === "save") { try { await saveDraft(P); dirty = false; say("💾 Αποθηκεύτηκε (πρόχειρο)"); redraw(); } catch (er) { say("⚠ " + er.message); } return; }
      if (ds.act === "print") { printLocal(P); return; }
      if (ds.act === "submit") { submit(); return; }
      if (ds.act === "xr") { xrQueue(); return; }
    });
    // ---- επιλογή ασθενή ----
    function sheet(inner) { const ov = document.createElement("div"); ov.style.cssText = "position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:9999;display:flex;align-items:flex-end"; ov.innerHTML = '<div style="background:#F2F2F7;width:100%;max-height:88vh;overflow:auto;border-radius:16px 16px 0 0;padding:12px 12px calc(16px + env(safe-area-inset-bottom));font:15px -apple-system,system-ui;color:#1F2933" id="orpSh">' + inner + "</div>"; ov.addEventListener("click", (e) => { if (e.target === ov) ov.remove(); }); document.body.appendChild(ov); return ov; }
    function pick(room) {
      const used = new Set(P.pts.map((p) => p.key).filter(Boolean));
      const L = inpatients().filter((x) => !used.has(x.k));
      const ov = sheet('<div style="display:flex;align-items:center;margin-bottom:8px"><b style="flex:1;font-size:17px">' + E(ORP.roomLabel(room)) + '</b><button data-x style="border:0;background:none;font-size:22px">✕</button></div><input data-q placeholder="Αναζήτηση ονόματος/κρεβατιού" style="width:100%;box-sizing:border-box;font-size:16px;padding:9px;border:1px solid #C7CDD4;border-radius:10px">'
        + '<div data-list></div><button data-man style="margin-top:10px;width:100%;min-height:44px;border:0;border-radius:12px;background:#fff;color:#0A84FF;font-weight:700">✍ Χειροκίνητη προσθήκη</button>');
      const list = ov.querySelector("[data-list]");
      const show = (q) => { const z = NZ(q); list.innerHTML = L.filter((x) => !z || NZ(x.d.name + " " + x.k).indexOf(z) >= 0).map((x) => '<div data-k="' + E(x.k) + '" style="background:#fff;margin-top:6px;padding:10px 12px;border-radius:10px;display:flex;gap:8px;align-items:center"><span style="flex:1"><b>' + E(x.d.name) + '</b><br><span style="color:#6B7280;font-size:13px">' + E([x.k.replace(/^extra-\d+$/, "εκτός θαλάμου"), x.d.therapist, x.d.surgery || x.d.surgeryCustom, x.d.pod === "Προς Χ/Ο" ? "προς Χ/Ο" : ""].filter(Boolean).join(" · ")) + "</span></span><span style=\"color:#0A84FF;font-weight:700\">＋</span></div>").join("") || '<div style="padding:12px;color:#6B7280">Κανένας</div>'; };
      show("");
      ov.querySelector("[data-q]").addEventListener("input", (e) => show(e.target.value));
      ov.addEventListener("click", (e) => {
        if (e.target.closest("[data-x]")) { ov.remove(); return; }
        const r = e.target.closest("[data-k]");
        if (r) { const x = L.find((y) => y.k === r.dataset.k); const p = fromPatient(x.k, x.d); p.room = room; p.n = P.pts.filter((q) => q.room === room).length + 1; P.pts.push(p); ov.remove(); touch(); if (!p.t) edit(p); return; }
        if (e.target.closest("[data-man]")) { const p = { id: "m" + Date.now().toString(36), key: "", room, n: P.pts.filter((q) => q.room === room).length + 1, bed: "", adm: "", nm: "", dx: "", an: "ΓΕΝΙΚΗ", bl: "", ca: "", t: "", side: "", t2: "", side2: "", surg: "" }; P.pts.push(p); ov.remove(); touch(); edit(p); }
      });
    }
    // ---- επεξεργασία ασθενή ----
    function edit(p) {
      if (!p) return;
      const box = (lbl, inner) => '<div style="margin-top:10px"><div style="font-weight:700;margin-bottom:4px">' + lbl + "</div>" + inner + "</div>";
      const ch = (a, v, on, t) => '<button data-' + a + '="' + E(v) + '" style="min-height:36px;margin:3px;padding:6px 12px;border-radius:18px;border:1px solid ' + (on ? "#0A84FF" : "#C7CDD4") + ";background:" + (on ? "#0A84FF" : "#fff") + ";color:" + (on ? "#fff" : "#1F2933") + ';font-weight:600">' + E(t == null ? (v || "—") : t) + "</button>";
      const inp = (f, v, ph) => '<input data-e="' + f + '" value="' + E(v) + '" placeholder="' + E(ph || "") + '" style="width:100%;box-sizing:border-box;font-size:16px;padding:9px;border:1px solid #C7CDD4;border-radius:10px">';
      let ov;
      const render = () => {
        const inner = '<div style="display:flex;align-items:center"><b style="flex:1;font-size:17px">' + E(p.nm || "Νέος ασθενής") + '</b><button data-x style="border:0;background:none;font-size:22px">✕</button></div>'
          + box("Αίθουσα", P.rooms.map((r) => ch("rm", r, p.room === r, ORP.roomLabel(r).replace("ΑΙΘΟΥΣΑ ", ""))).join(""))
          + '<div style="display:flex;gap:8px">' + '<div style="flex:1">' + box("Θάλαμος.κρεβάτι", inp("bed", p.bed, "341.1")) + '</div><div style="flex:1">' + box("Ημ. εισόδου", inp("adm", p.adm, "08/11")) + "</div></div>"
          + box("Ονοματεπώνυμο · του πατρός · ΑΜ", inp("nm", p.nm, "ΕΠΩΝΥΜΟ ΟΝΟΜΑ ΤΟΥ ΠΑΤΡΟΣ 12345678"))
          + box("Επέμβαση (έτοιμο σετ)", '<button data-pk="1" style="width:100%;text-align:left;min-height:44px;border:1px solid #C7CDD4;border-radius:10px;background:#fff;padding:8px 10px;font-size:15px">' + (p.t ? "<b>" + E(p.dx) + "</b><br>" + E(p.t) : "🔎 Διάλεξε…") + "</button>"
            + '<div style="margin-top:4px">' + SIDES.map((s) => ch("sd", s, (p.side || "") === s, s || "χωρίς πλευρά")).join("") + "</div>")
          + box("Δεύτερη επέμβαση", (p.t2 ? '<div style="display:flex;gap:6px;align-items:center"><span style="flex:1">' + E(p.t2) + '</span><button data-t2x style="border:0;background:#FDECEC;color:#B42318;border-radius:8px;min-height:34px;padding:0 10px">✕</button></div><div>' + SIDES.map((s) => ch("sd2", s, (p.side2 || "") === s, s || "χωρίς πλευρά")).join("") + "</div>" : '<button data-pk="2" style="min-height:40px;border:1px dashed #C7CDD4;border-radius:10px;background:none;padding:6px 12px;color:#0A84FF">＋ δεύτερη επέμβαση</button>'))
          + box("Διάγνωση", inp("dx", p.dx, "π.χ. ΑΠΟΦΡΑΚΤΙΚΗ ΟΥΡΟΠΑΘΕΙΑ"))
          + box("Αναισθησία", ANS.map((a) => ch("an", a, p.an === a)).join(""))
          + box("Αίμα", BLOOD.map((a) => ch("bl", a, (p.bl || "") === a, a || "—")).join("") + '<button data-ca style="margin-left:10px;min-height:36px;padding:6px 12px;border-radius:18px;border:1px solid ' + (p.ca ? "#0A84FF" : "#C7CDD4") + ";background:" + (p.ca ? "#0A84FF" : "#fff") + ";color:" + (p.ca ? "#fff" : "#1F2933") + ';font-weight:600">c-ARM ' + (p.ca ? "ΝΑΙ" : "—") + "</button>")
          + box("Χειρουργός", SURGEONS().map((s) => ch("sg", s, p.surg === s, s.replace(/^κ /, ""))).join("") + inp("surg", p.surg, "κ ΕΠΩΝΥΜΟ"))
          + '<div style="display:flex;gap:8px;margin-top:14px"><button data-del style="flex:1;min-height:46px;border:0;border-radius:12px;background:#FDECEC;color:#B42318;font-weight:700">Αφαίρεση</button><button data-x style="flex:2;min-height:46px;border:0;border-radius:12px;background:#0A84FF;color:#fff;font-weight:700">Τέλος</button></div>';
        if (!ov) ov = sheet(inner); else ov.querySelector("#orpSh").innerHTML = inner;
      };
      render();
      ov.addEventListener("input", (e) => { const f = e.target.dataset.e; if (!f) return; p[f] = f === "surg" || f === "dx" || f === "nm" ? e.target.value.toUpperCase().replace(/^Κ /, "κ ") : e.target.value; dirty = true; });
      ov.addEventListener("click", (e) => {
        const b = e.target.closest("button"); if (!b) return; const d = b.dataset;
        if (d.x !== undefined) { ov.remove(); touch(); return; }
        if (d.del !== undefined) { if (!confirm("Αφαίρεση από το πρόγραμμα;")) return; P.pts = P.pts.filter((x) => x !== p); renum(p.room); ov.remove(); touch(); return; }
        if (d.rm) { const old = p.room; p.room = d.rm; p.n = P.pts.filter((q) => q.room === d.rm && q !== p).length + 1; renum(old); }
        else if (d.sd !== undefined) p.side = d.sd;
        else if (d.sd2 !== undefined) p.side2 = d.sd2;
        else if (d.an) p.an = d.an;
        else if (d.bl !== undefined) p.bl = d.bl;
        else if (d.ca !== undefined) p.ca = p.ca ? "" : "ΝΑΙ";
        else if (d.sg) p.surg = d.sg;
        else if (d.t2x !== undefined) { p.t2 = ""; p.side2 = ""; }
        else if (d.pk) { pkPick((x) => { if (d.pk === "2") { p.t2 = x.t; } else { p.t = x.t; if (x.dx) p.dx = x.dx; if (x.an) p.an = x.an; p.bl = x.bl || ""; p.ca = x.ca || ""; } dirty = true; render(); }); return; }
        else return;
        dirty = true; render();
      });
    }
    // ---- διαλογή έτοιμου σετ: σετ κλινικής → «ΚΩΔΙΚΟΙ (ΝΕΟΙ)» → ΕΤΙΠ ----
    let etip = null;
    function pkPick(done) {
      const ov = sheet('<div style="display:flex;align-items:center;margin-bottom:8px"><b style="flex:1;font-size:17px">Επέμβαση</b><button data-x style="border:0;background:none;font-size:22px">✕</button></div><input data-q placeholder="Αναζήτηση (π.χ. νεφρεκτομή, pigtail, 357.03)" style="width:100%;box-sizing:border-box;font-size:16px;padding:9px;border:1px solid #C7CDD4;border-radius:10px"><div data-list></div>');
      const list = ov.querySelector("[data-list]"); let items = [];
      const row = (i, a, b, c) => '<div data-i="' + i + '" style="background:#fff;margin-top:6px;padding:9px 12px;border-radius:10px"><b>' + E(a) + '</b><br><span style="font-size:14px">' + E(b) + "</span>" + (c ? '<br><span style="color:#6B7280;font-size:12px">' + E(c) + "</span>" : "") + "</div>";
      const show = async (q) => {
        const z = NZ(q); items = []; let h = "";
        const L = PK().filter((x) => !z || NZ(x.dx + " " + x.t + " " + x.g).indexOf(z) >= 0);
        let g = ""; for (const x of L) { if (x.g !== g && !z) { g = x.g; h += '<div style="margin:12px 4px 2px;font-weight:700;color:#6B7280;font-size:13px">' + E(g) + "</div>"; } items.push(x); h += row(items.length - 1, x.dx, x.t, [x.an, x.bl && "αίμα " + x.bl, x.ca && "c-ARM"].filter(Boolean).join(" · ")); }
        if (z.length >= 3) {
          const C = ((window.ORP_DATA || {}).codes || []).filter((c) => NZ(c.d + " " + c.c + " " + c.s).indexOf(z) >= 0).slice(0, 30);
          if (C.length) h += '<div style="margin:12px 4px 2px;font-weight:700;color:#6B7280;font-size:13px">ΚΩΔΙΚΟΙ (ΝΕΟΙ) ΓΙΑ ΛΙΣΤΑ ΧΕΙΡΟΥΡΓΕΙΟΥ</div>';
          for (const c of C) { items.push({ dx: "", t: UP(c.d) + " " + c.c }); h += row(items.length - 1, UP(c.d), c.c, c.s); }
          if (!etip) { try { etip = await (await fetch("etip.json", { cache: "force-cache" })).json(); } catch (e) { etip = []; } }
          const T = etip.filter((x) => NZ(x[0] + " " + x[1]).indexOf(z) >= 0).slice(0, 40);
          if (T.length) h += '<div style="margin:12px 4px 2px;font-weight:700;color:#6B7280;font-size:13px">ΕΤΙΠ — συστηματικός κατάλογος</div>';
          for (const x of T) { const name = UP(String(x[1]).split(" — ").pop()); items.push({ dx: "", t: name + " " + x[0] }); h += row(items.length - 1, x[0], x[1]); }
          h += row(items.push({ dx: "", t: UP(q) }) - 1, "✍ Όπως το έγραψα", UP(q));
        }
        list.innerHTML = h || '<div style="padding:12px;color:#6B7280">Τίποτα — γράψτε περισσότερα</div>';
      };
      show("");
      ov.querySelector("[data-q]").addEventListener("input", (e) => show(e.target.value));
      ov.addEventListener("click", (e) => { if (e.target.closest("[data-x]")) { ov.remove(); return; } const r = e.target.closest("[data-i]"); if (r) { done(items[+r.dataset.i]); ov.remove(); } });
    }
    // ---- 🫁 ακτινογραφίες θώρακα στη σειρά του προγράμματος (για την καρδιολογική εκτίμηση) ----
    function xrQueue() {
      const items = [];
      for (const R of ORP.rooms(P)) for (const p of R.pts) {
        let d = null; try { d = p.key ? ((typeof data !== "undefined" && data[p.key]) || (typeof extraPatients !== "undefined" && extraPatients[p.key]) || null) : null; } catch (e) {}
        const am = (/\b(\d{8})\b/.exec(p.nm || "") || [])[1] || (d && d.am) || "";
        const name = d && d.name ? d.name : String(p.nm || "").replace(/\s+ΤΟΥ\s.*$/, "").replace(/\s*\d{8}.*$/, "").trim();
        items.push({ key: p.key || "", name, am, amka: (d && d.amka) || "", ad: (d && d.admitDate) || p.adm || "", label: R.label.replace("ΑΙΘΟΥΣΑ ", "Αίθ. ") + (p.bed ? " · " + p.bed : "") });
      }
      window.pdxXrQueue(items);
    }
    // ---- υποβολή: Excel στο δίκτυο + εκτύπωση (bot Μεσαίου) ----
    async function submit() {
      const bad = P.pts.filter((p) => !p.nm || !ORP.procText(p));
      if (!P.pts.length) { say("Το πρόγραμμα είναι άδειο"); return; }
      if (bad.length) { say("⚠ Λείπει όνομα ή επέμβαση σε " + bad.length + " ασθενή(ς)"); return; }
      if (!P.resp && !confirm("Χωρίς υπεύθυνο προγράμματος — συνέχεια;")) return;
      const fname = ORP.fileName(P.date, P.kind), dir = ORP.monthDir(P.date);
      if (!confirm("Υποβολή «" + fname + "» στον φάκελο " + dir + " και εκτύπωση στον " + PRINTER + ";")) return;
      try {
        say("✅ Υποβολή…");
        const x = await sealB64(ORP.xlsx(P)), hh = await sealB64(new TextEncoder().encode(ORP.html(P)));
        const ins = await SB().from("oncall_jobs").insert({ name_key: "orsubmit", name: "🗓 Πρόγραμμα " + ORP.dShort(P.date), room: null, kind: "orsubmit", status: "new",
          payload: { x, h: hh, dir, name: P.file || fname, overwrite: !!P.file, printer: PRINTER, only: BOT, quiet: true } }).select("id").single();
        if (ins.error) throw new Error(ins.error.message);
        for (let i = 0; i < 80; i++) {
          await new Promise((r) => setTimeout(r, 1500));
          const r = await SB().from("oncall_jobs").select("status,note,payload").eq("id", ins.data.id).maybeSingle(); const j = r.data || {};
          if (j.status === "done") { const res = (j.payload || {}).result || {}; P.status = "submitted"; P.file = res.name || fname; P.subAt = Date.now(); await saveDraft(P); dirty = false; redraw();
            SB().from("oncall_jobs").update({ payload: { dir, name: P.file, result: res } }).eq("id", ins.data.id).then(() => {}, () => {}); alert(j.note || "✅ Υποβλήθηκε"); return; }
          if (j.status === "failed") { alert("⚠ " + (j.note || "απέτυχε")); return; }
        }
        alert("Το bot του Μεσαίου δεν απάντησε ακόμα — θα το κάνει μόλις ξυπνήσει. Δείτε «Ενέργειες bot».");
      } catch (e) { alert("⚠ " + (e.message || e)); }
    }
    load(P.date);
    return pg;
  };
})();
