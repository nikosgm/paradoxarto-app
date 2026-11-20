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
    { let r = 2; const groups = []; const nz0 = (v) => String(v || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[^A-ZΑ-Ω0-9]/g, ""), seen0 = [];   // 17/11: χωρίς διπλά (όπως η λίστα επιλογής)
      for (const x of [...(d.lib || []), ...(d.extra || [])]) { const t0 = nz0(x.t), d0 = nz0(x.dx); if (seen0.some((y) => y[0] === t0 && (y[1] === d0 || y[1].startsWith(d0) || d0.startsWith(y[1])))) continue; seen0.push([t0, d0]); let g = groups.find((y) => y.g === x.g); if (!g) groups.push(g = { g: x.g, L: [] }); g.L.push(x); }
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
        // ✏ 16/11 (χρήστης: «να πατάω στο κελί και να το αλλάζω εύκολα»): στην προεπισκόπηση (opt.click) κάθε κελί ανοίγει την επεξεργασία στο πεδίο του
        const A = (f) => opt.click ? ' data-cell="' + HE(p.id) + "|" + f + '" style="cursor:pointer"' : "";
        h += '<tr class="p1"><td rowspan="2" class="b"' + A("bed") + ">" + HE(p.bed) + '</td><td rowspan="2" class="b"' + A("adm") + ">" + HE(p.adm) + '</td><td rowspan="2"' + A("nm") + ">" + hb(p.nm) + "</td><td" + A("dx") + ">" + HE(p.dx) + "</td><td" + A("an") + ">" + HE(p.an) + '</td><td class="b"' + A("bl") + ">" + HE(p.bl) + '</td><td class="b"' + A("ca") + ">" + HE(p.ca) + '</td><td rowspan="2" class="b"' + A("surg") + ">" + HE(p.surg) + "</td></tr>";
        h += '<tr class="p2"><td colspan="4"' + A("proc") + ">" + hb(procText(p)) + "</td></tr>";
      }
    }
    h += '<tr class="end"><td colspan="8"></td></tr><tr><td colspan="7" style="border:0"></td><td class="b c" style="border:0;font-size:14pt">Ο ΔΙΕΥΘΥΝΤΗΣ</td></tr></table>';
    // 🖨 16/11 (χρήστης: «η εκτύπωση κόβει τα γκρι — θέλω ακριβώς όπως το χαρτί»): οι browsers ΔΕΝ τυπώνουν φόντα από προεπιλογή → ρητό print-color-adjust
    const css = ".orp,.orp *{-webkit-print-color-adjust:exact;print-color-adjust:exact;color-adjust:exact}.orp{border-collapse:collapse;width:100%;table-layout:fixed;font:11pt Arial,Helvetica,sans-serif;color:#000}.orp td{border:1px solid #000;padding:2px 4px;text-align:center;vertical-align:middle;word-wrap:break-word}"
      + ".orp .b{font-weight:700}.orp .l{text-align:left;border:0}.orp tr:nth-child(1) td,.orp tr:nth-child(2) td,.orp tr:nth-child(3) td{border:0;font-weight:700}.orp tr:nth-child(2) td.l{font-weight:400}"
      + ".orp .hd td{background:#C0C0C0;font-weight:700}.orp .rm td{background:#BFBFBF;font-weight:700;text-align:center}.orp .p2 td{font-size:11pt}.orp .end td{border:0;height:8px}";
    if (opt.bodyOnly) return "<style>" + css + "</style>" + h;
    return '<!doctype html><html><head><meta charset="utf-8"><title>' + HE((prog.kind || "ΠΡΟΓΡΑΜΜΑ") + " " + dShort(prog.date)) + "</title><style>@page{size:A4 landscape;margin:4mm 3mm}html,body{margin:0;-webkit-print-color-adjust:exact;print-color-adjust:exact}" + css + "</style></head><body>" + h + "</body></html>";
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
  // 🧹 17/11 19:37 (χρήστης: «όταν επιλέγω χειρουργείο, έχεις πάρει στη λίστα και κάποια duplicate — σβήσ' τα»): ίδια επέμβαση (χωρίς κενά/παύλες/τόνους)
  // με ίδια διάγνωση (ή η μία συνέχεια της άλλης, π.χ. «CA ΚΥΣΤΗΣ» ↔ «CA ΚΥΣΤΗΣ-ΓΥΝΑΙΚΑ») → ΜΙΑ φορά (κρατά το πρώτο = έτοιμα σετ πριν από τα «επιπλέον»)
  let PKc = null, PKsrc = null;
  const PK = () => { const d = window.ORP_DATA || {}; if (PKc && PKsrc === d) return PKc; const out = [];
    for (const x of [...(d.lib || []), ...(d.extra || [])]) { const t = NZ(x.t), dx = NZ(x.dx);
      if (out.some((y) => NZ(y.t) === t && (NZ(y.dx) === dx || NZ(y.dx).startsWith(dx) || dx.startsWith(NZ(y.dx))))) continue; out.push(x); }
    PKsrc = d; PKc = out; return out; };
  window.__orpPK = () => PK(); window.__orpFind = (q, n) => pkFind(q, n);   // για τις δοκιμές
  // 🔎 16/11 (χρήστης: «αναζήτηση χειρουργείων με drop-down»): ταίριασμα ΑΝΑ ΛΕΞΗ (οποιαδήποτε σειρά, χωρίς τόνους) + συντομογραφίες της κλινικής
  const PKSYN = { TURB: ["ΔΙΟΥΡΗΘΡ", "ΚΥΣΤ"], TURBT: ["ΔΙΟΥΡΗΘΡ", "ΚΥΣΤ"], TUR: ["ΔΙΟΥΡΗΘΡ"], TURP: ["ΔΙΟΥΡΗΘΡ", "ΠΡΟΣΤΑΤ"], HOLEP: ["ΠΡΟΣΤΑΤ"], RIRS: ["ΟΥΡΗΤΗΡΟΣΚΟΠ"], URS: ["ΟΥΡΗΤΗΡΟΣΚΟΠ"], ΟΛΥ: ["ΟΥΡΗΤΗΡΟΛΙΘΟΤΡΙΨ"], ΟΝΥ: ["ΟΥΡΗΤΗΡΟΛΙΘΟΤΡΙΨ"],
    PCNL: ["ΔΙΑΔΕΡΜ", "ΝΕΦΡΟΛΙΘ"], PCN: ["ΝΕΦΡΟΣΤΟΜ"], RC: ["ΚΥΣΤΕΚΤΟΜ"], RP: ["ΠΡΟΣΤΑΤΕΚΤΟΜ"], RARP: ["ΠΡΟΣΤΑΤΕΚΤΟΜ"], NX: ["ΝΕΦΡΕΚΤΟΜ"], PN: ["ΜΕΡΙΚΗ", "ΝΕΦΡΕΚΤΟΜ"], NUX: ["ΝΕΦΡΟΟΥΡΗΤΗΡΕΚΤΟΜ"], NU: ["ΝΕΦΡΟΟΥΡΗΤΗΡΕΚΤΟΜ"],
    AUS: ["ΣΦΙΓΚΤΗΡ"], DJ: ["PIGTAIL"], JJ: ["PIGTAIL"], PIGTAIL: ["PIGTAIL"], STENT: ["PIGTAIL"], MALLEABLE: ["ΗΜΙΑΚΑΜΠΤ"], IPP: ["ΠΕΟΣ", "ΠΡΟΘΕΣ"], OPTILUME: ["ΑΥΧΕΝΑ"], CYSTO: ["ΚΥΣΤΕΟΣΚΟΠ"], ΚΥΣΤΕΟ: ["ΚΥΣΤΕΟΣΚΟΠ"], ΠΕΡΙΤΟΜΗ: ["ΠΕΡΙΤΟΜ"] };
  const PKSIDE = { ΑΡ: "ΑΡ", ΑΡΙΣΤΕΡΑ: "ΑΡ", ΑΡΙΣΤΕΡΗ: "ΑΡ", ΔΕ: "ΔΕ", ΔΕΞΙΑ: "ΔΕ", ΔΕΞΙΑΣ: "ΔΕ", ΑΜΦ: "ΑΜΦΩ", ΑΜΦΩ: "ΑΜΦΩ", ΑΜΦΟΤΕΡΟΠΛΕΥΡΗ: "ΑΜΦΩ" };
  const pkSide = (q) => { for (const t of UP(q).split(/[^A-ZΑ-Ω]+/)) if (PKSIDE[t]) return PKSIDE[t]; return ""; };
  const pkTok = (q) => UP(q).split(/[^A-ZΑ-Ω0-9.]+/).filter((t) => t && !PKSIDE[t]).map((t) => PKSYN[t.replace(/\./g, "")] || [t.replace(/\./g, "")]);
  // κάθε λέξη (ή η επέκτασή της) πρέπει να υπάρχει· βαθμός: πόσες λέξεις ταιριάζουν στην ΑΡΧΗ λέξης
  const pkHit = (toks, hay) => { const H = NZ(hay), W = UP(hay).split(/[^A-ZΑ-Ω0-9]+/); let sc = 0;
    for (const alt of toks) { if (!alt.every((t) => H.indexOf(NZ(t)) >= 0)) return -1; if (alt.some((t) => W.some((w) => w.startsWith(NZ(t))))) sc++; } return sc; };
  const pkFind = (q, max) => { const toks = pkTok(q); if (!toks.length) return []; const out = [];
    for (const x of PK()) { const sc = pkHit(toks, x.dx + " " + x.t + " " + x.g); if (sc >= 0) out.push({ x, sc: sc + 2 }); }
    for (const c of ((window.ORP_DATA || {}).codes || [])) { const sc = pkHit(toks, c.d + " " + c.c + " " + c.s); if (sc >= 0) out.push({ x: { dx: "", t: UP(c.d) + " " + c.c, code: c.c, sub: c.s }, sc }); }
    // κωδικός που γράφεται ΙΔΙΟΣ με έτοιμο σετ ή με άλλον κωδικό → μία φορά
    const seen = new Set(); return out.sort((a, b) => b.sc - a.sc).filter((o) => { const k = NZ(o.x.t) + "|" + NZ(o.x.dx); if (seen.has(k) || (!o.x.dx && [...seen].some((s0) => s0.split("|")[0] === NZ(o.x.t)))) return false; seen.add(k); return true; }).slice(0, max || 40).map((o) => o.x); };
  // 🗓 20/11 (χρήστης: «το πρόχειρο το καταλαβαίνει με πολλά λάθη — ό,τι μοιάζει με κάτι να το βάζει · ό,τι γράφουμε υπάρχει στη βάση μας ·
  // φτιάξε αρχείο αντιστοίχισης να ανατρέχει»): (α) αρχείο or-proc-alias.json (tools/oralias.js) + μαθημένα oncall_labs «orprog|alias»
  // (κάθε διόρθωση/✓ σε γραμμή από σάρωση) · (β) ανεκτικό ταίριασμα: η ΠΙΟ ΚΟΝΤΙΝΗ επέμβαση του καταλόγου (γράμματα/συντομογραφίες).
  const ALIAS = { m: {}, at: 0 };
  const aliasLoad = async () => { if (Date.now() - ALIAS.at < 60000) return ALIAS.m; ALIAS.at = Date.now(); const m = {};
    try { const r = await fetch("or-proc-alias.json?v=" + Math.floor(Date.now() / 600000)); if (r.ok) { const j = await r.json(); for (const [k, v] of Object.entries(Object.assign({}, j.alias || {}, j.learned || {}))) m[NZ(k)] = v; } } catch (e) {}
    try { const r = await SB().from("oncall_labs").select("data").eq("name_key", "orprog|alias").maybeSingle(); const a = (r.data && r.data.data && r.data.data.a) || {}; for (const [k, v] of Object.entries(a)) if (v && v.t) m[NZ(k)] = v.t; } catch (e) {}
    ALIAS.m = m; return m; };
  const aliasLearn = async (raw, t) => { const k = NZ(raw); if (!k || k.length < 2 || !t || ALIAS.m[k] === t) return; ALIAS.m[k] = t;
    try { const r = await SB().from("oncall_labs").select("data").eq("name_key", "orprog|alias").maybeSingle(); const d = (r.data && r.data.data) || {}; d.a = d.a || {};
      d.a[k] = { t, raw: String(raw).slice(0, 80), n: ((d.a[k] || {}).n || 0) + 1, at: Date.now() };
      await SB().from("oncall_labs").upsert({ name_key: "orprog|alias", name: "🗓 Αντιστοιχίσεις επεμβάσεων (χαρτί → κατάλογος)", pid: null, data: d, updated_at: new Date().toISOString() }, { onConflict: "name_key" }); } catch (e) {} };
  window.__orpAlias = ALIAS; window.__orpAliasLoad = aliasLoad;
  const pkLev = (a, b) => { if (Math.abs(a.length - b.length) > 6) return 99; let p = Array.from({ length: b.length + 1 }, (_, j) => j); for (let i = 1; i <= a.length; i++) { const q = [i]; for (let j = 1; j <= b.length; j++) q[j] = Math.min(p[j] + 1, q[j - 1] + 1, p[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)); p = q; } return p[b.length]; };
  const pkSim = (a, b) => { if (!a || !b) return 0; const sh = Math.min(a.length, b.length), lo = Math.max(a.length, b.length);
    if (sh >= 3 && (b.indexOf(a) === 0 || a.indexOf(b) === 0) && (a.length <= b.length ? true : sh / lo >= 0.6)) return 0.92;   /* συντομογραφία του χαρτιού = αρχή λέξης του σετ (όχι το ανάποδο με μικρή λέξη, π.χ. «ΟΥΡ.») */
    const d = pkLev(a, b); return Math.max(0, 1 - d / Math.max(a.length, b.length)); };
  // η ΠΙΟ ΚΟΝΤΙΝΗ επέμβαση: κάθε λέξη του χαρτιού (ή η επέκτασή της, π.χ. TURB → ΔΙΟΥΡΗΘΡ+ΚΥΣΤ) με την πιο κοντινή λέξη του σετ · μέσος όρος → {x, sc 0–1}
  const pkBest = (q) => { const toks = UP(q).split(/[^A-ZΑ-Ω0-9]+/).filter((t) => t.length >= 2 && !PKSIDE[t] && !/^\d+$/.test(t)); if (!toks.length) return null;
    const cand = [...PK().map((x) => ({ x, h: x.t + " " + x.dx })), ...(((window.ORP_DATA || {}).codes) || []).map((c) => ({ x: { dx: "", t: UP(c.d) + " " + c.c, code: c.c, sub: c.s }, h: c.d + " " + c.s, cd: 1 }))];
    let best = null;
    for (const c of cand) { const W = UP(c.h).split(/[^A-ZΑ-Ω0-9]+/).filter((w) => w.length >= 2 && !/^\d+$/.test(w)).map(NZ); if (!W.length) continue;
      let tot = 0; for (const t of toks) { const alts = PKSYN[t.replace(/\./g, "")] || [t]; tot += alts.reduce((s0, a) => s0 + Math.max(...W.map((w) => pkSim(NZ(a), w))), 0) / alts.length; }
      const sc = tot / toks.length - (c.cd ? 0.03 : 0); if (!best || sc > best.sc) best = { x: c.x, sc }; }
    return best; };
  window.__orpBest = (q) => pkBest(q);
  // γραμμή χαρτιού → σετ: αρχείο αντιστοίχισης → επιλογή του καταλόγου από τη σάρωση («χαρτί => κατάλογος») → αναζήτηση → πιο κοντινό
  const procResolve = (pq0) => { const s0 = String(pq0 || "").trim(); if (!s0) return null; const ix = s0.indexOf("=>"); const raw = (ix >= 0 ? s0.slice(0, ix) : s0).trim(), can = ix >= 0 ? s0.slice(ix + 2).trim() : "";
    const bySet = (t) => PK().find((y) => NZ(y.t) === NZ(t)) || null;
    const noSide = UP(raw).split(/[^A-ZΑ-Ω0-9]+/).filter((t) => t && !PKSIDE[t]).join("");
    const al = ALIAS.m[NZ(raw)] || ALIAS.m[noSide];
    if (al) { const k = bySet(al) || pkFind(al, 1)[0]; if (k) return { k, raw, how: "alias", sure: true }; }
    if (can) { const k = bySet(can) || pkFind(can, 1)[0]; if (k) { const k2 = raw ? pkFind(raw, 1)[0] : null; return { k, raw, how: "cat", sure: !!(k2 && NZ(k2.t) === NZ(k.t)) }; } }
    if (raw) { const k = pkFind(raw, 1)[0]; if (k) return { k, raw, how: "find", sure: true }; }
    const B = [can && pkBest(can), raw && pkBest(raw)].filter(Boolean).sort((a, b) => b.sc - a.sc)[0];
    if (B && B.sc >= 0.5) return { k: B.x, raw, how: "fuzzy", sure: false, sc: B.sc };
    return { k: null, raw }; };
  window.__orpResolve = (q) => procResolve(q);
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
    // 🏥 9/11 (χρήστης: «από όλο το ΝΠΣ, όχι μόνο Παραδόχαρτο»): + όσοι νοσηλεύονται στον ΝΠΣ αλλά δεν είναι στο Παραδόχαρτο (τελευταία λίστα ΝΠΣ 3449)
    try { const N = (window.pdxNpsBulkList && window.pdxNpsBulkList()) || []; for (const x of N) if (x && x.d && String(x.d.name || "").trim()) L.push({ k: x.k, d: x.d, nps: true }); } catch (e) {}
    // 🔤 9/11 (χρήστης): η λίστα επιλογής ΑΛΦΑΒΗΤΙΚΑ (επώνυμο) — Παραδόχαρτο και ΝΠΣ μαζί
    const C = new Intl.Collator("el", { sensitivity: "base", numeric: true });
    return L.sort((a, b) => C.compare(String(a.d.name || "").trim(), String(b.d.name || "").trim()) || C.compare(String(a.k), String(b.k)));
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
  // 🚨 16/11: χωριστό πρόχειρο ανά είδος — τακτικό «orprog|<ημ>» (όπως πριν), έκτακτο «orprog|<ημ>|ΕΚΤΑΚΤΟ», συμπληρωματικό «…|ΣΥΜΠΛΗΡΩΜΑΤΙΚΟ»
  const KINDS = [["ΠΡΟΓΡΑΜΜΑ", "📋 Τακτικό"], ["ΕΚΤΑΚΤΟ ΠΡΟΓΡΑΜΜΑ", "🚨 Έκτακτο"], ["ΣΥΜΠΛΗΡΩΜΑΤΙΚΟ ΠΡΟΓΡΑΜΜΑ", "➕ Συμπληρωματικό"]];
  const kTag = (kind) => (!kind || kind === "ΠΡΟΓΡΑΜΜΑ") ? "" : "|" + String(kind).replace(/ ΠΡΟΓΡΑΜΜΑ$/, "");
  const keyOf = (date, kind) => "orprog|" + date + kTag(kind);
  const todayIso = () => { const d = new Date(); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); };
  async function loadCfg() { try { const r = await SB().from("oncall_labs").select("data").eq("name_key", "orprog|cfg").maybeSingle(); return (r.data && r.data.data) || {}; } catch (e) { return {}; } }
  async function saveCfg(c) { await SB().from("oncall_labs").upsert({ name_key: "orprog|cfg", name: "Πρόγραμμα χειρουργείου", pid: null, data: c, updated_at: new Date().toISOString() }, { onConflict: "name_key" }); }
  async function loadDraft(date, kind) { const r = await SB().from("oncall_labs").select("data,updated_at").eq("name_key", keyOf(date, kind)).maybeSingle(); if (r.error) throw new Error(r.error.message); const env = r.data && r.data.data; if (!env) return null; return await CK().decrypt(env.enc || env); }
  async function saveDraft(P) { P.at = Date.now(); P.by = String(window.authName || "").replace(/@paradoxarto\.app$/, ""); const enc = await CK().encrypt(P); const r = await SB().from("oncall_labs").upsert({ name_key: keyOf(P.date, P.kind), name: (P.kind && P.kind !== "ΠΡΟΓΡΑΜΜΑ" ? P.kind.replace(/ ΠΡΟΓΡΑΜΜΑ$/, "") + " " : "Πρόγραμμα χειρουργείου ") + ORP.dShort(P.date), pid: null, data: { enc, st: P.status || "draft", n: (P.pts || []).length }, updated_at: new Date().toISOString() }, { onConflict: "name_key" }); if (r.error) throw new Error(r.error.message); }
  async function sealB64(u8) { const s = await CK().sealBytes(u8); let b = ""; for (let i = 0; i < s.length; i += 0x8000) b += String.fromCharCode.apply(null, s.subarray(i, i + 0x8000)); return btoa(b); }
  function printLocal(P) { const w = window.open("", "_blank"); if (!w) { say("Επιτρέψτε τα αναδυόμενα για εκτύπωση"); return; } w.document.write(ORP.html(P)); w.document.close(); setTimeout(() => { try { w.focus(); w.print(); } catch (e) {} }, 400); }

  window.pdxOrProg = async function (date0, kind0, opts0) {
    let P = { date: date0 || (kind0 === "ΕΚΤΑΚΤΟ ΠΡΟΓΡΑΜΜΑ" ? todayIso() : ORP.defDate()), kind: kind0 || "ΠΡΟΓΡΑΜΜΑ", resp: "", rooms: ["10", "8"], pts: [], status: "draft" }, cfg = {}, dirty = false, busy = "", loaded = false;
    const can = ["resident", "attending", "admin"].indexOf(window.authRole) >= 0;
    const css = '<style>#orpPg .ie{box-sizing:border-box;font:15px -apple-system,system-ui;padding:5px 7px;border:1px solid #C7CDD4;border-radius:8px;background:#fff;color:#1F2933}#orpPg .iel{display:inline-flex;align-items:center;gap:4px;color:#6B7280;font-size:13px}#orpPg [data-cell]:hover{background:#FFF7CC!important;outline:2px solid #0A84FF}#orpPg .chip{display:inline-flex;align-items:center;min-height:36px;padding:6px 12px;margin:3px;border-radius:18px;border:1px solid #C7CDD4;background:#fff;font:600 14px -apple-system,system-ui;color:#1F2933}#orpPg .chip.on{background:#0A84FF;border-color:#0A84FF;color:#fff}'
      + '#orpPg .card{margin:8px 12px;padding:10px 12px;border-radius:12px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.08);font:14px -apple-system,system-ui}#orpPg .card b{font-weight:700}#orpPg .rm{margin:16px 16px 4px;font:700 13px -apple-system,system-ui;color:#6B7280;letter-spacing:.4px}'
      + '#orpPg .mini{border:0;background:#EEF2F6;border-radius:8px;min-width:38px;min-height:34px;font-size:15px;margin-left:4px}#orpPg .bar{position:sticky;bottom:0;display:flex;gap:8px;padding:10px 12px calc(10px + env(safe-area-inset-bottom));background:rgba(242,242,247,.96);backdrop-filter:blur(8px);border-top:1px solid #DDE1E6}'
      + '#orpPg .bar button{flex:1;min-height:46px;border:0;border-radius:12px;font:700 14px -apple-system,system-ui}#orpPg input,#orpPg select{font-size:16px;padding:8px 10px;border:1px solid #C7CDD4;border-radius:10px;background:#fff}'
      + 'html.pdx-night #orpPg .card{background:#1C1C1E;color:#fff}html.pdx-night #orpPg .chip{background:#2C2C2E;color:#fff;border-color:#3A3A3C}html.pdx-night #orpPg .chip.on{background:#0A84FF}html.pdx-night #orpPg .bar{background:rgba(28,28,30,.96);border-color:#38383A}html.pdx-night #orpPg input,html.pdx-night #orpPg select{background:#2C2C2E;color:#fff;border-color:#3A3A3C}html.pdx-night #orpPg .mini{background:#2C2C2E;color:#fff}</style>';
    const chip = (attr, v, on, t) => '<button class="chip' + (on ? " on" : "") + '" ' + attr + '="' + E(v) + '">' + (t == null ? E(v || "—") : t) + "</button>";
    const draw = () => {
      if (!loaded) return css + '<div id="orpPg"><div class="card">Φόρτωση…</div></div>';
      let h = css + '<div id="orpPg">';
      // 🚨 16/11 (χρήστης: «αρχική επιλογή για έκτακτο — το έκτακτο να ΜΗΝ μπαίνει στο κύριο»): είδος πάνω-πάνω· κάθε είδος = δικό του πρόχειρο/αρχείο
      h += '<div class="card"><div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px">' + KINDS.map(([k, t]) => '<button class="chip' + (P.kind === k ? " on" : "") + '" data-kind="' + E(k) + '" style="flex:1;justify-content:center;min-height:44px' + (k === "ΕΚΤΑΚΤΟ ΠΡΟΓΡΑΜΜΑ" && P.kind === k ? ";background:#FF3B30;border-color:#FF3B30" : "") + '">' + t + "</button>").join("") + "</div>"
        + (P.kind === "ΕΚΤΑΚΤΟ ΠΡΟΓΡΑΜΜΑ" ? '<div style="color:#FF3B30;font-weight:700;margin-bottom:6px">🚨 ΕΚΤΑΚΤΟ — χωριστό από το κύριο πρόγραμμα της ημέρας</div>' : "")
        + '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap"><input type="date" data-f="date" value="' + E(P.date) + '">'
        + '<span style="color:#6B7280;font-size:13px">' + (P.status === "submitted" ? "✅ υποβλήθηκε" + (P.file ? " · " + E(P.file) : "") : dirty ? "● μη αποθηκευμένο" : P.at ? "💾 πρόχειρο " + new Date(P.at).toLocaleTimeString("el-GR", { hour: "2-digit", minute: "2-digit" }) + (P.by ? " · " + E(P.by) : "") : "νέο") + "</span></div>";
      // 8/11: οι 8 ειδικευόμενοι (από τα προγράμματα του φακέλου) σε λίστα επιλογής· «＋ άλλος…» μόνο για νέο όνομα
      const RES = [...new Set([...RESIDENTS, ...(cfg.residents || []), ...(P.resp ? [P.resp] : [])])];
      h += '<div style="margin-top:8px;font-weight:700">Υπεύθυνος προγράμματος</div><select data-f="resp" style="width:100%;margin-top:4px"><option value="">— διάλεξε ειδικευόμενο —</option>' + RES.map((n) => "<option" + (P.resp === n ? " selected" : "") + ">" + E(n) + "</option>").join("") + '<option value="__add">＋ άλλος…</option></select>';
      h += (P.pts.length && window.pdxXrQueue ? '<button class="chip" data-act="xr" style="margin-top:8px;background:#0F172A;color:#fff;border-color:#0F172A">🫁 Ακτινογραφίες θώρακα (' + P.pts.length + ') → Νίκος CT Viewer</button>' : "");
      h += '<label class="chip" style="margin-top:8px;cursor:pointer;' + (photoBusy ? "opacity:.6" : "") + '">' + (photoBusy ? "⏳ Διαβάζω το πρόχειρο…" : "📷 Από φωτογραφία χειρόγραφου προχείρου") + '<input type="file" accept="image/*" capture="environment" data-photo style="display:none"' + (photoBusy ? " disabled" : "") + "></label>";
      h += '<div style="margin-top:8px;font-weight:700">Αίθουσες</div><div>' + ROOMS_ALL.map((r) => chip("data-room", r, P.rooms.indexOf(r) >= 0, E(ORP.roomLabel(r).replace("ΑΙΘΟΥΣΑ ", "")))).join("") + "</div></div>";
      for (const R of ORP.rooms(P)) {
        h += '<div class="rm">' + E(R.label) + " · " + R.pts.length + "</div>";
        R.pts.forEach((p, i) => {
          h += '<div class="card" data-open="' + E(p.id) + '"><div style="display:flex;align-items:flex-start;gap:6px"><div style="flex:1;min-width:0"><b>' + (i + 1) + ". " + (p.key ? E(p.nm || "—") + "</b>" : '</b><input class="ie" data-ie="nm" data-pid="' + E(p.id) + '" value="' + E(p.nm) + '" placeholder="✍ ΕΠΩΝΥΜΟ ΟΝΟΜΑ ΤΟΥ ΠΑΤΡΟΣ ΑΜ (εκτός κλινικής)" style="width:calc(100% - 30px);font-weight:700">')
            // ✏ 16/11 (χρήστης: «να δηλώνω θάλαμο-κρεβάτι π.χ. 345-1 και ημ. εισόδου — προκαθορισμένη από το εισιτήριο, αλλά να την αλλάζω»): πεδία μέσα στην κάρτα
            + '<div style="display:flex;gap:6px;margin-top:4px;flex-wrap:wrap"><label class="iel">Θάλ.<input class="ie" data-ie="bed" data-pid="' + E(p.id) + '" value="' + E(p.bed) + '" placeholder="345-1" style="width:78px"></label><label class="iel">Εισ.<input class="ie" data-ie="adm" data-pid="' + E(p.id) + '" value="' + E(p.adm) + '" placeholder="ηη/μμ" style="width:70px"></label></div>'
            + (p.warn ? '<div style="margin-top:4px;padding:4px 8px;border-radius:8px;background:#FFF4E5;color:#9A3412;font-size:13px">' + E(p.warn) + ' <button class="mini" data-wok="' + E(p.id) + '" style="background:#fff;color:#9A3412">✓ ΟΚ</button></div>' : "") + '<div style="margin-top:3px">' + E(p.dx || "") + '</div><div style="font-weight:600">' + E(ORP.procText(p) || "⚠ χωρίς επέμβαση") + '</div><div style="color:#6B7280;font-size:13px">' + E([p.an, p.bl && "αίμα " + p.bl, p.ca && "c-ARM", p.surg].filter(Boolean).join(" · ")) + "</div></div>"
            + '<span style="display:flex;flex-direction:column"><button class="mini" data-up="' + E(p.id) + '">▲</button><button class="mini" data-dn="' + E(p.id) + '" style="margin-top:4px">▼</button>'
            // 🗑 16/11 (χρήστης: «στον υπολογιστή δεν μπορώ να αφαιρέσω ασθενή»): αφαίρεση από τη λίστα· 1ο πάτημα «Σίγουρα;», 2ο αφαιρεί (χωρίς παράθυρο browser)
            + '<button class="mini" data-rmp="' + E(p.id) + '" title="Αφαίρεση" style="margin-top:4px;' + (rmArm === p.id ? "background:#FF3B30;color:#fff;min-width:74px" : "background:#FDECEC;color:#B42318") + '">' + (rmArm === p.id ? "Σίγουρα;" : "🗑") + "</button></span></div></div>";
        });
        h += '<div style="margin:2px 12px"><button class="chip" data-add="' + E(R.r) + '">＋ Ασθενής στην ' + E(R.label.replace("ΑΙΘΟΥΣΑ", "αίθουσα")) + '</button><button class="chip" data-man="' + E(R.r) + '">✍ Εκτός κλινικής (γράφω στοιχεία)</button></div>';
      }
      if (!P.rooms.length) h += '<div class="card">Διαλέξτε αίθουσες.</div>';
      h += '<div class="card" style="overflow:auto;padding:6px"><div style="font-weight:700;margin:4px 6px">Προεπισκόπηση</div><div style="color:#6B7280;font-size:12px;margin:0 6px 4px">✏ Πατήστε σε οποιοδήποτε κελί για αλλαγή</div><div data-prev style="zoom:.65;min-width:1100px">' + ORP.html(P, { bodyOnly: true, click: true }) + "</div></div><div style=\"height:20px\"></div>";
      h += '<div class="bar"><button data-act="save" style="background:#EEF2F6;color:#0A84FF">💾 Πρόχειρο</button><button data-act="print" style="background:#EEF2F6;color:#0A84FF">🖨 Εκτύπωση</button><button data-act="submit" style="background:#34C759;color:#fff"' + (can ? "" : " disabled") + ">✅ Υποβολή</button></div></div>";
      return h;
    };
    const pg = window.iosPage("Πρόγραμμα χειρουργείου", "Ρυθμίσεις", draw);
    const redraw = () => { const y = pg.el.querySelector(".ios-body") && pg.el.querySelector(".ios-body").scrollTop; pg.redraw(); if (y) pg.el.querySelector(".ios-body").scrollTop = y; };
    const load = async (date, kind) => {
      kind = kind || P.kind || "ΠΡΟΓΡΑΜΜΑ"; loaded = false; redraw();
      try { if (CK() && CK().ready) await CK().ready(); cfg = await loadCfg(); const d = await loadDraft(date, kind);
        // παλιό πρόχειρο με άλλο είδος στο ίδιο κλειδί (π.χ. «ΕΚΤΑΚΤΟ» γραμμένο πάνω στο κύριο πριν το 16/11) → εμφανίζεται όπως είναι, με το είδος του
        P = d || { date, kind, resp: "", rooms: ["10", "8"], pts: [], status: "draft" }; if (!P.kind) P.kind = kind; dirty = false; }
      catch (e) { say("⚠ " + (e.message || e)); P = { date, kind, resp: "", rooms: ["10", "8"], pts: [], status: "draft" }; }
      loaded = true; redraw();
    };
    // 💾 16/11: ΠΡΩΤΑ αποθήκευση, ΜΕΤΑ αλλαγή είδους/ημερομηνίας (πριν: ερώτηση «χωρίς αποθήκευση;» → χανόταν ασθενής)
    const saveNow = async () => { clearTimeout(asT); if (!dirty || !loaded) return true; try { await saveDraft(P); dirty = false; return true; } catch (er) { say("⚠ δεν αποθηκεύτηκε: " + (er.message || er)); return false; } };
    let rmArm = "", rmT = 0;
    const removePt = (p) => { P.pts = P.pts.filter((x) => x !== p); renum(p.room); rmArm = ""; touch(); say("🗑 Αφαιρέθηκε: " + (p.nm || "ασθενής")); };
    const switchKind = async (k) => { if (k === P.kind) return; if (!(await saveNow())) return;
      const date = k === "ΕΚΤΑΚΤΟ ΠΡΟΓΡΑΜΜΑ" ? todayIso() : P.kind === "ΕΚΤΑΚΤΟ ΠΡΟΓΡΑΜΜΑ" ? ORP.defDate() : P.date; await load(date, k); };
    // 💾 9/11 (χρήστης: «με το που προσθέτω ασθενή να κάνει αυτόματα πρόχειρο — να μη χαθεί τίποτα»): αποθήκευση 0,6″ μετά από κάθε αλλαγή
    let asT = 0, asBusy = false, asAgain = false;
    function autoSave() { clearTimeout(asT); asT = setTimeout(async () => { if (!dirty || !loaded) return; if (asBusy) { asAgain = true; return; } asBusy = true;
      try { await saveDraft(P); dirty = false; window.__orpAutoAt = Date.now(); } catch (er) { say("⚠ αυτόματη αποθήκευση: " + (er.message || er)); }
      asBusy = false; if (asAgain) { asAgain = false; autoSave(); } }, 150); }
    const flush = () => { if (dirty && loaded && !asBusy) { clearTimeout(asT); saveDraft(P).then(() => { dirty = false; }, () => {}); } };
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") flush(); });
    window.addEventListener("pagehide", flush);
    const touch = () => { dirty = true; autoSave(); if (P.status === "submitted") P.status = "changed"; redraw(); };
    const renum = (room) => P.pts.filter((p) => p.room === room).sort((a, b) => (a.n || 0) - (b.n || 0)).forEach((p, i) => { p.n = i + 1; });
    pg.el.addEventListener("change", async (e) => { const f = e.target.dataset.f; if (!f) return;
      if (f === "resp" && e.target.value === "__add") { const n = UP(prompt("Επώνυμο ειδικευόμενου:") || "").trim(); if (n) { cfg.residents = [...new Set([...(cfg.residents || []), n])].sort((a, b) => a.localeCompare(b, "el")); try { await saveCfg(cfg); } catch (er) {} P.resp = n; } touch(); return; } if (f === "date") { const nd = e.target.value; if (!(await saveNow())) { e.target.value = P.date; return; } load(nd, P.kind); return; } if (f === "kind") { switchKind(e.target.value); return; } P[f] = e.target.value; touch(); });
    pg.el.addEventListener("click", async (e) => {
      if (e.target.closest("input,select,textarea,label.iel")) return;
      const c = e.target.closest("[data-cell]"); if (c) { const [pid, f] = c.dataset.cell.split("|"); const p = P.pts.find((x) => x.id === pid); if (p) edit(p, f); return; }
      const b = e.target.closest("button,[data-open]"); if (!b) return; const ds = b.dataset;
      if (ds.man) { addManual(ds.man); return; }
      if (ds.wok) { const p = P.pts.find((x) => x.id === ds.wok); if (p) { if (p.praw && p.t) aliasLearn(p.praw, p.t); delete p.warn; touch(); } return; }   // ✓ = σωστή αντιστοίχιση → μαθαίνεται
      if (ds.kind) { switchKind(ds.kind); return; }
      if (ds.rmp) { e.stopPropagation(); const p = P.pts.find((x) => x.id === ds.rmp); if (!p) return; clearTimeout(rmT);
        if (rmArm === ds.rmp) { removePt(p); return; } rmArm = ds.rmp; redraw(); rmT = setTimeout(() => { rmArm = ""; redraw(); }, 4000); return; }
      if (ds.room) { const i = P.rooms.indexOf(ds.room); if (i >= 0) { if (P.pts.some((p) => p.room === ds.room)) { say("Η αίθουσα έχει ασθενείς"); return; } P.rooms.splice(i, 1); } else P.rooms.push(ds.room); touch(); return; }
      if (ds.resp) { P.resp = P.resp === ds.resp ? "" : ds.resp; touch(); return; }
      if (ds.respAdd !== undefined) { const n = UP(prompt("Επώνυμο ειδικευόμενου (όπως θα γράφεται):") || "").trim(); if (!n) return; cfg.residents = [...new Set([...(cfg.residents || []), n])].sort((a, b) => a.localeCompare(b, "el")); try { await saveCfg(cfg); } catch (er) {} P.resp = n; touch(); return; }
      if (ds.up || ds.dn) { const p = P.pts.find((x) => x.id === (ds.up || ds.dn)); const L = P.pts.filter((x) => x.room === p.room).sort((a, b) => a.n - b.n); const i = L.indexOf(p), j = ds.up ? i - 1 : i + 1; if (j < 0 || j >= L.length) return; const t = L[j].n; L[j].n = p.n; p.n = t; touch(); return; }
      if (ds.add) { pick(ds.add); return; }
      if (ds.open) { edit(P.pts.find((x) => x.id === ds.open)); return; }
      if (ds.act === "save") { try { await saveDraft(P); dirty = false; say("💾 Αποθηκεύτηκε (πρόχειρο)"); redraw(); } catch (er) { say("⚠ " + er.message); } return; }
      if (ds.act === "print") { printViaExcel(); return; }
      if (ds.act === "submit") { submit(); return; }
      if (ds.act === "xr") { xrQueue(); return; }
    });
    // ---- επιλογή ασθενή ----
    function sheet(inner) { const ov = document.createElement("div"); ov.style.cssText = "position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:9999;display:flex;align-items:flex-end"; ov.innerHTML = '<div style="background:#F2F2F7;width:100%;max-height:88vh;overflow:auto;border-radius:16px 16px 0 0;padding:12px 12px calc(16px + env(safe-area-inset-bottom));font:15px -apple-system,system-ui;color:#1F2933" id="orpSh">' + inner + "</div>"; ov.addEventListener("click", (e) => { if (e.target === ov) ov.remove(); }); document.body.appendChild(ov); return ov; }
    function pick(room) {
      const used = new Set(P.pts.map((p) => p.key).filter(Boolean));
      const L = inpatients().filter((x) => !used.has(x.k));
      const ov = sheet('<div style="display:flex;align-items:center;margin-bottom:8px"><b style="flex:1;font-size:17px">' + E(ORP.roomLabel(room)) + '</b><button data-x style="border:0;background:none;font-size:22px">✕</button></div><input data-q placeholder="Αναζήτηση ονόματος/κρεβατιού" style="width:100%;box-sizing:border-box;font-size:16px;padding:9px;border:1px solid #C7CDD4;border-radius:10px">'
        + '<div data-list></div><div data-npsall></div><button data-man style="margin-top:10px;width:100%;min-height:44px;border:0;border-radius:12px;background:#fff;color:#0A84FF;font-weight:700">✍ Χειροκίνητη προσθήκη</button>');
      const list = ov.querySelector("[data-list]"), nb = ov.querySelector("[data-npsall]");
      // 🏥 12/11 (χρήστης: «ασθενείς από ΟΛΟ το ΝΠΣ — ναι ακριβώς»): αναζήτηση επωνύμου σε όλο το νοσοκομείο (εντολή npsfind, εκτός κλινικής, από 2000)
      let ns = { q: "", busy: false, msg: "" };
      const nbDraw = (q) => { const t = String(q || "").trim(); nb.innerHTML = t.length >= 3 ? '<button data-npsq style="margin-top:8px;width:100%;min-height:44px;border:1px dashed #0A84FF;border-radius:12px;background:#fff;color:#0A84FF;font-weight:700"' + (ns.busy ? " disabled" : "") + ">" + (ns.busy ? "⏳ " + E(ns.msg || "Το bot ψάχνει στο ΝΠΣ…") : "🏥 Αναζήτηση «" + E(t) + "» σε όλο το ΝΠΣ") + "</button>" + (ns.msg && !ns.busy ? '<div style="margin-top:6px;color:#6B7280;font-size:13px">' + E(ns.msg) + "</div>" : "") : ""; };
      const show = (q) => { const z = NZ(q); list.innerHTML = L.filter((x) => !z || NZ(x.d.name + " " + x.k).indexOf(z) >= 0).map((x) => '<div data-k="' + E(x.k) + '" style="background:#fff;margin-top:6px;padding:10px 12px;border-radius:10px;display:flex;gap:8px;align-items:center"><span style="flex:1"><b>' + E(x.d.name) + '</b><br><span style="color:#6B7280;font-size:13px">' + E([x.all ? "🏥 ΝΠΣ · εκτός κλινικής" + (x.d.lastAdm ? " · τελ. εισ. " + x.d.lastAdm : "") : x.nps ? "🏥 ΝΠΣ (όχι στο Παραδόχαρτο)" : x.k.replace(/^extra-\d+$/, "εκτός θαλάμου"), x.d.therapist, x.d.surgery || x.d.surgeryCustom, x.d.pod === "Προς Χ/Ο" ? "προς Χ/Ο" : ""].filter(Boolean).join(" · ")) + "</span></span><span style=\"color:#0A84FF;font-weight:700\">＋</span></div>").join("") || '<div style="padding:12px;color:#6B7280">Κανένας</div>'; };
      show(""); nbDraw("");
      ov.querySelector("[data-q]").addEventListener("input", (e) => { show(e.target.value); if (!ns.busy) ns.msg = ""; nbDraw(e.target.value); });
      const npsAll = async () => { const q = String(ov.querySelector("[data-q]").value || "").trim(); if (q.length < 3 || ns.busy) return; ns = { q, busy: true, msg: "" }; nbDraw(q);
        try { const ins = await SB().from("oncall_jobs").insert({ name_key: "zzz_find", name: "🔎 ΝΠΣ: " + q.slice(0, 40), room: null, kind: "npsfind", payload: { q, quiet: true }, status: "new" }).select("id").single(); if (ins.error) throw new Error(ins.error.message);
          let x = null; for (let i = 0; i < 90 && ov.isConnected; i++) { await new Promise((r) => setTimeout(r, 2000)); const r = await SB().from("oncall_jobs").select("status,note,payload").eq("id", ins.data.id).maybeSingle(); x = r.data;
            if (x && /^(done|failed|dismissed|cancelled)$/.test(x.status)) break; if (x && x.note) { ns.msg = String(x.note).slice(0, 80); nbDraw(q); } x = null; }
          if (!x) throw new Error("το bot δεν απάντησε"); if (x.status !== "done") throw new Error(String(x.note || "απέτυχε"));
          const R = await CK().decrypt(x.payload.result.enc); const H = R.heads || [], ix = (re) => H.findIndex((h) => re.test(h));
          const iN = ix(/Ον\/?νυμο|Επώνυμο|Όνομα/i), iA = ix(/ΑΜΚΑ/), iM = ix(/^Α\.?Μ\.?$|Μητρώ|Κωδικός/i), iI = ix(/Εισαγ/i);
          const have = new Set(L.map((y) => NZ(y.d.name))), seen = new Set(); let n = 0;
          for (const c of R.rows || []) { const nm = String(c[iN] || "").replace(/\s+/g, " ").trim(); if (!nm) continue; const am = iM >= 0 ? String(c[iM] || "").trim() : "", amka = iA >= 0 ? String(c[iA] || "").replace(/\D/g, "") : "";
            const id = "npsall:" + (am || amka || nm); if (seen.has(id) || used.has(id)) continue; seen.add(id);
            if (have.has(NZ(nm))) continue;   // ήδη στη λίστα (Παραδόχαρτο / ΝΠΣ κλινικής)
            L.push({ k: id, d: { name: nm, am, amka, lastAdm: iI >= 0 ? String(c[iI] || "").trim() : "" }, nps: true, all: true }); n++; }
          ns = { q, busy: false, msg: n ? "🏥 " + n + " από όλο το ΝΠΣ (στην κορυφή της λίστας)" : "Κανένας επιπλέον στο ΝΠΣ για «" + q + "»" };
          L.sort((a, b) => (b.all ? 1 : 0) - (a.all ? 1 : 0)); show(q); }
        catch (er) { ns = { q, busy: false, msg: "⚠ " + (er.message || er) }; }
        if (ov.isConnected) nbDraw(ov.querySelector("[data-q]").value); };
      // η λίστα ΝΠΣ φορτώνει την 1η φορά → ξανά σε λίγο (χωρίς να χαθεί η αναζήτηση)
      if (!L.some((x) => x.nps)) setTimeout(() => { if (!ov.isConnected) return; const add = inpatients().filter((x) => x.nps && !used.has(x.k) && !L.some((y) => y.k === x.k)); if (add.length) { L.push(...add); show(ov.querySelector("[data-q]").value); } }, 1800);
      ov.addEventListener("click", (e) => {
        if (e.target.closest("[data-x]")) { ov.remove(); return; }
        if (e.target.closest("[data-npsq]")) { npsAll(); return; }
        const r = e.target.closest("[data-k]");
        if (r) { const x = L.find((y) => y.k === r.dataset.k); const p = fromPatient(x.k, x.d); p.room = room; p.n = P.pts.filter((q) => q.room === room).length + 1; P.pts.push(p); ov.remove(); touch(); if (!p.t) edit(p); return; }
        if (e.target.closest("[data-man]")) { ov.remove(); addManual(room); }
      });
    }
    // ✍ 16/11 (χρήστης: «να μπορώ να βάλω ασθενή εκτός κλινικής — να γράφω εγώ στοιχεία»): κενή κάρτα, ημ. εισόδου = σήμερα (αλλάζει)
    function addManual(room) { const t = new Date(); const p = { id: "m" + Date.now().toString(36), key: "", room, n: P.pts.filter((q) => q.room === room).length + 1, bed: "", adm: String(t.getDate()).padStart(2, "0") + "/" + String(t.getMonth() + 1).padStart(2, "0"), nm: "", dx: "", an: "ΓΕΝΙΚΗ", bl: "", ca: "", t: "", side: "", t2: "", side2: "", surg: "" }; P.pts.push(p); touch(); edit(p, "nm"); }
    // 📷 16/11 (χρήστης: «σου δίνω το χειρόγραφο πρόχειρο και φτιάχνεις όσους γράφει, μετά προσθέτω όσους λείπουν — αποθήκευση σημαντικό»):
    // φωτογραφία → AI διαβάζει γραμμές (αίθουσα, σειρά, επώνυμο, θάλαμος, επέμβαση, πλευρά, χειρουργός, αναισθησία, αίμα, c-ARM) → ταίριασμα ΤΟΠΙΚΑ
    // με τους ασθενείς της κλινικής (κανένα όνομα δεν στέλνεται στο AI πέρα από τη φωτογραφία) και με τα έτοιμα σετ → μπαίνουν όλοι + αυτόματη αποθήκευση.
    const imgJpeg = (file) => new Promise((ok, no) => { const fr = new FileReader(); fr.onerror = () => no(new Error("δεν διαβάστηκε η φωτογραφία")); fr.onload = () => { const im = new Image(); im.onerror = () => no(new Error("μη έγκυρη εικόνα"));
      im.onload = () => { const M = 2200, k = Math.min(1, M / Math.max(im.width, im.height)), c = document.createElement("canvas"); c.width = Math.round(im.width * k); c.height = Math.round(im.height * k); c.getContext("2d").drawImage(im, 0, 0, c.width, c.height); ok(c.toDataURL("image/jpeg", 0.85).split(",")[1]); }; im.src = fr.result; }; fr.readAsDataURL(file); });
    const PHOTO_SYS = "Διαβάζεις χειρόγραφο πρόχειρο προγράμματος χειρουργείου ουρολογικής κλινικής. Επέστρεψε ΜΟΝΟ JSON: {\"rows\":[{\"room\":\"10|9|8|8Α ή κενό\",\"n\":σειρά,\"surname\":\"ΕΠΩΝΥΜΟ όπως γράφεται (κεφαλαία)\",\"first\":\"όνομα αν υπάρχει\",\"bed\":\"θάλαμος-κρεβάτι αν υπάρχει π.χ. 345-1\",\"proc\":\"επέμβαση όπως γράφεται (π.χ. TURB, Lap Nx, ΟΛΨ, Double J)\",\"side\":\"ΑΡ|ΔΕ|ΑΜΦΩ ή κενό\",\"surgeon\":\"χειρουργός/συντομογραφία αν υπάρχει\",\"an\":\"ΓΕΝΙΚΗ|ΤΟΠΙΚΗ ή κενό\",\"blood\":\"ml ή κενό\",\"carm\":true/false,\"unsure\":\"ό,τι δεν διαβάζεται καθαρά\"}]}. Μία γραμμή ανά ασθενή, με τη σειρά του χαρτιού. Αίθουσα: από επικεφαλίδες/στήλες (π.χ. «10», «Αιθ. 8», «8 απόγευμα» = 8Α). Μην επινοείς ό,τι δεν γράφεται — άφησέ το κενό.";
    async function fromPhoto(file) {
      if (!(window.pdxAiOn && window.pdxAiOn())) { say("Οι λειτουργίες AI είναι κλειστές (Ρυθμίσεις → 🤖) — προσθέστε τους ασθενείς με «＋ Ασθενής»"); return; }
      say("📷 Διαβάζω το πρόχειρο…"); photoBusy = true; redraw();
      try {
        const b64 = await imgJpeg(file);
        const r = await window.pdxAI("πρόγραμμα χειρουργείου", { model: "claude-sonnet-5-5", max_tokens: 4000, system: PHOTO_SYS, messages: [{ role: "user", content: [{ type: "image", source: { type: "base64", media_type: "image/jpeg", data: b64 } }, { type: "text", text: "Ημερομηνία προγράμματος: " + ORP.dShort(P.date) + ". Επέστρεψε το JSON." }] }] });
        if (!r.ok) throw new Error(r.error || "το AI δεν απάντησε");
        const t = String(r.text || ""), a = t.indexOf("{"), z = t.lastIndexOf("}"); let J; try { J = JSON.parse(t.slice(a, z + 1)); } catch (e) { throw new Error("η απάντηση δεν διαβάστηκε"); }
        const rows = (J.rows || []).filter((x) => x && String(x.surname || "").trim());
        if (!rows.length) throw new Error("δεν βρέθηκαν ασθενείς στη φωτογραφία");
        const res = photoApply(rows); touch(); await saveNow();
        photoReport(res);
      } catch (er) { say("⚠ " + (er.message || er)); }
      photoBusy = false; redraw();
    }
    // καθαρή συνάρτηση (δοκιμάζεται χωρίς AI): γραμμές → ασθενείς στο P
    function photoApply(rows) {
      const L = inpatients(), used = new Set(P.pts.map((p) => p.key).filter(Boolean)), usedNm = new Set(P.pts.map((p) => NZ(String(p.nm || "").split(" ")[0])));
      const bedK = (b) => String(b || "").replace(/[.\/ ]/g, "-").replace(/-+/g, "-").trim();
      const out = { added: [], skipped: [] };
      for (const x of rows) {
        const sn = NZ(x.surname), fn = NZ(x.first || ""), bk = bedK(x.bed);
        let room = String(x.room || "").toUpperCase().replace(/A$/, "Α").replace(/\s+/g, ""); if (ROOMS_ALL.indexOf(room) < 0) room = P.rooms[0] || "10";
        if (P.rooms.indexOf(room) < 0) P.rooms.push(room);
        let C = L.filter((y) => { const w = NZ(String(y.d.name || "").split(/\s+/)[0]); return w && sn && (w === sn || (sn.length >= 5 && (w.indexOf(sn) === 0 || sn.indexOf(w) === 0))); });
        if (C.length > 1 && fn) { const C2 = C.filter((y) => NZ(y.d.name).indexOf(fn.slice(0, 3)) >= 0); if (C2.length) C = C2; }
        if (C.length > 1 && bk) { const C2 = C.filter((y) => y.k === bk); if (C2.length) C = C2; }
        if (C.length > 1) { const C2 = C.filter((y) => !y.nps); if (C2.length) C = C2; }
        const hit = C.length === 1 ? C[0] : null;
        if (hit && used.has(hit.k)) { out.skipped.push(x.surname + " (ήδη στο πρόγραμμα)"); continue; }
        if (!hit && usedNm.has(sn)) { out.skipped.push(x.surname + " (ήδη στο πρόγραμμα)"); continue; }
        const p = hit ? fromPatient(hit.k, hit.d) : { id: "f" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), key: "", bed: bedNorm(x.bed), adm: "", nm: UP([x.surname, x.first].filter(Boolean).join(" ")), dx: "", an: "ΓΕΝΙΚΗ", bl: "", ca: "", t: "", side: "", t2: "", side2: "", surg: "" };
        if (hit) used.add(hit.k); usedNm.add(sn);
        p.room = room; p.n = P.pts.filter((q) => q.room === room).length + 1;
        const warn = [];
        // επέμβαση: από τη φωτογραφία (υπερισχύει της συντομογραφίας του Παραδόχαρτου), αλλιώς ό,τι έφερε ο ασθενής
        const pq0 = String(x.proc || "").trim(), R = procResolve(pq0), pq = R ? R.raw : "";
        if (R && R.k) { const k = R.k; p.t = k.t; if (k.dx) { p.dx = k.dx; p.bl = k.bl || ""; p.ca = k.ca || ""; } if (k.an) p.an = k.an; p.praw = R.raw;
          if (!R.sure) warn.push("επέμβαση «" + R.raw + "» → «" + String(k.t).replace(/\s+\d{3}\.[\d.\/ +]*$/, "") + "» (πιο κοντινή — έλεγξε)"); }
        else if (pq0) { p.praw = pq || pq0; warn.push("επέμβαση «" + (pq || pq0) + "» χωρίς σετ"); }
        const sd = SIDES.indexOf(String(x.side || "").toUpperCase()) > 0 ? String(x.side).toUpperCase() : pkSide(pq); if (sd) p.side = sd;
        if (/^ΤΟΠ/i.test(String(x.an || ""))) p.an = "ΤΟΠΙΚΗ"; else if (/^ΓΕΝ/i.test(String(x.an || ""))) p.an = "ΓΕΝΙΚΗ";
        const bl = String(x.blood || "").replace(/\D/g, ""); if (BLOOD.indexOf(bl) > 0) p.bl = bl;
        if (x.carm === true) p.ca = "ΝΑΙ";
        const sg = UP(x.surgeon || "").replace(/^Κ\s+/, ""); if (sg) { const F = FULL(), key = Object.keys(F).find((k) => UP(k) === sg || NZ(F[k]).indexOf(NZ(sg)) === 0); if (key) p.surg = "κ " + UP(F[key]); }
        if (!hit) warn.push("εκτός κλινικής / δεν βρέθηκε — συμπληρώστε στοιχεία");
        if (x.unsure) warn.push("δυσανάγνωστο: " + String(x.unsure).slice(0, 60));
        if (warn.length) p.warn = "📷 " + warn.join(" · ");
        P.pts.push(p); out.added.push({ nm: p.nm, ok: !!hit && !warn.length });
      }
      return out;
    }
    function photoReport(res) {
      const ok = res.added.filter((x) => x.ok).length, chk = res.added.length - ok;
      say("📷 Μπήκαν " + res.added.length + " ασθενείς" + (chk ? " · ⚠ " + chk + " για έλεγχο (πορτοκαλί)" : "") + (res.skipped.length ? " · " + res.skipped.length + " ήταν ήδη" : "") + " · 💾 αποθηκεύτηκε");
    }
    let photoBusy = false;
    window.__orpPhotoApply = (rows) => { const r = photoApply(rows); touch(); return r; };   // για τις δοκιμές
    // θάλαμος-κρεβάτι: «345-1» / «345 1» / «345/1» → «345.1» (όπως στο πρόγραμμα της κλινικής)
    const bedNorm = (v) => String(v || "").trim().toUpperCase().replace(/^(\d{3}Α?)\s*[-\/ ,]\s*(\d{1,2})$/, "$1.$2");
    const admNorm = (v) => String(v || "").trim().replace(/^(\d{1,2})[\/.\-](\d{1,2})(?:[\/.\-]\d{2,4})?$/, (x, a, b) => a.padStart(2, "0") + "/" + b.padStart(2, "0"));
    const prevUpd = () => { const pv = pg.el.querySelector("[data-prev]"); if (pv) pv.innerHTML = ORP.html(P, { bodyOnly: true, click: true }); };
    pg.el.addEventListener("input", (e) => { const t = e.target; if (!t.dataset || !t.dataset.ie) return; const p = P.pts.find((x) => x.id === t.dataset.pid); if (!p) return;
      p[t.dataset.ie] = t.dataset.ie === "nm" ? t.value.toUpperCase() : t.value; dirty = true; if (P.status === "submitted") P.status = "changed"; autoSave(); });
    pg.el.addEventListener("change", (e) => { if (e.target.dataset && e.target.dataset.photo !== undefined) { const f = e.target.files && e.target.files[0]; e.target.value = ""; if (f) fromPhoto(f); return; } });
    pg.el.addEventListener("change", (e) => { const t = e.target; if (!t.dataset || !t.dataset.ie) return; const p = P.pts.find((x) => x.id === t.dataset.pid); if (!p) return;
      if (t.dataset.ie === "bed") { p.bed = bedNorm(t.value); t.value = p.bed; } if (t.dataset.ie === "adm") { p.adm = admNorm(t.value); t.value = p.adm; } dirty = true; autoSave(); prevUpd(); });
    // ---- επεξεργασία ασθενή ----
    function edit(p, fld) {
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
          + box("Επέμβαση (έτοιμο σετ)", '<input data-pkq autocomplete="off" placeholder="🔎 Γράψε επέμβαση (π.χ. TURB, νεφρεκτομή αρ, 357.03)" style="width:100%;box-sizing:border-box;font-size:16px;padding:9px;border:1px solid #0A84FF;border-radius:10px;margin-bottom:4px"><div data-pkdd></div><button data-pk="1" style="width:100%;text-align:left;min-height:44px;border:1px solid #C7CDD4;border-radius:10px;background:#fff;padding:8px 10px;font-size:15px">' + (p.t ? "<b>" + E(p.dx) + "</b><br>" + E(p.t) : "🔎 Διάλεξε…") + "</button>"
            + '<div style="margin-top:4px">' + SIDES.map((s) => ch("sd", s, (p.side || "") === s, s || "χωρίς πλευρά")).join("") + "</div>")
          + box("Δεύτερη επέμβαση", (p.t2 ? '<div style="display:flex;gap:6px;align-items:center"><span style="flex:1">' + E(p.t2) + '</span><button data-t2x style="border:0;background:#FDECEC;color:#B42318;border-radius:8px;min-height:34px;padding:0 10px">✕</button></div><div>' + SIDES.map((s) => ch("sd2", s, (p.side2 || "") === s, s || "χωρίς πλευρά")).join("") + "</div>" : '<button data-pk="2" style="min-height:40px;border:1px dashed #C7CDD4;border-radius:10px;background:none;padding:6px 12px;color:#0A84FF">＋ δεύτερη επέμβαση</button>'))
          + box("Διάγνωση", inp("dx", p.dx, "π.χ. ΑΠΟΦΡΑΚΤΙΚΗ ΟΥΡΟΠΑΘΕΙΑ"))
          + box("Αναισθησία", ANS.map((a) => ch("an", a, p.an === a)).join(""))
          + box("Αίμα", BLOOD.map((a) => ch("bl", a, (p.bl || "") === a, a || "—")).join("") + '<button data-ca style="margin-left:10px;min-height:36px;padding:6px 12px;border-radius:18px;border:1px solid ' + (p.ca ? "#0A84FF" : "#C7CDD4") + ";background:" + (p.ca ? "#0A84FF" : "#fff") + ";color:" + (p.ca ? "#fff" : "#1F2933") + ';font-weight:600">c-ARM ' + (p.ca ? "ΝΑΙ" : "—") + "</button>")
          + box("Χειρουργός", SURGEONS().map((s) => ch("sg", s, p.surg === s, s.replace(/^κ /, ""))).join("") + inp("surg", p.surg, "κ ΕΠΩΝΥΜΟ"))
          + (window.pdxOpNote ? box("Πρακτικό χειρουργείου", '<button data-opn style="width:100%;text-align:left;min-height:44px;border:1px solid #C7CDD4;border-radius:10px;background:#fff;padding:8px 10px;font-size:15px">📝 ' + (p.opnote ? "<b>Αποθηκευμένο</b> · " + E(String(p.opnote).slice(0, 60)) + "…" : "Από τα πρακτικά της κλινικής ›") + "</button>") : "")
          + '<div style="display:flex;gap:8px;margin-top:14px"><button data-del style="flex:1;min-height:46px;border:0;border-radius:12px;background:#FDECEC;color:#B42318;font-weight:700">Αφαίρεση</button><button data-x style="flex:2;min-height:46px;border:0;border-radius:12px;background:#0A84FF;color:#fff;font-weight:700">Τέλος</button></div>';
        if (!ov) ov = sheet(inner); else ov.querySelector("#orpSh").innerHTML = inner;
      };
      render();
      if (fld) (() => { const el = ov.querySelector(fld === "proc" ? "[data-pkq]" : '[data-e="' + fld + '"]'); if (el) { el.focus(); if (el.select) el.select(); } else { const b = ov.querySelector(fld === "an" ? "[data-an]" : fld === "bl" ? "[data-bl]" : fld === "ca" ? "[data-ca]" : "x"); if (b) b.scrollIntoView({ block: "center" }); } })();
      let pkItems = [];
      ov.addEventListener("input", (e) => { if (e.target.dataset.pkq !== undefined) { const q = e.target.value, dd = ov.querySelector("[data-pkdd]"); pkItems = q.trim().length >= 2 ? pkFind(q, 8) : [];
          dd.innerHTML = pkItems.length ? '<div style="border:1px solid #C7CDD4;border-radius:10px;background:#fff;overflow:hidden;margin-bottom:6px">' + pkItems.map((x, i) => '<button data-pki="' + i + '" style="display:block;width:100%;text-align:left;border:0;border-bottom:1px solid #EEF2F6;background:#fff;padding:8px 10px;font-size:14px;cursor:pointer">' + (x.dx ? "<b>" + E(x.dx) + "</b><br>" : "") + E(x.t) + (x.sub ? '<br><span style="color:#6B7280;font-size:12px">' + E(x.sub) + "</span>" : "") + "</button>").join("") + "</div>"
            : q.trim().length >= 2 ? '<div style="color:#6B7280;font-size:13px;margin:2px 4px 6px">Τίποτα — δοκιμάστε άλλη λέξη ή «🔎 Διάλεξε…» για όλη τη λίστα</div>' : ""; return; }
        const f = e.target.dataset.e; if (!f) return; p[f] = f === "surg" || f === "dx" || f === "nm" ? e.target.value.toUpperCase().replace(/^Κ /, "κ ") : e.target.value; dirty = true; autoSave(); });
      ov.addEventListener("change", (e) => { const f = e.target.dataset.e; if (f === "bed") { p.bed = bedNorm(e.target.value); e.target.value = p.bed; } if (f === "adm") { p.adm = admNorm(e.target.value); e.target.value = p.adm; } dirty = true; autoSave(); });
      ov.addEventListener("click", (e) => {
        const b = e.target.closest("button"); if (!b) return; const d = b.dataset;
        if (d.x !== undefined) { ov.remove(); touch(); return; }
        if (d.pki !== undefined) { const x = pkItems[+d.pki]; const sd = pkSide((ov.querySelector("[data-pkq]") || {}).value || ""); if (x) { if (sd && SIDES.indexOf(sd) >= 0) p.side = sd; p.t = x.t; if (p.praw) { aliasLearn(p.praw, x.t); if (p.warn && /επέμβαση «/.test(p.warn)) { p.warn = p.warn.split(" · ").filter((w) => !/επέμβαση «/.test(w)).join(" · "); if (p.warn === "📷" || !p.warn.replace("📷", "").trim()) delete p.warn; } } if (x.dx) p.dx = x.dx; if (x.an) p.an = x.an; if (x.dx) { p.bl = x.bl || ""; p.ca = x.ca || ""; } dirty = true; autoSave(); render(); } return; }
        if (d.del !== undefined) { if (b.dataset.armed !== "1") { b.dataset.armed = "1"; b.textContent = "Σίγουρα; Πάτα ξανά"; b.style.background = "#FF3B30"; b.style.color = "#fff"; setTimeout(() => { if (b.isConnected) { b.dataset.armed = ""; b.textContent = "Αφαίρεση"; b.style.background = "#FDECEC"; b.style.color = "#B42318"; } }, 4000); return; }
          ov.remove(); removePt(p); return; }
        if (d.opn !== undefined && window.pdxOpNote) { const proc = [p.t, p.t2].filter(Boolean).join(" + ") || p.dx || ""; window.pdxOpNote({ name: p.nm, proc, side: p.side || "", surg: p.opsurg || "", text: p.opnote || "", onSave: (r) => { p.opnote = r.text; p.opsurg = r.surg; dirty = true; autoSave(); render(); touch(); } }); return; }
        if (d.rm) { const old = p.room; p.room = d.rm; p.n = P.pts.filter((q) => q.room === d.rm && q !== p).length + 1; renum(old); }
        else if (d.sd !== undefined) p.side = d.sd;
        else if (d.sd2 !== undefined) p.side2 = d.sd2;
        else if (d.an) p.an = d.an;
        else if (d.bl !== undefined) p.bl = d.bl;
        else if (d.ca !== undefined) p.ca = p.ca ? "" : "ΝΑΙ";
        else if (d.sg) p.surg = d.sg;
        else if (d.t2x !== undefined) { p.t2 = ""; p.side2 = ""; }
        else if (d.pk) { pkPick((x) => { if (d.pk === "2") { p.t2 = x.t; } else { p.t = x.t; if (p.praw) aliasLearn(p.praw, x.t); if (x.dx) p.dx = x.dx; if (x.an) p.an = x.an; p.bl = x.bl || ""; p.ca = x.ca || ""; } dirty = true; autoSave(); render(); }); return; }
        else return;
        dirty = true; autoSave(); render();
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
        const toks = pkTok(q), L = !toks.length ? PK() : PK().filter((x) => pkHit(toks, x.dx + " " + x.t + " " + x.g) >= 0);
        let g = ""; for (const x of L) { if (x.g !== g && !z) { g = x.g; h += '<div style="margin:12px 4px 2px;font-weight:700;color:#6B7280;font-size:13px">' + E(g) + "</div>"; } items.push(x); h += row(items.length - 1, x.dx, x.t, [x.an, x.bl && "αίμα " + x.bl, x.ca && "c-ARM"].filter(Boolean).join(" · ")); }
        if (z.length >= 3) {
          const C = ((window.ORP_DATA || {}).codes || []).filter((c) => pkHit(toks, c.d + " " + c.c + " " + c.s) >= 0).slice(0, 30);
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
    // 🖨 16/11 (χρήστης: «να τυπώνεται ΑΚΡΙΒΩΣ όπως το Excel μας»): το ίδιο .xlsx τυπώνεται μέσα από το Excel του υπολογιστή του εκτυπωτή
    // (ΜΟΝΟ εκτύπωση — κανένα αρχείο στο δίκτυο). Αν το bot δεν απαντήσει σε ~45″ ή αποτύχει → τοπική εκτύπωση όπως πριν.
    async function printViaExcel() {
      if (!P.pts.length) { say("Το πρόγραμμα είναι άδειο"); return; }
      try {
        say("🖨 Στέλνω το Excel για εκτύπωση…");
        const x = await sealB64(ORP.xlsx(P)), hh = await sealB64(new TextEncoder().encode(ORP.html(P)));
        const ins = await SB().from("oncall_jobs").insert({ name_key: "orprint", name: "🖨 Εκτύπωση " + (P.kind && P.kind !== "ΠΡΟΓΡΑΜΜΑ" ? P.kind.replace(/ ΠΡΟΓΡΑΜΜΑ$/, "") + " " : "") + ORP.dShort(P.date), room: null, kind: "orsubmit", status: "new",
          payload: { x, h: hh, dir: ORP.monthDir(P.date), name: ORP.fileName(P.date, P.kind), noSave: true, printer: PRINTER, only: BOT, quiet: true } }).select("id").single();
        if (ins.error) throw new Error(ins.error.message);
        for (let i = 0; i < 30; i++) {
          await new Promise((r) => setTimeout(r, 1500));
          const r = await SB().from("oncall_jobs").select("status,note").eq("id", ins.data.id).maybeSingle(); const j = r.data || {};
          if (j.status === "done") { say(j.note || "🖨 Τυπώθηκε"); SB().from("oncall_jobs").update({ payload: { wiped: true } }).eq("id", ins.data.id).then(() => {}, () => {}); return; }
          if (j.status === "failed") { if (confirm("⚠ " + (j.note || "η εκτύπωση απέτυχε") + "\n\nΕκτύπωση από εδώ;")) printLocal(P); return; }
        }
        if (confirm("Ο υπολογιστής του εκτυπωτή δεν απάντησε ακόμη. Εκτύπωση από εδώ;")) printLocal(P);
      } catch (er) { if (confirm("⚠ " + (er.message || er) + "\n\nΕκτύπωση από εδώ;")) printLocal(P); }
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
    // 🗓 16/11: πρόχειρο από τη Σάρωση (scanrx «ordraft») → οι γραμμές μπαίνουν μόλις φορτώσει το πρόγραμμα της ημέρας + αυτόματη αποθήκευση
    load(P.date, P.kind).then(async () => { if (opts0 && Array.isArray(opts0.rows) && opts0.rows.length) { await aliasLoad(); const res = photoApply(opts0.rows); touch(); await saveNow(); photoReport(res); } });
    return pg;
  };
})();
