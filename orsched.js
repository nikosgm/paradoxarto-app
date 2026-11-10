// 🗓 8/11: Πρόγραμμα χειρουργείου (Excel) → (α) ΜΟΝΟ η περιοχή εκτύπωσης σε HTML (όπως η προεπισκόπηση), (β) λίστα ασθενών με ΑΜ.
// Χρειάζεται το SheetJS (XLSX). Κοινό για Παραδόχαρτο (window.pdxOrSched) και δοκιμές (module.exports).
(function (root) {
  "use strict";
  const E = (x) => String(x == null ? "" : x).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  // περιοχή εκτύπωσης: όνομα _xlnm.Print_Area του φύλλου, αλλιώς όλο το φύλλο
  function printArea(XLSX, wb) {
    const names = (wb.Workbook && wb.Workbook.Names) || [];
    for (let si = 0; si < wb.SheetNames.length; si++) {
      const n = names.find((x) => /^_xlnm\.Print_Area$/i.test(x.Name) && (x.Sheet === si || x.Sheet == null));
      if (n) { const ref = String(n.Ref).split(",")[0].replace(/^'?[^!']*'?!/, "").replace(/\$/g, ""); return { si, sheet: wb.SheetNames[si], range: XLSX.utils.decode_range(ref) }; }
    }
    const ws = wb.Sheets[wb.SheetNames[0]]; return { si: 0, sheet: wb.SheetNames[0], range: XLSX.utils.decode_range(ws["!ref"] || "A1:A1") };
  }
  const txt = (c) => (c ? String(c.w != null ? c.w : c.v != null ? c.v : "") : "");
  function render(XLSX, wb) {
    const pa = printArea(XLSX, wb), ws = wb.Sheets[pa.sheet], R = pa.range;
    const cols = ws["!cols"] || [], merges = (ws["!merges"] || []).filter((m) => m.s.r >= R.s.r && m.s.c >= R.s.c && m.s.r <= R.e.r && m.s.c <= R.e.c);
    const cover = new Set(), span = {};
    for (const m of merges) { span[m.s.r + ":" + m.s.c] = [Math.min(m.e.r, R.e.r) - m.s.r + 1, Math.min(m.e.c, R.e.c) - m.s.c + 1]; for (let r = m.s.r; r <= Math.min(m.e.r, R.e.r); r++) for (let c = m.s.c; c <= Math.min(m.e.c, R.e.c); c++) if (r !== m.s.r || c !== m.s.c) cover.add(r + ":" + c); }
    const wpx = (c) => { const k = cols[c] || {}; return Math.round(k.wpx || (k.wch ? k.wch * 7 + 5 : k.width ? k.width * 7 : 64)); };
    let tw = 0; for (let c = R.s.c; c <= R.e.c; c++) tw += wpx(c);
    let h = '<table class="ors" style="width:' + tw + 'px"><colgroup>'; for (let c = R.s.c; c <= R.e.c; c++) h += '<col style="width:' + wpx(c) + 'px">'; h += "</colgroup>";
    for (let r = R.s.r; r <= R.e.r; r++) {
      const rowT = []; for (let c = R.s.c; c <= R.e.c; c++) rowT.push(txt(ws[XLSX.utils.encode_cell({ r, c })]));
      const all = rowT.join(" ");
      const cls = /ΟΝΟΜΑΤΕΠΩΝΥΜΟ|ΕΠΕΜΒΑΣΗ/.test(all) ? "hd" : /^\s*ΑΙΘΟΥΣΑ/i.test(rowT[0]) ? "rm" : "";
      const rh = (ws["!rows"] || [])[r], hp = rh && (rh.hpx || (rh.hpt ? rh.hpt * 96 / 72 : 0));   // 10/11 βράδυ: ύψος γραμμής όπως στο Excel (προεπισκόπηση)
      h += "<tr" + (cls ? ' class="' + cls + '"' : "") + (hp ? ' style="height:' + Math.round(hp) + 'px"' : "") + ">";
      for (let c = R.s.c; c <= R.e.c; c++) {
        const k = r + ":" + c; if (cover.has(k)) continue;
        const sp = span[k], t = rowT[c - R.s.c];
        const cs = (ws[XLSX.utils.encode_cell({ r, c })] || {}).s || {}, st = [];
        if (cs.font && cs.font.bold) st.push("font-weight:700"); if (cs.font && cs.font.italic) st.push("font-style:italic"); if (cs.font && cs.font.sz) st.push("font-size:" + Math.round(cs.font.sz * 96 / 72) + "px");
        if (cs.fgColor && cs.fgColor.rgb && cs.patternType === "solid" && !/^(FFFFFF|000000)$/i.test(String(cs.fgColor.rgb).slice(-6))) st.push("background:#" + String(cs.fgColor.rgb).slice(-6));
        const al = cs.alignment || {}; if (al.horizontal) st.push("text-align:" + (al.horizontal === "centerContinuous" ? "center" : al.horizontal)); if (al.vertical) st.push("vertical-align:" + (al.vertical === "center" ? "middle" : al.vertical));
        h += "<td" + (sp ? (sp[0] > 1 ? ' rowspan="' + sp[0] + '"' : "") + (sp[1] > 1 ? ' colspan="' + sp[1] + '"' : "") : "") + (st.length ? ' style="' + st.join(";") + '"' : "") + ">" + E(t).replace(/\n/g, "<br>") + "</td>";
      }
      h += "</tr>";
    }
    return { html: h + "</table>", width: tw, range: XLSX.utils.encode_range(R), sheet: pa.sheet };
  }
  // λίστα ασθενών: στήλη ΟΝΟΜΑΤΕΠΩΝΥΜΟ «ΕΠΩΝΥΜΟ ΟΝΟΜΑ ΤΟΥ ΠΑΤΡΟΣ 12345678», από κάτω η επέμβαση (ενωμένα κελιά)
  function patients(XLSX, wb) {
    const pa = printArea(XLSX, wb), ws = wb.Sheets[pa.sheet], R = pa.range, out = [];
    let hdr = -1, col = {}; let room = "";
    for (let r = R.s.r; r <= R.e.r && hdr < 0; r++) for (let c = R.s.c; c <= R.e.c; c++) { const t = txt(ws[XLSX.utils.encode_cell({ r, c })]).toUpperCase().trim();
      if (/ΟΝΟΜΑΤΕΠΩΝΥΜΟ/.test(t)) hdr = r; if (/ΟΝΟΜΑΤΕΠΩΝΥΜΟ/.test(t)) col.name = c; else if (/ΕΠΕΜΒΑΣΗ/.test(t)) col.op = c; else if (/ΑΝΑΙΣΘ/.test(t)) col.an = c; else if (/^ΑΙΜΑ/.test(t)) col.blood = c; else if (/ARM/.test(t)) col.carm = c; else if (/ΧΕΙΡΟΥΡΓ/.test(t)) col.surg = c; else if (/^ΘΑΛ/.test(t)) col.ward = c; else if (/ΕΙΣ/.test(t)) col.adm = c; }
    if (hdr < 0 || col.name == null) return out;
    const at = (r, c) => (c == null ? "" : txt(ws[XLSX.utils.encode_cell({ r, c })]).replace(/\s+/g, " ").trim());
    for (let r = hdr + 1; r <= R.e.r; r++) {
      const first = at(r, R.s.c); if (/^ΑΙΘΟΥΣΑ/i.test(first)) { room = first.replace(/\s+/g, " ").trim(); continue; }
      const nm = at(r, col.name); if (!nm || /ΔΙΕΥΘΥΝΤΗΣ/.test(nm)) continue;
      const am = (nm.match(/(\d{8})\s*$/) || nm.match(/\b(\d{8})\b/) || [])[1] || "";
      const clean = nm.replace(/\d{6,}/g, "").replace(/\s+/g, " ").trim();
      const m = clean.match(/^(.*?)\s+ΤΟΥ\s+(.+)$/);
      out.push({ room, name: m ? m[1] : clean, father: m ? m[2] : "", am, dx: at(r, col.op), op: at(r + 1, col.op), anesth: at(r, col.an), blood: at(r, col.blood), carm: at(r, col.carm), surgeon: at(r, col.surg).replace(/^κ\.?\s*/i, ""), row: r + 1 });
    }
    return out;
  }
  // ημερομηνία προγράμματος: κελί δίπλα στο «ΠΡΟΓΡΑΜΜΑ» (π.χ. 9.11.26) ή από το όνομα αρχείου
  function dateOf(XLSX, wb, fname) {
    const pa = printArea(XLSX, wb), ws = wb.Sheets[pa.sheet], R = pa.range;
    for (let r = R.s.r; r <= Math.min(R.e.r, R.s.r + 6); r++) for (let c = R.s.c; c <= R.e.c; c++) { const t = txt(ws[XLSX.utils.encode_cell({ r, c })]); const m = t.match(/^\s*(\d{1,2})[.\/](\d{1,2})[.\/](\d{2,4})\s*$/); if (m) return (m[3].length === 2 ? "20" + m[3] : m[3]) + "-" + m[2].padStart(2, "0") + "-" + m[1].padStart(2, "0"); }
    const f = String(fname || "").match(/(\d{1,2})[.\-_](\d{1,2})[.\-_](\d{2,4})/); if (f) return (f[3].length === 2 ? "20" + f[3] : f[3]) + "-" + f[2].padStart(2, "0") + "-" + f[1].padStart(2, "0");
    return "";
  }
  const CSS = ".ors{border-collapse:collapse;table-layout:fixed;font:12px/1.25 Calibri,Arial,sans-serif;color:#000;background:#fff}.ors td{border:1px solid #444;padding:2px 4px;vertical-align:middle;overflow:hidden;word-wrap:break-word}.ors tr.hd td{font-weight:700;background:#E8EEF5;text-align:center}.ors tr.rm td{font-weight:700;background:#F3F3F3}";
  const api = { printArea, render, patients, dateOf, CSS };
  if (typeof module === "object" && module.exports) module.exports = api;
  root.pdxOrSched = api;
})(typeof self !== "undefined" ? self : this);
