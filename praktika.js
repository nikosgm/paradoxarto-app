// 📝 8/11 (εντολή χρήστη): ΠΡΑΚΤΙΚΟ ΧΕΙΡΟΥΡΓΕΙΟΥ από τα πρακτικά της κλινικής (window.PRAKTIKA).
// pdxOpNote({name, proc, side, surg, text, onSave}) → προτεινόμενα πρακτικά για την επέμβαση (πιο ταιριαστά πρώτα) →
// επιλογή → επεξεργασία (ΑΡ⇄ΔΕ, χειρουργοί) → Αντιγραφή / Αποθήκευση στο πρόγραμμα.
(function () {
  "use strict";
  const E = (s) => String(s == null ? "" : s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const N = (s) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/ς/g, "σ");
  // συνώνυμα/συντομογραφίες του προγράμματος → λέξεις των πρακτικών
  const SYN = [
    [/turbt|turb\b|διουρηθρ.*(ογκ|κυστ)|εκτομη ογκ/, "turbt"], [/re-?turbt|re-?turb/, "re-turbt"], [/turp|διουρηθρ.*προστατ|τομη αυχενα|εκτομη αυχενα/, "turp"],
    [/pigtail|double ?j|\bjj\b|ουρητηρικ.*καθετηρ|stent/, "pigtail"], [/αλλαγη (pigtail|jj|stent)|αντικατασταση/, "αλλαγη pigtail"],
    [/urs|ουρητηροσκοπ/, "ουρητηροσκοπηση"], [/rirs|ευκαμπτ/, "ευκαμπτη"], [/laser|λιθοτριψ|holep/, "laser λιθοτριψια"], [/pcnl|διαδερμ/, "νεφρολιθοτριψια"],
    [/ριζικ.*νεφρεκτ|νεφρεκτομη/, "νεφρεκτομη"], [/μερικ.*νεφρεκτ|nss|nephron/, "μερικη νεφρεκτομη"], [/λαπαροσκ|lap\b/, "λαπαροσκοπικη"],
    [/rrp|ριζικ.*προστατεκτ/, "ριζικη προστατεκτομη"], [/κυστεκτομ/, "κυστεκτομη"], [/ορχεκτομ/, "ορχεκτομη"], [/ορχεοπηξ/, "ορχεοπηξια"],
    [/κιρσοκηλ|varicocel/, "κιρσοκηλεκτομη"], [/υδροκηλ/, "υδροκηλη"], [/περιτομ|circum/, "περιτομη"], [/optilume/, "optilume"], [/ουρηθροτομ|sachse/, "ουρηθροτομη"],
    [/ουρηθροπλαστ/, "ουρηθροπλαστικη"], [/βιοψ.*προστ|fusion/, "βιοψια προστατη"], [/κυστεοσκοπ|cysto/, "κυστεοσκοπηση"], [/αιματοκυστ|κενωση|εκκενωση/, "αιματοκυστη"],
    [/μετεμφυτευ|reimplant/, "μετεμφυτευση"], [/αποστημα/, "αποστημα"], [/κυστ.*νεφρου|αποστεγασ/, "αποστεγαση κυστεων"], [/ire\b|nanoknife/, "ire"],
  ];
  const keys = (s) => { const n = N(s); const out = new Set(n.split(/[^a-zα-ω0-9]+/).filter((w) => w.length >= 3)); for (const [re, k] of SYN) if (re.test(n)) N(k).split(" ").forEach((w) => out.add(w)); return out; };
  const sideOf = (s) => { const n = N(s); return /\bαμφω|αμφοτεροπλ/.test(n) ? "ΑΜΦΩ" : /\bαρ\b|αριστερ/.test(n) ? "ΑΡ" : /\bδε\b|δεξι/.test(n) ? "ΔΕ" : ""; };
  const IDX = () => (window.PRAKTIKA || []).map((x) => ({ ...x, k: keys(x.op + " " + (x.aka || []).join(" ") + " " + x.t.slice(0, 160)), side: sideOf(x.op) }));
  function rank(proc, side) {
    const q = keys(proc); const L = IDX().map((x) => { let sc = 0; q.forEach((w) => { if (x.k.has(w)) sc += 3; else if ([...x.k].some((k) => k.startsWith(w.slice(0, 5)))) sc += 1; }); if (side && x.side === side) sc += 1; if (/\[;\]|δυσανάγνωστο/.test(x.op)) sc -= 2; return { x, sc }; });
    return L.filter((o) => o.sc > 0).sort((a, b) => b.sc - a.sc || a.x.t.length - b.x.t.length).map((o) => o.x);
  }
  // ΑΡ ⇄ ΔΕ σε όλο το κείμενο (και «αριστερό/δεξιό»)
  const swapSide = (t, to) => { const from = to === "ΑΡ" ? "ΔΕ" : "ΑΡ"; const W = to === "ΑΡ" ? [["ΔΕ", "ΑΡ"], ["δεξι", "αριστερ"], ["Δεξι", "Αριστερ"]] : [["ΑΡ", "ΔΕ"], ["αριστερ", "δεξι"], ["Αριστερ", "Δεξι"]];
    let s = t; for (const [a, b] of W) s = s.replace(new RegExp("(^|[^Α-Ωα-ωΆ-ώA-Za-z])" + a + (a === a.toUpperCase() ? "(?![Α-Ωα-ωΆ-ώA-Za-z])" : ""), "g"), "$1" + b); return s; };
  window.pdxOpNoteRank = rank; window.pdxOpNoteSwap = swapSide;
  window.pdxOpNote = function (o) {
    o = o || {}; const side = o.side || sideOf(o.proc);
    let list = rank(o.proc || "", side), q = "", cur = o.text ? { t: o.text, op: "αποθηκευμένο", surg: o.surg || "" } : null, text = o.text || "", surg = o.surg || "";
    const ov = document.createElement("div"); ov.id = "opnOv";
    ov.style.cssText = "position:fixed;inset:0;z-index:9800;background:rgba(15,25,35,.45);display:flex;align-items:flex-end;justify-content:center";
    document.body.appendChild(ov);
    const draw = () => {
      const all = q ? IDX().filter((x) => N(x.op + " " + x.t).includes(N(q))) : list;
      ov.innerHTML = '<div style="background:var(--card,#fff);color:var(--t1,#1F2933);width:100%;max-width:760px;max-height:92vh;overflow:auto;border-radius:18px 18px 0 0;padding:14px 16px calc(16px + env(safe-area-inset-bottom,0px));font:15px -apple-system,system-ui">'
        + '<div style="display:flex;align-items:center;gap:8px"><b style="flex:1;font-size:17px">📝 Πρακτικό · ' + E(o.name || "") + '</b><button data-x style="border:0;background:#EEF2F6;border-radius:10px;padding:8px 12px">Κλείσιμο</button></div>'
        + '<div style="color:#6B7280;font-size:13px;margin:4px 0 8px">' + E(o.proc || "—") + (side ? " · " + side : "") + "</div>"
        + (cur ? (o.noSurg ? "" : '<div style="font-weight:700;margin-top:6px">Χειρουργοί</div><input data-sg value="' + E(surg) + '" style="width:100%;box-sizing:border-box;font-size:16px;padding:9px;border:1px solid #C7CDD4;border-radius:10px">')
          + '<div style="display:flex;gap:6px;margin:10px 0 4px;align-items:center"><b style="flex:1">Κείμενο</b><button data-sd="ΑΡ" style="border:1px solid #C7CDD4;background:#fff;border-radius:16px;padding:6px 12px">→ ΑΡ</button><button data-sd="ΔΕ" style="border:1px solid #C7CDD4;background:#fff;border-radius:16px;padding:6px 12px">→ ΔΕ</button><button data-back style="border:0;background:#EEF2F6;border-radius:16px;padding:6px 12px">Άλλο πρότυπο</button></div>'
          + '<textarea data-t style="width:100%;box-sizing:border-box;min-height:300px;font-size:15.5px;line-height:1.45;padding:10px;border:1px solid #C7CDD4;border-radius:10px">' + E(text) + "</textarea>"
          + '<div style="display:flex;gap:8px;margin-top:10px"><button data-cp style="flex:1;min-height:46px;border:0;border-radius:12px;background:#EEF2F6;color:#0A84FF;font-weight:700">📋 Αντιγραφή</button>' + (o.onSave ? '<button data-sv style="flex:1;min-height:46px;border:0;border-radius:12px;background:#0A84FF;color:#fff;font-weight:700">💾 Αποθήκευση</button>' : "") + "</div>"
          + '<div style="font-size:12px;color:#8E8E93;margin-top:8px">Από τα πρακτικά της κλινικής (ψηφιοποίηση — έλεγξε πριν από τη χρήση· τα «[;]» είναι αβέβαιες λέξεις).</div>'
        : '<input data-q type="search" placeholder="🔎 Αναζήτηση σε όλα τα πρακτικά (π.χ. pigtail, νεφρεκτομή)" value="' + E(q) + '" style="width:100%;box-sizing:border-box;font-size:16px;padding:9px;border:1px solid #C7CDD4;border-radius:10px">'
          + '<div style="font-weight:700;margin:12px 0 6px">' + (q ? "Αποτελέσματα · " + all.length : all.length ? "Προτεινόμενα για αυτή την επέμβαση · " + all.length : "Δεν βρέθηκε ταιριαστό — ψάξε παραπάνω") + "</div>"
          + all.slice(0, 25).map((x) => '<button data-pick="' + x.n + '" style="display:block;width:100%;text-align:left;border:1px solid #E5E7EB;background:#fff;border-radius:12px;padding:10px 12px;margin:0 0 8px"><b>' + E(x.op) + '</b>' + (o.noSurg ? "" : '<div style="font-size:12.5px;color:#6B7280;margin:2px 0 4px">' + E(x.surg) + '</div>') + '<div style="font-size:13.5px;color:#374151;max-height:3.9em;overflow:hidden">' + E(x.t) + "</div></button>").join(""))
        + "</div>";
    };
    draw();
    ov.addEventListener("input", (e) => { if (e.target.dataset.q != null) { q = e.target.value; const c = e.target.selectionStart; draw(); const i = ov.querySelector("[data-q]"); i.focus(); try { i.setSelectionRange(c, c); } catch (x) {} } if (e.target.dataset.t != null) text = e.target.value; if (e.target.dataset.sg != null) surg = e.target.value; });
    ov.addEventListener("click", (e) => {
      const t = e.target;
      if (t === ov || t.closest("[data-x]")) return ov.remove();
      const pk = t.closest("[data-pick]"); if (pk) { cur = (window.PRAKTIKA || []).find((x) => x.n === +pk.dataset.pick); const cs = sideOf(cur.op) || sideOf(cur.t); text = side && cs && cs !== "ΑΜΦΩ" && cs !== side && side !== "ΑΜΦΩ" ? swapSide(cur.t, side) : cur.t; /* 9/11: συγχωνευμένα → πλευρά από το κείμενο */ if (!surg) surg = cur.surg; return draw(); }
      if (t.closest("[data-back]")) { cur = null; return draw(); }
      const sd = t.closest("[data-sd]"); if (sd) { text = swapSide(text, sd.dataset.sd); return draw(); }
      if (t.closest("[data-cp]")) { const s = (surg && !o.noSurg ? surg + "\n\n" : "") + text; (navigator.clipboard ? navigator.clipboard.writeText(s) : Promise.reject()).then(() => { t.textContent = "✓ Αντιγράφηκε"; }, () => { const a = ov.querySelector("[data-t]"); a.select(); document.execCommand("copy"); t.textContent = "✓ Αντιγράφηκε"; }); return; }
      if (t.closest("[data-sv]") && o.onSave) { o.onSave({ text, surg }); ov.remove(); }
    });
    return ov;
  };
})();
