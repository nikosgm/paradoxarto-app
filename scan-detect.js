// 🗓 8/11: «Είναι πρόγραμμα χειρουργείου;» — χωρίς AI, από τα χαρακτηριστικά της σελίδας (στη συσκευή):
// (1) ΠΙΝΑΚΑΣ: πολλές μακριές οριζόντιες γραμμές (2 ανά ασθενή + κεφαλίδες) και αρκετές κάθετες στήλες (8 στήλες),
// (2) ΓΚΡΙΖΕΣ ΛΩΡΙΔΕΣ: κεφαλίδα «ΘΑΛ · ΗΜ.ΕΙΣ. · …» και γραμμές «ΑΙΘΟΥΣΑ N» με γκρι φόντο,
// (3) ΟΡΙΖΟΝΤΙΟ φύλλο (οι γραμμές του πίνακα κατά τη μεγάλη διάσταση· ελέγχονται και οι δύο προσανατολισμοί, αν η φωτογραφία είναι όρθια).
// Επιστρέφει {prog:bool, score, why}. Κοινό για scan.html (window.pdxIsProgram) και δοκιμές (module.exports).
(function (root) {
  "use strict";
  function gray(img, W, H) { const g = new Float32Array(W * H); for (let i = 0, j = 0; i < g.length; i++, j += 4) g[i] = 0.299 * img[j] + 0.587 * img[j + 1] + 0.114 * img[j + 2]; return g; }
  // τοπικό κατώφλι (μέσος όρος σε παράθυρο) με ολοκληρωτική εικόνα — αντέχει σκιές/φωτισμό φωτογραφίας
  function dark(g, W, H, r, c) {
    const I = new Float64Array((W + 1) * (H + 1));
    for (let y = 0; y < H; y++) { let s = 0; for (let x = 0; x < W; x++) { s += g[y * W + x]; I[(y + 1) * (W + 1) + x + 1] = I[y * (W + 1) + x + 1] + s; } }
    const D = new Uint8Array(W * H);
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      const x0 = Math.max(0, x - r), x1 = Math.min(W, x + r + 1), y0 = Math.max(0, y - r), y1 = Math.min(H, y + r + 1);
      const m = (I[y1 * (W + 1) + x1] - I[y0 * (W + 1) + x1] - I[y1 * (W + 1) + x0] + I[y0 * (W + 1) + x0]) / ((x1 - x0) * (y1 - y0));
      D[y * W + x] = g[y * W + x] < m - c ? 1 : 0;
    }
    return D;
  }
  // ΛΕΠΤΕΣ γραμμές (όχι γράμματα): σκοτεινό εικονοστοιχείο με ανοιχτό χρώμα ±4 κάθετα στη γραμμή · ανοχή κλίσης φωτογραφίας (έως ~5°)
  function thin(D, W, H, horiz) {
    const T = new Uint8Array(W * H), k = 4;
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) { if (!D[y * W + x]) continue;
      if (horiz) { const u = y - k >= 0 ? D[(y - k) * W + x] : 0, d = y + k < H ? D[(y + k) * W + x] : 0; if (!u && !d) T[y * W + x] = 1; }
      else { const l = x - k >= 0 ? D[y * W + x - k] : 0, r = x + k < W ? D[y * W + x + k] : 0; if (!l && !r) T[y * W + x] = 1; } }
    return T;
  }
  function lines(D, W, H, horiz, frac, gap) {
    const T = thin(D, W, H, horiz), N = horiz ? H : W, L = horiz ? W : H;
    let bestN = 0, bestPos = [];
    for (const t of [0, -0.035, 0.035, -0.07, 0.07]) {
      const hits = new Uint8Array(N);
      for (let a = 0; a < N; a++) {
        let best = 0, run = 0, miss = 0;
        for (let b = 0; b < L; b++) {
          const a2 = Math.round(a + (b - L / 2) * t); if (a2 < 0 || a2 >= N) { run = 0; miss = 0; continue; }
          const at = (aa) => aa >= 0 && aa < N && (horiz ? T[aa * W + b] : T[b * W + aa]);
          if (at(a2) || at(a2 - 1) || at(a2 + 1)) { run += 1 + miss; miss = 0; if (run > best) best = run; } else if (run && miss < gap) miss++; else { run = 0; miss = 0; }
        }
        hits[a] = best >= frac * L ? 1 : 0;
      }
      let n = 0; const pos = []; for (let a = 0; a < N; a++) if (hits[a] && !hits[a - 1] && !hits[a - 2]) { n++; pos.push(a / N); }
      if (n > bestN) { bestN = n; bestPos = pos; }
    }
    return { n: bestN, pos: bestPos };
  }
  // γκρίζες λωρίδες: σειρές όπου η πλειονότητα των εικονοστοιχείων είναι «μεσαίο γκρι» σε σχέση με το χαρτί (όχι λευκό, όχι μελάνι)
  function bands(g, W, H, horiz) {
    const N = horiz ? H : W, L = horiz ? W : H;
    const all = Array.from(g).sort((a, b) => a - b), paper = all[Math.floor(all.length * 0.9)] || 255;
    const rows = [];
    for (let a = 0; a < N; a++) {
      let mid = 0;
      for (let b = Math.floor(L * 0.1); b < L * 0.9; b++) { const v = (horiz ? g[a * W + b] : g[b * W + a]) / paper; if (v > 0.55 && v < 0.88) mid++; }
      rows.push(mid / (L * 0.8) > 0.55);
    }
    let n = 0, run = 0; for (let a = 0; a <= N; a++) { if (rows[a]) run++; else { if (run >= Math.max(3, N * 0.012)) n++; run = 0; } }
    return n;
  }
  function analyze(img, W, H) {
    const g = gray(img, W, H), D = dark(g, W, H, Math.max(6, Math.round(Math.max(W, H) / 60)), 12);
    const res = [];
    for (const horiz of [true, false]) {   // horiz: οι γραμμές του πίνακα κατά πλάτος της εικόνας
      const hl = lines(D, W, H, horiz, 0.5, 4), vl = lines(D, W, H, !horiz, 0.07, 2), gb = bands(g, W, H, horiz);
      const along = horiz ? W : H, across = horiz ? H : W;   // φύλλο οριζόντιο ως προς τις γραμμές του πίνακα
      const land = along / across >= 1.15;
      let score = 0; const why = [];
      if (hl.n >= 8) { score += 2; why.push(hl.n + " οριζόντιες"); } else if (hl.n >= 5) { score += 1; why.push(hl.n + " οριζόντιες"); }
      if (vl.n >= 6) { score += 2; why.push(vl.n + " κάθετες"); } else if (vl.n >= 4) { score += 1; why.push(vl.n + " κάθετες"); }
      if (gb >= 2) { score += 2; why.push(gb + " γκρι λωρίδες"); } else if (gb === 1) { score += 1; why.push("1 γκρι λωρίδα"); }
      if (land) { score += 1; why.push("οριζόντιο"); }
      res.push({ score, why: why.join(" · "), hl: hl.n, vl: vl.n, gb, land });
    }
    const best = res[0].score >= res[1].score ? res[0] : res[1];
    return { prog: best.score >= 6 && best.hl >= 5 && best.vl >= 4 && best.gb >= 1, score: best.score, why: best.why, best };
  }
  // από <canvas>/<img>: σμίκρυνση ώστε η μεγάλη πλευρά = 640
  function fromCanvas(src) {
    const w0 = src.naturalWidth || src.videoWidth || src.width, h0 = src.naturalHeight || src.videoHeight || src.height, k = 640 / Math.max(w0, h0);
    const W = Math.max(1, Math.round(w0 * k)), H = Math.max(1, Math.round(h0 * k));
    const c = document.createElement("canvas"); c.width = W; c.height = H; const x = c.getContext("2d", { willReadFrequently: true }); x.drawImage(src, 0, 0, W, H);
    return analyze(x.getImageData(0, 0, W, H).data, W, H);
  }
  const api = { analyze, fromCanvas };
  root.pdxIsProgram = api;
  if (typeof module !== "undefined") module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
