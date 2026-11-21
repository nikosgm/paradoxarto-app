/* 📝 docedit.js — 21/11 (χρήστης: «τροποποιώ Word από τα Αρχεία κλινικής και πηγαίνει άμεσα στον φάκελο ουρολογική κλινική, στο δίκτυο» · «το 3» · «ΟΛΟΙ»)
   Επεξεργασία .docx ΜΕΣΑ στην εφαρμογή χωρίς να χαλάει η μορφή του Word:
   • αλλάζει ΜΟΝΟ το κείμενο των παραγράφων που άγγιξες: κοινό πρόθεμα/κατάληξη → η αλλαγή μπαίνει στο κομμάτι (run) όπου έγινε,
     οπότε έντονα/πλάγια/μέγεθος/χρώμα γύρω της μένουν ίδια
   • Enter = νέα παράγραφος με την ΙΔΙΑ μορφή/αρίθμηση · ⌫ στην αρχή = ένωση με την προηγούμενη · πίνακες: + / − γραμμή
   • παράγραφοι με εικόνες / πεδία / συνδέσμους / παρακολούθηση αλλαγών = μόνο ανάγνωση (αλλαγή τους μόνο στο Word)
   • κεφαλίδες, υποσέλιδα, στυλ, αρίθμηση, εικόνες, ρυθμίσεις σελίδας: ΔΕΝ αγγίζονται (το υπόλοιπο zip μένει όπως ήταν)
   Η αποθήκευση γίνεται από τον καλούντα (opts.onSave(bytes)) → bot «fssave» → φάκελος κλινικής, με αντίγραφο πριν. */
(function(){
'use strict';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XMLNS = 'http://www.w3.org/XML/1998/namespace';
const RUN_OK = new Set(['rPr','t','tab','br','cr','lastRenderedPageBreak','softHyphen','noBreakHyphen']);
const P_ZERO = new Set(['pPr','proofErr','bookmarkStart','bookmarkEnd','commentRangeStart','commentRangeEnd','permStart','permEnd']);
const isW = (n, ln)=> n && n.nodeType===1 && n.namespaceURI===W && (!ln || n.localName===ln);
const kids = (n, ln)=> Array.from(n.childNodes).filter(c=>isW(c, ln));

function runText(r){ let s='';
  for(const c of kids(r)){ const n=c.localName;
    if(n==='t') s+=c.textContent; else if(n==='tab') s+='\t'; else if(n==='br'||n==='cr') s+='\n'; else if(n==='softHyphen') s+='­'; else if(n==='noBreakHyphen') s+='‑'; }
  return s; }
// γιατί μια παράγραφος ΔΕΝ γίνεται να αλλάξει εδώ ('' = γίνεται)
function lockWhy(p){
  for(const c of Array.from(p.childNodes)){ if(c.nodeType!==1) continue;
    if(!isW(c)) return 'ειδικό περιεχόμενο';
    if(P_ZERO.has(c.localName)) continue;
    if(c.localName!=='r') return c.localName==='hyperlink' ? 'σύνδεσμος' : /^(ins|del|moveFrom|moveTo)$/.test(c.localName) ? 'παρακολούθηση αλλαγών' : /^(fldSimple|smartTag|sdt)$/.test(c.localName) ? 'πεδίο' : 'ειδικό περιεχόμενο';
    for(const k of Array.from(c.childNodes)){ if(k.nodeType!==1) continue;
      if(!isW(k) || !RUN_OK.has(k.localName)) return /drawing|pict|object/.test(k.localName) ? 'εικόνα' : /fld|instr/.test(k.localName) ? 'πεδίο' : 'ειδικό περιεχόμενο';
      if(k.localName==='br' && /^(page|column)$/.test(k.getAttributeNS(W,'type')||'')) return 'αλλαγή σελίδας'; } }
  return ''; }
const pText = p=> kids(p,'r').map(runText).join('');
function runFill(r, s){ const d=r.ownerDocument;
  for(const c of Array.from(r.childNodes)) if(!isW(c,'rPr')) r.removeChild(c);
  for(const x of s.split(/(\t|\n)/)){ if(!x) continue;
    if(x==='\t') r.appendChild(d.createElementNS(W,'w:tab')); else if(x==='\n') r.appendChild(d.createElementNS(W,'w:br'));
    else { const t=d.createElementNS(W,'w:t'); t.setAttributeNS(XMLNS,'xml:space','preserve'); t.textContent=x; r.appendChild(t); } } }
function newRun(p){ const d=p.ownerDocument, r=d.createElementNS(W,'w:r'); const ppr=kids(p,'pPr')[0], mk=ppr && kids(ppr,'rPr')[0];
  if(mk){ const rp=d.createElementNS(W,'w:rPr'); for(const c of Array.from(mk.childNodes)) if(isW(c) && !/^(ins|del|moveFrom|moveTo|rPrChange)$/.test(c.localName)) rp.appendChild(c.cloneNode(true)); if(rp.childNodes.length) r.appendChild(rp); }
  return r; }
// αλλαγή κειμένου παραγράφου ΔΙΑΤΗΡΩΝΤΑΣ τη μορφή: μόνο το τμήμα που άλλαξε, μέσα στο run όπου έγινε
function setText(p, nt){
  nt = String(nt).replace(/\r\n?/g,'\n');
  const runs = kids(p,'r'), old = runs.map(runText), O = old.join('');
  if(O===nt) return false;
  if(!runs.length){ const r=newRun(p); p.appendChild(r); runFill(r, nt); return true; }
  let a=0; while(a<O.length && a<nt.length && O[a]===nt[a]) a++;
  let s=0; while(s<O.length-a && s<nt.length-a && O[O.length-1-s]===nt[nt.length-1-s]) s++;
  const dEnd = O.length-s, ins = nt.slice(a, nt.length-s);
  // run εισαγωγής: αυτό που περιέχει τον χαρακτήρα ΠΡΙΝ από την αλλαγή (συνεχίζεις να γράφεις με τη μορφή του)· στην αρχή: το 1ο με κείμενο
  let insRun=-1, insAt=0, off=0;
  if(a===0){ insRun = Math.max(0, old.findIndex(t=>t.length>0)); insAt = 0; }
  else for(let i=0;i<runs.length;i++){ const L=old[i].length; if(a>off && a<=off+L){ insRun=i; insAt=a-off; break; } off+=L; }
  if(insRun<0){ insRun=runs.length-1; insAt=old[insRun].length; }
  off=0; const ns = old.map(t=>{ const st=off, en=off+t.length; off=en; const ds=Math.max(a,st), de=Math.min(dEnd,en); return de>ds ? t.slice(0,ds-st)+t.slice(de-st) : t; });
  ns[insRun] = ns[insRun].slice(0,insAt) + ins + ns[insRun].slice(insAt);
  runs.forEach((r,i)=>{ if(ns[i]===old[i]) return; if(!ns[i] && i!==insRun) r.parentNode.removeChild(r); else runFill(r, ns[i]); });
  return true; }
// κλώνος παραγράφου χωρίς σελιδοδείκτες/σχόλια/ταυτότητες (αλλιώς το Word «βρίσκει πρόβλημα» με διπλά ids)
function cleanClone(el){ const c = el.cloneNode(true);
  c.querySelectorAll('*').forEach(x=>{ if(isW(x) && /^(bookmarkStart|bookmarkEnd|commentRangeStart|commentRangeEnd|permStart|permEnd)$/.test(x.localName)) x.parentNode && x.parentNode.removeChild(x); });
  [c].concat(Array.from(c.querySelectorAll('*'))).forEach(x=>{ for(const at of Array.from(x.attributes||[])) if(/^(paraId|textId)$/.test(at.localName) || (at.localName==='rsidR'||at.localName==='rsidRDefault'||at.localName==='rsidP')) x.removeAttributeNode(at); });
  return c; }
function splitP(p, pos){ const t=pText(p), c=cleanClone(p); setText(p, t.slice(0,pos)); setText(c, t.slice(pos)); p.parentNode.insertBefore(c, p.nextSibling); return c; }
function mergeP(prev, cur){ const keep = kids(prev,'r'); const at = keep.length ? keep[keep.length-1].nextSibling : null;
  for(const c of Array.from(cur.childNodes)){ if(isW(c,'pPr')) continue; if(isW(c) && /^(bookmarkStart|bookmarkEnd)$/.test(c.localName)) continue; prev.insertBefore(c, at); }
  cur.parentNode.removeChild(cur); }
function delP(p){ const par=p.parentNode; if(isW(par,'tc') && kids(par,'p').length<=1){ setText(p,''); return false; } par.removeChild(p); return true; }
function addRow(tr){ const c=cleanClone(tr); kids(c,'tc').forEach(tc=>{ const ps=kids(tc,'p'); ps.slice(1).forEach(x=>tc.removeChild(x)); if(ps[0]) setText(ps[0],''); kids(tc,'tbl').forEach(x=>tc.removeChild(x)); }); tr.parentNode.insertBefore(c, tr.nextSibling); return c; }
function delRow(tr){ const tb=tr.parentNode; if(kids(tb,'tr').length<=1) return false; tb.removeChild(tr); return true; }

// μορφή για την προβολή (μόνο ενδεικτικά: έντονα/πλάγια/μέγεθος/στοίχιση/επικεφαλίδα/αρίθμηση)
function look(p){ const ppr=kids(p,'pPr')[0]; const o={};
  if(ppr){ const ps=kids(ppr,'pStyle')[0], v=ps ? (ps.getAttributeNS(W,'val')||'') : ''; if(/heading|title|επικεφ|τίτλ/i.test(v)) o.h=1;
    const jc=kids(ppr,'jc')[0]; if(jc) o.jc=jc.getAttributeNS(W,'val'); if(kids(ppr,'numPr').length) o.num=1; }
  const r=kids(p,'r').find(x=>runText(x).trim()); const rp=r && kids(r,'rPr')[0];
  if(rp){ const on=n=>{ const e=kids(rp,n)[0]; return e && !/^(0|false|off)$/.test(e.getAttributeNS(W,'val')||''); }; if(on('b')) o.b=1; if(on('i')) o.i=1; if(on('u')) o.u=1;
    const sz=kids(rp,'sz')[0]; if(sz){ const v=+sz.getAttributeNS(W,'val'); if(v) o.sz=v/2; } }
  return o; }

async function loadZip(){ if(window.JSZip) return window.JSZip;
  await new Promise((ok,no)=>{ const s=document.createElement('script'); s.src='https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js'; s.onload=ok; s.onerror=()=>no(new Error('δεν φορτώθηκε το JSZip (δίκτυο)')); document.head.appendChild(s); });
  return window.JSZip; }

// ===== μοντέλο εγγράφου (χωρίς UI — δοκιμάζεται μόνο του) =====
async function parse(buf){
  const JSZip = await loadZip(); const zip = await JSZip.loadAsync(buf);
  const f = zip.file('word/document.xml'); if(!f) throw new Error('δεν είναι έγγραφο Word (λείπει word/document.xml)');
  const src = await f.async('string'); const decl = (src.match(/^\s*<\?xml[^>]*\?>\s*/)||[''])[0];
  const doc = new DOMParser().parseFromString(src, 'application/xml');
  if(doc.getElementsByTagName('parsererror').length) throw new Error('το έγγραφο δεν διαβάζεται');
  const body = doc.getElementsByTagNameNS(W,'body')[0]; if(!body) throw new Error('κενό έγγραφο');
  const M = {zip, doc, body, decl, dirty:false, rev:0,   // rev: αλλαγές που έγιναν ΕΝΩ αποθηκευόταν μένουν «μη αποθηκευμένες»
    blocks(){ const out=[]; for(const c of kids(body)){ if(c.localName==='p') out.push({k:'p', el:c}); else if(c.localName==='tbl') out.push({k:'tbl', el:c}); else if(c.localName!=='sectPr') out.push({k:'x', el:c}); } return out; },
    text: pText, lockWhy, look, setText(p,t){ if(setText(p,t)){ M.dirty=true; M.rev++; } }, split(p,pos){ M.dirty=true; M.rev++; return splitP(p,pos); }, merge(a,b){ M.dirty=true; M.rev++; mergeP(a,b); },
    del(p){ M.dirty=true; M.rev++; return delP(p); }, addRow(tr){ M.dirty=true; M.rev++; return addRow(tr); }, delRow(tr){ const r=delRow(tr); if(r){ M.dirty=true; M.rev++; } return r; },
    rows: tbl=>kids(tbl,'tr'), cells: tr=>kids(tr,'tc'), paras: tc=>kids(tc,'p'),
    async save(){ let xml = new XMLSerializer().serializeToString(doc); if(!/^<\?xml/.test(xml)) xml = (decl || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n') + xml;
      zip.file('word/document.xml', xml, {createFolders:false});
      const out = await zip.generateAsync({type:'uint8array', compression:'DEFLATE', compressionOptions:{level:6}, mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
      // έλεγχος: ξαναδιαβάζεται και έχει το ίδιο κείμενο
      const z2 = await JSZip.loadAsync(out), x2 = await z2.file('word/document.xml').async('string'), d2 = new DOMParser().parseFromString(x2,'application/xml');
      if(d2.getElementsByTagName('parsererror').length) throw new Error('το νέο έγγραφο δεν είναι έγκυρο — δεν αποθηκεύτηκε');
      return out; } };
  return M; }

// ===== οθόνη επεξεργασίας =====
const CSS = '#pdxDE{position:fixed;inset:0;z-index:2147483000;background:#F2F2F7;display:flex;flex-direction:column;font:17px/1.4 -apple-system,system-ui,sans-serif;color:#000}'
 +'#pdxDE .de-h{display:flex;align-items:center;gap:8px;padding:calc(env(safe-area-inset-top) + 8px) 12px 8px;background:rgba(249,249,249,.94);border-bottom:.5px solid #C6C6C8}'
 +'#pdxDE .de-h b{flex:1;min-width:0;text-align:center;font-size:16px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}'
 +'#pdxDE .de-h button{border:0;background:none;color:#007AFF;font-size:17px;min-height:44px;padding:0 6px}#pdxDE .de-h button.sv{font-weight:700}#pdxDE .de-h button:disabled{color:#C7C7CC}'
 +'#pdxDE .de-st{padding:6px 14px;font-size:13px;color:#6C6C70;background:#F2F2F7}#pdxDE .de-st.er{color:#D70015;font-weight:600}#pdxDE .de-st.ok{color:#248A3D;font-weight:600}'
 +'#pdxDE .de-q{display:flex;gap:6px;padding:6px 12px}#pdxDE .de-q input{flex:1;min-width:0;border:0;border-radius:10px;background:rgba(118,118,128,.12);padding:8px 10px;font-size:16px;color:inherit;outline:0}'
 +'#pdxDE .de-q button{border:0;border-radius:10px;background:rgba(118,118,128,.12);color:#007AFF;font-size:15px;font-weight:600;padding:0 10px;min-height:36px}'
 +'#pdxDE .de-b{flex:1;overflow:auto;-webkit-overflow-scrolling:touch;padding:8px 10px calc(env(safe-area-inset-bottom) + 40vh)}'
 +'#pdxDE .de-pg{background:#fff;border-radius:12px;padding:10px 12px;box-shadow:0 1px 2px rgba(0,0,0,.06)}'
 +'#pdxDE textarea{display:block;width:100%;box-sizing:border-box;border:0;resize:none;overflow:hidden;background:transparent;color:inherit;font:inherit;padding:3px 4px;margin:0;border-radius:6px;outline:0;min-height:1.5em}'
 +'#pdxDE textarea:focus{background:rgba(0,122,255,.07)}#pdxDE textarea.hit{background:rgba(255,204,0,.35)}'
 +'#pdxDE .lk{padding:3px 4px;color:#8E8E93;white-space:pre-wrap}#pdxDE .lk i{font-style:normal;font-size:11px;background:rgba(118,118,128,.14);border-radius:5px;padding:1px 5px;margin-right:4px}'
 +'#pdxDE .num{display:flex;gap:4px}#pdxDE .num>span{flex:none;color:#8E8E93;padding:3px 0 0}'
 +'#pdxDE .tb{margin:6px 0;border:1px solid #D1D1D6;border-radius:8px;overflow:hidden}#pdxDE .tr{display:flex;border-top:1px solid #E5E5EA}#pdxDE .tr:first-child{border-top:0}'
 +'#pdxDE .tc{flex:1;min-width:0;border-left:1px solid #E5E5EA;padding:2px}#pdxDE .tc:first-of-type{border-left:0}#pdxDE .tr>button{flex:none;width:34px;border:0;background:none;color:#FF3B30;font-size:20px}'
 +'#pdxDE .tadd{display:block;width:100%;border:0;border-top:1px solid #E5E5EA;background:rgba(0,122,255,.06);color:#007AFF;font-weight:600;font-size:14px;padding:8px}'
 +'#pdxDE .more{display:block;width:100%;margin:8px 0;border:0;border-radius:10px;background:rgba(118,118,128,.12);color:#007AFF;font-size:15px;font-weight:600;padding:10px}'
 +'html.pdx-night #pdxDE{background:#000;color:#fff}html.pdx-night #pdxDE .de-h{background:rgba(28,28,30,.94);border-color:#38383A}html.pdx-night #pdxDE .de-st{background:#000;color:#8E8E93}'
 +'html.pdx-night #pdxDE .de-pg{background:#1C1C1E}html.pdx-night #pdxDE .tb,html.pdx-night #pdxDE .tr,html.pdx-night #pdxDE .tc{border-color:#38383A}html.pdx-night #pdxDE .de-h button,html.pdx-night #pdxDE .de-q button,html.pdx-night #pdxDE .tadd,html.pdx-night #pdxDE .more{color:#0A84FF}'
 +'html.pdx-night #pdxDE .de-st.ok{color:#30D158}html.pdx-night #pdxDE .de-st.er{color:#FF6961}';
const WIN = 120;
const fold = s=>String(s||'').normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase();
const esc = s=>String(s).replace(/[&<>"]/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

// opts: {name, bytes(Uint8Array|ArrayBuffer), onSave:async(bytes)=>({ok, msg}) — throw Error(.locked/.conflict)}
async function open(opts){
  const M = await parse(opts.bytes);
  document.getElementById('pdxDE')?.remove();
  if(!document.getElementById('pdxDECss')){ const c=document.createElement('style'); c.id='pdxDECss'; c.textContent=CSS; document.head.appendChild(c); }
  const ov = document.createElement('div'); ov.id='pdxDE';
  ov.innerHTML = '<div class="de-h"><button data-x>Κλείσιμο</button><b></b><button class="sv" data-sv disabled>Αποθήκευση</button></div>'
    +'<div class="de-st"></div><form class="de-q"><input type="search" enterkeyhint="search" placeholder="🔍 Εύρεση στο έγγραφο"><button type="submit">Εύρεση</button><button type="button" data-end>Τέλος ⤓</button></form><div class="de-b"></div>';
  document.body.appendChild(ov);
  ov.querySelector('.de-h b').textContent = opts.name||'Έγγραφο';
  const body = ov.querySelector('.de-b'), stEl = ov.querySelector('.de-st'), svB = ov.querySelector('[data-sv]');
  let w0 = 0, saving = false, hitEl = null; const els = new Map(); let seq = 0;   // textarea id → στοιχείο XML
  const status = (t, cls)=>{ stEl.textContent = t; stEl.className = 'de-st'+(cls?' '+cls:''); };
  const dirty = ()=>{ svB.disabled = saving || !M.dirty; if(M.dirty && !saving) status('Αλλαγές που δεν έχουν αποθηκευτεί — «Αποθήκευση» για τον φάκελο της κλινικής'); };
  status('Πάτα πάνω στο κείμενο για αλλαγή · Enter = νέα παράγραφος · μορφή, εικόνες & κεφαλίδες μένουν όπως είναι');
  const fit = ta=>{ ta.style.height='auto'; ta.style.height = ta.scrollHeight+'px'; };
  const taOf = (p, cls)=>{ const why = lockWhy(p), t = pText(p), L = look(p);
    const sty = (L.b||L.h?'font-weight:700;':'')+(L.i?'font-style:italic;':'')+(L.u?'text-decoration:underline;':'')+(L.h?'font-size:1.12em;':L.sz&&L.sz>=14?'font-size:1.06em;':'')+(L.jc==='center'?'text-align:center;':L.jc==='right'?'text-align:right;':'');
    if(why) return '<div class="lk" style="'+sty+'"><i>🔒 '+esc(why)+'</i>'+esc(t)+'</div>';
    const id = 'p'+(++seq); els.set(id, p);
    const ta = '<textarea rows="1" data-id="'+id+'"'+(cls?' class="'+cls+'"':'')+' style="'+sty+'" autocapitalize="sentences" spellcheck="false">'+esc(t)+'</textarea>';
    return L.num ? '<div class="num"><span>•</span>'+ta+'</div>' : ta; };
  function draw(focusEl, caret){
    els.clear(); const B = M.blocks(); if(w0 > Math.max(0, B.length-1)) w0 = Math.max(0, B.length-WIN);
    let h = '<div class="de-pg">';
    if(w0>0) h += '<button class="more" data-up>▲ Προηγούμενα ('+w0+')</button>';
    B.slice(w0, w0+WIN).forEach(b=>{
      if(b.k==='p') h += taOf(b.el);
      else if(b.k==='tbl'){ const rows = M.rows(b.el); h += '<div class="tb">';
        rows.forEach((tr,ri)=>{ const rid='r'+(++seq); els.set(rid, tr); h += '<div class="tr">'+M.cells(tr).map(tc=>'<div class="tc">'+M.paras(tc).map(p=>taOf(p)).join('')+(Array.from(tc.childNodes).some(c=>isW(c,'tbl'))?'<div class="lk"><i>🔒 πίνακας μέσα σε πίνακα</i></div>':'')+'</div>').join('')+(rows.length>1?'<button type="button" data-delrow="'+rid+'" aria-label="Διαγραφή γραμμής">−</button>':'')+'</div>'; });
        const last = rows[rows.length-1]; const lid='r'+(++seq); els.set(lid, last); h += '<button type="button" class="tadd" data-addrow="'+lid+'">+ Γραμμή στον πίνακα</button></div>'; }
      else h += '<div class="lk"><i>🔒 ειδικό περιεχόμενο</i>'+esc(b.el.textContent||'')+'</div>';
    });
    if(w0+WIN < B.length) h += '<button class="more" data-dn>▼ Επόμενα ('+(B.length-w0-WIN)+')</button>';
    body.innerHTML = h + '</div>';
    body.querySelectorAll('textarea').forEach(fit);
    if(focusEl){ const id=[...els].find(([k,v])=>v===focusEl && k[0]==='p'); const ta = id && body.querySelector('textarea[data-id="'+id[0]+'"]');
      if(ta){ ta.focus({preventScroll:true}); try{ ta.setSelectionRange(caret||0, caret||0); }catch(e){} ta.scrollIntoView({block:'center'}); } }
    dirty(); }
  const allPs = ()=>{ const out=[]; for(const b of M.blocks()){ if(b.k==='p') out.push(b.el); else if(b.k==='tbl') M.rows(b.el).forEach(tr=>M.cells(tr).forEach(tc=>M.paras(tc).forEach(p=>out.push(p)))); } return out; };
  const blockIdx = el=>{ const B=M.blocks(); for(let i=0;i<B.length;i++){ if(B[i].el===el || B[i].el.contains(el)) return i; } return 0; };
  body.addEventListener('input', e=>{ const ta=e.target; if(ta.tagName!=='TEXTAREA') return; const p=els.get(ta.dataset.id); if(!p) return; M.setText(p, ta.value); fit(ta); dirty(); });
  const enter = (ta)=>{ const p=els.get(ta.dataset.id); if(!p) return; M.setText(p, ta.value); const c = M.split(p, ta.selectionStart); draw(c, 0); };
  const back = (ta)=>{ const p=els.get(ta.dataset.id); if(!p) return false; if(ta.selectionStart!==0 || ta.selectionEnd!==0) return false;
    const prev = p.previousSibling && isW(p.previousSibling,'p') ? p.previousSibling : null; if(!prev || lockWhy(prev)) return false;
    M.setText(p, ta.value); const at = pText(prev).length; M.merge(prev, p); draw(prev, at); return true; };
  body.addEventListener('keydown', e=>{ const ta=e.target; if(ta.tagName!=='TEXTAREA' || e.isComposing) return;
    if(e.key==='Enter' && !e.shiftKey){ e.preventDefault(); enter(ta); }
    else if(e.key==='Backspace' && back(ta)) e.preventDefault(); });
  body.addEventListener('beforeinput', e=>{ const ta=e.target; if(ta.tagName!=='TEXTAREA') return;
    if(e.inputType==='insertParagraph' || e.inputType==='insertLineBreak'){ e.preventDefault(); enter(ta); }
    else if(e.inputType==='deleteContentBackward' && back(ta)) e.preventDefault(); });
  body.addEventListener('click', e=>{ let b;
    if(e.target.closest('[data-up]')){ w0 = Math.max(0, w0-WIN); draw(); body.scrollTop = 0; return; }
    if(e.target.closest('[data-dn]')){ w0 += WIN; draw(); body.scrollTop = 0; return; }
    if((b=e.target.closest('[data-addrow]'))){ const tr=els.get(b.dataset.addrow); if(tr){ const c=M.addRow(tr); const p=M.paras(M.cells(c)[0]||c)[0]; draw(p, 0); } return; }
    if((b=e.target.closest('[data-delrow]'))){ const tr=els.get(b.dataset.delrow); if(tr && confirm('Διαγραφή αυτής της γραμμής του πίνακα;')){ M.delRow(tr); draw(); } return; } });
  ov.querySelector('.de-q').addEventListener('submit', e=>{ e.preventDefault(); const q = fold(ov.querySelector('.de-q input').value.trim()); if(!q) return;
    const P = allPs(); const start = hitEl ? P.indexOf(hitEl)+1 : 0; let k = -1;
    for(let i=0;i<P.length;i++){ const j=(start+i)%P.length; if(fold(pText(P[j])).includes(q)){ k=j; break; } }
    if(k<0){ status('Δεν βρέθηκε «'+ov.querySelector('.de-q input').value.trim()+'»', 'er'); return; }
    hitEl = P[k]; const bi = blockIdx(hitEl); if(bi < w0 || bi >= w0+WIN) w0 = Math.max(0, bi-10); draw();
    const id=[...els].find(([kk,v])=>v===hitEl && kk[0]==='p'); const ta = id && body.querySelector('textarea[data-id="'+id[0]+'"]');
    if(ta){ ta.classList.add('hit'); ta.scrollIntoView({block:'center'}); } status('Βρέθηκε — πάτα «Εύρεση» ξανά για το επόμενο'); });
  ov.querySelector('[data-end]').addEventListener('click', ()=>{ const B=M.blocks(); w0 = Math.max(0, B.length-WIN); draw(); body.scrollTop = body.scrollHeight; });
  ov.querySelector('[data-x]').addEventListener('click', ()=>{ if(M.dirty && !confirm('Υπάρχουν αλλαγές που δεν αποθηκεύτηκαν. Κλείσιμο χωρίς αποθήκευση;')) return; ov.remove(); opts.onClose && opts.onClose(); });
  async function doSave(){ if(saving || !M.dirty) return; const ae = document.activeElement; if(ae && ae.tagName==='TEXTAREA'){ const p=els.get(ae.dataset.id); if(p) M.setText(p, ae.value); }
    saving = true; dirty(); status('Αποθήκευση στον φάκελο της κλινικής…');
    const rev0 = M.rev;
    try{ const bytes = await M.save(); const r = await opts.onSave(bytes); M.dirty = M.rev !== rev0; saving = false; dirty(); if(M.dirty) status((r && r.msg || '✓ Αποθηκεύτηκε')+' · νεότερες αλλαγές σου δεν έχουν αποθηκευτεί ακόμα'); else status(r && r.msg || '✓ Αποθηκεύτηκε στον φάκελο της κλινικής', 'ok'); }
    catch(e){ saving = false; dirty(); status('⚠ '+(e.message||e)+' — οι αλλαγές σου είναι ακόμα εδώ', 'er'); if(e.conflict && opts.onConflict) opts.onConflict(await M.save().catch(()=>null)); } }
  svB.addEventListener('click', doSave);
  if(opts.atEnd){ const B=M.blocks(); w0 = Math.max(0, B.length-WIN); }
  draw(); if(opts.atEnd) body.scrollTop = body.scrollHeight;
  return {M, el:ov, save:doSave, close(){ ov.remove(); }};
}
window.pdxDocEdit = {parse, open, _setText:setText, _lockWhy:lockWhy};
})();
