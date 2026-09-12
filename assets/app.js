import { ROWS } from './rows.js';
import { parseNum, r3, fmt, fmtLei } from './num.js';
import { loadTemplate, buildXlsx, readMacheta, readBC39 } from './xlsx.js';
import * as store from './store.js';

const KEYS = ['e','f','g','h','i','j'];
const lunaDinText = t => { const i = LUNI.findIndex(l=>new RegExp(l,'i').test(String(t||''))); return i<0?null:i+1; };
const anDinText = t => { const m = String(t||'').match(/(20\d\d)/); return m?+m[1]:null; };
const LUNI = ['ianuarie','februarie','martie','aprilie','mai','iunie','iulie','august','septembrie','octombrie','noiembrie','decembrie'];
const byRow = Object.fromEntries(ROWS.map(r=>[r.r, r]));
const LEAVES = ROWS.filter(r=>!r.ch).map(r=>r.r);
const $ = s => document.querySelector(s);
const blank = () => Object.fromEntries(KEYS.map(k=>[k,0]));

let TPL = null;            // {buf, check}
let state = newState();
let prev = null;           // {label, leaves}
let bc39ref = null;        // {bugetar, angajament, sursa, coloane}

function newState(){
  const d = new Date();
  const leaves = {}; const din = {};
  LEAVES.forEach(r=>{ leaves[r]=blank(); din[r]={i:0,dj:0,f:0}; });
  return {
    meta:{ an:d.getFullYear(), luna:d.getMonth()+1, unitate:'SC SPINAL CARE DOBRECI SRL' },
    mode:'cumulat', linkEG:true, leaves, din, prevLabel:null
  };
}

/* ---------- calcul ---------- */
function recompute(){
  if(state.mode==='delta'){
    LEAVES.forEach(r=>{
      const p = prev && prev.leaves[r] ? prev.leaves[r] : blank();
      const d = state.din[r] || {i:0,dj:0,f:0};
      const L = state.leaves[r];
      L.i = r3(d.i);
      L.h = r3((p.h||0) + d.i);
      L.j = r3((p.j||0) + d.dj);
      L.f = r3(d.f);
      if(state.linkEG){ L.e = L.h; L.g = L.h; }
    });
  } else if(state.linkEG){
    LEAVES.forEach(r=>{ const L = state.leaves[r]; L.e = L.h; L.g = L.h; });
  }
  return aggregate(state.leaves);
}

/** Agregare bottom-up: fiecare rand de sinteza = suma copiilor sai. */
export function aggregate(leaves){
  const vals = {};
  LEAVES.forEach(r=>{ vals[r] = {...blank(), ...(leaves[r]||{})}; });
  [...ROWS].sort((a,b)=>b.r-a.r).forEach(row=>{
    if(!row.ch) return;
    const acc = blank();
    row.ch.forEach(c=>{ const v = vals[c]; if(v) KEYS.forEach(k=>{ acc[k] += (v[k]||0); }); });
    KEYS.forEach(k=>{ acc[k]=r3(acc[k]); });
    vals[row.r]=acc;
  });
  return vals;
}

/* ---------- tabel ---------- */
function buildTable(){
  const H = (t,n,cls='') => `<th class="num ${cls}">${t}<small>${n}</small></th>`;
  const head = `<thead>
    <tr>
      <th class="code">Sub&shy;cap.</th><th class="code">Titlu art.</th><th class="code">Alin.</th>
      <th class="lbl">Denumirea indicatorilor</th>
      ${H('Credite bugetare aprobate',3)}${H('Credite de angajament',4)}
      ${H('Credite trimestriale cumulate',5)}${H('Plăți nete de casă cumulat',6)}
      ${H('Luna curentă',7)}${H('Cheltuieli efective',8)}
      ${H('Δ efective luna','Δ','dcol')}
    </tr>
  </thead>`;
  const body = ROWS.map(row=>{
    const isSum = !!row.ch;
    const pad = 4 + row.lvl*11;
    const cells = KEYS.map((k,idx)=>{
      if(isSum) return `<td class="num calc" data-calc="${row.r}.${k}">0,000</td>`;
      const locked = state.linkEG && (k==='e'||k==='g');
      const derived = state.mode==='delta' && (k==='h'||k==='j');
      if(locked||derived) return `<td class="num calc" data-calc="${row.r}.${k}">0,000</td>`;
      const src = (state.mode==='delta' && k==='f') ? 'din.f' : (state.mode==='delta' && k==='i') ? 'din.i' : 'leaf.'+k;
      return `<td class="num"><input type="text" inputmode="decimal" data-r="${row.r}" data-k="${k}" data-src="${src}" data-col="${idx}" value="0,000"></td>`;
    }).join('');
    const dcell = isSum
      ? `<td class="num calc dcol">—</td>`
      : `<td class="num dcol"><input type="text" inputmode="decimal" data-r="${row.r}" data-k="dj" data-src="din.dj" data-col="6" value="0,000"></td>`;
    // in macheta, unele randuri au textul indicatorului in coloana de cod (ex. r.110): il mutam in coloana de denumire
    const codes = [row.sub, row.tit, row.alin].map(c=>String(c||''));
    const lung = codes.filter(c=>c.length>8);
    const lbl = [row.lbl, ...lung].filter(Boolean).join(' ');
    return `<tr data-row="${row.r}" class="${isSum?'sum l'+Math.min(row.lvl,3):''}">
      ${codes.map(c=>`<td class="code">${c.length>8?'':c}</td>`).join('')}
      <td class="lbl" style="padding-left:${pad}px"><span class="rn">${row.r}</span> ${lbl}</td>
      ${cells}${dcell}</tr>`;
  }).join('');
  $('#tbl').innerHTML = head + '<tbody>' + body + '</tbody>';
  requestAnimationFrame(()=>document.documentElement.style.setProperty('--thead-h', $('#tbl thead').offsetHeight+'px'));
  $('#tbl').querySelectorAll('input').forEach(inp=>{
    inp.addEventListener('focus', e=>e.target.select());
    inp.addEventListener('input', onInput);
    inp.addEventListener('blur', e=>{ e.target.value = fmt(readInput(e.target)); });
    inp.addEventListener('keydown', onKey);
  });
}

function readInput(inp){
  const r = +inp.dataset.r, k = inp.dataset.k, src = inp.dataset.src;
  return src.startsWith('din') ? (state.din[r][k]||0) : (state.leaves[r][k]||0);
}
function onInput(e){
  const inp = e.target, r = +inp.dataset.r, k = inp.dataset.k;
  const v = r3(parseNum(inp.value));
  if(inp.dataset.src.startsWith('din')) state.din[r][k] = v; else state.leaves[r][k] = v;
  render();
  scheduleSave();
}
function onKey(e){
  const move = e.key==='ArrowDown'||e.key==='Enter' ? 1 : e.key==='ArrowUp' ? -1 : 0;
  if(!move) return;
  e.preventDefault();
  const col = e.target.dataset.col;
  const all = [...$('#tbl').querySelectorAll(`input[data-col="${col}"]`)].filter(i=>i.offsetParent!==null);
  const i = all.indexOf(e.target);
  const n = all[i+move]; if(n){ n.focus(); }
}

/* ---------- randare ---------- */
function render(){
  const vals = recompute();
  document.querySelectorAll('#tbl td[data-calc]').forEach(td=>{
    const [r,k] = td.dataset.calc.split('.');
    td.textContent = fmt(vals[+r] ? vals[+r][k] : 0);
  });
  document.querySelectorAll('#tbl input').forEach(inp=>{
    const v = readInput(inp);
    if(document.activeElement!==inp) inp.value = fmt(v);
    inp.classList.toggle('z', Math.abs(v)<1e-9);
  });
  const q = ($('#filtru').value||'').trim().toLowerCase();
  const hz = $('#hideZero').checked;
  document.querySelectorAll('#tbl tbody tr').forEach(tr=>{
    const rn = +tr.dataset.row, row = byRow[rn], v = vals[rn] || blank();
    const gol = KEYS.every(k=>Math.abs(v[k])<1e-9);
    tr.classList.toggle('zero', gol);
    // randurile de sinteza raman vizibile ca reper, cu exceptia celor complet goale cand se ascund zerourile
    const potrivit = !q || row.lbl.toLowerCase().includes(q) || String(rn)===q;
    let vizibil = row.ch ? (!q || potrivit || true) : potrivit;
    if(hz && gol && document.activeElement?.dataset?.r !== String(rn)) vizibil = false;
    tr.hidden = !vizibil;
  });
  document.body.classList.toggle('mode-delta', state.mode==='delta');
  $('#t8e').textContent = fmt(vals[8].e); $('#t8f').textContent = fmt(vals[8].f);
  $('#t8h').textContent = fmt(vals[8].h); $('#t8i').textContent = fmt(vals[8].i);
  $('#bcBug').textContent = fmtLei(vals[8].e*1000); $('#bcAng').textContent = fmtLei(vals[8].f*1000);
  $('#phTitle').textContent = 'luna ' + LUNI[state.meta.luna-1].toUpperCase() + ' ' + state.meta.an;
  $('#phUnit').textContent = state.meta.unitate;
  runChecks(vals);
  return vals;
}

/* ---------- lista de control ---------- */
function runChecks(vals){
  const out = [];
  const t = vals[8];
  const near = (a,b,eps=0.0005)=>Math.abs(a-b)<=eps;

  // 1. corelatie BC39
  if(bc39ref && (bc39ref.bugetar!=null || bc39ref.angajament!=null)){
    const db = bc39ref.bugetar!=null ? Math.round(t.e*1000)-Math.round(bc39ref.bugetar) : 0;
    const da = bc39ref.angajament!=null ? Math.round(t.f*1000)-Math.round(bc39ref.angajament) : 0;
    const tinta = (col, cur, ref, d) => ref==null ? '' :
      `${col} este ${fmt(cur)}, trebuie ${fmt(ref/1000)} (${fmtLei(ref)} lei)` + (d===0 ? ' ✓' : d<0 ? ` — lipsesc ${fmt(-d/1000)}` : ` — în plus ${fmt(d/1000)}`);
    out.push(chk(db===0&&da===0,
      'Corelația cu fișa BC39',
      db===0&&da===0 ? `Bugetar ${fmtLei(t.e*1000)} lei și Angajament ${fmtLei(t.f*1000)} lei coincid cu BC39 (${bc39ref.sursa}).`
        : `Rândul 8: ${[tinta('col.3',t.e,bc39ref.bugetar,db), tinta('col.4',t.f,bc39ref.angajament,da)].filter(Boolean).join(' · ')}.`));
    $('#tbl').querySelectorAll('tr[data-row="8"] td[data-calc]').forEach(td=>{
      const k = td.dataset.calc.split('.')[1];
      const d = k==='e' ? db : k==='f' ? da : null;
      td.classList.toggle('bad', d!==null && d!==0); td.classList.toggle('good', d===0);
    });
  } else {
    $('#tbl').querySelectorAll('tr[data-row="8"] td.bad,tr[data-row="8"] td.good').forEach(td=>td.classList.remove('bad','good'));
    out.push(chk(null,'Corelația cu fișa BC39',
      `De verificat manual: Bugetar = ${fmtLei(t.e*1000)} lei, Angajament = ${fmtLei(t.f*1000)} lei. Importă fișa BC39 pentru verificare automată.`));
  }

  // 2. col.3 = col.5 = col.6 pe randul 8
  out.push(chk(near(t.e,t.g)&&near(t.e,t.h), 'Rândul 8: col.3 = col.5 = col.6',
    near(t.e,t.g)&&near(t.e,t.h) ? 'Creditele aprobate, cele trimestriale și plățile cumulate coincid.'
      : `col.3 ${fmt(t.e)} · col.5 ${fmt(t.g)} · col.6 ${fmt(t.h)} — diferențele înseamnă credite neconsumate; verifică dacă e intenționat.`, true));

  // 3. cumulat monoton fata de luna precedenta
  if(prev){
    const bad=[];
    LEAVES.forEach(r=>{
      const p=prev.leaves[r]||blank(), c=vals[r]||blank();
      ['e','g','h','j'].forEach(k=>{ if(c[k] < p[k]-0.0005 && !bad.includes(r)) bad.push(r); });
    });
    out.push(chk(bad.length===0,'Cumulat monoton față de '+prev.label,
      bad.length===0 ? 'Toate coloanele cumulate sunt ≥ luna precedentă.'
        : `${bad.length} rânduri scad față de luna precedentă: ${bad.slice(0,12).join(', ')}${bad.length>12?'…':''}`));
  } else {
    out.push(chk(null,'Cumulat monoton față de luna precedentă','Nicio lună precedentă încărcată — verificarea nu poate rula.'));
  }

  // 4. col.6 <= col.4
  const over = LEAVES.filter(r=>(vals[r].h - vals[r].f) > 0.0005);
  out.push(chk(over.length===0,'Plăți cumulate ≤ credite de angajament (col.6 ≤ col.4)',
    over.length===0 ? 'Nicio poziție cu plăți peste angajament.'
      : `${over.length} rânduri depășesc angajamentul: ${over.slice(0,12).map(r=>r+' ('+fmt(vals[r].h-vals[r].f)+')').join(', ')}${over.length>12?'…':''}`));

  // 5. col.8 >= col.6
  const neg = LEAVES.filter(r=>(vals[r].j - vals[r].h) < -0.0005);
  const gap = r3(t.j - t.h);
  out.push(chk(neg.length===0,'Cheltuieli efective ≥ plăți (col.8 ≥ col.6)',
    (neg.length===0 ? `Decalaj accrual–cash pe total: ${gap>=0?'+':''}${fmt(gap)} mii lei (facturi înregistrate și neplătite).`
      : `${neg.length} rânduri cu efective sub plăți: ${neg.slice(0,12).join(', ')}${neg.length>12?'…':''}`) +
     (prev ? ` În ${prev.label} decalajul era ${fmt(r3(prev.totals.j - prev.totals.h))}.` : ''), neg.length>0));

  // 6. integritatea sablonului
  if(TPL) out.push(chk(TPL.check.ok,'Integritatea șablonului (formula I117)',
    TPL.check.ok ? 'I117 are formulă proprie (=+I119+I118), nu copiază cumulatul.' : 'Șablon defect: '+TPL.check.reason));
  else out.push(chk(null,'Integritatea șablonului','Șablonul nu a fost încă încărcat.'));

  // 7. formatul BC39
  out.push(chk(null,'Formatul fișei BC39',
    bc39ref ? `Coloane de valori detectate: ${bc39ref.coloane.join(' · ')}. Numărul lor diferă de la o lună la alta; raportarea se face la același rând 8.`
            : 'Numărul coloanelor de valori din BC39 variază între luni (ex. „Spital" vs „Total" + „Spital (inclusiv PNS)"). Verifică antetul la fiecare depunere.'));

  $('#checks').innerHTML = out.map(chkHtml).join('');
  const errs = out.filter(c=>c.cls==='err').length;
  const g = guidance(vals, out);
  const gd = $('#guide');
  gd.className = 'guide '+g.cls;
  gd.innerHTML = `<span class="ic"></span><span class="tx">${g.text}</span>` + (g.more ? `<button class="lnk" id="guideMore">Toate verificările</button>` : '');
  const badge = $('#chkBadge');
  badge.textContent = errs ? String(errs) : '✓';
  badge.className = errs ? 'err' : 'ok';
  $('#btnSide').classList.toggle('has-err', errs>0);
  badge.title = errs ? errs + (errs===1?' problemă':' probleme') : 'Fără erori';
}
function chk(ok, title, det, warnIfFail){
  const cls = ok===null ? 'warn' : ok ? 'ok' : (warnIfFail?'warn':'err');
  const ic  = ok===null ? '?' : ok ? '✓' : '!';
  return { cls, ic, title, det };
}
const chkHtml = c => `<div class="chk ${c.cls}"><span class="ic">${c.ic}</span><span><b>${c.title}</b><span class="det">${c.det}</span></span></div>`;

/* ---------- ghidare: o singura linie care spune ce urmeaza ---------- */
function guidance(vals, checks){
  const t = vals[8];
  const gol = KEYS.every(k=>Math.abs(t[k])<1e-9);
  if(TPL && !TPL.check.ok) return { cls:'err', text:'Șablonul este defect: '+TPL.check.reason+'. Exportul ar produce o machetă greșită.' };
  if(gol && !prev && !bc39ref) return { cls:'info', text:'Începe prin a importa luna precedentă sau macheta lunii curente din meniul ⋯, apoi fișa BC39 pentru verificare. Sau completează direct rândurile.' };
  const err = checks.find(c=>c.cls==='err');
  if(err) return { cls:'err', text:err.title+': '+err.det, more:true };
  if(prev && state.mode==='cumulat' && gol) return { cls:'info', text:'Luna precedentă ('+prev.label+') e încărcată. Treci pe modul „delta" și completează doar plățile lunii (col. 7) și cheltuielile efective.' };
  if(!bc39ref) return { cls:'info', text:`Rândul 8: Bugetar ${fmtLei(t.e*1000)} lei, Angajament ${fmtLei(t.f*1000)} lei. Importă fișa BC39 din meniul ⋯ ca să verific corelația automat.` };
  const warn = checks.find(c=>c.cls==='warn' && c.ic==='!');
  if(warn) return { cls:'warn', text:warn.title+': '+warn.det, more:true };
  return { cls:'ok', text:'Se corelează cu BC39 ('+bc39ref.sursa+') și toate verificările trec. Poți exporta macheta .xlsx.' };
}

/* ---------- serializare ---------- */
function snapshot(withPrev){
  const vals = recompute();
  const totals = { e:vals[8].e, f:vals[8].f, g:vals[8].g, h:vals[8].h, i:vals[8].i, j:vals[8].j };
  return {
    versiune:1, generat:new Date().toISOString(),
    meta:state.meta, mode:state.mode, linkEG:state.linkEG, prevLabel:state.prevLabel,
    total_r8:totals,
    bc39:{ bugetar:Math.round(totals.e*1000), angajament:Math.round(totals.f*1000) },
    leaves:state.leaves, din:state.din,
    prev: (withPrev && prev) ? { label:prev.label, leaves:prev.leaves } : undefined
  };
}
function restore(data){
  if(!data||!data.leaves) throw new Error('Fișier de stare invalid.');
  state = newState();
  Object.assign(state.meta, data.meta||{});
  state.mode = data.mode||'cumulat'; state.linkEG = data.linkEG!==false; state.prevLabel = data.prevLabel||null;
  LEAVES.forEach(r=>{ if(data.leaves[r]) Object.assign(state.leaves[r], data.leaves[r]);
                      if(data.din&&data.din[r]) Object.assign(state.din[r], data.din[r]); });
  if(data.prev && data.prev.leaves) setPrev(data.prev.leaves, data.prev.label||'lună precedentă');
  // Modul delta se sprijina pe luna precedenta: fara ea, valorile cumulate ar fi recalculate gresit.
  if(state.mode==='delta' && !prev){
    state.mode='cumulat';
    toast('Luna precedentă lipsește din fișier — am trecut pe modul „cumulat", valorile rămân neschimbate.', true);
  }
  syncMetaInputs(); buildTable(); render();
}
function setPrev(leaves, label){
  const L = {}; LEAVES.forEach(r=>{ L[r] = {...blank(), ...(leaves[r]||{})}; });
  prev = { leaves:L, label, totals: aggregate(L)[8] };
  state.prevLabel = label;
  $('#prevInfo').textContent = 'Lună precedentă: ' + label;
  $('#modeDelta').disabled = false; $('#modeSeg').hidden = false;
  if($('#tbl').rows.length) render();
}

/* ---------- I/O ---------- */
async function exportXlsx(){
  if(!TPL) return toast('Șablonul nu este încărcat.', true);
  const vals = recompute();
  const all = {}; ROWS.forEach(r=>{ all[r.r] = vals[r.r]; });
  const luna = LUNI[state.meta.luna-1].toUpperCase();
  const blob = await buildXlsx(TPL.buf, all, { unitate:state.meta.unitate, titluLuna:`luna ${luna} ${state.meta.an}` });
  download(blob, `Macheta Executie spital ${luna} ${state.meta.an}.xlsx`);
  toast('Macheta exportată.');
}
function exportJson(){
  const b = new Blob([JSON.stringify(snapshot(),null,2)],{type:'application/json'});
  download(b, store.fileName(state.meta.an, state.meta.luna));
}
function download(blob, name){
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name; a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href), 4000);
}
async function onFile(input, target){
  const f = input.files[0]; if(!f) return;
  input.value='';
  $('#mainMenu').removeAttribute('open');
  try{
    if(f.name.toLowerCase().endsWith('.json')){
      const data = JSON.parse(await f.text());
      if(target==='prev') setPrev(data.leaves||{}, data.meta ? `${LUNI[(data.meta.luna||1)-1]} ${data.meta.an}` : f.name);
      else { restore(data); toast('Stare încărcată din JSON.'); }
      return;
    }
    const buf = await f.arrayBuffer();
    if(target==='bc39'){
      const b = await readBC39(buf);
      if(b.bugetar==null && b.angajament==null) throw new Error('Nu am găsit rândurile Bugetar/Angajament în fișier.');
      bc39ref = {...b, sursa:f.name};
      $('#bcRef').textContent = `BC39 importat: Bugetar ${fmtLei(b.bugetar||0)} · Angajament ${fmtLei(b.angajament||0)}`;
      render(); toast('Fișa BC39 importată.');
      return;
    }
    const m = await readMacheta(buf);
    if(Object.keys(m.values).length < 50) throw new Error('Fișierul nu pare a fi macheta CAS (foaia BC).');
    const lm = lunaDinText(m.luna), la = anDinText(m.luna);
    const eticheta = lm ? `${LUNI[lm-1]} ${la||''}`.trim() : (m.luna || f.name);
    if(target==='prev'){ setPrev(m.values, eticheta); toast('Lună precedentă încărcată: '+eticheta); }
    else {
      LEAVES.forEach(r=>{ if(m.values[r]) Object.assign(state.leaves[r], m.values[r]); });
      if(LEAVES.some(r=>{ const v=state.leaves[r]; return Math.abs(v.e-v.h)>0.0005 || Math.abs(v.g-v.h)>0.0005; })){
        state.linkEG = false; toast('Macheta are col. 3 sau 5 diferite de col. 6 pe unele rânduri — am dezlegat coloanele.');
      }
      if(lm) state.meta.luna = lm; if(la) state.meta.an = la;
      state.mode='cumulat'; syncMetaInputs(); buildTable(); render();
      toast('Machetă încărcată ca lună curentă'+(lm?' ('+eticheta+')':'')+'.');
    }
  }catch(err){ toast(err.message, true); }
}

/* ---------- GitHub ---------- */
async function ghRefresh(){
  const el = $('#ghList');
  el.innerHTML = '<div class="muted">Se încarcă…</div>';
  try{
    const acc = await store.checkAccess();
    const months = await store.listMonths();
    $('#ghStatus').textContent = `${acc.full} · ${acc.private?'privat':'public'} · ${acc.push?'scriere permisă':'doar citire'}`;
    el.innerHTML = months.length ? months.map(m=>`<div class="row-inline" style="justify-content:space-between;border-bottom:1px solid var(--line);padding:5px 0">
        <b>${m.name}</b>
        <span class="row-inline">
          <button class="btn" data-gh="cur" data-n="${m.name}">Deschide</button>
          <button class="btn" data-gh="prev" data-n="${m.name}">Ca lună precedentă</button>
        </span></div>`).join('')
      : '<div class="muted">Niciun fișier în directorul de date încă.</div>';
    el.querySelectorAll('button[data-gh]').forEach(b=>b.addEventListener('click', ()=>ghLoad(b.dataset.n, b.dataset.gh)));
  }catch(e){ el.innerHTML=''; $('#ghStatus').textContent = 'Eroare: '+e.message; }
}
async function ghLoad(name, as){
  try{
    const r = await store.loadMonth(name);
    if(!r) throw new Error('Fișierul nu există.');
    if(as==='prev'){ setPrev(r.data.leaves||{}, name.replace('.json','')); toast('Lună precedentă încărcată din GitHub.'); }
    else { restore(r.data); toast('Lună deschisă din GitHub.'); }
    $('#ghDlg').close();
  }catch(e){ toast(e.message, true); }
}
async function ghSave(){
  try{
    const name = store.fileName(state.meta.an, state.meta.luna);
    await store.saveMonth(name, snapshot(), `Execuție bugetară ${LUNI[state.meta.luna-1]} ${state.meta.an}`);
    toast('Salvat în GitHub: data/'+name);
    if($('#ghDlg').open) ghRefresh();
  }catch(e){ toast(e.message, true); }
}

/* ---------- diverse ---------- */
let saveTimer;
function scheduleSave(){ clearTimeout(saveTimer); saveTimer=setTimeout(()=>store.draft.save(snapshot(true)), 600); }
function toast(msg, bad){ const s=$('#status'); s.textContent=msg; s.classList.toggle('bad', !!bad); clearTimeout(toast._t); toast._t=setTimeout(()=>{s.textContent='';},6000); }
function setSide(open){ document.body.classList.toggle('side-open', open); }
function syncMetaInputs(){
  $('#an').value = state.meta.an; $('#luna').value = state.meta.luna; $('#unitate').value = state.meta.unitate;
  $('#modeCumulat').checked = state.mode==='cumulat'; $('#modeDelta').checked = state.mode==='delta';
  $('#linkEG').checked = state.linkEG;
}

/* ---------- pornire ---------- */
function wire(){
  $('#luna').innerHTML = LUNI.map((l,i)=>`<option value="${i+1}">${l}</option>`).join('');
  ['an','luna','unitate'].forEach(id=>$('#'+id).addEventListener('change', e=>{
    state.meta[id] = id==='unitate' ? e.target.value : +e.target.value; render(); scheduleSave();
  }));
  $('#modeCumulat').addEventListener('change', ()=>{ state.mode='cumulat'; buildTable(); render(); });
  $('#modeDelta').addEventListener('change', ()=>{
    if(!prev) return toast('Încarcă întâi luna precedentă.', true);
    state.mode='delta';
    LEAVES.forEach(r=>{ const p=prev.leaves[r]||blank(); if(!state.din[r].f) state.din[r].f = p.f||0; });
    buildTable(); render();
  });
  $('#linkEG').addEventListener('change', e=>{ state.linkEG=e.target.checked; buildTable(); render(); });
  $('#filtru').addEventListener('input', ()=>render());
  $('#hideZero').addEventListener('change', ()=>render());
  $('#btnXlsx').addEventListener('click', exportXlsx);
  $('#btnPdf').addEventListener('click', ()=>window.print());
  $('#btnJson').addEventListener('click', exportJson);
  $('#btnNew').addEventListener('click', ()=>{ if(confirm('Golești formularul curent?')){ state=newState(); store.draft.clear(); syncMetaInputs(); buildTable(); render(); }});
  $('#fPrev').addEventListener('change', e=>onFile(e.target,'prev'));
  $('#fCur').addEventListener('change', e=>onFile(e.target,'cur'));
  $('#fBc').addEventListener('change', e=>onFile(e.target,'bc39'));
  $('#btnGh').addEventListener('click', ()=>{ const c=store.cfg.get();
    $('#ghOwner').value=c.owner||''; $('#ghRepo').value=c.repo||''; $('#ghBranch').value=c.branch||'main';
    $('#ghDir').value=c.dir||'data'; $('#ghTok').value=store.cfg.token();
    $('#ghDlg').showModal(); if(c.owner&&c.repo) ghRefresh(); });
  $('#ghSaveCfg').addEventListener('click', ()=>{
    store.cfg.set({owner:$('#ghOwner').value.trim(), repo:$('#ghRepo').value.trim(), branch:$('#ghBranch').value.trim()||'main', dir:$('#ghDir').value.trim()||'data'});
    store.cfg.setToken($('#ghTok').value.trim()); ghRefresh(); });
  $('#ghPush').addEventListener('click', ghSave);
  $('#ghClose').addEventListener('click', ()=>$('#ghDlg').close());
  $('#btnSide').addEventListener('click', ()=>setSide(!document.body.classList.contains('side-open')));
  $('#guide').addEventListener('click', e=>{ if(e.target.id==='guideMore') setSide(true); });
  $('#btnSideClose').addEventListener('click', ()=>setSide(false));
  document.addEventListener('keydown', e=>{ if(e.key==='Escape' && document.body.classList.contains('side-open')) setSide(false); });
  document.addEventListener('click', e=>{ if(document.body.classList.contains('side-open') && !e.target.closest('.side') && !e.target.closest('#btnSide') && e.target.id!=='guideMore') setSide(false); });
  document.addEventListener('click', e=>{ const m=$('#mainMenu'); if(!m.open) return; if(!m.contains(e.target) || e.target.closest('.menu-list button')) m.removeAttribute('open'); });
  $('#btnTheme').addEventListener('click', ()=>{
    const cur = document.documentElement.getAttribute('data-theme') || (matchMedia('(prefers-color-scheme:dark)').matches ? 'dark' : 'light');
    const nxt = cur==='dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', nxt);
    try{ localStorage.setItem('execbug.theme', nxt); }catch{}
  });
}

async function init(){
  try{ const t=localStorage.getItem('execbug.theme'); if(t) document.documentElement.setAttribute('data-theme',t); }catch{}
  wire(); syncMetaInputs(); buildTable();
  const d = store.draft.load();
  if(d){ try{ restore(d); toast('Ciornă restaurată din acest browser.'); }catch{} }
  render();
  try{ TPL = await loadTemplate('./template/macheta.xlsx'); }
  catch(e){ toast('Șablon indisponibil: '+e.message, true); }
  render();
}
init();
