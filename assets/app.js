import { ROWS } from './rows.js?v=20260914-114551';
import { parseNum, r3, fmt, fmtLei } from './num.js?v=20260914-114551';
import { loadTemplate, buildXlsx, readMacheta, readBC39 } from './xlsx.js?v=20260914-114551';
import * as store from './store.js?v=20260914-114551';

const KEYS = ['e','f','g','h','i','j'];
const lunaDinText = t => { const i = LUNI.findIndex(l=>new RegExp(l,'i').test(String(t||''))); return i<0?null:i+1; };
const anDinText = t => { const m = String(t||'').match(/(20\d\d)/); return m?+m[1]:null; };
const LUNI = ['ianuarie','februarie','martie','aprilie','mai','iunie','iulie','august','septembrie','octombrie','noiembrie','decembrie'];
const byRow = Object.fromEntries(ROWS.map(r=>[r.r, r]));
const cuvCheie = t => String(t||'').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'')
  .replace(/[^A-Z0-9 ]/g,' ').split(/\s+/).filter(w=>w && !['SC','SRL','SA','S','C'].includes(w));
const aceeasiUnitate = (a,b) => { const A=cuvCheie(a), B=cuvCheie(b); return A.length && B.length && (A.every(w=>B.includes(w)) || B.every(w=>A.includes(w))); };
const LEAVES = ROWS.filter(r=>!r.ch).map(r=>r.r);
const $ = s => document.querySelector(s);
const blank = () => Object.fromEntries(KEYS.map(k=>[k,0]));

let TPL = null;            // {buf, check}
let state = newState();
let prev = null;           // {label, leaves}
let bc39ref = null;        // {bugetar, angajament, sursa, coloane, luna, an, denumire, detalii}
let importNote = null;     // {title, det} — inconsecvente gasite in ultima macheta importata

function newState(){
  const d = new Date();
  const leaves = {}; const din = {};
  LEAVES.forEach(r=>{ leaves[r]=blank(); din[r]={i:0,dj:null,f:null}; });   // f:null = col. 4 urmeaza col. 3; dj:null = Δ col. 8 urmeaza col. 7
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
      const d = state.din[r] || {i:0,dj:null,f:null};
      const L = state.leaves[r];
      L.i = r3(d.i);
      L.h = r3((p.h||0) + d.i);
      // col. 8 (cheltuieli efective): cheltuiala lunii = plata lunii (col. 7) pana cand utilizatorul tasteaza Δ pe rand;
      // in machetele reale col. 8 = col. 6 pe majoritatea randurilor, diferentele sunt facturi inregistrate si neplatite
      L.j = r3((p.j||0) + (d.dj==null ? d.i : d.dj));
      if(state.linkEG){ L.e = L.h; L.g = L.h; }
      // col. 4 (credite de angajament) se reface lunar din col. 3, cu cateva randuri rotunjite in sus:
      // urmeaza col. 3 pana cand utilizatorul tasteaza o valoare pe rand (din.f != null)
      L.f = d.f==null ? L.e : r3(d.f);
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
    <tr class="tinta" id="tintaRow" hidden>
      <th class="code" colspan="3">BC39</th><th class="lbl" id="tintaLbl"></th>
      ${KEYS.map(k=>`<th class="num" data-tinta="${k}"></th>`).join('')}<th class="num dcol"></th>
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
  masoaraAntet();
  legaEvidentiereColoana();
  $('#tbl').querySelectorAll('input').forEach(inp=>{
    inp.addEventListener('focus', e=>e.target.select());
    inp.addEventListener('input', onInput);
    inp.addEventListener('blur', e=>{ e.target.value = fmt(readInput(e.target)); });
    inp.addEventListener('keydown', onKey);
  });
}

/** Celula automata in delta: col. 4 fara valoare tastata urmeaza col. 3, Δ col. 8 fara valoare tastata urmeaza col. 7. */
const celulaAuto = (inp) => (inp.dataset.src==='din.f' && state.din[+inp.dataset.r].f==null) || (inp.dataset.src==='din.dj' && state.din[+inp.dataset.r].dj==null);
const TITLU_AUTO = { f:'Urmează col. 3. Tastează o valoare ca s-o fixezi; șterge-o ca să revină la col. 3.',
                     dj:'Urmează col. 7 (cheltuiala lunii = plata lunii). Tastează cheltuiala lunii ca s-o fixezi; șterge-o ca să revină.' };
function readInput(inp){
  const r = +inp.dataset.r, k = inp.dataset.k, src = inp.dataset.src;
  if(celulaAuto(inp)) return k==='f' ? (state.leaves[r].f||0) : (state.din[r].i||0);
  return src.startsWith('din') ? (state.din[r][k]||0) : (state.leaves[r][k]||0);
}
function onInput(e){
  const inp = e.target, r = +inp.dataset.r, k = inp.dataset.k;
  // col. 4 sau Δ golita revine la automat (urmeaza col. 3, respectiv col. 7)
  if((inp.dataset.src==='din.f' || inp.dataset.src==='din.dj') && inp.value.trim()===''){ state.din[r][k] = null; render(); scheduleSave(); return; }
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

function masoaraAntet(){
  requestAnimationFrame(()=>{
    const th = $('#tbl thead'); if(!th) return;
    document.documentElement.style.setProperty('--thead-h', th.offsetHeight+'px');
    document.documentElement.style.setProperty('--th1-h', th.rows[0].offsetHeight+'px');
  });
}
const bc39AltaLuna = () => !!(bc39ref && bc39ref.lunaNr && (bc39ref.lunaNr!==state.meta.luna || (bc39ref.an && +bc39ref.an!==+state.meta.an)));

/** Randul „Tinta BC39": valorile din fisa in coloanele lor (r.8 col.3 = Bugetar, col.4 = Angajament) si cat mai ramane. */
function renderTinta(vals){
  const tr = $('#tintaRow'); if(!tr) return;
  const arata = bc39ref && !bc39AltaLuna() && (bc39ref.bugetar!=null || bc39ref.angajament!=null);
  if(tr.hidden !== !arata){ tr.hidden = !arata; masoaraAntet(); }
  if(!arata) return;
  const bug = bc39ref.bugetar!=null ? bc39ref.bugetar/1000 : null;
  const ang = bc39ref.angajament!=null ? bc39ref.angajament/1000 : null;
  const t = { e:bug, f:ang, g:bug, h:bug, i:(prev && bug!=null) ? r3(bug - prev.totals.h) : null, j:null };
  const eticheta = bc39ref.lunaNr ? `${LUNI[bc39ref.lunaNr-1]} ${bc39ref.an||''}`.trim() : bc39ref.sursa;
  $('#tintaLbl').textContent = 'Țintă din fișa BC39 (' + eticheta + ')' + (prev ? ' · col. 7 = țintă − ' + prev.label : '');
  tr.querySelectorAll('th[data-tinta]').forEach(th=>{
    const k = th.dataset.tinta, tinta = t[k];
    if(tinta==null){ th.innerHTML = '<span class="muted">—</span>'; th.className='num'; return; }
    const d = Math.round(tinta*1000) - Math.round(vals[8][k]*1000);
    th.innerHTML = `<b>${fmt(tinta)}</b><small>${d===0 ? '✓ atins' : d>0 ? 'rămân '+fmt(d/1000) : 'în plus '+fmt(-d/1000)}</small>`;
    th.className = 'num ' + (d===0 ? 'good' : 'bad');
  });
}

/* Evidentierea coloanei sub cursor: un overlay translucid pozitionat pe coloana celulei peste care e mouse-ul. */
function legaEvidentiereColoana(){
  const tw = $('.tablewrap'), tbl = $('#tbl');
  let hl = tw.querySelector('.colhl');
  if(!hl){ hl = document.createElement('div'); hl.className='colhl'; tw.appendChild(hl); }
  let ultima = -1;
  const ths = () => tbl.tHead.rows[0].cells;
  const ascunde = () => { hl.style.display='none'; if(ultima>=0 && ths()[ultima]) ths()[ultima].classList.remove('hl'); ultima=-1; };
  tbl.onmouseover = e => {
    const cell = e.target.closest('td,th'); if(!cell || cell.parentElement.classList.contains('tinta')) return;
    const idx = cell.cellIndex;
    if(idx < 4){ ascunde(); return; }                  // doar coloanele de valori
    if(idx === ultima) return;
    if(ultima>=0 && ths()[ultima]) ths()[ultima].classList.remove('hl');
    const ref = ths()[idx]; if(!ref) return;
    ultima = idx; ref.classList.add('hl');
    hl.style.left = ref.offsetLeft+'px'; hl.style.width = ref.offsetWidth+'px'; hl.style.height = tbl.offsetHeight+'px'; hl.style.display='block';
  };
  tbl.onmouseleave = ascunde;
}

/* ---------- randare ---------- */
function render(){
  const vals = recompute();
  renderTinta(vals);
  document.querySelectorAll('#tbl td[data-calc]').forEach(td=>{
    const [r,k] = td.dataset.calc.split('.');
    td.textContent = fmt(vals[+r] ? vals[+r][k] : 0);
  });
  document.querySelectorAll('#tbl input').forEach(inp=>{
    const v = readInput(inp);
    if(document.activeElement!==inp) inp.value = fmt(v);
    inp.classList.toggle('z', Math.abs(v)<1e-9);
    const auto = celulaAuto(inp);
    inp.classList.toggle('auto', auto);
    inp.title = auto ? TITLU_AUTO[inp.dataset.k] : '';
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
  const nr = (n) => n===1 ? '1 rând' : n+' rânduri';

  // 1. corelatie BC39
  if(bc39ref && (bc39ref.bugetar!=null || bc39ref.angajament!=null)){
    const db = bc39ref.bugetar!=null ? Math.round(t.e*1000)-Math.round(bc39ref.bugetar) : 0;
    const da = bc39ref.angajament!=null ? Math.round(t.f*1000)-Math.round(bc39ref.angajament) : 0;
    const lunaBc = bc39ref.lunaNr ? `${LUNI[bc39ref.lunaNr-1]} ${bc39ref.an||''}`.trim() : null;
    const lunaForm = `${LUNI[state.meta.luna-1]} ${state.meta.an}`;
    const altaLuna = bc39AltaLuna();
    const peste = (ref) => (prev && ref!=null) ? ` — cu ${fmt(ref/1000 - prev.totals.h)} peste ${prev.label}` : '';
    const tinta = (col, cur, ref, d, extra, sfat) => ref==null ? '' :
      `${col} este ${fmt(cur)}, trebuie ${fmt(ref/1000)} (${fmtLei(ref)} lei${extra||''})` + (d===0 ? ' ✓' : (d<0 ? ` — lipsesc ${fmt(-d/1000)}` : ` — în plus ${fmt(d/1000)}`) + (sfat||''));
    // in delta, col. 4 porneste de la col. 3: diferenta fata de BC39 se acopera rotunjind cateva randuri
    const sfat4 = state.mode==='delta' ? ': col. 4 urmează col. 3, tastează în col. 4 pe rândurile de rotunjit' : '';
    if(altaLuna){
      out.push(chk(false,'Corelația cu fișa BC39',
        `Fișa BC39 este pentru ${lunaBc}, dar formularul este pe ${lunaForm}. Schimbă luna formularului sau importă fișa BC39 a lunii corecte.`));
    } else {
      const unit = bc39ref.denumire && !aceeasiUnitate(bc39ref.denumire, state.meta.unitate)
        ? ` Atenție: fișa este pentru „${bc39ref.denumire}", formularul pentru „${state.meta.unitate}".` : '';
      out.push(chk(db===0&&da===0,
        'Corelația cu fișa BC39',
        (db===0&&da===0 ? `Bugetar ${fmtLei(t.e*1000)} lei și Angajament ${fmtLei(t.f*1000)} lei coincid cu BC39 (${bc39ref.sursa}${lunaBc?', '+lunaBc:''}).`
          : `Rândul 8: ${[tinta('col.3',t.e,bc39ref.bugetar,db,peste(bc39ref.bugetar)), tinta('col.4',t.f,bc39ref.angajament,da,'',sfat4)].filter(Boolean).join(' · ')}.`) + unit));
    }
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
  const c2 = near(t.e,t.g)&&near(t.e,t.h);
  out.push(chk(c2, 'Rândul 8: col.3 = col.5 = col.6',
    c2 ? (state.linkEG ? 'Coloanele sunt legate prin opțiune.' : 'Creditele aprobate, cele trimestriale și plățile cumulate coincid.')
      : `col.3 ${fmt(t.e)} · col.5 ${fmt(t.g)} · col.6 ${fmt(t.h)} — diferențele înseamnă credite neconsumate; verifică dacă e intenționat.`, true, 'col. 3 = 5 = 6'));

  // 3. cumulat monoton fata de luna precedenta
  if(prev){
    const bad=[];
    LEAVES.forEach(r=>{
      const p=prev.leaves[r]||blank(), c=vals[r]||blank();
      ['e','g','h','j'].forEach(k=>{ if(c[k] < p[k]-0.0005 && !bad.includes(r)) bad.push(r); });
    });
    out.push(chk(bad.length===0,'Cumulat monoton față de '+prev.label,
      bad.length===0 ? 'Toate coloanele cumulate sunt ≥ luna precedentă.'
        : `${nr(bad.length)} ${bad.length===1?'scade':'scad'} față de ${prev.label}: ${bad.slice(0,12).join(', ')}${bad.length>12?'…':''}. Un cumulat nu poate scădea — verifică plățile lunii pe aceste rânduri.`,
      false, 'cumulat ≥ '+prev.label));
  }

  // 4. col.6 <= col.4
  const over = LEAVES.filter(r=>(vals[r].h - vals[r].f) > 0.0005);
  // in aprilie 2026 macheta reala avea r.52 cu plati peste angajament, deci e avertisment, nu eroare
  out.push(chk(over.length===0,'Plăți cumulate ≤ credite de angajament (col.6 ≤ col.4)',
    over.length===0 ? 'Nicio poziție cu plăți peste angajament.'
      : `${nr(over.length)} ${over.length===1?'are':'au'} plăți (col. 6) peste creditele de angajament (col. 4): ${over.slice(0,12).map(r=>r+' (+'+fmt(vals[r].h-vals[r].f)+')').join(', ')}${over.length>12?'…':''}. Rotunjește col. 4 în sus pe aceste rânduri sau verifică plățile.`,
    true, 'plăți ≤ angajament'));

  // 5. col.8 >= col.6
  const neg = LEAVES.filter(r=>(vals[r].j - vals[r].h) < -0.0005);
  const gap = r3(t.j - t.h);
  const reper = prev ? ` În ${prev.label} col. 8 depășea col. 6 cu ${fmt(r3(prev.totals.j - prev.totals.h))}.` : '';
  const deltaAuto = state.mode==='delta' && LEAVES.every(r=>(state.din[r]||{}).dj==null);
  if(state.mode==='cumulat' && Math.abs(t.j)<1e-9 && Math.abs(t.h)>1e-9){
    out.push(chk(false,'Cheltuieli efective (col. 8) necompletate','Col. 8 este goală. Completează cheltuielile efective cumulate pe rânduri.' + reper, true));
  } else if(deltaAuto){
    // nicio Δ tastata: col. 8 = luna precedenta + col. 7 pe toate randurile; decalajul fata de plati ramane cel din luna precedenta
    const cuFacturi = LEAVES.filter(r=>{ const p=prev.leaves[r]||blank(); return (p.j - p.h) > 0.0005; });
    out.push(chk(null,'Cheltuieli efective (col. 8)',
      `Urmează plățile lunii (col. 7) pe toate rândurile: col. 8 = ${fmt(t.j)}, cu ${fmt(gap)} peste plăți, ca în ${prev.label}. ` +
      `Dacă în ${LUNI[state.meta.luna-1]} s-au înregistrat facturi neplătite sau s-au plătit facturi mai vechi, tastează în coloana Δ cheltuiala lunii pe acele rânduri.` +
      (cuFacturi.length ? ` În ${prev.label} aveau facturi neplătite rândurile ${cuFacturi.join(', ')}.` : '')));
  } else {
    // in aprilie 2026 macheta reala avea 2 randuri cu col. 8 sub col. 6, deci e avertisment
    out.push(chk(neg.length===0,'Cheltuieli efective ≥ plăți (col.8 ≥ col.6)',
      (neg.length===0 ? `Col. 8 depășește plățile cu ${gap>=0?'+':''}${fmt(gap)} (facturi înregistrate și neplătite).` +
          (state.mode==='delta' ? ` Δ tastată pe ${nr(LEAVES.filter(r=>state.din[r].dj!=null).length)}, restul urmează col. 7.` : '')
        : `${nr(neg.length)} ${neg.length===1?'are':'au'} cheltuieli efective (col. 8) sub plăți (col. 6): ${neg.slice(0,12).join(', ')}${neg.length>12?'…':''}. De regulă col. 8 ≥ col. 6 — verifică dacă e corect.`) + reper, true));
  }

  // 6. integritatea sablonului
  if(TPL) out.push(chk(TPL.check.ok,'Integritatea șablonului (formula I117)',
    TPL.check.ok ? 'I117 are formulă proprie (=+I119+I118), nu copiază cumulatul.' : 'Șablon defect: '+TPL.check.reason, false, 'șablon'));
  else out.push(chk(null,'Integritatea șablonului','Șablonul nu a fost încă încărcat.'));

  // 7. formatul BC39: interesant doar cand coloanele fisei difera intre ele (PNS / alte unitati)
  if(bc39ref){
    const d = bc39ref.detalii||{};
    const difera = ['bugetar','angajament'].filter(k=>{ const v=Object.values(d[k]||{}); return v.length>1 && v.some(x=>Math.round(x)!==Math.round(v[0])); });
    const col = `Coloane de valori: ${(bc39ref.coloane||[]).join(' · ')}; referința este „${bc39ref.coloanaRef}" (rândul 8 include PNS).`;
    if(difera.length) out.push(chk(null,'Formatul fișei BC39', col + ` Coloanele diferă între ele la ${difera.join(' și ')}: ` + difera.map(k=>Object.entries(d[k]).map(([c,v])=>`${c} ${fmtLei(v)}`).join(' / ')).join('; ') + ' — diferența înseamnă PNS sau alte unități.'));
    else out.push(chk(true,'Formatul fișei BC39', col + ' Valorile coincid între coloane.', false, 'format BC39'));
  }

  if(importNote) out.push({ cls:'warn', ic:'!', title:importNote.title, det:importNote.det });
  const trecute = out.filter(c=>c.cls==='ok' && c.scurt).map(c=>c.scurt);
  $('#checks').innerHTML = out.filter(c=>!(c.cls==='ok' && c.scurt)).map(chkHtml).join('') +
    (trecute.length ? chkHtml({ cls:'ok', ic:'✓', title:'Trec fără observații', det:trecute.join(' · ') }) : '');
  const errs = out.filter(c=>c.cls==='err').length;
  const g = guidance(vals, out);
  const gd = $('#guide');
  gd.className = 'guide noprint '+g.cls;   // noprint trebuie pastrat: linia nu apare in PDF
  gd.innerHTML = `<span class="ic"></span><span class="tx">${g.text}</span>` + (g.more ? `<button class="lnk" id="guideMore">Toate verificările</button>` : '') + (g.reload ? `<button class="btn pri" id="guideReload">Reîncarcă aplicația</button>` : '');
  const badge = $('#chkBadge');
  badge.textContent = errs ? String(errs) : '✓';
  badge.className = errs ? 'err' : 'ok';
  $('#btnSide').classList.toggle('has-err', errs>0);
  badge.title = errs ? errs + (errs===1?' problemă':' probleme') : 'Fără erori';
}
/** scurt: eticheta cu care verificarea trecuta se strange in linia „Trec fara observatii" (fara ea ramane vizibila). */
function chk(ok, title, det, warnIfFail, scurt){
  const cls = ok===null ? 'warn' : ok ? 'ok' : (warnIfFail?'warn':'err');
  const ic  = ok===null ? '?' : ok ? '✓' : '!';
  return { cls, ic, title, det, scurt };
}
const chkHtml = c => `<div class="chk ${c.cls}"><span class="ic">${c.ic}</span><span><b>${c.title}</b><span class="det">${c.det}</span></span></div>`;

/* ---------- ghidare: o singura linie care spune ce urmeaza ---------- */
function guidance(vals, checks){
  const t = vals[8];
  const gol = KEYS.every(k=>Math.abs(t[k])<1e-9);
  if(versiuneNoua) return { cls:'warn', text:`Este publicată o versiune nouă a aplicației (${versiuneNoua}); rulezi ${VERSIUNE}. Ciorna se păstrează.`, reload:true };
  if(TPL && !TPL.check.ok) return { cls:'err', text:'Șablonul este defect: '+TPL.check.reason+'. Exportul ar produce o machetă greșită.' };
  if(gol && !prev && !bc39ref) return { cls:'info', text:'Începe prin a importa luna precedentă sau macheta lunii curente din meniul ⋯, apoi fișa BC39 pentru verificare. Sau completează direct rândurile.' };
  const err = checks.find(c=>c.cls==='err');
  if(err) return { cls:'err', text:err.title+': '+err.det, more:true };
  if(prev && state.mode==='cumulat' && gol) return { cls:'info', text:'Luna precedentă ('+prev.label+') e încărcată. Treci pe modul „delta" și completează doar plățile lunii (col. 7) și cheltuielile efective.' };
  if(!bc39ref) return { cls:'info', text:`Rândul 8: Bugetar ${fmtLei(t.e*1000)} lei, Angajament ${fmtLei(t.f*1000)} lei. Importă fișa BC39 din meniul ⋯ ca să verific corelația automat.` };
  const warn = checks.find(c=>c.cls==='warn' && c.ic==='!');
  if(warn) return { cls:'warn', text:warn.title+': '+warn.det, more:true };
  const deltaAuto = state.mode==='delta' && LEAVES.every(r=>(state.din[r]||{}).dj==null);
  return { cls:'ok', text:'Se corelează cu BC39 ('+bc39ref.sursa+') și toate verificările trec. ' +
    (deltaAuto ? 'Col. 8 urmează col. 7 (fără facturi neplătite noi); dacă e corect, poți exporta macheta .xlsx.' : 'Poți exporta macheta .xlsx.'), more: deltaAuto };
}

/* ---------- serializare ---------- */
function snapshot(withPrev){
  const vals = recompute();
  const totals = { e:vals[8].e, f:vals[8].f, g:vals[8].g, h:vals[8].h, i:vals[8].i, j:vals[8].j };
  return {
    versiune:3, generat:new Date().toISOString(),
    meta:state.meta, mode:state.mode, linkEG:state.linkEG, prevLabel:state.prevLabel,
    total_r8:totals,
    bc39:{ bugetar:Math.round(totals.e*1000), angajament:Math.round(totals.f*1000) },
    leaves:state.leaves, din:state.din,
    bc39ref: bc39ref || undefined,
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
  // versiunea 1 copia col. 4 din luna precedenta; acum col. 4 urmeaza col. 3 pana e tastata (din.f = null)
  if(!(data.versiune>=2) && state.mode==='delta' && prev){
    LEAVES.forEach(r=>{ const d=state.din[r], pf=(prev.leaves[r]||blank()).f||0; if(!d.f || Math.abs(d.f-pf)<1e-9) d.f=null; });
  }
  // pana la versiunea 2, Δ = 0 insemna „netastat"; acum Δ netastata (null) urmeaza col. 7
  if(!(data.versiune>=3) && state.mode==='delta'){
    LEAVES.forEach(r=>{ const d=state.din[r]; if(!d.dj) d.dj=null; });
  }
  if(data.bc39ref && (data.bc39ref.bugetar!=null || data.bc39ref.angajament!=null)) setBc39(data.bc39ref);
  else { bc39ref = null; $('#bcRef').textContent=''; }
  importNote = null;
  // Modul delta se sprijina pe luna precedenta: fara ea, valorile cumulate ar fi recalculate gresit.
  if(state.mode==='delta' && !prev){
    state.mode='cumulat';
    toast('Luna precedentă lipsește din fișier — am trecut pe modul „cumulat", valorile rămân neschimbate.', true);
  }
  syncMetaInputs(); buildTable(); render(); scheduleSave();
  if(prev && state.mode==='cumulat' && faraCumulat()) treciPeDelta(true);
}
/** Formularul nu are inca date cumulate (col. 3/6/8 goale pe toate randurile)? */
const faraCumulat = () => LEAVES.every(r=>{ const L=state.leaves[r]; return Math.abs(L.h)<1e-9 && Math.abs(L.j)<1e-9 && Math.abs(L.e)<1e-9; });
/** Trece pe modul delta fata de luna precedenta, pastrand ce a fost tastat in modul cumulat (col. 7, col. 4, col. 8). */
function treciPeDelta(anunta){
  if(!prev) return;
  state.mode = 'delta';
  LEAVES.forEach(r=>{
    const p = prev.leaves[r]||blank(), L = state.leaves[r], d = state.din[r];
    if(!d.i && L.i) d.i = L.i;                                   // platile lunii tastate in col. 7
    if(!d.f) d.f = (L.f && Math.abs(L.f-L.e)>1e-9) ? L.f : null;   // col. 4 tastata diferit de col. 3 se pastreaza; altfel urmeaza col. 3
    if(d.dj==null && L.j) d.dj = r3(L.j - (p.j||0));              // cheltuieli efective tastate cumulat; altfel Δ urmeaza col. 7
  });
  syncMetaInputs(); buildTable(); render(); scheduleSave();
  if(anunta) toast(`Mod „delta" față de ${prev.label}: completezi plățile lunii (col. 7) și delta cheltuielilor efective; cumulatele pornesc de la ${prev.label}. Col. 4 urmează col. 3 și Δ col. 8 urmează col. 7 — tastezi doar rândurile care diferă.`);
}
/** Dupa incarcarea lunii precedente: daca luna curenta nu are inca date cumulate, treci automat pe delta. */
function dupaPrev(){ if(state.mode==='cumulat' && faraCumulat()) setTimeout(()=>treciPeDelta(true), 50); }
function setPrev(leaves, label){
  const L = {}; LEAVES.forEach(r=>{ L[r] = {...blank(), ...(leaves[r]||{})}; });
  prev = { leaves:L, label, totals: aggregate(L)[8] };
  state.prevLabel = label;
  $('#prevInfo').textContent = 'Lună precedentă: ' + label;
  $('#prevHdr').textContent = 'față de ' + label; $('#prevHdr').hidden = false;
  $('#modeDelta').disabled = false; $('#modeSeg').hidden = false;
  if($('#tbl').rows.length) render();
}

function setBc39(b){
  const an = b.an ? (String(b.an).match(/\d{4}/)||[])[0] : null;
  bc39ref = { ...b, lunaNr: lunaDinText(b.luna), an: an ? +an : null };
  $('#bcRef').textContent = `BC39 (${b.sursa||'importat'}): Bugetar ${fmtLei(b.bugetar||0)} · Angajament ${fmtLei(b.angajament||0)} lei` +
    (bc39ref.lunaNr ? ` · ${LUNI[bc39ref.lunaNr-1]} ${bc39ref.an||''}` : '') + (b.denumire ? ` · ${b.denumire}` : '');
}
/** Compara totalurile statice din macheta importata cu suma randurilor lor. Intoarce lista diferentelor. */
function verificaConsecventa(m){
  const vals = aggregate(m.values);
  const dif = [];
  ROWS.filter(r=>r.ch).forEach(row=>{
    const st = m.statice[row.r]||{};
    KEYS.forEach((k,i)=>{ const f=st[k]; if(f==null) return;
      const c = vals[row.r][k]; if(Math.abs(f-c)>0.0005) dif.push(`r.${row.r} col.${i+3}: în fișier ${fmt(f)}, suma rândurilor ${fmt(c)}`); });
  });
  return dif;
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
      if(target==='prev'){ setPrev(data.leaves||{}, data.meta ? `${LUNI[(data.meta.luna||1)-1]} ${data.meta.an}` : f.name); dupaPrev(); }
      else { restore(data); toast('Stare încărcată din JSON.'); }
      return;
    }
    const buf = await f.arrayBuffer();
    if(target==='bc39'){
      const b = await readBC39(buf);
      if(b.bugetar==null && b.angajament==null) throw new Error('Nu am găsit rândurile Bugetar/Angajament în fișier.');
      setBc39({...b, sursa:f.name});
      // formular gol: preia luna si anul din fisa
      const vals = recompute();
      const gol = KEYS.every(k=>Math.abs(vals[8][k])<1e-9);
      let msg = 'Fișa BC39 importată';
      if(bc39ref.lunaNr){
        const eticheta = `${LUNI[bc39ref.lunaNr-1]} ${bc39ref.an||''}`.trim();
        if(gol && !prev){ state.meta.luna = bc39ref.lunaNr; if(bc39ref.an) state.meta.an = bc39ref.an; syncMetaInputs(); msg += ` — formularul a fost setat pe ${eticheta}`; }
        else msg += ` (${eticheta})`;
      }
      render(); scheduleSave(); toast(msg+'.');
      return;
    }
    const m = await readMacheta(buf);
    if(Object.keys(m.values).length < 50) throw new Error('Fișierul nu pare a fi macheta CAS (foaia BC).');
    const lm = lunaDinText(m.luna), la = anDinText(m.luna);
    const eticheta = lm ? `${LUNI[lm-1]} ${la||''}`.trim() : (m.luna || f.name);
    const dif = verificaConsecventa(m);
    importNote = dif.length ? { title:`Macheta importată (${eticheta}) are totaluri care nu corespund sumei rândurilor`,
      det: dif.slice(0,6).join(' · ') + (dif.length>6?` · încă ${dif.length-6}`:'') + '. Aplicația recalculează totalurile de jos în sus; verifică rândurile respective în fișierul sursă.' } : null;
    if(target==='prev'){ setPrev(m.values, eticheta); scheduleSave(); toast('Lună precedentă încărcată: '+eticheta+(dif.length?' — are totaluri inconsecvente, vezi verificările':'')); dupaPrev(); }
    else {
      LEAVES.forEach(r=>{ if(m.values[r]) Object.assign(state.leaves[r], m.values[r]); });
      if(LEAVES.some(r=>{ const v=state.leaves[r]; return Math.abs(v.e-v.h)>0.0005 || Math.abs(v.g-v.h)>0.0005; })){
        state.linkEG = false; toast('Macheta are col. 3 sau 5 diferite de col. 6 pe unele rânduri — am dezlegat coloanele.');
      }
      if(lm) state.meta.luna = lm; if(la) state.meta.an = la;
      if(m.unitate && m.unitate.trim()) state.meta.unitate = m.unitate.trim();
      state.mode='cumulat'; syncMetaInputs(); buildTable(); render(); scheduleSave();
      toast('Machetă încărcată ca lună curentă'+(lm?' ('+eticheta+')':'')+'.');
    }
  }catch(err){ toast(err.message, true); }
}

/* ---------- salvarea lunii: arhiva din browser + GitHub daca e configurat ---------- */
const cheiaLunii = () => `${state.meta.an}-${String(state.meta.luna).padStart(2,'0')}`;
const numeLunii = k => { const [a,l]=k.split('-'); return `${LUNI[(+l||1)-1]} ${a}`; };
const amprenta = () => JSON.stringify({ m:state.meta, mode:state.mode, l:state.linkEG, leaves:state.leaves, din:state.din, p:state.prevLabel });
let amprentaSalvata = null;
function marcheazaModificari(){
  const sal = store.luni.get(cheiaLunii());
  const dirty = !sal || amprentaSalvata !== amprenta();
  $('#btnSave').classList.toggle('dirty', dirty);
  $('#btnSave').title = dirty ? 'Modificări nesalvate — salvează luna (Ctrl/Cmd+S)' : 'Luna este salvată (' + new Date(sal.salvat||sal.generat).toLocaleString('ro-RO') + ')';
}
async function salveazaLuna(){
  const k = cheiaLunii();
  const data = snapshot(true); data.salvat = new Date().toISOString();
  try{ store.luni.set(k, data); }
  catch(e){ return toast('Nu am putut salva în browser: '+e.message, true); }
  amprentaSalvata = amprenta(); marcheazaModificari();
  const unde = ['în acest browser'];
  const erori = [];
  // 1. folderul de pe disc (fara token; permisiunea se cere la primul clic)
  if(store.folder.suportat() && await store.folder.get()){
    const p = await store.folder.permisiune(true);
    if(p==='granted'){
      try{ await store.folder.scrie(k+'.json', JSON.stringify(data, null, 2)); unde.push('în folderul „'+((await store.folder.get()).name||'ales')+'"'); }
      catch(e){ erori.push('folder: '+e.message); }
    } else erori.push('folderul de salvare nu are permisiune (alege-l din nou din meniu)');
  }
  // 2. GitHub, doar daca e configurat cu token
  const c = store.cfg.get();
  if(c.owner && c.repo && store.cfg.token()){
    try{ await store.saveMonth(k+'.json', snapshot(), `Execuție bugetară ${numeLunii(k)}`); unde.push(`în GitHub (${c.owner}/${c.repo})`); }
    catch(e){ erori.push('GitHub: '+e.message); }
  }
  let msg = `Salvat ${numeLunii(k)} ${unde.join(', ')}.`;
  if(unde.length===1 && store.folder.suportat() && !(await store.folder.get())) msg += ' Alege un folder de salvare din meniul ⋯ ca să ai fișierul și pe disc.';
  if(erori.length) msg += ' Nu s-a putut salva ' + erori.join('; ') + '.';
  toast(msg, erori.length>0);
}
async function alegeFolder(){
  if(!store.folder.suportat()) return toast('Browserul acesta nu permite scrierea în foldere. Folosește Chrome sau Edge; între timp lunile se salvează în browser.', true);
  try{
    const h = await store.folder.alege();
    await actualizeazaFolderInfo();
    toast(`Folder de salvare: „${h.name}". De acum „Salvează luna" scrie și fișierul AAAA-LL.json acolo.`);
  }catch(e){ if(e.name!=='AbortError') toast('Nu am putut alege folderul: '+e.message, true); }
}
async function actualizeazaFolderInfo(){
  const el = $('#folderInfo');
  if(!store.folder.suportat()){ el.textContent = 'Salvarea pe disc nu e disponibilă în acest browser.'; $('#btnFolder').hidden = true; return; }
  const h = await store.folder.get();
  if(!h){ el.textContent = 'Niciun folder ales — lunile se salvează doar în browser.'; return; }
  const p = await store.folder.permisiune(false);
  el.innerHTML = `Folder: <b>${h.name||'(ales)'}</b>` + (p==='granted' ? '' : ' · <span class="muted">permisiunea se cere la prima salvare</span>') +
    ` · <a href="#" id="folderUita">uită folderul</a>`;
  $('#folderUita').addEventListener('click', async e=>{ e.preventDefault(); await store.folder.uita(); actualizeazaFolderInfo(); toast('Folderul a fost uitat; lunile se salvează doar în browser.'); });
}
async function deschideLuna(k, ca){
  // fisierul de pe disc are prioritate (poate fi modificat si din alta parte), apoi arhiva din browser
  let d = null;
  if(await store.folder.get() && (await store.folder.permisiune(true))==='granted') d = await store.folder.citeste(k+'.json');
  if(!d) d = store.luni.get(k);
  if(!d) return toast('Luna nu există nici în folder, nici în arhiva din browser.', true);
  if(ca==='prev'){ setPrev(d.leaves||{}, numeLunii(k)); scheduleSave(); toast('Lună precedentă: '+numeLunii(k)); dupaPrev(); }
  else { restore(d); amprentaSalvata = amprenta(); marcheazaModificari(); toast('Deschis pentru modificări: '+numeLunii(k)); }
  $('#luniDlg').close();
}
async function listeazaLuni(){
  const el = $('#luniList'); const cur = cheiaLunii();
  el.innerHTML = '<div class="muted">Se încarcă…</div>';
  const surse = {};                                   // key -> { browser, disc }
  store.luni.keys().forEach(k=>{ surse[k] = { browser: store.luni.get(k), disc:null }; });
  const h = await store.folder.get();
  let sursa = 'Sursa: arhiva din acest browser.';
  if(h){
    const p = await store.folder.permisiune(true);
    if(p==='granted'){
      for(const f of await store.folder.lista()){ (surse[f.key] ||= {browser:null, disc:null}).disc = await store.folder.citeste(f.name) || {}; }
      sursa = `Sursa: folderul „${h.name||'ales'}" (are prioritate) și arhiva din acest browser.`;
    } else sursa = `Folderul „${h.name}" nu are permisiune — se afișează doar arhiva din browser.`;
  }
  $('#luniSursa').textContent = sursa;
  const keys = Object.keys(surse).sort().reverse();
  if(!keys.length){ el.innerHTML = '<div class="muted">Nicio lună salvată încă. Completează formularul și apasă „Salvează luna".</div>'; return; }
  el.innerHTML = keys.map(k=>{ const S=surse[k]; const d=S.disc||S.browser; const t=d.total_r8||{};
    const src = (S.disc?'<span class="src disc">disc</span>':'') + (S.browser?'<span class="src">browser</span>':'');
    return `<div class="luna-i ${k===cur?'cur':''}">
      <div><div class="n">${numeLunii(k)}${k===cur?' <span class="muted">(în lucru)</span>':''}${src}</div>
        <div class="d">plăți cumulate ${fmt(t.h||0)} · luna ${fmt(t.i||0)} · salvat ${d.salvat||d.generat ? new Date(d.salvat||d.generat).toLocaleString('ro-RO') : '—'}</div></div>
      <div class="acts">
        <button class="btn" data-act="open" data-k="${k}">Deschide</button>
        <button class="btn" data-act="prev" data-k="${k}">Ca lună precedentă</button>
        <button class="btn" data-act="json" data-k="${k}">JSON</button>
        <button class="btn" data-act="del" data-k="${k}">Șterge</button>
      </div></div>`; }).join('');
  el.querySelectorAll('button[data-act]').forEach(b=>b.addEventListener('click', ()=>{
    const k=b.dataset.k;
    if(b.dataset.act==='open') deschideLuna(k);
    else if(b.dataset.act==='prev') deschideLuna(k,'prev');
    else if(b.dataset.act==='json'){ const d=(surse[k].disc||surse[k].browser); download(new Blob([JSON.stringify(d,null,2)],{type:'application/json'}), k+'.json'); }
    else if(b.dataset.act==='del'){ if(confirm('Ștergi luna '+numeLunii(k)+(surse[k].disc?' din folder și':'')+' din arhiva acestui browser?')){ store.luni.remove(k); if(surse[k].disc) store.folder.sterge(k+'.json').catch(e=>toast('Nu am putut șterge fișierul: '+e.message,true)).then(listeazaLuni); else listeazaLuni(); marcheazaModificari(); } }
  }));
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
    if(as==='prev'){ setPrev(r.data.leaves||{}, name.replace('.json','')); toast('Lună precedentă încărcată din GitHub.'); dupaPrev(); }
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
function scheduleSave(){ clearTimeout(saveTimer); saveTimer=setTimeout(()=>{ store.draft.save(snapshot(true)); marcheazaModificari(); }, 600); }
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
    state.meta[id] = id==='unitate' ? e.target.value : +e.target.value; render(); scheduleSave(); amprentaSalvata=null; marcheazaModificari();
  }));
  $('#modeCumulat').addEventListener('change', ()=>{ state.mode='cumulat'; buildTable(); render(); });
  $('#modeDelta').addEventListener('change', ()=>{
    if(!prev) return toast('Încarcă întâi luna precedentă.', true);
    treciPeDelta(true);
  });
  $('#linkEG').addEventListener('change', e=>{ state.linkEG=e.target.checked; buildTable(); render(); });
  $('#filtru').addEventListener('input', ()=>render());
  $('#hideZero').addEventListener('change', ()=>render());
  $('#btnXlsx').addEventListener('click', exportXlsx);
  $('#btnSave').addEventListener('click', salveazaLuna);
  $('#btnLuni').addEventListener('click', ()=>{ $('#luniDlg').showModal(); listeazaLuni(); });
  $('#btnFolder').addEventListener('click', alegeFolder);
  $('#luniClose').addEventListener('click', ()=>$('#luniDlg').close());
  document.addEventListener('keydown', e=>{ if((e.metaKey||e.ctrlKey) && e.key.toLowerCase()==='s'){ e.preventDefault(); salveazaLuna(); } });
  $('#btnPdf').addEventListener('click', ()=>window.print());
  $('#btnJson').addEventListener('click', exportJson);
  $('#btnNew').addEventListener('click', ()=>{ if(confirm('Golești formularul curent?')){ state=newState(); store.draft.clear(); bc39ref=null; importNote=null; prev=null; $('#bcRef').textContent=''; $('#prevHdr').hidden=true; $('#modeSeg').hidden=true; $('#modeDelta').disabled=true; $('#prevInfo').textContent='Nicio lună precedentă încărcată'; syncMetaInputs(); buildTable(); render(); }});
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
  $('#guide').addEventListener('click', e=>{ if(e.target.id==='guideMore') setSide(true); if(e.target.id==='guideReload') reincarca(); });
  $('#btnSideClose').addEventListener('click', ()=>setSide(false));
  document.addEventListener('keydown', e=>{ if(e.key==='Escape' && document.body.classList.contains('side-open')) setSide(false); });
  document.addEventListener('click', e=>{ if(document.body.classList.contains('side-open') && !e.target.closest('.side') && !e.target.closest('#btnSide') && e.target.id!=='guideMore') setSide(false); });
  document.addEventListener('click', e=>{ const m=$('#mainMenu'); if(!m.open) return; if(!m.contains(e.target) || e.target.closest('.menu-list button')) m.removeAttribute('open'); });
  $('#versiune').textContent = 'versiune ' + VERSIUNE;
  $('#btnTheme').addEventListener('click', ()=>{
    const cur = document.documentElement.getAttribute('data-theme') || (matchMedia('(prefers-color-scheme:dark)').matches ? 'dark' : 'light');
    const nxt = cur==='dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', nxt);
    try{ localStorage.setItem('execbug.theme', nxt); }catch{}
  });
}

/* ---------- versiune: detecteaza codul nou publicat, chiar daca pagina a venit din cache ---------- */
const VERSIUNE = new URL(import.meta.url).searchParams.get('v') || 'necunoscută';
let versiuneNoua = null;
async function verificaVersiunea(){
  try{
    const html = await (await fetch('./index.html', {cache:'no-store'})).text();
    const m = html.match(/assets\/app\.js\?v=([0-9A-Za-z._-]+)/);
    if(m && m[1] !== VERSIUNE){ versiuneNoua = m[1]; render(); }
  }catch{}
}
function reincarca(){
  // reload-ul normal revalideaza pagina; daca browserul tot o tine, adresa cu parametru nou o forteaza
  try{ store.draft.save(snapshot(true)); }catch{}
  location.replace(location.pathname + '?r=' + Date.now());
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
  if(location.search.includes('r=')) history.replaceState(null, '', location.pathname);
  { const sal = store.luni.get(cheiaLunii()); if(sal && JSON.stringify(sal.leaves)===JSON.stringify(state.leaves)) amprentaSalvata = amprenta(); }
  marcheazaModificari();
  actualizeazaFolderInfo();
  verificaVersiunea();
  setInterval(verificaVersiunea, 10*60*1000);
  document.addEventListener('visibilitychange', ()=>{ if(!document.hidden) verificaVersiunea(); });
}
init();
