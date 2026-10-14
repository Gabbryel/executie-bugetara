import { ROWS } from './rows.js?v=20261014-132809';
import { parseNum, r3, fmt, fmtLei } from './num.js?v=20261014-132809';
import { loadTemplate, buildXlsx, readMacheta, readBC39 } from './xlsx.js?v=20261014-132809';
import * as store from './store.js?v=20261014-132809';
import { KEYS, LUNI, LEAVES, blank, dinGol, aggregate, recompute as calculeaza, dinDinCumulat, lunaDinText, anDinText,
  lunaAnDinEticheta, numeLunaAn, lunaUrmatoare, bazaInceputAn, starePrecedenta, precedentaAsteptata, col7Inconsecvent,
  faraCumulat as faraCumulatIn, aceleasiFrunze, repartizeazaCol4, curataStare } from './calc.js?v=20261014-132809';

const byRow = Object.fromEntries(ROWS.map(r=>[r.r, r]));
const cuvCheie = t => String(t||'').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'')
  .replace(/[^A-Z0-9 ]/g,' ').split(/\s+/).filter(w=>w && !['SC','SRL','SA','S','C'].includes(w));
const aceeasiUnitate = (a,b) => { const A=cuvCheie(a), B=cuvCheie(b); return A.length && B.length && (A.every(w=>B.includes(w)) || B.every(w=>A.includes(w))); };
// cautarea ignora diacriticele si majusculele si gaseste si dupa cod (sub-capitol, titlu, alineat)
const fold = s => String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
const CAUTA = Object.fromEntries(ROWS.map(r=>[r.r, fold([r.lbl, r.sub, r.tit, r.alin].join(' '))]));
const $ = s => document.querySelector(s);
const escHtml = s => String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
const COL_NAMES = ['Credite bugetare aprobate (col. 3)', 'Credite de angajament (col. 4)', 'Credite trimestriale cumulate (col. 5)', 'Plăți nete de casă cumulat (col. 6)', 'Luna curentă (col. 7)', 'Cheltuieli efective (col. 8)'];

let TPL = null;            // {buf, check}
let state = newState();
let prev = null;           // {label, leaves, an, luna, totals} — luna 0 = inceput de an (ianuarie porneste de la zero)
let bc39ref = null;        // {bugetar, angajament, sursa, coloane, luna, an, denumire, detalii}
let importNote = null;     // {title, det} — inconsecvente gasite in ultima macheta importata
let cheieDeschisa = null;  // luna din arhiva (AAAA-LL) din care provine formularul; salvarea peste alta luna existenta intreaba

function newState(){
  const d = new Date();
  const leaves = {}; const din = {};
  LEAVES.forEach(r=>{ leaves[r]=blank(); din[r]=dinGol(); });
  return {
    meta:{ an:d.getFullYear(), luna:d.getMonth()+1, unitate:'SC SPINAL CARE DOBRECI SRL' },
    mode:'cumulat', linkEG:true, leaves, din, prevLabel:null
  };
}

/* ---------- calcul (logica e in calc.js) ---------- */
const recompute = () => calculeaza(state, prev);
/** Formularul nu are inca date cumulate (col. 3/6/8 goale pe toate randurile)? */
const faraCumulat = () => faraCumulatIn(state.leaves);
const cap = s => s.charAt(0).toUpperCase() + s.slice(1);
const TEXT_DELTA = 'Col. 4 urmează col. 3 și Δ col. 8 urmează col. 7 — tastezi doar rândurile care diferă.';

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
      const aria = `aria-label="rând ${row.r} ${escHtml(row.lbl)} — ${COL_NAMES[idx]}"`;
      return `<td class="num"><input type="text" inputmode="decimal" ${aria} data-r="${row.r}" data-k="${k}" data-src="${src}" data-col="${idx}" value="0,000"></td>`;
    }).join('');
    const dcell = isSum
      ? `<td class="num calc dcol">—</td>`
      : `<td class="num dcol"><input type="text" inputmode="decimal" aria-label="rând ${row.r} ${escHtml(row.lbl)} — Δ cheltuieli efective luna" data-r="${row.r}" data-k="dj" data-src="din.dj" data-col="6" value="0,000"></td>`;
    // in macheta, unele randuri au textul indicatorului in coloana de cod (ex. r.110): il mutam in coloana de denumire
    const codes = [row.sub, row.tit, row.alin].map(c=>String(c||''));
    const lung = codes.filter(c=>c.length>8);
    const lbl = [row.lbl, ...lung].filter(Boolean).join(' ');
    return `<tr data-row="${row.r}" class="${isSum?'sum l'+Math.min(row.lvl,3):''}">
      ${codes.map(c=>`<td class="code">${c.length>8?'':escHtml(c)}</td>`).join('')}
      <td class="lbl" style="padding-left:${pad}px"><span class="rn">${row.r}</span> ${escHtml(lbl)}</td>
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
    inp.addEventListener('paste', onPaste);
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
/** Scrie textul tastat sau lipit in starea celulei. Col. 4 sau Δ golita revine la automat (urmeaza col. 3, respectiv col. 7). */
function scrieCelula(inp, text){
  const r = +inp.dataset.r, k = inp.dataset.k, src = inp.dataset.src;
  if((src==='din.f' || src==='din.dj') && String(text).trim()===''){ state.din[r][k] = null; return; }
  const v = r3(parseNum(text));
  if(src.startsWith('din')) state.din[r][k] = v; else state.leaves[r][k] = v;
}
function onInput(e){
  scrieCelula(e.target, e.target.value);
  render();
  scheduleSave();
}
/** Lipire din Excel a mai multor celule: randul k al textului merge pe randul machetei de dedesubt (r + k),
 *  coloana j pe coloana de valori din dreapta; randurile de total si celulele calculate se sar. */
function onPaste(e){
  const text = (e.clipboardData && e.clipboardData.getData('text')) || '';
  if(!/[\t\n\r]/.test(text.replace(/[\r\n]+$/,''))) return;     // o singura valoare: lipirea obisnuita
  e.preventDefault();
  const linii = text.replace(/\r/g,'').replace(/\n+$/,'').split('\n').map(l=>l.split('\t'));
  const r0 = +e.target.dataset.r, c0 = +e.target.dataset.col;
  let n = 0, ultim = r0, sarite = 0;
  linii.forEach((celule, di)=>{
    const r = r0 + di; if(!byRow[r]) return;
    celule.forEach((t, dj)=>{
      const inp = $(`#tbl input[data-r="${r}"][data-col="${c0+dj}"]`);
      if(!inp){ if(String(t).trim()!=='') sarite++; return; }
      scrieCelula(inp, t); n++; ultim = r;
    });
  });
  render(); scheduleSave();
  e.target.value = fmt(readInput(e.target));        // render() nu rescrie celula in care e cursorul
  toast(`Am lipit ${n} ${n===1?'valoare':'valori'} pe rândurile ${r0}–${ultim}.` + (sarite ? ` ${sarite} ${sarite===1?'valoare a căzut':'valori au căzut'} pe rânduri de total sau celule calculate și nu s-au scris.` : ''));
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
    // col. 4 sub tinta: diferenta se acopera rotunjind in sus cateva randuri, ca la contabil — un clic
    const rep = k==='f' && d>0 ? ` <button type="button" class="lnk" data-act="repartizeaza" title="Rotunjește în sus la mii întregi câteva rânduri din col. 4, ca să atingă ținta">repartizează</button>` : '';
    th.innerHTML = `<b>${fmt(tinta)}</b><small>${d===0 ? '✓ atins' : d>0 ? 'rămân '+fmt(d/1000) : 'în plus '+fmt(-d/1000)}${rep}</small>`;
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
  const q = fold(($('#filtru').value||'').trim());
  const hz = $('#hideZero').checked;
  document.querySelectorAll('#tbl tbody tr').forEach(tr=>{
    const rn = +tr.dataset.row, row = byRow[rn], v = vals[rn] || blank();
    const gol = KEYS.every(k=>Math.abs(v[k])<1e-9);
    tr.classList.toggle('zero', gol);
    // randurile de sinteza raman vizibile ca reper, cu exceptia celor complet goale cand se ascund zerourile
    const potrivit = !q || CAUTA[rn].includes(q) || String(rn)===q;
    let vizibil = row.ch ? true : potrivit;
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

  // 8. luna precedenta = luna dinainte; in ianuarie cumulatele pornesc de la zero, deci decembrie nu se aduna.
  // Prima in lista: o luna precedenta gresita strica toate cumulatele, deci e cauza celorlalte diferente.
  if(prev){
    const st = starePrecedenta(state.meta, prev);
    const a = precedentaAsteptata(state.meta.an, state.meta.luna);
    const asteptata = numeLunaAn(a.an, a.luna), form = numeLunaAn(state.meta.an, state.meta.luna);
    if(st==='ok') out.push(chk(true,'Luna precedentă', `${prev.label} este baza pentru ${form}.`, false, 'luna precedentă: '+prev.label));
    else if(st==='ianuarie') out.push({ ...chk(false,'Luna precedentă',
      `Formularul este pe ${form}: în ianuarie cumulatele pornesc de la zero, dar luna precedentă încărcată este ${prev.label} și s-ar aduna. Apasă „Pornește de la zero".`), act:'zero' });
    else if(st==='alta') out.push(chk(false,'Luna precedentă',
      `Luna precedentă este ${prev.label}, dar formularul este pe ${form}: cumulatele trebuie să pornească de la ${asteptata}. Încarcă ${asteptata} ca lună precedentă (meniul ⋯ → Luni salvate) sau schimbă luna formularului.`));
    else out.push(chk(null,'Luna precedentă', `Nu știu ce lună este „${prev.label}". Verifică să fie ${asteptata}.`));
  }

  // 1. corelatie BC39
  if(bc39ref && (bc39ref.bugetar!=null || bc39ref.angajament!=null)){
    const db = bc39ref.bugetar!=null ? Math.round(t.e*1000)-Math.round(bc39ref.bugetar) : 0;
    const da = bc39ref.angajament!=null ? Math.round(t.f*1000)-Math.round(bc39ref.angajament) : 0;
    const lunaBc = bc39ref.lunaNr ? `${LUNI[bc39ref.lunaNr-1]} ${bc39ref.an||''}`.trim() : null;
    const lunaForm = `${LUNI[state.meta.luna-1]} ${state.meta.an}`;
    const altaLuna = bc39AltaLuna();
    const peste = (ref) => (prev && prev.luna!==0 && ref!=null) ? ` — cu ${fmt(ref/1000 - prev.totals.h)} peste ${prev.label}` : '';
    const tinta = (col, cur, ref, d, extra, sfat) => ref==null ? '' :
      `${col} este ${fmt(cur)}, trebuie ${fmt(ref/1000)} (${fmtLei(ref)} lei${extra||''})` + (d===0 ? ' ✓' : (d<0 ? ` — lipsesc ${fmt(-d/1000)}` : ` — în plus ${fmt(d/1000)}`) + (sfat||''));
    // in delta, col. 4 porneste de la col. 3: diferenta fata de BC39 se acopera rotunjind cateva randuri
    const sfat4 = state.mode==='delta' ? ': col. 4 urmează col. 3; apasă „repartizează" în rândul țintă sau tastează în col. 4 pe rândurile de rotunjit' : '';
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

  // 9. in cumulat: col. 7 = col. 6 − col. 6 din luna precedenta (in ianuarie, fata de zero).
  // Verificat pe machetele reale iunie → iulie → august: se potriveste exact pe toate randurile.
  if(state.mode==='cumulat' && (prev || state.meta.luna===1) && !faraCumulat()){
    const bad = col7Inconsecvent(state.leaves, prev ? prev.leaves : null);
    const fata = prev ? prev.label : 'începutul anului';
    out.push(chk(bad.length===0, 'Luna curentă (col. 7) = cumulat − luna precedentă',
      bad.length===0 ? `Pe toate rândurile col. 7 = col. 6 − col. 6 din ${fata}.`
        : `${nr(bad.length)}: col. 7 nu este col. 6 minus ${fata}: ` +
          bad.slice(0,8).map(r=>{ const p = prev ? (prev.leaves[r]||blank()).h : 0; return `r.${r} (col. 6 − ${fata} = ${fmt(r3(vals[r].h - p))}, col. 7 = ${fmt(vals[r].i)})`; }).join(', ') +
          (bad.length>8?'…':'') + '. Corectează col. 7 sau col. 6 pe aceste rânduri.',
      false, 'col. 7 = Δ col. 6'));
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
  const reper = prev && prev.luna!==0 ? ` În ${prev.label} col. 8 depășea col. 6 cu ${fmt(r3(prev.totals.j - prev.totals.h))}.` : '';
  const deltaAuto = state.mode==='delta' && LEAVES.every(r=>(state.din[r]||{}).dj==null);
  if(state.mode==='cumulat' && Math.abs(t.j)<1e-9 && Math.abs(t.h)>1e-9){
    out.push(chk(false,'Cheltuieli efective (col. 8) necompletate','Col. 8 este goală. Completează cheltuielile efective cumulate pe rânduri.' + reper, true));
  } else if(deltaAuto){
    // nicio Δ tastata: col. 8 = luna precedenta + col. 7 pe toate randurile; decalajul fata de plati ramane cel din luna precedenta
    const cuFacturi = LEAVES.filter(r=>{ const p=prev.leaves[r]||blank(); return (p.j - p.h) > 0.0005; });
    out.push(chk(null,'Cheltuieli efective (col. 8)',
      `Urmează plățile lunii (col. 7) pe toate rândurile: col. 8 = ${fmt(t.j)}` + (prev.luna!==0 ? `, cu ${fmt(gap)} peste plăți, ca în ${prev.label}. ` : ', egal cu plățile. ') +
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
  gd.innerHTML = `<span class="ic"></span><span class="tx">${escHtml(g.text)}</span>` + (g.act ? `<button class="btn pri" id="guideAct" data-act="${g.act}">${ACTIUNI[g.act]}</button>` : '') +
    (g.more ? `<button class="lnk" id="guideMore">Toate verificările</button>` : '') + (g.reload ? `<button class="btn pri" id="guideReload">Reîncarcă aplicația</button>` : '');
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
const ACTIUNI = { zero:'Pornește de la zero' };
const chkHtml = c => `<div class="chk ${c.cls}"><span class="ic">${c.ic}</span><span><b>${escHtml(c.title)}</b><span class="det">${escHtml(c.det)}</span></span></div>`;

/* ---------- ghidare: o singura linie care spune ce urmeaza ---------- */
function guidance(vals, checks){
  const t = vals[8];
  const gol = KEYS.every(k=>Math.abs(t[k])<1e-9);
  if(versiuneNoua) return { cls:'warn', text:`Este publicată o versiune nouă a aplicației (${versiuneNoua}); rulezi ${VERSIUNE}. Ciorna se păstrează.`, reload:true };
  if(TPL && !TPL.check.ok) return { cls:'err', text:'Șablonul este defect: '+TPL.check.reason+'. Exportul ar produce o machetă greșită.' };
  if(state.meta.luna===1 && !prev && state.mode==='cumulat' && faraCumulat())
    return { cls:'info', act:'zero', text:`Ianuarie ${state.meta.an}: cumulatele pornesc de la zero, nu de la decembrie. Apasă „Pornește de la zero" și completează doar plățile lunii (col. 7), ca în celelalte luni.` };
  if(gol && !prev && !bc39ref) return { cls:'info', text:'Începe prin a importa luna precedentă sau macheta lunii curente din meniul ⋯, apoi fișa BC39 pentru verificare. Sau completează direct rândurile.' };
  const err = checks.find(c=>c.cls==='err');
  if(err) return { cls:'err', text:err.title+': '+err.det, more:true, act:err.act };
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
    prev: (withPrev && prev) ? { label:prev.label, an:prev.an, luna:prev.luna, leaves:prev.leaves } : undefined
  };
}
function restore(data){
  if(!data||typeof data!=='object'||!data.leaves||typeof data.leaves!=='object') throw new Error('Fișier de stare invalid.');
  state = newState();
  Object.assign(state.meta, data.meta||{});
  state.mode = data.mode||'cumulat'; state.linkEG = data.linkEG!==false;
  LEAVES.forEach(r=>{ if(data.leaves[r]) Object.assign(state.leaves[r], data.leaves[r]);
                      if(data.din&&data.din[r]) Object.assign(state.din[r], data.din[r]); });
  curataStare(state);                          // un fisier deteriorat nu trebuie sa blocheze aplicatia
  // luna precedenta vine doar din fisier: cea incarcata inainte apartine altei luni
  clearPrev();
  if(data.prev && data.prev.leaves) setPrev(data.prev.leaves, data.prev.label||'lună precedentă', data.prev.an, data.prev.luna);
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
  cheieDeschisa = null;
  // Modul delta se sprijina pe luna precedenta: fara ea, valorile cumulate ar fi recalculate gresit.
  let nota = null;
  if(state.mode==='delta' && !prev){
    state.mode='cumulat';
    nota = 'Luna precedentă lipsește din fișier — am trecut pe modul „cumulat", valorile rămân neschimbate.';
  }
  syncMetaInputs(); buildTable(); render(); scheduleSave();
  if(prev && state.mode==='cumulat' && faraCumulat()) treciPeDelta(true);
  return nota;
}
/** Restaureaza si anunta: avertismentul din restore() are prioritate fata de mesajul obisnuit. */
function deschide(data, mesaj){ const nota = restore(data); toast(nota || mesaj, !!nota); }
/** Trece pe modul delta fata de luna precedenta. Valorile din tabel sunt sursa (calc.dinDinCumulat): ce s-a tastat sau
 *  corectat in cumulat se pastreaza, iar celulele egale cu valoarea automata raman automate. */
function treciPeDelta(anunta){
  if(!prev) return;
  const inc = dinDinCumulat(state, prev);
  syncMetaInputs(); buildTable(); render(); scheduleSave();
  if(inc.length) toast(`Mod „delta" față de ${prev.label}. Pe ${inc.length===1?'rândul':'rândurile'} ${inc.slice(0,10).join(', ')}${inc.length>10?'…':''} col. 6 nu era ${prev.label} + col. 7: am păstrat col. 7 și am recalculat col. 6.`, true);
  else if(anunta) toast(`Mod „delta" față de ${prev.label}: completezi plățile lunii (col. 7) și delta cheltuielilor efective; cumulatele pornesc de la ${prev.label}. ${TEXT_DELTA}`);
}
/**
 * Incarca luna precedenta (an/luna pot lipsi; se deduc din eticheta).
 * Pe un formular gol, fara luna fixata de BC39, formularul trece pe luna de dupa cea precedenta.
 * In ianuarie, o luna din anul trecut nu se aduna: baza devine „inceput de an" (cumulatele pornesc de la zero).
 * Daca formularul nu are inca date cumulate, trece pe delta.
 */
function incarcaPrev(leaves, an, luna, eticheta){
  if(an==null || luna==null){ const la = lunaAnDinEticheta(eticheta); if(la){ an = la.an; luna = la.luna; } }
  const note = [];
  if(an && luna && faraCumulat() && !(bc39ref && bc39ref.lunaNr)){
    const u = lunaUrmatoare(+an, +luna);
    if(u.an!==state.meta.an || u.luna!==state.meta.luna){
      state.meta.an = u.an; state.meta.luna = u.luna; syncMetaInputs(); amprentaSalvata = null;
      note.push(`formularul a trecut pe ${numeLunaAn(u.an, u.luna)}`);
    }
  }
  let msg;
  if(state.meta.luna===1 && an && +an < state.meta.an){
    const z = bazaInceputAn(state.meta.an);
    setPrev(z.leaves, z.label, z.an, 0);
    msg = `Ianuarie ${state.meta.an}: cumulatele pornesc de la zero, deci ${numeLunaAn(+an, +luna)} nu se adună` + (note.length ? ` (${note.join('; ')})` : '') + '. Completezi plățile lunii (col. 7); col. 3, 5, 6 și 8 ies egale cu luna.';
  } else {
    setPrev(leaves, eticheta, an, luna);
    msg = 'Lună precedentă: ' + prev.label + (note.length ? ` — ${note.join('; ')}` : '') + '.';
  }
  scheduleSave();
  if(state.mode==='cumulat' && faraCumulat()){ treciPeDelta(false); msg += (prev.luna===0 ? ' ' : ' Mod „delta": completezi plățile lunii (col. 7). ') + TEXT_DELTA; }
  toast(msg);
}
/** Ianuarie: baza „inceput de an" (zero) in locul lunii precedente; in delta cumulatele devin chiar valorile lunii. */
function pornesteDeLaZero(){
  const z = bazaInceputAn(state.meta.an);
  setPrev(z.leaves, z.label, z.an, 0);
  if(state.mode==='cumulat' && faraCumulat()) treciPeDelta(false); else render();
  scheduleSave(); marcheazaModificari();
  toast(`Ianuarie ${state.meta.an}: cumulatele pornesc de la zero. Completezi plățile lunii (col. 7); col. 3, 5, 6 și 8 ies egale cu luna. ${TEXT_DELTA}`);
}
function setPrev(leaves, label, an, luna){
  const L = {}; LEAVES.forEach(r=>{ L[r] = blank(); KEYS.forEach(k=>{ const v = (leaves[r]||{})[k]; L[r][k] = typeof v==='number' ? (isFinite(v)?v:0) : parseNum(v); }); });
  const la = (an!=null && luna!=null && isFinite(+an) && isFinite(+luna)) ? { an:+an, luna:+luna } : lunaAnDinEticheta(label);
  prev = { leaves:L, an: la ? la.an : null, luna: la ? la.luna : null, label: la ? numeLunaAn(la.an, la.luna) : (label||'lună precedentă'), totals: aggregate(L)[8] };
  state.prevLabel = prev.label;
  $('#prevInfo').textContent = 'Lună precedentă: ' + prev.label;
  $('#prevHdr').textContent = 'față de ' + prev.label; $('#prevHdr').hidden = false;
  $('#modeDelta').disabled = false; $('#modeSeg').hidden = false;
  if($('#tbl').rows.length) render();
}
function clearPrev(){
  prev = null; state.prevLabel = null;
  $('#prevInfo').textContent = 'Nicio lună precedentă încărcată';
  $('#prevHdr').hidden = true; $('#modeSeg').hidden = true; $('#modeDelta').disabled = true;
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
  const nume = `Macheta Executie spital ${luna} ${state.meta.an}.xlsx`;
  download(blob, nume);
  // si in folderul de salvare, langa AAAA-LL.json, daca e ales
  const f = await scrieInFolder(nume, blob);
  toast('Macheta exportată' + (f.nume ? ' și salvată în folderul „'+f.nume+'"' : '') + '.' + (f.eroare ? ' Folder: '+f.eroare : ''), !!f.eroare);
}
/** Scrie un fisier in folderul de salvare (daca e ales si are permisiune). Returneaza {nume} sau {eroare}, sau {} fara folder. */
async function scrieInFolder(name, content){
  if(!store.folder.suportat() || !(await store.folder.get())) return {};
  const p = await store.folder.permisiune(true);
  if(p!=='granted') return { eroare:'folderul de salvare nu are permisiune (alege-l din nou din meniu)' };
  try{ await store.folder.scrie(name, content); return { nume:(await store.folder.get()).name||'ales' }; }
  catch(e){ return { eroare:e.message }; }
}
/** La tiparire, titlul paginii devine numele fisierului: „Salvează ca PDF" il propune ca nume. Folderul il alege browserul (il tine minte pe ultimul). */
function numePdf(){ return `Macheta Executie spital ${LUNI[state.meta.luna-1].toUpperCase()} ${state.meta.an}`; }
function exportJson(){
  // cu luna precedenta: fara ea, fisierul redeschis ar pierde modul delta
  const b = new Blob([JSON.stringify(snapshot(true),null,2)],{type:'application/json'});
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
    // importul ca luna curenta inlocuieste formularul: intai intrebam de modificarile nesalvate
    if(target==='cur' && !(await inainteDeInlocuire('imporți „'+f.name+'"'))) return;
    if(f.name.toLowerCase().endsWith('.json')){
      const data = JSON.parse(await f.text());
      if(target==='prev'){ if(!data || !data.leaves) throw new Error('Fișier de stare invalid.'); incarcaPrev(data.leaves, data.meta && data.meta.an, data.meta && data.meta.luna, f.name); }
      else deschide(data, 'Stare încărcată din JSON.');
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
    if(target==='prev'){
      incarcaPrev(m.values, la, lm, eticheta);
      if(dif.length) toast('Luna precedentă are totaluri inconsecvente în fișier — vezi verificările.', true);
    }
    else {
      // macheta inlocuieste luna: valorile vin toate din fisier, iar Δ-urile vechi nu mai corespund
      const nou = newState();
      LEAVES.forEach(r=>{ state.leaves[r] = { ...nou.leaves[r], ...(m.values[r]||{}) }; });
      state.din = nou.din;
      cheieDeschisa = null;
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
const amprenta = () => JSON.stringify({ m:state.meta, mode:state.mode, l:state.linkEG, leaves:state.leaves, din:state.din, p:state.prevLabel, pt:prev ? prev.totals : null });
let amprentaSalvata = null;
function marcheazaModificari(){
  const sal = store.luni.get(cheiaLunii());
  const dirty = !sal || amprentaSalvata !== amprenta();
  $('#btnSave').classList.toggle('dirty', dirty);
  $('#btnSave').title = dirty ? 'Modificări nesalvate — salvează luna (Ctrl/Cmd+S)' : 'Luna este salvată (' + new Date(sal.salvat||sal.generat).toLocaleString('ro-RO') + ')';
}
/** Ce s-a tastat in formular (o luna precedenta incarcata fara nimic tastat nu conteaza). */
const areDateTastate = () => state.mode==='delta'
  ? LEAVES.some(r=>{ const d=state.din[r]; return d.i || d.f!=null || d.dj!=null; })
  : LEAVES.some(r=>KEYS.some(k=>Math.abs(state.leaves[r][k])>1e-9));
const nesalvat = () => { const sal = store.luni.get(cheiaLunii()); return areDateTastate() && (!sal || amprentaSalvata !== amprenta()); };
/** Inainte ca formularul sa fie inlocuit (alta luna, import, GitHub): modificarile nesalvate nu se pierd fara intrebare. */
async function inainteDeInlocuire(ce){
  if(!nesalvat()) return true;
  const r = await intreaba(`Formularul pentru ${numeLunaAn(state.meta.an, state.meta.luna)} are modificări nesalvate. Înainte să ${ce}:`, [
    { id:'save', text:'Salvează luna și continuă', pri:true }, { id:'skip', text:'Continuă fără să salvez' }, { id:'cancel', text:'Anulează' }]);
  if(r==='save') return await salveazaLuna();
  return r==='skip';
}
/** Intrebare cu butoane, in pagina (nu confirm() al browserului). Intoarce id-ul butonului apasat sau null (Escape). */
function intreaba(text, butoane){
  const dlg = $('#askDlg'), bar = $('#askBtns');
  $('#askText').textContent = text;
  bar.innerHTML = '';
  return new Promise(res=>{
    const gata = v => { dlg.removeEventListener('close', laEscape); if(dlg.open) dlg.close(); res(v); };
    const laEscape = () => gata(null);
    butoane.forEach(b=>{
      const el = document.createElement('button');
      el.type = 'button'; el.className = 'btn' + (b.pri ? ' pri' : ''); el.textContent = b.text;
      el.addEventListener('click', ()=>gata(b.id));
      bar.appendChild(el);
    });
    dlg.addEventListener('close', laEscape);
    dlg.showModal();
    (bar.querySelector('.pri') || bar.firstChild)?.focus();
  });
}
/** Luna salvata: din folder (are prioritate, poate fi modificat si din alta parte), altfel din arhiva browserului. */
async function citesteLuna(k){
  let d = null;
  if(await store.folder.get() && (await store.folder.permisiune(true))==='granted') d = await store.folder.citeste(k+'.json');
  return d || store.luni.get(k);
}
/** Salveaza luna. Intoarce true daca s-a salvat (cel putin in browser). */
async function salveazaLuna(){
  const k = cheiaLunii();
  // peste o luna salvata din care formularul nu provine (ex. dupa schimbarea lunii din antet): intai intrebam
  if(k !== cheieDeschisa){
    const peDisc = (await store.folder.permisiune(false))==='granted' ? await store.folder.citeste(k+'.json') : null;
    const exista = store.luni.get(k) || peDisc;
    if(exista){
      const cand = exista.salvat||exista.generat ? ` (salvată ${new Date(exista.salvat||exista.generat).toLocaleString('ro-RO')})` : '';
      const r = await intreaba(`Există deja o versiune salvată pentru ${numeLunii(k)}${cand}, iar formularul nu a fost deschis din ea. O înlocuiești cu ce e acum în formular?`,
        [{ id:'da', text:'Înlocuiește', pri:true }, { id:'nu', text:'Anulează' }]);
      if(r!=='da'){ toast('Nu am salvat nimic.'); return false; }
    }
  }
  const data = snapshot(true); data.salvat = new Date().toISOString();
  try{ store.luni.set(k, data); }
  catch(e){ toast('Nu am putut salva în browser: '+e.message, true); return false; }
  cheieDeschisa = k;
  amprentaSalvata = amprenta(); marcheazaModificari(); scheduleSave();
  const unde = ['în acest browser'];
  const erori = [];
  // 1. folderul de pe disc (fara token; permisiunea se cere la primul clic)
  const f = await scrieInFolder(k+'.json', JSON.stringify(data, null, 2));
  if(f.nume) unde.push('în folderul „'+f.nume+'"'); else if(f.eroare) erori.push('folder: '+f.eroare);
  // 2. GitHub, doar daca e configurat cu token
  const c = store.cfg.get();
  if(c.owner && c.repo && store.cfg.token()){
    try{ await store.saveMonth(k+'.json', data, `Execuție bugetară ${numeLunii(k)}`); unde.push(`în GitHub (${c.owner}/${c.repo})`); }
    catch(e){ erori.push('GitHub: '+e.message); }
  }
  let msg = `Salvat ${numeLunii(k)} ${unde.join(', ')}.`;
  if(unde.length===1 && store.folder.suportat() && !(await store.folder.get())) msg += ' Alege un folder de salvare din meniul ⋯ ca să ai fișierul și pe disc.';
  if(erori.length) msg += ' Nu s-a putut salva ' + erori.join('; ') + '.';
  toast(msg, erori.length>0);
  return true;
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
  el.innerHTML = `Folder: <b>${escHtml(h.name||'(ales)')}</b>` + (p==='granted' ? '' : ' · <span class="muted">permisiunea se cere la prima salvare</span>') +
    ` · <a href="#" id="folderUita">uită folderul</a>`;
  $('#folderUita').addEventListener('click', async e=>{ e.preventDefault(); await store.folder.uita(); actualizeazaFolderInfo(); toast('Folderul a fost uitat; lunile se salvează doar în browser.'); });
}
async function deschideLuna(k, ca){
  if(ca!=='prev' && !(await inainteDeInlocuire('deschizi '+numeLunii(k)))) return;
  const d = await citesteLuna(k);
  if(!d) return toast('Luna nu există nici în folder, nici în arhiva din browser.', true);
  const [a, l] = k.split('-').map(Number);
  $('#luniDlg').close();
  if(ca==='prev'){ incarcaPrev(d.leaves||{}, a, l, numeLunii(k)); return; }
  deschide(d, 'Deschis pentru modificări: '+numeLunii(k));
  cheieDeschisa = k; amprentaSalvata = amprenta(); marcheazaModificari();
  await verificaBazaActualizata();
}
/**
 * Fiecare luna salvata isi tine copia lunii precedente. Daca luna precedenta a fost salvata din nou intre timp
 * (ex. iulie corectata dupa ce ai inceput august), luna deschisa s-ar calcula pe valorile vechi: propunem actualizarea.
 */
async function verificaBazaActualizata(){
  if(!prev || state.mode!=='delta' || !prev.luna) return;
  const kp = `${prev.an}-${String(prev.luna).padStart(2,'0')}`;
  const sal = await citesteLuna(kp);
  if(!sal || !sal.leaves) return;
  const L = {}; LEAVES.forEach(r=>{ L[r] = { ...blank(), ...(sal.leaves[r]||{}) }; });
  if(aceleasiFrunze(L, prev.leaves)) return;
  const luna = numeLunaAn(state.meta.an, state.meta.luna);
  const r = await intreaba(`${cap(prev.label)} a fost salvată din nou și are alte valori decât cele pe care se sprijină ${luna}. Recalculez ${luna} pe versiunea salvată a lunii ${prev.label}? (plățile lunii rămân aceleași, se schimbă cumulatele)`,
    [{ id:'da', text:'Recalculează', pri:true }, { id:'nu', text:'Păstrează versiunea veche' }]);
  if(r!=='da') return;
  setPrev(sal.leaves, null, prev.an, prev.luna);
  render(); scheduleSave(); marcheazaModificari();
  toast(`${cap(luna)} se calculează acum pe ${prev.label} salvată. Verifică și salvează luna.`);
}
/** Toate lunile din arhiva browserului intr-un singur fisier (copie de siguranta sau mutarea pe alt calculator / alta adresa). */
function exportArhiva(){
  const data = { tip:'execbug-arhiva', versiune:1, generat:new Date().toISOString(), luni:store.luni.all() };
  const n = Object.keys(data.luni).length;
  if(!n) return toast('Arhiva din browser e goală: salvează întâi o lună.', true);
  download(new Blob([JSON.stringify(data,null,2)],{type:'application/json'}), `execbug-arhiva-${new Date().toISOString().slice(0,10)}.json`);
  toast(`Am descărcat ${n} ${n===1?'lună':'luni'} într-un singur fișier.`);
}
/** Readuce lunile dintr-o arhiva descarcata. O luna existenta se inlocuieste doar cu o versiune salvata mai tarziu. */
async function importArhiva(input){
  const f = input.files[0]; if(!f) return;
  input.value = '';
  try{
    const data = JSON.parse(await f.text());
    if(!data || data.tip!=='execbug-arhiva' || !data.luni || typeof data.luni!=='object') throw new Error('Fișierul nu este o arhivă de luni (din „Descarcă toate lunile").');
    let noi = 0, inloc = 0, pastrate = 0;
    for(const [k, d] of Object.entries(data.luni)){
      if(!/^\d{4}-\d{2}$/.test(k) || !d || !d.leaves) continue;
      const ex = store.luni.get(k);
      if(!ex){ store.luni.set(k, d); noi++; }
      else if(String(d.salvat||d.generat||'') > String(ex.salvat||ex.generat||'')){ store.luni.set(k, d); inloc++; }
      else pastrate++;
    }
    const luni = n => n===1 ? '1 lună' : n+' luni';
    toast(`Arhivă importată: ${luni(noi)} ${noi===1?'nouă':'noi'}, ${luni(inloc)} ${inloc===1?'înlocuită':'înlocuite'} cu o versiune salvată mai târziu, ${luni(pastrate)} ${pastrate===1?'păstrată':'păstrate'} (în browser era deja o versiune la fel de nouă).`);
    listeazaLuni(); marcheazaModificari();
  }catch(e){ toast(e.message, true); }
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
    const safeK = escHtml(k);
    const safeNume = escHtml(numeLunii(k));
    return `<div class="luna-i ${k===cur?'cur':''}">
      <div><div class="n">${safeNume}${k===cur?' <span class="muted">(în lucru)</span>':''}${src}</div>
        <div class="d">plăți cumulate ${fmt(t.h||0)} · luna ${fmt(t.i||0)} · salvat ${d.salvat||d.generat ? new Date(d.salvat||d.generat).toLocaleString('ro-RO') : '—'}</div></div>
      <div class="acts">
        <button class="btn" data-act="open" data-k="${safeK}">Deschide</button>
        <button class="btn" data-act="prev" data-k="${safeK}">Ca lună precedentă</button>
        <button class="btn" data-act="json" data-k="${safeK}">JSON</button>
        <button class="btn" data-act="del" data-k="${safeK}">Șterge</button>
      </div></div>`; }).join('');
  el.querySelectorAll('button[data-act]').forEach(b=>b.addEventListener('click', ()=>{
    const k=b.dataset.k;
    if(b.dataset.act==='open') deschideLuna(k);
    else if(b.dataset.act==='prev') deschideLuna(k,'prev');
    else if(b.dataset.act==='json'){ const d=(surse[k].disc||surse[k].browser); download(new Blob([JSON.stringify(d,null,2)],{type:'application/json'}), k+'.json'); }
    else if(b.dataset.act==='del') stergeLuna(k, !!surse[k].disc);
  }));
}

async function stergeLuna(k, peDisc){
  const r = await intreaba(`Ștergi ${numeLunii(k)}${peDisc?' din folder și':''} din arhiva acestui browser? Nu se poate anula.`,
    [{ id:'da', text:'Șterge' }, { id:'nu', text:'Anulează', pri:true }]);
  if(r!=='da') return;
  store.luni.remove(k);
  if(k===cheieDeschisa) cheieDeschisa = null;
  if(peDisc){ try{ await store.folder.sterge(k+'.json'); }catch(e){ toast('Nu am putut șterge fișierul: '+e.message, true); } }
  listeazaLuni(); marcheazaModificari();
}

/* ---------- GitHub ---------- */
async function ghRefresh(){
  const el = $('#ghList');
  el.innerHTML = '<div class="muted">Se încarcă…</div>';
  try{
    const acc = await store.checkAccess();
    const months = await store.listMonths();
    $('#ghStatus').textContent = `${acc.full} · ${acc.private?'privat':'public'} · ${acc.push?'scriere permisă':'doar citire'}`;
    el.innerHTML = months.length ? months.map(m=>{
      const safeN = escHtml(m.name);
      return `<div class="row-inline" style="justify-content:space-between;border-bottom:1px solid var(--line);padding:5px 0">
        <b>${safeN}</b>
        <span class="row-inline">
          <button class="btn" data-gh="cur" data-n="${safeN}">Deschide</button>
          <button class="btn" data-gh="prev" data-n="${safeN}">Ca lună precedentă</button>
        </span></div>`;
    }).join('')
      : '<div class="muted">Niciun fișier în directorul de date încă.</div>';
    el.querySelectorAll('button[data-gh]').forEach(b=>b.addEventListener('click', ()=>ghLoad(b.dataset.n, b.dataset.gh)));
  }catch(e){ el.innerHTML=''; $('#ghStatus').textContent = 'Eroare: '+e.message; }
}
async function ghLoad(name, as){
  try{
    const r = await store.loadMonth(name);
    if(!r) throw new Error('Fișierul nu există.');
    $('#ghDlg').close();
    if(as==='prev'){ incarcaPrev(r.data.leaves||{}, null, null, name); return; }
    if(!(await inainteDeInlocuire('deschizi '+name))) return;
    deschide(r.data, 'Lună deschisă din GitHub.');
  }catch(e){ toast(e.message, true); }
}
async function ghSave(){
  try{
    const name = store.fileName(state.meta.an, state.meta.luna);
    await store.saveMonth(name, snapshot(true), `Execuție bugetară ${LUNI[state.meta.luna-1]} ${state.meta.an}`);
    toast('Salvat în GitHub: data/'+name);
    if($('#ghDlg').open) ghRefresh();
  }catch(e){ toast(e.message, true); }
}

/* ---------- col. 4: repartizarea pana la tinta BC39 ---------- */
function repartizeaza(){
  if(!bc39ref || bc39ref.angajament==null) return;
  const vals = recompute();
  const rez = repartizeazaCol4(state, vals, bc39ref.angajament/1000);
  if(!rez) return toast('Nu am ce repartiza: col. 4 atinge ținta sau nu are rânduri libere (toate sunt tastate).', true);
  Object.entries(rez.schimbari).forEach(([r, v])=>{ if(state.mode==='delta') state.din[r].f = v; else state.leaves[r].f = v; });
  render(); scheduleSave();
  const rot = rez.rotunjite.filter(r=>r!==rez.restPe).map(r=>'r.'+r);
  const parti = [];
  if(rot.length) parti.push(`am rotunjit în sus la mii întregi ${rot.join(', ')}`);
  if(rez.restPe) parti.push(`${rot.length ? 'restul' : 'diferența'} l-am pus pe r.${rez.restPe}`);
  toast(`Col. 4: ${parti.join(', iar ')}. Ținta BC39 e atinsă; poți schimba orice celulă, iar o celulă golită revine la col. 3.`);
}

/* ---------- diverse ---------- */
let saveTimer;
function scheduleSave(){
  clearTimeout(saveTimer);
  saveTimer=setTimeout(()=>{
    const ok = store.draft.save({ ...snapshot(true), cheie:cheieDeschisa });
    if(!ok) toast('Atenție: Cota localStorage a fost depășită. Ciorna curentă nu a putut fi salvată automat.', true);
    marcheazaModificari();
  }, 600);
}
function toast(msg, bad){ const s=$('#status'); s.textContent=msg; s.classList.toggle('bad', !!bad); clearTimeout(toast._t); toast._t=setTimeout(()=>{s.textContent='';},6000); }
function setSide(open){
  const deschis = document.body.classList.toggle('side-open', open);
  if(deschis){ $('#btnSideClose')?.focus(); }
  else { $('#btnSide')?.focus(); }
}
function syncMetaInputs(){
  $('#an').value = state.meta.an; $('#luna').value = state.meta.luna; $('#unitate').value = state.meta.unitate;
  $('#modeCumulat').checked = state.mode==='cumulat'; $('#modeDelta').checked = state.mode==='delta';
  $('#linkEG').checked = state.linkEG;
}

/* ---------- pornire ---------- */
function wire(){
  $('#luna').innerHTML = LUNI.map((l,i)=>`<option value="${i+1}">${l}</option>`).join('');
  ['an','luna','unitate'].forEach(id=>$('#'+id).addEventListener('change', e=>{
    const v = id==='unitate' ? e.target.value : parseInt(e.target.value, 10);
    if(id==='an' && !(v>=2000 && v<=2100)){ e.target.value = state.meta.an; return toast('Anul trebuie să fie între 2000 și 2100.', true); }
    state.meta[id] = v; render(); scheduleSave(); amprentaSalvata=null; marcheazaModificari();
  }));
  $('#modeCumulat').addEventListener('change', ()=>{ state.mode='cumulat'; buildTable(); render(); scheduleSave(); });
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
  const titlu = document.title;
  window.addEventListener('beforeprint', ()=>{ document.title = numePdf(); });
  window.addEventListener('afterprint', ()=>{ document.title = titlu; });
  $('#btnJson').addEventListener('click', exportJson);
  $('#btnNew').addEventListener('click', async ()=>{
    const r = await intreaba('Golești formularul? Lunile salvate rămân în arhivă.', nesalvat()
      ? [{ id:'save', text:'Salvează luna și golește', pri:true }, { id:'da', text:'Golește fără să salvez' }, { id:'nu', text:'Anulează' }]
      : [{ id:'da', text:'Golește', pri:true }, { id:'nu', text:'Anulează' }]);
    if(r==='save' ? !(await salveazaLuna()) : r!=='da') return;
    state = newState(); store.draft.clear(); bc39ref = null; importNote = null; clearPrev(); cheieDeschisa = null; amprentaSalvata = null;
    $('#bcRef').textContent = ''; $('#filtru').value = '';
    syncMetaInputs(); buildTable(); render(); marcheazaModificari();
  });
  $('#fPrev').addEventListener('change', e=>onFile(e.target,'prev'));
  $('#fCur').addEventListener('change', e=>onFile(e.target,'cur'));
  $('#fBc').addEventListener('change', e=>onFile(e.target,'bc39'));
  $('#btnGh').addEventListener('click', ()=>{ const c=store.cfg.get();
    $('#ghOwner').value=c.owner||''; $('#ghRepo').value=c.repo||''; $('#ghBranch').value=c.branch||'main';
    $('#ghDir').value=c.dir||'data'; $('#ghTok').value=store.cfg.token();
    // pe *.github.io toate site-urile contului impart aceeasi origine, deci acelasi localStorage
    $('#ghOrigine').hidden = !/\.github\.io$/i.test(location.hostname);
    $('#ghDlg').showModal(); if(c.owner&&c.repo) ghRefresh(); });
  $('#ghForm').addEventListener('submit', e=>e.preventDefault());   // Enter intr-un camp nu reincarca pagina
  $('#ghSaveCfg').addEventListener('click', ()=>{
    store.cfg.set({owner:$('#ghOwner').value.trim(), repo:$('#ghRepo').value.trim(), branch:$('#ghBranch').value.trim()||'main', dir:$('#ghDir').value.trim()||'data'});
    store.cfg.setToken($('#ghTok').value.trim()); ghRefresh(); });
  $('#ghPush').addEventListener('click', ghSave);
  $('#ghClose').addEventListener('click', ()=>$('#ghDlg').close());
  $('#btnSide').addEventListener('click', ()=>setSide(!document.body.classList.contains('side-open')));
  $('#guide').addEventListener('click', e=>{
    if(e.target.id==='guideMore') setSide(true);
    if(e.target.id==='guideReload') reincarca();
    if(e.target.id==='guideAct' && e.target.dataset.act==='zero') pornesteDeLaZero();
  });
  // butoanele din randul-tinta se refac la fiecare randare: un singur ascultator pe tabel
  $('.tablewrap').addEventListener('click', e=>{ if(e.target.closest('[data-act="repartizeaza"]')) repartizeaza(); });
  $('#btnArhivaExport').addEventListener('click', exportArhiva);
  $('#fArhiva').addEventListener('change', e=>importArhiva(e.target));
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
  try{ store.draft.save({ ...snapshot(true), cheie:cheieDeschisa }); }catch{}
  location.replace(location.pathname + '?r=' + Date.now());
}

async function init(){
  try{ const t=localStorage.getItem('execbug.theme'); if(t) document.documentElement.setAttribute('data-theme',t); }catch{}
  wire(); syncMetaInputs(); buildTable();
  const d = store.draft.load();
  if(d){
    try{ deschide(d, 'Ciornă restaurată din acest browser.'); cheieDeschisa = 'cheie' in d ? d.cheie : cheiaLunii(); }
    catch(e){
      // o ciorna deteriorata nu trebuie sa blocheze pornirea: o punem deoparte si pornim un formular gol
      store.draft.pastreazaDeteriorata(d);
      state = newState(); clearPrev(); bc39ref = null; importNote = null; cheieDeschisa = null;
      syncMetaInputs(); buildTable();
      toast('Ciorna din acest browser nu a putut fi citită; am pus-o deoparte (execbug.draft.deteriorata) și am pornit un formular gol.', true);
    }
  }
  render();
  // cere browserului sa nu stearga datele site-ului cand ramane fara spatiu (Chrome aproba fara intrebare)
  try{ navigator.storage?.persist?.().catch(()=>{}); }catch{}
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
