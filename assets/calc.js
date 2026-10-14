// Calculul machetei, fara DOM: agregare, modul delta, luna precedenta, repartizarea col. 4.
// Testat in tests/calc.test.js; app.js il foloseste pentru tabel si verificari.
import { ROWS } from './rows.js?v=20261014-132809';
import { r3, parseNum } from './num.js?v=20261014-132809';

export const KEYS = ['e','f','g','h','i','j'];
export const LUNI = ['ianuarie','februarie','martie','aprilie','mai','iunie','iulie','august','septembrie','octombrie','noiembrie','decembrie'];
export const LEAVES = ROWS.filter(r=>!r.ch).map(r=>r.r);
export const blank = () => Object.fromEntries(KEYS.map(k=>[k,0]));
export const dinGol = () => ({ i:0, dj:null, f:null });   // f:null = col. 4 urmeaza col. 3; dj:null = Δ col. 8 urmeaza col. 7
const lei = x => Math.round((x||0)*1000);                 // comparatiile se fac pe intregi (lei)

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

/** Recalculeaza frunzele (in delta: luna precedenta + ce s-a tastat in luna) si intoarce valorile agregate. */
export function recompute(state, prev){
  if(state.mode==='delta'){
    LEAVES.forEach(r=>{
      const p = prev && prev.leaves[r] ? prev.leaves[r] : blank();
      const d = state.din[r] || dinGol();
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

/**
 * Trecerea din cumulat pe delta. Valorile din tabel sunt sursa: plata lunii = col. 7; col. 4 ramane automata cand
 * e egala cu col. 3; Δ col. 8 ramane automata cand e egala cu col. 7. Asa, o comutare delta → cumulat → delta nu
 * schimba nimic, iar corecturile facute in cumulat se pastreaza.
 * Intoarce randurile pe care col. 6 nu era luna precedenta + col. 7 (acolo s-a pastrat col. 7).
 */
export function dinDinCumulat(state, prev){
  const inconsecvente = [];
  LEAVES.forEach(r=>{
    const p = (prev && prev.leaves[r]) || blank(), L = state.leaves[r];
    const i = r3(L.i||0);
    const cumulatTastat = Math.abs(L.h||0)>1e-9 || Math.abs(L.j||0)>1e-9;
    if(cumulatTastat && lei(L.h) !== lei((p.h||0) + i)) inconsecvente.push(r);
    const dj = r3((L.j||0) - (p.j||0));
    state.din[r] = {
      i,
      f: (Math.abs(L.f||0)>1e-9 && Math.abs((L.f||0)-(L.e||0))>1e-9) ? r3(L.f) : null,
      // col. 8 goala = netastata (un cumulat nu scade la zero), deci Δ urmeaza col. 7
      dj: (Math.abs(L.j||0)<1e-9 || lei(dj)===lei(i)) ? null : dj
    };
  });
  state.mode = 'delta';
  return inconsecvente;
}

/* ---------- luna precedenta ---------- */
// luna 0 = „inceput de an": in ianuarie cumulatele pornesc de la zero, nu de la decembrie
export const numeLunaAn = (an, luna) => luna===0 ? `început de an ${an}` : `${LUNI[luna-1]} ${an}`;
export const lunaDinText = t => {
  const s = String(t||'').toLowerCase();
  const i = LUNI.findIndex(l=>new RegExp('(^|[^a-zăâîșşțţ])'+l+'($|[^a-zăâîșşțţ])').test(s));
  return i<0 ? null : i+1;
};
export const anDinText = t => { const m = String(t||'').match(/(20\d\d)/); return m ? +m[1] : null; };
/** {an, luna} dintr-o eticheta: „iulie 2026", „2026-07", „2026-07.json", „început de an 2027". null daca nu se poate. */
export function lunaAnDinEticheta(t){
  const s = String(t||'');
  const m = s.match(/(20\d\d)-(\d{2})/);
  if(m && +m[2]>=1 && +m[2]<=12) return { an:+m[1], luna:+m[2] };
  if(/[iî]nceput de an/i.test(s)){ const a = anDinText(s); return a ? { an:a, luna:0 } : null; }
  const l = lunaDinText(s), a = anDinText(s);
  return l && a ? { an:a, luna:l } : null;
}
/** Luna precedenta asteptata pentru luna formularului; pentru ianuarie: inceputul anului (zero). */
export const precedentaAsteptata = (an, luna) => ({ an, luna: luna-1 });
export const lunaUrmatoare = (an, luna) => luna>=12 ? { an:an+1, luna:1 } : { an, luna:luna+1 };
/** Baza pentru ianuarie: toate cumulatele zero. */
export function bazaInceputAn(an){
  const leaves = {}; LEAVES.forEach(r=>{ leaves[r] = blank(); });
  return { leaves, an, luna:0, label: numeLunaAn(an, 0) };
}
/** 'ok' | 'ianuarie' (decembrie adunat la ianuarie) | 'alta' (nu e luna dinainte) | 'necunoscuta' */
export function starePrecedenta(meta, prev){
  if(!prev) return 'ok';
  if(prev.an==null || prev.luna==null) return 'necunoscuta';
  const a = precedentaAsteptata(+meta.an, +meta.luna);
  if(+prev.an===a.an && +prev.luna===a.luna) return 'ok';
  return a.luna===0 ? 'ianuarie' : 'alta';
}

/** Randurile pe care col. 7 nu e diferenta col. 6 fata de luna precedenta (in ianuarie, fata de zero). */
export function col7Inconsecvent(leaves, prevLeaves){
  return LEAVES.filter(r=>{
    const c = leaves[r]||blank(), p = (prevLeaves && prevLeaves[r]) || blank();
    return lei(c.h) - lei(p.h) !== lei(c.i);
  });
}
/** Formularul nu are inca date cumulate (col. 3/6/8 goale pe toate randurile)? */
export const faraCumulat = leaves => LEAVES.every(r=>{ const L=leaves[r]||blank(); return Math.abs(L.h)<1e-9 && Math.abs(L.j)<1e-9 && Math.abs(L.e)<1e-9; });
/** Aceleasi valori pe toate frunzele, la leu? (luna precedenta salvata vs. copia din luna curenta) */
export const aceleasiFrunze = (a, b) => LEAVES.every(r=>KEYS.every(k=>lei((a[r]||{})[k]) === lei((b[r]||{})[k])));

/**
 * Repartizeaza diferenta pana la tinta BC39 a col. 4 (r.8), cum face contabilul: rotunjeste in sus la mii intregi
 * randurile cu col. 4 automata (= col. 3), incepand cu cele mai mari, cat timp incape; restul il pune pe cel mai mare rand.
 * Intoarce { schimbari:{rand: valoare}, rotunjite:[...], restPe: rand|null } sau null daca nu e nimic de adaugat.
 */
export function repartizeazaCol4(state, vals, tinta){
  let rest = lei(tinta) - lei(vals[8].f);
  if(rest<=0) return null;
  const L = state.leaves;
  const libere = LEAVES.filter(r=>{
    const e = L[r].e;
    if(!(e>0)) return false;
    return state.mode==='delta' ? state.din[r].f==null : Math.abs((L[r].f||0)-e)<1e-9;
  }).sort((a,b)=>L[b].e-L[a].e || a-b);
  if(!libere.length) return null;
  const schimbari = {}, rotunjite = [];
  for(const r of libere){
    const e = lei(L[r].e), sus = Math.ceil(e/1000)*1000 - e;
    if(sus>0 && sus<=rest){ schimbari[r] = (e+sus)/1000; rotunjite.push(r); rest -= sus; }
    if(rest===0) break;
  }
  let restPe = null;
  if(rest>0){ restPe = libere[0]; schimbari[restPe] = (lei(schimbari[restPe] ?? L[restPe].e) + rest)/1000; }
  return { schimbari, rotunjite, restPe };
}

/** Aduce o stare citita din JSON (ciorna, fisier, GitHub) la tipurile asteptate: luna 1–12, an plauzibil, numere. */
export function curataStare(state){
  const azi = new Date();
  const luna = parseInt(state.meta.luna, 10), an = parseInt(state.meta.an, 10);
  state.meta.luna = luna>=1 && luna<=12 ? luna : azi.getMonth()+1;
  state.meta.an = an>=2000 && an<=2100 ? an : azi.getFullYear();
  state.meta.unitate = String(state.meta.unitate ?? '');
  state.mode = state.mode==='delta' ? 'delta' : 'cumulat';
  const nr = v => typeof v==='number' ? (isFinite(v) ? v : 0) : parseNum(v);
  LEAVES.forEach(r=>{
    const L = state.leaves[r]; KEYS.forEach(k=>{ L[k] = nr(L[k]); });
    const d = state.din[r]; d.i = nr(d.i);
    d.f = d.f==null || d.f==='' ? null : nr(d.f);
    d.dj = d.dj==null || d.dj==='' ? null : nr(d.dj);
  });
  return state;
}
