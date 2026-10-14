// Calculul machetei: agregare, modul delta, luna precedenta, col. 7, repartizarea col. 4. Ruleaza: npm test
import test from 'node:test';
import assert from 'node:assert/strict';
import { LEAVES, blank, dinGol, aggregate, recompute, dinDinCumulat, lunaAnDinEticheta, lunaDinText, starePrecedenta,
  lunaUrmatoare, bazaInceputAn, col7Inconsecvent, repartizeazaCol4, curataStare, aceleasiFrunze, faraCumulat } from '../assets/calc.js';

const stareGoala = (mode='cumulat') => {
  const leaves = {}, din = {};
  LEAVES.forEach(r=>{ leaves[r] = blank(); din[r] = dinGol(); });
  return { meta:{ an:2026, luna:8, unitate:'X' }, mode, linkEG:true, leaves, din };
};
const prevCu = (rows, an=2026, luna=7) => {
  const leaves = {}; LEAVES.forEach(r=>{ leaves[r] = { ...blank(), ...(rows[r]||{}) }; });
  return { leaves, an, luna, label:'iulie 2026', totals: aggregate(leaves)[8] };
};

test('agregarea aduna copiii pana la randul 8, cu 3 zecimale', ()=>{
  const s = stareGoala();
  s.leaves[14].h = 1.5; s.leaves[15].h = 2.25; s.leaves[128].h = 0.1; s.leaves[121] && (s.leaves[121].h = 0);
  const v = aggregate(s.leaves);
  assert.equal(v[13].h, 3.75);
  assert.equal(v[8].h, 3.85);
});

test('delta: cumulat = luna precedenta + luna; col. 4 si Δ col. 8 urmeaza automat', ()=>{
  const p = prevCu({ 14:{ e:500, f:500, g:500, h:500, i:100, j:520 } });
  const s = stareGoala('delta'); s.din[14].i = 100;
  const v = recompute(s, p);
  assert.deepEqual({ ...s.leaves[14] }, { e:600, f:600, g:600, h:600, i:100, j:620 });
  assert.equal(v[8].h, 600);
});

test('comutarea delta → cumulat → delta nu fixeaza celulele automate (regresie)', ()=>{
  const p = prevCu({ 14:{ e:500, f:500, g:500, h:500, i:100, j:520 } });
  const s = stareGoala('delta'); s.din[14].i = 100; recompute(s, p);
  s.mode = 'cumulat'; recompute(s, p);
  assert.deepEqual(dinDinCumulat(s, p), []);
  assert.deepEqual(s.din[14], { i:100, f:null, dj:null });
  s.din[14].i = 150; recompute(s, p);
  assert.equal(s.leaves[14].j, 670, 'col. 8 urmeaza col. 7 si dupa comutare');
  assert.equal(s.leaves[14].f, 650, 'col. 4 urmeaza col. 3 si dupa comutare');
});

test('corecturile facute in cumulat se pastreaza la intoarcerea pe delta', ()=>{
  const p = prevCu({ 14:{ e:500, f:500, g:500, h:500, i:100, j:520 } });
  const s = stareGoala('delta'); s.din[14] = { i:100, f:412.05, dj:5 }; recompute(s, p);
  s.mode = 'cumulat'; recompute(s, p);
  s.leaves[14].i = 120; s.leaves[14].h = 620; recompute(s, p);
  assert.deepEqual(dinDinCumulat(s, p), []);
  recompute(s, p);
  assert.equal(s.leaves[14].i, 120);
  assert.equal(s.leaves[14].h, 620);
  assert.equal(s.din[14].f, 412.05, 'col. 4 tastata ramane tastata');
  assert.equal(s.din[14].dj, 5, 'Δ tastata ramane tastata');
});

test('col. 6 si col. 7 nepotrivite in cumulat sunt semnalate la trecerea pe delta', ()=>{
  const p = prevCu({ 14:{ h:500, j:500 } });
  const s = stareGoala(); s.leaves[14] = { e:630, f:630, g:630, h:630, i:120, j:630 };
  assert.deepEqual(dinDinCumulat(s, p), [14]);
  assert.equal(s.din[14].i, 120);
});

test('formular gol: trecerea pe delta lasa totul automat, pastreaza col. 7 tastata', ()=>{
  const p = prevCu({ 14:{ e:500, f:500, g:500, h:500, j:520 } });
  const s = stareGoala(); s.leaves[15].i = 7;
  assert.ok(faraCumulat(s.leaves));
  dinDinCumulat(s, p);
  assert.deepEqual(s.din[14], { i:0, f:null, dj:null });
  assert.deepEqual(s.din[15], { i:7, f:null, dj:null });
});

test('eticheta lunii: nume, AAAA-LL, inceput de an; „mai" doar ca luna', ()=>{
  assert.deepEqual(lunaAnDinEticheta('iulie 2026'), { an:2026, luna:7 });
  assert.deepEqual(lunaAnDinEticheta('2026-07.json'), { an:2026, luna:7 });
  assert.deepEqual(lunaAnDinEticheta('luna AUGUST 2026'), { an:2026, luna:8 });
  assert.deepEqual(lunaAnDinEticheta('început de an 2027'), { an:2027, luna:0 });
  assert.equal(lunaDinText('Macheta MAI 2026'), 5);
  assert.equal(lunaDinText('str. Maiorescu'), null);
  assert.equal(lunaAnDinEticheta('fisier.xlsx'), null);
});

test('luna precedenta: luna dinainte; in ianuarie, inceput de an', ()=>{
  const aug = { an:2026, luna:8 }, ian = { an:2027, luna:1 }, feb = { an:2027, luna:2 };
  assert.equal(starePrecedenta(aug, { an:2026, luna:7 }), 'ok');
  assert.equal(starePrecedenta(aug, { an:2026, luna:6 }), 'alta');
  assert.equal(starePrecedenta(ian, { an:2026, luna:12 }), 'ianuarie');
  assert.equal(starePrecedenta(ian, bazaInceputAn(2027)), 'ok');
  assert.equal(starePrecedenta(feb, bazaInceputAn(2027)), 'alta');
  assert.equal(starePrecedenta(aug, { an:null, luna:null }), 'necunoscuta');
  assert.deepEqual(lunaUrmatoare(2026, 12), { an:2027, luna:1 });
});

test('ianuarie pe baza „inceput de an": cumulatele sunt chiar luna', ()=>{
  const s = stareGoala('delta'); s.meta = { an:2027, luna:1 }; s.din[14].i = 80;
  const v = recompute(s, bazaInceputAn(2027));
  assert.equal(v[14].h, 80); assert.equal(v[14].j, 80); assert.equal(v[14].e, 80);
});

test('col. 7 = col. 6 − col. 6 din luna precedenta (si fata de zero in ianuarie)', ()=>{
  const p = prevCu({ 14:{ h:500 }, 15:{ h:10 } });
  const s = stareGoala(); s.leaves[14].h = 600; s.leaves[14].i = 100; s.leaves[15].h = 12; s.leaves[15].i = 2.001;
  assert.deepEqual(col7Inconsecvent(s.leaves, p.leaves), [15]);
  s.leaves[15].i = 2; assert.deepEqual(col7Inconsecvent(s.leaves, p.leaves), []);
  assert.deepEqual(col7Inconsecvent(s.leaves, null), [14, 15], 'fara luna precedenta (ianuarie) col. 7 = col. 6');
});

test('repartizarea col. 4 atinge exact tinta, rotunjind in sus randurile mari', ()=>{
  const p = prevCu({});
  const s = stareGoala('delta');
  Object.assign(s.din[14], { i:410.256 }); Object.assign(s.din[33], { i:27.43 }); Object.assign(s.din[57], { i:98.611 });
  Object.assign(s.din[62], { i:7.904 }); Object.assign(s.din[44], { i:5, f:5.5 });          // r.44 tastat: nu se atinge
  let v = recompute(s, p);
  const tinta = Math.round(v[8].f*1000 + 1874)/1000;
  const rez = repartizeazaCol4(s, v, tinta);
  assert.ok(rez);
  assert.ok(!('44' in rez.schimbari));
  rez.rotunjite.forEach(r=>assert.ok(Number.isInteger(rez.schimbari[r]) || +r===rez.restPe, `r.${r} rotunjit la mii intregi`));
  Object.entries(rez.schimbari).forEach(([r, x])=>{ s.din[r].f = x; });
  v = recompute(s, p);
  assert.equal(Math.round(v[8].f*1000), Math.round(tinta*1000));
  LEAVES.forEach(r=>assert.ok(v[r].f >= v[r].h, `r.${r}: col. 4 ≥ col. 6`));
  assert.equal(repartizeazaCol4(s, v, tinta), null, 'tinta atinsa: nimic de repartizat');
});

test('o stare deteriorata se aduce la tipuri valide', ()=>{
  const s = stareGoala(); s.meta = { luna:13, an:'abc', unitate:null }; s.mode = 'x';
  s.leaves[14].h = '1,5'; s.din[14] = { i:'2', f:'', dj:undefined };
  curataStare(s);
  assert.ok(s.meta.luna>=1 && s.meta.luna<=12);
  assert.ok(s.meta.an>=2000);
  assert.equal(s.mode, 'cumulat');
  assert.equal(s.leaves[14].h, 1.5);
  assert.deepEqual(s.din[14], { i:2, f:null, dj:null });
});

test('aceleasiFrunze compara la leu', ()=>{
  const a = prevCu({ 14:{ h:1.0001 } }).leaves, b = prevCu({ 14:{ h:1 } }).leaves, c = prevCu({ 14:{ h:1.002 } }).leaves;
  assert.ok(aceleasiFrunze(a, b));
  assert.ok(!aceleasiFrunze(b, c));
});
