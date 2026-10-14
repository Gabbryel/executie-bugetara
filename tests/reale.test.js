// Verificari pe machetele reale, tinute in afara depozitului. Fara MACHETE_DIR testele se sar.
//   MACHETE_DIR=~/priv/machete npm test
// Directorul contine „Macheta Executie spital <LUNA> <AN>.xlsx", optional „BC39-<luna>.xlsx" si formularul CAS
// gol din care e facut sablonul („MACHETA Executie pe spital decembrie 2025 ….xlsx").
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'fs';
import path from 'path';
import '../tools/jszip-node.mjs';
import { buildXlsx, readMacheta, readBC39 } from '../assets/xlsx.js';
import { LUNI, LEAVES, KEYS, aggregate, col7Inconsecvent } from '../assets/calc.js';
import { ROWS } from '../assets/rows.js';
import { pregatesteSablon } from '../tools/pregateste-sablon.mjs';

const DIR = process.env.MACHETE_DIR;
const sari = !DIR && 'MACHETE_DIR nesetat';
const ab = b => b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
// Inconsecvente cunoscute in fisierele sursa (CLAUDE.md): aprilie 2026, r.32 col. 4 ≠ suma copiilor
const CUNOSCUTE = { '2026-04': ['f'] };

async function machete(){
  const out = [];
  for(const f of fs.readdirSync(DIR)){
    const m = f.match(/^Macheta Executie spital ([A-ZĂÂÎȘȚ]+) (\d{4})\.xlsx$/i); if(!m) continue;
    const luna = LUNI.indexOf(m[1].toLowerCase()) + 1; if(!luna) continue;
    const buf = fs.readFileSync(path.join(DIR, f));
    out.push({ an:+m[2], luna, cheie:`${m[2]}-${String(luna).padStart(2,'0')}`, buf, m: await readMacheta(ab(buf)) });
  }
  return out.sort((a,b)=>a.cheie.localeCompare(b.cheie));
}

test('agregarea aplicatiei = totalul din fisier (r.8)', { skip:sari }, async ()=>{
  for(const x of await machete()){
    if(!x.m.values[8]) continue;    // fisier exportat fara totaluri salvate (versiunile vechi ale aplicatiei)
    const v = aggregate(x.m.values)[8];
    KEYS.filter(k=>!(CUNOSCUTE[x.cheie]||[]).includes(k)).forEach(k=>
      assert.equal(Math.round(v[k]*1000), Math.round(x.m.values[8][k]*1000), `${x.cheie} r.8 ${k}`));
  }
});

test('col. 7 = col. 6 − col. 6 din luna precedenta, pe luni consecutive', { skip:sari }, async ()=>{
  const lista = await machete();
  for(let i=1; i<lista.length; i++){
    const a = lista[i-1], b = lista[i];
    const consecutive = (b.an*12 + b.luna) - (a.an*12 + a.luna) === 1;
    if(!consecutive) continue;
    assert.deepEqual(col7Inconsecvent(b.m.values, a.m.values), [], `${a.cheie} → ${b.cheie}`);
  }
});

test('BC39 = r.8 col. 3 si col. 4 × 1000', { skip:sari }, async ()=>{
  const lista = await machete();
  for(const f of fs.readdirSync(DIR).filter(f=>/^BC39-[a-z]+\.xlsx$/i.test(f))){
    const b = await readBC39(ab(fs.readFileSync(path.join(DIR, f))));
    const luna = LUNI.indexOf(String(b.luna).toLowerCase()) + 1;
    const x = lista.find(x=>x.luna===luna && x.an===+b.an); if(!x) continue;
    // totalul afisat de Excel (salvat in fisier); pentru aprilie e cel care include inconsecventa de la r.32
    const t = x.m.values[8] || aggregate(x.m.values)[8];
    assert.equal(Math.round(t.e*1000), b.bugetar, `${f} bugetar`);
    assert.equal(Math.round(t.f*1000), b.angajament, `${f} angajament`);
  }
});

test('import → export in sablon → import: aceleasi valori pe toate randurile', { skip:sari }, async ()=>{
  const TPL = fs.readFileSync(new URL('../template/macheta.xlsx', import.meta.url));
  for(const x of await machete()){
    const z = Object.fromEntries(KEYS.map(k=>[k, 0]));   // randurile goale in fisier nu apar in import
    const leaves = {}; LEAVES.forEach(r=>{ leaves[r] = x.m.values[r] || z; });
    const vals = aggregate(leaves), all = {}; ROWS.forEach(r=>{ all[r.r] = vals[r.r]; });
    const m2 = await readMacheta(ab(Buffer.from(await (await buildXlsx(ab(TPL), all, {})).arrayBuffer())));
    LEAVES.forEach(r=>KEYS.forEach(k=>assert.equal(m2.values[r][k], leaves[r][k], `${x.cheie} r.${r}.${k}`)));
  }
});

test('sablonul din depozit = formularul CAS (decembrie 2025) pregatit cu tools/pregateste-sablon.mjs', { skip:sari }, async ()=>{
  const f = fs.readdirSync(DIR).find(f=>/^MACHETA .*decembrie 2025.*\.xlsx$/i.test(f)); if(!f) return;
  const JSZip = globalThis.JSZip;
  const foaie = async buf => (await JSZip.loadAsync(buf)).file('xl/worksheets/sheet1.xml').async('string');
  const rez = await pregatesteSablon(fs.readFileSync(path.join(DIR, f)));
  assert.equal(await foaie(rez.buf), await foaie(fs.readFileSync(new URL('../template/macheta.xlsx', import.meta.url))));
});
