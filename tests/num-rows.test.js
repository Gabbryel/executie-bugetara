// Parsarea numerelor si modelul randurilor (rows.js generat din sablon). Ruleaza: npm test
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'fs';
import { parseNum, numSauNull, r3, fmt } from '../assets/num.js';
import { ROWS } from '../assets/rows.js';
import { genereazaRows, textRows } from '../tools/genereaza-rows.mjs';

test('parseNum: format romanesc, anglo-saxon, separatori de mii', ()=>{
  assert.equal(parseNum('2.468,135'), 2468.135);
  assert.equal(parseNum('412,050'), 412.05);
  assert.equal(parseNum('2,468.135'), 2468.135);
  assert.equal(parseNum('3.456.789'), 3456789, 'mai multe puncte = separatori de mii');
  assert.equal(parseNum('2 468,135'), 2468.135);
  assert.equal(parseNum('-5'), -5);
  assert.equal(parseNum('abc'), 0);
  assert.equal(parseNum(''), 0);
});

test('numSauNull: text si gol nu sunt numere', ()=>{
  assert.equal(numSauNull('Bugetar'), null);
  assert.equal(numSauNull(''), null);
  assert.equal(numSauNull('3.456.789'), 3456789);
  assert.equal(numSauNull('3456789'), 3456789);
});

test('r3 rotunjeste simetric la 3 zecimale', ()=>{
  assert.equal(r3(0.1+0.2), 0.3);
  assert.equal(r3(1.0005), 1.001);
  assert.equal(r3(-1.0005), -1.001);
  assert.equal(fmt(2468.135), '2.468,135');
});

test('rows.js = generat din template/macheta.xlsx', async ()=>{
  const rows = await genereazaRows(fs.readFileSync(new URL('../template/macheta.xlsx', import.meta.url)));
  assert.equal(textRows(rows), fs.readFileSync(new URL('../assets/rows.js', import.meta.url), 'utf8'),
    'ruleaza: node tools/genereaza-rows.mjs --scrie (dupa ce verifici sablonul)');
});

test('arborele de agregare: 121 de randuri, fiecare cu un singur parinte, copii sub parinte', ()=>{
  assert.equal(ROWS.length, 121);
  assert.deepEqual(ROWS.map(r=>r.r), Array.from({ length:121 }, (_, i)=>i+8));
  const parinti = {};
  ROWS.forEach(r=>(r.ch||[]).forEach(c=>{ assert.ok(c>r.r, `r.${r.r} → ${c}`); parinti[c] = (parinti[c]||0) + 1; }));
  ROWS.filter(r=>r.r!==8).forEach(r=>assert.equal(parinti[r.r], 1, `r.${r.r} are un singur parinte`));
  assert.deepEqual(ROWS.find(r=>r.r===38).ch, [39], 'r.38 e sinteza lui 39 (static in macheta)');
});
