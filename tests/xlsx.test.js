// Exportul in sablon si citirea machetei / fisei BC39. Ruleaza: npm test
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'fs';
import JSZip from '../tools/jszip-node.mjs';
import { buildXlsx, readMacheta, readBC39, templateIsClean } from '../assets/xlsx.js';
import { LEAVES, blank, aggregate } from '../assets/calc.js';
import { ROWS } from '../assets/rows.js';
import { pregatesteSablon } from '../tools/pregateste-sablon.mjs';

const TPL = fs.readFileSync(new URL('../template/macheta.xlsx', import.meta.url));
const ab = b => b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
const foaie = async buf => (await JSZip.loadAsync(buf)).file('xl/worksheets/sheet1.xml').async('string');
const formule = xml => Object.fromEntries([...xml.matchAll(/<c\b[^>]*?\br="([A-Z]+\d+)"[^>]*?>(<f\b[^>]*\/>|<f\b[^>]*>[^<]*<\/f>)/g)].map(m=>[m[1], m[2]]));

async function exporta(leaves){
  const vals = aggregate(leaves);
  const all = {}; ROWS.forEach(r=>{ all[r.r] = vals[r.r]; });
  const blob = await buildXlsx(ab(TPL), all, { unitate:'UNITATE & TEST <1>', titluLuna:'luna AUGUST 2026' });
  return { vals, buf: Buffer.from(await blob.arrayBuffer()) };
}

test('sablonul trece verificarea de integritate (I117)', async ()=>{
  assert.deepEqual(templateIsClean(await foaie(TPL)), { ok:true });
});

test('export → import: aceleasi valori, formulele neatinse, totalurile scrise langa formule', async ()=>{
  const leaves = {}; let x = 1;
  LEAVES.forEach(r=>{ leaves[r] = blank(); ['e','f','g','h','i','j'].forEach(k=>{ leaves[r][k] = Math.round((x = (x*7919) % 100003) ) / 1000; }); });
  const { vals, buf } = await exporta(leaves);
  const m = await readMacheta(ab(buf));
  LEAVES.forEach(r=>['e','f','g','h','i','j'].forEach(k=>assert.equal(m.values[r][k], leaves[r][k], `r.${r}.${k}`)));
  assert.equal(m.luna, 'luna AUGUST 2026');
  assert.equal(m.unitate, 'UNITATE & TEST <1>');
  const xml = await foaie(buf);
  assert.deepEqual(formule(xml), formule(await foaie(TPL)), 'formulele sunt identice cu sablonul');
  // valoarea calculata (cache) din celulele de formula = agregarea aplicatiei
  for(const row of ROWS.filter(r=>r.ch)){
    for(const [c, k] of [['E','e'], ['F','f'], ['H','h'], ['I','i'], ['J','j']]){
      const cel = xml.match(new RegExp('<c\\b[^>]*\\br="'+c+row.r+'"[^>]*>([\\s\\S]*?)<\\/c>'));
      if(!cel || !/<f[\s>\/]/.test(cel[1])) continue;
      const v = (cel[1].match(/<\/f><v>([^<]*)<\/v>|<f[^>]*\/><v>([^<]*)<\/v>/)||[]).slice(1).find(Boolean);
      assert.equal(+v, vals[row.r][k], `${c}${row.r}`);
    }
  }
});

test('importul citeste valoarea salvata dintr-o celula-frunza cu formula (H54 = G54+I54 in machetele reale)', async ()=>{
  const zip = await JSZip.loadAsync(TPL);
  let xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
  xml = xml.replace(/<c r="H54"([^>]*?)(\/>|>[\s\S]*?<\/c>)/, '<c r="H54"$1><f>G54+I54</f><v>12.5</v></c>'.replace('$1', ' s="55"'));
  zip.file('xl/worksheets/sheet1.xml', xml);
  const m = await readMacheta(ab(await zip.generateAsync({ type:'nodebuffer' })));
  assert.equal(m.values[54].h, 12.5);
  assert.equal(m.statice[54].h, null, 'pentru verificarea consecventei, formula ramane „nestatica"');
});

async function bc39(randuri, { sharedStrings=false }={}){
  const ss = [];
  const cel = (ref, v) => typeof v==='number' ? `<c t="n" r="${ref}"><v>${v}</v></c>`
    : sharedStrings ? `<c r="${ref}" t="s"><v>${ss.push(v)-1}</v></c>` : `<c r="${ref}" t="inlineStr"><is><t>${v}</t></is></c>`;
  const rows = randuri.map((cells, i)=>`<row r="${i+2}">${cells.map((v, j)=>cel(String.fromCharCode(65+j)+(i+2), v)).join('')}</row>`).join('');
  const zip = new JSZip();
  zip.file('xl/worksheets/sheet1.xml', `<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Raport</t></is></c></row>${rows}</sheetData></worksheet>`);
  if(sharedStrings) zip.file('xl/sharedStrings.xml', `<sst>${ss.map(s=>`<si><t>${s}</t></si>`).join('')}</sst>`);
  return ab(await zip.generateAsync({ type:'nodebuffer' }));
}

test('BC39: antet pe alt rand, atribute in alta ordine, coloana Total ca referinta', async ()=>{
  const b = await readBC39(await bc39([
    ['An', 'Luna (cumulat)', 'Denumire', 'Cod', 'Tip credit', 'Total', 'Spital (inclusiv PNS)'],
    [2026, 'Iulie', 'UNITATE TEST SRL', 1, 'Angajament', 1234567, 1234567],
    [2026, 'Iulie', 'UNITATE TEST SRL', 1, 'Bugetar', 1233210, 1233210],
  ]));
  assert.equal(b.bugetar, 1233210); assert.equal(b.angajament, 1234567);
  assert.equal(b.coloanaRef, 'Total'); assert.equal(b.luna, 'Iulie'); assert.equal(String(b.an), '2026');
  assert.equal(b.denumire, 'UNITATE TEST SRL');
});

test('BC39: sume scrise ca text cu separatori de mii', async ()=>{
  const b = await readBC39(await bc39([
    ['Tip credit', 'Spital'],
    ['Bugetar', '3.456.789'],
    ['Angajament', '3.458.001'],
  ], { sharedStrings:true }));
  assert.equal(b.bugetar, 3456789); assert.equal(b.angajament, 3458001);
});

test('pregatirea sablonului e idempotenta pe sablonul din depozit', async ()=>{
  const rez = await pregatesteSablon(TPL);
  assert.deepEqual(rez.jurnal, []);
  assert.equal(await foaie(rez.buf), await foaie(TPL));
});
