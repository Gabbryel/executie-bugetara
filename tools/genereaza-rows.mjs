// Genereaza assets/rows.js din sablon (foaia BC, randurile 8–128): codurile din A–C, denumirea din D,
// iar arborele de agregare din formulele coloanei E (=E14+E15+... -> copiii randului).
//   node tools/genereaza-rows.mjs                      verifica: rows.js generat = cel din depozit
//   node tools/genereaza-rows.mjs --scrie [sablon.xlsx] rescrie assets/rows.js
import fs from 'fs';
import JSZip from './jszip-node.mjs';

const PRIMUL = 8, ULTIMUL = 128;
// Randuri de sinteza care in sablon au valoare statica in coloana E, nu formula.
// R.38 („Contributii, din care:") e static in macheta CAS, dar e suma randului 39.
const SINTEZE_STATICE = { 38:[39] };

const dezXml = s => String(s).replace(/<[^>]+>/g,'')
  .replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&');

export async function genereazaRows(buf){
  const zip = await JSZip.loadAsync(buf);
  const xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
  const ssf = zip.file('xl/sharedStrings.xml');
  const ss = ssf ? [...(await ssf.async('string')).matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m=>dezXml(m[1])) : [];
  const celula = ref => xml.match(new RegExp('<c\\b([^>]*?\\br="'+ref+'"[^>]*?)(\\/>|>([\\s\\S]*?)<\\/c>)'));
  const text = ref => {
    const m = celula(ref); if(!m || m[2]==='/>') return '';
    const body = m[3]||'';
    const is = body.match(/<is>([\s\S]*?)<\/is>/); if(is) return dezXml(is[1]).trim();
    const v = (body.match(/<v>([^<]*)<\/v>/)||[])[1]; if(v===undefined) return '';
    return (/\bt="s"/.test(m[1]) ? ss[+v]||'' : dezXml(v)).trim();
  };
  const formula = ref => { const m = celula(ref); return m && m[3] ? (m[3].match(/<f[^>]*>([^<]*)<\/f>/)||[])[1] ?? null : null; };

  const rows = [];
  for(let r=PRIMUL; r<=ULTIMUL; r++){
    const f = formula('E'+r);
    const ch = SINTEZE_STATICE[r] || (f ? [...f.matchAll(/E(\d+)/g)].map(m=>+m[1]) : null);
    // denumirea cu spatiile multiple restranse; codurile raman cum sunt in macheta
    rows.push({ r, sub:text('A'+r), tit:text('B'+r), alin:text('C'+r), lbl:text('D'+r).replace(/\s+/g,' '), lvl:0, ch });
  }
  // nivelul = adancimea in arbore; copiii au mereu numar de rand mai mare decat parintele
  const by = Object.fromEntries(rows.map(x=>[x.r, x]));
  rows.forEach(x=>{ (x.ch||[]).forEach(c=>{
    if(!by[c] || c<=x.r) throw new Error(`r.${x.r}: copil invalid ${c}`);
    by[c].lvl = x.lvl + 1;
  }); });
  return rows;
}

/** Textul fisierului rows.js, in formatul existent (un singur rand JSON). */
export const textRows = rows => '// Generat din macheta CAS (foaia BC, randurile 8-128). Nu edita manual.\nexport const ROWS = [' +
  rows.map(o=>'{'+Object.entries(o).map(([k,v])=>JSON.stringify(k)+': '+JSON.stringify(v)).join(',')+'}').join(',') + '];\n';

if(import.meta.url === `file://${process.argv[1]}`){
  const scrie = process.argv.includes('--scrie');
  const sursa = process.argv.slice(2).find(a=>!a.startsWith('--')) || new URL('../template/macheta.xlsx', import.meta.url);
  const tinta = new URL('../assets/rows.js', import.meta.url);
  const nou = textRows(await genereazaRows(fs.readFileSync(sursa)));
  if(scrie){ fs.writeFileSync(tinta, nou); console.log('assets/rows.js rescris.'); }
  else if(nou === fs.readFileSync(tinta, 'utf8')) console.log('rows.js corespunde sablonului.');
  else { console.log('rows.js NU corespunde sablonului — ruleaza cu --scrie dupa ce verifici sablonul.'); process.exit(1); }
}
