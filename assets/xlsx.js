// Citire/scriere .xlsx prin editare directa a XML-ului din pachet (JSZip).
// Scopul: exportul pastreaza intact formatarea, formulele si setarile de tiparire ale machetei CAS.
import { r3, numSauNull } from './num.js?v=20261014-132140';

const COLS = ['E','F','G','H','I','J'];
const KEYS = ['e','f','g','h','i','j'];

const esc = s => String(s)
  .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
  .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
// textul din XML: entitatile inapoi in caractere (&amp; -> &), fara etichetele interioare (<r>, <rPh>)
const dezXml = s => String(s)
  .replace(/<rPh\b[\s\S]*?<\/rPh>/g,'').replace(/<[^>]+>/g,'')
  .replace(/&#x([0-9a-f]+);/gi,(_,h)=>String.fromCodePoint(parseInt(h,16))).replace(/&#(\d+);/g,(_,d)=>String.fromCodePoint(+d))
  .replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&');

// Localizeaza <c ... r="REF" ...> fie self-closing, fie cu continut, indiferent de ordinea atributelor.
function cellRe(ref){
  return new RegExp('<c\\b([^>]*?\\br="'+ref+'"[^>]*?)(\\/>|>([\\s\\S]*?)<\\/c>)');
}

function hasFormula(xml, ref){
  const m = xml.match(cellRe(ref));
  return !!(m && m[2]!=='/>' && /<f[\s>\/]/.test(m[3]||''));
}

function setValue(xml, ref, num){
  const m = xml.match(cellRe(ref));
  if(!m) return xml;
  const attrs = m[1].replace(/\st="[^"]*"/g,'').replace(/\br="[^"]*"/g,'').trim();
  const repl = '<c r="'+ref+'"'+(attrs?' '+attrs:'')+'><v>'+num+'</v></c>';
  return xml.slice(0,m.index)+repl+xml.slice(m.index+m[0].length);
}

function setText(xml, ref, text){
  const m = xml.match(cellRe(ref));
  if(!m) return xml;
  const attrs = m[1].replace(/\st="[^"]*"/g,'').replace(/\br="[^"]*"/g,'').trim();
  const repl = '<c r="'+ref+'"'+(attrs?' '+attrs:'')+' t="inlineStr"><is><t xml:space="preserve">'+esc(text)+'</t></is></c>';
  return xml.slice(0,m.index)+repl+xml.slice(m.index+m[0].length);
}

/** Pune valoarea calculata (<v>) intr-o celula de formula, imediat dupa <f>; formula nu se schimba. */
function setCached(xml, ref, num){
  const m = xml.match(cellRe(ref));
  if(!m || m[2]==='/>') return xml;
  const body = (m[3]||'').replace(/<v>[^<]*<\/v>|<v\/>/g,'');
  const f = body.match(/<f\b[^>]*\/>|<f\b[^>]*>[\s\S]*?<\/f>/);
  if(!f) return xml;
  const attrs = m[1].replace(/\st="[^"]*"/g,'').replace(/\br="[^"]*"/g,'').trim();
  const nb = body.slice(0, f.index+f[0].length) + '<v>'+num+'</v>' + body.slice(f.index+f[0].length);
  const repl = '<c r="'+ref+'"'+(attrs?' '+attrs:'')+'>'+nb+'</c>';
  return xml.slice(0,m.index)+repl+xml.slice(m.index+m[0].length);
}

/** Valoarea numerica a celulei. Fara cuFormula, celulele de formula intorc null (valori „statice");
 *  cu cuFormula, intorc valoarea salvata de Excel (cache-ul <v>), daca exista. */
function getNumber(xml, ref, cuFormula){
  const m = xml.match(cellRe(ref));
  if(!m || m[2]==='/>') return null;
  const body = m[3]||'';
  if(/<f[\s>\/]/.test(body) && !cuFormula) return null;
  if(/\bt="(s|str|inlineStr|e|b)"/.test(m[1])) return null;   // text, eroare sau boolean: nu e o suma
  const v = body.match(/<v>([^<]*)<\/v>/);
  if(!v) return null;
  const n = parseFloat(v[1]);
  return isFinite(n)?n:null;
}

async function sharedStrings(zip){
  const f = zip.file('xl/sharedStrings.xml');
  if(!f) return [];
  return [...(await f.async('string')).matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m=>dezXml(m[1]));
}

async function sheetOf(buf){
  const zip = await JSZip.loadAsync(buf);
  const name = zip.file('xl/worksheets/sheet1.xml') ? 'xl/worksheets/sheet1.xml'
             : Object.keys(zip.files).find(n=>/^xl\/worksheets\/sheet\d+\.xml$/.test(n));
  if(!name) throw new Error('Fisierul nu contine o foaie de calcul valida.');
  return { zip, name, xml: await zip.file(name).async('string') };
}

/** Verifica daca sablonul este cel corectat (I117 are formula proprie, nu =H117). */
export function templateIsClean(xml){
  const m = xml.match(cellRe('I117'));
  if(!m) return { ok:false, reason:'celula I117 lipseste din sablon' };
  const body = m[3]||'';
  if(/<f[^>]*>\s*H117\s*<\/f>/.test(body)) return { ok:false, reason:'I117 = H117 (regresia din macheta de iulie)' };
  if(!/<f[\s>]/.test(body)) return { ok:false, reason:'I117 nu mai contine formula' };
  return { ok:true };
}

export async function loadTemplate(url){
  const res = await fetch(url, {cache:'no-cache'});
  if(!res.ok) throw new Error('Nu pot incarca sablonul ('+res.status+').');
  const buf = await res.arrayBuffer();
  const { xml } = await sheetOf(buf);
  return { buf, check: templateIsClean(xml) };
}

/** Scrie valorile in sablon si intoarce un Blob .xlsx. */
export async function buildXlsx(templateBuf, values, meta){
  const { zip, name, xml } = await sheetOf(templateBuf.slice(0));
  let s = xml;
  for(const r of Object.keys(values)){
    const v = values[r];
    COLS.forEach((c,i)=>{
      const ref = c+r;
      // formula ramane neatinsa; langa ea se scrie valoarea calculata, ca previzualizarile (Quick Look, e-mail)
      // sa arate totalurile. Excel le recalculeaza oricum la deschidere (fullCalcOnLoad).
      s = hasFormula(s, ref) ? setCached(s, ref, r3(v[KEYS[i]]||0)) : setValue(s, ref, r3(v[KEYS[i]]||0));
    });
  }
  if(meta.unitate) s = setText(s,'A1', meta.unitate);
  if(meta.titluLuna) s = setText(s,'B4', meta.titluLuna);
  zip.file(name, s);
  return zip.generateAsync({type:'blob', compression:'DEFLATE', mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}

/** Textul unei celule (inline, shared string sau valoare). */
function readCellText(xml, ref, shared){
  const m = xml.match(cellRe(ref));
  if(!m || m[2]==='/>') return '';
  const body = m[3]||'';
  const inline = body.match(/<is>([\s\S]*?)<\/is>/);
  if(inline) return dezXml(inline[1]);
  const v = body.match(/<v>([^<]*)<\/v>/);
  if(!v) return '';
  if(/t="s"/.test(m[1])){
    return (shared && shared[parseInt(v[1],10)]) || '';
  }
  return v[1];
}

/** Citeste valorile E..J de pe randurile 8..128 dintr-o macheta existenta. */
export async function readMacheta(buf){
  const { zip, xml } = await sheetOf(buf);
  const ss = await sharedStrings(zip);
  const out = {};
  for(let r=8;r<=128;r++){
    const v={};
    let any=false;
    // cu valoarea salvata si din celulele de formula: in machetele reale H54, H63 (=G+I) sunt formule pe randuri-frunza
    COLS.forEach((c,i)=>{ const n=getNumber(xml,c+r,true); v[KEYS[i]] = n===null?0:n; if(n!==null) any=true; });
    if(any) out[r]=v;
  }
  // valorile statice din fisier, cu null pentru celulele cu formula (folosite la verificarea consecventei)
  const statice = {};
  for(let r=8;r<=128;r++){ const v={}; COLS.forEach((c,i)=>{ v[KEYS[i]] = getNumber(xml,c+r); }); statice[r]=v; }
  return { values: out, statice, luna: readCellText(xml, 'B4', ss), unitate: readCellText(xml, 'A1', ss) };
}

/** Citeste fisa BC39 (xlsx): intoarce {bugetar, angajament, coloane:[...]} in lei. */
export async function readBC39(buf){
  const zip = await JSZip.loadAsync(buf);
  const shared = await sharedStrings(zip);
  const name = zip.file('xl/worksheets/sheet1.xml') ? 'xl/worksheets/sheet1.xml'
             : Object.keys(zip.files).find(n=>/^xl\/worksheets\/sheet\d+\.xml$/.test(n));
  if(!name) throw new Error('Fișierul nu conține o foaie de calcul validă.');
  const xml = await zip.file(name).async('string');
  const rows = {};
  for(const rm of xml.matchAll(/<row\b[^>]*?\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)){
    const cells = {};
    for(const cm of rm[2].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)){
      const col = (cm[1].match(/\br="([A-Z]+)\d+"/)||[])[1]; if(!col) continue;
      const body = cm[2]||'';
      const inline = body.match(/<is>([\s\S]*?)<\/is>/);
      if(inline){ cells[col] = dezXml(inline[1]); continue; }
      const vv = (body.match(/<v>([^<]*)<\/v>/)||[])[1];
      if(vv===undefined){ cells[col]=''; continue; }
      cells[col] = /\bt="s"/.test(cm[1]) ? (shared[parseInt(vv,10)]||'') : dezXml(vv);
    }
    rows[rm[1]] = cells;
  }
  const rowKeys = Object.keys(rows).sort((a,b)=>+a - +b);
  const headKey = rowKeys.find(rn => Object.values(rows[rn]||{}).some(v => /tip credit/i.test(String(v)))) || rowKeys[0] || '1';
  const head = rows[headKey]||{};
  const ordine = (a,b) => a.length-b.length || (a<b?-1:a>b?1:0);   // A..Z, apoi AA..
  const cols = Object.keys(head).sort(ordine);
  const tipCol = cols.find(c=>/tip credit/i.test(String(head[c]))) || 'E';
  const valCols = cols.filter(c=>ordine(c,tipCol)>0);
  // randul 8 al machetei include programele nationale, deci referinta e coloana „Total" daca exista
  const refCol = valCols.find(c=>/^total/i.test(String(head[c]))) || valCols[0];
  const out = { bugetar:null, angajament:null, coloane: valCols.map(c=>String(head[c])), coloanaRef: String(head[refCol]||''),
                detalii:{ bugetar:{}, angajament:{} }, luna:null, an:null, denumire:null };
  const lc = cols.find(c=>/luna/i.test(String(head[c])));
  const ac = cols.find(c=>/^an$/i.test(String(head[c])));
  const dc = cols.find(c=>/denumire/i.test(String(head[c])));
  for(const rn of rowKeys){
    if(+rn <= +headKey) continue;
    const rw = rows[rn];
    const tip = String(rw[tipCol]||'').toLowerCase();
    // valorile pot fi numere sau text („1.234.567"); celulele fara numar se ignora
    const val = numSauNull(rw[refCol]);
    if(val===null) continue;
    const key = tip.includes('bugetar') ? 'bugetar' : tip.includes('angajament') ? 'angajament' : null;
    if(!key) continue;
    out[key] = val;
    valCols.forEach(c=>{ const n=numSauNull(rw[c]); if(n!==null) out.detalii[key][String(head[c])] = n; });
    if(lc) out.luna = rw[lc]; if(ac) out.an = rw[ac]; if(dc) out.denumire = rw[dc];
  }
  return out;
}
