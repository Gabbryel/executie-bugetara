// Citire/scriere .xlsx prin editare directa a XML-ului din pachet (JSZip).
// Scopul: exportul pastreaza intact formatarea, formulele si setarile de tiparire ale machetei CAS.
import { r3 } from './num.js?v=20260914-114551';

const COLS = ['E','F','G','H','I','J'];
const KEYS = ['e','f','g','h','i','j'];

const esc = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

// Localizeaza <c r="REF" ...> fie self-closing, fie cu continut.
function cellRe(ref){
  return new RegExp('<c r="'+ref+'"([^>]*?)(\\/>|>([\\s\\S]*?)<\\/c>)');
}

function hasFormula(xml, ref){
  const m = xml.match(cellRe(ref));
  return !!(m && m[2]!=='/>' && /<f[\s>]/.test(m[3]||''));
}

function setValue(xml, ref, num){
  const m = xml.match(cellRe(ref));
  if(!m) return xml;
  const attrs = m[1].replace(/\st="[^"]*"/,'');
  const repl = '<c r="'+ref+'"'+attrs+'><v>'+num+'</v></c>';
  return xml.slice(0,m.index)+repl+xml.slice(m.index+m[0].length);
}

function setText(xml, ref, text){
  const m = xml.match(cellRe(ref));
  if(!m) return xml;
  const attrs = m[1].replace(/\st="[^"]*"/,'');
  const repl = '<c r="'+ref+'"'+attrs+' t="inlineStr"><is><t xml:space="preserve">'+esc(text)+'</t></is></c>';
  return xml.slice(0,m.index)+repl+xml.slice(m.index+m[0].length);
}

function getNumber(xml, ref){
  const m = xml.match(cellRe(ref));
  if(!m || m[2]==='/>') return null;
  const body = m[3]||'';
  if(/<f[\s>]/.test(body)) return null;          // celula de formula -> se recalculeaza
  const v = body.match(/<v>([^<]*)<\/v>/);
  if(!v) return null;
  const n = parseFloat(v[1]);
  return isFinite(n)?n:null;
}

async function sharedStrings(zip){
  const f = zip.file('xl/sharedStrings.xml');
  if(!f) return [];
  return [...(await f.async('string')).matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m=>m[1].replace(/<[^>]+>/g,''));
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
      // celulele de formula raman neatinse: Excel le recalculeaza la deschidere
      if(hasFormula(s, ref)) return;
      s = setValue(s, ref, r3(v[KEYS[i]]||0));
    });
  }
  if(meta.unitate) s = setText(s,'A1', meta.unitate);
  if(meta.titluLuna) s = setText(s,'B4', meta.titluLuna);
  zip.file(name, s);
  return zip.generateAsync({type:'blob', compression:'DEFLATE', mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}

/** Citeste valorile E..J de pe randurile 8..128 dintr-o macheta existenta. */
async function readCellText(buf, xml, ref){
  const m = xml.match(cellRe(ref));
  if(!m || m[2]==='/>') return '';
  const body = m[3]||'';
  const inline = body.match(/<t[^>]*>([^<]*)<\/t>/);
  if(inline) return inline[1];
  const v = body.match(/<v>([^<]*)<\/v>/);
  if(!v) return '';
  if(/t="s"/.test(m[1])){
    const zip = await JSZip.loadAsync(buf.slice(0));
    const ss = await sharedStrings(zip);
    return ss[parseInt(v[1],10)]||'';
  }
  return v[1];
}

export async function readMacheta(buf){
  const { xml } = await sheetOf(buf);
  const out = {};
  for(let r=8;r<=128;r++){
    const v={};
    let any=false;
    COLS.forEach((c,i)=>{ const n=getNumber(xml,c+r); v[KEYS[i]] = n===null?0:n; if(n!==null) any=true; });
    if(any) out[r]=v;
  }
  // valorile statice din fisier, cu null pentru celulele cu formula (folosite la verificarea consecventei)
  const statice = {};
  for(let r=8;r<=128;r++){ const v={}; COLS.forEach((c,i)=>{ v[KEYS[i]] = getNumber(xml,c+r); }); statice[r]=v; }
  return { values: out, statice, luna: await readCellText(buf, xml, 'B4'), unitate: await readCellText(buf, xml, 'A1') };
}

/** Citeste fisa BC39 (xlsx): intoarce {bugetar, angajament, coloane:[...]} in lei. */
export async function readBC39(buf){
  const zip = await JSZip.loadAsync(buf);
  const shared = zip.file('xl/sharedStrings.xml')
    ? [...(await zip.file('xl/sharedStrings.xml').async('string')).matchAll(/<si>([\s\S]*?)<\/si>/g)]
        .map(m=>m[1].replace(/<[^>]+>/g,''))
    : [];
  const name = Object.keys(zip.files).find(n=>/^xl\/worksheets\/sheet\d+\.xml$/.test(n));
  const xml = await zip.file(name).async('string');
  const rows = {};
  for(const rm of xml.matchAll(/<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)){
    const cells = {};
    for(const cm of rm[2].matchAll(/<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)){
      const body = cm[3]||'';
      const vv = (body.match(/<v>([^<]*)<\/v>/)||[])[1];
      if(vv===undefined){ cells[cm[1]]=''; continue; }
      cells[cm[1]] = /t="s"/.test(cm[2]) ? (shared[parseInt(vv,10)]||'') : vv;
    }
    rows[rm[1]] = cells;
  }
  const head = rows['1']||{};
  const cols = Object.keys(head).sort();
  const tipCol = cols.find(c=>/tip credit/i.test(String(head[c]))) || 'E';
  const valCols = cols.filter(c=>c>tipCol);
  // randul 8 al machetei include programele nationale, deci referinta e coloana „Total" daca exista
  const refCol = valCols.find(c=>/^total/i.test(String(head[c]))) || valCols[0];
  const out = { bugetar:null, angajament:null, coloane: valCols.map(c=>String(head[c])), coloanaRef: String(head[refCol]||''),
                detalii:{ bugetar:{}, angajament:{} }, luna:null, an:null, denumire:null };
  const lc = cols.find(c=>/luna/i.test(String(head[c])));
  const ac = cols.find(c=>/^an$/i.test(String(head[c])));
  const dc = cols.find(c=>/denumire/i.test(String(head[c])));
  for(const rn of Object.keys(rows)){
    if(rn==='1') continue;
    const rw = rows[rn];
    const tip = String(rw[tipCol]||'').toLowerCase();
    const val = parseFloat(rw[refCol]);
    if(!isFinite(val)) continue;
    const key = tip.includes('bugetar') ? 'bugetar' : tip.includes('angajament') ? 'angajament' : null;
    if(!key) continue;
    out[key] = val;
    valCols.forEach(c=>{ const n=parseFloat(rw[c]); if(isFinite(n)) out.detalii[key][String(head[c])] = n; });
    if(lc) out.luna = rw[lc]; if(ac) out.an = rw[ac]; if(dc) out.denumire = rw[dc];
  }
  return out;
}
