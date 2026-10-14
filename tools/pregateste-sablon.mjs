// Pregateste template/macheta.xlsx dintr-o macheta primita de la CAS (foaia BC, randurile 8–128, coloanele E–J):
//  - golește valorile (celulele cu valoare primesc 0, cele goale raman goale), pastrand stilurile;
//  - pe randurile-frunza scoate formulele (in machetele reale: H54, H63 = G+I; J50, J63 trimit la coloana K, care nu exista);
//  - pe randurile de sinteza, o formula care nu urmeaza structura din coloana E se reface din ea (ex. I117 = H117);
//  - sterge valorile salvate ale formulelor si cere recalcularea la deschidere (fullCalcOnLoad), fara calcChain.
// Formatarea, setarile de tiparire si restul foii raman neatinse.
//   node tools/pregateste-sablon.mjs "Macheta CAS.xlsx" template/macheta.xlsx
// Dupa inlocuirea sablonului: node tools/genereaza-rows.mjs (si --scrie, daca s-a schimbat structura).
import fs from 'fs';
import JSZip from './jszip-node.mjs';
import { genereazaRows } from './genereaza-rows.mjs';

const COLS = ['E','F','G','H','I','J'];
const SHEET = 'xl/worksheets/sheet1.xml';

export async function pregatesteSablon(buf){
  const zip = await JSZip.loadAsync(buf);
  const rows = await genereazaRows(buf);
  const by = Object.fromEntries(rows.map(r=>[r.r, r]));
  let xml = await zip.file(SHEET).async('string');
  const jurnal = [];
  const formulaE = {};
  for(const r of rows){ const m = celula(xml, 'E'+r.r); formulaE[r.r] = m && m[3] ? (m[3].match(/<f[^>]*>([^<]*)<\/f>/)||[])[1] ?? null : null; }

  for(const row of rows){
    // sinteza „adevarata" = are formula in E; r.38 e sinteza doar in model (static in macheta), deci se trateaza ca frunza
    const sinteza = !!formulaE[row.r];
    for(const c of COLS){
      const ref = c+row.r;
      const m = celula(xml, ref); if(!m) continue;
      const attrs = m[1].replace(/\st="[^"]*"/g,'').replace(/\br="[^"]*"/g,'').trim();
      const deschis = '<c r="'+ref+'"'+(attrs?' '+attrs:'');
      const body = m[2]==='/>' ? '' : (m[3]||'');
      const f = body.match(/<f\b([^>]*)\/>|<f\b([^>]*)>([\s\S]*?)<\/f>/);
      let nou;
      if(!sinteza){
        if(f){ jurnal.push(`${ref}: formula „${f[3]??'(partajata)'}" scoasa — rândul ${row.r} se completează`); nou = deschis+'><v>0</v></c>'; }
        else nou = /<v>/.test(body) ? deschis+'><v>0</v></c>' : deschis+'/>';
      } else if(!f){
        if(/<v>/.test(body)) jurnal.push(`${ref}: valoare statică pe un rând de sinteză — golită`);
        nou = deschis+'/>';
      } else {
        let fx = f[0];
        if(f[3]!==undefined && c!=='E'){
          const asteptat = formulaE[row.r].replace(/\bE(\d+)/g, c+'$1');
          if(f[3].trim()!==asteptat){ jurnal.push(`${ref}: formula „${f[3]}" refăcută ca „${asteptat}"`); fx = '<f'+f[2]+'>'+asteptat+'</f>'; }
        }
        nou = deschis+'>'+fx+'</c>';   // fara <v>: Excel recalculeaza
      }
      xml = xml.slice(0,m.index) + nou + xml.slice(m.index+m[0].length);
    }
  }
  // A1 = denumirea unitatii: in sablonul public ramane generica (exportul o completeaza din formular)
  const a1 = celula(xml, 'A1');
  if(a1){
    const attrs = a1[1].replace(/\st="[^"]*"/g,'').replace(/\br="[^"]*"/g,'').trim();
    xml = xml.slice(0,a1.index) + '<c r="A1"'+(attrs?' '+attrs:'')+' t="inlineStr"><is><t>UNITATEA SANITARA</t></is></c>' + xml.slice(a1.index+a1[0].length);
  }
  zip.file(SHEET, xml);

  // recalculare la deschidere; calcChain ar indica formule care nu mai exista
  let wb = await zip.file('xl/workbook.xml').async('string');
  if(/<calcPr\b[^>]*fullCalcOnLoad=/.test(wb)) wb = wb.replace(/fullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"');
  else if(/<calcPr\b/.test(wb)) wb = wb.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"');
  else wb = wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
  zip.file('xl/workbook.xml', wb);
  if(zip.file('xl/calcChain.xml')){
    zip.remove('xl/calcChain.xml');
    const rels = 'xl/_rels/workbook.xml.rels', ct = '[Content_Types].xml';
    zip.file(rels, (await zip.file(rels).async('string')).replace(/<Relationship\b[^>]*calcChain[^>]*\/>/g,''));
    zip.file(ct, (await zip.file(ct).async('string')).replace(/<Override\b[^>]*calcChain[^>]*\/>/g,''));
    jurnal.push('calcChain.xml șters');
  }
  return { buf: await zip.generateAsync({ type:'nodebuffer', compression:'DEFLATE' }), jurnal, rows: rows.length, frunze: rows.filter(r=>!r.ch).length };
}

function celula(xml, ref){
  return xml.match(new RegExp('<c\\b([^>]*?\\br="'+ref+'"[^>]*?)(\\/>|>([\\s\\S]*?)<\\/c>)'));
}

if(import.meta.url === `file://${process.argv[1]}`){
  const [intrare, iesire] = process.argv.slice(2);
  if(!intrare || !iesire){ console.log('Folosire: node tools/pregateste-sablon.mjs <macheta CAS.xlsx> <iesire.xlsx>'); process.exit(2); }
  const rez = await pregatesteSablon(fs.readFileSync(intrare));
  fs.writeFileSync(iesire, rez.buf);
  console.log(`${rez.rows} rânduri, ${rez.frunze} de completat. Modificări:`);
  rez.jurnal.forEach(l=>console.log('  ' + l));
  console.log(`Scris: ${iesire}. Verifică apoi: node tools/genereaza-rows.mjs`);
}
