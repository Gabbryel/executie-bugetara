// Parsare / formatare numerica pentru mii lei cu 3 zecimale (format romanesc).
export function parseNum(v){
  if(v===null||v===undefined) return 0;
  if(typeof v==='number') return isFinite(v)?v:0;
  let s=String(v).trim().replace(/[\s\u00A0\u202F]/g,'');
  if(!s) return 0;
  const lastComma = s.lastIndexOf(',');
  const lastDot = s.lastIndexOf('.');
  if(lastComma > -1 && lastDot > -1){
    if(lastComma > lastDot){
      // Format 1.234,56 (romanesc: puncte pentru mii, virgula zecimala)
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      // Format 1,234.56 (anglo-saxon: virgule pentru mii, punct zecimal)
      s = s.replace(/,/g, '');
    }
  } else if(lastComma > -1){
    // Doar virgule: ultima virgula e zecimala, restul separatori de mii
    const parts = s.split(',');
    if(parts.length > 2){
      const dec = parts.pop();
      s = parts.join('') + '.' + dec;
    } else {
      s = s.replace(',', '.');
    }
  } else if(s.split('.').length > 2){
    // Mai multe puncte, fara virgula: separatori de mii (1.234.567 = 1234567)
    s = s.replace(/\./g, '');
  }
  const n=parseFloat(s);
  return isFinite(n)?n:0;
}
/** Numarul dintr-o celula sau null daca celula nu contine un numar (text, gol). */
export function numSauNull(v){
  if(typeof v==='number') return isFinite(v)?v:null;
  const s = String(v??'').trim();
  if(!/^[-+]?[\d\s  .,]*\d[\d\s  .,]*(e[-+]?\d+)?$/i.test(s)) return null;
  return parseNum(s);
}
export const r3 = n => n === 0 ? 0 : Math.sign(n) * (Math.round((Math.abs(n)+Number.EPSILON)*1000)/1000);
export function fmt(n,dec=3){
  if(n===null||n===undefined||!isFinite(n)) return '';
  return n.toLocaleString('ro-RO',{minimumFractionDigits:dec,maximumFractionDigits:dec});
}
export const fmtLei = n => Math.round(n).toLocaleString('ro-RO');

