// Parsare / formatare numerica pentru mii lei cu 3 zecimale (format romanesc).
export function parseNum(v){
  if(v===null||v===undefined) return 0;
  if(typeof v==='number') return isFinite(v)?v:0;
  let s=String(v).trim().replace(/\s| /g,'');
  if(!s) return 0;
  if(s.includes(',')){ s=s.replace(/\./g,'').replace(',','.'); }
  const n=parseFloat(s);
  return isFinite(n)?n:0;
}
export const r3 = n => Math.round((n+Number.EPSILON)*1000)/1000;
export function fmt(n,dec=3){
  if(n===null||n===undefined||!isFinite(n)) return '';
  return n.toLocaleString('ro-RO',{minimumFractionDigits:dec,maximumFractionDigits:dec});
}
export const fmtLei = n => Math.round(n).toLocaleString('ro-RO');
