// JSZip vendorizat (assets/jszip.min.js) incarcat in Node, ca global, cum il vede si browserul.
// Fara npm: fisierul UMD primeste un `module` CommonJS si isi pune exportul acolo.
import fs from 'fs';
const src = fs.readFileSync(new URL('../assets/jszip.min.js', import.meta.url), 'utf8');
const module = { exports:{} };
new Function('module', 'exports', src)(module, module.exports);
globalThis.JSZip = module.exports;
export default module.exports;
