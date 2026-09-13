// "Mini baza de date" = un depozit GitHub. Fiecare luna e un fisier JSON in data/.
// Citire: publica (fara token) sau privata (cu token). Scriere: doar cu token.
const LS_CFG = 'execbug.cfg';
const LS_TOK = 'execbug.tok';
const LS_DRAFT = 'execbug.draft';
const API = 'https://api.github.com';

const b64enc = s => btoa(String.fromCharCode(...new TextEncoder().encode(s)));
const b64dec = s => new TextDecoder().decode(Uint8Array.from(atob(s.replace(/\n/g,'')), c=>c.charCodeAt(0)));

export const cfg = {
  get(){ try{ return JSON.parse(localStorage.getItem(LS_CFG)||'{}'); }catch{ return {}; } },
  set(v){ try{ localStorage.setItem(LS_CFG, JSON.stringify(v)); }catch{} },
  token(){ try{ return localStorage.getItem(LS_TOK)||''; }catch{ return ''; } },
  setToken(t){ try{ t ? localStorage.setItem(LS_TOK,t) : localStorage.removeItem(LS_TOK); }catch{} }
};

// Arhiva lunilor salvate in acest browser: { 'AAAA-LL': snapshot }
const LS_LUNI = 'execbug.luni';
export const luni = {
  all(){ try{ return JSON.parse(localStorage.getItem(LS_LUNI)||'{}'); }catch{ return {}; } },
  keys(){ return Object.keys(this.all()).sort().reverse(); },
  get(k){ return this.all()[k] || null; },
  set(k, data){ const a=this.all(); a[k]=data; localStorage.setItem(LS_LUNI, JSON.stringify(a)); },
  remove(k){ const a=this.all(); delete a[k]; try{ localStorage.setItem(LS_LUNI, JSON.stringify(a)); }catch{} }
};

// Folder de salvare pe disc (File System Access API, Chrome/Edge). Handle-ul se pastreaza in IndexedDB.
function idb(){ return new Promise((res,rej)=>{ const r=indexedDB.open('execbug',1); r.onupgradeneeded=()=>r.result.createObjectStore('kv'); r.onsuccess=()=>res(r.result); r.onerror=()=>rej(r.error); }); }
async function kv(op, k, v){ const db=await idb(); return new Promise((res,rej)=>{ const st=db.transaction('kv', op==='get'?'readonly':'readwrite').objectStore('kv'); const t = op==='get'?st.get(k): op==='set'?st.put(v,k): st.delete(k); t.onsuccess=()=>res(t.result); t.onerror=()=>rej(t.error); }); }
const NUME_LUNA = /^(\d{4})-(\d{2})\.json$/;
let folderHandle = null;
export const folder = {
  suportat(){ return typeof window!=='undefined' && 'showDirectoryPicker' in window; },
  async get(){ if(folderHandle) return folderHandle; try{ folderHandle = (await kv('get','folder')) || null; }catch{ folderHandle=null; } return folderHandle; },
  async alege(){
    const h = await window.showDirectoryPicker({ mode:'readwrite', id:'execbug', startIn:'documents' });
    folderHandle = h; try{ await kv('set','folder',h); }catch{}
    return h;
  },
  async uita(){ folderHandle=null; try{ await kv('del','folder'); }catch{} },
  /** 'granted' | 'prompt' | 'denied' | null (fara folder). Cu cere=true incearca sa obtina permisiunea (necesita un clic al utilizatorului). */
  async permisiune(cere){
    const h = await this.get(); if(!h) return null;
    if(!h.queryPermission) return 'granted';
    let p = await h.queryPermission({mode:'readwrite'});
    if(p==='prompt' && cere){ try{ p = await h.requestPermission({mode:'readwrite'}); }catch{ p='denied'; } }
    return p;
  },
  async scrie(name, text){
    const h = await this.get(); if(!h) throw new Error('Niciun folder ales.');
    const f = await h.getFileHandle(name, {create:true});
    const w = await f.createWritable(); await w.write(text); await w.close();
  },
  async citeste(name){
    const h = await this.get(); if(!h) return null;
    try{ const f = await h.getFileHandle(name); return JSON.parse(await (await f.getFile()).text()); }catch{ return null; }
  },
  async sterge(name){ const h = await this.get(); if(h) await h.removeEntry(name); },
  /** Lista lunilor din folder: [{key:'AAAA-LL', name, modificat}] */
  async lista(){
    const h = await this.get(); if(!h) return [];
    const out = [];
    for await (const e of h.values()){
      const m = e.kind==='file' && e.name.match(NUME_LUNA); if(!m) continue;
      const f = await e.getFile(); out.push({ key:`${m[1]}-${m[2]}`, name:e.name, modificat:f.lastModified });
    }
    return out.sort((a,b)=>b.key.localeCompare(a.key));
  }
};

export const draft = {
  save(state){ try{ localStorage.setItem(LS_DRAFT, JSON.stringify(state)); }catch{} },
  load(){ try{ return JSON.parse(localStorage.getItem(LS_DRAFT)||'null'); }catch{ return null; } },
  clear(){ try{ localStorage.removeItem(LS_DRAFT); }catch{} }
};

function conf(){
  const c = cfg.get();
  if(!c.owner || !c.repo) throw new Error('Configureaza depozitul GitHub (Setari).');
  return { owner:c.owner, repo:c.repo, branch:c.branch||'main', dir:(c.dir||'data').replace(/^\/|\/$/g,'') };
}

async function api(path, opts={}){
  const t = cfg.token();
  const h = { 'Accept':'application/vnd.github+json', 'X-GitHub-Api-Version':'2022-11-28', ...(opts.headers||{}) };
  if(t) h.Authorization = 'Bearer '+t;
  const res = await fetch(API+path, {...opts, headers:h});
  if(res.status===404) return null;
  if(!res.ok){
    let msg = res.status+' '+res.statusText;
    try{ const j = await res.json(); if(j.message) msg = j.message; }catch{}
    if(res.status===401||res.status===403) msg += ' — verifica tokenul si drepturile (Contents: read and write).';
    throw new Error(msg);
  }
  return res.json();
}

export const fileName = (an, luna) => `${an}-${String(luna).padStart(2,'0')}.json`;

export async function listMonths(){
  const c = conf();
  const j = await api(`/repos/${c.owner}/${c.repo}/contents/${c.dir}?ref=${encodeURIComponent(c.branch)}`);
  if(!j) return [];
  return j.filter(f=>f.type==='file' && f.name.endsWith('.json'))
          .map(f=>({name:f.name, path:f.path, sha:f.sha}))
          .sort((a,b)=>b.name.localeCompare(a.name));
}

export async function loadMonth(name){
  const c = conf();
  const j = await api(`/repos/${c.owner}/${c.repo}/contents/${c.dir}/${name}?ref=${encodeURIComponent(c.branch)}`);
  if(!j) return null;
  return { data: JSON.parse(b64dec(j.content)), sha: j.sha };
}

export async function saveMonth(name, data, message){
  const c = conf();
  if(!cfg.token()) throw new Error('Salvarea in GitHub necesita un token cu drept de scriere.');
  const cur = await api(`/repos/${c.owner}/${c.repo}/contents/${c.dir}/${name}?ref=${encodeURIComponent(c.branch)}`);
  const body = {
    message: message || `Executie bugetara ${name}`,
    content: b64enc(JSON.stringify(data, null, 2)),
    branch: c.branch
  };
  if(cur && cur.sha) body.sha = cur.sha;
  const j = await api(`/repos/${c.owner}/${c.repo}/contents/${c.dir}/${name}`, {
    method:'PUT', headers:{'Content-Type':'application/json'}, body: JSON.stringify(body)
  });
  return j.content.sha;
}

export async function checkAccess(){
  const c = conf();
  const j = await api(`/repos/${c.owner}/${c.repo}`);
  if(!j) throw new Error('Depozitul nu a fost gasit (sau tokenul nu are acces la el).');
  return { private: j.private, push: !!(j.permissions && j.permissions.push), full: j.full_name };
}
