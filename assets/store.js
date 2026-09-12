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
