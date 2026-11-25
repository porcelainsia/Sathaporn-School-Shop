const DB = window.__FB = window.__FB || {};
const LOG = window.__FBLOG = [];
const subs = [];
export function initializeFirestore(){ return {}; }
export function persistentLocalCache(){ return {}; }
export function persistentMultipleTabManager(){ return {}; }
export function collection(db, name){ return {col:name}; }
export function doc(db, col, id){ return {col, id}; }
export function serverTimestamp(){ return {__ts:Date.now()}; }
export function increment(n){ return {__inc:n}; }
function snapOf(col){ const m = DB[col] || {}; const docs = Object.keys(m).map(id=>({id, data:()=>JSON.parse(JSON.stringify(m[id]))}));
  return {docs, metadata:{hasPendingWrites:false}, forEach:(f)=>docs.forEach(f)}; }
export async function getDocs(c){ return snapOf(c.col); }
export function onSnapshot(c, opts, cb){ subs.push({col:c.col, cb}); cb(snapOf(c.col)); return ()=>{}; }
function notify(col){ subs.filter(s=>s.col===col).forEach(s=>setTimeout(()=>s.cb(snapOf(col)),0)); }
function apply(old, data){ const o = Object.assign({}, old); Object.keys(data).forEach(k=>{ const v = data[k]; if (v && v.__inc !== undefined) o[k] = (Number(o[k])||0) + v.__inc; else o[k] = v; }); return o; }
export async function setDoc(r, data, opt){ DB[r.col] = DB[r.col] || {}; DB[r.col][r.id] = apply(opt && opt.merge ? DB[r.col][r.id] : {}, data); LOG.push(['set', r.col, r.id, data]); }
export async function updateDoc(r, data){ if (!DB[r.col] || !DB[r.col][r.id]) throw new Error('no doc '+r.col+'/'+r.id); DB[r.col][r.id] = apply(DB[r.col][r.id], data); LOG.push(['update', r.col, r.id, data]); }
export async function deleteDoc(r){ if (DB[r.col]) delete DB[r.col][r.id]; LOG.push(['delete', r.col, r.id]); }
export function writeBatch(){ const ops=[]; return {set:(r,d,o)=>ops.push(()=>setDoc(r,d,o)), update:(r,d)=>ops.push(()=>updateDoc(r,d)), delete:(r)=>ops.push(()=>deleteDoc(r)), commit:async()=>{ for (const f of ops) await f(); }}; }
window.__remoteWrite = function(col, id, data){ DB[col] = DB[col] || {}; DB[col][id] = Object.assign({}, DB[col][id]||{}, data); notify(col); };
