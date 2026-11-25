// Data layer: Firebase sign-in, loading everything the app shows, and saving changes.
//
// The app's logic (js/app-logic.js) works on plain objects in its state (sales, stock,
// prices, receipts, ...). This file turns the Firestore documents into those objects
// when the app starts, keeps them up to date when another phone or the old apps change
// something, and writes back whatever the logic changes (see persist()).
//
// The uniform and stationery collections keep the same field names the old apps use,
// so both old and new apps can run side by side during the change-over week.
import { initializeApp } from 'https://www.gstatic.com/firebasejs/12.18.0/firebase-app.js';
import { getAuth, GoogleAuthProvider, signInWithPopup, signInWithRedirect, signOut as fbSignOut, onAuthStateChanged }
  from 'https://www.gstatic.com/firebasejs/12.18.0/firebase-auth.js';
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager, collection, doc, getDocs, onSnapshot,
  setDoc, updateDoc, deleteDoc, increment, serverTimestamp, writeBatch }
  from 'https://www.gstatic.com/firebasejs/12.18.0/firebase-firestore.js';
import { UCAT, UORDER, SMETA, SORDER } from './catalog.js';

const firebaseConfig = {
  apiKey: 'AIzaSyBsOmyOALjussJ81waMLjgJBQvVFnQ3xIc',
  authDomain: 'sathaporn-school-shop.firebaseapp.com',
  projectId: 'sathaporn-school-shop',
  storageBucket: 'sathaporn-school-shop.firebasestorage.app',
  messagingSenderId: '868289495495',
  appId: '1:868289495495:web:589e4d5d919241976919e2'
};
export const OWNER_EMAIL = 'a.satrabhaya@gmail.com';

// ?test=1 in the address uses separate "test_" collections, so a trial run never touches real data.
const TEST = new URLSearchParams(location.search).has('test');
const P = TEST ? 'test_' : '';
const C = {
  uProducts: P + 'uniform_products', uSales: P + 'uniform_transactions', uRestocks: P + 'uniform_restocks',
  sProducts: P + 'stationary_products', sSales: P + 'sales', sRestocks: P + 'stationary_restocks',
  batches: P + 'shop_batches', demand: P + 'shop_demand', closes: P + 'shop_closes', drawer: P + 'shop_drawer',
  merges: P + 'shop_merges', meta: P + 'shop_meta'
};
export const COLLECTIONS = C;
export const IS_TEST = TEST;

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = initializeFirestore(app, { localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }) });

// ---------- dates ----------
export function bangkokDate(d){
  const t = new Date((d ? d.getTime() : Date.now()) + 7 * 3600 * 1000);
  return t.toISOString().slice(0, 10);
}
function bangkokHM(d){
  const t = new Date(d.getTime() + 7 * 3600 * 1000);
  return t.toISOString().slice(11, 16);
}
function hm(s){ // old apps store th-TH times like "09:12" or "9:12"
  const m = String(s || '').match(/(\d{1,2})[:.](\d{2})/);
  return m ? (m[1].length === 1 ? '0' + m[1] : m[1]) + ':' + m[2] : '';
}
const THM = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
function thaiLabel(iso){ const p = iso.split('-').map(Number); return p[2] + ' ' + THM[p[1] - 1] + ' ' + (p[0] + 543); }
function clean(o){ return JSON.parse(JSON.stringify(o === undefined ? null : o)); }

// ---------- what the app reads ----------
export const DATA = {
  today: bangkokDate(), test: TEST,
  auth: 'login', role: 'none', me: '', myName: '', user: null,
  products: [], sales: [], receipts: [], batches: {}, demand: [], names: {}, rosters: [],
  closes: [], dMoves: [], merges: [], dStart: { date: bangkokDate(), amt: 0 }, staff: [],
  lastBackup: '', lastBackupAt: '',
  signIn, signOut
};

// raw documents by collection, kept so saves can compare against what is stored
const RAW = {};

// ---------- mapping: Firestore -> app objects ----------
function sizeRank(cat, size){
  const list = (UCAT[cat] && UCAT[cat].sizes) || [];
  const i = list.indexOf(String(size));
  if (i >= 0) return i;
  const n = parseFloat(size);
  return 100 + (isNaN(n) ? 500 : n);
}
function uProduct(id, d){
  const m = UCAT[d.category] || {};
  return {
    id, shop: 'u', groups: d.groups || m.groups || ['student'], gender: d.gender || m.gender || '',
    part: d.part || m.part || 'top', tone: d.tone || m.tone || (m.gender || 'boy'), icon: d.icon || m.icon || 'shirt',
    name: d.categoryLabel || m.name || id, size: d.size == null ? '' : String(d.size), price: Number(d.price) || 0,
    stock: Number(d.remaining) || 0, received: Number(d.received) || 0, cat: d.category || '',
    _order: (UORDER.indexOf(d.category) < 0 ? 99 : UORDER.indexOf(d.category)) * 1000 + sizeRank(d.category, d.size)
  };
}
function guessIcon(name){
  if (/สมุด/.test(name)) return 'book'; if (/ปากกา/.test(name)) return 'pen'; if (/ดินสอ|กบ/.test(name)) return 'pencil';
  if (/ยางลบ|ลบคำผิด/.test(name)) return 'eraser'; if (/ไม้บรรทัด/.test(name)) return 'ruler';
  if (/บัตร/.test(name)) return 'glue'; if (/วอกเกิ้ล|ผ้าพันคอ|เชือก/.test(name)) return 'scissors'; return 'color';
}
function guessCat(d){
  const g = guessGrp(d);
  return ['วอกเกิ้ล', 'ผ้าพันคอ', 'หมวก', 'เข็ม', 'เชือก'].indexOf(g) >= 0 ? 'tools' : (g === 'สมุด' || g === 'บัตร' ? 'paper' : 'write');
}
function guessGrp(d){
  const n = d.name || '', c = d.cat || '';
  for (const g of ['วอกเกิ้ล', 'ผ้าพันคอ', 'หมวก', 'เข็ม', 'เชือก']) if (n.indexOf(g) >= 0) return g;
  if (/บัตร/.test(n + c)) return 'บัตร';
  if (/สมุด/.test(n + c)) return 'สมุด';
  return 'เครื่องเขียน';
}
function sProduct(id, d){
  const m = SMETA[d.name] || {};
  const oi = SORDER.indexOf(d.name);
  return {
    id, shop: 's', cat: d.appCat || m.cat || guessCat(d), icon: d.icon || m.icon || guessIcon(d.name || ''),
    name: d.name || id, size: '', price: Number(d.price) || 0, stock: Number(d.remaining) || 0, received: Number(d.received) || 0,
    grp: d.grp != null ? d.grp : (m.grp || guessGrp(d)), short: d.short != null ? d.short : (m.short || ''), sub: d.sub != null ? d.sub : (m.sub || ''),
    sg: d.sg != null ? d.sg : (m.sg || ''), lv: d.lv != null ? d.lv : (m.lv || ''), hidden: !!d.hidden,
    _order: 50000 + (oi < 0 ? 500 + (Number(d.order) || 0) : oi)
  };
}
function uSale(id, d){
  const a = d.app || {};
  const transfer = d.transferAmount != null ? Number(d.transferAmount) : (d.payment === 'transfer' ? Number(d.total) || 0 : 0);
  return Object.assign({}, a, {
    id, shop: 'u', date: d.date || (d.createdAt && d.createdAt.toDate ? bangkokDate(d.createdAt.toDate()) : DATA.today),
    time: a.time || hm(d.time), code: d.studentCode != null ? String(d.studentCode) : (a.code || ''),
    items: (d.items || []).map(i => ({ pid: i.productId, qty: Number(i.qty) || 0, price: Number(i.price) || 0 })),
    transfer, seller: d.seller || a.seller || '', voided: a.voided || (d.voided ? true : false)
  });
}
function sSale(id, d){
  const a = d.app || {};
  const ts = d.ts ? new Date(d.ts) : null;
  const date = /^\d{4}-\d{2}-\d{2}$/.test(d.dateKey || '') ? d.dateKey : (ts ? bangkokDate(ts) : DATA.today);
  return Object.assign({}, a, {
    id, shop: 's', date, time: a.time || (ts ? bangkokHM(ts) : ''), code: a.code || '',
    items: (d.items || []).map(i => ({ pid: i.id, qty: Number(i.qty) || 0, price: Number(i.price) || 0 })),
    transfer: Number(a.transfer) || 0, seller: d.seller || a.seller || '', voided: a.voided || (d.voided ? true : false)
  });
}
function receiptOf(id, d, shop){
  const a = d.app || {};
  const p = PMAP()[d.productId];
  return Object.assign({ rid: id }, a, {
    date: d.date || DATA.today, pid: d.productId, qty: Number(d.qty) || 0,
    label: a.label || (p ? (p.size ? p.name + ' #' + p.size : p.name) : (d.name || d.productId)),
    memo: a.memo || '', cost: a.cost != null ? a.cost : null, batch: a.batch || null, shop
  });
}
let _pmap = {};
function PMAP(){ return _pmap; }

function buildProducts(){
  const list = [];
  Object.entries(RAW[C.uProducts] || {}).forEach(([id, d]) => list.push(uProduct(id, d)));
  Object.entries(RAW[C.sProducts] || {}).forEach(([id, d]) => list.push(sProduct(id, d)));
  list.sort((a, b) => a._order - b._order);
  _pmap = {}; list.forEach(p => { _pmap[p.id] = p; });
  (DATA.merges || []).forEach(m => { if (_pmap[m.from]) _pmap[m.from].mergedInto = m.to; });
  return list;
}
function buildAll(){
  // keep the same product objects the logic already holds, so its PMAP stays valid
  const fresh = buildProducts();
  const byId = {}; DATA.products.forEach(p => { byId[p.id] = p; });
  const merged = fresh.map(p => byId[p.id] ? Object.assign(byId[p.id], p) : p);
  DATA.products.length = 0; merged.forEach(p => DATA.products.push(p));
  _pmap = {}; DATA.products.forEach(p => { _pmap[p.id] = p; });

  const sales = [];
  Object.entries(RAW[C.uSales] || {}).forEach(([id, d]) => sales.push(uSale(id, d)));
  Object.entries(RAW[C.sSales] || {}).forEach(([id, d]) => sales.push(sSale(id, d)));
  sales.sort((a, b) => (b.date + b.time).localeCompare(a.date + a.time));
  DATA.sales = sales.filter(x => x.items.every(i => _pmap[i.pid]));
  const rec = [];
  Object.entries(RAW[C.uRestocks] || {}).forEach(([id, d]) => rec.push(receiptOf(id, d, 'u')));
  Object.entries(RAW[C.sRestocks] || {}).forEach(([id, d]) => rec.push(receiptOf(id, d, 's')));
  rec.sort((a, b) => b.date.localeCompare(a.date));
  DATA.receipts = rec.filter(r => _pmap[r.pid]);
  DATA.batches = {}; Object.entries(RAW[C.batches] || {}).forEach(([id, d]) => { DATA.batches[id] = Object.assign({ id }, d); });
  DATA.demand = Object.entries(RAW[C.demand] || {}).map(([id, d]) => Object.assign({ id }, d)).filter(d => _pmap[d.pid]);
  DATA.closes = Object.values(RAW[C.closes] || {}).sort((a, b) => (b.date + b.at).localeCompare(a.date + a.at));
  DATA.dMoves = Object.entries(RAW[C.drawer] || {}).map(([id, d]) => Object.assign({ id }, d)).sort((a, b) => (b.date + b.at).localeCompare(a.date + a.at));
  const st = (RAW[C.meta] || {}).students || {};
  DATA.names = st.names || {}; DATA.rosters = st.rosters || [];
  const set = (RAW[C.meta] || {}).settings || {};
  DATA.dStart = set.dStart || { date: DATA.today, amt: 0 };
  DATA.staff = set.staff || [];
  DATA.lastBackup = set.lastBackup || ''; DATA.lastBackupAt = set.lastBackupAt || '';
}

// ---------- sign-in ----------
export function signIn(){
  const pr = new GoogleAuthProvider(); pr.setCustomParameters({ prompt: 'select_account' });
  signInWithPopup(auth, pr).then(() => location.reload()).catch(e => {
    if (e && (e.code === 'auth/popup-blocked' || e.code === 'auth/operation-not-supported-in-this-environment')) signInWithRedirect(auth, pr);
    else console.error('sign-in failed', e);
  });
}
export function signOut(){ fbSignOut(auth).then(() => location.reload()); }

function waitForUser(){
  return new Promise(res => { const off = onAuthStateChanged(auth, u => { off(); res(u); }); });
}

// ---------- start ----------
const LISTEN = ['uProducts', 'sProducts', 'uSales', 'sSales', 'uRestocks', 'sRestocks', 'batches', 'demand', 'closes', 'drawer', 'merges', 'meta'];

export async function start(){
  const u = await waitForUser();
  if (!u || u.isAnonymous){ DATA.auth = 'login'; return false; }
  DATA.user = u; DATA.me = (u.email || '').toLowerCase();
  DATA.myName = DATA.me === OWNER_EMAIL ? 'Porcelain' : (u.displayName || DATA.me.split('@')[0] || '').split(' ')[0];
  await Promise.all(LISTEN.map(async k => {
    const snap = await getDocs(collection(db, C[k]));
    const m = {}; snap.forEach(d => { m[d.id] = d.data(); }); RAW[C[k]] = m;
  }));
  DATA.merges = Object.values(RAW[C.merges] || {});
  const set = (RAW[C.meta] || {}).settings || {};
  const staff = (set.staff || []).map(x => (x.email || '').toLowerCase());
  if (DATA.me === OWNER_EMAIL) DATA.role = 'owner';
  else if (staff.indexOf(DATA.me) >= 0) DATA.role = 'helper';
  else { DATA.auth = 'denied'; return false; }
  DATA.auth = 'in';
  buildAll();
  return true;
}

// live updates from other devices and the old apps
let APP = null, applying = false;
export function listen(appLogic){
  APP = appLogic;
  LISTEN.forEach(k => {
    let first = true;
    onSnapshot(collection(db, C[k]), { includeMetadataChanges: false }, snap => {
      if (first){ first = false; return; }
      if (snap.metadata.hasPendingWrites) return; // our own save; state already has it
      const m = {}; snap.forEach(d => { m[d.id] = d.data(); }); RAW[C[k]] = m;
      if (k === 'merges') DATA.merges = Object.values(m);
      buildAll();
      const s = APP.state, stock = {}, prices = {};
      DATA.products.forEach(p => { stock[p.id] = p.stock; prices[p.id] = p.price; });
      applying = true;
      try {
        APP.setState({ sales: DATA.sales, stock, prices, receipts: DATA.receipts, batches: DATA.batches, demand: DATA.demand,
          closes: DATA.closes, dMoves: DATA.dMoves, merges: DATA.merges, names: Object.assign({}, s.names, DATA.names),
          rosters: DATA.rosters.length ? DATA.rosters : s.rosters, staff: DATA.staff, dStart: DATA.dStart });
      } finally { applying = false; }
    }, err => console.error('listen failed', k, err));
  });
}

// ---------- saving: compare the state before and after each change ----------
const queue = [];
let flushing = false;
function q(fn){ queue.push(fn); if (!flushing) flush(); }
async function flush(){
  flushing = true;
  while (queue.length){
    const fn = queue.shift();
    try { await fn(); } catch (e){ console.error('save failed', e); if (APP && APP.flash) APP.flash('บันทึกไม่สำเร็จ ลองอีกครั้ง'); }
  }
  flushing = false;
}

function uSaleDoc(x){
  const items = x.items.map(i => { const p = _pmap[i.pid] || {}; return { productId: i.pid, category: p.cat || '', categoryLabel: p.name || '', size: p.size || '', price: i.price, qty: i.qty, subtotal: i.price * i.qty }; });
  const total = items.reduce((a, i) => a + i.subtotal, 0);
  const transfer = Math.min(Number(x.transfer) || 0, total);
  const app = Object.assign({}, x); delete app.items; delete app.id; delete app.shop;
  return { date: x.date, time: x.time, seller: x.seller || '', cashAmount: total - transfer, transferAmount: transfer,
    studentCode: x.code || '', items, total, voided: !!x.voided, app: clean(app) };
}
function sSaleDoc(x){
  const items = x.items.map(i => { const p = _pmap[i.pid] || {}; return { id: i.pid, name: p.name || '', price: i.price, qty: i.qty }; });
  const total = items.reduce((a, i) => a + i.price * i.qty, 0);
  const app = Object.assign({}, x); delete app.items; delete app.id; delete app.shop;
  return { id: x.id, dateKey: x.date, dateLabel: thaiLabel(x.date), items, total,
    ts: new Date(x.date + 'T' + (x.time || '00:00') + ':00+07:00').toISOString(), seller: x.seller || '', voided: !!x.voided, app: clean(app) };
}
function byKey(list, key){ const m = new Map(); (list || []).forEach(x => m.set(key(x), x)); return m; }

export function persist(prev, next){
  if (applying || DATA.auth !== 'in') return;
  // sales
  if (prev.sales !== next.sales){
    const pm = byKey(prev.sales, x => x.id);
    (next.sales || []).forEach(x => {
      if (pm.get(x.id) === x) return;
      const isNew = !pm.has(x.id);
      if (x.shop === 'u'){
        const d = uSaleDoc(x); if (isNew) d.createdAt = serverTimestamp();
        q(() => setDoc(doc(db, C.uSales, x.id), d, { merge: true }));
      } else {
        const d = sSaleDoc(x);
        q(() => setDoc(doc(db, C.sSales, x.id), d, { merge: true }));
      }
    });
    const nm = byKey(next.sales, x => x.id);
    (prev.sales || []).forEach(x => { if (!nm.has(x.id)) q(() => deleteDoc(doc(db, x.shop === 'u' ? C.uSales : C.sSales, x.id))); });
  }
  // stock: always change by the difference, so two phones selling at once both count
  if (prev.stock !== next.stock){
    Object.keys(next.stock || {}).forEach(id => {
      const d = (Number(next.stock[id]) || 0) - (Number((prev.stock || {})[id]) || 0);
      if (!d || !_pmap[id]) return;
      q(() => updateDoc(doc(db, _pmap[id].shop === 'u' ? C.uProducts : C.sProducts, id), { remaining: increment(d) }));
    });
  }
  if (prev.prices !== next.prices){
    Object.keys(next.prices || {}).forEach(id => {
      if ((prev.prices || {})[id] === next.prices[id] || !_pmap[id]) return;
      q(() => updateDoc(doc(db, _pmap[id].shop === 'u' ? C.uProducts : C.sProducts, id), { price: Number(next.prices[id]) || 0 }));
    });
  }
  // deliveries received: one restock row per item, plus the product's received total
  if (prev.receipts !== next.receipts){
    const old = new Set(prev.receipts || []);
    (next.receipts || []).forEach(r => {
      if (old.has(r)) return;
      const p = _pmap[r.pid]; if (!p) return;
      const rid = r.rid || ('r' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6));
      r.rid = rid;
      const app = clean({ label: r.label, memo: r.memo, cost: r.cost, batch: r.batch });
      const base = p.shop === 'u'
        ? { productId: r.pid, category: p.cat, categoryLabel: p.name, size: p.size, name: p.name }
        : { productId: r.pid, name: p.name, cat: p.cat };
      q(async () => {
        await setDoc(doc(db, p.shop === 'u' ? C.uRestocks : C.sRestocks, rid), Object.assign(base, {
          date: r.date, time: '', qty: r.qty, createdAt: serverTimestamp(), app }));
        await updateDoc(doc(db, p.shop === 'u' ? C.uProducts : C.sProducts, r.pid), { received: increment(r.qty), receivedAt: serverTimestamp(), lastRestockQty: r.qty });
      });
    });
  }
  syncMap(prev.batches, next.batches, C.batches);
  syncList(prev.demand, next.demand, C.demand, x => x.id);
  syncList(prev.closes, next.closes, C.closes, x => x.date + '_' + x.shop);
  syncList(prev.dMoves, next.dMoves, C.drawer, x => x.id);
  syncList(prev.merges, next.merges, C.merges, x => x.from);
  if (prev.names !== next.names || prev.rosters !== next.rosters){
    q(() => setDoc(doc(db, C.meta, 'students'), { names: clean(next.names || {}), rosters: clean(next.rosters || []) }));
  }
  const S = ['dStart', 'staff', 'lastBackup', 'lastBackupAt'];
  if (S.some(k => prev[k] !== next[k])){
    const o = {}; S.forEach(k => { if (next[k] !== undefined) o[k] = clean(next[k]); });
    q(() => setDoc(doc(db, C.meta, 'settings'), o, { merge: true }));
  }
}
function syncMap(a, b, col){
  if (a === b) return;
  Object.keys(b || {}).forEach(id => { if ((a || {})[id] !== b[id]) q(() => setDoc(doc(db, col, id), clean(b[id]))); });
  Object.keys(a || {}).forEach(id => { if (!(b || {})[id]) q(() => deleteDoc(doc(db, col, id))); });
}
function syncList(a, b, col, key){
  if (a === b) return;
  const am = byKey(a, key), bm = byKey(b, key);
  bm.forEach((x, k) => { if (am.get(k) !== x) q(() => setDoc(doc(db, col, String(k)), clean(x))); });
  am.forEach((x, k) => { if (!bm.has(k)) q(() => deleteDoc(doc(db, col, String(k)))); });
}

// used by the data-move script and the test-mode copy
export const _internals = { db, collection, doc, getDocs, setDoc, writeBatch, C };
