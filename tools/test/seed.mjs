import { UCAT, SMETA } from '/home/claude/sathaporn-school-shop/js/catalog.js';
const today = new Date(Date.now()+7*3600e3).toISOString().slice(0,10);
const yest = new Date(Date.now()-17*3600e3).toISOString().slice(0,10);
const DB = {uniform_products:{}, stationary_products:{}, uniform_transactions:{}, sales:{}, uniform_restocks:{}, stationary_restocks:{}};
let n=0;
for (const [k,m] of Object.entries(UCAT)) for (const sz of m.sizes){ n++; DB.uniform_products[k+'_'+sz] = {category:k, categoryLabel:m.name, categorySub:'ประถม', size:sz, price:200+10*(n%5), received:20, remaining:(n*7)%23}; }
let i=0; for (const name of Object.keys(SMETA)){ i++; DB.stationary_products['p'+i] = {name, price:5+5*(i%3), cat:'เครื่องเขียน', received:50, remaining:(i*11)%40, order:i}; }
DB.stationary_products['p99'] = {name:'ปากกาน้ำเงิน ธรรมดา', price:5, cat:'เครื่องเขียน', received:150, remaining:48, order:99};
DB.uniform_transactions['t1'] = {date:today, time:'09:12', seller:'Porcelain', cashAmount:220, transferAmount:0, studentCode:'6900', items:[{productId:'shirt_f_34', category:'shirt_f', categoryLabel:'เสื้อนักเรียนหญิง', size:'34', price:220, qty:1, subtotal:220}], total:220, createdAt:{}};
DB.uniform_transactions['t2'] = {date:yest, time:'13:40', seller:'ครูอ้อย', cashAmount:0, transferAmount:430, studentCode:'7018', items:[{productId:'shirt_m_34', category:'shirt_m', categoryLabel:'เสื้อนักเรียนชาย', size:'34', price:215, qty:2, subtotal:430}], total:430, createdAt:{}};
DB.sales['1764040000000-abcde'] = {id:'1764040000000-abcde', dateKey:today, dateLabel:'x', items:[{id:'p1', name:'x', price:15, qty:2},{id:'p99', name:'ปากกาน้ำเงิน ธรรมดา', price:5, qty:3}], total:45, ts:new Date().toISOString(), seller:'ครูอ้อย'};
DB.uniform_restocks['r1'] = {date:yest, time:'15:20', productId:'shirt_f_34', category:'shirt_f', categoryLabel:'เสื้อนักเรียนหญิง', size:'34', qty:6, receivedTotal:26, createdAt:{}};
DB.shop_meta = {students:{names:{'6900':'ด.ช. ธนกฤต เทพรัตนวิชัย','7018':'ด.ช. ชนาวินทร์ ศรีนอก'}, rosters:[{id:'r2', term:'2/2569', eff:'2026-11-01', classes:{'6900':'อ.3','7018':'อ.1'}}]}, settings:{dStart:{date:yest, amt:2871}, staff:[]}};
DB.shop_closes = {}; DB.shop_closes[yest+'_s'] = {date:yest, at:'16:30', shop:'s', opening:2700, sales:171, out:0, in:0, expected:2871, counted:2871, diff:0, reason:'', note:'', by:'Porcelain'};
console.log(JSON.stringify(DB));
