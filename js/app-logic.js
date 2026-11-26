import { DCLogic } from './runtime.js';
import { DATA } from './data.js';

var TODAY = DATA.today;
var THM = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
function thDate(iso){ if(!iso) return ''; var p = iso.split('-').map(Number); return p[2]+' '+THM[p[1]-1]+' '+(p[0]+543); }
function thShort(iso){ var p = iso.split('-').map(Number); return p[2]+' '+THM[p[1]-1]; }
function baht(n){ return '฿'+Math.round(Number(n)||0).toLocaleString('en-US'); }
function addDays(iso,n){ var d = new Date(iso+'T00:00:00Z'); d.setUTCDate(d.getUTCDate()+n); return d.toISOString().slice(0,10); }

function U(id, groups, gender, part, tone, icon, name, size, price, stock){
  return {id:id, shop:'u', groups:groups, gender:gender, part:part, tone:tone, icon:icon, name:name, size:size, price:price, stock:stock};
}
var PRODUCTS = DATA.products;
function S(id, cat, icon, name, price, stock, grp, short, sub, sg, lv){ return {id:id, shop:'s', cat:cat, icon:icon, name:name, size:'', price:price, stock:stock, grp:grp||'', short:short||'', sub:sub||'', sg:sg||'', lv:lv||''}; }
function sFlags(k, p){ return {iPencil:k==='pencil', iPen:k==='pen', iEraser:k==='eraser', iBook:k==='book', iPaper:k==='paper', iRuler:k==='ruler', iGlue:k==='glue', iScissors:k==='scissors', iColor:k==='color', icoM: icoStyle(p ? icoKey(p) : k)}; }
var PMAP = {}; PRODUCTS.forEach(function(p){ PMAP[p.id] = p; });
var ALLP = PRODUCTS;
function bn(v){ v = String(v == null ? '' : v); var u = v.charAt(0) === '฿'; return {n: u ? v.slice(1) : v, u: u}; }
function itemLabel(id){ var p = PMAP[id]; return p.size ? p.name+' #'+p.size : p.name; }

var BASE_NAMES = DATA.names;
var R1 = DATA.rosters[DATA.rosters.length - 1] || {id:"r0", term:"", eff:DATA.today, classes:{}};
var NEW_FILE = {fileName:'รายชื่อนักเรียน ภาคเรียน 2-2569.xlsx', classes:{'10231':'ป.4/1','10232':'ป.4/1','10318':'ป.5/2','10405':'ป.6/1','20112':'อ.2','30017':'ป.4/2'}, names:{'30017':'ด.ช. วรวิช นาคดี'}};

function gx(n){ n = n || ''; return /หญิง/.test(n) ? ' gf' : (/ชาย/.test(n) ? ' gm' : ''); }

var ICO_D = {"notebook":["M10 11H13.5M10 7H17","M13 2H12.5C8.72877 2 6.84314 2 5.67157 3.17158C4.5 4.34315 4.5 6.22877 4.5 10V14C4.5 17.7712 4.5 19.6569 5.67157 20.8284C6.84315 22 8.72876 22 12.5 22H13C16.7712 22 18.6569 22 19.8284 20.8284C21 19.6569 21 17.7712 21 14V10C21 6.22877 21 4.34315 19.8284 3.17157C18.6569 2 16.7712 2 13 2Z","M6 6H3M6 12H3M6 18H3"],"pencil":["M6 22V13.6944C6 12.1287 6 11.3459 6.21454 10.6077C6.42908 9.86948 6.84589 9.21812 7.6795 7.91542L10.3359 3.76419C11.0885 2.58806 11.4648 2 12 2C12.5352 2 12.9115 2.58806 13.6641 3.76419L16.3205 7.91542C17.1541 9.21812 17.5709 9.86948 17.7855 10.6077C18 11.3459 18 12.1287 18 13.6944V22","M7 11C7.63152 11.3231 8.4887 11.9732 9.28009 11.9991C10.2988 12.0324 10.9868 11.1372 12 11.1372C13.0132 11.1372 13.7012 12.0324 14.7199 11.9991C15.5113 11.9732 16.3685 11.3231 17 11","M12 12L12 22","M10 5H14"],"pen":["M3.49977 18.9853V20.5H5.01449C6.24074 20.5 6.85387 20.5 7.40518 20.2716C7.9565 20.0433 8.39004 19.6097 9.25713 18.7426L19.1211 8.87868C20.0037 7.99612 20.4449 7.55483 20.4937 7.01325C20.5018 6.92372 20.5018 6.83364 20.4937 6.74411C20.4449 6.20253 20.0037 5.76124 19.1211 4.87868C18.2385 3.99612 17.7972 3.55483 17.2557 3.50605C17.1661 3.49798 17.0761 3.49798 16.9865 3.50605C16.4449 3.55483 16.0037 3.99612 15.1211 4.87868L5.25713 14.7426C4.39004 15.6097 3.9565 16.0433 3.72813 16.5946C3.49977 17.1459 3.49977 17.759 3.49977 18.9853Z","M13.5 6.5L17.5 10.5"],"eraser":["M8.73792 7.78021L4.84693 11.7574C3.67722 12.953 3.09236 13.5508 3.01263 14.2802C2.99579 14.4343 2.99579 14.5899 3.01263 14.744C3.09236 15.4733 3.67722 16.0711 4.84693 17.2668L4.99601 17.4191C5.62049 18.0575 5.93274 18.3766 6.30638 18.5911C6.5236 18.7157 6.75482 18.8134 6.99505 18.882C7.40827 19 7.85149 19 8.73792 19C9.62436 19 10.0676 19 10.4808 18.882C10.721 18.8134 10.9522 18.7157 11.1695 18.5911C11.5431 18.3766 11.8554 18.0575 12.4798 17.4191L15.3239 14.5121M8.73792 7.78021L12.3199 4.12313C13.7065 2.70754 14.3997 1.99974 15.2627 2C16.1256 2.00026 16.8185 2.70846 18.2042 4.12487L18.9473 4.8845C20.3159 6.28342 21.0002 6.98288 21 7.85008C20.9997 8.71728 20.315 9.41633 18.9456 10.8144L15.3239 14.5121M8.73792 7.78021L15.3239 14.5121","M10 22L21 22"],"ruler":["M17.5 10.5L19.5 12.5M14 14L16 16M10.5 17.5L12.5 19.5","M10.5355 4.67767C11.9002 3.31296 12.5826 2.6306 13.3438 2.3153C14.3587 1.8949 15.4991 1.8949 16.5141 2.3153C17.2753 2.6306 17.9576 3.31296 19.3223 4.67767C20.687 6.04238 21.3694 6.72474 21.6847 7.48594C22.1051 8.50088 22.1051 9.64126 21.6847 10.6562C21.3694 11.4174 20.687 12.0998 19.3223 13.4645L13.4645 19.3223C12.0998 20.687 11.4174 21.3694 10.6562 21.6847C9.64126 22.1051 8.50088 22.1051 7.48594 21.6847C6.72474 21.3694 6.04238 20.687 4.67767 19.3223C3.31296 17.9576 2.6306 17.2753 2.3153 16.5141C1.8949 15.4991 1.8949 14.3587 2.3153 13.3438C2.6306 12.5826 3.31296 11.9002 4.67767 10.5355L10.5355 4.67767Z"],"shorts":["M10.9876 15.1231L10.2566 20.2254C10.0233 21.8538 9.63333 22.1467 8.0011 21.9433L3.79203 21.4186C2.20615 21.221 1.90147 20.854 2.0244 19.2605L2.76037 9.72027C2.81727 8.98263 2.84572 8.61382 2.91847 8.25272C3.10263 7.33855 3.61242 6.53567 3.82223 5.63542C3.94435 5.11145 3.90106 4.55591 3.90106 4.01893C3.90106 2.34611 4.25316 2 5.92579 2H18.0742C19.7468 2 20.0989 2.34611 20.0989 4.01893C20.0989 4.55591 20.0557 5.11145 20.1778 5.63542C20.3876 6.53567 20.8974 7.33855 21.0815 8.25272C21.7973 11.8057 21.6973 15.6526 21.9756 19.2605C22.0985 20.854 21.7938 21.221 20.208 21.4186L15.9989 21.9433C14.3667 22.1467 13.9767 21.8538 13.7434 20.2254L13.0124 15.1231","M9 13.5C9.76053 14.2605 10.1976 15 11.3284 15H12.6716C13.8024 15 14.2395 14.2605 15 13.5","M4 5H20","M3 11C7 11.0068 7 7.00225 7 5","M17 5C17 7.00225 17 11.0068 21 11"],"pepants":["M18.1604 5V4C18.1604 3.05719 18.1604 2.58579 17.8596 2.29289C17.5589 2 17.0749 2 16.1069 2H7.8931C6.92509 2 6.44109 2 6.14037 2.29289C5.83965 2.58579 5.83965 3.05719 5.83965 4V5M18.1604 5L20.8152 19.6524C21.0124 20.7411 21.1111 21.2855 20.8033 21.6427C20.4956 22 19.928 22 18.793 22H17.6235C16.8926 22 16.5272 22 16.2552 21.8044C15.9833 21.6088 15.8758 21.2686 15.6609 20.5882L13.9626 15.2126C13.1598 12.6711 12.7583 11.4004 12 11.4004C11.2417 11.4004 10.8402 12.6711 10.0374 15.2126L8.33912 20.5882C8.12417 21.2686 8.01669 21.6088 7.74476 21.8044C7.47284 22 7.10738 22 6.37647 22H5.20702C4.07196 22 3.50443 22 3.19668 21.6427C2.88893 21.2855 2.98756 20.7411 3.18482 19.6524L5.83965 5M18.1604 5H5.83965","M10 7.77778L12 5L15 10"],"schoolshirt":["M6 9.5v7.2c0 1.8 0 2.7.59 3.27 1.31 1.26 9.27 1.5 10.82 0 .59-.57.59-1.47.59-3.27V9.5","M5.74 12 3.04 9.39C2.35 8.72 2 8.38 2 7.96s.35-.75 1.04-1.43l2-1.94c.35-.33.52-.5.73-.62.2-.12.44-.2.9-.35L9 2.75","M18.26 12l2.7-2.61c.7-.67 1.04-1.01 1.04-1.43s-.35-.75-1.04-1.43l-2-1.94c-.35-.33-.52-.5-.73-.62-.2-.12-.44-.2-.9-.35L15 2.75","M9 2.75l3 3.75-2.15 1.9c-.33.28-.78.17-.93-.22L8 4.5M15 2.75l-3 3.75 2.15 1.9c.33.28.78.17.93-.22L16 4.5","M9 2.75c1 .33 2 .5 3 .5s2-.17 3-.5","M12 6.5v14.6","M13.4 11h.01M13.4 15h.01M13.4 19h.01"],"polo":["M5.99756 12V17.6841C5.99756 19.4952 5.99756 20.4008 6.58285 20.9635C7.24789 21.6028 9.6154 21.9785 11.9925 21.9991C14.3085 22.0192 16.6337 21.7022 17.4021 20.9635C17.9874 20.4008 17.9874 19.4952 17.9874 17.6841V12","M5.72 14.02c-.86-.31-2.49-.58-3.22-1.32-1.01-1.01-.36-2.38.88-4.96.56-1.16 1.62-2.02 2.87-2.36.17-.04.31-.15.41-.3L7.94 3.09","M18.25 14.03c.87-.31 2.52-.6 3.25-1.34 1.01-1.01.39-2.54-1-5.07-.56-1.16-1.57-1.9-2.82-2.24-.17-.05-.31-.15-.41-.3l-1.32-1.98","M7.94 3.09C9.2 2.45 10.5 2.25 12 2.25s2.8.2 4.06.84","M8.6 2.75 7.85 6c-.1.45.22.78.65.62L12 5.5l3.5 1.12c.43.16.75-.17.65-.62L15.4 2.75","M10.75 5.9v3.6c0 .55.45 1 1 1h.5c.55 0 1-.45 1-1V5.9"],"sketchbook":["M6.75 5h10.5c1.705 0 2.75 .38 2.75 2.75v10.75c0 1.705-.38 2.75-2.75 2.75h-10.5c-1.705 0-2.75-.38-2.75-2.75v-10.75c0-1.705 .38-2.75 2.75-2.75Z","M8 2.75v4M12 2.75v4M16 2.75v4","M7.25 17.5l3.25-3.75 2.25 2.5 1.6-1.6 2.4 2.85","M16.25 10.5a1.25 1.25 0 1 1-2.5 0 1.25 1.25 0 0 1 2.5 0Z"],"cardsleeve":["M5 6.25c0-1.65 0-2.47.51-2.98S6.85 2.75 8.5 2.75h7c1.65 0 2.47 0 2.98.52.52.5.52 1.33.52 2.98v12.5c0 1.65 0 2.47-.52 2.98-.5.52-1.33.52-2.98.52h-7c-1.65 0-2.47 0-2.98-.52C5 21.22 5 20.4 5 18.75V6.25Z","M10.5 5.75h3","M9.25 8.75h5.5c0.9299999999999999 0 1.5 .38 1.5 1.5v7.25c0 0.9299999999999999-.38 1.5-1.5 1.5h-5.5c-0.9299999999999999 0-1.5-.38-1.5-1.5v-7.25c0-0.9299999999999999 .38-1.5 1.5-1.5Z","M13.25 12.25a1.25 1.25 0 1 1-2.5 0 1.25 1.25 0 0 1 2.5 0Z","M9.9 16.5c.4-.9 1.15-1.4 2.1-1.4s1.7.5 2.1 1.4"],"lanyard":["M6 2.75v6.5c0 3.2 2.7 5.75 6 5.75s6-2.55 6-5.75v-6.5","M9 2.75v6c0 1.75 1.35 3.25 3 3.25s3-1.5 3-3.25v-6","M11.25 15h1.5c0.62 0 1 .38 1 1v0.5c0 0.62-.38 1-1 1h-1.5c-0.62 0-1-.38-1-1v-0.5c0-0.62 .38-1 1-1Z","M10.75 17.5h2.5v2.25c0 .9-.55 1.5-1.25 1.5s-1.25-.6-1.25-1.5V17.5Z"],"skirt":["M7.35 2.75h9.3c0.992 0 1.6 .38 1.6 1.6v0.7999999999999998c0 0.992-.38 1.6-1.6 1.6h-9.3c-0.992 0-1.6-.38-1.6-1.6v-0.7999999999999998c0-0.992 .38-1.6 1.6-1.6Z","M6 6.75 3.15 17.9c-.33 1.25.06 2.03 1.3 2.4 2.1.62 4.73.95 7.55.95s5.45-.33 7.55-.95c1.24-.37 1.63-1.15 1.3-2.4L18 6.75","M12 6.75v14.5M9 6.75 7.9 20.7M15 6.75l1.1 13.95"],"cap":["M3 15.5c0-5 3.6-9 8-9s7.5 3.6 8 8.25","M3 15.5h16c1.65 0 2.75.5 2.75 1.25S20.65 18 19 18H5.5C4.12 18 3 16.88 3 15.5Z","M11 6.5c-1.5 2.4-2.25 5.4-2.25 9M11 6.5c1.25 2.6 1.75 5.4 1.75 8.25","M11 6.5V5.25"],"scarf":["M2.9 5.9c-.55-.95.13-2.15 1.2-2.15h15.8c1.07 0 1.75 1.2 1.2 2.15l-7.9 14.3c-.53.95-1.87.95-2.4 0Z","M4.1 7.75h15.8","M12 10l2.25 2.25L12 14.5l-2.25-2.25Z"],"wogleather":["M6 5.5c0-.97 2.69-1.75 6-1.75s6 .78 6 1.75v13c0 .97-2.69 1.75-6 1.75s-6-.78-6-1.75v-13Z","M6 5.5c0 .97 2.69 1.75 6 1.75s6-.78 6-1.75","M14.5 13.25a2.5 2.5 0 1 1-5 0 2.5 2.5 0 0 1 5 0Z","M15.5 7.1v11.9"],"wogscout":["M21.25 12a9.25 9.25 0 1 1-18.5 0 9.25 9.25 0 0 1 18.5 0Z","M18.75 12a6.75 6.75 0 1 1-13.5 0 6.75 6.75 0 0 1 13.5 0Z","M12 7.25c-.9.95-1.35 1.95-1.35 3 0 .95.35 1.8.9 2.45h.9c.55-.65.9-1.5.9-2.45 0-1.05-.45-2.05-1.35-3Z","M11.1 12.7c-1.45 0-2.7-.75-3-1.95M12.9 12.7c1.45 0 2.7-.75 3-1.95","M10 13.85h4M12 13.85v1.9M10.75 17c.35-.6.75-1 1.25-1s.9.4 1.25 1","M12 2.75v2.5M12 18.75v2.5M2.75 12h2.5M18.75 12h2.5"],"woguide":["M12 2.75c2.35 0 4.25 1.9 4.25 4.25 0 .4-.05.8-.16 1.17 2.2-.2 4.16 1.62 4.16 3.83 0 2.2-1.8 4-4 4-.75 0-1.45-.2-2.05-.57L13.5 18.5l1 2.75h-5l1-2.75-.7-3.07c-.6.37-1.3.57-2.05.57-2.2 0-4-1.8-4-4 0-2.21 1.96-4.03 4.16-3.83A4.25 4.25 0 0 1 12 2.75Z","M12 5c-.85.8-1.3 1.65-1.3 2.5 0 .55.2 1.05.55 1.45M12 5c.85.8 1.3 1.65 1.3 2.5 0 .55-.2 1.05-.55 1.45M12 6.75v2.2","M10.6 9.4h2.8L12 15.5Z","M7.25 12h2M16.75 12h-2"],"rope":["M8.5 3.75c0-.55 1.57-1 3.5-1s3.5.45 3.5 1v9.5c0 .55-1.57 1-3.5 1s-3.5-.45-3.5-1v-9.5Z","M8.5 6.5c0 .55 1.57 1 3.5 1s3.5-.45 3.5-1M8.5 9.5c0 .55 1.57 1 3.5 1s3.5-.45 3.5-1","M9.75 14.1c-1.75 2.05-2.75 3.95-2.75 5.35 0 1.15.75 1.8 1.65 1.8 1.05 0 1.85-1.6 2.35-6.9M14.25 14.1c1.75 2.05 2.75 3.95 2.75 5.35 0 1.15-.75 1.8-1.65 1.8-1.05 0-1.85-1.6-2.35-6.9"],"guidepin":["M12 2.75c1.8 0 3.25 1.45 3.25 3.25 0 .3-.04.6-.12.88 1.75-.2 3.37 1.18 3.37 2.97 0 1.65-1.35 3-3 3-.55 0-1.07-.15-1.52-.42L13.25 14.5h-2.5l-.73-2.07c-.45.27-.97.42-1.52.42-1.65 0-3-1.35-3-3 0-1.8 1.62-3.17 3.37-2.97-.08-.28-.12-.58-.12-.88 0-1.8 1.45-3.25 3.25-3.25Z","M12 5.25v6.75M8.5 9.75h2M15.5 9.75h-2","M7.5 15.75h9c.55 0 1 .45 1 1v2.75c0 .55-.45 1-1 1h-9c-.55 0-1-.45-1-1v-2.75c0-.55.45-1 1-1Z","M6.5 17H3.25l1.5 1.88-1.5 1.87H6.5M17.5 17h3.25l-1.5 1.88 1.5 1.87H17.5"],"angelpin":["M10.5 3.1c.93-.53 2.07-.53 3 0l5 2.88c.93.54 1.5 1.53 1.5 2.6v5.84c0 1.07-.57 2.06-1.5 2.6l-5 2.88c-.93.53-2.07.53-3 0l-5-2.88c-.93-.54-1.5-1.53-1.5-2.6V8.58c0-1.07.57-2.06 1.5-2.6Z","M13.25 7.75a1.25 1.25 0 1 1-2.5 0 1.25 1.25 0 0 1 2.5 0Z","M12 9v4.75M9 11.25c1.05-.5 2.05-.75 3-.75s1.95.25 3 .75M10 16.5l2-2.75 2 2.75"],"sharpener":["M4 6.75c0-1.1.75-1.85 1.75-1.85.9 0 1.55.5 2.4.75 1.15.35 2.45.5 3.85.5s2.7-.15 3.85-.5c.85-.25 1.5-.75 2.4-.75 1 0 1.75.75 1.75 1.85v10.5c0 1.1-.75 1.85-1.75 1.85-.9 0-1.55-.5-2.4-.75-1.15-.35-2.45-.5-3.85-.5s-2.7.15-3.85.5c-.85.25-1.5.75-2.4.75-1 0-1.75-.75-1.75-1.85V6.75Z","M8 8.25h8c0.775 0 1.25 .38 1.25 1.25v5c0 0.775-.38 1.25-1.25 1.25h-8c-0.775 0-1.25-.38-1.25-1.25v-5c0-0.775 .38-1.25 1.25-1.25Z","M6.75 12h10.5","M12.85 10.1a0.85 0.85 0 1 1-1.7 0 0.85 0.85 0 0 1 1.7 0Z"],"correction":["M8 7.75V5.1c0-1.3.95-2.35 2.25-2.35h3.5c1.3 0 2.25 1.05 2.25 2.35v2.65","M12 3.75v1.5M10.75 7.75v-.75c0-.55.45-1 1-1h.5c.55 0 1 .45 1 1v.75","M9.25 7.75h5.5c1.085 0 1.75 .38 1.75 1.75v7.5c0 1.085-.38 1.75-1.75 1.75h-5.5c-1.085 0-1.75-.38-1.75-1.75v-7.5c0-1.085 .38-1.75 1.75-1.75Z","M12 10.25h0c1.395 0 2.25 .38 2.25 2.25v1.5c0 1.395-.38 2.25-2.25 2.25h-0c-1.395 0-2.25-.38-2.25-2.25v-1.5c0-1.395 .38-2.25 2.25-2.25Z","M8.5 18.75v.75c0 .97.78 1.75 1.75 1.75h3.5c.97 0 1.75-.78 1.75-1.75v-.75"],"dress":["M9 2.75v3.1c0 .85-.4 1.6-1.05 2.1L6.6 9.05c-.62.5-.92 1.28-.8 2.06l1.2 8.04c.15 1.1.9 2.1 2.25 2.1h5.5c1.35 0 2.1-1 2.25-2.1l1.2-8.04c.12-.78-.18-1.56-.8-2.06l-1.35-1.1C15.4 7.45 15 6.7 15 5.85v-3.1","M9 4.5c.9.55 1.9.85 3 .85s2.1-.3 3-.85","M6.5 12.25h11","M12 7.5h.01M12 9.75h.01"],"kgset":["M9 2.75 6.4 3.85c-.45.2-.8.55-1 1L4.5 6.9l2 1.5 1-.65V12h9V7.75l1 .65 2-1.5-.9-2.05c-.2-.45-.55-.8-1-1L15 2.75c-.65.8-1.7 1.3-3 1.3s-2.35-.5-3-1.3Z","M7.6 14.25h8.8c.4 0 .7.3.75.68l.6 5.45c.05.47-.3.87-.77.87h-2.6c-.33 0-.62-.2-.73-.52L12 17.5l-1.65 3.23c-.11.32-.4.52-.73.52h-2.6c-.47 0-.82-.4-.77-.87l.6-5.45c.05-.38.35-.68.75-.68Z"]};
Object.assign(ICO_D, {"db":["M3 12C3 7.75736 3 5.63604 4.31802 4.31802C5.63604 3 7.75736 3 12 3C16.2426 3 18.364 3 19.682 4.31802C21 5.63604 21 7.75736 21 12C21 16.2426 21 18.364 19.682 19.682C18.364 21 16.2426 21 12 21C7.75736 21 5.63604 21 4.31802 19.682C3 18.364 3 16.2426 3 12Z","M3 12H21","M11 7.5L17 7.5","M7.125 7.5H7M7.25 7.5C7.25 7.63807 7.13807 7.75 7 7.75C6.86193 7.75 6.75 7.63807 6.75 7.5C6.75 7.36193 6.86193 7.25 7 7.25C7.13807 7.25 7.25 7.36193 7.25 7.5Z","M11 16.5L17 16.5","M7.125 16.5H7M7.25 16.5C7.25 16.6381 7.13807 16.75 7 16.75C6.86193 16.75 6.75 16.6381 6.75 16.5C6.75 16.3619 6.86193 16.25 7 16.25C7.13807 16.25 7.25 16.3619 7.25 16.5Z"],"receipt":["M4 18.6458V8.05426C4 5.20025 4 3.77325 4.87868 2.88663C5.75736 2 7.17157 2 10 2H14C16.8284 2 18.2426 2 19.1213 2.88663C20 3.77325 20 5.20025 20 8.05426V18.6458C20 20.1575 20 20.9133 19.538 21.2108C18.7831 21.6971 17.6161 20.6774 17.0291 20.3073C16.5441 20.0014 16.3017 19.8485 16.0325 19.8397C15.7417 19.8301 15.4949 19.9768 14.9709 20.3073L13.06 21.5124C12.5445 21.8374 12.2868 22 12 22C11.7132 22 11.4555 21.8374 10.94 21.5124L9.02913 20.3073C8.54415 20.0014 8.30166 19.8485 8.03253 19.8397C7.74172 19.8301 7.49493 19.9768 6.97087 20.3073C6.38395 20.6774 5.21687 21.6971 4.46195 21.2108C4 20.9133 4 20.1575 4 18.6458Z","M11 11H8","M14 7L8 7"],"box":["M12 22C11.1818 22 10.4002 21.6698 8.83693 21.0095C4.94564 19.3657 3 18.5438 3 17.1613C3 16.7742 3 10.0645 3 7M12 22C12.8182 22 13.5998 21.6698 15.1631 21.0095C19.0544 19.3657 21 18.5438 21 17.1613V7M12 22L12 11.3548","M8.32592 9.69138L5.40472 8.27785C3.80157 7.5021 3 7.11423 3 6.5C3 5.88577 3.80157 5.4979 5.40472 4.72215L8.32592 3.30862C10.1288 2.43621 11.0303 2 12 2C12.9697 2 13.8712 2.4362 15.6741 3.30862L18.5953 4.72215C20.1984 5.4979 21 5.88577 21 6.5C21 7.11423 20.1984 7.5021 18.5953 8.27785L15.6741 9.69138C13.8712 10.5638 12.9697 11 12 11C11.0303 11 10.1288 10.5638 8.32592 9.69138Z","M6 12L8 13","M17 4L7 9"],"dl":["M16.9504 12.1817C17.1981 12.814 16.5076 13.5726 15.1267 15.0899C13.6702 16.6902 12.9201 17.4904 12 17.5C11.0799 17.4904 10.3298 16.6902 8.87331 15.0899C7.49239 13.5726 6.80193 12.814 7.04964 12.1817C7.05868 12.1586 7.06851 12.1359 7.0791 12.1135C7.34928 11.542 8.24477 11.5029 10 11.5002V4.99998C10 4.53501 10 4.30253 10.0511 4.11179C10.1898 3.59414 10.5941 3.1898 11.1118 3.05111C11.3025 3 11.535 3 12 3C12.4649 3 12.6974 3 12.8882 3.05111C13.4058 3.1898 13.8102 3.59414 13.9489 4.11179C14 4.30253 14 4.53501 14 4.99998V11.5002C15.7552 11.5029 16.6507 11.542 16.9209 12.1135C16.9315 12.1359 16.9413 12.1586 16.9504 12.1817Z","M5.00006 21H19.0001"],"people":["M18.5 7C18.5 9.20914 16.7091 11 14.5 11C12.2909 11 10.5 9.20914 10.5 7C10.5 4.79086 12.2909 3 14.5 3C16.7091 3 18.5 4.79086 18.5 7Z","M8 11C6.34315 11 5 9.65685 5 8C5 6.34315 6.34315 5 8 5","M14.5 14C9.8125 14 7 16.5 7 19C7 20.1046 7.83947 21 8.875 21H20.125C21.1605 21 22 20.1046 22 19C22 16.5 19.1875 14 14.5 14Z","M4 19.9992H3.5C2.67157 19.9992 2 19.2317 2 18.2849C2 16.673 3.27307 15.0612 5.5 14.3613"],"merge":["M7 20C8.10457 20 9 19.1046 9 18C9 16.8954 8.10457 16 7 16C5.89543 16 5 16.8954 5 18C5 19.1046 5.89543 20 7 20Z","M7 8C8.10457 8 9 7.10457 9 6C9 4.89543 8.10457 4 7 4C5.89543 4 5 4.89543 5 6C5 7.10457 5.89543 8 7 8Z","M17 14C18.1046 14 19 13.1046 19 12C19 10.8954 18.1046 10 17 10C15.8954 10 15 10.8954 15 12C15 13.1046 15.8954 14 17 14Z","M7.02116 8.2793V15.4073M14.4113 12.0047L10.0193 12.0048C8.92158 12.0048 6.86182 11.1254 7.01818 8.78001"],"file":["M8 7L16 7","M8 11L12 11","M13 21.5V21C13 18.1716 13 16.7574 13.8787 15.8787C14.7574 15 16.1716 15 19 15H19.5M20 13.3431V10C20 6.22876 20 4.34315 18.8284 3.17157C17.6569 2 15.7712 2 12 2C8.22877 2 6.34315 2 5.17157 3.17157C4 4.34314 4 6.22876 4 10L4 14.5442C4 17.7892 4 19.4117 4.88607 20.5107C5.06508 20.7327 5.26731 20.9349 5.48933 21.1139C6.58831 22 8.21082 22 11.4558 22C12.1614 22 12.5141 22 12.8372 21.886C12.9044 21.8623 12.9702 21.835 13.0345 21.8043C13.3436 21.6564 13.593 21.407 14.0919 20.9081L18.8284 16.1716C19.4065 15.5935 19.6955 15.3045 19.8478 14.9369C20 14.5694 20 14.1606 20 13.3431Z"],"alert":["M13.9248 21H10.0752C5.44476 21 3.12955 21 2.27636 19.4939C1.42317 17.9879 2.60736 15.9914 4.97574 11.9985L6.90057 8.75333C9.17559 4.91778 10.3131 3 12 3C13.6869 3 14.8244 4.91777 17.0994 8.75332L19.0243 11.9985C21.3926 15.9914 22.5768 17.9879 21.7236 19.4939C20.8704 21 18.5552 21 13.9248 21Z","M12 9V13","M12.125 16.75H12M12.25 16.75C12.25 16.8881 12.1381 17 12 17C11.8619 17 11.75 16.8881 11.75 16.75C11.75 16.6119 11.8619 16.5 12 16.5C12.1381 16.5 12.25 16.6119 12.25 16.75Z"]});
Object.assign(ICO_D, {"undo":["M12 21C16.9706 21 21 16.9706 21 12C21 7.02944 16.9706 3 12 3C8.66873 3 5.76018 4.80989 4.20404 7.5","M3 3V4.27816C3 6.47004 3 7.56599 3.70725 8.16512C4.4145 8.76425 5.49553 8.58408 7.6576 8.22373L9 8"],"cash":["M14.4998 12.001C14.4998 13.3817 13.3805 14.501 11.9998 14.501C10.6191 14.501 9.49982 13.3817 9.49982 12.001C9.49982 10.6203 10.6191 9.50098 11.9998 9.50098C13.3805 9.50098 14.4998 10.6203 14.4998 12.001Z","M21.1329 5.6773C20.1903 5.38518 18.4794 5.00098 16 5.00098C11.4209 5.00098 10.1967 6.67844 3.87798 5.42458C2.92079 5.23464 2 5.94628 2 6.92214V16.9392C2 17.6275 2.47265 18.232 3.1448 18.3802C10.1096 19.9161 11.2491 18.1104 16 18.1104C17.5107 18.1104 18.7361 18.253 19.6762 18.4305C20.7719 18.6374 22 17.7984 22 16.6833V6.92803C22 6.35987 21.6756 5.84549 21.1329 5.6773Z","M2.00049 14.4971C4.20963 14.4971 6.00049 16.2879 6.00049 18.4971M22 9.46997C20.2655 9.46997 18.5005 7.54062 18.5005 5.50098M5.92901 5.75514C5.70483 7.40605 3.95133 9.00098 2 9.00098M18.102 18.0993C18.2601 16.3112 20.1009 15.001 22 15.001"],"wait":["M6.5 6a4 4 0 1 0 8 0a4 4 0 1 0 -8 0","M10.5 13C10.1446 13.0097 9.81335 13.0226 9.50082 13.0379C5.8 13.2193 2.78417 16.3058 2.5 20.0002","M18.5 18L17 17.5V15.5M21.5 17.5C21.5 19.9853 19.4853 22 17 22C14.5147 22 12.5 19.9853 12.5 17.5C12.5 15.0147 14.5147 13 17 13C19.4853 13 21.5 15.0147 21.5 17.5Z"],"offline":["M8.5 15.9996C9.36651 15.1331 10.4207 14.642 11.5 14.5264","M5.5 12.5C8.67327 9.32673 12.6221 8.67087 16 10.5324","M2 8.5C8.31579 3.16669 15.6842 3.16668 22 8.49989","M21.0002 13.5L15.0002 19.5M21.0002 19.5L15.0002 13.5"],"help":["M2 12a10 10 0 1 0 20 0a10 10 0 1 0 -20 0","M9.5 9.5C9.5 8.11929 10.6193 7 12 7C13.3807 7 14.5 8.11929 14.5 9.5C14.5 10.3569 14.0689 11.1131 13.4117 11.5636C12.7283 12.0319 12 12.6716 12 13.5","M12.125 16.75H12M12.25 16.75C12.25 16.8881 12.1381 17 12 17C11.8619 17 11.75 16.8881 11.75 16.75C11.75 16.6119 11.8619 16.5 12 16.5C12.1381 16.5 12.25 16.6119 12.25 16.75Z"],"search":["M17 17L21 21","M19 11C19 6.58172 15.4183 3 11 3C6.58172 3 3 6.58172 3 11C3 15.4183 6.58172 19 11 19C15.4183 19 19 15.4183 19 11Z"],"transfer":["M19 9H6.65856C5.65277 9 5.14987 9 5.02472 8.69134C4.89957 8.38268 5.25517 8.01942 5.96637 7.29289L8.21091 5","M5 15H17.3414C18.3472 15 18.8501 15 18.9753 15.3087C19.1004 15.6173 18.7448 15.9806 18.0336 16.7071L15.7891 19"],"support":["M17 10.8045C17 10.4588 17 10.286 17.052 10.132C17.2032 9.68444 17.6018 9.51076 18.0011 9.32888C18.45 9.12442 18.6744 9.02219 18.8968 9.0042C19.1493 8.98378 19.4022 9.03818 19.618 9.15929C19.9041 9.31984 20.1036 9.62493 20.3079 9.87302C21.2513 11.0188 21.7229 11.5918 21.8955 12.2236C22.0348 12.7334 22.0348 13.2666 21.8955 13.7764C21.6438 14.6979 20.8485 15.4704 20.2598 16.1854C19.9587 16.5511 19.8081 16.734 19.618 16.8407C19.4022 16.9618 19.1493 17.0162 18.8968 16.9958C18.6744 16.9778 18.45 16.8756 18.0011 16.6711C17.6018 16.4892 17.2032 16.3156 17.052 15.868C17 15.714 17 15.5412 17 15.1955V10.8045Z","M7 10.8046C7 10.3694 6.98778 9.97821 6.63591 9.6722C6.50793 9.5609 6.33825 9.48361 5.99891 9.32905C5.55001 9.12458 5.32556 9.02235 5.10316 9.00436C4.43591 8.9504 4.07692 9.40581 3.69213 9.87318C2.74875 11.019 2.27706 11.5919 2.10446 12.2237C1.96518 12.7336 1.96518 13.2668 2.10446 13.7766C2.3562 14.6981 3.15152 15.4705 3.74021 16.1856C4.11129 16.6363 4.46577 17.0475 5.10316 16.996C5.32556 16.978 5.55001 16.8757 5.99891 16.6713C6.33825 16.5167 6.50793 16.4394 6.63591 16.3281C6.98778 16.0221 7 15.631 7 15.1957V10.8046Z","M5 9C5 5.68629 8.13401 3 12 3C15.866 3 19 5.68629 19 9","M19 17V17.8C19 19.5673 17.2091 21 15 21H13"]});
var ICO_OLD = {shirt:'schoolshirt', tee:'polo', pants:'shorts', skirt:'skirt', dress:'dress', set:'kgset', book:'notebook', paper:'sketchbook', glue:'cardsleeve', scissors:'wogscout', color:'cap', pencil:'pencil', pen:'pen', eraser:'eraser', ruler:'ruler', write:'pencil', tools:'wogscout'};
function icoKey(p){ var n = p.name || '';
  if (p.shop === 'u') return p.icon === 'pants' && p.tone === 'pe' ? 'pepants' : (ICO_OLD[p.icon] || 'schoolshirt');
  if (/วาดรูป/.test(n)) return 'sketchbook'; if (/^สมุด/.test(n)) return 'notebook';
  if (/สายห้อย/.test(n)) return 'lanyard'; if (/บัตร/.test(n)) return 'cardsleeve';
  if (/กบเหลา/.test(n)) return 'sharpener'; if (/ดินสอ/.test(n)) return 'pencil'; if (/ปากกา/.test(n)) return 'pen';
  if (/น้ำยาลบ/.test(n)) return 'correction'; if (/ยางลบ/.test(n)) return 'eraser'; if (/บรรทัด/.test(n)) return 'ruler';
  if (/^หมวก/.test(n)) return 'cap'; if (/เข็มติดหมวก/.test(n)) return 'woguide'; if (/เข็มนางฟ้า สำรอง/.test(n)) return 'angelpin'; if (/เข็มนางฟ้า/.test(n)) return 'guidepin';
  if (/ผ้าพันคอ/.test(n)) return 'scarf'; if (/หนัง/.test(n)) return 'wogleather'; if (/วอกเกิ้ลลูกเสือ/.test(n)) return 'wogscout'; if (/วอกเกิ้ล/.test(n)) return 'woguide';
  if (/เชือก/.test(n)) return 'rope'; return ICO_OLD[p.icon] || 'notebook'; }
var ICO_M = {};
function icoStyle(k){ k = ICO_D[k] ? k : (ICO_OLD[k] || 'notebook'); if (ICO_M[k]) return ICO_M[k];
  var svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#000" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">'+ICO_D[k].map(function(d){ return '<path d="'+d+'"/>'; }).join('')+'</svg>';
  var u = "url('data:image/svg+xml,"+encodeURIComponent(svg)+"')";
  return ICO_M[k] = '-webkit-mask-image:'+u+';mask-image:'+u; }
function iconFlags(k, p){ return {iShirt:k==='shirt', iTee:k==='tee', iPants:k==='pants', iSkirt:k==='skirt', iDress:k==='dress', iSet:k==='set', icoM: icoStyle(p ? icoKey(p) : k)}; }
function mkSale(id,date,time,shop,code,items,transfer,seller,tempName,transferTime){
  return {id:id, date:date, time:time, shop:shop, code:code||'', items:items.map(function(x){ return {pid:x[0], qty:x[1], price:PMAP[x[0]].price}; }), transfer:transfer||0, transferTime:transferTime||'', seller:seller||'ครู ก.', tempName:tempName||''};
}
var SEED_SALES = DATA.sales;
var SEED_RECEIPTS = DATA.receipts;
var SEED_BATCHES = DATA.batches;
var DEMO_PO_DUP = {no:'สถพ สซ 69/1002-01', date:'2026-10-19', file:'สถพ สซ 69-1002-01 ร้านธนากรพาณิชย์.pdf', lines:[
  {k:'x1', doc:'เสื้อนักเรียนหญิง ประถม เบอร์ 34', pid:'shirt_f_34', qty:4, cost:175},
  {k:'x2', doc:'เสื้อนักเรียนชาย ประถม เบอร์ 38', pid:'shirt_m_38', qty:6, cost:170},
  {k:'x3', doc:'กระโปรงนักเรียน เบอร์ 30', pid:'skirt_30', qty:5, cost:190}]};
var DEMO_MEMO_OLD = {no:'สถพ ซ 69/0921', date:'2026-09-12', file:'สถพ ซ 69-0921 ขออนุมัติจัดซื้อเครื่องแบบนักเรียน.pdf', lines:[
  {k:'y1', doc:'เสื้อนักเรียนชาย ประถม เบอร์ 38', pid:'shirt_m_38', qty:20, cost:null},
  {k:'y2', doc:'กระโปรงนักเรียน เบอร์ 28', pid:'skirt_28', qty:10, cost:null}]};
function thD2(iso){ if (!iso) return ''; var p = iso.split('-').map(Number); return p[2]+' '+THM[p[1]-1]+' '+String(p[0]+543).slice(-2); }
function shortNo(no){ return (no||'').replace(/^สถพ\s*(?:สซ|[ซบท])\s*/,'').replace(/^PO\s*/,''); }
function docTags(b, memoTxt){
  var tg = [];
  if (!b){ if (memoTxt) tg.push({label:memoTxt, cls:'dtag'}); tg.push({label:'ไม่มีไฟล์', cls:'dtag amber'}); return tg; }
  if (b.voided) tg.push({label:'ยกเลิกแล้ว', cls:'dtag red'});
  if (b.src === 'manual'){
    tg.push({label:'รับเข้าเอง', cls:'dtag grey'});
    if (b.memoNo) tg.push({label:'บันทึก '+shortNo(b.memoNo), cls:'dtag'});
    if (!b.voided) tg.push({label:'ไม่มีไฟล์', cls:'dtag amber'});
    return tg;
  }
  if (b.memo) tg.push({label:'บันทึก '+shortNo(b.memo.no), cls:'dtag'});
  if (b.po) tg.push({label:'PO '+shortNo(b.po.no), cls:'dtag'});
  else if (!b.voided) tg.push({label:'ยังไม่มี PO', cls:'dtag amber'});
  if (b.log && b.log.length && !b.voided) tg.push({label:'แก้ไขแล้ว', cls:'dtag grey'});
  return tg;
}
function batchTitle(b){ return b.memo ? b.memo.no : (b.po ? b.po.no : (b.memoNo || 'รับเข้าเอง')); }
function findDup(batches, receipts, doc, excl){
  var best = null, set = function(o){ if (!best || best.rank < o.rank) best = o; };
  Object.keys(batches).forEach(function(id){ var b = batches[id]; if (b.voided || id === excl) return;
    [b.memo, b.po].forEach(function(d){ if (!d) return; if (d.file === doc.file) set({reason:'file', id:id, rank:3}); else if (d.no === doc.no) set({reason:'number', id:id, rank:2}); });
    var lines = doc.lines || IMP_LINES, rs = receipts.filter(function(r){ return r.batch === id; }), m = 0;
    lines.forEach(function(l){ if (rs.some(function(r){ return r.pid === l.pid && r.qty === Number(l.qty); })) m++; });
    if (m >= 2 && m >= 0.7*Math.min(lines.length, rs.length)) set({reason:'data', id:id, rank:1, m:m, n:lines.length});
  });
  return best;
}
var SEED_DEMAND = DATA.demand;
var PICKUP = 'ห้องธุรการ';   // จุดรับของ ใช้ในข้อความแจ้งนักเรียน
function bahtText(n){
  var D = ['','หนึ่ง','สอง','สาม','สี่','ห้า','หก','เจ็ด','แปด','เก้า'], U = ['','สิบ','ร้อย','พัน','หมื่น','แสน'];
  var rd = function(x, big){ if (x >= 1000000) return rd(Math.floor(x/1000000))+'ล้าน'+rd(x % 1000000, true); var s = String(x), L = s.length, o = '';
    for (var i = 0; i < L; i++){ var g = +s[i], pos = L-i-1; if (!g) continue;
      if (pos === 1 && g === 1) o += 'สิบ'; else if (pos === 1 && g === 2) o += 'ยี่สิบ'; else if (pos === 0 && g === 1 && (L > 1 || big)) o += 'เอ็ด'; else o += D[g]+U[pos]; }
    return o; };
  n = Math.round((Number(n)||0)*100); var b = Math.floor(n/100), st = n % 100;
  return (b ? rd(b) : 'ศูนย์')+'บาท'+(st ? rd(st)+'สตางค์' : 'ถ้วน');
}
var MEMO_RE = /^สถพ\s*[ซบท]\s*(\d{2})\/(\d{1,4})$/;
function asksOf(d){ return d.asks || [{date:d.date, time:d.time, got:d.got||null}]; }
function dayGap(a, b){ return Math.round((Date.parse(b+'T00:00:00Z') - Date.parse(a+'T00:00:00Z'))/86400000); }
function addDemand(list, e){
  var who = (e.code||'') || (e.name||'');
  if (who){
    var ix = -1; list.forEach(function(d, k){ if (ix < 0 && d.pid === e.pid && ((d.code||'') || (d.name||'')) === who && !d.fulfilled && !d.dropped && !d.notified) ix = k; });
    if (ix >= 0){ var out = list.slice(), o = list[ix], h = asksOf(o).concat([{date:e.date, time:e.time, got:e.got||null}]);
      out[ix] = Object.assign({}, o, {got: e.got || o.got, asks:h, asked:h.length, confirmed:e.date}); return out; }
  }
  return [Object.assign({id:'d'+Date.now()+Math.floor(Math.random()*1000)}, e)].concat(list);
}
function nowHM(){ var d = new Date(); return ('0'+d.getHours()).slice(-2)+':'+('0'+d.getMinutes()).slice(-2); }
var IMP_MEMO = {no:'สถพ ซ 69/1007', date:'2026-10-14', file:'สถพ ซ 69-1007 ขออนุมัติจัดซื้อเครื่องแบบนักเรียน.pdf'};
var IMP_PO = {no:'สถพ สซ 69/1007-01', date:'2026-10-13', file:'สถพ สซ 69-1007-01 ร้านธนากรพาณิชย์.pdf'};
var IMP_LINES = [
  {k:'a', doc:'เสื้อนักเรียนหญิง ประถม เบอร์ 40', pid:'shirt_f_40', qty:5, cost:185},
  {k:'b', doc:'เสื้อพละ ประถม ไซส์ M', pid:'pe_shirt_M', qty:6, cost:170},
  {k:'c', doc:'กางเกงพละ ประถม ไซส์ M', pid:'pe_pants_M', qty:6, cost:150},
  {k:'d', doc:'ชุดนักเรียนอนุบาลชาย เบอร์ 3', pid:'kg_boy_3', qty:1, cost:null},
  {k:'e', doc:'ชุดนักเรียนอนุบาลชาย เบอร์ 4', pid:'kg_boy_4', qty:2, cost:null},
  {k:'f', doc:'เสื้อนักเรียนชาย ประถม เบอร์ 52', pid:'shirt_m_52', qty:3, cost:210},
  {k:'g', doc:'กางเกงนักเรียนชาย เบอร์ 46', pid:'pants_m_46', qty:3, cost:null},
  {k:'h', doc:'เสื้อพละ ประถม ไซส์ XXL', pid:'pe_shirt_XXL', qty:3, cost:180},
  {k:'i', doc:'กางเกงพละ ประถม ไซส์ XXL', pid:'pe_pants_XXL', base:'pe_pants_M', size:'XXL', qty:3, cost:160}
];
function saleTotal(s){ return s.items.reduce(function(a,i){ return a+i.qty*i.price; },0); }
function itemsText(s){ return s.items.map(function(i){ return itemLabel(i.pid)+' ×'+i.qty; }).join(', '); }

export class Component extends DCLogic {
  takeInstead(wanted, got){
    var st = this.state;
    if (!(st.stock[got] > 0)){ this.flash('สต็อกไม่พอ'); return; }
    var key = PMAP[got].shop === 'u' ? 'cartU' : 'cartS', c = Object.assign({}, st[key]); c[got] = (c[got]||0) + 1;
    var upd = {wantPop:null, dockOpen:true}; upd[key] = c;
    if (PMAP[got].shop === 'u'){ var cw = Object.assign({}, st.cartWant); cw[got] = wanted; upd.cartWant = cw; }
    else upd.demand = addDemand(st.demand, {date:TODAY, time:nowHM(), pid:wanted, got:got, code:(st.wantCode||'').trim(), name:(st.wantName||'').trim()});
    if (PMAP[got].shop === 'u' && (st.wantCode||'').trim() && !(st.codeInput||'').trim()) upd.codeInput = st.wantCode.trim();
    this.setState(upd);
    this.flash('เพิ่ม '+itemLabel(got)+' แทน '+itemLabel(wanted)+' (จดความต้องการไว้แล้ว)');
  }
  quickAdd(pid){
    var st = this.state, p = PMAP[pid];
    if (!(st.stock[pid] > 0)){ this.flash(itemLabel(pid)+' หมด'); return; }
    if (!st.prices[pid]){ this.flash(itemLabel(pid)+' ยังไม่ตั้งราคา'); return; }
    var c = Object.assign({}, st.cartU), next = (c[pid]||0) + 1;
    if (next > st.stock[pid]){ this.flash('สต็อกไม่พอ'); return; }
    c[pid] = next;
    var kg = p.groups.indexOf('kg') >= 0;
    this.setState({page:'sellU', cartU:c, dockOpen:true, q:'', searchOpen:false, drawerOpen:false, filterOpen:false, uSel:p.name});
    window.scrollTo(0, 0);
    this.flash('เพิ่ม '+itemLabel(pid)+' ลงตะกร้าแล้ว');
  }
  addSaleItem(pid){
    var ed = this.state.saleEdit, items = ed.items.slice(), ix = -1, price = this.state.prices[pid] || PMAP[pid].price;
    items.forEach(function(i, k){ if (i.pid === pid) ix = k; });
    if (ed.swapIx != null){
      var sw = ed.swapIx, old = items[sw];
      if (old.pid === pid){ this.setState({saleEdit: Object.assign({}, ed, {adding:false, addName:'', swapIx:null})}); return; }
      if ((this.state.stock[pid]||0) < old.qty){ this.flash(itemLabel(pid)+' เหลือแค่ '+(this.state.stock[pid]||0)+' ชิ้น ไม่พอเปลี่ยน '+old.qty+' ชิ้น'); return; }
      if (ix >= 0){ items[ix] = Object.assign({}, items[ix], {qty: items[ix].qty + old.qty}); items.splice(sw, 1); }
      else items[sw] = {pid:pid, qty:old.qty, price:price};
      this.setState({saleEdit: Object.assign({}, ed, {items:items, adding:false, addName:'', swapIx:null})});
      this.flash('เปลี่ยน '+itemLabel(old.pid)+' เป็น '+itemLabel(pid)+' แล้ว');
      return;
    }
    if (ix >= 0) items[ix] = Object.assign({}, items[ix], {qty: items[ix].qty + 1});
    else items.push({pid:pid, qty:1, price:price});
    this.setState({saleEdit: Object.assign({}, ed, {items:items, adding:false, addName:''})});
    this.flash('เพิ่ม '+itemLabel(pid)+' แล้ว');
  }
  pickDoc(kind, doc){
    var st = this.state, p = {impLines:null};
    if (kind === 'memo'){ p.memoLoaded = true; p.impMemoDoc = doc; } else { p.poLoaded = true; p.impPoDoc = doc; }
    var d = findDup(st.batches, st.receipts, doc, st.replaceId);
    p.dupAlert = d ? {kind:kind, doc:doc, id:d.id, reason:d.reason, m:d.m, n:d.n} : null;
    this.setState(p);
  }
  constructor(props){
    super(props);
    var stock = {}, prices = {};
    PRODUCTS.forEach(function(p){ stock[p.id] = p.stock; prices[p.id] = p.price; });
    this.state = {
      page:DATA.role === 'owner' ? 'dash' : 'sellU', role:DATA.role, me:DATA.me, auth:DATA.auth, q:'', searchOpen:false, period:'week',
      cartU:{}, cartS:{}, uFilter:'all', gFilter:'all', codeInput:'', tempName:'', transferInput:'', payMode:'cash', transferTime:'', dockOpen:false, sideMini:false, lFilter:'all', drawerOpen:false, filterOpen:false, pFilter:'all', editId:null, editPrice:'', editQty:'', editMemo:'', editDate:TODAY, imp2Open:false, imp2Step:0, memoLoaded:false, poLoaded:false, rDate:TODAY, impLines:null, impDone:'',
      sales: SEED_SALES.slice(), stock: stock, prices: prices,
      histFilter:'all', stockShop:'u', stockFilter:'all', restockOpen:null, restockQty:'', restockMemo:'',
      receipts: SEED_RECEIPTS.slice(), demand: SEED_DEMAND.slice(), cartWant:{}, cartPick:null, cartPickPos:null, wantPop:null, wantFor:null, wantCode:'', wantName:'', waitOpen:false, waitTab:'list', expOpen:false, expFmt:'pdf', expShop:null, expPrice:true, expCost:true, expZero:true, expDmd:true, expVer:'school', stuPop:null, waitShop:'all', waitQ:'',  memoOpen:false, memoNo:'', memoDate:TODAY, memoQty:{}, memoOff:{}, memoLast:null, askHist:null, wxOpen:false, wxScope:'all', wxSort:'item', wxFmt:'pdf', stuWait:false, resvAsk:null, resvOk:{}, batches: Object.assign({}, SEED_BATCHES), rcvId:null, rcvDoc:'memo', rcvPid:null, rcvVoidAsk:false, replaceId:null, impMemoDoc:null, impPoDoc:null, dupAlert:null, studentQ:'', classFilter:'all', studentId:null,
      rosters:DATA.rosters.slice(), names: Object.assign({}, BASE_NAMES), lastBackup:DATA.lastBackup, lastBackupAt:DATA.lastBackupAt, dStart:DATA.dStart, dMoves:DATA.dMoves, closes:DATA.closes, merges:DATA.merges, mv:null,
      importStep:0, fileChosen:false, termInput:'2/2569', effDate:TODAY, doneTerm:'', doneEff:'',
      staff:DATA.staff, newStaff:'',
      transferEmail:'', transferStep:0, toast:''
    };
  }
  componentDidMount(){
    var selfL = this; window.addEventListener('resize', function(){ selfL.lockCheck(); });
    document.addEventListener('focusin', function(e){ var t = e.target; if (t && t.tagName === 'INPUT' && t.type === 'number'){ setTimeout(function(){ try { t.select(); } catch(x){} }, 0); } });
    try { var dh0 = Number(localStorage.getItem('dockH')); if (dh0 >= 96) this.setState({dockH:dh0}); } catch(e){}
    try { var uc0 = JSON.parse(localStorage.getItem('uClosed') || 'null'); if (uc0) this.setState({uClosed:uc0}); } catch(e){}
    try { var tl0 = JSON.parse(localStorage.getItem('tlPrefs') || 'null'); if (tl0) this.setState({tlFold:tl0.fold||{}, tlHi:tl0.hi||{}}); } catch(e){}
    var self = this;
    this._onScroll = function(){
      var el = document.getElementById('stuband');
      var compact = false;
      if (el){ var r = el.getBoundingClientRect(); var lim = window.innerWidth <= 760 ? 125 : 133; compact = r.bottom < lim; }
      if (compact !== !!self.state.stuCompact) self.setState({stuCompact: compact});
      var sc = (window.scrollY || document.documentElement.scrollTop || 0) > 4;
      if (sc !== document.documentElement.hasAttribute('data-scrolled')){ if (sc) document.documentElement.setAttribute('data-scrolled', ''); else document.documentElement.removeAttribute('data-scrolled'); }
      var st = false;
      [].forEach.call(document.querySelectorAll('.fbar, .pinbar'), function(fb){ var on = false;
        if (sc){ var tp = parseFloat(window.getComputedStyle(fb).top) || 0, y = fb.getBoundingClientRect().top; on = Math.abs(y - tp) < 1.5; }
        if (on) st = true; if (on !== fb.hasAttribute('data-stuck')){ if (on) fb.setAttribute('data-stuck', ''); else fb.removeAttribute('data-stuck'); } });
      if (st !== document.documentElement.hasAttribute('data-fbstuck')){ if (st) document.documentElement.setAttribute('data-fbstuck', ''); else document.documentElement.removeAttribute('data-fbstuck'); }
    };
    this._onPanelScroll = function(e){ var el = e.target; if (!el || !el.classList) return;
      if (el.classList.contains('mbody') || el.classList.contains('exmain')){ var pn = el.closest('.mpanel'); if (pn) self._panelShade(pn); return; }
      if (!(el.classList.contains('mpanel') || el.classList.contains('fpop'))) return;
      var sc = el.scrollTop > 4; if (sc !== el.hasAttribute('data-scrolled')){ if (sc) el.setAttribute('data-scrolled', ''); else el.removeAttribute('data-scrolled'); } };
    this._panelShade = function(pn){ var bd = pn.querySelector(':scope > .mbody'); if (!bd) return; var sc = bd.querySelector('.exmain'); var cs = window.getComputedStyle && sc && window.getComputedStyle(sc).overflowY === 'auto' ? sc : bd;
      var top = cs.scrollTop > 4, more = cs.scrollTop + cs.clientHeight < cs.scrollHeight - 4;
      if (top !== pn.hasAttribute('data-scrolled')){ if (top) pn.setAttribute('data-scrolled', ''); else pn.removeAttribute('data-scrolled'); }
      if (more !== pn.hasAttribute('data-more')){ if (more) pn.setAttribute('data-more', ''); else pn.removeAttribute('data-more'); } };
    document.addEventListener('scroll', this._onPanelScroll, true);
    this.applyTheme();
    window.addEventListener('scroll', this._onScroll, {passive:true});
    window.addEventListener('resize', this._onScroll);
    // swipe down to dismiss bottom sheets (phone)
    var drag = null;
    var scrimOf = function(el){ if (el.classList.contains('dock')) return el.querySelector('.dockt'); var uw = el.closest('.uwrap'); if (uw) return uw.querySelector('.uscrim'); var mb = el.closest('.mback'); if (mb) return mb.querySelector('.mscrim'); var p = el.parentElement; return p ? p.querySelector('.fscrim') : null; };
    var backOf = function(el){ if (el.classList.contains('dock')) return null; if (el.closest('.uwrap')) return el.closest('.uwrap').querySelector('.uscrim'); return el.closest('.mback') ? el.closest('.mback').querySelector('.mscrim') : scrimOf(el); };
    this._onTS = function(e){ if (window.innerWidth > 760 || e.touches.length !== 1) return;
      var el = e.target.closest && e.target.closest('.mpanel, .fpop, .uwrap.open .upane, .dock:has(.dockbody)'); if (!el || e.target.closest('input[type=range], textarea, .mtip')) { drag = null; return; }
      var scE = el.classList.contains('dock') ? (e.target.closest('.dockrows') || el) : ((el.querySelector && el.querySelector(':scope > .mbody')) || el); drag = {el:el, sc:scE, y0:e.touches[0].clientY, t0:Date.now(), top0:scE.scrollTop, dy:0, on:false}; };
    this._onTM = function(e){ if (!drag) return; var dy = e.touches[0].clientY - drag.y0;
      if (!drag.on){ if (dy > 6 && drag.top0 <= 0 && drag.sc.scrollTop <= 0){ drag.on = true; drag.y0 = e.touches[0].clientY; dy = 0; } else if (Math.abs(dy) > 6){ drag = null; return; } else return; }
      if (e.cancelable) e.preventDefault(); drag.dy = Math.max(0, dy);
      drag.el.style.transition = 'none'; drag.el.style.transform = 'translateY(' + drag.dy + 'px)';
      var sc = backOf(drag.el); if (sc) { sc.style.transition = 'none'; sc.style.opacity = String(Math.max(0.2, 1 - drag.dy / (drag.el.offsetHeight || 600))); } };
    this._onTE = function(){ if (!drag) return; var d = drag; drag = null; if (!d.on) return;
      var v = d.dy / Math.max(1, Date.now() - d.t0), sc = backOf(d.el), close = d.dy > 120 || (d.dy > 40 && v > 0.5);
      d.el.style.transition = 'transform .2s ease-out'; if (sc) sc.style.transition = 'opacity .2s ease-out';
      if (close && d.el.classList.contains('dock')){ d.el.style.transform = ''; setTimeout(function(){ d.el.style.transition = ''; }, 220); var tg = scrimOf(d.el); if (tg) tg.click(); }
      else if (close){ d.el.style.transform = 'translateY(100%)'; if (sc) sc.style.opacity = '0';
        setTimeout(function(){ var s2 = scrimOf(d.el); d.el.style.transform = ''; d.el.style.transition = ''; if (sc){ sc.style.opacity = ''; sc.style.transition = ''; } if (s2) s2.click(); }, 200); }
      else { d.el.style.transform = ''; if (sc) sc.style.opacity = ''; setTimeout(function(){ d.el.style.transition = ''; if (sc) sc.style.transition = ''; }, 220); } };
    document.addEventListener('touchstart', this._onTS, {passive:true});
    document.addEventListener('touchmove', this._onTM, {passive:false});
    document.addEventListener('touchend', this._onTE); document.addEventListener('touchcancel', this._onTE);
    this._onKey = function(e){ if (e.key !== 'Escape') return; var ss = document.querySelectorAll('.mback .mscrim, .fscrim'); if (ss.length) ss[ss.length - 1].click(); };
    document.addEventListener('keydown', this._onKey);
    // swipe right to open / left to close the side menu (phone)
    var sw = null;
    var hScroll = function(n){ for (; n && n !== document.body; n = n.parentElement){ if (n.scrollWidth > n.clientWidth + 2){ var ox = window.getComputedStyle(n).overflowX; if (ox === 'auto' || ox === 'scroll') return true; } } return false; };
    this._onSwS = function(e){ sw = null; if (window.innerWidth > 760 || e.touches.length !== 1) return;
      if (document.querySelector('.mback, .fpop, .trpop')) return;
      var tg = e.target; if (tg.closest && tg.closest('input, textarea, select, [type=range]')) return; if (hScroll(tg)) return;
      var el = document.querySelector('.side'); if (!el) return;
      sw = {el:el, x0:e.touches[0].clientX, y0:e.touches[0].clientY, t0:Date.now(), open:!!self.state.drawerOpen, axis:null, dx:0}; };
    this._onSwM = function(e){ if (!sw) return; var dx = e.touches[0].clientX - sw.x0, dy = e.touches[0].clientY - sw.y0;
      if (!sw.axis){ if (Math.abs(dx) < 10 && Math.abs(dy) < 10) return; sw.axis = Math.abs(dx) > Math.abs(dy) * 1.2 ? 'x' : 'y';
        if (sw.axis === 'x' && ((sw.open && dx > 0) || (!sw.open && dx < 0))) sw.axis = 'y'; if (sw.axis === 'y'){ sw = null; return; } }
      if (e.cancelable) e.preventDefault(); sw.dx = dx;
      var w = sw.el.offsetWidth, x = sw.open ? Math.min(0, Math.max(-w, dx)) : Math.min(0, Math.max(-w, -w + dx));
      sw.el.style.transition = 'none'; sw.el.style.transform = 'translateX(' + x + 'px)'; sw.el.style.boxShadow = '8px 0 32px rgba(32,55,79,' + (0.25 * (1 + x / w)).toFixed(3) + ')'; };
    this._onSwE = function(){ if (!sw) return; var d = sw; sw = null; if (d.axis !== 'x') return;
      var w = d.el.offsetWidth, v = d.dx / Math.max(1, Date.now() - d.t0);
      var open = d.open ? !(-d.dx > w * 0.35 || (d.dx < -30 && v < -0.4)) : (d.dx > w * 0.35 || (d.dx > 30 && v > 0.4));
      d.el.style.transition = 'transform .22s ease';
      d.el.style.transform = open ? 'translateX(0px)' : 'translateX(-100%)';
      setTimeout(function(){ if (open !== !!self.state.drawerOpen) self.setState({drawerOpen: open});
        setTimeout(function(){ d.el.style.transition = ''; d.el.style.transform = ''; d.el.style.boxShadow = ''; }, 240); }, 0); };
    document.addEventListener('touchstart', this._onSwS, {passive:true});
    document.addEventListener('touchmove', this._onSwM, {passive:false});
    document.addEventListener('touchend', this._onSwE); document.addEventListener('touchcancel', this._onSwE);
  }
  componentWillUnmount(){
    document.removeEventListener('touchstart', this._onTS); document.removeEventListener('touchmove', this._onTM); document.removeEventListener('touchend', this._onTE); document.removeEventListener('touchcancel', this._onTE); document.removeEventListener('keydown', this._onKey); document.removeEventListener('touchstart', this._onSwS); document.removeEventListener('touchmove', this._onSwM); document.removeEventListener('touchend', this._onSwE); document.removeEventListener('touchcancel', this._onSwE);
    window.removeEventListener('scroll', this._onScroll);
    document.removeEventListener('scroll', this._onPanelScroll, true);
    window.removeEventListener('resize', this._onScroll);
  }
  applyTheme(){ var s = this.state, th = s.page==='sellU' ? 'u' : (s.page==='sellS' ? 's' : (s.page==='stock' ? (s.stockShop||'u') : 'app'));
    if (s.waitOpen && s.waitShop && s.waitShop !== 'all') th = s.waitShop;
    if (s.expOpen){ var es = s.expShop || s.stockShop || 'u'; th = es === 'both' ? 'app' : es; }
    if (document.documentElement.getAttribute('data-theme') !== th) document.documentElement.setAttribute('data-theme', th); }
  lockCheck(){ if (typeof document === 'undefined') return; var r = document.documentElement, lock = false; if (window.innerWidth <= 760){ document.querySelectorAll('.mback .mpanel, .fpop, .uwrap.open .upane').forEach(function(e){ if (e.offsetHeight > window.innerHeight * 0.5) lock = true; }); } r.classList.toggle('lockbg', lock); }
  componentDidUpdate(){ var self0 = this; setTimeout(function(){ self0.lockCheck(); }, 0); this.applyTheme(); if (this._onScroll) this._onScroll(); var self = this; if (this._panelShade && typeof document !== 'undefined' && document.querySelectorAll) setTimeout(function(){ document.querySelectorAll('.mpanel').forEach(function(pn){ self._panelShade(pn); }); }, 0); }
  go(page, extra){ if (this.state.role !== 'owner' && ['dash','students','student','import','settings'].indexOf(page) >= 0) page = 'sellU';
    var u = Object.assign({page:page, q:'', searchOpen:false, drawerOpen:false}, extra||{}); var k = page==='student' ? 'students' : page;
    if (this._navSets && this._navSets[k]){ u.seen = Object.assign({}, this.state.seen||{}); u.seen[k] = this._navSets[k].slice(); } this.setState(u); }
  okFlash(text, amt, actLabel, act, sub, cls, undo){ var self = this; clearTimeout(this._ot); var t = {text:text, amt:amt||'', actLabel:actLabel||'', act:act||null, sub:sub||'', cls:cls||'', undo:undo||null, k:Date.now()}; this.setState({toastOk:null}); setTimeout(function(){ self.setState({toastOk:t}); }, 0); this._ot = setTimeout(function(){ self.setState({toastOk:null}); }, undo ? 10000 : 4200); }
  flash(msg){ var self = this; this.setState({toast:msg}); clearTimeout(this._t); this._t = setTimeout(function(){ self.setState({toast:''}); }, 2600); }
  classFor(code, date, rosters){
    var rs = (rosters || this.state.rosters).slice().sort(function(a,b){ return b.eff.localeCompare(a.eff); });
    for (var k=0; k<rs.length; k++){ var r = rs[k]; if (r.eff <= date && r.classes[code]) return r.classes[code]; }
    return null;
  }
  latestRoster(){ return this.state.rosters.slice().sort(function(a,b){ return b.eff.localeCompare(a.eff); })[0]; }
  isUnmatched(sale){ return sale.shop==='u' && sale.code && !this.classFor(sale.code, sale.date); }

  renderVals(){
    function icoFlags(list, names){ return list.map(function(k, i){ if (names[i]) k['i_'+names[i]] = true; return k; }); }
    var self = this, s0 = this.state, s = Object.assign({}, s0, {sales: s0.sales.filter(function(x){ return !x.voided; })});
    var PRODUCTS = ALLP.filter(function(p){ return !p.mergedInto; });
    var isOwner = s.role === 'owner';
    var OWNER = ['dash','students','student','import','settings'];
    var locked = false;
    var show = {
      locked: locked,
      dash: s.page==='dash' && !locked,
      sell: s.page==='sellU' || s.page==='sellS',
      history: s.page==='history',
      stock: s.page==='stock',
      students: s.page==='students' && !locked,
      student: s.page==='student' && !locked,
      import: s.page==='import' && !locked,
      settings: s.page==='settings' && !locked,
      help: s.page==='help'
    };

    var unmatchedSales = s.sales.filter(function(x){ return self.isUnmatched(x); });
    // low stock = will run out within 14 days at the average pace of the last 4 weeks
    var pace = {}, paceFrom = addDays(TODAY, -27);
    s.sales.forEach(function(x){ if (x.date < paceFrom || x.date > TODAY) return; x.items.forEach(function(i){ pace[i.pid] = (pace[i.pid]||0) + i.qty; }); });
    var daysLeft = function(pid){ var st = s.stock[pid]||0, wk = (pace[pid]||0)/4; return wk > 0 ? Math.max(1, Math.round(st/(wk/7))) : Infinity; };
    var isLow = function(pid){ return (s.stock[pid]||0) > 0 && s.prices[pid] > 0 && daysLeft(pid) <= 14; };
    var lowCount = PRODUCTS.filter(function(p){ return isLow(p.id); }).length;
    var zeroCount = PRODUCTS.filter(function(p){ return !s.prices[p.id]; }).length;

    // ---------- nav ----------
    // a menu dot means "new since you last opened this page"
    var navSets = {stock: PRODUCTS.filter(function(p){ return isLow(p.id) || !s.prices[p.id]; }).map(function(p){ return p.id; }),
      students: unmatchedSales.map(function(x){ return x.code; }).filter(function(c, i, a){ return a.indexOf(c) === i; })};
    self._navSets = navSets;
    var newIn = function(k){ var cur = s.page==='student' ? 'students' : s.page; if (cur === k) return 0; var seen = (s.seen||{})[k] || [];
      return navSets[k].filter(function(x){ return seen.indexOf(x) < 0; }).length; };
    var NAV = [
      {title:'ภาพรวม', items:[['dash','แดชบอร์ด',true,0,'dash']]},
      {title:'ขายหน้าร้าน', items:[['sellU','ขายเครื่องแบบ',false,0,'shirt'],['sellS','ขายเครื่องเขียน',false,0,'pencil']]},
      {title:'ข้อมูล', items:[['history','ประวัติการขาย',false,0,'receipt'],['stock','สต็อกสินค้า',false,newIn('stock'),'box'],['students','รายชื่อนักเรียน',true,newIn('students'),'people']]},
      {title:'ระบบ', items:[['import','นำเข้ารายชื่อ',true,0,'upload'],['help','วิธีใช้',false,0,'help'],['settings','ผู้ใช้และการส่งต่อ',true,(isOwner && s.lastBackup && dayGap(s.lastBackup, TODAY) > 30) ? 1 : 0,'users']]}
    ];
    var active = s.page==='student' ? 'students' : s.page;
    var navGroups = NAV.map(function(g){
      var items = g.items.filter(function(it){ return !it[2] || isOwner; }).map(function(it){
        var k = it[4];
        return {label:it[1], cls:'nav'+(active===it[0]?' on':''), go:function(){ self.go(it[0]); }, badge:it[3], hasBadge: it[3]>0, dotCls:'ndot'+(it[0]==='settings' ? ' red' : ''), dotLabel: it[0]==='settings' ? 'ถึงเวลาสำรองข้อมูล' : 'มีเรื่องต้องดู',
          nDash:k==='dash', nShirt:k==='shirt', nPencil:k==='pencil', nReceipt:k==='receipt', nBox:k==='box', nPeople:k==='people', nUpload:k==='upload', nUsers:k==='users', nHelp:k==='help', helpM: icoStyle('help')};
      });
      return {title:g.title, items:items};
    }).filter(function(g){ return g.items.length; });

    var roleChips = [['owner','เจ้าของ'],['staff','ครูผู้ช่วย']].map(function(r){
      return {label:r[1], cls:'segb'+(s.role===r[0]?' on':''), pick:function(){ var u = {role:r[0]}; if (r[0]!=='owner' && OWNER.indexOf(self.state.page) >= 0) u.page = 'sellU'; self.setState(u); self.flash(r[0]==='owner'?'ดูในฐานะเจ้าของ':'ดูในฐานะครูผู้ช่วย: เห็นหน้าขาย ประวัติ และสต็อก'); }};
    });
    var user = {name:DATA.myName, role: isOwner ? 'เจ้าของระบบ' : 'ครูผู้ช่วย', initial:(DATA.myName || '?').charAt(0).toUpperCase()};

    // ---------- search ----------
    var q = (s.q||'').trim().toLowerCase();
    var sections = [];
    // size lookup: type a size, see it plus the nearest sizes that are in stock
    var SZORD = ['1','2','3','4','5','S','M','L','XL','XXL','XXXL'];
    var szKey = function(z){ var n = parseFloat(z); return isNaN(n) ? 1000+SZORD.indexOf(String(z).toUpperCase()) : (SZORD.indexOf(String(z)) >= 0 && n <= 5 ? n : 100+n); };
    var sizeHits = [];
    var QALIAS = [
      [['g','girl','girls','หญิง','ญ','ผู้หญิง'], function(p){ return p.gender === 'girl' || p.gender === 'both'; }],
      [['b','boy','boys','ชาย','ช','ผู้ชาย'], function(p){ return p.gender === 'boy' || p.gender === 'both'; }],
      [['skirt','skirts','กระโปรง'], function(p){ return p.icon === 'skirt'; }],
      [['shirt','shirts','top','tops','เสื้อ','บน','ท่อนบน'], function(p){ return p.part === 'top'; }],
      [['pants','shorts','short','trousers','กางเกง'], function(p){ return p.icon === 'pants'; }],
      [['bottom','bottoms','ล่าง','ท่อนล่าง'], function(p){ return p.part === 'bottom'; }],
      [['pe','sport','sports','พละ','พล'], function(p){ return p.groups.indexOf('pe') >= 0; }],
      [['student','uniform','นักเรียน','นร'], function(p){ return p.groups.indexOf('student') >= 0 && p.groups.indexOf('kg') < 0; }],
      [['k','kg','kindergarten','อ','อนุบาล'], function(p){ return p.groups.indexOf('kg') >= 0; }],
      [['p','primary','ป','ประถม'], function(p){ return p.groups.indexOf('kg') < 0; }],
      [['set','dress','ชุด','เต็มตัว'], function(p){ return p.part === 'set'; }]
    ];
    var ALLSZ = {}; PRODUCTS.forEach(function(p){ if (p.shop === 'u' && p.size) ALLSZ[String(p.size).toLowerCase()] = 1; });
    var qParts = q.replace(/([^\d\s])(\d)/g, '$1 $2').split(/\s+/).filter(Boolean), qSize = '', qTests = [];
    for (var qi = qParts.length-1; qi >= 0; qi--) if (!qSize && ALLSZ[qParts[qi]]){ qSize = qParts[qi]; qParts.splice(qi,1); }
    qParts.forEach(function(tok){
      var hit = QALIAS.filter(function(a){ return a[0].indexOf(tok) >= 0; })[0];
      qTests.push(hit ? hit[1] : function(p){ return p.name.toLowerCase().indexOf(tok) >= 0; });
    });
    if (qSize){
      var byName = {}, order = [];
      PRODUCTS.forEach(function(p){ if (p.shop !== 'u' || !p.size) return; if (!byName[p.name]){ byName[p.name] = []; order.push(p.name); } byName[p.name].push(p); });
      order.forEach(function(nm){
        if (!qTests.every(function(f){ return f(byName[nm][0]); })) return;
        var list = byName[nm].slice().sort(function(a,b){ return szKey(a.size) - szKey(b.size); });
        var ix = -1; list.forEach(function(p, k){ if (String(p.size).toLowerCase() === qSize) ix = k; });
        if (ix < 0) return;
        var hit = list[ix], below = [], above = [];
        for (var k = ix-1; k >= 0 && below.length < 2; k--) if (s.stock[list[k].id] > 0) below.unshift(list[k]);
        for (var k2 = ix+1; k2 < list.length && above.length < 2; k2++) if (s.stock[list[k2].id] > 0) above.push(list[k2]);
        var tile = function(p, exact){
          var st0 = s.stock[p.id]||0, pr = s.prices[p.id], low = isLow(p.id);
          return {size: p.size, stock: st0 > 0 ? 'เหลือ '+st0 : 'หมด', price: pr ? baht(pr) : 'ไม่มีราคา',
            cls: 'sztile t-'+p.tone+(exact ? ' exact' : ' near')+(st0 > 0 ? '' : ' out')+(low ? ' low' : '')+(pr ? '' : ' noprice'),
            add: function(){ self.quickAdd(p.id); }};
        };
        var blank = {size:'', stock:'', price:'', cls:'sztile blank', add:function(){}};
        var hst = s.stock[hit.id]||0;
        sizeHits.push(Object.assign({name:nm, giCls:'gico gi-'+hit.tone+gx(nm),
          note: hst > 0 ? 'ไซส์ '+hit.size+' เหลือ '+hst : 'ไซส์ '+hit.size+' หมด'+(below.length || above.length ? ': ดูไซส์ใกล้เคียง' : ''),
          noteCls: 'szn-note '+(hst > 5 ? 'ok' : (hst > 0 ? 'low' : 'out')),
          tiles: [0,1].map(function(k){ var p = below[k - (2 - below.length)]; return p ? tile(p,false) : blank; }).concat([tile(hit,true)]).concat([0,1].map(function(k){ return above[k] ? tile(above[k],false) : blank; }))}, iconFlags(hit.icon, hit), sFlags(hit.icon, hit)));
      });
    }
    if (q){
      var prodRows = PRODUCTS.filter(function(p){ return itemLabel(p.id).toLowerCase().indexOf(q) >= 0; }).slice(0,4).map(function(p){
        return {main:itemLabel(p.id), sub:(p.shop==='u'?'เครื่องแบบ':'เครื่องเขียน')+'\u2003คงเหลือ '+s.stock[p.id]+'\u2003'+(s.prices[p.id]?baht(s.prices[p.id]):'ยังไม่ตั้งราคา'), go:function(){ self.go('stock',{stockShop:p.shop, stockFilter:'all'}); }};
      });
      if (prodRows.length && !sizeHits.length) sections.push({title:'สินค้า', rows:prodRows});
      if (isOwner){
        var stuRows = Object.keys(s.names).filter(function(c){ return c.indexOf(q) >= 0 || s.names[c].toLowerCase().indexOf(q) >= 0; }).slice(0,4).map(function(c){
          var cl = self.classFor(c,'9999-12-31');
          return {main:s.names[c], sub:'รหัส '+c+'\u2003'+(cl||'ไม่อยู่ในรายชื่อ'), go:function(){ self.go('student',{studentId:c}); }};
        });
        if (stuRows.length) sections.push({title:'นักเรียน', rows:stuRows});
      }
      var saleRows = s.sales.filter(function(x){ return x.code && x.code.indexOf(q) >= 0; }).slice(0,3).map(function(x){
        return {main:thDate(x.date)+' '+x.time+'\u2003'+baht(saleTotal(x)), sub:'รหัส '+x.code+': '+itemsText(x), go:function(){ self.go('history',{saleId:x.id, hDays:(function(){ var o = {}; o[x.date] = true; return o; })(), hMonths:(function(){ var o = {}; o[x.date.slice(0,7)] = true; return o; })()}); }};
      });
      if (saleRows.length) sections.push({title:'การขาย', rows:saleRows});
    }

    // ---------- dashboard ----------
    var sinceMap = {today:TODAY, week:addDays(TODAY,-6), month:'2026-10-01', term:'2026-05-16'};
    var since = sinceMap[s.period];
    var ps = s.sales.filter(function(x){ return x.date >= since && x.date <= TODAY; });
    var uS = ps.filter(function(x){ return x.shop==='u'; }), sS = ps.filter(function(x){ return x.shop==='s'; });
    var sum = function(arr){ return arr.reduce(function(a,x){ return a+saleTotal(x); },0); };
    var transferSum = ps.reduce(function(a,x){ return a+x.transfer; },0);
    var prevSince = addDays(since, -((s.period==='today' ? 1 : (s.period==='week' ? 7 : 0))));
    var prevPs = s.period==='today' || s.period==='week' ? s.sales.filter(function(x){ return x.date >= prevSince && x.date < since; }) : [];
    var curT = sum(ps), prevT = sum(prevPs), cashT = curT - transferSum;
    var cashPct = curT ? Math.round(cashT/curT*100) : 0, circ = 2*Math.PI*40;
    var hr = new Date().getHours();
    var dash = {
      total: baht(curT), count: ps.length, items: ps.reduce(function(a,x){ return a+x.items.reduce(function(b,i){ return b+i.qty; },0); },0),
      u: baht(sum(uS)), uCount: uS.length, s: baht(sum(sS)), sCount: sS.length,
      cash: baht(cashT), transfer: baht(transferSum), cashPct: cashPct+'%', dash: (circ*cashPct/100).toFixed(1)+' '+circ.toFixed(1),
      totalLabel: 'ยอดขายรวม'+({today:'วันนี้', week:' 7 วัน', month:'เดือนนี้', term:'ภาคเรียนนี้'})[s.period],
      deltaCls: 'dchip'+(curT < prevT ? ' down' : ''), hasDelta: prevPs.length > 0 && prevT > 0, delta: (curT >= prevT ? '+' : '−')+Math.abs(Math.round((curT-prevT)/(prevT||1)*100))+'% จาก'+(s.period==='today' ? 'เมื่อวาน' : ' 7 วันก่อน'),
      greet: (hr < 12 ? 'สวัสดีตอนเช้า' : (hr < 17 ? 'สวัสดีตอนบ่าย' : 'สวัสดีตอนเย็น'))+', '+(isOwner ? DATA.myName : 'ครู ก.'),
      rangeLabel: s.period==='today' ? 'วันนี้ '+thDate(TODAY) : thDate(since)+' – '+thDate(TODAY)
    };
    var periods = [['today','วันนี้'],['week','7 วัน'],['month','เดือนนี้'],['term','ภาคเรียนนี้']].map(function(p){
      return {label:p[1], cls:'segb'+(s.period===p[0]?' on':''), pick:function(){ self.setState({period:p[0]}); }};
    });
    var dayList = []; for (var i=6; i>=0; i--) dayList.push(addDays(TODAY,-i));
    var dayVals = dayList.map(function(d){
      var ds = s.sales.filter(function(x){ return x.date===d; });
      return {d:d, u:sum(ds.filter(function(x){ return x.shop==='u'; })), s:sum(ds.filter(function(x){ return x.shop==='s'; }))};
    });
    var maxV = Math.max.apply(null, dayVals.map(function(v){ return Math.max(v.u, v.s); }).concat([1]));
    var AW = 560, AH = 200, pt = function(k){ return dayVals.map(function(v,i){ return [Math.round(i*AW/6), Math.round(AH-10-(v[k]/maxV*(AH-36)))]; }); };
    var curve = function(p){ var d = 'M'+p[0][0]+' '+p[0][1]; for (var i=1; i<p.length; i++){ var x = (p[i-1][0]+p[i][0])/2; d += ' C'+x+' '+p[i-1][1]+' '+x+' '+p[i][1]+' '+p[i][0]+' '+p[i][1]; } return d; };
    var pu = pt('u'), pst = pt('s'), lastV = dayVals[6];
    var area = {uLine:curve(pu), uFill:curve(pu)+' L'+AW+' '+AH+' L0 '+AH+' Z', sLine:curve(pst), sFill:curve(pst)+' L'+AW+' '+AH+' L0 '+AH+' Z',
      uToday:baht(lastV.u), sToday:baht(lastV.s), range: thShort(dayList[0])+' ถึง '+thShort(TODAY),
      days: dayVals.map(function(v,i){ var dn = ['อา.','จ.','อ.','พ.','พฤ.','ศ.','ส.'][new Date(v.d+'T00:00:00').getDay()]; return {label: v.d===TODAY ? 'วันนี้' : dn, x: Math.round(i/6*1000)/10, cls:'ad'+(i===0?' first':(i===6?' last today':''))}; })};
    var days = [];
    var agg = {};
    ps.forEach(function(x){ x.items.forEach(function(i){ var k = PMAP[i.pid].mergedInto || i.pid; agg[k] = agg[k] || {qty:0, rev:0}; agg[k].qty += i.qty; agg[k].rev += i.qty*i.price; }); });
    var pInfoT = function(pid){ var p = PMAP[pid]; return Object.assign({pname:p.name, psize: p.size ? '#'+p.size : '', giCls:'gico gi-'+(p.shop==='u' ? p.tone : p.cat)+gx(p.name)}, iconFlags(p.icon, p), sFlags(p.icon, p)); };
    var topIds = Object.keys(agg).sort(function(a,b){ return agg[b].qty - agg[a].qty || agg[b].rev - agg[a].rev; });
    var topMax = topIds.length ? agg[topIds[0]].qty : 1;
    var top = topIds.slice(0,6).map(function(id,ix){
      return Object.assign(pInfoT(id), {rank:ix+1, qty:agg[id].qty, pct: Math.round(agg[id].qty/topMax*100), barCls: PMAP[id].shop==='u' ? 'u' : 's'});
    });
    var bestOf = function(sh){ var id = topIds.filter(function(k){ return PMAP[k].shop===sh; })[0]; return id ? itemLabel(id) : ''; };
    var shopCards = [['u','เครื่องแบบ',uS,'sellU'],['s','เครื่องเขียน',sS,'sellS']].map(function(c){
      var n = c[2].reduce(function(a,x){ return a+x.items.reduce(function(b,i){ return b+i.qty; },0); },0), bst = bestOf(c[0]);
      return {cls:'dbx dshop '+c[0], isU:c[0]==='u', isS:c[0]==='s', title:c[1], amt:baht(sum(c[2])), amtB:bn(baht(sum(c[2]))), sub:c[2].length+' รายการ\u2003'+n+' ชิ้น', best:bst, hasBest:!!bst, go:function(){ self.go(c[3]); }};
    });
    var todayU = s.sales.filter(function(x){ return x.date===TODAY && x.shop==='u'; }), todayUT = sum(todayU), todayUTr = todayU.reduce(function(a,x){ return a+x.transfer; },0);
    var drawerSince = s.drawerSince || '2026-09-30';
    var drawerT = sum(s.sales.filter(function(x){ return x.shop==='s' && x.date > drawerSince && x.date <= TODAY; }));
    var FLOAT = 0, DCOL = {1000:'#B7C0C9',500:'#9C88C4',100:'#D9776D',50:'#7FA6D9',20:'#86B98A',10:'#D8B957',5:'#C9CDD2',1:'#DADDE1'};
    var closeOf = function(sh){ return (s.closes||[]).filter(function(c){ return c.date===TODAY && c.shop===sh; })[0]; };
    var closeNote = function(c){ return 'ปิดยอดแล้ว '+c.at+' '+(c.diff ? (c.diff<0 ? 'ขาด ' : 'เกิน ')+baht(Math.abs(c.diff)) : 'ตรงพอดี'); };
    var cU = closeOf('u'), cS = closeOf('s');
    // stationery drawer: opening = last evening count, then today's cash sales and money in/out
    var sToday = s.sales.filter(function(x){ return x.shop==='s' && x.date===TODAY && !x.voided; });
    var sCash = sum(sToday) - sToday.reduce(function(a,x){ return a+(x.transfer||0); },0);
    var prevS = (s.closes||[]).filter(function(c){ return c.shop==='s' && c.date < TODAY; }).sort(function(a,b){ return b.date.localeCompare(a.date); })[0];
    var dOpen = prevS ? {amt:prevS.counted, date:prevS.date} : (s.dStart || {amt:0, date:TODAY});
    var mvToday = (s.dMoves||[]).filter(function(m){ return m.date===TODAY; });
    var mvSum = function(l, d){ return l.reduce(function(a,m){ return a + (m.dir===d ? m.amt : 0); },0); };
    var mvPre = mvToday.filter(function(m){ return !m.after; }), mvPost = mvToday.filter(function(m){ return m.after; });
    var outT = mvSum(mvPre,'out'), inT = mvSum(mvPre,'in');
    var dExp = dOpen.amt + sCash - outT + inT;
    var dNow = cS ? cS.counted - mvSum(mvPost,'out') + mvSum(mvPost,'in') : dExp;
    var mvNet = mvSum(mvToday,'in') - mvSum(mvToday,'out');
    var kSum = function(k){ return (s.dMoves||[]).reduce(function(a,m){ return a + (m.kind===k ? m.amt : 0); },0); };
    var oweAdv = Math.max(0, kSum('สำรองจ่าย') - kSum('เบิกคืนสำรองจ่าย')), oweLoan = Math.max(0, kSum('ยืมชั่วคราว') - kSum('คืนเงินที่ยืม')), owe = oweAdv + oweLoan;
    var oweTxt = owe ? 'ค้างคืน '+baht(owe) : '';
    var openMv = function(dir, kind, from){ return function(){ self.setState({mv:{dir:dir, kind:kind||'', amt:'', note:'', from:from||''}}); }; };
    var money = {uCash:baht(todayUT-todayUTr), uTr:baht(todayUTr), uTot:baht(todayUT), sTot:baht(dNow), sCash:baht(sCash),
      owe:!!owe, oweTxt:oweTxt, mvbCls:'mval mvb'+(owe ? ' owe' : ''), sNetTxt: mvNet ? (mvNet<0 ? '−' : '+')+Math.abs(mvNet).toLocaleString() : '0', mvOpen:openMv('out'),
      copy:function(){ self.setState({cpv:{k:'u', date:TODAY}}); },
      hand:openMv('out', 'นำส่งผู้บริหาร')};
    var openClose = function(sh){ return function(){ self.setState({clsShop:sh, clsCnt:{}, clsReason:'', clsNote:''}); }; };
    Object.assign(money, {uOpen:!cU, uClosed:!!cU, sOpen:!cS, sClosed:!!cS, uNote: cU ? closeNote(cU) : 'นำส่งผู้บริหารทุกเย็น', sNote: cS ? closeNote(cS) : 'ตั้งต้น '+baht(dOpen.amt)+' จากยอดนับ '+thD2(dOpen.date), closeU:openClose('u'), closeS:openClose('s')});

    var hTopic = function(ico, tone, title, steps, tip){ return {icoM: icoStyle(ico), boxCls:'hpi '+tone, title:title, steps: steps.map(function(t, i){ return {n:i+1, text:t}; }), tip:tip||'', hasTip:!!tip}; };
    var helpTopics = [
      hTopic('schoolshirt','u','ขายเครื่องแบบ',['กรอกรหัสนักเรียน ระบบจะขึ้นชื่อให้ตรวจ','แตะชื่อสินค้าทางซ้าย แล้วแตะไซส์เพื่อเพิ่มลงรายการ','ไซส์ที่นักเรียนเคยซื้อจะขึ้นในขนาดแนะนำ แตะ + เพื่อเพิ่มได้เลย','เลือกเงินสดหรือโอน แล้วกดบันทึกการขาย'],'เปลี่ยนไซส์ในรายการได้ที่ป้ายไซส์ เช่น #34'),
      hTopic('pencil','s','ขายเครื่องเขียน',['เลือกหมวดทางซ้าย เช่น สมุด หรือวอกเกิ้ล','แตะสินค้าเพื่อเพิ่มลงรายการ แตะซ้ำเพื่อเพิ่มจำนวน','กดบันทึกการขาย'],'เครื่องเขียนรับเงินสดเท่านั้น'),
      hTopic('undo','n','กดผิด แก้อย่างไร',['หลังบันทึก จะมีแถบบันทึกการขายแล้วขึ้นด้านล่าง','กดเลิกทำภายใน 10 วินาที ของจะกลับเข้ารายการให้แก้ต่อ','ถ้าเลยเวลาแล้ว แจ้งเจ้าของระบบให้แก้ในประวัติการขาย']),
      hTopic('transfer','n','รับเงินโอน',['เลือกโอนที่แถบด้านล่าง','ใส่เวลาตามสลิป หรือกดใช้เวลาตอนนี้','ตรวจยอดในสลิปให้ตรงกับยอดรวมก่อนบันทึก']),
      hTopic('wait','a','ไซส์ที่หมด',['ไซส์ที่หมดเป็นกรอบเส้นประ แตะเพื่อดูไซส์ใกล้เคียงที่มีของ','ถ้านักเรียนอยากรอ ใส่รหัสหรือชื่อไว้','ระบบจะจองให้เมื่อของมาถึง']),
      hTopic('search','n','ค้นหาไซส์เร็ว',['พิมพ์ไซส์ในช่องค้นหาด้านบน เช่น ญ 36 หรือกระโปรง 30','แตะไซส์ที่ขึ้นมาเพื่อเพิ่มลงรายการ'],'ใช้ได้ทั้งภาษาไทยและอังกฤษ เช่น g skirt 36'),
      hTopic('offline','n','เน็ตหลุด',['ขายต่อได้ตามปกติ ระบบเก็บการขายไว้ในเครื่อง','เมื่อกลับมาออนไลน์ ระบบส่งยอดให้เองและบอกว่าส่งกี่รายการ'])
    ];
    if (isOwner) helpTopics.push(hTopic('cash','n','ปิดยอดตอนเย็น',['เครื่องแบบ: นับเงินในถุงเครื่องแบบ แล้วคัดลอกยอดนำส่งผู้บริหาร','เครื่องเขียน: บันทึกเงินเข้า/ออกของวันนี้ให้ครบก่อน เช่น นำส่งผู้บริหาร หรือสำรองจ่าย','นับธนบัตรและเหรียญในลิ้นชัก ยอดที่นับได้จะเป็นเงินตั้งต้นของวันพรุ่งนี้','ถ้ายอดไม่ตรง เลือกสาเหตุก่อนปิดยอด แล้วคัดลอกยอดขายวันนี้ลงสมุดเงินสด'],'เฉพาะเจ้าของระบบ'));
    var help = {topics: helpTopics, lede: isOwner ? 'คู่มือสั้นสำหรับครูผู้ช่วย และขั้นตอนปิดยอดของเจ้าของระบบ' : 'ขายของ รับเงินโอน และแก้ไขเมื่อกดผิด',
      shirtM: icoStyle('schoolshirt'), pencilM: icoStyle('pencil'), supportM: icoStyle('support'), goU:function(){ self.go('sellU'); }, goS:function(){ self.go('sellS'); }, phone: window.innerWidth <= 760};

    var bkDays = s.lastBackup ? dayGap(s.lastBackup, TODAY) : 999, bkDue = isOwner && bkDays > 30;
    var bkLast = s.lastBackup ? (bkDays === 0 ? 'วันนี้'+(s.lastBackupAt ? ' '+s.lastBackupAt : '') : thD2(s.lastBackup)) : 'ยังไม่เคยสำรอง';
    var logN = s0.sales.reduce(function(a, x){ return a + ((x.log||[]).length); }, 0) + Object.keys(s.batches||{}).reduce(function(a, k){ return a + (((s.batches[k]||{}).log||[]).length); }, 0);
    var bk = {due: bkDue, pillCls:'bkpill '+(bkDue ? 'warn' : 'ok'), pillText:'สำรองล่าสุด '+bkLast, dueText:'ไม่ได้สำรองข้อมูลมา '+bkDays+' วันแล้ว ควรดาวน์โหลดเก็บไว้เดือนละครั้ง',
      dbM: icoStyle('db'), dlM: icoStyle('dl'), alertM: icoStyle('alert'), fileName:'สำรองข้อมูลร้านค้า '+thD2(TODAY)+'.xlsx',
      items: [['receipt','การขาย', s0.sales.length+' รายการ'],['box','สินค้าและสต็อก', ALLP.length+' รายการ'],['dl','การรับเข้า', Object.keys(s.batches||{}).length+' ครั้ง'],['people','นักเรียน', Object.keys(s.names||{}).length+' คน'],['wait','รายการรอของ', (s.demand||[]).length+' รายการ'],['cash','การปิดยอด', (s.closes||[]).length+' ครั้ง'],['merge','การรวมสินค้า', (s.merges||[]).length+' รายการ'],['file','ประวัติการแก้ไข', logN+' ครั้ง']].map(function(r){ return {icoM: icoStyle(r[0]), label:r[1], n:r[2]}; }),
      download:function(){ self.setState({lastBackup:TODAY, lastBackupAt:nowHM()}); self.okFlash('ดาวน์โหลดไฟล์สำรองแล้ว', '', '', null, 'สำรองข้อมูลร้านค้า '+thD2(TODAY)+'.xlsx 8 ชีต'); }};
    var cls = {open:false, notes:[], coins:[], chips:[], parts:[], moves:[]};
    if (s.clsShop){
      var csh = s.clsShop, isS = csh==='s', ccnt = s.clsCnt||{}, csum = 0, cany = false;
      var cSales = isS ? sToday.filter(function(x){ return (x.transfer||0) < sum([x]); }) : todayU.filter(function(x){ return !x.voided && x.transfer < sum([x]); });
      var cAmt = isS ? sCash : todayUT - todayUTr, cExp = isS ? dExp : cAmt;
      var cRow = function(d, coin){ var v = ccnt[d], has = v !== undefined && v !== '', n = has ? Math.max(0, parseInt(v,10)||0) : 0; if (has) cany = true; csum += n*d;
        return {den:d.toLocaleString(), swCls:'sw'+(coin?' c':''), swStyle:'background:'+DCOL[d], val: has ? v : '', sub: baht(n*d), subCls:'cz-sub'+(has?' on':''), aria:'จำนวน '+(coin ? 'เหรียญ ' : 'ธนบัตร ')+d+' บาท',
          on:function(e){ var c2 = Object.assign({}, self.state.clsCnt||{}); c2[d] = e.target.value; self.setState({clsCnt:c2}); }}; };
      var cNotes = [1000,500,100,50,20].map(function(d){ return cRow(d,0); }), cCoins = [10,5,1].map(function(d){ return cRow(d,1); });
      var cdiff = csum - cExp, cbad = cany && cdiff !== 0;
      var salesLbl = 'ขายเงินสดวันนี้ '+cSales.length+' รายการ';
      var parts = isS ? [{k:'ตั้งต้น (นับ '+thD2(dOpen.date)+')', v:baht(dOpen.amt), cls:''}, {k:salesLbl, v:'+'+baht(cAmt), cls:''}, {k:'เงินออก', v:'−'+baht(outT), cls:outT ? 'out' : ''}, {k:'เงินเข้า', v:'+'+baht(inT), cls:inT ? 'in' : ''}]
                      : [{k:salesLbl, v:baht(cAmt), cls:''}];
      var moves = mvPre.map(function(m){ return {at:m.at, kind:m.kind, note:m.note||'', hasNote:!!m.note, amt:(m.dir==='out' ? '−' : '+')+baht(m.amt), amtCls:'mono cz-ma '+m.dir}; });
      var REAS = isS ? ['ทอนเงินผิด','ลืมบันทึกการขาย','ลืมบันทึกเงินเข้า/ออก','ไม่ทราบสาเหตุ'] : ['ทอนเงินผิด','ลืมบันทึกการขาย','ลูกค้าโอนแต่บันทึกเป็นเงินสด','ไม่ทราบสาเหตุ'];
      var forgot = cbad && s.clsReason === 'ลืมบันทึกเงินเข้า/ออก', needNote = cbad && s.clsReason === 'ไม่ทราบสาเหตุ', noteMiss = needNote && !(s.clsNote||'').trim();
      cls = {open:true, isS:isS, panelCls:'mpanel mmid cz cnr '+csh, eyebrow:(isS ? 'ร้านเครื่องเขียน' : 'ร้านเครื่องแบบ')+' '+thDate(TODAY),
        expLabel: isS ? 'ควรมีในลิ้นชัก' : 'ควรมีในถุงเงินเครื่องแบบ', countLabel: isS ? 'นับเงินในลิ้นชัก' : 'นับเงินในถุงเครื่องแบบ',
        exp:baht(cExp), parts:parts, moves:moves, noMoves:!moves.length, addMove:openMv('out', '', 'cls'), owe:!!owe, oweTxt:oweTxt, sum:baht(csum),
        notes:cNotes, coins:cCoins, wait:!cany, ok:cany && cdiff===0, bad:cbad, stCls:'cz-st '+(!cany ? 'wait' : (cdiff===0 ? 'ok' : 'bad')),
        stTitle: !cany ? 'ยังไม่ได้นับ' : (cdiff===0 ? 'ตรงพอดี' : (cdiff<0 ? 'ขาด ' : 'เกิน ')+baht(Math.abs(cdiff))),
        stSub: !cany ? 'กรอกจำนวนที่นับได้ในแต่ละช่อง' : (cdiff===0 ? 'นับได้ '+baht(csum)+(isS ? ' ใช้เป็นเงินตั้งต้นพรุ่งนี้' : '') : 'นับได้ '+baht(csum)+' ควรมี '+baht(cExp)),
        whyTitle:(cdiff<0 ? 'เงินขาด ' : 'เงินเกิน ')+baht(Math.abs(cdiff))+' เกิดจากอะไร',
        chips:REAS.map(function(r){ return {label:r, cls:s.clsReason===r ? 'on' : '', pick:function(){ self.setState({clsReason:r}); }}; }),
        forgot:forgot, unknown:needNote, hint: forgot ? 'บันทึกรายการที่ลืม แล้วยอดจะคำนวณใหม่' : (needNote ? 'บันทึกเป็นเงินขาด/เกิน พรุ่งนี้เริ่มจากยอดที่นับได้' : (s.clsReason ? 'ใส่หมายเหตุเพิ่มได้ถ้ามี' : 'เลือกสาเหตุก่อนปิดยอด')),
        notePh: needNote ? 'หมายเหตุ (ต้องใส่) เช่น เงินทอนที่ไม่ได้บันทึกก่อนหน้านี้' : 'บันทึกเพิ่มเติม (ไม่บังคับ)',
        note:s.clsNote||'', onNote:function(e){ self.setState({clsNote:e.target.value}); },
        btnLabel: cbad ? (needNote ? 'ปิดยอดและบันทึกเงิน'+(cdiff<0 ? 'ขาด' : 'เกิน') : 'ปิดยอดพร้อมบันทึกส่วนต่าง') : (!isS && cany ? 'ปิดยอดและคัดลอกยอดนำส่ง' : 'ปิดยอด'),
        btnCls:'btn'+((!cany || (cbad && (!s.clsReason || forgot || noteMiss))) ? ' dis' : ''),
        close:function(){ self.setState({clsShop:null}); },
        go:function(){ var st = self.state; if (!cany){ self.flash('กรอกจำนวนที่นับได้ก่อน'); return; } if (cbad && !st.clsReason){ self.flash('เลือกสาเหตุของส่วนต่างก่อน'); return; }
          if (forgot){ self.flash('บันทึกเงินเข้า/ออกที่ลืมก่อน แล้วยอดจะคำนวณใหม่'); return; } if (noteMiss){ self.flash('ใส่หมายเหตุก่อน'); return; }
          var rec = {date:TODAY, at:nowHM(), shop:csh, opening: isS ? dOpen.amt : 0, sales:cAmt, out: isS ? outT : 0, in: isS ? inT : 0, expected:cExp, counted:csum, diff:cdiff, reason: cdiff ? st.clsReason : '', note: cdiff ? (st.clsNote||'') : '', by: DATA.myName};
          self.setState({closes:[rec].concat(st.closes||[]), clsShop:null, clsCnt:{}, clsReason:'', clsNote:''});
          if (!isS) self.setState({cpv:{k:'u', date:TODAY}});
          else self.okFlash('ปิดยอดแล้ว', baht(csum), 'คัดลอกยอดวันนี้', function(){ self.setState({toastOk:null, cpv:{k:'s', date:TODAY}}); }, (cdiff ? (cdiff<0 ? 'ขาด ' : 'เกิน ')+baht(Math.abs(cdiff))+' ' : '')+'พรุ่งนี้เริ่มที่ '+baht(csum)); }};
    }
    // money in/out of the stationery drawer
    var mvp = {open:false, kinds:[]};
    if (s.mv){
      var mm = s.mv, mIn = mm.dir==='in';
      var KO = ['นำส่งผู้บริหาร','สำรองจ่าย','ยืมชั่วคราว','ส่งทั้งลิ้นชัก (สิ้นภาคเรียน)'], KI = ['เบิกคืนสำรองจ่าย','คืนเงินที่ยืม','อื่นๆ'];
      var HINT = {'นำส่งผู้บริหาร':'ใส่เงินในถุงนำส่ง แล้วให้ผู้บริหาร', 'สำรองจ่าย':'จ่ายไปก่อนบันทึกอนุมัติ เมื่ออนุมัติแล้วให้บันทึก "เบิกคืนสำรองจ่าย"',
        'ยืมชั่วคราว':'เช่น ยืมเงินไปสมทบจ่ายเงินอุดหนุน แล้วนำมาคืนภายหลัง', 'ส่งทั้งลิ้นชัก (สิ้นภาคเรียน)':'ส่งเงินทั้งหมดให้ผู้บริหาร ภาคเรียนถัดไปเริ่มที่ 0 บาท',
        'เบิกคืนสำรองจ่าย': oweAdv ? 'สำรองจ่ายที่ยังค้างคืน '+baht(oweAdv)+' ได้คืนหลังบันทึกอนุมัติ' : 'ไม่มีสำรองจ่ายที่ค้างคืน', 'คืนเงินที่ยืม': oweLoan ? 'เงินที่ยืมไปยังค้างคืน '+baht(oweLoan) : 'ไม่มีเงินยืมที่ค้างคืน', 'อื่นๆ':'ใส่หมายเหตุว่าเงินมาจากไหน'};
      var isAll = mm.kind==='ส่งทั้งลิ้นชัก (สิ้นภาคเรียน)', mAmt = isAll ? dNow : (parseInt(mm.amt,10)||0);
      var mNeed = mm.kind==='อื่นๆ' || mm.kind==='ยืมชั่วคราว', mNoteMiss = mNeed && !(mm.note||'').trim(), over = !mIn && mAmt > dNow;
      var mOk = !!mm.kind && mAmt > 0 && !mNoteMiss && !over;
      var setMv = function(p){ self.setState({mv:Object.assign({}, self.state.mv, p)}); };
      mvp = {open:true, outCls:'segb'+(!mIn ? ' on' : ''), inCls:'segb'+(mIn ? ' on' : ''),
        setOut:function(){ setMv({dir:'out', kind:''}); }, setIn:function(){ setMv({dir:'in', kind:''}); },
        kinds:(mIn ? KI : KO).map(function(k){ return {label:k, cls:mm.kind===k ? 'on' : '', pick:function(){ setMv({kind:k}); }}; }),
        hint: HINT[mm.kind] || 'เลือกรายการด้านบน', hasHint:true, afterLbl: over ? 'หลังบันทึก (เงินไม่พอ)' : 'หลังบันทึก', isAll:isAll, notAll:!isAll, allAmt:baht(dNow),
        amt:mm.amt||'', onAmt:function(e){ setMv({amt:e.target.value}); }, note:mm.note||'', onNote:function(e){ setMv({note:e.target.value}); },
        notePh: mNeed ? 'หมายเหตุ (ต้องใส่)' : 'หมายเหตุ (ไม่บังคับ)',
        now:baht(dNow), after:baht(dNow + (mIn ? mAmt : -mAmt)), afterCls:'mono'+(over ? ' neg' : ''), over:over,
        btnCls:'btn'+(mOk ? '' : ' dis'), btnLabel: mIn ? 'บันทึกเงินเข้า' : 'บันทึกเงินออก',
        close:function(){ self.setState({mv:null}); },
        go:function(){ if (!mm.kind){ self.flash('เลือกรายการก่อน'); return; } if (mAmt <= 0){ self.flash('ใส่จำนวนเงินก่อน'); return; } if (over){ self.flash('เงินในลิ้นชักมีไม่พอ'); return; } if (mNoteMiss){ self.flash('ใส่หมายเหตุก่อน'); return; }
          var st = self.state, rec = {id:'m'+Date.now(), date:TODAY, at:nowHM(), dir:mm.dir, kind:mm.kind, amt:mAmt, note:(mm.note||'').trim(), after:!!cS, by: DATA.myName};
          self.setState({dMoves:[rec].concat(st.dMoves||[]), mv:null});
          self.okFlash(mIn ? 'บันทึกเงินเข้าแล้ว' : 'บันทึกเงินออกแล้ว', baht(mAmt), '', null, rec.kind+(rec.note ? ' '+rec.note : ''), '', function(){ self.setState({dMoves:(self.state.dMoves||[]).filter(function(x){ return x.id !== rec.id; })}); }); }};
    }
    // demand counts since each size was last restocked (cancelled deliveries ignored)
    var LASTRC = {};
    s.receipts.forEach(function(r){ var b = r.batch && s.batches[r.batch]; if (b && b.voided) return;
      var k = r.date+' '+((b && b.at) || '00:00'); if (!LASTRC[r.pid] || k > LASTRC[r.pid]) LASTRC[r.pid] = k; });
    var DMD = {};
    var dmdOpen = function(d){ return !LASTRC[d.pid] || (d.date+' '+(d.time||'00:00')) > LASTRC[d.pid]; };
    (s.demand||[]).forEach(function(d){ if (d.dropped || !dmdOpen(d)) return; var o = DMD[d.pid] = DMD[d.pid] || {n:0, sub:0, lost:0}; o.n++; if (d.got) o.sub++; else o.lost++; });
    var sinceLabel = function(pid){ return LASTRC[pid] ? ' ตั้งแต่รับเข้าครั้งล่าสุด ('+thD2(LASTRC[pid].slice(0,10))+')' : ' (ยังไม่มีประวัติรับเข้า)'; };
    var whoOf = function(d){ var nm = d.code && s.names[d.code], cl = d.code ? self.classFor(d.code, TODAY) : null;
      return {who: isOwner ? (nm || d.name || ('รหัส '+d.code)) : (d.code ? 'รหัส '+d.code : (d.name || 'ไม่ระบุ')), cls: cl || ''}; };
    var withWho = function(d){ return !!(d.code || d.name); };
    var live = function(d){ return !d.dropped && !d.fulfilled && !d.released; };
    var resvUntil = function(d){ return LASTRC[d.pid] ? addDays(LASTRC[d.pid].slice(0,10), 7) : null; };
    var arrived = (s.demand||[]).filter(function(d){ return live(d) && withWho(d) && !dmdOpen(d); })
      .sort(function(x,y){ return (x.date+x.time).localeCompare(y.date+y.time); });
    var RESV = {};
    arrived.forEach(function(d){ var r = RESV[d.pid] = RESV[d.pid] || []; if (r.length < (s.stock[d.pid]||0)) r.push(d); });
    var resvFor = function(pid, code){ return (RESV[pid]||[]).filter(function(d){ return !code || d.code !== code; }).length; };
    var isHeld = function(d){ return (RESV[d.pid]||[]).indexOf(d) >= 0; };
    var readyList = arrived.filter(function(d){ return !d.notified && isHeld(d); });
    var heldList = arrived.filter(function(d){ return d.notified && isHeld(d); });
    var staleCut = addDays(TODAY, -120);
    var waitingList = (s.demand||[]).filter(function(d){ return live(d) && withWho(d) && dmdOpen(d); });
    var setD = function(id, patch, msg){ self.setState({demand: self.state.demand.map(function(x){ return x.id===id ? Object.assign({}, x, patch) : x; })}); if (msg) self.flash(msg); };
    var pInfo = function(pid){ var p = PMAP[pid]; return Object.assign({pname:p.name, psize: p.size ? '#'+p.size : '', giCls:'gico gi-'+(p.shop==='u' ? p.tone : p.cat)+gx(p.name)}, iconFlags(p.icon, p), sFlags(p.icon, p)); };
    var waitRow = function(d){ var w = whoOf(d), stale = dmdOpen(d) && (d.confirmed || d.date) < staleCut;
      return Object.assign(pInfo(d.pid), {who:w.who, cls:w.cls || (d.code ? 'รหัส '+d.code : 'ไม่มีรหัส'), item:itemLabel(d.pid), whoLine: w.who+(w.cls ? '\u2003'+w.cls : (d.code ? ', รหัส '+d.code : '')), asked: 'ขอ '+thShort(d.date), subTxt: d.got ? 'ได้ '+(PMAP[d.got].size ? '#'+PMAP[d.got].size : PMAP[d.got].name)+' ไปก่อน' : '', hasExtra: !!(d.got || isHeld(d) || asksOf(d).length > 1),
        gotTag: d.got ? 'ได้ '+(PMAP[d.got].size ? '#'+PMAP[d.got].size : PMAP[d.got].name)+' ไปก่อน' : 'ยังไม่ได้ซื้อของ',
        gotCls: d.got ? 'dtag' : 'dtag amber', date: 'ขอเมื่อ '+thD2(d.date), askN: asksOf(d).length, multi: asksOf(d).length > 1, askLabel: 'ถาม '+asksOf(d).length+' ครั้ง', askCls: 'dtag askb'+(asksOf(d).length >= 3 ? ' hot' : ''), openAsks: function(){ self.setState({askHist:d.id}); }, canStu: isOwner && !!(d.code && s.names[d.code]), noStu: !(isOwner && d.code && s.names[d.code]), openStu: function(){ if (d.code && s.names[d.code]) self.setState({stuPop:d.code}); },
        resvText: isHeld(d) ? 'จองของไว้แล้ว' : '', hasResv: isHeld(d),
        stale: stale && isOwner, cardCls:'wrow cnr '+PMAP[d.pid].shop+(stale ? ' stale' : ''), canAct: isOwner,
        done: function(){ setD(d.id, {notified:TODAY}, 'แจ้ง '+w.who+' แล้ว (จองของไว้จนกว่าจะมารับของ)'); },
        collect: function(){ var p0 = PMAP[d.pid], key = p0.shop === 'u' ? 'cartU' : 'cartS', c = {}; c[d.pid] = 1; var ok = Object.assign({}, self.state.resvOk); ok[d.pid] = true;
          var u = {page: p0.shop === 'u' ? 'sellU' : 'sellS', sSel: p0.grp || null, dockOpen:true, waitOpen:false, stuPop:null, resvOk:ok, codeInput: d.code || '', q:'', searchOpen:false}; u[key] = c; self.setState(u); if (window.scrollTo) window.scrollTo(0,0);
          self.flash('เพิ่ม '+itemLabel(d.pid)+' ลงตะกร้าแล้ว บันทึกการขายเมื่อรับเงิน'); },
        release: function(){ setD(d.id, {released:TODAY}, 'ยกเลิกการจองของ '+w.who+' แล้ว'); },
        keep: function(){ setD(d.id, {confirmed:TODAY}, 'ยืนยันว่ายังต้องการ'); },
        drop: function(){ setD(d.id, {dropped:TODAY}, 'เอาออกจากรายการรอแล้ว'); }}); };
    var byClass = function(list){ var g = [], m = {};
      list.forEach(function(d){ var k = whoOf(d).cls || 'ไม่ทราบชั้น'; if (!m[k]){ m[k] = {cls:k, rows:[]}; g.push(m[k]); } m[k].rows.push(waitRow(d)); });
      return g.sort(function(x,y){ return x.cls.localeCompare(y.cls,'th'); }); };
    var shopOrd = function(d){ return PMAP[d.pid] && PMAP[d.pid].shop === 's' ? 1 : 0; };
    var byShop = function(list){ return list.slice().sort(function(a,b){ return shopOrd(a) - shopOrd(b); }); };
    var wq = (s.waitQ||'').trim().toLowerCase(), wshop = s.waitShop || 'all';
    var wfl = function(d){ if (wshop !== 'all' && PMAP[d.pid].shop !== wshop) return false; if (!wq) return true; var w = whoOf(d);
      return [itemLabel(d.pid), w.who, w.cls, d.code||'', d.name||''].join(' ').toLowerCase().indexOf(wq) >= 0; };
    var fReady = byShop(readyList.filter(wfl)), fHeld = byShop(heldList.filter(wfl)), fWaiting = byShop(waitingList.filter(wfl));
    var segIco = function(list, names){ return list.map(function(k, i){ if (names[i]) k['i_'+names[i]] = true; k.cls += ' segic'; return k; }); };
    var tileOf = function(list){ return list.map(function(k){ k.cls = k.cls.replace('segb', 'ftile').replace(' segic', ''); return k; }); };
    var splitB = function(list, sh){ var n = list.filter(function(d){ return PMAP[d.pid].shop === sh; }).length; return {n:String(n), cls:'wsb '+sh+(n ? '' : ' off')}; };
    // term stats: what was asked for most this term
    var termStart = '2026-05-16', tmap = {}, torder = [];
    (s.demand||[]).forEach(function(d){ if (d.dropped || d.date < termStart) return; var p0 = PMAP[d.pid]; if (!p0) return;
      if (!tmap[p0.name]){ tmap[p0.name] = {name:p0.name, total:0, sz:{}, shop:p0.shop}; torder.push(p0.name); }
      var key = p0.size ? '#'+p0.size : p0.name; var na = asksOf(d).filter(function(a){ return a.date >= termStart; }).length || 1; tmap[p0.name].sz[key] = (tmap[p0.name].sz[key]||0) + na; tmap[p0.name].total += na; });
    var TERMP = {}; (s.demand||[]).forEach(function(d){ if (d.dropped) return; var na = asksOf(d).filter(function(a){ return a.date >= termStart; }).length; if (na) TERMP[d.pid] = (TERMP[d.pid]||0) + na; });
    var termRows = torder.map(function(n){ return tmap[n]; }).sort(function(x,y){ return y.total - x.total; }).map(function(r){
      return {name:r.name, total:r.total+' ครั้ง', sizes: Object.keys(r.sz).sort(function(a,b){ return r.sz[b]-r.sz[a]; }).map(function(k){ return k+' ×'+r.sz[k]; }).join(', '), cls:'trow cnr '+r.shop}; });
    // suggested order quantity: sold since last restock + unmet requests − stock on hand
    var sugg = function(pid){
      var since = LASTRC[pid] || '0000', sold = 0;
      s.sales.forEach(function(x){ if ((x.date+' '+x.time) <= since) return; x.items.forEach(function(i){ if (i.pid === pid) sold += i.qty; }); });
      var unmet = DMD[pid] ? DMD[pid].n : 0, have = s.stock[pid]||0, q = Math.max(0, sold + unmet - have);
      return {q:q, sold:sold, unmet:unmet, have:have, since: LASTRC[pid] ? thD2(LASTRC[pid].slice(0,10)) : '', text: 'ควรสั่ง ~'+q+' ชิ้น', why: 'ขายไป '+sold+' + ถามหา '+unmet+' − มีอยู่ '+have+(LASTRC[pid] ? ' (ตั้งแต่รับเข้า '+thD2(LASTRC[pid].slice(0,10))+')' : '')}; };
    var wantedOut = Object.keys(DMD).filter(function(k){ return PMAP[k] && (s.stock[k]||0) <= 0; });
    var lowRows = PRODUCTS.filter(function(p){ return isLow(p.id); })
      .map(function(p){ var st = s.stock[p.id], wk = (pace[p.id]||0)/4, dleft = daysLeft(p.id);
        return Object.assign(pInfoT(p.id), {pid:p.id, shop:p.shop, dleft:dleft, days:'~'+dleft+' วัน', dCls:'dleft '+(dleft <= 7 ? 'r' : 'a'), meta:'ขายสัปดาห์ละ ~'+(Math.round(wk*10)/10)+', เหลือ '+st}); })
      .sort(function(x,y){ return (x.shop===y.shop ? 0 : (x.shop==='u' ? -1 : 1)) || x.dleft - y.dleft; });
    var alerts = [
      {key:'ready', tone:'g', count:readyList.length, label:'ของมาแล้ว รอแจ้ง', note: readyList.length ? whoOf(readyList[0]).who+': '+itemLabel(readyList[0].pid) : 'ไม่มีรายการ', go:function(){ self.setState({waitOpen:true, waitTab:'ready', bellOpen:false}); }},
      {key:'low', tone:'a', count:lowRows.length, label:'ใกล้หมด', note:'หมดภายใน 2 สัปดาห์', go:function(){ self.setState({lowOpen:true, bellOpen:false}); }},
      {key:'want', tone:'r', count:wantedOut.length, label:'ถูกถามหาแต่หมด', note: wantedOut.length ? itemLabel(wantedOut[0])+' ×'+DMD[wantedOut[0]].n : 'ไม่มีรายการ', go:function(){ self.go('stock',{stockFilter:'want', bellOpen:false}); }},
      {key:'price', tone:'n', count:zeroCount, label:'ยังไม่ตั้งราคา', note:'ขายไม่ได้จนกว่าจะใส่ราคา', go:function(){ self.go('stock',{stockShop:'u', stockFilter:'zero', bellOpen:false}); }}
    ].map(function(a){ a.cls = 'atile '+a.tone+(a.count ? '' : ' off'); a.ncls = 'nrow'+(a.count ? '' : ' off'); a.icls = 'nnum '+a.tone; return a; });
    alerts.forEach(function(a){ a.num = String(a.count); });
    var bellAlerts = (bkDue ? [{key:'backup', tone:'r', count:1, num:'!', label:'ถึงเวลาสำรองข้อมูล', note:'สำรองล่าสุด '+bkLast+'\u2003'+bkDays+' วันก่อน', ncls:'nrow urgent', icls:'nnum r', go:function(){ self.go('settings', {bellOpen:false}); setTimeout(function(){ var el = document.getElementById('bkcard'); if (el) el.scrollIntoView({behavior:'smooth', block:'start'}); }, 80); }}] : []).concat(alerts);
    var bellN = bellAlerts.filter(function(a){ return a.count > 0; }).length;
    var bell = {n:String(bellN), has: bellN > 0, label: bellN+' เรื่อง', cls:'bell'+(s.bellOpen ? ' open' : '')+(bkDue ? ' urgent' : ''), open: !!s.bellOpen,
      toggle:function(){ self.setState({bellOpen: !self.state.bellOpen}); }, close:function(){ self.setState({bellOpen:false}); }};
    var lowGroups = [['ภายใน 1 สัปดาห์', function(r){ return r.dleft <= 7; }], ['ภายใน 2 สัปดาห์', function(r){ return r.dleft > 7; }]]
      .map(function(g){ return {title:'หมด'+g[0], rows: lowRows.filter(g[1])}; }).filter(function(g){ return g.rows.length; });
    var lowPop = {open: !!s.lowOpen, groups: lowGroups, empty: !lowRows.length, close:function(){ self.setState({lowOpen:false}); },
      draft:function(){ self.setState({lowOpen:false}); self.go('stock',{expOpen:true, expKind:'order'}); self.flash('เปิดร่างใบสั่งซื้อจากรายการใกล้หมด'); }};


    // ---------- wanted-size popup ----------
    var want = {open:false, tiles:[], sizes:[], alts:[]};
    var wp = s.wantPop && PMAP[s.wantPop], wf = s.wantFor && PMAP[s.wantFor];
    var wcode = (s.wantCode||'').trim(), wNameFound = wcode && s.names[wcode], wCls = wcode ? self.classFor(wcode, TODAY) : null;
    if (wp){
      var dm = DMD[wp.id];
      want = {open:true, isOut:true, isFor:false, title:itemLabel(wp.id)+' หมด',
        sub: (dm ? 'ถูกถามหาแล้ว '+dm.n+' ครั้ง'+sinceLabel(wp.id)+' ' : '')+'จดไว้ช่วยตัดสินใจว่าควรสั่งไซส์ไหนเพิ่ม',
        tiles:[], alts:[], hasTiles:false, hasAlts:false,
        code: s.wantCode || '', nameVal: s.wantName || '', found: !!wNameFound, notFound: !!wcode && !wNameFound,
        foundText: wNameFound ? (isOwner ? wNameFound : 'พบในรายชื่อ')+(wCls ? '\u2003'+wCls : '') : '',
        whoHint: wcode || s.wantName ? 'จะบันทึกไว้ แจ้งนักเรียนได้เมื่อของมา' : 'ไม่บังคับ (ใส่ไว้เพื่อแจ้งนักเรียนเมื่อของมา)'};
      if (wp.shop === 'u'){
        var sib = PRODUCTS.filter(function(p){ return p.shop==='u' && p.name===wp.name; }).sort(function(a,b){ return szKey(a.size)-szKey(b.size); });
        var wi = sib.indexOf(wp), lo = [], hi = [];
        for (var a1 = wi-1; a1 >= 0 && lo.length < 2; a1--) if (s.stock[sib[a1].id] > 0 && s.prices[sib[a1].id]) lo.unshift(sib[a1]);
        for (var a2 = wi+1; a2 < sib.length && hi.length < 2; a2++) if (s.stock[sib[a2].id] > 0 && s.prices[sib[a2].id]) hi.push(sib[a2]);
        var wt = function(p, me){ var st0 = s.stock[p.id]||0;
          return {size:p.size, stock: me ? 'ที่ต้องการ' : 'เหลือ '+st0, price: me ? 'หมด' : baht(s.prices[p.id]), cls:'sztile t-'+p.tone+(me ? ' exact out' : ' near')+(!me && st0 <= 5 ? ' low' : ''),
            add: me ? function(){} : function(){ self.takeInstead(wp.id, p.id); }}; };
        var bl = {size:'', stock:'', price:'', cls:'sztile blank', add:function(){}};
        want.tiles = [0,1].map(function(k){ var q0 = lo[k - (2 - lo.length)]; return q0 ? wt(q0,false) : bl; }).concat([wt(wp,true)]).concat([0,1].map(function(k){ return hi[k] ? wt(hi[k],false) : bl; }));
        want.hasTiles = lo.length + hi.length > 0;
      } else {
        want.alts = PRODUCTS.filter(function(p){ return p.shop==='s' && p.cat===wp.cat && p.id!==wp.id && s.stock[p.id] > 0 && s.prices[p.id]; }).slice(0,4)
          .map(function(p){ return {label:p.name, cls:'chip', pick:function(){ self.takeInstead(wp.id, p.id); }}; });
        want.hasAlts = want.alts.length > 0;
      }
      want.noAlt = !want.hasTiles && !want.hasAlts;
    } else if (wf){
      var cur = s.cartWant && s.cartWant[wf.id];
      want = {open:true, isOut:false, isFor:true, title:'ลูกค้าอยากได้ไซส์ไหน?', sub:'ได้ '+itemLabel(wf.id)+' ไปแทน — เลือกไซส์ที่ต้องการจริง บันทึกกับนักเรียนในช่องด้านบนตอนกดบันทึกการขาย',
        sizes: PRODUCTS.filter(function(p){ return p.shop==='u' && p.name===wf.name && p.id!==wf.id; }).sort(function(a,b){ return szKey(a.size)-szKey(b.size); }).map(function(p){
          var st0 = s.stock[p.id]||0;
          return {label:'#'+p.size, sub: st0 > 0 ? 'เหลือ '+st0 : 'หมด', cls:'chip szchip'+(cur===p.id ? ' on' : ''), pick:function(){ var cw = Object.assign({}, self.state.cartWant); cw[wf.id] = p.id; self.setState({cartWant:cw, wantFor:null}); self.flash('จดไว้: ลูกค้าต้องการ '+itemLabel(p.id)); }}; }),
        hasCur: !!cur, clear:function(){ var cw = Object.assign({}, self.state.cartWant); delete cw[wf.id]; self.setState({cartWant:cw, wantFor:null}); }};
    }

    // ---------- sell ----------
    var shop = s.page==='sellS' ? 's' : 'u';
    var cartKey = shop==='u' ? 'cartU' : 'cartS';
    var cart = s[cartKey];
    var cartIds = Object.keys(cart).filter(function(id){ return cart[id] > 0; });
    var total = cartIds.reduce(function(a,id){ return a+cart[id]*s.prices[id]; },0);
    var setCart = function(id, delta){
      var c = Object.assign({}, self.state[cartKey]);
      var next = (c[id]||0) + delta;
      if (next > self.state.stock[id]){ self.flash('สต็อกไม่พอ'); return; }
      if (delta > 0 && !(self.state.resvOk||{})[id] && next > (self.state.stock[id]||0) - resvFor(id, (self.state.codeInput||'').trim())){ self.setState({resvAsk:id}); return; }
      var wasEmpty = Object.keys(c).filter(function(k){ return c[k] > 0; }).length === 0;
      if (next <= 0) delete c[id]; else c[id] = next;
      var nowEmpty = Object.keys(c).filter(function(k){ return c[k] > 0; }).length === 0;
      var upd = {}; upd[cartKey] = c;
      if (wasEmpty && !nowEmpty) upd.dockOpen = true;
      if (nowEmpty) upd.dockOpen = false;
      self.setState(upd);
    };
    var code = (s.codeInput||'').trim();
    var curCls = code ? this.classFor(code, TODAY) : null;
    var allCls = {};
    Object.keys(s.names).forEach(function(c){ var k = self.classFor(c,'9999-12-31'); if (k) allCls[k] = 1; });
    var classOptions = Object.keys(allCls).sort().map(function(k){ return {v:k}; });
    var clIn = (s.classInput||'').trim();
    var nameOptions = Object.keys(s.names).filter(function(c){ var k = self.classFor(c,'9999-12-31')||''; return !clIn || k.indexOf(clIn)===0; }).map(function(c){ return {v:s.names[c], c:self.classFor(c,'9999-12-31')||''}; }).sort(function(a,b){ return a.v.localeCompare(b.v,'th'); });
    var pickF = function(list, cur, kind){ var q = (cur||'').trim(), exact = list.some(function(o){ return o.v === q; });
      return list.filter(function(o){ return !q || exact || o.v.indexOf(q) >= 0; }).map(function(o){ return {v:o.v, sub:o.c||'', hasSub:!!o.c, cls:'cbopt'+(o.v===q ? ' on' : ''), pick:function(){ if (kind==='c') self._onClass({target:{value:o.v}}); else self._onName({target:{value:o.v}}); self.setState({pickOpen:null}); }}; }); };
    var pickC = pickF(classOptions, s.classInput, 'c'), pickN = pickF(nameOptions, s.nameInput, 'n');
    var pick = {classOpen: s.pickOpen==='c', nameOpen: s.pickOpen==='n', anyOpen: !!s.pickOpen, classOpts: pickC, nameOpts: pickN, classNone: !pickC.length, nameNone: !pickN.length,
      openC: function(){ self.setState({pickOpen:'c'}); }, openN: function(){ self.setState({pickOpen:'n'}); }, close: function(){ self.setState({pickOpen:null}); }};
    var lookup = {
      empty: !code,
      hint: !code && !(s.nameInput||'').trim(),
      noMatch: !code && !!(s.nameInput||'').trim(),
      found: !!(code && curCls),
      inactive: !!(code && !curCls && s.names[code]),
      notFound: !!(code && !curCls && !s.names[code]),
      name: s.names[code] || '', cls: curCls || ''
    };
    var prevSet = {}, prevList = [];
    if (code){
      s.sales.filter(function(x){ return x.code===code; }).sort(function(a,b){ return (b.date+b.time).localeCompare(a.date+a.time); }).forEach(function(x){
        x.items.forEach(function(i){ if (!prevSet[i.pid]){ prevSet[i.pid] = true; prevList.push(itemLabel(i.pid)); } });
      });
    }
    lookup.prevText = prevList.length ? 'เคยซื้อ: '+prevList.join(', ') : 'ยังไม่เคยซื้อ';
    lookup.full = lookup.name+'\u2003'+lookup.cls+'\u2003'+lookup.prevText;
    if (!code){ lookup.dockText = 'ยังไม่ได้กรอกรหัสนักเรียน (ช่องด้านบน)'; lookup.dockCls = 'lookup warn'; }
    else if (lookup.found){ lookup.dockText = lookup.name+'\u2003'+lookup.cls+'\u2003รหัส '+code; lookup.dockCls = 'lookup ok'; }
    else if (lookup.inactive){ lookup.dockText = lookup.name+', รหัส '+code+' (ไม่อยู่ในรายชื่อ)'; lookup.dockCls = 'lookup off'; }
    else { lookup.dockText = 'รหัส '+code+(s.nameInput ? '\u2003'+s.nameInput : '')+'\u2003ยังไม่มีในรายชื่อ'; lookup.dockCls = 'lookup warn'; }
    lookup.dockLk = (lookup.dockCls||'lookup off').replace('lookup','lk');
    var todayShop = s.sales.filter(function(x){ return x.date===TODAY && x.shop===shop; });
    var todayTr = todayShop.reduce(function(a,x){ return a+x.transfer; },0);
    var tally = {cash: baht(sum(todayShop)-todayTr), transfer: baht(todayTr), total: baht(sum(todayShop)), count: todayShop.length, trVal: shop === 's' ? '฿0' : baht(todayTr), trCls: 'tst tstr'+(shop === 's' ? ' off' : ''), trTitle: shop === 's' ? 'เครื่องเขียนรับเงินสดเท่านั้น' : 'โอน'};

    var uView = (function(){
      if (shop !== 'u') return {groups:[], hints:[], hasHint:false, sel:{tiles:[]}, paneCls:'uwrap'};
      var stuName = code ? (s.names[code]||'') : '', stuCls = code ? (curCls||self.classFor(code,'9999-12-31')||'') : '';
      var sg = /^(ด\.ญ\.|เด็กหญิง|นางสาว|น\.ส\.)/.test(stuName) ? 'girl' : (/^(ด\.ช\.|เด็กชาย|นาย)/.test(stuName) ? 'boy' : '');
      var sl = stuCls ? (/^อ/.test(stuCls) ? 'kg' : 'primary') : '';
      var fits = function(p){ if (!sg && !sl) return true; var lv = p.groups.indexOf('kg') >= 0 ? 'kg' : 'primary'; return (!sl || lv===sl) && (!sg || p.gender===sg || p.gender==='both'); };
      var closed = s.uClosed || {};
      var mkTile = function(p){
        var st = s.stock[p.id], price = s.prices[p.id], out = st <= 0, dis = out || !price, qty = cart[p.id]||0;
        return {size:p.size, price: price?baht(price):'ไม่มีราคา', stockLabel: out?'หมด':'เหลือ '+st,
          stockCls:'tstock'+(out?' out':(isLow(p.id)?' low':'')), cls:'st t-'+p.tone+(dis?' dis':'')+(out?' sold':'')+(prevSet[p.id]?' prev':'')+(qty?' inq':'')+(s.jumpHi===p.id?' jhi':''), jhi: s.jumpHi===p.id, prev: !!prevSet[p.id], disabledStr: dis?'true':'false',
          aria: itemLabel(p.id)+' '+(price?baht(price):'ยังไม่ตั้งราคา')+' '+(out?'หมด':'เหลือ '+st),
          qty:qty, hasQty:qty>0, resv: resvFor(p.id, code), hasResv: resvFor(p.id, code) > 0,
          add:function(){ if (out){ self.setState({wantPop:p.id, wantFor:null, wantCode:self.state.codeInput||'', wantName:''}); return; } if (dis){ self.flash('ยังไม่ตั้งราคา ตั้งได้ที่หน้าสต็อก'); return; } setCart(p.id, 1); }};
      };
      var byName = {}, order = [];
      PRODUCTS.forEach(function(p){ if (p.shop !== 'u') return; if (!byName[p.name]){ byName[p.name] = {name:p.name, p:p, prods:[]}; order.push(p.name); } byName[p.name].prods.push(p); });
      var mine = order.filter(function(n){ return fits(byName[n].p); }), others = order.filter(function(n){ return !fits(byName[n].p); });
      var selName = s.uSel && byName[s.uSel] ? s.uSel : (mine[0] || order[0]);
      var mkItem = function(n){ var g = byName[n], q = g.prods.reduce(function(a,p){ return a+(cart[p.id]||0); },0), on = n===selName;
        return Object.assign({name:n, giCls:'gico gi-'+g.p.tone+gx(n), qty:q, hasQty:q>0, cls:'uitem'+(on?' on':''), cur: on?'true':'false', pick:function(){ self.setState({uSel:n, uSheet:true}); }}, iconFlags(g.p.icon, g.p)); };
      var setClosed = function(k, v){ var c = Object.assign({}, self.state.uClosed||{}); c[k] = v; try { localStorage.setItem('uClosed', JSON.stringify(c)); } catch(e){} self.setState({uClosed:c}); };
      var mkGroup = function(k, title, names, defClosed){ var isClosed = closed[k] != null ? closed[k] : defClosed;
        return {title:title, count:String(names.length), open:!isClosed, exp: isClosed?'false':'true', wrapCls:'ugrp'+(isClosed?' closed':'')+(k==='other'?' other':''), items:names.map(mkItem), toggle:function(){ setClosed(k, !isClosed); }}; };
      var groups = [['top','เสื้อ'],['bottom','ท่อนล่าง'],['set','ชุดเต็มตัว']].map(function(c){ return mkGroup(c[0], c[1], mine.filter(function(n){ return byName[n].p.part===c[0]; }), false); }).filter(function(g){ return g.items.length; });
      if (others.length) groups.push(mkGroup('other', 'อื่นๆ', others, true));
      var hints = []; if (sg) hints.push({v: sg==='girl' ? 'หญิง' : 'ชาย'}); if (sl) hints.push({v: sl==='kg' ? 'อนุบาล' : 'ประถม'});
      var sg0 = byName[selName];
      var sel = Object.assign({name:selName, giCls:'gico gi-'+sg0.p.tone+gx(selName), tiles: sg0.prods.map(mkTile)}, iconFlags(sg0.p.icon, sg0.p));
      return {groups:groups, hints:hints, hasHint: hints.length > 0, sel:sel, paneCls:'uwrap'+(s.uSheet?' open':''), close:function(){ self.setState({uSheet:false}); }};
    })();
    var sView = (function(){
      if (shop !== 's') return {groups:[], sel:{tiles:[], gOpts:[], lOpts:[]}, paneCls:'uwrap'};
      var GI = {'สมุด':'notebook','เครื่องเขียน':'pencil','บัตร':'cardsleeve','วอกเกิ้ล':'wogscout','ผ้าพันคอ':'scarf','หมวก':'cap','เข็ม':'guidepin','เชือก':'rope'};
      var byG = {}; PRODUCTS.forEach(function(p){ if (p.shop === 's' && p.grp) (byG[p.grp] = byG[p.grp] || []).push(p); });
      var selName = s.sSel && byG[s.sSel] ? s.sSel : 'เครื่องเขียน';
      var closed = s.sClosed || {};
      var mkItem = function(n){ var ps = byG[n] || [], q = ps.reduce(function(a,p){ return a+(cart[p.id]||0); },0), on = n===selName;
        return Object.assign({name:n, giCls:'gico gi-'+(ps[0] ? ps[0].cat : 'write'), qty:q, hasQty:q>0, cls:'uitem'+(on?' on':''), cur: on?'true':'false', pick:function(){ self.setState({sSel:n, sSheet:true}); }}, sFlags(GI[n])); };
      var groups = [['gen','ทั่วไป',['สมุด','เครื่องเขียน','บัตร']],['scout','ลูกเสือ-เนตรนารี',['วอกเกิ้ล','ผ้าพันคอ','หมวก','เข็ม','เชือก']]].map(function(g){
        var isClosed = !!closed[g[0]];
        return {title:g[1], count:String(g[2].length), open:!isClosed, exp: isClosed?'false':'true', wrapCls:'ugrp'+(isClosed?' closed':''), items:g[2].map(mkItem),
          toggle:function(){ var c = Object.assign({}, self.state.sClosed||{}); c[g[0]] = !isClosed; self.setState({sClosed:c}); }};
      });
      var prods = byG[selName] || [], fg = s.sFG || '', fl = s.sFL || '';
      var uniq = function(k){ var o = {}; prods.forEach(function(p){ if (p[k]) o[p[k]] = 1; }); return Object.keys(o).length; }, hasG = uniq('sg') > 1, hasL = uniq('lv') > 1;
      var vis = prods.slice();
      var rk = function(p){ return (p.sg==='m'?0:(p.sg==='f'?1:2))*10 + (p.lv==='สามัญ'?0:(p.lv==='สำรอง'?1:2)); };
      vis = vis.map(function(p,i){ return [p,i]; }).sort(function(x,y){ return rk(x[0])-rk(y[0]) || x[1]-y[1]; }).map(function(x){ return x[0]; });
      var tiles = vis.map(function(p){
        var st = s.stock[p.id], price = s.prices[p.id], out = st <= 0, dis = out || !price, qty = cart[p.id]||0, nm = p.short || p.name;
        return Object.assign({_sg:p.sg, name:nm, sub:p.sub, hasSub:!!p.sub, price: price?baht(price):'ไม่มีราคา', stockLabel: out?'หมด':'เหลือ '+st,
          stockCls:'tstock'+(out?' out':(isLow(p.id)?' low':'')), disabledStr: dis?'true':'false', wrapCls:'sti'+(out?' out':(dis?' dis':''))+(qty?' inq':''), giCls:'gico gi-'+p.cat+(p.sg==='m'?' gm':(p.sg==='f'?' gf':'')), noQty: qty===0,
          inc:function(){ setCart(p.id, 1); }, dec:function(){ setCart(p.id, -1); },
          aria: p.name+' '+(price?baht(price):'ยังไม่ตั้งราคา')+' '+(out?'หมด':'เหลือ '+st), qty:qty, hasQty:qty>0,
          add:function(){ if (out){ self.setState({wantPop:p.id, wantFor:null, wantCode:'', wantName:''}); return; } if (dis){ self.flash('ยังไม่ตั้งราคา ตั้งได้ที่หน้าสต็อก'); return; } setCart(p.id, 1); }}, sFlags(p.icon, p));
      });
      var opt = function(list, cur, key){ return list.map(function(o){ var on = cur===o[0]; return {label:o[1], cls:'sfb'+(on?' on':''), pr: on?'true':'false', mark:o[2]||'', hasMark:!!o[2], markCls:'sfm '+(o[3]||''),
        pick:function(){ var u = {}; u[key] = on ? '' : o[0]; self.setState(u); }}; }); };
      var secs = [];
      if (hasG && tiles.length >= 4){ [['m','ลูกเสือ','♂'],['f','เนตรนารี','♀'],['','ทั้งสองแบบ','']].forEach(function(g){ var ts = tiles.filter(function(t){ return t._sg === g[0]; }); if (ts.length) secs.push({hasHead:true, title:g[1], mark:g[2], markCls:'sfm '+(g[0]||'n'), count:ts.length+' แบบ', tiles:ts}); }); }
      else secs.push({hasHead:false, tiles:tiles});
      var sel = Object.assign({name:selName, giCls:'gico gi-'+(prods[0] ? prods[0].cat : 'write'), tiles:tiles, sections:secs, none: tiles.length===0,
        hasFil: hasG || hasL, hasG:hasG, hasL:hasL, filOn: !!((hasG && fg) || (hasL && fl)),
        gOpts: opt([['m','ลูกเสือ','♂','m'],['f','เนตรนารี','♀','f']], fg, 'sFG'), lOpts: opt([['สามัญ','สามัญ'],['สำรอง','สำรอง']], fl, 'sFL'),
        clearFil: function(){ self.setState({sFG:'', sFL:''}); }}, sFlags(GI[selName]));
      return {groups:groups, sel:sel, paneCls:'uwrap'+(s.sSheet?' open':''), close:function(){ self.setState({sSheet:false}); }};
    })();
    var sell = {
      isU: shop==='u',
      title: shop==='u' ? 'ขายเครื่องแบบ' : 'ขายเครื่องเขียน',
      note: shop==='u' ? 'กรอกรหัสนักเรียนทุกครั้ง' : 'ขายเงินสด ไม่บันทึกข้อมูลนักเรียน',
      isS: shop==='s',
      chips: [['all','ทั้งหมด'],['student','ชุดนักเรียน'],['pe','ชุดพละ']].map(function(c){
        return {label:c[1], cls:'chip'+(c[0]==='all'?' all':'')+(s.uFilter===c[0]?' on':''), pick:function(){ self.setState({uFilter:c[0]}); }};
      }),
      showParts: true,
      pChips: [['top','ท่อนบน'],['bottom','ท่อนล่าง']].map(function(c){
        var on = s.pFilter===c[0], dis = (s.lFilter||'all')==='kg';
        return {label:c[1], cls:'chip sub'+(on?' on':'')+(dis?' dis':''), pressed: on?'true':'false', pick:function(){
          if (dis){ self.flash('ท่อนบน/ท่อนล่างใช้ได้เฉพาะประถม'); return; }
          self.setState(on ? {pFilter:'all'} : {pFilter:c[0], lFilter:'primary'});
        }};
      }),
      lChips: [['all','ทั้งหมด'],['primary','ประถม'],['kg','อนุบาล']].map(function(c){
        return {label:c[1], cls:'chip'+(c[0]==='all'?' all':'')+((s.lFilter||'all')===c[0]?' on':''), pick:function(){
          var pf = self.state.pFilter || 'all';
          if (c[0] !== 'primary') pf = 'all';
          self.setState({lFilter:c[0], pFilter:pf});
        }};
      }),
      gChips: [['all','ทั้งหมด'],['boy','ชาย'],['girl','หญิง']].map(function(c){
        return {label:c[1], cls:'chip'+(c[0]==='all'?' all':'')+(s.gFilter===c[0]?' on':''), pick:function(){ self.setState({gFilter:c[0]}); }};
      }),
      ucols: (function(){
        var vis = PRODUCTS.filter(function(p){
          if (p.shop !== 'u' || !p.groups) return false;
          var lvl = p.groups.indexOf('kg') >= 0 ? 'kg' : 'primary';
          var lf = s.lFilter || 'all';
          return (s.uFilter==='all' || p.groups.indexOf(s.uFilter) >= 0) && (lf==='all' || lvl===lf) && (s.gFilter==='all' || p.gender===s.gFilter || p.gender==='both');
        });
        var mkTile = function(p){
          var st = s.stock[p.id], price = s.prices[p.id], out = st <= 0, dis = out || !price, qty = cart[p.id]||0;
          return {size:p.size, price: price?baht(price):'ไม่มีราคา', stockLabel: out?'หมด':'เหลือ '+st,
            stockCls:'tstock'+(out?' out':(isLow(p.id)?' low':'')), cls:'st t-'+p.tone+(dis?' dis':'')+(out?' sold':'')+(prevSet[p.id]?' prev':''), prev: !!prevSet[p.id], disabledStr: dis?'true':'false',
            aria: itemLabel(p.id)+' '+(price?baht(price):'ยังไม่ตั้งราคา')+' '+(out?'หมด':'เหลือ '+st),
            qty:qty, hasQty:qty>0, resv: resvFor(p.id, code), hasResv: resvFor(p.id, code) > 0,
            add:function(){ if (out){ self.setState({wantPop:p.id, wantFor:null, wantCode:self.state.codeInput||'', wantName:''}); return; } if (dis){ self.flash(out?'สินค้าหมด':'ยังไม่ตั้งราคา ตั้งได้ที่หน้าสต็อก'); return; } setCart(p.id, 1); }};
        };
        return [['top','เสื้อ','shirt'],['bottom','ท่อนล่าง','pants'],['set','ชุดเต็มตัว','set']].map(function(c){
          var groups = [], byName = {};
          vis.filter(function(p){ return p.part===c[0]; }).forEach(function(p){
            if (!byName[p.name]){ byName[p.name] = Object.assign({name:p.name, giCls:'gico gi-'+p.tone+gx(p.name), tiles:[]}, iconFlags(p.icon, p)); groups.push(byName[p.name]); }
            byName[p.name].tiles.push(mkTile(p));
          });
          return Object.assign({key:c[0], title:c[1], groups:groups}, iconFlags(c[2]));
        }).filter(function(c){ return c.groups.length && (!s.pFilter || s.pFilter==='all' || s.pFilter===c.key); });
      })(),
      sCols: [['write','เขียนและลบ','pencil'],['paper','สมุดและบัตร','book'],['tools','ลูกเสือและเนตรนารี','scissors']].map(function(c){
        var tiles = PRODUCTS.filter(function(p){ return p.shop==='s' && p.cat===c[0]; }).map(function(p){
          var st = s.stock[p.id], price = s.prices[p.id], out = st <= 0, dis = out || !price, qty = cart[p.id]||0;
          return Object.assign({name:p.name, price: price?baht(price):'ไม่มีราคา', stockLabel: out?'หมด':'เหลือ '+st,
            stockCls:'tstock'+(out?' out':(isLow(p.id)?' low':'')), cls:'st t-'+p.cat+(dis?' dis':'')+(out?' sold':''), disabledStr: dis?'true':'false', wrapCls:'sti'+(out?' out':(dis?' dis':''))+(qty?' inq':''), giCls:'gico gi-'+p.cat, noQty: qty===0, inc:function(){ setCart(p.id, 1); }, dec:function(){ setCart(p.id, -1); },
            aria: p.name+' '+(price?baht(price):'ยังไม่ตั้งราคา')+' '+(out?'หมด':'เหลือ '+st), qty:qty, hasQty:qty>0,
            add:function(){ if (out){ self.setState({wantPop:p.id, wantFor:null, wantCode:self.state.codeInput||'', wantName:''}); return; } if (dis){ self.flash(out?'สินค้าหมด':'ยังไม่ตั้งราคา ตั้งได้ที่หน้าสต็อก'); return; } setCart(p.id, 1); }}, sFlags(p.icon, p));
        });
        return Object.assign({title:c[1], giCls:'gico gi-'+c[0], tiles:tiles}, sFlags(c[2]));
      }),
      cart: cartIds.map(function(id){
        var wnt = shop==='u' && s.cartWant && s.cartWant[id];
        var P0 = PMAP[id], sized = shop==='u' && !!P0.size, sib = sized ? PRODUCTS.filter(function(q){ return q.shop==='u' && q.name===P0.name && q.size; }).map(function(q){ return q.id; }) : [];
        var swap = function(pid){ if (pid === id){ self.setState({cartPick:null}); return; }
          var c = Object.assign({}, self.state[cartKey]), q = c[id]||0, room = (self.state.stock[pid]||0) - (c[pid]||0), mv = Math.min(q, room);
          if (mv <= 0){ self.flash(itemLabel(pid)+' หมด'); return; }
          if (mv < q){ c[id] = q - mv; self.flash('#'+PMAP[pid].size+' มีแค่ '+room+' ชิ้น ย้ายให้ '+mv+' ชิ้น'); } else delete c[id];
          c[pid] = (c[pid]||0) + mv;
          var upd = {cartPick:null}; upd[cartKey] = c;
          if (mv === q && self.state.cartWant && self.state.cartWant[id]){ var cw = Object.assign({}, self.state.cartWant); if (cw[id] !== pid) cw[pid] = cw[id]; delete cw[id]; upd.cartWant = cw; }
          self.setState(upd); };
        return {name:itemLabel(id), pname:P0.name, size: sized ? '#'+P0.size : '', sized:sized, plain:!sized, unit:baht(s.prices[id])+' / ชิ้น',
          zbCls:'zb btn'+(s.cartPick===id ? ' on' : ''), pickOpen: s.cartPick===id, pickStyle: s.cartPickPos ? 'left: '+s.cartPickPos.l+'px; top: '+s.cartPickPos.t+'px; width: '+s.cartPickPos.w+'px' : '',
          openPick:function(e){ if (self.state.cartPick===id){ self.setState({cartPick:null}); return; }
            var r = e && e.currentTarget ? e.currentTarget.getBoundingClientRect() : {left:16, top:400}, vw = window.innerWidth, w = vw <= 760 ? vw - 32 : Math.min(400, sib.length*64 + 22);
            var l = Math.max(16, Math.min(r.left - 14, vw - 16 - w));
            self.setState({cartPick:id, cartPickPos:{l:Math.round(l), t:Math.round(r.top - 8), w:Math.round(w)}}); }, closePick:function(){ self.setState({cartPick:null}); },
          opts: sib.map(function(pid){ var so = s.stock[pid]||0, inCart = (pid===id ? 0 : (cart[pid]||0)), avail = so - inCart > 0 && s.prices[pid] > 0;
            var sold = !(so > 0 && s.prices[pid] > 0), isW = wnt === pid;
            return {size:'#'+PMAP[pid].size, stk: sold ? (isW ? 'ต้องการ' : 'หมด') : 'เหลือ '+so, stCls: avail && isLow(pid) ? 'low' : (isW ? 'want' : ''), tag: pid===id ? 'ในรายการ' : (inCart ? 'มีแล้ว '+inCart : ''), tgCls:'tg'+(pid===id ? ' rec' : ''),
              cls:'zc'+(pid===id ? ' on' : '')+(sold ? (isW ? ' wantz' : ' soldz') : (avail || pid===id ? '' : ' out')),
              pick:function(){ if (sold){ var cw = Object.assign({}, self.state.cartWant); if (cw[id] === pid) delete cw[id]; else cw[id] = pid; self.setState({cartWant:cw, cartPick:null}); return; } swap(pid); }}; }), qty:cart[id], sub:baht(cart[id]*s.prices[id]), inc:function(){ setCart(id,1); }, dec:function(){ setCart(id,-1); },
          hasWant: !!wnt, wantLabel: wnt ? 'ต้องการ #'+PMAP[wnt].size : '',
          wantOpen:function(){ self.setState({wantFor:id, wantPop:null}); }};
      }),
      hasCart: cartIds.length > 0, cartEmpty: cartIds.length === 0,
      total: baht(total),
      clear: function(){ var u = {}; u[cartKey] = {}; self.setState(u); },
      save: function(){
        var st = self.state, c = st[cartKey], ids = Object.keys(c).filter(function(id){ return c[id] > 0; });
        if (!ids.length){ self.flash('ยังไม่มีสินค้าในรายการ'); return; }
        var cd = shop==='u' ? (st.codeInput||'').trim() : '';
        if (shop==='u' && !cd){ self.flash('กรอกรหัสนักเรียนที่ช่องด้านบนก่อนบันทึก'); return; }
        var items = ids.map(function(id){ return {pid:id, qty:c[id], price:st.prices[id]}; });
        var tot = items.reduce(function(a,i){ return a+i.qty*i.price; },0);
        var isTr = shop==='u' && st.payMode==='transfer';
        if (isTr && !st.transferTime){ self.setState({dockOpen:true}); self.flash('กรอกเวลาที่โอนตามสลิปก่อนบันทึก'); return; }
        var tr = isTr ? tot : 0;
        var now = new Date(), time = String(now.getHours()).padStart(2,'0')+':'+String(now.getMinutes()).padStart(2,'0');
        var sale = {id:'n'+Date.now(), date:TODAY, time:time, shop:shop, code:cd, items:items, transfer:tr, transferTime: isTr ? st.transferTime : '', slip: isTr && st.trSlip ? st.trSlip.name : '', seller: DATA.myName, tempName: st.names[cd] ? '' : (st.nameInput||'')};
        var stock = Object.assign({}, st.stock); items.forEach(function(i){ stock[i.pid] -= i.qty; });
        var flagged = shop==='u' && !self.classFor(cd, TODAY);
        var newDemand = shop==='u' ? ids.filter(function(id){ return st.cartWant && st.cartWant[id]; }).map(function(id){ return {id:'d'+Date.now()+id, date:TODAY, time:time, pid:st.cartWant[id], got:id, code:cd}; }) : [];
        items.forEach(function(i){ if (shop==='u' && st.cartWant && st.cartWant[i.pid]) i.wanted = st.cartWant[i.pid]; });
        var dm2 = st.demand; newDemand.forEach(function(e){ dm2 = addDemand(dm2, e); });
        if (cd) dm2 = dm2.map(function(d){ return (d.code === cd && !d.fulfilled && items.some(function(i){ return i.pid === d.pid; })) ? Object.assign({}, d, {fulfilled:TODAY}) : d; });
        var upd = {sales:[sale].concat(st.sales), stock:stock, codeInput:'', classInput:'', nameInput:'', tempName:'', payMode:'cash', transferTime:'', trSlip:null, dockOpen:false, demand: dm2, resvOk:{}}; upd[cartKey] = {}; if (shop==='u') upd.cartWant = {};
        self.setState(upd); var upd2 = {};
        if (st.offline) upd2.pending = (st.pending||0) + 1;
        if (upd2.pending) self.setState(upd2);
        var nPcs = items.reduce(function(a,i){ return a+i.qty; },0), prevForm = {codeInput:st.codeInput, classInput:st.classInput, nameInput:st.nameInput, tempName:st.tempName, payMode:st.payMode, transferTime:st.transferTime, trSlip:st.trSlip, cartWant:st.cartWant, demand:st.demand};
        var undoSale = function(){ var s2 = self.state; if (!s2.sales.some(function(x){ return x.id === sale.id; })) return; var sk2 = Object.assign({}, s2.stock); items.forEach(function(i){ sk2[i.pid] += i.qty; });
          var u2 = Object.assign({}, prevForm, {sales: s2.sales.filter(function(x){ return x.id !== sale.id; }), stock: sk2}); u2[cartKey] = c; if (shop !== 'u'){ delete u2.cartWant; } if (st.offline && s2.pending) u2.pending = s2.pending - 1; self.setState(u2);
          self.okFlash('เลิกทำแล้ว', '', '', null, 'คืนสต็อก '+nPcs+' ชิ้น ของกลับมาอยู่ในตะกร้าให้แก้ต่อ', 'info'); };
        self.okFlash(st.offline ? 'บันทึกในเครื่องแล้ว' : 'บันทึกการขายแล้ว', baht(tot), 'ดูรายการ', function(){ self.setState({toastOk:null}); self.go('history'); self.setState({saleId:sale.id}); }, (isTr ? 'โอน '+sale.transferTime : 'เงินสด')+(cd ? '\u2003'+(st.names[cd] || st.nameInput || 'รหัส '+cd) : '')+(flagged ? '\u2003ติดธง: รหัสยังไม่มีในรายชื่อ' : ''), '', undoSale);
      }
    };

    sell.ucolsEmpty = sell.isU && sell.ucols.length === 0;
    var cartQty = cartIds.reduce(function(a,id){ return a+cart[id]; },0);
    sell.count = cartQty;
    sell.countCls = 'dcount'+(cartQty?'':' zero');
    var fitRows = window.innerWidth > 760 ? Math.max(2, Math.round(((s.dockDrag ? s.dockDrag.h : s.dockH) || 96) / 48)) : 2;
    sell.moreHint = cartIds.length > fitRows ? 'เลื่อนดูอีก '+(cartIds.length-fitRows)+' รายการ' : '';
    sell.summary = cartIds.length ? cartIds.map(function(id){ return itemLabel(id)+' ×'+cart[id]; }).join(', ') : 'ยังไม่มีสินค้า แตะสินค้าเพื่อเพิ่ม';
    var needsDetail = shop==='u' && !s.dockOpen;
    sell.primaryLabel = needsDetail ? 'ชำระเงิน' : 'บันทึกการขาย';
    sell.primary = needsDetail ? function(){
      if (!cartIds.length){ self.flash('ยังไม่มีสินค้าในรายการ'); return; }
      self.setState({dockOpen:true});
    } : sell.save;
    var FL = {u:{student:'ชุดนักเรียน', pe:'ชุดพละ'}, l:{primary:'ประถม', kg:'อนุบาล'}, p:{top:'ท่อนบน', bottom:'ท่อนล่าง'}, g:{boy:'ชาย', girl:'หญิง'}};
    var activePills = [];
    if (s.uFilter && s.uFilter!=='all') activePills.push({label:FL.u[s.uFilter], remove:function(){ self.setState({uFilter:'all'}); }});
    if (s.lFilter && s.lFilter!=='all') activePills.push({label:FL.l[s.lFilter], remove:function(){ self.setState({lFilter:'all', pFilter:'all'}); }});
    if (s.pFilter && s.pFilter!=='all') activePills.push({label:FL.p[s.pFilter], remove:function(){ self.setState({pFilter:'all'}); }});
    if (s.gFilter && s.gFilter!=='all') activePills.push({label:FL.g[s.gFilter], remove:function(){ self.setState({gFilter:'all'}); }});
    activePills.forEach(function(p){ p.aria = 'ลบตัวกรอง '+p.label; });
    var visibleCount = (sell.ucols||[]).reduce(function(a,c){ return a + c.groups.reduce(function(b,g){ return b + g.tiles.length; },0); },0);
    var payChips = [['cash','เงินสด'],['transfer','โอน']].map(function(c){
      return {label:c[1], cls:'segb'+(s.payMode===c[0]?' on':''), pick:function(){ self.setState({payMode:c[0]}); }};
    });

    // ---------- history (month > day > sale) ----------
    var allSales = s0.sales;
    var hRange = s.hRange || 'all', hFrom = s.hFrom || '', hTo = s.hTo || '';
    var rng = hRange==='today' ? [TODAY, TODAY] : hRange==='7d' ? [addDays(TODAY,-6), TODAY] : hRange==='month' ? [TODAY.slice(0,8)+'01', TODAY] : hRange==='lastmonth' ? (function(){ var f = new Date(TODAY.slice(0,8)+'01T00:00:00'); f.setMonth(f.getMonth()-1); var a = f.toISOString().slice(0,10); return [a, addDays(TODAY.slice(0,8)+'01',-1)]; })() : hRange==='custom' ? [hFrom || '0000-01-01', hTo || '9999-12-31'] : null;
    var hShop = s.hShop || 'all', hq = (s.hQ||'').trim().toLowerCase(), hPay = s.hPay || 'all', hStuA = s.hStuSel || [], hSellerA = s.hSellerSel || [];
    var sellerNames = []; allSales.forEach(function(x){ if (sellerNames.indexOf(x.seller) < 0) sellerNames.push(x.seller); });
    var hList = allSales.filter(function(x){
      if (hShop !== 'all' && x.shop !== hShop) return false;
      if (rng && (x.date < rng[0] || x.date > rng[1])) return false;
      if (hPay === 'cash' && x.transfer > 0) return false;
      if (hPay === 'transfer' && !(x.transfer > 0)) return false;
      var un = self.isUnmatched(x);
      if (hStuA.length && !hStuA.some(function(k){ return k === 'has' ? !!(x.code && !un) : (k === 'unmatched' ? un : !x.code); })) return false;
      if (hSellerA.length && hSellerA.indexOf(x.seller) < 0) return false;
      if (hq){ var hay = [x.code, isOwner ? (s.names[x.code]||'') : '', x.tempName, x.seller, itemsText(x)].join(' ').toLowerCase(); if (hay.indexOf(hq) < 0) return false; }
      return true;
    }).sort(function(a,b){ return (b.date+b.time).localeCompare(a.date+a.time); });
    var hFiltering = !!(hq || rng || hPay !== 'all' || hStuA.length || hSellerA.length);
    var SHOPN = {u:'เครื่องแบบ', s:'เครื่องเขียน'};
    var saleRow = function(x){
      var tot = saleTotal(x), un = self.isUnmatched(x), cl = x.code ? self.classFor(x.code, x.date) : null, v = !!x.voided;
      var who, whoSub;
      if (x.shop === 's' && !x.code){ who = 'ขายหน้าร้าน'; whoSub = 'ไม่บันทึกนักเรียน'; }
      else if (!x.code){ who = 'ไม่ระบุนักเรียน'; whoSub = 'ไม่มีรหัสนักเรียน'; }
      else if (un){ who = 'รหัส '+x.code; whoSub = 'ยังไม่อยู่ในรายชื่อ'+(x.tempName ? '\u2003จดไว้: '+x.tempName : ''); }
      else if (isOwner){ who = s.names[x.code] || ('รหัส '+x.code); whoSub = 'รหัส '+x.code+(cl ? '\u2003'+cl : ''); }
      else { who = 'รหัส '+x.code; whoSub = cl || 'ในรายชื่อ'; }
      var tags = [];
      if (v) tags.push({label:'ยกเลิกแล้ว', cls:'dtag red'});
      if (x.transfer > 0 && x.transfer < tot) tags.push({label:'สด '+baht(tot-x.transfer)+', โอน '+baht(x.transfer), cls:'dtag'});
      else if (x.transfer > 0) tags.push({label:'โอน'+(x.transferTime ? ' '+x.transferTime : ''), cls:'dtag'});
      if (un && !v) tags.push({label:'รหัสยังไม่จับคู่', cls:'dtag amber'});
      if (x.log && x.log.length && !v) tags.push({label:'แก้ไขแล้ว', cls:'dtag grey'});
      return {time:x.time, who:who, whoSub:whoSub, items:itemsText(x), tags:tags, total:baht(tot), shopLabel:SHOPN[x.shop], shopCls:'shoplbl '+x.shop,
        cls:'srow2 '+x.shop+(v ? ' void' : ''), icoCls:'sico '+x.shop, isU:x.shop==='u', isS:x.shop==='s', open:function(){ self.setState({saleId:x.id, saleEdit:null, saleVoidAsk:false}); }};
    };
    var THMF = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    var THWD = ['อา.','จ.','อ.','พ.','พฤ.','ศ.','ส.'];
    var hMonths = [], mmap = {};
    hList.forEach(function(x){
      var mk = x.date.slice(0,7);
      if (!mmap[mk]){ var mp = mk.split('-').map(Number); mmap[mk] = {key:mk, title:THMF[mp[1]-1]+' '+(mp[0]+543), days:[], dmap:{}, n:0, t:0}; hMonths.push(mmap[mk]); }
      var M = mmap[mk];
      if (!M.dmap[x.date]){ var dd = new Date(x.date+'T00:00:00'); M.dmap[x.date] = {key:x.date, title:THWD[dd.getDay()]+' '+thD2(x.date), sales:[], n:0, t:0, u:0, sv:0, cash:0, tr:0, pcs:0}; M.days.push(M.dmap[x.date]); }
      var D = M.dmap[x.date]; D.sales.push(x);
      if (!x.voided){ var tt = saleTotal(x); D.n++; D.t += tt; M.n++; M.t += tt; if (x.shop==='u') D.u += tt; else D.sv += tt; D.tr += (x.transfer||0); D.cash += tt - (x.transfer||0); D.pcs += x.items.reduce(function(a,i){ return a+i.qty; },0); }
    });
    var firstMonth = hMonths[0] && hMonths[0].key, firstDay = hList[0] && hList[0].date;
    hMonths = hMonths.map(function(M){
      var mo = (s.hMonths && M.key in s.hMonths) ? s.hMonths[M.key] : (hFiltering || M.key === firstMonth);
      return {title:M.title, sub:M.n+' รายการ', total:baht(M.t), isOpen:mo, chevCls:'hchev'+(mo ? ' up' : ''),
        toggle:function(){ var o = Object.assign({}, self.state.hMonths||{}); o[M.key] = !mo; self.setState({hMonths:o}); },
        days: M.days.map(function(D){
          var dop = (s.hDays && D.key in s.hDays) ? s.hDays[D.key] : (hFiltering || D.key === TODAY || D.key === firstDay);
          return {title:D.title, count:D.n+' รายการ'+(D.sales.length > D.n ? ', ยกเลิก '+(D.sales.length-D.n) : ''), total:baht(D.t), isOpen:dop, cls:'hday'+(dop ? ' open' : ''), chevCls:'hchev'+(dop ? ' up' : ''),
            hasU: D.u > 0, hasS: D.sv > 0, uTot:baht(D.u), sTot:baht(D.sv),
            stats:'เงินสด '+baht(D.cash)+'\u2003โอน '+baht(D.tr)+'\u2003'+D.pcs+' ชิ้น',
            rows: D.sales.map(saleRow),
            toggle:function(){ var o = Object.assign({}, self.state.hDays||{}); o[D.key] = !dop; self.setState({hDays:o}); },
            copyU:function(){ self.setState({cpv:{k:'u', date:D.key}}); },
            copyS:function(){ self.setState({cpv:{k:'s', date:D.key}}); },
            xlsx:function(){ self.flash('ต้นแบบ: ดาวน์โหลดรายงานประจำวัน '+thDate(D.key)+' (แยกตารางเครื่องแบบและเครื่องเขียนให้แล้ว)'); }};
        })};
    });
    var hPills = [];
    var chipSet = function(list, cur, key){ return list.map(function(c){ return {label:c[1], cls:'chip'+(c[0]==='all' ? ' all' : '')+(cur===c[0] ? ' on' : ''), pick:function(){ var o = {}; o[key] = c[0]; self.setState(o); }}; }); };
    var mkMulti = function(id, list, sel, key, allLabel, unit){ var on = function(v){ return sel.indexOf(v) >= 0; };
      return {label: !sel.length ? allLabel : (sel.length <= 2 ? list.filter(function(c){ return on(c[0]); }).map(function(c){ return c[1]; }).join(', ') : sel.length+' '+unit),
        openP: s.mOpen === id+'P', openR: s.mOpen === id+'R', toggleP: function(){ self.setState({mOpen: self.state.mOpen === id+'P' ? null : id+'P'}); }, toggleR: function(){ self.setState({mOpen: self.state.mOpen === id+'R' ? null : id+'R'}); }, close: function(){ self.setState({mOpen:null}); },
        allLabel: allLabel, allCls: 'mopt mall'+(!sel.length ? ' on' : ''), pickAll: function(){ var o = {}; o[key] = []; self.setState(o); },
        opts: list.map(function(c){ return {label:c[1], dot:c[2]||'', cls:'mopt'+(on(c[0]) ? ' on' : ''), pick:function(){ var cur = (self.state[key]||[]).slice(), i = cur.indexOf(c[0]); if (i >= 0) cur.splice(i,1); else cur.push(c[0]); var o = {}; o[key] = cur; self.setState(o); }}; })}; };
    var hPayList = [['all','ทั้งหมด'],['cash','เงินสด'],['transfer','โอน']];
    var hStuList = [['all','ทั้งหมด'],['has','มีรหัสนักเรียน'],['unmatched','รหัสยังไม่จับคู่'+(unmatchedSales.length ? ' ('+unmatchedSales.length+')' : '')],['none','ไม่ระบุนักเรียน']];
    var hSellerList = [['all','ทุกคน']].concat(sellerNames.map(function(n){ return [n, n]; }));
    [[hPayList,hPay,'hPay']].forEach(function(g){
      if (g[1] !== 'all'){ var lb = g[0].filter(function(c){ return c[0]===g[1]; })[0]; hPills.push({label: (g[2]==='hSeller' ? 'ผู้ขาย: ' : '')+(lb ? lb[1] : g[1]), aria:'เอาตัวกรองออก', remove:function(){ var o = {}; o[g[2]] = 'all'; self.setState(o); }}); }
    });
    hStuA.forEach(function(k){ var lb = hStuList.filter(function(c){ return c[0]===k; })[0]; hPills.push({label: lb ? lb[1] : k, aria:'เอาตัวกรองออก', remove:function(){ self.setState({hStuSel:(self.state.hStuSel||[]).filter(function(x){ return x !== k; })}); }}); });
    hSellerA.forEach(function(k){ hPills.push({label:'ผู้ขาย: '+k, aria:'เอาตัวกรองออก', remove:function(){ self.setState({hSellerSel:(self.state.hSellerSel||[]).filter(function(x){ return x !== k; })}); }}); });
    var hRangeList = [['all','ทั้งหมด'],['today','วันนี้'],['7d','7 วันล่าสุด'],['month','เดือนนี้'],['lastmonth','เดือนที่แล้ว'],['custom','กำหนดเอง']];
    if (hRange !== 'all'){ var rl = hRange==='custom' ? (hFrom ? thD2(hFrom) : 'เริ่มต้น')+' – '+(hTo ? thD2(hTo) : 'ล่าสุด') : hRangeList.filter(function(c){ return c[0]===hRange; })[0][1];
      hPills.unshift({label:rl, aria:'เอาช่วงวันที่ออก', remove:function(){ self.setState({hRange:'all'}); }}); }
    var hist = {
      months: hMonths, empty: hMonths.length === 0, emptyQ: hMonths.length === 0 && !!(s.hQ||'').trim(), emptyF: hMonths.length === 0 && !(s.hQ||'').trim(),
      clearQ: function(){ self.setState({hQ:''}); }, clearAll: function(){ self.setState({hQ:'', hStuSel:[], hSellerSel:[], hShop:'all', hPay:'all', hRange:'all', hFrom:'', hTo:''}); }, shop: hShop, shopBox: 'sbox'+(hShop==='all' ? '' : ' '+hShop), q: s.hQ || '',
      rangeChips: chipSet(hRangeList, hRange, 'hRange'), isCustom: hRange === 'custom', from: hFrom, to: hTo,
      payChips: icoFlags(chipSet(hPayList, hPay, 'hPay'), ['', 'cash', 'transfer']), stuM: mkMulti('hstu', hStuList.slice(1), hStuA, 'hStuSel', 'ทั้งหมด', 'แบบ'), sellerM: mkMulti('hsel', hSellerList.slice(1), hSellerA, 'hSellerSel', 'ทุกคน', 'คน'),
      rangeOpts: (function(l){ return l.map(function(c){ return {v:c[0], label:c[1]}; }); })(hRangeList), range: hRange,
      on_range: function(e){ self.setState({hRange:e.target.value}); }, 
      paySeg: hPayList.map(function(c){ return {label:c[1], cls:'segb'+(hPay===c[0] ? ' on' : ''), pick:function(){ self.setState({hPay:c[0]}); }}; }),
      pills: hPills, hasPills: hPills.length > 0, pillCount: hPills.length, fOpen: !!s.hFilterOpen, fbtnCls: 'fbtn'+(hPills.length ? ' act' : ''),
      resultText: 'แสดงผล', resultN: String(hList.filter(function(x){ return !x.voided; }).length),
      summary: hList.length ? (hList.filter(function(x){ return !x.voided; }).length+' รายการ, '+baht(hList.reduce(function(a,x){ return a+(x.voided ? 0 : saleTotal(x)); },0))) : ''
    };
    // ---------- sale detail ----------
    var mg = {open:false, opts:[]};
    if (s.mgFrom && PMAP[s.mgFrom]){
      var mf = PMAP[s.mgFrom], mt = s.mgTo && PMAP[s.mgTo];
      var mCand = PRODUCTS.filter(function(p){ return p.id !== mf.id && p.shop === mf.shop && (mf.shop === 's' ? (mf.grp ? p.grp === mf.grp : p.cat === mf.cat) : p.part === mf.part); });
      var mfSt = s.stock[mf.id]||0, mtSt = mt ? (s.stock[mt.id]||0) : 0, mfP = s.prices[mf.id]||0, mtP = mt ? (s.prices[mt.id]||0) : 0;
      var mSales = s0.sales.filter(function(x){ return x.items.some(function(i){ return i.pid === mf.id; }); }).length;
      mg = {open:true, panelCls:'mpanel mmid cnr '+mf.shop, shopName: mf.shop === 'u' ? 'ร้านเครื่องแบบ' : 'ร้านเครื่องเขียน', from:itemLabel(mf.id), fromSub:'คงเหลือ '+mfSt+' ชิ้น\u2003'+(mfP ? baht(mfP) : 'ยังไม่ตั้งราคา'),
        opts: mCand.map(function(p){ return {cls:'mgopt'+(s.mgTo === p.id ? ' on' : ''), label:itemLabel(p.id), sub:'คงเหลือ '+(s.stock[p.id]||0)+' ชิ้น', price: s.prices[p.id] ? baht(s.prices[p.id]) : 'ยังไม่ตั้งราคา', pick:function(){ self.setState({mgTo:p.id}); }}; }),
        none: !mCand.length, has: !!mt, pickHint: !mt && mCand.length > 0,
        mvTitle:'ย้ายสต็อก '+mfSt+' ชิ้น', mvSub: mt ? itemLabel(mt.id)+' คงเหลือ '+mtSt+' → '+(mtSt+mfSt)+' ชิ้น' : '',
        hTitle:'ประวัติการขาย '+mSales+' รายการยังอยู่ครบ', hSub: mt ? 'แสดงชื่อเดิม พร้อมหมายเหตุว่ารวมเข้ากับ '+itemLabel(mt.id)+' แล้ว ยอดขายนับรวมกัน' : '',
        rTitle:'ซ่อน '+itemLabel(mf.id), rSub:'ไม่แสดงในหน้าขาย สต็อก และการค้นหาอีก',
        priceWarn: !!(mt && mtP && mfP && mtP !== mfP), priceText: mt ? 'ราคาไม่เท่ากัน ต่อไปจะขายในราคาของ '+itemLabel(mt.id)+' '+baht(mtP)+' (เดิม '+baht(mfP)+')' : '',
        btnCls:'btn'+(mt ? '' : ' dis'),
        close:function(){ self.setState({mgFrom:null, mgTo:null}); },
        go:function(){ var st = self.state, from = st.mgFrom, to = st.mgTo; if (!to){ self.flash('เลือกสินค้าที่จะเก็บไว้ก่อน'); return; }
          var q = st.stock[from]||0, sk = Object.assign({}, st.stock); sk[to] = (sk[to]||0) + q; sk[from] = 0; PMAP[from].mergedInto = to;
          var rec = {from:from, to:to, qty:q, date:TODAY, at:nowHM()};
          self.setState({stock:sk, merges:[rec].concat(st.merges||[]), mgFrom:null, mgTo:null, editId:null});
          self.okFlash('รวมสินค้าแล้ว', '', '', null, itemLabel(from)+' → '+itemLabel(to), '', function(){ var s2 = self.state; delete PMAP[from].mergedInto; var sk2 = Object.assign({}, s2.stock); sk2[to] -= q; sk2[from] = q;
            self.setState({stock:sk2, merges:(s2.merges||[]).filter(function(m){ return m !== rec; })}); self.okFlash('เลิกรวมสินค้าแล้ว', '', '', null, 'คืนสต็อก '+q+' ชิ้นให้ '+itemLabel(from), 'info'); }); }};
    }
    var ep0 = s.editId && PMAP[s.editId], epOld = ep0 ? (s.prices[s.editId]||0) : 0, epV = Number(s.editPrice), epHas = s.editPrice !== '' && s.editPrice != null && epV > 0;
    var epBig = epHas && epOld > 0 && (epV >= epOld*1.5 || epV <= epOld*0.5), epAck = epBig && s.editPriceAck === epV;
    var epQ = parseInt(s.editQty,10)||0, epSt = ep0 ? (s.stock[s.editId]||0) : 0;
    var epMax = ep0 ? s.receipts.filter(function(r){ return r.pid === s.editId; }).reduce(function(a,r){ return Math.max(a, r.qty||0); }, 0) : 0;
    var epQBig = epQ > Math.max(100, epMax*3), epQAck = epQBig && s.editQtyAck === epQ;
    var npe = {pHolder: epOld ? 'ตอนนี้ '+epOld : 'เช่น 240', btnCls:'btn ghost mbtn'+(epHas ? '' : ' dis'),
      btn: !epHas ? 'บันทึกราคา' : (epAck ? 'ยืนยัน '+baht(epV) : (epOld ? baht(epOld)+' → '+baht(epV) : 'ตั้งราคา '+baht(epV))),
      warn: epBig, warnText:'ราคาใหม่ต่างจากเดิม '+baht(epOld)+' มากกว่าครึ่งหนึ่ง ตรวจจำนวนเลขศูนย์อีกครั้ง',
      qPrev: epQ > 0, qPrevText:'คงเหลือ '+epSt+' → '+(epSt+epQ)+' ชิ้น', qWarn: epQBig, qWarnText:'มากกว่าที่เคยรับเข้าต่อครั้ง ('+(epMax ? epMax+' ชิ้น' : 'ยังไม่มีประวัติ')+') หลายเท่า ตรวจอีกครั้ง',
      recvLabel: epQ > 0 ? (epQAck ? 'ยืนยันรับเข้า '+epQ+' ชิ้น' : 'รับเข้า '+epQ+' ชิ้น') : 'รับเข้า', recvCls:'btn mbig'+(epQ > 0 ? '' : ' dis'),
      mergeOpen:function(){ self.setState({mgFrom:self.state.editId, mgTo:null, editId:null}); }};
    var sale = {open:false, rows:[], eRows:[], tags:[], diff:[], log:[], payTabs:[], back:[]};
    var sx = s.saleId && allSales.filter(function(x){ return x.id === s.saleId; })[0];
    if (sx){
      var sr = saleRow(sx), stot = saleTotal(sx), ed = s.saleEdit;
      var payMain = sx.transfer > 0 && sx.transfer < stot ? 'สด + โอน' : (sx.transfer > 0 ? 'โอน' : 'เงินสด');
      var paySub = sx.transfer > 0 && sx.transfer < stot ? 'สด '+baht(stot-sx.transfer)+', โอน '+baht(sx.transfer) : (sx.transfer > 0 ? (sx.transferTime ? 'เวลาในสลิป '+sx.transferTime : 'ไม่ได้ระบุเวลาในสลิป') : '');
      var eTotal = 0, diff = [];
      if (ed){
        var om = {}, nm = {}, ks = [];
        sx.items.forEach(function(i){ om[i.pid] = (om[i.pid]||0) + i.qty; if (ks.indexOf(i.pid) < 0) ks.push(i.pid); });
        ed.items.forEach(function(i){ nm[i.pid] = (nm[i.pid]||0) + i.qty; eTotal += i.qty*i.price; if (ks.indexOf(i.pid) < 0) ks.push(i.pid); });
        ks.forEach(function(k){ var d0 = (om[k]||0) - (nm[k]||0); if (d0) diff.push({label:itemLabel(k), text: d0 > 0 ? 'คืนเข้าสต็อก +'+d0 : 'หักสต็อกเพิ่ม '+(-d0), cls:'rvrow vrow '+(d0 > 0 ? 'dplus' : 'dminus'), now:s.stock[k]||0, after:(s.stock[k]||0)+d0}); });
      }
      sale = {open:true, panelCls:'mpanel wide mfix cnr '+sx.shop, when: thDate(sx.date)+' เวลา '+sx.time+' น. โดย '+sx.seller, total: baht(stot), shopLabel: SHOPN[sx.shop], shopCls:'shoplbl '+sx.shop,
        tags: sr.tags, who: sr.who, whoSub: sr.whoSub || '—', payMain: payMain, paySub: paySub || '—',
        pieces: sx.items.reduce(function(a,i){ return a+i.qty; },0)+' ชิ้น', lines: sx.items.length+' รายการ',
        canOpenStudent: isOwner && !!sx.code && !self.isUnmatched(sx),
        openStudent: function(){ self.setState({saleId:null}); self.go('student',{studentId:sx.code}); },
        rows: sx.items.map(function(i){ var mi = PMAP[i.pid].mergedInto; return {label:itemLabel(i.pid), merged: mi ? 'รวมเข้ากับ '+itemLabel(mi)+' แล้ว' : '', qty:i.qty, price:baht(i.price), line:baht(i.qty*i.price)}; }),
        voided: !!sx.voided, voidText: sx.voided ? 'เมื่อ '+sx.voided.at+' โดย '+sx.voided.by : '',
        notEditing: !ed, editing: !!ed,
        noAsk: isOwner && !sx.voided && !ed && !s.saleVoidAsk, voidAsk: isOwner && !sx.voided && !!s.saleVoidAsk, staffNote: !isOwner && !sx.voided,
        back: sx.items.map(function(i){ var nw = s.stock[i.pid]||0; return {label:itemLabel(i.pid), now:nw, after:nw+i.qty, cls:'rvrow vrow dplus'}; }),
        log: (sx.log||[]).slice().reverse(), hasLog: !!(sx.log && sx.log.length),
        isU: sx.shop === 'u', isS: sx.shop === 's'};
      if (ed){
        sale.eRows = ed.items.map(function(i, ix){
          var upd = function(q){ var it = self.state.saleEdit.items.slice(); it[ix] = Object.assign({}, it[ix], {qty: Math.max(0, q)}); self.setState({saleEdit: Object.assign({}, self.state.saleEdit, {items: it})}); };
          return {label:itemLabel(i.pid), price:baht(i.price)+'/ชิ้น', qty:i.qty, line:baht(i.qty*i.price), cls:'erow'+(i.qty ? '' : ' off')+(ed.swapIx === ix ? ' swapping' : ''),
            swap:function(){ self.setState({saleEdit: Object.assign({}, self.state.saleEdit, {adding:true, swapIx:ix, addName: sx.shop === 'u' ? PMAP[i.pid].name : ''})}); }, inc:function(){ upd(i.qty+1); }, dec:function(){ upd(i.qty-1); }};
        });
        sale.eTotal = baht(eTotal); sale.eDelta = eTotal === stot ? 'ยอดเท่าเดิม' : (eTotal > stot ? 'เพิ่มขึ้น '+baht(eTotal-stot) : 'ลดลง '+baht(stot-eTotal));
        sale.eCode = ed.code; sale.eIsTransfer = ed.pay === 'transfer'; sale.eSlip = ed.slip;
        (function(){ var si = ed.slipImg, isT = ed.pay === 'transfer';
          sale.slipRowCls = 'eslip'+(isT ? '' : ' dis'); sale.slipHas = !!(si && isT);
          if (!isT){ sale.slipCls = 'slipcard'; sale.slipTitle = 'ไม่ต้องใช้สลิป'; sale.slipNote = 'เลือก โอน เพื่อแนบสลิปและเวลา'; }
          else if (!si){ sale.slipCls = 'slipcard'; sale.slipTitle = sx.slip ? 'มีสลิปเดิม: '+sx.slip : 'ยังไม่มีสลิป'; sale.slipNote = 'ถ่ายหรือเลือกรูปสลิป แล้วเวลาจะใส่ให้เอง หรือพิมพ์เวลาเอง'; }
          else { var ok = si.amt === eTotal; sale.slipCls = 'slipcard on'+(ok ? '' : ' warn'); sale.slipTitle = 'อ่านจากสลิป: เวลา '+si.time+' น., ยอด '+baht(si.amt);
            sale.slipNote = ok ? 'ยอดตรงกับรายการขาย จะแนบสลิปไว้กับการขายนี้' : 'ยอดในสลิปไม่ตรงกับยอดขาย '+baht(eTotal)+' ตรวจอีกครั้ง'; }
          sale.slipClear = function(){ self.setState({saleEdit: Object.assign({}, self.state.saleEdit, {slipImg:null})}); }; })();
        sale.payTabs = [['cash','เงินสด'],['transfer','โอน']].map(function(p){ return {label:p[1], i_cash: p[0]==='cash', i_transfer: p[0]==='transfer', cls:'segb segic'+(ed.pay===p[0] ? ' on' : ''), pick:function(){ self.setState({saleEdit: Object.assign({}, self.state.saleEdit, {pay:p[0]})}); }}; });
        sale.diff = diff; sale.hasDiff = diff.length > 0; sale.noDiff = diff.length === 0;
        var sxProds = PRODUCTS.filter(function(p){ return p.shop === sx.shop; }), pNames = [];
        sxProds.forEach(function(p){ if (pNames.indexOf(p.name) < 0) pNames.push(p.name); });
        sale.adding = !!ed.adding; sale.notAdding = !ed.adding;
        var swapItem = ed.swapIx != null ? ed.items[ed.swapIx] : null;
        sale.pickTitle = swapItem ? 'เปลี่ยน '+itemLabel(swapItem.pid)+' เป็น' : 'เลือกสินค้า';
        sale.pickHint = swapItem ? (sx.shop === 'u' ? 'เลือกไซส์ใหม่ หรือแตะสินค้าอื่นด้านบนถ้าคีย์ผิดชนิด' : 'แตะสินค้าที่ลูกค้าซื้อจริง จำนวนเท่าเดิม') : '';
        sale.hasHint = !!swapItem;
        var grpDef = sx.shop === 'u' ? [['top','เสื้อ'],['bottom','ท่อนล่าง'],['set','ชุดเต็มตัว']] : [['write','เขียนและลบ'],['paper','สมุดและบัตร'],['tools','ลูกเสือและเนตรนารี']];
        var mkName = function(n){ var p0 = sxProds.filter(function(p){ return p.name === n; })[0]; return Object.assign({label:n, cls:'chip ichip'+(ed.addName===n ? ' on' : '')+(ed.swapIx != null && sx.shop === 's' && PMAP[ed.items[ed.swapIx].pid].name === n ? ' cur' : ''), giCls:'gico gi-'+(p0.shop==='u' ? p0.tone : p0.cat)+gx(n), pick:function(){
          var sizes = sxProds.filter(function(p){ return p.name === n; });
          if (sizes.length === 1 && !sizes[0].size){ self.addSaleItem(sizes[0].id); return; }
          self.setState({saleEdit: Object.assign({}, self.state.saleEdit, {addName:n})}); }}, iconFlags(p0.icon, p0), sFlags(p0.icon, p0)); };
        sale.addGroups = grpDef.map(function(g){ var names = pNames.filter(function(n){ var p0 = sxProds.filter(function(p){ return p.name === n; })[0]; return (sx.shop === 'u' ? p0.part : p0.cat) === g[0]; });
          return {title:g[1], names:names.map(mkName), show:names.length > 0}; }).filter(function(g){ return g.show; });
        sale.hasSizes = !!ed.addName;
        sale.addSizes = ed.addName ? sxProds.filter(function(p){ return p.name === ed.addName; }).map(function(p){ var st0 = s.stock[p.id]||0;
          var isCur = ed.swapIx != null && ed.items[ed.swapIx].pid === p.id;
          return {label:'#'+p.size, sub: isCur ? 'ที่คีย์ไว้' : (st0 ? 'เหลือ '+st0 : 'หมด'), cls:'chip szchip'+(isCur ? ' cur' : (st0 ? '' : ' dis')), pick:function(){ if (isCur){ return; } if (!st0){ self.flash(itemLabel(p.id)+' หมดสต็อก'); return; } self.addSaleItem(p.id); }}; }) : [];
      }
    }

    // ---------- stock ----------
    var stockShopChips = [['u','เครื่องแบบ'],['s','เครื่องเขียน']].map(function(c){
      return {label:c[1], cls:'segb'+(s.stockShop===c[0]?' on':''), pick:function(){ self.setState({stockShop:c[0], restockOpen:null}); }};
    });
    var stockChips = [['all','ทั้งหมด'],['low','ใกล้หมด'],['out','หมด'],['zero','ยังไม่ตั้งราคา'],['want','ถูกถามหา']].map(function(c){
      return {label:c[1], dot: c[0]==='low' ? 'cdot low' : (c[0]==='out' ? 'cdot out' : ''), i_zero: c[0]==='zero', i_want: c[0]==='want', cls:'chip'+(c[0]==='all'?' all':'')+(s.stockFilter===c[0]?' on':''), pick:function(){ self.setState({stockFilter:c[0], restockOpen:null}); }};
    });
    var stockRows = PRODUCTS.filter(function(p){
      if (p.shop !== s.stockShop) return false;
      if (s.stockFilter==='low') return s.stock[p.id] <= 5 && s.prices[p.id] > 0;
      if (s.stockFilter==='zero') return !s.prices[p.id];
      return true;
    }).map(function(p){
      var st = s.stock[p.id], pr = s.prices[p.id];
      var status, scls;
      if (!pr){ status = 'ยังไม่ตั้งราคา'; scls = 'tag warn'; }
      else if (st <= 0){ status = 'หมด'; scls = 'tag out'; }
      else if (st <= 5){ status = 'ใกล้หมด'; scls = 'tag warn'; }
      else { status = 'ปกติ'; scls = 'tag ok'; }
      var open = s.restockOpen === p.id;
      return {name:p.name, size:p.size||'—', price: pr?baht(pr):'—', stock:st, status:status, statusCls:scls, isOpen:open, btn: open?'ปิด':'รับเข้า',
        toggle:function(){ self.setState({restockOpen: open?null:p.id, restockQty:'', restockMemo:''}); }};
    });
    var receipts = s.receipts.map(function(r){ return {label:r.label, qty:'+'+r.qty, cost: r.cost ? '(ทุน '+baht(r.cost)+'/ชิ้น)' : '', date:thDate(r.date), memo:r.memo||'ไม่ระบุเลขที่'}; });

    // ---------- stock v2 ----------
    var sf = s.stockFilter || 'all';
    var stPass = function(p){ var st = s.stock[p.id], pr = s.prices[p.id];
      if (sf==='low') return isLow(p.id);
      if (sf==='out') return pr > 0 && st <= 0;
      if (sf==='zero') return !pr;
      if (sf==='want') return !!DMD[p.id];
      return true; };
    var stTone = function(p){ return p.shop==='u' ? p.tone : p.cat; };
    var openEditFor = function(p){ return function(){ self.setState({editId:p.id, editPrice: s.prices[p.id] ? String(s.prices[p.id]) : '', editQty:'', editMemo:'', editDate:TODAY}); }; };
    var mkStockTile = function(p){
      var st = s.stock[p.id], pr = s.prices[p.id], out = st <= 0;
      return {big: p.shop==='u' ? p.size : p.name, bigCls: p.shop==='u' ? 'ss' : 'sn',
        price: pr ? baht(pr) : 'ตั้งราคา', priceCls: 'sp'+(pr ? '' : ' need'),
        stock: st, stockTxt: out ? 'หมด' : st+' ชิ้น', stockCls: 'tstock'+(out ? ' out' : (isLow(p.id) ? ' low' : '')),
        want: DMD[p.id] ? DMD[p.id].n : 0, hasWant: !!DMD[p.id],
        cls: 'st t-'+stTone(p)+(out ? ' outt' : ''), cellCls: 'lc '+(out ? 'out' : (isLow(p.id) ? 'low' : 'ok'))+(p.shop==='u' ? '' : ' solo'), hasTop: p.shop==='u', cellTitle: (pr ? baht(pr) : 'ยังไม่ตั้งราคา')+(DMD[p.id] ? ', ถามหา '+DMD[p.id].n+' ครั้ง' : ''), aria: itemLabel(p.id)+' คงเหลือ '+st+(pr ? ' ราคา '+baht(pr) : ' ยังไม่ตั้งราคา'), open: openEditFor(p)};
    };
    var shopU = (s.stockShop||'u')==='u';
    var mergedRows = (s.merges||[]).filter(function(m){ return PMAP[m.from] && PMAP[m.from].shop === (shopU ? 'u' : 's') && PMAP[m.from].mergedInto; }).map(function(m){
      var p = PMAP[m.from], ask = s.unmergeAsk === m.from;
      return Object.assign({name:itemLabel(m.from), sub:'รวมเข้ากับ '+itemLabel(m.to), date:thD2(m.date), giCls:'gico gi-'+(p.shop==='u' ? p.tone : p.cat)+gx(p.name),
        btnCls:'btn sm '+(ask ? 'warnb' : 'ghost'), btnLabel: ask ? 'แตะอีกครั้งเพื่อยืนยัน' : 'ยกเลิกการรวม',
        undo:function(){ var st = self.state; if (st.unmergeAsk !== m.from){ self.setState({unmergeAsk:m.from}); return; }
          var sk = Object.assign({}, st.stock), back = Math.min(m.qty, sk[m.to]||0); sk[m.to] = (sk[m.to]||0) - back; sk[m.from] = (sk[m.from]||0) + back; delete PMAP[m.from].mergedInto;
          self.setState({stock:sk, merges:(st.merges||[]).filter(function(x){ return x !== m; }), unmergeAsk:null}); self.okFlash('ยกเลิกการรวมแล้ว', '', '', null, itemLabel(m.from)+' กลับมาขายได้ คืนสต็อก '+back+' ชิ้น', 'info'); }}, iconFlags(p.icon, p), sFlags(p.icon, p));
    });
    var colsDef = shopU ? [['top','เสื้อ','shirt'],['bottom','ท่อนล่าง','pants'],['set','ชุดเต็มตัว','set']] : [['write','เขียนและลบ','pencil'],['paper','สมุดและบัตร','book'],['tools','ลูกเสือและเนตรนารี','scissors']];
    // stock timeline: deliveries placed on real dates, red where a size was sold out
    var tlOn = s.stockView === 'tl', TL0 = termStart, TLN = Math.max(1, dayGap(TL0, TODAY));
    var tlPct = function(iso){ return 3 + 94 * Math.max(0, Math.min(1, dayGap(TL0, iso) / TLN)); };
    var tlRec = s.receipts.filter(function(r){ var b = r.batch && s.batches[r.batch]; return !(b && b.voided) && PMAP[r.pid] && PMAP[r.pid].shop === (shopU ? 'u' : 's') && r.date >= TL0; });
    var tlDates = []; tlRec.forEach(function(r){ if (tlDates.indexOf(r.date) < 0) tlDates.push(r.date); }); tlDates.sort();
    var tlHead = tlDates.map(function(d, i){ var r = tlRec.filter(function(x){ return x.date === d; })[0], gap = i ? dayGap(tlDates[i-1], d) : 0;
      return {pos:'left:'+tlPct(d)+'%', bcls:'tlbat'+(i % 2 ? ' odd' : '')+(tlPct(d) < 7 ? ' st' : (tlPct(d) > 93 ? ' en' : '')), date: Number(d.slice(8))+'/'+Number(d.slice(5,7)), memo: (r.memo||'').split(',')[0].replace('สถพ ซ ',''), gap: i ? '+'+gap+' วัน' : '', gapCls:'tlgap'+(gap > 40 ? ' w' : ''),
        open: function(){ if (r.batch) self.setState({rcvId:r.batch, rcvDoc:'memo', rcvPid:null}); }}; });
    var tlMonths = []; (function(){ var d = TL0.slice(0,8)+'01'; while (d <= TODAY){ var nx = addDays(d.slice(0,8)+'28', 4).slice(0,8)+'01'; var a = d < TL0 ? TL0 : d, b = nx > TODAY ? TODAY : nx;
      tlMonths.push({label:(tlPct(b)-tlPct(a)) >= 12 ? THM[Number(d.slice(5,7))-1] : '', pos:'left:'+tlPct(a)+'%;width:'+(tlPct(b)-tlPct(a))+'%', cls:'tlmb'+(tlMonths.length % 2 ? '' : ' alt')}); d = nx; } })();
    var tlRow = function(p){
      var lvl = s.stock[p.id]||0, ev = [];
      s0.sales.forEach(function(x){ if (x.voided || x.date > TODAY) return; x.items.forEach(function(i){ if (i.pid === p.id) ev.push({d:x.date, q:-i.qty}); }); });
      tlRec.forEach(function(r){ if (r.pid === p.id) ev.push({d:r.date, q:r.qty}); });
      ev.sort(function(a, b){ return b.d.localeCompare(a.d); });
      var spans = [], cur = TODAY;
      for (var k = 0; k < ev.length && cur > TL0; k++){ var d = ev[k].d < TL0 ? TL0 : ev[k].d; if (lvl === 0 && d < cur) spans.push([d, cur]); lvl -= ev[k].q; cur = d; }
      if (cur > TL0 && lvl === 0) spans.push([TL0, cur]);
      spans.reverse(); spans = spans.reduce(function(a, sp){ var l = a[a.length-1]; if (l && dayGap(l[1], sp[0]) <= 1) l[1] = sp[1]; else a.push(sp.slice()); return a; }, []);
      var boxes = tlRec.filter(function(r){ return r.pid === p.id; }).map(function(r){ return {q:r.qty, pos:'left:'+tlPct(r.date)+'%', title:thD2(r.date)+' '+r.memo+' รับ '+r.qty, open:function(){ if (r.batch) self.setState({rcvId:r.batch, rcvDoc:'memo', rcvPid:p.id}); }}; });
      var st = s.stock[p.id]||0, never = !boxes.length;
      var hi = !!tlHi[p.id];
      return {id:p.id, hi:hi, rowCls:'tlrow'+(hi ? ' hi' : ''), starCls:'tlstar'+(hi ? ' on' : ''), starAria:(hi ? 'เลิกไฮไลต์ ' : 'ไฮไลต์ ')+itemLabel(p.id),
        star:function(){ var h = Object.assign({}, self.state.tlHi||{}); if (h[p.id]) delete h[p.id]; else h[p.id] = 1; self.setState({tlHi:h}); tlSave(self.state.tlFold||{}, h); },
        label: p.shop === 'u' ? (p.size ? '#'+p.size : p.name) : p.name, nameCls:'tlname'+(never ? ' never' : ''), open: openEditFor(p), boxes:boxes,
        outs: spans.map(function(sp){ return {pos:'left:'+tlPct(sp[0])+'%;width:'+Math.max(0.6, tlPct(sp[1])-tlPct(sp[0]))+'%', title:'หมด '+thD2(sp[0])+' ถึง '+(sp[1] === TODAY ? 'วันนี้' : thD2(sp[1]))+' ('+dayGap(sp[0], sp[1])+' วัน)'}; }),
        stock: st, nowCls: 'tlnow'+(st <= 0 ? ' z' : (isLow(p.id) ? ' l' : ''))};
    };
    var tlFold = s.tlFold || {}, tlHi = s.tlHi || {}, tlOnlyHi = !!s.tlOnlyHi;
    var tlSave = function(fold, hi){ try { localStorage.setItem('tlPrefs', JSON.stringify({fold:fold, hi:hi})); } catch(e){} };
    var tlSetFold = function(k, v){ var f = Object.assign({}, self.state.tlFold||{}); if (v) f[k] = 1; else delete f[k]; self.setState({tlFold:f}); tlSave(f, self.state.tlHi||{}); };
    var tlHiN = Object.keys(tlHi).filter(function(id){ return PMAP[id] && PMAP[id].shop === (shopU ? 'u' : 's'); }).length;
    var stockView = [['grid','คงเหลือ'],['tl','ไทม์ไลน์']].map(function(v){ var on = (s.stockView||'grid') === v[0]; return {label:v[1], cls:'segb'+(on ? ' on' : ''), pick:function(){ self.setState({stockView:v[0]}); }}; });
    var stockCols = colsDef.map(function(c){
      var groups = [], by = {}, total = 0;
      PRODUCTS.filter(function(p){ return p.shop===(shopU?'u':'s') && (shopU ? p.part===c[0] : p.cat===c[0]); }).forEach(function(p){
        total += s.stock[p.id];
        if (!stPass(p)) return;
        if (!by[p.name]){ by[p.name] = Object.assign({name:p.name, giCls:'gico gi-'+stTone(p)+gx(p.name), tiles:[]}, iconFlags(p.icon, p), sFlags(p.icon, p)); groups.push(by[p.name]); }
        by[p.name].tiles.push(mkStockTile(p)); if (tlOn){ by[p.name].tl = by[p.name].tl || []; by[p.name].tl.push(tlRow(p)); } by[p.name].sum = (by[p.name].sum||0) + s.stock[p.id];
      });
      groups.forEach(function(g){ g.sumTxt = shopU ? g.sum+' ชิ้น' : (g.tiles[0].price); });
      return Object.assign({title:c[1], tlCls:'tlgrid'+(shopU ? '' : ' wide'), tlHeads: shopU, giCls:'gico '+(shopU ? 'gi-boy' : 'gi-'+c[0]), total:total, ledCls: shopU ? 'led' : 'led two', groups:groups, tilesCls: shopU ? 'stiles' : 'stiles wide'}, iconFlags(c[2]), sFlags(c[2]));
    }).filter(function(c){ return c.groups.length; });
    if (tlOn) stockCols = stockCols.map(function(c){
      var ck = (shopU ? 'u' : 's')+':'+c.title, cf = !!tlFold[ck];
      var gs = c.groups.map(function(g){
        var gk = ck+':'+g.name, rows = (g.tl||[]).filter(function(r){ return !tlOnlyHi || r.hi; }), gf = !!tlFold[gk] && !tlOnlyHi, nh = rows.filter(function(r){ return r.hi; }).length;
        return Object.assign({}, g, {tl: gf ? [] : rows, tlShow: rows.length > 0, gOpen: !gf, gCls:'tlg'+(gf ? ' shut' : ''), gExp: gf ? 'false' : 'true', gMeta: (shopU ? (g.tl||[]).length+' ไซส์' : '')+(nh ? (shopU ? ' · ' : '')+'★ '+nh : ''), gTgl: function(){ tlSetFold(gk, !gf); }});
      }).filter(function(g){ return g.tlShow; });
      return Object.assign({}, c, {groups: cf && !tlOnlyHi ? [] : gs, cOpen: !cf || tlOnlyHi, cCls:'ucolh lh tlch'+(cf && !tlOnlyHi ? ' shut' : ''), cExp: cf ? 'false' : 'true', cTgl: function(){ tlSetFold(ck, !cf); }, tlHas: gs.length > 0});
    }).filter(function(c){ return c.tlHas; });
    var tlAllKeys = []; if (tlOn) stockCols.forEach(function(c){ tlAllKeys.push((shopU ? 'u' : 's')+':'+c.title); });
    var tlAnyFold = tlAllKeys.some(function(k){ return tlFold[k]; }) || Object.keys(tlFold).some(function(k){ return k.indexOf((shopU ? 'u' : 's')+':') === 0; });
    var shopProds = PRODUCTS.filter(function(p){ return p.shop===(shopU?'u':'s'); });
    var cLow = shopProds.filter(function(p){ return isLow(p.id); }).length;
    var cOut = shopProds.filter(function(p){ return s.prices[p.id] > 0 && s.stock[p.id] <= 0; }).length;
    var cZero = shopProds.filter(function(p){ return !s.prices[p.id]; }).length;
    var lastRc = s.receipts.filter(function(r){ return PMAP[r.pid] && PMAP[r.pid].shop===(shopU?'u':'s'); })[0];
    var stockSum = [
      {key:'low', value:cLow, label:'ใกล้หมด', tone:'w'},
      {key:'out', value:cOut, label:'หมด', tone:'r'},
      {key:'zero', value:cZero, label:'ยังไม่ตั้งราคา', tone:'n'},
      {key:'want', value: shopProds.filter(function(p){ return DMD[p.id]; }).length, label:'ถูกถามหา', tone:'n'}
    ].map(function(k){ var tone = {low:'a', out:'r', zero:'n', want:'v'}[k.key]; return {value:k.value, label:k.label, i_low:k.key==='low', i_out:k.key==='out', i_zero:k.key==='zero', i_want:k.key==='want', pressed: sf===k.key ? 'true' : 'false', cls:'ktile '+tone+(sf===k.key ? ' on' : '')+(k.value ? '' : ' zero'), pick:function(){ self.setState({stockFilter: sf===k.key ? 'all' : k.key, restockOpen:null}); }}; });
    var stockTabs = [['u','เครื่องแบบ'],['s','เครื่องเขียน']].map(function(t){
      var on = (s.stockShop||'u')===t[0];
      return {label:t[1], isU:t[0]==='u', isS:t[0]==='s', cls:'stab '+t[0]+(on?' on':''), sel: on?'true':'false', go:function(){ self.setState({stockShop:t[0], rcShop:t[0]}); }};
    });
    // edit dialog
    var edit = {open:false, panelCls:'mpanel', title:'', sub:'', stock:0, price:'', cost:'', history:[], noHistory:true, giCls:'gico'};
    if (s.editId && PMAP[s.editId]){
      var ep = PMAP[s.editId];
      var eh = s.receipts.filter(function(r){ return r.pid===ep.id; });
      var ec = eh.filter(function(r){ return r.cost; })[0];
      var edPast = (s.demand||[]).filter(function(d){ return d.pid===ep.id && !dmdOpen(d); }).length, edm = DMD[ep.id], edl = (s.demand||[]).filter(function(d){ return d.pid===ep.id && dmdOpen(d); }).sort(function(a,b){ return (b.date+b.time).localeCompare(a.date+a.time); });
      var esg = sugg(ep.id);
      var esince = LASTRC[ep.id] ? 'ตั้งแต่รับเข้าครั้งล่าสุด '+esg.since : '';
      edit = Object.assign({open:true, panelCls:'mpanel cnr '+ep.shop, hasWant: !!edm || edPast > 0, suggText: esg.text, suggWhy: esg.why,
        sgCls:'nx-sug'+(esg.q > 0 ? ' go' : ''), sgTitle: esg.q > 0 ? 'ควรสั่ง '+esg.q+' ชิ้น' : 'ยังไม่ต้องสั่ง',
        sgWhy: esg.q > 0 ? 'ขายไป '+esg.sold+' ชิ้น'+(esg.unmet ? ' ถูกถามหา '+esg.unmet+' ครั้ง' : '')+' เหลือ '+esg.have+' ชิ้น' : 'มีอยู่ '+esg.have+' ชิ้น พอสำหรับที่ขาย'+(esg.unmet ? 'และถูกถามหา' : ''),
        sgSince: esince, hasSince: !!esince, sgTipOn: !!s.sgTip, sgTipTxt:'ขายไป '+esg.sold+' + ถามหา '+esg.unmet+' − มีอยู่ '+esg.have+' = '+esg.q+' ชิ้น', sgTipTog:function(){ self.setState({sgTip:!self.state.sgTip}); },
        askOn: !!edm, askOff: !edm, askN: edm ? edm.n : 0, askSub: edm ? edm.sub : 0, askLost: edm ? edm.lost : 0, askNote: LASTRC[ep.id] ? 'ตั้งแต่รับเข้า '+esg.since : 'ยังไม่เคยรับเข้า',
        askPast: edPast ? 'ก่อนรับเข้าครั้งล่าสุดเคยถูกถามหาอีก '+edPast+' ครั้ง' : '', hasAskPast: edPast > 0,
        wantText: (edm ? 'ลูกค้าถามหา '+edm.n+' ครั้ง'+sinceLabel(ep.id)+': ได้ไซส์อื่นแทน '+edm.sub+', ไม่ได้ซื้อ '+edm.lost : 'ยังไม่มีคนถามหาตั้งแต่รับเข้าครั้งล่าสุด')+(edPast ? ' (ก่อนรับเข้าครั้งนั้นเคยถามหา '+edPast+' ครั้ง)' : ''),
        pastWant: (s.demand||[]).filter(function(d){ return d.pid===ep.id && !dmdOpen(d); }).length,
        wantRows: edl.map(function(d){ return {date:thD2(d.date)+' '+d.time, text: d.got ? 'ได้ '+itemLabel(d.got)+' แทน' : 'ไม่ได้ซื้อของ'}; }), title:itemLabel(ep.id), sub:(ep.shop==='u'?'เครื่องแบบ':'เครื่องเขียน'), stock:s.stock[ep.id], price: s.prices[ep.id] ? baht(s.prices[ep.id]) : 'ยังไม่ตั้ง', cost: ec ? baht(ec.cost) : '—',
        history: eh.map(function(r){ var bb = r.batch && s.batches[r.batch]; return {tags: docTags(bb, r.memo), d2: thD2(r.date), cls:'ihrow'+(bb && bb.voided ? ' void' : ''), open:function(){ if (r.batch) self.setState({rcvId:r.batch, rcvDoc:'memo', rcvPid:r.pid, rcvVoidAsk:false}); }, hasDoc: !!(bb && (bb.memo || bb.po)), qty:'+'+r.qty, date:thDate(r.date), memo:(r.memo||'ไม่ระบุเลขที่')+(r.batch && s.batches[r.batch] && s.batches[r.batch].voided ? ' (ยกเลิกแล้ว)' : '')}; }), noHistory: eh.length===0, giCls:'gico gi-'+stTone(ep)}, iconFlags(ep.icon, ep), sFlags(ep.icon, ep));
    }
    // import dialog
    var lines0 = s.impLines || [];
    var incLines = lines0.filter(function(l){ return l.inc && Number(l.qty) > 0; });
    var CM = s.impMemoDoc || IMP_MEMO, CP = s.impPoDoc || IMP_PO;
    var memoD = CM.date, poD = CP.date;
    var dateChips = [['upload','วันอัปโหลด', TODAY]];
    if (s.memoLoaded) dateChips.push(['memo','วันที่บันทึก', memoD]);
    if (s.poLoaded) dateChips.push(['po','วันที่ใบสั่งซื้อ', poD]);
    var impLinesView = lines0.map(function(l, ix){
      var exists = !!PMAP[l.pid], pr = exists ? s.prices[l.pid] : 0, st = exists ? s.stock[l.pid] : 0, q = Number(l.qty)||0;
      var setLine = function(patch){ var arr = self.state.impLines.slice(); arr[ix] = Object.assign({}, arr[ix], patch); self.setState({impLines:arr}); };
      var tag, tagCls;
      if (!exists){ tag = 'ไซส์ใหม่ (สร้างให้)'; tagCls = 'tag warn'; }
      else if (!pr){ tag = 'ต้องตั้งราคา'; tagCls = 'tag warn'; }
      else { tag = 'ตรงกัน'; tagCls = 'tag ok'; }
      var match = (exists ? itemLabel(l.pid)+'\u2003'+st+' → '+(st+q) : PMAP[l.base].name+' #'+l.size+'\u2003ใหม่ 0 → '+q)+(DMD[l.pid] ? '\u2003ลูกค้าถามหา '+DMD[l.pid].n+' ครั้ง' : '')+(exists ? '\u2003'+sugg(l.pid).text : '');
      return {doc:l.doc, match:match, qty:l.qty, cost: (s.poLoaded && l.cost) ? baht(l.cost) : '—', tag:tag, tagCls:tagCls,
        needPrice: !exists || !pr, price:l.price||'', pid:'ip_'+l.k, qid:'iq_'+l.k, inc:!!l.inc, incAria:'รวมรายการ '+l.doc,
        rowCls:'ilrow'+(l.inc?'':' off'),
        toggle:function(){ setLine({inc:!l.inc}); }, onQty:function(e){ setLine({qty:e.target.value}); }, onPrice:function(e){ setLine({price:e.target.value}); }};
    });
    var totQty = incLines.reduce(function(a,l){ return a+(Number(l.qty)||0); },0);
    var withCost = incLines.filter(function(l){ return s.poLoaded && l.cost; });
    var totCost = withCost.reduce(function(a,l){ return a+l.cost*(Number(l.qty)||0); },0);
    var newCount = incLines.filter(function(l){ return !PMAP[l.pid]; }).length;
    var needP = incLines.filter(function(l){ return (!PMAP[l.pid] || !s.prices[l.pid]) && !(Number(l.price) > 0); }).length;
    var rpB = s.replaceId && s.batches[s.replaceId], rpOld = rpB ? s.receipts.filter(function(r){ return r.batch === rpB.id; }) : [], rpDiff = [];
    if (rpB){
      var om = {}, nm = {}, keys = [];
      rpOld.forEach(function(r){ om[r.pid] = (om[r.pid]||0) + r.qty; if (keys.indexOf(r.pid) < 0) keys.push(r.pid); });
      incLines.forEach(function(l){ nm[l.pid] = (nm[l.pid]||0) + Number(l.qty); if (keys.indexOf(l.pid) < 0) keys.push(l.pid); });
      rpDiff = keys.map(function(k){ var o = om[k]||0, n2 = nm[k]||0, dl = n2 - o;
        return {label: PMAP[k] ? itemLabel(k) : k, old: o, nw: n2, delta: dl === 0 ? 'ไม่เปลี่ยน' : (dl > 0 ? '+'+dl : ''+dl), cls:'rvrow '+(dl > 0 ? 'dplus' : (dl < 0 ? 'dminus' : 'dsame'))}; });
    }
    var da = s.dupAlert, dup = {open:false};
    if (da && s.batches[da.id]){
      var db = s.batches[da.id], drs = s.receipts.filter(function(r){ return r.batch === da.id; }), dLines = da.doc.lines || IMP_LINES;
      var kn = da.kind === 'po' ? 'PO' : 'บันทึกข้อความ', lacks = !db[da.kind];
      dup = {open:true, kindName: kn,
        title: {file:'ไฟล์นี้เคยอัปโหลดแล้ว', number:'เลขที่เอกสารนี้มีในระบบแล้ว', data:'รายการตรงกับล็อตที่รับเข้าแล้ว'}[da.reason],
        body: da.reason === 'file' ? 'ไฟล์ "'+da.doc.file+'" ถูกใช้รับเข้าไปแล้ว ถ้ารับเข้าอีกครั้ง สต็อกจะถูกนับซ้ำ'
          : (da.reason === 'number' ? 'มีการรับเข้าด้วยเลขที่ '+da.doc.no+' อยู่แล้ว ถ้าไฟล์นี้เป็นฉบับแก้ไข ให้ใช้แทนที่ของเดิม'
          : 'สินค้า '+da.m+' จาก '+da.n+' รายการในไฟล์นี้ ตรงกับล็อตที่รับเข้าไปแล้ว อาจเป็นของล็อตเดียวกัน'+(lacks ? ' ซึ่งยังไม่มี '+kn+' แนบไว้' : '')),
        oldTitle: batchTitle(db), oldSub: 'รับเข้า '+thDate(db.date)+'\u2003'+drs.length+' รายการ\u2003+'+drs.reduce(function(a,r){ return a+r.qty; },0)+' ชิ้น โดย '+db.by,
        newTitle: da.doc.no, newSub: 'ลงวันที่ '+thDate(da.doc.date)+'\u2003'+dLines.length+' รายการ\u2003+'+dLines.reduce(function(a,l){ return a+Number(l.qty); },0)+' ชิ้น',
        canAttach: da.reason === 'data' && lacks, attachLabel: 'แนบกับล็อตเดิม',
        canReplace: da.reason === 'number', canNew: da.reason !== 'file',
        dropCls: da.reason === 'file' ? 'btn' : 'btn ghost'};
    }
    var imp2 = {
      open: !!s.imp2Open, s0: s.imp2Step===0, s1: s.imp2Step===1, s2: s.imp2Step===2,
      replace: !!rpB, rpTitle: rpB ? batchTitle(rpB) : '', rpSub: rpB ? 'รับเข้าเมื่อ '+thDate(rpB.date)+'\u2003'+rpOld.length+' รายการ\u2003+'+rpOld.reduce(function(a,r){ return a+r.qty; },0)+' ชิ้น' : '',
      title: rpB ? 'แก้ไขล็อต / แทนที่ไฟล์' : 'นำเข้าเอกสาร', confirmLabel: rpB ? 'ยืนยันแก้ไข' : 'ยืนยันรับของเข้า',
      diff: rpDiff,
      steps: (rpB ? ['อัปโหลดไฟล์ที่ถูกต้อง','เทียบและตรวจ','เสร็จ'] : ['อัปโหลด','ตรวจรายการ','เสร็จ']).map(function(l,ix){ var dn = (s.imp2Step||0) > ix || (s.imp2Step||0) === 2; return {label:l, n:String(ix+1), done:dn, notDone:!dn, cls:'pst'+((s.imp2Step||0)===ix && !dn ? ' on' : '')+(dn ? ' done' : '')}; }),
      fill: 'width: '+[0,50,100][s.imp2Step||0]+'%',
      memoLoaded: !!s.memoLoaded, poLoaded: !!s.poLoaded, memoCls:'dz'+(s.memoLoaded?' ok':''), poCls:'dz'+(s.poLoaded?' ok':''),
      memoFile: CM.file, memoNo: CM.no, memoDate: thDate(memoD), poFile: CP.file, poNo: CP.no, poDate: thDate(poD), lineCount: ((s.poLoaded && CP.lines) || (s.memoLoaded && CM.lines) || IMP_LINES).length,
      memoBtn: s.memoLoaded ? 'เปลี่ยนไฟล์' : 'เลือกไฟล์', poBtn: s.poLoaded ? 'เปลี่ยนไฟล์' : 'เลือกไฟล์',
      dateVal: s.rDate || TODAY,
      dateChips: dateChips.map(function(d){ return {label:d[1], val:thShort(d[2]), cls:'chip'+((s.rDate||TODAY)===d[2]?' on':''), pick:function(){ self.setState({rDate:d[2]}); }}; }),
      lines: impLinesView,
      summary: 'รับเข้า '+incLines.length+' รายการ, +'+totQty+' ชิ้น, ลงวันที่ '+thDate(s.rDate||TODAY)
        + (s.poLoaded ? ', ต้นทุนรวม '+baht(totCost)+(withCost.length < incLines.length ? ' (ไม่รวม '+(incLines.length-withCost.length)+' รายการที่ไม่มีราคาใน PO)' : '') : ' (ไม่มี PO จึงไม่มีต้นทุน)')
        + (newCount ? ', สร้างไซส์ใหม่ '+newCount : '') + (needP ? ', ยังไม่ตั้งราคา '+needP+' ไซส์ (ตั้งทีหลังได้)' : ''),
      doneText: s.impDone || ''
    };
    // receipts history: one card per delivery
    var rcMode = s.rcMode || 'date', rq = (s.rcQ||'').trim().toLowerCase();
    var rcShop = s.rcShop || s.stockShop || 'all';
    var rcList = s.receipts.filter(function(r){ var sh = PMAP[r.pid] ? PMAP[r.pid].shop : 'u'; if (rcShop !== 'all' && sh !== rcShop) return false; return !rq || (r.label||'').toLowerCase().indexOf(rq) >= 0 || (r.memo||'').toLowerCase().indexOf(rq) >= 0; });
    var chevOpen = function(bid, pid){ return function(){ if (bid) self.setState({rcvId:bid, rcvDoc:'memo', rcvPid:pid||null, rcvVoidAsk:false}); }; };
    var rcGroups = [], rcItems = [];
    if (rcMode === 'item'){
      var imap = {};
      rcList.forEach(function(r){
        var nm = PMAP[r.pid] ? PMAP[r.pid].name : r.label, bb = r.batch && s.batches[r.batch];
        if (!imap[nm]){ imap[nm] = {title:nm, rows:[], q:0, cls:'dgrp icard cnr '+((PMAP[r.pid] && PMAP[r.pid].shop) || 'u'), shopLabel:{u:'เครื่องแบบ', s:'เครื่องเขียน'}[(PMAP[r.pid] && PMAP[r.pid].shop) || 'u'], shopCls:'shoplbl '+((PMAP[r.pid] && PMAP[r.pid].shop) || 'u')}; rcItems.push(imap[nm]); }
        imap[nm].rows.push({size: PMAP[r.pid] && PMAP[r.pid].size ? '#'+PMAP[r.pid].size : r.label, qty:'+'+r.qty, date:thD2(r.date), tags:docTags(bb, r.memo), open:chevOpen(r.batch, r.pid), cls:'ihrow'+(bb && bb.voided ? ' void' : '')});
        if (!(bb && bb.voided)) imap[nm].q += r.qty;
      });
      rcItems.sort(function(x,y){ return x.title.localeCompare(y.title,'th'); });
      
    } else {
      var cards = [], cmap = {};
      rcList.forEach(function(r, ix){
        var key = r.batch || ('x'+ix), bb = r.batch && s.batches[r.batch];
        if (!cmap[key]){ cmap[key] = {key:key, date:r.date, b:bb, memo:r.memo, items:[], q:0}; cards.push(cmap[key]); }
        cmap[key].items.push(r); cmap[key].q += r.qty;
      });
      var mkCard = function(c, showDate){
        var shp = (PMAP[c.items[0].pid] && PMAP[c.items[0].pid].shop) || 'u', v = !!(c.b && c.b.voided), names = c.items.slice(0,3).map(function(r){ return r.label+' ×'+r.qty; });
        return {qty:'+'+c.q, count:c.items.length+' รายการ', date: showDate ? thD2(c.date) : '', hasDate: !!showDate,
          title: (showDate ? thD2(c.date)+', ' : '')+(c.b ? 'โดย '+c.b.by+(c.b.at ? ' เวลา '+c.b.at+' น.' : '') : 'รับเข้าเอง'), tags:docTags(c.b, c.memo), preview: names.join(', ')+(c.items.length > 3 ? ' และอีก '+(c.items.length-3)+' รายการ' : ''),
          by: c.b ? 'โดย '+c.b.by : '', cls:'dlv cnr '+shp+(v ? ' void' : ''), shopLabel:{u:'เครื่องแบบ', s:'เครื่องเขียน'}[shp], shopCls:'shoplbl '+shp, open:chevOpen(c.b && c.b.id, null)};
      };
      if (rcMode === 'date'){
        var dmap = {};
        cards.forEach(function(c){ if (!dmap[c.date]){ dmap[c.date] = {key:c.date, title:thD2(c.date), cards:[], q:0}; rcGroups.push(dmap[c.date]); }
          dmap[c.date].cards.push(mkCard(c, false)); if (!(c.b && c.b.voided)) dmap[c.date].q += c.q; });
        rcGroups.sort(function(x,y){ return y.key.localeCompare(x.key); });
        rcGroups.forEach(function(g){ g.sub = g.q ? '' : 'ยกเลิกทั้งหมด'; g.hasHead = true; });
      } else {
        cards.sort(function(x,y){ var a1 = x.b ? batchTitle(x.b) : (x.memo||''), b1 = y.b ? batchTitle(y.b) : (y.memo||''); return b1.localeCompare(a1); });
        rcGroups = [{key:'all', title:'', sub:'', hasHead:false, cards: cards.map(function(c){ return mkCard(c, true); })}];
      }
    }
    var rcv = {open:false, rows:[], docTabs:[]};
    var rb = s.rcvId && s.batches[s.rcvId];
    if (rb){
      var rItems = s.receipts.filter(function(r){ return r.batch===rb.id; }), tq = 0, tc = 0, anyCost = false;
      
      var rvRows = rItems.map(function(r, i){ tq += r.qty; if (r.cost){ tc += r.cost*r.qty; anyCost = true; }
        return {n:i+1, label:r.label, qty:r.qty, cost: r.cost ? baht(r.cost) : '—', line: r.cost ? baht(r.cost*r.qty) : '—', cls:'rvrow'+(r.pid===s.rcvPid ? ' hl' : '')}; });
      var docs = []; if (rb.memo) docs.push(['memo','บันทึกข้อความ']); if (rb.po) docs.push(['po','ใบสั่งซื้อ']);
      var cur = docs.filter(function(d){ return d[0]===s.rcvDoc; })[0] || docs[0], dd = cur ? rb[cur[0]] : null;
      rcv = {open:true, tags: docTags(rb), panelCls:'mpanel wide mfix cnr '+(((rItems[0] && PMAP[rItems[0].pid]) || {shop:'u'}).shop),
        title: rb.memo ? rb.memo.no : (rb.po ? rb.po.no : (rb.memoNo ? rb.memoNo : 'รับเข้าเอง')),
        tag: rb.src==='import' ? 'นำเข้าจากเอกสาร' : 'รับเข้าเอง',
        recvDate: thDate(rb.date), dateFrom: rb.dateFrom || '', by: rb.by, at: rb.at+' น.',
        hasDocs: docs.length > 0, noDocs: docs.length === 0,
        noDocText: rb.memoNo ? 'อ้างอิงเลขที่ '+rb.memoNo+' แต่ไม่ได้แนบไฟล์ เพราะรับเข้าเองจากหน้าสินค้า' : 'รับเข้าเองจากหน้าสินค้า ไม่มีเลขที่เอกสารหรือไฟล์แนบ',
        multiDoc: docs.length > 1,
        docTabs: docs.map(function(d){ return {label:d[1], cls:'segb'+(cur && cur[0]===d[0] ? ' on' : ''), pick:function(){ self.setState({rcvDoc:d[0]}); }}; }),
        isMemo: !!(cur && cur[0]==='memo'), isPo: !!(cur && cur[0]==='po'),
        docFile: dd ? dd.file : '', docNo: dd ? dd.no : '', docDate: dd ? thDate(dd.date) : '',
        rows: rvRows, count: rvRows.length, qty: '+'+tq+' ชิ้น', totalCost: anyCost ? baht(tc) : '—',
        costNote: anyCost ? '' : 'ยังไม่มีราคาทุน (ไม่มี PO)',
        voided: !!rb.voided, voidText: rb.voided ? 'เมื่อ '+rb.voided.at+' โดย '+rb.voided.by : '',
        canAct: !rb.voided && isOwner, voidAsk: isOwner && !rb.voided && !!s.rcvVoidAsk, noAsk: isOwner && !rb.voided && !s.rcvVoidAsk, staffNote: !isOwner && !rb.voided,
        impact: rItems.map(function(r){ var nw = s.stock[r.pid]||0, af = nw - r.qty; return {label:r.label, now:nw, after:af, cls:'rvrow vrow'+(af < 0 ? ' neg' : '')}; }),
        anyNeg: rItems.some(function(r){ return (s.stock[r.pid]||0) - r.qty < 0; }),
        log: (rb.log||[]).slice().reverse(), hasLog: !!(rb.log && rb.log.length)};
    }
    var rcShops = [['all','ทั้งหมด'],['u','เครื่องแบบ'],['s','เครื่องเขียน']].map(function(c){
      return {label:c[1], cls:'chip'+(c[0]==='all' ? ' all' : ' shopchip '+c[0])+(rcShop===c[0] ? ' on' : ''), pick:function(){ self.setState({rcShop:c[0]}); }};
    });
    var rcModes = [['date','ตามวันที่'],['doc','ตามเอกสาร'],['item','ตามสินค้า']].map(function(m){
      var o = {label:m[1], cls:'segb segic'+(rcMode===m[0]?' on':''), pick:function(){ self.setState({rcMode:m[0]}); }}; o['i_'+m[0]] = true; return o;
    });
    // ---------- students ----------
    var latest = this.latestRoster();
    var sq = (s.studentQ||'').trim().toLowerCase();
    var WSTU = {};
    [[readyList,'r'],[heldList,'h'],[waitingList,'w']].forEach(function(gp){ gp[0].forEach(function(d){ if (!d.code) return; var o = WSTU[d.code] = WSTU[d.code] || {n:0, ready:0, r:0, h:0, w:0}; o.n++; o[gp[1]]++; if (!dmdOpen(d)) o.ready++; }); });
    var studentRows = Object.keys(s.names).sort().map(function(c){
      var cl = self.classFor(c,'9999-12-31');
      var buys = s.sales.filter(function(x){ return x.code===c; }).sort(function(a,b){ return b.date.localeCompare(a.date); });
      var activeNow = !!(latest && latest.classes[c]);
      return {code:c, name:s.names[c], cls:cl||'—', status: activeNow?'อยู่ในรายชื่อ':'ไม่อยู่ในรายชื่อ', statusCls: activeNow?'tag ok':'tag off',
        buys:buys.length, last: buys.length?thDate(buys[0].date):'—', open:function(){ self.go('student',{studentId:c}); }, _cl:cl||'',
        inactive: !activeNow, wTags: WSTU[c] ? [['r','wdot ok','แจ้งรับของ'],['h','wdot held','รอรับของ'],['w','wdot no','รอของ']].filter(function(k){ return WSTU[c][k[0]]; }).map(function(k){ return {dot:k[1], label:k[2], n:String(WSTU[c][k[0]])}; }) : [], hasW: !!WSTU[c]};
    }).filter(function(r){
      if (sq && r.code.indexOf(sq) < 0 && r.name.toLowerCase().indexOf(sq) < 0) return false;
      if ((s.classSel||[]).length && (s.classSel||[]).indexOf(r._cl) < 0) return false;
      var stS = s.stuStSel || [];
      if (stS.length && !stS.some(function(k){ return k === 'off' ? r.inactive : !!(WSTU[r.code] && WSTU[r.code][k]); })) return false;
      return true;
    });
    var ROOMS = ['อ.1','อ.2','อ.3','ป.1','ป.2','ป.3','ป.4/1','ป.4/2','ป.5/1','ป.5/2','ป.6/1','ป.6/2'];
    var stuSt = s.stuStatus || 'all', stA = s.stuStSel || [];
    var STU_ST = [['all','ทั้งหมด',''],['r','แจ้งรับของ','wdot ok'],['h','รอรับของ','wdot held'],['w','รอของ','wdot no'],['off','ไม่อยู่ในรายชื่อ','']];
    var stuPills = [];
    var csel = s.classSel || []; csel.forEach(function(c){ stuPills.push({label:'ชั้น '+c, remove:function(){ self.setState({classSel:(self.state.classSel||[]).filter(function(x){ return x !== c; })}); }}); });
    stA.forEach(function(k){ stuPills.push({label:STU_ST.filter(function(x){ return x[0]===k; })[0][1], remove:function(){ self.setState({stuStSel:(self.state.stuStSel||[]).filter(function(x){ return x !== k; })}); }}); });
    var stuF = {statM: mkMulti('sst', STU_ST.slice(1), stA, 'stuStSel', 'ทั้งหมด', 'สถานะ'), roomLabel: !csel.length ? 'ทุกชั้น' : (csel.length <= 3 ? ROOMS.filter(function(r){ return csel.indexOf(r) >= 0; }).join(', ') : csel.length+' ชั้น'),
      roomOpenP: s.mOpen === 'roomP', roomOpenR: s.mOpen === 'roomR', toggleP: function(){ self.setState({mOpen: self.state.mOpen === 'roomP' ? null : 'roomP'}); }, toggleR: function(){ self.setState({mOpen: self.state.mOpen === 'roomR' ? null : 'roomR'}); }, closeRoom: function(){ self.setState({mOpen:null}); },
      allCls: 'mopt mall'+(!csel.length ? ' on' : ''), pickAll: function(){ self.setState({classSel:[]}); },
      roomGroups: [['อนุบาล', ROOMS.slice(0,3)], ['ประถม', ROOMS.slice(3)]].map(function(g){ return {label:g[0], rooms:g[1].map(function(r){ var on = csel.indexOf(r) >= 0; return {label:r, cls:'mopt'+(on ? ' on' : ''), pick:function(){ var cur = (self.state.classSel||[]).slice(), i = cur.indexOf(r); if (i >= 0) cur.splice(i,1); else cur.push(r); self.setState({classSel:cur}); }}; })}; }),
      room: s.classFilter || 'all', stat: stuSt, onRoom: function(e){ self.setState({classFilter:e.target.value}); }, onStat: function(e){ self.setState({stuStatus:e.target.value}); }, open: !!s.stuFOpen, btnCls:'fbtn'+(stuPills.length ? ' act' : ''), pills:stuPills, hasPills: stuPills.length > 0, pillCount:String(stuPills.length), n:String(studentRows.length),
      toggle:function(){ self.setState({stuFOpen:!self.state.stuFOpen}); }, close:function(){ self.setState({stuFOpen:false}); }, clear:function(){ self.setState({classSel:[], classFilter:'all', stuStSel:[]}); },
      rooms: [{id:'all', label:'ทุกชั้น', cls:'chip all wide'+(s.classFilter==='all' ? ' on' : '')}].concat(ROOMS.map(function(r){ return {id:r, label:r, sub: r === 'อ.1' ? 'อนุบาล' : (r === 'ป.1' ? 'ประถม' : ''), cls:'chip'+(r.indexOf('/') > 0 ? ' half' : '')+(s.classFilter===r ? ' on' : '')}; })).map(function(o){ o.pick = function(){ self.setState({classFilter:o.id}); }; return o; }),
      stats: STU_ST.map(function(x){ return {label:x[1], dot:x[2], i_off: x[0]==='off', cls:'chip'+(x[0]==='all' ? ' all wide' : (x[0]==='off' ? ' wide' : ''))+(stuSt===x[0] ? ' on' : ''), pick:function(){ self.setState({stuStatus:x[0]}); }}; })};
    var classChips = [['all','ทุกชั้น'],['อ.','อนุบาล'],['ป.4','ป.4'],['ป.5','ป.5'],['ป.6','ป.6']].map(function(c){
      return {label:c[1], cls:'chip'+(c[0]==='all'?' all':'')+(s.classFilter===c[0]?' on':''), pick:function(){ self.setState({classFilter:c[0]}); }};
    });
    var waitChip = {n: String(Object.keys(WSTU).filter(function(c){ return s.names[c]; }).length), cls:'chip wchip'+(s.stuWait ? ' on' : ''), pick:function(){ self.setState({stuWait: !self.state.stuWait}); }};
    var seenU = {};
    var unmatchedList = unmatchedSales.filter(function(x){ if (seenU[x.code]) return false; seenU[x.code] = 1; return true; }).map(function(x){
      return {code:x.code, name: x.tempName ? '('+x.tempName+')' : '', date:thDate(x.date)};
    });

    // ---------- student detail ----------
    var sid = s.studentId;
    var stu = {code:'', name:'', cls:'', status:'', statusCls:'', spent:'', buys:0, items:0, sizes:'', history:[], purchases:[], noBuys:true, waits:[], hasWaits:false, waitTitle:''};
    if (sid && s.names[sid]){
      var sb = s.sales.filter(function(x){ return x.code===sid; }).sort(function(a,b){ return (b.date+b.time).localeCompare(a.date+a.time); });
      var sizes = []; sb.forEach(function(x){ x.items.forEach(function(i){ var l = itemLabel(i.pid); if (sizes.indexOf(l) < 0) sizes.push(l); }); });
      var act = !!(latest && latest.classes[sid]);
      stu = {code:sid, name:s.names[sid], cls: this.classFor(sid,'9999-12-31') || '—',
        status: act?'อยู่ในรายชื่อ':'ไม่อยู่ในรายชื่อ', statusCls: act?'tag ok':'tag off', inactive: !act,
        spent: baht(sb.reduce(function(a,x){ return a+saleTotal(x); },0)), buys: sb.length,
        items: sb.reduce(function(a,x){ return a+x.items.reduce(function(b,i){ return b+i.qty; },0); },0),
        sizes: sizes.length ? sizes.join(', ') : '—',
        sizeRows: (function(){ var by = {}, rows = [];
          sb.slice().sort(function(x,y){ return (y.date+y.time).localeCompare(x.date+x.time); }).forEach(function(x){ x.items.forEach(function(i){ var p = PMAP[i.pid]; if (!p || !p.size) return;
            var g = by[p.name]; if (!g){ g = by[p.name] = Object.assign({name:p.name, giCls:'gico gi-'+stTone(p)+gx(p.name), seen:{}, sizes:[], shop:p.shop}, iconFlags(p.icon, p), sFlags(p.icon, p)); rows.push(g); }
            if (!g.seen[p.size]){ g.seen[p.size] = 1; var first = g.sizes.length === 0; g.sizes.push({label:'#'+p.size, cls:'szc'+(first ? ' latest' : ''), title: first ? 'ไซส์ล่าสุด' : 'เคยซื้อ'}); } }); });
          return rows; })(),
        history: s.rosters.slice().sort(function(a,b){ return a.eff.localeCompare(b.eff); }).filter(function(r){ return r.classes[sid]; }).map(function(r){ return {cls:r.classes[sid], term:r.term, eff:thDate(r.eff)}; }),
        purchases: sb.slice().sort(function(x,y){ return (y.date+y.time).localeCompare(x.date+x.time); }).map(function(x){ return {items:itemsText(x), date:thDate(x.date)+' '+x.time, cls: self.classFor(sid, x.date) || 'ยังไม่จับคู่', total:baht(saleTotal(x)), lines: x.items.map(function(i){ var p = PMAP[i.pid]; return Object.assign({name:p.name, size: p.size ? '#'+p.size : '', qty:String(i.qty), giCls:'gico gi-'+stTone(p)}, iconFlags(p.icon, p), sFlags(p.icon, p)); })}; }),
        noBuys: sb.length === 0, noSizes: !sb.some(function(x){ return x.items.some(function(i){ return PMAP[i.pid] && PMAP[i.pid].size; }); }),
        waits: byShop([].concat(readyList, heldList, waitingList).filter(function(d){ return d.code === sid; })).map(function(d){ var r = waitRow(d), st = !dmdOpen(d) ? (d.notified ? 'แจ้งนักเรียนแล้ว, จองของไว้แล้ว' : 'ของมาแล้ว, ยังไม่ได้แจ้งนักเรียน') : 'รอของ';
          return Object.assign(pInfoT(d.pid), {meta: r.date+', '+r.gotTag, item:r.item, gotTag:r.gotTag, gotCls:r.gotCls, date:r.date, status: dmdOpen(d) ? 'รอของ' : (d.notified ? 'รอรับของ' : 'แจ้งรับของ'), dot: 'wdot '+(dmdOpen(d) ? 'no' : (d.notified ? 'held' : 'ok')), stCls:'wtag', askLabel:r.askLabel, askCls:r.askCls, openAsks:r.openAsks, multi:r.multi}); })};
      stu.hasWaits = stu.waits.length > 0; stu.waitTitle = 'กำลังรอของ ('+stu.waits.length+')';
    }

    // ---------- import ----------
    var effL = thDate(s.effDate);
    var curClasses = {};
    Object.keys(s.names).forEach(function(c){ var cl = self.classFor(c,'9999-12-31'); if (cl) curClasses[c] = cl; });
    var newRows = [], changes = [], missing = [], same = 0, matchTotal = 0;
    Object.keys(NEW_FILE.classes).forEach(function(c){
      var to = NEW_FILE.classes[c];
      if (!s.names[c]){
        var m = s.sales.filter(function(x){ return x.code===c && self.isUnmatched(x) && x.date >= s.effDate; }).length;
        matchTotal += m;
        newRows.push({code:c, name:NEW_FILE.names[c]||'', cls:to, match: m ? 'จับคู่ยอดขาย '+m+' รายการ' : 'ยังไม่มียอดขาย', matchCls: m ? 'tag ok' : 'tag off'});
      } else if (curClasses[c] && curClasses[c] !== to){
        changes.push({code:c, name:s.names[c], from:curClasses[c], to:to});
      } else same++;
    });
    if (latest) Object.keys(latest.classes).forEach(function(c){ if (!NEW_FILE.classes[c]) missing.push({code:c, name:s.names[c]||''}); });
    var hasR2 = s.rosters.length > 1;
    var imp = {
      s0: s.importStep===0, s1: s.importStep===1, s2: s.importStep===2,
      noFile: !s.fileChosen, hasFile: s.fileChosen, fileName: NEW_FILE.fileName, fileCount: Object.keys(NEW_FILE.classes).length,
      effLabel: effL, newRows:newRows, newCount:newRows.length, changes:changes, changeCount:changes.length, missing:missing, missingCount:missing.length,
      sameCount: same, matchTotal: matchTotal, doneTerm: s.doneTerm, doneEff: s.doneEff
    };
    var importSteps = ['เลือกไฟล์และวันที่','ตรวจสอบ','เสร็จสิ้น'].map(function(l,ix){
      var dn = s.importStep>ix || s.importStep===2;
      return {label:l, n:String(ix+1), done:dn, notDone:!dn, cls:'pst'+(s.importStep===ix && !dn?' on':'')+(dn?' done':'')};
    });
    var rosterRows = s.rosters.slice().sort(function(a,b){ return b.eff.localeCompare(a.eff); }).map(function(r, ix){
      var future = r.eff > TODAY;
      var st = future ? 'เริ่มใช้ในอนาคต' : (ix === s.rosters.filter(function(z){ return z.eff > TODAY; }).length ? 'ใช้อยู่' : 'ก่อนหน้า');
      return {term:r.term, count:Object.keys(r.classes).length, eff:thDate(r.eff), status:st, statusCls: st==='ใช้อยู่' ? 'tag ok' : 'tag off'};
    });

    // ---------- settings ----------
    var staffRows = s.staff.map(function(st, ix){
      return {name:st.name, email:st.email, remove:function(){ var l = self.state.staff.slice(); l.splice(ix,1); self.setState({staff:l}); self.flash('นำ '+st.name+' ออกแล้ว'); }};
    });
    var xfer = {s0: s.transferStep===0, s1: s.transferStep===1, s2: s.transferStep===2};

    return {
      navGroups: navGroups, roleChips: roleChips, user: user,
      q: s.q, searchPh: isOwner ? 'ค้นหาสินค้า รหัสนักเรียน หรือชื่อ' : 'ค้นหาสินค้าหรือรหัสนักเรียน',
      onQ: function(e){ self.setState({q:e.target.value, searchOpen:true}); },
      openSearch: function(){ self.setState({searchOpen:true}); },
      closeSearch: function(){ self.setState({searchOpen:false, q:''}); },
      showResults: s.searchOpen && q.length > 0, searchSections: sections, searchEmpty: q.length > 0 && sections.length === 0 && sizeHits.length === 0, sizeHits: sizeHits, hasSizeHits: sizeHits.length > 0, sizeTitle: 'ไซส์ “'+(s.q||'').trim()+'” และไซส์ใกล้เคียงที่มีของ',
      show: show,
      goSellU: function(){ self.go('sellU'); },
      dash: dash, periods: periods, days: days, alerts: alerts, top: top, topEmpty: top.length === 0, shopCards: shopCards, money: money, area: area, bell: bell, lowPop: lowPop,
      sell: sell, uv: uView, sv: sView,
      offline: !!s.offline, pending: s.pending||0, hasPending: (s.pending||0) > 0,
      okt: (function(){ var t = s.toastOk; if (!t) return {on:false}; var sellPg = s.page==='sellU' || s.page==='sellS';
        return {on:true, cls:'okt'+(t.cls?' '+t.cls:'')+(sellPg?' up':''), text:t.text, amt:t.amt, hasAmt:!!t.amt, sub:t.sub, hasSub:!!t.sub, actLabel:t.actLabel, hasAct:!!(t.act && t.actLabel), act:t.act, hasUndo:!!t.undo, undo:function(){ var u = t.undo; clearTimeout(self._ot); self.setState({toastOk:null}); if (u) u(); }, isInfo: t.cls==='info', notInfo: t.cls!=='info', hasActs: !!(t.undo || (t.act && t.actLabel))}; })(),
      lg: {on: !!s.auth && s.auth !== 'in', login: s.auth === 'login', denied: s.auth === 'denied',
        signIn: function(){ DATA.signIn(); }, other: function(){ DATA.signOut(); }, retry: function(){ location.reload(); }},
      demo: {offLabel: s.offline ? 'กลับมาออนไลน์' : 'จำลองออฟไลน์',
        toggleOff: function(){ if (self.state.offline){ var n = self.state.pending||0; self.setState({offline:false, pending:0}); self.okFlash(n ? 'กลับมาออนไลน์ ส่งยอดที่ค้างไว้ '+n+' รายการแล้ว' : 'กลับมาออนไลน์แล้ว', '', '', null, '', 'sync'); } else { self.setState({offline:true}); } },
        showLogin: function(){ self.setState({auth:'login'}); }, showDenied: function(){ self.setState({auth:'denied'}); }},
 codeInput: s.codeInput, tempName: s.tempName, transferInput: s.transferInput, lookup: lookup,
      pick: pick, classOptions: classOptions, nameOptions: nameOptions, classInput: s.classInput||'', nameInput: s.nameInput||'',
      onClass: (self._onClass = function(e){
        var v = e.target.value, st = self.state, upd = {classInput:v};
        var cc = st.codeInput && st.names[st.codeInput] ? self.classFor(st.codeInput,'9999-12-31') : null;
        if (cc && v && cc.indexOf(v) !== 0){ upd.nameInput = ''; upd.codeInput = ''; }
        self.setState(upd);
      }),
      onName: (self._onName = function(e){
        var v = e.target.value, st = self.state, upd = {nameInput:v};
        var cl = (st.classInput||'').trim();
        var hits = Object.keys(st.names).filter(function(c){ var k = self.classFor(c,'9999-12-31')||''; return st.names[c]===v && (!cl || k.indexOf(cl)===0); });
        if (hits.length === 1){ upd.codeInput = hits[0]; if (!cl) upd.classInput = self.classFor(hits[0],'9999-12-31')||''; }
        else if (st.codeInput && st.names[st.codeInput]) upd.codeInput = '';
        self.setState(upd);
      }),
      onCode: function(e){
        var v = e.target.value.trim(), st = self.state, upd = {codeInput:v};
        if (st.names[v]){ upd.nameInput = st.names[v]; upd.classInput = self.classFor(v,'9999-12-31')||''; }
        self.setState(upd);
      },
      onTempName: function(e){ self.setState({tempName:e.target.value}); },
      payChips: payChips, payIsTransfer: s.payMode==='transfer', transferTime: s.transferTime, transferShow: s.transferTime || 'เวลา', trBtnCls: 'trbtn'+(s.transferTime ? '' : ' empty'), trSheetOpen: !!s.trSheet, nowLabel: 'ตอนนี้', nowTime: nowHM(),
      onSlipFile: function(e){ var f = e && e.target && e.target.files && e.target.files[0]; var d = new Date(Date.now() - 2*60000), tm = ('0'+d.getHours()).slice(-2)+':'+('0'+d.getMinutes()).slice(-2);
        var ct = self.state.cartU || {}, amt = Object.keys(ct).reduce(function(a,id){ return a + ct[id]*(self.state.prices[id]||0); }, 0);
        self.setState({trSlip:{name: f ? f.name : 'slip.jpg', time: tm, amt: amt}, transferTime: tm}); if (e && e.target) e.target.value = ''; },
      slip: (function(){ var sl = s.trSlip, ct = s.cartU || {}, tot = Object.keys(ct).reduce(function(a,id){ return a + ct[id]*(s.prices[id]||0); }, 0);
        if (!sl) return {has:false, cls:'slipcard', title:'ยังไม่มีสลิป', note:'หรือพิมพ์เวลาเองด้านล่าง', pcCls:'slipic', clear:function(){}};
        var ok = sl.amt === tot;
        return {has:true, cls:'slipcard on'+(ok ? '' : ' warn'), title:'อ่านจากสลิป: เวลา '+sl.time+' น., ยอด '+baht(sl.amt), note: ok ? 'ยอดตรงกับรายการขาย, แนบสลิปไว้กับการขายนี้' : 'ยอดในสลิปไม่ตรงกับยอดขาย '+baht(tot)+' ตรวจอีกครั้ง',
          pcCls:'slipic on', clear:function(){ self.setState({trSlip:null}); }}; })(),
      sug: (function(){
        var empty = function(txt){ return {has:false, none:true, tiles:[], emptyText:txt, closePick:function(){}}; };
        if (shop !== 'u') return empty('');
        if (!code) return empty('ใส่รหัสนักเรียน แล้วจะแนะนำขนาดจากที่เคยซื้อ');
        var last = {}, ord = [];
        s.sales.filter(function(x){ return x.code === code && !x.voided && x.shop === 'u'; }).sort(function(a,b){ return (b.date+b.time).localeCompare(a.date+a.time); })
          .forEach(function(x){ x.items.forEach(function(i){ var p = PMAP[i.pid]; if (!p || last[p.name]) return; last[p.name] = {pid:i.pid, date:x.date}; ord.push(p.name); }); });
        if (!ord.length) return empty('ยังไม่เคยซื้อ เลือกจากชั้นวางด้านล่างได้เลย');
        var fmtSz = function(pid){ return PMAP[pid].size || PMAP[pid].name; };
        var tiles = ord.slice(0, 3).map(function(nm){
          var L = last[nm], sizes = PRODUCTS.filter(function(q){ return q.shop === 'u' && q.name === nm; }).map(function(q){ return q.id; });
          var li = sizes.indexOf(L.pid), up = dayGap(L.date, TODAY) >= 180 && li >= 0 && li < sizes.length - 1;
          var sugPid = up ? sizes[li + 1] : L.pid, key = code+'|'+nm, sel = (s.sugSel && s.sugSel[key]) || sugPid;
          var st = s.stock[sel] || 0, ok = st > 0 && s.prices[sel] > 0, low = ok && isLow(sel), open = s.sugOpen === key;
          var lo = Math.max(0, Math.min(li, sizes.indexOf(sugPid)) - 1), hi = Math.min(sizes.length - 1, Math.max(li, sizes.indexOf(sugPid)) + 1);
          while (hi - lo < 3 && (lo > 0 || hi < sizes.length - 1)){ if (hi < sizes.length - 1) hi++; else lo--; }
          var setSel = function(pid){ var o = Object.assign({}, self.state.sugSel); o[key] = pid; self.setState({sugSel:o, sugOpen:null}); };
          var differ = L.pid !== sugPid;
          return Object.assign(pInfo(sel), {pname:nm, size: fmtSz(sel), isUp: sel === sugPid && differ,
            zbCls:'zb btn'+(sel === sugPid && differ ? ' up' : '')+(open ? ' on' : ''), stk: ok ? 'เหลือ '+st : 'หมด', skCls:'sk'+(ok ? (low ? ' low' : '') : ' out'),
            cls:'sg2'+(ok ? '' : ' off'), adCls:'ad'+(ok ? '' : ' dis'),
            add: function(){ if (!ok){ self.flash(itemLabel(sel)+' หมด'); return; } self.quickAdd(sel); },
            jump: function(){ self.setState({uSel:nm, uSheet:true, jumpHi:sel}); setTimeout(function(){ var e = document.querySelector('.st.jhi'); if (e && window.innerWidth > 760) e.scrollIntoView({behavior:'smooth', block:'center'}); }, 80); clearTimeout(self._jt); self._jt = setTimeout(function(){ self.setState({jumpHi:null}); }, 6000); },
            hasUse: differ, noUse: !differ, why: 'ครั้งก่อน '+fmtSz(L.pid), lastDate: thD2(L.date),
            useLabel: sel === L.pid ? 'ใช้ ↑'+fmtSz(sugPid)+' ที่แนะนำ' : 'ใช้ '+fmtSz(L.pid)+' เหมือนครั้งก่อน',
            use: function(){ setSel(sel === L.pid ? sugPid : L.pid); },
            open: open, openPick: function(){ self.setState({sugOpen: self.state.sugOpen === key ? null : key}); },
            opts: sizes.slice(lo, hi + 1).map(function(pid){ var so = s.stock[pid] || 0, avail = so > 0 && s.prices[pid] > 0;
              return {size: fmtSz(pid), stk: avail ? 'เหลือ '+so : 'หมด', stCls: avail && isLow(pid) ? 'low' : '', tag: pid === L.pid ? 'ครั้งก่อน' : (pid === sugPid ? '↑ แนะนำ' : ''), tgCls:'tg'+(pid === L.pid ? ' last' : (pid === sugPid ? ' rec' : '')),
                cls:'zc'+(pid === sel ? ' on' : '')+(avail ? '' : ' out'), pick: function(){ if (!avail){ self.flash(itemLabel(pid)+' หมด'); return; } setSel(pid); }}; })}); });
        return {has:true, none:false, tiles:tiles, emptyText:'', closePick: function(){ self.setState({sugOpen:null}); }};
      })(),
      tally: tally,
      rcGroups: rcGroups, rcModes: rcModes, rcShop: rcShop, rcShopBox: 'sbox'+(rcShop==='all' ? '' : ' '+rcShop), onRcShop: function(e){ self.setState({rcShop:e.target.value}); }, rcv: rcv, dup: dup,
      rcvAskVoid: function(){ if (self.state.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; } self.setState({rcvVoidAsk:true}); },
      rcvCancelVoid: function(){ self.setState({rcvVoidAsk:false}); },
      rcvConfirmVoid: function(){if (self.state.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; } 
        var st = self.state, id = st.rcvId, b = st.batches[id], sk = Object.assign({}, st.stock), bt = Object.assign({}, st.batches), q = 0, by = DATA.myName, at = thDate(TODAY)+' '+nowHM();
        st.receipts.forEach(function(r){ if (r.batch === id){ sk[r.pid] = (sk[r.pid]||0) - r.qty; q += r.qty; } });
        bt[id] = Object.assign({}, b, {voided:{at:at, by:by}, log:(b.log||[]).concat([{at:at, by:by, text:'ยกเลิกการรับเข้า: หักสต็อกออก '+q+' ชิ้น'}])});
        self.setState({stock:sk, batches:bt, rcvVoidAsk:false}); self.flash('ยกเลิกการรับเข้า '+batchTitle(b)+' แล้ว (หักสต็อก '+q+' ชิ้น)'); },
      rcvEdit: function(){ if (self.state.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; } var st = self.state, b = st.batches[st.rcvId];
        self.setState({rcvId:null, rcvVoidAsk:false, imp2Open:true, imp2Step:0, replaceId:b.id, memoLoaded:false, poLoaded:false, impMemoDoc:null, impPoDoc:null, impLines:null, dupAlert:null, rDate:b.date}); },
      dupDrop: function(){ var a = self.state.dupAlert, p = {dupAlert:null, impLines:null}; if (a.kind==='memo'){ p.memoLoaded = false; p.impMemoDoc = null; } else { p.poLoaded = false; p.impPoDoc = null; } self.setState(p); },
      dupKeep: function(){ self.setState({dupAlert:null}); },
      dupView: function(){ var a = self.state.dupAlert;
        self.setState({dupAlert:null, imp2Open:false, imp2Step:0, memoLoaded:false, poLoaded:false, impMemoDoc:null, impPoDoc:null, impLines:null, replaceId:null, rcvId:a.id, rcvDoc:'memo', rcvPid:null}); },
      dupReplace: function(){ var a = self.state.dupAlert; self.setState({dupAlert:null, replaceId:a.id}); self.flash('เปลี่ยนเป็นโหมดแก้ไขล็อต '+batchTitle(self.state.batches[a.id])); },
      dupAttach: function(){if (self.state.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; } 
        var st = self.state, a = st.dupAlert, bt = Object.assign({}, st.batches), b = Object.assign({}, bt[a.id]), by = DATA.myName, at = thDate(TODAY)+' '+nowHM();
        b[a.kind] = a.doc; b.log = (b.log||[]).concat([{at:at, by:by, text:'แนบ '+(a.kind==='po' ? '' : 'บันทึกข้อความ ')+a.doc.no+' เข้ากับล็อตนี้ (สต็อกไม่เปลี่ยน'+(a.kind==='po' ? ', เติมราคาทุน' : '')+')'}]); bt[a.id] = b;
        var memoTxt = [b.memo && b.memo.no, b.po && b.po.no].filter(Boolean).join(', ');
        var rc = st.receipts.map(function(r){ if (r.batch !== a.id) return r; var l = (a.doc.lines||[]).filter(function(x){ return x.pid === r.pid; })[0];
          return Object.assign({}, r, {memo:memoTxt, cost:(a.kind==='po' && l && l.cost) ? l.cost : r.cost}); });
        self.setState({batches:bt, receipts:rc, dupAlert:null, imp2Open:false, imp2Step:0, memoLoaded:false, poLoaded:false, impMemoDoc:null, impPoDoc:null, impLines:null, replaceId:null, rcvId:a.id, rcvDoc:a.kind, rcvPid:null});
        self.flash('แนบ '+a.doc.no+' เข้ากับ '+batchTitle(b)+' แล้ว'); },
      pickDupPo: function(){ self.pickDoc('po', DEMO_PO_DUP); },
      pickOldMemo: function(){ self.pickDoc('memo', DEMO_MEMO_OLD); },
      closeRcv: function(){ self.setState({rcvId:null, rcvPid:null, rcvVoidAsk:false}); },
      openDocFile: function(){ self.flash('เวอร์ชันจริงจะเปิดไฟล์ PDF ต้นฉบับในแท็บใหม่'); },
      dlDocFile: function(){ self.flash('เวอร์ชันจริงจะดาวน์โหลดไฟล์ต้นฉบับ'); }, rcEmpty: rcMode === 'item' ? rcItems.length === 0 : !(rcGroups.length && rcGroups[0].cards.length), rcItems: rcItems, rcIsItem: rcMode === 'item', rcNotItem: rcMode !== 'item', rcQ: s.rcQ||'', onRcQ: function(e){ self.setState({rcQ:e.target.value}); },
      bk: bk, bellAlerts: bellAlerts, help: help, B: {sCash:bn(money.sCash), uCash:bn(money.uCash), uTr:bn(money.uTr), uTot:bn(money.uTot), sTot:bn(money.sTot), dTotal:bn(dash.total), tTotal:bn(tally.total), sTotal:bn(sell.total), ePrice:bn(edit.price), eCost:bn(edit.cost), saleTotal:bn(sale.total), clsExp:bn(cls.exp)}, cls: cls, mvp: mvp, mg: mg, ne: npe, canEdit: isOwner, viewOnly: !isOwner, stockCols: stockCols, mergedRows: mergedRows, hasMerged: mergedRows.length > 0, mergedCount: mergedRows.length+' รายการ', stockSum: stockSum, stockTabs: stockTabs, edit: edit, imp2: imp2,
      editPrice: s.editPrice, editQty: s.editQty, editMemo: s.editMemo, editDate: s.editDate,
      closeEdit: function(){ self.setState({editId:null}); },
      onEditPrice: function(e){ self.setState({editPrice:e.target.value, editPriceAck:null}); },
      onEditQty: function(e){ self.setState({editQty:e.target.value, editQtyAck:null}); },
      onEditMemo: function(e){ self.setState({editMemo:e.target.value}); },
      onEditDate: function(e){ self.setState({editDate:e.target.value}); },
      savePrice: function(){ if (self.state.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; } var st = self.state, v = Number(st.editPrice); if (!(v > 0)){ self.flash('ใส่ราคามากกว่า 0'); return; } var old = st.prices[st.editId]||0; if (old > 0 && (v >= old*1.5 || v <= old*0.5) && st.editPriceAck !== v){ self.setState({editPriceAck:v}); self.flash('ราคาต่างจากเดิมมาก ถ้าถูกต้องให้กดยืนยันอีกครั้ง'); return; } var pr = Object.assign({}, st.prices); pr[st.editId] = v; self.setState({prices:pr, editPrice:'', editPriceAck:null}); self.flash('ตั้งราคา '+itemLabel(st.editId)+' เป็น '+baht(v)+' แล้ว'); },
      saveReceive: function(){ if (self.state.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; } var st = self.state, q = parseInt(st.editQty,10); if (!q || q < 1){ self.flash('กรอกจำนวนรับเข้า'); return; } var qmax = st.receipts.filter(function(r){ return r.pid === st.editId; }).reduce(function(a,r){ return Math.max(a, r.qty||0); }, 0); if (q > Math.max(100, qmax*3) && st.editQtyAck !== q){ self.setState({editQtyAck:q}); self.flash('จำนวนมากผิดปกติ ถ้าถูกต้องให้กดยืนยันอีกครั้ง'); return; }
        var sk = Object.assign({}, st.stock); sk[st.editId] = (sk[st.editId]||0) + q;
        var bid = 'm'+Date.now(), bt = Object.assign({}, st.batches), mno = (st.editMemo||'').trim();
        bt[bid] = {id:bid, date:st.editDate||TODAY, at:nowHM(), by:DATA.myName, src:'manual', memoNo:mno, memo:null, po:null, dateFrom:'กรอกเอง'};
        self.setState({stock:sk, editQty:'', editQtyAck:null, batches:bt, receipts:[{date:st.editDate||TODAY, pid:st.editId, label:itemLabel(st.editId), qty:q, memo:mno, cost:null, batch:bid}].concat(st.receipts)});
        self.flash('รับเข้า '+itemLabel(st.editId)+' +'+q+' แล้ว'); },
      openImport: function(){ if (self.state.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; } var st = self.state; self.setState(st.replaceId ? {imp2Open:true, imp2Step:0, memoLoaded:false, poLoaded:false, impLines:null, rDate:TODAY, replaceId:null, impMemoDoc:null, impPoDoc:null, dupAlert:null} : {imp2Open:true}); },
      closeImport: function(){ var st = self.state; self.setState(st.imp2Step===2 || st.replaceId ? {imp2Open:false, imp2Step:0, memoLoaded:false, poLoaded:false, impLines:null, rDate:TODAY, replaceId:null, impMemoDoc:null, impPoDoc:null, dupAlert:null} : {imp2Open:false}); },
      pickMemo: function(){ self.pickDoc('memo', IMP_MEMO); },
      pickPo: function(){ self.pickDoc('po', IMP_PO); },
      toReview: function(){ var st = self.state; if (!st.memoLoaded && !st.poLoaded){ self.flash('เลือกบันทึกข้อความหรือ PO อย่างน้อยหนึ่งไฟล์'); return; }
        var src = (st.poLoaded && st.impPoDoc && st.impPoDoc.lines) || (st.memoLoaded && st.impMemoDoc && st.impMemoDoc.lines) || IMP_LINES;
        self.setState({imp2Step:1, impLines: st.impLines || src.map(function(l){ return Object.assign({}, l, {inc:true, price:''}); })}); },
      backUpload: function(){ self.setState({imp2Step:0}); },
      onRDate: function(e){ self.setState({rDate:e.target.value}); },
      confirmReceive: function(){if (self.state.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; } 
        var st = self.state, d = st.rDate || TODAY, sk = Object.assign({}, st.stock), pr = Object.assign({}, st.prices), rc = [], n = 0, made = 0, q0 = 0;
        var CM = st.impMemoDoc || IMP_MEMO, CP = st.impPoDoc || IMP_PO, rpId = st.replaceId, old = rpId ? st.batches[rpId] : null;
        var bid = rpId || ('b'+Date.now()), bt = Object.assign({}, st.batches), by = DATA.myName;
        var memoDoc = st.memoLoaded ? CM : (old ? old.memo : null), poDoc = st.poLoaded ? CP : (old ? old.po : null);
        var memoTxt = [memoDoc && memoDoc.no, poDoc && poDoc.no].filter(Boolean).join(', ');
        var base = st.receipts, oq = 0;
        if (old){ base = st.receipts.filter(function(r){ return r.batch !== rpId; }); st.receipts.forEach(function(r){ if (r.batch === rpId){ sk[r.pid] = (sk[r.pid]||0) - r.qty; oq += r.qty; } }); }
        (st.impLines||[]).forEach(function(l){
          var q = Number(l.qty)||0; if (!l.inc || q <= 0) return;
          if (!PMAP[l.pid]){ var b = PMAP[l.base]; var np = Object.assign({}, b, {id:l.pid, size:l.size, price:0, stock:0});
            var at = 0; ALLP.forEach(function(p, i){ if (p.name===b.name) at = i; }); ALLP.splice(at+1, 0, np); PMAP[np.id] = np; sk[np.id] = 0; pr[np.id] = 0; made++; }
          sk[l.pid] = (sk[l.pid]||0) + q; if (Number(l.price) > 0) pr[l.pid] = Number(l.price);
          rc.push({date:d, pid:l.pid, label:itemLabel(l.pid), qty:q, memo:memoTxt, cost: poDoc ? (l.cost||null) : null, batch:bid}); n++; q0 += q;
        });
        var stamp = thDate(TODAY)+' '+nowHM();
        var log = old ? (old.log||[]).concat([{at:stamp, by:by, text:'แทนที่ '+[old.memo && old.memo.file, old.po && old.po.file].filter(Boolean).join(', ')+' ด้วย '+[st.memoLoaded && CM.file, st.poLoaded && CP.file].filter(Boolean).join(', ')+': สต็อก '+(q0-oq >= 0 ? '+' : '')+(q0-oq)+' ชิ้น'}]) : [];
        bt[bid] = {id:bid, date:d, at: old ? old.at : nowHM(), by: old ? old.by : by, src:'import', log:log,
          dateFrom: old && d===old.date ? old.dateFrom : memoDoc && d===memoDoc.date ? 'วันที่ในบันทึกข้อความ' : (poDoc && d===poDoc.date ? 'วันที่ใน PO' : (d===TODAY ? 'วันที่อัปโหลด' : 'กำหนดเอง')),
          memo: memoDoc, po: poDoc};
        var left = (st.impLines||[]).filter(function(l){ return l.inc && !(pr[l.pid] > 0); }).length;
        self.setState({stock:sk, prices:pr, batches:bt, receipts: rc.concat(base), imp2Step:2,
          impDone: (old ? 'แก้ไขล็อต '+memoTxt+' แล้ว: เดิม +'+oq+' ชิ้น → ใหม่ +'+q0+' ชิ้น (สต็อกปรับตามส่วนต่าง และบันทึกไว้ในประวัติการแก้ไขของล็อต)'
            : '+'+q0+' ชิ้น ใน '+n+' รายการ, ลงวันที่ '+thDate(d)+', อ้างอิง '+memoTxt)+(made ? ', สร้างไซส์ใหม่ '+made : '')+(left ? ', ยังต้องตั้งราคา '+left+' ไซส์ ดูได้จากการ์ด "ยังไม่ตั้งราคา"' : '')});
      },
      resetImport2: function(){ self.setState({imp2Step:0, memoLoaded:false, poLoaded:false, impLines:null, rDate:TODAY, replaceId:null, impMemoDoc:null, impPoDoc:null, dupAlert:null}); },
      launch: [['u','sellU','ขายเครื่องแบบ'],['s','sellS','ขายเครื่องเขียน']].map(function(L){
        var td = s.sales.filter(function(x){ return x.date===TODAY && x.shop===L[0]; });
        var low = PRODUCTS.filter(function(p){ return p.shop===L[0] && s.prices[p.id] > 0 && s.stock[p.id] <= 5; }).length;
        var cartN = Object.keys(s[L[0]==='u'?'cartU':'cartS']||{}).length;
        var note = cartN ? 'มีสินค้าค้างในตะกร้า '+cartN+' รายการ' : (low ? 'ใกล้หมด '+low+' รายการ' : 'สต็อกปกติ');
        return {cls:'lcard cnr '+L[0], isU:L[0]==='u', isS:L[0]==='s', title:L[2], today:baht(sum(td)), count:td.length, note:note, go:function(){ self.go(L[1]); }};
      }),
      stuChip: (function(){
        var hasCart = Object.keys(s.cartU||{}).some(function(k){ return s.cartU[k] > 0; });
        var goUp = function(){ var el = document.getElementById('pclass'); if (el){ el.scrollIntoView({behavior:'smooth', block:'center'}); setTimeout(function(){ try{ el.focus({preventScroll:true}); }catch(e){} }, 350); } };
        var c1 = lookup.found ? lookup.cls : ((s.classInput||'').trim() || 'ชั้น');
        var nm = (lookup.found || lookup.inactive) ? lookup.name : ((s.nameInput||'').trim() || 'ชื่อนักเรียน');
        var cd = code || 'รหัส';
        var st, note;
        if (code && lookup.found){ st = 'ok'; note = 'พบในรายชื่อ'; }
        else if (code){ st = 'warn'; note = lookup.inactive ? 'ไม่อยู่ในรายชื่อ' : 'ยังไม่มีในรายชื่อ'; }
        else if (hasCart){ st = 'need'; note = 'ยังไม่ได้เลือกนักเรียน'; }
        else { st = 'idle'; note = 'แตะเพื่อเลือกนักเรียน'; }
        return {show:!!s.stuCompact, cls:'stuchip '+st, cls1:c1, name:nm, code:cd, note:note, title:'แตะเพื่อเลือกหรือแก้ไขนักเรียน', go:goUp};
      })(),
      shopTabs: [['sellU','เครื่องแบบ','cartU'],['sellS','เครื่องเขียน','cartS']].map(function(t){
        var on = s.page===t[0], c = s[t[2]]||{}, n = Object.keys(c).reduce(function(a,k){ return a+(c[k]||0); },0);
        return {label:t[1], cls:'stab '+(t[0]==='sellU'?'u':'s')+(on?' on':''), sel: on?'true':'false', isU: t[0]==='sellU', isS: t[0]==='sellS', cart:n, hasCart: n>0 && !on, go:function(){ self.go(t[0]); }};
      }),
      activePills: activePills, activeCount: activePills.length, hasActive: activePills.length > 0, noActive: activePills.length === 0, visibleCount: visibleCount,
      filterOpen: !!s.filterOpen, filterExpanded: s.filterOpen ? 'true' : 'false',
      fbtnCls: 'fbtn'+(activePills.length ? ' act' : ''),
      fbarCls: 'fbar'+(s.filterOpen ? ' open' : ''),
      toggleFilter: function(){ self.setState({filterOpen: !self.state.filterOpen}); },
      closeFilter: function(){ self.setState({filterOpen:false}); },
      clearFilters: function(){ self.setState({uFilter:'all', lFilter:'all', pFilter:'all', gFilter:'all'}); },
      payIsCash: s.payMode==='cash', payCashCls: 'segb'+(s.payMode==='cash'?' on':''),
      pickCash: function(){ self.setState({payMode:'cash'}); },
      pickTransfer: function(){ self.setState({payMode:'transfer', trSheet: window.innerWidth <= 760}); },
      openTrSheet: function(){ self.setState({trSheet:true}); },
      trNow: function(){ self.setState({transferTime: nowHM()}); },
      trCancel: function(){ self.setState({trSheet:false, payMode: self.state.transferTime ? self.state.payMode : 'cash'}); },
      trDone: function(){ if (!self.state.transferTime){ self.flash('ใส่เวลาที่โอนตามสลิปก่อน'); return; } self.setState({trSheet:false}); },
      chevCls: 'chev'+(s.dockOpen?' up':''),
      decoCls: 'deco '+(s.page==='sellS'?'s':'u')+(s.sideMini?' mini':''),
      dockCls: 'dock'+(s.sideMini?' mini':''), dockSpaceCls: 'dockspace'+(s.dockOpen?' open':''),
      dockGripCls: 'dgrip'+(s.dockDrag ? ' on' : ''),
      dockRowsStyle: (function(){ var h = s.dockDrag ? s.dockDrag.h : s.dockH; return h && window.innerWidth > 760 ? 'height: '+h+'px' : ''; })(),
      dockSpaceStyle: (function(){ var h = s.dockH; return s.dockOpen && h && window.innerWidth > 760 ? 'height: '+(300 + h - 96)+'px' : ''; })(),
      dockGrip: function(e){ if (window.innerWidth <= 760 || (e.button != null && e.button !== 0)) return; e.preventDefault();
        var y0 = e.clientY, h0 = self.state.dockH || 96, maxH = function(){ return Math.max(96, Math.floor((window.innerHeight * 0.62 - 210) / 48) * 48); };
        var mv = function(ev){ var h = Math.max(96, Math.min(maxH(), h0 + (y0 - ev.clientY))); self.setState({dockDrag:{h:h}}); };
        var up = function(){ window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up); document.documentElement.removeAttribute('data-dragging');
          var d = self.state.dockDrag, h = d ? Math.max(96, Math.min(maxH(), Math.round(d.h / 48) * 48)) : h0;
          try { localStorage.setItem('dockH', String(h)); } catch(err){}
          self.setState({dockH:h, dockDrag:null}); };
        document.documentElement.setAttribute('data-dragging', '');
        window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up); },
      dockGripReset: function(){ try { localStorage.setItem('dockH', '96'); } catch(err){} self.setState({dockH:96, dockDrag:null}); },
      sideCls: 'side'+(s.sideMini?' mini':'')+(s.drawerOpen?' open':''),
      drawerOpen: !!s.drawerOpen,
      openDrawer: function(){ self.setState({drawerOpen:true}); },
      closeDrawer: function(){ self.setState({drawerOpen:false}); }, sideMini: s.sideMini, sideFull: !s.sideMini,
      sideTogLabel: s.sideMini ? 'ขยายเมนู' : 'ย่อเมนูเหลือไอคอน',
      toggleSide: function(){
        var phone = window.matchMedia && window.matchMedia('(max-width: 760px)').matches;
        if (phone) self.setState({drawerOpen:false}); else self.setState({sideMini: !self.state.sideMini});
      },
      dockOpen: s.dockOpen, dockExpanded: s.dockOpen ? 'true' : 'false',
      dockLabel: s.dockOpen ? 'ซ่อนรายละเอียด' : 'ดูรายละเอียด',
      toggleDock: function(){ self.setState({dockOpen: !self.state.dockOpen}); },
      onTTime: function(e){ self.setState({transferTime:e.target.value}); },
      hist: hist, sale: sale, want: want, waitChip: waitChip,
      waitSum: (function(){ var all = [].concat(fReady, fHeld, fWaiting), stu = {}, rows = [];
        all.forEach(function(d){ stu[d.code || d.name] = 1; });
        [[fReady,'แจ้งรับของ','ok'],[fHeld,'แจ้งแล้ว รอรับของ','held'],[fWaiting,'รอของ','no']].forEach(function(gp, gi){
          var m = {}, order = [];
          gp[0].forEach(function(d){ if (!m[d.pid]){ m[d.pid] = {pid:d.pid, n:0, who:{}}; order.push(m[d.pid]); } m[d.pid].n++; m[d.pid].who[d.code || d.name] = 1; });
          order.sort(function(a,b){ return ((PMAP[a.pid].shop === 's') - (PMAP[b.pid].shop === 's')) || b.n - a.n || itemLabel(a.pid).localeCompare(itemLabel(b.pid),'th'); });
          rows.push({isHead:true, isRow:false, isEmpty:false, head:gp[1], headN:String(gp[0].length), dot:'wdot '+gp[2], cls:'wshead'+(gi ? ' sep' : '')});
          if (!order.length) rows.push({isHead:false, isRow:false, isEmpty:true, cls:'wsempty'});
          order.forEach(function(o){ var u = PMAP[o.pid].shop === 'u'; rows.push({isHead:false, isRow:true, isEmpty:false, label:itemLabel(o.pid), pcs:o.n+' ชิ้น', ppl:Object.keys(o.who).length+' คน', icoCls:'wsico '+(u ? 'u' : 's'), isU:u, isS:!u, cls:'wsrow'}); });
        });
        return {sumTotal: 'รวม '+all.length+' ชิ้น จากนักเรียน '+Object.keys(stu).length+' คน', sumRows: rows, hasSum: true};
      })(),
      wait: {open: !!s.waitOpen,
        readyGroups: byClass(fReady), held: fHeld.map(waitRow), waiting: fWaiting.map(waitRow),
        waitingGroups: (function(){ var g = [], m = {}; fWaiting.forEach(function(d){ if (!m[d.pid]){ m[d.pid] = Object.assign(pInfo(d.pid), {item:itemLabel(d.pid), rows:[], cls:'wgcard cnr '+PMAP[d.pid].shop}); g.push(m[d.pid]); } m[d.pid].rows.push(waitRow(d)); });
          g.forEach(function(x){ x.count = x.rows.length+' คน'; });
          return g.sort(function(a,b){ return (a.cls.slice(-1) === 's') - (b.cls.slice(-1) === 's') || b.rows.length - a.rows.length; }); })(),
        hasReady: fReady.length > 0, noReady: fReady.length === 0, hasHeld: fHeld.length > 0, hasWaiting: fWaiting.length > 0, noHeld: fHeld.length === 0, noWaiting: fWaiting.length === 0,
        readyN: String(fReady.length), heldN: String(fHeld.length), waitingN: String(fWaiting.length),
        readyU: splitB(fReady,'u'), readyS: splitB(fReady,'s'), heldU: splitB(fHeld,'u'), heldS: splitB(fHeld,'s'), waitingU: splitB(fWaiting,'u'), waitingS: splitB(fWaiting,'s'),
        q: s.waitQ || '', shop: wshop, shopBox: 'sbox'+(wshop === 'all' ? '' : ' '+wshop), filtered: !!(wq || wshop !== 'all'), noMatch: !!(wq || wshop !== 'all') && !(fReady.length + fHeld.length + fWaiting.length), readyTitle: 'ของมาแล้ว, รอแจ้งนักเรียน ('+readyList.length+')', heldTitle: 'แจ้งแล้ว, รอรับของ ('+heldList.length+')', waitingTitle: 'ยังรอของ ('+waitingList.length+')',
        canAct: isOwner},
      waitBtn: 'นักเรียนรอของ', waitBtnN: String(readyList.length + heldList.length + waitingList.length), waitBtnCls: 'btn ghost waitbtn'+(readyList.length ? ' hot' : ''), waitTileCls: readyList.length ? 'hot' : '',
      wx: (function(){
        var scope = s.wxScope || 'all', sort = s.wxSort || 'item', fmt = s.wxFmt || 'pdf';
        var all = [].concat(readyList.map(function(d){ return [d,'r']; }), heldList.map(function(d){ return [d,'h']; }), waitingList.map(function(d){ return [d,'w']; }));
        var list = all.filter(function(x){ return scope === 'all' || (scope === 'wait' ? x[1] === 'w' : (scope === 'ready' ? x[1] === 'r' : x[1] !== 'w')); });
        var rows = list.map(function(x){ var d = x[0], w = whoOf(d), p0 = PMAP[d.pid];
          return {who:w.who, cls:w.cls || '—', item:itemLabel(d.pid), date:thD2(d.date), asks: asksOf(d).length > 1 ? 'ถาม '+asksOf(d).length+' ครั้ง' : '', _d:d.date+(d.time||''), _n:p0.name, _z:p0.size||'', _s: p0.shop === 's' ? 1 : 0, _c:w.cls || 'ฮ',
            got: d.got ? 'ได้ '+(PMAP[d.got].size ? '#'+PMAP[d.got].size : PMAP[d.got].name)+' ไปก่อน' : '',
            status: x[1] === 'w' ? 'รอของ' : (x[1] === 'r' ? 'ของมาแล้ว, รอแจ้งนักเรียน' : 'แจ้งนักเรียนแล้ว, จองของไว้แล้ว'), stCls: 'wxst '+x[1]}; });
        rows.sort(function(a,b){ return (a._s - b._s) || (sort === 'cls' ? (a._c.localeCompare(b._c,'th') || a.who.localeCompare(b.who,'th')) : (sort === 'date' ? a._d.localeCompare(b._d) : (a._n.localeCompare(b._n,'th') || a._z.localeCompare(b._z,undefined,{numeric:true}) || a._d.localeCompare(b._d)))); });
        var seg = function(list0, cur, key){ return list0.map(function(o){ return {label:o[1], cls:'segb'+(cur===o[0] ? ' on' : ''), pick:function(){ var u = {}; u[key] = o[0]; self.setState(u); }}; }); };
        var anon = (s.demand||[]).filter(function(d){ return live(d) && !withWho(d) && dmdOpen(d); }).length;
        var ttl = scope === 'wait' ? 'รอของ' : (scope === 'arrived' ? 'ของมาแล้ว' : (scope === 'ready' ? 'ของมาแล้ว ยังไม่แจ้งนักเรียน' : 'ทั้งหมด'));
        return {open: !!s.wxOpen, scopeTabs: seg([['all','ทั้งหมด'],['ready','แจ้งรับของ'],['wait','รอของ']], scope, 'wxScope').map(function(k, i){ k.dots = [[],['wdot ok'],['wdot no']][i]; return k; }),
          sortTabs: segIco(seg([['item','ตามสินค้า'],['cls','ตามชั้น'],['date','ตามวันที่ขอ']], sort, 'wxSort'), ['item','cls','date']), fmtTabs: tileOf(segIco(seg([['text','ข้อความ'],['pdf','PDF'],['jpg','JPG']], fmt, 'wxFmt'), ['text','pdf','jpg'])), isText: fmt === 'text', isFile: fmt !== 'text', goLabel: fmt === 'text' ? 'คัดลอก' : 'ดาวน์โหลด',
          text: (function(){ var out = [], cur = null; rows.forEach(function(r){ var head = sort === 'cls' ? r.cls : (sort === 'item' ? r._n : thD2(r._d.slice(0,10))); if (head !== cur){ if (out.length) out.push(''); out.push(head); cur = head; } out.push('- '+r.who+(sort === 'cls' || r.cls === '—' ? '' : ' ('+r.cls+')')+': '+r.item+(scope === 'ready' ? ' มาแล้ว รับได้ที่'+PICKUP : '')); }); return out.join('\n') || 'ไม่มีรายชื่อ'; })(),
          rows: rows, empty: rows.length === 0, has: rows.length > 0, title: 'รายชื่อนักเรียนที่รอของ: '+ttl, date: 'ณ วันที่ '+thDate(TODAY)+' เวลา '+nowHM()+' น.',
          total: 'รวม '+rows.length+' รายการ, '+Object.keys(rows.reduce(function(m,r){ m[r.who] = 1; return m; },{})).length+' คน', anon: anon ? 'ไม่รวมคนที่แค่ถามโดยไม่ได้ให้ชื่อ '+anon+' ครั้ง' : '',
          file: 'นักเรียนรอของ '+ttl+' '+thD2(TODAY)+(fmt === 'pdf' ? '.pdf' : '.jpg')};
      })(),
      ah: (function(){ var d = (s.demand||[]).filter(function(x){ return x.id === s.askHist; })[0]; if (!d) return {open:false, rows:[]};
        var w = whoOf(d), h = asksOf(d).slice().sort(function(a,b){ return (a.date+a.time).localeCompare(b.date+b.time); });
        var first = h[0].date, last = h[h.length-1].date, ev = [];
        h.forEach(function(a, k){ ev.push({k:a.date+' '+a.time, kind:'a', a:a, n:k+1}); });
        var rcSeen = {}; s.receipts.forEach(function(r){ var b = r.batch && s.batches[r.batch]; if (b && b.voided) return; if (r.pid !== d.pid || r.date < first) return; var key = r.date+(r.batch||''); if (rcSeen[key]) return; rcSeen[key] = 1; ev.push({k:r.date+' '+((b && b.at) || '00:00'), kind:'r', r:r}); });
        ev.sort(function(a,b){ return a.k.localeCompare(b.k); });
        var prev = null, rows = ev.map(function(e){
          if (e.kind === 'r'){ var rk = e.k, rs = s.stock[d.pid]||0;
            s.receipts.forEach(function(r){ var b = r.batch && s.batches[r.batch]; if (b && b.voided) return; if (r.pid === d.pid && (r.date+' '+((b && b.at) || '00:00')) > rk) rs -= r.qty; });
            s.sales.forEach(function(x){ if (x.voided || (x.date+' '+x.time) <= rk) return; x.items.forEach(function(i){ if (i.pid === d.pid) rs += i.qty; }); });
            rs = Math.max(0, rs);
            return {stockTxt:'', stockCls:'', cls:'ahrow r', title:'ของเข้า '+e.r.qty+' ชิ้น', when:thD2(e.r.date), tag:'มีในสต็อก '+rs, tagCls:'dtag '+(rs > 5 ? 'ok' : (rs > 0 ? 'amber' : 'red')), gap:'', hasGap:false}; }
          var g = prev ? dayGap(prev, e.a.date) : null; prev = e.a.date;
          var key = e.a.date+' '+(e.a.time||'23:59'), sk = s.stock[d.pid]||0;
          s.receipts.forEach(function(r){ var b = r.batch && s.batches[r.batch]; if (b && b.voided) return; if (r.pid === d.pid && (r.date+' '+((b && b.at) || '00:00')) > key) sk -= r.qty; });
          s.sales.forEach(function(x){ if (x.voided || (x.date+' '+x.time) <= key) return; x.items.forEach(function(i){ if (i.pid === d.pid) sk += i.qty; }); });
          sk = Math.max(0, sk);
          return {stockTxt:'สต็อกตอนนั้น '+sk, stockCls:'dtag '+(sk > 5 ? 'ok' : (sk > 0 ? 'amber' : 'red')), cls:'ahrow', title:'ครั้งที่ '+e.n, when:thD2(e.a.date)+' '+(e.a.time||''), tag: e.a.got ? 'ได้ '+(PMAP[e.a.got].size ? '#'+PMAP[e.a.got].size : PMAP[e.a.got].name)+' ไปก่อน' : 'ไม่ได้ซื้อของ', tagCls: 'dtag grey', gap: g !== null ? '+'+g+' วัน' : '', hasGap: g !== null}; });
        var span = dayGap(first, last), avg = h.length > 1 ? Math.round(span/(h.length-1)) : 0, waited = dayGap(first, TODAY);
        return {open:true, title:itemLabel(d.pid), sub:w.who+(w.cls ? ', '+w.cls : ''), rows:rows,
          k1: h.length+' ครั้ง', k2: h.length > 1 ? 'ทุก ~'+avg+' วัน' : '—', k3: waited+' วัน',
          note: dmdOpen(d) ? 'ยังไม่มีของเข้าหลังจากถามครั้งแรก' : 'มีของเข้าแล้ว'+(d.notified ? ', แจ้งนักเรียนแล้ว' : ', ยังไม่ได้แจ้งนักเรียน')};
      })(),
      closeAh: function(){ self.setState({askHist:null}); },
      cpv: (function(){
        var c = s.cpv; if (!c) return {open:false, cells:[]};
        var EN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        var sheet = function(iso){ var p = iso.split('-').map(Number); return EN[p[1]-1]+', '+p[2]; };
        var ds = s.drawerSince || '2026-09-30';
        var list = s.sales.filter(function(x){ return !x.voided && x.shop === c.k && (c.hand ? (x.date > ds && x.date <= TODAY) : x.date === c.date); })
          .sort(function(a,b){ return (a.date+a.time).localeCompare(b.date+b.time); });
        var rows = [], heads, tot = 0;
        if (c.k === 'u'){
          heads = ['รหัส','ชื่อสินค้า','ไซส์','ราคา','จำนวน','สด','โอน','เวลา','อ้างอิงการโอน'];
          list.forEach(function(x){ var rem = x.transfer || 0; x.items.forEach(function(i){ var p = PMAP[i.pid], amt = i.qty*i.price, tr = Math.min(amt, rem); rem -= tr; tot += amt;
            rows.push([x.code || '', p.name, p.size || '', String(i.price), String(i.qty), amt-tr ? String(amt-tr) : '', tr ? String(tr) : '', x.time, tr && x.transferTime ? 'สลิป '+x.transferTime : '']); }); });
        } else {
          heads = ['เลขที่บันทึก','รายการ','ราคา','จำนวน','สด','โอน','หมายเหตุ','หมวด'];
          var M = {}, ord = [];
          list.forEach(function(x){ x.items.forEach(function(i){ var p = PMAP[i.pid], k = p.name+'|'+i.price; if (!M[k]){ M[k] = {name:p.name, price:i.price, q:0, who:[]}; ord.push(k); } M[k].q += i.qty; tot += i.qty*i.price; if (M[k].who.indexOf(x.seller) < 0) M[k].who.push(x.seller); }); });
          ord.forEach(function(k){ var r = M[k]; rows.push(['', r.name, String(r.price), String(r.q), String(r.q*r.price), '', r.who.join(', ')+' ขาย', 'ขายเครื่องเขียน']); });
        }
        var NUM = c.k === 'u' ? [3,4,5,6] : [2,3,4,5];
        var cells = heads.map(function(h, i){ return {v:h, cls:'cbc cbh'+(NUM.indexOf(i) >= 0 ? ' n' : '')}; });
        rows.forEach(function(r){ r.forEach(function(v, i){ cells.push({v:v, cls:'cbc'+(NUM.indexOf(i) >= 0 ? ' n mono' : '')}); }); });
        var day = c.hand ? TODAY : c.date, table = c.k === 'u' ? 'ยอดขายเครื่องแบบนักเรียน' : 'ค่าสมุดอุปกรณ์ / รายรับอื่น ๆ';
        var title = c.hand ? 'นำส่งเงินเครื่องเขียน' : (c.k === 'u' ? 'คัดลอกเครื่องแบบ' : 'คัดลอกเครื่องเขียน')+' '+thD2(c.date);
        var sub = (c.hand ? 'ขายตั้งแต่ '+thD2(addDays(ds, 1))+' ถึง '+thD2(TODAY)+', ' : '')+'วางในตาราง '+table+' แผ่น '+sheet(day)+' ของ Cash Book';
        var empty = rows.length === 0;
        return {open:true, title:title, sub:sub, cells:cells, gridCls:'cbt '+c.k, empty:empty, has:!empty, count: rows.length+' แถว', total: baht(tot), totLabel: c.hand ? 'เงินที่นำส่ง' : 'รวม',
          goLabel: c.hand ? 'คัดลอกและนำส่ง' : 'คัดลอก', goCls: 'btn'+(empty ? ' dis' : ''),
          close: function(){ self.setState({cpv:null}); },
          go: function(){ if (empty){ self.flash('ไม่มีรายการ'); return; } var txt = rows.map(function(r){ return r.join('\t'); }).join('\n'); try { navigator.clipboard.writeText(txt); } catch(e){}
            if (c.hand){ self.setState({cpv:null, drawerSince:TODAY}); self.flash('คัดลอก '+rows.length+' แถวและบันทึกนำส่ง '+baht(tot)+' แล้ว'); }
            else { self.setState({cpv:null}); self.flash('คัดลอก '+rows.length+' แถวแล้ว วางในแผ่น '+sheet(day)+' ได้เลย'); } }};
      })(),
      memo: (function(){
        if (!s.memoOpen) return {open:false, rows:[], sum:[], poRows:[]};
        var fmt = function(n){ return Math.round(Number(n)||0).toLocaleString('en-US'); };
        var es = s.expShop || (s.stockShop || 'u'), shops = es === 'both' ? ['u','s'] : [es];
        var md = s.memoDate || TODAY, be = Number(md.slice(0,4))+543, yy = String(be).slice(-2), mon = md.slice(5,7), yearly = be >= 2570;
        var used = {}; Object.keys(s.batches).forEach(function(k){ var b = s.batches[k]; if (b.memo && b.memo.no) used[b.memo.no.replace(/\s+/g,' ')] = 1; });
        var raw = (s.memoNo||'').trim().replace(/\s+/g,' '), full = '', bad = !raw, chk;
        if (!raw) chk = {cls:'mchk', text:'ใช้เลขจากทะเบียนบันทึกข้อความ: กรอกแถวว่างแถวแรก ประเภท ซ - จัดซื้อ แล้วนำเลขในคอลัมน์ A มาใส่'};
        else if (!yearly){
          var m = raw.match(/^(?:สถพ\s*([ก-ฮ]{1,2})\s*(\d{2})\s*\/\s*)?(\d{4})$/);
          if (!m){ chk = {cls:'mchk bad', text:'รูปแบบไม่ตรง ควรเป็น สถพ ซ '+yy+'/'+mon+'01 หรือพิมพ์แค่ 4 หลัก (เดือน + ลำดับในเดือน)'}; bad = true; }
          else { var lt = m[1] || 'ซ', y = m[2] || yy, n = m[3]; full = 'สถพ ซ '+y+'/'+n;
            if (lt !== 'ซ'){ chk = {cls:'mchk bad', text:'บันทึกขอซื้อใช้ประเภท ซ (จัดซื้อ) ไม่ใช่ '+lt}; bad = true; }
            else if (Number(n.slice(0,2)) < 1 || Number(n.slice(0,2)) > 12 || n.slice(2) === '00'){ chk = {cls:'mchk bad', text:'สองหลักแรกคือเดือน (01–12) สองหลักหลังคือลำดับในเดือน (01 ขึ้นไป)'}; bad = true; }
            else if (used[full]){ chk = {cls:'mchk bad', text:'เลขนี้ใช้กับล็อตรับเข้าแล้ว ('+full+')'}; bad = true; }
            else if (y !== yy){ chk = {cls:'mchk warn', text:'ปีในเลขที่ ('+y+') ไม่ตรงกับวันที่บันทึก ('+yy+') ตรวจอีกครั้ง'}; }
            else if (n.slice(0,2) !== mon){ chk = {cls:'mchk warn', text:'เลขนี้เป็นของเดือน'+THMF[Number(n.slice(0,2))-1]+' แต่บันทึกลงวันที่เดือน'+THMF[Number(mon)-1]+' ตรวจอีกครั้ง'}; }
            else chk = {cls:'mchk ok', text:'จะใช้ '+full+' (ฉบับที่ '+Number(n.slice(2))+' ของเดือน'+THMF[Number(mon)-1]+')'}; }
        } else {
          var m2 = raw.match(/^(?:สถพ\s*([ก-ฮ]{1,2})\s*)?(\d{1,3})(?:\s*\/\s*(\d{4}))?$/);
          if (!m2){ chk = {cls:'mchk bad', text:'รูปแบบไม่ตรง ควรเป็น สถพ ซ 12/'+be+' หรือพิมพ์แค่ลำดับในปี'}; bad = true; }
          else { var lt2 = m2[1] || 'ซ', y2 = m2[3] || String(be), n2 = Number(m2[2]); full = 'สถพ ซ '+n2+'/'+y2;
            if (lt2 !== 'ซ'){ chk = {cls:'mchk bad', text:'บันทึกขอซื้อใช้ประเภท ซ (จัดซื้อ) ไม่ใช่ '+lt2}; bad = true; }
            else if (!n2){ chk = {cls:'mchk bad', text:'ลำดับในปีเริ่มที่ 1'}; bad = true; }
            else if (used[full]){ chk = {cls:'mchk bad', text:'เลขนี้ใช้กับล็อตรับเข้าแล้ว ('+full+')'}; bad = true; }
            else if (y2 !== String(be)){ chk = {cls:'mchk warn', text:'ปีในเลขที่ ('+y2+') ไม่ตรงกับวันที่บันทึก ('+be+') ตรวจอีกครั้ง'}; }
            else chk = {cls:'mchk ok', text:'จะใช้ '+full+' (ฉบับที่ '+n2+' ของปี '+be+')'}; }
        }
        var lc = function(pid){ var c0 = null; s.receipts.forEach(function(r){ var b = r.batch && s.batches[r.batch]; if (b && b.voided) return; if (r.pid === pid && r.cost && c0 === null) c0 = r.cost; }); return c0; };
        var tot = 0, nOn = 0, noCost = 0, picked = [];
        var rows = PRODUCTS.filter(function(p){ return shops.indexOf(p.shop) >= 0 && sugg(p.id).q > 0; }).map(function(p){
          var sg = sugg(p.id), q = s.memoQty[p.id] != null ? s.memoQty[p.id] : sg.q, on = !s.memoOff[p.id] && q > 0, c0 = lc(p.id), est = '';
          if (!c0){ PRODUCTS.forEach(function(o2){ if (c0 || o2.name !== p.name || o2.id === p.id) return; var c2 = lc(o2.id); if (c2){ c0 = c2; est = ' (ประมาณจาก #'+o2.size+')'; } }); }
          if (on){ nOn++; if (c0) tot += q*c0; else noCost++; picked.push({p:p, q:q, c:c0}); }
          return {label:itemLabel(p.id), why:sg.why, q:String(q), cost: c0 ? 'ทุน '+baht(c0)+(est ? ' (จาก'+est.replace(' (ประมาณจาก','').replace(')','')+')' : '') : 'ทุน — (ไม่มีบันทึก)',
            tipTitle: 'ทำไมแนะนำ '+sg.q+' ชิ้น',
            tip: [sg.since ? 'ขายไปแล้ว '+sg.sold+' ชิ้น หลังของเข้ารอบล่าสุด ('+sg.since+')' : 'ขายไปแล้ว '+sg.sold+' ชิ้น (ยังไม่เคยรับของเข้า)', 'มีนักเรียนขอไว้แต่ยังไม่ได้ของ '+sg.unmet+' คน', 'เหลือในสต็อก '+sg.have+' ชิ้น'],
            tipSum: [sg.sold ? 'ถ้าขายได้เท่ารอบที่แล้ว' : '', sg.unmet ? 'ถ้าให้คนที่ขอไว้ได้ครบ' : ''].filter(Boolean).join(' และ')+' ต้องใช้ '+(sg.sold+sg.unmet)+' ชิ้น ตอนนี้เหลือ '+sg.have+' จึงควรสั่งอย่างน้อย '+sg.q+' ชิ้น',
            tipNote: est ? 'ราคาทุนใช้ของไซส์'+est.replace(' (ประมาณจาก','').replace(')','')+' เพราะไซส์นี้ยังไม่มีบันทึก' : (c0 ? '' : 'ยังไม่มีราคาทุน ใส่เองในไฟล์ Word ได้'),
            cls:'moi'+(on ? '' : ' off'), on:on,
            boxLabel: on ? 'ไม่สั่งรายการนี้' : 'สั่งรายการนี้',
            tog: function(){ var o = Object.assign({}, self.state.memoOff); if (o[p.id]) delete o[p.id]; else o[p.id] = 1; self.setState({memoOff:o}); },
            onQty: function(e){ var o = Object.assign({}, self.state.memoQty); o[p.id] = Math.max(0, Math.floor(Number(e.target.value)||0)); self.setState({memoQty:o}); }}; });
        // summary by item (as in the memo body), sizes of the same price share a line (as in attachment 1)
        var G = {}, gord = [];
        picked.forEach(function(x){ var k = x.p.name; if (!G[k]){ G[k] = {name:k, unit: /ชุด/.test(k) ? 'ชุด' : (x.p.shop === 's' ? 'ชิ้น' : 'ตัว'), q:0, amt:0, lo:null, hi:null, byPrice:{}, pord:[]}; gord.push(k); }
          var g = G[k]; g.q += x.q; if (x.c){ g.amt += x.q*x.c; g.lo = g.lo === null ? x.c : Math.min(g.lo, x.c); g.hi = g.hi === null ? x.c : Math.max(g.hi, x.c); }
          var pk = String(x.c || 0); if (!g.byPrice[pk]){ g.byPrice[pk] = []; g.pord.push(pk); } g.byPrice[pk].push((x.p.size || '')+(x.p.size ? ' ' : '')+'('+x.q+')'); });
        var sum = gord.map(function(k, i){ var g = G[k];
          return {no:String(i+1), name:g.name, qty:g.q+' '+g.unit, price: g.lo === null ? '—' : (g.lo === g.hi ? fmt(g.lo) : fmt(g.lo)+' – '+fmt(g.hi)), amt: g.amt ? fmt(g.amt) : '—',
            sizes: g.pord.map(function(pk){ return (pk !== '0' ? fmt(Number(pk))+' บาท: ' : 'ไม่มีราคา: ')+g.byPrice[pk].join(', '); }).join('; ')}; });
        var subject = 'ขออนุมัติจัดซื้อ'+(es === 'u' ? 'เครื่องแบบนักเรียน' : (es === 's' ? 'เครื่องเขียน' : 'เครื่องแบบนักเรียนและเครื่องเขียน'));
        var vendor = s.memoVendor != null ? s.memoVendor : (es === 's' ? '' : 'ร้านธนากรพาณิชย์');
        var poNo = (full && !bad ? full.replace('สถพ ซ','สถพ สซ') : 'สถพ สซ '+(yearly ? '…/'+be : yy+'/…'))+'-01';
        var poName = function(p){ return p.shop === 'u' && p.size ? p.name+' ขนาด '+p.size : p.name; };
        var poRows = picked.map(function(x, i){ return {no:String(i+1), name:poName(x.p), q:String(x.q), price: x.c ? fmt(x.c) : '—', amt: x.c ? fmt(x.q*x.c) : '—'}; });
        return {open:true, no: s.memoNo, chkCls: chk.cls, chkText: chk.text, tickCls: 'mtick hide', tick: function(){}, tickLabel:'',
          date: s.memoDate, rows: rows, empty: rows.length === 0, vendor: vendor, subject: subject, sum: sum, hasSum: sum.length > 0,
          total: baht(tot), words: '('+bahtText(tot)+')', count: 'สั่ง '+nOn+' รายการ'+(noCost ? ', ยังไม่มีราคาทุน '+noCost+' รายการ (ใส่ในไฟล์ Word ได้)' : ''),
          poNo: poNo, poRows: poRows, poInfo: poRows.length+' รายการ, '+Math.max(1, Math.ceil(poRows.length/13))+' หน้า (หน้าละ 13 รายการ)', poTab: poNo.replace(/^สถพ สซ [^/]*\//,''),
          poCopy: function(){ if (!poRows.length){ self.flash('ยังไม่มีรายการที่จะสั่ง'); return; }
            var txt = picked.map(function(x){ return [poName(x.p), x.q, x.c || ''].join('\t'); }).join('\n'); try { navigator.clipboard.writeText(txt); } catch(e){}
            self.flash('คัดลอก '+poRows.length+' รายการแล้ว วางในคอลัมน์ รายการ ของแท็บ PO ใหม่ใน Purchase Order Log'); },
          file: full && !bad ? full.replace('/','-')+' '+subject+'.docx' : '(ใส่เลขที่ก่อน) '+subject+'.docx',
          go: function(){ if (bad){ self.flash(raw ? 'ตรวจเลขที่บันทึกข้อความอีกครั้ง' : 'ใส่เลขที่จากทะเบียนก่อน'); return; } if (!nOn){ self.flash('ยังไม่มีรายการที่จะสั่ง'); return; }
            self.setState({memoOpen:false, memoNo:'', memoQty:{}, memoOff:{}});
            self.flash('ต้นแบบ: ดาวน์โหลด '+full.replace('/','-')+' '+subject+'.docx (ของจริงเติมลงแม่แบบ Template - Memo ของโรงเรียน)'); }};
      })(),
      openMemo: function(){ if (self.state.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; } self.setState({memoOpen:true}); },
      closeMemo: function(){ self.setState({memoOpen:false}); },
      onMemoNo: function(e){ self.setState({memoNo:e.target.value}); },
      onMemoDate: function(e){ self.setState({memoDate:e.target.value}); },
      onMemoVendor: function(e){ self.setState({memoVendor:e.target.value}); },
      sp: (function(){ var c = s.stuPop; if (!c || !s.names[c]) return {open:false, buys:[], waits:[]};
        var sb = s.sales.filter(function(x){ return x.code === c && !x.voided; }).sort(function(a,b){ return (b.date+b.time).localeCompare(a.date+a.time); });
        var act = !!(latest && latest.classes[c]);
        var wl = byShop([].concat(readyList, heldList, waitingList).filter(function(d){ return d.code === c; }));
        return {open:true, name:s.names[c], subA:'รหัส '+c, subB:(self.classFor(c,'9999-12-31') || 'ไม่ทราบชั้น'), status: act ? 'อยู่ในรายชื่อ' : 'ไม่อยู่ในรายชื่อ', statusCls: act ? 'tag ok' : 'tag off',
          spent: baht(sb.reduce(function(a,x){ return a+saleTotal(x); },0)), cA: sb.length+' ครั้ง', cB: sb.reduce(function(a,x){ return a+x.items.reduce(function(b,i){ return b+i.qty; },0); },0)+' ชิ้น',
          waits: wl.map(function(d){ return {item:itemLabel(d.pid), st: dmdOpen(d) ? 'รอของ' : (d.notified ? 'แจ้งนักเรียนแล้ว' : 'ของมาแล้ว'), stCls: dmdOpen(d) ? 'dtag amber' : 'dtag ok'}; }), hasWaits: wl.length > 0,
          buys: sb.slice(0,5).map(function(x){ return {items:itemsText(x), date:thD2(x.date), total:baht(saleTotal(x))}; }), hasBuys: sb.length > 0, noBuys: sb.length === 0, more: sb.length > 5 ? 'และอีก '+(sb.length-5)+' ครั้ง' : '',
          full: function(){ self.setState({stuPop:null, waitOpen:false, wxOpen:false, askHist:null}); self.go('student',{studentId:c}); }};
      })(),
      closeSp: function(){ self.setState({stuPop:null}); },
      openWx: function(){ if (self.state.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; } self.setState({wxOpen:true, wxScope:'all', wxSort:'item', wxFmt:'pdf'}); },
      closeWx: function(){ self.setState({wxOpen:false}); },
      doWx: function(){ var st = self.state, f = st.wxFmt || 'pdf'; self.setState({wxOpen:false}); self.flash(f === 'text' ? 'คัดลอกแล้ว วางในไลน์ได้เลย' : 'ต้นแบบ: ดาวน์โหลดไฟล์แล้ว'); },
      resv: (function(){ var id = s.resvAsk; if (!id || !PMAP[id]) return {open:false, who:[]};
        return {open:true, title:itemLabel(id)+' จองไว้ให้นักเรียน', who:(RESV[id]||[]).map(function(d){ var w = whoOf(d); return {who:w.who, cls:w.cls, until:'ขอเมื่อ '+thShort(d.date)}; }),
          sub:'มีของ '+(s.stock[id]||0)+' ชิ้น จองไว้ '+(RESV[id]||[]).length+' ชิ้น ถ้าขายให้คนนี้ คนที่จองอาจไม่ได้ของ'}; })(),
      exp: (function(){
        var es = s.expShop || (s.stockShop || 'u'), fmt = s.expVer === 'own' ? 'pdf' : (s.expFmt || 'pdf'), ext = fmt === 'pdf' ? '.pdf' : '.jpg', d2 = thD2(TODAY);
        var shopsOut = es === 'both' ? ['u','s'] : [es];
        var files = shopsOut.map(function(k){ return 'สต็อก'+(k==='u' ? 'เครื่องแบบ' : 'เครื่องเขียน')+(s.expVer === 'own' ? ' สำหรับสั่งซื้อ' : '')+' '+d2+ext; });
        var seg = function(list, cur, key){ return list.map(function(o){ return {label:o[1], cls:'segb'+(cur===o[0] ? ' on' : ''), pick:function(){ var u = {}; u[key] = o[0]; self.setState(u); }}; }); };
        var pshop = shopsOut[0], prods = PRODUCTS.filter(function(p){ return p.shop === pshop; });
        var stc = function(n){ return 'ecell'+(n <= 0 ? ' z' : (n <= 5 ? ' l' : ' g')); };
        var lastCost = function(pid){ var c0 = null; s.receipts.forEach(function(r){ var b = r.batch && s.batches[r.batch]; if (b && b.voided) return; if (r.pid === pid && r.cost && c0 === null) c0 = r.cost; }); return c0; };
        var own = s.expVer === 'own', showC = own, showP = true, showD = own, tot = 0, valP = 0, valC = 0, dTot = 0;
        var mkRow = function(p){ var n = s.stock[p.id]||0, c0 = lastCost(p.id), pr = s.prices[p.id]||0, dn = (DMD[p.id] && DMD[p.id].n) || 0, tn = TERMP[p.id] || 0; tot += n; valP += n*pr; valC += n*(c0||0); dTot += dn;
          return {label: pshop === 'u' ? p.size : p.name, n:n, cls:stc(n), dmd: dn ? String(dn) : '—', dmdCls: dn ? 'ecell d' : '', term: tn ? String(tn) : '—', cost: c0 ? baht(c0) : '—', price: pr ? baht(pr) : '—'}; };
        var groups = [], gm = {};
        if (pshop === 'u'){
          prods.forEach(function(p){ if (!gm[p.name]){ gm[p.name] = {name:p.name, rows:[], total:0}; groups.push(gm[p.name]); } var r = mkRow(p); gm[p.name].rows.push(r); gm[p.name].total += r.n; });
          groups.forEach(function(g){ g.totalText = 'รวม '+g.total+' ชิ้น'; });
        }
        var rows = pshop === 's' ? prods.map(mkRow) : [];
        var cols = 'excols '+(own ? 'n5' : 'n2');
        return {open: !!s.expOpen, fmtTabs: tileOf(segIco(seg([['pdf','PDF'],['jpg','JPG']], fmt, 'expFmt'), ['pdf','jpg'])).map(function(k, i){ if (s.expVer === 'own' && i === 1){ k.cls += ' dis'; k.pick = function(){ self.flash('ฉบับสั่งซื้อส่งออกเป็น PDF เท่านั้น'); }; } return k; }),
          shopTabs: segIco(seg([['both','ทั้งหมด'],['u','เครื่องแบบ'],['s','เครื่องเขียน']], es, 'expShop'), ['','u','s']),
          showFmt: true, isOwn: s.expVer === 'own', memoCls: 'btn ghost'+(s.expVer === 'own' ? '' : ' dis'), memoGo: function(){ if (self.state.expVer !== 'own'){ self.flash('เลือก ใช้สำหรับ: สั่งซื้อ ก่อน'); return; } self.setState({memoOpen:true}); }, verTabs: segIco(seg([['school','ส่งโรงเรียน'],['own','สั่งซื้อ']], own ? 'own' : 'school', 'expVer'), ['school','cart']), showDmd: showD, listCls: 'exlist'+(own ? ' exwide' : ''),
          verNote: own ? 'เพิ่มคอลัมน์รอของ ขอทั้งเทอม และราคาทุน ใช้ประกอบการสั่งซื้อ ไม่ควรส่งต่อ (PDF เท่านั้น)' : 'คงเหลือและราคาขายเท่านั้น ไม่มีราคาทุนและชื่อนักเรียน', rowCls: 'exrow '+cols, headCls: 'exrow exth '+cols,
          files: files.map(function(f){ return {name:f}; }), fileNote: es === 'both' ? 'ได้ 2 ไฟล์ แยกเครื่องแบบกับเครื่องเขียน (ตัวอย่างด้านบนแสดงไฟล์เครื่องแบบ)' : (fmt === 'jpg' ? 'ถ้ายาวเกินหนึ่งหน้า จะแบ่งเป็นหลายรูปตามหน้า' : ''),
          previewTitle: 'รายงานสต็อก'+(pshop==='u' ? 'เครื่องแบบนักเรียน' : 'เครื่องเขียน')+(own ? ' (สำหรับสั่งซื้อ)' : ''), previewDate: 'ณ วันที่ '+thDate(TODAY)+' เวลา '+nowHM()+' น.',
          previewNote: es === 'both' ? 'ตัวอย่างไฟล์แรก (เครื่องแบบ)' : '', isU: pshop === 'u', isS: pshop === 's', showPrice: showP, showCost: showC, firstCol: pshop === 'u' ? 'ไซส์' : 'รายการ',
          groups: groups, rows: rows, total: 'รวม '+tot+' ชิ้น', value: [showD ? 'รอของ '+dTot : '', showC ? 'มูลค่าตามทุน '+baht(valC) : '', 'มูลค่าตามราคาขาย '+baht(valP)].filter(Boolean).join(', '),
          legend: own ? 'รอของ = ขอหลังของเข้าครั้งล่าสุด; ขอทั้งเทอม = นับตั้งแต่เปิดเทอม '+thD2(termStart)+' รวมคนที่แค่ถาม; ราคาทุนจาก PO ล่าสุด' : '',
          by: 'พิมพ์โดย '+DATA.myName};
      })(),
      openExp: function(){ if (self.state.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; } self.setState({expOpen:true, expShop:null, expVer:'school'}); },
      closeExp: function(){ self.setState({expOpen:false}); },
      doExp: function(){ var st = self.state, es = st.expShop || (st.stockShop||'u'), ext = (st.expVer === 'own' || (st.expFmt||'pdf') === 'pdf') ? '.pdf' : '.jpg';
        var names = (es === 'both' ? ['u','s'] : [es]).map(function(k){ return 'สต็อก'+(k==='u' ? 'เครื่องแบบ' : 'เครื่องเขียน')+(st.expVer === 'own' ? ' สำหรับสั่งซื้อ' : '')+' '+thD2(TODAY)+ext; });
        self.setState({expOpen:false}); self.flash('ต้นแบบ: ดาวน์โหลด '+names.join(' และ ')); },
      resvNo: function(){ self.setState({resvAsk:null}); },
      resvYes: function(){ var id = self.state.resvAsk, ok = Object.assign({}, self.state.resvOk); ok[id] = true; self.setState({resvAsk:null, resvOk:ok}); setCart(id, 1); },
      closeWant: function(){ self.setState({wantPop:null, wantFor:null}); },
      onWantCode: function(e){ self.setState({wantCode:e.target.value}); },
      onWantName: function(e){ self.setState({wantName:e.target.value}); },
      onWaitShop: function(e){ self.setState({waitShop:e.target.value}); },
      onWaitQ: function(e){ self.setState({waitQ:e.target.value}); },
      openWait: function(){ self.setState({waitOpen:true}); },
      closeWait: function(){ self.setState({waitOpen:false}); },
      copyReady: function(){ self.setState({wxOpen:true, wxScope:'ready', wxSort:'cls', wxFmt:'text'}); },
      doneAll: function(){ if (self.state.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; } var ids = {}; readyList.forEach(function(d){ ids[d.id] = 1; }); self.setState({demand: self.state.demand.map(function(x){ return ids[x.id] ? Object.assign({}, x, {notified:TODAY}) : x; })}); self.flash('แจ้งนักเรียนครบแล้ว (จองของไว้จนกว่าจะมารับของ)'); },
      quickAsk: function(){ var st = self.state, pid = st.wantPop; self.setState({demand: addDemand(st.demand, {date:TODAY, time:nowHM(), pid:pid, got:null, code:'', name:''}), wantPop:null, wantCode:'', wantName:''}); self.flash('จดไว้: มีคนถามหา '+itemLabel(pid)); },
      recordWant: function(){ var st = self.state, pid = st.wantPop;
        self.setState({demand: addDemand(st.demand, {date:TODAY, time:nowHM(), pid:pid, got:null, code:(st.wantCode||'').trim(), name:(st.wantName||'').trim()}), wantPop:null, wantCode:'', wantName:''});
        self.flash('จดไว้แล้ว: ลูกค้าต้องการ '+itemLabel(pid)+' (ไม่ได้ซื้อ)'); },
      onHQ: function(e){ self.setState({hQ:e.target.value}); },
      onHShop: function(e){ self.setState({hShop:e.target.value}); },
      toggleHFilter: function(){ self.setState({hFilterOpen:!self.state.hFilterOpen}); },
      closeHFilter: function(){ self.setState({hFilterOpen:false}); },
      clearHFilters: function(){ self.setState({hPay:'all', hStuSel:[], hSellerSel:[], hRange:'all'}); },
      onHFrom: function(e){ self.setState({hFrom:e.target.value}); },
      onHTo: function(e){ self.setState({hTo:e.target.value}); },
      closeSale: function(){ self.setState({saleId:null, saleEdit:null, saleVoidAsk:false}); },
      saleStartEdit: function(){ var st = self.state; if (st.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; }
        var x = st.sales.filter(function(y){ return y.id === st.saleId; })[0];
        self.setState({saleVoidAsk:false, saleEdit:{items: x.items.map(function(i){ return Object.assign({}, i); }), code: x.code, pay: x.transfer > 0 ? 'transfer' : 'cash', slip: x.transferTime || ''}}); },
      saleEditCancel: function(){ self.setState({saleEdit:null}); },
      saleAddOpen: function(){ self.setState({saleEdit: Object.assign({}, self.state.saleEdit, {adding:true, addName:'', swapIx:null})}); },
      saleAddClose: function(){ self.setState({saleEdit: Object.assign({}, self.state.saleEdit, {adding:false, addName:'', swapIx:null})}); },
      onSaleCode: function(e){ self.setState({saleEdit: Object.assign({}, self.state.saleEdit, {code:e.target.value})}); },
      onSaleSlip: function(e){ self.setState({saleEdit: Object.assign({}, self.state.saleEdit, {slip:e.target.value})}); },
      onSaleSlipFile: function(e){ var f = e && e.target && e.target.files && e.target.files[0], ed = self.state.saleEdit || {}, sx2 = self.state.sales.filter(function(x){ return x.id === self.state.saleId; })[0];
        var tm = (sx2 && sx2.time) || nowHM(), amt = (ed.items || []).reduce(function(a,i){ return a + i.qty*i.price; }, 0);
        self.setState({saleEdit: Object.assign({}, ed, {slipImg:{name: f ? f.name : 'slip.jpg', time: tm, amt: amt}, slip: tm})}); if (e && e.target) e.target.value = ''; },
      saleEditSave: function(){
        var st = self.state; if (st.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; }
        var ed = st.saleEdit, sk = Object.assign({}, st.stock), at = thDate(TODAY)+' '+nowHM();
        var items = ed.items.filter(function(i){ return i.qty > 0; });
        if (!items.length){ self.flash('ไม่มีสินค้าเหลือ ใช้ "ยกเลิกการขาย" แทน'); return; }
        var sales = st.sales.map(function(x){
          if (x.id !== st.saleId) return x;
          x.items.forEach(function(i){ sk[i.pid] = (sk[i.pid]||0) + i.qty; });
          items.forEach(function(i){ sk[i.pid] = (sk[i.pid]||0) - i.qty; });
          var tot = items.reduce(function(a,i){ return a+i.qty*i.price; },0), was = saleTotal(x);
          return Object.assign({}, x, {items: items, code: (ed.code||'').trim(), transfer: ed.pay === 'transfer' ? tot : 0, transferTime: ed.pay === 'transfer' ? ed.slip : '',
            log: (x.log||[]).concat([{at:at, by:DATA.myName, text:'แก้ไขรายการ: ยอด '+baht(was)+' → '+baht(tot)}])});
        });
        self.setState({sales:sales, stock:sk, saleEdit:null}); self.flash('บันทึกการแก้ไขแล้ว และปรับสต็อกตามส่วนต่าง'); },
      saleAskVoid: function(){ if (self.state.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; } self.setState({saleVoidAsk:true}); },
      saleCancelVoid: function(){ self.setState({saleVoidAsk:false}); },
      saleConfirmVoid: function(){
        var st = self.state; if (st.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; }
        var sk = Object.assign({}, st.stock), at = thDate(TODAY)+' '+nowHM(), q = 0;
        var sales = st.sales.map(function(x){
          if (x.id !== st.saleId) return x;
          x.items.forEach(function(i){ sk[i.pid] = (sk[i.pid]||0) + i.qty; q += i.qty; });
          return Object.assign({}, x, {voided:{at:at, by:DATA.myName}, log:(x.log||[]).concat([{at:at, by:DATA.myName, text:'ยกเลิกการขาย: คืนสต็อก '+q+' ชิ้น'}])});
        });
        self.setState({sales:sales, stock:sk, saleVoidAsk:false}); self.flash('ยกเลิกการขายแล้ว (คืนสต็อก '+q+' ชิ้น)'); },
      exportFake: function(){ if (self.state.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; } self.flash('ต้นแบบ: ส่งออก Excel ตามตัวกรองที่เลือก'); },
      stockView: stockView, tl: {on: tlOn, off: !tlOn, head: tlHead, months: tlMonths, range: thD2(TL0)+' ถึง '+thD2(TODAY), none: !tlHead.length,
        hiCls:'btn ghost sm tlhib'+(tlOnlyHi ? ' on' : ''), hiLabel: 'เฉพาะที่ไฮไลต์'+(tlHiN ? ' ('+tlHiN+')' : ''), hiPressed: tlOnlyHi ? 'true' : 'false',
        hiTgl: function(){ if (!tlHiN && !tlOnlyHi){ self.flash('แตะดาวหน้าไซส์เพื่อไฮไลต์ก่อน'); return; } self.setState({tlOnlyHi: !tlOnlyHi}); },
        foldLabel: tlAnyFold ? 'กางทั้งหมด' : 'พับทั้งหมด', foldAll: function(){ var f = Object.assign({}, self.state.tlFold||{}), pre = (shopU ? 'u' : 's')+':';
          Object.keys(f).forEach(function(k){ if (k.indexOf(pre) === 0) delete f[k]; }); if (!tlAnyFold) tlAllKeys.forEach(function(k){ f[k] = 1; }); self.setState({tlFold:f}); tlSave(f, self.state.tlHi||{}); },
        empty: tlOnlyHi && !stockCols.length}, stockShopChips: stockShopChips, stockChips: stockChips, stockRows: stockRows, stockEmpty: stockCols.length === 0,
      restockQty: s.restockQty, restockMemo: s.restockMemo, receipts: receipts,
      onRQty: function(e){ self.setState({restockQty:e.target.value}); },
      onRMemo: function(e){ self.setState({restockMemo:e.target.value}); },
      saveRestock: function(){if (self.state.role !== 'owner'){ self.flash('เฉพาะเจ้าของระบบ'); return; } 
        var st = self.state, id = st.restockOpen, qn = parseInt(st.restockQty,10);
        if (!id || !qn || qn < 1){ self.flash('กรอกจำนวนรับเข้า'); return; }
        var stock = Object.assign({}, st.stock); stock[id] += qn;
        self.setState({stock:stock, receipts:[{date:TODAY, label:itemLabel(id), qty:qn, memo:(st.restockMemo||'').trim()}].concat(st.receipts), restockOpen:null, restockQty:'', restockMemo:''});
        self.flash('รับเข้า '+itemLabel(id)+' +'+qn+' แล้ว');
      },
      hasUnmatched: unmatchedList.length > 0 && Object.keys(s.names||{}).length > 0, unmatchedCount: unmatchedList.length, unmatchedList: unmatchedList,
      studentQ: s.studentQ, onStudentQ: function(e){ self.setState({studentQ:e.target.value}); },
      classChips: classChips, stuF: stuF, studentRows: studentRows, studentsEmpty: studentRows.length === 0, stuNoRoster: studentRows.length === 0 && !Object.keys(s.names||{}).length, stuNoMatch: studentRows.length === 0 && Object.keys(s.names||{}).length > 0,
      stuClear: function(){ self.setState({studentQ:'', classSel:[], stuStSel:[], stuStatus:'all'}); },
      goImport: function(){ self.go('import'); }, goStudents: function(){ self.go('students'); },
      stu: stu,
      rosterRows: rosterRows, importSteps: importSteps, importNow: String((s.importStep||0)+1), importFill: 'width: '+[0,50,100][s.importStep||0]+'%', imp: imp, termInput: s.termInput, effDate: s.effDate,
      onTerm: function(e){ self.setState({termInput:e.target.value}); },
      onEff: function(e){ self.setState({effDate:e.target.value}); },
      chooseFile: function(){ self.setState({fileChosen:true}); },
      toPreview: function(){ if (!self.state.fileChosen){ self.flash('เลือกไฟล์ก่อน'); return; } if (!self.state.effDate){ self.flash('เลือกวันที่เริ่มใช้'); return; } if (self.state.rosters.length > 1){ self.flash('ต้นแบบนี้นำเข้าได้ครั้งเดียว กดรีเซ็ตก่อน'); return; } self.setState({importStep:1}); },
      backToForm: function(){ self.setState({importStep:0}); },
      confirmImport: function(){
        var st = self.state;
        var r2 = {id:'r2', term:st.termInput||'ไม่ระบุ', eff:st.effDate, classes:Object.assign({}, NEW_FILE.classes)};
        self.setState({rosters: st.rosters.concat([r2]), names: Object.assign({}, st.names, NEW_FILE.names), importStep:2, doneTerm:r2.term, doneEff:thDate(r2.eff)});
        self.flash('นำเข้ารายชื่อแล้ว');
      },
      resetImport: function(){ self.setState({rosters:DATA.rosters.slice(), names:Object.assign({}, BASE_NAMES), importStep:0, fileChosen:false}); },
      staffRows: staffRows, newStaff: s.newStaff,
      onNewStaff: function(e){ self.setState({newStaff:e.target.value}); },
      addStaff: function(){ var v = (self.state.newStaff||'').trim(); if (!v){ self.flash('กรอกอีเมลก่อน'); return; } self.setState({staff:self.state.staff.concat([{name:'ครูผู้ช่วยใหม่', email:v}]), newStaff:''}); self.flash('เพิ่ม '+v+' แล้ว'); },
      transferEmail: s.transferEmail, xfer: xfer,
      onXEmail: function(e){ self.setState({transferEmail:e.target.value}); },
      xferNext: function(){ if (!(self.state.transferEmail||'').trim()){ self.flash('กรอกอีเมลก่อน'); return; } self.setState({transferStep:1}); },
      xferBack: function(){ self.setState({transferStep:0}); },
      xferConfirm: function(){ self.setState({transferStep:2}); },
      xferReset: function(){ self.setState({transferStep:0, transferEmail:''}); },
      toast: s.toast, hasToast: !!s.toast
    };
  }
}
