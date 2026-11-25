"""Split the design prototype (Main.dc.html) into the app's files.

usage: python3 tools/from-prototype.py path/to/Main.dc.html
writes: css/app.css, index.html (template), js/app-logic.js
"""
import re, sys, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
src = pathlib.Path(sys.argv[1]).read_text(encoding='utf-8')

helmet = src[src.index('<helmet>') + 8: src.index('</helmet>')]
styles = re.findall(r'<style>(.*?)</style>', helmet, re.S)
links = re.findall(r'<link[^>]*>', helmet)
body = src[src.index('</helmet>') + 9: src.index('</x-dc>')].strip()
m = re.search(r'<script type="text/x-dc"[^>]*>(.*?)</script>', src, re.S)
logic = m.group(1)

(root / 'css').mkdir(exist_ok=True)
(root / 'css' / 'app.css').write_text('\n'.join(styles), encoding='utf-8')

logic = logic.replace('class Component extends DCLogic', 'export class Component extends DCLogic')
proto_logic = logic

# ---- live data: swap the prototype's sample data for what js/data.js loaded ----
def sub_block(text, start, end, repl):
    i = text.index(start)
    j = text.index(end, i) + len(end)
    return text[:i] + repl + text[j:]
def sub_once(text, old, new):
    assert text.count(old) == 1, ('not found once', old[:60], text.count(old))
    return text.replace(old, new)
live = logic
live = re.sub(r"var TODAY = '[0-9-]+';", "var TODAY = DATA.today;", live, count=1)
live = sub_block(live, 'var PRODUCTS = [', '\n];', 'var PRODUCTS = DATA.products;')
live = sub_block(live, 'var SEED_SALES = [', '\n];', 'var SEED_SALES = DATA.sales;')
live = sub_block(live, 'var SEED_RECEIPTS = [', '\n];', 'var SEED_RECEIPTS = DATA.receipts;')
live = sub_block(live, 'var SEED_BATCHES = {', '\n};', 'var SEED_BATCHES = DATA.batches;')
live = sub_block(live, 'var SEED_DEMAND = [', '\n];', 'var SEED_DEMAND = DATA.demand;')
live = sub_block(live, 'var BASE_NAMES = {', '};', 'var BASE_NAMES = DATA.names;')
live = sub_block(live, 'var R1 = {', '};', 'var R1 = DATA.rosters[DATA.rosters.length - 1] || {id:"r0", term:"", eff:DATA.today, classes:{}};')
live = live.replace("rosters:[R1],", "rosters:DATA.rosters.slice(),")
live = sub_once(live, "demand: SEED_DEMAND.map(function(d, i){ return Object.assign({id:'d'+i}, d); })", "demand: SEED_DEMAND.slice()")
live = re.sub(r"lastBackup:'[0-9-]+', lastBackupAt:'',", "lastBackup:DATA.lastBackup, lastBackupAt:DATA.lastBackupAt,", live, count=1)
live = re.sub(r"dStart:\{[^}]*\}, dMoves:\[\{.*?\}\],", "dStart:DATA.dStart, dMoves:DATA.dMoves, closes:DATA.closes, merges:DATA.merges,", live, count=1, flags=re.S)
live = re.sub(r"staff:\[\{name:.*?\}\],", "staff:DATA.staff,", live, count=1, flags=re.S)
live = sub_once(live, "page:'dash', role:'owner',", "page:DATA.role === 'owner' ? 'dash' : 'sellU', role:DATA.role, me:DATA.me,")
live = re.sub(r"(memoDate|effDate|editDate|rDate):'[0-9-]+'", r"\1:TODAY", live)
live = sub_once(live, "role:DATA.role, me:DATA.me,", "role:DATA.role, me:DATA.me, auth:DATA.auth,")
live = sub_once(live, "signIn: function(){ self.setState({auth:'in'}); self.okFlash('เข้าสู่ระบบแล้ว'); }, other: function(){ self.setState({auth:'login'}); }, retry: function(){ self.setState({auth:'in'}); }",
  "signIn: function(){ DATA.signIn(); }, other: function(){ DATA.signOut(); }, retry: function(){ location.reload(); }")
live = re.sub(r"\(?st\.role\s*===\s*'owner'\s*\?\s*'Porcelain'\s*:\s*'สมใจ'\)?", "DATA.myName", live)
live = live.replace("'Porcelain'", "DATA.myName").replace("(isOwner ? DATA.myName : 'สมใจ')", "DATA.myName")
live = sub_once(live, "var user = isOwner ? {name:DATA.myName, role:'เจ้าของระบบ', initial:'P'} : {name:'สมใจ', role:'ครูผู้ช่วย', initial:'ส'};",
  "var user = {name:DATA.myName, role: isOwner ? 'เจ้าของระบบ' : 'ครูผู้ช่วย', initial:(DATA.myName || '?').charAt(0).toUpperCase()};")
for k in ['DATA.products', 'DATA.sales', 'DATA.receipts', 'DATA.batches', 'DATA.demand', 'DATA.names', 'DATA.rosters.slice()', 'DATA.dStart', 'DATA.staff', 'DATA.lastBackup', 'DATA.role']:
    assert k in live, k
(root / 'js' / 'app-logic.js').write_text("import { DCLogic } from './runtime.js';\nimport { DATA } from './data.js';\n" + live, encoding='utf-8')

body = body.replace('<span class="av">ค</span><span class="wn">ครู ค.<small>teacher.c@sathaporn.ac.th</small></span>',
                    '<span class="av">{{lg.initial}}</span><span class="wn">{{lg.name}}<small>{{lg.email}}</small></span>')
assert '{{lg.email}}' in body
html = f'''<!doctype html>
<html lang="th">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>ร้านค้าโรงเรียน สถาพรศึกษา</title>
<link rel="icon" href="img/icon.png">
{chr(10).join(links)}
<link rel="stylesheet" href="css/app.css">\n<link rel="stylesheet" href="css/live.css">
</head>
<body>
<div id="root"></div>
<template id="app-template">
{body}
</template>
<script type="module" src="js/main.js"></script>
</body>
</html>
'''
(root / 'index.html').write_text(html, encoding='utf-8')
print('ok', len(styles), 'style blocks,', len(body), 'chars of markup,', len(logic), 'chars of logic')

# ---- catalog: product display info taken from the prototype's sample products ----
import json
U = re.findall(r"U\('([^']*)',\[([^\]]*)\],'([^']*)','([^']*)','([^']*)','([^']*)','([^']*)','([^']*)',(\d+),(\d+)\)", logic)
S = re.findall(r"S\('([^']*)','([^']*)','([^']*)','([^']*)',(\d+),(\d+),'([^']*)','([^']*)','([^']*)','([^']*)','([^']*)'\)", logic)
ucat, uorder = {}, []
for pid, groups, gender, part, tone, icon, name, size, price, stock in U:
    key = pid[:-(len(size) + 1)]
    if key not in ucat:
        ucat[key] = {'groups': [g.strip().strip("'") for g in groups.split(',') if g.strip()], 'gender': gender, 'part': part,
                     'tone': tone, 'icon': icon, 'name': name, 'sizes': []}
        uorder.append(key)
    ucat[key]['sizes'].append(size)
smeta, sorder = {}, []
for sid, cat, icon, name, price, stock, grp, short, sub, sg, lv in S:
    smeta[name] = {'cat': cat, 'icon': icon, 'grp': grp, 'short': short, 'sub': sub, 'sg': sg, 'lv': lv}
    sorder.append(name)
(root / 'js' / 'catalog.js').write_text(
    '// Display info for products (groups, icon, colour tone), generated from the design prototype.\n'
    'export const UCAT = ' + json.dumps(ucat, ensure_ascii=False) + ';\n'
    'export const UORDER = ' + json.dumps(uorder, ensure_ascii=False) + ';\n'
    'export const SMETA = ' + json.dumps(smeta, ensure_ascii=False) + ';\n'
    'export const SORDER = ' + json.dumps(sorder, ensure_ascii=False) + ';\n', encoding='utf-8')
print('catalog:', len(ucat), 'uniform categories,', len(smeta), 'stationery items')
