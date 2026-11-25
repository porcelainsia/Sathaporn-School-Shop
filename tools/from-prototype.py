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
(root / 'js' / 'app-logic.js').write_text("import { DCLogic } from './runtime.js';\n" + logic, encoding='utf-8')

html = f'''<!doctype html>
<html lang="th">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>ร้านค้าโรงเรียน สถาพรศึกษา</title>
<link rel="icon" href="img/icon.png">
{chr(10).join(links)}
<link rel="stylesheet" href="css/app.css">
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
