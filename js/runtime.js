// Tiny template runtime for the shop app.
// The page markup lives in <template id="app-template"> and uses three things:
//   {{path.to.value}}   a value from renderVals() (or from a loop variable)
//   <sc-for list="{{items}}" as="x">   repeat the children for each item
//   <sc-if value="{{flag}}">           show the children only when the value is truthy
// The logic class (js/app.js) keeps state, calls setState, and returns every value
// the template needs from renderVals(). Rendering is done with Preact.
import { h, render, Component as PComponent } from '../vendor/preact.module.js';

const BIND = /\{\{\s*([^}]+?)\s*\}\}/g;
const ONLY = /^\{\{\s*([^}]+?)\s*\}\}$/;

// HTML parsing lowercases attribute names; put back the names Preact expects.
const ATTR = {
  onclick: 'onClick', onchange: 'onChange', onfocus: 'onFocus', onpointerdown: 'onPointerDown',
  ondoubleclick: 'onDblClick', ondblclick: 'onDblClick', tabindex: 'tabIndex', autocomplete: 'autoComplete',
  inputmode: 'inputMode', readonly: 'readOnly', maxlength: 'maxLength', for: 'for'
};
const SKIP_ATTR = /^hint-placeholder-/;
const TEXT_INPUT = /^(text|number|search|tel|email|url|password|date|time|)$/;

function getter(path){
  const parts = path.split('.');
  return function(scope){
    let v;
    const head = parts[0];
    for (let i = scope.length - 1; i >= 0; i--){
      if (scope[i] && Object.prototype.hasOwnProperty.call(scope[i], head)){ v = scope[i][head]; break; }
    }
    for (let j = 1; j < parts.length; j++){ if (v == null) return undefined; v = v[parts[j]]; }
    return v;
  };
}

function textOf(str){
  const one = str.match(ONLY);
  if (one) { const g = getter(one[1]); return { raw: true, get: g }; }
  if (str.indexOf('{{') < 0) return null;
  const pieces = []; let last = 0; let m;
  BIND.lastIndex = 0;
  while ((m = BIND.exec(str))){
    if (m.index > last) pieces.push(str.slice(last, m.index));
    pieces.push(getter(m[1]));
    last = m.index + m[0].length;
  }
  if (last < str.length) pieces.push(str.slice(last));
  return { raw: false, get: function(scope){
    let out = '';
    for (const p of pieces) out += typeof p === 'string' ? p : (function(v){ return v == null ? '' : String(v); })(p(scope));
    return out;
  } };
}

function compileChildren(nodes){
  const out = [];
  nodes.forEach(function(n){ const c = compile(n); if (c) out.push(c); });
  return function(scope){
    const res = [];
    for (const c of out){ const v = c(scope); if (Array.isArray(v)) res.push.apply(res, v); else if (v != null && v !== false) res.push(v); }
    return res;
  };
}

function compile(node){
  if (node.nodeType === 3){
    const t = node.nodeValue;
    const b = textOf(t);
    if (!b) return function(){ return t; };
    return function(scope){ const v = b.get(scope); return v == null || v === false ? '' : (typeof v === 'object' ? '' : String(v)); };
  }
  if (node.nodeType !== 1) return null;
  const tag = node.localName;
  const kids = compileChildren(Array.from(node.childNodes));

  if (tag === 'sc-if'){
    const g = textOf(node.getAttribute('value') || '');
    return function(scope){ return g && g.get(scope) ? kids(scope) : null; };
  }
  if (tag === 'sc-for'){
    const g = textOf(node.getAttribute('list') || '');
    const as = node.getAttribute('as') || 'item';
    return function(scope){
      const list = g ? g.get(scope) : null;
      if (!list || !list.length) return null;
      const res = [];
      for (let i = 0; i < list.length; i++){
        const frame = {}; frame[as] = list[i]; frame[as + 'Index'] = i;
        const r = kids(scope.concat([frame]));
        for (const x of r) res.push(x);
      }
      return res;
    };
  }

  const isInput = tag === 'input' || tag === 'textarea';
  const statics = {}; const dyn = [];
  Array.from(node.attributes).forEach(function(a){
    if (SKIP_ATTR.test(a.name)) return;
    let name = ATTR[a.name] || a.name;
    if (name === 'onChange' && isInput && TEXT_INPUT.test((node.getAttribute('type') || '').toLowerCase())) name = 'onInput';
    const b = textOf(a.value);
    if (!b) statics[name] = a.value; else dyn.push([name, b]);
  });
  return function(scope){
    const props = Object.assign({}, statics);
    for (const d of dyn){
      const v = d[1].get(scope);
      if (d[1].raw && (v === undefined || v === null || v === false) && /^on|^value$/.test(d[0]) === false){ continue; }
      props[d[0]] = v;
    }
    if (isInput && props.value === undefined && 'value' in statics === false && dyn.some(function(d){ return d[0] === 'value'; })) props.value = '';
    const c = kids(scope);
    return h(tag, props, c.length ? c : undefined);
  };
}

// Base class the app's logic extends (same surface as the design prototype).
export class DCLogic {
  constructor(){ this.state = {}; this._host = null; }
  setState(patch, cb){
    const next = typeof patch === 'function' ? patch(this.state) : patch;
    this.state = Object.assign({}, this.state, next);
    if (this._host) this._host.forceUpdate(cb);
    else if (cb) cb();
  }
  forceUpdate(){ if (this._host) this._host.forceUpdate(); }
}

export function mount(LogicClass, templateEl, rootEl){
  const tpl = templateEl.content;
  const roots = Array.from(tpl.childNodes);
  const renderAll = compileChildren(roots);
  class Host extends PComponent {
    constructor(){ super(); this.logic = new LogicClass(); this.logic._host = this; }
    componentDidMount(){ if (this.logic.componentDidMount) this.logic.componentDidMount(); }
    componentDidUpdate(){ if (this.logic.componentDidUpdate) this.logic.componentDidUpdate(); }
    componentWillUnmount(){ if (this.logic.componentWillUnmount) this.logic.componentWillUnmount(); }
    render(){ const vals = this.logic.renderVals(); return h('div', { id: 'app-root' }, renderAll([vals])); }
  }
  render(h(Host), rootEl);
}
