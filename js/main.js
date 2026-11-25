import { mount, DCLogic } from './runtime.js';
import { start, persist, listen, DATA } from './data.js';

// Shown before sign-in, or when the signed-in account has no access.
class LoginOnly extends DCLogic {
  renderVals(){
    return { lg: { on: true, login: DATA.auth !== 'denied', denied: DATA.auth === 'denied',
      signIn: () => DATA.signIn(), other: () => DATA.signOut(), retry: () => location.reload(),
      name: DATA.user ? (DATA.user.displayName || DATA.me) : '', email: DATA.me, initial: (DATA.me || '?').charAt(0).toUpperCase() } };
  }
}

(async function(){
  try { await start(); } catch (e){ console.error('start failed', e); DATA.auth = 'login'; }
  const tpl = document.getElementById('app-template'), root = document.getElementById('root');
  if (DATA.auth !== 'in'){ mount(LoginOnly, tpl, root); return; }
  const { Component } = await import('./app-logic.js');
  Component.prototype.afterState = function(prev, next){ persist(prev, next); };
  const origMount = Component.prototype.componentDidMount;
  Component.prototype.componentDidMount = function(){
    if (origMount) origMount.call(this);
    listen(this);
  };
  mount(Component, tpl, root);
})();
