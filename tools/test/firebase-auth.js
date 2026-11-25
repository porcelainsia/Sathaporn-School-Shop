const U = window.__MOCK_USER === undefined ? {email:'a.satrabhaya@gmail.com', isAnonymous:false} : window.__MOCK_USER;
export function getAuth(){ return {}; }
export class GoogleAuthProvider { setCustomParameters(){} }
export function onAuthStateChanged(a, cb){ setTimeout(()=>cb(U),0); return ()=>{}; }
export async function signInWithPopup(){ window.__signedIn = true; }
export async function signInWithRedirect(){}
export async function signOut(){ window.__signedOut = true; }
