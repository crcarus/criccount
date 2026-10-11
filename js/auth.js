// ─── auth.js — autenticación con clave fija + localStorage ───────────────

const AUTH_KEY  = 'cc_auth';
const AUTH_PASS = '1234';   // ← cambia esto por tu clave real

function isAuthenticated() {
  return localStorage.getItem(AUTH_KEY) === 'ok';
}

function login(clave) {
  if (clave === AUTH_PASS) {
    localStorage.setItem(AUTH_KEY, 'ok');
    return true;
  }
  return false;
}

function logout() {
  localStorage.removeItem(AUTH_KEY);
  window.location.href = 'index.html';
}

// Llama esto al inicio de cada página (excepto index.html)
function requireAuth() {
  if (!isAuthenticated()) {
    window.location.href = 'index.html';
  }
}
