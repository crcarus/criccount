// ─── utils.js — helpers compartidos ───────────────────────────────────────

function formatCLP(n) {
  return '$' + Math.round(n).toLocaleString('es-CL');
}

function formatPeriodo(p) {
  const meses = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  const [y, m] = p.split('-');
  return meses[parseInt(m) - 1] + ' ' + y;
}

function showToast(msg, type = 'info') {
  let t = document.getElementById('toast');
  if (!t) {
    t = document.createElement('div');
    t.id = 'toast';
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.className = 'toast show ' + type;
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.classList.remove('show'), 3200);
}

function irA(page) {
  window.location.href = page + '.html';
}
