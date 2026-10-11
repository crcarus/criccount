// ─── api.js — capa de comunicación con el backend GAS ────────────────────

const GAS_URL = 'https://script.google.com/macros/s/REEMPLAZA_CON_TU_URL/exec';

async function gasCall(params) {
  const url = new URL(GAS_URL);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
  const res = await fetch(url.toString());
  if (!res.ok) throw new Error('Error HTTP ' + res.status);
  return res.json();
}

async function gasPost(action, body) {
  const res = await fetch(GAS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify({ action, ...body })
  });
  if (!res.ok) throw new Error('Error HTTP ' + res.status);
  return res.json();
}

// ── API pública ────────────────────────────────────────────────────────────

const API = {
  getDashboard:   (periodo) => gasCall({ action: 'getDashboard', periodo }),
  getRegistros:   (periodo) => gasCall({ action: 'getRegistros', periodo }),
  getPresupuesto: (periodo) => gasCall({ action: 'getPresupuesto', periodo }),
  getConfig:      ()        => gasCall({ action: 'getConfig' }),

  guardarGasto:       (data) => gasPost('guardarGasto', data),
  eliminarGasto:      (id)   => gasPost('eliminarGasto', { id }),
  editarGasto:        (data) => gasPost('editarGasto', data),
  guardarPresupuesto: (data) => gasPost('guardarPresupuesto', data),
  guardarConfig:      (data) => gasPost('guardarConfig', data),
};
