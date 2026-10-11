// ============================================================
// CriCount — Codigo.gs
// Backend como API JSON pura (GitHub Pages + GAS)
// ============================================================

const CONFIG = {
  SPREADSHEET_NAME   : "ControlBudget",
  SHEET_GASTOS       : "Gastos",
  SHEET_PRESUPUESTO  : "Presupuesto",
  SHEET_CATEGORIAS   : "Categorias",
  SHEET_PARTICIPANTES: "Participantes",
  EMAIL_NOTIFICACIONES: [
    "cristian.carusm@gmail.com",
    "hs.cmorales@gmail.com"
  ],
};

// ── Helpers CORS + JSON ──────────────────────────────────────
function _jsonOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function _jsonErr(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Punto de entrada GET ─────────────────────────────────────
// Uso: ?action=getDashboard&periodo=2026-10
//      ?action=getRegistros&periodo=2026-10
//      ?action=getPresupuesto&periodo=2026-10
//      ?action=getConfig
//      ?action=getResumenPeriodos
function doGet(e) {
  try {
    var p      = e && e.parameter ? e.parameter : {};
    var action = p.action || "";
    var periodo = p.periodo || _getMesActual();

    if (action === "getDashboard") {
      return _jsonOk(getDatosDashboard(periodo));
    }
    if (action === "getRegistros") {
      return _jsonOk(getGastosPorPeriodo(periodo));
    }
    if (action === "getResumenPeriodos") {
      return _jsonOk(getResumenPeriodos());
    }
    if (action === "getPresupuesto") {
      return _jsonOk({
        presupuesto : getPresupuestoPorPeriodo(periodo),
        categorias  : getCategorias(),
        participantes: getParticipantes(),
        periodos    : _getPeriodos(),
        mesActual   : _getMesActual(),
        periodoActivo: periodo,
      });
    }
    if (action === "getConfig") {
      return _jsonOk({
        categorias   : getCategorias(),
        participantes: getParticipantes(),
        mesActual    : _getMesActual(),
        periodos     : _getPeriodos(),
      });
    }
    if (action === "getPeriodos") {
      return _jsonOk({ periodos: getPeriodosConDatos() });
    }

    return _jsonErr("action desconocida: " + action);
  } catch (err) {
    return _jsonErr(err.message);
  }
}

// ── Punto de entrada POST ────────────────────────────────────
// Body: JSON con { action, ...datos }
function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action || "";

    if (action === "guardarGasto")       return _jsonOk(guardarGasto(body));
    if (action === "eliminarGasto")      return _jsonOk(eliminarGasto(body.id));
    if (action === "editarGasto")        return _jsonOk(editarGasto(body));
    if (action === "guardarPresupuesto") return _jsonOk(guardarPresupuesto(body));
    if (action === "eliminarPeriodo")    return _jsonOk(eliminarPeriodo(body.periodo));
    if (action === "replicarPeriodo")    return _jsonOk(replicarPeriodo(body.periodoOrigen, body.periodoDestino));
    if (action === "agregarCategoria")   return _jsonOk(agregarCategoria(body.nombre, body.emoji));
    if (action === "agregarParticipante")return _jsonOk(agregarParticipante(body.nombre));

    return _jsonErr("action desconocida: " + action);
  } catch (err) {
    return _jsonErr(err.message);
  }
}

// ── Spreadsheet ──────────────────────────────────────────────
function inicializarSpreadsheet() {
  var ss;
  var arch = DriveApp.getFilesByName(CONFIG.SPREADSHEET_NAME);
  ss = arch.hasNext()
    ? SpreadsheetApp.open(arch.next())
    : SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
  _crearHojaGastos(ss);
  _crearHojaPresupuesto(ss);
  _crearHojaCategorias(ss);
  _crearHojaParticipantes(ss);
  return ss.getId();
}
function _getOrCreate(ss, nombre) {
  return ss.getSheetByName(nombre) || ss.insertSheet(nombre);
}
function _crearHojaGastos(ss) {
  var h = _getOrCreate(ss, CONFIG.SHEET_GASTOS);
  if (h.getLastRow() === 0) {
    var cols = ["ID","Periodo","Fecha","Titulo","Categoria","Monto Total CLP",
                "Responsable","Division","Monto P1","Monto P2","Notas","Timestamp"];
    h.appendRow(cols);
    h.getRange(1,1,1,cols.length).setBackground("#1a1a2e").setFontColor("#fff").setFontWeight("bold");
    h.setFrozenRows(1);
  }
}
function _crearHojaPresupuesto(ss) {
  var h = _getOrCreate(ss, CONFIG.SHEET_PRESUPUESTO);
  if (h.getLastRow() === 0) {
    var cols = ["Periodo","Categoria","Presupuesto CLP","Timestamp"];
    h.appendRow(cols);
    h.getRange(1,1,1,cols.length).setBackground("#1a1a2e").setFontColor("#fff").setFontWeight("bold");
    h.setFrozenRows(1);
  }
}
function _crearHojaCategorias(ss) {
  var h = _getOrCreate(ss, CONFIG.SHEET_CATEGORIAS);
  if (h.getLastRow() === 0) {
    h.appendRow(["Nombre","Emoji","Activa"]);
    h.getRange(1,1,1,3).setBackground("#1a1a2e").setFontColor("#fff").setFontWeight("bold");
    [
      ["Supermercado","🛒"],["Almacén","🏪"],["Panadería","🥐"],["Feria","🥬"],
      ["Arriendo","🏠"],["Gasto Común","🏢"],["Luz","💡"],["Agua","💧"],
      ["Gas","🔥"],["Uber/Cabi","🚗"],["Viajes","✈️"],["Veterinario","🐾"],
      ["Regalos","🎁"],["Aseo","🧹"],["Paseos","🌿"],["Juntas amigos","🍻"],
      ["Restaurante","🍽️"],["Otro","📌"]
    ].forEach(function(c) { h.appendRow([c[0], c[1], true]); });
    h.setFrozenRows(1);
  }
}
function _crearHojaParticipantes(ss) {
  var h = _getOrCreate(ss, CONFIG.SHEET_PARTICIPANTES);
  if (h.getLastRow() === 0) {
    h.appendRow(["Nombre","Activo"]);
    h.getRange(1,1,1,2).setBackground("#1a1a2e").setFontColor("#fff").setFontWeight("bold");
    h.appendRow(["Claush", true]);
    h.appendRow(["Cri", true]);
    h.setFrozenRows(1);
  }
}

// ── Helpers internos ─────────────────────────────────────────
function _normalizarPeriodo(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return val.getFullYear() + "-" + String(val.getMonth()+1).padStart(2,"0");
  }
  var s = String(val).trim();
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  var d = new Date(s);
  if (!isNaN(d)) return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0");
  return s;
}
function _getSS() {
  var f = DriveApp.getFilesByName(CONFIG.SPREADSHEET_NAME);
  if (f.hasNext()) return SpreadsheetApp.open(f.next());
  inicializarSpreadsheet();
  return SpreadsheetApp.open(DriveApp.getFilesByName(CONFIG.SPREADSHEET_NAME).next());
}
function _getMesActual() {
  var h = new Date();
  return h.getFullYear() + "-" + String(h.getMonth()+1).padStart(2,"0");
}
function _getPeriodos() {
  var parts = _getMesActual().split("-").map(Number);
  var y = parts[0], m = parts[1];
  return Array.from({length:24}, function(_,i) {
    var mes = m - i, yr = y;
    if (mes <= 0) { mes += 12; yr--; }
    return yr + "-" + String(mes).padStart(2,"0");
  });
}
function _generarId() { return "G" + Date.now().toString(36).toUpperCase(); }
function _formatFecha(f) {
  if (!f) return "";
  var d = new Date(f);
  if (isNaN(d)) return String(f);
  return String(d.getDate()).padStart(2,"0") + "/" + String(d.getMonth()+1).padStart(2,"0");
}
function _formatFechaCompleta(f) {
  if (!f) return "";
  var d = new Date(f);
  if (isNaN(d)) return String(f);
  return String(d.getDate()).padStart(2,"0") + "/" + String(d.getMonth()+1).padStart(2,"0") + "/" + d.getFullYear();
}
function _formatFechaISO(f) {
  if (!f) return "";
  var d = new Date(f);
  if (isNaN(d)) return "";
  return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0");
}

// ── Datos base ───────────────────────────────────────────────
function getCategorias() {
  return _getSS().getSheetByName(CONFIG.SHEET_CATEGORIAS)
    .getDataRange().getValues().slice(1)
    .filter(function(r) { return r[2]===true || r[2]==="TRUE" || r[2]==="true"; })
    .map(function(r) { return { nombre: r[0], emoji: r[1] }; });
}
function getParticipantes() {
  return _getSS().getSheetByName(CONFIG.SHEET_PARTICIPANTES)
    .getDataRange().getValues().slice(1)
    .filter(function(r) { return r[1]===true || r[1]==="TRUE" || r[1]==="true"; })
    .map(function(r) { return r[0]; });
}

// ── Notificación email ───────────────────────────────────────
function _enviarNotificacionGasto(gasto, id, mp1, mp2) {
  try {
    var partes = getParticipantes();
    var meses  = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
                  "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
    var ym     = gasto.periodo.split("-");
    var periodoLabel = meses[parseInt(ym[1])-1] + " " + ym[0];
    var divMap = { "50/50":"50 / 50", "p1":"Solo "+(partes[0]||"P1"),
                   "p2":"Solo "+(partes[1]||"P2"), "personalizado":"Personalizado" };
    var divisionLabel = divMap[gasto.division] || gasto.division;
    var montoFmt = function(n) { return "$" + Math.round(n).toLocaleString("es-CL"); };
    var asunto = "Nuevo gasto: " + gasto.titulo + " - " + montoFmt(gasto.monto);
    var html = "<div style='font-family:Arial,sans-serif;max-width:520px;margin:0 auto'>"
      + "<h2 style='background:#6366f1;color:#fff;padding:16px;border-radius:8px 8px 0 0;margin:0'>CriCount - Nuevo gasto</h2>"
      + "<div style='border:1px solid #e2e8f0;border-top:none;padding:20px;border-radius:0 0 8px 8px'>"
      + "<table style='width:100%;border-collapse:collapse;font-size:14px'>"
      + "<tr><td style='padding:8px;color:#64748b'>Titulo</td><td style='padding:8px;font-weight:bold'>" + gasto.titulo + "</td></tr>"
      + "<tr style='background:#f8fafc'><td style='padding:8px;color:#64748b'>Monto</td><td style='padding:8px;font-weight:bold;font-size:16px;color:#6366f1'>" + montoFmt(gasto.monto) + "</td></tr>"
      + "<tr><td style='padding:8px;color:#64748b'>Categoria</td><td style='padding:8px'>" + gasto.categoria + "</td></tr>"
      + "<tr style='background:#f8fafc'><td style='padding:8px;color:#64748b'>Fecha</td><td style='padding:8px'>" + _formatFechaCompleta(gasto.fecha) + "</td></tr>"
      + "<tr><td style='padding:8px;color:#64748b'>Periodo</td><td style='padding:8px'>" + periodoLabel + "</td></tr>"
      + "<tr style='background:#f8fafc'><td style='padding:8px;color:#64748b'>Responsable</td><td style='padding:8px'>" + gasto.responsable + "</td></tr>"
      + "<tr><td style='padding:8px;color:#64748b'>Division</td><td style='padding:8px'>" + divisionLabel + "</td></tr>"
      + "<tr style='background:#f8fafc'><td style='padding:8px;color:#64748b'>" + (partes[0]||"P1") + "</td><td style='padding:8px'>" + montoFmt(mp1) + "</td></tr>"
      + "<tr><td style='padding:8px;color:#64748b'>" + (partes[1]||"P2") + "</td><td style='padding:8px'>" + montoFmt(mp2) + "</td></tr>"
      + (gasto.notas ? "<tr style='background:#f8fafc'><td style='padding:8px;color:#64748b'>Notas</td><td style='padding:8px;font-style:italic'>" + gasto.notas + "</td></tr>" : "")
      + "</table>"
      + "<p style='margin:16px 0 0;font-size:11px;color:#94a3b8;text-align:center'>ID: " + id + " · " + _formatFechaCompleta(new Date()) + "</p>"
      + "</div></div>";
    CONFIG.EMAIL_NOTIFICACIONES.forEach(function(email) {
      MailApp.sendEmail({ to: email, subject: asunto, htmlBody: html });
    });
  } catch(err) {
    Logger.log("Error email: " + err.message);
  }
}

// ── Guardar Gasto ────────────────────────────────────────────
function guardarGasto(gasto) {
  try {
    var hoja  = _getSS().getSheetByName(CONFIG.SHEET_GASTOS);
    var id    = _generarId();
    var monto = parseFloat(gasto.monto) || 0;
    var periodo = String(gasto.periodo).trim();
    var mp1 = 0, mp2 = 0;
    if (gasto.division === "50/50") {
      mp1 = Math.round(monto/2); mp2 = monto - mp1;
    } else if (gasto.division === "p1") {
      mp1 = monto; mp2 = 0;
    } else if (gasto.division === "p2") {
      mp1 = 0; mp2 = monto;
    } else if (gasto.division === "personalizado") {
      mp1 = parseFloat(gasto.montoClaush) || 0;
      mp2 = parseFloat(gasto.montoCri)    || 0;
    }
    hoja.appendRow([id, periodo, gasto.fecha, gasto.titulo, gasto.categoria,
      monto, gasto.responsable, gasto.division, mp1, mp2, gasto.notas||"", new Date()]);
    _enviarNotificacionGasto(gasto, id, mp1, mp2);
    return { ok: true, id: id, periodo: periodo };
  } catch(err) { return { ok: false, error: err.message }; }
}

// ── Eliminar Gasto ───────────────────────────────────────────
function eliminarGasto(id) {
  try {
    var hoja  = _getSS().getSheetByName(CONFIG.SHEET_GASTOS);
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0] === id) { hoja.deleteRow(i+1); return { ok: true }; }
    }
    return { ok: false, error: "No encontrado" };
  } catch(err) { return { ok: false, error: err.message }; }
}

// ── Editar Gasto ─────────────────────────────────────────────
function editarGasto(gasto) {
  try {
    var hoja   = _getSS().getSheetByName(CONFIG.SHEET_GASTOS);
    var datos  = hoja.getDataRange().getValues();
    var monto  = parseFloat(gasto.monto) || 0;
    var periodo = String(gasto.periodo).trim();
    var mp1 = 0, mp2 = 0;
    if (gasto.division === "50/50") {
      mp1 = Math.round(monto/2); mp2 = monto - mp1;
    } else if (gasto.division === "p1") {
      mp1 = monto; mp2 = 0;
    } else if (gasto.division === "p2") {
      mp1 = 0; mp2 = monto;
    } else if (gasto.division === "personalizado") {
      mp1 = parseFloat(gasto.montoClaush) || 0;
      mp2 = parseFloat(gasto.montoCri)    || 0;
    }
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0] === gasto.id) {
        var fila = i + 1;
        hoja.getRange(fila,2).setValue(periodo);
        hoja.getRange(fila,3).setValue(gasto.fecha);
        hoja.getRange(fila,4).setValue(gasto.titulo);
        hoja.getRange(fila,5).setValue(gasto.categoria);
        hoja.getRange(fila,6).setValue(monto);
        hoja.getRange(fila,7).setValue(gasto.responsable);
        hoja.getRange(fila,8).setValue(gasto.division);
        hoja.getRange(fila,9).setValue(mp1);
        hoja.getRange(fila,10).setValue(mp2);
        hoja.getRange(fila,11).setValue(gasto.notas||"");
        return { ok: true, periodo: periodo };
      }
    }
    return { ok: false, error: "Gasto no encontrado" };
  } catch(err) { return { ok: false, error: err.message }; }
}

// ── Eliminar período ─────────────────────────────────────────
function eliminarPeriodo(periodo) {
  try {
    var ss = _getSS();
    var hG = ss.getSheetByName(CONFIG.SHEET_GASTOS);
    var hP = ss.getSheetByName(CONFIG.SHEET_PRESUPUESTO);
    var gastos = hG.getDataRange().getValues();
    for (var i = gastos.length-1; i >= 1; i--) {
      if (_normalizarPeriodo(gastos[i][1]) === periodo) hG.deleteRow(i+1);
    }
    var presup = hP.getDataRange().getValues();
    for (var j = presup.length-1; j >= 1; j--) {
      if (_normalizarPeriodo(presup[j][0]) === periodo) hP.deleteRow(j+1);
    }
    return { ok: true };
  } catch(err) { return { ok: false, error: err.message }; }
}

// ── Replicar período ─────────────────────────────────────────
function replicarPeriodo(periodoOrigen, periodoDestino) {
  try {
    var hoja  = _getSS().getSheetByName(CONFIG.SHEET_PRESUPUESTO);
    var all   = hoja.getDataRange().getValues();
    var items = [];
    all.slice(1).forEach(function(r) {
      if (_normalizarPeriodo(r[0]) === periodoOrigen) {
        items.push({ categoria: r[1], monto: parseFloat(r[2])||0 });
      }
    });
    if (items.length === 0) return { ok: false, error: "El periodo origen no tiene presupuesto" };
    guardarPresupuesto({ periodo: periodoDestino, items: items });
    return { ok: true, copiados: items.length };
  } catch(err) { return { ok: false, error: err.message }; }
}

// ── Presupuesto ──────────────────────────────────────────────
function guardarPresupuesto(datos) {
  try {
    var hoja   = _getSS().getSheetByName(CONFIG.SHEET_PRESUPUESTO);
    var all    = hoja.getDataRange().getValues();
    var ts     = new Date();
    var periodo = String(datos.periodo).trim();
    datos.items.forEach(function(item) {
      var monto = parseFloat(item.monto);
      var valor = isNaN(monto) ? 0 : monto;
      var fila  = -1;
      for (var i = 1; i < all.length; i++) {
        if (_normalizarPeriodo(all[i][0]) === periodo && all[i][1] === item.categoria) {
          fila = i+1; break;
        }
      }
      if (fila > 0) {
        hoja.getRange(fila,3).setValue(valor);
        hoja.getRange(fila,4).setValue(ts);
      } else {
        hoja.appendRow([periodo, item.categoria, valor, ts]);
      }
    });
    return { ok: true };
  } catch(err) { return { ok: false, error: err.message }; }
}
function getPresupuestoPorPeriodo(periodo) {
  var result = {};
  _getSS().getSheetByName(CONFIG.SHEET_PRESUPUESTO)
    .getDataRange().getValues().slice(1)
    .forEach(function(r) {
      if (_normalizarPeriodo(r[0]) === String(periodo).trim())
        result[r[1]] = parseFloat(r[2]) || 0;
    });
  return result;
}

// ── Registros ────────────────────────────────────────────────
function getResumenPeriodos() {
  var hoja   = _getSS().getSheetByName(CONFIG.SHEET_GASTOS);
  var partes = getParticipantes();
  var map    = {};
  hoja.getDataRange().getValues().slice(1)
    .filter(function(r) { return r[1]; })
    .forEach(function(r) {
      var p = _normalizarPeriodo(r[1]);
      if (!p) return;
      if (!map[p]) map[p] = { periodo:p, total:0, p1:0, p2:0, count:0 };
      map[p].total += parseFloat(r[5]) || 0;
      map[p].p1    += parseFloat(r[8]) || 0;
      map[p].p2    += parseFloat(r[9]) || 0;
      map[p].count++;
    });
  var actual = _getMesActual();
  if (!map[actual]) map[actual] = { periodo:actual, total:0, p1:0, p2:0, count:0 };
  var hP = _getSS().getSheetByName(CONFIG.SHEET_PRESUPUESTO);
  hP.getDataRange().getValues().slice(1).forEach(function(r) {
    var p = _normalizarPeriodo(r[0]);
    if (p && !map[p]) map[p] = { periodo:p, total:0, p1:0, p2:0, count:0 };
  });
  return {
    periodos     : Object.values(map).sort(function(a,b){ return b.periodo.localeCompare(a.periodo); }),
    participantes: partes,
    categorias   : getCategorias(),
    mesActual    : _getMesActual(),
  };
}

function getGastosPorPeriodo(periodo) {
  var partes = getParticipantes();
  var gastos = _getSS().getSheetByName(CONFIG.SHEET_GASTOS)
    .getDataRange().getValues().slice(1)
    .filter(function(r) { return _normalizarPeriodo(r[1]) === String(periodo).trim(); })
    .map(function(r) {
      return {
        id         : r[0],
        periodo    : _normalizarPeriodo(r[1]),
        fecha      : r[2] ? _formatFechaCompleta(r[2]) : "",
        fechaRaw   : r[2] ? _formatFechaISO(r[2]) : "",
        titulo     : r[3],
        categoria  : r[4],
        monto      : parseFloat(r[5]) || 0,
        responsable: r[6],
        division   : r[7],
        montoP1    : parseFloat(r[8]) || 0,
        montoP2    : parseFloat(r[9]) || 0,
        notas      : r[10] || "",
      };
    })
    .sort(function(a,b) { return a.fechaRaw.localeCompare(b.fechaRaw); });
  return {
    gastos       : gastos,
    total        : gastos.reduce(function(s,g){ return s+g.monto; }, 0),
    totalP1      : gastos.reduce(function(s,g){ return s+g.montoP1; }, 0),
    totalP2      : gastos.reduce(function(s,g){ return s+g.montoP2; }, 0),
    periodo      : periodo,
    participantes: partes,
    categorias   : getCategorias(),
  };
}

// ── Dashboard ────────────────────────────────────────────────
function getDatosDashboard(periodo) {
  var partes = getParticipantes();
  var todos  = _getSS().getSheetByName(CONFIG.SHEET_GASTOS)
    .getDataRange().getValues().slice(1)
    .filter(function(r) { return _normalizarPeriodo(r[1]) === String(periodo).trim(); });

  var totalGastado = todos.reduce(function(s,r){ return s+(parseFloat(r[5])||0); }, 0);
  var totalP1      = todos.reduce(function(s,r){ return s+(parseFloat(r[8])||0); }, 0);
  var totalP2      = todos.reduce(function(s,r){ return s+(parseFloat(r[9])||0); }, 0);
  var pagadoP1     = todos.filter(function(r){ return r[6]===partes[0]; })
                          .reduce(function(s,r){ return s+(parseFloat(r[5])||0); }, 0);
  var pagadoP2     = todos.filter(function(r){ return r[6]===partes[1]; })
                          .reduce(function(s,r){ return s+(parseFloat(r[5])||0); }, 0);
  var deudaP1      = pagadoP1 - totalP1;

  var porCategoria = {};
  todos.forEach(function(r) {
    var cat = r[4];
    if (!porCategoria[cat]) porCategoria[cat] = 0;
    porCategoria[cat] += parseFloat(r[5]) || 0;
  });

  var presupuesto      = getPresupuestoPorPeriodo(periodo);
  var totalPresupuesto = Object.values(presupuesto).reduce(function(s,v){ return s+v; }, 0);

  var diasSet = new Set(todos.map(function(r){ return _formatFechaISO(r[2]); }).filter(Boolean));
  var diasConGasto = diasSet.size || 1;
  var promDiario   = totalGastado / diasConGasto;

  var porDia = {};
  todos.forEach(function(r) {
    var f = _formatFecha(r[2]);
    if (!porDia[f]) porDia[f] = 0;
    porDia[f] += parseFloat(r[5]) || 0;
  });
  var acumulado = {};
  var acc = 0;
  Object.keys(porDia).sort().forEach(function(d) { acc += porDia[d]; acumulado[d] = acc; });

  var recientes = todos.slice(-8).reverse().map(function(r) {
    return {
      fecha      : _formatFecha(r[2]),
      titulo     : r[3],
      categoria  : r[4],
      monto      : parseFloat(r[5]) || 0,
      responsable: r[6],
    };
  });

  var gastosDetalle = todos.map(function(r) {
    return {
      titulo     : r[3],
      categoria  : r[4],
      monto      : parseFloat(r[5]) || 0,
      responsable: r[6],
      fecha      : _formatFecha(r[2]),
    };
  });

  return {
    periodo, totalGastado, totalP1, totalP2,
    pagadoP1, pagadoP2, deudaP1,
    totalPresupuesto, porCategoria, presupuesto,
    promDiario, porDia, acumulado, recientes, gastosDetalle,
    categorias   : getCategorias(),
    participantes: partes,
    periodos     : getPeriodosConDatos(),
  };
}

function getPeriodosConDatos() {
  var set = {};
  _getSS().getSheetByName(CONFIG.SHEET_GASTOS)
    .getDataRange().getValues().slice(1)
    .map(function(r){ return _normalizarPeriodo(r[1]); })
    .filter(Boolean)
    .forEach(function(p){ set[p] = true; });
  var periodos = Object.keys(set);
  periodos.sort().reverse();
  var actual = _getMesActual();
  if (periodos.indexOf(actual) === -1) periodos.unshift(actual);
  return periodos;
}

// ── Config / extras ──────────────────────────────────────────
function agregarCategoria(nombre, emoji) {
  try {
    _getSS().getSheetByName(CONFIG.SHEET_CATEGORIAS).appendRow([nombre, emoji||"📌", true]);
    return { ok: true };
  } catch(err) { return { ok: false, error: err.message }; }
}
function agregarParticipante(nombre) {
  try {
    _getSS().getSheetByName(CONFIG.SHEET_PARTICIPANTES).appendRow([nombre, true]);
    return { ok: true };
  } catch(err) { return { ok: false, error: err.message }; }
}

function setup() {
  Logger.log("ID: " + inicializarSpreadsheet());
}
