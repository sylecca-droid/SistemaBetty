// ============================================================
// ApiOffline.gs  — REEMPLAZA TODO el archivo ApiOffline.gs
// Incluye: POS offline + PENDIENTES (CANCELADO / PENDIENTE)
// ============================================================
const SYNC_DIAS = 3;
const SHEET_PENDIENTES = "PENDIENTES";

function doPost(e) {
  let out;
  try {
    const b = JSON.parse((e.postData && e.postData.contents) || "{}");
    switch (b.accion) {
      case "tipos":      out = { ok: true, tipos: obtenerTiposUsuario() }; break;
      case "login":      out = apiLogin_(b); break;
      case "inventario": out = apiInventario_(b); break;
      case "venta":      out = apiVenta_(b); break;
      case "pendiente":  apiAuth_(b.sync); out = pendienteGuardar_(b); break;
      case "pendientes": apiAuth_(b.sync); out = pendientesListar_(); break;
      case "logout":     PropertiesService.getScriptProperties().deleteProperty("sync_" + b.sync); out = { ok: true }; break;
      default:           out = { ok: false, error: "Acción no válida." };
    }
  } catch (err) {
    const m = String(err && err.message || err);
    out = { ok: false, error: m === "AUTH" ? "Sesión de sincronización vencida. Ingrese con internet." : m, auth: m === "AUTH" };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function apiLogin_(b) {
  const tipo = String(b.tipo || "").trim().toUpperCase();
  const cache = CacheService.getScriptCache(), k = "fall_" + tipo;
  const fallos = Number(cache.get(k)) || 0;
  if (fallos >= 5) return { ok: false, error: "Demasiados intentos. Espere 5 minutos." };
  const r = validarAccesoPOS(tipo, b.clave);
  if (!r.ok) { cache.put(k, String(fallos + 1), 300); return { ok: false, error: r.error || "Contraseña incorrecta." }; }
  cache.remove(k);
  return { ok: true, tipo: r.tipo, usuario: r.usuario, token: r.token, sync: crearSync_(r.tipo, r.usuario) };
}

// Token de sincronización: dura SYNC_DIAS días
function crearSync_(tipo, usuario) {
  const p = PropertiesService.getScriptProperties(), ahora = Date.now(), todo = p.getProperties();
  Object.keys(todo).forEach(function (k) {
    if (k.indexOf("sync_") !== 0) return;
    try { if (JSON.parse(todo[k]).exp < ahora) p.deleteProperty(k); } catch (e) { p.deleteProperty(k); }
  });
  const t = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, "");
  p.setProperty("sync_" + t, JSON.stringify({ tipo: tipo, usuario: usuario, exp: ahora + SYNC_DIAS * 864e5 }));
  return t;
}

function apiAuth_(t) {
  const raw = t ? PropertiesService.getScriptProperties().getProperty("sync_" + t) : null;
  if (!raw) throw new Error("AUTH");
  const s = JSON.parse(raw);
  if (s.exp < Date.now()) throw new Error("AUTH");
  return s;
}

// Inventario + marca PEND (columna T = "SI" -> pregunta CANCELADO/PENDIENTE)
function apiInventario_(b) {
  apiAuth_(b.sync);
  const r = obtenerDatosPOS();
  if (!r || !r.ok || !r.inventario || !r.inventario.length) return r;
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_INVENTARIO);
  const last = sh.getLastRow();
  if (last < 2) return r;
  const t = sh.getRange(2, 20, last - 1, 1).getValues(); // columna T
  r.inventario.forEach(function (p) {
    const fila = Number(p.fila);
    const v = (fila >= 2 && t[fila - 2]) ? String(t[fila - 2][0] || "").trim().toUpperCase() : "";
    p.PEND = v !== "" && v !== "NO" && v !== "0";
  });
  return r;
}

// Registra una venta hecha offline. El ID local evita duplicados si se reenvía.
function apiVenta_(b) {
  apiAuth_(b.sync);
  const id = String(b.idLocal || "");
  if (!id) return { ok: false, error: "Falta el ID de la venta." };

  const COL_ID = 21; // columna U de VENTAS guarda el ID local
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_VENTAS);
  const last = sh.getLastRow();
  if (last >= 2) {
    const ids = sh.getRange(2, COL_ID, last - 1, 1).getValues();
    const nums = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === id) {
        try { if (b.payload && b.payload.pendIds) marcarPendientesCancelados_(b.payload.pendIds, nums[i][0]); } catch (e) {}
        return { ok: true, numVenta: nums[i][0], duplicada: true };
      }
    }
  }

  const r = registrarVenta(b.payload);
  if (!r.ok) return r;

  const l2 = sh.getLastRow();
  const nums2 = sh.getRange(1, 1, l2, 1).getValues();
  for (let i = l2 - 1; i >= 1; i--) {
    if (nums2[i][0] === r.numVenta) {
      sh.getRange(i + 1, COL_ID).setValue(id);
      if (b.fechaLocal) sh.getRange(i + 1, 2).setValue(new Date(b.fechaLocal));
      break;
    }
  }

  // Los pendientes cobrados pasan a CANCELADO
  try { if (b.payload && b.payload.pendIds) marcarPendientesCancelados_(b.payload.pendIds, r.numVenta); } catch (e) {}

  return { ok: true, numVenta: r.numVenta, total: r.total };
}

// ============================================================
// PENDIENTES
// ============================================================
function hojaPendientes_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_PENDIENTES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PENDIENTES);
    sh.appendRow(["ID", "FECHA", "CLIENTE", "COD", "PRODUCTO", "PRECIO",
                  "VENDEDOR", "ESTADO", "FECHA CANCELADO", "N° VENTA"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function pendienteGuardar_(b) {
  const id = String(b.idLocal || "").trim();
  if (!id) return { ok: false, error: "Pendiente sin ID." };
  const cliente = String(b.cliente || "").trim().toUpperCase();
  if (!cliente) return { ok: false, error: "Falta el nombre del cliente." };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const sh = hojaPendientes_();
    const last = sh.getLastRow();
    if (last >= 2) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim().toLowerCase() === id.toLowerCase()) return { ok: true, repetido: true };
      }
    }
    let f = new Date(b.fechaLocal);
    if (isNaN(f.getTime())) f = new Date();
    sh.appendRow([id, f, cliente, String(b.productoId || ""), String(b.nombre || ""),
                  Number(b.precio) || 0, String(b.vendedor || ""), "PENDIENTE", "", ""]);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function pendientesListar_() {
  const sh = hojaPendientes_();
  const last = sh.getLastRow();
  const lista = [];
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 8).getValues().forEach(function (r) {
      if (String(r[7]).trim().toUpperCase() !== "PENDIENTE") return;
      lista.push({
        id: String(r[0]),
        fecha: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
        cliente: String(r[2]), productoId: String(r[3]), nombre: String(r[4]),
        precio: Number(r[5]) || 0, vendedor: String(r[6])
      });
    });
  }
  return { ok: true, pendientes: lista };
}

function marcarPendientesCancelados_(ids, numVenta) {
  if (!Array.isArray(ids) || !ids.length) return;
  const set = {};
  ids.forEach(function (x) { set[String(x).trim().toLowerCase()] = true; });
  const sh = hojaPendientes_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const data = sh.getRange(2, 1, last - 1, 1).getValues();
  const ahora = new Date();
  for (let i = 0; i < data.length; i++) {
    if (set[String(data[i][0]).trim().toLowerCase()]) {
      sh.getRange(i + 2, 8, 1, 3).setValues([["CANCELADO", ahora, numVenta]]);
    }
  }
}
