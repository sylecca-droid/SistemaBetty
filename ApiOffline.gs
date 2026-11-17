// ============================================================
// ApiOffline.gs  — pegar como archivo NUEVO en Apps Script
// Recibe las llamadas del POS offline (GitHub Pages)
// ============================================================
const SYNC_DIAS = 3;

function doPost(e) {
  let out;
  try {
    const b = JSON.parse((e.postData && e.postData.contents) || "{}");
    switch (b.accion) {
      case "tipos":      out = { ok: true, tipos: obtenerTiposUsuario() }; break;
      case "login":      out = apiLogin_(b); break;
      case "inventario": apiAuth_(b.sync); out = obtenerDatosPOS(); break;
      case "venta":      out = apiVenta_(b); break;
      case "ticketpdf":  apiAuth_(b.sync); out = generarTicketPDF(b.venta); break;
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

// Token de sincronización: dura SYNC_DIAS días (el de sesión normal vence en minutos)
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
      if (String(ids[i][0]) === id) return { ok: true, numVenta: nums[i][0], duplicada: true };
    }
  }

  const r = registrarVenta(b.payload);
  if (!r.ok) return r;

  const l2 = sh.getLastRow();
  const nums2 = sh.getRange(1, 1, l2, 1).getValues();
  for (let i = l2 - 1; i >= 1; i--) {
    if (nums2[i][0] === r.numVenta) {
      sh.getRange(i + 1, COL_ID).setValue(id);
      if (b.fechaLocal) sh.getRange(i + 1, 2).setValue(new Date(b.fechaLocal)); // fecha real de la venta
      break;
    }
  }
  return { ok: true, numVenta: r.numVenta, total: r.total };
}
