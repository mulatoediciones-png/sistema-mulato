/**
 * Sistema MULATO EDICIONES — puente entre la página (GitHub Pages) y este Google Sheet.
 *
 * Instalación (una sola vez):
 *  1. En el Sheet: Extensiones > Apps Script. Borra lo que haya y pega este archivo completo.
 *  2. Cambia CLAVE por tu clave secreta (la misma que escribes en la página).
 *  3. Guarda. Arriba elige la función "instalar" y presiona Ejecutar. Acepta los permisos.
 *  4. Implementar > Nueva implementación > tipo "Aplicación web".
 *     Ejecutar como: Yo. Quién tiene acceso: Cualquier usuario.
 *  5. Copia la URL de la aplicación web y pégala en la página junto con tu clave.
 *
 * Tus datos nunca salen de este Sheet sin la clave.
 */

const CLAVE = 'PEGA_AQUI_TU_CLAVE';
const ZONA = 'America/Caracas';

const HOJAS = {
  tasas: 'Tasas', clientes: 'Clientes', cobros: 'Cobros', gastos: 'Gastos',
  nomina: 'Nomina', planes: 'Planes', herramientas: 'Herramientas'
};
// Columnas que siempre se guardan como texto (para que Sheets no las convierta en fechas o números)
const TEXTO = ['id', 'fecha', 'mes', 'clienteId', 'desde', 'telefono', 'fechaPago', 'nombre', 'plan', 'pago',
  'estado', 'categoria', 'moneda', 'nota', 'metodo', 'fuente'];

/* ---------------- API web ---------------- */
function doGet(e) { return responder_(e && e.parameter ? e.parameter : {}); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { body = {}; }
  return responder_(body);
}

function responder_(p) {
  if (String(p.key || '') !== clave_()) return json_({ ok: false, error: 'clave' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (p.action) {
      case 'all': return json_({ ok: true, data: leerTodo_() });
      case 'set': guardarFila_(p.col, p.id, p.data || {}); return json_({ ok: true });
      case 'del': borrarFila_(p.col, p.id); return json_({ ok: true });
      case 'config': guardarConfig_(p.data || {}); return json_({ ok: true });
      case 'tasasHoy': lock.releaseLock(); return json_(actualizarTasas());
      default: return json_({ ok: false, error: 'accion' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function clave_() {
  return PropertiesService.getScriptProperties().getProperty('CLAVE') || CLAVE;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- Lectura ---------------- */
function hoja_(nombre) {
  const sh = SpreadsheetApp.getActive().getSheetByName(nombre);
  if (!sh) throw new Error('Falta la pestaña ' + nombre);
  return sh;
}

function filas_(nombre) {
  const sh = hoja_(nombre);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0].map(String);
  return values.slice(1).filter(r => String(r[0]) !== '').map(r => {
    const o = {};
    head.forEach((h, i) => {
      let v = r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, ZONA, h === 'mes' ? 'yyyy-MM' : 'yyyy-MM-dd');
      if (TEXTO.indexOf(h) >= 0 && v !== '' && v !== null) v = String(v);
      o[h] = v;
    });
    return o;
  });
}

function leerConfig_() {
  const cfg = { reparto: {} };
  filas_('Config').forEach(r => {
    const k = String(r.clave), v = Number(r.valor);
    if (k.indexOf('reparto_') === 0) cfg.reparto[k.slice(8)] = v; else cfg[k] = v;
  });
  cfg.planes = filas_(HOJAS.planes).map(p => ({
    id: p.id, nombre: p.nombre, base: Number(p.base), cm: Number(p.cm), pautas: Number(p.pautas), lista: Number(p.lista)
  }));
  cfg.herramientas = filas_(HOJAS.herramientas).map(h => ({ id: h.id, nombre: h.nombre, monto: Number(h.monto) }));
  return cfg;
}

function leerTodo_() {
  return {
    config: leerConfig_(),
    tasas: filas_(HOJAS.tasas),
    clientes: filas_(HOJAS.clientes),
    cobros: filas_(HOJAS.cobros),
    gastos: filas_(HOJAS.gastos),
    nomina: filas_(HOJAS.nomina)
  };
}

/* ---------------- Escritura ---------------- */
function guardarFila_(col, id, data) {
  const nombre = HOJAS[col];
  if (!nombre || !id) throw new Error('Datos incompletos');
  const sh = hoja_(nombre);
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const obj = Object.assign({}, data, { id: id });
  const fila = head.map(h => (obj[h] === undefined || obj[h] === null) ? '' : obj[h]);
  let n = buscar_(sh, id);
  if (n < 0) n = sh.getLastRow() + 1;
  const rango = sh.getRange(n, 1, 1, head.length);
  rango.setNumberFormats([head.map(h => TEXTO.indexOf(h) >= 0 ? '@' : 'General')]);
  rango.setValues([fila]);
}

function borrarFila_(col, id) {
  const sh = hoja_(HOJAS[col]);
  const n = buscar_(sh, id);
  if (n > 1) sh.deleteRow(n);
}

function buscar_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function guardarConfig_(cfg) {
  // Valores sueltos
  const sh = hoja_('Config');
  const valores = Object.assign({}, cfg);
  delete valores.planes; delete valores.herramientas; delete valores.reparto;
  Object.keys(cfg.reparto || {}).forEach(k => valores['reparto_' + k] = cfg.reparto[k]);
  const data = sh.getDataRange().getValues();
  Object.keys(valores).forEach(k => {
    let fila = -1;
    for (let i = 1; i < data.length; i++) if (String(data[i][0]) === k) { fila = i + 1; break; }
    if (fila < 0) { sh.appendRow([k, valores[k], '']); }
    else sh.getRange(fila, 2).setValue(valores[k]);
  });
  // Planes y herramientas se reescriben completos
  reescribir_(HOJAS.planes, ['id', 'nombre', 'base', 'cm', 'pautas', 'lista'], cfg.planes || []);
  reescribir_(HOJAS.herramientas, ['id', 'nombre', 'monto'], cfg.herramientas || []);
}

function reescribir_(nombre, head, rows) {
  const sh = hoja_(nombre);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent();
  if (!rows.length) return;
  sh.getRange(2, 1, rows.length, head.length).setValues(rows.map(r => head.map(h => r[h] === undefined ? '' : r[h])));
}

/* ---------------- Tasas automáticas ---------------- */
function actualizarTasas() {
  const hoy = Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd');
  const sh = hoja_(HOJAS.tasas);
  const existente = filas_(HOJAS.tasas).find(t => t.id === hoy);
  if (existente && existente.fuente === 'manual') return { ok: true, data: existente, nota: 'Ya hay una tasa manual para hoy' };

  const bcv = tasaBCV_();
  const usdt = tasaUSDT_();
  if (!bcv) return { ok: false, error: 'No se pudo leer la tasa BCV' };
  if (!usdt.valor) return { ok: false, error: 'No se pudo leer el precio del USDT' };
  if (usdt.valor < bcv * 0.95 || usdt.valor > bcv * 2.5) return { ok: false, error: 'Tasas fuera de rango, no se guardaron', bcv: bcv, usdt: usdt.valor };

  const fila = { fecha: hoy, bcv: Math.round(bcv * 10000) / 10000, usdt: Math.round(usdt.valor * 100) / 100, fuente: 'auto · ' + usdt.fuente };
  guardarFila_('tasas', hoy, fila);
  return { ok: true, data: Object.assign({ id: hoy }, fila) };
}

function tasaBCV_() {
  try {
    const r = UrlFetchApp.fetch('https://ve.dolarapi.com/v1/dolares/oficial', { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return null;
    const v = Number(JSON.parse(r.getContentText()).promedio);
    return v > 0 ? v : null;
  } catch (e) { return null; }
}

// Precio al que se VENDE USDT en Binance P2P (mediana de los 10 primeros anuncios); si falla, el paralelo de DolarApi.
function tasaUSDT_() {
  try {
    const r = UrlFetchApp.fetch('https://p2p.binance.com/bapi/c2c/v2/friendly/c2c/adv/search', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ asset: 'USDT', fiat: 'VES', tradeType: 'SELL', page: 1, rows: 10, payTypes: [], publisherType: null, merchantCheck: false })
    });
    if (r.getResponseCode() === 200) {
      const precios = (JSON.parse(r.getContentText()).data || [])
        .map(x => Number(x.adv && x.adv.price)).filter(n => n > 0).sort((a, b) => a - b);
      if (precios.length >= 3) return { valor: precios[Math.floor(precios.length / 2)], fuente: 'Binance P2P' };
    }
  } catch (e) {}
  try {
    const r2 = UrlFetchApp.fetch('https://ve.dolarapi.com/v1/dolares/paralelo', { muteHttpExceptions: true });
    if (r2.getResponseCode() === 200) {
      const v = Number(JSON.parse(r2.getContentText()).promedio);
      if (v > 0) return { valor: v, fuente: 'DolarApi paralelo' };
    }
  } catch (e) {}
  return { valor: null };
}

/* ---------------- Instalación ---------------- */
function instalar() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'actualizarTasas')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('actualizarTasas').timeBased().everyDays(1).atHour(9).inTimezone(ZONA).create();
  const r = actualizarTasas();
  Logger.log('Listo. Tasas se actualizarán cada día a las 9 am (Caracas). Prueba de hoy: ' + JSON.stringify(r));
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Sistema')
    .addItem('Actualizar tasas ahora', 'actualizarTasas')
    .addToUi();
}
