/**
 * MAKEN METAL · Backend Ventas v2 (Google Apps Script)
 * Modelo de dos niveles: Cuentas (clientes/prospectos) y Oportunidades (proyectos).
 * Las pestañas y columnas se crean solas.
 *
 * MIGRACIÓN (una sola vez): si vienes de la versión anterior, ejecuta la función
 * migrarDesdePipeline() desde el editor (botón Ejecutar) para pasar tus datos.
 *
 * OJO al agregar columnas: esta lista manda sobre la hoja. Si agregas una columna
 * en la hoja pero no aquí, el script la ignora en silencio —acepta el dato y lo
 * tira—. Y el orden de esta lista debe coincidir con el de la hoja: agrega
 * siempre AL FINAL, porque sheetFor_ reescribe los encabezados si no cuadran.
 */

const TABS = {
  Cuentas: ['ID','Empresa','Tipo','Sector','Vendedor','Estado','Contactos',
            'ProximaAccion','FechaProximaAccion','FechaUltimoContacto','Notas',
            'FechaCreado','FechaActualizado','Origen'],
  Oportunidades: ['ID','CuentaID','Empresa','Contacto','Descripcion','Etapa','Monto',
                  'Periodicidad','Vendedor','ProximaAccion','FechaProximaAccion',
                  'FechaUltimoContacto','MotivoPerdido','Notas','FechaCreado','FechaActualizado']
};

function sheetFor_(tab) {
  const headers = TABS[tab];
  if (!headers) throw new Error('tab inválido: ' + tab);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(tab);
  if (!sh) { sh = ss.insertSheet(tab); }
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#23272f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  } else {
    if (sh.getMaxColumns() < headers.length) {
      sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
    }
    const hdr = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    let needs = false;
    for (let i = 0; i < headers.length; i++) { if (hdr[i] !== headers[i]) { needs = true; break; } }
    if (needs) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setBackground('#23272f').setFontColor('#ffffff');
    }
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function todayStr_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function nowStr_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }

function readTab_(tab) {
  const sh = sheetFor_(tab), headers = TABS[tab];
  if (sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  values.shift();
  return values.filter(r => String(r[0]).trim() !== '').map(r => {
    const o = {};
    headers.forEach((h, i) => {
      let v = r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      o[h] = (v === undefined || v === null) ? '' : v;
    });
    return o;
  });
}

function findRow_(sh, id) {
  const ids = sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) { if (String(ids[i][0]) === String(id)) return i + 2; }
  return -1;
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'all';
    if (action === 'ping') return json_({ ok: true, pong: true });
    if (action === 'all') { autoMigrate_(); return json_({ ok: true, cuentas: readTab_('Cuentas'), oportunidades: readTab_('Oportunidades') }); }
    if (action === 'list') return json_({ ok: true, rows: readTab_(e.parameter.tab || 'Cuentas') });
    return json_({ ok: false, error: 'accion desconocida: ' + action });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = body.action;
    const tab = body.tab;
    const headers = TABS[tab];
    if (!headers) return json_({ ok: false, error: 'tab inválido' });
    const sh = sheetFor_(tab);

    if (action === 'add') {
      const rec = body.rec || {};
      rec.ID = (tab === 'Cuentas' ? 'C' : 'OP') + Date.now() + Math.floor(Math.random() * 1000);
      rec.FechaCreado = todayStr_();
      rec.FechaActualizado = nowStr_();
      if (tab === 'Oportunidades' && !rec.Etapa) rec.Etapa = 'RFQ / Cotizando';
      if (tab === 'Cuentas' && !rec.Estado) rec.Estado = 'En acercamiento';
      if (rec.Notas) rec.FechaUltimoContacto = todayStr_();
      sh.appendRow(headers.map(h => rec[h] !== undefined ? rec[h] : ''));
      return json_({ ok: true, id: rec.ID });
    }

    if (action === 'update') {
      const rowIdx = findRow_(sh, body.id);
      if (rowIdx < 0) return json_({ ok: false, error: 'no encontrado: ' + body.id });
      const current = sh.getRange(rowIdx, 1, 1, headers.length).getValues()[0];
      const rec = body.rec || {};
      headers.forEach((h, i) => {
        if (rec[h] !== undefined) {
          let v = rec[h];
          if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          current[i] = v;
        }
      });
      current[headers.indexOf('FechaActualizado')] = nowStr_();
      sh.getRange(rowIdx, 1, 1, headers.length).setValues([current]);
      return json_({ ok: true });
    }

    if (action === 'delete') {
      const rowIdx = findRow_(sh, body.id);
      if (rowIdx < 0) return json_({ ok: false, error: 'no encontrado' });
      sh.deleteRow(rowIdx);
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'accion desconocida: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally { try { lock.releaseLock(); } catch (e2) {} }
}

/**
 * MIGRACIÓN AUTOMÁTICA desde la pestaña vieja "Pipeline".
 * Se dispara sola la primera vez que se abre la app (en doGet 'all').
 * Es idempotente: no duplica cuentas ni oportunidades aunque se ejecute varias veces.
 * No borra "Pipeline" (queda de respaldo).
 */
function autoMigrate_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('migrated') === '1') return;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (props.getProperty('migrated') === '1') return;
    doMigrate_();
    props.setProperty('migrated', '1');
  } catch (e) {
    // si algo falla, no marcamos como migrado para reintentar luego
  } finally { try { lock.releaseLock(); } catch (e2) {} }
}

// Ejecútala manualmente desde el editor si quieres forzar la migración.
function migrarDesdePipeline() {
  doMigrate_();
  Logger.log('Migración lista. Cuentas: ' + readTab_('Cuentas').length + ' · Oportunidades: ' + readTab_('Oportunidades').length);
}

function doMigrate_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName('Pipeline');
  if (!old || old.getLastRow() < 2) return;
  const vals = old.getDataRange().getValues();
  const head = vals.shift();
  const col = {}; head.forEach((h, i) => col[h] = i);
  const get = (r, h) => (col[h] !== undefined && r[col[h]] !== undefined && r[col[h]] !== null) ? r[col[h]] : '';
  const fmt = (v) => (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : v;

  const shC = sheetFor_('Cuentas');
  const shO = sheetFor_('Oportunidades');
  const OPEN_OPP = ['RFQ / Cotizando', 'Cotización enviada', 'Ganado', 'Perdido'];

  // Sembrar con lo que YA existe para no duplicar
  const cuentaByName = {};
  readTab_('Cuentas').forEach(c => { cuentaByName[String(c.Empresa).toLowerCase().trim()] = { id: c.ID }; });
  const oppKey = {};
  readTab_('Oportunidades').forEach(o => { oppKey[(String(o.Empresa) + '|' + String(o.Descripcion)).toLowerCase()] = true; });

  vals.forEach(r => {
    const empresa = String(get(r, 'Empresa')).trim();
    if (!empresa) return;
    const key = empresa.toLowerCase();
    const tipo = get(r, 'Tipo') || 'Prospecto nuevo';

    if (!cuentaByName[key]) {
      const cid = 'C' + Date.now() + Math.floor(Math.random() * 100000);
      const contacto = String(get(r, 'Contacto')).trim();
      const contactos = contacto ? [{ n: contacto, p: get(r, 'Puesto'), t: get(r, 'Telefono'), e: get(r, 'Email') }] : [];
      const estado = (tipo === 'Cliente actual') ? 'Cliente activo' : 'En acercamiento';
      shC.appendRow(TABS.Cuentas.map(h => {
        switch (h) {
          case 'ID': return cid;
          case 'Empresa': return empresa;
          case 'Tipo': return tipo;
          case 'Sector': return get(r, 'Sector');
          case 'Vendedor': return get(r, 'Vendedor');
          case 'Estado': return estado;
          case 'Contactos': return JSON.stringify(contactos);
          case 'FechaCreado': return todayStr_();
          case 'FechaActualizado': return nowStr_();
          default: return '';
        }
      }));
      cuentaByName[key] = { id: cid };
    }
    const cuenta = cuentaByName[key];

    const etapa = String(get(r, 'Etapa')).trim();
    const desc = String(get(r, 'Descripcion')).trim();
    const ok2 = (empresa + '|' + desc).toLowerCase();
    if (OPEN_OPP.indexOf(etapa) >= 0 && !oppKey[ok2]) {
      shO.appendRow(TABS.Oportunidades.map(h => {
        switch (h) {
          case 'ID': return 'OP' + Date.now() + Math.floor(Math.random() * 100000);
          case 'CuentaID': return cuenta.id;
          case 'Empresa': return empresa;
          case 'Contacto': return get(r, 'Contacto');
          case 'Descripcion': return desc;
          case 'Etapa': return etapa;
          case 'Monto': return get(r, 'Monto');
          case 'Periodicidad': return get(r, 'Periodicidad') || 'Único';
          case 'Vendedor': return get(r, 'Vendedor');
          case 'ProximaAccion': return get(r, 'ProximaAccion');
          case 'FechaProximaAccion': return fmt(get(r, 'FechaProximaAccion'));
          case 'FechaUltimoContacto': return fmt(get(r, 'FechaUltimoContacto'));
          case 'MotivoPerdido': return get(r, 'MotivoPerdido');
          case 'Notas': return get(r, 'Notas');
          case 'FechaCreado': return fmt(get(r, 'FechaCreado')) || todayStr_();
          case 'FechaActualizado': return nowStr_();
          default: return '';
        }
      }));
      oppKey[ok2] = true;
    }
  });
}
