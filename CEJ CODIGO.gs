// ===== APP CEJ - Backend Apps Script (Cobranza Extrajudicial) =====
// Planilla: SEGUIMIENTO CEJ 2026  |  Publicar como Web App (Ejecutar como: yo / Acceso: cualquiera)
var SHEET_ID = '1lvZxUv-El8-R86HYMaprp9Lg5Uom5guPN1pn3w4N-eE';
var SHEET_NAME = 'SEGUIMIENTO';
var NCOLS = 18;
var DATE_COLS = [8, 9, 10, 11, 12, 13, 16]; // indices 0-based de columnas de fecha
var EMAIL_INSTRUCCION = ''; // opcional: correo(s) separados por coma para avisar nuevas instrucciones. Vacio = desactivado
var TZ = 'America/Santiago';

// EJECUTAR UNA VEZ (editando las claves) antes de publicar. Varias claves separadas por coma.
function configurarClaves() {
  PropertiesService.getScriptProperties().setProperties({
    CLAVES_EDIT: 'Raot7012',
    CLAVES_LECTURA: 'DSTEMUCO'
  });
}

function sheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}
function rol_(clave) {
  clave = String(clave || '').trim();
  if (!clave) return null;
  var pr = PropertiesService.getScriptProperties();
  var ed = String(pr.getProperty('CLAVES_EDIT') || '').split(',').map(function (s) { return s.trim(); });
  var le = String(pr.getProperty('CLAVES_LECTURA') || '').split(',').map(function (s) { return s.trim(); });
  if (ed.indexOf(clave) >= 0) return 'edit';
  if (le.indexOf(clave) >= 0) return 'read';
  return null;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function fmt_(v, isDate) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  if (isDate && v !== '' && v !== null) {
    var d = new Date(v);
    if (!isNaN(d)) return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  }
  return v === null ? '' : String(v);
}

function doGet(e) {
  try {
    var p = e.parameter || {};
    var rol = rol_(p.clave);
    if (!rol) return out_({ ok: false, error: 'Clave inválida' });
    if (p.action === 'login') return out_({ ok: true, rol: rol });
    var sh = sheet_();
    var last = sh.getLastRow();
    var rows = [];
    if (last >= 2) {
      var data = sh.getRange(2, 1, last - 1, NCOLS).getValues();
      data.forEach(function (r, i) {
        if (String(r[0]).trim() === '' && String(r[1]).trim() === '') return;
        rows.push({ row: i + 2, v: r.map(function (c, j) { return fmt_(c, DATE_COLS.indexOf(j) >= 0); }) });
      });
    }
    return out_({ ok: true, rol: rol, rows: rows });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function toDate_(s) {
  s = String(s || '').trim();
  if (!s) return '';
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], 12);
  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], 12);
  return s;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var p = e.parameter || {};
    if (rol_(p.clave) !== 'edit') return out_({ ok: false, error: 'Sin permiso de edición' });
    var sh = sheet_();
    var fila = parseInt(p.row, 10);
    var esEdicion = !isNaN(fila) && fila >= 2;
    if (!esEdicion) fila = Math.max(sh.getLastRow(), 1) + 1;

    var oldObs = '', oldInstr = '';
    if (esEdicion) {
      oldObs = String(sh.getRange(fila, 15).getValue()).trim();
      oldInstr = String(sh.getRange(fila, 16).getValue()).trim();
    }
    var newObs = String(p.obs || '').trim();
    var newInstr = String(p.instr || '').trim();
    var hoy = new Date();

    var deudaTxt = String(p.deuda || '').replace(/[^0-9]/g, '');
    var vals = [
      p.deudor, p.rut, p.cuenta, p.nid,
      deudaTxt === '' ? '' : Number(deudaTxt), p.telefono, p.correo, p.direccion,
      toDate_(p.f_asig), toDate_(p.f_mail), toDate_(p.f_tel), toDate_(p.f_wsp), toDate_(p.f_carta),
      toDate_(p.f_obs), p.obs, p.instr, toDate_(p.f_instr), p.encargado
    ].map(function (v) { return v === undefined || v === null ? '' : v; });

    if (newObs !== oldObs) vals[13] = newObs === '' ? '' : hoy;       // FECHA ULTIMA OBS
    if (newInstr !== oldInstr) vals[16] = newInstr === '' ? '' : hoy; // FECHA INSTRUCCION

    sh.getRange(fila, 1, 1, NCOLS).setValues([vals]);
    DATE_COLS.forEach(function (c) { sh.getRange(fila, c + 1).setNumberFormat('dd-MM-yyyy'); });
    lock.releaseLock();

    if (EMAIL_INSTRUCCION && newInstr !== '' && newInstr !== oldInstr) {
      try {
        MailApp.sendEmail(EMAIL_INSTRUCCION, 'Nueva instrucción CEJ: ' + p.deudor,
          'Deudor: ' + p.deudor + '\nRUT: ' + p.rut + '\nEncargado: ' + (p.encargado || '') +
          '\n\nInstrucción:\n' + newInstr);
      } catch (errMail) { Logger.log(errMail); }
    }
    return out_({ ok: true, fila: fila });
  } catch (err) {
    try { lock.releaseLock(); } catch (x) {}
    return out_({ ok: false, error: String(err) });
  }
}
