/**
 * ============================================================
 *  FICHA SOCIAL UNCP — SCRIPT v8 (COPIAR TODO ESTE ARCHIVO)
 * ============================================================
 *  Hoja:  1obx8kVIXXxk2P65LVXxHju3I1iUf3SsPgozlxKmZ5ek
 *  Drive: 1dqy2HuWf6IYVGasuyrtff2id7Jf3aRXv
 *
 *  CÓMO INSTALAR (si no, la fila NO se guarda):
 *  1. Abra la hoja → Extensiones → Apps Script
 *  2. Borre TODO el código → pegue ESTE archivo → Ctrl+S
 *  3. Implementar → Nueva implementación → Aplicación web
 *       Ejecutar como: Yo
 *       Quién tiene acceso: Cualquier persona
 *  4. Implementar → Permitir permisos
 *  5. COPIE la URL que termina en /exec
 *  6. En la entrevista: Admin → URL / config → pegar URL → Guardar → Diagnosticar
 *  7. La URL /exec DEBE decir: "version":"2026-03-pdf-v10-excel"
 * ============================================================
 */

var SCRIPT_VERSION = '2026-03-pdf-v10-excel';
var SPREADSHEET_ID = '1obx8kVIXXxk2P65LVXxHju3I1iUf3SsPgozlxKmZ5ek';
var PDF_FOLDER_ID = '1dqy2HuWf6IYVGasuyrtff2id7Jf3aRXv';
var SHEET_TAB = 'Fichas'; // pestaña limpia (se crea sola)

function doGet(e) {
  var o = {
    ok: true,
    version: SCRIPT_VERSION,
    pdfReady: true,
    acceptsPdfBase64: true,
    acceptsExcelBase64: true,
    message: 'OK v10 — pestaña Fichas + PDF + Excel + Word en Drive',
    spreadsheetId: SPREADSHEET_ID,
    pdfFolderId: PDF_FOLDER_ID,
    sheetTab: SHEET_TAB,
    sheetOk: false,
    folderAccessible: false,
    sheetError: '',
    folderError: ''
  };
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    getFichasSheet_(ss);
    o.sheetOk = true;
  } catch (err) {
    o.sheetError = String(err.message || err);
  }
  try {
    var f = DriveApp.getFolderById(PDF_FOLDER_ID);
    o.folderAccessible = true;
    o.pdfFolderUrl = f.getUrl();
  } catch (err2) {
    o.folderError = String(err2.message || err2);
  }
  return out_(o);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return out_({ ok: false, version: SCRIPT_VERSION, error: 'Sin datos' });
    }
    var p = JSON.parse(e.postData.contents);
    var action = p.action || 'append';

    if (action === 'ping') {
      return out_({ ok: true, version: SCRIPT_VERSION, message: 'pong', pdfReady: true });
    }

    if (action === 'save_pdf_only' || action === 'save_files') {
      var pdf = null;
      var word = null;
      var excel = null;
      var pdfErr = '';
      var wordErr = '';
      var excelErr = '';
      try {
        if (p.pdfBase64 || !p.skipPdf) pdf = savePdf_(p);
      } catch (eP) {
        pdfErr = String(eP.message || eP);
      }
      try {
        if (p.excelBase64 && String(p.excelBase64).length > 100) {
          excel = saveExcel_(p);
        }
      } catch (eX) {
        excelErr = String(eX.message || eX);
      }
      try {
        if (p.wordBase64 && String(p.wordBase64).length > 100) {
          word = saveWord_(p);
        }
      } catch (eW) {
        wordErr = String(eW.message || eW);
      }
      if (!pdf && !word && !excel) {
        return out_({
          ok: false,
          version: SCRIPT_VERSION,
          error: 'No se guardó PDF, Excel ni Word. ' + pdfErr + ' ' + excelErr + ' ' + wordErr
        });
      }
      return out_({
        ok: true,
        version: SCRIPT_VERSION,
        message: 'Archivos guardados en Drive',
        pdfUrl: pdf ? pdf.url : '',
        pdfId: pdf ? pdf.id : '',
        pdfName: pdf ? pdf.name : '',
        pdfFolderUrl: pdf ? pdf.folderUrl : (excel ? excel.folderUrl : (word ? word.folderUrl : '')),
        pdfOrigen: pdf ? pdf.origen : '',
        pdfError: pdfErr,
        excelUrl: excel ? excel.url : '',
        excelId: excel ? excel.id : '',
        excelName: excel ? excel.name : '',
        excelError: excelErr,
        wordUrl: word ? word.url : '',
        wordId: word ? word.id : '',
        wordName: word ? word.name : '',
        wordError: wordErr
      });
    }

    // Fila en hoja (y opcionalmente PDF si viene base64)
    if (action === 'append' || action === 'append_row_only') {
      return appendRow_(p);
    }

    return out_({ ok: false, version: SCRIPT_VERSION, error: 'Accion desconocida: ' + action });
  } catch (err) {
    return out_({ ok: false, version: SCRIPT_VERSION, error: String(err.message || err) });
  }
}

function appendRow_(p) {
  var headers = (p.headers || []).slice();
  var values = (p.values || []).slice();
  if (!headers.length || headers.length !== values.length) {
    return out_({
      ok: false,
      version: SCRIPT_VERSION,
      error: 'headers/values invalidos'
    });
  }

  // Columnas de enlace al PDF y Word
  function ensureCol(name, val) {
    if (headers.indexOf(name) < 0) {
      headers.push(name);
      values.push(val || '');
    } else if (val) {
      values[headers.indexOf(name)] = val;
    }
  }
  ensureCol('Link PDF', p.pdfUrlPrevio || '');
  ensureCol('Nombre PDF', p.pdfNamePrevio || '');
  ensureCol('Link Excel', p.excelUrlPrevio || '');
  ensureCol('Nombre Excel', p.excelNamePrevio || '');
  ensureCol('Link Word', p.wordUrlPrevio || '');
  ensureCol('Nombre Word', p.wordNamePrevio || '');

  var pdfInfo = null;
  if (p.pdfBase64 && String(p.pdfBase64).length > 200) {
    try {
      pdfInfo = savePdf_(p);
      values[headers.indexOf('Link PDF')] = pdfInfo.url;
      values[headers.indexOf('Nombre PDF')] = pdfInfo.name;
    } catch (e1) { /* la fila se guarda igual */ }
  }

  var excelInfo = null;
  if (p.excelBase64 && String(p.excelBase64).length > 200) {
    try {
      excelInfo = saveExcel_(p);
      values[headers.indexOf('Link Excel')] = excelInfo.url;
      values[headers.indexOf('Nombre Excel')] = excelInfo.name;
    } catch (eX) { /* la fila se guarda igual */ }
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getFichasSheet_(ss);
  var row = writeCells_(sheet, headers, values);

  return out_({
    ok: true,
    version: SCRIPT_VERSION,
    message: 'Fila guardada en pestaña Fichas',
    sheet: sheet.getName(),
    spreadsheetId: ss.getId(),
    row: row,
    pdfUrl: pdfInfo ? pdfInfo.url : (p.pdfUrlPrevio || ''),
    pdfName: pdfInfo ? pdfInfo.name : (p.pdfNamePrevio || ''),
    pdfId: pdfInfo ? pdfInfo.id : '',
    pdfFolderUrl: pdfInfo ? pdfInfo.folderUrl : '',
    pdfOrigen: pdfInfo ? pdfInfo.origen : '',
    excelUrl: excelInfo ? excelInfo.url : (p.excelUrlPrevio || ''),
    excelName: excelInfo ? excelInfo.name : (p.excelNamePrevio || ''),
    wordUrl: p.wordUrlPrevio || '',
    wordName: p.wordNamePrevio || ''
  });
}

/** Guarda Excel (.xlsx) — misma ficha que el PDF */
function saveExcel_(p) {
  p = p || {};
  var b64 = String(p.excelBase64 || '')
    .replace(/^data:application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet;base64,/i, '')
    .replace(/^data:application\/octet-stream;base64,/i, '')
    .replace(/\s/g, '');
  if (!b64 || b64.length < 50) throw new Error('excelBase64 vacío');

  var folder = DriveApp.getFolderById(PDF_FOLDER_ID);
  var name = safe_(p.excelName || p.pdfName || ('Ficha_' + Date.now()));
  name = name.replace(/\.pdf$/i, '').replace(/\.docx?$/i, '');
  if (!/\.xlsx$/i.test(name)) name += '.xlsx';

  var mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var file = folder.createFile(
    Utilities.newBlob(Utilities.base64Decode(b64), mime, name)
  );
  return {
    url: file.getUrl(),
    id: file.getId(),
    name: file.getName(),
    folderUrl: folder.getUrl()
  };
}

/** Guarda archivo Word (.doc) en la carpeta de Drive */
function saveWord_(p) {
  p = p || {};
  var b64 = String(p.wordBase64 || '')
    .replace(/^data:application\/msword;base64,/i, '')
    .replace(/^data:application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.document;base64,/i, '')
    .replace(/\s/g, '');
  if (!b64 || b64.length < 50) throw new Error('wordBase64 vacío');

  var folder = DriveApp.getFolderById(PDF_FOLDER_ID);
  var name = safe_(p.wordName || p.pdfName || ('Ficha_' + Date.now()));
  name = name.replace(/\.pdf$/i, '');
  if (!/\.docx?$/i.test(name)) name += '.doc';

  var mime = /\.docx$/i.test(name)
    ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    : 'application/msword';

  var file = folder.createFile(
    Utilities.newBlob(Utilities.base64Decode(b64), mime, name)
  );
  return {
    url: file.getUrl(),
    id: file.getId(),
    name: file.getName(),
    folderUrl: folder.getUrl()
  };
}

function getFichasSheet_(ss) {
  var sh = ss.getSheetByName(SHEET_TAB);
  if (!sh) sh = ss.insertSheet(SHEET_TAB);
  return sh;
}

/** SOLO setValue por celda — nunca setValues multi-columna */
function writeCells_(sheet, headers, values) {
  var n = headers.length;
  while (values.length < n) values.push('');

  for (var c = 0; c < n; c++) {
    sheet.getRange(1, c + 1).setValue(String(headers[c] == null ? '' : headers[c]));
  }
  try {
    sheet.getRange(1, 1, 1, n).setFontWeight('bold');
    sheet.getRange(1, 1, 1, n).setBackground('#1a365d');
    sheet.getRange(1, 1, 1, n).setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  } catch (e) {}

  var row = Math.max(sheet.getLastRow() + 1, 2);
  for (var i = 0; i < n; i++) {
    var v = values[i];
    if (v == null) v = '';
    if (Object.prototype.toString.call(v) === '[object Array]') v = v.join(', ');
    if (typeof v === 'object') {
      try { v = JSON.stringify(v); } catch (e2) { v = String(v); }
    }
    v = String(v);
    if (v.length > 49000) v = v.substring(0, 49000);
    sheet.getRange(row, i + 1).setValue(v);
  }
  return row;
}

function savePdf_(p) {
  p = p || {};
  var folder = DriveApp.getFolderById(PDF_FOLDER_ID);
  var b64 = p.pdfBase64 ? String(p.pdfBase64) : '';

  if (b64.length > 200) {
    b64 = b64.replace(/^data:application\/pdf;base64,/i, '').replace(/\s/g, '');
    var name = safe_(p.pdfName || ('Ficha_' + Date.now()));
    if (!/\.pdf$/i.test(name)) name += '.pdf';
    var file = folder.createFile(
      Utilities.newBlob(Utilities.base64Decode(b64), MimeType.PDF, name)
    );
    return {
      url: file.getUrl(),
      id: file.getId(),
      name: file.getName(),
      folderUrl: folder.getUrl(),
      origen: 'ficha_impresa'
    };
  }

  // Respaldo texto
  var meta = p.meta || {};
  var docName = safe_((p.pdfName || meta.nombres || 'Ficha')).replace(/\.pdf$/i, '');
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();
  body.appendParagraph('FICHA SOCIAL UNCP');
  body.appendParagraph('Socio: ' + (meta.nombres || ''));
  body.appendParagraph('DNI: ' + (meta.dni || ''));
  var hs = p.headers || [];
  var vs = p.values || [];
  for (var i = 0; i < hs.length; i++) {
    if (vs[i] == null || vs[i] === '') continue;
    body.appendParagraph(String(hs[i]) + ': ' + String(vs[i]));
  }
  doc.saveAndClose();
  var docFile = DriveApp.getFileById(doc.getId());
  var pdfFile = folder.createFile(docFile.getAs(MimeType.PDF).setName(docName + '.pdf'));
  try { docFile.setTrashed(true); } catch (e3) {}
  return {
    url: pdfFile.getUrl(),
    id: pdfFile.getId(),
    name: pdfFile.getName(),
    folderUrl: folder.getUrl(),
    origen: 'datos_texto'
  };
}

function safe_(s) {
  return String(s || 'ficha').replace(/[\\/:*?"<>|]+/g, '_').replace(/\s+/g, '_').substring(0, 120);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Ejecutar en el editor para probar la fila sin la entrevista */
function PROBAR_FILA() {
  var r = appendRow_({
    headers: ['Fecha registro', 'Apellidos y nombres', 'DNI', 'Link PDF', 'Nombre PDF'],
    values: [new Date().toISOString(), 'PRUEBA V8', '11111111', 'https://example.com', 'prueba.pdf'],
    meta: { nombres: 'PRUEBA V8', dni: '11111111' }
  });
  Logger.log(r.getContent());
}
