/**
 * Backend da Roleta Di Lione — Google Apps Script
 *
 * Como publicar:
 *  1. Crie uma Planilha Google → Extensões → Apps Script → cole este arquivo.
 *  2. Troque CLIENT_ID pelo mesmo ID usado no index.html.
 *  3. Implantar → Nova implantação → Tipo "App da Web"
 *       Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
 *  4. Copie a URL /exec e cole em SURVEY_ENDPOINT no index.html.
 *
 * O que ele garante:
 *  - O token do Google é verificado no servidor (não dá para falsificar o e-mail).
 *  - Uma pesquisa e um giro por conta Google.
 */

const CLIENT_ID = '7212727243-n2uqn1dds2cq02vql1fdtn290er0s1bj.apps.googleusercontent.com';
const SHEET_NAME = 'Pesquisas';
const HEADERS = ['Data', 'E-mail', 'Nome', 'Nota', 'Comentário', 'Prêmio', 'Código', 'Data do giro'];
const COL = { email: 2, prize: 6, code: 7, spinAt: 8 };
const PRIZES = ['SOB', 'BEB', 'D10', 'FAT'];

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, error: 'bad_request' });
  }

  const user = verifyToken(body.credential);
  if (!user) return json({ ok: false, error: 'auth' });

  const email = String(user.email).toLowerCase();
  const sheet = getSheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const rowIdx = findRow(sheet, email);
    const row = rowIdx ? sheet.getRange(rowIdx, 1, 1, HEADERS.length).getValues()[0] : null;

    if (body.action === 'status') {
      return json({
        ok: true,
        surveyDone: !!row,
        prizeCode: row && row[COL.prize - 1] || null,
        code: row && row[COL.code - 1] || null
      });
    }

    if (body.action === 'survey') {
      if (row) return json({ ok: false, error: 'duplicate' });
      const d = body.data || {};
      sheet.appendRow([
        new Date(), email, user.name || '',
        stars(d.nota),
        String(d.comentario || '').slice(0, 500),
        '', '', ''
      ]);
      return json({ ok: true });
    }

    if (body.action === 'spin') {
      if (!row) return json({ ok: false, error: 'no_survey' });
      if (row[COL.code - 1]) {
        return json({ ok: true, existing: true, prizeCode: row[COL.prize - 1], code: row[COL.code - 1] });
      }
      const d = body.data || {};
      if (PRIZES.indexOf(d.prizeCode) < 0) return json({ ok: false, error: 'bad_prize' });
      sheet.getRange(rowIdx, COL.prize, 1, 3).setValues([[d.prizeCode, String(d.code || '').slice(0, 40), new Date()]]);
      return json({ ok: true });
    }

    return json({ ok: false, error: 'unknown_action' });
  } finally {
    lock.releaseLock();
  }
}

function verifyToken(token) {
  if (!token) return null;
  const res = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token),
    { muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) return null;
  const p = JSON.parse(res.getContentText());
  if (p.aud !== CLIENT_ID) return null;
  if (p.iss !== 'accounts.google.com' && p.iss !== 'https://accounts.google.com') return null;
  if (String(p.email_verified) !== 'true') return null;
  return p;
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function findRow(sheet, email) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const emails = sheet.getRange(2, COL.email, last - 1, 1).getValues();
  for (let i = 0; i < emails.length; i++) {
    if (String(emails[i][0]).toLowerCase() === email) return i + 2;
  }
  return 0;
}

function stars(v) {
  const n = Number(v);
  return n >= 1 && n <= 5 ? n : '';
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
