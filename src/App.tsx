const SHEET_ID = '1iXtou2JY2z8-w-sQAsXiLP84JznEUQ2mOEfwX9egOvo';

function doGet(e) { return handleGet(e); }
function doPost(e) { return handlePost(e); }

function handleGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID);
    const action = e.parameter.action;

    if (action === 'fetch') {
      const log = sheet.getSheetByName('Log');
      const rows = log.getDataRange().getValues();
      if (rows.length <= 1) return buildResponse([]);
      const headers = rows[0];
      const data = rows.slice(1)
        .filter(r => r[0] !== '' && r[0] !== null && r[0] !== undefined)
        .map(row => {
          const obj = {};
          headers.forEach((h, i) => obj[String(h).trim()] = row[i]);
          obj.rate        = Number(obj.rate)        || 0;
          obj.ts          = Number(obj.ts)           || 0;
          obj.index_score = Number(obj.index_score)  || 0;
          obj.recurring   = String(obj.recurring||"").toLowerCase() === "true";
          obj.notes       = obj.notes || "";
          return obj;
        });
      return buildResponse(data);
    }

    if (action === 'deliverables') {
      const ws = sheet.getSheetByName('Deliverables');
      if (!ws) return buildResponse([]);
      const rows = ws.getDataRange().getValues();
      if (rows.length <= 1) return buildResponse([]);
      const headers = rows[0];
      const data = rows.slice(1)
        .filter(r => r[0] !== '' && r[0] !== null && r[0] !== undefined)
        .map(row => {
          const obj = {};
          headers.forEach((h, i) => obj[String(h).trim()] = row[i]);
          obj.rate        = Number(obj.rate)        || 0;
          obj.index_score = Number(obj.index_score) || 1;
          return obj;
        });
      return buildResponse(data);
    }

    if (action === 'clubs') {
      const ws = sheet.getSheetByName('Clubs');
      if (!ws) return buildResponse([]);
      const rows = ws.getDataRange().getValues();
      if (rows.length <= 1) return buildResponse([]);
      const headers = rows[0];
      const data = rows.slice(1)
        .filter(r => r[0] !== '' && r[0] !== null && r[0] !== undefined)
        .map(row => {
          const obj = {};
          headers.forEach((h, i) => obj[String(h).trim()] = row[i]);
          return obj;
        });
      return buildResponse(data);
    }

    return buildResponse({ error: 'Unknown action: ' + action });
  } catch(err) {
    return buildResponse({ error: err.toString() });
  }
}

function handlePost(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID);
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'insert') {
      const log = sheet.getSheetByName('Log');
      const row = data.row;
      const lastRow = log.getLastRow();
      // Columns: id, ts, staff, dept, name, rate, type, cat, club, league, recurring, notes
      log.getRange(lastRow + 1, 1, 1, 12).setValues([[
        row.id, row.ts, row.staff, row.dept, row.name,
        row.rate, row.type, row.cat, row.club, row.league,
        row.recurring || false, row.notes || ""
      ]]);
      return buildResponse({ success: true });
    }

    if (action === 'insertBatch') {
      const log = sheet.getSheetByName('Log');
      const rows = data.rows;
      const lastRow = log.getLastRow();
      const values = rows.map(row => [
        row.id, row.ts, row.staff, row.dept, row.name,
        row.rate, row.type, row.cat, row.club, row.league,
        row.recurring || false, row.notes || ""
      ]);
      log.getRange(lastRow + 1, 1, values.length, 12).setValues(values);
      return buildResponse({ success: true, count: rows.length });
    }

    if (action === 'delete') {
      const log = sheet.getSheetByName('Log');
      const ids = data.ids.map(String);
      const rows = log.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) {
        if (ids.includes(String(rows[i][0]))) log.deleteRow(i + 1);
      }
      return buildResponse({ success: true });
    }

    return buildResponse({ error: 'Unknown action: ' + action });
  } catch(err) {
    return buildResponse({ error: err.toString() });
  }
}

function buildResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
