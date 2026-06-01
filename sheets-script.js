const SHEET_NAME = 'מופעים';

const STATUS_LABELS = {
  done: 'התקיים',
  approved: 'מאושר',
  pending: 'ממתין אישור',
  open: 'לא סגור',
  signature: 'ממתין חתימה',
  postponed: 'נדחה',
  moved: 'הוזז',
  cancelled: 'בוטל',
  '': 'ללא סטטוס'
};

function doPost(e) {
  try {
    let shows;
    // ניסיון 1: גוף טקסט ישיר (text/plain)
    if (e.postData && e.postData.contents) {
      shows = JSON.parse(e.postData.contents);
    }
    // ניסיון 2: פרמטר form
    else if (e.parameter && e.parameter.data) {
      shows = JSON.parse(e.parameter.data);
    } else {
      return buildResponse({ success: false, error: 'no data' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
    syncSheet(sheet, shows);
    return buildResponse({ success: true, count: shows.length });
  } catch (err) {
    return buildResponse({ success: false, error: err.message });
  }
}

function doGet(e) {
  if (e.parameter && e.parameter.data) {
    try {
      const shows = JSON.parse(e.parameter.data);
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
      syncSheet(sheet, shows);
      return buildResponse({ success: true, count: shows.length });
    } catch (err) {
      return buildResponse({ success: false, error: err.message });
    }
  }
  return buildResponse({ status: 'ok' });
}

function buildResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function syncSheet(sheet, shows) {
  sheet.clearContents();
  sheet.setRightToLeft(true);

  const headers = [
    'תאריך', 'שעה', 'מיקום', 'שם המופע', 'שכבת גיל',
    'מי מביא', 'איש קשר', 'טלפון', 'מס׳ סבבים',
    'בוקר/ערב', 'סטטוס', 'צוות', 'הערות'
  ];

  sheet.appendRow(headers);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1E3A5F');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  const rows = shows.map(s => [
    s.date ? s.date.split('-').reverse().join('/') : '',
    s.time || '',
    s.location || '',
    s.show || '',
    s.type || '',
    s.presenter || '',
    s.contact || '',
    s.phone || '',
    s.count || 1,
    s.morning ? 'בוקר' : 'ערב',
    STATUS_LABELS[s.status] || '',
    s.team ? s.team.join(', ') : '',
    s.notes || ''
  ]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.autoResizeColumns(1, headers.length);
}
