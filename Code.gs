const SHEET_ID = '1uY2mSkIoZtmI7nnu4QUDfDdnK_ZFhG5GOTPnJSMMfUY';
const SHEET_NAME = 'lefoo_personal';

function doGet(e) {
  const action = e.parameter.action || 'read';

  if (action === 'read') {
    return handleRead();
  } else if (action === 'toggle') {
    return handleToggle(e.parameter.name);
  }

  return jsonResponse({ error: 'Unknown action' });
}

function handleRead() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const nameIdx    = headers.indexOf('Name');
  const roleIdx    = headers.indexOf('Zuständigkeit');
  const presentIdx = headers.indexOf('Anwesend');
  const photoIdx   = headers.indexOf('Foto_URL');

  const staff = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[nameIdx]) continue;
    staff.push({
      name:    row[nameIdx],
      role:    row[roleIdx],
      present: row[presentIdx] === true || String(row[presentIdx]).toUpperCase() === 'TRUE',
      photo:   row[photoIdx] || ''
    });
  }

  return jsonResponse({ staff: staff, updated: new Date().toISOString() });
}

function handleToggle(name) {
  if (!name) return jsonResponse({ error: 'No name provided' });

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const nameIdx    = headers.indexOf('Name');
  const presentIdx = headers.indexOf('Anwesend');

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameIdx] === name) {
      const currentValue = data[i][presentIdx];
      const isPresent = currentValue === true || String(currentValue).toUpperCase() === 'TRUE';
      const newValue = isPresent ? 'FALSE' : 'TRUE';
      sheet.getRange(i + 1, presentIdx + 1).setValue(newValue);
      return jsonResponse({ success: true, name: name, present: !isPresent });
    }
  }

  return jsonResponse({ error: 'Person not found' });
}

function jsonResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
