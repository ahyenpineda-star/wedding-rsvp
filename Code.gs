function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('RSVPs');
    if (!sheet) {
      sheet = ss.insertSheet('RSVPs');
      sheet.appendRow(['Timestamp', 'Name', 'Attending', 'Message']);
    }

    let data;
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      data = e.parameter;
    } else {
      throw new Error('No data received');
    }

    sheet.appendRow([
      new Date(),
      data.name || '',
      (data.attending || '').toUpperCase(),
      data.message || ''
    ]);

    updateSummary();

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.ping) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService
    .createTextOutput('RSVP webhook is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function updateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let summary = ss.getSheetByName('Summary');
  if (!summary) {
    summary = ss.insertSheet('Summary');
    summary.setTabColor('#5c7a5c');
  }

  summary.clear();

  summary.getRange('A1').setValue('RSVP Response Tracker');
  summary.getRange('A1:B1').merge();
  summary.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#2a3a2a');
  summary.getRange('A1:B1').setHorizontalAlignment('center').setBackground('#d8e2d3');

  summary.getRange('A3').setValue('Total Responses');
  summary.getRange('B3').setFormula('=COUNTA(RSVPs!B2:B)');
  summary.getRange('A4').setValue('Accepted (YES)');
  summary.getRange('B4').setFormula('=COUNTIF(RSVPs!C2:C,"YES")');
  summary.getRange('A5').setValue('Declined (NO)');
  summary.getRange('B5').setFormula('=COUNTIF(RSVPs!C2:C,"NO")');

  summary.getRange('A3:A5').setFontWeight('bold').setFontColor('#2a3a2a');
  summary.getRange('A3:A5').setBackground('#f2f5ee');
  summary.getRange('B3:B5').setFontSize(14).setFontWeight('bold').setFontColor('#5c7a5c');
  summary.getRange('B3:B5').setHorizontalAlignment('center').setBackground('#ffffff');
  summary.getRange('B3:B5').setBorder(true, true, true, true, false, false);

  summary.setColumnWidths(1, 2, 220);
}
