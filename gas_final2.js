// 現場検査記録 - Google Apps Script
const FOLDER_NAME = '現場検査記録';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents.replace(/^[^{]*/, ""));
    const outputType = data.type || 'excel';
    const folder = getOrCreateFolder(FOLDER_NAME);
    const ss = buildSpreadsheet(data, folder);
    let url;
    if (outputType === 'pdf') {
      const pdfFile = exportAsPdf(ss, data, folder);
      url = pdfFile.getUrl();
      DriveApp.getFileById(ss.getId()).setTrashed(true);
    } else {
      const xlsxFile = exportAsXlsx(ss, data, folder);
      url = xlsxFile.getUrl();
      DriveApp.getFileById(ss.getId()).setTrashed(true);
    }
    return buildResponse({ url });
  } catch (err) {
    return buildResponse({ error: err.toString() });
  }
}

function doGet(e) {
  return buildResponse({ status: 'OK', message: '現場検査記録GASは正常に動作しています' });
}

function buildResponse(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function buildSpreadsheet(data, folder) {
  const ss = SpreadsheetApp.create(buildFileName(data, 'temp'));
  const defaultSheet = ss.getActiveSheet();
  data.rows.forEach((row, idx) => {
    let sheet;
    if (idx === 0) { sheet = defaultSheet; sheet.setName('No' + (idx + 1)); }
    else { sheet = ss.insertSheet('No' + (idx + 1)); }
    buildSheet(sheet, data, row, idx);
  });
  const file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  return ss;
}

function buildSheet(sheet, data, row, idx) {
  sheet.setColumnWidth(1, 270);
  sheet.setColumnWidth(2, 270);
  const headerRange = sheet.getRange(1, 1, 1, 2);
  headerRange.merge();
  const inspStr = [data.insp1, data.insp2].filter(Boolean).join(' / ');
  headerRange.setValue(data.owner + ' 様邸　' + formatDate(data.date) + '　' + data.inspType + '　担当：' + inspStr);
  headerRange.setBackground('#1b3a6b').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  headerRange.setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 30);
  sheet.getRange(2, 1).setValue('指摘事項 No.' + (idx + 1));
  sheet.getRange(2, 1).setBackground('#e6edf8').setFontWeight('bold').setFontSize(9).setFontColor('#1b3a6b');
  sheet.getRange(2, 2).setValue('↓↓是正写真はこの面に貼付けて提出してください。');
  sheet.getRange(2, 2).setBackground('#e6f4f0').setFontWeight('bold').setFontSize(9).setFontColor('#0a7c5c');
  sheet.setRowHeight(2, 22);
  const PHOTO_ROW_HEIGHT = 150;
  const PHOTO_START_ROW = 3;
  const TOTAL_ROWS = 6;
  for (var i = 0; i < 4; i++) {
    var r = PHOTO_START_ROW + i;
    sheet.setRowHeight(r, PHOTO_ROW_HEIGHT);
    var leftCell = sheet.getRange(r, 1);
    leftCell.setBackground('#f0f3f8').setVerticalAlignment('middle').setHorizontalAlignment('center');
    if (row.leftPhotos[i]) {
      insertImage(sheet, row.leftPhotos[i], r, 1);
    } else {
      leftCell.setValue('（写真なし）').setFontColor('#9098b0').setFontSize(9);
    }
    var rightCell = sheet.getRange(r, 2);
    rightCell.setBackground('#f0f8f4').setVerticalAlignment('middle').setHorizontalAlignment('center');
    if (row.rightPhotos[i]) {
      insertImage(sheet, row.rightPhotos[i], r, 2);
    } else if (row.rightTexts[i]) {
      rightCell.setValue(row.rightTexts[i]).setFontSize(10).setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('top');
    } else {
      rightCell.setValue('（建設会社が是正後に記入）').setFontColor('#c0c4d0').setFontSize(8).setWrap(true).setHorizontalAlignment('center');
    }
    sheet.getRange(r, 1, 1, 2).setBorder(true, true, true, true, true, false, '#8a90a8', SpreadsheetApp.BorderStyle.SOLID);
  }
  sheet.getRange(1, 1, TOTAL_ROWS, 2).setBorder(true, true, true, true, false, false, '#1b3a6b', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function insertImage(sheet, base64DataUrl, row, col) {
  try {
    var parts = base64DataUrl.split(',');
    var mime = (parts[0].match(/:(.*?);/) || [])[1] || 'image/jpeg';
    var blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), mime, 'photo.jpg');
    var image = sheet.insertImage(blob, col, row);
    var cw = sheet.getColumnWidth(col);
    var rh = sheet.getRowHeight(row);
    var m = 4;
    var maxW = cw - m * 2;
    var maxH = rh - m * 2;
    var scale = Math.min(maxW / image.getWidth(), maxH / image.getHeight(), 1);
    image.setWidth(Math.floor(image.getWidth() * scale));
    image.setHeight(Math.floor(image.getHeight() * scale));
    image.setAnchorCell(sheet.getRange(row, col));
    image.setAnchorCellXOffset(Math.floor((maxW - image.getWidth()) / 2));
    image.setAnchorCellYOffset(m);
  } catch(err) {
    sheet.getRange(row, col).setValue('（写真エラー）').setFontSize(8).setFontColor('#c0392b');
  }
}

function exportAsXlsx(ss, data, folder) {
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  var blob = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName(buildFileName(data, 'excel') + '.xlsx');
  return folder.createFile(blob);
}

function exportAsPdf(ss, data, folder) {
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
    '/export?format=pdf&size=A4&portrait=true&fitw=true&fith=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false';
  var blob = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName(buildFileName(data, 'pdf') + '.pdf');
  return folder.createFile(blob);
}

function getOrCreateFolder(name) {
  var f = DriveApp.getFoldersByName(name);
  return f.hasNext() ? f.next() : DriveApp.createFolder(name);
}

function buildFileName(data, type) {
  return (data.date||'').replace(/-/g,'') + '_' + (data.owner||'物件') + '様_' +
    (data.inspType||'') + '_検査記録_' + (type==='pdf' ? 'PDF' : 'Excel');
}

function formatDate(s) {
  if (!s) return '';
  var d = new Date(s);
  return '令和' + (d.getFullYear()-2018) + '年' + (d.getMonth()+1) + '月' + d.getDate() + '日';
}
