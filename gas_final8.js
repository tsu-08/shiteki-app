// 現場検査記録 - Google Apps Script
const FOLDER_NAME = '現場検査記録';
const TEMP_FOLDER = '現場検査記録_一時ファイル';

function doPost(e) {
  try {
    var raw = e.postData.contents;
    var data = JSON.parse(raw);
    return processData(data);
  } catch (err) {
    Logger.log('doPostエラー:' + err.toString());
    return buildResponse({ error: 'doPost error: ' + err.toString() });
  }
}

function doGet(e) {
  return buildResponse({ status: 'OK', message: '現場検査記録GASは正常に動作しています' });
}

function processData(data) {
  var outputType = data.type || 'excel';
  var folder = getOrCreateFolder(FOLDER_NAME);
  var tempFolder = getOrCreateFolder(TEMP_FOLDER);

  // 写真を一時フォルダに保存してURLを取得
  var photoUrls = buildPhotoUrls(data, tempFolder);

  // スプレッドシートを作成
  var ss = buildSpreadsheet(data, folder, photoUrls);
  var ssId = ss.getId();
  var ssFile = DriveApp.getFileById(ssId);

  var url;

  if (outputType === 'pdf') {
    // PDFはスプレッドシートから直接生成
    var pdfFile = exportAsPdf(ss, data, folder);
    url = pdfFile.getUrl();
    // スプレッドシートも保存（後でExcel変換できるように）
    Logger.log('PDF生成完了: ' + url);
  } else {
    // Excelはスプレッドシートのダウンロードリンクを返す
    // スプレッドシートをフォルダに移動して保存
    Logger.log('スプレッドシート保存: ' + ssId);
    url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=xlsx';
    // ダウンロード用の共有リンクも生成
    ssFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    url = ssFile.getUrl();
  }

  return buildResponse({ url: url, type: outputType });
}

function buildPhotoUrls(data, tempFolder) {
  var photoUrls = [];
  data.rows.forEach(function(row, idx) {
    var rowUrls = { left: [], right: [] };
    for (var i = 0; i < 4; i++) {
      rowUrls.left.push(savePhotoToTemp(row.leftPhotos && row.leftPhotos[i], tempFolder, 'L' + idx + '_' + i));
      rowUrls.right.push(savePhotoToTemp(row.rightPhotos && row.rightPhotos[i], tempFolder, 'R' + idx + '_' + i));
    }
    photoUrls.push(rowUrls);
  });
  return photoUrls;
}

function savePhotoToTemp(base64DataUrl, tempFolder, name) {
  if (!base64DataUrl || typeof base64DataUrl !== 'string' || base64DataUrl.indexOf('data:') !== 0) {
    return null;
  }
  try {
    var parts = base64DataUrl.split(',');
    var mime = (parts[0].match(/:(.*?);/) || [])[1] || 'image/jpeg';
    var ext = mime.indexOf('png') >= 0 ? '.png' : '.jpg';
    var blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), mime, name + ext);
    var file = tempFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  } catch(err) {
    Logger.log('写真保存エラー: ' + err.toString());
    return null;
  }
}

function buildResponse(obj) {
  var output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function buildSpreadsheet(data, folder, photoUrls) {
  var ss = SpreadsheetApp.create(buildFileName(data, 'sheet'));
  var defaultSheet = ss.getActiveSheet();
  for (var idx = 0; idx < data.rows.length; idx++) {
    var sheet;
    if (idx === 0) {
      sheet = defaultSheet;
      sheet.setName('No' + (idx + 1));
    } else {
      sheet = ss.insertSheet('No' + (idx + 1));
    }
    buildSheet(sheet, data, data.rows[idx], idx, photoUrls[idx]);
  }
  SpreadsheetApp.flush();
  Utilities.sleep(2000);
  var file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  return ss;
}

function buildSheet(sheet, data, row, idx, urls) {
  sheet.setColumnWidth(1, 270);
  sheet.setColumnWidth(2, 270);
  var headerRange = sheet.getRange(1, 1, 1, 2);
  headerRange.merge();
  var inspStr = [data.insp1, data.insp2].filter(Boolean).join(' / ');
  headerRange.setValue(data.owner + ' 様邸　' + formatDate(data.date) + '　' + data.inspType + '　担当：' + inspStr);
  headerRange.setBackground('#1b3a6b').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  headerRange.setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 30);
  sheet.getRange(2, 1).setValue('指摘事項 No.' + (idx + 1));
  sheet.getRange(2, 1).setBackground('#e6edf8').setFontWeight('bold').setFontSize(9).setFontColor('#1b3a6b');
  sheet.getRange(2, 2).setValue('↓↓是正写真はこの面に貼付けて提出してください。');
  sheet.getRange(2, 2).setBackground('#e6f4f0').setFontWeight('bold').setFontSize(9).setFontColor('#0a7c5c');
  sheet.setRowHeight(2, 22);
  var PHOTO_ROW_HEIGHT = 150;
  var PHOTO_START_ROW = 3;
  var TOTAL_ROWS = 6;
  for (var i = 0; i < 4; i++) {
    var r = PHOTO_START_ROW + i;
    sheet.setRowHeight(r, PHOTO_ROW_HEIGHT);
    var leftCell = sheet.getRange(r, 1);
    leftCell.setBackground('#f0f3f8').setVerticalAlignment('middle').setHorizontalAlignment('center');
    var leftUrl = urls && urls.left && urls.left[i];
    if (leftUrl) {
      leftCell.setFormula('=IMAGE("' + leftUrl + '",4,' + (PHOTO_ROW_HEIGHT - 10) + ',260)');
    } else {
      leftCell.setValue('（写真なし）').setFontColor('#9098b0').setFontSize(9);
    }
    var rightCell = sheet.getRange(r, 2);
    rightCell.setBackground('#f0f8f4').setVerticalAlignment('middle').setHorizontalAlignment('center');
    var rightUrl = urls && urls.right && urls.right[i];
    var rightText = row.rightTexts && row.rightTexts[i];
    if (rightUrl) {
      rightCell.setFormula('=IMAGE("' + rightUrl + '",4,' + (PHOTO_ROW_HEIGHT - 10) + ',260)');
    } else if (rightText) {
      rightCell.setValue(rightText).setFontSize(10).setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('top');
    } else {
      rightCell.setValue('（建設会社が是正後に記入）').setFontColor('#c0c4d0').setFontSize(8).setWrap(true).setHorizontalAlignment('center');
    }
    sheet.getRange(r, 1, 1, 2).setBorder(true, true, true, true, true, false, '#8a90a8', SpreadsheetApp.BorderStyle.SOLID);
  }
  sheet.getRange(1, 1, TOTAL_ROWS, 2).setBorder(true, true, true, true, false, false, '#1b3a6b', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function exportAsPdf(ss, data, folder) {
  SpreadsheetApp.flush();
  Utilities.sleep(3000);
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
    '/export?format=pdf&size=A4&portrait=true&fitw=true&fith=true' +
    '&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false';
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
  return (data.date || '').replace(/-/g, '') + '_' +
    (data.owner || '物件') + '様_' +
    (data.inspType || '') + '_検査記録_' +
    (type === 'pdf' ? 'PDF' : type === 'sheet' ? 'Sheet' : 'Excel');
}

function formatDate(s) {
  if (!s) return '';
  var d = new Date(s);
  return '令和' + (d.getFullYear() - 2018) + '年' + (d.getMonth() + 1) + '月' + d.getDate() + '日';
}

function testRun() {
  var data = {
    type: 'excel',
    owner: 'テスト',
    date: '2026-05-04',
    insp1: '鶴岡 諭',
    insp2: '',
    inspType: '配筋',
    rows: [{
      no: 1,
      leftPhotos: [null, null, null, null],
      rightPhotos: [null, null, null, null],
      rightTexts: ['是正期限：5月末', '', '', '']
    }]
  };
  var result = processData(data);
  Logger.log(result.getContent());
}
