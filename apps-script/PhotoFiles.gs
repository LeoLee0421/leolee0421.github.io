/**
 * 網站照片清單：把 Photos 分頁每個活動 drive_url 資料夾裡的照片，寫進 PhotoFiles 分頁。
 * 網站會從 PhotoFiles 隨機挑照片放進各活動的藝廊。
 *
 * 安裝：試算表 → 擴充功能 → Apps Script，把這整個檔案貼上並儲存。
 * 使用：重新整理試算表後，上方選單「網站 → 更新照片清單」。
 * 自動更新：在 Apps Script 執行一次 installDailyTrigger，之後每天會自動更新。
 */

var MAX_PER_ALBUM = 500;   // 每個資料夾最多記幾張
var SUBFOLDER_DEPTH = 2;   // 往下讀幾層子資料夾

function onOpen() {
  SpreadsheetApp.getUi().createMenu('網站')
    .addItem('更新照片清單', 'updatePhotoFiles')
    .addToUi();
}

function updatePhotoFiles() {
  var ss = SpreadsheetApp.getActive();
  var photos = ss.getSheetByName('Photos');
  var data = photos.getDataRange().getValues();
  var head = data[0].map(function (h) { return String(h).trim().toLowerCase(); });
  var cUrl = head.indexOf('drive_url'), cTitle = head.indexOf('title');
  var out = [['folder_id', 'title', 'file_id', 'name']];
  var problems = [];

  for (var r = 1; r < data.length; r++) {
    var m = String(data[r][cUrl] || '').match(/folders\/([\w-]+)/);
    if (!m) continue;
    var title = data[r][cTitle], found = [];
    try {
      collect(DriveApp.getFolderById(m[1]), found, 0);
    } catch (e) {
      problems.push(title + '：' + e.message);
      continue;
    }
    // 依檔名排序，讓沒指定 cover 的相簿封面固定是第一張
    found.sort(function (a, b) { return a.name < b.name ? -1 : a.name > b.name ? 1 : 0; });
    found.slice(0, MAX_PER_ALBUM).forEach(function (f) { out.push([m[1], title, f.id, f.name]); });
  }

  var sheet = ss.getSheetByName('PhotoFiles') || ss.insertSheet('PhotoFiles');
  sheet.clearContents();
  sheet.getRange(1, 1, out.length, 4).setValues(out);
  sheet.setFrozenRows(1);

  var msg = '已更新 ' + (out.length - 1) + ' 張照片。';
  if (problems.length) msg += '\n\n讀不到的資料夾：\n' + problems.join('\n');
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function collect(folder, found, depth) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    if (String(f.getMimeType()).indexOf('image/') === 0) found.push({ id: f.getId(), name: f.getName() });
  }
  if (depth >= SUBFOLDER_DEPTH) return;
  var subs = folder.getFolders();
  while (subs.hasNext()) collect(subs.next(), found, depth + 1);
}

// 執行一次就好：之後每天清晨自動更新照片清單
function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'updatePhotoFiles') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('updatePhotoFiles').timeBased().everyDays(1).atHour(5).create();
}
