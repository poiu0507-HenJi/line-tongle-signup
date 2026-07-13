// 幸福同樂會 報名表　接收程式（貼到 Google Apps Script）
// 功能：收到報名表送出的資料後，自動寫入這個試算表的第一個工作表

// 欄位順序（跟報名表送出的名稱一致，也會當成試算表標題列）
const HEADERS = ['姓名', '綽號', '身分', '輔導員', '小天使', '圓夢據計劃', '五感恩', '如何得知本活動', '填寫時間'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000); // 避免多人同時送出時搶寫
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    // 試算表還是空的話，先寫入標題列
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
    }

    const p = (e && e.parameter) ? e.parameter : {};
    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

    // 按照 HEADERS 順序組出一列資料
    const row = HEADERS.map(function (h) {
      if (h === '填寫時間') return p[h] || now; // 沒帶時間就用伺服器時間
      return p[h] || '';
    });

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// 用瀏覽器直接開網址時，回一句話確認服務有在運作
function doGet() {
  return ContentService.createTextOutput('幸福同樂會報名表 webhook 正常運作中');
}
