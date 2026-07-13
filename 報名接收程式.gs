// 幸福同樂會 報名表　接收程式（貼到 Google Apps Script）
// 功能：收到報名表送出的資料後，自動寫入這個試算表的第一個工作表
// 特色：每次送出都會自動校正標題列（欄位改名也會自動更新，不必手動改）

// 欄位順序（跟報名表送出的名稱一致，也是試算表標題列）
const HEADERS = ['姓名', '綽號', '身分', '輔導員', '小天使', '圓夢計劃', '五感恩', '如何得知本活動', '填寫時間'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000); // 避免多人同時送出時搶寫
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    // 每次都把標題列寫成最新的（可自動修正欄位名稱，例如把舊錯字改掉）
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

    const p = (e && e.parameter) ? e.parameter : {};
    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

    // 按 HEADERS 順序組出一列資料
    const row = HEADERS.map(function (h) {
      return (h === '填寫時間') ? (p[h] || now) : (p[h] || '');
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
