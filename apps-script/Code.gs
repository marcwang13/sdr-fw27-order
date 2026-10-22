/**
 * SUN DAY RED FW27 線上訂貨：收單程式（Google Apps Script，綁在 Google 試算表上）
 *
 * 安裝：
 * 1. 新增一個 Google 試算表 → 選單「擴充功能」→「Apps Script」
 * 2. 把這個檔案的內容整份貼上取代原本的程式碼，儲存
 * 3. 右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *    執行身分：我；誰可以存取：所有人 → 部署 → 依畫面授權
 * 4. 複製「網頁應用程式網址」，貼到網站的 config.js 的 SUBMIT_URL
 *
 * 每張訂單會寫入兩個分頁：
 * - 訂單：一張單一列（同一個訂單編號再次送出會覆蓋舊資料）
 * - 明細：每個款色尺寸一列（連鎖分店會分店分開列）
 */

const ORDER_HEAD = ['送出時間', '訂單編號', '經銷商／店名', '聯絡人', '電話', 'Email', '訂單備註', '款色數', '件數', '零售金額', '分店數', '分店', '狀態'];
const LINE_HEAD = ['送出時間', '訂單編號', '經銷商／店名', '分店', '分頁', '月份', '款號', '色碼 SKU', '英文品名', '中文品名', '顏色', '尺寸', '數量', '台灣建議售價', '零售金額'];

function doGet() {
  return ContentService.createTextOutput('SDR FW27 收單程式運作中');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d || !d.oid || !Array.isArray(d.lines)) return out_('invalid');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const os = sheet_(ss, '訂單', ORDER_HEAD);
    const ls = sheet_(ss, '明細', LINE_HEAD);
    removeOid_(os, d.oid);
    removeOid_(ls, d.oid);

    const t = d.submittedAt ? new Date(d.submittedAt) : new Date();
    const branches = Array.isArray(d.branches) ? d.branches : [];
    os.appendRow([t, d.oid, d.store || '', d.contact || '', d.phone || '', d.email || '', d.note || '',
      d.styles || 0, d.qty || 0, d.amount || 0, branches.length || '', branches.map(b => b.name).join('、'), d.status || '已送出']);

    const info = {};
    d.lines.forEach(l => { info[l.k] = l; });
    const rows = [];
    const push = (branch, l, qty) => {
      const x = info[l.k] || l;
      rows.push([t, d.oid, d.store || '', branch, x.sheet || '', x.mon || '', x.style || '', x.sku || '', x.en || '', x.zh || '', x.color || '',
        l.size, qty, x.price || 0, qty * (x.price || 0)]);
    };
    if (branches.length) {
      branches.forEach(b => (b.lines || []).forEach(l => push(b.name, l, l.qty)));
    } else {
      d.lines.forEach(l => push('', l, l.qty));
    }
    if (rows.length) ls.getRange(ls.getLastRow() + 1, 1, rows.length, LINE_HEAD.length).setValues(rows);
    return out_('ok');
  } catch (err) {
    return out_('error: ' + err);
  } finally {
    lock.releaseLock();
  }
}

function sheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#610e0e').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function removeOid_(sh, oid) {
  const n = sh.getLastRow();
  if (n < 2) return;
  const ids = sh.getRange(2, 2, n - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(oid)) sh.deleteRow(i + 2);
  }
}

function out_(s) {
  return ContentService.createTextOutput(s);
}
