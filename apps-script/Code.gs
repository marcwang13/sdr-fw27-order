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
 *
 * 刪除／作廢訂單：
 * - 直接刪掉「訂單」分頁裡那一列（明細可一起刪），或把「狀態」改成「作廢」
 * - 之後客人就算從「我的訂單」修改再送出，也不會再寫進來，網頁會請客人重新下新單
 * - 隱藏分頁「_紀錄」記錄收過的訂單編號，請不要刪除
 *
 * Email 通知：每張新訂單／修改都會寄信到 NOTIFY_TO（多個收件人用逗號分隔，留空就不寄）
 */

const NOTIFY_TO = 'marc_wang@emers.com';

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
    const log = logSheet_(ss);
    const seen = log.getLastRow() > 0 && log.getRange(1, 1, log.getLastRow(), 1).getValues().some(r => String(r[0]) === String(d.oid));
    const row = findOid_(os, d.oid);
    if (row && String(os.getRange(row, ORDER_HEAD.indexOf('狀態') + 1).getValue()).trim() === '作廢') return out_('void');
    if (seen && !row) return out_('deleted');
    if (!seen) log.appendRow([d.oid, new Date()]);
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
    try { notify_(ss, d, rows, seen); } catch (err) { console.error('notify failed: ' + err); }
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

function logSheet_(ss) {
  let sh = ss.getSheetByName('_紀錄');
  if (!sh) { sh = ss.insertSheet('_紀錄'); sh.hideSheet(); }
  return sh;
}

function findOid_(sh, oid) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const ids = sh.getRange(2, 2, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(oid)) return i + 2;
  return 0;
}

function removeOid_(sh, oid) {
  const n = sh.getLastRow();
  if (n < 2) return;
  const ids = sh.getRange(2, 2, n - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(oid)) sh.deleteRow(i + 2);
  }
}

function notify_(ss, d, rows, isEdit) {
  if (!NOTIFY_TO) return;
  const esc = v => String(v == null ? '' : v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const fmt = n => Number(n || 0).toLocaleString('en-US');
  const kind = isEdit ? '訂單修改' : '新訂單';
  const subject = `【SDR FW27 ${kind}】${d.store || '（未填店名）'}｜${fmt(d.qty)} 件｜NT$${fmt(d.amount)}`;
  // 依分店、款色彙整成一行一款：M×4 L×2
  const groups = {};
  rows.forEach(r => { const key = [r[3], r[6], r[7], r[9], r[10]].join('|'); (groups[key] = groups[key] || { branch: r[3], style: r[6], sku: r[7], zh: r[9], color: r[10], sizes: [], qty: 0, amt: 0 });
    groups[key].sizes.push(`${r[11]}×${r[12]}`); groups[key].qty += r[12]; groups[key].amt += r[14]; });
  const list = Object.values(groups);
  const hasBranch = list.some(g => g.branch);
  const tr = list.map(g => `<tr>${hasBranch ? `<td>${esc(g.branch)}</td>` : ''}<td>${esc(g.style)}<br><span style="color:#888">${esc(g.sku)}</span></td><td>${esc(g.zh)}</td><td>${esc(g.color)}</td><td>${esc(g.sizes.join('  '))}</td><td style="text-align:right">${g.qty}</td><td style="text-align:right">${fmt(g.amt)}</td></tr>`).join('');
  const th = s => `<th style="background:#610e0e;color:#fff;padding:6px 8px;text-align:left;font-weight:600">${s}</th>`;
  const html = `<div style="font-family:Arial,'Microsoft JhengHei',sans-serif;font-size:14px;color:#222">
  <h2 style="color:#610e0e;margin:0 0 8px">SUN DAY RED FW27 ${kind}</h2>
  <table style="border-collapse:collapse;margin-bottom:12px">
  <tr><td style="color:#666;padding:2px 12px 2px 0">訂單編號</td><td><b>${esc(d.oid)}</b></td></tr>
  <tr><td style="color:#666;padding:2px 12px 2px 0">經銷商／店名</td><td>${esc(d.store)}</td></tr>
  <tr><td style="color:#666;padding:2px 12px 2px 0">聯絡人</td><td>${esc(d.contact)} ${esc(d.phone)} ${esc(d.email)}</td></tr>
  ${d.note ? `<tr><td style="color:#666;padding:2px 12px 2px 0">備註</td><td>${esc(d.note)}</td></tr>` : ''}
  ${(d.branches || []).length ? `<tr><td style="color:#666;padding:2px 12px 2px 0">分店</td><td>${esc(d.branches.map(b => b.name).join('、'))}</td></tr>` : ''}
  <tr><td style="color:#666;padding:2px 12px 2px 0">合計</td><td><b>${fmt(d.styles)} 款色、${fmt(d.qty)} 件、零售金額 NT$${fmt(d.amount)}</b></td></tr>
  </table>
  <table style="border-collapse:collapse;font-size:13px" border="1" bordercolor="#ddd" cellpadding="6">
  <tr>${hasBranch ? th('分店') : ''}${th('款號')}${th('品名')}${th('顏色')}${th('尺寸×數量')}${th('件數')}${th('零售金額')}</tr>${tr}</table>
  <p><a href="${ss.getUrl()}" style="color:#610e0e">開啟 Google 試算表看完整訂單</a></p></div>`;
  MailApp.sendEmail({ to: NOTIFY_TO, subject: subject, htmlBody: html, name: 'SDR FW27 線上訂貨' });
}

function out_(s) {
  return ContentService.createTextOutput(s);
}
