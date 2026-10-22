# SUN DAY RED FW27 經銷線上訂貨

經銷商打開網站即可挑款、依尺寸填數量、送出訂單或下載 Excel 訂單本，不需要登入。

- `index.html`：下單頁（商品資料在 `data.js`，縮圖在 `imgs.js`，大圖在 `big/`）
- `config.js`：收單網址設定（`SUBMIT_URL`）。留空時為「傳單模式」，經銷商下載 Excel 後用 LINE / Email 傳給業務。
- `apps-script/Code.gs`：貼到 Google 試算表的 Apps Script，部署成網頁應用程式後即可收單，訂單寫入「訂單」與「明細」分頁。
