/**
 * 💰 Google Sheet 雙幣記帳與資產管理助手 (後端 Google Apps Script)
 * 
 * 【安裝說明】
 * 1. 在 Google 試算表中點選「擴充功能」>「Apps Script」
 * 2. 將本檔案內容完整貼入預設的 `Code.gs`（或 `程式碼.gs`）
 * 3. 點選左側「+」新增「HTML」檔案，檔案名稱請務必設為 `Page`（系統會自動加上 .html）
 * 4. 將 `Page.html` 的內容貼入並儲存
 */

// 網頁應用程式 (Web App) 進入點；HTML 檔案名稱請設為 Page
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Page').setTitle('記帳助手 🌍')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 試算表開啟時自動建立上方「💰 記帳助手」自訂選單
function onOpen() {
  SpreadsheetApp.getUi().createMenu('💰 記帳助手')
    .addItem('📱 開啟記帳側邊欄', 'showSidebar')
    .addItem('🔄 重算今日起始額度 (G欄)', 'calculateDailyBudget')
    .addItem('📅 執行固定扣款檢查', 'autoProcessFixedExpenses')
    .addItem('⏰ 重新建立自動定時器', 'createTriggers')
    .addToUi();
}

// 於電腦版 Google 試算表右側開啟記帳側邊欄
function showSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('Page').setTitle('記帳助手 🌍')
  );
}

// 數值或含貨幣符號、千分位的金額文字均可安全解析
function parseMoney_(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : NaN;
  if (value == null || String(value).trim() === '') return NaN;
  let text = String(value).normalize('NFKC').trim().replace(/−/g, '-');
  const negative = /^\(.*\)$/.test(text);
  text = text.replace(/(?:NT\$|TWD|USD|JPY|US\$|NTD|新台幣|台幣|元|圓|日圓|美元)/gi, '')
    .replace(/[,\s$¥￥()]/g, '');
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(text)) return NaN;
  const result = Number(text);
  return Number.isFinite(result) ? (negative ? -Math.abs(result) : result) : NaN;
}

function moneyOrZero_(value) {
  const n = parseMoney_(value);
  return Number.isFinite(n) ? n : 0;
}

function sheet_(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('找不到「' + name + '」工作表，請確認工作表名稱是否正確');
  return s;
}

function today_() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
}

function amount_(value) {
  const n = parseMoney_(value);
  if (!Number.isFinite(n) || n <= 0) throw new Error('請輸入大於 0 的有效金額');
  return n;
}

// 讀取「預算進度」工作表中的即時匯率（J1: 美金匯率, J2: 日幣匯率）
function rates_() {
  const s = sheet_('預算進度');
  return {
    TWD: 1,
    USD: moneyOrZero_(s.getRange('J1').getValue()) || 30,
    JPY: moneyOrZero_(s.getRange('J2').getValue()) || 0.21
  };
}

// 每日自動根據「當月剩餘預算 ÷ 本月剩餘天數」重算今日可用起始額度 (寫入 G6, G8)
function calculateDailyBudget() {
  const s = sheet_('預算進度');
  const parts = today_().split('/').map(Number);
  const remainingDays = new Date(parts[0], parts[1], 0).getDate() - parts[2] + 1;
  s.getRange('G6').setValue(Math.max(0, Math.round(moneyOrZero_(s.getRange('D6').getValue()) / remainingDays)));
  s.getRange('G8').setValue(Math.max(0, Math.round(moneyOrZero_(s.getRange('D8').getValue()) / remainingDays)));
  SpreadsheetApp.flush();
}

// 儲存支出紀錄（E欄寫入折合台幣預算金額，H欄寫入原幣別扣款金額）
function saveTransaction(data) {
  const amt = amount_(data.amount), rates = rates_();
  if (!rates[data.currency]) throw new Error('不支援的幣別');
  sheet_('流水帳').appendRow([
    data.date,
    data.category,
    '支出',
    data.subItem,
    amt * rates[data.currency],
    data.account,
    data.note,
    amt,
    data.category
  ]);
  SpreadsheetApp.flush();
  return '✨ 紀錄成功！';
}

// 儲存收入紀錄（E欄為 0 避免影響支出預算，H欄寫入實際入帳金額）
function saveIncome(data) {
  const amt = amount_(data.amount), rates = rates_();
  const cur = data.currency || 'TWD';
  if (!rates[cur]) throw new Error('不支援的幣別');
  const note = cur === 'TWD'
    ? 'TWD 收入'
    : cur + ' 收入 (折合台幣約 NT$' + Math.round(amt * rates[cur]) + ')';
  sheet_('流水帳').appendRow([
    data.date,
    '收入',
    '收入',
    data.source,
    0,
    data.account,
    note,
    amt
  ]);
  SpreadsheetApp.flush();
  return '🎊 收入已入帳！';
}

// 執行帳戶間轉帳（一次新增一筆轉出與一筆轉入紀錄）
function saveTransfer(data) {
  const amt = amount_(data.amount), rates = rates_();
  const cur = data.currency || 'TWD';
  if (!rates[cur]) throw new Error('不支援的幣別');
  if (!data.fromAccount || !data.toAccount || data.fromAccount === data.toAccount) {
    throw new Error('請選擇不同的轉出及轉入帳戶');
  }
  const note = cur === 'TWD'
    ? 'TWD 轉帳'
    : cur + ' 轉帳 (折合台幣約 NT$' + Math.round(amt * rates[cur]) + ')';
  const s = sheet_('流水帳');
  s.getRange(s.getLastRow() + 1, 1, 2, 8).setValues([
    [data.date, '轉帳', '支出', '轉出至 ' + data.toAccount, 0, data.fromAccount, note, amt],
    [data.date, '轉帳', '收入', '來自 ' + data.fromAccount + ' 轉入', 0, data.toAccount, note, amt]
  ]);
  SpreadsheetApp.flush();
  return '✅ 轉帳成功！';
}

// 帳戶餘額手動校準（自動計算差額並補上一筆餘額校準紀錄）
function processAdjustment(name, diff, finalVal) {
  diff = parseMoney_(diff);
  finalVal = parseMoney_(finalVal);
  if (!Number.isFinite(diff) || !Number.isFinite(finalVal)) throw new Error('校準金額無效');
  sheet_('流水帳').appendRow([
    today_(),
    '餘額校準',
    diff >= 0 ? '收入' : '支出',
    '手動校準誤差',
    0,
    name,
    '校準至 ' + finalVal,
    Math.abs(diff)
  ]);
  SpreadsheetApp.flush();
  return '✅ 【' + name + '】已校準！';
}

function validRow_(sheet, value) {
  const row = Number(value);
  if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) throw new Error('找不到該筆紀錄');
  return row;
}

// 更新歷史紀錄明細（保留原始幣別匯率比例及收入/轉帳的 E 欄 0，避免編輯時改壞帳戶餘額）
function updateTransaction(data) {
  const s = sheet_('流水帳'), row = validRow_(s, data.rowIndex), amt = amount_(data.amount);
  const old = s.getRange(row, 1, 1, 8).getValues()[0];
  const oldTwd = moneyOrZero_(old[4]), oldOriginal = moneyOrZero_(old[7]);
  const ratio = oldOriginal ? oldTwd / oldOriginal : 1;
  old[0] = data.date;
  old[3] = data.item;
  old[4] = oldTwd === 0 ? 0 : amt;
  old[5] = data.account;
  old[6] = data.note;
  old[7] = oldTwd === 0 ? amt : (ratio ? amt / ratio : amt);
  s.getRange(row, 1, 1, 8).setValues([old]);
  SpreadsheetApp.flush();
  return '✅ 紀錄已成功更新！';
}

// 刪除單筆歷史紀錄
function deleteTransaction(rowIndex) {
  const s = sheet_('流水帳');
  s.deleteRow(validRow_(s, rowIndex));
  SpreadsheetApp.flush();
  return '🗑️ 紀錄已成功刪除！';
}

// 取得首頁初始化資料（今日一般預算剩餘、特支本月剩餘、即時匯率、帳戶清單）
function getInitData() {
  const s = sheet_('預算進度'), a = sheet_('帳戶名單');
  const accounts = a.getLastRow() > 1
    ? a.getRange(2, 1, a.getLastRow() - 1, 1).getValues().flat().filter(v => String(v).trim() !== '' && String(v).trim() !== '帳戶名稱')
    : [];
  return {
    budget: {
      general: moneyOrZero_(s.getRange('F6').getValue()),
      special: moneyOrZero_(s.getRange('D8').getValue())
    },
    rates: rates_(),
    accounts: accounts
  };
}

// 取得所有帳戶餘額、總資產匯總、保底總資產與當月存錢目標
function getFullAccountBalances() {
  SpreadsheetApp.flush();
  const s = sheet_('帳戶名單'), data = s.getDataRange().getValues();
  let twd = 0, usd = 0, usdToTwd = 0, total = 0;
  for (let i = 0; i < data.length - 1; i++) {
    for (let j = 0; j < data[i].length; j++) {
      const label = String(data[i][j]).trim();
      if (label === '台幣總資產') twd = moneyOrZero_(data[i + 1][j]);
      if (label === '美金總資產') {
        usd = moneyOrZero_(data[i + 1][j]);
        usdToTwd = moneyOrZero_(data[i + 1][j + 1]);
      }
      if (label === '現金總資產') total = moneyOrZero_(data[i + 1][j]);
    }
  }

  // 保底總資產標題動態讀取 I21（若空白則顯示預設名稱），金額固定讀取 I22
  const rawLabel = String(s.getRange('I21').getValue() || '').trim();
  const secureLabel = rawLabel || '保底總資產 (預扣專款基金)';
  const cell = s.getRange('I22');
  const secureTotal = parseMoney_(cell.getValue());
  if (!Number.isFinite(secureTotal)) {
    throw new Error('帳戶名單!I22 為空白或不是有效金額，請檢查數值或公式');
  }

  // 讀取當月保底存錢目標（涵蓋 I28:J45 完整 1~12 月底列範圍）
  let savingsGoal = '未設定';
  const parts = today_().split('/').map(Number), target = parts[0] + '年' + parts[1] + '月底';
  const goals = s.getRange('I28:J45').getValues();
  for (let k = 0; k < goals.length; k++) {
    if (String(goals[k][0]).trim() === target) {
      const n = parseMoney_(goals[k][1]);
      if (Number.isFinite(n)) savingsGoal = Number(n.toFixed(2)) + ' 萬';
      break;
    }
  }

  const detail = s.getRange('A2:C25').getValues()
    .filter(r => r[0] !== '' && r[0] !== '帳戶名稱')
    .map(r => ({ name: String(r[0]), current: moneyOrZero_(r[2]) }));

  return {
    detail: detail,
    secureLabel: secureLabel,
    secureTotal: secureTotal,
    savingsGoal: savingsGoal,
    summary: [
      { label: '台幣總資產', value: twd },
      { label: '美金總資產', value: usd, subValue: usdToTwd },
      { label: '現金總資產', value: total }
    ]
  };
}

// 讀取「選單設定」工作表中的「預算分類 -> 詳細項目」對應表
function getCategoryMapping() {
  const s = sheet_('選單設定'), map = {};
  if (s.getLastRow() < 2) return map;
  s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(r => {
    if (r[0]) {
      if (!map[r[0]]) map[r[0]] = [];
      map[r[0]].push(r[1]);
    }
  });
  return map;
}

// 依年月字串 (例如 "2026/10") 查詢歷史流水帳紀錄
function getHistoryByDate(targetString) {
  const s = sheet_('流水帳');
  if (s.getLastRow() < 2) return [];
  const data = s.getRange(2, 1, s.getLastRow() - 1, 8).getValues(), result = [];
  const target = String(targetString).replace(/-/g, '/');
  data.forEach((r, i) => {
    let date = r[0] instanceof Date
      ? Utilities.formatDate(r[0], 'Asia/Taipei', 'yyyy/MM/dd')
      : String(r[0] || '').replace(/-/g, '/');
    const match = date.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
    if (match) date = match[1] + '/' + match[2].padStart(2, '0') + '/' + match[3].padStart(2, '0');
    if (date.startsWith(target)) {
      result.push({
        rowIndex: i + 2,
        date: date,
        category: r[1],
        type: r[2],
        item: r[3],
        amount: moneyOrZero_(r[4]) !== 0 ? moneyOrZero_(r[4]) : moneyOrZero_(r[7]),
        account: r[5],
        note: r[6]
      });
    }
  });
  return result.sort((a, b) => b.date.localeCompare(a.date) || b.rowIndex - a.rowIndex);
}

// 每日自動檢查「固定扣款設定」工作表，若今日為扣款日且狀態為「啟用」則自動寫入流水帳（含防重複機制）
function autoProcessFixedExpenses() {
  const s = sheet_('固定扣款設定');
  if (s.getLastRow() < 2) return;
  const tx = sheet_('流水帳'), rates = rates_(), today = today_(), day = Number(today.split('/')[2]);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const existing = tx.getLastRow() > 1 ? tx.getRange(2, 1, tx.getLastRow() - 1, 8).getValues() : [];
    s.getRange(2, 1, s.getLastRow() - 1, 8).getValues().forEach(r => {
      const [name, category, amount, currency, debitDay, account, note, status] = r;
      if (status !== '啟用' || parseInt(debitDay, 10) !== day) return;
      const autoNote = '(自動) ' + (note || '');
      const duplicate = existing.some(e => {
        const date = e[0] instanceof Date
          ? Utilities.formatDate(e[0], 'Asia/Taipei', 'yyyy/MM/dd')
          : String(e[0]).replace(/-/g, '/');
        return date === today && e[3] === name && e[5] === account && e[6] === autoNote;
      });
      if (duplicate) return;
      const amt = amount_(amount);
      if (!rates[currency]) throw new Error('固定扣款幣別無效：' + currency);
      const row = [today, category, '支出', name, amt * rates[currency], account, autoNote, amt, category || '一般花費'];
      tx.appendRow(row);
      existing.push(row);
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

// 一鍵建立每日自動觸發器（凌晨 1 點檢查固定扣款、早上 6 點重算今日起始預算額度）
function createTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['calculateDailyBudget', 'autoProcessFixedExpenses'].includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('calculateDailyBudget')
    .timeBased().inTimezone('Asia/Taipei').atHour(6).everyDays(1).create();
  ScriptApp.newTrigger('autoProcessFixedExpenses')
    .timeBased().inTimezone('Asia/Taipei').atHour(1).everyDays(1).create();
}
