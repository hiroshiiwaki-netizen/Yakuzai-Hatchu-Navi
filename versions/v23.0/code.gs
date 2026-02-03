/**
 * 薬剤発注ナビ v23.0 (患者検索修正版)
 * C列(漢字)とD列(カナ)を正しく読み込むように修正
 */

const CONFIG = {
  APP_NAME: '薬剤発注ナビ',
  SHEET_NAME: 'シート1',
  ALLOWED_DOMAIN: 'nhw.jp',
  
  CALENDAR_ID: 'c_fbaa34d73eb9d9fada0ce4f25fcaefa4bcc3ba7626d6194d50c1778fef9d0244@group.calendar.google.com',
  CHAT_WEBHOOK_URL: 'https://chat.googleapis.com/v1/spaces/AAQAdWFuXFM/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=DTkSi3XOT1oyoM14hhEryyYOaWbnbnz4dTDtIqAzKOA',
  
  STAFF_MASTER_ID: '1wR-ht-2MFrf2NQFPef2Xol3iQ7qIwXsQv_9T2IImoTM',
  PATIENT_MASTER_ID: '1zD7lIxWrMEzma9GDP0Yltp_YG56pOxVUsYYPNJ3a-nE',
  DATA_DB_ID: '1ffuhDYZSts3u6YN0vzQX_CwVXWDqRLD2zn3vkFbqtxQ',

  STATUS: { PENDING: '未発注', ORDERED: '発注済', DELIVERED: '納品済' }
};

function doGet() {
  const html = HtmlService.createTemplateFromFile('index').evaluate();
  return html.setTitle(CONFIG.APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- リスト更新用 ---
function getLatestOrders() { return step3_GetOrders(); }

// --- ステップ1: 認証 ---
function step1_SimpleAuth() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (CONFIG.ALLOWED_DOMAIN && !email.endsWith('@' + CONFIG.ALLOWED_DOMAIN)) {
      return { success: false, message: `権限エラー: ${CONFIG.ALLOWED_DOMAIN} のアカウントのみ利用可能です。` };
    }
    return { success: true, email: email };
  } catch (e) { return { success: false, message: '認証エラー: ' + e.message }; }
}

// --- ステップ2: マスターデータ取得 ---
function step2_GetMasters() {
  try {
    const staff = getStaffListSafe();
    const patients = getPatientListSafe();
    return { success: true, staff: staff, patients: patients };
  } catch (e) {
    return { success: true, staff: [], patients: [], error: e.message };
  }
}

// --- ステップ3: 発注リスト取得 ---
function step3_GetOrders() {
  try {
    const staffList = getStaffListSafe();
    const orders = getActiveOrdersSafe(staffList);
    return { success: true, orders: orders };
  } catch (e) {
    return { success: true, orders: [], error: e.message };
  }
}

// --- 通知送信 ---
function sendToChat(message) {
  try {
    UrlFetchApp.fetch(CONFIG.CHAT_WEBHOOK_URL, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify({ text: message })
    });
  } catch(e) { console.error('通知エラー', e); }
}

// === データ取得関数 ===
function getStaffListSafe() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.STAFF_MASTER_ID);
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const limit = Math.min(lastRow - 1, 1000);
    const data = sheet.getRange(2, 1, limit, 2).getValues();
    return data.map(r => ({ email: String(r[0]), name: String(r[1]) })).filter(s => s.email.includes('@'));
  } catch (e) { return []; }
}

// ★ここを修正しました
function getPatientListSafe() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.PATIENT_MASTER_ID);
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const limit = lastRow - 1; // 全件取得
    
    // C列(3列目:漢字)とD列(4列目:カナ)を取得
    // getRange(row, col, numRows, numCols) -> col=3 (C列) から 2列分 (C, D)
    const data = sheet.getRange(2, 3, limit, 2).getValues();
    
    return data
      .map(r => ({ 
        name: String(r[0]), // C列: 漢字氏名
        kana: String(r[1])  // D列: カナ
      }))
      .filter(p => p.name && p.name.trim() !== '');
  } catch (e) { 
    console.error('Patient Load Error:', e);
    return []; 
  }
}

function getActiveOrdersSafe(staffList) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const limit = Math.min(lastRow - 1, 1000);
    const data = sheet.getRange(2, 1, limit, 11).getValues();
    
    const formatDateSafe = (val) => {
      if (!val) return '';
      try {
        const d = new Date(val);
        if (isNaN(d.getTime())) return '';
        return Utilities.formatDate(d, "JST", "yyyy-MM-dd");
      } catch(e) { return ''; }
    };

    const orders = data
      .filter(r => r[1] !== CONFIG.STATUS.DELIVERED)
      .map(r => {
        const staff = staffList ? staffList.find(s => s.email === r[6]) : null;
        const staffName = staff ? staff.name : r[6];
        return {
          id: r[0], status: r[1], patient: r[2], drug: r[3],
          deadline: formatDateSafe(r[4]),
          adminDate: formatDateSafe(r[5]),
          pic: staffName, picEmail: r[6], eventId: r[7],
          sortDate: r[4] ? new Date(r[4]).getTime() : 9999999999999
        };
      });

    orders.sort((a, b) => a.sortDate - b.sortDate);
    return orders;
  } catch (e) { return []; }
}

// --- 登録・編集・削除・更新 ---
function registerOrder(form) {
  const now = new Date();
  const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const event = cal.createAllDayEvent(`【${CONFIG.STATUS.PENDING}】${form.patient} / ${form.drug}`, new Date(form.deadline), {
    description: `投与予定: ${form.adminDate}\n担当: ${form.picName} (${form.picEmail})`,
    guests: form.picEmail
  });
  const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  sheet.appendRow([Utilities.getUuid(), CONFIG.STATUS.PENDING, form.patient, form.drug, form.deadline, form.adminDate, form.picEmail, event.getId(), now, '', '']);
  SpreadsheetApp.flush();
  const msg = `🟢 【新規登録】薬剤の発注予定が登録されました。\n患者: ${form.patient} 様\n薬剤: ${form.drug}\n期限: ${form.deadline}\n担当: ${form.picName}\n<users/all>`;
  sendToChat(msg);
  return { success: true };
}

function updateOrderData(form) {
  const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === form.id) {
      sheet.getRange(i + 1, 3).setValue(form.patient);
      sheet.getRange(i + 1, 4).setValue(form.drug);
      sheet.getRange(i + 1, 5).setValue(new Date(form.deadline));
      sheet.getRange(i + 1, 6).setValue(form.adminDate ? new Date(form.adminDate) : '');
      try {
        const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
        const event = cal.getEventById(data[i][7]);
        if (event) {
          event.setTitle(`【${data[i][1]}】${form.patient} / ${form.drug}`);
          event.setAllDayDate(new Date(form.deadline));
          event.setDescription(`投与予定: ${form.adminDate}\n担当: ${form.picName} (${data[i][6]})`);
        }
      } catch(e) {}
      SpreadsheetApp.flush();
      const msg = `✏️ 【修正】登録内容が変更されました。\n患者: ${form.patient} 様\n薬剤: ${form.drug}\n期限: ${form.deadline}\n<users/all>`;
      sendToChat(msg);
      return { success: true };
    }
  }
  return { success: false, message: 'データが見つかりません' };
}

function deleteOrder(id, patientName, drugName, deleterName) {
  const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      try {
        const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
        const event = cal.getEventById(data[i][7]);
        if (event) event.deleteEvent();
      } catch(e) {}
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      const msg = `🗑️ 【削除】以下の予定が削除されました。\n患者: ${patientName} 様\n薬剤: ${drugName}\n実行者: ${deleterName}\n<users/all>`;
      sendToChat(msg);
      return { success: true };
    }
  }
  return { success: false, message: 'データが見つかりません' };
}

function updateStatus(id, newStatus, confirmPerson = '', updaterName = '') {
  const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const row = data[i];
      sheet.getRange(i + 1, 2).setValue(newStatus);
      let msg = "";
      if (newStatus === CONFIG.STATUS.ORDERED) {
        msg = `🟠 【発注しました！】\n以下の発注処理が完了しました。\n患者: ${row[2]} 様\n薬剤: ${row[3]}\n担当: ${updaterName}\n<users/all>`;
      } 
      else if (newStatus === CONFIG.STATUS.DELIVERED) {
        sheet.getRange(i + 1, 10).setValue(now);
        sheet.getRange(i + 1, 11).setValue(confirmPerson);
        msg = `🔵 【納品されました！】\n在庫を確認しました。\n患者: ${row[2]} 様\n薬剤: ${row[3]}\n確認者: ${confirmPerson}\n<users/all>`;
      }
      try {
        const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
        const event = cal.getEventById(row[7]);
        if (event) event.setTitle(event.getTitle().replace(/【.*】/, `【${newStatus}】`));
      } catch(e) {}
      SpreadsheetApp.flush();
      if (msg) sendToChat(msg);
      return { success: true };
    }
  }
  return { success: false, message: 'データが見つかりません' };
}

function hourlyAlertTask() {
  const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const currentHour = now.getHours();
  // 22時〜6時は停止
  if (currentHour >= 22 || currentHour < 6) return;

  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  data.slice(1).forEach(row => {
    if (row[1] === CONFIG.STATUS.PENDING) {
      const deadline = new Date(row[4]);
      const diffTime = deadline.getTime() - today.getTime();
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      let msg = "";
      if (diffDays >= 1 && diffDays <= 3) {
        if (currentHour === 9) msg = `🟡 【リマインド】発注期限が近づいています（あと${diffDays}日）\n患者: ${row[2]} 様 / 薬剤: ${row[3]}\n<users/all>`;
      }
      else if (diffDays === 0) {
        if ([9, 12, 16].includes(currentHour)) msg = `🟠 【本日発注日】今日が発注期限です！忘れていませんか？\n患者: ${row[2]} 様 / 薬剤: ${row[3]}\n<users/all>`;
      }
      else if (diffDays < 0) {
        msg = `🔴 【緊急：発注超過】至急発注してください！！\n患者: ${row[2]} 様 / 薬剤: ${row[3]}\n<users/all>`;
      }
      if (msg) sendToChat(msg);
    }
  });
}
