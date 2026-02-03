/**
 * 薬剤発注ナビ v24.0 - 通知機能
 * Notify.gs
 * 最終更新: 2026-01-02
 * 
 * Google Chat通知と定期リマインド
 */

/**
 * Google Chatに通知を送信
 * @param {string} message - 送信するメッセージ
 * @returns {boolean} 送信成功可否
 */
function sendToChat(message) {
  try {
    UrlFetchApp.fetch(CONFIG.CHAT_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: message }),
      muteHttpExceptions: true
    });
    return true;
  } catch(e) {
    console.error('Chat通知エラー:', e);
    return false;
  }
}

/**
 * 新規登録通知
 */
function notifyNewOrder(patient, drug, deadline, picName) {
  const msg = `🟢 【新規登録】薬剤の発注予定が登録されました。
患者: ${patient} 様
薬剤: ${drug}
期限: ${deadline}
担当: ${picName}
<users/all>`;
  sendToChat(msg);
}

/**
 * 更新通知
 */
function notifyOrderUpdated(patient, drug, deadline) {
  const msg = `✏️ 【修正】登録内容が変更されました。
患者: ${patient} 様
薬剤: ${drug}
期限: ${deadline}
<users/all>`;
  sendToChat(msg);
}

/**
 * 削除通知
 */
function notifyOrderDeleted(patient, drug, deleterName) {
  const msg = `🗑️ 【削除】以下の予定が削除されました。
患者: ${patient} 様
薬剤: ${drug}
実行者: ${deleterName}
<users/all>`;
  sendToChat(msg);
}

/**
 * 発注完了通知
 */
function notifyOrdered(patient, drug, updaterName) {
  const msg = `🟠 【発注しました！】
以下の発注処理が完了しました。
患者: ${patient} 様
薬剤: ${drug}
担当: ${updaterName}
<users/all>`;
  sendToChat(msg);
}

/**
 * 納品完了通知
 */
function notifyDelivered(patient, drug, confirmPerson) {
  const msg = `🔵 【納品されました！】
在庫を確認しました。
患者: ${patient} 様
薬剤: ${drug}
確認者: ${confirmPerson}
<users/all>`;
  sendToChat(msg);
}

/**
 * 担当者変更通知
 */
function notifyPicChanged(patient, drug, oldPic, newPic, changerName) {
  const msg = `👤 【担当者変更】
以下の発注の担当者が変更されました。
患者: ${patient} 様
薬剤: ${drug}
旧担当: ${oldPic}
新担当: ${newPic}
変更者: ${changerName}
<users/all>`;
  sendToChat(msg);
}

/**
 * 一括操作通知
 */
function notifyBulkOperation(action, count, updaterName) {
  const msg = `📋 【一括${action}】
${count}件のデータを${action}しました。
実行者: ${updaterName}
<users/all>`;
  sendToChat(msg);
}

/**
 * 定期リマインドタスク (毎時トリガー用)
 * 22時〜6時は停止
 */
function hourlyAlertTask() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      console.error('シートが見つかりません');
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const currentHour = now.getHours();
    
    // 22時〜6時は停止
    if (currentHour >= 22 || currentHour < 6) {
      console.log('夜間のためリマインドをスキップ');
      return;
    }

    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    
    data.slice(1).forEach(row => {
      // 未発注のみ対象
      if (row[1] !== CONFIG.STATUS.PENDING) return;
      
      const deadline = new Date(row[4]);
      const diffTime = deadline.getTime() - today.getTime();
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      
      let msg = "";
      
      // 期限1〜3日前: 9時に通知
      if (diffDays >= 1 && diffDays <= 3) {
        if (currentHour === 9) {
          msg = `🟡 【リマインド】発注期限が近づいています（あと${diffDays}日）
患者: ${row[2]} 様 / 薬剤: ${row[3]}
<users/all>`;
        }
      }
      // 期限当日: 9時, 12時, 16時に通知
      else if (diffDays === 0) {
        if ([9, 12, 16].includes(currentHour)) {
          msg = `🟠 【本日発注日】今日が発注期限です！忘れていませんか？
患者: ${row[2]} 様 / 薬剤: ${row[3]}
<users/all>`;
        }
      }
      // 期限超過: 毎時通知
      else if (diffDays < 0) {
        msg = `🔴 【緊急：発注超過】至急発注してください！！
患者: ${row[2]} 様 / 薬剤: ${row[3]}
<users/all>`;
      }
      
      if (msg) {
        sendToChat(msg);
      }
    });
    
  } catch (e) {
    console.error('リマインドエラー:', e);
  }
}
