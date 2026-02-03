/**
 * 薬剤発注ナビ v24.0 - 発注データ管理
 * Orders.gs
 * 最終更新: 2026-01-02
 * 
 * 発注データのCRUD操作、一括操作、担当者変更
 */

// ===== データ取得 =====

/**
 * 発注リスト取得 (フロントエンド用)
 * @returns {Object} { success: boolean, orders: Array }
 */
function step3_GetOrders() {
  try {
    const staffList = getStaffListCached();
    const orders = getActiveOrders(staffList);
    return createSuccessResponse({ orders: orders });
  } catch (e) {
    console.error('発注リスト取得エラー:', e);
    return createSuccessResponse({ orders: [], error: e.message });
  }
}

/**
 * リスト更新用 (エイリアス)
 */
function getLatestOrders() {
  return step3_GetOrders();
}

/**
 * アクティブな発注データを取得（内部関数）
 * @param {Array} staffList - スタッフリスト
 * @returns {Array} 発注データ配列
 */
function getActiveOrders(staffList) {
  const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const limit = Math.min(lastRow - 1, CONFIG.VALIDATION.MAX_ORDER_FETCH);
  const data = sheet.getRange(2, 1, limit, 11).getValues();

  const orders = data
    .filter(r => r[1] !== CONFIG.STATUS.DELIVERED)
    .map(r => {
      const staffName = resolveStaffName(r[6], staffList);
      return {
        id: r[0],
        status: r[1],
        patient: r[2],
        drug: r[3],
        deadline: formatDateSafe(r[4]),
        adminDate: formatDateSafe(r[5]),
        pic: staffName,
        picEmail: r[6],
        eventId: r[7],
        sortDate: r[4] ? new Date(r[4]).getTime() : 9999999999999
      };
    });

  orders.sort((a, b) => a.sortDate - b.sortDate);
  return orders;
}

// ===== 登録・更新・削除 =====

/**
 * 新規登録
 * @param {Object} form - フォームデータ
 * @returns {Object} 処理結果
 */
function registerOrder(form) {
  // バリデーション
  const validation = validateOrderForm(form);
  if (!validation.valid) {
    return createErrorResponse(
      validation.errors.join('\n'),
      CONFIG.ERROR_CODES.VALIDATION_ERROR,
      validation.errors
    );
  }
  
  try {
    const now = new Date();
    
    // カレンダーイベント作成
    const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    const event = cal.createAllDayEvent(
      `【${CONFIG.STATUS.PENDING}】${form.patient} / ${form.drug}`,
      new Date(form.deadline),
      {
        description: `投与予定: ${form.adminDate || '未設定'}\n担当: ${form.picName} (${form.picEmail})`,
        guests: form.picEmail
      }
    );
    
    // データ登録
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const newId = Utilities.getUuid();
    
    sheet.appendRow([
      newId,
      CONFIG.STATUS.PENDING,
      sanitizeInput(form.patient),
      sanitizeInput(form.drug),
      form.deadline,
      form.adminDate || '',
      form.picEmail,
      event.getId(),
      now,
      '',
      ''
    ]);
    
    SpreadsheetApp.flush();
    
    // ログ記録
    writeLog(LOG_ACTIONS.REGISTER, newId, 
      `患者: ${form.patient}, 薬剤: ${form.drug}`, form.picEmail);
    
    // 通知
    notifyNewOrder(form.patient, form.drug, form.deadline, form.picName);
    
    return createSuccessResponse({ id: newId });
    
  } catch (e) {
    console.error('登録エラー:', e);
    return createErrorResponse('登録に失敗しました: ' + e.message, CONFIG.ERROR_CODES.UNKNOWN_ERROR);
  }
}

/**
 * データ更新
 * @param {Object} form - フォームデータ
 * @returns {Object} 処理結果
 */
function updateOrderData(form) {
  // バリデーション
  const validation = validateUpdateForm(form);
  if (!validation.valid) {
    return createErrorResponse(
      validation.errors.join('\n'),
      CONFIG.ERROR_CODES.VALIDATION_ERROR,
      validation.errors
    );
  }
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === form.id) {
        // データ更新
        sheet.getRange(i + 1, 3).setValue(sanitizeInput(form.patient));
        sheet.getRange(i + 1, 4).setValue(sanitizeInput(form.drug));
        sheet.getRange(i + 1, 5).setValue(new Date(form.deadline));
        sheet.getRange(i + 1, 6).setValue(form.adminDate ? new Date(form.adminDate) : '');
        
        // 担当者変更があれば更新
        if (form.newPicEmail && form.newPicEmail !== data[i][6]) {
          const oldPicEmail = data[i][6];
          sheet.getRange(i + 1, 7).setValue(form.newPicEmail);
          
          // ログ記録（担当者変更）
          writeLog(LOG_ACTIONS.PIC_CHANGE, form.id,
            `旧: ${oldPicEmail}, 新: ${form.newPicEmail}`, form.updaterEmail || '');
          
          // 担当者変更通知
          const oldPicName = resolveStaffName(oldPicEmail);
          const newPicName = resolveStaffName(form.newPicEmail);
          notifyPicChanged(form.patient, form.drug, oldPicName, newPicName, form.picName);
        }
        
        // カレンダー更新
        try {
          const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
          const event = cal.getEventById(data[i][7]);
          if (event) {
            event.setTitle(`【${data[i][1]}】${form.patient} / ${form.drug}`);
            event.setAllDayDate(new Date(form.deadline));
            const picEmail = form.newPicEmail || data[i][6];
            event.setDescription(`投与予定: ${form.adminDate || '未設定'}\n担当: ${form.picName} (${picEmail})`);
            
            // ゲスト更新
            if (form.newPicEmail && form.newPicEmail !== data[i][6]) {
              try {
                event.removeGuest(data[i][6]);
                event.addGuest(form.newPicEmail);
              } catch(ge) {
                console.warn('ゲスト更新エラー:', ge);
              }
            }
          }
        } catch(e) {
          console.warn('カレンダー更新エラー:', e);
        }
        
        SpreadsheetApp.flush();
        
        // ログ記録
        writeLog(LOG_ACTIONS.UPDATE, form.id,
          `患者: ${form.patient}, 薬剤: ${form.drug}`, form.updaterEmail || '');
        
        // 通知
        notifyOrderUpdated(form.patient, form.drug, form.deadline);
        
        return createSuccessResponse();
      }
    }
    
    return createErrorResponse('データが見つかりません', CONFIG.ERROR_CODES.NOT_FOUND);
    
  } catch (e) {
    console.error('更新エラー:', e);
    return createErrorResponse('更新に失敗しました: ' + e.message, CONFIG.ERROR_CODES.UNKNOWN_ERROR);
  }
}

/**
 * データ削除
 * @param {string} id - 対象ID
 * @param {string} patientName - 患者名
 * @param {string} drugName - 薬剤名
 * @param {string} deleterName - 削除者名
 * @returns {Object} 処理結果
 */
function deleteOrder(id, patientName, drugName, deleterName) {
  if (!id) {
    return createErrorResponse('IDは必須です', CONFIG.ERROR_CODES.VALIDATION_ERROR);
  }
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        // カレンダーイベント削除
        try {
          const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
          const event = cal.getEventById(data[i][7]);
          if (event) event.deleteEvent();
        } catch(e) {
          console.warn('カレンダー削除エラー:', e);
        }
        
        // 行削除
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        
        // ログ記録
        writeLog(LOG_ACTIONS.DELETE, id,
          `患者: ${patientName}, 薬剤: ${drugName}`, deleterName);
        
        // 通知
        notifyOrderDeleted(patientName, drugName, deleterName);
        
        return createSuccessResponse();
      }
    }
    
    return createErrorResponse('データが見つかりません', CONFIG.ERROR_CODES.NOT_FOUND);
    
  } catch (e) {
    console.error('削除エラー:', e);
    return createErrorResponse('削除に失敗しました: ' + e.message, CONFIG.ERROR_CODES.UNKNOWN_ERROR);
  }
}

/**
 * ステータス更新
 * @param {string} id - 対象ID
 * @param {string} newStatus - 新ステータス
 * @param {string} confirmPerson - 確認者 (納品済時)
 * @param {string} updaterName - 更新者名
 * @returns {Object} 処理結果
 */
function updateStatus(id, newStatus, confirmPerson = '', updaterName = '') {
  if (!id || !newStatus) {
    return createErrorResponse('ID・ステータスは必須です', CONFIG.ERROR_CODES.VALIDATION_ERROR);
  }
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        const row = data[i];
        
        // ステータス更新
        sheet.getRange(i + 1, 2).setValue(newStatus);
        
        // 納品済の場合は確認情報も記録
        if (newStatus === CONFIG.STATUS.DELIVERED) {
          sheet.getRange(i + 1, 10).setValue(now);
          sheet.getRange(i + 1, 11).setValue(confirmPerson);
        }
        
        // カレンダー更新
        try {
          const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
          const event = cal.getEventById(row[7]);
          if (event) {
            event.setTitle(event.getTitle().replace(/【.*】/, `【${newStatus}】`));
          }
        } catch(e) {
          console.warn('カレンダー更新エラー:', e);
        }
        
        SpreadsheetApp.flush();
        
        // ログ記録
        writeLog(LOG_ACTIONS.STATUS_CHANGE, id,
          `${row[1]} → ${newStatus}`, updaterName);
        
        // 通知
        if (newStatus === CONFIG.STATUS.ORDERED) {
          notifyOrdered(row[2], row[3], updaterName);
        } else if (newStatus === CONFIG.STATUS.DELIVERED) {
          notifyDelivered(row[2], row[3], confirmPerson);
        }
        
        return createSuccessResponse();
      }
    }
    
    return createErrorResponse('データが見つかりません', CONFIG.ERROR_CODES.NOT_FOUND);
    
  } catch (e) {
    console.error('ステータス更新エラー:', e);
    return createErrorResponse('ステータス更新に失敗しました: ' + e.message, CONFIG.ERROR_CODES.UNKNOWN_ERROR);
  }
}

// ===== 一括操作 =====

/**
 * 一括ステータス更新
 * @param {Array} ids - 対象ID配列
 * @param {string} newStatus - 新ステータス
 * @param {string} updaterName - 更新者名
 * @param {string} confirmPerson - 確認者 (納品済時)
 * @returns {Object} 処理結果
 */
function bulkUpdateStatus(ids, newStatus, updaterName, confirmPerson = '') {
  if (!ids || !Array.isArray(ids) || ids.length === 0) {
    return createErrorResponse('対象を選択してください', CONFIG.ERROR_CODES.VALIDATION_ERROR);
  }
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    
    let successCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      if (ids.includes(data[i][0])) {
        // ステータス更新
        sheet.getRange(i + 1, 2).setValue(newStatus);
        
        // 納品済の場合は確認情報も記録
        if (newStatus === CONFIG.STATUS.DELIVERED) {
          sheet.getRange(i + 1, 10).setValue(now);
          sheet.getRange(i + 1, 11).setValue(confirmPerson);
        }
        
        // カレンダー更新
        try {
          const event = cal.getEventById(data[i][7]);
          if (event) {
            event.setTitle(event.getTitle().replace(/【.*】/, `【${newStatus}】`));
          }
        } catch(e) {}
        
        successCount++;
      }
    }
    
    SpreadsheetApp.flush();
    
    // ログ記録
    writeLog(LOG_ACTIONS.BULK_STATUS_CHANGE, ids.join(','),
      `${successCount}件を${newStatus}に変更`, updaterName);
    
    // 通知
    notifyBulkOperation(newStatus, successCount, updaterName);
    
    return createSuccessResponse({ count: successCount });
    
  } catch (e) {
    console.error('一括ステータス更新エラー:', e);
    return createErrorResponse('一括更新に失敗しました: ' + e.message, CONFIG.ERROR_CODES.UNKNOWN_ERROR);
  }
}

/**
 * 一括削除
 * @param {Array} ids - 対象ID配列
 * @param {string} deleterName - 削除者名
 * @returns {Object} 処理結果
 */
function bulkDelete(ids, deleterName) {
  if (!ids || !Array.isArray(ids) || ids.length === 0) {
    return createErrorResponse('対象を選択してください', CONFIG.ERROR_CODES.VALIDATION_ERROR);
  }
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    
    // 削除対象の行番号を収集（後ろから削除するため逆順）
    const rowsToDelete = [];
    for (let i = 1; i < data.length; i++) {
      if (ids.includes(data[i][0])) {
        rowsToDelete.push({
          row: i + 1,
          eventId: data[i][7]
        });
      }
    }
    
    // 後ろから削除
    rowsToDelete.sort((a, b) => b.row - a.row);
    
    for (const item of rowsToDelete) {
      // カレンダー削除
      try {
        const event = cal.getEventById(item.eventId);
        if (event) event.deleteEvent();
      } catch(e) {}
      
      // 行削除
      sheet.deleteRow(item.row);
    }
    
    SpreadsheetApp.flush();
    
    // ログ記録
    writeLog(LOG_ACTIONS.BULK_DELETE, ids.join(','),
      `${rowsToDelete.length}件を削除`, deleterName);
    
    // 通知
    notifyBulkOperation('削除', rowsToDelete.length, deleterName);
    
    return createSuccessResponse({ count: rowsToDelete.length });
    
  } catch (e) {
    console.error('一括削除エラー:', e);
    return createErrorResponse('一括削除に失敗しました: ' + e.message, CONFIG.ERROR_CODES.UNKNOWN_ERROR);
  }
}

// ===== 統計データ取得 =====

/**
 * 統計データ取得
 * @param {number} fiscalYear - 年度 (例: 2025 = 2025年4月〜2026年3月)
 * @returns {Object} 統計データ
 */
function getStatistics(fiscalYear) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      return createErrorResponse('シートが見つかりません', CONFIG.ERROR_CODES.SHEET_ERROR);
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return createSuccessResponse({
        fiscalYear: fiscalYear,
        monthly: [],
        byPic: {},
        byStatus: {},
        overdueCount: 0
      });
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const staffList = getStaffListCached();
    
    // 年度の開始・終了日
    const startDate = new Date(fiscalYear, 3, 1); // 4月1日
    const endDate = new Date(fiscalYear + 1, 2, 31); // 3月31日
    const today = new Date();
    
    // 集計用オブジェクト
    const monthly = {};
    const byPic = {};
    const byStatus = {
      [CONFIG.STATUS.PENDING]: 0,
      [CONFIG.STATUS.ORDERED]: 0,
      [CONFIG.STATUS.DELIVERED]: 0
    };
    let overdueCount = 0;
    
    // 月別初期化 (4月〜翌3月)
    for (let m = 4; m <= 12; m++) {
      monthly[`${fiscalYear}-${String(m).padStart(2, '0')}`] = 0;
    }
    for (let m = 1; m <= 3; m++) {
      monthly[`${fiscalYear + 1}-${String(m).padStart(2, '0')}`] = 0;
    }
    
    data.forEach(row => {
      const registeredDate = row[8] ? new Date(row[8]) : null;
      const deadline = row[4] ? new Date(row[4]) : null;
      const status = row[1];
      const picEmail = row[6];
      
      // 年度内のデータのみ集計
      if (registeredDate && registeredDate >= startDate && registeredDate <= endDate) {
        // 月別
        const monthKey = Utilities.formatDate(registeredDate, "JST", "yyyy-MM");
        if (monthly[monthKey] !== undefined) {
          monthly[monthKey]++;
        }
        
        // 担当者別
        const picName = resolveStaffName(picEmail, staffList);
        if (!byPic[picName]) byPic[picName] = 0;
        byPic[picName]++;
        
        // ステータス別
        if (byStatus[status] !== undefined) {
          byStatus[status]++;
        }
      }
      
      // 期限超過 (未発注のみ)
      if (status === CONFIG.STATUS.PENDING && deadline && deadline < today) {
        overdueCount++;
      }
    });
    
    // 月別を配列に変換
    const monthlyArray = Object.entries(monthly).map(([month, count]) => ({
      month: month,
      count: count
    }));
    
    return createSuccessResponse({
      fiscalYear: fiscalYear,
      monthly: monthlyArray,
      byPic: byPic,
      byStatus: byStatus,
      overdueCount: overdueCount
    });
    
  } catch (e) {
    console.error('統計データ取得エラー:', e);
    return createErrorResponse('統計データの取得に失敗しました: ' + e.message, CONFIG.ERROR_CODES.UNKNOWN_ERROR);
  }
}

// ===== 納品済み履歴取得 =====

/**
 * 納品済みデータ取得
 * @param {number} days - 過去何日分を取得するか
 * @returns {Object} { success: boolean, orders: Array }
 */
function getDeliveredOrders(days) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      return createErrorResponse('シートが見つかりません', CONFIG.ERROR_CODES.SHEET_ERROR);
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return createSuccessResponse({ orders: [] });
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    const staffList = getStaffListCached();
    
    // 期間の計算
    const today = new Date();
    const cutoffDate = new Date();
    cutoffDate.setDate(today.getDate() - days);
    
    const orders = data
      .filter(r => {
        if (r[1] !== CONFIG.STATUS.DELIVERED) return false;
        const deliveredDate = r[9] ? new Date(r[9]) : null;
        return deliveredDate && deliveredDate >= cutoffDate;
      })
      .map(r => ({
        id: r[0],
        patient: r[2],
        drug: r[3],
        deadline: formatDateSafe(r[4]),
        adminDate: formatDateSafe(r[5]),
        pic: resolveStaffName(r[6], staffList),
        deliveredDate: formatDateSafe(r[9]),
        confirmedBy: r[10] || ''
      }))
      .sort((a, b) => {
        // 納品日の新しい順
        const dateA = a.deliveredDate ? new Date(a.deliveredDate).getTime() : 0;
        const dateB = b.deliveredDate ? new Date(b.deliveredDate).getTime() : 0;
        return dateB - dateA;
      });
    
    return createSuccessResponse({ orders: orders });
    
  } catch (e) {
    console.error('納品済み履歴取得エラー:', e);
    return createErrorResponse('履歴の取得に失敗しました: ' + e.message, CONFIG.ERROR_CODES.UNKNOWN_ERROR);
  }
}
