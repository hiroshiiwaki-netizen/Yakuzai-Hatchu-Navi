/**
 * 薬剤発注ナビ v25.0 - 発注データ管理
 * Orders.gs
 * 最終更新: 2026-01-16
 * 
 * 発注データのCRUD操作、一括操作、担当者変更
 * v25.0: 定期発注機能追加
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
  // v25.0: 16列まで取得（定期発注情報含む）
  const data = sheet.getRange(2, 1, limit, 16).getValues();

  const orders = data
    // v25.0: Template（定期親）、キャンセル、納品済を除外
    .filter(r => r[CONFIG.COL.STATUS] !== CONFIG.STATUS.DELIVERED 
              && r[CONFIG.COL.STATUS] !== CONFIG.STATUS.TEMPLATE
              && r[CONFIG.COL.STATUS] !== CONFIG.STATUS.CANCELLED
              && r[CONFIG.COL.IS_CANCELLED] !== true
              && r[CONFIG.COL.IS_CANCELLED] !== 'TRUE')
    .map(r => {
      const staffName = resolveStaffName(r[CONFIG.COL.PIC_EMAIL], staffList);
      return {
        id: r[CONFIG.COL.ID],
        status: r[CONFIG.COL.STATUS],
        patient: r[CONFIG.COL.PATIENT],
        drug: r[CONFIG.COL.DRUG],
        deadline: formatDateSafe(r[CONFIG.COL.DEADLINE]),
        adminDate: formatDateSafe(r[CONFIG.COL.ADMIN_DATE]),
        pic: staffName,
        picEmail: r[CONFIG.COL.PIC_EMAIL],
        eventId: r[CONFIG.COL.EVENT_ID],
        sortDate: r[CONFIG.COL.DEADLINE] ? new Date(r[CONFIG.COL.DEADLINE]).getTime() : 9999999999999,
        // v25.0: 定期発注情報
        recurrenceType: r[CONFIG.COL.RECURRENCE_TYPE] || null,
        parentOrderId: r[CONFIG.COL.PARENT_ORDER_ID] || null,
        isRecurring: !!r[CONFIG.COL.PARENT_ORDER_ID]
      };
    });

  orders.sort((a, b) => a.sortDate - b.sortDate);
  return orders;
}

// ===== 登録・更新・削除 =====

/**
 * 新規登録
 * @param {Object} form - フォームデータ
 * @param {boolean} form.isRecurring - 定期発注フラグ (v25.0)
 * @param {string} form.recurrenceType - 定期パターン (v25.0)
 * @param {string} form.recurrenceValue - 定期設定値 (v25.0)
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
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    // v25.0: 定期発注の場合
    if (form.isRecurring && form.recurrenceType) {
      return registerRecurringOrder(form, ss, sheet, now);
    }
    
    // 通常の単発発注
    const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    const event = cal.createAllDayEvent(
      `【${CONFIG.STATUS.PENDING}】${form.patient} / ${form.drug}`,
      new Date(form.deadline),
      {
        description: `投与予定: ${form.adminDate || '未設定'}\n担当: ${form.picName} (${form.picEmail})`,
        guests: form.picEmail
      }
    );
    
    const newId = Utilities.getUuid();
    
    // v25.0: 16列分のデータを追加（L-P列は空）
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
      '',  // 納品確認日時
      '',  // 納品確認者
      '',  // RECURRENCE_TYPE
      '',  // RECURRENCE_VALUE
      '',  // PARENT_ORDER_ID
      '',  // IS_CANCELLED
      ''   // SERIES_CANCELLED
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
 * v25.0: 定期発注の登録
 * @param {Object} form - フォームデータ
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {Sheet} sheet - シート
 * @param {Date} now - 現在日時
 * @returns {Object} 処理結果
 */
function registerRecurringOrder(form, ss, sheet, now) {
  try {
    // 1. 親レコード（Template）を作成
    const templateId = Utilities.getUuid();
    
    sheet.appendRow([
      templateId,
      CONFIG.STATUS.TEMPLATE,  // ステータス: Template
      sanitizeInput(form.patient),
      sanitizeInput(form.drug),
      form.deadline,  // 次回生成基準日として使用
      form.adminDate || '',
      form.picEmail,
      '',  // イベントIDなし (Templateはカレンダー不要)
      now,
      '',  // 納品確認日時
      '',  // 納品確認者
      form.recurrenceType,    // L: RECURRENCE_TYPE
      form.recurrenceValue,   // M: RECURRENCE_VALUE
      '',                      // N: PARENT_ORDER_ID (親自身は空)
      '',                      // O: IS_CANCELLED
      ''                       // P: SERIES_CANCELLED
    ]);
    
    // 2. 初回の子レコードを作成
    const firstChildId = createChildOrderFromTemplate(
      sheet, templateId, form, new Date(form.deadline), now
    );
    
    SpreadsheetApp.flush();
    
    // ログ記録
    writeLog(LOG_ACTIONS.REGISTER, templateId, 
      `定期発注登録: ${form.patient}, ${form.drug}, パターン: ${form.recurrenceType}`, form.picEmail);
    
    // 通知
    notifyNewOrder(form.patient, form.drug, form.deadline, form.picName);
    
    return createSuccessResponse({ 
      templateId: templateId,
      firstOrderId: firstChildId,
      isRecurring: true
    });
    
  } catch (e) {
    console.error('定期発注登録エラー:', e);
    return createErrorResponse('定期発注の登録に失敗しました: ' + e.message, CONFIG.ERROR_CODES.UNKNOWN_ERROR);
  }
}

/**
 * v25.0: Templateから子レコードを作成
 * @param {Sheet} sheet - シート
 * @param {string} parentId - 親Template ID
 * @param {Object} form - フォームデータ
 * @param {Date} deadline - 発注期限日
 * @param {Date} now - 現在日時
 * @returns {string} 作成した子レコードのID
 */
function createChildOrderFromTemplate(sheet, parentId, form, deadline, now) {
  const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const event = cal.createAllDayEvent(
    `【${CONFIG.STATUS.PENDING}】🔄 ${form.patient} / ${form.drug}`,
    deadline,
    {
      description: `投与予定: ${form.adminDate || '未設定'}\n担当: ${form.picName} (${form.picEmail})\n※定期発注`,
      guests: form.picEmail
    }
  );
  
  const childId = Utilities.getUuid();
  
  sheet.appendRow([
    childId,
    CONFIG.STATUS.PENDING,
    sanitizeInput(form.patient),
    sanitizeInput(form.drug),
    deadline,
    form.adminDate || '',
    form.picEmail,
    event.getId(),
    now,
    '',  // 納品確認日時
    '',  // 納品確認者
    '',  // RECURRENCE_TYPE (子は空)
    '',  // RECURRENCE_VALUE (子は空)
    parentId,  // N: PARENT_ORDER_ID
    '',        // O: IS_CANCELLED
    ''         // P: SERIES_CANCELLED
  ]);
  
  return childId;
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

// ===== v25.0: 定期発注 自動生成・キャンセル =====

/**
 * v25.0: 定期発注レコードを自動生成（日次バッチ）
 * @returns {Object} 処理結果
 */
function generateRecurringOrders() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return createSuccessResponse({ generated: 0 });
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    const today = new Date();
    const generateLimit = new Date();
    generateLimit.setDate(today.getDate() + CONFIG.RECURRENCE.GENERATE_DAYS_AHEAD);
    
    let generatedCount = 0;
    const existingOrders = new Set();
    
    // 既存の子レコード（PARENT_ORDER_ID + 発注期限日）を収集
    data.forEach(row => {
      const parentId = row[CONFIG.COL.PARENT_ORDER_ID];
      const deadline = row[CONFIG.COL.DEADLINE];
      if (parentId && deadline) {
        const key = `${parentId}_${formatDateKey(deadline)}`;
        existingOrders.add(key);
      }
    });
    
    // Templateレコードを処理
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      
      // Templateでなければスキップ
      if (row[CONFIG.COL.STATUS] !== CONFIG.STATUS.TEMPLATE) continue;
      
      // シリーズキャンセル済みならスキップ
      if (row[CONFIG.COL.SERIES_CANCELLED] === true || row[CONFIG.COL.SERIES_CANCELLED] === 'TRUE') {
        continue;
      }
      
      const templateId = row[CONFIG.COL.ID];
      const recurrenceType = row[CONFIG.COL.RECURRENCE_TYPE];
      const recurrenceValue = row[CONFIG.COL.RECURRENCE_VALUE];
      const baseDate = row[CONFIG.COL.DEADLINE] ? new Date(row[CONFIG.COL.DEADLINE]) : today;
      
      // 次回以降の発注日を計算
      const nextDates = calculateNextDates(recurrenceType, recurrenceValue, baseDate, generateLimit);
      
      for (const nextDate of nextDates) {
        const key = `${templateId}_${formatDateKey(nextDate)}`;
        
        // 既に存在していればスキップ
        if (existingOrders.has(key)) continue;
        
        // 新規子レコードを作成
        const form = {
          patient: row[CONFIG.COL.PATIENT],
          drug: row[CONFIG.COL.DRUG],
          adminDate: row[CONFIG.COL.ADMIN_DATE],
          picEmail: row[CONFIG.COL.PIC_EMAIL],
          picName: resolveStaffName(row[CONFIG.COL.PIC_EMAIL])
        };
        
        createChildOrderFromTemplate(sheet, templateId, form, nextDate, new Date());
        existingOrders.add(key);
        generatedCount++;
      }
    }
    
    SpreadsheetApp.flush();
    
    if (generatedCount > 0) {
      console.log(`定期発注 ${generatedCount}件を生成しました`);
    }
    
    return createSuccessResponse({ generated: generatedCount });
    
  } catch (e) {
    console.error('定期発注生成エラー:', e);
    return createErrorResponse('定期発注の生成に失敗しました: ' + e.message, CONFIG.ERROR_CODES.UNKNOWN_ERROR);
  }
}

/**
 * v25.0: キャンセル処理（定期発注対応）
 * @param {string} id - キャンセル対象ID
 * @param {string} cancelType - 'single' または 'series'
 * @param {string} cancelerName - キャンセル実行者名
 * @returns {Object} 処理結果
 */
function cancelOrder(id, cancelType, cancelerName) {
  if (!id) {
    return createErrorResponse('IDは必須です', CONFIG.ERROR_CODES.VALIDATION_ERROR);
  }
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][CONFIG.COL.ID] === id) {
        const row = data[i];
        const parentId = row[CONFIG.COL.PARENT_ORDER_ID];
        
        // 1. 対象レコードをキャンセル
        sheet.getRange(i + 1, CONFIG.COL.STATUS + 1).setValue(CONFIG.STATUS.CANCELLED);
        sheet.getRange(i + 1, CONFIG.COL.IS_CANCELLED + 1).setValue('TRUE');
        
        // カレンダーイベントを削除
        try {
          const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
          const event = cal.getEventById(row[CONFIG.COL.EVENT_ID]);
          if (event) event.deleteEvent();
        } catch(e) {
          console.warn('カレンダー削除エラー:', e);
        }
        
        // 2. シリーズキャンセルの場合は親も更新
        if (cancelType === 'series' && parentId) {
          for (let j = 1; j < data.length; j++) {
            if (data[j][CONFIG.COL.ID] === parentId) {
              sheet.getRange(j + 1, CONFIG.COL.SERIES_CANCELLED + 1).setValue('TRUE');
              break;
            }
          }
        }
        
        SpreadsheetApp.flush();
        
        // ログ記録
        const logDetail = cancelType === 'series' 
          ? `シリーズキャンセル: ${row[CONFIG.COL.PATIENT]}, ${row[CONFIG.COL.DRUG]}`
          : `単独キャンセル: ${row[CONFIG.COL.PATIENT]}, ${row[CONFIG.COL.DRUG]}`;
        writeLog(LOG_ACTIONS.DELETE, id, logDetail, cancelerName);
        
        return createSuccessResponse({ 
          cancelled: true, 
          cancelType: cancelType,
          seriesCancelled: cancelType === 'series'
        });
      }
    }
    
    return createErrorResponse('データが見つかりません', CONFIG.ERROR_CODES.NOT_FOUND);
    
  } catch (e) {
    console.error('キャンセルエラー:', e);
    return createErrorResponse('キャンセルに失敗しました: ' + e.message, CONFIG.ERROR_CODES.UNKNOWN_ERROR);
  }
}

/**
 * v25.0: 次回発注日を計算
 * @param {string} recurrenceType - 定期パターン
 * @param {string} recurrenceValue - 設定値
 * @param {Date} baseDate - 基準日
 * @param {Date} limitDate - 生成上限日
 * @returns {Date[]} 次回発注日の配列
 */
function calculateNextDates(recurrenceType, recurrenceValue, baseDate, limitDate) {
  const dates = [];
  let currentDate = new Date(baseDate);
  
  // 過去の日付なら今日から開始
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  if (currentDate < today) {
    currentDate = today;
  }
  
  switch (recurrenceType) {
    case CONFIG.RECURRENCE_PATTERNS.WEEKLY: {
      // recurrenceValue: 曜日(0-6)
      const targetDay = parseInt(recurrenceValue, 10);
      while (currentDate <= limitDate) {
        if (currentDate.getDay() === targetDay && currentDate >= today) {
          dates.push(new Date(currentDate));
        }
        currentDate.setDate(currentDate.getDate() + 1);
      }
      break;
    }
    
    case CONFIG.RECURRENCE_PATTERNS.BIWEEKLY: {
      // recurrenceValue: "曜日,基準日" (例: "1,2026-01-01")
      const parts = recurrenceValue.split(',');
      const targetDay = parseInt(parts[0], 10);
      const referenceDate = parts[1] ? new Date(parts[1]) : baseDate;
      
      while (currentDate <= limitDate) {
        if (currentDate.getDay() === targetDay && currentDate >= today) {
          // 基準日からの週数が偶数週かチェック
          const weekDiff = Math.floor((currentDate - referenceDate) / (7 * 24 * 60 * 60 * 1000));
          if (weekDiff % 2 === 0) {
            dates.push(new Date(currentDate));
          }
        }
        currentDate.setDate(currentDate.getDate() + 1);
      }
      break;
    }
    
    case CONFIG.RECURRENCE_PATTERNS.MONTHLY_DATE: {
      // recurrenceValue: 日付(1-31)
      const targetDate = parseInt(recurrenceValue, 10);
      while (currentDate <= limitDate) {
        const daysInMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 0).getDate();
        const actualDate = Math.min(targetDate, daysInMonth);
        const candidateDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), actualDate);
        
        if (candidateDate >= today && candidateDate <= limitDate && !dates.some(d => d.getTime() === candidateDate.getTime())) {
          dates.push(candidateDate);
        }
        
        // 次月へ
        currentDate.setMonth(currentDate.getMonth() + 1);
        currentDate.setDate(1);
      }
      break;
    }
    
    case CONFIG.RECURRENCE_PATTERNS.MONTHLY_WEEK: {
      // recurrenceValue: "第N-曜日" (例: "2-3" = 第2水曜日)
      const parts = recurrenceValue.split('-');
      const weekNum = parseInt(parts[0], 10);  // 第N週
      const targetDay = parseInt(parts[1], 10); // 曜日
      
      while (currentDate <= limitDate) {
        const candidateDate = getNthWeekdayOfMonth(currentDate.getFullYear(), currentDate.getMonth(), weekNum, targetDay);
        
        if (candidateDate && candidateDate >= today && candidateDate <= limitDate) {
          if (!dates.some(d => d.getTime() === candidateDate.getTime())) {
            dates.push(candidateDate);
          }
        }
        
        // 次月へ
        currentDate.setMonth(currentDate.getMonth() + 1);
        currentDate.setDate(1);
      }
      break;
    }
  }
  
  return dates;
}

/**
 * v25.0: 月の第N曜日を取得
 * @param {number} year - 年
 * @param {number} month - 月 (0-11)
 * @param {number} weekNum - 第N週 (1-5)
 * @param {number} dayOfWeek - 曜日 (0-6)
 * @returns {Date|null} 該当日またはnull
 */
function getNthWeekdayOfMonth(year, month, weekNum, dayOfWeek) {
  const firstDay = new Date(year, month, 1);
  const firstDayOfWeek = firstDay.getDay();
  
  // 最初のtarget曜日の日付
  let firstTargetDay = 1 + ((7 + dayOfWeek - firstDayOfWeek) % 7);
  
  // 第N週の日付
  const targetDay = firstTargetDay + (weekNum - 1) * 7;
  
  // 月内かチェック
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  if (targetDay > daysInMonth) {
    return null;
  }
  
  return new Date(year, month, targetDay);
}

/**
 * v25.0: 日付キー生成（重複チェック用）
 * @param {Date} date - 日付
 * @returns {string} YYYY-MM-DD形式の文字列
 */
function formatDateKey(date) {
  if (!date) return '';
  const d = new Date(date);
  return Utilities.formatDate(d, 'JST', 'yyyy-MM-dd');
}
