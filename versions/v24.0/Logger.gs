/**
 * 薬剤発注ナビ v24.0 - ログ機能
 * Logger.gs
 * 最終更新: 2026-01-02
 * 
 * 操作ログの記録と取得
 * 
 * ログシートカラム構成:
 * A: 日時 | B: 操作 | C: 対象ID | D: 詳細 | E: 実行者
 */

/**
 * 操作ログを書き込む
 * @param {string} action - 操作種別 (登録/更新/削除/ステータス変更)
 * @param {string} targetId - 対象データID
 * @param {string} details - 詳細情報 (JSON文字列可)
 * @param {string} userEmail - 実行者メールアドレス
 */
function writeLog(action, targetId, details, userEmail) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
    
    // ログシートが存在しない場合はスキップ（エラーにしない）
    if (!logSheet) {
      console.warn('操作ログシートが見つかりません: ' + CONFIG.LOG_SHEET_NAME);
      return;
    }
    
    const now = new Date();
    const timestamp = Utilities.formatDate(now, "JST", "yyyy-MM-dd HH:mm:ss");
    
    logSheet.appendRow([
      timestamp,
      action,
      targetId || '',
      details || '',
      userEmail || ''
    ]);
    
  } catch (e) {
    // ログ書き込みエラーは握りつぶす（本体処理に影響させない）
    console.error('ログ書き込みエラー:', e);
  }
}

/**
 * 直近の操作ログを取得
 * @param {number} limit - 取得件数 (デフォルト50)
 * @returns {Array} ログ配列
 */
function getRecentLogs(limit = 50) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
    
    if (!logSheet) {
      return createErrorResponse('操作ログシートが見つかりません', CONFIG.ERROR_CODES.SHEET_ERROR);
    }
    
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) {
      return createSuccessResponse({ logs: [] });
    }
    
    // 最新のログから取得
    const fetchCount = Math.min(lastRow - 1, limit);
    const startRow = Math.max(2, lastRow - fetchCount + 1);
    
    const data = logSheet.getRange(startRow, 1, fetchCount, 5).getValues();
    
    const logs = data.reverse().map(row => ({
      timestamp: row[0],
      action: row[1],
      targetId: row[2],
      details: row[3],
      user: row[4]
    }));
    
    return createSuccessResponse({ logs: logs });
    
  } catch (e) {
    console.error('ログ取得エラー:', e);
    return createErrorResponse('ログの取得に失敗しました: ' + e.message, CONFIG.ERROR_CODES.SHEET_ERROR);
  }
}

/**
 * ログアクション定義
 */
const LOG_ACTIONS = {
  REGISTER: '登録',
  UPDATE: '更新',
  DELETE: '削除',
  STATUS_CHANGE: 'ステータス変更',
  BULK_STATUS_CHANGE: '一括ステータス変更',
  BULK_DELETE: '一括削除',
  PIC_CHANGE: '担当者変更'
};

/**
 * 期間指定で操作ログを取得（フロントエンド用）
 * @param {number} days - 過去何日分を取得するか
 * @returns {Object} { success: boolean, logs: Array }
 */
function getOperationLogs(days) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.DATA_DB_ID);
    const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
    
    if (!logSheet) {
      return createErrorResponse('操作ログシートが見つかりません', CONFIG.ERROR_CODES.SHEET_ERROR);
    }
    
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) {
      return createSuccessResponse({ logs: [] });
    }
    
    // 全データ取得
    const data = logSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    
    // 期間フィルタリング
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - days);
    cutoffDate.setHours(0, 0, 0, 0);
    
    const logs = data
      .filter(row => {
        if (!row[0]) return false;
        const logDate = new Date(row[0]);
        return logDate >= cutoffDate;
      })
      .map(row => {
        const logDate = new Date(row[0]);
        return {
          timestamp: Utilities.formatDate(logDate, "JST", "MM/dd HH:mm"),
          action: row[1] || '',
          targetId: row[2] || '',
          details: row[3] || '',
          operator: row[4] || ''
        };
      })
      .reverse(); // 新しい順
    
    return createSuccessResponse({ logs: logs });
    
  } catch (e) {
    console.error('操作ログ取得エラー:', e);
    return createErrorResponse('操作ログの取得に失敗しました: ' + e.message, CONFIG.ERROR_CODES.SHEET_ERROR);
  }
}
