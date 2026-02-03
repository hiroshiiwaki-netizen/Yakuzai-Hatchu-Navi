/**
 * 薬剤発注ナビ v24.0 - 設定ファイル
 * Config.gs
 * 最終更新: 2026-01-02
 * 
 * アプリケーション全体の設定定数を管理
 */

const CONFIG = {
  // === アプリケーション基本設定 ===
  APP_NAME: '薬剤発注ナビ',
  VERSION: '24.0',
  
  // === シート設定 ===
  SHEET_NAME: 'シート1',
  LOG_SHEET_NAME: '操作ログ',
  
  // === アクセス制御 ===
  ALLOWED_DOMAIN: 'nhw.jp',
  
  // === キャッシュ設定 ===
  CACHE_TTL_SECONDS: 600, // 10分
  CACHE_KEY_STAFF: 'master_staff',
  CACHE_KEY_PATIENTS: 'master_patients',
  
  // === パフォーマンス設定 ===
  MAX_PATIENT_FETCH: 99999, // 患者マスター取得上限（無制限）
  
  // === 外部連携ID ===
  CALENDAR_ID: 'c_fbaa34d73eb9d9fada0ce4f25fcaefa4bcc3ba7626d6194d50c1778fef9d0244@group.calendar.google.com',
  CHAT_WEBHOOK_URL: 'https://chat.googleapis.com/v1/spaces/AAQAdWFuXFM/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=DTkSi3XOT1oyoM14hhEryyYOaWbnbnz4dTDtIqAzKOA',
  
  // === スプレッドシートID ===
  STAFF_MASTER_ID: '1wR-ht-2MFrf2NQFPef2Xol3iQ7qIwXsQv_9T2IImoTM',
  PATIENT_MASTER_ID: '1zD7lIxWrMEzma9GDP0Yltp_YG56pOxVUsYYPNJ3a-nE',
  DATA_DB_ID: '1ffuhDYZSts3u6YN0vzQX_CwVXWDqRLD2zn3vkFbqtxQ',

  // === ステータス定義 ===
  STATUS: {
    PENDING: '未発注',
    ORDERED: '発注済',
    DELIVERED: '納品済'
  },
  
  // === バリデーション設定 ===
  VALIDATION: {
    MAX_PATIENT_NAME_LENGTH: 50,
    MAX_DRUG_NAME_LENGTH: 100,
    MAX_STAFF_FETCH: 1000,
    MAX_ORDER_FETCH: 1000
  },
  
  // === エラーコード ===
  ERROR_CODES: {
    AUTH_FAILED: 'AUTH_FAILED',
    DOMAIN_NOT_ALLOWED: 'DOMAIN_NOT_ALLOWED',
    VALIDATION_ERROR: 'VALIDATION_ERROR',
    NOT_FOUND: 'NOT_FOUND',
    CALENDAR_ERROR: 'CALENDAR_ERROR',
    SHEET_ERROR: 'SHEET_ERROR',
    UNKNOWN_ERROR: 'UNKNOWN_ERROR'
  }
};

/**
 * 統一レスポンス生成: 成功
 */
function createSuccessResponse(data = {}) {
  return {
    success: true,
    ...data
  };
}

/**
 * 統一レスポンス生成: 失敗
 */
function createErrorResponse(message, code = CONFIG.ERROR_CODES.UNKNOWN_ERROR, details = null) {
  const response = {
    success: false,
    message: message,
    code: code
  };
  if (details) {
    response.details = details;
  }
  return response;
}
