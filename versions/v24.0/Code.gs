/**
 * 薬剤発注ナビ v24.0 - エントリーポイント
 * Code.gs
 * 最終更新: 2026-01-02
 * 
 * Webアプリケーションのメインエントリーポイント
 * 認証とHTMLサービスの提供
 */

/**
 * Webアプリケーションのエントリーポイント
 */
function doGet(e) {
  const html = HtmlService.createTemplateFromFile('index').evaluate();
  return html
    .setTitle(CONFIG.APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTMLテンプレート内で別ファイルを読み込む
 * @param {string} filename - 読み込むファイル名
 * @returns {string} ファイル内容
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 認証処理
 * @returns {Object} 認証結果
 */
function step1_SimpleAuth() {
  try {
    const email = Session.getActiveUser().getEmail();
    
    // ドメイン検証
    if (CONFIG.ALLOWED_DOMAIN && !email.endsWith('@' + CONFIG.ALLOWED_DOMAIN)) {
      return createErrorResponse(
        `権限エラー: ${CONFIG.ALLOWED_DOMAIN} のアカウントのみ利用可能です。`,
        CONFIG.ERROR_CODES.DOMAIN_NOT_ALLOWED
      );
    }
    
    return createSuccessResponse({ email: email });
    
  } catch (e) {
    console.error('認証エラー:', e);
    return createErrorResponse('認証エラー: ' + e.message, CONFIG.ERROR_CODES.AUTH_FAILED);
  }
}

/**
 * 現在の年度を取得
 * @returns {number} 年度 (4月始まり)
 */
function getCurrentFiscalYear() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1; // 0-indexed to 1-indexed
  
  // 1-3月は前年度
  return month >= 4 ? year : year - 1;
}
