/**
 * 薬剤発注ナビ v24.0 - マスターデータ管理
 * Masters.gs
 * 最終更新: 2026-01-02
 * 
 * スタッフ・患者マスターの取得とキャッシュ管理
 */

/**
 * キャッシュ付きスタッフリスト取得
 * @returns {Array} スタッフ配列
 */
function getStaffListCached() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(CONFIG.CACHE_KEY_STAFF);
    
    if (cached) {
      console.log('スタッフマスター: キャッシュから取得');
      return JSON.parse(cached);
    }
    
    console.log('スタッフマスター: シートから取得');
    const data = getStaffListFromSheet();
    
    // キャッシュに保存 (10分)
    cache.put(CONFIG.CACHE_KEY_STAFF, JSON.stringify(data), CONFIG.CACHE_TTL_SECONDS);
    
    return data;
  } catch (e) {
    console.error('スタッフマスター取得エラー:', e);
    return [];
  }
}

/**
 * シートからスタッフリストを取得（内部関数）
 * A列:Email, B列:氏名, C列:職種, D列:所属デポ
 * @returns {Array} スタッフ配列
 */
function getStaffListFromSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.STAFF_MASTER_ID);
  const sheet = ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return [];
  
  const limit = Math.min(lastRow - 1, CONFIG.VALIDATION.MAX_STAFF_FETCH);
  // A〜D列（4列）を取得
  const data = sheet.getRange(2, 1, limit, 4).getValues();
  
  return data
    .map(r => ({
      email: String(r[0]).trim(),
      name: String(r[1]).trim(),
      role: String(r[2]).trim(),    // 職種
      dept: String(r[3]).trim()     // 所属デポ
    }))
    .filter(s => s.email.includes('@'));
}

/**
 * キャッシュ付き患者リスト取得
 * ※開発中のためキャッシュをスキップして直接取得
 * @returns {Array} 患者配列
 */
function getPatientListCached() {
  try {
    // キャッシュをスキップして直接取得（往診予定チェッカーと同じ方式）
    console.log('患者マスター: シートから直接取得');
    return getPatientListFromSheet();
  } catch (e) {
    console.error('患者マスター取得エラー:', e);
    return [];
  }
}

/**
 * シートから患者リストを取得（内部関数）
 * C列(漢字氏名)とD列(カナ)を取得
 * @returns {Array} 患者配列
 */
function getPatientListFromSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.PATIENT_MASTER_ID);
  const sheet = ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return [];
  
  // 全件取得
  const rowCount = lastRow - 1;
  
  // C列(3列目:漢字)とD列(4列目:カナ)を取得
  const data = sheet.getRange(2, 3, rowCount, 2).getValues();
  
  return data
    .map(r => ({
      name: String(r[0]).trim(),
      kana: String(r[1]).trim()
    }))
    .filter(p => p.name && p.name.length > 0);
}

/**
 * キャッシュを無効化
 * @param {string} type - 'staff' | 'patients' | 'all'
 */
function invalidateCache(type = 'all') {
  const cache = CacheService.getScriptCache();
  
  if (type === 'staff' || type === 'all') {
    cache.remove(CONFIG.CACHE_KEY_STAFF);
    console.log('スタッフキャッシュを無効化');
  }
  
  if (type === 'patients' || type === 'all') {
    cache.remove(CONFIG.CACHE_KEY_PATIENTS);
    console.log('患者キャッシュを無効化');
  }
  
  return { success: true, message: 'キャッシュを無効化しました' };
}

// ===== API関数 =====

/**
 * マスターデータ取得 (フロントエンド用)
 * @returns {Object} { success: boolean, staff: Array, patients: Array }
 */
function step2_GetMasters() {
  try {
    const staff = getStaffListCached();
    const patients = getPatientListCached();
    
    return createSuccessResponse({
      staff: staff,
      patients: patients
    });
  } catch (e) {
    console.error('マスターデータ取得エラー:', e);
    return createSuccessResponse({
      staff: [],
      patients: [],
      error: e.message
    });
  }
}

/**
 * スタッフ名をメールアドレスから解決
 * @param {string} email - メールアドレス
 * @param {Array} staffList - スタッフリスト (省略可)
 * @returns {string} スタッフ名またはメールアドレス
 */
function resolveStaffName(email, staffList = null) {
  if (!email) return '';
  
  const list = staffList || getStaffListCached();
  const staff = list.find(s => s.email === email);
  
  return staff ? staff.name : email;
}
