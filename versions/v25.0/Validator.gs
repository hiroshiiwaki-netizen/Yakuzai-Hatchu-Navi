/**
 * 薬剤発注ナビ v24.0 - バリデーション
 * Validator.gs
 * 最終更新: 2026-01-02
 * 
 * 入力値の検証とサニタイズ処理
 */

/**
 * 発注フォームのバリデーション
 * @param {Object} form - フォームデータ
 * @returns {Object} { valid: boolean, errors: string[] }
 */
function validateOrderForm(form) {
  const errors = [];
  
  // 患者名: 必須、最大50文字
  if (!form.patient || !form.patient.trim()) {
    errors.push('患者名は必須です');
  } else if (form.patient.length > CONFIG.VALIDATION.MAX_PATIENT_NAME_LENGTH) {
    errors.push(`患者名は${CONFIG.VALIDATION.MAX_PATIENT_NAME_LENGTH}文字以内で入力してください`);
  }
  
  // 薬剤名: 最大100文字（任意）
  if (form.drug && form.drug.length > CONFIG.VALIDATION.MAX_DRUG_NAME_LENGTH) {
    errors.push(`薬剤名は${CONFIG.VALIDATION.MAX_DRUG_NAME_LENGTH}文字以内で入力してください`);
  }
  
  // 発注期限日: 必須、日付形式
  if (!form.deadline) {
    errors.push('発注期限日は必須です');
  } else if (!isValidDateString(form.deadline)) {
    errors.push('発注期限日の形式が正しくありません (yyyy-MM-dd)');
  }
  
  // 投与予定日: 日付形式（任意）
  if (form.adminDate && !isValidDateString(form.adminDate)) {
    errors.push('投与予定日の形式が正しくありません (yyyy-MM-dd)');
  }
  
  // 担当者メール: 必須、メール形式
  if (!form.picEmail) {
    errors.push('担当者メールは必須です');
  } else if (!isValidEmail(form.picEmail)) {
    errors.push('担当者メールの形式が正しくありません');
  }
  
  return {
    valid: errors.length === 0,
    errors: errors
  };
}

/**
 * 更新フォームのバリデーション
 * @param {Object} form - フォームデータ
 * @returns {Object} { valid: boolean, errors: string[] }
 */
function validateUpdateForm(form) {
  const errors = [];
  
  // ID: 必須
  if (!form.id) {
    errors.push('IDは必須です');
  }
  
  // 患者名: 必須、最大50文字
  if (!form.patient || !form.patient.trim()) {
    errors.push('患者名は必須です');
  } else if (form.patient.length > CONFIG.VALIDATION.MAX_PATIENT_NAME_LENGTH) {
    errors.push(`患者名は${CONFIG.VALIDATION.MAX_PATIENT_NAME_LENGTH}文字以内で入力してください`);
  }
  
  // 発注期限日: 必須、日付形式
  if (!form.deadline) {
    errors.push('発注期限日は必須です');
  } else if (!isValidDateString(form.deadline)) {
    errors.push('発注期限日の形式が正しくありません');
  }
  
  return {
    valid: errors.length === 0,
    errors: errors
  };
}

/**
 * 日付文字列の検証 (yyyy-MM-dd)
 * @param {string} dateStr - 日付文字列
 * @returns {boolean}
 */
function isValidDateString(dateStr) {
  if (!dateStr) return false;
  const regex = /^\d{4}-\d{2}-\d{2}$/;
  if (!regex.test(dateStr)) return false;
  
  const date = new Date(dateStr);
  return !isNaN(date.getTime());
}

/**
 * メールアドレスの検証
 * @param {string} email - メールアドレス
 * @returns {boolean}
 */
function isValidEmail(email) {
  if (!email) return false;
  const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regex.test(email);
}

/**
 * 入力値のサニタイズ
 * @param {string} str - 入力文字列
 * @returns {string} サニタイズ済み文字列
 */
function sanitizeInput(str) {
  if (!str) return '';
  return String(str)
    .trim()
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/**
 * 日付を安全にフォーマット
 * @param {*} val - 日付値
 * @returns {string} yyyy-MM-dd形式の文字列または空文字
 */
function formatDateSafe(val) {
  if (!val) return '';
  try {
    const d = new Date(val);
    if (isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, "JST", "yyyy-MM-dd");
  } catch(e) {
    return '';
  }
}
