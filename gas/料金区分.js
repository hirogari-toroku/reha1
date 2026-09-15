// Blank policies preserve pre-existing contracts. Newly imported records require review.
const PRICING_POLICY_VERSION = "2026-09-15-v1";
const USER_PRICING_HEADER = "利用料金区分";
const STAFF_PRICING_HEADER = "報酬区分";
const TRAVEL_PRICING_HEADER = "交通費区分";
const NEW_PRICING = "新料金";
const LEGACY_PRICING = "旧料金";
const PENDING_PRICING = "要確認";

function pricingPolicy_(value) {
  const policy = String(value || "").trim();
  if (!policy || policy === LEGACY_PRICING) return LEGACY_PRICING;
  if (policy === NEW_PRICING || policy === PENDING_PRICING) return policy;
  throw new Error("料金区分が不正です: " + policy);
}

function pricingColumn_(headers, name) {
  return headers.indexOf(name);
}

function pricingTravel_(policy, distance, fallback) {
  policy = pricingPolicy_(policy);
  if (policy === PENDING_PRICING) throw new Error("交通費区分を確認してください。");
  if (policy === LEGACY_PRICING) {
    return {
      payroll: calculatePayrollTravelCost_(Number(distance) || 0, Number(fallback) || 0),
      direct: calculateDirectTravelCostForUser_(Number(distance) || 0)
    };
  }
  if (distance === "" || distance == null || !Number.isFinite(Number(distance)) || Number(distance) < 0) {
    throw new Error("新料金の片道距離を入力してください。未入力を0kmとして計算しません。");
  }
  return { payroll: Math.round(Number(distance) * 20), direct: 0 };
}

function pricingUserMap_(ss) {
  const sheet = ss.getSheetByName(USER_MASTER_SHEET_NAME);
  const map = {};
  if (!sheet) return map;
  const values = sheet.getDataRange().getValues();
  const col = pricingColumn_(values[0], USER_PRICING_HEADER);
  values.slice(1).forEach(row => {
    if (row[0]) map[normalizeName_(row[0])] = pricingPolicy_(row[col]);
  });
  return map;
}

function pricingStaffMap_(ss) {
  const sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  const map = {};
  if (!sheet) return map;
  const values = sheet.getDataRange().getValues();
  const col = pricingColumn_(values[0], STAFF_PRICING_HEADER);
  values.slice(1).forEach(row => {
    if (row[0]) map[normalizeName_(row[0])] = pricingPolicy_(row[col]);
  });
  return map;
}

function validateNewCouponPayments_(users, payments) {
  const byName = {};
  users.forEach(user => { byName[normalizeName_(user.userName)] = user; });
  Object.keys(payments).forEach(key => {
    const user = byName[key];
    if (!user || user.pricingPolicy === LEGACY_PRICING) return;
    payments[key].forEach(payment => {
      if (user.pricingPolicy === PENDING_PRICING) {
        throw new Error(user.userName + "の利用料金区分を確認してください。入金は未反映です。");
      }
      if (payment.amount > 0 && payment.amount % 37000 !== 0) {
        throw new Error(user.userName + "の入金は37,000円単位ではありません。単発・分割入金等を確認してください。回数券に自動換算しません。");
      }
    });
  });
}

function setupPricingPolicyColumnsCore_(ss) {
  const specs = [
    [USER_MASTER_SHEET_NAME, USER_PRICING_HEADER],
    [STAFF_SHEET_NAME, STAFF_PRICING_HEADER],
    [STAFF_USER_MASTER_SHEET_NAME, TRAVEL_PRICING_HEADER]
  ];
  specs.forEach(([name, header]) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(name + "がありません。");
  });
  specs.forEach(([name, header]) => {
    const sheet = ss.getSheetByName(name);
    const col = ensureHeaderColumn_(sheet, header);
    sheet.getRange(1, col + 1).setNote(
      "空欄は既存の旧料金を維持。新料金・旧料金・要確認から選択。利用開始後の区分変更は過去の再計算に影響するため行わず、管理者へ確認してください。"
    );
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList([LEGACY_PRICING, NEW_PRICING, PENDING_PRICING], true)
      .setAllowInvalid(false).build();
    if (sheet.getMaxRows() > 1) sheet.getRange(2, col + 1, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
  });
  return { success: true, version: PRICING_POLICY_VERSION, message: "料金区分の列を追加しました。既存の金額・区分・残数は変更していません。" };
}

function setupPricingPolicyColumns() {
  const result = setupPricingPolicyColumnsCore_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert(result.message);
}

function adminSetupPricingPolicy_(lineUserId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!isAdminLiffUser_(ss, lineUserId)) return adminDeniedResponse_(lineUserId);
  return setupPricingPolicyColumnsCore_(ss);
}
