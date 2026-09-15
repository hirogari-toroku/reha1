const ANNUAL_RAISE_SHEET = "担当継続昇給";
const ANNUAL_RAISE_HEADERS = ["スタッフ名", "利用者名", "担当開始日", "給与区分", "本人紹介", "運用開始月", "有効"];

function annualRaiseMonth_(startDate) {
  const text = Object.prototype.toString.call(startDate) === "[object Date]" && !isNaN(startDate.getTime())
    ? Utilities.formatDate(startDate, "Asia/Tokyo", "yyyy-MM-dd") : String(startDate || "").trim();
  const match = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (!match) throw new Error("担当開始日は年月日で入力してください。");
  const year = Number(match[1]), month = Number(match[2]), day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (year < 1900 || date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    throw new Error("担当開始日が正しくありません。");
  }
  // First day of the month following the one-year anniversary, including leap days.
  const index = year * 12 + month - 1 + 13;
  return Math.floor(index / 12) + "-" + String(index % 12 + 1).padStart(2, "0");
}

function getAnnualRaisePolicies_(ss) {
  const sheet = ss.getSheetByName(ANNUAL_RAISE_SHEET);
  const map = {};
  if (!sheet) return map;
  const rows = sheet.getDataRange().getValues();
  const cols = ANNUAL_RAISE_HEADERS.map(header => rows[0].indexOf(header));
  if (cols.some(col => col < 0)) throw new Error("担当継続昇給の見出しを確認してください。");
  rows.slice(1).forEach(row => {
    const enabled = row[cols[6]];
    if (enabled === false || enabled === "" || enabled == null) return;
    if (enabled !== true) throw new Error("担当継続昇給の有効欄はチェックボックスを使用してください。");
    const staff = normalizeName_(row[cols[0]]), user = normalizeName_(row[cols[1]]);
    const category = String(row[cols[3]] || "").trim();
    const referred = String(row[cols[4]] || "").trim();
    if (!staff || !user || !["運営スタッフ", "他スタッフ"].includes(category) || !["はい", "いいえ"].includes(referred)) {
      throw new Error("担当継続昇給の対象・給与区分・本人紹介を確認してください。");
    }
    const anniversaryMonth = annualRaiseMonth_(row[cols[2]]);
    const activeMonth = payrollHistoryMonth_(row[cols[5]]);
    const key = staff + "_" + user;
    if (map[key]) throw new Error("担当継続昇給に同じ組み合わせの有効行が重複しています。");
    map[key] = {
      effectiveMonth: anniversaryMonth > activeMonth ? anniversaryMonth : activeMonth,
      category, referred: referred === "はい"
    };
  });
  return map;
}

function annualRaiseUnitPay_(policies, key, month, currentPay) {
  const policy = policies[key];
  if (!policy || policy.referred || payrollHistoryMonth_(month) < policy.effectiveMonth) return currentPay;
  const base = policy.category === "運営スタッフ" ? 3500 : 3000;
  const raised = base + 500;
  if (!Number.isFinite(currentPay) || currentPay < base) {
    throw new Error("担当継続昇給の前に新給与体系への手動切替を確認してください。");
  }
  return Math.max(currentPay, raised);
}

function setupAnnualRaiseSheetCore_(ss) {
  let sheet = ss.getSheetByName(ANNUAL_RAISE_SHEET);
  if (sheet) return { success: true, message: "担当継続昇給は作成済みです。既存設定は変更していません。" };
  sheet = ss.insertSheet(ANNUAL_RAISE_SHEET);
  sheet.getRange(1, 1, 1, ANNUAL_RAISE_HEADERS.length).setValues([ANNUAL_RAISE_HEADERS]).setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, 7, 150);
  sheet.getRange("C2:C").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("F2:F").setNumberFormat("@");
  sheet.getRange("G2:G").setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build());
  sheet.getRange("D2:D").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["運営スタッフ", "他スタッフ"], true).setAllowInvalid(false).build());
  sheet.getRange("E2:E").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["はい", "いいえ"], true).setAllowInvalid(false).build());
  sheet.getRange("F1").setNote("YYYY-MM。新給与体系へ手動切替した月以降を指定。1年経過翌月との遅い方から昇給。過去の支払済み月へ遡らせない。");
  sheet.getRange("G1").setNote("担当開始日・本人紹介・給与区分・運用開始月を確認してから有効化。11月分の手動切替を自動実行する設定ではありません。");
  return { success: true, message: "担当継続昇給の設定欄を作成しました。対象は未登録・無効のままです。", sheetId: sheet.getSheetId() };
}

function setupAnnualRaiseSheet() {
  const result = setupAnnualRaiseSheetCore_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert(result.message);
}

function adminSetupAnnualRaise_(lineUserId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!isAdminLiffUser_(ss, lineUserId)) return adminDeniedResponse_(lineUserId);
  return setupAnnualRaiseSheetCore_(ss);
}
