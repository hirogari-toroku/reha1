// 毎月1日の給与自動処理。前月分について、予定・実績照合 → 給与集計 → 給与チェックを行い、
// 結果を「給与チェック」シートに書いて止まる。管理者が確認したあと「確認済み」の操作
// （メニューまたは管理画面）で、賃金台帳・給与明細PDF・GMO振込CSVを作り、PDFをスタッフフォルダへ移す。
// LINEでの通知はしない（管理者・スタッフとも）。銀行への振込登録は管理者が行う。
const PAYROLL_CHECK_SHEET_NAME = "給与チェック";
const PAYROLL_AUTO_HANDLER = "runMonthlyPayrollAuto";
const PAYROLL_AUTO_DAY = 1;
const PAYROLL_AUTO_HOUR = 8;
// トリガーは前後15分ずれる（7:45〜8:15ごろ）。Google側で動くのでPCは不要。
// 8:30にPCを開いた時点でチェック結果がそろっているよう、8:00ごろに動かす。
const PAYROLL_AUTO_MINUTE = 0;
// true にすると、要確認がない月は確認を待たずに明細まで作る。運用に慣れるまでは false。
const PAYROLL_AUTO_CREATE_WITHOUT_REVIEW = false;
const PAYROLL_AUTO_STATE_PROPERTY = "PAYROLL_AUTO_STATE";
const PAYROLL_REVIEW = "要確認";
const PAYROLL_INFO = "参考";
const PAYROLL_WAITING = "確認待ち";
// 1回の訪問として不自然な長さ（分）。
const PAYROLL_SHORT_VISIT_MINUTES = 20;
const PAYROLL_LONG_VISIT_MINUTES = 180;
// 同じ日の開始・終了がこの分数以内に複数あるだけなら、二重登録として扱う（給与は1回）。
const PAYROLL_DUPLICATE_WINDOW_MINUTES = 15;
// 訪問回数が前月からこの割合以上、かつこの回数以上変わったら知らせる。
const PAYROLL_COUNT_CHANGE_RATIO = 0.5;
const PAYROLL_COUNT_CHANGE_MIN = 4;

function payrollMonthLabel_(ym) {
  const parts = String(ym).split("-");
  return parts[0] + "年" + Number(parts[1]) + "月分";
}

// トリガーから呼ばれる。対象は前月分。
function runMonthlyPayrollAuto() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  return withPayrollLock_(() => {
    const ym = previousMonthKey_(new Date());
    const prepared = preparePayrollForMonth_(ss, ym);
    if (!prepared.success) return finishPayrollWithError_(ym, prepared.message);
    if (prepared.staffCount === 0) {
      savePayrollAutoState_(ym, "対象なし");
      return { success: true, ym: ym, status: "対象なし" };
    }

    const reviews = prepared.issues.filter(issue => issue.level === PAYROLL_REVIEW);
    if (reviews.length > 0 || !PAYROLL_AUTO_CREATE_WITHOUT_REVIEW) {
      const status = reviews.length > 0 ? PAYROLL_REVIEW : PAYROLL_WAITING;
      savePayrollAutoState_(ym, status);
      return { success: true, ym: ym, status: status, issues: prepared.issues };
    }

    return createPayrollFilesForMonth_(ss, ym, prepared);
  });
}

// 管理者が確認・修正したあとに実行する。修正を反映するため集計からやり直す。
function approvePayrollAndCreateFilesCore_(ss) {
  return withPayrollLock_(() => {
    const ym = payrollTargetMonth_();
    const prepared = preparePayrollForMonth_(ss, ym);
    if (!prepared.success) return finishPayrollWithError_(ym, prepared.message);
    if (prepared.staffCount === 0) {
      return { success: false, ym: ym, message: payrollMonthLabel_(ym) + "は支給対象の実績がありません。明細は作っていません。" };
    }
    return createPayrollFilesForMonth_(ss, ym, prepared);
  });
}

// 修正のあと、明細を作る前にもう一度チェックだけ行う。
function recheckPayrollCore_(ss) {
  return withPayrollLock_(() => {
    const ym = payrollTargetMonth_();
    const prepared = preparePayrollForMonth_(ss, ym);
    if (!prepared.success) return { success: false, ym: ym, message: prepared.message };
    const reviews = prepared.issues.filter(issue => issue.level === PAYROLL_REVIEW);
    return {
      success: true,
      ym: ym,
      issues: prepared.issues,
      message: payrollMonthLabel_(ym) + "の給与チェック：要確認 " + reviews.length + "件" +
        (reviews.length ? "（" + summarizePayrollIssues_(reviews) + "）" : "") +
        "\n詳細は「" + PAYROLL_CHECK_SHEET_NAME + "」シートを見てください。明細はまだ作っていません。"
    };
  });
}

// 確認待ちで止まっている月があればその月、なければ前月。
function payrollTargetMonth_() {
  const state = loadPayrollAutoState_();
  if (state && state.ym && (state.status === PAYROLL_REVIEW || state.status === PAYROLL_WAITING)) return state.ym;
  return previousMonthKey_(new Date());
}

function preparePayrollForMonth_(ss, ym) {
  try {
    const comparison = updateScheduleVisitComparisonCore_(ss);
    if (!comparison.success) return { success: false, message: comparison.message };
    createPayrollSummary(ss);
  } catch (error) {
    return { success: false, message: "給与集計でエラー：" + (error && error.message ? error.message : String(error)) };
  }

  const checked = buildPayrollChecks_(ss, ym);
  writePayrollCheckSheet_(ss, ym, checked.issues);
  return { success: true, ym: ym, issues: checked.issues, staffCount: checked.staffCount, totalPayment: checked.totalPayment };
}

function createPayrollFilesForMonth_(ss, ym, prepared) {
  createWageLedger(ss, ym);
  const pdf = createPayrollPdfFromTemplate(ss, ym) || {};
  const csv = createGmoTransferCsv(ss, ym) || {};
  const moved = movePayrollPdfsToStaffFoldersCore_(ss);

  const problems = [];
  (pdf.failedFiles || []).forEach(text => problems.push("明細PDF作成失敗：" + text));
  (pdf.noStaffFolderFiles || []).forEach(text => problems.push("スタッフフォルダ未登録：" + text));
  (pdf.skippedFiles || []).forEach(text => problems.push("同じ月の明細がすでにあるため作り直していません：" + text));
  (moved.skippedFiles || []).forEach(text => problems.push("フォルダへ移せなかった明細：" + text));
  (csv.unregisteredStaffs || []).forEach(name => problems.push("口座未登録のため振込CSVに入っていない：" + name));
  (csv.skippedCsvFiles || []).forEach(text => problems.push("振込CSVは作り直していません（同名あり）：" + text));
  if (csv.success === false && csv.message) problems.push("振込CSV：" + csv.message);
  if (csv.warningCount) problems.push("受取人名とスタッフ名が一致しない可能性：" + csv.warningCount + "件（GMO振込CSVシートの確認メモ列）");

  const status = problems.length ? "作成済（一部要確認）" : "作成済";
  savePayrollAutoState_(ym, status);

  const message =
    payrollMonthLabel_(ym) + "の給与明細を作成しました。\n" +
    "明細PDF：作成 " + (pdf.createdCount || 0) + "件、スタッフフォルダへ移動 " + (moved.movedCount || 0) + "件\n" +
    "振込CSV：作成 " + (csv.createdCsvCount || 0) + "件（確定フォルダへ移動 " + (moved.csvMovedCount || 0) + "件）\n" +
    "支給日 " + getPayrollTransferDate_(ym) + "。振込CSVの銀行への登録をお願いします。" +
    (problems.length ? "\n\n確認してほしいこと：\n" + problems.join("\n") : "");

  return { success: problems.length === 0, ym: ym, status: status, message: message, issues: prepared.issues };
}

function finishPayrollWithError_(ym, message) {
  savePayrollAutoState_(ym, "エラー");
  return { success: false, ym: ym, status: "エラー", message: message + "\n明細・振込CSVは作っていません。" };
}

// 「未実施2・口座1」のように項目ごとの件数にまとめる。
function summarizePayrollIssues_(issues) {
  const counts = {};
  const order = [];
  issues.forEach(issue => {
    if (!counts[issue.item]) order.push(issue.item);
    counts[issue.item] = (counts[issue.item] || 0) + 1;
  });
  return order.map(item => item + counts[item]).join("・");
}

// ---- チェック ----

function buildPayrollChecks_(ss, ym) {
  const issues = [];
  checkScheduleVisitsForPayroll_(ss, ym, issues);
  const payroll = checkPayrollSummaryForPayroll_(ss, ym, issues);
  checkPaidMonthsForPayroll_(ss, ym, issues);
  checkStaffBankForPayroll_(ss, payroll.staffNames, issues);
  return { issues: issues, staffCount: payroll.staffNames.length, totalPayment: payroll.totalPayment };
}

function addPayrollIssue_(issues, level, item, date, staffName, userName, detail) {
  issues.push({ level: level, item: item, date: date || "", staffName: staffName || "", userName: userName || "", detail: detail || "" });
}

// 予定と実績の突き合わせ（対象月の訪問日のもの）。実績の利用者名は「白井」のような
// 苗字だけのこともあるため、担当の正式名へそろえてから予定と比べる。
function checkScheduleVisitsForPayroll_(ss, ym, issues) {
  const masterMap = getStaffUserMasterMap_(ss);
  const resolvedCache = {};
  const resolve = (staffName, userName) => {
    const key = normalizeName_(staffName) + "|" + normalizeName_(userName);
    if (!(key in resolvedCache)) resolvedCache[key] = resolveUserName_(ss, staffName, userName);
    return resolvedCache[key];
  };

  const merged = {};
  buildScheduleVisitComparisonItems_(
    ss.getSheetByName(SCHEDULE_SHEET_NAME),
    ss.getSheetByName(VISIT_RESULT_SHEET_NAME)
  ).forEach(item => {
    if (Utilities.formatDate(item.date, "Asia/Tokyo", "yyyy-MM") !== ym) return;
    const userName = resolve(item.staffName, item.userName);
    const key = Utilities.formatDate(item.date, "Asia/Tokyo", "yyyy-MM-dd") + "|" + normalizeName_(item.staffName) + "|" + normalizeName_(userName);
    if (!merged[key]) {
      merged[key] = {
        date: item.date, staffName: item.staffName, userName: userName,
        known: !!masterMap[normalizeName_(item.staffName) + "_" + normalizeName_(userName)],
        planned: false, startTimes: [], endTimes: [], otherVisitTypes: []
      };
    }
    const target = merged[key];
    target.planned = target.planned || item.planned;
    item.startTimes.forEach(t => addUniqueText_(target.startTimes, t));
    item.endTimes.forEach(t => addUniqueText_(target.endTimes, t));
    item.otherVisitTypes.forEach(t => addUniqueText_(target.otherVisitTypes, t));
  });

  Object.keys(merged).map(key => merged[key])
    .sort((a, b) => a.date.getTime() - b.date.getTime())
    .forEach(item => checkOneVisitForPayroll_(item, issues));
}

function checkOneVisitForPayroll_(item, issues) {
  const date = Utilities.formatDate(item.date, "Asia/Tokyo", "yyyy/MM/dd");
  const hasStart = item.startTimes.length > 0;
  const hasEnd = item.endTimes.length > 0;
  const hasVisit = hasStart || hasEnd || item.otherVisitTypes.length > 0;
  const add = (level, name, detail) => addPayrollIssue_(issues, level, name, date, item.staffName, item.userName, detail);

  // 担当として登録のない名前（「お疲れ様です」などの連絡文が実績として入ったもの）は給与に入らない。
  if (!item.planned && !item.known) {
    add(PAYROLL_INFO, "利用者以外", "担当の利用者にない名前の実績（連絡文などの誤登録の可能性。給与には入らない）");
    return;
  }

  if (item.planned && !hasVisit) add(PAYROLL_REVIEW, "未実施", "予定はあるが実績がない（実施していれば報告漏れ、中止ならキャンセル登録漏れ）");
  if (!item.planned && hasVisit) add(PAYROLL_REVIEW, "予定外", "予定にない実績（日付や利用者の登録違いがないか）");
  if (hasStart && !hasEnd) add(PAYROLL_REVIEW, "開始のみ", "開始だけで終了の登録がない（給与に入らない）");
  if (!hasStart && hasEnd) add(PAYROLL_REVIEW, "終了のみ", "終了だけで開始の登録がない");

  if (item.startTimes.length > 1 || item.endTimes.length > 1) {
    const detail = "同じ日に開始" + item.startTimes.length + "件（" + item.startTimes.join("、") + "）・終了" +
      item.endTimes.length + "件（" + item.endTimes.join("、") + "）";
    if (payrollTimesWithin_(item.startTimes) && payrollTimesWithin_(item.endTimes)) {
      add(PAYROLL_INFO, "二重登録", detail + "。同じ訪問の二重登録とみられ、給与は1回で計算");
    } else {
      add(PAYROLL_REVIEW, "重複", detail + "。別の訪問の可能性があるが、給与は1回で計算");
    }
  }

  if (hasStart && hasEnd) {
    const minutes = payrollMinutesBetween_(item.startTimes[0], item.endTimes[0]);
    if (minutes !== null && (minutes < PAYROLL_SHORT_VISIT_MINUTES || minutes > PAYROLL_LONG_VISIT_MINUTES)) {
      add(PAYROLL_REVIEW, "訪問時間", item.startTimes[0] + "〜" + item.endTimes[0] + "（" + minutes + "分）");
    }
  }
}

function payrollTextToMinutes_(text) {
  const match = String(text || "").match(/^(\d{1,2}):(\d{2})$/);
  return match ? Number(match[1]) * 60 + Number(match[2]) : null;
}

function payrollMinutesBetween_(startText, endText) {
  const start = payrollTextToMinutes_(startText);
  const end = payrollTextToMinutes_(endText);
  if (start === null || end === null) return null;
  return end - start;
}

function payrollTimesWithin_(times) {
  if (times.length <= 1) return true;
  const minutes = times.map(payrollTextToMinutes_);
  if (minutes.some(m => m === null)) return false;
  return Math.max.apply(null, minutes) - Math.min.apply(null, minutes) <= PAYROLL_DUPLICATE_WINDOW_MINUTES;
}

// 給与集計の結果を前月と比べる。
function checkPayrollSummaryForPayroll_(ss, ym, issues) {
  const sheet = ss.getSheetByName("給与集計");
  const result = { staffNames: [], totalPayment: 0 };
  if (!sheet || sheet.getLastRow() < 2) return result;

  const values = sheet.getDataRange().getValues();
  const prevYm = previousMonthKey_(new Date(Number(ym.slice(0, 4)), Number(ym.slice(5, 7)) - 1, 1));
  const byMonth = { [ym]: {}, [prevYm]: {} };

  values.slice(1).forEach(row => {
    const rowYm = formatPayrollYmKey_(row[0]);
    if (!byMonth[rowYm] || !row[2]) return;
    byMonth[rowYm][normalizeName_(row[2])] = {
      staffName: String(row[2]),
      count: Number(row[3]) || 0,
      basePay: Number(row[4]) || 0,
      payment: Number(row[11]) || 0,
      memo: String(row[12] || "").trim()
    };
  });

  const current = byMonth[ym];
  const previous = byMonth[prevYm];

  Object.keys(current).forEach(key => {
    const now = current[key];
    const before = previous[key];
    result.staffNames.push(now.staffName);
    result.totalPayment += now.payment;

    if (now.memo) {
      // 同じ訪問の終了二重登録（重複除外）だけなら、給与は正しく1回で数えている。
      const onlyDuplicates = now.memo.split("、").every(part => /重複除外）$/.test(part));
      addPayrollIssue_(issues, onlyDuplicates ? PAYROLL_INFO : PAYROLL_REVIEW, "集計から除外", "", now.staffName, "",
        "給与に入らなかった実績：" + now.memo + (onlyDuplicates ? "（同じ日の二重登録を1回として数えたもの）" : "（担当の組み合わせ未登録・同姓など）"));
    }
    if (!before) {
      addPayrollIssue_(issues, PAYROLL_INFO, "前月なし", "", now.staffName, "", "前月は支給なし（新しいスタッフなら問題なし）");
      return;
    }
    const diff = now.count - before.count;
    if (Math.abs(diff) >= PAYROLL_COUNT_CHANGE_MIN && Math.abs(diff) >= before.count * PAYROLL_COUNT_CHANGE_RATIO) {
      addPayrollIssue_(issues, PAYROLL_REVIEW, "回数の変化", "", now.staffName, "",
        "訪問回数 " + before.count + "回 → " + now.count + "回");
    }
    const nowUnit = now.count ? Math.round(now.basePay / now.count) : 0;
    const beforeUnit = before.count ? Math.round(before.basePay / before.count) : 0;
    if (nowUnit && beforeUnit && nowUnit !== beforeUnit) {
      addPayrollIssue_(issues, PAYROLL_REVIEW, "単価の変化", "", now.staffName, "",
        "1回あたりの基本給 " + beforeUnit + "円 → " + nowUnit + "円（単価変更の月か、担当利用者の構成が変わったか確認）");
    }
  });

  Object.keys(previous).forEach(key => {
    if (current[key]) return;
    addPayrollIssue_(issues, PAYROLL_INFO, "今月なし", "", previous[key].staffName, "", "前月は支給あり、この月は実績なし");
  });

  return result;
}

// 支払済みの月（賃金台帳に記録のある月）を計算し直した結果が、台帳の支給額と違えば知らせる。
// 単価をマスタだけで書き換えた、支払後に実績が追加・削除された、などで起きる。
function checkPaidMonthsForPayroll_(ss, ym, issues) {
  const ledger = ss.getSheetByName("賃金台帳");
  const payroll = ss.getSheetByName("給与集計");
  if (!ledger || ledger.getLastRow() < 3 || !payroll || payroll.getLastRow() < 2) return;

  const paid = {};
  const paidDates = {};
  ledger.getDataRange().getValues().slice(2).forEach(row => {
    const payDate = formatDateForKey_(row[0]);
    const staffKey = normalizeName_(row[1]);
    if (!payDate || !staffKey) return;
    paid[payDate + "_" + staffKey] = { staffName: String(row[1]), payment: Number(row[15]) || 0 };
    paidDates[payDate] = true;
  });

  const recalculated = {};
  payroll.getDataRange().getValues().slice(1).forEach(row => {
    const rowYm = formatPayrollYmKey_(row[0]);
    const payDate = formatDateForKey_(row[1]);
    if (!rowYm || rowYm >= ym || !paidDates[payDate] || !row[2]) return;
    recalculated[payDate + "_" + normalizeName_(row[2])] = {
      ym: rowYm, staffName: String(row[2]), payment: Number(row[11]) || 0
    };
  });

  const keys = {};
  Object.keys(paid).concat(Object.keys(recalculated)).forEach(key => { keys[key] = true; });
  Object.keys(keys).sort().forEach(key => {
    const before = paid[key];
    const now = recalculated[key];
    const beforePayment = before ? before.payment : 0;
    const nowPayment = now ? now.payment : 0;
    if (beforePayment === nowPayment) return;
    const payDate = key.split("_")[0];
    addPayrollIssue_(issues, PAYROLL_REVIEW, "支払済み月の変化", "", (now || before).staffName, "",
      "支給日 " + payDate + " の分を計算し直すと " + beforePayment + "円 → " + nowPayment +
      "円（単価をマスタだけで書き換えた、または支払後に実績が追加・削除された可能性。支払済みの明細は変わらない）");
  });
}

// 振込先。登録状況・口座の有無・受取人名の簡易チェック。
function checkStaffBankForPayroll_(ss, staffNames, issues) {
  if (!staffNames.length) return;
  const sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  if (!sheet) return;
  const values = sheet.getDataRange().getValues();
  const cols = getStaffMasterColumnMap_(sheet);
  const statusByStaff = {};
  const folderByStaff = {};
  values.slice(1).forEach(row => {
    const name = normalizeName_(row[cols.name]);
    if (!name) return;
    if (cols.bankRegistrationStatus >= 0) statusByStaff[name] = String(row[cols.bankRegistrationStatus] || "").trim();
    folderByStaff[name] = String(row[cols.payrollFolderId] || "").trim();
  });
  const bankMap = getStaffBankMap_(ss);

  staffNames.forEach(staffName => {
    const key = normalizeName_(staffName);
    if (!folderByStaff[key]) {
      addPayrollIssue_(issues, PAYROLL_REVIEW, "スタッフフォルダ", "", staffName, "", "スタッフフォルダIDが未登録（明細をフォルダへ移せない）");
    }
    const bank = bankMap[key];
    const status = statusByStaff[key];
    if (!bank) {
      addPayrollIssue_(issues, PAYROLL_REVIEW, "口座", "", staffName, "", "口座情報が未登録（振込CSVに入らない）");
      return;
    }
    if (status === "未登録" || status === "要確認") {
      addPayrollIssue_(issues, PAYROLL_REVIEW, "口座", "", staffName, "", "振込先登録状況が「" + status + "」");
    }
    const warning = makeGmoReceiverWarning_(staffName, normalizeGmoText_(bank.receiverName));
    if (warning) addPayrollIssue_(issues, PAYROLL_REVIEW, "口座", "", staffName, "", warning);
  });
}

function writePayrollCheckSheet_(ss, ym, issues) {
  const headers = ["対象月", "判定", "項目", "訪問日", "スタッフ", "利用者", "内容", "チェック日時"];
  const sheet = ss.getSheetByName(PAYROLL_CHECK_SHEET_NAME) || ss.insertSheet(PAYROLL_CHECK_SHEET_NAME);
  const checkedAt = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
  const sorted = issues.slice().sort((a, b) => (a.level === b.level ? 0 : a.level === PAYROLL_REVIEW ? -1 : 1));
  const rows = sorted.length
    ? sorted.map(issue => [payrollMonthLabel_(ym), issue.level, issue.item, issue.date, issue.staffName, issue.userName, issue.detail, checkedAt])
    : [[payrollMonthLabel_(ym), "問題なし", "", "", "", "", "気になる点はありませんでした", checkedAt]];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
}

// ---- 状態・トリガー ----

function savePayrollAutoState_(ym, status) {
  PropertiesService.getScriptProperties().setProperty(PAYROLL_AUTO_STATE_PROPERTY, JSON.stringify({
    ym: ym,
    status: status,
    at: Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm")
  }));
}

function loadPayrollAutoState_() {
  try {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty(PAYROLL_AUTO_STATE_PROPERTY) || "null");
  } catch (error) {
    return null;
  }
}

// 同時に2つ動くと明細や振込CSVが重複するため、1つずつ実行する。
function withPayrollLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: "給与処理が実行中です。少し待ってからもう一度お試しください。" };
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function payrollAutoTimeText_() {
  return PAYROLL_AUTO_HOUR + ":" + String(PAYROLL_AUTO_MINUTE).padStart(2, "0");
}

function setupPayrollAutoTriggerCore_() {
  let deletedCount = 0;
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === PAYROLL_AUTO_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });
  // nearMinute は前後15分の幅で実行される。
  ScriptApp.newTrigger(PAYROLL_AUTO_HANDLER).timeBased()
    .onMonthDay(PAYROLL_AUTO_DAY).atHour(PAYROLL_AUTO_HOUR).nearMinute(PAYROLL_AUTO_MINUTE).create();
  return {
    success: true,
    message: "毎月の給与自動処理を設定しました（既存の設定の置き換え：" + deletedCount + "件）。\n" +
      "毎月" + PAYROLL_AUTO_DAY + "日 " + payrollAutoTimeText_() + "ごろ（前後15分）に、前月分の照合・集計・チェックを行い、" +
      "結果を「" + PAYROLL_CHECK_SHEET_NAME + "」シートに書いて止まります。確認後に「確認済み→明細作成」を実行してください。"
  };
}

function payrollAutoStatusMessage_() {
  const registered = ScriptApp.getProjectTriggers().some(trigger => trigger.getHandlerFunction() === PAYROLL_AUTO_HANDLER);
  const state = loadPayrollAutoState_();
  return (registered
    ? "給与の自動処理：設定済み（毎月" + PAYROLL_AUTO_DAY + "日 " + payrollAutoTimeText_() + "ごろ）"
    : "給与の自動処理：未設定") +
    (state ? "\n直近：" + payrollMonthLabel_(state.ym) + " " + state.status + "（" + state.at + "）" : "");
}

// ---- メニュー ----

function setupPayrollAutoTrigger() {
  payrollAlert_(setupPayrollAutoTriggerCore_().message);
}

function recheckPayroll() {
  payrollAlert_(recheckPayrollCore_(SpreadsheetApp.getActiveSpreadsheet()).message);
}

function approvePayrollAndCreateFiles() {
  const ym = payrollTargetMonth_();
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert("給与明細の作成",
    payrollMonthLabel_(ym) + "を集計し直して、給与明細PDFと振込CSVを作り、PDFをスタッフフォルダへ入れます。\n" +
    "スタッフから明細が見えるようになります。よろしいですか。", ui.ButtonSet.OK_CANCEL);
  if (answer !== ui.Button.OK) return;
  ui.alert(approvePayrollAndCreateFilesCore_(SpreadsheetApp.getActiveSpreadsheet()).message);
}

// ---- 管理画面（LIFF） ----

function adminPayrollRecheckFromLiff_(lineUserId) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  if (!isAdminLiffUser_(ss, lineUserId)) return adminDeniedResponse_(lineUserId);
  return recheckPayrollCore_(ss);
}

function adminPayrollApproveFromLiff_(lineUserId) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  if (!isAdminLiffUser_(ss, lineUserId)) return adminDeniedResponse_(lineUserId);
  return approvePayrollAndCreateFilesCore_(ss);
}

function adminSetupPayrollAutoFromLiff_(lineUserId) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  if (!isAdminLiffUser_(ss, lineUserId)) return adminDeniedResponse_(lineUserId);
  const result = setupPayrollAutoTriggerCore_();
  return { success: true, message: result.message + "\n\n" + payrollAutoStatusMessage_() };
}

function adminPayrollStatusFromLiff_(lineUserId) {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  if (!isAdminLiffUser_(ss, lineUserId)) return adminDeniedResponse_(lineUserId);
  return { success: true, message: payrollAutoStatusMessage_() };
}
