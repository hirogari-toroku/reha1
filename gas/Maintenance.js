// 日次バックアップ・LINE送信通数の見張り・月次点検。いずれもGASの時間主導トリガーで
// Google側で動くため、端末やこのリポジトリの操作は不要。
const BACKUP_FOLDER_NAME = "自費リハ管理バックアップ";
const BACKUP_KEEP_GENERATIONS = 7;
const BACKUP_TRIGGER_HOUR = 3;
// 点検結果の通知は、気づいてもらいやすい朝に送る。
const MONTHLY_CHECK_TRIGGER_HOUR = 7;

// 無料のコミュニケーションプランは送信（プッシュ）が月200通まで。返信は無料。
const LINE_FREE_PUSH_LIMIT = 200;
const LINE_PUSH_WARNING_RATIO = 0.75;
const MONTHLY_CHECK_SHEET_NAME = "月次点検";

function backupFolder_() {
  const parents = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  return parents.hasNext() ? parents.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

// 毎日1回、本体スプレッドシートを複製し、古い世代を消す。トリガーから呼ばれる。
function runDailyBackup() {
  const folder = backupFolder_();
  const stamp = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd_HHmm");
  const name = "自費リハ管理_" + stamp;
  let result;

  try {
    const copy = DriveApp.getFileById(MAIN_SPREADSHEET_ID).makeCopy(name, folder);
    const removed = removeOldBackups_(folder);
    result = {
      success: true,
      message: "バックアップを作成しました：" + name + "（古い世代の削除：" + removed + "件）",
      fileId: copy.getId()
    };
  } catch (error) {
    result = { success: false, message: "バックアップに失敗しました：" + (error && error.message ? error.message : String(error)) };
  }

  try {
    saveLiffOperationLog_(SpreadsheetApp.openById(MAIN_SPREADSHEET_ID), "backup",
      "", "", "", "", "", "", result.success ? "成功" : "失敗", result.message, "");
  } catch (logError) {
    // ログに残せなくてもバックアップ自体の結果は返す。
  }
  if (!result.success) notifyAdminMaintenance_("バックアップ失敗", result.message);
  return result;
}

// 新しい順に BACKUP_KEEP_GENERATIONS 件だけ残し、それより古いコピーをゴミ箱へ移す。
function removeOldBackups_(folder) {
  const files = [];
  const iterator = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (iterator.hasNext()) {
    const file = iterator.next();
    if (file.getName().indexOf("自費リハ管理_") === 0) {
      files.push({ file: file, at: file.getDateCreated().getTime() });
    }
  }
  files.sort((a, b) => b.at - a.at);
  let removed = 0;
  files.slice(BACKUP_KEEP_GENERATIONS).forEach(item => {
    item.file.setTrashed(true);
    removed++;
  });
  return removed;
}

function setupDailyBackupTriggerCore_() {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "runDailyBackup" || trigger.getHandlerFunction() === "runMonthlyMaintenanceCheck") {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });

  ScriptApp.newTrigger("runDailyBackup").timeBased().atHour(BACKUP_TRIGGER_HOUR).everyDays(1).create();
  ScriptApp.newTrigger("runMonthlyMaintenanceCheck").timeBased().atHour(MONTHLY_CHECK_TRIGGER_HOUR).everyDays(1).create();

  return {
    success: true,
    message:
      "日次バックアップと点検のトリガーを設定しました。\n" +
      "既存トリガー削除：" + deletedCount + "件\n" +
      "毎日" + BACKUP_TRIGGER_HOUR + "時台にバックアップ（" + BACKUP_KEEP_GENERATIONS + "世代保持）、" +
      MONTHLY_CHECK_TRIGGER_HOUR + "時台にLINE送信通数と月次点検を確認します。\n" +
      "Google側で自動実行されるため、パソコンを開いておく必要はありません。"
  };
}

function setupDailyBackupTrigger() {
  const result = setupDailyBackupTriggerCore_();
  SpreadsheetApp.getUi().alert(result.message);
}

// 送受信ログの「送信」行を月ごとに数える。返信（応答）は無料のため、プッシュ送信だけを
// 通数として扱う。sendPushMessage_ 系は「送信」「送信失敗」で記録される。
function countLinePushMessages_(ss, monthKey) {
  const sheet = ss.getSheetByName(LINE_MESSAGE_LOG_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { month: monthKey, sent: 0, failed: 0, rows: 0 };

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  let sent = 0;
  let failed = 0;
  const reasons = {};
  values.forEach(row => {
    const at = row[0];
    if (!(Object.prototype.toString.call(at) === "[object Date]") || isNaN(at.getTime())) return;
    if (Utilities.formatDate(at, "Asia/Tokyo", "yyyy-MM") !== monthKey) return;
    const direction = String(row[1] || "").trim();
    if (direction === "送信") sent++;
    else if (direction === "送信失敗") {
      failed++;
      const reason = summarizePushFailureReason_(row[4]);
      reasons[reason] = (reasons[reason] || 0) + 1;
    }
  });
  return {
    month: monthKey,
    sent: sent,
    failed: failed,
    rows: values.length,
    failureReasons: Object.keys(reasons).map(reason => reason + "：" + reasons[reason] + "件")
  };
}

// 失敗理由をまとめる。LINEユーザーIDや氏名はまとめ結果に残さない。
function summarizePushFailureReason_(message) {
  const text = String(message || "").replace(/U[0-9a-f]{32}/g, "（ID）").trim();
  if (!text) return "理由不明";
  if (text.indexOf("LINE_CHANNEL_ACCESS_TOKEN") !== -1) return "チャネルアクセストークン未設定";
  if (/LINEユーザーID(が)?未登録/.test(text)) return "送信先のLINEユーザーID未登録";
  const status = text.match(/(\d{3})/);
  if (/送信失敗/.test(text) && status) return "LINE APIエラー " + status[1];
  return text.slice(0, 40);
}

function getLinePushUsage_(ss) {
  const now = new Date();
  const thisMonth = Utilities.formatDate(now, "Asia/Tokyo", "yyyy-MM");
  const previous = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const lastMonth = Utilities.formatDate(previous, "Asia/Tokyo", "yyyy-MM");
  const current = countLinePushMessages_(ss, thisMonth);
  const before = countLinePushMessages_(ss, lastMonth);
  const limit = LINE_FREE_PUSH_LIMIT;
  return {
    month: thisMonth,
    sent: current.sent,
    failed: current.failed,
    lastMonth: lastMonth,
    lastMonthSent: before.sent,
    limit: limit,
    failureReasons: current.failureReasons || [],
    remaining: Math.max(limit - current.sent, 0),
    warning: current.sent >= Math.floor(limit * LINE_PUSH_WARNING_RATIO),
    overLimit: current.sent >= limit
  };
}

function adminLineUsageFromLiff_(lineUserId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!isAdminLiffUser_(ss, lineUserId)) return adminDeniedResponse_(lineUserId);
  const usage = getLinePushUsage_(ss);
  return {
    success: true,
    lineUsage: usage,
    message:
      usage.month + " のLINE送信：" + usage.sent + "通（無料枠" + usage.limit + "通、残り" + usage.remaining + "通）\n" +
      usage.lastMonth + " は " + usage.lastMonthSent + "通。送信失敗：" + usage.failed + "件。\n" +
      (usage.failureReasons.length ? "失敗の内訳：" + usage.failureReasons.join(" / ") + "\n" : "") +
      (usage.overLimit
        ? "無料枠を超えています。超過分は送信されないか、有料プランが必要です。"
        : usage.warning ? "無料枠の残りが少なくなっています。" : "余裕があります。")
  };
}

// 月次点検: 件数・合計と前月差を1行にまとめ、点検シートへ追記する。異常は管理者へLINE。
function buildMonthlyCheck_(ss, monthKey) {
  const visits = countVisitRowsForMonth_(ss, monthKey);
  const schedules = countScheduleRowsForMonth_(ss, monthKey);
  const deposits = sumDepositsForMonth_(ss, monthKey);
  const coupon = summarizeCoupons_(ss);
  const usage = countLinePushMessages_(ss, monthKey);
  const issues = [];

  if (visits.endCount > visits.startCount) issues.push("終了実績が開始実績より多い（" + visits.startCount + "→" + visits.endCount + "）");
  if (coupon.negativeCount > 0) issues.push("回数券残数がマイナスの利用者：" + coupon.negativeCount + "名");
  if (usage.sent >= Math.floor(LINE_FREE_PUSH_LIMIT * LINE_PUSH_WARNING_RATIO)) {
    issues.push("LINE送信が無料枠に近い（" + usage.sent + "/" + LINE_FREE_PUSH_LIMIT + "通）");
  }
  if (usage.failed > 0) issues.push("LINE送信失敗：" + usage.failed + "件");

  return {
    month: monthKey,
    startCount: visits.startCount,
    endCount: visits.endCount,
    scheduleCount: schedules.count,
    completedCount: schedules.completed,
    depositTotal: deposits.total,
    depositCount: deposits.count,
    couponNegative: coupon.negativeCount,
    linePush: usage.sent,
    lineFailed: usage.failed,
    issues: issues
  };
}

function countVisitRowsForMonth_(ss, monthKey) {
  const sheet = ss.getSheetByName(VISIT_RESULT_SHEET_NAME);
  const result = { startCount: 0, endCount: 0 };
  if (!sheet || sheet.getLastRow() < 2) return result;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  values.forEach(row => {
    const date = parseComparisonDate_(row[4], row[0] instanceof Date ? row[0] : undefined);
    if (!date || Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM") !== monthKey) return;
    const type = String(row[1] || "").trim();
    if (type === "開始") result.startCount++;
    if (type === "終了") result.endCount++;
  });
  return result;
}

function countScheduleRowsForMonth_(ss, monthKey) {
  const sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  const result = { count: 0, completed: 0 };
  if (!sheet || sheet.getLastRow() < 2) return result;
  const values = sheet.getDataRange().getValues().slice(1);
  values.forEach(row => {
    const date = parseComparisonDate_(row[3], row[0] instanceof Date ? row[0] : undefined);
    if (!date || Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM") !== monthKey) return;
    const status = String(row[8] || "予定").trim();
    if (isCancelledScheduleStatus_(status)) return;
    result.count++;
    if (status === "完了") result.completed++;
  });
  return result;
}

function sumDepositsForMonth_(ss, monthKey) {
  const sheet = ss.getSheetByName("入出金明細");
  const result = { total: 0, count: 0 };
  if (!sheet || sheet.getLastRow() < 2) return result;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  values.forEach(row => {
    const at = row[0];
    if (!(Object.prototype.toString.call(at) === "[object Date]") || isNaN(at.getTime())) return;
    if (Utilities.formatDate(at, "Asia/Tokyo", "yyyy-MM") !== monthKey) return;
    const amount = Number(row[2]) || 0;
    if (amount <= 0) return;
    result.total += amount;
    result.count++;
  });
  return result;
}

function summarizeCoupons_(ss) {
  const map = getCouponDisplayMap_(ss);
  let negativeCount = 0;
  Object.keys(map).forEach(key => {
    if (Number(map[key].balance) < 0) negativeCount++;
  });
  return { negativeCount: negativeCount };
}

function monthlyCheckSheet_(ss) {
  const headers = ["点検日時", "対象月", "開始実績", "終了実績", "予定(有効)", "完了", "入金件数", "入金合計", "回数券マイナス", "LINE送信", "LINE失敗", "気になる点"];
  let sheet = ss.getSheetByName(MONTHLY_CHECK_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MONTHLY_CHECK_SHEET_NAME);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 毎日動かして構わない（同じ月に何度書いても履歴として残る）。月初は前月分も点検する。
function runMonthlyMaintenanceCheck() {
  const ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  const now = new Date();
  const months = [Utilities.formatDate(now, "Asia/Tokyo", "yyyy-MM")];
  if (now.getDate() <= 3) {
    months.push(Utilities.formatDate(new Date(now.getFullYear(), now.getMonth() - 1, 1), "Asia/Tokyo", "yyyy-MM"));
  }

  const sheet = monthlyCheckSheet_(ss);
  const results = months.map(monthKey => buildMonthlyCheck_(ss, monthKey));
  const rows = results.map(check => [
    now, check.month, check.startCount, check.endCount, check.scheduleCount, check.completedCount,
    check.depositCount, check.depositTotal, check.couponNegative, check.linePush, check.lineFailed,
    check.issues.join(" / ")
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

  const problems = results.filter(check => check.issues.length);
  if (problems.length) {
    notifyAdminMaintenance_("月次点検で気になる点",
      problems.map(check => check.month + "：" + check.issues.join(" / ")).join("\n"));
  }
  return { success: true, checks: results };
}

function adminRunMonthlyCheckFromLiff_(lineUserId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!isAdminLiffUser_(ss, lineUserId)) return adminDeniedResponse_(lineUserId);
  const result = runMonthlyMaintenanceCheck();
  const check = result.checks[0];
  return {
    success: true,
    message:
      check.month + " の点検結果\n" +
      "開始 " + check.startCount + "件 / 終了 " + check.endCount + "件 / 予定 " + check.scheduleCount + "件（完了 " + check.completedCount + "件）\n" +
      "入金 " + check.depositCount + "件 " + check.depositTotal.toLocaleString() + "円\n" +
      "LINE送信 " + check.linePush + "通（失敗 " + check.lineFailed + "件）\n" +
      (check.issues.length ? "気になる点：" + check.issues.join(" / ") : "気になる点はありません。") +
      "\n詳細は「" + MONTHLY_CHECK_SHEET_NAME + "」シートに追記しました。"
  };
}

// 管理者への連絡は1日1回まで（同じ内容で何度も届かないように）。
function notifyAdminMaintenance_(title, body) {
  try {
    if (!ADMIN_LINE_USER_ID) return;
    const cache = CacheService.getScriptCache();
    const key = "maintenanceNotified:" + title + ":" + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd");
    if (cache.get(key)) return;
    cache.put(key, "1", 86400);
    sendPushMessage_(SpreadsheetApp.openById(MAIN_SPREADSHEET_ID), ADMIN_LINE_USER_ID, "システム点検",
      "【" + title + "】\n" + body);
  } catch (error) {
    console.error("notifyAdminMaintenance_ failed: " + error.message);
  }
}

function adminSetupDailyBackupFromLiff_(lineUserId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!isAdminLiffUser_(ss, lineUserId)) return adminDeniedResponse_(lineUserId);
  return setupDailyBackupTriggerCore_();
}
