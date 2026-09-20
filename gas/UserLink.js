const USER_LINK_HEADERS = ["利用者ID", "公式LINEユーザーID", "続柄", "本人家族確認", "LIFF用LINEユーザーID", "状態", "招待ハッシュ", "有効期限", "連携日時", "確認者", "送信日時", "利用者名", "閲覧者名", "メモ"];
const USER_LINK_CHANNEL = "2010856600";

function findConfirmedUserDirectoryMatch_(ss, id) {
  if (!/^U[0-9a-f]{32}$/.test(String(id || ""))) return null;
  try {
    const rows = userLinkSheet_(ss).getDataRange().getValues().slice(1)
      .filter(row => row[4] === id);
    if (rows.length !== 1 || rows[0][3] !== "確認済み" || rows[0][5] !== "連携済み") return null;
    const row = rows[0], user = userLinkMaster_(ss, String(row[0]));
    return { name: user.name, userId: user.id, lineDisplayName: row[12] || "",
      liffLineUserId: id, messagingLineUserId: "" };
  } catch (error) {
    if (error.userLinkSafe) return null;
    throw error;
  }
}

function userLinkFail_(message) {
  const error = new Error(message);
  error.userLinkSafe = true;
  throw error;
}

function userLinkSheet_(ss) {
  const sheet = ss.getSheetByName("利用者LINE連携");
  if (!sheet || JSON.stringify(sheet.getRange(1, 1, 1, 14).getValues()[0]) !== JSON.stringify(USER_LINK_HEADERS)) {
    userLinkFail_("連携表の列構成を管理者へご確認ください。");
  }
  return sheet;
}

function userLinkMaster_(ss, id) {
  const sheet = ss.getSheetByName(USER_MASTER_SHEET_NAME);
  if (!sheet) userLinkFail_("利用者マスタを確認できません。");
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const idCol = headers.indexOf(USER_ID_HEADER), nameCol = headers.indexOf("利用者名");
  const rows = values.filter(row => id && idCol >= 0 && String(row[idCol]) === id);
  if (rows.length !== 1 || nameCol < 0 || !rows[0][nameCol]) userLinkFail_("利用者IDを一意に確認できません。");
  const name = String(rows[0][nameCol]);
  // Existing schedules use names, so duplicate names must not share authorization.
  if (values.filter(row => normalizeName_(row[nameCol]) === normalizeName_(name)).length !== 1) userLinkFail_("同名の利用者がいます。管理者へご確認ください。");
  const statusCol = headers.indexOf("状態");
  if (statusCol >= 0 && !isActiveUserStatus_(rows[0][statusCol])) userLinkFail_("現在この利用者の予定は表示できません。");
  return { id: id, name: name };
}

// This dialog is available to spreadsheet editors, not through the public web API.
function openUserLinkDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  userLinkSheet_(ss);
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile("UserLinkDialog").setWidth(560).setHeight(650), "利用者・家族のLINE連携");
}

function userLinkDialogOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(USER_MASTER_SHEET_NAME).getDataRange().getValues();
  const h = master.shift();
  const idCol = h.indexOf(USER_ID_HEADER), nameCol = h.indexOf("利用者名");
  const directory = ss.getSheetByName(LINE_USER_DIRECTORY_SHEET_NAME);
  const contacts = directory && directory.getLastRow() > 1 ? directory.getRange(2, 1, directory.getLastRow() - 1, 13).getValues() : [];
  return {
    users: master.filter(row => row[idCol] && row[nameCol]).map(row => ({ id: String(row[idCol]), name: String(row[nameCol]) })),
    contacts: contacts.filter(row => /^U[0-9a-f]{32}$/.test(String(row[2]))).map(row => ({ id: String(row[2]), name: String(row[4] || "表示名なし"), message: String(row[11] || "").slice(0, 100) })),
    logins: userLinkSheet_(ss).getDataRange().getValues().slice(1)
      .filter(row => row[4] && row[5] === "未連携" && row[3] === "未確認")
      .map(row => ({ id: String(row[4]), name: String(row[12] || "表示名なし") }))
  };
}

function userLinkConfirmAccount(input) {
  if (!input || input.confirmed !== true) userLinkFail_("本人・家族と対象アカウントの確認が必要です。");
  const relation = String(input.relation || "").trim();
  if (!relation || relation.length > 30) userLinkFail_("続柄を入力してください。");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const options = userLinkDialogOptions();
  const matches = options.contacts.filter(c => c.id === input.messagingId);
  if (matches.length !== 1) userLinkFail_("公式LINEの受信記録を一意に確認できません。");
  const user = userLinkMaster_(ss, String(input.userId || ""));
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = userLinkSheet_(ss);
    const rows = sheet.getDataRange().getValues().slice(1);
    const targets = rows.map((r, i) => ({ r, i })).filter(x => input.liffId && x.r[4] === input.liffId);
    if (targets.length !== 1) userLinkFail_("予約確認のログイン記録を一意に確認できません。");
    const target = targets[0], row = target.r;
    if (row[5] !== "未連携" || row[3] !== "未確認") userLinkFail_("このログインは確認済み、または無効です。自動で上書きしません。");
    if ((row[0] && String(row[0]) !== user.id) || (row[1] && row[1] !== input.messagingId)) userLinkFail_("既存の利用者・LINE情報と一致しません。");
    if (rows.some((r, i) => i !== target.i && r[1] === input.messagingId)) userLinkFail_("公式LINEは別の連携記録に登録されています。管理者へ確認してください。");
    const rowNumber = target.i + 2;
    // Preserve captured LIFF ID, legacy invitation columns, notes and sent timestamps.
    sheet.getRange(rowNumber, 1, 1, 4).setValues([[user.id, input.messagingId, relation, "確認済み"]]);
    sheet.getRange(rowNumber, 9, 1, 2).setValues([[new Date(), Session.getActiveUser().getEmail() || "シート編集者"]]);
    sheet.getRange(rowNumber, 12).setValue(user.name);
    sheet.getRange(rowNumber, 6).setValue("連携済み");
    SpreadsheetApp.flush();
    return { message: user.name + "様の「" + relation + "」として登録しました。予約確認を開き直すか、予約を更新すると表示されます。メッセージは送信していません。" };
  } finally { lock.releaseLock(); }
}

function userLinkVerify_(token) {
  if (typeof token !== "string" || !token || token.length > 4096) userLinkFail_("LINEでログインし直してください。");
  // Send both LINE API calls together instead of waiting for verify before starting
  // profile: they are independent reads, and the profile is only trusted below once
  // verify has also passed.
  const [verified, response] = UrlFetchApp.fetchAll([
    { url: "https://api.line.me/oauth2/v2.1/verify?access_token=" + encodeURIComponent(token), muteHttpExceptions: true },
    { url: "https://api.line.me/v2/profile", headers: { Authorization: "Bearer " + token }, muteHttpExceptions: true }
  ]);
  if (verified.getResponseCode() !== 200) userLinkFail_("LINEログインの有効期限が切れています。開き直してください。");
  const result = JSON.parse(verified.getContentText());
  if (String(result.client_id) !== USER_LINK_CHANNEL || !(Number(result.expires_in) > 0)) userLinkFail_("LINE認証を確認できません。");
  if (response.getResponseCode() !== 200) userLinkFail_("LINEプロフィールを確認できません。");
  const profile = JSON.parse(response.getContentText());
  if (!/^U[0-9a-f]{32}$/.test(profile.userId || "")) userLinkFail_("LINE認証を確認できません。");
  return profile;
}

function userLinkCapturePending_(ss, profile) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = userLinkSheet_(ss), rows = sheet.getDataRange().getValues().slice(1);
    if (rows.some(row => row[4] === profile.userId)) return;
    const row = ["", "", "", "未確認", profile.userId, "未連携", "", "", "", "", "", "", profile.displayName || "", "LINE認証済み・予約確認の初回ログイン"];
    const rowNumber = sheet.getLastRow() + 1;
    if (rowNumber > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 100);
    sheet.getRange(rowNumber, 1, 1, 14).setValues([row]);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
}

function userLinkResolve_(ss, id) {
  const rows = userLinkSheet_(ss).getDataRange().getValues().slice(1).filter(row => row[4] === id);
  if (!rows.length) return null;
  if (rows.length === 1 && rows[0][3] === "未確認" && rows[0][5] === "未連携") return null;
  if (rows.length !== 1 || rows[0][3] !== "確認済み" || rows[0][5] !== "連携済み") userLinkFail_("連携状態を管理者へご確認ください。");
  return userLinkMaster_(ss, String(rows[0][0]));
}

function userLinkRequest_(data) {
  let diagnosticProfile = null;
  try {
    const profile = userLinkVerify_(data.accessToken);
    diagnosticProfile = profile;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (getStaffNameCached_(ss, profile.userId) !== "未登録") userLinkFail_("スタッフの方は通常の予約確認ページから開き直してください。");
    const user = userLinkResolve_(ss, profile.userId);
    if (!user) {
      userLinkCapturePending_(ss, profile);
      logUserLinkStage_(data, "pending_saved", profile);
      return { success: false, pending: true, schedules: [], message: "LINE情報を受け付けました。管理者が確認・登録後に予定を表示します。公式LINEへ利用者様のお名前と続柄をお知らせください。" };
    }
    const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
    if (!scheduleSheet) userLinkFail_("訪問予定を確認できません。管理者へご連絡ください。");
    const visits = buildVisitStatusIndex_(ss.getSheetByName(VISIT_RESULT_SHEET_NAME), "", user.name);
    const items = collectActiveSchedulesForUser_(scheduleSheet, user.name, visits);
    // Return only the fields rendered to users; never expose chart or staff-folder URLs.
    const schedules = items.map(item => ({ visitDate: item.visitDate, status: item.status, staffName: item.staffName, updatedAt: item.updatedAt, kind: item.kind, lastVisitText: item.lastVisitText }));
    const coupon = getCouponDisplayMap_(ss)[normalizeName_(user.name)] || null;
    // Normal viewing is read-only; pending registrations and failures retain logs.
    return { success: true, linked: true, userName: user.name, schedules: schedules, coupon: coupon, message: schedules.length ? "" : "対象期間内の予約はありません。" };
  } catch (error) {
    logUserLinkStage_(data, "request_failed", diagnosticProfile);
    return { success: false, schedules: [], message: error.userLinkSafe ? error.message : "通信または連携処理に失敗しました。時間をおいて再度お試しください。" };
  }
}
