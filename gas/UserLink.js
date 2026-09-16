const USER_LINK_HEADERS = ["利用者ID", "公式LINEユーザーID", "続柄", "本人家族確認", "LIFF用LINEユーザーID", "状態", "招待ハッシュ", "有効期限", "連携日時", "確認者", "送信日時", "利用者名", "閲覧者名", "メモ"];
const USER_LINK_CHANNEL = "2010856600";

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

function userLinkDigest_(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8)
    .map(n => ("0" + ((n + 256) % 256).toString(16)).slice(-2)).join("");
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
    contacts: contacts.filter(row => /^U[0-9a-f]{32}$/.test(String(row[2]))).map(row => ({ id: String(row[2]), name: String(row[4] || "表示名なし"), message: String(row[11] || "").slice(0, 100) }))
  };
}

function userLinkCreateInvitation(input) {
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
    const existing = rows.map((r, i) => ({ r: r, i: i })).filter(x => x.r[1] === input.messagingId);
    if (existing.length > 1 || existing.some(x => String(x.r[0]) !== user.id)) userLinkFail_("このLINEには別の連携記録があります。上書きせず管理者へ確認してください。");
    let target = existing[0];
    if (target && (target.r[4] || target.r[5] === "連携済み")) userLinkFail_("連携済みのアカウントです。予約確認をご利用ください。");
    if (!target) {
      const prepared = rows.map((r, i) => ({ r: r, i: i })).filter(x => String(x.r[0]) === user.id && !x.r[1] && !x.r[4] && x.r[2] === relation && x.r[5] === "未連携");
      if (prepared.length === 1) target = prepared[0];
    }
    const token = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
    const expires = new Date(Date.now() + 86400000);
    const row = target ? target.r.slice(0, 14) : new Array(14).fill("");
    while (row.length < 14) row.push("");
    row[0] = user.id; row[1] = input.messagingId; row[2] = relation; row[3] = "確認済み";
    row[5] = "招待中"; row[6] = userLinkDigest_(token); row[7] = expires;
    row[9] = Session.getActiveUser().getEmail() || "シート編集者";
    row[11] = user.name; row[12] = matches[0].name;
    // No send call or sent timestamp: the operator sends the displayed text manually.
    const rowNumber = target ? target.i + 2 : sheet.getLastRow() + 1;
    if (rowNumber > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 100);
    sheet.getRange(rowNumber, 1, 1, 4).setValues([row.slice(0, 4)]);
    sheet.getRange(rowNumber, 6, 1, 3).setValues([row.slice(5, 8)]);
    sheet.getRange(rowNumber, 10).setValue(row[9]);
    sheet.getRange(rowNumber, 12, 1, 2).setValues([row.slice(11, 13)]);
    SpreadsheetApp.flush();
    return { recipient: matches[0].name, userName: user.name, message:
      "いつもご利用ありがとうございます。ひろがりです。\nLINEから訪問予定や回数券を確認するため、以下の専用リンクを24時間以内に開いてください。\nご本人・ご家族専用です。他の方へ転送しないでください。\n\nhttps://hirogari-toroku.github.io/reha1/?mode=user#link=" + token +
      "\n\n連携後は、画面下の「予定確認」から確認できます。操作が難しい場合は、このLINEでお知らせください。" };
  } finally { lock.releaseLock(); }
}

function userLinkVerify_(token) {
  if (typeof token !== "string" || !token || token.length > 4096) userLinkFail_("LINEでログインし直してください。");
  const verified = UrlFetchApp.fetch("https://api.line.me/oauth2/v2.1/verify?access_token=" + encodeURIComponent(token), { muteHttpExceptions: true });
  if (verified.getResponseCode() !== 200) userLinkFail_("LINEログインの有効期限が切れています。開き直してください。");
  const result = JSON.parse(verified.getContentText());
  if (String(result.client_id) !== USER_LINK_CHANNEL || !(Number(result.expires_in) > 0)) userLinkFail_("LINE認証を確認できません。");
  const response = UrlFetchApp.fetch("https://api.line.me/v2/profile", { headers: { Authorization: "Bearer " + token }, muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) userLinkFail_("LINEプロフィールを確認できません。");
  const profile = JSON.parse(response.getContentText());
  if (!/^U[0-9a-f]{32}$/.test(profile.userId || "")) userLinkFail_("LINE認証を確認できません。");
  return profile;
}

function userLinkClaim_(ss, token, profile) {
  if (!/^[0-9a-f]{64}$/.test(String(token))) userLinkFail_("専用リンクを確認してください。");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = userLinkSheet_(ss), rows = sheet.getDataRange().getValues().slice(1);
    const hash = userLinkDigest_(token);
    const matches = rows.map((r, i) => ({ r: r, i: i })).filter(x => x.r[6] === hash);
    if (matches.length !== 1) userLinkFail_("リンクは無効です。管理者へ再発行をご依頼ください。");
    const target = matches[0], row = target.r;
    if (row[3] !== "確認済み") userLinkFail_("管理者による本人・家族の確認が必要です。");
    userLinkMaster_(ss, String(row[0]));
    if (row[5] === "連携済み" && row[4] === profile.userId) return; // Lost-response retry by the same identity.
    if (row[5] !== "招待中" || row[4] || !(new Date(row[7]).getTime() > Date.now())) userLinkFail_("リンクは使用済み、または期限切れです。管理者へご連絡ください。");
    if (rows.some((other, i) => i !== target.i && (other[4] === profile.userId || other[1] === row[1]))) userLinkFail_("既存の連携があります。管理者へご確認ください。");
    // Only the link-state cells are updated; master data, names and notes remain intact.
    sheet.getRange(target.i + 2, 5, 1, 2).setValues([[profile.userId, "連携済み"]]);
    sheet.getRange(target.i + 2, 9).setValue(new Date());
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
}

function userLinkResolve_(ss, id) {
  const rows = userLinkSheet_(ss).getDataRange().getValues().slice(1).filter(row => row[4] === id);
  if (!rows.length) return null;
  if (rows.length !== 1 || rows[0][3] !== "確認済み" || rows[0][5] !== "連携済み") userLinkFail_("連携状態を管理者へご確認ください。");
  return userLinkMaster_(ss, String(rows[0][0]));
}

function userLinkRequest_(data) {
  try {
    const profile = userLinkVerify_(data.accessToken);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (data.inviteToken) userLinkClaim_(ss, data.inviteToken, profile);
    const user = userLinkResolve_(ss, profile.userId);
    if (!user) return { success: false, schedules: [], message: "まだ連携されていません。ひろがりから届いた専用リンクを開いてください。" };
    const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
    if (!scheduleSheet) userLinkFail_("訪問予定を確認できません。管理者へご連絡ください。");
    const visits = buildVisitStatusIndex_(ss.getSheetByName(VISIT_RESULT_SHEET_NAME));
    const items = collectActiveSchedulesForUser_(scheduleSheet, user.name, visits);
    // Return only the fields rendered to users; never expose chart or staff-folder URLs.
    const schedules = items.map(item => ({ visitDate: item.visitDate, status: item.status, staffName: item.staffName, updatedAt: item.updatedAt, kind: item.kind, lastVisitText: item.lastVisitText }));
    const coupon = getCouponDisplayMap_(ss)[normalizeName_(user.name)] || null;
    return { success: true, linked: true, userName: user.name, schedules: schedules, coupon: coupon, message: schedules.length ? "" : "対象期間内の予約はありません。" };
  } catch (error) {
    return { success: false, schedules: [], message: error.userLinkSafe ? error.message : "通信または連携処理に失敗しました。時間をおいて再度お試しください。" };
  }
}
