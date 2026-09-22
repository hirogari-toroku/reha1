// 初回訪問の候補日: the admin offers up to three dates/times for a new user's first
// visit, the user (or a confirmed family account) picks one in the booking LIFF, and the
// pick becomes a normal 訪問予定 row. Staff and admin are told by LINE.
const FIRST_VISIT_SHEET_NAME = "初回訪問候補";
const FIRST_VISIT_HEADERS = [
  "候補ID", "作成日時", "利用者ID", "利用者名", "スタッフID", "スタッフ名",
  "候補1", "候補2", "候補3", "状態", "選択候補", "選択日時", "選択者LIFF ID", "予定行", "通知結果"
];
const FIRST_VISIT_OPEN = "候補提示中";
const FIRST_VISIT_BOOKED = "確定";
const FIRST_VISIT_REPLACED = "差し替え";
const FIRST_VISIT_MAX_OPTIONS = 3;
const FIRST_VISIT_BOOKING_LINK = "https://liff.line.me/2010856600-yxOaV2Np";

function firstVisitSheet_(ss) {
  let sheet = ss.getSheetByName(FIRST_VISIT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FIRST_VISIT_SHEET_NAME);
    sheet.getRange(1, 1, 1, FIRST_VISIT_HEADERS.length).setValues([FIRST_VISIT_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function firstVisitRows_(sheet) {
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, FIRST_VISIT_HEADERS.length).getValues()
    .map((row, index) => ({ rowNumber: index + 2, row: row }));
}

// Accepts "yyyy-MM-ddTHH:mm" (datetime-local), "yyyy-MM-dd HH:mm", or a Date (Sheets turns
// the stored text into a date cell); returns a Date or null.
function parseFirstVisitOption_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return isNaN(value.getTime()) ? null : new Date(value.getTime());
  }
  const match = String(value || "").trim().match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})$/);
  if (!match) return null;
  const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), Number(match[4]), Number(match[5]));
  if (date.getMonth() !== Number(match[2]) - 1 || date.getDate() !== Number(match[3])) return null;
  return date;
}

function formatFirstVisitOption_(date) {
  const weekdays = ["日", "月", "火", "水", "木", "金", "土"];
  const pad = n => String(n).padStart(2, "0");
  return (date.getMonth() + 1) + "/" + date.getDate() + "（" + weekdays[date.getDay()] + "） " +
    date.getHours() + ":" + pad(date.getMinutes());
}

function firstVisitOptionKey_(date) {
  const pad = n => String(n).padStart(2, "0");
  return date.getFullYear() + "-" + pad(date.getMonth() + 1) + "-" + pad(date.getDate()) + " " +
    pad(date.getHours()) + ":" + pad(date.getMinutes());
}

function firstVisitOptionsFromRow_(row) {
  return [row[6], row[7], row[8]]
    .map((value, index) => ({ index: index + 1, date: parseFirstVisitOption_(value) }))
    .filter(option => option.date);
}

// Official-account (Messaging API) IDs of the user and family accounts an admin confirmed.
function getConfirmedUserMessagingIds_(ss, userId) {
  const sheet = userLinkSheet_(ss);
  const ids = sheet.getDataRange().getValues().slice(1)
    .filter(row => String(row[0]) === String(userId) && row[3] === "確認済み" && row[5] === "連携済み")
    .map(row => String(row[1] || "").trim())
    .filter(id => /^U[0-9a-f]{32}$/.test(id));
  return ids.filter((id, index) => ids.indexOf(id) === index);
}

function adminCreateFirstVisitCandidatesFromLiff_(lineUserId, displayName, params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!isAdminLiffUser_(ss, lineUserId)) return adminDeniedResponse_(lineUserId);

  let user;
  try {
    user = userLinkMaster_(ss, String(params.userId || ""));
  } catch (error) {
    if (error.userLinkSafe) return { success: false, message: error.message };
    throw error;
  }

  const staff = getAdminStaffMap_(ss).byId[String(params.staffId || "")];
  if (!staff) return { success: false, message: "担当スタッフを選んでください。" };
  if (staff.lineLinkPending) {
    return { success: false, message: staff.name + "さんのLINE紐づけが承認待ちです。先に承認してください。" };
  }

  const now = new Date();
  const options = [];
  const seen = {};
  for (let i = 1; i <= FIRST_VISIT_MAX_OPTIONS; i++) {
    const raw = String(params["option" + i] || "").trim();
    if (!raw) continue;
    const date = parseFirstVisitOption_(raw);
    if (!date) return { success: false, message: "候補" + i + "の日時を確認してください。" };
    if (date.getTime() <= now.getTime()) return { success: false, message: "候補" + i + "は過去の日時です。" };
    const key = firstVisitOptionKey_(date);
    if (seen[key]) continue;
    seen[key] = true;
    options.push(date);
  }
  if (!options.length) return { success: false, message: "候補日時を1つ以上入力してください。" };
  options.sort((a, b) => a.getTime() - b.getTime());

  let recipients = [];
  try {
    recipients = getConfirmedUserMessagingIds_(ss, user.id);
  } catch (error) {
    if (!error.userLinkSafe) throw error;
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let candidateId;
  let rowNumber;
  try {
    const sheet = firstVisitSheet_(ss);
    // One open offer per user: a new offer replaces the previous one.
    firstVisitRows_(sheet).forEach(item => {
      if (String(item.row[2]) === user.id && item.row[9] === FIRST_VISIT_OPEN) {
        sheet.getRange(item.rowNumber, 10).setValue(FIRST_VISIT_REPLACED);
      }
    });
    candidateId = "FV" + Utilities.formatDate(now, "Asia/Tokyo", "yyyyMMddHHmmss") + Math.floor(Math.random() * 90 + 10);
    const optionCells = [0, 1, 2].map(i => (options[i] ? firstVisitOptionKey_(options[i]) : ""));
    rowNumber = sheet.getLastRow() + 1;
    sheet.getRange(rowNumber, 1, 1, FIRST_VISIT_HEADERS.length).setValues([[
      candidateId, now, user.id, user.name, staff.id, staff.name,
      optionCells[0], optionCells[1], optionCells[2], FIRST_VISIT_OPEN, "", "", "", "", ""
    ]]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  const message =
    user.name + " 様\n" +
    "初回訪問の候補日をお送りします。ご都合のよい日時を1つお選びください。\n\n" +
    options.map((date, i) => "候補" + (i + 1) + "：" + formatFirstVisitOption_(date)).join("\n") +
    "\n\n下のリンクから選べます。\n" + FIRST_VISIT_BOOKING_LINK +
    "\n\nどの日程も難しい場合は、このLINEにご返信ください。";
  const results = recipients.map(id => sendPushMessage_(ss, id, "初回訪問候補", message));
  const sentCount = results.filter(result => result.success).length;
  const notifyText = recipients.length
    ? "LINE送信 " + sentCount + "/" + recipients.length + "件"
    : "LINE連携済みの利用者・家族がいないため未送信";
  firstVisitSheet_(ss).getRange(rowNumber, 15).setValue(notifyText);

  return {
    success: true,
    candidateId: candidateId,
    message: user.name + " 様へ候補日を" + options.length + "件作成しました（" + notifyText + "）。" +
      (recipients.length ? "" : "電話などで予約確認ページの案内をお願いします。")
  };
}

function getOpenFirstVisitCandidatesForAdmin_(ss) {
  const sheet = ss.getSheetByName(FIRST_VISIT_SHEET_NAME);
  if (!sheet) return [];
  return firstVisitRows_(sheet)
    .filter(item => item.row[9] === FIRST_VISIT_OPEN || item.row[9] === FIRST_VISIT_BOOKED)
    .slice(-20)
    .reverse()
    .map(item => {
      const selected = parseFirstVisitOption_(item.row[10]);
      return {
        candidateId: String(item.row[0]),
        userName: String(item.row[3]),
        staffName: String(item.row[5]),
        status: String(item.row[9]),
        options: firstVisitOptionsFromRow_(item.row).map(option => formatFirstVisitOption_(option.date)),
        selected: selected ? formatFirstVisitOption_(selected) : "",
        notifyResult: String(item.row[14] || "")
      };
    });
}

// Open offer for the verified user, shaped for the booking screen (no staff contact data).
function getFirstVisitOfferForUser_(ss, user) {
  const sheet = ss.getSheetByName(FIRST_VISIT_SHEET_NAME);
  if (!sheet) return null;
  const now = Date.now();
  const open = firstVisitRows_(sheet).filter(item => String(item.row[2]) === user.id && item.row[9] === FIRST_VISIT_OPEN);
  if (!open.length) return null;
  const item = open[open.length - 1];
  const options = firstVisitOptionsFromRow_(item.row)
    .filter(option => option.date.getTime() > now)
    .map(option => ({ index: option.index, label: formatFirstVisitOption_(option.date) }));
  if (!options.length) return null;
  return { candidateId: String(item.row[0]), staffName: String(item.row[5]), options: options };
}

function userSelectFirstVisitRequest_(data) {
  try {
    const profile = userLinkVerify_(data.accessToken);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const user = userLinkResolve_(ss, profile.userId);
    if (!user) userLinkFail_("管理者の確認が終わってから選べるようになります。");
    return bookFirstVisitOption_(ss, user, profile.userId, String(data.candidateId || ""), Number(data.optionIndex));
  } catch (error) {
    return { success: false, message: error.userLinkSafe ? error.message : "選択できませんでした。時間をおいて再度お試しください。" };
  }
}

function bookFirstVisitOption_(ss, user, liffLineUserId, candidateId, optionIndex) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let booked;
  try {
    const sheet = firstVisitSheet_(ss);
    const item = firstVisitRows_(sheet).filter(entry => String(entry.row[0]) === candidateId)[0];
    if (!item || String(item.row[2]) !== user.id) userLinkFail_("候補日が見つかりません。画面を開き直してください。");
    if (item.row[9] !== FIRST_VISIT_OPEN) userLinkFail_("この候補日はすでに確定済み、または差し替えられています。画面を開き直してください。");

    const option = firstVisitOptionsFromRow_(item.row).filter(entry => entry.index === optionIndex)[0];
    if (!option) userLinkFail_("選んだ候補を確認できません。画面を開き直してください。");
    if (option.date.getTime() <= Date.now()) userLinkFail_("この候補日時は過ぎています。別の候補を選ぶか、公式LINEでご連絡ください。");

    const staffName = String(item.row[5]);
    const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
    if (!scheduleSheet) userLinkFail_("予定を登録できません。公式LINEでご連絡ください。");
    ensureScheduleStatusColumns_(scheduleSheet);
    const now = new Date();
    const dateText = (option.date.getMonth() + 1) + "/" + option.date.getDate();
    const timeText = option.date.getHours() + ":" + String(option.date.getMinutes()).padStart(2, "0");
    let scheduleRow = "";
    if (!isDuplicateSchedule_(scheduleSheet, staffName, user.name, dateText)) {
      bumpReadCache_("assignmentSchedule");
      scheduleSheet.appendRow([
        now, staffName, user.name, dateText,
        "初回訪問（利用者選択） " + timeText,
        liffLineUserId, "", "", "予定", now, "初回訪問 " + timeText
      ]);
      scheduleRow = scheduleSheet.getLastRow();
    }

    sheet.getRange(item.rowNumber, 10, 1, 5).setValues([[
      FIRST_VISIT_BOOKED, firstVisitOptionKey_(option.date), now, liffLineUserId, scheduleRow
    ]]);
    SpreadsheetApp.flush();
    booked = { staffName: staffName, date: option.date, staffId: String(item.row[4]) };
  } finally {
    lock.releaseLock();
  }

  const label = formatFirstVisitOption_(booked.date);
  notifyFirstVisitBooked_(ss, user, booked, label);
  return {
    success: true,
    message: "初回訪問を " + label + " で予約しました。当日は担当スタッフ（" + booked.staffName + "）が伺います。"
  };
}

function notifyFirstVisitBooked_(ss, user, booked, label) {
  const text = "初回訪問の日程が決まりました\n利用者：" + user.name + " 様\n日時：" + label + "\n担当：" + booked.staffName;
  try {
    const staff = getAdminStaffMap_(ss).byName[normalizeName_(booked.staffName)];
    if (staff && staff.messagingLineUserId && !staff.lineLinkPending) {
      sendPushMessage_(ss, staff.messagingLineUserId, staff.name, text);
    }
    if (ADMIN_LINE_USER_ID) sendPushMessage_(ss, ADMIN_LINE_USER_ID, "管理者通知", text);
  } catch (error) {
    console.error("notifyFirstVisitBooked_ failed: " + error.message);
  }
}
