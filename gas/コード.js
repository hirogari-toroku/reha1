const RAW_LOG_SHEET_NAME = "LINE原文ログ";
const LINE_MESSAGE_LOG_SHEET_NAME = "LINE送受信ログ";
const SCHEDULE_SHEET_NAME = "訪問予定";
const VISIT_RESULT_SHEET_NAME = "訪問実績";
const SCHEDULE_VISIT_COMPARISON_SHEET_NAME = "予定・実績照合";
const STAFF_SHEET_NAME = "スタッフマスタ";
const STAFF_USER_MASTER_SHEET_NAME = "スタッフ利用者マスタ";
const LIFF_DISPLAY_MASTER_SHEET_NAME = "LIFF表示用マスタ";
const LINE_USER_DIRECTORY_SHEET_NAME = "LINEユーザー一覧";
const PAYROLL_FOLDER_ID = "1V29zEuKH4XnPy2cJUTBsv-mTfYKv2E_s";
const GMO_TRANSFER_CSV_CONFIRMED_FOLDER_ID = "1vTWdykjuj7fO27lmoCImPjy3TdNEb9Jp";
const GMO_PASTE_SHEET_NAME = "GMO入出金CSV貼付";
const COUPON_SHEET_NAME = "回数券管理";
const STAFF_QUESTIONNAIRE_SPREADSHEET_ID = "1sQzA9oHCMH712l8gNuoZxrB29eX_UeQg5zHgy_1_tQw";
const STAFF_QUESTIONNAIRE_SHEET_ID = 1876133724;
const DEFAULT_REHAB_UNIT_PRICE = 7500;
const COUPON_START_DATE_TEXT = "2026/04/20";
const PDF_EXPORT_WAIT_MS = 1000;
const PDF_EXPORT_MAX_RETRIES = 5;
const TEST_USER_NAME = "テスト 利用者";
const ADMIN_LINE_USER_ID = "Uc21fa34144f5bc50c6e5324d5e4de344";
const LIFF_DUPLICATE_VISIT_WINDOW_MS = 10 * 60 * 1000;
const LIFF_INIT_CACHE_SECONDS = 1800;
const LIFF_UNREGISTERED_CACHE_SECONDS = 30;
const LIFF_SCHEDULE_LOOKBACK_MONTHS = 2;
const LIFF_SCHEDULE_FUTURE_MONTHS = 6;

// スタッフマスタ固定列
// A:スタッフ名 B:LINE表示名 C:LINEユーザーID D:LIFF用LINEユーザーID E:カレンダーID
// F:銀行コード G:支店番号 H:預金種目 I:口座番号 J:受取人名 K:給与明細フォルダID L:住所
const STAFF_COL_NAME = 0;
const STAFF_COL_LINE_DISPLAY_NAME = 1;
const STAFF_COL_LINE_USER_ID = 2;
const STAFF_COL_LIFF_LINE_USER_ID = 3;
const STAFF_COL_CALENDAR_ID = 4;
const STAFF_COL_BANK_CODE = 5;
const STAFF_COL_BRANCH_CODE = 6;
const STAFF_COL_ACCOUNT_TYPE = 7;
const STAFF_COL_ACCOUNT_NUMBER = 8;
const STAFF_COL_RECEIVER_NAME = 9;
const STAFF_COL_PAYROLL_FOLDER_ID = 10;
const STAFF_COL_ADDRESS = 11;


/**
 * LINE Webhook
 * LINEからのメッセージ受信時に起動するメイン処理
 */
function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const replyMessages = [];
  const NL = String.fromCharCode(10);
  const rawSheet = ss.getSheetByName(RAW_LOG_SHEET_NAME);
  const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  const resultSheet = ss.getSheetByName(VISIT_RESULT_SHEET_NAME);

  const json = e && e.postData && e.postData.contents
    ? JSON.parse(e.postData.contents)
    : {};

  if (json.action) {
    return doGet({
      parameter: json
    });
  }

  const events = json.events || [];

  events.forEach(event => {
    if (event.type !== "message") return;
    if (event.message.type !== "text") return;

    const receivedAt = new Date();
    const userId = event.source.userId || "";
    const staffName = getStaffNameFromLineEvent_(ss, event);
    const isRegisteredStaff = staffName !== "未登録";
    const calendarId = getStaffCalendarIdFromLineEvent_(ss, event);
    const displayName = getLineDisplayNameFromEvent_(event);
    const text = event.message.text;
    const replyToken = event.replyToken;

    saveLineUserDirectory_(ss, {
      source: "公式LINEメッセージ",
      messagingLineUserId: userId,
      displayName: displayName,
      detectedStaffName: staffName,
      lastMessage: text,
      checkedAt: receivedAt
    });

    rawSheet.appendRow([receivedAt, staffName, userId, text]);
    saveLineMessageLog_(
      ss,
      receivedAt,
      "受信",
      staffName,
      userId,
      text
    );

    if (!isRegisteredStaff) {
      logStaffLookupFailure_(ss, userId, text);
      replyMessages.push({
        replyToken,
        userId,
        staffName,
        message:
          "スタッフが未登録です。" + NL +
          "管理者へ確認してください。" + NL +
          "LINEユーザーID：" + userId
      });
      return;
    }

    // 1. 訪問実績（開始・終了）の解析と登録
    const visitRows = parseVisitResult_(ss, text, receivedAt, staffName, userId);

    if (visitRows.length > 0) {
      const firstRow = visitRows[0];

      if (isAmbiguousUserName_(firstRow[3])) {
        replyMessages.push({
          replyToken,
          userId,
          staffName,
          message:
            "同じ苗字の利用者が複数います。" + NL +
            "フルネームで送信してください。" + NL +
            "入力内容：" + getAmbiguousInputName_(firstRow[3])
        });
        return;
      }

      visitRows.forEach(row => resultSheet.appendRow(row));

      if (
        firstRow[3] === "不明" ||
        !isRegisteredStaffUser_(ss, staffName, firstRow[3])
      ) {
        replyMessages.push({
          replyToken,
          userId,
          staffName,
          message:
            "利用者が未登録です。" + NL +
            "管理者へ確認してください。" + NL +
            "入力内容：" + firstRow[3]
        });
      } else {
        replyMessages.push({
          replyToken,
          userId,
          staffName,
          message:
            firstRow[1] === "開始"
              ? "よろしくお願いします。" + NL +
                "利用者：" + firstRow[3] + " 様"
              : "お疲れ様でした。" + NL +
                "利用者：" + firstRow[3] + " 様の実績を登録しました。"
        });
      }

      return;
    }

    // 2. 予定連絡の解析とカレンダー連携・予定シートへの書き込み
    const scheduleRows = parseSchedule_(ss, text, receivedAt, staffName, userId);

    if (scheduleRows.length > 0) {
      const firstRow = scheduleRows[0];

      // 同姓複数の曖昧さチェック
      if (isAmbiguousUserName_(firstRow[2])) {
        replyMessages.push({
          replyToken,
          userId,
          staffName,
          message:
            "同じ苗字の利用者が複数います。" + NL +
            "フルネームで送信してください。" + NL +
            "入力内容：" + getAmbiguousInputName_(firstRow[2])
        });
        return;
      }

      const registeredDates = [];
      const duplicateDates = [];

      scheduleRows.forEach(row => {
        // 重複チェック
        if (isDuplicateSchedule_(scheduleSheet, row[1], row[2], row[3])) {
          duplicateDates.push(row[3]);
          return;
        }

        let calendarStatus = "";
        let eventId = "";

        if (calendarId) {
          const result = createCalendarEvent_(calendarId, row[2], row[3]);
          calendarStatus = result.status;
          eventId = result.eventId;
        } else {
          calendarStatus = "カレンダーID未登録";
        }

        scheduleSheet.appendRow([
          row[0],
          row[1],
          row[2],
          row[3],
          row[4],
          row[5],
          calendarStatus,
          eventId
        ]);

        registeredDates.push(row[3]);
      });

      // 登録できた日程があれば、まとめて1通で返信する
      if (registeredDates.length > 0) {
        replyMessages.push({
          replyToken,
          userId,
          staffName,
          message:
            "予定を登録しました" + NL +
            "利用者：" + firstRow[2] + NL +
            "日付：" + registeredDates.join("、")
        });
      } else if (duplicateDates.length > 0) {
        // すべて重複していた場合もまとめて通知
        replyMessages.push({
          replyToken,
          userId,
          staffName,
          message:
            "重複のため登録しませんでした" + NL +
            "利用者：" + firstRow[2] + NL +
            "日付：" + duplicateDates.join("、")
        });
      }
    }
  });

  sendReplyMessages_(replyMessages);

  return ContentService.createTextOutput("OK");
}

function doGet(e) {
  const action = e && e.parameter ? e.parameter.action || "ping" : "ping";

  if (action === "ping") {
    return liffResponse_(e, {
      success: true,
      message: "Unified GAS LIFF API OK",
      time: new Date()
    });
  }

  if (action === "getUsers") {
    return liffResponse_(e, getLiffUserList_(e.parameter.lineUserId));
  }

  if (action === "init") {
    return liffResponse_(e, initLiffApp_(
      e.parameter.lineUserId,
      e.parameter.displayName
    ));
  }

  if (action === "logLogin") {
    return liffResponse_(e, logLiffLogin_(
      e.parameter.lineUserId,
      e.parameter.displayName
    ));
  }

  if (action === "recordVisit") {
    return liffResponse_(e, recordVisitFromLiff_(
      e.parameter.lineUserId,
      e.parameter.userName,
      e.parameter.visitType,
      e.parameter.visitDate,
      e.parameter.visitTime,
      e.parameter.scheduleId
    ));
  }

  if (action === "recordSchedule") {
    return liffResponse_(e, recordScheduleFromLiff_(
      e.parameter.lineUserId,
      e.parameter.userName,
      e.parameter.dates || e.parameter.visitDates
    ));
  }

  if (action === "getSchedules") {
    return liffResponse_(e, getSchedulesForLiff_(
      e.parameter.lineUserId
    ));
  }

  if (action === "cancelSchedule") {
    return liffResponse_(e, cancelScheduleFromLiff_(
      e.parameter.lineUserId,
      e.parameter.scheduleId
    ));
  }

  if (action === "updateSchedule") {
    return liffResponse_(e, updateScheduleFromLiff_(
      e.parameter.lineUserId,
      e.parameter.scheduleId,
      e.parameter.userName,
      e.parameter.visitDate || e.parameter.date
    ));
  }

  if (action === "testPush") {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return liffResponse_(e, testPushFromLiff_(ss, e.parameter));
  }

  return liffResponse_(e, {
    success: false,
    message: "不明なactionです: " + action
  });
}

function liffResponse_(e, data) {
  const callback = e && e.parameter ? e.parameter.callback : "";

  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(data) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function logLiffLogin_(lineUserId, displayName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffName = getStaffName_(ss, lineUserId);
  saveLineUserDirectory_(ss, {
    source: "LIFF",
    liffLineUserId: lineUserId,
    displayName: displayName,
    detectedStaffName: staffName,
    checkedAt: new Date()
  });
  const message =
    staffName === "未登録"
      ? "LIFFログインを確認しました。スタッフマスタのLIFF用LINEユーザーIDへ登録してください。表示名：" + (displayName || "")
      : "LIFFログインを確認しました。表示名：" + (displayName || "");

  if (staffName === "未登録") {
    saveUnregisteredLiffLogin_(ss, lineUserId, displayName, "logLogin");
  }

  return {
    success: staffName !== "未登録",
    staffName: staffName,
    lineUserId: lineUserId || "",
    displayName: displayName || "",
    message: message
  };
}

function initLiffApp_(lineUserId, displayName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const displayMasterData = getLiffInitDataFromDisplayMaster_(lineUserId);

  if (displayMasterData) {
    saveLineUserDirectory_(ss, {
      source: "LIFF",
      liffLineUserId: lineUserId,
      displayName: displayName,
      detectedStaffName: displayMasterData.staffName,
      checkedAt: new Date()
    });
    return {
      success: true,
      staffName: displayMasterData.staffName,
      users: displayMasterData.users,
      lineUserId: lineUserId || "",
      displayName: displayName || "",
      message: ""
    };
  }

  const staffName = getStaffNameCached_(ss, lineUserId);
  saveLineUserDirectory_(ss, {
    source: "LIFF",
    liffLineUserId: lineUserId,
    displayName: displayName,
    detectedStaffName: staffName,
    checkedAt: new Date()
  });

  if (staffName === "未登録") {
    const message = "スタッフが未登録です。管理者へ確認してください。\nLIFF用LINEユーザーID：" + lineUserId;

    saveUnregisteredLiffLogin_(ss, lineUserId, displayName, "init");

    return {
      success: false,
      staffName: "",
      users: [],
      lineUserId: lineUserId || "",
      displayName: displayName || "",
      message: message
    };
  }

  const users = getLiffUserListFast_(ss, staffName);

  return {
    success: true,
    staffName: staffName,
    users: users,
    lineUserId: lineUserId || "",
    displayName: displayName || "",
    message: ""
  };
}

function saveUnregisteredLiffLogin_(ss, lineUserId, displayName, action) {
  saveLiffOperationLog_(
    ss,
    action || "init",
    lineUserId,
    "未登録",
    "",
    "",
    "",
    "",
    "未登録",
    "LIFFログインを確認しました。スタッフマスタのLIFF用LINEユーザーIDへ登録してください。表示名：" + (displayName || ""),
    displayName || ""
  );

  notifyAdminUnregisteredLiffLogin_(ss, lineUserId, displayName, action);
}

function notifyAdminUnregisteredLiffLogin_(ss, lineUserId, displayName, action) {
  const targetLineUserId = String(lineUserId || "").trim();
  if (!targetLineUserId || !ADMIN_LINE_USER_ID) return;

  const props = PropertiesService.getScriptProperties();
  const notifiedKey = "UNREGISTERED_LIFF_ADMIN_NOTIFIED_" + targetLineUserId;

  if (props.getProperty(notifiedKey)) return;

  const token = props.getProperty("LINE_CHANNEL_ACCESS_TOKEN");

  if (!token) {
    saveLiffOperationLog_(
      ss,
      "notifyAdminUnregisteredLiff",
      targetLineUserId,
      "未登録",
      "",
      "",
      "",
      "",
      "失敗",
      "LINE_CHANNEL_ACCESS_TOKENが未設定のため、管理者へ初回ログイン通知を送信できませんでした。",
      displayName || ""
    );
    return;
  }

  const message =
    "未登録スタッフがWebアプリへ初回ログインしました。\n\n" +
    "LIFF用LINEユーザーID：\n" + targetLineUserId + "\n" +
    "LINE表示名：" + (displayName || "未取得") + "\n" +
    "操作：" + (action || "init") + "\n\n" +
    "スタッフマスタのD列「LIFF用LINEユーザーID」へ登録してください。";

  try {
    const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + token
      },
      payload: JSON.stringify({
        to: ADMIN_LINE_USER_ID,
        messages: [
          {
            type: "text",
            text: message
          }
        ]
      }),
      muteHttpExceptions: true
    });
    const responseCode = response.getResponseCode();
    const success = responseCode >= 200 && responseCode < 300;

    if (success) {
      props.setProperty(notifiedKey, new Date().toISOString());
    }

    saveLiffOperationLog_(
      ss,
      "notifyAdminUnregisteredLiff",
      targetLineUserId,
      "未登録",
      "",
      "",
      "",
      "",
      success ? "成功" : "失敗",
      success
        ? "管理者へ初回ログイン通知を送信しました。"
        : "管理者への初回ログイン通知に失敗しました：" + responseCode + " / " + response.getContentText(),
      displayName || ""
    );

    saveLineMessageLog_(
      ss,
      new Date(),
      success ? "送信" : "送信失敗",
      "管理者通知",
      ADMIN_LINE_USER_ID,
      success ? message : "管理者への初回ログイン通知に失敗：" + responseCode + " " + response.getContentText()
    );
  } catch (error) {
    saveLiffOperationLog_(
      ss,
      "notifyAdminUnregisteredLiff",
      targetLineUserId,
      "未登録",
      "",
      "",
      "",
      "",
      "失敗",
      error && error.stack ? error.stack : String(error),
      displayName || ""
    );
  }
}

function getLiffUserList_(lineUserId) {
  const displayMasterData = getLiffInitDataFromDisplayMaster_(lineUserId);

  if (displayMasterData) {
    return {
      success: true,
      staffName: displayMasterData.staffName,
      users: displayMasterData.users,
      message: ""
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffName = getStaffNameCached_(ss, lineUserId);

  if (staffName === "未登録") {
    const message = "スタッフが未登録です。管理者へ確認してください。\nLINEユーザーID：" + lineUserId;

    saveLiffOperationLog_(
      ss,
      "getUsers",
      lineUserId,
      staffName,
      "",
      "",
      "",
      "",
      "失敗",
      message
    );

    return {
      success: false,
      staffName: "",
      users: [],
      message: message
    };
  }

  const users = getLiffUserListFast_(ss, staffName);

  saveLiffOperationLog_(
    ss,
    "getUsers",
    lineUserId,
    staffName,
    "",
    "",
    "",
    "",
    "成功",
    "利用者取得：" + users.length + "件"
  );

  return {
    success: true,
    staffName: staffName,
    users: users,
    message: ""
  };
}

function getLiffInitDataFromDisplayMaster_(lineUserId) {
  const targetUserId = String(lineUserId || "").trim();
  if (!targetUserId) return null;

  const cache = CacheService.getScriptCache();
  const cacheKey = "liffInit:" + getLiffDisplayMasterCacheVersion_() + ":" + targetUserId;
  const cached = cache.get(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  const propertyData = getLiffInitDataFromProperties_(targetUserId);
  if (propertyData) {
    cache.put(cacheKey, JSON.stringify(propertyData), LIFF_INIT_CACHE_SECONDS);
    return propertyData;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LIFF_DISPLAY_MASTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();

  for (let i = 0; i < values.length; i++) {
    const rowUserId = String(values[i][0] || "").trim();
    if (rowUserId !== targetUserId) continue;

    const staffName = String(values[i][1] || "").trim();
    const userText = String(values[i][2] || "").trim();
    if (!staffName) return null;

    const users = userText
      .split("\n")
      .map(name => String(name || "").trim())
      .filter(name => name !== "")
      .map(name => ({ name: name }));
    const data = {
      staffName: staffName,
      users: users
    };

    cache.put(cacheKey, JSON.stringify(data), LIFF_INIT_CACHE_SECONDS);
    return data;
  }

  return null;
}

function getLiffInitDataFromProperties_(lineUserId) {
  const props = PropertiesService.getScriptProperties();
  const chunkCount = Number(props.getProperty("LIFF_DISPLAY_MASTER_JSON_CHUNK_COUNT") || 0);
  let json = "";

  if (chunkCount > 0) {
    for (let i = 0; i < chunkCount; i++) {
      json += props.getProperty("LIFF_DISPLAY_MASTER_JSON_" + i) || "";
    }
  } else {
    json = props.getProperty("LIFF_DISPLAY_MASTER_JSON") || "";
  }

  if (!json) return null;

  try {
    const map = JSON.parse(json);
    return map[lineUserId] || null;
  } catch (error) {
    return null;
  }
}

function saveLiffDisplayMasterProperties_(initDataMap) {
  const props = PropertiesService.getScriptProperties();
  const oldChunkCount = Number(props.getProperty("LIFF_DISPLAY_MASTER_JSON_CHUNK_COUNT") || 0);
  const json = JSON.stringify(initDataMap);
  const chunkSize = 8000;
  const chunks = [];

  for (let i = 0; i < json.length; i += chunkSize) {
    chunks.push(json.slice(i, i + chunkSize));
  }

  const properties = {
    LIFF_DISPLAY_MASTER_CACHE_VERSION: String(Date.now()),
    LIFF_DISPLAY_MASTER_JSON_CHUNK_COUNT: String(chunks.length)
  };

  chunks.forEach((chunk, index) => {
    properties["LIFF_DISPLAY_MASTER_JSON_" + index] = chunk;
  });

  props.setProperties(properties);
  props.deleteProperty("LIFF_DISPLAY_MASTER_JSON");

  for (let i = chunks.length; i < oldChunkCount; i++) {
    props.deleteProperty("LIFF_DISPLAY_MASTER_JSON_" + i);
  }
}

function getLiffUserListFast_(ss, staffName) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "liffUsers:" + normalizeName_(staffName);
  const cached = cache.get(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  const sheet = ss.getSheetByName(STAFF_USER_MASTER_SHEET_NAME);
  if (!sheet) {
    return [];
  }

  const values = sheet.getDataRange().getValues();
  const users = [];
  const seen = {};

  seen[normalizeName_(TEST_USER_NAME)] = true;
  users.push({
    name: TEST_USER_NAME
  });

  for (let i = 1; i < values.length; i++) {
    const masterStaff = values[i][0];
    const formalName = values[i][1];

    if (!formalName) continue;
    if (normalizeName_(masterStaff) !== normalizeName_(staffName)) continue;
    if (seen[normalizeName_(formalName)]) continue;

    seen[normalizeName_(formalName)] = true;
    users.push({
      name: formalName
    });
  }

  cache.put(cacheKey, JSON.stringify(users), LIFF_INIT_CACHE_SECONDS);
  return users;
}

function recordVisitFromLiff_(lineUserId, userName, visitType, visitDate, visitTime, scheduleId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultSheet = ss.getSheetByName(VISIT_RESULT_SHEET_NAME);
  const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  const staffName = getStaffName_(ss, lineUserId);
  const resolvedUserName = resolveUserName_(ss, staffName, userName);
  const type = visitType === "開始" ? "開始" : "終了";
  const now = new Date();
  const targetDate = normalizeLiffDateText_(visitDate, now);
  const targetTime = normalizeLiffTimeText_(visitTime, now);
  const originalText = "LIFF実績登録：" + resolvedUserName + " " + type;
  let linkedScheduleRow = null;

  if (!resultSheet) {
    const message = "訪問実績シートがありません。";

    saveLiffOperationLog_(
      ss,
      "recordVisit",
      lineUserId,
      staffName,
      userName,
      visitType,
      visitDate,
      visitTime,
      "失敗",
      message
    );

    return {
      success: false,
      message: message
    };
  }

  if (staffName === "未登録") {
    const message = "スタッフが未登録です。管理者へ確認してください。\nLINEユーザーID：" + lineUserId;

    saveLiffOperationLog_(
      ss,
      "recordVisit",
      lineUserId,
      staffName,
      userName,
      visitType,
      visitDate,
      visitTime,
      "失敗",
      message
    );

    return {
      success: false,
      message: message
    };
  }

  if (!userName) {
    const message = "利用者を選択してください。";

    saveLiffOperationLog_(
      ss,
      "recordVisit",
      lineUserId,
      staffName,
      userName,
      visitType,
      visitDate,
      visitTime,
      "失敗",
      message
    );

    return {
      success: false,
      message: message
    };
  }

  if (isAmbiguousUserName_(resolvedUserName)) {
    const message = "同じ苗字の利用者が複数います。フルネームで選択してください。";

    saveLiffOperationLog_(
      ss,
      "recordVisit",
      lineUserId,
      staffName,
      userName,
      visitType,
      visitDate,
      visitTime,
      "失敗",
      message
    );

    return {
      success: false,
      message: message
    };
  }

  if (!isRegisteredStaffUser_(ss, staffName, resolvedUserName)) {
    const message = "利用者が未登録です。管理者へ確認してください。";

    saveUnknownUser_(ss, now, staffName, resolvedUserName, originalText);
    saveLiffOperationLog_(
      ss,
      "recordVisit",
      lineUserId,
      staffName,
      resolvedUserName,
      visitType,
      visitDate,
      visitTime,
      "失敗",
      message
    );

    return {
      success: false,
      message: message
    };
  }

  if (scheduleId) {
    if (!scheduleSheet) {
      const message = "訪問予定シートがありません。";
      saveLiffOperationLog_(ss, "recordVisit:schedule", lineUserId, staffName, resolvedUserName, type, targetDate, targetTime, "失敗", message);
      return {
        success: false,
        message: message
      };
    }

    ensureScheduleStatusColumns_(scheduleSheet);
    const targetSchedule = getEditableScheduleRow_(scheduleSheet, staffName, scheduleId);
    if (!targetSchedule.success) {
      saveLiffOperationLog_(ss, "recordVisit:schedule", lineUserId, staffName, resolvedUserName, type, targetDate, targetTime, "失敗", targetSchedule.message);
      return targetSchedule;
    }

    linkedScheduleRow = targetSchedule.row;
    if (
      normalizeName_(linkedScheduleRow.userName) !== normalizeName_(resolvedUserName) ||
      linkedScheduleRow.visitDateValue !== targetDate
    ) {
      const message = "予定と実績の利用者または日付が一致しません。画面を更新してください。";
      saveLiffOperationLog_(ss, "recordVisit:schedule", lineUserId, staffName, resolvedUserName, type, targetDate, targetTime, "失敗", message);
      return {
        success: false,
        message: message
      };
    }
  }

  if (isRecentDuplicateVisit_(resultSheet, staffName, resolvedUserName, type, targetDate, targetTime, now)) {
    const message = "同じ内容の実績がすでに登録されています。重複登録を防止しました。";

    saveLiffOperationLog_(
      ss,
      "recordVisit:duplicate",
      lineUserId,
      staffName,
      resolvedUserName,
      type,
      targetDate,
      targetTime,
      "重複",
      message
    );

    return {
      success: false,
      duplicate: true,
      message: message
    };
  }

  resultSheet.appendRow([
    now,
    type,
    staffName,
    resolvedUserName,
    targetDate,
    targetTime,
    originalText,
    lineUserId
  ]);

  saveLineMessageLog_(
    ss,
    now,
    "LIFF",
    staffName,
    lineUserId,
    originalText
  );

  if (linkedScheduleRow) {
    updateLinkedScheduleVisitStatus_(scheduleSheet, linkedScheduleRow.rowNumber, type, targetTime, now);
  }

  const message = "利用者：" + resolvedUserName + " 様の" + type + "実績を登録しました。";

  saveLiffOperationLog_(
    ss,
    "recordVisit",
    lineUserId,
    staffName,
    resolvedUserName,
    type,
    targetDate,
    targetTime,
    "成功",
    message
  );

  sendVisitConfirmationPushFromLiff_(
    ss,
    lineUserId,
    staffName,
    resolvedUserName,
    type
  );

  return {
    success: true,
    message: message
  };
}

function recordScheduleFromLiff_(lineUserId, userName, dates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  const staffName = getStaffName_(ss, lineUserId);
  const calendarId = getStaffCalendarId_(ss, lineUserId);
  const now = new Date();

  if (!scheduleSheet) {
    const message = "訪問予定シートがありません。";

    saveLiffOperationLog_(
      ss,
      "recordSchedule",
      lineUserId,
      staffName,
      userName,
      "",
      dates,
      "",
      "失敗",
      message
    );

    return {
      success: false,
      message: message
    };
  }

  ensureScheduleStatusColumns_(scheduleSheet);

  if (staffName === "未登録") {
    const message = "スタッフが未登録です。管理者へ確認してください。\nLINEユーザーID：" + lineUserId;

    saveLiffOperationLog_(
      ss,
      "recordSchedule",
      lineUserId,
      staffName,
      userName,
      "",
      dates,
      "",
      "失敗",
      message
    );

    return {
      success: false,
      message: message
    };
  }

  if (!userName || !dates) {
    const message = "利用者と予定日を入力してください。";

    saveLiffOperationLog_(
      ss,
      "recordSchedule",
      lineUserId,
      staffName,
      userName,
      "",
      dates,
      "",
      "失敗",
      message
    );

    return {
      success: false,
      message: message
    };
  }

  const resolvedUserName = resolveUserName_(ss, staffName, userName);

  if (isAmbiguousUserName_(resolvedUserName)) {
    const message = "同じ苗字の利用者が複数います。フルネームで選択してください。";

    saveLiffOperationLog_(
      ss,
      "recordSchedule",
      lineUserId,
      staffName,
      userName,
      "",
      dates,
      "",
      "失敗",
      message
    );

    return {
      success: false,
      message: message
    };
  }

  if (!isRegisteredStaffUser_(ss, staffName, resolvedUserName)) {
    const message = "利用者が未登録です。管理者へ確認してください。";

    saveUnknownUser_(ss, now, staffName, resolvedUserName, "LIFF予定登録：" + userName + " " + dates);
    saveLiffOperationLog_(
      ss,
      "recordSchedule",
      lineUserId,
      staffName,
      resolvedUserName,
      "",
      dates,
      "",
      "失敗",
      message
    );

    return {
      success: false,
      message: message
    };
  }

  const dateList = splitScheduleDateText_(dates);

  if (dateList.length === 0) {
    const message = "訪問予定日を正しい形式で入力してください。";

    saveLiffOperationLog_(
      ss,
      "recordSchedule",
      lineUserId,
      staffName,
      resolvedUserName,
      "",
      dates,
      "",
      "失敗",
      message
    );

    return {
      success: false,
      message: message
    };
  }

  const registeredDates = [];
  const duplicateDates = [];

  dateList.forEach(dateText => {
    if (isDuplicateSchedule_(scheduleSheet, staffName, resolvedUserName, dateText)) {
      duplicateDates.push(dateText);
      return;
    }

    let calendarStatus = "";
    let eventId = "";

    if (calendarId) {
      const result = createCalendarEvent_(calendarId, resolvedUserName, dateText);
      calendarStatus = result.status;
      eventId = result.eventId;
    } else {
      calendarStatus = "カレンダーID未登録";
    }

    scheduleSheet.appendRow([
      now,
      staffName,
      resolvedUserName,
      dateText,
      "LIFF予定登録：" + dates,
      lineUserId,
      calendarStatus,
      eventId,
      "予定",
      now,
      ""
    ]);

    registeredDates.push(dateText);
  });

  const message =
    registeredDates.length > 0
      ? "予定を登録しました。日付：" + registeredDates.join("、")
      : "重複のため登録しませんでした。日付：" + duplicateDates.join("、");

  saveLiffOperationLog_(
    ss,
    "recordSchedule",
    lineUserId,
    staffName,
    resolvedUserName,
    "",
    dateList.join("、"),
    "",
    registeredDates.length > 0 ? "成功" : "重複",
    message
  );

  if (registeredDates.length > 0) {
    sendScheduleConfirmationPushFromLiff_(
      ss,
      lineUserId,
      staffName,
      resolvedUserName,
      registeredDates
    );
  }

  return {
    success: true,
    staffName: staffName,
    message: message
  };
}

function getSchedulesForLiff_(lineUserId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  const staffName = getStaffName_(ss, lineUserId);

  if (!scheduleSheet) {
    return {
      success: false,
      message: "訪問予定シートがありません。",
      schedules: []
    };
  }

  if (staffName === "未登録") {
    return {
      success: false,
      message: "スタッフが未登録です。管理者へ確認してください。\nLINEユーザーID：" + lineUserId,
      schedules: []
    };
  }

  ensureScheduleStatusColumns_(scheduleSheet);
  const schedules = collectActiveSchedulesForStaff_(scheduleSheet, staffName);

  return {
    success: true,
    staffName: staffName,
    schedules: schedules
  };
}

function cancelScheduleFromLiff_(lineUserId, scheduleId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  const staffName = getStaffName_(ss, lineUserId);
  const now = new Date();

  if (!scheduleSheet) {
    return {
      success: false,
      message: "訪問予定シートがありません。"
    };
  }

  ensureScheduleStatusColumns_(scheduleSheet);
  const target = getEditableScheduleRow_(scheduleSheet, staffName, scheduleId);

  if (!target.success) {
    saveLiffOperationLog_(ss, "cancelSchedule", lineUserId, staffName, "", "", "", "", "失敗", target.message);
    return target;
  }

  const row = target.row;
  const calendarId = getStaffCalendarId_(ss, lineUserId);
  const calendarStatus = cancelCalendarEvent_(calendarId, row.eventId);

  scheduleSheet.getRange(row.rowNumber, 9, 1, 3).setValues([[
    "キャンセル",
    now,
    "LIFFキャンセル：" + row.userName + " " + row.visitDate + " / " + calendarStatus
  ]]);
  if (calendarStatus) {
    scheduleSheet.getRange(row.rowNumber, 7).setValue(calendarStatus);
  }

  const message = "予定をキャンセルしました。利用者：" + row.userName + " 様、日付：" + row.visitDate;

  saveLiffOperationLog_(
    ss,
    "cancelSchedule",
    lineUserId,
    staffName,
    row.userName,
    "",
    row.visitDate,
    "",
    "成功",
    message
  );

  return {
    success: true,
    message: message
  };
}

function updateScheduleFromLiff_(lineUserId, scheduleId, userName, visitDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  const staffName = getStaffName_(ss, lineUserId);
  const calendarId = getStaffCalendarId_(ss, lineUserId);
  const now = new Date();

  if (!scheduleSheet) {
    return {
      success: false,
      message: "訪問予定シートがありません。"
    };
  }

  ensureScheduleStatusColumns_(scheduleSheet);
  const target = getEditableScheduleRow_(scheduleSheet, staffName, scheduleId);

  if (!target.success) {
    saveLiffOperationLog_(ss, "updateSchedule", lineUserId, staffName, userName, "", visitDate, "", "失敗", target.message);
    return target;
  }

  const dateList = splitScheduleDateText_(visitDate);
  if (dateList.length !== 1) {
    const message = "変更後の予定日を1つだけ選択してください。";
    saveLiffOperationLog_(ss, "updateSchedule", lineUserId, staffName, userName, "", visitDate, "", "失敗", message);
    return {
      success: false,
      message: message
    };
  }

  const resolvedUserName = resolveUserName_(ss, staffName, userName);
  if (isAmbiguousUserName_(resolvedUserName)) {
    const message = "同じ苗字の利用者が複数います。フルネームで選択してください。";
    saveLiffOperationLog_(ss, "updateSchedule", lineUserId, staffName, userName, "", visitDate, "", "失敗", message);
    return {
      success: false,
      message: message
    };
  }

  if (!isRegisteredStaffUser_(ss, staffName, resolvedUserName)) {
    const message = "利用者が未登録です。管理者へ確認してください。";
    saveUnknownUser_(ss, now, staffName, resolvedUserName, "LIFF予定変更：" + userName + " " + visitDate);
    saveLiffOperationLog_(ss, "updateSchedule", lineUserId, staffName, resolvedUserName, "", visitDate, "", "失敗", message);
    return {
      success: false,
      message: message
    };
  }

  const targetRow = target.row;
  const newDate = dateList[0];

  if (isDuplicateSchedule_(scheduleSheet, staffName, resolvedUserName, newDate, targetRow.rowNumber)) {
    const message = "同じ予定がすでに登録されています。重複登録を防止しました。";
    saveLiffOperationLog_(ss, "updateSchedule:duplicate", lineUserId, staffName, resolvedUserName, "", newDate, "", "重複", message);
    return {
      success: false,
      duplicate: true,
      message: message
    };
  }

  const calendarResult = updateCalendarEvent_(calendarId, targetRow.eventId, resolvedUserName, newDate);

  scheduleSheet.getRange(targetRow.rowNumber, 3, 1, 9).setValues([[
    resolvedUserName,
    newDate,
    "LIFF予定変更：" + targetRow.userName + " " + targetRow.visitDate + " → " + resolvedUserName + " " + newDate,
    lineUserId,
    calendarResult.status,
    calendarResult.eventId,
    "予定",
    now,
    "LIFF変更：" + targetRow.userName + " " + targetRow.visitDate + " → " + resolvedUserName + " " + newDate
  ]]);

  const message = "予定を変更しました。利用者：" + resolvedUserName + " 様、日付：" + newDate;

  saveLiffOperationLog_(
    ss,
    "updateSchedule",
    lineUserId,
    staffName,
    resolvedUserName,
    "",
    newDate,
    "",
    "成功",
    message
  );

  sendScheduleConfirmationPushFromLiff_(ss, lineUserId, staffName, resolvedUserName, [newDate]);

  return {
    success: true,
    message: message
  };
}

function collectActiveSchedulesForStaff_(sheet, staffName) {
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(sheet.getLastColumn(), 11)).getValues();
  const targetStaff = normalizeName_(staffName);
  const dateWindow = getLiffScheduleDateWindow_();
  const schedules = [];

  values.forEach((row, index) => {
    const rowNumber = index + 2;
    const rowStaffName = String(row[1] || "").trim();
    const userName = String(row[2] || "").trim();
    const visitDate = row[3];
    const status = String(row[8] || "予定").trim();

    if (normalizeName_(rowStaffName) !== targetStaff) return;
    if (!userName || !visitDate) return;
    if (isCancelledScheduleStatus_(status)) return;

    const date = parseComparisonDate_(visitDate, row[0]);
    if (!isDateInLiffScheduleWindow_(date, dateWindow)) return;

    schedules.push({
      id: String(rowNumber),
      userName: userName,
      visitDate: formatScheduleDateForLiff_(visitDate, row[0]),
      visitDateValue: formatScheduleDateValueForLiff_(visitDate, row[0]),
      status: status || "予定",
      registeredAt: formatComparisonDateTime_(row[0]),
      calendarStatus: String(row[6] || ""),
      updatedAt: formatComparisonDateTime_(row[9]),
      lastVisitText: String(row[10] || "")
    });
  });

  schedules.sort((a, b) => String(b.visitDateValue).localeCompare(String(a.visitDateValue)));
  return schedules;
}

function getLiffScheduleDateWindow_() {
  const today = new Date();
  const start = new Date(
    today.getFullYear(),
    today.getMonth() - LIFF_SCHEDULE_LOOKBACK_MONTHS,
    today.getDate()
  );
  const end = new Date(
    today.getFullYear(),
    today.getMonth() + LIFF_SCHEDULE_FUTURE_MONTHS,
    today.getDate()
  );

  return {
    startTime: new Date(start.getFullYear(), start.getMonth(), start.getDate()).getTime(),
    endTime: new Date(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59, 999).getTime()
  };
}

function isDateInLiffScheduleWindow_(date, dateWindow) {
  if (!date) return false;
  const time = date.getTime();
  if (isNaN(time)) return false;
  return time >= dateWindow.startTime && time <= dateWindow.endTime;
}

function ensureScheduleStatusColumns_(sheet) {
  const labels = [
    { column: 9, name: "状態" },
    { column: 10, name: "更新日時" },
    { column: 11, name: "更新内容" }
  ];

  labels.forEach(label => {
    const cell = sheet.getRange(1, label.column);
    if (!String(cell.getValue() || "").trim()) {
      cell.setValue(label.name);
    }
  });
}

function updateLinkedScheduleVisitStatus_(sheet, rowNumber, visitType, visitTime, now) {
  const type = visitType === "終了" ? "終了" : "開始";
  const nextStatus = type === "終了" ? "完了" : "訪問中";
  const note = "LIFF" + type + "実績：" + visitTime;

  sheet.getRange(rowNumber, 9, 1, 3).setValues([[
    nextStatus,
    now,
    note
  ]]);
}

function getEditableScheduleRow_(sheet, staffName, scheduleId) {
  if (staffName === "未登録") {
    return {
      success: false,
      message: "スタッフが未登録です。管理者へ確認してください。"
    };
  }

  const rowNumber = Number(scheduleId);
  if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    return {
      success: false,
      message: "予定が見つかりませんでした。画面を更新してください。"
    };
  }

  const row = sheet.getRange(rowNumber, 1, 1, Math.min(sheet.getLastColumn(), 11)).getValues()[0];
  const rowStaffName = String(row[1] || "").trim();
  const userName = String(row[2] || "").trim();
  const visitDate = row[3];
  const status = String(row[8] || "予定").trim();

  if (normalizeName_(rowStaffName) !== normalizeName_(staffName)) {
    return {
      success: false,
      message: "この予定を操作する権限がありません。"
    };
  }

  if (!userName || !visitDate || isCancelledScheduleStatus_(status)) {
    return {
      success: false,
      message: "予定が見つかりませんでした。画面を更新してください。"
    };
  }

  return {
    success: true,
    row: {
      rowNumber: rowNumber,
      registeredAt: row[0],
      staffName: rowStaffName,
      userName: userName,
      visitDate: formatScheduleDateForLiff_(visitDate, row[0]),
      visitDateValue: formatScheduleDateValueForLiff_(visitDate, row[0]),
      eventId: String(row[7] || "")
    }
  };
}

function isCancelledScheduleStatus_(status) {
  return /キャンセル|取消|中止/.test(String(status || "").trim());
}

function formatScheduleDateForLiff_(value, fallbackDate) {
  const date = parseComparisonDate_(value, fallbackDate);
  if (!date) return String(value || "");
  return Utilities.formatDate(date, "Asia/Tokyo", "M/d");
}

function formatScheduleDateValueForLiff_(value, fallbackDate) {
  const date = parseComparisonDate_(value, fallbackDate);
  if (!date) return "";
  return Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM-dd");
}

function isRecentDuplicateVisit_(sheet, staffName, userName, visitType, visitDate, visitTime, now) {
  if (!sheet || sheet.getLastRow() < 2) return false;

  const lastRow = sheet.getLastRow();
  const startRow = Math.max(2, lastRow - 99);
  const values = sheet
    .getRange(startRow, 1, lastRow - startRow + 1, Math.min(sheet.getLastColumn(), 8))
    .getValues();
  const targetStaff = normalizeName_(staffName);
  const targetUser = normalizeName_(userName);
  const targetType = String(visitType || "").trim();
  const targetDate = normalizeVisitDateKey_(visitDate);
  const targetTime = normalizeVisitTimeKey_(visitTime);
  const nowTime = now instanceof Date ? now.getTime() : new Date(now).getTime();

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const recordedAt = row[0] instanceof Date ? row[0].getTime() : new Date(row[0]).getTime();

    if (isNaN(recordedAt) || nowTime - recordedAt > LIFF_DUPLICATE_VISIT_WINDOW_MS) continue;
    if (String(row[1] || "").trim() !== targetType) continue;
    if (normalizeName_(row[2]) !== targetStaff) continue;
    if (normalizeName_(row[3]) !== targetUser) continue;
    if (normalizeVisitDateKey_(row[4]) !== targetDate) continue;
    if (normalizeVisitTimeKey_(row[5]) !== targetTime) continue;

    return true;
  }

  return false;
}

function normalizeVisitDateKey_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, "Asia/Tokyo", "yyyy-MM-dd");
  }

  const text = String(value || "").trim();
  if (!text) return "";

  const date = new Date(text);
  if (!isNaN(date.getTime()) && /^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}$/.test(text)) {
    return Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM-dd");
  }

  return text.replace(/\//g, "-");
}

function normalizeVisitTimeKey_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, "Asia/Tokyo", "HH:mm");
  }

  const text = String(value || "").trim();
  const match = text.match(/^(\d{1,2}):(\d{2})/);

  if (match) {
    return String(match[1]).padStart(2, "0") + ":" + match[2];
  }

  return text;
}

function saveLiffOperationLog_(ss, action, lineUserId, staffName, userName, visitType, visitDate, visitTime, result, message, displayName) {
  if (!ss) return;

  let sheet = ss.getSheetByName("LIFF操作ログ");

  if (!sheet) {
    sheet = ss.insertSheet("LIFF操作ログ");
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "日時",
      "操作",
      "LINEユーザーID",
      "スタッフ名",
      "利用者名",
      "種別",
      "実施日",
      "時刻",
      "結果",
      "メッセージ",
      "LINE表示名"
    ]);
  } else if (sheet.getLastColumn() < 11 || !sheet.getRange(1, 11).getValue()) {
    sheet.getRange(1, 11).setValue("LINE表示名");
  }

  sheet.appendRow([
    new Date(),
    action || "",
    lineUserId || "",
    staffName || "",
    userName || "",
    visitType || "",
    visitDate || "",
    visitTime || "",
    result || "",
    message || "",
    displayName || ""
  ]);
}

function normalizeLiffDateText_(value, fallbackDate) {
  const text = String(value || "").trim();
  const match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);

  if (match) {
    return Number(match[2]) + "/" + Number(match[3]);
  }

  const date = fallbackDate || new Date();
  return Utilities.formatDate(date, "Asia/Tokyo", "M/d");
}

function normalizeLiffTimeText_(value, fallbackDate) {
  const text = String(value || "").trim();

  if (/^\d{1,2}:\d{2}$/.test(text)) return text;

  const date = fallbackDate || new Date();
  return Utilities.formatDate(date, "Asia/Tokyo", "H:mm");
}

function getHeaderColumnMap_(sheet) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const map = {};

  headers.forEach((header, index) => {
    const key = String(header || "").trim();
    if (key) map[key] = index;

    const normalizedKey = normalizeHeaderName_(key);
    if (normalizedKey) map[normalizedKey] = index;
  });

  return map;
}

function getColumnIndex_(headerMap, names, fallbackIndex) {
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    if (headerMap[name] !== undefined) return headerMap[name];

    const normalizedName = normalizeHeaderName_(name);
    if (headerMap[normalizedName] !== undefined) return headerMap[normalizedName];
  }

  return fallbackIndex;
}

function normalizeHeaderName_(name) {
  return String(name || "")
    .replace(/[\s　]/g, "")
    .replace(/[：:;；、,。\.]/g, "")
    .trim();
}

function getStaffMasterColumnMap_(sheet) {
  const headerMap = getHeaderColumnMap_(sheet);

  return {
    name: getColumnIndex_(headerMap, ["スタッフ名", "氏名"], STAFF_COL_NAME),
    lineDisplayName: getColumnIndex_(headerMap, ["LINE表示名"], STAFF_COL_LINE_DISPLAY_NAME),
    lineUserId: getColumnIndex_(headerMap, ["LINEユーザーID", "Messaging API LINEユーザーID"], STAFF_COL_LINE_USER_ID),
    liffLineUserId: getColumnIndex_(headerMap, ["LIFF用LINEユーザーID", "LIFF LINEユーザーID"], STAFF_COL_LIFF_LINE_USER_ID),
    calendarId: getColumnIndex_(headerMap, ["カレンダーID", "GoogleカレンダーID"], STAFF_COL_CALENDAR_ID),
    bankCode: getColumnIndex_(headerMap, ["銀行コード"], STAFF_COL_BANK_CODE),
    branchCode: getColumnIndex_(headerMap, ["支店番号"], STAFF_COL_BRANCH_CODE),
    accountType: getColumnIndex_(headerMap, ["預金種目"], STAFF_COL_ACCOUNT_TYPE),
    accountNumber: getColumnIndex_(headerMap, ["口座番号"], STAFF_COL_ACCOUNT_NUMBER),
    receiverName: getColumnIndex_(headerMap, ["受取人名"], STAFF_COL_RECEIVER_NAME),
    payrollFolderId: getColumnIndex_(headerMap, ["給与明細フォルダID"], STAFF_COL_PAYROLL_FOLDER_ID),
    address: getColumnIndex_(headerMap, ["住所"], STAFF_COL_ADDRESS)
  };
}

function getStaffNameCached_(ss, userId) {
  const targetUserId = String(userId || "").trim();
  if (!targetUserId) return "未登録";

  const cache = CacheService.getScriptCache();
  const cacheKey = "staffName:" + targetUserId;
  const cached = cache.get(cacheKey);

  if (cached) return cached;

  const staffName = getStaffName_(ss, targetUserId);
  cache.put(
    cacheKey,
    staffName,
    staffName === "未登録" ? LIFF_UNREGISTERED_CACHE_SECONDS : LIFF_INIT_CACHE_SECONDS
  );
  return staffName;
}

/**
 * スタッフ名取得（Messaging APIまたはLIFFのLINEユーザーIDからマスタを走査）
 */
function getStaffName_(ss, userId) {
  const sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  if (!sheet) return "未登録";

  const values = sheet.getDataRange().getValues();
  const cols = getStaffMasterColumnMap_(sheet);
  const targetUserId = String(userId || "").trim();

  if (!targetUserId) return "未登録";

  const directStaffName = findStaffNameByAnyCellValue_(values, cols.name, targetUserId);
  if (directStaffName) return directStaffName;

  const lineUserIdCols = getUniqueIndexes_([
    cols.lineUserId,
    cols.liffLineUserId,
    STAFF_COL_LINE_USER_ID,
    STAFF_COL_LIFF_LINE_USER_ID
  ]);

  for (let i = 1; i < values.length; i++) {
    const staffName = values[i][cols.name];

    if (!staffName) continue;

    const matched = lineUserIdCols.some(col => {
      return String(values[i][col] || "").trim() === targetUserId;
    });

    if (matched) {
      return staffName;
    }
  }

  return "未登録";
}

function getStaffNameFromLineEvent_(ss, event) {
  const userId = event && event.source ? event.source.userId || "" : "";
  const staffName = getStaffName_(ss, userId);

  if (staffName !== "未登録") return staffName;

  const displayName = getLineDisplayNameFromEvent_(event);
  if (!displayName) return "未登録";

  return getStaffNameByDisplayName_(ss, displayName);
}

function getStaffCalendarIdFromLineEvent_(ss, event) {
  const userId = event && event.source ? event.source.userId || "" : "";
  const calendarId = getStaffCalendarId_(ss, userId);

  if (calendarId) return calendarId;

  const displayName = getLineDisplayNameFromEvent_(event);
  if (!displayName) return "";

  const staffName = getStaffNameByDisplayName_(ss, displayName);
  if (staffName === "未登録") return "";

  return getStaffCalendarIdByStaffName_(ss, staffName);
}

function getStaffNameByDisplayName_(ss, displayName) {
  const sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  if (!sheet) return "未登録";

  const values = sheet.getDataRange().getValues();
  const cols = getStaffMasterColumnMap_(sheet);
  const target = normalizeName_(displayName);
  const matches = [];

  for (let i = 1; i < values.length; i++) {
    const staffName = values[i][cols.name];
    const lineDisplayName = values[i][cols.lineDisplayName];

    if (!staffName) continue;

    if (
      normalizeName_(staffName) === target ||
      normalizeName_(lineDisplayName) === target
    ) {
      matches.push(staffName);
    }
  }

  const uniqueMatches = [...new Set(matches.filter(name => name))];

  if (uniqueMatches.length === 1) return uniqueMatches[0];

  return "未登録";
}

function getStaffCalendarIdByStaffName_(ss, staffName) {
  const sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  if (!sheet) return "";

  const values = sheet.getDataRange().getValues();
  const cols = getStaffMasterColumnMap_(sheet);
  const target = normalizeName_(staffName);

  for (let i = 1; i < values.length; i++) {
    if (normalizeName_(values[i][cols.name]) === target) {
      return values[i][cols.calendarId] || "";
    }
  }

  return "";
}

function getLineDisplayNameFromEvent_(event) {
  if (event && event._lineDisplayNameFetched) {
    return event._lineDisplayName || "";
  }

  const token = PropertiesService
    .getScriptProperties()
    .getProperty("LINE_CHANNEL_ACCESS_TOKEN");

  if (!token || !event || !event.source || !event.source.userId) {
    if (event) {
      event._lineDisplayNameFetched = true;
      event._lineDisplayName = "";
    }
    return "";
  }

  try {
    const source = event.source;
    let url = "";

    if (source.type === "group" && source.groupId) {
      url =
        "https://api.line.me/v2/bot/group/" +
        encodeURIComponent(source.groupId) +
        "/member/" +
        encodeURIComponent(source.userId);
    } else if (source.type === "room" && source.roomId) {
      url =
        "https://api.line.me/v2/bot/room/" +
        encodeURIComponent(source.roomId) +
        "/member/" +
        encodeURIComponent(source.userId);
    } else {
      url =
        "https://api.line.me/v2/bot/profile/" +
        encodeURIComponent(source.userId);
    }

    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: {
        Authorization: "Bearer " + token
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      event._lineDisplayNameFetched = true;
      event._lineDisplayName = "";
      return "";
    }

    const profile = JSON.parse(response.getContentText());
    event._lineDisplayNameFetched = true;
    event._lineDisplayName = profile.displayName || "";
    return event._lineDisplayName;

  } catch (error) {
    Logger.log(error);
    if (event) {
      event._lineDisplayNameFetched = true;
      event._lineDisplayName = "";
    }
    return "";
  }
}

/**
 * カレンダーIDの取得
 */
function getStaffCalendarId_(ss, userId) {
  const sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  if (!sheet) return "";

  const values = sheet.getDataRange().getValues();
  const cols = getStaffMasterColumnMap_(sheet);
  const targetUserId = String(userId || "").trim();

  if (!targetUserId) return "";

  const directMatchedRow = findRowByAnyCellValue_(values, targetUserId);
  if (directMatchedRow) return directMatchedRow[cols.calendarId] || "";

  const lineUserIdCols = getUniqueIndexes_([
    cols.lineUserId,
    cols.liffLineUserId,
    STAFF_COL_LINE_USER_ID,
    STAFF_COL_LIFF_LINE_USER_ID
  ]);

  for (let i = 1; i < values.length; i++) {
    const calendarId = values[i][cols.calendarId];

    const matched = lineUserIdCols.some(col => {
      return String(values[i][col] || "").trim() === targetUserId;
    });

    if (matched) {
      return calendarId || "";
    }
  }

  return "";
}

function logStaffLookupFailure_(ss, userId, text) {
  const sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  const targetUserId = String(userId || "").trim();
  let detail = "スタッフ照合失敗";

  if (!sheet) {
    detail += "：スタッフマスタシートがありません";
  } else {
    const values = sheet.getDataRange().getValues();
    const header = values.length > 0 ? values[0].join(" / ") : "";
    const matchedRowNumber = findRowNumberByAnyCellValue_(values, targetUserId);

    detail += "：LINEユーザーID=" + targetUserId;
    detail += " / スタッフマスタ内一致行=" + (matchedRowNumber || "なし");
    detail += " / ヘッダー=" + header;
  }

  saveLineMessageLog_(
    ss,
    new Date(),
    "診断",
    "システム",
    targetUserId,
    detail + " / 入力=" + String(text || "")
  );
}

function findStaffNameByAnyCellValue_(values, staffNameCol, targetValue) {
  const matchedRow = findRowByAnyCellValue_(values, targetValue);
  if (!matchedRow) return "";

  return matchedRow[staffNameCol] || "";
}

function findRowByAnyCellValue_(values, targetValue) {
  const target = String(targetValue || "").trim();
  if (!target) return null;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    for (let j = 0; j < row.length; j++) {
      if (String(row[j] || "").trim() === target) {
        return row;
      }
    }
  }

  return null;
}

function findRowNumberByAnyCellValue_(values, targetValue) {
  const target = String(targetValue || "").trim();
  if (!target) return "";

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    for (let j = 0; j < row.length; j++) {
      if (String(row[j] || "").trim() === target) {
        return i + 1;
      }
    }
  }

  return "";
}

function getUniqueIndexes_(indexes) {
  const map = {};
  const result = [];

  indexes.forEach(index => {
    const number = Number(index);

    if (isNaN(number) || number < 0 || map[number]) return;

    map[number] = true;
    result.push(number);
  });

  return result;
}

/**
 * Googleカレンダー予定登録（重複防止機能付き）
 */
function createCalendarEvent_(calendarId, userName, dateText) {
  try {
    const calendar = CalendarApp.getCalendarById(calendarId);

    if (!calendar) {
      return {
        status: "カレンダー取得失敗",
        eventId: ""
      };
    }

    const date = convertDateTextToDate_(dateText);
    const title = userName + "様 訪問";

    const events = calendar.getEventsForDay(date);
    const duplicate = events.some(event => event.getTitle() === title);

    if (duplicate) {
      return {
        status: "重複のため未登録",
        eventId: ""
      };
    }

    const event = calendar.createAllDayEvent(title, date, {
      description: "LINE予定連絡から自動登録"
    });

    return {
      status: "登録済",
      eventId: event.getId()
    };

  } catch (error) {
    return {
      status: "エラー：" + error.message,
      eventId: ""
    };
  }
}

function updateCalendarEvent_(calendarId, eventId, userName, dateText) {
  if (!calendarId) {
    return {
      status: "カレンダーID未登録",
      eventId: eventId || ""
    };
  }

  try {
    const calendar = CalendarApp.getCalendarById(calendarId);

    if (!calendar) {
      return {
        status: "カレンダー取得失敗",
        eventId: eventId || ""
      };
    }

    const date = convertDateTextToDate_(dateText);
    const title = userName + "様 訪問";
    let event = eventId ? calendar.getEventById(eventId) : null;

    if (event) {
      event.setTitle(title);
      event.setAllDayDate(date);
      return {
        status: "変更済",
        eventId: event.getId()
      };
    }

    const created = createCalendarEvent_(calendarId, userName, dateText);
    return {
      status: "再登録：" + created.status,
      eventId: created.eventId
    };

  } catch (error) {
    return {
      status: "変更エラー：" + error.message,
      eventId: eventId || ""
    };
  }
}

function cancelCalendarEvent_(calendarId, eventId) {
  if (!calendarId) return "カレンダーID未登録";
  if (!eventId) return "カレンダー予定IDなし";

  try {
    const calendar = CalendarApp.getCalendarById(calendarId);

    if (!calendar) return "カレンダー取得失敗";

    const event = calendar.getEventById(eventId);

    if (!event) return "カレンダー予定なし";

    event.deleteEvent();
    return "キャンセル済";

  } catch (error) {
    return "キャンセルエラー：" + error.message;
  }
}

function convertDateTextToDate_(dateText) {
  const now = new Date();
  const year = now.getFullYear();
  const parts = dateText.split("/");

  return new Date(
    year,
    Number(parts[0]) - 1,
    Number(parts[1])
  );
}

/**
 * 訪問実績（LINEテキスト）の解析
 */
function parseVisitResult_(ss, text, receivedAt, staffName, userId) {
  const rows = [];

  // 「7月訪問予定」などの予定連絡を、訪問開始の実績として誤判定しない
  if (isScheduleLikeMessage_(text)) return rows;

  const compact = text.replace(/\s/g, "");

  let type = "";
  let userName = "";

  const targetDate = Utilities.formatDate(receivedAt, "Asia/Tokyo", "M/d");
  const targetTime = Utilities.formatDate(receivedAt, "Asia/Tokyo", "H:mm");

  if (/終了|終わり|完了|終えました|終わりました/.test(compact)) {
    type = "終了";
  } else if (/開始|始めます|入ります|入りました|介入|訪問|リハ開始|リハビリ開始/.test(compact)) {
    type = "開始";
  }

  if (!type) return rows;

  userName = extractUserNameFromText_(compact);

  // 「終了」連絡で名前が省略されている場合、直前の「開始」ユーザーを特定する
  if (!userName && type === "終了") {
    userName = findLastStartedUser_(ss, userId);
  }

  if (userName) {
    userName = resolveUserName_(ss, staffName, userName);
  }

  if (!userName) {
    userName = "不明";
  }

  // マスタ未登録または不明の場合は専用シートに記録
  if (
    userName === "不明" ||
    !isRegisteredStaffUser_(ss, staffName, userName)
  ) {
    saveUnknownUser_(
      ss,
      receivedAt,
      staffName,
      userName,
      text
    );
  }

  rows.push([
    receivedAt,
    type,
    staffName,
    userName,
    targetDate,
    targetTime,
    text,
    userId
  ]);

  return rows;
}

function isScheduleLikeMessage_(text) {
  const normalized = String(text || "").replace(/\r/g, "");
  const hasDate = extractDates_(normalized).length > 0;

  if (!hasDate) return false;

  return /予定|訪問予定|送らせて|送ります|お願い致します|お願いいたします|お願いします/.test(normalized);
}

/**
 * メッセージからユーザー名を抽出（ノイズ除去）
 */
function extractUserNameFromText_(text) {
  let name = text;

  name = name
    .replace(/只今より/g, "")
    .replace(/リハビリ/g, "")
    .replace(/リハ/g, "")
    .replace(/介入を始めます/g, "")
    .replace(/介入開始します/g, "")
    .replace(/介入します/g, "")
    .replace(/開始しました/g, "")
    .replace(/開始します/g, "")
    .replace(/始めます/g, "")
    .replace(/入ります/g, "")
    .replace(/入りました/g, "")
    .replace(/訪問します/g, "")
    .replace(/訪問しました/g, "")
    .replace(/終了しました/g, "")
    .replace(/終了します/g, "")
    .replace(/終わりました/g, "")
    .replace(/完了しました/g, "")
    .replace(/次回.*$/g, "")
    .replace(/[!！。]/g, "")
    .trim();

  const match = name.match(/(.+?様)/);

  if (match) {
    return cleanupUserName_(match[1]);
  }

  return cleanupUserName_(name);
}

/**
 * 直前に「開始」した利用者を逆引き
 */
function findLastStartedUser_(ss, userId) {
  const sheet = ss.getSheetByName(VISIT_RESULT_SHEET_NAME);
  if (!sheet) return "";

  const values = sheet.getDataRange().getValues();

  for (let i = values.length - 1; i >= 1; i--) {
    const type = values[i][1];
    const userName = values[i][3];
    const rowUserId = values[i][7];

    if (
      rowUserId === userId &&
      type === "開始" &&
      userName &&
      userName !== "不明"
    ) {
      return userName;
    }
  }

  return "";
}

/**
 * 訪問予定の解析（複数行・複数日付対応）
 */
function parseSchedule_(ss, text, receivedAt, staffName, userId) {
  const rows = [];

  const lines = text
    .replace(/\r/g, "")
    .split("\n")
    .map(line => line.trim())
    .filter(line => line !== "");

  let currentUserName = "";

  lines.forEach(line => {
    const dates = extractDates_(line);

    if (dates.length === 0) {
      if (isUserNameLine_(line)) {
        const candidateName = extractScheduleUserNameFromLine_(line);
        if (candidateName) {
          currentUserName = resolveUserName_(ss, staffName, candidateName);
        }
      }
      return;
    }

    const lineUserName = extractScheduleUserNameFromLine_(line);
    if (lineUserName) {
      currentUserName = resolveUserName_(ss, staffName, lineUserName);
    }

    if (!currentUserName) return;

    dates.forEach(date => {
      rows.push([
        receivedAt,
        staffName,
        currentUserName,
        date,
        text,
        userId
      ]);
    });
  });

  return rows;
}

function extractScheduleUserNameFromLine_(line) {
  const text = String(line || "").trim();

  if (!text) return "";
  if (/お疲れ|ありがとう|連絡|予定|よろしく|お願い|訪問/.test(text)) return "";

  const match = text.match(/^(.+?(?:様|さん|氏))/);
  if (match) return cleanupUserName_(match[1]);

  const withoutDates = text
    .replace(/(\d{1,2})\s*[\/月]\s*(\d{1,2})\s*日?/g, "")
    .replace(/[、,，]/g, "")
    .trim();

  if (!withoutDates || withoutDates.length > 10) return "";

  return cleanupUserName_(withoutDates);
}

/**
 * 予定の重複チェック
 */
function isDuplicateSchedule_(sheet, staffName, userName, dateText, excludeRowNumber) {
  if (!sheet) return false;

  const values = sheet.getDataRange().getValues();
  const targetDateKey = normalizeScheduleDateKey_(dateText);

  for (let i = 1; i < values.length; i++) {
    const rowNumber = i + 1;
    const rowStaff = values[i][1];
    const rowUser = values[i][2];
    const rowDate = values[i][3];
    const rowStatus = values[i][8];

    if (excludeRowNumber && rowNumber === Number(excludeRowNumber)) continue;
    if (isCancelledScheduleStatus_(rowStatus)) continue;

    if (
      normalizeName_(rowStaff) === normalizeName_(staffName) &&
      normalizeName_(rowUser) === normalizeName_(userName) &&
      normalizeScheduleDateKey_(rowDate) === targetDateKey
    ) {
      return true;
    }
  }

  return false;
}

function normalizeScheduleDateKey_(value) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, "Asia/Tokyo", "M/d");
  }

  const text = String(value).trim();
  const match = text.match(/(\d{1,2})\s*[\/月]\s*(\d{1,2})\s*日?/);

  if (match) {
    return Number(match[1]) + "/" + Number(match[2]);
  }

  const date = new Date(text);
  if (!isNaN(date.getTime())) {
    return Utilities.formatDate(date, "Asia/Tokyo", "M/d");
  }

  return text;
}

/**
 * 表記揺れ・苗字省略の解決（マスタ照合）
 */
function resolveUserName_(ss, staffName, inputName) {
  if (isTestUserName_(inputName)) return TEST_USER_NAME;

  const sheet = ss.getSheetByName(STAFF_USER_MASTER_SHEET_NAME);
  if (!sheet) return cleanupUserName_(inputName);

  const values = sheet.getDataRange().getValues();
  const target = normalizeName_(inputName);
  const exactMatches = [];
  const partialMatches = [];

  for (let i = 1; i < values.length; i++) {
    const masterStaff = values[i][0];
    const formalName = values[i][1];
    const lineName = values[i][2];

    if (normalizeName_(masterStaff) !== normalizeName_(staffName)) continue;
    if (!formalName) continue;

    const formalNormalized = normalizeName_(formalName);
    const lineNormalized = normalizeName_(lineName);

    if (formalNormalized === target || lineNormalized === target) {
      exactMatches.push(formalName);
      continue;
    }

    // 苗字だけの入力にも前方一致で対応（例:「白井」→「白井裕二」）
    if (
      target &&
      (
        formalNormalized.indexOf(target) === 0 ||
        lineNormalized.indexOf(target) === 0
      )
    ) {
      partialMatches.push(formalName);
    }
  }

  const uniqueExactMatches = [...new Set(exactMatches.filter(name => name))];
  if (uniqueExactMatches.length === 1) return uniqueExactMatches[0];
  if (uniqueExactMatches.length >= 2) return "要フルネーム:" + cleanupUserName_(inputName);

  const uniquePartialMatches = [...new Set(partialMatches.filter(name => name))];
  if (uniquePartialMatches.length === 1) return uniquePartialMatches[0];
  if (uniquePartialMatches.length >= 2) return "要フルネーム:" + cleanupUserName_(inputName);

  return cleanupUserName_(inputName);
}

function isAmbiguousUserName_(userName) {
  return String(userName || "").indexOf("要フルネーム:") === 0;
}

function getAmbiguousInputName_(userName) {
  return String(userName || "").replace("要フルネーム:", "");
}

function normalizeName_(name) {
  return String(name || "")
    .replace(/\s/g, "")
    .replace(/様|さん|氏/g, "")
    .replace(/[：:;；、,。\.]/g, "")
    .trim();
}

function isUserNameLine_(line) {
  if (extractDates_(line).length > 0) return false;
  if (line.length > 25) return false;

  return /様$|さん$/.test(line) || line.length <= 10;
}

function cleanupUserName_(name) {
  return String(name || "")
    .replace(/\s/g, "")
    .replace(/様|さん|氏/g, "")
    .replace(/[：:;；、,。\.]/g, "")
    .trim();
}

function extractDates_(text) {
  const dates = [];
  const regex = /(\d{1,2})\s*[\/月]\s*(\d{1,2})\s*日?/g;

  let match;

  while ((match = regex.exec(text)) !== null) {
    dates.push(match[1] + "/" + match[2]);
  }

  return dates;
}

function splitScheduleDateText_(text) {
  const normalized = String(text || "")
    .replace(/[，、\n\r]/g, ",")
    .replace(/\s+/g, ",");

  const parts = normalized
    .split(",")
    .map(value => value.trim())
    .filter(Boolean);

  const dates = [];

  parts.forEach(part => {
    const normalizedDate = normalizeScheduleDateText_(part);
    if (normalizedDate) dates.push(normalizedDate);
  });

  if (dates.length > 0) return [...new Set(dates)];

  return [...new Set(extractDates_(text))];
}

function normalizeScheduleDateText_(value) {
  const text = String(value || "").trim();
  let match = text.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);

  if (match) {
    return Number(match[2]) + "/" + Number(match[3]);
  }

  match = text.match(/^(\d{1,2})\s*[\/月]\s*(\d{1,2})\s*日?$/);

  if (match) {
    return Number(match[1]) + "/" + Number(match[2]);
  }

  return "";
}

/**
 * 給与振込日算出（対象月の翌月15日、土日祝の場合は前営業日）
 */
function getPayrollTransferDate_(ym) {
  const parts = ym.split("-");
  const year = Number(parts[0]);
  const month = Number(parts[1]);

  let payDate = new Date(year, month, 15);

  while (isHolidayOrWeekend_(payDate)) {
    payDate.setDate(payDate.getDate() - 1);
  }

  return Utilities.formatDate(payDate, "Asia/Tokyo", "yyyy/MM/dd");
}

function isHolidayOrWeekend_(date) {
  const day = date.getDay();

  if (day === 0 || day === 6) return true;

  const holidayCalendar = CalendarApp.getCalendarById(
    "ja.japanese#holiday@group.v.calendar.google.com"
  );

  if (!holidayCalendar) return false;

  const events = holidayCalendar.getEventsForDay(date);
  return events.length > 0;
}

/**
 * 手当マスタ（給与手当入力）のマッピング情報作成
 */
function getPayrollAllowanceMap_(ss) {
  const sheet = ss.getSheetByName("給与手当入力");
  const map = {};

  if (!sheet) return map;

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const ym = String(values[i][0] || "").trim();
    const staffName = String(values[i][1] || "").trim();

    if (!ym || !staffName) continue;

    const key = ym + "_" + normalizeName_(staffName);

    map[key] = {
      incentive: Number(values[i][2]) || 0,
      referral: Number(values[i][3]) || 0,
      management: Number(values[i][4]) || 0,
      memo: values[i][5] || ""
    };
  }

  return map;
}

function getStaffFormalNameMap_(ss) {
  const sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  const map = {};

  if (!sheet) return map;

  const values = sheet.getDataRange().getValues();
  const cols = getStaffMasterColumnMap_(sheet);

  for (let i = 1; i < values.length; i++) {
    const staffName = values[i][cols.name];
    if (!staffName) continue;

    map[normalizeName_(staffName)] = staffName;
  }

  return map;
}

/**
 * 給与集計処理（メインロジック）
 */
function createPayrollSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const visitSheet = ss.getSheetByName("訪問実績");
  const payrollSheet = ss.getSheetByName("給与集計");

  const masterMap = getStaffUserMasterMap_(ss);
  const allowanceMap = getPayrollAllowanceMap_(ss);
  const staffFormalNameMap = getStaffFormalNameMap_(ss);
  const values = visitSheet.getDataRange().getValues();

  const TAX_RATE = 0.03063; // 源泉徴収税率
  const summary = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const receivedAt = row[0];
    const type = row[1];
    const staffName = row[2];
    const userName = row[3];

    if (type !== "終了") continue;
    if (!staffName || !userName) continue;

    const staffKey = normalizeName_(staffName);
    const formalStaffName = staffFormalNameMap[staffKey] || staffName;

    const ym = Utilities.formatDate(
      new Date(receivedAt),
      "Asia/Tokyo",
      "yyyy-MM"
    );

    const key = ym + "_" + staffKey;

    if (!summary[key]) {
      summary[key] = {
        ym: ym,
        staffName: formalStaffName,
        count: 0,
        basePay: 0,
        travelCost: 0,
        missing: []
      };
    }

    const resolvedUserName = resolveUserName_(ss, staffName, userName);

    if (isAmbiguousUserName_(resolvedUserName)) {
      summary[key].missing.push(userName + "（同姓複数の可能性あり）");
      continue;
    }

    const masterKey = normalizeName_(staffName) + "_" + normalizeName_(resolvedUserName);
    const master = masterMap[masterKey];

    if (!master) {
      summary[key].missing.push(userName);
      continue;
    }

    summary[key].count += 1;
    summary[key].basePay += master.unitPay;
    summary[key].travelCost += master.travelCost;
  }

  payrollSheet.clearContents();

  const output = [[
    "年月",
    "振込日",
    "スタッフ名",
    "訪問件数",
    "基本給",
    "交通費",
    "奨励手当",
    "紹介手当",
    "運営報酬",
    "課税対象額",
    "所得税",
    "支給額",
    "備考"
  ]];

  Object.keys(summary).forEach(key => {
    const item = summary[key];
    const transferDate = getPayrollTransferDate_(item.ym);
    const allowanceKey = item.ym + "_" + normalizeName_(item.staffName);
    const allowance = allowanceMap[allowanceKey] || {
      incentive: 0,
      referral: 0,
      management: 0,
      memo: ""
    };

    const taxablePay = item.basePay + allowance.incentive + allowance.referral + allowance.management;
    const tax = Math.floor(taxablePay * TAX_RATE);
    const payment = taxablePay + item.travelCost - tax;

    output.push([
      item.ym,
      transferDate,
      item.staffName,
      item.count,
      item.basePay,
      item.travelCost,
      allowance.incentive,
      allowance.referral,
      allowance.management,
      taxablePay,
      tax,
      payment,
      allowance.memo || [...new Set(item.missing)].join("、")
    ]);
  });

  payrollSheet
    .getRange(1, 1, output.length, output[0].length)
    .setValues(output);

  payrollSheet
    .getRange(1, 1, 1, output[0].length)
    .setFontWeight("bold");

  if (output.length > 1) {
    payrollSheet
      .getRange(2, 1, output.length - 1, output[0].length)
      .setFontWeight("normal");
  }
}

function getStaffUserMasterMap_(ss) {
  const sheet = ss.getSheetByName(STAFF_USER_MASTER_SHEET_NAME);
  const map = {};

  if (!sheet) return map;

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const staffName = values[i][0];
    const userName = values[i][1];
    const unitPay = Number(values[i][3]);

    const manualTravelCost = Number(values[i][4]) || 0;
    const oneWayDistanceKm = Number(values[i][5]) || 0;
    const travelCost = calculatePayrollTravelCost_(oneWayDistanceKm, manualTravelCost);

    if (!staffName || !userName) continue;

    map[normalizeName_(staffName) + "_" + normalizeName_(userName)] = {
      unitPay: unitPay || 0,
      travelCost: travelCost || 0
    };
  }

  return map;
}

/**
 * 距離に応じたスタッフ給与交通費の算出規則
 * - 10km未満：片道距離 × 20円
 * - 10km以上：一律200円
 */
function calculatePayrollTravelCost_(oneWayDistanceKm, fallbackTravelCost) {
  if (!oneWayDistanceKm) return fallbackTravelCost || 0;

  if (oneWayDistanceKm < 10) {
    return Math.round(oneWayDistanceKm * 20);
  }

  return 200;
}

/**
 * 距離に応じた利用者直接請求交通費の算出規則
 * - 10km未満：請求なし（0円）
 * - 10km以上：5kmごとに200円加算
 */
function calculateDirectTravelCostForUser_(oneWayDistanceKm) {
  if (!oneWayDistanceKm) return 0;

  if (oneWayDistanceKm < 10) return 0;

  return Math.floor((oneWayDistanceKm - 10) / 5) * 200 + 200;
}

/**
 * Google Maps APIによる距離計算と各種交通費の自動更新
 */
function updateDistanceAndTravelCosts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffUserSheet = ss.getSheetByName(STAFF_USER_MASTER_SHEET_NAME);
  const staffSheet = ss.getSheetByName(STAFF_SHEET_NAME);
  const userSheet = ss.getSheetByName("利用者マスタ");

  if (!staffUserSheet || !staffSheet || !userSheet) {
    SpreadsheetApp.getUi().alert("必要なマスタシートがありません。");
    return;
  }

  const staffValues = staffSheet.getDataRange().getValues();
  const userValues = userSheet.getDataRange().getValues();
  const staffUserValues = staffUserSheet.getDataRange().getValues();
  const staffCols = getStaffMasterColumnMap_(staffSheet);

  const staffAddressMap = {};
  const userAddressMap = {};

  for (let i = 1; i < staffValues.length; i++) {
    const staffName = staffValues[i][staffCols.name];
    const address = staffValues[i][staffCols.address];

    if (staffName && address) {
      staffAddressMap[normalizeName_(staffName)] = address;
    }
  }

  for (let i = 1; i < userValues.length; i++) {
    const userName = userValues[i][0];
    const address = userValues[i][5]; // F列住所

    if (userName && address) {
      userAddressMap[normalizeName_(userName)] = address;
    }
  }

  let calculatedCount = 0;
  let skippedCount = 0;

  for (let i = 1; i < staffUserValues.length; i++) {
    const staffName = staffUserValues[i][0];
    const userName = staffUserValues[i][1];

    const existingDistance = Number(staffUserValues[i][5]) || 0;
    if (existingDistance > 0) {
      skippedCount++;
      continue;
    }

    const staffAddress = staffAddressMap[normalizeName_(staffName)];
    const userAddress = userAddressMap[normalizeName_(userName)];

    if (!staffAddress || !userAddress) continue;

    try {
      const directions = Maps.newDirectionFinder()
        .setOrigin(staffAddress)
        .setDestination(userAddress)
        .setMode(Maps.DirectionFinder.Mode.DRIVING)
        .getDirections();

      const route = directions.routes[0];
      if (!route) continue;

      const meters = route.legs[0].distance.value;
      const km = Math.round((meters / 1000) * 10) / 10;

      staffUserSheet.getRange(i + 1, 6).setValue(km); // F列：片道距離km
      calculatedCount++;

    } catch (e) {
      Logger.log(e);
    }
  }

  updateTravelCostsInStaffUserMaster();

  SpreadsheetApp.getUi().alert(
    "距離と交通費を更新しました。" + String.fromCharCode(10) +
    "新しく距離を計算した件数：" + calculatedCount + "件" + String.fromCharCode(10) +
    "既存距離のため上書きしなかった件数：" + skippedCount + "件"
  );
}

function updateTravelCostsInStaffUserMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STAFF_USER_MASTER_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert("スタッフ利用者マスタがありません。");
    return;
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    SpreadsheetApp.getUi().alert("スタッフ利用者マスタにデータがありません。");
    return;
  }

  sheet.getRange(1, 7).setValue("利用者請求交通費");
  sheet.getRange(1, 8).setValue("交通費備考");

  const output = [];

  for (let i = 1; i < values.length; i++) {
    const oneWayDistanceKm = Number(values[i][5]) || 0;
    const fallbackTravelCost = Number(values[i][4]) || 0;

    const payrollTravelCost = calculatePayrollTravelCost_(oneWayDistanceKm, fallbackTravelCost);
    const directTravelCost = calculateDirectTravelCostForUser_(oneWayDistanceKm);

    let memo = "";

    if (!oneWayDistanceKm) {
      memo = "距離未入力のためE列を使用";
    } else if (oneWayDistanceKm < 10) {
      memo = "給与交通費に反映";
    } else {
      memo = "給与交通費200円＋利用者へ直接請求";
    }

    output.push([
      payrollTravelCost,
      oneWayDistanceKm,
      directTravelCost,
      memo
    ]);
  }

  sheet.getRange(2, 5, output.length, 4).setValues(output);
}

/**
 * 給与明細PDFをテンプレートシートをベースに自動出力
 */
function createPayrollPdfFromTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const detailSheet = ss.getSheetByName("給与集計");
  const templateSheet = ss.getSheetByName("給与明細テンプレート");

  if (!detailSheet || !templateSheet) {
    SpreadsheetApp.getUi().alert("給与集計または給与明細テンプレートシートがありません。");
    return;
  }

  const values = detailSheet.getDataRange().getValues();
  const folder = DriveApp.getFolderById(PAYROLL_FOLDER_ID);
  const staffFolderMap = getStaffPayrollFolderMap_(ss);
  const targetYmKeys = getRecentPayrollYmKeys_(values, 2);
  const failedFiles = [];
  const skippedFiles = [];
  const noStaffFolderFiles = [];
  let oldMonthSkippedCount = 0;
  let createdCount = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const ymKey = formatPayrollYmKey_(row[0]);
    const ym = formatPayrollMonth_(row[0]);
    const staffName = row[2];
    const count = row[3];
    const basePay = row[4];
    const travelCost = row[5];
    const incentive = row[6];
    const referral = row[7];
    const management = row[8];
    const taxablePay = row[9];
    const tax = row[10];
    const payment = row[11];

    if (!ym || !staffName) continue;

    if (!targetYmKeys[ymKey]) {
      oldMonthSkippedCount++;
      continue;
    }

    const fileName = "給与明細_" + staffName + "_" + ym + ".pdf";
    const staffFolderId = staffFolderMap[normalizeName_(staffName)];

    if (staffFolderId) {
      try {
        const staffFolder = DriveApp.getFolderById(staffFolderId);

        if (hasPayrollPdfForMonth_(staffFolder, ym)) {
          skippedFiles.push(fileName);
          continue;
        }

      } catch (error) {
        failedFiles.push(fileName + "：スタッフフォルダ確認失敗 " + error.message);
        Logger.log(error);
        continue;
      }
    } else {
      noStaffFolderFiles.push(fileName);
    }

    templateSheet.getRange("D2").setValue(ym);
    templateSheet.getRange("A4").setValue("氏名：" + staffName);

    templateSheet.getRange("B7").setValue(count);
    templateSheet.getRange("B8").setValue(count + ":00");

    templateSheet.getRange("C7:C17").clearContent();

    templateSheet.getRange("E7").setValue(basePay);
    templateSheet.getRange("E8").setValue(taxablePay);
    templateSheet.getRange("E13").setValue(travelCost);

    templateSheet.getRange("E14").setValue(referral);
    templateSheet.getRange("E15").setValue(incentive);
    templateSheet.getRange("E16").setValue(management);

    templateSheet.getRange("F7:F16").clearContent();

    templateSheet.getRange("H7").setValue(taxablePay);
    templateSheet.getRange("H8").setValue(tax);
    templateSheet.getRange("H21").setValue(payment);

    templateSheet.getRange("B7:B17").setHorizontalAlignment("right").setFontSize(10);
    templateSheet.getRange("E7:E16").setHorizontalAlignment("right").setFontSize(10);
    templateSheet.getRange("H7:H8").setHorizontalAlignment("right").setFontSize(10);
    templateSheet.getRange("H21").setHorizontalAlignment("right").setFontSize(10);

    SpreadsheetApp.flush();

    try {
      const pdfBlob = exportSheetToPdfWithRetry_(ss, templateSheet);

      deleteExistingFileByName_(folder, fileName);

      folder
        .createFile(pdfBlob)
        .setName(fileName);

      createdCount++;
      Utilities.sleep(PDF_EXPORT_WAIT_MS);

    } catch (error) {
      failedFiles.push(fileName + "：" + error.message);
      Logger.log(error);
      Utilities.sleep(PDF_EXPORT_WAIT_MS * 2);
    }
  }

  if (failedFiles.length > 0) {
    SpreadsheetApp.getUi().alert(
      "給与明細PDFの作成が一部失敗しました。" + String.fromCharCode(10) +
      "作成済：" + createdCount + "件" + String.fromCharCode(10) +
      "既存明細ありでスキップ：" + skippedFiles.length + "件" + String.fromCharCode(10) +
      "直近2ヶ月以外でスキップ：" + oldMonthSkippedCount + "件" + String.fromCharCode(10) +
      "失敗：" + failedFiles.length + "件" + String.fromCharCode(10) +
      String.fromCharCode(10) +
      failedFiles.join(String.fromCharCode(10)) + String.fromCharCode(10) +
      String.fromCharCode(10) +
      "429エラーの場合は、数分待ってから再実行してください。"
    );
  } else {
    let message =
      "給与明細PDFの作成が完了しました。" + String.fromCharCode(10) +
      "作成済：" + createdCount + "件" + String.fromCharCode(10) +
      "既存明細ありでスキップ：" + skippedFiles.length + "件" + String.fromCharCode(10) +
      "直近2ヶ月以外でスキップ：" + oldMonthSkippedCount + "件";

    if (noStaffFolderFiles.length > 0) {
      message +=
        String.fromCharCode(10) + String.fromCharCode(10) +
        "スタッフフォルダID未登録のため、既存確認できず作成した明細：" + String.fromCharCode(10) +
        noStaffFolderFiles.join(String.fromCharCode(10));
    }

    SpreadsheetApp.getUi().alert(message);
  }
}

function getRecentPayrollYmKeys_(values, monthCount) {
  const months = [];
  const seen = {};

  for (let i = 1; i < values.length; i++) {
    const ymKey = formatPayrollYmKey_(values[i][0]);

    if (!ymKey || seen[ymKey]) continue;

    seen[ymKey] = true;
    months.push(ymKey);
  }

  months.sort().reverse();

  const map = {};
  months.slice(0, monthCount).forEach(ymKey => {
    map[ymKey] = true;
  });

  return map;
}

function formatPayrollYmKey_(value) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, "Asia/Tokyo", "yyyy-MM");
  }

  const text = String(value).trim();
  let match = text.match(/^(\d{4})[-\/](\d{1,2})/);

  if (match) {
    return match[1] + "-" + String(Number(match[2])).padStart(2, "0");
  }

  match = text.match(/^(\d{4})年\s*(\d{1,2})月/);

  if (match) {
    return match[1] + "-" + String(Number(match[2])).padStart(2, "0");
  }

  const date = new Date(text);
  if (!isNaN(date.getTime())) {
    return Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM");
  }

  return "";
}

function formatPayrollMonth_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(
      value,
      "Asia/Tokyo",
      "yyyy年M月分"
    );
  }

  const text = String(value);
  const match = text.match(/^(\d{4})-(\d{1,2})$/);

  if (match) {
    return match[1] + "年" + Number(match[2]) + "月分";
  }

  return text;
}

/**
 * シートを指定 of 印刷設定でPDF化してBlobとして取得
 */
function exportSheetToPdf_(ss, sheet) {
  const spreadsheetId = ss.getId();
  const sheetId = sheet.getSheetId();

  const url =
    "https://docs.google.com/spreadsheets/d/" +
    spreadsheetId +
    "/export?format=pdf" +
    "&gid=" + sheetId +
    "&size=A4" +
    "&portrait=false" +
    "&fitw=true" +
    "&scale=3" +
    "&top_margin=0.25" +
    "&bottom_margin=0.25" +
    "&left_margin=0.25" +
    "&right_margin=0.25" +
    "&sheetnames=false" +
    "&printtitle=false" +
    "&pagenumbers=false" +
    "&gridlines=false" +
    "&fzr=false";

  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + token
    }
  });

  return response.getBlob();
}

function exportSheetToPdfWithRetry_(ss, sheet) {
  let lastError = null;

  for (let attempt = 1; attempt <= PDF_EXPORT_MAX_RETRIES; attempt++) {
    try {
      if (attempt > 1) {
        Utilities.sleep(PDF_EXPORT_WAIT_MS * attempt);
      }

      return exportSheetToPdf_(ss, sheet);

    } catch (error) {
      lastError = error;

      const message = String(error && error.message ? error.message : error);
      const shouldRetry =
        message.indexOf("429") !== -1 ||
        message.indexOf("Service invoked too many times") !== -1 ||
        message.indexOf("Exception") !== -1 ||
        message.indexOf("リクエストに失敗") !== -1;

      if (!shouldRetry || attempt === PDF_EXPORT_MAX_RETRIES) {
        break;
      }
    }
  }

  throw lastError || new Error("PDF出力に失敗しました。");
}

function hasPayrollPdfForMonth_(folder, ymText) {
  const files = folder.getFilesByType(MimeType.PDF);
  const suffix = "_" + ymText + ".pdf";

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    if (
      fileName.indexOf("給与明細_") === 0 &&
      fileName.lastIndexOf(suffix) === fileName.length - suffix.length
    ) {
      return true;
    }
  }

  return false;
}

/**
 * GMOあおぞらネット銀行対応 総合振込CSVの自動作成
 *
 * - 給与集計のスタッフ名をキーにして銀行情報を取得
 * - 手数料負担区分は 1 = 当方負担
 * - 確認用シートにスタッフ名と警告を表示
 * - 実際のCSVファイルにはスタッフ名・警告列は含めない
 */
function createGmoTransferCsv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const payrollSheet = ss.getSheetByName("給与集計");
  const gmoSheet = ss.getSheetByName("GMO振込CSV");

  if (!payrollSheet) {
    SpreadsheetApp.getUi().alert("給与集計シートがありません。");
    return;
  }

  if (!gmoSheet) {
    SpreadsheetApp.getUi().alert("GMO振込CSVシートがありません。");
    return;
  }

  const staffBankMap = getStaffBankMap_(ss);
  const payrollValues = payrollSheet.getDataRange().getValues();

  if (payrollValues.length < 2) {
    SpreadsheetApp.getUi().alert("給与集計にデータがありません。");
    return;
  }

  const header = payrollValues[0];

  const ymCol = header.indexOf("年月");
  const staffNameCol = header.indexOf("スタッフ名");
  const paymentCol = header.indexOf("支給額");

  if (ymCol === -1 || staffNameCol === -1 || paymentCol === -1) {
    SpreadsheetApp.getUi().alert(
      "給与集計シートのヘッダーが見つかりません。\n\n" +
      "必要な列名：年月、スタッフ名、支給額"
    );
    return;
  }

  const REQUESTER_NAME = "ﾄﾞ)ﾋﾛｶﾞﾘ";
  const FEE_TYPE = 1; // 1:依頼人負担、2:受取人負担

  gmoSheet.clearContents();

  gmoSheet.appendRow([
    "年月",
    "スタッフ名",
    "銀行コード",
    "支店番号",
    "預金種目",
    "口座番号",
    "受取人名",
    "振込金額",
    "手数料負担区分",
    "振込依頼人名",
    "EDI情報",
    "確認メモ"
  ]);

  const csvRowsByYm = {};
  const unregisteredStaffs = [];
  const warningMessages = [];
  const targetYmKeys = getRecentPayrollYmKeys_(payrollValues, 2);
  let oldMonthSkippedRowCount = 0;

  for (let i = 1; i < payrollValues.length; i++) {
    const row = payrollValues[i];

    const ym = row[ymCol];
    const staffName = row[staffNameCol];
    const payment = row[paymentCol];

    if (!ym || !staffName || !payment) continue;

    const ymKey = formatPayrollYmKey_(ym);
    if (!targetYmKeys[ymKey]) {
      oldMonthSkippedRowCount++;
      continue;
    }

    const staffKey = normalizeName_(staffName);
    const bank = staffBankMap[staffKey];

    if (!bank) {
      unregisteredStaffs.push(staffName);
      continue;
    }

    const paymentNumber = Math.floor(Number(payment));

    if (!paymentNumber || paymentNumber <= 0) continue;

    const receiverName = normalizeGmoText_(bank.receiverName);
    const warning = makeGmoReceiverWarning_(staffName, receiverName);

    if (warning) {
      warningMessages.push(warning);
    }

    const outputRow = [
      bank.bankCode,
      bank.branchCode,
      bank.accountType,
      bank.accountNumber,
      receiverName,
      paymentNumber,
      FEE_TYPE,
      normalizeGmoText_(REQUESTER_NAME),
      ""
    ];

    gmoSheet.appendRow([
      ym,
      staffName
    ].concat(outputRow).concat([
      warning
    ]));

    const ymText = formatPayrollMonth_(ym);
    if (!csvRowsByYm[ymText]) csvRowsByYm[ymText] = [];
    csvRowsByYm[ymText].push(outputRow);
  }

  const folder = DriveApp.getFolderById(PAYROLL_FOLDER_ID);
  const confirmedFolder = DriveApp.getFolderById(GMO_TRANSFER_CSV_CONFIRMED_FOLDER_ID);
  const ymList = Object.keys(csvRowsByYm);
  let createdCsvCount = 0;
  let skippedCsvCount = 0;
  const skippedCsvFiles = [];

  if (ymList.length === 0) {
    SpreadsheetApp.getUi().alert("GMO振込CSV対象データがありません。");
    return;
  }

  ymList.forEach(ymText => {
    const fileName = "GMO振込CSV_" + ymText + ".csv";

    if (
      hasFileByName_(confirmedFolder, fileName) ||
      hasFileByName_(folder, fileName)
    ) {
      skippedCsvCount++;
      skippedCsvFiles.push(fileName);
      return;
    }

    const csvText = csvRowsByYm[ymText]
      .map(row => row.map(value => String(value)).join(","))
      .join(String.fromCharCode(13, 10)) + String.fromCharCode(13, 10);

    const blob = Utilities
      .newBlob("", "text/csv", fileName)
      .setDataFromString(csvText, "Shift_JIS");

    folder.createFile(blob);
    createdCsvCount++;
  });

  let message =
    "GMO振込CSV処理が完了しました。" + String.fromCharCode(10) +
    "作成：" + createdCsvCount + "件" + String.fromCharCode(10) +
    "既存CSVありでスキップ：" + skippedCsvCount + "件" + String.fromCharCode(10) +
    "直近2ヶ月以外でスキップ：" + oldMonthSkippedRowCount + "行";

  if (skippedCsvFiles.length > 0) {
    message +=
      "\n\n作成しなかったCSV：\n" +
      skippedCsvFiles.join("\n");
  }

  if (unregisteredStaffs.length > 0) {
    message +=
      "\n\n⚠️ 口座情報未登録：\n" +
      unregisteredStaffs.join("、");
  }

  if (warningMessages.length > 0) {
    message +=
      "\n\n⚠️ スタッフ名と受取人名が一致していない可能性があります。\n" +
      "GMO振込CSVシートの確認メモ列を確認してください。";
  }

  SpreadsheetApp.getUi().alert(message);
}

/**
 * スタッフ名と受取人名の明らかなズレを検出する
 * 完全判定ではなく、取り違え防止用の簡易チェック
 */
function makeGmoReceiverWarning_(staffName, receiverName) {
  const staff = normalizeName_(staffName);
  const receiver = String(receiverName || "");

  const checks = [
    {
      staffContains: "高橋",
      receiverContains: "ﾀｶﾊｼ"
    },
    {
      staffContains: "小澤",
      receiverContains: "ｺｻﾞﾜ"
    },
    {
      staffContains: "石川",
      receiverContains: "ｲｼｶﾜ"
    },
    {
      staffContains: "石倉",
      receiverContains: "ｲｼｸﾗ"
    }
  ];

  for (let i = 0; i < checks.length; i++) {
    const item = checks[i];

    if (
      staff.indexOf(item.staffContains) !== -1 &&
      receiver.indexOf(item.receiverContains) === -1
    ) {
      return "要確認：スタッフ名と受取人名が一致していない可能性があります";
    }
  }

  return "";
}

/**
 * GMOあおぞら銀行CSV用の文字列整形
 * - 全角スペースを半角スペースへ変換
 * - カンマを除去
 * - 全角カッコを半角カッコへ変換
 * - 前後の空白を削除
 */
function normalizeGmoText_(value) {
  return String(value || "")
    .replace(/　/g, " ")
    .replace(/[，、]/g, "")
    .replace(/[（]/g, "(")
    .replace(/[）]/g, ")")
    .trim();
}

/**
 * 日付オブジェクトまたは文字列を YYYYMMDD 形式に変換するヘルパー関数
 */
function formatToYyyymmdd_(dateValue) {
  if (!dateValue) return "";
  let date;
  if (Object.prototype.toString.call(dateValue) === "[object Date]") {
    date = dateValue;
  } else {
    const cleanStr = String(dateValue).replace(/-/g, "/").trim();
    date = new Date(cleanStr);
  }

  if (isNaN(date.getTime())) return "";

  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return yyyy + mm + dd;
}

function deleteExistingFileByName_(folder, fileName) {
  const files = folder.getFilesByName(fileName);

  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
}

function hasFileByName_(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  return files.hasNext();
}

/**
 * スタッフマスタから、スタッフ名をキーにして銀行情報を取得する
 *
 * スタッフマスタ固定列：
 * A列：スタッフ名
 * B列：LINE表示名
 * C列：LINEユーザーID
 * D列：LIFF用LINEユーザーID
 * E列：カレンダーID
 * F列：銀行コード
 * G列：支店番号
 * H列：預金種目
 * I列：口座番号
 * J列：受取人名
 * K列：給与明細フォルダID
 * L列：住所
 */
function getStaffBankMap_(ss) {
  const sheet = ss.getSheetByName(STAFF_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert("スタッフマスタシートがありません。");
    return {};
  }

  const values = sheet.getDataRange().getValues();
  const cols = getStaffMasterColumnMap_(sheet);
  const map = {};
  const duplicateStaffs = [];

  for (let i = 1; i < values.length; i++) {
    const staffName = values[i][cols.name];

    if (!staffName) continue;

    const staffKey = normalizeName_(staffName);

    if (map[staffKey]) {
      duplicateStaffs.push(staffName);
      continue;
    }

    map[staffKey] = {
      bankCode: formatCode_(values[i][cols.bankCode], 4),
      branchCode: formatCode_(values[i][cols.branchCode], 3),
      accountType: values[i][cols.accountType],
      accountNumber: formatCode_(values[i][cols.accountNumber], 7),
      receiverName: values[i][cols.receiverName]
    };
  }

  if (duplicateStaffs.length > 0) {
    SpreadsheetApp.getUi().alert(
      "スタッフマスタに同じスタッフ名が複数あります。\n\n" +
      duplicateStaffs.join("、") + "\n\n" +
      "口座情報の取り違え防止のため、スタッフマスタを確認してください。"
    );
  }

  return map;
}

function formatCode_(value, length) {
  return String(value || "").padStart(length, "0");
}

/**
 * 給与関連一括処理の実行
 */
function runPayrollAll() {
  createPayrollSummary();
  createWageLedger();
  createPayrollPdfFromTemplate();
  createGmoTransferCsv();
}

/**
 * 賃金台帳の自動更新・蓄積
 */
function createWageLedger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const payrollSheet = ss.getSheetByName("給与集計");
  let ledgerSheet = ss.getSheetByName("賃金台帳");

  if (!payrollSheet) {
    SpreadsheetApp.getUi().alert("給与集計シートがありません。");
    return;
  }

  if (!ledgerSheet) {
    SpreadsheetApp.getUi().alert("賃金台帳シートがありません。先に手動で作成してください。");
    return;
  }

  const header1 = [
    "支払月日", "氏名", "性別", "労働日数", "労働時間数", "早出残業時間数",
    "深夜労働時間数", "基本賃金", "賃金所定時間外割増", "通勤手当",
    "奨励手当", "紹介手当", "運営報酬", "合計", "控除額", "実物給与", "備考"
  ];

  const header2 = [
    "", "", "", "", "", "", "", "", "", "手当/報酬", "手当/報酬",
    "手当/報酬", "手当/報酬", "", "", "", ""
  ];

  if (ledgerSheet.getLastRow() < 2) {
    ledgerSheet.getRange(1, 1, 1, header1.length).setValues([header1]);
    ledgerSheet.getRange(2, 1, 1, header2.length).setValues([header2]);
    ledgerSheet.setFrozenRows(2);
  } else {
    ledgerSheet.getRange(1, 1, 1, header1.length).setValues([header1]);
    ledgerSheet.getRange(2, 1, 1, header2.length).setValues([header2]);
  }

  const payrollValues = payrollSheet.getDataRange().getValues();
  if (payrollValues.length < 2) return;

  const existingValues = ledgerSheet.getDataRange().getValues();
  const existingRowMap = {};

  for (let i = 2; i < existingValues.length; i++) {
    const payDate = formatDateForKey_(existingValues[i][0]);
    const staffName = normalizeName_(existingValues[i][1]);
    if (!payDate || !staffName) continue;
    existingRowMap[payDate + "_" + staffName] = i + 1;
  }

  for (let i = 1; i < payrollValues.length; i++) {
    const row = payrollValues[i];

    const payDate = row[1];
    const staffName = row[2];
    const visitCount = Number(row[3]) || 0;
    const basePay = Number(row[4]) || 0;
    const travelCost = Number(row[5]) || 0;
    const incentive = Number(row[6]) || 0;
    const referral = Number(row[7]) || 0;
    const management = Number(row[8]) || 0;
    const tax = Number(row[10]) || 0;
    const payment = Number(row[11]) || 0;
    const memo = row[12] || "";

    if (!payDate || !staffName) continue;

    const total = basePay + travelCost + incentive + referral + management;
    const outputRow = [
      payDate,
      staffName,
      "",
      visitCount,
      visitCount,
      0,
      0,
      basePay,
      0,
      travelCost,
      incentive,
      referral,
      management,
      total,
      tax,
      payment,
      memo
    ];

    const key = formatDateForKey_(payDate) + "_" + normalizeName_(staffName);
    const existingRow = existingRowMap[key];

    if (existingRow) {
      ledgerSheet.getRange(existingRow, 1, 1, outputRow.length).setValues([outputRow]);
    } else {
      ledgerSheet.appendRow(outputRow);
    }
  }
}

function formatDateForKey_(value) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, "Asia/Tokyo", "yyyy/MM/dd");
  }

  return String(value).replace(/-/g, "/").trim();
}

/**
 * 共有フォルダに作成されたPDFを、スタッフ個別Googleドライブフォルダへ自動移動
 */
function movePayrollPdfsToStaffFolders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffFolderMap = getStaffPayrollFolderMap_(ss);
  const sourceFolder = DriveApp.getFolderById(PAYROLL_FOLDER_ID);
  const files = sourceFolder.getFilesByType(MimeType.PDF);

  let movedCount = 0;
  let skippedCount = 0;
  const skippedFiles = [];

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    if (fileName.indexOf("給与明細_") !== 0) continue;

    const staffName = extractStaffNameFromPayrollPdfName_(fileName);
    const staffKey = normalizeName_(staffName);
    const targetFolderId = staffFolderMap[staffKey];

    if (!staffName || !targetFolderId) {
      skippedCount++;
      skippedFiles.push(fileName + "（抽出名：" + staffName + "）");
      continue;
    }

    const targetFolder = DriveApp.getFolderById(targetFolderId);
    file.moveTo(targetFolder);
    movedCount++;
  }

  let message = "給与明細PDFの移動が完了しました。" + String.fromCharCode(10) +
    "移動件数：" + movedCount + "件" + String.fromCharCode(10) +
    "未移動件数：" + skippedCount + "件";

  if (skippedFiles.length > 0) {
    message += String.fromCharCode(10) + String.fromCharCode(10) +
      "未移動ファイル：" + String.fromCharCode(10) +
      skippedFiles.join(String.fromCharCode(10));
  }

  SpreadsheetApp.getUi().alert(message);
}

function getStaffPayrollFolderMap_(ss) {
  const sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const cols = getStaffMasterColumnMap_(sheet);
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const staffName = values[i][cols.name];
    const folderId = values[i][cols.payrollFolderId];

    if (!staffName || !folderId) continue;

    const staffKey = normalizeName_(staffName);
    map[staffKey] = folderId;
  }

  return map;
}

function extractStaffNameFromPayrollPdfName_(fileName) {
  const text = String(fileName || "");
  const prefix = "給与明細_";
  const suffixIndex = text.lastIndexOf("_");

  if (text.indexOf(prefix) !== 0 || suffixIndex <= prefix.length) return "";

  return text.substring(prefix.length, suffixIndex);
}

/**
 * 登録スタッフの紐付け有無確認
 */
function isRegisteredStaffUser_(ss, staffName, userName) {
  if (isTestUserName_(userName)) return true;

  const sheet = ss.getSheetByName(STAFF_USER_MASTER_SHEET_NAME);
  if (!sheet) return false;

  const values = sheet.getDataRange().getValues();
  const targetStaff = normalizeName_(staffName);
  const targetUser = normalizeName_(userName);

  for (let i = 1; i < values.length; i++) {
    const masterStaff = normalizeName_(values[i][0]);
    const formalName = normalizeName_(values[i][1]);
    const lineName = normalizeName_(values[i][2]);

    if (masterStaff !== targetStaff) continue;

    if (formalName === targetUser || lineName === targetUser) {
      return true;
    }
  }

  return false;
}

function isTestUserName_(userName) {
  const normalized = normalizeName_(userName);
  return normalized === normalizeName_(TEST_USER_NAME) || normalized === "テスト";
}

function saveUnknownUser_(ss, receivedAt, staffName, userName, text) {
  const sheet = ss.getSheetByName("未登録利用者");

  if (!sheet) return;

  sheet.appendRow([
    receivedAt,
    staffName,
    userName || "不明",
    text,
    "未登録"
  ]);
}

function saveLineMessageLog_(ss, dateTime, direction, senderName, lineUserId, message) {
  const sheet = ss.getSheetByName(LINE_MESSAGE_LOG_SHEET_NAME);

  if (!sheet) return;

  sheet.appendRow([
    dateTime,
    direction,
    senderName,
    lineUserId,
    message
  ]);
}

function saveLineUserDirectory_(ss, data) {
  const sheet = ensureLineUserDirectorySheet_(ss);
  const now = data.checkedAt || new Date();
  const messagingLineUserId = String(data.messagingLineUserId || "").trim();
  const liffLineUserId = String(data.liffLineUserId || "").trim();
  const displayName = String(data.displayName || "").trim();
  const detectedStaffName = String(data.detectedStaffName || "").trim();
  const source = String(data.source || "").trim();
  const lastMessage = String(data.lastMessage || "").trim();

  if (!messagingLineUserId && !liffLineUserId && !displayName) return;

  const rowNumber = findLineUserDirectoryRow_(sheet, messagingLineUserId, liffLineUserId, displayName);
  const existing = rowNumber ? sheet.getRange(rowNumber, 1, 1, 13).getValues()[0] : [];
  const match = resolveLineUserDirectoryMatch_(ss, displayName, detectedStaffName);
  const firstSeenAt = existing[0] || now;
  const mergedMessagingLineUserId = messagingLineUserId || existing[2] || "";
  const mergedLiffLineUserId = liffLineUserId || existing[3] || match.liffLineUserId || "";
  const mergedDisplayName = displayName || existing[4] || "";
  const mergedSource = mergeDirectoryText_(existing[10], source);
  const mergedMessage = lastMessage || existing[11] || "";
  const note = match.note || existing[12] || "";

  const values = [[
    firstSeenAt,
    now,
    mergedMessagingLineUserId,
    mergedLiffLineUserId,
    mergedDisplayName,
    match.type,
    match.name,
    match.staffRow || "",
    match.userRow || "",
    match.status,
    mergedSource,
    mergedMessage,
    note
  ]];

  if (rowNumber) {
    sheet.getRange(rowNumber, 1, 1, values[0].length).setValues(values);
  } else {
    sheet.appendRow(values[0]);
  }
}

function ensureLineUserDirectorySheet_(ss) {
  const headers = [
    "初回確認日時",
    "最終確認日時",
    "Messaging API LINEユーザーID",
    "LIFF用LINEユーザーID",
    "LINE表示名",
    "種別候補",
    "マスタ候補名",
    "スタッフマスタ行",
    "利用者マスタ行",
    "紐づけ状態",
    "取得元",
    "最終メッセージ",
    "メモ"
  ];
  let sheet = ss.getSheetByName(LINE_USER_DIRECTORY_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(LINE_USER_DIRECTORY_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeaders = headers.some((header, index) => String(existingHeaders[index] || "") !== header);
  if (needsHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function findLineUserDirectoryRow_(sheet, messagingLineUserId, liffLineUserId, displayName) {
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  const targetMessagingId = String(messagingLineUserId || "").trim();
  const targetLiffId = String(liffLineUserId || "").trim();
  const targetDisplayName = normalizeName_(displayName);

  for (let i = 0; i < values.length; i++) {
    if (targetMessagingId && String(values[i][2] || "").trim() === targetMessagingId) return i + 2;
    if (targetLiffId && String(values[i][3] || "").trim() === targetLiffId) return i + 2;
  }

  if (!targetDisplayName) return 0;

  for (let i = 0; i < values.length; i++) {
    if (normalizeName_(values[i][4]) === targetDisplayName) return i + 2;
  }

  return 0;
}

function resolveLineUserDirectoryMatch_(ss, displayName, detectedStaffName) {
  const staffMatch = findStaffDirectoryMatch_(ss, displayName, detectedStaffName);
  const userMatch = findUserDirectoryMatch_(ss, displayName);

  if (staffMatch && userMatch) {
    return {
      type: "要確認",
      name: staffMatch.name + " / " + userMatch.name,
      staffRow: staffMatch.rowNumber,
      userRow: userMatch.rowNumber,
      liffLineUserId: staffMatch.liffLineUserId || "",
      status: "スタッフ・利用者の両方に一致",
      note: "同じLINE表示名がスタッフと利用者の両方にあります。手動確認してください。"
    };
  }

  if (staffMatch) {
    return {
      type: "スタッフ",
      name: staffMatch.name,
      staffRow: staffMatch.rowNumber,
      userRow: "",
      liffLineUserId: staffMatch.liffLineUserId || "",
      status: staffMatch.liffLineUserId ? "スタッフ候補・LIFF紐づけ済み" : "スタッフ候補",
      note: ""
    };
  }

  if (userMatch) {
    return {
      type: "利用者",
      name: userMatch.name,
      staffRow: "",
      userRow: userMatch.rowNumber,
      liffLineUserId: "",
      status: "利用者候補",
      note: ""
    };
  }

  return {
    type: "",
    name: "",
    staffRow: "",
    userRow: "",
    liffLineUserId: "",
    status: "未紐づけ",
    note: ""
  };
}

function findStaffDirectoryMatch_(ss, displayName, detectedStaffName) {
  const sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  if (!sheet) return null;

  const targetDisplayName = normalizeName_(displayName);
  const targetStaffName = normalizeName_(detectedStaffName);
  const values = sheet.getDataRange().getValues();
  const cols = getStaffMasterColumnMap_(sheet);
  const matches = [];

  for (let i = 1; i < values.length; i++) {
    const staffName = values[i][cols.name];
    const lineDisplayName = values[i][cols.lineDisplayName];

    if (!staffName) continue;

    if (
      (targetStaffName && normalizeName_(staffName) === targetStaffName) ||
      (targetDisplayName && (
        normalizeName_(staffName) === targetDisplayName ||
        normalizeName_(lineDisplayName) === targetDisplayName
      ))
    ) {
      matches.push({
        name: staffName,
        rowNumber: i + 1,
        liffLineUserId: values[i][cols.liffLineUserId] || ""
      });
    }
  }

  const unique = dedupeDirectoryMatches_(matches);
  return unique.length === 1 ? unique[0] : null;
}

function findUserDirectoryMatch_(ss, displayName) {
  const target = normalizeName_(displayName);
  if (!target) return null;

  const matches = [];
  collectUserDirectoryMatchesFromUserMaster_(ss, target, matches);
  collectUserDirectoryMatchesFromStaffUserMaster_(ss, target, matches);

  const unique = dedupeDirectoryMatches_(matches);
  return unique.length === 1 ? unique[0] : null;
}

function collectUserDirectoryMatchesFromUserMaster_(ss, target, matches) {
  const sheet = ss.getSheetByName("利用者マスタ");
  if (!sheet || sheet.getLastRow() < 2) return;

  const values = sheet.getDataRange().getValues();
  const headerMap = getHeaderColumnMap_(sheet);
  const nameCol = getColumnIndex_(headerMap, ["利用者名", "氏名", "名前"], 0);
  const lineNameCol = getColumnIndex_(headerMap, ["LINE表示名", "LINE名"], -1);

  for (let i = 1; i < values.length; i++) {
    const userName = values[i][nameCol];
    const lineDisplayName = lineNameCol >= 0 ? values[i][lineNameCol] : "";

    if (!userName) continue;

    if (
      normalizeName_(userName) === target ||
      normalizeName_(lineDisplayName) === target
    ) {
      matches.push({
        name: userName,
        rowNumber: i + 1
      });
    }
  }
}

function collectUserDirectoryMatchesFromStaffUserMaster_(ss, target, matches) {
  const sheet = ss.getSheetByName(STAFF_USER_MASTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return;

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const userName = values[i][1];
    const lineDisplayName = values[i][2];

    if (!userName) continue;

    if (
      normalizeName_(userName) === target ||
      normalizeName_(lineDisplayName) === target
    ) {
      matches.push({
        name: userName,
        rowNumber: ""
      });
    }
  }
}

function dedupeDirectoryMatches_(matches) {
  const seen = {};
  const unique = [];

  matches.forEach(match => {
    const key = normalizeName_(match.name);
    if (!key || seen[key]) return;
    seen[key] = true;
    unique.push(match);
  });

  return unique;
}

function mergeDirectoryText_(currentValue, nextValue) {
  const values = String(currentValue || "")
    .split("、")
    .map(value => value.trim())
    .filter(value => value);
  const next = String(nextValue || "").trim();

  if (next && values.indexOf(next) === -1) {
    values.push(next);
  }

  return values.join("、");
}

function updateLineUserDirectoryLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureLineUserDirectorySheet_(ss);

  if (sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("LINEユーザー一覧にデータがありません。");
    return;
  }

  const updatedCount = updateLineUserDirectoryLinksWithoutAlert_(ss, sheet);
  SpreadsheetApp.getUi().alert("LINEユーザー一覧の紐づけ候補を更新しました。件数：" + updatedCount + "件");
}

function applyLineUserDirectoryToMasters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const directorySheet = ensureLineUserDirectorySheet_(ss);

  if (directorySheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("LINEユーザー一覧にデータがありません。");
    return;
  }

  updateLineUserDirectoryLinksWithoutAlert_(ss, directorySheet);

  const rows = directorySheet.getRange(2, 1, directorySheet.getLastRow() - 1, 13).getValues();
  const staffSheet = ss.getSheetByName(STAFF_SHEET_NAME);
  const userSheet = ss.getSheetByName("利用者マスタ");
  const staffCols = staffSheet ? getStaffMasterColumnMap_(staffSheet) : null;
  const userCols = userSheet ? ensureUserMasterLineColumns_(userSheet) : null;
  let staffUpdatedCount = 0;
  let userUpdatedCount = 0;
  let skippedCount = 0;

  rows.forEach((row, index) => {
    const rowNumber = index + 2;
    const messagingLineUserId = String(row[2] || "").trim();
    const liffLineUserId = String(row[3] || "").trim();
    const displayName = String(row[4] || "").trim();
    const type = String(row[5] || "").trim();
    const staffRow = Number(row[7]) || 0;
    const userRow = Number(row[8]) || 0;

    if (type === "スタッフ" && staffSheet && staffCols && staffRow >= 2) {
      const result = applyLineUserDirectoryToStaffRow_(
        staffSheet,
        staffCols,
        staffRow,
        displayName,
        messagingLineUserId,
        liffLineUserId
      );
      staffUpdatedCount += result.updated ? 1 : 0;
      skippedCount += result.skipped ? 1 : 0;
      directorySheet.getRange(rowNumber, 10).setValue(result.status);
      directorySheet.getRange(rowNumber, 13).setValue(result.note);
      return;
    }

    if (type === "利用者" && userSheet && userCols && userRow >= 2) {
      const result = applyLineUserDirectoryToUserRow_(
        userSheet,
        userCols,
        userRow,
        displayName,
        messagingLineUserId,
        liffLineUserId
      );
      userUpdatedCount += result.updated ? 1 : 0;
      skippedCount += result.skipped ? 1 : 0;
      directorySheet.getRange(rowNumber, 10).setValue(result.status);
      directorySheet.getRange(rowNumber, 13).setValue(result.note);
      return;
    }

    if (type) {
      skippedCount++;
      directorySheet.getRange(rowNumber, 10).setValue("反映対象外");
      directorySheet.getRange(rowNumber, 13).setValue("候補が複数、またはマスタ行が特定できないため手動確認してください。");
    }
  });

  SpreadsheetApp.getUi().alert(
    "LINEユーザー一覧からマスタへ反映しました。\n" +
    "スタッフ更新：" + staffUpdatedCount + "件\n" +
    "利用者更新：" + userUpdatedCount + "件\n" +
    "確認・スキップ：" + skippedCount + "件"
  );
}

function updateLineUserDirectoryLinksWithoutAlert_(ss, sheet) {
  const targetSheet = sheet || ensureLineUserDirectorySheet_(ss);
  if (targetSheet.getLastRow() < 2) return 0;

  const values = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 13).getValues();
  const updatedRows = values.map(row => {
    const displayName = row[4];
    const match = resolveLineUserDirectoryMatch_(ss, displayName, "");

    row[3] = row[3] || match.liffLineUserId || "";
    row[5] = match.type;
    row[6] = match.name;
    row[7] = match.staffRow || "";
    row[8] = match.userRow || "";
    row[9] = match.status;
    row[12] = match.note || row[12] || "";

    return row;
  });

  targetSheet.getRange(2, 1, updatedRows.length, 13).setValues(updatedRows);
  return updatedRows.length;
}

function applyLineUserDirectoryToStaffRow_(sheet, cols, rowNumber, displayName, messagingLineUserId, liffLineUserId) {
  const row = sheet.getRange(rowNumber, 1, 1, Math.max(sheet.getLastColumn(), 12)).getValues()[0];
  const updates = [];
  const conflicts = [];

  addCellUpdateIfEmpty_(updates, conflicts, row, rowNumber, cols.lineDisplayName, displayName, "LINE表示名");
  addCellUpdateIfEmpty_(updates, conflicts, row, rowNumber, cols.lineUserId, messagingLineUserId, "LINEユーザーID");
  addCellUpdateIfEmpty_(updates, conflicts, row, rowNumber, cols.liffLineUserId, liffLineUserId, "LIFF用LINEユーザーID");
  updates.forEach(update => sheet.getRange(update.rowNumber, update.column).setValue(update.value));

  return {
    updated: updates.length > 0,
    skipped: updates.length === 0 || conflicts.length > 0,
    status: conflicts.length > 0 ? "スタッフマスタ一部確認" : "スタッフマスタ反映済み",
    note: conflicts.join(" / ")
  };
}

function applyLineUserDirectoryToUserRow_(sheet, cols, rowNumber, displayName, messagingLineUserId, liffLineUserId) {
  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  const updates = [];
  const conflicts = [];

  addCellUpdateIfEmpty_(updates, conflicts, row, rowNumber, cols.lineDisplayName, displayName, "LINE表示名");
  addCellUpdateIfEmpty_(updates, conflicts, row, rowNumber, cols.lineUserId, messagingLineUserId, "LINEユーザーID");
  addCellUpdateIfEmpty_(updates, conflicts, row, rowNumber, cols.liffLineUserId, liffLineUserId, "LIFF用LINEユーザーID");
  updates.forEach(update => sheet.getRange(update.rowNumber, update.column).setValue(update.value));

  return {
    updated: updates.length > 0,
    skipped: updates.length === 0 || conflicts.length > 0,
    status: conflicts.length > 0 ? "利用者マスタ一部確認" : "利用者マスタ反映済み",
    note: conflicts.join(" / ")
  };
}

function addCellUpdateIfEmpty_(updates, conflicts, row, rowNumber, zeroBasedColumn, value, label) {
  const nextValue = String(value || "").trim();
  if (!nextValue || zeroBasedColumn < 0) return;

  const currentValue = String(row[zeroBasedColumn] || "").trim();
  if (!currentValue) {
    updates.push({
      rowNumber: rowNumber,
      column: zeroBasedColumn + 1,
      value: nextValue
    });
    return;
  }

  if (currentValue !== nextValue) {
    conflicts.push(label + "は既存値があるため未上書き");
  }
}

function ensureUserMasterLineColumns_(sheet) {
  const lineDisplayName = ensureHeaderColumn_(sheet, "LINE表示名");
  const lineUserId = ensureHeaderColumn_(sheet, "LINEユーザーID");
  const liffLineUserId = ensureHeaderColumn_(sheet, "LIFF用LINEユーザーID");

  return {
    lineDisplayName: lineDisplayName,
    lineUserId: lineUserId,
    liffLineUserId: liffLineUserId
  };
}

function ensureHeaderColumn_(sheet, headerName) {
  const headerMap = getHeaderColumnMap_(sheet);
  const normalizedHeaderName = normalizeHeaderName_(headerName);

  if (headerMap[headerName] !== undefined) return headerMap[headerName];
  if (headerMap[normalizedHeaderName] !== undefined) return headerMap[normalizedHeaderName];

  const column = sheet.getLastColumn() + 1;
  sheet.getRange(1, column).setValue(headerName);
  return column - 1;
}

function getMessagingLineUserIdByLiffId_(ss, lineUserId) {
  if (!lineUserId) return "";

  const sheet = ss.getSheetByName(STAFF_SHEET_NAME);
  if (!sheet) return "";

  const values = sheet.getDataRange().getValues();
  const cols = getStaffMasterColumnMap_(sheet);
  const targetUserId = String(lineUserId || "").trim();

  for (let i = 1; i < values.length; i++) {
    const messagingLineUserId = String(values[i][cols.lineUserId] || "").trim();
    const liffLineUserId = String(values[i][cols.liffLineUserId] || "").trim();

    if (liffLineUserId === targetUserId || messagingLineUserId === targetUserId) {
      return messagingLineUserId;
    }
  }

  return "";
}

function sendVisitConfirmationPushFromLiff_(ss, liffLineUserId, staffName, userName, visitType) {
  const typeText = visitType === "終了" ? "終了" : "開始";

  try {
    const token = PropertiesService
      .getScriptProperties()
      .getProperty("LINE_CHANNEL_ACCESS_TOKEN");
    const toLineUserId = getMessagingLineUserIdByLiffId_(ss, liffLineUserId);
    const message = "利用者：" + userName + " 様の" + typeText + "の実績を登録しました。";

    if (!token) {
      saveLiffOperationLog_(
        ss,
        "pushVisitConfirmation",
        liffLineUserId,
        staffName,
        userName,
        typeText,
        "",
        "",
        "失敗",
        "LINE_CHANNEL_ACCESS_TOKENが未設定のため、確認LINEを送信できませんでした。"
      );
      return;
    }

    if (!toLineUserId) {
      saveLiffOperationLog_(
        ss,
        "pushVisitConfirmation",
        liffLineUserId,
        staffName,
        userName,
        typeText,
        "",
        "",
        "失敗",
        "スタッフマスタC列のLINEユーザーIDが未登録のため、確認LINEを送信できませんでした。"
      );
      return;
    }

    const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + token
      },
      payload: JSON.stringify({
        to: toLineUserId,
        messages: [
          {
            type: "text",
            text: message
          }
        ]
      }),
      muteHttpExceptions: true
    });
    const responseCode = response.getResponseCode();
    const success = responseCode >= 200 && responseCode < 300;

    saveLiffOperationLog_(
      ss,
      "pushVisitConfirmation",
      liffLineUserId,
      staffName,
      userName,
      typeText,
      "",
      "",
      success ? "成功" : "失敗",
      success
        ? "確認LINEを送信しました。送信先LINEユーザーID：" + toLineUserId
        : "確認LINE送信失敗：" + responseCode + " / 送信先LINEユーザーID：" + toLineUserId + " / " + response.getContentText()
    );

    saveLineMessageLog_(
      ss,
      new Date(),
      success ? "送信" : "送信失敗",
      staffName || "",
      toLineUserId,
      success ? message : "LINE Push送信失敗：" + responseCode + " " + response.getContentText()
    );
  } catch (error) {
    saveLiffOperationLog_(
      ss,
      "pushVisitConfirmation",
      liffLineUserId,
      staffName,
      userName,
      typeText,
      "",
      "",
      "失敗",
      error && error.stack ? error.stack : String(error)
    );
  }
}

function sendScheduleConfirmationPushFromLiff_(ss, liffLineUserId, staffName, userName, dates) {
  const dateText = (dates || []).join("、");

  try {
    const token = PropertiesService
      .getScriptProperties()
      .getProperty("LINE_CHANNEL_ACCESS_TOKEN");
    const toLineUserId = getMessagingLineUserIdByLiffId_(ss, liffLineUserId);
    const message =
      "予定を登録しました\n" +
      "利用者：" + userName + "\n" +
      "日付：" + dateText;

    if (!token) {
      saveLiffOperationLog_(
        ss,
        "pushScheduleConfirmation",
        liffLineUserId,
        staffName,
        userName,
        "",
        dateText,
        "",
        "失敗",
        "LINE_CHANNEL_ACCESS_TOKENが未設定のため、予定確認LINEを送信できませんでした。"
      );
      return;
    }

    if (!toLineUserId) {
      saveLiffOperationLog_(
        ss,
        "pushScheduleConfirmation",
        liffLineUserId,
        staffName,
        userName,
        "",
        dateText,
        "",
        "失敗",
        "スタッフマスタC列のLINEユーザーIDが未登録のため、予定確認LINEを送信できませんでした。"
      );
      return;
    }

    const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + token
      },
      payload: JSON.stringify({
        to: toLineUserId,
        messages: [
          {
            type: "text",
            text: message
          }
        ]
      }),
      muteHttpExceptions: true
    });
    const responseCode = response.getResponseCode();
    const success = responseCode >= 200 && responseCode < 300;

    saveLiffOperationLog_(
      ss,
      "pushScheduleConfirmation",
      liffLineUserId,
      staffName,
      userName,
      "",
      dateText,
      "",
      success ? "成功" : "失敗",
      success
        ? "予定確認LINEを送信しました。送信先LINEユーザーID：" + toLineUserId
        : "予定確認LINE送信失敗：" + responseCode + " / 送信先LINEユーザーID：" + toLineUserId + " / " + response.getContentText()
    );

    saveLineMessageLog_(
      ss,
      new Date(),
      success ? "送信" : "送信失敗",
      staffName || "",
      toLineUserId,
      success ? message : "予定確認LINE送信失敗：" + responseCode + " " + response.getContentText()
    );
  } catch (error) {
    saveLiffOperationLog_(
      ss,
      "pushScheduleConfirmation",
      liffLineUserId,
      staffName,
      userName,
      "",
      dateText,
      "",
      "失敗",
      error && error.stack ? error.stack : String(error)
    );
  }
}

function testPushFromLiff_(ss, params) {
  const lineUserId = params.lineUserId || "";
  const token = PropertiesService
    .getScriptProperties()
    .getProperty("LINE_CHANNEL_ACCESS_TOKEN");

  if (!lineUserId) {
    return {
      success: false,
      message: "lineUserIdが指定されていません。"
    };
  }

  if (!token) {
    return {
      success: false,
      message: "LINE_CHANNEL_ACCESS_TOKENが未設定です。"
    };
  }

  const message = "LINE返信テストです。連携できています。";
  const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + token
    },
    payload: JSON.stringify({
      to: lineUserId,
      messages: [
        {
          type: "text",
          text: message
        }
      ]
    }),
    muteHttpExceptions: true
  });
  const responseCode = response.getResponseCode();
  const success = responseCode >= 200 && responseCode < 300;

  saveLiffOperationLog_(
    ss,
    "testPush",
    lineUserId,
    "",
    "",
    "",
    "",
    "",
    success ? "成功" : "失敗",
    success
      ? "テストLINEを送信しました。"
      : "テストLINE送信失敗：" + responseCode + " " + response.getContentText()
  );

  return {
    success: success,
    statusCode: responseCode,
    responseText: response.getContentText(),
    message: success
      ? "テストLINEを送信しました。"
      : "テストLINE送信に失敗しました。"
  };
}

function sendPushMessage_(ss, lineUserId, staffName, message) {
  const token = PropertiesService
    .getScriptProperties()
    .getProperty("LINE_CHANNEL_ACCESS_TOKEN");

  if (!lineUserId || !message) return;

  if (!token) {
    saveLineMessageLog_(
      ss,
      new Date(),
      "送信失敗",
      staffName || "",
      lineUserId || "",
      "LINE_CHANNEL_ACCESS_TOKENが未設定のため、確認LINEを送信できませんでした。"
    );
    return;
  }

  const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + token
    },
    payload: JSON.stringify({
      to: lineUserId,
      messages: [
        {
          type: "text",
          text: message
        }
      ]
    }),
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();

  if (responseCode < 200 || responseCode >= 300) {
    saveLineMessageLog_(
      ss,
      new Date(),
      "送信失敗",
      staffName || "",
      lineUserId || "",
      "LINE Push送信失敗：" + responseCode + " " + response.getContentText()
    );
    return;
  }

  saveLineMessageLog_(
    ss,
    new Date(),
    "送信",
    staffName || "",
    lineUserId || "",
    message
  );
}

/**
 * LINE応答メッセージ配信
 */
function sendReplyMessages_(replyMessages) {
  if (!replyMessages || replyMessages.length === 0) return;

  const token = PropertiesService
    .getScriptProperties()
    .getProperty("LINE_CHANNEL_ACCESS_TOKEN");

  if (!token) {
    saveLineMessageLog_(
      SpreadsheetApp.getActiveSpreadsheet(),
      new Date(),
      "送信失敗",
      "",
      "",
      "LINE_CHANNEL_ACCESS_TOKENが未設定です。"
    );
    return;
  }

  const url = "https://api.line.me/v2/bot/message/reply";
  const usedTokens = {};

  replyMessages.forEach(item => {
    if (!item.replyToken || !item.message) return;

    if (usedTokens[item.replyToken]) return;
    usedTokens[item.replyToken] = true;

    const payload = {
      replyToken: item.replyToken,
      messages: [
        {
          type: "text",
          text: item.message
        }
      ]
    };

    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + token
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();

    if (responseCode < 200 || responseCode >= 300) {
      saveLineMessageLog_(
        SpreadsheetApp.getActiveSpreadsheet(),
        new Date(),
        "送信失敗",
        item.staffName || "",
        item.userId || "",
        "LINE返信失敗：" + responseCode + " " + response.getContentText()
      );
      return;
    }

    saveLineMessageLog_(
      SpreadsheetApp.getActiveSpreadsheet(),
      new Date(),
      "送信",
      item.staffName || "",
      item.userId || "",
      item.message
    );
  });
}

/**
 * カタカナ名義の標準化（GMO入出金チェック用）
 */
function normalizeKana_(text) {
  return String(text || "")
    .replace(/\s/g, "")
    .replace(/[ｰー－]/g, "ー")
    .replace(/振込/g, "")
    .trim();
}

/**
 * コピペしたGMOネット銀行入出金データを「入出金明細」へ反映、利用者自動照合、回数券集計までを実行
 */
function importGmoNyushukkinFromPaste() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pasteSheet = ss.getSheetByName(GMO_PASTE_SHEET_NAME);

  if (!pasteSheet) {
    SpreadsheetApp.getUi().alert("GMO入出金CSV貼付シートがありません。" + String.fromCharCode(10) + "CSVを貼り付けるシートを作成してください。");
    return;
  }

  const startRow = 17; // 17行目から明細開始と想定
  const lastRow = pasteSheet.getLastRow();

  const values = pasteSheet
    .getRange(startRow, 1, Math.max(lastRow - startRow + 1, 1), 6)
    .getValues();

  if (values.length < 2) {
    SpreadsheetApp.getUi().alert("GMO入出金CSV貼付シートにCSVデータを貼り付けてください。");
    return;
  }

  importGmoRows_(ss, values);
  reconcileNyushukkinUsers();
  updateCouponManagement();

  SpreadsheetApp.getUi().alert("GMO入出金CSVを取り込み、回数券管理を更新しました。");
}

function importGmoRows_(ss, rows) {
  const importSheet = ss.getSheetByName("入出金明細");
  const userSheet = ss.getSheetByName("利用者マスタ");

  if (!importSheet) {
    SpreadsheetApp.getUi().alert("入出金明細シートがありません。");
    return;
  }

  if (!userSheet) {
    SpreadsheetApp.getUi().alert("利用者マスタシートがありません。");
    return;
  }

  ensureNyushukkinHeader_(importSheet);

  const userMap = getUserKanaMap_(userSheet);
  const existingKeys = getExistingNyushukkinKeys_(importSheet);
  const output = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    const tradeDate = String(row[0] || "").trim();
    const summary = String(row[1] || "").trim();

    if (!tradeDate || tradeDate === "日付" || tradeDate === "取引日") continue;
    if (!summary) continue;

    const inAmount = Number(row[2]) || 0;
    const outAmount = Number(row[3]) || 0;
    const balance = Number(row[4]) || 0;
    const memo = row[5] || "";

    const key = makeNyushukkinKey_(tradeDate, summary, inAmount, outAmount, balance);
    if (existingKeys[key]) continue;

    let matchedUser = "";
    let status = "未照合";

    if (inAmount > 0) {
      const normalized = normalizeKana_(summary);

      Object.keys(userMap).forEach(kana => {
        if (normalized.indexOf(kana) !== -1) {
          matchedUser = userMap[kana];
          status = "照合済";
        }
      });
    }

    output.push([
      tradeDate,
      summary,
      inAmount,
      outAmount,
      balance,
      memo,
      matchedUser,
      status
    ]);
  }

  if (output.length > 0) {
    importSheet
      .getRange(importSheet.getLastRow() + 1, 1, output.length, output[0].length)
      .setValues(output);
  }
}

function ensureNyushukkinHeader_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "取引日", "摘要", "入金額", "出金額", "残高", "メモ", "利用者名", "照合結果"
    ]);
    return;
  }

  const firstCell = sheet.getRange(1, 1).getValue();

  if (!firstCell) {
    sheet.getRange(1, 1, 1, 8).setValues([[
      "取引日", "摘要", "入金額", "出金額", "残高", "メモ", "利用者名", "照合結果"
    ]]);
  }
}

function getUserKanaMap_(userSheet) {
  const userValues = userSheet.getDataRange().getValues();
  const userMap = {};

  for (let i = 1; i < userValues.length; i++) {
    const formalName = userValues[i][0];

    if (!formalName) continue;

    const kanaCandidates = [
      userValues[i][1], // B:振込名義カナ1
      userValues[i][2], // C:振込名義カナ2
      userValues[i][3]  // D:振込名義カナ3
    ];

    kanaCandidates.forEach(kana => {
      const kanaName = normalizeKana_(kana);
      if (!kanaName) return;
      userMap[kanaName] = formalName;
    });
  }

  return userMap;
}

function getExistingNyushukkinKeys_(sheet) {
  const values = sheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const key = makeNyushukkinKey_(
      values[i][0],
      values[i][1],
      values[i][2],
      values[i][3],
      values[i][4]
    );

    map[key] = true;
  }

  return map;
}

function makeNyushukkinKey_(tradeDate, summary, inAmount, outAmount, balance) {
  return [
    String(tradeDate || "").trim(),
    String(summary || "").trim(),
    Number(inAmount) || 0,
    Number(outAmount) || 0,
    Number(balance) || 0
  ].join("|");
}

function reconcileNyushukkinUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nyushukkinSheet = ss.getSheetByName("入出金明細");
  const userSheet = ss.getSheetByName("利用者マスタ");

  if (!nyushukkinSheet || !userSheet) {
    SpreadsheetApp.getUi().alert("入出金明細または利用者マスタがありません。");
    return;
  }

  const values = nyushukkinSheet.getDataRange().getValues();
  if (values.length < 2) return;

  const userMap = getUserKanaMap_(userSheet);
  let updatedCount = 0;

  for (let i = 1; i < values.length; i++) {
    const summary = values[i][1];
    const inAmount = Number(values[i][2]) || 0;

    if (inAmount <= 0) continue;

    const normalized = normalizeKana_(summary);
    let matchedUser = "";

    Object.keys(userMap).forEach(kana => {
      if (normalized.indexOf(kana) !== -1) {
        matchedUser = userMap[kana];
      }
    });

    if (matchedUser) {
      nyushukkinSheet.getRange(i + 1, 7).setValue(matchedUser);
      nyushukkinSheet.getRange(i + 1, 8).setValue("照合済");
      updatedCount++;
    }
  }

  SpreadsheetApp.getUi().alert("入出金明細の再照合が完了しました。更新件数：" + updatedCount + "件");
}

/**
 * 回数券残数及び未入金額、不足額、直近入金実績の自動集計
 */
function updateCouponManagement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const couponSheet = ss.getSheetByName(COUPON_SHEET_NAME);
  const userSheet = ss.getSheetByName("利用者マスタ");
  const nyushukkinSheet = ss.getSheetByName("入出金明細");
  const visitSheet = ss.getSheetByName(VISIT_RESULT_SHEET_NAME);

  if (!couponSheet) {
    SpreadsheetApp.getUi().alert("回数券管理シートがありません。");
    return;
  }

  if (!userSheet || !nyushukkinSheet || !visitSheet) {
    SpreadsheetApp.getUi().alert("利用者マスタ・入出金明細・訪問実績のいずれかがありません。");
    return;
  }

  const existingMemoMap = getExistingCouponMemoMap_(couponSheet);
  const users = getCouponUsers_(userSheet);
  const paymentMap = getUserPaymentMap_(nyushukkinSheet);
  const usageMap = getUserUsageMap_(visitSheet);

  const output = [[
    "利用者名", "回数券残数", "未払い残高", "最終入金日", "最終入金額", "1回単価", "不足金額", "メモ"
  ]];

  users.forEach(user => {
    const userName = user.userName;
    const unitPrice = user.unitPrice || existingMemoMap[userName]?.unitPrice || DEFAULT_REHAB_UNIT_PRICE;
    const paidAmount = paymentMap[userName]?.total || 0;
    const lastPaymentDate = paymentMap[userName]?.lastDate || "";
    const lastPaymentAmount = paymentMap[userName]?.lastAmount || "";
    const usedCount = usageMap[userName] || 0;

    const paidCount = Math.floor(paidAmount / unitPrice);
    const couponBalance = paidCount - usedCount;
    const unpaidAmount = Math.max((usedCount * unitPrice) - paidAmount, 0);
    const shortageAmount = unpaidAmount;
    const memo = existingMemoMap[userName]?.memo || "";

    output.push([
      userName,
      couponBalance,
      unpaidAmount,
      lastPaymentDate,
      lastPaymentAmount,
      unitPrice,
      shortageAmount,
      memo
    ]);
  });

  couponSheet.clearContents();
  couponSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}

function getCouponUsers_(userSheet) {
  const values = userSheet.getDataRange().getValues();
  const users = [];
  const seen = {};

  for (let i = 1; i < values.length; i++) {
    const userName = values[i][0];
    if (!userName || seen[userName]) continue;

    seen[userName] = true;
    users.push({
      userName,
      unitPrice: Number(values[i][4]) || DEFAULT_REHAB_UNIT_PRICE // E列：回数券単価
    });
  }

  return users;
}

function getUserPaymentMap_(nyushukkinSheet) {
  const values = nyushukkinSheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const tradeDate = values[i][0];
    const inAmount = Number(values[i][2]) || 0;
    const userName = values[i][6];
    const status = values[i][7];

    if (!isOnOrAfterCouponStart_(tradeDate)) continue;
    if (!userName || inAmount <= 0 || status !== "照合済") continue;

    if (!map[userName]) {
      map[userName] = {
        total: 0,
        lastDate: "",
        lastAmount: ""
      };
    }

    map[userName].total += inAmount;
    map[userName].lastDate = tradeDate;
    map[userName].lastAmount = inAmount;
  }

  return map;
}

function getUserUsageMap_(visitSheet) {
  const values = visitSheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const visitDate = values[i][0];
    const type = values[i][1];
    const userName = values[i][3];

    if (!isOnOrAfterCouponStart_(visitDate)) continue;
    if (type !== "終了") continue;
    if (!userName || isAmbiguousUserName_(userName) || userName === "不明") continue;

    if (!map[userName]) map[userName] = 0;
    map[userName] += 1;
  }

  return map;
}

function parseDateForCoupon_(value) {
  if (!value) return null;

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const text = String(value).trim();

  if (text.length === 8 && !isNaN(Number(text))) {
    return new Date(
      Number(text.substring(0, 4)),
      Number(text.substring(4, 6)) - 1,
      Number(text.substring(6, 8))
    );
  }

  const normalized = text.replace(/-/g, "/");
  const parts = normalized.split("/");

  if (parts.length >= 3) {
    return new Date(
      Number(parts[0]),
      Number(parts[1]) - 1,
      Number(parts[2])
    );
  }

  return null;
}

function isOnOrAfterCouponStart_(value) {
  const targetDate = parseDateForCoupon_(value);
  const startDate = parseDateForCoupon_(COUPON_START_DATE_TEXT);

  if (!targetDate || !startDate) return false;

  return targetDate.getTime() >= startDate.getTime();
}

function getExistingCouponMemoMap_(couponSheet) {
  const values = couponSheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const userName = values[i][0];
    if (!userName) continue;

    map[userName] = {
      unitPrice: Number(values[i][5]) || DEFAULT_REHAB_UNIT_PRICE,
      memo: values[i][7] || ""
    };
  }

  return map;
}

function runCouponAll() {
  importGmoNyushukkinFromPaste();
}

/**
 * 練習用のテスト利用者を全スタッフの担当利用者に追加する
 */
function addTestUserForAllStaff() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = ss.getSheetByName(STAFF_SHEET_NAME);
  const staffUserSheet = ss.getSheetByName(STAFF_USER_MASTER_SHEET_NAME);

  if (!staffSheet || !staffUserSheet) {
    SpreadsheetApp.getUi().alert("スタッフマスタ、またはスタッフ利用者マスタがありません。");
    return;
  }

  const staffValues = staffSheet.getDataRange().getValues();
  const staffCols = getStaffMasterColumnMap_(staffSheet);
  const staffUserValues = staffUserSheet.getDataRange().getValues();
  const existing = {};
  const rows = [];

  for (let i = 1; i < staffUserValues.length; i++) {
    const staffName = staffUserValues[i][0];
    const userName = staffUserValues[i][1];
    const key = normalizeName_(staffName) + "|" + normalizeName_(userName);
    existing[key] = true;
  }

  for (let i = 1; i < staffValues.length; i++) {
    const staffName = String(staffValues[i][staffCols.name] || "").trim();
    if (!staffName) continue;

    const key = normalizeName_(staffName) + "|" + normalizeName_(TEST_USER_NAME);
    if (existing[key]) continue;

    rows.push([
      staffName,
      TEST_USER_NAME,
      TEST_USER_NAME
    ]);
    existing[key] = true;
  }

  if (rows.length > 0) {
    staffUserSheet
      .getRange(staffUserSheet.getLastRow() + 1, 1, rows.length, rows[0].length)
      .setValues(rows);
  }

  updateLiffDisplayMaster(true);
  SpreadsheetApp.getUi().alert("全スタッフへテスト利用者を追加しました。追加：" + rows.length + "件");
}

/**
 * LIFF起動時に読む軽量マスタを作成・更新する
 */
function updateLiffDisplayMaster(suppressAlert) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = ss.getSheetByName(STAFF_SHEET_NAME);
  const staffUserSheet = ss.getSheetByName(STAFF_USER_MASTER_SHEET_NAME);

  if (!staffSheet || !staffUserSheet) {
    SpreadsheetApp.getUi().alert("スタッフマスタ、またはスタッフ利用者マスタがありません。");
    return;
  }

  ensureUserStatusColumn_(ss);
  const usersByStaff = buildUsersByStaffMap_(ss, staffUserSheet);
  const staffValues = staffSheet.getDataRange().getValues();
  const cols = getStaffMasterColumnMap_(staffSheet);
  const now = new Date();
  const rows = [];
  const initDataMap = {};

  for (let i = 1; i < staffValues.length; i++) {
    const staffName = String(staffValues[i][cols.name] || "").trim();
    const liffLineUserId = String(staffValues[i][cols.liffLineUserId] || "").trim();

    if (!staffName) continue;

    const users = usersByStaff[normalizeName_(staffName)] || [];
    const userNames = [TEST_USER_NAME].concat(users.filter(name => normalizeName_(name) !== normalizeName_(TEST_USER_NAME)));

    rows.push([
      liffLineUserId,
      staffName,
      userNames.join("\n"),
      now
    ]);

    if (liffLineUserId) {
      initDataMap[liffLineUserId] = {
        staffName: staffName,
        users: userNames.map(name => ({ name: name }))
      };
    }
  }

  const sheet = ss.getSheetByName(LIFF_DISPLAY_MASTER_SHEET_NAME) ||
    ss.insertSheet(LIFF_DISPLAY_MASTER_SHEET_NAME);
  const headers = [
    "LIFF用LINEユーザーID",
    "スタッフ名",
    "利用者一覧",
    "更新日時"
  ];

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#1f2933")
    .setFontColor("#ffffff")
    .setFontWeight("bold");
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), headers.length).setWrap(true);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 280);
  sheet.setColumnWidth(4, 150);
  saveLiffDisplayMasterProperties_(initDataMap);

  if (!suppressAlert) {
    SpreadsheetApp.getUi().alert("LIFF表示用マスタを更新しました。スタッフ：" + rows.length + "件");
  }
}

function setupUserStatusColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = ensureUserStatusColumn_(ss);
  SpreadsheetApp.getUi().alert(result.message);
}

function ensureUserStatusColumn_(ss) {
  const sheet = ss.getSheetByName("利用者マスタ");
  if (!sheet) {
    return {
      ok: false,
      message: "利用者マスタシートがありません。"
    };
  }

  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1).setValue("利用者名");
  }

  const headerMap = getHeaderColumnMap_(sheet);
  let statusCol = getColumnIndex_(headerMap, ["状態", "ステータス", "利用状態"], -1);
  const lastRow = Math.max(sheet.getLastRow(), 2);

  if (statusCol < 0) {
    statusCol = sheet.getLastColumn();
    sheet.getRange(1, statusCol + 1).setValue("状態");
    sheet.getRange(2, statusCol + 1, lastRow - 1, 1).setValue("利用中");
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["利用中", "終了", "休止"], true)
    .setAllowInvalid(true)
    .build();

  sheet.getRange(2, statusCol + 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  sheet.setColumnWidth(statusCol + 1, 100);
  sheet.getRange(1, statusCol + 1)
    .setBackground("#1f2933")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  return {
    ok: true,
    message: "利用者マスタに状態列を設定しました。利用中・終了・休止を選べます。"
  };
}

function getLiffDisplayMasterCacheVersion_() {
  return PropertiesService
    .getScriptProperties()
    .getProperty("LIFF_DISPLAY_MASTER_CACHE_VERSION") || "1";
}

function buildUsersByStaffMap_(ss, staffUserSheet) {
  const values = staffUserSheet.getDataRange().getValues();
  const activeUserMap = getActiveUserMap_(ss);
  const map = {};
  const seen = {};

  for (let i = 1; i < values.length; i++) {
    const staffName = String(values[i][0] || "").trim();
    const userName = String(values[i][1] || "").trim();
    const staffKey = normalizeName_(staffName);
    const userKey = normalizeName_(userName);

    if (!staffKey || !userName) continue;
    if (!isActiveUserForLiffDisplay_(activeUserMap, userName)) continue;
    if (!map[staffKey]) map[staffKey] = [];
    if (!seen[staffKey]) seen[staffKey] = {};
    if (seen[staffKey][userKey]) continue;

    seen[staffKey][userKey] = true;
    map[staffKey].push(userName);
  }

  Object.keys(map).forEach(staffKey => {
    map[staffKey].sort((a, b) => a.localeCompare(b, "ja"));
  });

  return map;
}

function getActiveUserMap_(ss) {
  const sheet = ss.getSheetByName("利用者マスタ");
  if (!sheet || sheet.getLastRow() < 2) return null;

  const headerMap = getHeaderColumnMap_(sheet);
  const nameCol = getColumnIndex_(headerMap, ["利用者名", "氏名", "名前"], 0);
  const statusCol = getColumnIndex_(headerMap, ["状態", "ステータス", "利用状態"], -1);

  if (statusCol < 0) return null;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const map = {};

  values.forEach(row => {
    const userName = String(row[nameCol] || "").trim();
    if (!userName) return;

    map[normalizeName_(userName)] = isActiveUserStatus_(row[statusCol]);
  });

  return map;
}

function isActiveUserForLiffDisplay_(activeUserMap, userName) {
  if (isTestUserName_(userName)) return true;
  if (!activeUserMap) return true;

  const key = normalizeName_(userName);
  if (activeUserMap[key] === undefined) return true;

  return activeUserMap[key];
}

function isActiveUserStatus_(status) {
  const text = normalizeName_(status);
  if (!text) return true;

  return !/終了|利用終了|停止|中止|休止|介入終了|退会/.test(text);
}

/**
 * 訪問予定と訪問実績を1枚の照合シートにまとめる
 */
function updateScheduleVisitComparison() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  const visitSheet = ss.getSheetByName(VISIT_RESULT_SHEET_NAME);

  if (!scheduleSheet || !visitSheet) {
    SpreadsheetApp.getUi().alert("訪問予定シート、または訪問実績シートがありません。");
    return;
  }

  const comparisonMap = {};

  collectScheduleComparisonRows_(scheduleSheet, comparisonMap);
  collectVisitComparisonRows_(visitSheet, comparisonMap);

  const rows = Object.keys(comparisonMap)
    .map(key => comparisonMap[key])
    .sort((a, b) => {
      const dateDiff = b.date.getTime() - a.date.getTime();
      if (dateDiff !== 0) return dateDiff;

      const staffDiff = a.staffName.localeCompare(b.staffName, "ja");
      if (staffDiff !== 0) return staffDiff;

      return a.userName.localeCompare(b.userName, "ja");
    })
    .map(item => {
      const status = getScheduleVisitStatus_(item);
      const note = getScheduleVisitNote_(item);

      return [
        Utilities.formatDate(item.date, "Asia/Tokyo", "yyyy/MM/dd"),
        getJapaneseWeekday_(item.date),
        item.userName,
        item.staffName,
        item.planned ? "予定あり" : "",
        item.startTimes.join("、"),
        item.endTimes.join("、"),
        status,
        item.scheduleRegisteredAts.join("、"),
        item.visitRegisteredAts.join("、"),
        note
      ];
    });

  const sheet = ss.getSheetByName(SCHEDULE_VISIT_COMPARISON_SHEET_NAME) ||
    ss.insertSheet(SCHEDULE_VISIT_COMPARISON_SHEET_NAME);
  const headers = [
    "日付",
    "曜日",
    "利用者",
    "担当スタッフ",
    "予定",
    "開始実績",
    "終了実績",
    "判定",
    "予定登録日時",
    "実績登録日時",
    "備考"
  ];

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  formatScheduleVisitComparisonSheet_(sheet, headers.length);
  SpreadsheetApp.getUi().alert("予定・実績照合シートを更新しました。件数：" + rows.length + "件");
}

function collectScheduleComparisonRows_(sheet, comparisonMap) {
  if (!sheet || sheet.getLastRow() < 2) return;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(sheet.getLastColumn(), 11)).getValues();

  values.forEach(row => {
    const registeredAt = row[0];
    const staffName = String(row[1] || "").trim();
    const userName = String(row[2] || "").trim();
    const visitDate = row[3];
    const status = row[8];
    if (!staffName || !userName || !visitDate) return;
    if (isTestUserName_(userName)) return;
    if (isCancelledScheduleStatus_(status)) return;

    const date = parseComparisonDate_(visitDate, registeredAt);
    if (!date) return;

    const item = getScheduleVisitComparisonItem_(comparisonMap, date, staffName, userName);
    item.planned = true;
    addUniqueText_(item.scheduleRegisteredAts, formatComparisonDateTime_(registeredAt));
  });
}

function collectVisitComparisonRows_(sheet, comparisonMap) {
  if (!sheet || sheet.getLastRow() < 2) return;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(sheet.getLastColumn(), 8)).getValues();

  values.forEach(row => {
    const registeredAt = row[0];
    const type = String(row[1] || "").trim();
    const staffName = String(row[2] || "").trim();
    const userName = String(row[3] || "").trim();
    const visitDate = row[4];
    const visitTime = row[5];
    if (!staffName || !userName || !visitDate) return;
    if (isTestUserName_(userName)) return;

    const date = parseComparisonDate_(visitDate, registeredAt);
    if (!date) return;

    const item = getScheduleVisitComparisonItem_(comparisonMap, date, staffName, userName);
    const timeText = formatComparisonTime_(visitTime || registeredAt);

    if (type === "開始") {
      addUniqueText_(item.startTimes, timeText);
    } else if (type === "終了") {
      addUniqueText_(item.endTimes, timeText);
    } else {
      addUniqueText_(item.otherVisitTypes, type);
    }

    addUniqueText_(item.visitRegisteredAts, formatComparisonDateTime_(registeredAt));
  });
}

function getScheduleVisitComparisonItem_(comparisonMap, date, staffName, userName) {
  const key = [
    Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM-dd"),
    normalizeName_(staffName),
    normalizeName_(userName)
  ].join("|");

  if (!comparisonMap[key]) {
    comparisonMap[key] = {
      date: date,
      staffName: staffName,
      userName: userName,
      planned: false,
      startTimes: [],
      endTimes: [],
      otherVisitTypes: [],
      scheduleRegisteredAts: [],
      visitRegisteredAts: []
    };
  }

  return comparisonMap[key];
}

function getScheduleVisitStatus_(item) {
  const hasStart = item.startTimes.length > 0;
  const hasEnd = item.endTimes.length > 0;
  const hasVisit = hasStart || hasEnd || item.otherVisitTypes.length > 0;

  if (item.planned && hasStart && hasEnd) return "OK";
  if (item.planned && !hasVisit) return "未実施";
  if (!item.planned && hasVisit) return "予定外";
  if (item.planned && hasVisit) return "不完全";

  return "";
}

function getScheduleVisitNote_(item) {
  const notes = [];

  if (item.planned && item.startTimes.length > 0 && item.endTimes.length === 0) {
    notes.push("開始のみ");
  }

  if (item.planned && item.startTimes.length === 0 && item.endTimes.length > 0) {
    notes.push("終了のみ");
  }

  if (!item.planned && (item.startTimes.length > 0 || item.endTimes.length > 0)) {
    notes.push("予定にない実績");
  }

  if (item.startTimes.length > 1) {
    notes.push("開始が複数");
  }

  if (item.endTimes.length > 1) {
    notes.push("終了が複数");
  }

  if (item.otherVisitTypes.length > 0) {
    notes.push("開始/終了以外：" + item.otherVisitTypes.join("、"));
  }

  return notes.join("、");
}

function parseComparisonDate_(value, fallbackDate) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const text = String(value || "").trim();
  if (!text) return null;

  const ymd = text.match(/^(\d{4})[-\/年](\d{1,2})[-\/月](\d{1,2})日?/);
  if (ymd) {
    return new Date(Number(ymd[1]), Number(ymd[2]) - 1, Number(ymd[3]));
  }

  const md = text.match(/^(\d{1,2})[\/月](\d{1,2})日?/);
  if (md) {
    const base = Object.prototype.toString.call(fallbackDate) === "[object Date]" && !isNaN(fallbackDate.getTime())
      ? fallbackDate
      : new Date();
    return new Date(base.getFullYear(), Number(md[1]) - 1, Number(md[2]));
  }

  const date = new Date(text);
  if (!isNaN(date.getTime())) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  return null;
}

function formatComparisonTime_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, "Asia/Tokyo", "H:mm");
  }

  const text = String(value || "").trim();
  const match = text.match(/^(\d{1,2}):(\d{2})/);
  if (match) return Number(match[1]) + ":" + match[2];

  return text;
}

function formatComparisonDateTime_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, "Asia/Tokyo", "yyyy/MM/dd H:mm");
  }

  return String(value || "").trim();
}

function getJapaneseWeekday_(date) {
  return ["日", "月", "火", "水", "木", "金", "土"][date.getDay()];
}

function addUniqueText_(list, value) {
  const text = String(value || "").trim();
  if (text && list.indexOf(text) === -1) {
    list.push(text);
  }
}

function formatScheduleVisitComparisonSheet_(sheet, columnCount) {
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  const statusRange = sheet.getRange(2, 1, maxRows, columnCount);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$H2="OK"')
      .setBackground("#dff6e5")
      .setRanges([statusRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$H2="未実施"')
      .setBackground("#fde2e2")
      .setRanges([statusRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$H2="予定外"')
      .setBackground("#fff1c2")
      .setRanges([statusRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$H2="不完全"')
      .setBackground("#ffe4c7")
      .setRanges([statusRange])
      .build()
  ];

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, columnCount)
    .setBackground("#1f2933")
    .setFontColor("#ffffff")
    .setFontWeight("bold");
  sheet.getRange(1, 1, sheet.getMaxRows(), columnCount).setVerticalAlignment("middle");
  sheet.getRange(1, 1, sheet.getMaxRows(), columnCount).setWrap(true);
  sheet.setConditionalFormatRules(rules);

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), columnCount).createFilter();
  sheet.autoResizeColumns(1, columnCount);
}

/**
 * Googleフォームの回答シートから、スタッフマスタに必要な項目だけを追記する
 *
 * 使い方：
 * 1. フォーム回答シートを開く
 * 2. メニューから「スタッフアンケートをスタッフマスタへ取込」を実行
 *
 * 既にスタッフマスタに同じ氏名がある行は追記しない。
 */
function importStaffQuestionnaireToStaffMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = getStaffQuestionnaireSourceSheet_();
  const staffSheet = ss.getSheetByName(STAFF_SHEET_NAME);

  if (!staffSheet) {
    SpreadsheetApp.getUi().alert("スタッフマスタシートがありません。");
    return;
  }

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(
      "スタッフアンケートの集計スプレッドシートを開けませんでした。\n\n" +
      "GAS実行アカウントに、アンケート集計スプレッドシートの閲覧権限があるか確認してください。"
    );
    return;
  }

  if (sourceSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("取込対象の回答データがありません。");
    return;
  }

  ensureStaffMasterHeaders_(staffSheet);

  const sourceValues = sourceSheet.getDataRange().getValues();
  const sourceHeaderMap = getHeaderColumnMap_(sourceSheet);
  const staffCols = getStaffMasterColumnMap_(staffSheet);
  const existingStaffs = getExistingStaffNameMap_(staffSheet, staffCols.name);
  const rows = [];
  const skipped = [];

  for (let i = 1; i < sourceValues.length; i++) {
    const sourceRow = sourceValues[i];
    const staffName = getQuestionnaireValue_(sourceRow, sourceHeaderMap, [
      "スタッフ名",
      "氏名",
      "お名前",
      "名前",
      "フルネーム"
    ], 1);

    if (!staffName) continue;

    const staffKey = normalizeName_(staffName);

    if (existingStaffs[staffKey]) {
      skipped.push(staffName);
      continue;
    }

    const row = createBlankStaffMasterRow_(staffSheet);

    row[staffCols.name] = staffName;
    row[staffCols.address] = getQuestionnaireValue_(sourceRow, sourceHeaderMap, [
      "住所",
      "現住所",
      "ご住所"
    ], 4);
    row[staffCols.bankCode] = getQuestionnaireBankCode_(sourceRow, sourceHeaderMap);
    row[staffCols.branchCode] = formatCode_(getQuestionnaireValue_(sourceRow, sourceHeaderMap, [
      "支店番号",
      "支店コード",
      "店番"
    ], 13), 3);
    row[staffCols.accountType] = getQuestionnaireValue_(sourceRow, sourceHeaderMap, [
      "預金種目",
      "口座種別"
    ], -1) || "1";
    row[staffCols.accountNumber] = formatCode_(getQuestionnaireValue_(sourceRow, sourceHeaderMap, [
      "口座番号"
    ], 14), 7);
    row[staffCols.receiverName] = getQuestionnaireValue_(sourceRow, sourceHeaderMap, [
      "受取人名",
      "口座名義",
      "口座名義人",
      "振込名義"
    ], -1);

    rows.push(row);
    existingStaffs[staffKey] = true;
  }

  if (rows.length > 0) {
    staffSheet
      .getRange(staffSheet.getLastRow() + 1, 1, rows.length, rows[0].length)
      .setValues(rows);
  }

  SpreadsheetApp.getUi().alert(
    "スタッフマスタへの取込が完了しました。\n\n" +
    "取込元：" + sourceSheet.getParent().getName() + " / " + sourceSheet.getName() + "\n" +
    "追加：" + rows.length + "件\n" +
    "既存のためスキップ：" + skipped.length + "件\n\n" +
    "LINEユーザーID、LIFF用LINEユーザーID、カレンダーIDは必要に応じて後で入力してください。"
  );
}

function getStaffQuestionnaireSourceSheet_() {
  try {
    const sourceSs = SpreadsheetApp.openById(STAFF_QUESTIONNAIRE_SPREADSHEET_ID);
    const sheets = sourceSs.getSheets();

    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === STAFF_QUESTIONNAIRE_SHEET_ID) {
        return sheets[i];
      }
    }

    return sheets[0] || null;
  } catch (error) {
    return null;
  }
}

function ensureStaffMasterHeaders_(sheet) {
  const headers = [
    "スタッフ名",
    "LINE表示名",
    "LINEユーザーID",
    "LIFF用LINEユーザーID",
    "カレンダーID",
    "銀行コード",
    "支店番号",
    "預金種目",
    "口座番号",
    "受取人名",
    "給与明細フォルダID",
    "住所"
  ];

  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];

  headers.forEach((header, index) => {
    if (!current[index]) {
      sheet.getRange(1, index + 1).setValue(header);
    }
  });
}

function getExistingStaffNameMap_(sheet, staffNameCol) {
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  values.forEach(row => {
    const staffName = String(row[staffNameCol] || "").trim();
    if (staffName) map[normalizeName_(staffName)] = true;
  });

  return map;
}

function createBlankStaffMasterRow_(sheet) {
  const length = Math.max(sheet.getLastColumn(), STAFF_COL_ADDRESS + 1);
  return new Array(length).fill("");
}

function getQuestionnaireValue_(row, headerMap, headerNames, fallbackIndex) {
  const col = getColumnIndex_(headerMap, headerNames, fallbackIndex);
  if (col < 0 || col >= row.length) return "";

  return String(row[col] || "").trim();
}

function getQuestionnaireBankCode_(row, headerMap) {
  const explicitBankCode = getQuestionnaireValue_(row, headerMap, [
    "銀行コード",
    "金融機関コード"
  ], -1);

  if (explicitBankCode) {
    return formatCode_(explicitBankCode, 4);
  }

  const bankName = getQuestionnaireValue_(row, headerMap, [
    "銀行名",
    "金融機関名",
    "振込先銀行",
    "振込先金融機関"
  ], 12);

  return getBankCodeByName_(bankName);
}

function getBankCodeByName_(bankName) {
  const normalized = normalizeName_(bankName);
  const map = {
    "ゆうちょ銀行": "9900",
    "三菱UFJ銀行": "0005",
    "三井住友銀行": "0009",
    "みずほ銀行": "0001",
    "りそな銀行": "0010",
    "楽天銀行": "0036",
    "PayPay銀行": "0033",
    "住信SBIネット銀行": "0038",
    "碧海信用金庫": "1560"
  };

  return map[normalized] || "";
}

/**
 * スプレッドシート起動時処理（独自カスタムメニューの設置）
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📌 自費リハ管理")
    .addItem("🧾 スタッフアンケートをスタッフマスタへ取込", "importStaffQuestionnaireToStaffMaster")
    .addSeparator()
    .addItem("📅 予定・実績照合を更新", "updateScheduleVisitComparison")
    .addItem("⚡ LIFF表示用マスタを更新", "updateLiffDisplayMaster")
    .addItem("🧑 LINEユーザー一覧を更新", "updateLineUserDirectoryLinks")
    .addItem("🔗 LINEユーザー一覧をマスタへ反映", "applyLineUserDirectoryToMasters")
    .addItem("👤 利用者状態列を設定", "setupUserStatusColumn")
    .addItem("🧪 テスト利用者を全スタッフに追加", "addTestUserForAllStaff")
    .addSeparator()
    .addItem("💰 給与関連を全て実行", "runPayrollAll")
    .addItem("📒 賃金台帳を更新", "createWageLedger")
    .addItem("📍 距離と交通費を更新", "updateDistanceAndTravelCosts")
    .addItem("📂 給与PDFをスタッフフォルダへ移動", "movePayrollPdfsToStaffFolders")
    .addSeparator()
    .addItem("🎫 回数券を更新", "runCouponAll")
    .addToUi();
}
