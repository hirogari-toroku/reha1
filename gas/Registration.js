// Registration captures an unverified contact only; it never grants schedule access.
function saveRegistrationContact_(ss, lineUserId, displayName) {
  const startedAt = Date.now();
  const id = String(lineUserId || "").trim();
  const name = String(displayName || "").trim();
  if (!/^U[0-9a-f]{32}$/.test(id) || name.length > 200) {
    return { success: false, message: "LINE情報を確認できません。LINEから開き直してください。" };
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { success: false, message: "登録が混み合っています。少し待って開き直してください。" };
  try {
    const sheet = ensureLineUserDirectorySheet_(ss);
    const count = sheet.getLastRow() - 1;
    const rows = count > 0 ? sheet.getRange(2, 1, count, 13).getValues() : [];
    const matches = rows.map((row, index) => ({row, index})).filter(item => String(item.row[3] || "").trim() === id);
    if (matches.length > 1) return { success: false, message: "LINE情報が重複しています。管理者へご連絡ください。" };
    const target = matches[0];
    const row = target ? target.row.slice() : Array(13).fill("");
    const now = new Date();
    row[0] = row[0] || now;
    row[1] = now;
    row[3] = id;
    row[4] = name || row[4] || "";
    if (!target) row[9] = "未確認";
    row[10] = mergeDirectoryText_(row[10], "新規利用者登録LIFF");
    // Preserve the user's submitted name/relation message and any existing master hints.
    row[11] = row[11] || "新規利用者登録フォームへ遷移";
    if (target) sheet.getRange(target.index + 2, 1, 1, 13).setValues([row]);
    else sheet.appendRow(row);
    SpreadsheetApp.flush();
    return { success: true, lineUserId: id, displayName: name, serverMs: Date.now() - startedAt,
      message: "LINE情報を保存しました。利用申請フォームへ移動します。" };
  } finally { lock.releaseLock(); }
}
