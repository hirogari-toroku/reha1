function getAdminUserLinkSummary_(ss) {
  const sheet = ss.getSheetByName("利用者LINE連携");
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 14).getValues();
  if (JSON.stringify(values.shift()) !== JSON.stringify(USER_LINK_HEADERS)) return map;
  const counts = {};
  values.forEach(row => { if (row[4]) counts[row[4]] = (counts[row[4]] || 0) + 1; });
  values.forEach(row => {
    const id = String(row[0] || "").trim();
    if (!id) return;
    const item = map[id] || (map[id] = { liffLinked: false, messagingLinked: false, viewers: [] });
    const approved = row[3] === "確認済み" && row[5] === "連携済み" &&
      /^U[0-9a-f]{32}$/.test(String(row[4] || "")) && counts[row[4]] === 1;
    item.liffLinked = item.liffLinked || approved;
    item.messagingLinked = item.messagingLinked || (approved && /^U[0-9a-f]{32}$/.test(String(row[1] || "")));
    item.viewers.push({ name: String(row[12] || "表示名未取得"), relation: String(row[2] || "続柄未確認"),
      state: approved ? "確認済み・連携済み" : String(row[5] || "未連携") + "・" + String(row[3] || "未確認") });
  });
  return map;
}
