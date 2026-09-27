// 訪問の開始を登録したときに、回数券が今回で使い切りになる（残り1回）か、すでに残りがない
// （0回以下）場合、LINE連携済みの利用者・家族へお知らせを送り、スタッフへの返信に一文を足す。
// 管理者へは送らない（LINEの無料枠を使わないため）。回数券の残数は「終了」の登録と入金取込で
// 更新されるので、開始の時点では今回の訪問はまだ差し引かれていない。
const COUPON_NOTICE_PROPERTY_PREFIX = "COUPON_NOTICE:";

function couponNoticeOnVisitStart_(ss, userName, visitDate) {
  try {
    if (isTestUserName_(userName)) return "";
    const coupon = readCouponDisplayMap_(ss)[normalizeName_(userName)];
    if (!coupon) return "";
    const balance = Number(coupon.balance) || 0;
    if (balance > 1) return "";

    const name = String(userName || "").replace(/\s*様$/, "");
    const dateText = couponNoticeDateText_(visitDate);
    if (claimCouponNotice_(name, dateText)) {
      const text = buildCouponNoticeUserMessage_(name, dateText, balance);
      couponNoticeRecipients_(ss, name).forEach(id => sendPushMessage_(ss, id, name + "（回数券のお知らせ）", text));
    }
    return buildCouponNoticeStaffNote_(name, balance);
  } catch (error) {
    console.error("couponNoticeOnVisitStart_ failed: " + error.message);
    return "";
  }
}

function buildCouponNoticeUserMessage_(name, dateText, balance) {
  const body = balance >= 1
    ? "本日（" + dateText + "）の訪問で、お手元の回数券をすべてご利用いただくことになります。\n" +
      "次回以降の訪問に向けて、次の回数券分のお振込みをお願いいたします。お振込み先はいつもの口座です。\n"
    : "お手元の回数券はすでにすべてご利用いただいているため、本日（" + dateText + "）の訪問は未精算となります" +
      "（本日分を含めて未精算 " + (1 - balance) + "回分）。\n" +
      "お手数ですが、お振込みをお願いいたします。お振込み先はいつもの口座です。\n";
  return name + " 様\n" +
    "いつもひろがりの訪問リハビリをご利用いただき、ありがとうございます。\n" +
    body +
    "行き違いでお振込み済みの場合は、ご容赦ください。\n" +
    "ご不明な点がありましたら、このLINEにご返信ください。\n" +
    "合同会社ひろがり";
}

function buildCouponNoticeStaffNote_(name, balance) {
  return "\n\n※" + name + "様の回数券は" +
    (balance >= 1 ? "今回で使い切りになります。" : "残りがなく、今回は未精算になります。") +
    "お金のことを聞かれたら、運営から案内があるとお伝えください（その場で受け取らないでください）。";
}

// 同じ利用者・同じ訪問日には1回だけ送る（開始の登録し直しで二重に届かないように）。
function claimCouponNotice_(name, dateText) {
  const props = PropertiesService.getScriptProperties();
  const key = COUPON_NOTICE_PROPERTY_PREFIX + normalizeName_(name);
  const stamp = new Date().getFullYear() + ":" + dateText;
  if (props.getProperty(key) === stamp) return false;
  props.setProperty(key, stamp);
  return true;
}

function couponNoticeRecipients_(ss, name) {
  const sheet = ss.getSheetByName(USER_MASTER_SHEET_NAME);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const idCol = headers.indexOf(USER_ID_HEADER);
  const nameCol = headers.indexOf("利用者名");
  if (idCol < 0 || nameCol < 0) return [];
  const rows = values.filter(row => normalizeName_(row[nameCol]) === normalizeName_(name));
  if (rows.length !== 1 || !rows[0][idCol]) return [];
  try {
    return getConfirmedUserMessagingIds_(ss, String(rows[0][idCol]));
  } catch (error) {
    if (error.userLinkSafe) return [];
    throw error;
  }
}

function couponNoticeDateText_(visitDate) {
  const match = String(visitDate || "").match(/^(\d{1,2})\/(\d{1,2})$/);
  if (match) return Number(match[1]) + "月" + Number(match[2]) + "日";
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "M月d日");
}
