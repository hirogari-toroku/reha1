function liffDiagnosticId_(value) {
  return /^[a-z0-9-]{8,60}$/.test(String(value || "")) ? String(value) : "";
}

function recordLiffDiagnostic_(data) {
  const stages = ["page_open", "init_ok", "login_redirect", "profile_ok", "init_failed", "user_request", "user_pending", "user_shown", "user_rejected", "user_network_failed", "page_ready"];
  const id = liffDiagnosticId_(data.traceId);
  if (!id || stages.indexOf(data.stage) < 0) return { success: false };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return { success: false };
  try {
    const cache = CacheService.getScriptCache();
    const key = "liffDiagnostic:" + id + ":" + data.stage;
    if (cache.get(key)) return { success: true };
    const dayKey = "liffDiagnosticDay:" + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
    const count = Number(cache.get(dayKey) || 0);
    if (count >= 1000) return { success: false };
    cache.put(key, "1", 21600);
    cache.put(dayKey, String(count + 1), 21600);
    // Browser reports are untrusted: never accept identities, URLs or error text.
    saveLiffOperationLog_(SpreadsheetApp.getActiveSpreadsheet(), "client:" + data.stage, "", "", "", "", "", "", "端末申告・未認証", "受付番号：" + id, "");
    return { success: true };
  } catch (_) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function logUserLinkStage_(data, stage, profile) {
  try {
    saveLiffOperationLog_(SpreadsheetApp.getActiveSpreadsheet(), "userLink:" + stage,
      profile ? profile.userId : "", "", "", "", "", "", profile ? "LINE認証済み" : "本人未確認",
      "受付番号：" + (liffDiagnosticId_(data.traceId) || "なし"), profile ? profile.displayName || "" : "");
  } catch (_) { /* Diagnostic failures must not block reservation access. */ }
}
