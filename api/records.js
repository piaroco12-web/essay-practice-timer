const { timingSafeEqual } = require("node:crypto");

module.exports = async function records(req, res) {
  res.setHeader("Cache-Control", "no-store");
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ ok: false });
  }

  const { RECORDS_ACCESS_KEY, APPS_SCRIPT_URL, APPS_SCRIPT_SECRET } = process.env;
  if (!RECORDS_ACCESS_KEY || !APPS_SCRIPT_SECRET ||
      !/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(APPS_SCRIPT_URL || "")) {
    return res.status(503).json({ ok: false, error: "저장 서버가 아직 설정되지 않았습니다." });
  }
  const supplied = Buffer.from(req.headers.authorization || "");
  const expected = Buffer.from("Bearer " + RECORDS_ACCESS_KEY);
  if (supplied.length !== expected.length || !timingSafeEqual(supplied, expected)) {
    return res.status(401).json({ ok: false, error: "개인 접근키를 확인해 주세요." });
  }

  const payload = req.body;
  if (!payload || !["save", "list"].includes(payload.action) || JSON.stringify(payload).length > 8192) {
    return res.status(400).json({ ok: false, error: "요청 형식을 확인해 주세요." });
  }

  try {
    const upstream = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: payload.action, record: payload.record, secret: APPS_SCRIPT_SECRET }),
      signal: AbortSignal.timeout(20000)
    });
    if (!upstream.ok) throw new Error("Apps Script request failed");
    const result = await upstream.json();
    if (!result.ok) {
      if (result.code === "INVALID_RECORD") {
        return res.status(400).json({ ok: false, error: "이름·주제·작성일과 기록 값을 확인해 주세요." });
      }
      throw new Error("Apps Script rejected request");
    }
    if (payload.action === "save" && result.id === payload.record?.id) {
      return res.status(200).json({ ok: true, id: result.id });
    }
    if (payload.action === "list" && Array.isArray(result.records)) {
      return res.status(200).json({ ok: true, records: result.records });
    }
    throw new Error("Invalid Apps Script response");
  } catch {
    return res.status(502).json({ ok: false, error: "시트 응답을 확인하지 못했습니다. 잠시 뒤 다시 시도해 주세요." });
  }
};
