const { createHmac, randomBytes, scrypt, timingSafeEqual } = require("node:crypto");
const derive = require("node:util").promisify(scrypt);
const usernamePattern = /^[a-z0-9][a-z0-9._-]{2,31}$/;
const cookieName = "__Host-timer-session";
const sessionSeconds = 43200;

function sign(value, secret) {
  return createHmac("sha256", secret).update("timer-account-v1:" + value).digest("base64url");
}
function same(a, b) {
  const left = Buffer.from(a);
  const right = Buffer.from(b);
  return left.length === right.length && timingSafeEqual(left, right);
}
function userFromSession(req, secret) {
  const cookie = (req.headers.cookie || "").split(/;\s*/).find(s => s.startsWith(cookieName + "="));
  const parts = cookie?.slice(cookieName.length + 1).split("~") || [];
  const [user, expires, signature] = parts;
  return parts.length === 3 && usernamePattern.test(user) && /^\d+$/.test(expires) &&
    Number(expires) > Date.now() && same(signature, sign(user + "~" + expires, secret)) ? user : null;
}
function setSession(res, user, secret) {
  const payload = user ? user + "~" + (Date.now() + sessionSeconds * 1000) : "";
  const token = user ? payload + "~" + sign(payload, secret) : "";
  res.setHeader("Set-Cookie", cookieName + "=" + token + "; Path=/; HttpOnly; Secure; SameSite=Strict; Max-Age=" + (user ? sessionSeconds : 0));
}

module.exports = async function records(req, res) {
  res.setHeader("Cache-Control", "no-store");
  const fail = (status, error) => res.status(status).json({ ok: false, error });
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return fail(405, "POST 요청이 필요합니다.");
  }
  if (req.headers.origin && req.headers.origin !== "https://" + req.headers.host) {
    return fail(403, "타이머 화면에서 다시 시도해 주세요.");
  }
  const { APPS_SCRIPT_URL, APPS_SCRIPT_SECRET } = process.env;
  if (!APPS_SCRIPT_SECRET || !/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(APPS_SCRIPT_URL || "")) {
    return fail(503, "저장 서버가 아직 설정되지 않았습니다.");
  }
  const payload = req.body;
  const action = payload?.action;
  if (!["signup", "login", "logout", "session", "save", "list"].includes(action) || JSON.stringify(payload).length > 8192) {
    return fail(400, "요청 형식을 확인해 주세요.");
  }

  const sessionUser = userFromSession(req, APPS_SCRIPT_SECRET);
  if (action === "session") return res.status(200).json({ ok: true, user: sessionUser });
  if (action === "logout") {
    setSession(res, null, APPS_SCRIPT_SECRET);
    return res.status(200).json({ ok: true });
  }
  if ((action === "list" || action === "save") && !sessionUser) return fail(401, "먼저 로그인해 주세요.");

  async function callScript(data) {
    const response = await fetch(APPS_SCRIPT_URL, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ...data, secret: APPS_SCRIPT_SECRET }),
      signal: AbortSignal.timeout(20000)
    });
    if (!response.ok) throw new Error("Apps Script request failed");
    return response.json();
  }

  try {
    if (action === "signup" || action === "login") {
      const user = typeof payload.user === "string" ? payload.user.trim().toLowerCase() : "";
      const password = payload.password;
      if (!usernamePattern.test(user) || typeof password !== "string" || password.length < 8 || password.length > 128) {
        return fail(400, "아이디는 영문·숫자와 . _ - 로 3~32자, 비밀번호는 8~128자로 입력해 주세요.");
      }
      const ip = String(req.headers["x-vercel-forwarded-for"] || req.headers["x-forwarded-for"] || "").split(",")[0].trim();
      const client = createHmac("sha256", APPS_SCRIPT_SECRET).update("timer-limit:" + ip).digest("hex");
      if (action === "signup") {
        const salt = randomBytes(16).toString("hex");
        const hash = (await derive(password, salt, 64)).toString("hex");
        const result = await callScript({ action, user, salt, hash, client });
        if (result.code === "USER_EXISTS") return fail(409, "이미 사용 중인 아이디입니다.");
        if (result.code === "RATE_LIMIT") return fail(429, "시도가 많습니다. 잠시 뒤 다시 시도해 주세요.");
        if (!result.ok) throw new Error("Account creation failed");
      } else {
        const result = await callScript({ action: "account", user, client });
        if (result.code === "RATE_LIMIT") return fail(429, "시도가 많습니다. 잠시 뒤 다시 시도해 주세요.");
        if (!result.ok) throw new Error("Account lookup failed");
        const account = result.account;
        const salt = account?.salt || "0".repeat(32);
        const hash = (await derive(password, salt, 64)).toString("hex");
        if (!account || !same(hash, account.hash)) return fail(401, "아이디 또는 비밀번호를 확인해 주세요.");
      }
      setSession(res, user, APPS_SCRIPT_SECRET);
      return res.status(200).json({ ok: true, user });
    }

    const result = await callScript({ action, user: sessionUser, record: action === "save" ? { ...payload.record, name: sessionUser } : payload.record });
    if (result.code === "INVALID_RECORD") return fail(400, "주제·작성일과 기록 값을 확인해 주세요.");
    if (result.ok && action === "save" && result.id === payload.record?.id) {
      return res.status(200).json({ ok: true, id: result.id });
    }
    if (result.ok && action === "list" && Array.isArray(result.records)) {
      return res.status(200).json({ ok: true, records: result.records });
    }
    throw new Error("Invalid Apps Script response");
  } catch {
    return fail(502, "시트 응답을 확인하지 못했습니다. 잠시 뒤 다시 시도해 주세요.");
  }
};
