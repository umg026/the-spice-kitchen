const crypto = require("crypto");

const COOKIE_NAME = "tsk_admin";
const SESSION_MAX_AGE = 12 * 60 * 60;

function getSessionSecret() {
  return process.env.ADMIN_SESSION_SECRET || "local-dev-session-secret-change-me";
}

function sign(value) {
  return crypto
    .createHmac("sha256", getSessionSecret())
    .update(value)
    .digest("base64url");
}

function createToken(username) {
  const payload = JSON.stringify({
    username,
    exp: Math.floor(Date.now() / 1000) + SESSION_MAX_AGE
  });
  const body = Buffer.from(payload).toString("base64url");
  return `${body}.${sign(body)}`;
}

function readCookies(req) {
  return Object.fromEntries(
    String(req.headers.cookie || "")
      .split(";")
      .map(cookie => cookie.trim())
      .filter(Boolean)
      .map(cookie => {
        const index = cookie.indexOf("=");
        return index === -1
          ? [cookie, ""]
          : [cookie.slice(0, index), decodeURIComponent(cookie.slice(index + 1))];
      })
  );
}

function verifyToken(token) {
  if (!token || !token.includes(".")) return false;
  const [body, signature] = token.split(".");
  if (signature !== sign(body)) return false;

  try {
    const payload = JSON.parse(Buffer.from(body, "base64url").toString("utf8"));
    return payload.exp && payload.exp > Math.floor(Date.now() / 1000);
  } catch (error) {
    return false;
  }
}

function isAuthenticated(req) {
  return verifyToken(readCookies(req)[COOKIE_NAME]);
}

function setSessionCookie(res, username) {
  res.setHeader(
    "Set-Cookie",
    `${COOKIE_NAME}=${encodeURIComponent(createToken(username))}; HttpOnly; SameSite=Lax; Secure; Path=/; Max-Age=${SESSION_MAX_AGE}`
  );
}

function clearSessionCookie(res) {
  res.setHeader(
    "Set-Cookie",
    `${COOKIE_NAME}=; HttpOnly; SameSite=Lax; Secure; Path=/; Max-Age=0`
  );
}

function adminUsername() {
  return process.env.ADMIN_USERNAME || "umang";
}

function adminPassword() {
  return process.env.ADMIN_PASSWORD || "Umang@123";
}

module.exports = {
  adminPassword,
  adminUsername,
  clearSessionCookie,
  isAuthenticated,
  setSessionCookie
};
