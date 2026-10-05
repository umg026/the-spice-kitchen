const http = require("http");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const ROOT = __dirname;
const MENU_FILE = path.join(ROOT, "menu.json");
const PORT = Number(process.env.PORT || 3000);
const HOST = process.env.HOST || "127.0.0.1";
const USERNAME = "umang";
const PASSWORD = "Umang@123";
const sessions = new Set();

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".webp": "image/webp",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon"
};

function send(res, status, body, headers = {}) {
  res.writeHead(status, headers);
  res.end(body);
}

function sendJSON(res, status, data, headers = {}) {
  send(res, status, JSON.stringify(data), {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    ...headers
  });
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", chunk => {
      body += chunk;
      if (body.length > 1024 * 1024) {
        reject(new Error("Request body is too large"));
        req.destroy();
      }
    });
    req.on("end", () => resolve(body));
    req.on("error", reject);
  });
}

function getCookie(req, name) {
  const cookies = req.headers.cookie || "";
  return cookies
    .split(";")
    .map(cookie => cookie.trim())
    .find(cookie => cookie.startsWith(name + "="))
    ?.slice(name.length + 1);
}

function isAuthed(req) {
  const token = getCookie(req, "tsk_admin");
  return Boolean(token && sessions.has(token));
}

function slugify(value) {
  return String(value || "item")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 60) || "item";
}

function normalizeItem(item) {
  const name = String(item.name || "New item").trim();
  const id = slugify(item.id || name);
  const price = Number(item.price);
  const next = {
    id,
    name,
    desc: String(item.desc || "").trim(),
    price: Number.isFinite(price) && price >= 0 ? Math.round(price * 100) / 100 : 0,
    veg: Boolean(item.veg)
  };
  if (item.soldOut) next.soldOut = true;
  return next;
}

function normalizeMenu(menu) {
  if (!Array.isArray(menu)) throw new Error("Menu must be an array");

  return menu.map((category, index) => {
    const next = {
      id: slugify(category.id || category.label || "category-" + (index + 1)),
      label: String(category.label || "Category").trim(),
      num: String(category.num || String(index + 1).padStart(2, "0")).trim(),
      note: String(category.note || "").trim()
    };

    if (Array.isArray(category.items)) {
      next.items = category.items.map(normalizeItem);
    }

    if (Array.isArray(category.subgroups)) {
      next.subgroups = category.subgroups.map(subgroup => ({
        label: String(subgroup.label || "Group").trim(),
        items: Array.isArray(subgroup.items) ? subgroup.items.map(normalizeItem) : [],
        ...(subgroup.addon ? { addon: normalizeItem(subgroup.addon) } : {})
      }));
    }

    if (category.addon) next.addon = normalizeItem(category.addon);
    if (!next.items && !next.subgroups) next.items = [];
    return next;
  });
}

async function handleApi(req, res, pathname) {
  if (pathname === "/api/session" && req.method === "GET") {
    sendJSON(res, 200, { authenticated: isAuthed(req) });
    return;
  }

  if (pathname === "/api/login" && req.method === "POST") {
    const body = JSON.parse(await readBody(req) || "{}");
    if (body.username === USERNAME && body.password === PASSWORD) {
      const token = crypto.randomBytes(32).toString("hex");
      sessions.add(token);
      sendJSON(res, 200, { success: true }, {
        "Set-Cookie": `tsk_admin=${token}; HttpOnly; SameSite=Lax; Path=/; Max-Age=43200`
      });
      return;
    }
    sendJSON(res, 401, { success: false, error: "Invalid username or password" });
    return;
  }

  if (pathname === "/api/logout" && req.method === "POST") {
    const token = getCookie(req, "tsk_admin");
    if (token) sessions.delete(token);
    sendJSON(res, 200, { success: true }, {
      "Set-Cookie": "tsk_admin=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0"
    });
    return;
  }

  if (pathname === "/api/menu" && req.method === "GET") {
    send(res, 200, fs.readFileSync(MENU_FILE), {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store"
    });
    return;
  }

  if (!isAuthed(req)) {
    sendJSON(res, 401, { error: "Not authenticated" });
    return;
  }

  if (pathname === "/api/menu" && req.method === "PUT") {
    const body = JSON.parse(await readBody(req) || "{}");
    const menu = normalizeMenu(body.menu);
    const json = JSON.stringify(menu, null, 2) + "\n";
    const tmpFile = MENU_FILE + ".tmp";
    fs.writeFileSync(tmpFile, json);
    fs.renameSync(tmpFile, MENU_FILE);
    sendJSON(res, 200, { success: true, menu });
    return;
  }

  sendJSON(res, 404, { error: "API route not found" });
}

function serveStatic(req, res, pathname) {
  const requested = pathname === "/" ? "/index.html" : pathname;
  const filePath = path.resolve(ROOT, "." + decodeURIComponent(requested));

  if (!filePath.startsWith(ROOT) || !fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
    send(res, 404, "Not found", { "Content-Type": "text/plain; charset=utf-8" });
    return;
  }

  const ext = path.extname(filePath).toLowerCase();
  const headers = {
    "Content-Type": mimeTypes[ext] || "application/octet-stream"
  };
  if (path.basename(filePath) === "menu.json") headers["Cache-Control"] = "no-store";
  fs.createReadStream(filePath).pipe(res.writeHead(200, headers));
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host}`);
    if (url.pathname.startsWith("/api/")) {
      await handleApi(req, res, url.pathname);
      return;
    }
    serveStatic(req, res, url.pathname);
  } catch (error) {
    sendJSON(res, 500, { error: error.message || "Server error" });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`The Spice Kitchen server running at http://localhost:${PORT}`);
  console.log(`Admin dashboard: http://localhost:${PORT}/admin.html`);
});
