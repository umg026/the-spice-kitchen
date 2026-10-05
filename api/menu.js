const { isAuthenticated } = require("./_auth");

const DEFAULT_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyS1wUpuiv3mY5Esaro37jhdpY4NBG6U9YUG8K5AhS1fD5nmR1kS68fYvL7bmMUKNyI/exec";

function scriptUrl() {
  return process.env.MENU_SCRIPT_URL || process.env.GOOGLE_SCRIPT_URL || DEFAULT_SCRIPT_URL;
}

function menuSecret() {
  return process.env.MENU_API_SECRET || "umang@9328641633";
}

async function readMenu() {
  const url = new URL(scriptUrl());
  url.searchParams.set("action", "menu");
  url.searchParams.set("v", Date.now().toString());

  const response = await fetch(url);
  const data = await response.json();
  if (!response.ok || data.success === false) {
    throw new Error(data.error || "Could not load menu from Google Sheet");
  }

  return data.menu || [];
}

async function saveMenu(menu) {
  if (!menuSecret()) {
    throw new Error("MENU_API_SECRET is not set in Vercel environment variables");
  }

  const response = await fetch(scriptUrl(), {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({
      type: "menu_save",
      secret: menuSecret(),
      menu
    })
  });

  const data = await response.json();
  if (!response.ok || data.success === false) {
    throw new Error(data.error || "Could not save menu to Google Sheet");
  }

  return data.menu || menu;
}

module.exports = async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");

  try {
    if (req.method === "GET") {
      res.status(200).json(await readMenu());
      return;
    }

    if (req.method === "PUT") {
      if (!isAuthenticated(req)) {
        res.status(401).json({ error: "Not authenticated" });
        return;
      }

      const menu = await saveMenu((req.body || {}).menu);
      res.status(200).json({ success: true, menu });
      return;
    }

    res.setHeader("Allow", "GET, PUT");
    res.status(405).json({ error: "Method not allowed" });
  } catch (error) {
    res.status(500).json({ error: error.message || "Menu API failed" });
  }
};
