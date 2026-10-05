import { promises as fs } from "fs";
import path from "path";
import { NextResponse } from "next/server";

const MENU_FILE = path.join(process.cwd(), "..", "menu.json");
const MENU_BLOB_PATH = process.env.MENU_BLOB_PATH || "menu.json";

export const dynamic = "force-dynamic";

function json(data, status = 200) {
  return NextResponse.json(data, {
    status,
    headers: { "Cache-Control": "no-store" }
  });
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
  const price = Number(item.price);
  const next = {
    id: slugify(item.id || name),
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

function hasBlobConfig() {
  return Boolean(process.env.BLOB_READ_WRITE_TOKEN || process.env.BLOB_STORE_ID || process.env.VERCEL_OIDC_TOKEN);
}

async function blobSdk() {
  return import("@vercel/blob");
}

async function readMenuFile() {
  return JSON.parse(await fs.readFile(MENU_FILE, "utf8"));
}

async function writeMenuFile(menu) {
  const normalized = normalizeMenu(menu);
  await fs.writeFile(MENU_FILE, JSON.stringify(normalized, null, 2) + "\n");
  return normalized;
}

async function readMenu() {
  if (hasBlobConfig()) {
    const { get } = await blobSdk();
    const result = await get(MENU_BLOB_PATH, { access: "private", useCache: false });
    if (result?.statusCode === 200 && result.stream) {
      return JSON.parse(await new Response(result.stream).text());
    }
  }

  return readMenuFile();
}

async function writeMenu(menu) {
  const normalized = normalizeMenu(menu);

  if (hasBlobConfig()) {
    const { put } = await blobSdk();
    await put(MENU_BLOB_PATH, JSON.stringify(normalized, null, 2) + "\n", {
      access: "private",
      allowOverwrite: true,
      contentType: "application/json",
      cacheControlMaxAge: 60
    });
    return normalized;
  }

  if (process.env.VERCEL) {
    throw new Error("Vercel Blob is not configured. Create a Vercel Blob store for this project so BLOB_READ_WRITE_TOKEN is available.");
  }

  return writeMenuFile(normalized);
}

function deleteMenuEntry(menu, categoryId, itemId) {
  const categoryIndex = menu.findIndex(category => category.id === categoryId);
  if (categoryIndex === -1) throw new Error("Category not found");

  if (!itemId) {
    menu.splice(categoryIndex, 1);
    return menu;
  }

  const category = menu[categoryIndex];
  const collections = [];
  if (Array.isArray(category.items)) collections.push(category.items);
  if (Array.isArray(category.subgroups)) {
    category.subgroups.forEach(subgroup => {
      if (Array.isArray(subgroup.items)) collections.push(subgroup.items);
    });
  }

  for (const items of collections) {
    const itemIndex = items.findIndex(item => item.id === itemId);
    if (itemIndex !== -1) {
      items.splice(itemIndex, 1);
      return menu;
    }
  }

  throw new Error("Item not found");
}

function isAuthorized(request, body = {}) {
  const secret = process.env.MENU_API_SECRET || "umang@9328641633";
  return request.headers.get("x-menu-api-secret") === secret || body.secret === secret;
}

export async function GET() {
  try {
    return json(await readMenu());
  } catch (error) {
    return json({ error: error.message || "Menu API failed" }, 500);
  }
}

export async function PUT(request) {
  try {
    const body = await request.json().catch(() => ({}));
    if (!isAuthorized(request, body)) return json({ error: "Not authorized" }, 401);
    const menu = await writeMenu(body.menu);
    return json({ success: true, menu });
  } catch (error) {
    return json({ error: error.message || "Menu API failed" }, 500);
  }
}

export async function DELETE(request) {
  try {
    const body = await request.json().catch(() => ({}));
    if (!isAuthorized(request, body)) return json({ error: "Not authorized" }, 401);
    const url = new URL(request.url);
    const menu = await writeMenu(deleteMenuEntry(
      await readMenu(),
      body.categoryId || url.searchParams.get("categoryId"),
      body.itemId || url.searchParams.get("itemId")
    ));
    return json({ success: true, menu });
  } catch (error) {
    return json({ error: error.message || "Menu API failed" }, 500);
  }
}
