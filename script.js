function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // ---------------------------------------------------------
    // VISITOR TRACKING
    // ---------------------------------------------------------
    if (data && data.type === "track") {
      return handleTrack_(data);
    }

    if (data && data.type === "menu_save") {
      return handleMenuSave_(data);
    }

    // ---------------------------------------------------------
    // ORDERS - KEEP EXISTING BEHAVIOUR
    // ---------------------------------------------------------
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Orders");

    if (!sheet) {
      throw new Error('Sheet "Orders" not found.');
    }

    let itemsText = "";

    if (Array.isArray(data.items)) {
      itemsText = data.items
        .map(item => {
          return `${item.name} x ${item.quantity}`;
        })
        .join(" | ");
    } else {
      itemsText = data.items || "";
    }

    sheet.appendRow([
      data.orderId || "",
      data.timestamp || new Date().toISOString(),
      data.customerName || "",
      data.phone || "",
      data.address || "",
      data.postcode || "",
      data.deliveryTime || "",
      itemsText,
      data.total || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {

    console.error(error);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/* =========================================================
   MENU TAB
   Production menu storage for the Vercel site/admin.
   Sheet tab name: Menu
========================================================= */

const MENU_TAB_NAME = "Menu";
const MENU_HEADERS = [
  "Category ID",
  "Category Label",
  "Category Num",
  "Category Note",
  "Group Label",
  "Item ID",
  "Item Name",
  "Description",
  "Price",
  "Veg",
  "Sold Out",
  "Sort"
];

const MENU_SEED = [
  {
    "id": "starters",
    "label": "Starters",
    "num": "01",
    "note": "Wok-fired and tandoor-fried, made to order.",
    "items": [
      {
        "id": "p-chilli",
        "name": "Paneer Chilli",
        "desc": "Crispy paneer tossed in a spicy chilli, garlic & soy sauce.",
        "price": 6.99,
        "veg": true
      },
      {
        "id": "veg-manch",
        "name": "Veg Manchurian",
        "desc": "Crispy vegetable dumplings tossed in a rich Manchurian sauce.",
        "price": 6.99,
        "veg": true
      },
      {
        "id": "chilli-chicken",
        "name": "Chilli Chicken",
        "desc": "Crispy chicken tossed with fresh chilli, peppers & Indo-Chinese sauce.",
        "price": 6.99,
        "veg": false
      },
      {
        "id": "chicken-manch",
        "name": "Chicken Manchurian",
        "desc": "Tender chicken tossed in a rich, tangy Manchurian gravy.",
        "price": 6.99,
        "veg": false
      },
      {
        "id": "chicken-majestic",
        "name": "Dragon Chicken",
        "desc": "Crispy chicken strips tossed in a creamy, spicy house sauce.",
        "price": 6.99,
        "veg": false
      },
      {
        "id": "egg-bhaji",
        "name": "Egg Pakoda",
        "desc": "Hard-boiled eggs coated in a thick, seasoned chickpea flour (gram flour/besan) batter and deep-fried until golden brown and intensely crispy.",
        "price": 1.99,
        "veg": false
      },
      {
        "id": "honey-chilli-potato",
        "name": "Honey Chilli Potato",
        "desc": "Crispy potato tossed in a sweet, spicy honey-chilli glaze.",
        "price": 6.49,
        "veg": true
      },
      {
        "id": "dragon-paneer",
        "name": "Dragon Paneer",
        "desc": "Crispy paneer tossed in a fiery, tangy dragon sauce with peppers & spring onion.",
        "price": 6.99,
        "veg": true
      }
    ]
  },
  {
    "id": "fried-rice",
    "label": "Fried Rice",
    "num": "04",
    "note": "Fragrant basmati, wok-tossed to order.",
    "items": [
      {
        "id": "veg-fried-rice",
        "name": "Veg Fried Rice",
        "desc": "Fragrant rice wok-tossed with garden vegetables & oriental seasoning.",
        "price": 5.99,
        "veg": true
      },
      {
        "id": "chicken-fried-rice",
        "name": "Chicken Fried Rice",
        "desc": "Fragrant rice wok-tossed with tender chicken & fresh vegetables.",
        "price": 6.99,
        "veg": false
      },
      {
        "id": "cegg-fried-rice",
        "name": "C & Egg Fried Rice",
        "desc": "Wok-tossed rice with chicken, egg & aromatic oriental seasoning.",
        "price": 7.99,
        "veg": false
      },
      {
        "id": "egg-fried-rice",
        "name": "Egg Fried Rice",
        "desc": "Fluffy egg and fragrant rice wok-tossed with fresh vegetables.",
        "price": 5.99,
        "veg": true
      },
      {
        "id": "paneer-fried-rice",
        "name": "Paneer Fried Rice",
        "desc": "Fragrant rice tossed with golden paneer, vegetables & oriental sauces.",
        "price": 6.99,
        "veg": true
      },
      {
        "id": "veg-manch-fried-rice",
        "name": "Veg Manchurian Fried Rice",
        "desc": "Wok-tossed fried rice topped with crispy vegetable Manchurian.",
        "price": 6.49,
        "veg": true
      },
      {
        "id": "mushroom-rice",
        "name": "Mushroom Rice",
        "desc": "Fragrant rice tossed with juicy mushrooms, vegetables & aromatic seasoning.",
        "price": 5.99,
        "veg": true
      }
    ]
  },
  {
    "id": "noodles",
    "label": "Noodles",
    "num": "02",
    "note": "Hakka-style noodles, wok-tossed to order.",
    "items": [
      {
        "id": "veg-noodles",
        "name": "Veg Noodles",
        "desc": "Wok-tossed noodles with crunchy vegetables & oriental sauces.",
        "price": 5.99,
        "veg": true
      },
      {
        "id": "chicken-noodles",
        "name": "Chicken Noodles",
        "desc": "Wok-tossed noodles with tender chicken, vegetables & spicy house sauce.",
        "price": 6.99,
        "veg": false
      },
      {
        "id": "cegg-noodles",
        "name": "C & Egg Noodles",
        "desc": "Noodles tossed with chicken, egg, vegetables & oriental seasoning.",
        "price": 7.99,
        "veg": false
      },
      {
        "id": "egg-noodles",
        "name": "Egg Noodles",
        "desc": "Noodles tossed with fluffy egg, vegetables & aromatic sauces.",
        "price": 5.99,
        "veg": true
      },
      {
        "id": "paneer-noodles",
        "name": "Paneer Noodles",
        "desc": "Noodles tossed with paneer, crunchy vegetables & spicy Indo-Chinese sauce.",
        "price": 6.99,
        "veg": true
      },
      {
        "id": "veg-manch-noodles",
        "name": "Veg Manchurian Noodles",
        "desc": "Wok-tossed noodles topped with crispy vegetable Manchurian.",
        "price": 6.49,
        "veg": true
      },
      {
        "id": "mushroom-noodles",
        "name": "Mushroom Noodles",
        "desc": "Wok-tossed noodles with juicy mushrooms, vegetables & aromatic seasoning.",
        "price": 5.99,
        "veg": true
      },
      {
        "id": "veg-hakka-noodles",
        "name": "Veg Hakka Noodles",
        "desc": "Classic Hakka-style noodles tossed with crisp vegetables & soy seasoning.",
        "price": 5.99,
        "veg": true
      },
      {
        "id": "chicken-hakka-noodles",
        "name": "Chicken Hakka Noodles",
        "desc": "Classic Hakka-style noodles wok-tossed with chicken & soy seasoning.",
        "price": 6.99,
        "veg": false
      },
      {
        "id": "egg-hakka-noodles",
        "name": "Egg Hakka Noodles",
        "desc": "Classic Hakka-style noodles wok-tossed with fluffy egg & soy seasoning.",
        "price": 6.49,
        "veg": false
      }
    ]
  },
  {
    "id": "curries",
    "label": "Curries",
    "num": "05",
    "note": "Slow-cooked gravies, finished to order.",
    "subgroups": [
      {
        "label": "Veg Curry",
        "items": [
          {
            "id": "paneer-butter-masala",
            "name": "Paneer Butter Masala",
            "desc": "Soft paneer in a rich tomato-butter gravy finished with cream.",
            "price": 7.99,
            "veg": true
          },
          {
            "id": "kadai-paneer",
            "name": "Kadai Paneer",
            "desc": "Paneer cooked with peppers, onion & freshly ground kadai spices.",
            "price": 7.99,
            "veg": true
          },
          {
            "id": "sev-tamatar",
            "name": "Sev Tamatar",
            "desc": "Tangy tomato gravy topped with crispy sev & aromatic Indian spices.",
            "price": 6.99,
            "veg": true
          },
          {
            "id": "bhindi-masala",
            "name": "Bhindi Masala",
            "desc": "Tender okra cooked with onion, tomato & traditional Indian spices.",
            "price": 6.99,
            "veg": true
          },
          {
            "id": "rajma",
            "name": "Rajma",
            "desc": "Slow-cooked kidney beans in a rich, homestyle onion-tomato gravy.",
            "price": 6.99,
            "veg": true
          },
          {
            "id": "amritsari-chole-bhature",
            "name": "Amritsari Chole Bhature",
            "desc": "Slow-cooked chickpeas in a bold Amritsari masala, served with fluffy fried bhature.",
            "price": 7.99,
            "veg": true
          }
        ]
      },
      {
        "label": "Non-Veg Curry",
        "items": [
          {
            "id": "chicken-bhuna",
            "name": "Chicken Bhuna",
            "desc": "Tender chicken simmered in a thick, richly spiced onion-tomato masala.",
            "price": 7.99,
            "veg": false
          },
          {
            "id": "chicken-butter-masala",
            "name": "Chicken Butter Masala",
            "desc": "Tender chicken in a rich buttery tomato gravy finished with cream.",
            "price": 7.99,
            "veg": false
          },
          {
            "id": "chicken-karai",
            "name": "Chicken Karai",
            "desc": "Tender chicken cooked with peppers, tomato, onion & bold kadai spices.",
            "price": 7.99,
            "veg": false
          },
          {
            "id": "egg-kari",
            "name": "Egg Kari",
            "desc": "Boiled eggs simmered in a rich, aromatic onion-tomato curry.",
            "price": 6.99,
            "veg": false
          },
          {
            "id": "spice-special-egg-curry",
            "name": "Spice Special Egg Curry",
            "desc": "Our house-special egg curry, simmered in a rich, bold masala gravy.",
            "price": 7.49,
            "veg": false
          }
        ]
      }
    ]
  },
  {
    "id": "dal-rice",
    "label": "Dal & Rice",
    "num": "06",
    "note": "Comforting, homestyle & tempered with ghee.",
    "items": [
      {
        "id": "rajma-chawal",
        "name": "Rajma Chawal",
        "desc": "Comforting rajma served with fragrant steamed basmati rice.",
        "price": 7.99,
        "veg": true
      },
      {
        "id": "chole-chawal",
        "name": "Chole Chawal",
        "desc": "Spiced Amritsari chole served with fluffy basmati rice.",
        "price": 7.99,
        "veg": true
      },
      {
        "id": "jeera-dal-tadka",
        "name": "Jeera Rice & Dal Tadka",
        "desc": "Fragrant cumin rice served with creamy dal finished with a sizzling tadka.",
        "price": 6.99,
        "veg": true
      },
      {
        "id": "dal",
        "name": "Dal Tadka",
        "desc": "Homestyle lentils slow-cooked with aromatic spices & a traditional tadka.",
        "price": 5.99,
        "veg": true
      },
      {
        "id": "plain-rice",
        "name": "Plain Rice",
        "desc": "Fluffy steamed basmati rice, perfect with any curry.",
        "price": 4.99,
        "veg": true
      },
      {
        "id": "jeera-rice",
        "name": "Jeera Rice",
        "desc": "Fragrant basmati rice tempered with roasted cumin & ghee.",
        "price": 5.99,
        "veg": true
      }
    ]
  },
  {
    "id": "specials",
    "label": "Chef's Specials",
    "num": "07",
    "note": "Signature dishes you won't find on every menu.",
    "items": [
      {
        "id": "spice-special-paneer-masala",
        "name": "The Spice Special Paneer Masala",
        "desc": "Our house-special paneer curry, simmered in a rich, bold masala gravy.",
        "price": 7.99,
        "veg": true
      },
      {
        "id": "spice-special-chicken-masala",
        "name": "The Spice Special Chicken Masala",
        "desc": "Our house-special chicken curry, simmered in a rich, bold masala gravy.",
        "price": 7.99,
        "veg": false
      },
      {
        "id": "ghee-pepper-papadam-curry",
        "name": "Ghee Pepper Papadam Curry",
        "desc": "Crisp papadam simmered in a peppery ghee-roasted curry, finished with cracked pepper.",
        "price": 6.99,
        "veg": true
      },
      {
        "id": "cheese-anguri",
        "name": "Cheese Anguri (Butter Masala)",
        "desc": "Soft cheese dumplings in a rich, creamy butter masala gravy.",
        "price": 8.99,
        "veg": true
      }
    ]
  },
  {
    "id": "sides",
    "label": "Sides & Snacks",
    "num": "08",
    "note": "Perfect on the side, or on their own.",
    "items": [
      {
        "id": "french-fries",
        "name": "French Fries 🍟",
        "desc": "Golden, crispy fries, lightly salted.",
        "price": 2.99,
        "veg": true
      },
      {
        "id": "loaded-cheese-fries",
        "name": "Loaded Cheese French Fries",
        "desc": "Crispy fries loaded with melted cheese.",
        "price": 4.49,
        "veg": true
      },
      {
        "id": "peri-peri-chips",
        "name": "Peri Peri Chips",
        "desc": "Crispy chips tossed in a bold, tangy peri peri seasoning.",
        "price": 3.49,
        "veg": true
      },
      {
        "id": "papadam-2pc",
        "name": "Papadam (2pcs)",
        "desc": "Crisp, roasted papadam, served plain.",
        "price": 1.99,
        "veg": true
      },
      {
        "id": "chicken-samosa-3pc",
        "name": "Chicken Samosa (3 pcs)",
        "desc": "Crispy pastry parcels filled with a savoury spiced chicken filling.",
        "price": 3.49,
        "veg": false
      },
      {
        "id": "veg-samosa-3pc",
        "name": "Veg Samosa (3 pcs)",
        "desc": "Crispy pastry parcels filled with a savoury spiced vegetable filling.",
        "price": 3.49,
        "veg": true
      }
    ]
  },
  {
    "id": "breads",
    "label": "Breads",
    "num": "03",
    "note": "Fresh from the tandoor, to go with any curry.",
    "items": [
      {
        "id": "butter-naan",
        "name": "Butter Naan",
        "desc": "Soft, fluffy tandoor-baked naan brushed generously with melted butter.",
        "price": 1.99,
        "veg": true
      },
      {
        "id": "bhature",
        "name": "Bhature (1pc)",
        "desc": "Soft, fluffy deep-fried bread, the classic partner to chole.",
        "price": 1.99,
        "veg": true
      }
    ]
  },
  {
    "id": "desserts",
    "label": "Sweets & Lassi",
    "num": "09",
    "note": "A sweet finish, chilled or warm.",
    "items": [
      {
        "id": "salted-lassi",
        "name": "Salted Lassi",
        "desc": "Chilled, churned yoghurt drink, lightly salted.",
        "price": 1.99,
        "veg": true
      },
      {
        "id": "sweet-lassi",
        "name": "Sweet Lassi",
        "desc": "Chilled, churned yoghurt drink, lightly sweetened.",
        "price": 2.99,
        "veg": true
      },
      {
        "id": "mango-lassi",
        "name": "Mango Lassi",
        "desc": "Chilled, churned yoghurt drink blended with sweet mango.",
        "price": 3.99,
        "veg": true
      }
    ]
  }
];

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === "menu") {
    return jsonOutput_({
      success: true,
      menu: getMenu_(),
      updatedAt: new Date().toISOString()
    });
  }

  return jsonOutput_({
    success: true,
    message: "The Spice Kitchen Apps Script is running"
  });
}

function handleMenuSave_(data) {
  assertMenuSecret_(data.secret);
  const menu = normalizeMenuForSheet_(data.menu);
  writeMenuToSheet_(menu);
  return jsonOutput_({ success: true, menu: getMenu_() });
}

function assertMenuSecret_(secret) {
  const expected = PropertiesService
    .getScriptProperties()
    .getProperty("MENU_API_SECRET");

  if (!expected) {
    throw new Error('Set Script property "MENU_API_SECRET" before saving menu changes.');
  }

  if (String(secret || "") !== String(expected)) {
    throw new Error("Invalid menu API secret.");
  }
}

function setupMenuSheet() {
  writeMenuToSheet_(MENU_SEED);
}

function getMenu_() {
  const sh = getMenuSheet_();
  ensureMenuHeaders_(sh);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const rows = sh.getRange(2, 1, lastRow - 1, MENU_HEADERS.length).getValues();
  const categories = [];
  const categoryMap = {};

  rows.forEach(function (row) {
    const categoryId = String(row[0] || "").trim();
    const itemId = String(row[5] || "").trim();
    if (!categoryId || !itemId) return;

    let category = categoryMap[categoryId];
    if (!category) {
      category = {
        id: categoryId,
        label: String(row[1] || "").trim(),
        num: String(row[2] || "").trim(),
        note: String(row[3] || "").trim()
      };
      categoryMap[categoryId] = category;
      categories.push(category);
    }

    const item = {
      id: itemId,
      name: String(row[6] || "").trim(),
      desc: String(row[7] || "").trim(),
      price: Number(row[8]) || 0,
      veg: row[9] === true || String(row[9]).toLowerCase() === "true" || String(row[9]).toLowerCase() === "yes"
    };

    const soldOut = row[10] === true || String(row[10]).toLowerCase() === "true" || String(row[10]).toLowerCase() === "yes";
    if (soldOut) item.soldOut = true;

    const groupLabel = String(row[4] || "").trim();
    if (groupLabel) {
      if (!category.subgroups) category.subgroups = [];
      let group = category.subgroups.find(function (g) { return g.label === groupLabel; });
      if (!group) {
        group = { label: groupLabel, items: [] };
        category.subgroups.push(group);
      }
      group.items.push(item);
    } else {
      if (!category.items) category.items = [];
      category.items.push(item);
    }
  });

  return categories;
}

function writeMenuToSheet_(menu) {
  const normalized = normalizeMenuForSheet_(menu);
  const sh = getMenuSheet_();
  sh.clear();
  sh.getRange(1, 1, 1, MENU_HEADERS.length).setValues([MENU_HEADERS]).setFontWeight("bold");
  sh.setFrozenRows(1);

  const rows = [];
  normalized.forEach(function (category) {
    let sort = 1;

    (category.items || []).forEach(function (item) {
      rows.push(menuRow_(category, "", item, sort++));
    });

    (category.subgroups || []).forEach(function (group) {
      (group.items || []).forEach(function (item) {
        rows.push(menuRow_(category, group.label, item, sort++));
      });
    });
  });

  if (rows.length) {
    sh.getRange(2, 1, rows.length, MENU_HEADERS.length).setValues(rows);
  }

  sh.autoResizeColumns(1, MENU_HEADERS.length);
}

function menuRow_(category, groupLabel, item, sort) {
  return [
    category.id,
    category.label,
    category.num,
    category.note || "",
    groupLabel || "",
    item.id,
    item.name,
    item.desc || "",
    Number(item.price) || 0,
    Boolean(item.veg),
    Boolean(item.soldOut),
    sort
  ];
}

function normalizeMenuForSheet_(menu) {
  if (!Array.isArray(menu)) {
    throw new Error("Menu must be an array.");
  }

  return menu.map(function (category, index) {
    const next = {
      id: slugifyMenu_(category.id || category.label || "category-" + (index + 1)),
      label: String(category.label || "Category").trim(),
      num: String(category.num || String(index + 1).padStart(2, "0")).trim(),
      note: String(category.note || "").trim()
    };

    if (Array.isArray(category.items)) {
      next.items = category.items.map(normalizeMenuItem_);
    }

    if (Array.isArray(category.subgroups)) {
      next.subgroups = category.subgroups.map(function (group) {
        return {
          label: String(group.label || "Group").trim(),
          items: Array.isArray(group.items) ? group.items.map(normalizeMenuItem_) : []
        };
      });
    }

    if (!next.items && !next.subgroups) next.items = [];
    return next;
  });
}

function normalizeMenuItem_(item) {
  const name = String(item.name || "New item").trim();
  const price = Number(item.price);
  const next = {
    id: slugifyMenu_(item.id || name),
    name: name,
    desc: String(item.desc || "").trim(),
    price: !isNaN(price) && price >= 0 ? Math.round(price * 100) / 100 : 0,
    veg: Boolean(item.veg)
  };

  if (item.soldOut) next.soldOut = true;
  return next;
}

function slugifyMenu_(value) {
  const slug = String(value || "item")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 60);

  return slug || "item";
}

function getMenuSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(MENU_TAB_NAME) || ss.insertSheet(MENU_TAB_NAME);
}

function ensureMenuHeaders_(sh) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, MENU_HEADERS.length).setValues([MENU_HEADERS]).setFontWeight("bold");
    sh.setFrozenRows(1);
    return;
  }

  const firstCell = String(sh.getRange(1, 1).getValue() || "");
  if (firstCell !== MENU_HEADERS[0]) {
    throw new Error('The "Menu" sheet exists but does not use the expected menu headers.');
  }
}

function jsonOutput_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}


/* =========================================================
   TRACK TAB
   ONE ROW PER VISITOR
   ========================================================= */

const TRACK_TAB_NAME = "Track";

const TRACK_HEADERS = [
  "Visitor ID",
  "First Seen",
  "Last Seen",
  "Visit #",
  "New/Returning",

  "IP",
  "City",
  "Region",
  "Country",
  "ISP",

  "Device",
  "Device Model",
  "OS",
  "OS Version",
  "Browser",
  "Browser Version",
  "In-App Browser",

  "Screen",
  "Viewport",
  "Pixel Ratio",
  "Language",
  "Timezone",
  "Connection",
  "Colour Mode",

  "Referrer",
  "Source",
  "Medium",
  "Campaign",
  "Term",
  "Content",
  "Click ID",
  "First Source",
  "Landing URL",

  "Sessions",
  "Events Count",
  "Events",
  "Labels",
  "Sections Viewed",

  "Max Scroll %",
  "Total Time (sec)",
  "Cart Items",
  "Last Event"
];


/* =========================================================
   MAIN TRACK HANDLER
   ========================================================= */

function handleTrack_(data) {

  const events = Array.isArray(data.events)
    ? data.events.slice(0, 50)
    : [];

  const c = data.ctx || {};

  // Visitor ID is required
  if (!c.vid || !events.length) {
    return ContentService.createTextOutput("ignored");
  }

  const lock = LockService.getScriptLock();

  // Prevent two simultaneous requests creating two rows
  lock.waitLock(15000);

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let sh = ss.getSheetByName(TRACK_TAB_NAME);

    if (!sh) {
      sh = ss.insertSheet(TRACK_TAB_NAME);
    }

    ensureTrackHeaders_(sh);

    const now = new Date();

    /*
     * -------------------------------------------------------
     * Find existing visitor
     * -------------------------------------------------------
     */

    const visitorRow = findVisitorRow_(sh, c.vid);

    /*
     * -------------------------------------------------------
     * Build visitor information
     * -------------------------------------------------------
     */

    const firstSeen = visitorRow
      ? sh.getRange(visitorRow, 2).getValue() || now
      : now;

    const previousVisitCount = visitorRow
      ? Number(sh.getRange(visitorRow, 34).getValue()) || 0
      : 0;

    const previousEventCount = visitorRow
      ? Number(sh.getRange(visitorRow, 35).getValue()) || 0
      : 0;

    const previousEvents = visitorRow
      ? String(sh.getRange(visitorRow, 36).getValue() || "")
      : "";

    const previousLabels = visitorRow
      ? String(sh.getRange(visitorRow, 37).getValue() || "")
      : "";

    const previousSections = visitorRow
      ? String(sh.getRange(visitorRow, 38).getValue() || "")
      : "";

    const previousMaxScroll = visitorRow
      ? Number(sh.getRange(visitorRow, 39).getValue()) || 0
      : 0;

    const previousTotalTime = visitorRow
      ? Number(sh.getRange(visitorRow, 40).getValue()) || 0
      : 0;

    const previousCartItems = visitorRow
      ? Number(sh.getRange(visitorRow, 41).getValue()) || 0
      : 0;


    /*
     * -------------------------------------------------------
     * Aggregate events
     * -------------------------------------------------------
     */

    const eventNames = splitList_(previousEvents);
    const labels = splitList_(previousLabels);
    const sections = splitList_(previousSections);

    let maxScroll = previousMaxScroll;
    let totalTime = previousTotalTime;
    let cartItems = previousCartItems;

    let lastEvent = "";

    events.forEach(function (ev) {

      const eventName = safe_(ev.e);
      const label = safe_(ev.l);

      if (eventName) {
        addUnique_(eventNames, eventName);
        lastEvent = eventName;
      }

      if (label) {
        addUnique_(labels, label);
      }

      /*
       * Section tracking
       *
       * Your current frontend sends events such as:
       *
       * section_view + categories
       * section_view + popular
       * section_view + menu
       */
      if (eventName === "section_view" && label) {
        addUnique_(sections, label);
      }

      /*
       * Read engagement details
       *
       * Example:
       * {
       *   totalSec: 125,
       *   maxScroll: 75,
       *   cartItems: 2
       * }
       */
      const details = parseDetails_(ev.d);

      if (details) {

        if (details.maxScroll !== undefined) {
          const scroll = Number(details.maxScroll);

          if (!isNaN(scroll)) {
            maxScroll = Math.max(maxScroll, scroll);
          }
        }

        if (details.totalSec !== undefined) {
          const sec = Number(details.totalSec);

          if (!isNaN(sec)) {
            /*
             * Use the largest engagement time rather than
             * adding the same timer multiple times.
             */
            totalTime = Math.max(totalTime, sec);
          }
        }

        if (details.cartItems !== undefined) {
          const cart = Number(details.cartItems);

          if (!isNaN(cart)) {
            cartItems = Math.max(cartItems, cart);
          }
        }
      }
    });


    /*
     * -------------------------------------------------------
     * Sessions
     *
     * Your frontend already keeps a visit/session number.
     * We use the highest visit number received.
     * -------------------------------------------------------
     */

    const currentVisit = Number(c.visit) || 1;

    const sessions = Math.max(
      previousVisitCount,
      currentVisit
    );


    /*
     * -------------------------------------------------------
     * Create/update row
     * -------------------------------------------------------
     */

    const row = [

      // 1
      safe_(c.vid),

      // 2
      firstSeen,

      // 3
      now,

      // 4
      currentVisit,

      // 5
      safe_(c.returning),

      // 6 - 10
      safe_(c.ip),
      safe_(c.city),
      safe_(c.region),
      safe_(c.country),
      safe_(c.isp),

      // 11 - 17
      safe_(c.device),
      safe_(c.model),
      safe_(c.os),
      safe_(c.osVer),
      safe_(c.browser),
      safe_(c.bver),
      safe_(c.inApp),

      // 18 - 24
      safe_(c.screen),
      safe_(c.viewport),
      safe_(c.dpr),
      safe_(c.lang),
      safe_(c.tz),
      safe_(c.conn),
      safe_(c.dark),

      // 25 - 33
      safe_(c.ref),
      safe_(c.source),
      safe_(c.medium),
      safe_(c.campaign),
      safe_(c.term),
      safe_(c.content),
      safe_(c.clickId),
      safe_(c.firstSource),
      safe_(c.landing),

      // 34
      sessions,

      // 35
      previousEventCount + events.length,

      // 36
      eventNames.join(", "),

      // 37
      labels.join(", "),

      // 38
      sections.join(", "),

      // 39
      maxScroll,

      // 40
      totalTime,

      // 41
      cartItems,

      // 42
      lastEvent
    ];


    /*
     * -------------------------------------------------------
     * UPDATE EXISTING VISITOR
     * -------------------------------------------------------
     */

    if (visitorRow) {

      sh.getRange(
        visitorRow,
        1,
        1,
        TRACK_HEADERS.length
      ).setValues([row]);

    }

    /*
     * -------------------------------------------------------
     * CREATE NEW VISITOR
     * -------------------------------------------------------
     */

    else {

      sh.getRange(
        sh.getLastRow() + 1,
        1,
        1,
        TRACK_HEADERS.length
      ).setValues([row]);

    }

  } finally {

    lock.releaseLock();

  }

  return ContentService.createTextOutput("ok");
}


/* =========================================================
   FIND VISITOR ROW
   ========================================================= */

function findVisitorRow_(sh, visitorId) {

  const lastRow = sh.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  /*
   * Visitor ID is column 1 in the new Track sheet.
   */

  const values = sh
    .getRange(2, 1, lastRow - 1, 1)
    .getValues();

  for (let i = 0; i < values.length; i++) {

    if (String(values[i][0]) === String(visitorId)) {
      return i + 2;
    }

  }

  return null;
}


/* =========================================================
   ENSURE HEADERS
   ========================================================= */

function ensureTrackHeaders_(sh) {

  if (sh.getLastRow() === 0) {

    sh
      .getRange(1, 1, 1, TRACK_HEADERS.length)
      .setValues([TRACK_HEADERS])
      .setFontWeight("bold");

    sh.setFrozenRows(1);

    return;
  }

  /*
   * If the sheet is empty except for an old header,
   * don't automatically mix the old and new structures.
   */

  const firstCell = sh.getRange(1, 1).getValue();

  if (String(firstCell) !== "Visitor ID") {

    /*
     * The old tracking format is event-based.
     *
     * We intentionally don't try to convert it automatically
     * because the old rows may contain multiple events for
     * the same visitor.
     *
     * Clear the old Track sheet once before starting the
     * new visitor-based tracking.
     */
    throw new Error(
      'The "Track" sheet is using the old tracking format. ' +
      'Please clear/delete the old Track sheet and let the script recreate it.'
    );
  }
}


/* =========================================================
   PARSE EVENT DETAILS
   ========================================================= */

function parseDetails_(value) {

  if (!value) {
    return null;
  }

  try {

    if (typeof value === "object") {
      return value;
    }

    return JSON.parse(value);

  } catch (e) {

    return null;

  }
}


/* =========================================================
   UNIQUE LIST HELPERS
   ========================================================= */

function splitList_(value) {

  if (!value) {
    return [];
  }

  return String(value)
    .split(",")
    .map(function (x) {
      return x.trim();
    })
    .filter(Boolean);

}


function addUnique_(array, value) {

  if (!value) {
    return;
  }

  if (array.indexOf(value) === -1) {
    array.push(value);
  }

}


/* =========================================================
   SAFE SHEET VALUE
   ========================================================= */

function safe_(v) {

  if (v === null || v === undefined) {
    return "";
  }

  v = String(v).slice(0, 500);

  /*
   * Prevent visitor supplied text from becoming
   * a Google Sheets formula.
   */

  return /^[=+\-@]/.test(v)
    ? "'" + v
    : v;
}