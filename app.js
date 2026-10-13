const SPREADSHEET_ID = "1HEefTFbxDTZSZ6nWfjJUHf4g8pLoOasV86y-jTOIpts";
const WORKBOOK_URL = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=xlsx`;

const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit?usp=sharing`;

const HTMLVIEW_URL = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/htmlview`;

// Column headers (case-insensitive) that get special treatment. Any other
// column in a tab is shown as a plain labeled detail, so new columns in the
// sheet appear automatically.
const TITLE_COLUMNS = ["name"];
const PRICE_COLUMNS = ["price", "rate", "cost"];
const PHONE_COLUMNS = ["phone", "number"];
const EMAIL_COLUMNS = ["email"];
const WEBSITE_COLUMNS = ["website", "site", "link"];
const NOTES_COLUMNS = ["notes", "note"];

const loadingEl = document.getElementById("loading");
const errorEl = document.getElementById("error");
const listEl = document.getElementById("vendor-list");

function normalizeHeader(value) {
  return String(value || "").trim().toLowerCase();
}

function formatPhone(value) {
  const digits = value.replace(/\D/g, "");
  const local = digits.length === 11 && digits.startsWith("1") ? digits.slice(1) : digits;
  if (local.length === 10) {
    return `${local.slice(0, 3)}-${local.slice(3, 6)}-${local.slice(6)}`;
  }
  return value;
}

function formatPrice(value) {
  const plain = value.replace(/[$,\s]/g, "");
  if (plain !== "" && !Number.isNaN(Number(plain))) {
    return `$${Number(plain).toLocaleString("en-US")}`;
  }
  return value;
}

function rowsToSheet(sheetName, rows) {
  const headers = (rows[0] || []).map((header) => String(header || "").trim());
  const items = rows.slice(1)
    .map((row) => headers
      .map((header, index) => ({ header, key: normalizeHeader(header), value: String(row[index] ?? "").trim() }))
      .filter((field) => field.header))
    .filter((fields) => fields.some((field) => field.value));

  const nameOf = (fields) => (fields.find((field) => TITLE_COLUMNS.includes(field.key)) || {}).value || "";
  items.sort((a, b) => nameOf(a).localeCompare(nameOf(b), "en", { sensitivity: "base" }));

  return { name: sheetName, items };
}

async function loadWorkbookSheets() {
  if (!window.XLSX) {
    throw new Error("The workbook parser did not load.");
  }

  const controller = new AbortController();
  const timeoutId = window.setTimeout(() => controller.abort(), 15000);
  let response;
  try {
    response = await fetch(WORKBOOK_URL, { signal: controller.signal });
  } finally {
    window.clearTimeout(timeoutId);
  }
  if (!response.ok) {
    throw new Error("Could not load the Google Sheet workbook.");
  }

  const workbook = window.XLSX.read(await response.arrayBuffer(), { type: "array" });
  return workbook.SheetNames.map((sheetName) => {
    const rows = window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: "",
      raw: false,
    });
    return rowsToSheet(sheetName, rows);
  });
}

function createLink(href, text) {
  const link = document.createElement("a");
  link.href = href;
  link.textContent = text;
  if (/^https?:/i.test(href)) {
    link.target = "_blank";
    link.rel = "noopener noreferrer";
  }
  return link;
}

function websiteHref(value) {
  const href = /^[a-z][a-z0-9+.-]*:/i.test(value) ? value : `https://${value}`;
  return /^https?:/i.test(href) ? href : null;
}

function createDetail(label, field) {
  const wrapper = document.createElement("div");
  wrapper.className = "detail";

  const labelEl = document.createElement("span");
  labelEl.className = "label";
  labelEl.textContent = label;

  const valueEl = document.createElement("span");
  valueEl.className = "value";
  const value = field.value;

  if (PHONE_COLUMNS.includes(field.key)) {
    const phone = formatPhone(value);
    valueEl.appendChild(createLink(`tel:${phone.replace(/[^\d+]/g, "")}`, phone));
  } else if (EMAIL_COLUMNS.includes(field.key)) {
    valueEl.appendChild(createLink(`mailto:${value}`, value));
  } else if (WEBSITE_COLUMNS.includes(field.key) && websiteHref(value)) {
    valueEl.appendChild(createLink(websiteHref(value), value.replace(/^https?:\/\//i, "").replace(/\/$/, "")));
  } else {
    valueEl.textContent = value;
  }

  wrapper.append(labelEl, valueEl);
  return wrapper;
}

function createCard(fields) {
  const card = document.createElement("article");
  card.className = "vendor-card";

  const titleField = fields.find((field) => TITLE_COLUMNS.includes(field.key) && field.value);
  const priceField = fields.find((field) => PRICE_COLUMNS.includes(field.key) && field.value);

  const header = document.createElement("div");
  header.className = "card-header";

  const title = document.createElement("h3");
  title.className = "provider";
  title.textContent = titleField ? titleField.value : "Unnamed";
  header.appendChild(title);

  if (priceField) {
    const price = document.createElement("span");
    price.className = "rate";
    price.textContent = formatPrice(priceField.value);
    header.appendChild(price);
  }

  const details = document.createElement("div");
  details.className = "details";

  // Blank fields are omitted; notes go last in a highlighted box.
  fields
    .filter((field) => field !== titleField && field !== priceField && field.value)
    .sort((a, b) => NOTES_COLUMNS.includes(a.key) - NOTES_COLUMNS.includes(b.key))
    .forEach((field) => {
      const detail = createDetail(field.header, field);
      if (NOTES_COLUMNS.includes(field.key)) {
        detail.classList.add("notes-detail");
      }
      details.appendChild(detail);
    });

  card.appendChild(header);
  if (details.children.length) {
    card.appendChild(details);
  }
  return card;
}

function createSection(sheet) {
  const section = document.createElement("section");
  section.className = "cat-section";
  section.dataset.category = sheet.name;

  const header = document.createElement("div");
  header.className = "cat-header";

  const title = document.createElement("h2");
  title.className = "cat-title";
  title.textContent = sheet.name;

  const count = document.createElement("span");
  count.className = "count";
  count.textContent = `${sheet.items.length} ${sheet.items.length === 1 ? "listing" : "listings"}`;

  const edit = document.createElement("a");
  edit.className = "section-edit";
  edit.href = SHEET_URL;
  edit.dataset.tab = sheet.name;
  edit.target = "_blank";
  edit.rel = "noopener noreferrer";
  edit.textContent = "Add or update";

  header.append(title, count, edit);
  section.appendChild(header);

  const cards = document.createElement("div");
  cards.className = "cards";
  if (sheet.items.length) {
    sheet.items.forEach((fields) => {
      const card = createCard(fields);
      card.dataset.search = fields.map((field) => field.value).join(" ").toLowerCase();
      cards.appendChild(card);
    });
  } else {
    const empty = document.createElement("p");
    empty.className = "empty";
    empty.textContent = "Nothing listed yet.";
    cards.appendChild(empty);
  }
  section.appendChild(cards);
  return section;
}

function setupFilters(sheets) {
  const chipsEl = document.getElementById("chips");
  const searchEl = document.getElementById("search");
  const noResultsEl = document.getElementById("no-results");
  let activeCategory = null;

  function apply() {
    const query = searchEl.value.trim().toLowerCase();
    let anyVisible = false;

    listEl.querySelectorAll(".cat-section").forEach((section) => {
      const inCategory = !activeCategory || section.dataset.category === activeCategory;
      const cards = section.querySelectorAll(".vendor-card");
      let matches = 0;
      cards.forEach((card) => {
        const match = !query || card.dataset.search.includes(query);
        card.classList.toggle("hidden", !match);
        if (match) {
          matches += 1;
        }
      });
      // While searching, hide categories with no hits (including empty ones).
      const visible = inCategory && (!query || matches > 0);
      section.classList.toggle("hidden", !visible);
      anyVisible = anyVisible || visible;
    });

    noResultsEl.classList.toggle("hidden", anyVisible);
  }

  const chips = [{ label: "All", value: null }, ...sheets.map((sheet) => ({ label: sheet.name, value: sheet.name }))];
  chips.forEach(({ label, value }) => {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "chip";
    chip.textContent = label;
    chip.setAttribute("aria-pressed", String(value === activeCategory));
    chip.addEventListener("click", () => {
      activeCategory = value;
      chipsEl.querySelectorAll(".chip").forEach((el) => el.setAttribute("aria-pressed", String(el === chip)));
      apply();
    });
    chipsEl.appendChild(chip);
  });

  searchEl.addEventListener("input", apply);
  document.getElementById("controls").classList.remove("hidden");
}

// Point each section's button at its own tab. The public htmlview page lists
// every tab with its gid; if it can't be read, buttons keep the main sheet link.
async function linkEditButtonsToTabs() {
  try {
    const response = await fetch(HTMLVIEW_URL);
    if (!response.ok) {
      return;
    }
    const html = await response.text();
    const gids = new Map();
    for (const match of html.matchAll(/items\.push\(\{name: "((?:[^"\\]|\\.)*)"[^}]*?gid: "(\d+)"/g)) {
      let name = match[1];
      try {
        name = JSON.parse(`"${name}"`);
      } catch (error) {
        // Keep the raw name; plain tab names have no escapes anyway.
      }
      gids.set(name, match[2]);
    }
    document.querySelectorAll(".section-edit").forEach((link) => {
      const gid = gids.get(link.dataset.tab);
      if (gid) {
        link.href = `${SHEET_URL}&gid=${gid}#gid=${gid}`;
      }
    });
  } catch (error) {
    console.warn("Could not look up tab links.", error);
  }
}

async function init() {
  try {
    const sheets = await loadWorkbookSheets();
    sheets.forEach((sheet) => listEl.appendChild(createSection(sheet)));
    setupFilters(sheets);
    linkEditButtonsToTabs();
    loadingEl.classList.add("hidden");
    listEl.classList.remove("hidden");
  } catch (error) {
    loadingEl.classList.add("hidden");
    errorEl.textContent = "Could not load the Google Sheet. Make sure the sheet is shared with \"anyone with the link\".";
    errorEl.classList.remove("hidden");
    console.error(error);
  }
}

init();
