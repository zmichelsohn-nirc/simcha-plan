const SPREADSHEET_ID = "1HEefTFbxDTZSZ6nWfjJUHf4g8pLoOasV86y-jTOIpts";
const WORKBOOK_URL = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=xlsx`;

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

function createChevron() {
  const chevron = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  chevron.setAttribute("class", "chevron");
  chevron.setAttribute("viewBox", "0 0 24 24");
  chevron.setAttribute("aria-hidden", "true");
  chevron.innerHTML = '<path d="M9 6l6 6-6 6" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>';
  return chevron;
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

function createSection(sheet, index) {
  const section = document.createElement("details");
  section.className = "cat-section";
  section.open = index === 0;

  const summary = document.createElement("summary");
  summary.className = "cat-summary";

  const title = document.createElement("h2");
  title.className = "cat-title";
  title.textContent = sheet.name;

  const count = document.createElement("span");
  count.className = "count";
  count.textContent = `${sheet.items.length} ${sheet.items.length === 1 ? "listing" : "listings"}`;

  summary.append(createChevron(), title, count);
  section.appendChild(summary);

  const cards = document.createElement("div");
  cards.className = "cards";
  if (sheet.items.length) {
    sheet.items.forEach((fields) => cards.appendChild(createCard(fields)));
  } else {
    const empty = document.createElement("p");
    empty.className = "empty";
    empty.textContent = "Nothing listed yet.";
    cards.appendChild(empty);
  }
  section.appendChild(cards);
  return section;
}

async function init() {
  try {
    const sheets = await loadWorkbookSheets();
    sheets.forEach((sheet, index) => listEl.appendChild(createSection(sheet, index)));
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
