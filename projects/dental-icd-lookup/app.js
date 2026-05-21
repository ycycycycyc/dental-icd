const DATA_URL = "./data/K00-K14_master_table_v3_semantic_en.tsv";

const SEARCH_FIELDS = [
  "diagnosis_code",
  "diagnosis_name_cn",
  "semantic_name_en",
  "who_icd_code",
  "who_name_en",
  "category_name_cn",
  "subcategory_name_cn",
  "category_name_en",
  "subcategory_name_en",
];

const RESULT_COLUMNS = [
  "diagnosis_code",
  "diagnosis_name_cn",
  "semantic_name_en",
  "who_icd_code",
  "who_name_en",
  "english_mapping_confidence",
  "english_mapping_type",
];

const DETAIL_FIELDS = [
  ["chapter", "Chapter"],
  ["section", "Section"],
  ["category_code", "Category code"],
  ["category_name_cn", "Category name CN"],
  ["category_name_en", "Category name EN"],
  ["subcategory_code", "Subcategory code"],
  ["subcategory_name_cn", "Subcategory name CN"],
  ["subcategory_name_en", "Subcategory name EN"],
  ["parent_code", "Parent code"],
  ["is_subtype", "Subtype"],
  ["structural_name_en", "Structural English"],
  ["semantic_source", "Semantic source"],
];

const state = {
  rows: [],
  filtered: [],
  expandedCode: null,
  query: "",
  confidence: "",
  type: "",
  subtypeOnly: false,
};

const elements = {
  status: document.querySelector("#datasetStatus"),
  search: document.querySelector("#searchInput"),
  confidence: document.querySelector("#confidenceFilter"),
  type: document.querySelector("#typeFilter"),
  subtypeOnly: document.querySelector("#subtypeOnly"),
  clear: document.querySelector("#clearButton"),
  count: document.querySelector("#resultCount"),
  label: document.querySelector("#resultLabel"),
  activeQuery: document.querySelector("#activeQuery"),
  body: document.querySelector("#resultsBody"),
};

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function parseTsv(text) {
  const lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").trimEnd().split("\n");
  const headers = lines.shift().split("\t");
  return lines
    .filter((line) => line.trim().length > 0)
    .map((line) => {
      const cells = line.split("\t");
      return Object.fromEntries(headers.map((header, index) => [header, cells[index] ?? ""]));
    });
}

function normalize(value) {
  return String(value ?? "").toLocaleLowerCase();
}

function tokenize(query) {
  return normalize(query)
    .split(/\s+/)
    .map((token) => token.trim())
    .filter(Boolean);
}

function matchesQuery(row, tokens) {
  if (!tokens.length) return true;
  const haystack = SEARCH_FIELDS.map((field) => normalize(row[field])).join(" ");
  return tokens.every((token) => haystack.includes(token));
}

function confidenceClass(value) {
  return normalize(value);
}

function highlight(value) {
  const safeValue = escapeHtml(value);
  const terms = tokenize(state.query);
  if (!terms.length) return safeValue;

  return terms.reduce((html, term) => {
    if (!term) return html;
    const pattern = new RegExp(`(${term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")})`, "gi");
    return html.replace(pattern, "<mark>$1</mark>");
  }, safeValue);
}

function populateTypeFilter(rows) {
  const types = [...new Set(rows.map((row) => row.english_mapping_type).filter(Boolean))].sort();
  elements.type.innerHTML = '<option value="">All</option>';
  for (const type of types) {
    const option = document.createElement("option");
    option.value = type;
    option.textContent = type;
    elements.type.append(option);
  }
}

function applyFilters() {
  const tokens = tokenize(state.query);
  state.filtered = state.rows.filter((row) => {
    if (!matchesQuery(row, tokens)) return false;
    if (state.confidence && row.english_mapping_confidence !== state.confidence) return false;
    if (state.type && row.english_mapping_type !== state.type) return false;
    if (state.subtypeOnly && row.is_subtype !== "TRUE") return false;
    return true;
  });
  render();
}

function renderTag(value, extraClass = "") {
  return `<span class="tag ${extraClass}">${escapeHtml(value)}</span>`;
}

function renderResultRow(row) {
  const isExpanded = state.expandedCode === row.diagnosis_code;
  const confidence = row.english_mapping_confidence;
  const cells = [
    `<span class="code">${highlight(row.diagnosis_code)}</span>`,
    highlight(row.diagnosis_name_cn),
    highlight(row.semantic_name_en),
    `<span class="code">${highlight(row.who_icd_code)}</span>`,
    highlight(row.who_name_en),
    renderTag(confidence, confidenceClass(confidence)),
    renderTag(row.english_mapping_type),
  ];

  const result = `
    <tr class="result-row">
      ${cells.map((cell) => `<td>${cell}</td>`).join("")}
      <td>
        <button
          class="details-button"
          type="button"
          data-code="${escapeHtml(row.diagnosis_code)}"
          aria-expanded="${isExpanded ? "true" : "false"}"
        >
          ${isExpanded ? "Hide" : "Show"}
        </button>
      </td>
    </tr>
  `;

  if (!isExpanded) return result;

  const details = DETAIL_FIELDS.map(([field, label]) => {
    return `
      <div class="detail-item">
        <dt>${escapeHtml(label)}</dt>
        <dd>${highlight(row[field])}</dd>
      </div>
    `;
  }).join("");

  return `
    ${result}
    <tr class="details-row">
      <td colspan="8">
        <dl class="details-panel">${details}</dl>
      </td>
    </tr>
  `;
}

function render() {
  const total = state.rows.length;
  const visible = state.filtered.length;
  elements.count.textContent = visible.toLocaleString();
  elements.label.textContent = visible === 1 ? "record" : "records";

  const queryParts = [];
  if (state.query) queryParts.push(`query: ${state.query}`);
  if (state.confidence) queryParts.push(`confidence: ${state.confidence}`);
  if (state.type) queryParts.push(`type: ${state.type}`);
  if (state.subtypeOnly) queryParts.push("subtypes only");
  elements.activeQuery.textContent = queryParts.length
    ? queryParts.join(" · ")
    : `${total.toLocaleString()} total records`;

  if (!visible) {
    elements.body.innerHTML = `
      <tr>
        <td colspan="8" class="empty-state">No matching records.</td>
      </tr>
    `;
    return;
  }

  elements.body.innerHTML = state.filtered.map(renderResultRow).join("");
}

function wireEvents() {
  elements.search.addEventListener("input", (event) => {
    state.query = event.target.value.trim();
    state.expandedCode = null;
    applyFilters();
  });

  elements.confidence.addEventListener("change", (event) => {
    state.confidence = event.target.value;
    state.expandedCode = null;
    applyFilters();
  });

  elements.type.addEventListener("change", (event) => {
    state.type = event.target.value;
    state.expandedCode = null;
    applyFilters();
  });

  elements.subtypeOnly.addEventListener("change", (event) => {
    state.subtypeOnly = event.target.checked;
    state.expandedCode = null;
    applyFilters();
  });

  elements.clear.addEventListener("click", () => {
    state.query = "";
    state.confidence = "";
    state.type = "";
    state.subtypeOnly = false;
    state.expandedCode = null;
    elements.search.value = "";
    elements.confidence.value = "";
    elements.type.value = "";
    elements.subtypeOnly.checked = false;
    applyFilters();
    elements.search.focus();
  });

  elements.body.addEventListener("click", (event) => {
    const button = event.target.closest(".details-button");
    if (!button) return;
    const code = button.dataset.code;
    state.expandedCode = state.expandedCode === code ? null : code;
    render();
  });
}

async function init() {
  wireEvents();
  try {
    const response = await fetch(DATA_URL);
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    const text = await response.text();
    state.rows = parseTsv(text);
    state.filtered = [...state.rows];
    populateTypeFilter(state.rows);
    elements.status.textContent = `${state.rows.length.toLocaleString()} records loaded`;
    render();
  } catch (error) {
    elements.status.textContent = "Dataset unavailable";
    elements.body.innerHTML = `
      <tr>
        <td colspan="8" class="empty-state">
          Could not load the TSV dataset. Serve this folder with a static web server for local testing.
        </td>
      </tr>
    `;
    console.error(error);
  }
}

init();
