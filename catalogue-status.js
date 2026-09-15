const resultsEl = document.getElementById("results");
const searchEl = document.getElementById("search");
const metaLineEl = document.getElementById("meta-line");
const emptyStateEl = document.getElementById("empty-state");

let records = [];
let fuse = null;

function escapeHtml(s) {
  return (s ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  })[c]);
}

function statusClass(code) {
  if (code === "none") return "status-none";
  if (code === "online") return "status-online";
  return "status-ref"; // M23 / M25 / M23 & M25
}

function renderRecord(r) {
  const li = document.createElement("li");
  li.className = "record";
  const sub = [r.city, r.country].filter(Boolean).join(", ");
  const scatterNote = r.has_scattered_references
    ? `<div class="scatter-note">Some general references exist for this collection —
       <a href="index.html?q=${encodeURIComponent(r.main_bibliography_query)}">search the main bibliography</a>.</div>`
    : "";

  li.innerHTML = `
    <div class="record-head" style="cursor:default;">
      <span class="tag-code">${escapeHtml(r.abbreviation || "—")}</span>
      <span class="record-name">${escapeHtml(r.name)}</span>
      <span class="status-badge ${statusClass(r.status_code)}">${escapeHtml(r.status_label)}</span>
    </div>
    <div class="record-sub">${escapeHtml(sub)}</div>
    ${scatterNote}
  `;
  return li;
}

function render(list) {
  resultsEl.innerHTML = "";
  emptyStateEl.hidden = list.length > 0;
  const frag = document.createDocumentFragment();
  list.forEach((r) => frag.appendChild(renderRecord(r)));
  resultsEl.appendChild(frag);
}

function runSearch(query) {
  if (!query.trim()) {
    render(records);
    return;
  }
  render(fuse.search(query).map((r) => r.item));
}

async function init() {
  const res = await fetch("data/uncatalogued.json");
  records = await res.json();

  fuse = new Fuse(records, {
    keys: [
      { name: "name", weight: 2 },
      { name: "abbreviation", weight: 2 },
      { name: "city", weight: 1.4 },
      { name: "country", weight: 1 },
    ],
    threshold: 0.32,
    ignoreLocation: true,
  });

  metaLineEl.textContent = `${records.length} collections`;
  render(records);

  searchEl.addEventListener("input", (e) => runSearch(e.target.value));
}

init().catch((err) => {
  metaLineEl.textContent = "Could not load the directory data.";
  console.error(err);
});
