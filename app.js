// ---------------------------------------------------------------------
// CONFIG
// ---------------------------------------------------------------------
// Replace with your actual "submit a correction" form once built (Google
// Forms / MS Forms). If the form supports pre-filled fields, add the
// institution name as a query param matching your form's field ID -
// see https://support.google.com/docs/answer/160000 ("Get pre-filled link").
const CORRECTION_FORM_BASE_URL = "https://forms.gle/REPLACE_ME";

// ---------------------------------------------------------------------

const resultsEl = document.getElementById("results");
const searchEl = document.getElementById("search");
const metaLineEl = document.getElementById("meta-line");
const emptyStateEl = document.getElementById("empty-state");

let institutions = [];
let pubsById = new Map();
let fuse = null;

function escapeHtml(s) {
  return (s ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  })[c]);
}

function correctionLink(record) {
  const label = encodeURIComponent(`${record.name}${record.abbreviation ? " (" + record.abbreviation + ")" : ""}`);
  return `${CORRECTION_FORM_BASE_URL}?entry.institution=${label}`;
}

function renderPublication(pub) {
  const titleHtml = pub.link
    ? `<a href="${escapeHtml(pub.link)}" target="_blank" rel="noopener">${escapeHtml(pub.title)}</a>`
    : escapeHtml(pub.title);
  const venueBits = [pub.venue, pub.volume && `vol. ${pub.volume}`, pub.pages]
    .filter(Boolean)
    .join(", ");
  const cite = [pub.authors, pub.year].filter(Boolean).join(" — ") + (venueBits ? `. ${venueBits}` : "");
  const notes = pub.notes && pub.notes.length
    ? `<div class="pub-note">${pub.notes.map(escapeHtml).join(" · ")}</div>`
    : "";
  return `
    <div class="pub">
      <div class="pub-title">${titleHtml}</div>
      <div class="pub-cite">${escapeHtml(cite)}</div>
      ${notes}
    </div>`;
}

function renderCard(record) {
  const li = document.createElement("li");
  li.className = "record";
  const codeLabel = record.type === "institution" ? record.abbreviation : "country";
  const sub = record.type === "institution"
    ? [record.city, record.country].filter(Boolean).join(", ")
    : "General / not yet linked to a specific collection";

  li.innerHTML = `
    <button class="record-head" aria-expanded="false">
      <span class="disclosure">▸</span>
      <span class="tag-code ${record.type === "country" ? "country" : ""}">${escapeHtml(codeLabel)}</span>
      <span class="record-name">${escapeHtml(record.name)}</span>
      <span class="record-count">${record.publication_count} ref${record.publication_count === 1 ? "" : "s"}</span>
    </button>
    <div class="record-body">
      <div class="record-sub">${escapeHtml(sub)}</div>
      <div class="pub-list"></div>
      <a class="correction-link" href="${correctionLink(record)}" target="_blank" rel="noopener">Submit a correction for this entry</a>
    </div>
  `;

  const head = li.querySelector(".record-head");
  const body = li.querySelector(".record-body");
  const pubList = li.querySelector(".pub-list");
  let rendered = false;

  head.addEventListener("click", () => {
    const isOpen = li.classList.toggle("open");
    head.setAttribute("aria-expanded", String(isOpen));
    li.querySelector(".disclosure").textContent = isOpen ? "▾" : "▸";
    if (isOpen && !rendered) {
      const pubs = record.publications
        .map((id) => pubsById.get(id))
        .filter(Boolean)
        .sort((a, b) => (b.year || "0").localeCompare(a.year || "0"));
      pubList.innerHTML = pubs.map(renderPublication).join("") || "<p>No linked references.</p>";
      rendered = true;
    }
  });

  return li;
}

function render(list) {
  resultsEl.innerHTML = "";
  emptyStateEl.hidden = list.length > 0;
  const frag = document.createDocumentFragment();
  list.forEach((r) => frag.appendChild(renderCard(r)));
  resultsEl.appendChild(frag);
}

function runSearch(query) {
  if (!query.trim()) {
    render(institutions);
    return;
  }
  const hits = fuse.search(query).map((r) => r.item);
  render(hits);
}

async function init() {
  const [instRes, pubRes] = await Promise.all([
    fetch("data/institutions.json"),
    fetch("data/publications.json"),
  ]);
  institutions = await instRes.json();
  const pubs = await pubRes.json();
  pubsById = new Map(pubs.map((p) => [p.id, p]));

  fuse = new Fuse(institutions, {
    keys: [
      { name: "name", weight: 2 },
      { name: "abbreviation", weight: 2 },
      { name: "city", weight: 1.4 },
      { name: "country", weight: 1 },
      { name: "verbatim_tags", weight: 1.2 },
    ],
    threshold: 0.32,
    ignoreLocation: true,
  });

  const totalPubs = pubs.length;
  const linkedPubs = pubs.filter((p) => p.institutions.length || p.countries_only.length).length;
  metaLineEl.textContent =
    `${institutions.length} collections · ${linkedPubs} of ${totalPubs} references linked`;

  render(institutions);

  searchEl.addEventListener("input", (e) => runSearch(e.target.value));
}

init().catch((err) => {
  metaLineEl.textContent = "Could not load the directory data.";
  console.error(err);
});
