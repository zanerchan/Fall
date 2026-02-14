// Footer year
document.addEventListener('DOMContentLoaded', () => {
  const y = document.getElementById('year');
  if (y) y.textContent = new Date().getFullYear();
});

// Scroll fade for hero image (Home and slim hero)
document.addEventListener('DOMContentLoaded', () => {
  const heroImg = document.querySelector('#hero .hero-img');
  if (!heroImg) return;

  const fadeDistance = 300; // px of scroll from 1 -> ~0 opacity
  const onScroll = () => {
    const y = window.scrollY || window.pageYOffset;
    const opacity = Math.max(0, 1 - (y / fadeDistance));
    heroImg.style.opacity = opacity.toFixed(3);
  };
  window.addEventListener('scroll', onScroll, { passive: true });
  onScroll();
});

// Demo: simulate "login" to reveal menus on Home
document.addEventListener('DOMContentLoaded', () => {
  const btn = document.getElementById('demoLogin');
  if (!btn) return;
  btn.addEventListener('click', () => {
    const body = document.body;
    const nowLogged = !body.classList.contains('logged-in');
    body.classList.toggle('logged-in', nowLogged);
    btn.setAttribute('aria-pressed', String(nowLogged));
    btn.textContent = nowLogged ? 'Hide Menus (simulate logout)' : 'Show Menus (simulate login)';
  });
});
/* Excel Finder & Viewer (static HTML, local only)
   - Finds Excel files via folder picker (File System Access API) or file input (fallback)
   - Lets you search filenames
   - Click any result to render its first sheet as an HTML table (SheetJS)
*/

const SUP_EXT = ['.xlsx', '.xls', '.xlsb', '.csv'];
const pickFolderBtn = document.getElementById('pickFolder');
const pickFilesInput = document.getElementById('pickFiles');
const loadDefaultBtn = document.getElementById('loadDefault');
const searchBox = document.getElementById('searchBox');
const resultsEl = document.getElementById('fileResults');
const countEl = document.getElementById('fileCount');
const infoEl = document.getElementById('excelInfo');
const tableWrap = document.getElementById('excelTableWrap');

let entries = []; // { name, path, handle? (FileSystemFileHandle), file? (File) }

/* ---------- Utils ---------- */
function extOf(name) {
  const dot = name.lastIndexOf('.');
  return dot >= 0 ? name.slice(dot).toLowerCase() : '';
}
function looksExcel(name) {
  return SUP_EXT.includes(extOf(name));
}
function escapeHTML(s) {
  return s.replace(/[&<>"']/g, m => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[m]));
}
function setCount(n) {
  countEl.textContent = n ? `${n} file${n===1?'':'s'} found` : 'No files found';
}
function renderList(list) {
  resultsEl.innerHTML = '';
  setCount(list.length);
  list.forEach((e, i) => {
    const li = document.createElement('li');
    li.innerHTML = `
      <div>
        <div><strong>${escapeHTML(e.name)}</strong></div>
        <div class="path">${escapeHTML(e.path || '')}</div>
      </div>
      <button class="open" data-index="${i}">Open</button>
    `;
    resultsEl.appendChild(li);
  });
}
function filterEntries(q) {
  q = q.trim().toLowerCase();
  if (!q) return entries;
  return entries.filter(e =>
    e.name.toLowerCase().includes(q) ||
    (e.path && e.path.toLowerCase().includes(q))
  );
}
function updateList() {
  const q = searchBox.value || '';
  renderList(filterEntries(q));
}

/* ---------- Reading & rendering with SheetJS ---------- */
function sheetToHTMLTable(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  if (!rows.length) return '<p>No data.</p>';

  let html = '<table class="spec-table"><thead><tr>';
  const header = rows[0];
  header.forEach(h => html += `<th>${escapeHTML(String(h ?? ''))}</th>`);
  html += '</tr></thead><tbody>';

  for (let r = 1; r < rows.length; r++) {
    html += '<tr>';
    (rows[r] || []).forEach(cell => html += `<td>${escapeHTML(String(cell ?? ''))}</td>`);
    html += '</tr>';
  }
  html += '</tbody></table>';
  return html;
}
function renderWorkbook(wb, label = '') {
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  tableWrap.innerHTML = sheetToHTMLTable(ws);
  infoEl.textContent = `Showing: ${label || sheetName} — ${new Date().toLocaleString()}`;
}
async function openFileAndRender(getFileFn, label) {
  try {
    const file = await getFileFn();
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(new Uint8Array(ab), { type: 'array' });
    renderWorkbook(wb, label || file.name);
  } catch (err) {
    alert('Could not open file. ' + (err?.message || err));
  }
}

/* ---------- Folder scanning (File System Access API) ---------- */
async function scanDirectory(dirHandle, basePath = '') {
  // Recursively walk the directory and collect supported files
  for await (const entry of dirHandle.values()) {
    const currPath = basePath ? `${basePath}/${entry.name}` : entry.name;
    if (entry.kind === 'file') {
      if (looksExcel(entry.name)) {
        entries.push({ name: entry.name, path: currPath, handle: entry });
      }
    } else if (entry.kind === 'directory') {
      await scanDirectory(entry, currPath);
    }
  }
}

/* ---------- Event wiring ---------- */
if (pickFolderBtn) {
  // Disable button if API not supported
  if (!('showDirectoryPicker' in window)) {
    pickFolderBtn.title = 'Your browser does not support folder picking (use the “Pick Files / Folder” button).';
  }

  pickFolderBtn.addEventListener('click', async () => {
    entries = [];
    resultsEl.innerHTML = '';
    setCount(0);
    tableWrap.innerHTML = '';
    infoEl.textContent = '';

    try {
      if (!('showDirectoryPicker' in window)) {
        alert('Folder picker not supported in this browser. Use the “Pick Files / Folder” button.');
        return;
      }
      const dir = await window.showDirectoryPicker({ mode: 'read' });
      await scanDirectory(dir);
      updateList();
    } catch (e) {
      if (e?.name !== 'AbortError') alert('Folder pick cancelled or failed.');
    }
  });
}

// Fallback: input type="file" (supports multiple and webkitdirectory)
pickFilesInput?.addEventListener('change', (ev) => {
  entries = [];
  resultsEl.innerHTML = '';
  setCount(0);
  tableWrap.innerHTML = '';
  infoEl.textContent = '';

  const files = Array.from(ev.target.files || []);
  // Build pseudo "path" from the webkitRelativePath if present
  files.forEach(f => {
    if (looksExcel(f.name)) {
      entries.push({
        name: f.name,
        path: f.webkitRelativePath || f.name,
        file: f
      });
    }
  });
  updateList();
});

// Click handler: open a result
resultsEl?.addEventListener('click', (ev) => {
  const btn = ev.target.closest('button.open');
  if (!btn) return;
  const idx = Number(btn.dataset.index);
  const item = filterEntries(searchBox.value || '')[idx];
  if (!item) return;

  if (item.handle) {
    openFileAndRender(() => item.handle.getFile(), item.path);
  } else if (item.file) {
    openFileAndRender(() => Promise.resolve(item.file), item.path);
  } else {
    alert('No file handle available.');
  }
});

// Search box
searchBox?.addEventListener('input', updateList);

// Optional: load a default workbook from /data/sample.xlsx (needs a local server for fetch)
loadDefaultBtn?.addEventListener('click', async () => {
  try {
    const res = await fetch('data/sample.xlsx');
    const ab = await res.arrayBuffer();
    const wb = XLSX.read(new Uint8Array(ab), { type: 'array' });
    renderWorkbook(wb, 'sample.xlsx');
  } catch (e) {
    alert('To use “Load Default Workbook”, open this site via a local server (e.g., VS Code Live Server).');
  }
});