// Membaca folder Google Drive (tanpa API key bila folder dibagikan "Siapa saja yang memiliki link"),
// mengubah semua tab desainer menjadi data.json, lalu menyiapkan folder dist/ untuk GitHub Pages.
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const ROOT_DIR = path.join(__dirname, '..');
const cfg = JSON.parse(fs.readFileSync(path.join(ROOT_DIR, 'config.json'), 'utf8'));
const ROOT = process.env.ROOT_FOLDER_ID || cfg.rootFolderId;
const KEY = process.env.GOOGLE_API_KEY || '';
const START_MONTH = cfg.startMonth || '2026-08';
const DATA_FILE = path.join(ROOT_DIR, 'data.json');
const DIST = path.join(ROOT_DIR, 'dist');
const UA = { 'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36', 'Accept-Language': 'id,en;q=0.8' };

const FOLDER = 'application/vnd.google-apps.folder';
const GSHEET = 'application/vnd.google-apps.spreadsheet';
const XLSXM = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const warnings = [];

/* ---------- nama bulan dari judul folder/file ---------- */
const MNAMES = [['januari','jan','january'],['februari','feb','february'],['maret','mar','march'],['april','apr'],['mei','may'],['juni','jun','june'],['juli','jul','july'],['agustus','agu','ags','agt','aug','august'],['september','sept','sep'],['oktober','okt','oct','october'],['november','nov','nop'],['desember','des','dec','december']];
function monthFromTitle(title) {
  const t = String(title || '').toLowerCase();
  const words = t.split(/[^a-z]+/).filter(Boolean);
  let mi = -1;
  for (const w of words) { const i = MNAMES.findIndex(a => a.includes(w)); if (i >= 0) { mi = i; break; } }
  if (mi < 0) return null;
  const y = (/(20\d{2})/.exec(t) || [])[1] || String(new Date().getFullYear());
  return y + '-' + String(mi + 1).padStart(2, '0');
}

/* ---------- daftar isi folder ---------- */
function decodeHtml(s) {
  return s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&#39;|&#x27;/g, "'").replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n));
}
function parseEmbedded(html) {
  const out = [], re = /id="entry-([A-Za-z0-9_-]{10,})"/g, hits = [];
  let m;
  while ((m = re.exec(html))) hits.push({ id: m[1], pos: m.index });
  hits.forEach((h, i) => {
    const chunk = html.slice(h.pos, (hits[i + 1] || { pos: html.length }).pos);
    const t = /flip-entry-title">([^<]*)</.exec(chunk);
    if (!t) return;
    const href = decodeHtml((/href="([^"]+)"/.exec(chunk) || [])[1] || '');
    const ty = /\/type\/([^"'?&\s]+)/.exec(chunk);
    let mime = ty ? decodeURIComponent(ty[1]) : '';
    if (!mime && /\/folders\//.test(href)) mime = FOLDER;
    out.push({ id: h.id, title: decodeHtml(t[1]).trim(), mime, href });
  });
  return out;
}
async function listEmbedded(id) {
  const r = await fetch('https://drive.google.com/embeddedfolderview?id=' + encodeURIComponent(id), { headers: UA });
  if (!r.ok) throw new Error('HTTP ' + r.status);
  return parseEmbedded(await r.text());
}
async function listApi(id) {
  const out = []; let tok = '';
  do {
    const p = new URLSearchParams({
      q: "'" + id + "' in parents and trashed = false", pageSize: '200',
      fields: 'nextPageToken,files(id,name,mimeType)', supportsAllDrives: 'true', includeItemsFromAllDrives: 'true', key: KEY
    });
    if (tok) p.set('pageToken', tok);
    const r = await fetch('https://www.googleapis.com/drive/v3/files?' + p.toString());
    if (!r.ok) throw new Error('Drive API HTTP ' + r.status);
    const j = await r.json();
    (j.files || []).forEach(f => out.push({ id: f.id, title: f.name, mime: f.mimeType, href: '' }));
    tok = j.nextPageToken || '';
  } while (tok);
  return out;
}
const listFolder = id => (KEY ? listApi(id) : listEmbedded(id));
function kindOf(f) {
  if (/folder/.test(f.mime) || /\/folders\//.test(f.href)) return 'folder';
  if (f.mime === GSHEET || /spreadsheets\/d\//.test(f.href)) return 'gsheet';
  if (f.mime === XLSXM || /\.xlsx?$/i.test(f.title)) return 'xlsx';
  return null;
}
async function crawl() {
  const found = [];
  const queue = [{ id: ROOT, month: null, depth: 0 }];
  while (queue.length) {
    const cur = queue.shift();
    let items;
    try { items = await listFolder(cur.id); }
    catch (e) { warnings.push('Folder ' + cur.id.slice(0, 8) + ' gagal dibaca: ' + e.message); continue; }
    if (!items.length && cur.depth === 0) warnings.push('Folder utama kosong atau tidak terbaca (belum dibagikan ke umum, atau tampilan Google berubah).');
    for (const f of items) {
      const kind = kindOf(f);
      if (kind === 'folder') { if (cur.depth < 3) queue.push({ id: f.id, month: monthFromTitle(f.title) || cur.month, depth: cur.depth + 1 }); }
      else if (kind) {
        const month = cur.month || monthFromTitle(f.title);
        if (month) found.push({ month, id: f.id, title: f.title, kind });
      }
    }
  }
  return found;
}

/* ---------- unduh file ---------- */
async function download(src) {
  const gs = id => KEY
    ? 'https://www.googleapis.com/drive/v3/files/' + id + '/export?mimeType=' + encodeURIComponent(XLSXM) + '&key=' + KEY
    : 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx';
  const fx = id => KEY
    ? 'https://www.googleapis.com/drive/v3/files/' + id + '?alt=media&supportsAllDrives=true&key=' + KEY
    : 'https://drive.usercontent.google.com/download?id=' + id + '&export=download&confirm=t';
  const fx2 = id => 'https://drive.google.com/uc?export=download&id=' + id;
  const tries = src.kind === 'gsheet' ? [gs(src.id)]
    : src.kind === 'xlsx' ? [fx(src.id), fx2(src.id)]
    : [gs(src.id), fx(src.id), fx2(src.id)];
  let last = '';
  for (const u of tries) {
    try {
      const r = await fetch(u, { headers: UA, redirect: 'follow' });
      if (!r.ok) { last = 'HTTP ' + r.status; continue; }
      const buf = Buffer.from(await r.arrayBuffer());
      if (buf.length > 4 && buf[0] === 0x50 && buf[1] === 0x4b) return buf;
      last = 'bukan file Excel';
    } catch (e) { last = e.message; }
  }
  throw new Error(last || 'gagal diunduh');
}

/* ---------- baca tab desainer ---------- */
const pad2 = n => String(n).padStart(2, '0');
function cellStr(v) {
  if (v instanceof Date) {
    const d = new Date(v.getTime() + 12 * 36e5);
    return pad2(d.getUTCDate()) + '/' + pad2(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
  }
  return v == null ? '' : String(v);
}
function dateCell(v) {
  if (/^\d{5}(\.\d+)?$/.test(v)) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(parseFloat(v)) * 864e5);
    return pad2(d.getUTCDate()) + '/' + pad2(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
  }
  return v;
}
function rowsToRecords(rows) {
  const hi = rows.findIndex(r => r.some(c => String(c).trim().toLowerCase() === 'campaign name'));
  if (hi < 0) return null;
  const h = rows[hi].map(c => String(c).trim().toLowerCase());
  const col = (...n) => { for (const x of n) { const i = h.indexOf(x); if (i >= 0) return i; } return -1; };
  const ix = {
    name: col('campaign name'), brand: col('brand'), marcom: col('marcom'), status: col('status'),
    ticket: col('tikect number', 'ticket number'), turunan: col('turunan'),
    acc: col('ticket accepted'), fa: col('ticket fa'), dur: col('work duration')
  };
  const g = (r, k) => (ix[k] >= 0 ? String(r[ix[k]] == null ? '' : r[ix[k]]) : '');
  const out = [];
  for (let i = hi + 1; i < rows.length; i++) {
    const r = rows[i];
    const name = g(r, 'name').trim();
    if (!name || name.toLowerCase() === 'campaign name') continue;
    out.push({
      name, brand: g(r, 'brand').trim(), marcom: g(r, 'marcom').trim(), status: g(r, 'status').trim(),
      ticket: g(r, 'ticket').replace(/\s/g, ''), turunan: g(r, 'turunan').trim(),
      acc: dateCell(g(r, 'acc').trim()), fa: dateCell(g(r, 'fa').trim()), dur: g(r, 'dur').trim()
    });
  }
  return out;
}
function parseWorkbook(buf) {
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
  const tabs = {};
  wb.SheetNames.forEach(sn => {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: '' }).map(r => r.map(cellStr));
    const recs = rowsToRecords(rows);
    if (recs) tabs[sn.trim()] = recs;
  });
  return tabs;
}

/* ---------- utama ---------- */
(async () => {
  const prev = fs.existsSync(DATA_FILE) ? JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')) : { months: {} };
  let sources = [];
  try { sources = await crawl(); } catch (e) { warnings.push('Penelusuran folder gagal: ' + e.message); }
  const seen = new Set();
  sources = sources.filter(s => s.month >= START_MONTH && !seen.has(s.id) && seen.add(s.id));

  const fresh = {}, report = [];
  async function load(s) {
    try {
      const tabs = parseWorkbook(await download(s));
      const n = Object.keys(tabs).length;
      if (!n) throw new Error('tidak ada tab desainer');
      fresh[s.month] = Object.assign(fresh[s.month] || {}, tabs);
      report.push({ month: s.month, title: s.title, tabs: n, ok: true });
      console.log('OK  ', s.month, s.title, n + ' tab');
    } catch (e) {
      warnings.push(s.title + ' (' + s.month + '): ' + e.message);
      report.push({ month: s.month, title: s.title, ok: false, error: e.message });
      console.log('GAGAL', s.month, s.title, e.message);
    }
  }
  for (const s of sources) await load(s);
  // sumber cadangan (config.json) hanya dipakai untuk bulan yang belum berhasil terbaca dari folder
  for (const x of (cfg.extraSheets || [])) {
    if (x.month < START_MONTH || fresh[x.month] || seen.has(x.id)) continue;
    seen.add(x.id);
    await load({ month: x.month, id: x.id, title: x.title || x.id, kind: 'auto' });
  }

  // bulan yang gagal dibaca tetap memakai data lama
  const merged = Object.assign({}, prev.months || {});
  Object.keys(fresh).forEach(k => { merged[k] = Object.assign({}, merged[k] || {}, fresh[k]); });
  const keep = Object.keys(merged).filter(k => k >= START_MONTH).sort();
  const months = {}; keep.forEach(k => { months[k] = merged[k]; });

  const now = new Date().toISOString();
  const changed = JSON.stringify(months) !== JSON.stringify(prev.months || {});
  const persistent = { updated: changed ? now : (prev.updated || now), startMonth: START_MONTH, months };
  fs.writeFileSync(DATA_FILE, JSON.stringify(persistent));

  fs.rmSync(DIST, { recursive: true, force: true });
  fs.mkdirSync(DIST, { recursive: true });
  fs.copyFileSync(path.join(ROOT_DIR, 'index.html'), path.join(DIST, 'index.html'));
  fs.writeFileSync(path.join(DIST, 'data.json'), JSON.stringify(Object.assign({ checked: now, sources: report, warnings }, persistent)));
  console.log('Selesai.', Object.keys(months).length + ' bulan,', warnings.length + ' peringatan,', changed ? 'ada perubahan data' : 'tidak ada perubahan');
})().catch(e => { console.error(e); process.exit(1); });
