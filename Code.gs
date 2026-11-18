/**
 * SERVER Kendali Proyek TA PKT 2026 — Cold Insulation (PT. Bumimarga Kharismaperkasa)
 * Pasang: Google Sheets baru › Ekstensi › Apps Script › tempel file ini › ganti KODE_AKSES & ADMIN_KEY
 * › Terapkan › Deployment baru › Jenis: Aplikasi web › Jalankan sebagai: Saya › Akses: Siapa saja › Terapkan › salin URL /exec.
 * Design by Ariel Panda
 */
const KODE_AKSES = 'GANTI-KODE-PROYEK';   // dibagikan ke semua pengguna aplikasi (lewat tautan setup)
const ADMIN_KEY  = 'GANTI-KUNCI-ADMIN';   // HANYA Project Control — untuk membaca GPS, audit & status online
const FOLDER_FOTO = 'Foto TA PKT 2026';   // folder Google Drive untuk foto

function doGet() { return out_({ok: true, app: 'TA-PKT server', t: Date.now()}); }
function doPost(e) {
  let b; try { b = JSON.parse(e.postData.contents); } catch (x) { return out_({ok: false, error: 'Format salah'}); }
  if (b.kode !== KODE_AKSES) return out_({ok: false, error: 'Kode akses salah'});
  const admin = b.adminKey === ADMIN_KEY;
  const lock = LockService.getScriptLock();
  try {
    switch (b.action) {
      case 'ping': return out_({ok: true, admin, t: Date.now()});
      case 'push': lock.waitLock(25000); return out_(push_(b.items || []));
      case 'pull': return out_(pull_(Number(b.since) || 0));
      case 'gps_push': lock.waitLock(25000); return out_(gpsPush_(b.points || []));
      case 'gps_pull': if (!admin) return out_({ok: false, error: 'Kunci Admin salah'}); return out_(gpsPull_(b.date, Number(b.since) || 0));
      case 'audit_push': lock.waitLock(25000); return out_(auditPush_(b.rows || []));
      case 'audit_pull': if (!admin) return out_({ok: false, error: 'Kunci Admin salah'}); return out_(auditPull_());
      case 'presence': return out_(presence_(b));
      case 'presence_list': if (!admin) return out_({ok: false, error: 'Kunci Admin salah'}); return out_(presenceList_());
      case 'pos_now': return out_(posNow_());
      case 'upload': return out_(upload_(b.name, b.data, b.folder));
      case 'pin_set': lock.waitLock(25000); return out_(pinSet_(b.items || []));
      case 'pin_list': if (!admin) return out_({ok: false, error: 'Kunci Admin salah'}); return out_(pinList_());
      case 'stats': return out_(stats_());
      default: return out_({ok: false, error: 'Aksi tidak dikenal'});
    }
  } catch (err) { return out_({ok: false, error: String(err && err.message || err)}); }
  finally { try { lock.releaseLock(); } catch (x) {} }
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sh_(name, head) {
  const ss = SpreadsheetApp.getActive(); let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(head); s.setFrozenRows(1); s.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#0e2747').setFontColor('#ffffff'); }
  return s;
}
/* ---------- data ---------- */
function slim_(j) {
  if (j.length < 48000) return j;
  const o = JSON.parse(j); Object.keys(o).forEach(k => { if (typeof o[k] === 'string' && o[k].indexOf('data:') === 0) o[k] = ''; });
  return JSON.stringify(o).slice(0, 48000);
}
function push_(items) {
  const s = sh_('data', ['koleksi', 'id', 'diubah', 'json', 'server_ts']);
  const n = s.getLastRow(); const keys = n > 1 ? s.getRange(2, 1, n - 1, 3).getValues() : [];
  const idx = {}; keys.forEach((r, i) => idx[r[0] + '|' + r[1]] = {row: i + 2, u: Number(r[2]) || 0});
  const now = Date.now(); const add = []; let up = 0;
  items.forEach(it => {
    const k = it.c + '|' + it.id, j = slim_(String(it.j || '{}'));
    const ex = idx[k];
    if (ex) { if (Number(it.u) >= ex.u) { s.getRange(ex.row, 3, 1, 3).setValues([[Number(it.u), j, now]]); up++; } }
    else { add.push([it.c, it.id, Number(it.u), j, now]); idx[k] = {row: -1, u: Number(it.u)}; }
  });
  if (add.length) s.getRange(s.getLastRow() + 1, 1, add.length, 5).setValues(add);
  return {ok: true, added: add.length, updated: up};
}
function pull_(since) {
  const s = sh_('data', ['koleksi', 'id', 'diubah', 'json', 'server_ts']); const n = s.getLastRow(); const now = Date.now();
  if (n < 2) return {ok: true, items: [], now};
  const v = s.getRange(2, 1, n - 1, 5).getValues(); const items = [];
  v.forEach(r => { if (Number(r[4]) > since - 2000) items.push({c: r[0], id: r[1], j: r[3]}); });
  return {ok: true, items, now};
}
/* ---------- GPS ---------- */
function gpsPush_(pts) {
  if (!pts.length) return {ok: true, n: 0};
  const s = sh_('gps', ['waktu', 't', 'tanggal', 'trip', 'kendaraan', 'driver', 'lat', 'lng', 'akurasi_m', 'kecepatan_kmj']);
  const tz = Session.getScriptTimeZone();
  const rows = pts.map(p => [new Date(p.t), p.t, Utilities.formatDate(new Date(p.t), tz, 'yyyy-MM-dd'), p.trip, p.kend, p.drv, p.lat, p.lng, p.acc, p.spd]);
  s.getRange(s.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return {ok: true, n: rows.length};
}
function gpsPull_(date, since) {
  const s = sh_('gps', ['waktu', 't', 'tanggal', 'trip', 'kendaraan', 'driver', 'lat', 'lng', 'akurasi_m', 'kecepatan_kmj']); const n = s.getLastRow();
  if (n < 2) return {ok: true, points: []};
  const st = since ? Math.max(2, n - 1500) : 2; // mode live: cukup baca baris terbaru
  const v = s.getRange(st, 1, n - st + 1, 10).getValues(); const tz = Session.getScriptTimeZone(); const pts = [];
  v.forEach(r => { if (since && Number(r[1]) < since) return; const d = r[2] instanceof Date ? Utilities.formatDate(r[2], tz, 'yyyy-MM-dd') : String(r[2]); if (!date || d === date) pts.push({t: Number(r[1]), trip: r[3], kend: r[4], drv: r[5], lat: Number(r[6]), lng: Number(r[7]), acc: Number(r[8]), spd: Number(r[9])}); });
  return {ok: true, points: pts};
}
function posNow_() { // posisi TERAKHIR per kendaraan (3 jam terakhir) — untuk layar driver, tanpa riwayat rute
  const s = sh_('gps', ['waktu', 't', 'tanggal', 'trip', 'kendaraan', 'driver', 'lat', 'lng', 'akurasi_m', 'kecepatan_kmj']); const n = s.getLastRow();
  if (n < 2) return {ok: true, list: []};
  const st = Math.max(2, n - 3000); const v = s.getRange(st, 1, n - st + 1, 10).getValues(); const last = {}; const lim = Date.now() - 3 * 3600e3;
  v.forEach(r => { const t = Number(r[1]); if (t > lim && (!last[r[4]] || last[r[4]].t < t)) last[r[4]] = {kend: r[4], t, lat: Number(r[6]), lng: Number(r[7])}; });
  return {ok: true, list: Object.keys(last).map(k => last[k])};
}
/* ---------- audit & online ---------- */
function auditPush_(rows) {
  if (!rows.length) return {ok: true};
  const s = sh_('audit', ['waktu', 't', 'pengguna', 'nama', 'aksi', 'koleksi', 'id', 'sebelum', 'sesudah', 'perangkat']);
  const v = rows.map(r => [new Date(r.t), r.t, r.u, r.un, r.act, r.c, r.id, JSON.stringify(r.b || ''), JSON.stringify(r.a || ''), r.dev]);
  s.getRange(s.getLastRow() + 1, 1, v.length, 10).setValues(v); return {ok: true};
}
function auditPull_() {
  const s = sh_('audit', ['waktu', 't', 'pengguna', 'nama', 'aksi', 'koleksi', 'id', 'sebelum', 'sesudah', 'perangkat']); const n = s.getLastRow();
  if (n < 2) return {ok: true, rows: []};
  const st = Math.max(2, n - 3000); const v = s.getRange(st, 1, n - st + 1, 10).getValues();
  const p = x => { try { return JSON.parse(x) || null; } catch (e) { return null; } };
  return {ok: true, rows: v.map(r => ({t: Number(r[1]), u: r[2], un: r[3], act: r[4], c: r[5], id: r[6], b: p(r[7]), a: p(r[8]), dev: r[9]}))};
}
function presence_(b) {
  const pr = PropertiesService.getScriptProperties(); const k = 'pr_' + b.dev;
  pr.setProperty(k, JSON.stringify({u: b.u, n: b.n, r: b.r, m: b.m, dev: b.dev, t: Date.now()})); return {ok: true};
}
function presenceList_() {
  const all = PropertiesService.getScriptProperties().getProperties(); const now = Date.now(); const list = [];
  Object.keys(all).forEach(k => { if (k.indexOf('pr_') !== 0) return; try { const o = JSON.parse(all[k]); if (now - o.t < 10 * 60e3) list.push(o); } catch (e) {} });
  return {ok: true, list};
}
/* ---------- catatan PIN (hanya terbaca dengan ADMIN_KEY) ---------- */
function pinSet_(items) {
  const s = sh_('pin', ['uid', 'nama', 'pin', 'diubah', 't']); const n = s.getLastRow(); const ids = n > 1 ? s.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0])) : [];
  items.forEach(it => { const row = [it.uid, it.nama, "'" + String(it.pin), new Date(Number(it.t)), Number(it.t)]; const i = ids.indexOf(String(it.uid)); if (i >= 0) s.getRange(i + 2, 1, 1, 5).setValues([row]); else { s.appendRow(row); ids.push(String(it.uid)); } });
  return {ok: true};
}
function pinList_() {
  const s = sh_('pin', ['uid', 'nama', 'pin', 'diubah', 't']); const n = s.getLastRow(); if (n < 2) return {ok: true, list: []};
  return {ok: true, list: s.getRange(2, 1, n - 1, 5).getValues().map(r => ({uid: String(r[0]), nama: r[1], pin: ('000000' + String(r[2]).replace(/\D/g, '')).slice(-6), t: Number(r[4])}))};
}
/* ---------- foto ---------- */
function sub_(parent, nama) { const it = parent.getFoldersByName(nama); return it.hasNext() ? it.next() : parent.createFolder(nama); }
function upload_(name, data, folder) {
  const m = String(data || '').match(/^data:(.+?);base64,(.+)$/); if (!m) return {ok: false, error: 'Foto tidak valid'};
  const it = DriveApp.getFoldersByName(FOLDER_FOTO); let f = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_FOTO);
  String(folder || '').split('/').map(x => x.trim().replace(/[\\:*?"<>|]/g, '_')).filter(Boolean).slice(0, 3).forEach(x => { f = sub_(f, x); }); // mis. Dokumentasi Progres/2026-11-17
  const file = f.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], String(name || 'foto.jpg').replace(/[^\w.\-]/g, '_')));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {ok: true, url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200'};
}
function stats_() {
  const s = sh_('data', ['koleksi', 'id', 'diubah', 'json', 'server_ts']); const n = s.getLastRow(); const counts = {};
  if (n > 1) s.getRange(2, 1, n - 1, 1).getValues().forEach(r => counts[r[0]] = (counts[r[0]] || 0) + 1);
  const g = sh_('gps', ['waktu', 't', 'tanggal', 'trip', 'kendaraan', 'driver', 'lat', 'lng', 'akurasi_m', 'kecepatan_kmj']);
  return {ok: true, counts, gps: Math.max(0, g.getLastRow() - 1)};
}
