// Karyamas Transkrip Nilai (Offline-first)
// Frontend: Bootstrap 5 + IndexedDB (idb) + XLSX + jsPDF
// Backend: Google Apps Script (see Code.gs)

// =====================
// CONFIG
// =====================
const GAS_URL_DEFAULT =
  "https://script.google.com/macros/s/AKfycbxA0ZVZYpGq4ePqFwHsFUGyOoVn0vwf4XyvdCtycPLxo05WI4bw0mURT10iMBnE1txm/exec";

function getGasUrl() {
  const u = (localStorage.getItem("GAS_URL") || "").trim();
  return u || GAS_URL_DEFAULT;
}

const { openDB } = window.idb;
if (!openDB) {
  throw new Error(
    "Library idb (openDB) tidak ditemukan. Pastikan idb UMD sudah dimuat sebelum app.js."
  );
}

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

// =====================
// STATE
// =====================
const state = {
  user: null, // {username, name, role, must_change}
  masters: { peserta: [], materi: [], pelatihan: [], bobot: [], predikat: [] },
  filters: {
    tahun: new Date().getFullYear(),
    jenis: "",
    nik: "",
    test: "",
    materi: "",
  },
};

// =====================
// UI: TOAST + BUSY + PROGRESS
// =====================
function toast(msg) {
  const el = $("#toastBody");
  if (!el) return alert(msg);
  el.textContent = msg;
  const t = new bootstrap.Toast($("#appToast"), { delay: 2500 });
  t.show();
}

let __busyCount = 0;

function setBtnBusy(btn, busy = true, textBusy = "Memproses…") {
  if (!btn) return;
  if (busy) {
    if (btn.dataset._busy === "1") return;
    btn.dataset._busy = "1";
    btn.dataset._origHtml = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = `
      <span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span>
      <span>${textBusy}</span>`;
  } else {
    if (btn.dataset._busy !== "1") return;
    btn.dataset._busy = "0";
    btn.disabled = false;
    btn.innerHTML = btn.dataset._origHtml || btn.innerHTML;
  }
}

function ensureProgressUI() {
  if ($("#progressWrapLogin")) return;
  if ($("#progressWrap")) return;

  const host = $("#viewContainer") || document.body;
  const el = document.createElement("div");
  el.id = "progressWrap";
  el.className = "d-none";
  el.innerHTML = `
    <div class="mt-2">
      <div class="d-flex align-items-center justify-content-between">
        <div class="small text-muted" id="progressText">Memproses…</div>
        <div class="small text-muted"><span id="progressNow">0</span>/<span id="progressTotal">0</span></div>
      </div>
      <div class="progress mt-1" style="height:10px;">
        <div class="progress-bar progress-bar-striped progress-bar-animated" id="progressBar" style="width:0%"></div>
      </div>
    </div>
  `;
  host.appendChild(el);
}

function _pIds() {
  if ($("#progressWrapLogin")) {
    return {
      wrap: "#progressWrapLogin",
      text: "#progressTextLogin",
      now: "#progressNowLogin",
      total: "#progressTotalLogin",
      bar: "#progressBarLogin",
    };
  }
  return {
    wrap: "#progressWrap",
    text: "#progressText",
    now: "#progressNow",
    total: "#progressTotal",
    bar: "#progressBar",
  };
}

function progressStart(total = 0, text = "Memproses…") {
  ensureProgressUI();
  const ids = _pIds();
  $(ids.wrap).classList.remove("d-none");
  $(ids.text).textContent = text;
  $(ids.total).textContent = String(total || 0);
  $(ids.now).textContent = "0";
  $(ids.bar).style.width = "0%";
}

function progressSet(now, total, text) {
  ensureProgressUI();
  const ids = _pIds();
  $(ids.wrap).classList.remove("d-none");
  if (typeof text === "string") $(ids.text).textContent = text;
  $(ids.now).textContent = String(now || 0);
  $(ids.total).textContent = String(total || 0);
  const pct = total ? Math.round((now / total) * 100) : 0;
  $(ids.bar).style.width = `${Math.max(0, Math.min(100, pct))}%`;
}

function progressDone(textDone = "Selesai.") {
  const ids = _pIds();
  const wrap = $(ids.wrap);
  if (!wrap) return;

  $(ids.text).textContent = textDone;
  $(ids.bar).classList.remove("progress-bar-animated");
  $(ids.bar).style.width = "100%";

  setTimeout(() => {
    wrap.classList.add("d-none");
    $(ids.bar).classList.add("progress-bar-animated");
  }, 700);
}

async function runBusy(
  btn,
  fn,
  { busyText = "Memproses…", progressText = null, total = 0 } = {}
) {
  try {
    __busyCount++;
    setBtnBusy(btn, true, busyText);
    if (progressText) progressStart(total, progressText);
    const out = await fn();
    if (progressText) progressDone("Selesai.");
    return out;
  } finally {
    __busyCount = Math.max(0, __busyCount - 1);
    setBtnBusy(btn, false);
  }
}

// =====================
// UTILS: STRING + DATE + ROW NORMALIZER
// =====================
function normStr(v) {
  return String(v ?? "").trim();
}
function normNik(v) {
  return normStr(v).replace(/\s+/g, "");
}

function uniq(arr, key) {
  const s = new Set();
  const out = [];
  for (const a of arr) {
    const v = (a[key] || "").toString().trim();
    if (v && !s.has(v)) {
      s.add(v);
      out.push(v);
    }
  }
  return out;
}

function _pad2(n) {
  return String(n).padStart(2, "0");
}

function toIsoDateTimeLocal(d) {
  const y = d.getFullYear();
  const m = _pad2(d.getMonth() + 1);
  const day = _pad2(d.getDate());
  const hh = _pad2(d.getHours());
  const mm = _pad2(d.getMinutes());
  const ss = _pad2(d.getSeconds());
  return `${y}-${m}-${day}T${hh}:${mm}:${ss}`;
}

function parseAnyDate(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v;

  // Excel serial
  if (typeof v === "number" && Number.isFinite(v)) {
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const d = new Date(ms);
    if (!Number.isNaN(d.getTime())) return d;
  }

  const s = String(v).trim();
  if (!s) return null;

  // ISO-ish
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const d = new Date(s);
    if (!Number.isNaN(d.getTime())) return d;
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) {
      const d2 = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
      if (!Number.isNaN(d2.getTime())) return d2;
    }
  }

  // DD/MM/YYYY (optional time)
  const m1 = s.match(
    /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/
  );
  if (m1) {
    const dd = Number(m1[1]);
    const mm = Number(m1[2]);
    const yy = Number(m1[3]);
    const hh = Number(m1[4] || 0);
    const mi = Number(m1[5] || 0);
    const ss = Number(m1[6] || 0);
    const d = new Date(yy, mm - 1, dd, hh, mi, ss);
    if (!Number.isNaN(d.getTime())) return d;
  }

  const d3 = new Date(s);
  if (!Number.isNaN(d3.getTime())) return d3;

  return null;
}

function fmtTanggalDisplay(tanggal) {
  const d = parseAnyDate(tanggal);
  if (!d) return "";
  return `${_pad2(d.getDate())}/${_pad2(d.getMonth() + 1)}/${d.getFullYear()}`;
}

function normalizeNilaiRow(input) {
  const r = { ...(input || {}) };

  r.jenis_pelatihan = normStr(r.jenis_pelatihan);
  r.test_type = normStr(r.test_type);
  r.nik = normNik(r.nik);
  r.nama = normStr(r.nama);
  r.materi_kode = normStr(r.materi_kode);
  r.materi_nama = normStr(r.materi_nama);

  if (typeof r.nilai !== "number") r.nilai = parseFloat(r.nilai);
  if (Number.isNaN(r.nilai)) r.nilai = null;

  const d = parseAnyDate(r.tanggal);
  if (d) {
    r.tanggal = toIsoDateTimeLocal(d);
    r.tanggal_ts = d.getTime();
  } else {
    r.tanggal = "";
    r.tanggal_ts = 0;
  }

  const y = parseInt(r.tahun, 10);
  if (Number.isNaN(y) || y < 1900 || y > 2100) {
    r.tahun = d ? d.getFullYear() : new Date().getFullYear();
  } else {
    r.tahun = y;
  }

  return r;
}

function materiKeyFromRow(r) {
  return normStr(r?.materi_kode) || normStr(r?.materi_nama);
}

// =====================
// TRANSKRIP HELPERS (NO DUPLICATE)
// =====================
function trxR1(n) {
  const x = typeof n === "number" ? n : parseFloat(n);
  if (!Number.isFinite(x)) return "";
  return (Math.round(x * 10) / 10).toFixed(1);
}

function trxFmtTanggalIndoLong(v) {
  const d = parseAnyDate(v);
  if (!d) return "";
  const bulan = [
    "Januari","Februari","Maret","April","Mei","Juni",
    "Juli","Agustus","September","Oktober","November","Desember",
  ];
  return `${String(d.getDate()).padStart(2, "0")} ${bulan[d.getMonth()]} ${d.getFullYear()}`;
}

function trxPickPredikatFromMaster(v) {
  const n = typeof v === "number" ? v : parseFloat(v);
  if (!Number.isFinite(n)) return "-";

  const rules = (state.masters.predikat || [])
    .map((x) => ({
      nama: normStr(x.nama),
      min: parseFloat(x.min),
      max: parseFloat(x.max),
    }))
    .filter((x) => x.nama && Number.isFinite(x.min) && Number.isFinite(x.max));

  if (rules.length) {
    const found = rules.find((r) => n >= r.min && n <= r.max);
    if (found) return found.nama;
  }

  if (n > 90) return "Sangat Memuaskan";
  if (n > 80) return "Memuaskan";
  if (n >= 76) return "Baik";
  if (n >= 70) return "Kurang";
  return "Sangat Kurang";
}

function trxGetLulusStatus(total) {
  const n = typeof total === "number" ? total : parseFloat(total);
  if (!Number.isFinite(n)) return { lulus: false, label: "TIDAK", badge: "text-bg-danger" };
  const ok = n >= 70;
  return ok
    ? { lulus: true, label: "LULUS", badge: "text-bg-success" }
    : { lulus: false, label: "TIDAK", badge: "text-bg-danger" };
}

function trxJenisDenganTahun(jenis, tahun) {
  const j = normStr(jenis);
  const y = String(normStr(tahun));
  if (!j) return y ? `Tahun ${y}` : "-";
  if (!y) return j;
  if (/tahun\s+\d{4}/i.test(j)) return j;
  return `${j} Tahun ${y}`;
}

function trxGetKategoriMateri(materiKode, materiNama) {
  const kode = normStr(materiKode);
  const nama = normStr(materiNama);

  let m = null;
  if (kode) m = (state.masters.materi || []).find((x) => normStr(x.kode) === kode) || null;
  if (!m && nama) m = (state.masters.materi || []).find((x) => normStr(x.nama) === nama) || null;

  const kat = normStr(m?.kategori);
  if (!kat) return "";
  const k = kat.toLowerCase();
  if (k.includes("man")) return "Managerial";
  if (k.includes("tek")) return "Teknis";
  if (k.includes("sup")) return "Support";
  if (k.includes("ojt")) return "OJT";
  if (k.includes("pres")) return "Presentasi";
  return kat;
}

function trxGetBobotByJenis(jenis) {
  const j = normStr(jenis);
  const b = (state.masters.bobot || []).find((x) => normStr(x.jenis_pelatihan) === j) || null;

  let obj = {
    managerial: parseFloat(b?.managerial) || 0,
    teknis: parseFloat(b?.teknis) || 0,
    support: parseFloat(b?.support) || 0,
    ojt: parseFloat(b?.ojt) || 0,
    presentasi: parseFloat(b?.presentasi) || 0,
  };

  const sum = obj.managerial + obj.teknis + obj.support + obj.ojt + obj.presentasi;
  if (sum > 0 && Math.abs(sum - 100) > 0.01) {
    const k = 100 / sum;
    obj = {
      managerial: obj.managerial * k,
      teknis: obj.teknis * k,
      support: obj.support * k,
      ojt: obj.ojt * k,
      presentasi: obj.presentasi * k,
    };
  }
  return obj;
}

function trxGetBobotPercentForKategori(bobotObj, kategori) {
  const k = normStr(kategori).toLowerCase();
  if (k === "managerial") return bobotObj.managerial;
  if (k === "teknis") return bobotObj.teknis;
  if (k === "support") return bobotObj.support;
  if (k === "ojt") return bobotObj.ojt;
  if (k === "presentasi") return bobotObj.presentasi;
  return 0;
}

function trxBuildAvgMap(finalRows) {
  const avgMap = new Map(); // key -> {total,n}
  for (const r of finalRows || []) {
    const jenis = normStr(r.jenis_pelatihan);
    const tahun = String(normStr(r.tahun));
    const mkey = materiKeyFromRow(r);
    if (!jenis || !tahun || !mkey) continue;

    const key = `${jenis}||${tahun}||${mkey}`;
    if (!avgMap.has(key)) avgMap.set(key, { total: 0, n: 0 });

    const v = typeof r.nilai === "number" ? r.nilai : parseFloat(r.nilai);
    if (Number.isFinite(v)) {
      const o = avgMap.get(key);
      o.total += v;
      o.n += 1;
    }
  }
  return avgMap;
}

function trxGetRerataKelasFromMap(avgMap, jenis, tahun, materiKey) {
  const key = `${normStr(jenis)}||${String(normStr(tahun))}||${normStr(materiKey)}`;
  const o = avgMap.get(key);
  if (!o || !o.n) return null;
  return o.total / o.n;
}

function trxBuildTranscriptData({ nik, finalRowsAll }) {
  const mine = (finalRowsAll || []).filter((r) => normNik(r.nik) === normNik(nik));
  if (!mine.length) return null;

  const peserta =
    (state.masters.peserta || []).find((p) => String(p.nik) === String(nik)) || {};

  const tahun = mine[0]?.tahun ?? peserta.tahun ?? state.filters.tahun ?? "";
  const jenisRaw = mine[0]?.jenis_pelatihan || peserta.jenis_pelatihan || "-";
  const jenisLabel = trxJenisDenganTahun(jenisRaw, tahun);

  const avgMap = trxBuildAvgMap(finalRowsAll);
  const bobotObj = trxGetBobotByJenis(jenisRaw);

  const listSorted = mine
    .slice()
    .sort((a, b) => String(a.materi_kode || "").localeCompare(String(b.materi_kode || "")));

  const lines = [];
  let totalWeighted = 0;

  for (const r of listSorted) {
    const mKey = materiKeyFromRow(r);
    const rk = trxGetRerataKelasFromMap(avgMap, jenisRaw, tahun, mKey);

    const poin = typeof r.nilai === "number" ? r.nilai : parseFloat(r.nilai);
    const poinOk = Number.isFinite(poin) ? poin : 0;

    const kategori = trxGetKategoriMateri(r.materi_kode, r.materi_nama);
    const bobotPct = trxGetBobotPercentForKategori(bobotObj, kategori);

    const nilaiWeighted = poinOk * (bobotPct / 100);
    if (Number.isFinite(nilaiWeighted)) totalWeighted += nilaiWeighted;

    lines.push({
      materi_kode: r.materi_kode || "",
      materi_nama: r.materi_nama || "",
      rerata_kelas: rk,
      poin: poinOk,
      nilai: nilaiWeighted,
    });
  }

  return { nik, peserta, tahun, jenisRaw, jenisLabel, lines, totalWeighted };
}

// =====================
// NETWORK BADGE
// =====================
function setNetBadge() {
  const online = navigator.onLine;
  const badge = $("#netBadge");
  if (!badge) return;
  badge.textContent = online ? "Online" : "Offline";
  badge.classList.toggle("badge-online", online);
  badge.classList.toggle("badge-offline", !online);
  badge.onclick = () => {
    try {
      toast("GAS: " + getGasUrl());
    } catch (e) {}
  };
  if ($("#btnSync")) $("#btnSync").classList.toggle("d-none", !online || !state.user);
}

window.addEventListener("online", () => {
  setNetBadge();
  syncQueue().catch(() => {});
});
window.addEventListener("offline", setNetBadge);

// =====================
// INDEXEDDB
// =====================
const dbPromise = openDB("karyamas_transkrip_db", 2, {
  upgrade(db, oldVersion, newVersion, tx) {
    if (!db.objectStoreNames.contains("masters")) {
      db.createObjectStore("masters", { keyPath: "key" });
    }

    if (!db.objectStoreNames.contains("nilai")) {
      const nilai = db.createObjectStore("nilai", { keyPath: "id" });
      try { nilai.createIndex("by_tahun", "tahun"); } catch (e) {}
      try { nilai.createIndex("by_nik", "nik"); } catch (e) {}
      try { nilai.createIndex("by_jenis", "jenis_pelatihan"); } catch (e) {}
    } else {
      const store = tx.objectStore("nilai");
      if (!store.indexNames.contains("by_tahun")) { try { store.createIndex("by_tahun", "tahun"); } catch (e) {} }
      if (!store.indexNames.contains("by_nik")) { try { store.createIndex("by_nik", "nik"); } catch (e) {} }
      if (!store.indexNames.contains("by_jenis")) { try { store.createIndex("by_jenis", "jenis_pelatihan"); } catch (e) {} }
    }

    if (!db.objectStoreNames.contains("users")) {
      db.createObjectStore("users", { keyPath: "username" });
    }
    if (!db.objectStoreNames.contains("queue")) {
      db.createObjectStore("queue", { keyPath: "qid" });
    }
    if (!db.objectStoreNames.contains("settings")) {
      db.createObjectStore("settings", { keyPath: "key" });
    }
  },
});

async function dbGet(store, key) { return (await dbPromise).get(store, key); }
async function dbPut(store, val) { return (await dbPromise).put(store, val); }
async function dbDel(store, key) { return (await dbPromise).delete(store, key); }
async function dbAll(store) { return (await dbPromise).getAll(store); }

async function loadMastersFromDB() {
  const rows = await dbAll("masters");
  const map = Object.fromEntries(rows.map((r) => [r.key, r.data]));
  state.masters.peserta = map.peserta || [];
  state.masters.materi = map.materi || [];
  state.masters.pelatihan = map.pelatihan || [];
  state.masters.bobot = map.bobot || [];
  state.masters.predikat = map.predikat || [];
}

async function saveMastersToDB() {
  await dbPut("masters", { key: "peserta", data: state.masters.peserta });
  await dbPut("masters", { key: "materi", data: state.masters.materi });
  await dbPut("masters", { key: "pelatihan", data: state.masters.pelatihan });
  await dbPut("masters", { key: "bobot", data: state.masters.bobot });
  await dbPut("masters", { key: "predikat", data: state.masters.predikat });
}

// =====================
// MIGRATION
// =====================
async function migrateNilaiTanggalIfNeeded() {
  const db = await dbPromise;
  const all = await db.getAll("nilai");
  if (!all.length) return;

  let changed = 0;
  for (const r of all) {
    const need =
      !("tanggal_ts" in r) ||
      (typeof r.tanggal === "string" && /^\d{1,2}\/\d{1,2}\/\d{4}/.test(r.tanggal));

    if (!need) continue;

    const fixed = normalizeNilaiRow(r);
    fixed.id = r.id;
    await db.put("nilai", fixed);
    changed++;
  }

  if (changed) console.log(`migrateNilaiTanggalIfNeeded: updated ${changed} rows`);
}

// =====================
// SESSION (Remember login 30 days)
// =====================
const SESSION_KEY = "session_user_v1";
const SESSION_DAYS = 30;

function nowIso() { return new Date().toISOString(); }
function addDaysIso(days) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  return d.toISOString();
}

async function saveSession(user) {
  const payload = {
    user: {
      username: user.username,
      role: user.role,
      name: user.name,
      must_change: !!user.must_change,
    },
    created_at: nowIso(),
    expires_at: addDaysIso(SESSION_DAYS),
  };
  await dbPut("settings", { key: SESSION_KEY, value: payload });
}

async function loadSession() {
  const row = await dbGet("settings", SESSION_KEY);
  const sess = row?.value;
  if (!sess || !sess.user || !sess.expires_at) return null;

  const exp = new Date(sess.expires_at).getTime();
  if (Number.isNaN(exp) || Date.now() > exp) {
    await clearSession();
    return null;
  }
  return sess;
}

async function clearSession() {
  await dbDel("settings", SESSION_KEY);
}

async function restoreSessionIntoState() {
  const sess = await loadSession();
  if (!sess) return false;
  state.user = sess.user;
  return true;
}

async function fixSessionMustChangeFromLocal() {
  if (!state.user?.username) return;
  try {
    const uLocal = await dbGet("users", state.user.username);
    if (!uLocal) return;

    if (state.user.must_change && uLocal.must_change === false) {
      state.user.must_change = false;
      await saveSession(state.user);
    }
  } catch (e) {
    console.warn("fixSessionMustChangeFromLocal failed:", e);
  }
}

// =====================
// CRYPTO (SHA-256)
// =====================
async function sha256Hex(str) {
  if (globalThis.crypto && crypto.subtle && globalThis.isSecureContext) {
    const enc = new TextEncoder().encode(str);
    const buf = await crypto.subtle.digest("SHA-256", enc);
    return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
  }

  // fallback JS SHA-256 ringkas
  function rrot(n, x) { return (x >>> n) | (x << (32 - n)); }
  function toHex(n) { return (n >>> 0).toString(16).padStart(8, "0"); }

  const msg = new TextEncoder().encode(str);
  const l = msg.length * 8;

  const with1 = new Uint8Array(((msg.length + 9 + 63) >> 6) << 6);
  with1.set(msg, 0);
  with1[msg.length] = 0x80;

  const dv = new DataView(with1.buffer);
  dv.setUint32(with1.length - 4, l >>> 0, false);
  dv.setUint32(with1.length - 8, Math.floor(l / 2 ** 32) >>> 0, false);

  const K = [
    0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,
    0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,
    0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,
    0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,
    0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,
    0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,
    0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,
    0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2,
  ];

  let h0=0x6a09e667,h1=0xbb67ae85,h2=0x3c6ef372,h3=0xa54ff53a,h4=0x510e527f,h5=0x9b05688c,h6=0x1f83d9ab,h7=0x5be0cd19;
  const W = new Uint32Array(64);

  for (let i = 0; i < with1.length; i += 64) {
    for (let t = 0; t < 16; t++) W[t] = dv.getUint32(i + t * 4, false);
    for (let t = 16; t < 64; t++) {
      const s0 = rrot(7, W[t - 15]) ^ rrot(18, W[t - 15]) ^ (W[t - 15] >>> 3);
      const s1 = rrot(17, W[t - 2]) ^ rrot(19, W[t - 2]) ^ (W[t - 2] >>> 10);
      W[t] = (W[t - 16] + s0 + W[t - 7] + s1) >>> 0;
    }

    let a=h0,b=h1,c=h2,d=h3,e=h4,f=h5,g=h6,h=h7;

    for (let t = 0; t < 64; t++) {
      const S1 = rrot(6, e) ^ rrot(11, e) ^ rrot(25, e);
      const ch = (e & f) ^ (~e & g);
      const temp1 = (h + S1 + ch + K[t] + W[t]) >>> 0;
      const S0 = rrot(2, a) ^ rrot(13, a) ^ rrot(22, a);
      const maj = (a & b) ^ (a & c) ^ (b & c);
      const temp2 = (S0 + maj) >>> 0;

      h=g; g=f; f=e; e=(d + temp1)>>>0;
      d=c; c=b; b=a; a=(temp1 + temp2)>>>0;
    }

    h0=(h0+a)>>>0; h1=(h1+b)>>>0; h2=(h2+c)>>>0; h3=(h3+d)>>>0;
    h4=(h4+e)>>>0; h5=(h5+f)>>>0; h6=(h6+g)>>>0; h7=(h7+h)>>>0;
  }

  return (toHex(h0)+toHex(h1)+toHex(h2)+toHex(h3)+toHex(h4)+toHex(h5)+toHex(h6)+toHex(h7));
}

// =====================
// GAS HELPERS (JSONP)
// =====================
async function gasCall(action, payload = {}) {
  const url = getGasUrl();
  if (!url || /PASTE_YOUR_GAS_WEBAPP_URL/i.test(url)) {
    throw new Error("GAS_URL belum diisi. Isi di Setting Admin atau hardcode default.");
  }
  return await gasJsonp(url, action, payload);
}

async function gasPing() {
  return await gasCall("ping", {});
}

function gasJsonp(GAS_URL, action, payload) {
  return new Promise((resolve, reject) => {
    const cb = "cb_" + Date.now() + "_" + Math.random().toString(16).slice(2);

    let script = null;
    let done = false;

    const timeout = setTimeout(() => {
      cleanup();
      reject(new Error("JSONP timeout"));
    }, 25000);

    window[cb] = (data) => {
      cleanup();
      if (!data || data.ok === false) return reject(new Error((data && data.error) || "GAS error"));
      resolve(data);
    };

    function cleanup() {
      if (done) return;
      done = true;

      clearTimeout(timeout);
      try { delete window[cb]; } catch (e) { window[cb] = undefined; }
      try { if (script && script.parentNode) script.parentNode.removeChild(script); } catch (e) {}
    }

    const params = new URLSearchParams();
    params.set("action", action);
    params.set("payload", JSON.stringify(payload || {}));
    params.set("callback", cb);
    params.set("_ts", String(Date.now()));

    const url = GAS_URL + (GAS_URL.includes("?") ? "&" : "?") + params.toString();

    script = document.createElement("script");
    script.src = url;
    script.async = true;

    script.onerror = () => {
      cleanup();
      reject(new Error("Gagal memuat GAS (JSONP). URL: " + GAS_URL + " | Pastikan WebApp publik & URL deployment terbaru."));
    };

    document.body.appendChild(script);
  });
}

// =====================
// AUTH + USERS
// =====================
async function ensureDefaultUsers() {
  const db = await dbPromise;
  const admin = await db.get("users", "admin");
  if (!admin) {
    await db.put("users", {
      username: "admin",
      role: "admin",
      name: "Administrator",
      pass_hash: await sha256Hex("123456"),
      must_change: false,
    });
  }
}

async function syncUsersFromMasterPeserta() {
  const db = await dbPromise;
  for (const p of state.masters.peserta) {
    const username = String(p.nik || "").trim();
    if (!username) continue;
    const u = await db.get("users", username);
    if (!u) {
      await db.put("users", {
        username,
        role: "user",
        name: p.nama || ("Peserta " + username),
        pass_hash: await sha256Hex(username),
        must_change: true,
      });
    }
  }
}

async function login(username, password) {
  username = String(username || "").trim();
  const passHash = await sha256Hex(password);
  const isDefault = username !== "admin" && passHash === (await sha256Hex(username));

  if (navigator.onLine) {
    try {
      const res = await gasCall("login", { username, pass_hash: passHash });
      await dbPut("users", {
        username,
        role: res.user.role,
        name: res.user.name,
        pass_hash: res.user.pass_hash,
        must_change: !!res.user.must_change,
      });

      state.user = {
        username,
        role: res.user.role,
        name: res.user.name,
        must_change: isDefault || !!res.user.must_change,
      };

      return { source: "gas", user: state.user };
    } catch (e) {
      console.warn("login gas failed, fallback local:", e);
    }
  }

  const uLocal = await dbGet("users", username);
  if (uLocal && uLocal.pass_hash === passHash) {
    state.user = {
      username,
      role: uLocal.role,
      name: uLocal.name,
      must_change: isDefault || !!uLocal.must_change,
    };
    return { source: "local", user: state.user };
  }

  throw new Error("Login gagal. Jika user baru, tarik master dulu atau pastikan online.");
}

async function changePassword(newPass) {
  const newHash = await sha256Hex(newPass);

  if (!state.user?.username) throw new Error("Belum login.");
  if (!navigator.onLine) throw new Error("Ganti password wajib Online (tidak memakai antrian).");

  const username = state.user.username;

  let currentLocal = await dbGet("users", username);
  if (!currentLocal) {
    const guessedHash = await sha256Hex(username === "admin" ? "123456" : username);
    currentLocal = {
      username,
      role: username === "admin" ? "admin" : "user",
      name: state.user.name || (username === "admin" ? "Administrator" : ("Peserta " + username)),
      pass_hash: guessedHash,
      must_change: username !== "admin",
    };
    await dbPut("users", currentLocal);
  }

  try {
    await gasCall("login", { username, pass_hash: currentLocal.pass_hash });
  } catch (e) {
    console.warn("ensure via login failed, try change_password anyway:", e);
  }

  await gasCall("change_password", { username, pass_hash: newHash });

  currentLocal.pass_hash = newHash;
  currentLocal.must_change = false;
  await dbPut("users", currentLocal);

  state.user.must_change = false;
  try { await saveSession(state.user); } catch (e) {}

  toast("Password berhasil diubah (server & lokal).");
}

async function purgePasswordQueue() {
  const db = await dbPromise;
  const items = await db.getAll("queue");
  for (const it of items) {
    if (it.type === "change_password" || it.type === "upsert_user") {
      await db.delete("queue", it.qid);
    }
  }
  await refreshQueueBadge();
}

// =====================
// QUEUE & SYNC
// =====================
function qid() {
  return "Q" + Date.now() + "_" + Math.random().toString(16).slice(2);
}

function ensureNilaiId(row) {
  if (!row) row = {};
  if (!row.id) row.id = "N" + Date.now() + "_" + Math.random().toString(16).slice(2);
  return row;
}

async function enqueue(type, payload) {
  await dbPut("queue", { qid: qid(), type, payload, created_at: new Date().toISOString() });
  await refreshQueueBadge();
}

async function refreshQueueBadge() {
  const q = await dbAll("queue");
  const c = $("#queueCount");
  if (c) c.textContent = String(q.length);
  const s = $("#syncState");
  if (s) s.textContent = navigator.onLine ? "Online" : "Offline";
}

async function syncQueue() {
  if (!navigator.onLine) return;
  const db = await dbPromise;
  const items = await db.getAll("queue");
  if (!items.length) return;

  const syncState = $("#syncState");
  if (syncState) syncState.textContent = "Syncing…";

  const total = items.length;
  progressStart(total, "Sinkronisasi antrian…");

  let done = 0;
  for (const it of items) {
    try {
      await gasCall(it.type, it.payload);
      await db.delete("queue", it.qid);
      done++;
      if (done % 5 === 0 || done === total) {
        progressSet(done, total, `Sync… (${done}/${total})`);
      }
    } catch (e) {
      console.warn("queue sync failed", it, e);
      progressDone("Sync berhenti (ada error).");
      break;
    }
  }

  await refreshQueueBadge();
  if (syncState) syncState.textContent = "Online";
  progressDone("Sync selesai.");
  toast("Sync selesai.");
}

// =====================
// MASTERS: OPTIONS + CACHE
// =====================
function optYears() {
  const now = new Date().getFullYear();
  const years = [];
  for (let y = now; y >= now - 10; y--) years.push(y);
  return years.map((y) => `<option value="${y}">${y}</option>`).join("");
}

async function rebuildPelatihanCache() {
  const fromPeserta = uniq(state.masters.peserta || [], "jenis_pelatihan");

  let fromNilai = [];
  try {
    const allNilai = await dbAll("nilai");
    fromNilai = uniq(allNilai || [], "jenis_pelatihan");
  } catch (e) {}

  const merged = [...new Set([...fromPeserta, ...fromNilai].map(normStr).filter(Boolean))].sort(
    (a, b) => a.localeCompare(b)
  );

  state.masters.pelatihan = merged.map((nama) => ({ nama }));
  await saveMastersToDB();
}

function optPelatihan() {
  let items = Array.isArray(state.masters.pelatihan) ? state.masters.pelatihan : [];
  items = items
    .map((x) => (typeof x === "string" ? { nama: x } : x))
    .filter((x) => normStr(x?.nama));

  if (!items.length) {
    const u = uniq(state.masters.peserta || [], "jenis_pelatihan");
    items = u.map((j) => ({ nama: j }));
  }

  return [`<option value="">Semua</option>`]
    .concat(items.map((p) => `<option value="${p.nama}">${p.nama}</option>`))
    .join("");
}

function optTestTypes() {
  const items = ["PreTest", "PostTest", "Final", "OJT", "Presentasi"];
  return [`<option value="">Semua</option>`]
    .concat(items.map((t) => `<option value="${t}">${t}</option>`))
    .join("");
}

function optMateriFromNilai(nilaiRows) {
  const map = new Map();
  for (const r of nilaiRows || []) {
    const key = materiKeyFromRow(r);
    if (!key) continue;
    const kode = normStr(r.materi_kode);
    const nama = normStr(r.materi_nama);
    const label = kode && nama ? `${kode} - ${nama}` : (nama || kode || key);
    if (!map.has(key)) map.set(key, label);
  }

  if (map.size === 0) {
    for (const m of state.masters.materi || []) {
      const key = normStr(m.kode) || normStr(m.nama);
      if (!key) continue;
      const label = m.kode && m.nama ? `${m.kode} - ${m.nama}` : (m.nama || m.kode);
      if (!map.has(key)) map.set(key, label);
    }
  }

  const arr = [...map.entries()]
    .map(([value, label]) => ({ value, label }))
    .sort((a, b) => a.label.localeCompare(b.label));

  return [`<option value="">Semua</option>`]
    .concat(arr.map((x) => `<option value="${x.value}">${x.label}</option>`))
    .join("");
}

// =====================
// UI ROUTING
// =====================
function mount(html) {
  $("#viewContainer").innerHTML = html;
}

function setWhoAmI() {
  const box = $("#whoami");
  if (!state.user) { if (box) box.textContent = ""; return; }
  if (box) box.textContent = `${state.user.name} • ${state.user.role.toUpperCase()} • ${state.user.username}`;
  $("#btnLogout")?.classList.remove("d-none");
  $("#btnSync")?.classList.toggle("d-none", !navigator.onLine);
}

function buildMenu() {
  const menu = $("#sideMenu");
  if (!menu) return;
  menu.innerHTML = "";

  const add = (id, icon, label) => {
    const a = document.createElement("a");
    a.href = "#";
    a.className = "list-group-item list-group-item-action d-flex align-items-center gap-2";
    a.dataset.view = id;
    a.innerHTML = `<i class="bi ${icon}"></i><span>${label}</span>`;
    menu.appendChild(a);
  };

  add("dashboard", "bi-speedometer2", "Dashboard");
  add("nilai", "bi-table", "Daftar Nilai");

  if (state.user.role === "admin") {
    add("input", "bi-pencil-square", "Input Data Nilai");
    add("master", "bi-database", "Master Data");
    add("setting", "bi-gear", "Setting");
  } else {
    add("setting_user", "bi-gear", "Setting");
  }

  menu.addEventListener("click", (e) => {
    const a = e.target.closest("a[data-view]");
    if (!a) return;
    e.preventDefault();
    renderView(a.dataset.view);
  });
}

function renderView(view) {
  if (view === "dashboard") return renderDashboard();
  if (view === "nilai") return renderNilaiList();
  if (view === "input") return renderInputNilai();
  if (view === "master") return renderMaster();
  if (view === "setting") return renderSettingAdmin();
  if (view === "setting_user") return renderSettingUser();
}

// =====================
// DASHBOARD
// =====================
function dashPickMateriHighLow(arr) {
  const n = arr.length;
  let nHigh = 0, nLow = 0;

  if (n >= 6) { nHigh = 3; nLow = 3; }
  else if (n === 5) { nHigh = 2; nLow = 3; }
  else if (n === 4) { nHigh = 2; nLow = 2; }
  else if (n === 3) { nHigh = 1; nLow = 2; }
  else if (n === 2) { nHigh = 1; nLow = 1; }
  else if (n === 1) { nHigh = 1; nLow = 1; }
  else { nHigh = 0; nLow = 0; }

  const high = arr.slice(0, nHigh);
  const low = nLow ? arr.slice(Math.max(0, n - nLow)).reverse() : [];
  return { high, low };
}

async function renderDashboard() {
  const allNilai = await dbAll("nilai");
  const f = state.filters;

  const filtered = allNilai.filter((r) => {
    const rNik = normNik(r.nik);
    const uNik = normNik(state.user?.username);

    if (state.user.role !== "admin" && rNik !== uNik) return false;
    if (f.tahun && String(normStr(r.tahun)) !== String(normStr(f.tahun))) return false;
    if (f.jenis && normStr(r.jenis_pelatihan) !== normStr(f.jenis)) return false;

    if (state.user.role === "admin" && f.nik && rNik !== normNik(f.nik)) return false;

    if (f.test && normStr(r.test_type) !== normStr(f.test)) return false;
    if (f.materi) {
      const mk = materiKeyFromRow(r);
      if (mk !== normStr(f.materi)) return false;
    }
    return true;
  });

  const byPeserta = new Map();
  for (const r of filtered) {
    const key = r.nik;
    if (!byPeserta.has(key)) byPeserta.set(key, { nik: r.nik, nama: r.nama, jenis: r.jenis_pelatihan, total: 0, n: 0 });
    const o = byPeserta.get(key);
    if (typeof r.nilai === "number") { o.total += r.nilai; o.n += 1; }
  }

  const pesertaArr = [...byPeserta.values()]
    .map((o) => ({ ...o, avg: o.n ? o.total / o.n : 0 }))
    .sort((a, b) => b.avg - a.avg);

  const top3 = pesertaArr.slice(0, 3);
  const low3 = pesertaArr.slice(-3).reverse();

  const byMateri = new Map();
  for (const r of filtered) {
    const key = r.materi_kode || r.materi_nama;
    if (!key) continue;
    if (!byMateri.has(key)) byMateri.set(key, { kode: r.materi_kode, nama: r.materi_nama, total: 0, n: 0 });
    const o = byMateri.get(key);
    if (typeof r.nilai === "number") { o.total += r.nilai; o.n += 1; }
  }

  const materiArr = [...byMateri.values()]
    .map((o) => ({ ...o, avg: o.n ? o.total / o.n : 0 }))
    .sort((a, b) => b.avg - a.avg);

  const { high: materiHighList, low: materiLowList } = dashPickMateriHighLow(materiArr);
  const gagal = pesertaArr.filter((p) => p.avg < 70);

  mount(`
    <div class="d-flex flex-wrap gap-2 align-items-end mb-3">
      <div class="me-auto">
        <div class="h4 mb-0">Dashboard</div>
        <div class="text-muted small">Ringkasan nilai berdasarkan filter.</div>
      </div>
      <div class="row g-2" style="min-width:420px;">
        <div class="col-4">
          <label class="form-label small mb-1">Tahun</label>
          <select id="fTahun" class="form-select form-select-sm">${optYears()}</select>
        </div>
        <div class="col-8">
          <label class="form-label small mb-1">Jenis Pelatihan</label>
          <select id="fJenis" class="form-select form-select-sm">${optPelatihan()}</select>
        </div>

        <div class="col-5">
          <label class="form-label small mb-1">Jenis Test</label>
          <select id="fTest" class="form-select form-select-sm">${optTestTypes()}</select>
        </div>
        <div class="col-7">
          <label class="form-label small mb-1">Materi</label>
          <select id="fMateri" class="form-select form-select-sm">${optMateriFromNilai(filtered)}</select>
        </div>

        <div class="col-12 ${state.user.role !== "admin" ? "d-none" : ""}">
          <label class="form-label small mb-1">NIK Peserta</label>
          <input id="fNik" class="form-control form-control-sm" placeholder="Kosongkan untuk semua">
        </div>
      </div>
    </div>

    <div class="row g-3">
      <div class="col-12 col-md-6">
        <div class="card border-0 shadow-soft kpi-card">
          <div class="card-body">
            <div class="text-muted small">Jumlah Data Nilai (filtered)</div>
            <div class="big">${filtered.length}</div>
            <div class="text-muted small">Jumlah Peserta</div>
            <div class="fw-semibold">${pesertaArr.length}</div>
          </div>
        </div>
      </div>

      <div class="col-12 col-md-6">
        <div class="card border-0 shadow-soft kpi-card">
          <div class="card-body">
            <div class="text-muted small">Antrian Belum Terkirim</div>
            <div class="big" id="kQueue">0</div>
            <div class="text-muted small">Klik tombol Sync saat online.</div>
          </div>
        </div>
      </div>

      <div class="col-12">
        <div class="card border-0 shadow-soft">
          <div class="card-body">
            <div class="fw-semibold mb-2">3 Peserta Terbaik</div>
            ${
              top3.length
                ? top3.map((p, i) => `
                  <div class="d-flex justify-content-between border-bottom py-2">
                    <div><span class="badge text-bg-primary me-2">#${i + 1}</span>${p.nama} <span class="text-muted">(${p.nik})</span></div>
                    <div class="fw-bold">${p.avg.toFixed(1)}</div>
                  </div>
                `).join("")
                : `<div class="text-muted">Belum ada data.</div>`
            }
          </div>
        </div>
      </div>

      <div class="col-12 col-md-6">
        <div class="card border-0 shadow-soft">
          <div class="card-body">
            <div class="fw-semibold mb-2">Materi Rata-rata Tertinggi</div>
            ${
              materiHighList.length
                ? materiHighList.map((m, i) => `
                  <div class="d-flex justify-content-between ${i < materiHighList.length - 1 ? "border-bottom" : ""} py-2">
                    <div class="small">${m.kode || ""} ${m.nama || ""}</div>
                    <div class="fw-bold">${m.avg.toFixed(1)}</div>
                  </div>
                `).join("")
                : `<div class="text-muted">—</div>`
            }

            <hr>
            <div class="fw-semibold mb-2">Materi Rata-rata Terendah</div>
            ${
              materiLowList.length
                ? materiLowList.map((m, i) => `
                  <div class="d-flex justify-content-between ${i < materiLowList.length - 1 ? "border-bottom" : ""} py-2">
                    <div class="small">${m.kode || ""} ${m.nama || ""}</div>
                    <div class="fw-bold">${m.avg.toFixed(1)}</div>
                  </div>
                `).join("")
                : `<div class="text-muted">—</div>`
            }
          </div>
        </div>
      </div>

      <div class="col-12 col-md-6">
        <div class="card border-0 shadow-soft">
          <div class="card-body">
            <div class="fw-semibold mb-2">3 Peserta Terendah</div>
            ${
              low3.length
                ? low3.map((p, i) => `
                  <div class="d-flex justify-content-between border-bottom py-2">
                    <div><span class="badge text-bg-warning me-2">#${i + 1}</span>${p.nama} <span class="text-muted">(${p.nik})</span></div>
                    <div class="fw-bold">${p.avg.toFixed(1)}</div>
                  </div>
                `).join("")
                : `<div class="text-muted">—</div>`
            }

            <hr>
            <div class="fw-semibold mb-2">Peserta Gagal / Tidak Lulus</div>
            ${
              gagal.length
                ? gagal.slice(0, 5).map((p) => `<div class="small">${p.nama} (${p.nik}) • ${p.avg.toFixed(1)}</div>`).join("")
                : `<div class="text-muted small">Tidak ada (nilai avg<70).</div>`
            }
          </div>
        </div>
      </div>
    </div>
  `);

  $("#fTahun").value = String(state.filters.tahun);
  $("#fJenis").value = state.filters.jenis;
  $("#fTest").value = state.filters.test;
  $("#fMateri").value = state.filters.materi;
  if ($("#fNik")) $("#fNik").value = state.filters.nik;

  $("#fTahun").addEventListener("change", (e) => { state.filters.tahun = e.target.value; renderDashboard(); });
  $("#fJenis").addEventListener("change", (e) => { state.filters.jenis = e.target.value; renderDashboard(); });
  $("#fTest").addEventListener("change", (e) => { state.filters.test = e.target.value; renderDashboard(); });
  $("#fMateri").addEventListener("change", (e) => { state.filters.materi = e.target.value; renderDashboard(); });
  if ($("#fNik")) $("#fNik").addEventListener("input", (e) => { state.filters.nik = e.target.value.trim(); });
  if ($("#fNik")) $("#fNik").addEventListener("change", () => renderDashboard());

  const q = await dbAll("queue");
  $("#kQueue").textContent = String(q.length);
}

// =====================
// NILAI LIST + TRANSKRIP (PREVIEW + PDF) -- NO DUPLICATE HELPERS
// =====================
async function renderNilaiList() {
  const allNilai = await dbAll("nilai");
  const f = state.filters;

  function baseFilterForTranscript() {
    return allNilai.filter((r) => {
      const rNik = normNik(r.nik);
      const uNik = normNik(state.user?.username);

      if (state.user.role !== "admin" && rNik !== uNik) return false;
      if (f.tahun && String(normStr(r.tahun)) !== String(normStr(f.tahun))) return false;
      if (f.jenis && normStr(r.jenis_pelatihan) !== normStr(f.jenis)) return false;

      if (f.materi) {
        const mk = materiKeyFromRow(r);
        if (mk !== normStr(f.materi)) return false;
      }

      if (state.user.role === "admin" && f.nik && rNik !== normNik(f.nik)) return false;
      return true;
    });
  }

  let rows = allNilai
    .filter((r) => {
      const rNik = normNik(r.nik);
      const uNik = normNik(state.user?.username);

      if (state.user.role !== "admin" && rNik !== uNik) return false;
      if (f.tahun && String(normStr(r.tahun)) !== String(normStr(f.tahun))) return false;
      if (f.jenis && normStr(r.jenis_pelatihan) !== normStr(f.jenis)) return false;
      if (f.test && normStr(r.test_type) !== normStr(f.test)) return false;

      if (f.materi) {
        const mk = materiKeyFromRow(r);
        if (mk !== normStr(f.materi)) return false;
      }

      if (state.user.role === "admin" && f.nik && rNik !== normNik(f.nik)) return false;
      return true;
    })
    .sort(
      (a, b) =>
        (b.tanggal_ts || 0) - (a.tanggal_ts || 0) ||
        (b.tanggal || "").localeCompare(a.tanggal || "")
    );

  let limit = 20;

  mount(`
    <div class="d-flex flex-wrap gap-2 align-items-end mb-3">
      <div class="me-auto">
        <div class="h4 mb-0">Daftar Nilai</div>
        <div class="text-muted small">Filter dan export data.</div>
      </div>
      <div class="row g-2">
        <div class="col-4">
          <label class="form-label small mb-1">Tahun</label>
          <select id="fTahun" class="form-select form-select-sm">${optYears()}</select>
        </div>
        <div class="col-8">
          <label class="form-label small mb-1">Jenis Pelatihan</label>
          <select id="fJenis" class="form-select form-select-sm">${optPelatihan()}</select>
        </div>
        <div class="col-4">
          <label class="form-label small mb-1">Jenis Test</label>
          <select id="fTest" class="form-select form-select-sm">${optTestTypes()}</select>
        </div>
        <div class="col-8">
          <label class="form-label small mb-1">Materi</label>
          <select id="fMateri" class="form-select form-select-sm">${optMateriFromNilai(rows)}</select>
        </div>
        <div class="col-6 ${state.user.role !== "admin" ? "d-none" : ""}">
          <label class="form-label small mb-1">NIK</label>
          <input id="fNik" class="form-control form-control-sm" placeholder="Semua">
        </div>
        <div class="col-6">
          <label class="form-label small mb-1">Tampil</label>
          <select id="fLimit" class="form-select form-select-sm">
            <option>20</option><option>50</option><option>100</option><option>500</option><option>1000</option>
          </select>
        </div>
      </div>
    </div>

    <div class="card border-0 shadow-soft">
      <div class="card-body">
        <div class="d-flex flex-wrap gap-2 mb-2">
          <button class="btn btn-outline-primary btn-sm" id="btnExportXlsx"><i class="bi bi-file-earmark-spreadsheet"></i> Export XLSX</button>
          <button class="btn btn-outline-secondary btn-sm" id="btnPreviewTrx"><i class="bi bi-eye"></i> Preview Transkrip</button>
          <button class="btn btn-outline-danger btn-sm" id="btnExportPdf"><i class="bi bi-file-earmark-pdf"></i> Export Transkrip PDF</button>
          <div class="ms-auto d-flex align-items-center gap-2">
            <label class="small text-muted">Tanggal Transkrip</label>
            <input type="date" id="trxDate" class="form-control form-control-sm" style="width:160px;">
          </div>
        </div>

        <div class="table-responsive">
          <table class="table table-sm align-middle">
            <thead>
              <tr>
                <th>Tanggal</th><th>NIK</th><th>Nama</th><th>Pelatihan</th><th>Test</th><th>Materi</th><th class="text-end">Nilai</th>
              </tr>
            </thead>
            <tbody id="tblNilai"></tbody>
          </table>
        </div>
        <div class="small text-muted">Total: <span id="rowCount">0</span></div>
      </div>
    </div>
  `);

  // defaults
  $("#fTahun").value = String(state.filters.tahun);
  $("#fJenis").value = state.filters.jenis;
  $("#fTest").value = state.filters.test;
  $("#fMateri").value = state.filters.materi;
  if ($("#fNik")) $("#fNik").value = state.filters.nik;
  $("#fLimit").value = String(limit);
  $("#trxDate").valueAsDate = new Date();

  function renderTable() {
    const tbody = $("#tblNilai");
    tbody.innerHTML = "";
    const show = rows.slice(0, limit);

    for (const r of show) {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td class="small">${fmtTanggalDisplay(r.tanggal)}</td>
        <td>${r.nik}</td>
        <td>${r.nama || ""}</td>
        <td class="small">${r.jenis_pelatihan || ""}</td>
        <td>${r.test_type || ""}</td>
        <td class="small">${(r.materi_kode || "")} ${(r.materi_nama || "")}</td>
        <td class="text-end fw-semibold">${(r.nilai ?? "")}</td>`;
      tbody.appendChild(tr);
    }

    $("#rowCount").textContent = String(rows.length);
  }

  renderTable();

  $("#fTahun").addEventListener("change", (e) => { state.filters.tahun = e.target.value; renderNilaiList(); });
  $("#fJenis").addEventListener("change", (e) => { state.filters.jenis = e.target.value; renderNilaiList(); });
  $("#fTest").addEventListener("change", (e) => { state.filters.test = e.target.value; renderNilaiList(); });
  $("#fMateri").addEventListener("change", (e) => { state.filters.materi = e.target.value; renderNilaiList(); });
  if ($("#fNik")) $("#fNik").addEventListener("change", (e) => { state.filters.nik = e.target.value.trim(); renderNilaiList(); });
  $("#fLimit").addEventListener("change", (e) => { limit = parseInt(e.target.value, 10); renderTable(); });

  $("#btnExportXlsx").addEventListener("click", () => {
    const data = rows.map((r) => ({
      tahun: r.tahun,
      jenis_pelatihan: r.jenis_pelatihan,
      nik: r.nik,
      nama: r.nama,
      test_type: r.test_type,
      materi_kode: r.materi_kode,
      materi_nama: r.materi_nama,
      nilai: r.nilai,
      tanggal: r.tanggal,
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "nilai");
    XLSX.writeFile(wb, `nilai_${state.filters.tahun || "all"}.xlsx`);
  });

  // ---------- Preview Modal ----------
  function ensureTrxPreviewModal() {
    if (document.getElementById("modalTrxPreview")) return;

    const el = document.createElement("div");
    el.innerHTML = `
    <div class="modal fade" id="modalTrxPreview" tabindex="-1">
      <div class="modal-dialog modal-xl modal-dialog-scrollable">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title">Preview Transkrip Nilai</h5>
            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
          </div>

          <div class="modal-body">
            <div class="row g-2 align-items-end mb-2">
              <div class="col-12 col-md-6">
                <label class="form-label small mb-1">Pilih Peserta</label>
                <select id="trxNikPick" class="form-select form-select-sm"></select>
              </div>
              <div class="col-6 col-md-3">
                <label class="form-label small mb-1">Tanggal Transkrip</label>
                <input type="date" id="trxDate2" class="form-control form-control-sm">
              </div>
              <div class="col-6 col-md-3 d-grid">
                <button class="btn btn-primary btn-sm" id="btnTrxRender">
                  <i class="bi bi-arrow-repeat"></i> Tampilkan
                </button>
              </div>
            </div>

            <div id="trxMeta" class="small"></div>
            <div class="mt-2" id="trxBadges"></div>

            <div class="table-responsive mt-2">
              <table class="table table-sm table-bordered align-middle">
                <thead class="table-light">
                  <tr>
                    <th style="width:50px;">No</th>
                    <th style="width:90px;">Kode</th>
                    <th>Jenis Materi</th>
                    <th class="text-end" style="width:120px;">Rerata Kelas</th>
                    <th class="text-end" style="width:90px;">Poin</th>
                    <th class="text-end" style="width:90px;">Nilai</th>
                  </tr>
                </thead>
                <tbody id="trxTblBody"></tbody>
                <tfoot class="table-light">
                  <tr>
                    <th colspan="5" class="text-end">Total Nilai</th>
                    <th class="text-end" id="trxTotal">0.0</th>
                  </tr>
                </tfoot>
              </table>
            </div>

            <div class="d-flex justify-content-between align-items-center">
              <div class="small text-muted" id="trxFootNote"></div>
              <div class="d-flex gap-2">
                <button class="btn btn-outline-success btn-sm" id="btnXlsxFromPreview">
                  <i class="bi bi-file-earmark-spreadsheet"></i> Export Excel
                </button>
                <button class="btn btn-outline-danger btn-sm" id="btnPdfFromPreview">
                  <i class="bi bi-file-earmark-pdf"></i> Cetak PDF
                </button>
              </div>
            </div>
          </div>

        </div>
      </div>
    </div>`;
    document.body.appendChild(el.firstElementChild);
  }

  $("#btnPreviewTrx").addEventListener("click", () => {
    ensureTrxPreviewModal();

    const m = new bootstrap.Modal(document.getElementById("modalTrxPreview"));
    const trxDate = $("#trxDate").value;
    $("#trxDate2").value = trxDate || new Date().toISOString().slice(0, 10);

    const baseFiltered = baseFilterForTranscript();
    const finalRowsAll = baseFiltered.filter((r) => normStr(r.test_type) === "Final");

    const nikSet = [...new Set(finalRowsAll.map((r) => normNik(r.nik)).filter(Boolean))];

    let nikOptions = nikSet;
    if (state.user.role !== "admin") {
      nikOptions = [normNik(state.user.username)];
    }

    const sel = $("#trxNikPick");
    sel.innerHTML = nikOptions
      .map((n) => {
        const p = (state.masters.peserta || []).find((x) => String(x.nik) === String(n)) || {};
        const nm = p.nama || (finalRowsAll.find((x) => normNik(x.nik) === n)?.nama) || "";
        return `<option value="${n}">${n} - ${nm}</option>`;
      })
      .join("");

    if (!sel.value && nikOptions[0]) sel.value = nikOptions[0];

    let __lastTrxData = null;

    const renderPreview = () => {
      const nik = sel.value;
      const data = trxBuildTranscriptData({ nik, finalRowsAll });
      __lastTrxData = data;

      const meta = $("#trxMeta");
      const body = $("#trxTblBody");
      body.innerHTML = "";

      if (!data) {
        meta.innerHTML = `<div class="text-danger small">Tidak ada data Final untuk peserta ini.</div>`;
        $("#trxTotal").textContent = "0.0";
        return;
      }

      const p = data.peserta || {};
      meta.innerHTML = `
        <div class="row g-1 small">
          <div class="col-12 col-md-6">
            <div><b>Nama</b>: ${p.nama || ""}</div>
            <div><b>NIK</b>: ${data.nik}</div>
            <div><b>Tanggal Presentasi</b>: ${trxFmtTanggalIndoLong(p.tanggal_presentasi || "")}</div>
            <div><b>Judul Makalah</b>: ${normStr(p.judul_presentasi || "")}</div>
          </div>
          <div class="col-12 col-md-6">
            <div><b>Lokasi OJT</b>: ${normStr(p.lokasi_ojt || "")}</div>
            <div><b>Unit</b>: ${normStr(p.unit || "")}</div>
            <div><b>Region</b>: ${normStr(p.region || "")}</div>
            <div><b>Jenis Pelatihan</b>: ${data.jenisLabel}</div>
          </div>
        </div>
      `;

      data.lines.forEach((ln, i) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${i + 1}</td>
          <td>${ln.materi_kode}</td>
          <td>${ln.materi_nama}</td>
          <td class="text-end">${ln.rerata_kelas == null ? "" : trxR1(ln.rerata_kelas)}</td>
          <td class="text-end fw-semibold">${trxR1(ln.poin)}</td>
          <td class="text-end fw-semibold">${trxR1(ln.nilai)}</td>
        `;
        body.appendChild(tr);
      });

      $("#trxTotal").textContent = trxR1(data.totalWeighted);

      const total = data.totalWeighted;
      const status = trxGetLulusStatus(total);
      const pred = trxPickPredikatFromMaster(total);

      $("#trxBadges").innerHTML = `
        <div class="d-flex flex-wrap gap-2 align-items-center">
          <span class="badge ${status.badge}">DINYATAKAN: ${status.label}</span>
          <span class="badge text-bg-primary">PREDIKAT: ${pred}</span>
          <span class="badge text-bg-dark">TOTAL: ${trxR1(total)}</span>
        </div>
      `;

      $("#trxFootNote").textContent = `Tanggal Transkrip: ${trxFmtTanggalIndoLong($("#trxDate2").value)}`;
    };

    $("#btnTrxRender").onclick = renderPreview;
    $("#trxNikPick").onchange = renderPreview;

    $("#btnPdfFromPreview").onclick = async () => {
      const keepDate = $("#trxDate2").value;
      state.filters.nik = sel.value;

      await renderNilaiList();

      const trxDateEl = document.getElementById("trxDate");
      if (trxDateEl && keepDate) trxDateEl.value = keepDate;

      document.getElementById("btnExportPdf").click();
    };

    $("#btnXlsxFromPreview").onclick = () => {
      if (!__lastTrxData) return toast("Data transkrip belum ada.");

      const d = __lastTrxData;
      const total = d.totalWeighted;
      const status = trxGetLulusStatus(total);
      const pred = trxPickPredikatFromMaster(total);

      const ringkas = [
        {
          nik: d.nik,
          nama: d.peserta?.nama || "",
          jenis_pelatihan: d.jenisLabel,
          tanggal_transkrip: $("#trxDate2").value,
          lokasi_ojt: d.peserta?.lokasi_ojt || "",
          unit: d.peserta?.unit || "",
          region: d.peserta?.region || "",
          tanggal_presentasi: d.peserta?.tanggal_presentasi || "",
          judul_makalah: d.peserta?.judul_presentasi || "",
          total_nilai: Math.round(total * 10) / 10,
          dinyatakan: status.label,
          predikat: pred,
        },
      ];

      const detail = (d.lines || []).map((x, i) => ({
        no: i + 1,
        materi_kode: x.materi_kode,
        materi_nama: x.materi_nama,
        rerata_kelas: x.rerata_kelas == null ? "" : Math.round(x.rerata_kelas * 10) / 10,
        poin: Math.round((x.poin || 0) * 10) / 10,
        nilai: Math.round((x.nilai || 0) * 10) / 10,
      }));

      const ws1 = XLSX.utils.json_to_sheet(ringkas);
      const ws2 = XLSX.utils.json_to_sheet(detail);

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws1, "Ringkasan");
      XLSX.utils.book_append_sheet(wb, ws2, "Detail");

      XLSX.writeFile(wb, `transkrip_${d.nik}_${String(d.tahun || state.filters.tahun)}.xlsx`);
    };

    renderPreview();
    m.show();
  });

  // ---------- Export PDF ----------
  $("#btnExportPdf").addEventListener("click", async () => {
    const trxDate = $("#trxDate").value;
    if (!rows.length) return toast("Tidak ada data untuk dibuat transkrip.");

    const baseFiltered = baseFilterForTranscript();
    const finalRowsAll = baseFiltered.filter((r) => normStr(r.test_type) === "Final");
    if (!finalRowsAll.length) return toast("Tidak ada data Final untuk dibuat transkrip.");

    const grouped = new Map();
    for (const r of finalRowsAll) {
      if (!grouped.has(r.nik)) grouped.set(r.nik, []);
      grouped.get(r.nik).push(r);
    }
    if (!grouped.size) return toast("Tidak ada data Final untuk dibuat transkrip.");

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: "pt", format: "a4" });

    let first = true;

    // prebuild avg map once (speed)
    const avgMap = trxBuildAvgMap(finalRowsAll);

    function getRerataKelas(jenis, tahun, materiKey) {
      return trxGetRerataKelasFromMap(avgMap, jenis, tahun, materiKey);
    }

    for (const [nik, list] of grouped) {
      if (!first) doc.addPage();
      first = false;

      const peserta =
        (state.masters.peserta || []).find((p) => String(p.nik) === String(nik)) || {};

      const tahun = list[0]?.tahun ?? peserta.tahun ?? state.filters.tahun ?? "";
      const jenisRaw = list[0]?.jenis_pelatihan || peserta.jenis_pelatihan || "-";
      const jenisLabel = trxJenisDenganTahun(jenisRaw, tahun);

      doc.setFont("helvetica", "bold");
      doc.setFontSize(14);
      doc.text("KARYAMAS PLANTATION", 40, 40);
      doc.setFontSize(12);
      doc.text("SERIANG TRAINING CENTER", 40, 58);
      doc.setFontSize(16);
      doc.text("TRANSKRIP NILAI", 40, 86);

      doc.setFontSize(11);
      doc.setFont("helvetica", "normal");
      const left = 40;
      const right = 320;

      doc.text(`Nama : ${peserta.nama || list[0]?.nama || ""}`, left, 120);
      doc.text(`NIK  : ${nik}`, left, 136);

      doc.text(`Tanggal Presentasi : ${trxFmtTanggalIndoLong(peserta.tanggal_presentasi || "")}`, left, 152);
      doc.text(`Judul Makalah      : ${normStr(peserta.judul_presentasi || "")}`, left, 168);

      doc.text(`Lokasi OJT : ${normStr(peserta.lokasi_ojt || "")}`, right, 120);
      doc.text(`Unit       : ${normStr(peserta.unit || "")}`, right, 136);
      doc.text(`Region     : ${normStr(peserta.region || "")}`, right, 152);

      doc.setFont("helvetica", "bold");
      doc.text(`${jenisLabel}`, left, 192);

      const bobotObj = trxGetBobotByJenis(jenisRaw);

      const listSorted = list
        .slice()
        .sort((a, b) => String(a.materi_kode || "").localeCompare(String(b.materi_kode || "")));

      const rowsTbl = [];
      let totalNilaiWeighted = 0;

      for (let idx = 0; idx < listSorted.length; idx++) {
        const r = listSorted[idx];
        const mKey = materiKeyFromRow(r);

        const rk = getRerataKelas(jenisRaw, tahun, mKey);

        const poin = typeof r.nilai === "number" ? r.nilai : parseFloat(r.nilai);
        const poinOk = Number.isFinite(poin) ? poin : 0;

        const kategori = trxGetKategoriMateri(r.materi_kode, r.materi_nama);
        const bobotPct = trxGetBobotPercentForKategori(bobotObj, kategori);

        const nilaiWeighted = poinOk * (bobotPct / 100);
        totalNilaiWeighted += Number.isFinite(nilaiWeighted) ? nilaiWeighted : 0;

        rowsTbl.push([
          String(idx + 1),
          r.materi_kode || "",
          r.materi_nama || "",
          rk == null ? "" : trxR1(rk),
          trxR1(poinOk),
          trxR1(nilaiWeighted),
        ]);
      }

      doc.autoTable({
        startY: 206,
        head: [["No", "Kode", "Jenis Materi", "Rerata Kelas", "Poin", "Nilai"]],
        body: rowsTbl,
        styles: { fontSize: 9, cellPadding: 3 },
        headStyles: { fillColor: [11, 58, 103] },
        margin: { left: 40, right: 40 },
      });

      const y = doc.lastAutoTable.finalY + 18;

      const total = totalNilaiWeighted;
      const pred = trxPickPredikatFromMaster(total);

      doc.setFont("helvetica", "bold");
      doc.text(
        `Total Nilai: ${trxR1(total)}     Dinyatakan: ${total >= 70 ? "LULUS" : "TIDAK LULUS"}     Predikat: ${pred}`,
        40,
        y
      );

      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.text(`Tanggal Transkrip: ${trxFmtTanggalIndoLong(trxDate)}`, 40, y + 18);
      doc.text("Dokumen ini dicetak secara komputerisasi", 40, 800);
    }

    doc.save(`transkrip_${state.filters.tahun || "all"}.pdf`);
  });
}

// =====================
// INPUT NILAI (ADMIN)
// =====================
async function renderInputNilai() {
  mount(`
    <div class="h4 mb-2">Input Data Nilai</div>
    <div class="text-muted small mb-3">Input manual atau upload XLSX.</div>

    <div class="row g-3">
      <div class="col-12">
        <div class="card border-0 shadow-soft">
          <div class="card-body">
            <div class="fw-semibold mb-2">Upload Excel (.xlsx)</div>
            <p class="small text-muted mb-2">Template kolom: tahun, jenis_pelatihan, nik, nama, test_type, materi_kode, materi_nama, nilai, tanggal</p>
            <input type="file" id="fileXlsx" class="form-control" accept=".xlsx">
            <div class="d-flex gap-2 mt-2">
              <button class="btn btn-outline-primary btn-sm" id="btnImport"><i class="bi bi-upload"></i> Import ke Lokal</button>
            </div>
          </div>
        </div>
      </div>

      <div class="col-12">
        <div class="card border-0 shadow-soft">
          <div class="card-body">
            <div class="fw-semibold mb-2">Input Manual</div>
            <form id="formNilai" class="row g-2">
              <div class="col-6 col-md-3">
                <label class="form-label small">Tahun</label>
                <select id="inTahun" class="form-select form-select-sm">${optYears()}</select>
              </div>
              <div class="col-6 col-md-5">
                <label class="form-label small">Jenis Pelatihan</label>
                <input id="inJenis" class="form-control form-control-sm" list="dlJenis" placeholder="Ketik…">
                <datalist id="dlJenis">${optPelatihan().replace('<option value="">Semua</option>', "")}</datalist>
              </div>
              <div class="col-6 col-md-4">
                <label class="form-label small">Jenis Test</label>
                <select id="inTest" class="form-select form-select-sm">
                  <option>PreTest</option><option>PostTest</option><option>Final</option><option>OJT</option><option>Presentasi</option>
                </select>
              </div>

              <div class="col-6 col-md-3">
                <label class="form-label small">NIK</label>
                <input id="inNik" class="form-control form-control-sm" list="dlNik" required>
                <datalist id="dlNik"></datalist>
              </div>
              <div class="col-12 col-md-5">
                <label class="form-label small">Nama</label>
                <input id="inNama" class="form-control form-control-sm" readonly>
              </div>
              <div class="col-12 col-md-4">
                <label class="form-label small">Tanggal</label>
                <input id="inTanggal" type="date" class="form-control form-control-sm" required>
              </div>

              <div class="col-6 col-md-3">
                <label class="form-label small">Kode Materi</label>
                <input id="inKode" class="form-control form-control-sm" list="dlMateriKode">
                <datalist id="dlMateriKode"></datalist>
              </div>
              <div class="col-6 col-md-6">
                <label class="form-label small">Nama Materi</label>
                <input id="inMateri" class="form-control form-control-sm" list="dlMateriNama">
                <datalist id="dlMateriNama"></datalist>
              </div>
              <div class="col-6 col-md-3">
                <label class="form-label small">Nilai</label>
                <input id="inNilai" type="number" min="0" max="100" step="0.1" class="form-control form-control-sm" required>
              </div>

              <div class="col-12 d-flex gap-2 mt-2">
                <button class="btn btn-primary" type="submit"><i class="bi bi-save"></i> Simpan </button>
                <button class="btn btn-outline-secondary" type="button" id="btnClear">Clear</button>
              </div>
            </form>
          </div>
        </div>
      </div>
    </div>
  `);

  $("#inTanggal").valueAsDate = new Date();

  $("#dlNik").innerHTML = state.masters.peserta.map((p) => `<option value="${p.nik}">${p.nama}</option>`).join("");
  $("#dlMateriKode").innerHTML = state.masters.materi.map((m) => `<option value="${m.kode}">${m.nama}</option>`).join("");
  $("#dlMateriNama").innerHTML = state.masters.materi.map((m) => `<option value="${m.nama}">${m.kode}</option>`).join("");

  $("#inNik").addEventListener("input", () => {
    const nik = $("#inNik").value.trim();
    const p = state.masters.peserta.find((x) => String(x.nik) === String(nik));
    $("#inNama").value = p?.nama || "";
    if (p?.jenis_pelatihan && !$("#inJenis").value) $("#inJenis").value = p.jenis_pelatihan;
  });

  $("#inKode").addEventListener("input", () => {
    const kode = $("#inKode").value.trim();
    const m = state.masters.materi.find((x) => String(x.kode) === String(kode));
    if (m) $("#inMateri").value = m.nama;
  });

  $("#inMateri").addEventListener("input", () => {
    const nama = $("#inMateri").value.trim();
    const m = state.masters.materi.find((x) => String(x.nama) === String(nama));
    if (m) $("#inKode").value = m.kode;
  });

  $("#formNilai").addEventListener("submit", async (e) => {
    e.preventDefault();

    let row = {
      id: "N" + Date.now() + "_" + Math.random().toString(16).slice(2),
      tahun: parseInt($("#inTahun").value, 10),
      jenis_pelatihan: $("#inJenis").value.trim(),
      test_type: $("#inTest").value,
      nik: $("#inNik").value.trim(),
      nama: $("#inNama").value.trim(),
      tanggal: $("#inTanggal").value, // YYYY-MM-DD
      materi_kode: $("#inKode").value.trim(),
      materi_nama: $("#inMateri").value.trim(),
      nilai: parseFloat($("#inNilai").value),
    };

    row = normalizeNilaiRow(row);

    await dbPut("nilai", row);
    await enqueue("upsert_nilai", row);
    toast("Nilai disimpan ke lokal & masuk antrian sync.");
    $("#inNilai").value = "";
  });

  $("#btnClear").addEventListener("click", () => {
    $("#formNilai").reset();
    $("#inTanggal").valueAsDate = new Date();
  });

  $("#btnImport").addEventListener("click", async () => {
    const btn = $("#btnImport");
    runBusy(btn, async () => {
      const f = $("#fileXlsx").files[0];
      if (!f) return toast("Pilih file xlsx dulu.");

      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows2 = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      const headerRaw2 = (rows2[0] || []).map((h) => String(h || ""));

      const norm2 = (s) => {
        s = String(s || "").trim().toLowerCase().replace(/\s+/g, "_").replace(/[^a-z0-9_]/g, "");
        if (s === "nik") return "nik";
        if (s === "nama") return "nama";
        if (s === "tahun") return "tahun";
        if (s === "jenis_pelatihan") return "jenis_pelatihan";
        if (s === "test_type") return "test_type";
        if (s === "materi_kode") return "materi_kode";
        if (s === "materi_nama") return "materi_nama";
        if (s === "nilai") return "nilai";
        if (s === "tanggal") return "tanggal";
        return s;
      };
      const header2 = headerRaw2.map(norm2);

      const data = [];
      for (let i = 1; i < rows2.length; i++) {
        const arr = rows2[i];
        if (!arr || arr.every((v) => String(v).trim() === "")) continue;
        const obj = {};
        header2.forEach((k, idx) => { if (k) obj[k] = arr[idx]; });
        data.push(obj);
      }

      const total = data.length;
      progressStart(total, `Import ${total} baris…`);

      let n = 0;
      for (const r of data) {
        let row = {
          id: "N" + Date.now() + "_" + Math.random().toString(16).slice(2),
          tahun: r.tahun,
          jenis_pelatihan: r.jenis_pelatihan,
          test_type: r.test_type,
          nik: r.nik,
          nama: r.nama,
          tanggal: r.tanggal,
          materi_kode: r.materi_kode,
          materi_nama: r.materi_nama,
          nilai: r.nilai,
        };

        row = normalizeNilaiRow(row);
        await dbPut("nilai", row);
        await enqueue("upsert_nilai", row);

        n++;
        if (n % 10 === 0 || n === total) progressSet(n, total, `Mengimpor… (${n}/${total})`);
      }

      await rebuildPelatihanCache();
      progressDone("Import selesai.");
      toast(`Import selesai: ${n} baris (masuk antrian sync).`);
    }, { busyText: "Import…" }).catch((e) => toast(e.message));
  });
}

// =====================
// MASTER
// =====================
async function renderMaster() {
  mount(`
    <div class="h4 mb-2">Master Data</div>
    <div class="text-muted small mb-3">Upload master ke Google Sheet (disarankan) atau tarik dari Google Sheet.</div>

    <div class="card border-0 shadow-soft mb-3 d-none">
      <div class="card-body">
        <div class="fw-semibold mb-2">Tarik Master dari Google Sheet</div>
        <div class="row g-2">
          <div class="col-4">
            <label class="form-label small mb-1">Tahun</label>
            <select id="mTahun" class="form-select form-select-sm">${optYears()}</select>
          </div>
          <div class="col-8">
            <label class="form-label small mb-1">Jenis Pelatihan</label>
            <select id="mJenis" class="form-select form-select-sm">${optPelatihan()}</select>
          </div>
        </div>
        <button class="btn btn-primary btn-sm mt-2" id="btnPullMaster"><i class="bi bi-cloud-download"></i> Tarik</button>
      </div>
    </div>

    <div class="row g-3">
      <div class="col-12 col-md-6">
        <div class="card border-0 shadow-soft">
          <div class="card-body">
            <div class="fw-semibold mb-2">Upload Master Peserta (.xlsx)</div>
            <p class="small text-muted mb-2">Kolom: nik, nama, jenis_pelatihan, lokasi_ojt, unit, region, group, tanggal_presentasi, judul_presentasi</p>
            <input type="file" id="upPeserta" class="form-control" accept=".xlsx">
            <button class="btn btn-outline-primary btn-sm mt-2" id="btnUpPeserta"><i class="bi bi-upload"></i> Upload ke GAS</button>
          </div>
        </div>
      </div>
      <div class="col-12 col-md-6">
        <div class="card border-0 shadow-soft">
          <div class="card-body">
            <div class="fw-semibold mb-2">Upload Master Materi (.xlsx)</div>
            <p class="small text-muted mb-2">Kolom: kode, nama, kategori (Managerial/Teknis/Support/OJT/Presentasi)</p>
            <input type="file" id="upMateri" class="form-control" accept=".xlsx">
            <button class="btn btn-outline-primary btn-sm mt-2" id="btnUpMateri"><i class="bi bi-upload"></i> Upload ke GAS</button>
          </div>
        </div>
      </div>
    </div>
  `);

  $("#mTahun")?.value = String(state.filters.tahun);
  $("#mJenis")?.value = state.filters.jenis;

  $("#btnPullMaster")?.addEventListener("click", async () => {
    const btn = $("#btnPullMaster");
    runBusy(btn, async () => {
      const tahun = $("#mTahun").value;
      const jenis = $("#mJenis").value;

      const res = await gasCall("pull_master", { tahun, jenis });

      state.masters = res.masters;
      await saveMastersToDB();
      await rebuildPelatihanCache();
      await syncUsersFromMasterPeserta();

      toast("Master ditarik & disimpan ke offline.");
    }, { busyText: "Menarik…" }).catch((e) => toast(e.message));
  });

  async function uploadXlsx(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (!rows.length) return [];

    const headerRaw = rows[0].map((h) => String(h || ""));
    const norm = (s) => {
      s = String(s || "").trim().toLowerCase().replace(/\s+/g, "_").replace(/[^a-z0-9_]/g, "");
      if (s.startsWith("kategori")) return "kategori";
      return s;
    };
    const header = headerRaw.map(norm);

    const out = [];
    for (let i = 1; i < rows.length; i++) {
      const arr = rows[i];
      if (!arr || arr.every((v) => String(v).trim() === "")) continue;
      const obj = {};
      header.forEach((k, idx) => { if (k) obj[k] = arr[idx]; });
      out.push(obj);
    }
    return out;
  }

  $("#btnUpPeserta")?.addEventListener("click", async () => {
    const btn = $("#btnUpPeserta");
    runBusy(btn, async () => {
      const f = $("#upPeserta").files[0];
      if (!f) return toast("Pilih file dulu.");

      const rows = await uploadXlsx(f);
      const total = rows.length;

      progressStart(total, `Menyiapkan ${total} baris master peserta…`);
      let n = 0;
      for (const r of rows) {
        await enqueue("upsert_master_peserta", r);
        n++;
        if (n % 20 === 0 || n === total) progressSet(n, total, `Masuk antrian… (${n}/${total})`);
      }

      state.masters.peserta = state.masters.peserta.concat(
        rows.map((r) => ({
          nik: (r.nik || "").toString(),
          nama: (r.nama || "").toString(),
          jenis_pelatihan: (r.jenis_pelatihan || "").toString(),
          lokasi_ojt: (r.lokasi_ojt || "").toString(),
          unit: (r.unit || "").toString(),
          region: (r.region || "").toString(),
          group: (r.group || "").toString(),
          tanggal_presentasi: (r.tanggal_presentasi || "").toString(),
          judul_presentasi: (r.judul_presentasi || "").toString(),
          tahun: (r.tahun || "").toString(),
        }))
      );

      await saveMastersToDB();
      await rebuildPelatihanCache();
      await syncUsersFromMasterPeserta();

      progressDone("Upload master peserta selesai.");
      toast(`Master peserta masuk antrian sync: ${total} baris.`);
      if (navigator.onLine) syncQueue().catch(() => {});
    }, { busyText: "Memproses…" }).catch((e) => toast(e.message));
  });

  $("#btnUpMateri")?.addEventListener("click", async () => {
    const btn = $("#btnUpMateri");
    runBusy(btn, async () => {
      const f = $("#upMateri").files[0];
      if (!f) return toast("Pilih file dulu.");

      const rows = await uploadXlsx(f);
      const total = rows.length;

      progressStart(total, `Menyiapkan ${total} baris master materi…`);
      let n = 0;
      for (const r of rows) {
        await enqueue("upsert_master_materi", r);
        n++;
        if (n % 20 === 0 || n === total) progressSet(n, total, `Masuk antrian… (${n}/${total})`);
      }

      state.masters.materi = state.masters.materi.concat(
        rows.map((r) => ({
          kode: (r.kode || "").toString(),
          nama: (r.nama || "").toString(),
          kategori: (r.kategori || "").toString(),
        }))
      );

      await saveMastersToDB();

      progressDone("Upload master materi selesai.");
      toast(`Master materi masuk antrian sync: ${total} baris.`);
      if (navigator.onLine) syncQueue().catch(() => {});
    }, { busyText: "Memproses…" }).catch((e) => toast(e.message));
  });
}

// =====================
// SETTINGS (ADMIN / USER)
// =====================
async function renderSettingAdmin() {
  mount(`
    <div class="h4 mb-2">Setting (Admin)</div>

    <div class="card border-0 shadow-soft mb-3 d-none">
      <div class="card-body">
        <div class="fw-semibold mb-2">Google Apps Script WebApp URL</div>
        <input id="setGas" class="form-control" placeholder="Paste URL WebApp GAS di sini">
        <button class="btn btn-primary btn-sm mt-2" id="btnSaveGas"><i class="bi bi-save"></i> Simpan</button>
      </div>
    </div>

    <div class="row g-3">
      <div class="col-12 col-md-6">
        <div class="card border-0 shadow-soft">
          <div class="card-body">
            <div class="fw-semibold mb-2">Tarik Data Aktual (Nilai) dari Google Sheet</div>
            <div class="row g-2">
              <div class="col-4">
                <label class="form-label small mb-1">Tahun</label>
                <select id="aTahun" class="form-select form-select-sm">${optYears()}</select>
              </div>
              <div class="col-8">
                <label class="form-label small mb-1">Jenis Pelatihan</label>
                <select id="aJenis" class="form-select form-select-sm">${optPelatihan()}</select>
              </div>
            </div>
            <button class="btn btn-outline-primary btn-sm mt-2" id="btnPullNilai"><i class="bi bi-cloud-download"></i> Tarik Nilai</button>
          </div>
        </div>
      </div>

      <div class="col-12 col-md-6">
        <div class="card border-0 shadow-soft">
          <div class="card-body">
            <div class="fw-semibold mb-2">Reset Password User</div>
            <p class="small text-muted mb-2">Set password kembali ke default (username/NIK).</p>
            <input id="resetUser" class="form-control form-control-sm" placeholder="username / NIK">
            <button class="btn btn-danger btn-sm mt-2" id="btnReset"><i class="bi bi-arrow-counterclockwise"></i> Reset</button>
          </div>
        </div>
      </div>

      <div class="col-12 col-md-6">
        <div class="card border-0 shadow-soft">
          <div class="card-body">
            <div class="fw-semibold mb-2">Ganti Password Saya</div>
            <p class="small text-muted mb-2">Ganti password akun yang sedang login.</p>
            <input id="admNewPass" type="password" class="form-control form-control-sm" placeholder="Password baru (min 6)">
            <button class="btn btn-primary btn-sm mt-2" id="btnAdmChangePass">
              <i class="bi bi-key"></i> Simpan
            </button>
          </div>
        </div>
      </div>

      <div class="col-12">
        <div class="card border-0 shadow-soft">
          <div class="card-body">
            <div class="fw-semibold mb-2">Hapus Data Lokal</div>
            <button class="btn btn-outline-danger btn-sm" id="btnWipe"><i class="bi bi-trash"></i> Hapus IndexedDB</button>
          </div>
        </div>
      </div>
    </div>
  `);

  $("#setGas")?.setAttribute("value", localStorage.getItem("GAS_URL") || "");
  $("#aTahun").value = String(state.filters.tahun);
  $("#aJenis").value = state.filters.jenis;

  $("#btnSaveGas")?.addEventListener("click", () => {
    localStorage.setItem("GAS_URL", $("#setGas").value.trim());
    toast("GAS_URL disimpan.");
  });

  $("#btnPullNilai")?.addEventListener("click", async () => {
    const btn = $("#btnPullNilai");
    runBusy(btn, async () => {
      const res = await gasCall("pull_nilai", { tahun: $("#aTahun").value, jenis: $("#aJenis").value });
      const total = res.rows?.length || 0;
      progressStart(total, "Menarik nilai dari Google Sheet…");

      let n = 0;
      for (const r of res.rows || []) {
        const row = ensureNilaiId(normalizeNilaiRow(r));
        await dbPut("nilai", row);
        n++;
        if (n % 10 === 0 || n === total) {progressSet(n, total, `Menyimpan ke offline… (${n}/${total})`);}
      }

      await rebuildPelatihanCache();
      progressDone("Tarik nilai selesai.");
      toast(`Tarik nilai selesai: ${n} baris.`);
    }, { busyText: "Menarik…" }).catch((e) => toast(e.message));
  });

  $("#btnReset")?.addEventListener("click", async () => {
    const btn = $("#btnReset");
    runBusy(btn, async () => {
      const username = $("#resetUser").value.trim();
      if (!username) return toast("Isi username/NIK.");
      const newHash = await sha256Hex(username === "admin" ? "123456" : username);

      const u = await dbGet("users", username);
      if (u) {
        u.pass_hash = newHash;
        u.must_change = username !== "admin";
        await dbPut("users", u);
      }

      await gasCall("admin_reset_password", { username, pass_hash: newHash, must_change: username !== "admin" });
      toast("Reset password berhasil.");
    }, { busyText: "Reset…" }).catch((e) => toast(e.message));
  });

  $("#btnAdmChangePass")?.addEventListener("click", async () => {
    const btn = $("#btnAdmChangePass");
    runBusy(btn, async () => {
      const np = $("#admNewPass").value.trim();
      if (np.length < 6) return toast("Minimal 6 karakter.");
      await changePassword(np);
      $("#admNewPass").value = "";
    }, { busyText: "Menyimpan…" }).catch((e) => toast(e.message));
  });

  $("#btnWipe")?.addEventListener("click", async () => {
    if (!confirm("Yakin hapus semua data lokal?")) return;
    const db = await dbPromise;
    await db.clear("masters");
    await db.clear("nilai");
    await db.clear("queue");
    toast("Data lokal dihapus. Silakan tarik master & nilai dari Google Sheet.");
  });
}

async function renderSettingUser() {
  mount(`
    <div class="h4 mb-2">Setting</div>
    <div class="card border-0 shadow-soft">
      <div class="card-body">
        <div class="fw-semibold mb-2">Tarik Data dari Google Sheet</div>
        <button class="btn btn-outline-primary btn-sm" id="btnPullMy"><i class="bi bi-cloud-download"></i> Tarik Master + Nilai Saya</button>
        <hr>
        <div class="fw-semibold mb-2">Ganti Password</div>
        <input id="newPass" type="password" class="form-control form-control-sm" placeholder="Password baru">
        <button class="btn btn-primary btn-sm mt-2" id="btnChangePass"><i class="bi bi-key"></i> Simpan</button>
      </div>
    </div>
  `);

  $("#btnPullMy")?.addEventListener("click", async () => {
    const btn = $("#btnPullMy");
    runBusy(btn, async () => {
      const res = await gasCall("pull_user_bundle", { username: state.user.username });

      state.masters = res.masters;
      await saveMastersToDB();
      await syncUsersFromMasterPeserta();

      const total = res.nilai?.length || 0;
      progressStart(total, "Menarik nilai Anda…");

      let n = 0;
      for (const r of res.nilai || []) {
        const row = ensureNilaiId(normalizeNilaiRow(r));
        await dbPut("nilai", row);
        n++;
        if (n % 10 === 0 || n === total) {progressSet(n, total, `Menyimpan ke offline… (${n}/${total})`);}
      }

      progressDone("Tarik data selesai.");
      toast("Data berhasil ditarik.");

      state.filters.nik = "";
      const all = await dbAll("nilai");
      const mine = all.filter((x) => normNik(x.nik) === normNik(state.user.username));
      const latestYear = mine
        .map((x) => parseInt(x.tahun, 10))
        .filter((n) => !Number.isNaN(n))
        .sort((a, b) => b - a)[0];
      if (latestYear) state.filters.tahun = latestYear;

      await rebuildPelatihanCache();
      renderView("nilai");
    }, { busyText: "Menarik…" }).catch((e) => toast(e.message));
  });

  $("#btnChangePass")?.addEventListener("click", async () => {
    const btn = $("#btnChangePass");
    runBusy(btn, async () => {
      const np = $("#newPass").value.trim();
      if (np.length < 6) return toast("Minimal 6 karakter.");
      await changePassword(np);
      $("#newPass").value = "";
    }, { busyText: "Menyimpan…" }).catch((e) => toast(e.message));
  });
}

// =====================
// FIRST LOGIN: FORCE CHANGE PASSWORD
// =====================
async function promptChangePasswordIfNeeded() {
  if (!state.user?.must_change) return;

  const modalHtml = `
  <div class="modal fade" id="modalPwd" tabindex="-1">
    <div class="modal-dialog">
      <div class="modal-content">
        <div class="modal-header"><h5 class="modal-title">Ganti Password</h5></div>
        <div class="modal-body">
          <p class="small text-muted">Demi keamanan, silakan ganti password default.</p>
          <input type="password" id="pwdNew1" class="form-control mb-2" placeholder="Password baru (min 6)">
          <input type="password" id="pwdNew2" class="form-control" placeholder="Ulangi password">
        </div>
        <div class="modal-footer">
          <button class="btn btn-primary" id="btnPwdSave">Simpan</button>
        </div>
      </div>
    </div>
  </div>`;
  document.body.insertAdjacentHTML("beforeend", modalHtml);

  const m = new bootstrap.Modal($("#modalPwd"), { backdrop: "static", keyboard: false });
  m.show();

  $("#btnPwdSave").addEventListener("click", async () => {
    const btn = $("#btnPwdSave");
    runBusy(btn, async () => {
      const a = $("#pwdNew1").value.trim();
      const b = $("#pwdNew2").value.trim();
      if (a.length < 6) return toast("Minimal 6 karakter.");
      if (a !== b) return toast("Password tidak sama.");
      await changePassword(a);
      m.hide();
      $("#modalPwd").remove();
    }, { busyText: "Menyimpan…" }).catch((e) => toast(e.message));
  });
}

// =====================
// PUBLIC PULL FOR LOGIN SCREEN
// =====================
async function pullPublicMasters() {
  try {
    const res = await gasCall("pull_public_master", {});
    state.masters = res.masters;
    await saveMastersToDB();
    await rebuildPelatihanCache();
    await syncUsersFromMasterPeserta();
    toast("Master publik ditarik. Sekarang user bisa login offline.");
  } catch (e) {
    toast(e.message);
  }
}

// =====================
// BOOT
// =====================

async function boot(){
    // ✅ Chrome Android sering gagal untuk ESM+SW+IndexedDB jika dibuka dari file://
  if(location.protocol === "file:"){
    // tampilkan toast bila UI sudah siap (bootstrap toast ada)
    try{ toast("Aplikasi dibuka dari file://. Di Chrome HP harus lewat https atau http://localhost (tidak bisa dari File Manager)."); }catch(e){}
    console.warn("Running from file:// is not supported on Chrome Android for module+SW.");
    // tetap lanjut, tapi user sudah dapat instruksi jelas
  }
  // register SW
  if("serviceWorker" in navigator){
    try{ await navigator.serviceWorker.register("./sw.js"); }catch(e){ console.warn(e); }
  }

  setNetBadge();
    // ✅ TEST GAS cepat saat online (agar di HP ketahuan masalahnya di awal)
  if(navigator.onLine){
    try{
      await gasPing();
      console.log("GAS ping OK:", getGasUrl());
    }catch(e){
      console.warn("GAS ping failed:", e);
      toast("GAS tidak bisa diakses: " + e.message);
    }
  }

  await ensureDefaultUsers();
  await loadMastersFromDB();
  await syncUsersFromMasterPeserta();
  await rebuildPelatihanCache();
  await refreshQueueBadge();
  await purgePasswordQueue();
  await migrateNilaiTanggalIfNeeded();


  // ============================================================
  // PASANG EVENT LISTENER SEKALI (WAJIB sebelum kemungkinan return)
  // ============================================================

  const formLogin = $("#formLogin");
  if(formLogin){
    formLogin.addEventListener("submit", async (e)=>{
      e.preventDefault();
      const btn = $("#btnLogin") || $("#formLogin button[type='submit']"); // aman bila id beda
      runBusy(btn, async ()=>{
        const u = $("#loginUser")?.value;
        const p = $("#loginPass")?.value;
        await login(u, p);

        await saveSession(state.user);

        $("#viewLogin")?.classList.add("d-none");
        $("#viewApp")?.classList.remove("d-none");
        $("#btnSync")?.classList.toggle("d-none", !navigator.onLine);
        $("#btnLogout")?.classList.remove("d-none");

        setWhoAmI();
        buildMenu();
        renderView("dashboard");

        setTimeout(()=>{ promptChangePasswordIfNeeded().catch(()=>{}); }, 200);

        await refreshQueueBadge();
        if(navigator.onLine) syncQueue().catch(()=>{});
      }, { busyText:"Login…" }).catch(err=>{
        toast(err.message);
      });
    });
  }

  const btnLogout = $("#btnLogout");
  if(btnLogout){
    btnLogout.addEventListener("click", async ()=>{
      state.user = null;

      // HAPUS SESSION (agar refresh kembali login)
      await clearSession();

      $("#viewApp")?.classList.add("d-none");
      $("#viewLogin")?.classList.remove("d-none");
      $("#btnLogout")?.classList.add("d-none");
      $("#btnSync")?.classList.add("d-none");

      const who = $("#whoami");
      if(who) who.textContent = "";

      toast("Logout.");
    });
  }

  const btnSync = $("#btnSync");
  if(btnSync){
    btnSync.addEventListener("click", ()=>{
      runBusy(btnSync, ()=>syncQueue(), { busyText:"Sync…" })
        .catch(e=>toast(e.message));
    });
  }


  const btnPullPublic = $("#btnPullPublic");
  if(btnPullPublic){
    btnPullPublic.addEventListener("click", ()=>{
      runBusy(btnPullPublic, ()=>pullPublicMasters(), { busyText:"Menarik…" })
        .catch(e=>toast(e.message));
    });
  }


  // ============================================================
  // AUTO RESTORE SESSION (30 hari)
  // ============================================================
  const restored = await restoreSessionIntoState();
  if(restored){
    await fixSessionMustChangeFromLocal();
    // langsung masuk app
    $("#viewLogin")?.classList.add("d-none");
    $("#viewApp")?.classList.remove("d-none");
    $("#btnSync")?.classList.toggle("d-none", !navigator.onLine);

    setWhoAmI();
    buildMenu();
    renderView("dashboard");

    // tetap paksa ganti password bila wajib
    setTimeout(()=>{ promptChangePasswordIfNeeded().catch(()=>{}); }, 200);

    await refreshQueueBadge();
    if(navigator.onLine) syncQueue().catch(()=>{});

    // rolling expiry: perpanjang 30 hari setiap app dibuka (hanya jika ada user)
    if(state.user) await saveSession(state.user);

    return; // penting: stop boot di sini, tidak perlu tampil login lagi
  }
}

boot();
