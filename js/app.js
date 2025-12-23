// Karyamas Transkrip Nilai (Offline-first)
// Frontend: Bootstrap 5 + IndexedDB (idb) + XLSX + jsPDF
// Backend: Google Apps Script (see Code.gs)
console.warn("APP.JS LOADED ✅", new Date().toISOString());


// ✅ GAS URL dikunci (hardcode) — tidak bisa diubah via setting/localStorage
const GAS_URL = "https://script.google.com/macros/s/AKfycbxA0ZVZYpGq4ePqFwHsFUGyOoVn0vwf4XyvdCtycPLxo05WI4bw0mURT10iMBnE1txm/exec";

function getGasUrl(){
  return GAS_URL;
}

const { openDB } = window.idb;
if (!openDB) {
  throw new Error("Library idb (openDB) tidak ditemukan. Pastikan idb UMD sudah dimuat sebelum app.js.");
}


const $ = (sel, root=document)=>root.querySelector(sel);
const $$ = (sel, root=document)=>[...root.querySelectorAll(sel)];

// =====================
// DEBUG REMOVED (CCTV cleaned)
// semua call dbg(...) tetap aman (no-op)
// =====================
function dbgOn(){}
function dbgOff(){}
function dbg(){ /* no-op */ }
function dbgLast(){ return []; }

// ✅ Tambahkan ini: agar pemanggilan renderDbgBox() tidak error
function renderDbgBox(){
  // Karena debug sudah dihapus, pastikan elemen debug tetap tersembunyi
  const btn = document.querySelector("#btnDbgToggle");
  const box = document.querySelector("#dbgBox");
  if(btn) btn.classList.add("d-none");
  if(box){
    box.classList.add("d-none");
    box.innerHTML = "";
  }
}


// ---------- UI Busy (Spinner) + Progress ----------
let __busyCount = 0;

function setBtnBusy(btn, busy=true, textBusy="Memproses…"){
  if(!btn) return;
  if(busy){
    if(btn.dataset._busy === "1") return; // already busy
    btn.dataset._busy = "1";
    btn.dataset._origHtml = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = `
      <span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span>
      <span>${textBusy}</span>`;
  }else{
    if(btn.dataset._busy !== "1") return;
    btn.dataset._busy = "0";
    btn.disabled = false;
    btn.innerHTML = btn.dataset._origHtml || btn.innerHTML;
  }
}

function ensureProgressUI(){
  // Prioritas: kalau ada progress khusus login, pakai itu
  if($("#progressWrapLogin")) return;

  // default: progress di bawah viewContainer (app)
  if($("#progressWrap")) return;

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

function _pIds(){
  // jika ada progress login, pakai itu
  if($("#progressWrapLogin")){
    return {
      wrap:"#progressWrapLogin",
      text:"#progressTextLogin",
      now:"#progressNowLogin",
      total:"#progressTotalLogin",
      bar:"#progressBarLogin"
    };
  }
  return {
    wrap:"#progressWrap",
    text:"#progressText",
    now:"#progressNow",
    total:"#progressTotal",
    bar:"#progressBar"
  };
}


function progressStart(total=0, text="Memproses…"){
  ensureProgressUI();
  const ids = _pIds();
  $(ids.wrap).classList.remove("d-none");
  $(ids.text).textContent = text;
  $(ids.total).textContent = String(total || 0);
  $(ids.now).textContent = "0";
  $(ids.bar).style.width = "0%";
}


function progressSet(now, total, text){
  ensureProgressUI();
  const ids = _pIds();
  $(ids.wrap).classList.remove("d-none");
  if(typeof text === "string") $(ids.text).textContent = text;
  $(ids.now).textContent = String(now || 0);
  $(ids.total).textContent = String(total || 0);
  const pct = total ? Math.round((now/total)*100) : 0;
  $(ids.bar).style.width = `${Math.max(0, Math.min(100, pct))}%`;
}


function progressDone(textDone="Selesai."){
  const wrapLogin = $("#progressWrapLogin");
  const ids = _pIds();
  const wrap = $(ids.wrap);
  if(!wrap) return;

  $(ids.text).textContent = textDone;
  $(ids.bar).classList.remove("progress-bar-animated");
  $(ids.bar).style.width = "100%";

  setTimeout(()=>{
    wrap.classList.add("d-none");
    $(ids.bar).classList.add("progress-bar-animated");
  }, 700);
}


async function runBusy(btn, fn, {busyText="Memproses…", progressText=null, total=0} = {}){
  try{
    __busyCount++;
    setBtnBusy(btn, true, busyText);
    if(progressText) progressStart(total, progressText);
    const out = await fn();
    if(progressText) progressDone("Selesai.");
    return out;
  }finally{
    __busyCount = Math.max(0, __busyCount-1);
    setBtnBusy(btn, false);
  }
}


const state = {
  user: null, // {username, name, role}
  masters: { peserta: [], materi: [], pelatihan: [], bobot: [], predikat: [] },
  // ✅ tambah: test & materi
  filters: {
    tahun: new Date().getFullYear(),
    jenis: "",
    nik: "",
    test: "",     // PreTest/PostTest/Final/OJT/Presentasi
    materi: ""    // key materi (kode atau nama)
  }
};

function toast(msg){
  $("#toastBody").textContent = msg;
  const t = new bootstrap.Toast($("#appToast"), { delay: 2500 });
  t.show();
}

// ===============================
// MODAL DATA TABLE (reuse) + XLSX
// ===============================
let __dataModalState = {
  title: "",
  filename: "export.xlsx",
  rows: [],
  columns: [] // [{key,label}]
};

function ensureDataTableModal(){
  if(document.getElementById("modalDataTable")) return;

  const wrap = document.createElement("div");
  wrap.innerHTML = `
  <div class="modal fade" id="modalDataTable" tabindex="-1" aria-hidden="true">
    <div class="modal-dialog modal-xl modal-dialog-scrollable">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="dtmTitle">Data</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
        </div>

        <div class="modal-body">
          <div class="d-flex flex-wrap gap-2 align-items-center mb-2">
            <div class="small text-muted" id="dtmMeta">—</div>
            <div class="ms-auto d-flex gap-2">
              <button class="btn btn-outline-success btn-sm" id="dtmExport">
                <i class="bi bi-file-earmark-spreadsheet"></i> Export XLSX
              </button>
            </div>
          </div>

          <div class="table-responsive">
            <table class="table table-sm table-bordered align-middle table-nowrap">
              <thead class="table-light">
                <tr id="dtmHead"></tr>
              </thead>
              <tbody id="dtmBody"></tbody>
            </table>
          </div>

          <div class="small text-muted">Total baris: <span id="dtmCount">0</span></div>
        </div>
      </div>
    </div>
  </div>
  `;
  document.body.appendChild(wrap.firstElementChild);

  // bind export sekali
  document.getElementById("dtmExport").addEventListener("click", ()=>{
    const cols = __dataModalState.columns || [];
    const rows = __dataModalState.rows || [];
    if(!rows.length) return toast("Tidak ada data untuk diexport.");

    // export sebagai json dengan urutan kolom sesuai columns
    const data = rows.map(r=>{
      const o = {};
      for(const c of cols){
        o[c.label] = (r && Object.prototype.hasOwnProperty.call(r, c.key)) ? r[c.key] : "";
      }
      return o;
    });

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "data");
    XLSX.writeFile(wb, __dataModalState.filename || "export.xlsx");
  });
}

function showDataTableModal({ title, meta, filename, columns, rows }){
  ensureDataTableModal();

  __dataModalState = {
    title: title || "Data",
    filename: filename || "export.xlsx",
    rows: Array.isArray(rows) ? rows : [],
    columns: Array.isArray(columns) ? columns : []
  };

  document.getElementById("dtmTitle").textContent = __dataModalState.title;
  document.getElementById("dtmMeta").textContent = meta || "—";
  document.getElementById("dtmCount").textContent = String(__dataModalState.rows.length);

  // head
  const head = document.getElementById("dtmHead");
  head.innerHTML = "";
  for(const c of __dataModalState.columns){
    const th = document.createElement("th");
    th.textContent = c.label;
    head.appendChild(th);
  }

  // body
  const body = document.getElementById("dtmBody");
  body.innerHTML = "";
  for(const r of __dataModalState.rows){
    const tr = document.createElement("tr");
    tr.innerHTML = __dataModalState.columns.map(c=>{
      const v = (r && Object.prototype.hasOwnProperty.call(r, c.key)) ? r[c.key] : "";
      return `<td>${escapeHtml(String(v ?? ""))}</td>`;
    }).join("");
    body.appendChild(tr);
  }

  const m = new bootstrap.Modal(document.getElementById("modalDataTable"));
  m.show();
}

function escapeHtml(s){
  return String(s)
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

function setNetBadge(){
  const online = navigator.onLine;
  const badge = $("#netBadge");
  if(!badge) return;
  badge.textContent = online ? "Online" : "Offline";
  badge.classList.toggle("badge-online", online);
  badge.classList.toggle("badge-offline", !online);
  badge.onclick = null; // jangan tampilkan GAS URL
  if ($("#btnSync")) $("#btnSync").classList.toggle("d-none", !online || !state.user);
}
window.addEventListener("online", ()=>{ setNetBadge(); syncQueue().catch(()=>{}); });
window.addEventListener("offline", setNetBadge);

async function migrateNilaiTanggalIfNeeded(){
  const db = await dbPromise;
  const all = await db.getAll("nilai");
  if(!all.length) return;

  let changed = 0;
  for(const r of all){
    // migrasi jika belum punya tanggal_ts atau tanggal masih format lama
    const need = (!("tanggal_ts" in r)) || (typeof r.tanggal === "string" && /^\d{1,2}\/\d{1,2}\/\d{4}/.test(r.tanggal));
    if(!need) continue;

    const fixed = normalizeNilaiRow(r);
    // pastikan id tetap sama
    fixed.id = r.id;

    await db.put("nilai", fixed);
    changed++;
  }

  if(changed){
    console.log(`migrateNilaiTanggalIfNeeded: updated ${changed} rows`);
  }
}

// ---------- IndexedDB ----------
const dbPromise = openDB("karyamas_transkrip_db", 3, {
  upgrade(db, oldVersion, newVersion, tx){
    // ✅ Gunakan tx (versionchange transaction) yang diberikan oleh idb

    if(!db.objectStoreNames.contains("masters")){
      db.createObjectStore("masters", { keyPath:"key" });
    }

    if(!db.objectStoreNames.contains("nilai")){
      const nilai = db.createObjectStore("nilai", { keyPath:"id" });
      try { nilai.createIndex("by_tahun", "tahun"); } catch(e) {}
      try { nilai.createIndex("by_nik", "nik"); } catch(e) {}
      try { nilai.createIndex("by_jenis", "jenis_pelatihan"); } catch(e) {}
    }else{
      // ✅ pastikan index ada (Chrome-safe)
      const store = tx.objectStore("nilai");
      if(!store.indexNames.contains("by_tahun")) { try{ store.createIndex("by_tahun","tahun"); }catch(e){} }
      if(!store.indexNames.contains("by_nik")) { try{ store.createIndex("by_nik","nik"); }catch(e){} }
      if(!store.indexNames.contains("by_jenis")) { try{ store.createIndex("by_jenis","jenis_pelatihan"); }catch(e){} }
    }

    if(!db.objectStoreNames.contains("users")){
      db.createObjectStore("users", { keyPath:"username" });
    }
    if(!db.objectStoreNames.contains("queue")){
      db.createObjectStore("queue", { keyPath:"qid" });
    }
    if(!db.objectStoreNames.contains("failed")){
      db.createObjectStore("failed", { keyPath:"fid" });
    }
    if(!db.objectStoreNames.contains("settings")){
      db.createObjectStore("settings", { keyPath:"key" });
    }
  }
});


async function dbGet(store, key){ return (await dbPromise).get(store, key); }
async function dbPut(store, val){ return (await dbPromise).put(store, val); }
async function dbDel(store, key){ return (await dbPromise).delete(store, key); }
async function dbAll(store){ return (await dbPromise).getAll(store); }
async function dbClear(store){ return (await dbPromise).clear(store); }

async function loadMastersFromDB(){
  const rows = await dbAll("masters");
  const map = Object.fromEntries(rows.map(r=>[r.key, r.data]));
  state.masters.peserta = map.peserta || [];
  state.masters.materi = map.materi || [];
  state.masters.pelatihan = map.pelatihan || [];
  state.masters.bobot = map.bobot || [];
  state.masters.predikat = map.predikat || [];

  normalizeMastersPesertaInState();
}

async function saveMastersToDB(){
  await dbPut("masters", { key:"peserta", data: state.masters.peserta });
  await dbPut("masters", { key:"materi", data: state.masters.materi });
  await dbPut("masters", { key:"pelatihan", data: state.masters.pelatihan });
  await dbPut("masters", { key:"bobot", data: state.masters.bobot });
  await dbPut("masters", { key:"predikat", data: state.masters.predikat });
}

// ---------- Session (Remember login 30 days) ----------
const SESSION_KEY = "session_user_v1";
const SESSION_DAYS = 30;

function nowIso(){ return new Date().toISOString(); }

function addDaysIso(days){
  const d = new Date();
  d.setDate(d.getDate() + days);
  return d.toISOString();
}

async function saveSession(user){
  // user: { username, role, name, must_change }
  const payload = {
    user: {
      username: user.username,
      role: user.role,
      name: user.name,
      must_change: !!user.must_change
    },
    created_at: nowIso(),
    expires_at: addDaysIso(SESSION_DAYS)
  };
  await dbPut("settings", { key: SESSION_KEY, value: payload });
}

async function loadSession(){
  const row = await dbGet("settings", SESSION_KEY);
  const sess = row?.value;
  if(!sess || !sess.user || !sess.expires_at) return null;

  const exp = new Date(sess.expires_at).getTime();
  if(Number.isNaN(exp) || Date.now() > exp){
    // expired → clean
    await clearSession();
    return null;
  }
  return sess;
}

async function clearSession(){
  await dbDel("settings", SESSION_KEY);
}

async function restoreSessionIntoState(){
  const sess = await loadSession();
  if(!sess) return false;
  state.user = sess.user;
  return true;
}

async function fixSessionMustChangeFromLocal(){
  if(!state.user?.username) return;
  try{
    const uLocal = await dbGet("users", state.user.username);
    if(!uLocal) return;

    // Kalau session bilang must_change true tapi di local users sudah false → benarkan
    if(state.user.must_change && uLocal.must_change === false){
      state.user.must_change = false;
      await saveSession(state.user);
    }
  }catch(e){
    console.warn("fixSessionMustChangeFromLocal failed:", e);
  }
}



// ---------- Crypto (SHA-256 hash) ----------
async function sha256Hex(str){
  // ✅ Jika secure context OK, pakai WebCrypto
  if(globalThis.crypto && crypto.subtle && globalThis.isSecureContext){
    const enc = new TextEncoder().encode(str);
    const buf = await crypto.subtle.digest("SHA-256", enc);
    return [...new Uint8Array(buf)].map(b=>b.toString(16).padStart(2,"0")).join("");
  }

  // ✅ Fallback JS SHA-256 (minimal, untuk kasus Chrome file:// / non-secure)
  // Sumber algoritma: implementasi ringkas SHA-256 (tanpa dependency)
  function rrot(n,x){ return (x>>>n) | (x<<(32-n)); }
  function toHex(n){ return (n>>>0).toString(16).padStart(8,"0"); }

  const msg = new TextEncoder().encode(str);
  const l = msg.length * 8;

  const with1 = new Uint8Array(((msg.length + 9 + 63) >> 6) << 6);
  with1.set(msg, 0);
  with1[msg.length] = 0x80;

  const dv = new DataView(with1.buffer);
  dv.setUint32(with1.length - 4, l >>> 0, false);
  dv.setUint32(with1.length - 8, Math.floor(l / 2**32) >>> 0, false);

  const K = [
    0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,
    0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,
    0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,
    0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,
    0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,
    0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,
    0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,
    0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2
  ];

  let h0=0x6a09e667,h1=0xbb67ae85,h2=0x3c6ef372,h3=0xa54ff53a,h4=0x510e527f,h5=0x9b05688c,h6=0x1f83d9ab,h7=0x5be0cd19;

  const W = new Uint32Array(64);

  for(let i=0;i<with1.length;i+=64){
    for(let t=0;t<16;t++) W[t] = dv.getUint32(i + t*4, false);
    for(let t=16;t<64;t++){
      const s0 = rrot(7,W[t-15]) ^ rrot(18,W[t-15]) ^ (W[t-15]>>>3);
      const s1 = rrot(17,W[t-2]) ^ rrot(19,W[t-2]) ^ (W[t-2]>>>10);
      W[t] = (W[t-16] + s0 + W[t-7] + s1) >>> 0;
    }

    let a=h0,b=h1,c=h2,d=h3,e=h4,f=h5,g=h6,h=h7;

    for(let t=0;t<64;t++){
      const S1 = rrot(6,e) ^ rrot(11,e) ^ rrot(25,e);
      const ch = (e & f) ^ (~e & g);
      const temp1 = (h + S1 + ch + K[t] + W[t]) >>> 0;
      const S0 = rrot(2,a) ^ rrot(13,a) ^ rrot(22,a);
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


// ---------- GAS helpers (POST text/plain first, fallback JSONP) ----------
async function gasCall(action, payload = {}) {
  const url = getGasUrl();
  if (!url || /PASTE_YOUR_GAS_WEBAPP_URL/i.test(url)) {
  throw new Error("GAS WebApp URL belum valid (hardcode). Periksa konstanta GAS_URL di app.js.");
}

  // 1) Coba POST text/plain (lebih stabil di Chrome mobile)
  try{
    return await gasPostPlain(url, action, payload, { timeoutMs: 25000 });
  }catch(e){
    console.warn("gasPostPlain failed, fallback JSONP:", e);
  }

  // 2) Fallback: JSONP (cara lama)
  return await gasJsonp(url, action, payload);
}

// (Ping dihapus / tidak dipakai lagi)
// async function gasPing(){ return await gasCall("ping", {}); }

async function gasPostPlain(GAS_URL, action, payload, { timeoutMs = 25000 } = {}){
  const ctrl = new AbortController();
  const t = setTimeout(()=>ctrl.abort(), timeoutMs);

  try{
    const res = await fetch(GAS_URL, {
      method: "POST",
      headers: {
        // kunci di sini: text/plain (metode yang Anda bilang paling aman di mobile)
        "Content-Type": "text/plain;charset=utf-8"
      },
      body: JSON.stringify({ action, payload: payload || {} }),
      cache: "no-store",
      signal: ctrl.signal
    });

    // kalau server balas non-200, tetap coba baca teks untuk error
    const txt = await res.text();
    let data = null;
    try{ data = JSON.parse(txt || "{}"); }catch(_){}

    if(!res.ok){
      throw new Error(`HTTP ${res.status} ${res.statusText}` + (txt ? ` | ${txt.slice(0,200)}` : ""));
    }
    if(!data || data.ok === false){
      throw new Error((data && data.error) || "GAS error");
    }
    return data;
  }catch(err){
    if(String(err?.name||"") === "AbortError"){
      throw new Error("Request timeout");
    }
    throw err;
  }finally{
    clearTimeout(t);
  }
}

function gasJsonp(GAS_URL, action, payload){
  return new Promise((resolve, reject)=>{
    const cb = "cb_"+Date.now()+"_"+Math.random().toString(16).slice(2);

    let script = null;
    let done = false;

    const timeout = setTimeout(()=>{
      cleanup();
      reject(new Error("JSONP timeout"));
    }, 25000);

    window[cb] = (data)=>{
      cleanup();
      if(!data || data.ok===false) return reject(new Error((data && data.error) || "GAS error"));
      resolve(data);
    };

    function cleanup(){
      if(done) return;
      done = true;

      clearTimeout(timeout);

      try{ delete window[cb]; }catch(e){ window[cb]=undefined; }

      try{
        if(script && script.parentNode) script.parentNode.removeChild(script);
      }catch(e){}
    }

    const params = new URLSearchParams();
    params.set("action", action);

    // ✅ FIX: jangan double-encode, biarkan URLSearchParams yang encode
    params.set("payload", JSON.stringify(payload||{}));

    params.set("callback", cb);

    // ✅ cache buster (Chrome mobile kadang agresif caching script)
    params.set("_ts", String(Date.now()));

    const url = GAS_URL + (GAS_URL.includes("?") ? "&" : "?") + params.toString();

    script = document.createElement("script");
    script.src = url;
    script.async = true;

    script.onerror = ()=>{
    cleanup();
    reject(new Error("Gagal memuat GAS (JSONP). URL: " + GAS_URL + " | Pastikan WebApp publik & URL deployment terbaru."));
  };


    document.body.appendChild(script);
  });
}


// ---------- Auth ----------
async function ensureDefaultUsers(){
  // Ensure admin exists in local users store for offline login scenario.
  const db = await dbPromise;
  const admin = await db.get("users","admin");
  if(!admin){
    await db.put("users", {
      username:"admin",
      role:"admin",
      name:"Administrator",
      pass_hash: await sha256Hex("123456"),
      must_change:false
    });
  }
}

async function syncUsersFromMasterPeserta(){
  // create local user accounts for each peserta if not exist yet
  const db = await dbPromise;
  for(const p of state.masters.peserta){
    const username = String(p.nik||"").trim();
    if(!username) continue;
    const u = await db.get("users", username);
    if(!u){
      await db.put("users", {
        username,
        role:"user",
        name: p.nama || ("Peserta "+username),
        pass_hash: await sha256Hex(username),
        must_change:true
      });
    }
  }
}

async function login(username, password){
  username = String(username||"").trim();
  const passHash = await sha256Hex(password);
  const isDefault = (username !== "admin") && (passHash === await sha256Hex(username));

  // 1) SERVER FIRST (kalau online)
  if(navigator.onLine){
    try{
      const res = await gasCall("login", { username, pass_hash: passHash });

      // cache ke lokal (agar bisa offline berikutnya)
      await dbPut("users", {
        username,
        role: res.user.role,
        name: res.user.name,
        pass_hash: res.user.pass_hash,
        must_change: !!res.user.must_change
      });

      state.user = {
        username,
        role: res.user.role,
        name: res.user.name,
        must_change: (isDefault || !!res.user.must_change)
      };
      return { source:"gas", user: state.user };
    }catch(e){
      // kalau gagal (misal jaringan jelek / server error), lanjut fallback lokal
      console.warn("login gas failed, fallback local:", e);
    }
  }

  // 2) FALLBACK LOCAL
  const uLocal = await dbGet("users", username);
  if(uLocal && uLocal.pass_hash === passHash){
    state.user = {
      username,
      role: uLocal.role,
      name: uLocal.name,
      must_change: (isDefault || !!uLocal.must_change)
    };
    return { source:"local", user: state.user };
  }

  throw new Error("Login gagal. Jika user baru, tarik master dulu atau pastikan online.");
}


// Pastikan akun user ada di GAS (khusus untuk user yang dibuat offline dari master peserta)
async function ensureUserExistsOnGAS(username){
  if(!navigator.onLine) return;
  username = String(username||"").trim();
  if(!username) return;

  const u = await dbGet("users", username);
  if(!u) return; // kalau tidak ada di lokal juga, biarkan error

  // Kirim info user minimum untuk dibuat/diupdate di GAS
  // NOTE: Butuh endpoint GAS action: "upsert_user" (lihat catatan di bawah)
  await gasCall("upsert_user", {
    username: u.username,
    name: u.name || "",
    role: u.role || "user",
    pass_hash: u.pass_hash || "",
    must_change: !!u.must_change
  });
}

async function changePassword(newPass){
  const newHash = await sha256Hex(newPass);

  if(!state.user?.username) throw new Error("Belum login.");

  // Wajib online karena tidak boleh queue
  if(!navigator.onLine){
    throw new Error("Ganti password wajib Online (tidak memakai antrian).");
  }

  const username = state.user.username;

  // 1) Pastikan user ada di server
  //    Manfaatkan login_() yang auto-create user jika belum ada.
  //    Untuk admin: login_() akan auto-create admin juga.
  //    Catatan: kita kirim pass_hash yang saat ini diketahui dari lokal (atau default).
  let currentLocal = await dbGet("users", username);

  // Jika belum ada di lokal (kasus jarang), buat perkiraan minimal:
  if(!currentLocal){
    const guessedHash = await sha256Hex(username === "admin" ? "123456" : username);
    currentLocal = {
      username,
      role: (username === "admin") ? "admin" : "user",
      name: state.user.name || (username === "admin" ? "Administrator" : ("Peserta " + username)),
      pass_hash: guessedHash,
      must_change: (username !== "admin")
    };
    await dbPut("users", currentLocal);
  }

  // Pastikan ada di server dengan memanggil login (auto-create).
  // Jika password saat ini di server beda (karena user sudah pernah ganti),
  // login ini bisa gagal. Karena itu kita bungkus try/catch:
  try{
    await gasCall("login", { username, pass_hash: currentLocal.pass_hash });
  }catch(e){
    // Kalau gagal, kita tetap coba change_password langsung.
    // Karena user mungkin sudah ada di server, hanya hash lokal tidak cocok.
    console.warn("ensure via login failed, try change_password anyway:", e);
  }

  // 2) Update password di server (langsung)
  await gasCall("change_password", { username, pass_hash: newHash });

  // 3) Update local cache
  currentLocal.pass_hash = newHash;
  currentLocal.must_change = false;
  await dbPut("users", currentLocal);

  state.user.must_change = false;

    // ✅ penting: update session agar refresh tidak memunculkan popup lagi
  try{
    await saveSession(state.user);
  }catch(e){
    console.warn("saveSession after changePassword failed:", e);
  }

  toast("Password berhasil diubah (server & lokal).");
}

async function purgePasswordQueue(){
  const db = await dbPromise;
  const items = await db.getAll("queue");
  for(const it of items){
    if(it.type === "change_password" || it.type === "upsert_user"){
      await db.delete("queue", it.qid);
    }
  }
  await refreshQueueBadge();
}


// ---------- Queue & Sync ----------
function qid(){ return "Q"+Date.now()+"_"+Math.random().toString(16).slice(2); }

async function enqueue(type, payload){
  await dbPut("queue", { qid: qid(), type, payload, created_at: new Date().toISOString() });
  await refreshQueueBadge();
}

async function refreshQueueBadge(){
  const q = await dbAll("queue");
  $("#queueCount").textContent = String(q.length);
  $("#syncState").textContent = navigator.onLine ? "Online" : "Offline";
}

async function syncQueue(){
  if(!navigator.onLine) return;

  const db = await dbPromise;
  const items = await db.getAll("queue");
  if(!items.length) return;

  $("#syncState").textContent = "Syncing…";
  const total = items.length;
  progressStart(total, "Sinkronisasi antrian…");

  let done = 0;
  let failed = 0;

  for(const it of items){
    try{
      await gasCall(it.type, it.payload);

      // sukses → hapus dari queue
      await db.delete("queue", it.qid);
      done++;
    }catch(e){
      console.warn("queue sync failed", it, e);

      failed++;

      // simpan ke failed store, lalu hapus dari queue
      const fid = "F" + Date.now() + "_" + Math.random().toString(16).slice(2);
      await db.put("failed", {
        fid,
        qid: it.qid,
        type: it.type,
        created_at: it.created_at || new Date().toISOString(),
        failed_at: new Date().toISOString(),
        error: (e && e.message) ? String(e.message) : String(e),
        payload_json: JSON.stringify(it.payload || {})
      });

      await db.delete("queue", it.qid);

      // tetap lanjut item berikutnya
      done++;
    }

    if(done % 5 === 0 || done === total){
      progressSet(done, total, `Sync… (${done}/${total}) | Gagal: ${failed}`);
    }
  }

  await refreshQueueBadge();
  $("#syncState").textContent = "Online";
  progressDone(failed ? `Sync selesai (Gagal: ${failed}).` : "Sync selesai.");
  toast(failed ? `Sync selesai. Gagal: ${failed}` : "Sync selesai.");
}

// ---------- UI Routing ----------
function setWhoAmI(){
  const box = $("#whoami");
  if(!state.user){ box.textContent=""; return; }
  box.textContent = `${state.user.name} • ${state.user.role.toUpperCase()} • ${state.user.username}`;
  $("#btnLogout").classList.remove("d-none");
  $("#btnSync").classList.toggle("d-none", !navigator.onLine);
}

function buildMenu(){
  const menu = $("#sideMenu");
  menu.innerHTML = "";
  const add = (id, icon, label)=> {
    const a = document.createElement("a");
    a.href="#";
    a.className="list-group-item list-group-item-action d-flex align-items-center gap-2";
    a.dataset.view=id;
    a.innerHTML = `<i class="bi ${icon}"></i><span>${label}</span>`;
    menu.appendChild(a);
  };

  add("dashboard","bi-speedometer2","Dashboard");
  add("nilai","bi-table","Daftar Nilai");

  if(state.user.role==="admin"){
    add("input","bi-pencil-square","Input Data Nilai");
    add("master","bi-database","Master Data");
    add("setting","bi-gear","Setting");
  } else {
    add("setting_user","bi-gear","Setting");
  }

  menu.addEventListener("click",(e)=>{
    const a = e.target.closest("a[data-view]");
    if(!a) return;
    e.preventDefault();
    renderView(a.dataset.view);
  });
}

function renderView(view){
  if(view==="dashboard") return renderDashboard();
  if(view==="nilai") return renderNilaiList();
  if(view==="input") return renderInputNilai();
  if(view==="master") return renderMaster();
  if(view==="setting") return renderSettingAdmin();
  if(view==="setting_user") return renderSettingUser();
}

function mount(html){
  $("#viewContainer").innerHTML = html;
}

function optYears(){
  const now = new Date().getFullYear();
  const years = [];
  for(let y=now; y>=now-10; y--) years.push(y);
  return years.map(y=>`<option value="${y}">${y}</option>`).join("");
}

function uniq(arr, key){
  const s = new Set();
  const out=[];
  for(const a of arr){ const v = (a[key]||"").toString().trim(); if(v && !s.has(v)){ s.add(v); out.push(v);} }
  return out;
}

async function rebuildPelatihanCache(){
  // gabungkan dari master peserta + data nilai (agar tetap muncul walau peserta belum di-pull)
  const fromPeserta = uniq(state.masters.peserta || [], "jenis_pelatihan");

  let fromNilai = [];
  try{
    const allNilai = await dbAll("nilai");
    fromNilai = uniq(allNilai || [], "jenis_pelatihan");
  }catch(e){}

  const merged = [...new Set([...fromPeserta, ...fromNilai].map(normStr).filter(Boolean))].sort((a,b)=>a.localeCompare(b));

  state.masters.pelatihan = merged.map(nama => ({ nama }));
  await saveMastersToDB(); // sekarang pelatihan ikut tersimpan ✅
}

// =====================
// SAFE MERGE MASTERS (anti ketimpa)
// =====================
function mergeMasters(prev, incoming){
  const base = prev || {};
  const inc = incoming || {};
  return {
    peserta: Array.isArray(inc.peserta) ? inc.peserta : (Array.isArray(base.peserta) ? base.peserta : []),
    materi: Array.isArray(inc.materi) ? inc.materi : (Array.isArray(base.materi) ? base.materi : []),
    pelatihan: Array.isArray(inc.pelatihan) ? inc.pelatihan : (Array.isArray(base.pelatihan) ? base.pelatihan : []),
    bobot: Array.isArray(inc.bobot) ? inc.bobot : (Array.isArray(base.bobot) ? base.bobot : []),
    predikat: Array.isArray(inc.predikat) ? inc.predikat : (Array.isArray(base.predikat) ? base.predikat : []),
  };
}

// =====================
// SAFE MERGE MASTERS (anti ketimpa) - IN PLACE + APPLY
// =====================

// Merge in-place: menjaga object state.masters tidak diganti (menghindari referensi UI/cache ikut rusak)
function mergeMastersInPlace(target, incoming){
  if(!target) return;
  if(!incoming) return;

  for(const k of Object.keys(incoming)){
    const v = incoming[k];

    // kalau array → overwrite aman (master terbaru menang)
    if(Array.isArray(v)){
      target[k] = v;
      continue;
    }

    // kalau object → shallow merge
    if(v && typeof v === "object"){
      target[k] = Object.assign({}, target[k] || {}, v);
      continue;
    }

    // primitive
    target[k] = v;
  }
}

// Helper terpusat: apply masters dari server + simpan + rebuild cache + sync user lokal
async function applyMastersFromServer(mastersIncoming, { rebuildCache=true, syncUsers=true } = {}){
  // pastikan bentuk final selalu lengkap (pakai fallback mergeMasters existing)
  const merged = mergeMasters(state.masters, mastersIncoming);

  // pakai in-place agar state.masters tetap referensi yang sama
  mergeMastersInPlace(state.masters, merged);

  normalizeMastersPesertaInState();

  await saveMastersToDB();
  if(rebuildCache) await rebuildPelatihanCache();
  if(syncUsers) await syncUsersFromMasterPeserta();

  // debug cepat (aktif jika DEBUG_TRX=1)
  dbg("masters.apply", {
    peserta: (state.masters.peserta||[]).length,
    materi: (state.masters.materi||[]).length,
    pelatihan: (state.masters.pelatihan||[]).length,
    bobot: (state.masters.bobot||[]).length,
    predikat: (state.masters.predikat||[]).length,
  });
}

// =====================
// NUMBER PARSER (dukung koma desimal: "89,9" -> 89.9)
// =====================
function numVal(v){
  if(v == null || v === "") return NaN;
  if(typeof v === "number") return v;
  const s = String(v).trim().replace(/\./g, "").replace(",", "."); 
  // note: replace(/\./g,"") aman utk kasus "1.234,5" -> "1234,5" -> "1234.5"
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : NaN;
}

function normalizeJenisForBobot(jenis){
  let s = normStr(jenis);

  // buang pola umum: "... 10 Tahun 2020" atau "... Tahun 2020"
  s = s.replace(/\s+\d+\s*tahun\s+\d{4}\s*$/i, "").trim();
  s = s.replace(/\s*tahun\s+\d{4}\s*$/i, "").trim();

  // kalau setelah itu masih ada angka sisa di belakang, buang (mis. "KLP1 AGRO 10")
  s = s.replace(/\s+\d+\s*$/i, "").trim();

  // rapikan spasi
  s = s.replace(/\s+/g, " ").trim();
  return s;
}


function normStr(v){
  return String(v ?? "").trim();
}

// untuk pencocokan key teks (jenis pelatihan, kategori, dll)
function normKey(v){
  return normStr(v).toLowerCase().replace(/\s+/g," ").trim();
}

// parse angka aman: dukung "89,9"
function num(v){
  if(v == null || v === "") return NaN;
  if(typeof v === "number") return v;
  const s = String(v).trim().replace(/\./g, "").replace(",", ".");
  const x = parseFloat(s);
  return Number.isFinite(x) ? x : NaN;
}

function normNik(v){
  // NIK dari Sheet/XLSX kadang kebaca 11001.0 → normalkan ke 11001
  let s = normStr(v).replace(/\s+/g,"");
  if(/^\d+(\.0+)?$/.test(s)) s = s.replace(/\.0+$/,"");
  return s;
}

// =====================
// NORMALISASI MASTER PESERTA (agar meta transkrip terbaca walau header beda-beda)
// =====================
function _pickAny(obj, keys){
  for(const k of keys){
    if(obj && Object.prototype.hasOwnProperty.call(obj, k)){
      const v = obj[k];
      if(v != null && String(v).trim() !== "") return v;
    }
  }
  return "";
}

function normalizePesertaMasterRow(raw){
  const r = raw || {};

  // buat akses key yang fleksibel (as-is)
  const o = { ...r };

  // fallback: coba juga versi lowercase + underscore (kalau GAS mengirim "Tanggal Presentasi")
  // (kita tidak ubah semua key, cuma ambil nilai)
  const tanggal_presentasi = _pickAny(o, [
    "tanggal_presentasi","Tanggal Presentasi","tanggal presentasi","tgl_presentasi","Tgl Presentasi","tgl presentasi",
    "tanggalPresentasi","TanggalPresentasi","Tanggal_Presentasi"
  ]);

  const judul_presentasi = _pickAny(o, [
    "judul_presentasi","Judul Presentasi","judul presentasi",
    "judul_makalah","Judul Makalah","judul makalah",
    "judulPresentasi","JudulPresentasi","Judul_Presentasi",
    "judulMakalah","JudulMakalah","Judul_Makalah"
  ]);

  const lokasi_ojt = _pickAny(o, [
    "lokasi_ojt","Lokasi OJT","lokasi ojt",
    "lokasi_praktek","Lokasi Praktek","lokasi praktek",
    "lokasi","Lokasi"
  ]);

  const unit = _pickAny(o, [
    "unit","Unit","unit_kerja","Unit Kerja","unit kerja","unitKerja","UnitKerja"
  ]);

  const region = _pickAny(o, [
    "region","Region","wilayah","Wilayah"
  ]);

  // pastikan output konsisten
  return {
    ...o,
    nik: normNik(o.nik),
    nama: normStr(o.nama),
    jenis_pelatihan: normStr(o.jenis_pelatihan),

    lokasi_ojt: normStr(lokasi_ojt),
    unit: normStr(unit),
    region: normStr(region),

    // penting: tanggal_presentasi biarkan RAW (bisa excel serial / Date / ISO string)
    tanggal_presentasi: (tanggal_presentasi ?? ""),
    judul_presentasi: normStr(judul_presentasi),

    // (opsional) kalau ada field lain biarkan tetap
  };
}

function normalizeMastersPesertaInState(){
  if(!Array.isArray(state.masters.peserta)) state.masters.peserta = [];
  state.masters.peserta = state.masters.peserta.map(normalizePesertaMasterRow);
}


// =====================
// TRANSKRIP HELPERS (GLOBAL) - anti duplikat
// =====================
function r1(n){
  const x = (typeof n==="number") ? n : parseFloat(n);
  if(!Number.isFinite(x)) return "";
  return (Math.round(x*10)/10).toFixed(1);
}

function fmtTanggalIndoLong(v){
  // output: "17 Desember 2025"
  const d = parseAnyDate(v);
  if(!d) return "";
  const bulan = [
    "Januari","Februari","Maret","April","Mei","Juni",
    "Juli","Agustus","September","Oktober","November","Desember"
  ];
  return `${String(d.getDate()).padStart(2,"0")} ${bulan[d.getMonth()]} ${d.getFullYear()}`;
}

function pickPredikatFromMaster(v){
  const n = Number.isFinite(v) ? v : numVal(v);
  if(!Number.isFinite(n)) return "-";

  const rules = (state.masters.predikat||[])
    .map(x=>({
      nama: normStr(x.nama),
      min: num(x.min),
      max: num(x.max)
    }))
    .filter(x=>x.nama && Number.isFinite(x.min) && Number.isFinite(x.max));

  if(rules.length){
    const found = rules.find(r => n >= r.min && n <= r.max);
    if(found) return found.nama;
  }

  // fallback default
  if(n>90) return "Sangat Memuaskan";
  if(n>80) return "Memuaskan";
  if(n>=76) return "Baik";
  if(n>=70) return "Kurang";
  return "Sangat Kurang";
}

function getLulusStatus(total){
  const n = (typeof total==="number") ? total : parseFloat(total);
  if(!Number.isFinite(n)) return { lulus:false, label:"TIDAK LULUS", badge:"text-bg-danger" };
  const ok = n >= 70;
  return ok
    ? { lulus:true, label:"LULUS", badge:"text-bg-success" }
    : { lulus:false, label:"TIDAK LULUS", badge:"text-bg-danger" };
}

function jenisDenganTahun(jenis, tahun){
  const j = normStr(jenis);
  const y = String(normStr(tahun));
  if(!j) return y ? `Tahun ${y}` : "-";
  if(!y) return j;
  if(/tahun\s+\d{4}/i.test(j)) return j;
  return `${j} Tahun ${y}`;
}

function getKategoriMateri(materiKode, materiNama){
  const kode = normStr(materiKode);
  const nama = normStr(materiNama);
  let m = null;

  // 1) Prioritas: master materi
  if(kode){
    m = (state.masters.materi||[]).find(x => normKey(x.kode) === normKey(kode)) || null;
  }
  if(!m && nama){
    m = (state.masters.materi||[]).find(x => normKey(x.nama) === normKey(nama)) || null;
  }

  let kat = normStr(m?.kategori);

  // 2) Jika master tidak lengkap → fallback dari KODE / NAMA
  if(!kat){
    const k = kode.toUpperCase();

    // fallback dari prefix kode (sesuaikan jika ada pola lain)
    if(/^MD/.test(k)) kat = "Managerial";
    // ✅ UPDATE: Teknis untuk AG (AGRO) + MI (MILL) + AD (ADMN)
    else if(/^(AG|MI|AD)/.test(k)) kat = "Teknis";
    else if(/^SU/.test(k)) kat = "Support";
    else if(/^KL/.test(k)){
      const nm = nama.toUpperCase();
      if(nm.includes("ON THE JOB") || nm.includes("OJT")) kat = "OJT";
      else if(nm.includes("PRESENTASI")) kat = "Presentasi";
      else kat = "OJT"; // default KL biasanya terkait OJT/presentasi
    }else{
      // fallback dari nama bila kode tidak jelas
      const nm = nama.toUpperCase();
      if(nm.includes("ON THE JOB") || nm.includes("OJT")) kat = "OJT";
      else if(nm.includes("PRESENTASI")) kat = "Presentasi";
      else kat = ""; // tidak ketahuan → bobot 0
    }
  }

  // 3) Normalisasi output agar konsisten
  if(!kat) return "";
  const kk = kat.toLowerCase();
  if(kk.includes("man")) return "Managerial";
  if(kk.includes("tek")) return "Teknis";
  if(kk.includes("sup")) return "Support";
  if(kk.includes("ojt")) return "OJT";
  if(kk.includes("pres")) return "Presentasi";
  return kat;
}

// =====================
// BOBOT LOOKUP HELPER (lebih toleran untuk MILL/ADMN)
// =====================
function findBobotRowByJenis(jenisInput){
  const src = (state.masters.bobot || []);
  if(!src.length) return null;

  const j0 = normStr(jenisInput);
  const jClean = normalizeJenisForBobot(j0);

  const k0 = normKey(j0);
  const k1 = normKey(jClean);

  // 1) exact match (paling ketat)
  let hit = src.find(x=>{
    const a0 = normKey(x.jenis_pelatihan);
    const a1 = normKey(normalizeJenisForBobot(x.jenis_pelatihan));
    return a0===k0 || a0===k1 || a1===k0 || a1===k1;
  });
  if(hit) return hit;

  // 2) fuzzy match: contains / substring (toleran "KLP1 MILL" vs "KLP1 MILL 04 Tahun 2022")
  hit = src.find(x=>{
    const a1 = normKey(normalizeJenisForBobot(x.jenis_pelatihan));
    return (a1 && (a1.includes(k1) || k1.includes(a1)));
  });
  if(hit) return hit;

  // 3) fuzzy match: contains versi raw (kadang angka "04" masih menempel)
  hit = src.find(x=>{
    const a0 = normKey(x.jenis_pelatihan);
    return (a0 && (a0.includes(k0) || k0.includes(a0)));
  });
  return hit || null;
}

function getBobotByJenis(jenis){
  const j0 = normStr(jenis);

  // ✅ cari bobot dengan lookup yang lebih toleran
  const b = findBobotRowByJenis(j0);

  // parse angka (dukung koma/format ribuan)
  let obj = {
    managerial: numVal(b?.managerial) || 0,
    teknis:     numVal(b?.teknis) || 0,
    support:    numVal(b?.support) || 0,
    ojt:        numVal(b?.ojt) || 0,
    presentasi: numVal(b?.presentasi) || 0
  };

  // ✅ fallback bila master bobot belum ada / tidak ketemu
  // (agar transkrip MILL & ADMN tidak 0 semua)
  const sum = obj.managerial + obj.teknis + obj.support + obj.ojt + obj.presentasi;
  if(sum <= 0){
    // default aman (silakan ubah sesuai kebijakan TC)
    obj = { managerial: 20, teknis: 60, support: 10, ojt: 5, presentasi: 5 };
  }

  return obj;
}

function getBobotPercentForKategori(bobotObj, kategori){
  const k = normStr(kategori).toLowerCase();
  if(k === "managerial") return bobotObj.managerial;
  if(k === "teknis") return bobotObj.teknis;
  if(k === "support") return bobotObj.support;
  if(k === "ojt") return bobotObj.ojt;
  if(k === "presentasi") return bobotObj.presentasi;
  return 0;
}

// base filter untuk transkrip: abaikan filter test (supaya PDF/preview selalu ambil Final)
function filterBaseForTranscript(allNilai, f){
  return (allNilai||[]).filter(r=>{
    const rNik = normNik(r.nik);
    const uNik = normNik(state.user?.username);

    if(state.user.role !== "admin" && rNik !== uNik) return false;

    if(f?.tahun && String(normStr(r.tahun)) !== String(normStr(f.tahun))) return false;
    if(f?.jenis && normStr(r.jenis_pelatihan) !== normStr(f.jenis)) return false;

    if(f?.materi){
      const mk = materiKeyFromRow(r);
      if(mk !== normStr(f.materi)) return false;
    }

    if(state.user.role === "admin" && f?.nik && rNik !== normNik(f.nik)) return false;

    return true;
  });
}

// hitung rerata kelas per materi dari Final rows
function buildAvgMap(finalRows){
  const avgMap = new Map(); // key -> {total,n}
  for(const r of (finalRows||[])){
    const jenis = normStr(r.jenis_pelatihan);
    const tahun = String(normStr(r.tahun));
    const mkey = materiKeyFromRow(r);
    if(!jenis || !tahun || !mkey) continue;
    const key = `${jenis}||${tahun}||${mkey}`;
    if(!avgMap.has(key)) avgMap.set(key, { total:0, n:0 });

    const v = (typeof r.nilai==="number") ? r.nilai : parseFloat(r.nilai);
    if(Number.isFinite(v)){
      const o = avgMap.get(key);
      o.total += v;
      o.n += 1;
    }
  }
  return avgMap;
}

function getRerataKelasFromMap(avgMap, jenis, tahun, materiKey){
  const key = `${normStr(jenis)}||${String(normStr(tahun))}||${normStr(materiKey)}`;
  const o = avgMap.get(key);
  if(!o || !o.n) return null;
  return (o.total / o.n);
}

function buildTranscriptData({ nik, finalRowsAll }){
  const mine = (finalRowsAll||[]).filter(r => normNik(r.nik) === normNik(nik));
  if(!mine.length) return null;

  const pesertaMaster = (state.masters.peserta||[]).find(p => normNik(p.nik) === normNik(nik)) || {};
    const contoh = mine[0] || {};
    const peserta = {
      ...pesertaMaster,
      // fallback minimal dari data nilai bila master kosong
      nik: pesertaMaster.nik ?? nik,
      nama: pesertaMaster.nama ?? contoh.nama ?? "",
      jenis_pelatihan: pesertaMaster.jenis_pelatihan ?? contoh.jenis_pelatihan ?? "",
      lokasi_ojt: pesertaMaster.lokasi_ojt ?? contoh.lokasi_ojt ?? "",
      unit: pesertaMaster.unit ?? contoh.unit ?? "",
      region: pesertaMaster.region ?? contoh.region ?? "",
      tanggal_presentasi: pesertaMaster.tanggal_presentasi ?? contoh.tanggal_presentasi ?? "",
      judul_presentasi: pesertaMaster.judul_presentasi ?? contoh.judul_presentasi ?? "",
    };
  const tahun = mine[0]?.tahun ?? peserta.tahun ?? state.filters.tahun ?? "";
  const jenisRaw = mine[0]?.jenis_pelatihan || peserta.jenis_pelatihan || "-";
  const jenisLabel = jenisDenganTahun(jenisRaw, tahun);

  const avgMap = buildAvgMap(finalRowsAll);
  const bobotObj = getBobotByJenis(jenisRaw.replace(/\s+tahun\s+\d{4}\s*$/i,"").trim());

    const listSorted = mine.slice().sort(cmpMateriKodeTranskrip);

  // ✅ HITUNG JUMLAH MATERI PER KATEGORI (untuk pembagian bobot per materi)
  const catCount = {};
  for(const rr of listSorted){
    const kk = getKategoriMateri(rr.materi_kode, rr.materi_nama);
    const kNorm = normStr(kk);
    if(!kNorm) continue;
    catCount[kNorm] = (catCount[kNorm] || 0) + 1;
  }

  const lines = [];
  let totalWeighted = 0;

  dbg("trx.catCount", { nik, jenisRaw, tahun, catCount });

  for(const r of listSorted){
    const mKey = materiKeyFromRow(r);
    const rk = getRerataKelasFromMap(avgMap, jenisRaw, tahun, mKey);

    const poin = (typeof r.nilai==="number") ? r.nilai : parseFloat(r.nilai);
    const poinOk = Number.isFinite(poin) ? poin : 0;

    const kategori = getKategoriMateri(r.materi_kode, r.materi_nama);
    const bobotPct = getBobotPercentForKategori(bobotObj, kategori);

    // ✅ bobot per materi = bobot kategori / jumlah materi kategori
    const nKat = catCount[normStr(kategori)] || 0;

    // rumus: poin * bobotPct / nKat / 100
    const nilaiWeighted = (nKat > 0 && bobotPct > 0)
      ? (poinOk * bobotPct / nKat / 100)
      : 0;

    if(Number.isFinite(nilaiWeighted)) totalWeighted += nilaiWeighted;

    lines.push({
      materi_kode: r.materi_kode || "",
      materi_nama: r.materi_nama || "",
      rerata_kelas: rk,
      poin: poinOk,
      nilai: nilaiWeighted
    });

    dbg("trx.row", {
      nik,
      jenisRaw,
      tahun,
      materi_kode: r.materi_kode,
      materi_nama: r.materi_nama,
      kategori,
      poinOk,
      bobotPct,
      nKat,
      nilaiWeighted
    });
  }


  return { nik, peserta, tahun, jenisRaw, jenisLabel, lines, totalWeighted };
      
}

// =====================
// NORMALISASI TANGGAL & NILAI ROW
// =====================
function _pad2(n){ return String(n).padStart(2,"0"); }

function toIsoDateLocal(d){
  // YYYY-MM-DD (local)
  const y = d.getFullYear();
  const m = _pad2(d.getMonth()+1);
  const day = _pad2(d.getDate());
  return `${y}-${m}-${day}`;
}

function toIsoDateTimeLocal(d){
  // YYYY-MM-DDTHH:mm:ss (local)
  const y = d.getFullYear();
  const m = _pad2(d.getMonth()+1);
  const day = _pad2(d.getDate());
  const hh = _pad2(d.getHours());
  const mm = _pad2(d.getMinutes());
  const ss = _pad2(d.getSeconds());
  return `${y}-${m}-${day}T${hh}:${mm}:${ss}`;
}

function parseAnyDate(v){
  // return Date | null
  if(v == null || v === "") return null;

  // Date object
  if(v instanceof Date && !Number.isNaN(v.getTime())) return v;

  // Excel serial number (common when importing XLSX)
  if(typeof v === "number" && Number.isFinite(v)){
    // Excel serial date: 25569 = 1970-01-01
    // Handles both date and date-time serials
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const d = new Date(ms);
    if(!Number.isNaN(d.getTime())) return d;
  }

  const s = String(v).trim();
  if(!s) return null;

  // ISO / YYYY-MM-DD / YYYY-MM-DDTHH:mm:ss
  if(/^\d{4}-\d{2}-\d{2}/.test(s)){
    const d = new Date(s);
    if(!Number.isNaN(d.getTime())) return d;

    // fallback manual parse YYYY-MM-DD
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if(m){
      const d2 = new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));
      if(!Number.isNaN(d2.getTime())) return d2;
    }
  }

  // DD/MM/YYYY or MM/DD/YYYY (auto) + optional time
  const m1 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if(m1){
    let a = Number(m1[1]); // could be DD or MM
    let b = Number(m1[2]); // could be MM or DD
    const yy = Number(m1[3]);
    const hh = Number(m1[4] || 0);
    const mi = Number(m1[5] || 0);
    const ss = Number(m1[6] || 0);

    // Heuristic:
    // - if b > 12 => assume MM/DD (US) e.g. 9/28/2020
    // - else default DD/MM (ID) e.g. 15/12/2025
    let dd, mm;
    if(b > 12){
      mm = a; dd = b;   // MM/DD/YYYY
    }else{
      dd = a; mm = b;   // DD/MM/YYYY
    }

    const d = new Date(yy, mm-1, dd, hh, mi, ss);
    if(!Number.isNaN(d.getTime())) return d;
  }


  return null;
}

function normalizeNilaiRow(input){
  // memastikan tanggal konsisten + tahun valid
  const r = { ...(input || {}) };

  // trim string fields
  r.jenis_pelatihan = normStr(r.jenis_pelatihan);
  r.test_type = normStr(r.test_type);
  r.nik = normNik(r.nik);
  r.nama = normStr(r.nama);
  r.materi_kode = normStr(r.materi_kode);
  r.materi_nama = normStr(r.materi_nama);

  // nilai number
  if(typeof r.nilai !== "number") r.nilai = parseFloat(r.nilai);
  if(Number.isNaN(r.nilai)) r.nilai = null;

  // tanggal normalize
  const d = parseAnyDate(r.tanggal);
  if(d){
    // simpan iso datetime agar "unik" bila ada beda detik (sesuai kebutuhan Anda)
    r.tanggal = toIsoDateTimeLocal(d);  // contoh: 2025-12-17T15:14:57
    r.tanggal_ts = d.getTime();
  }else{
    // fallback: kosongkan agar tidak bikin filter/sort kacau
    r.tanggal = "";
    r.tanggal_ts = 0;
  }

  // tahun: kalau kosong / tidak valid → ambil dari tanggal
  const y = parseInt(r.tahun, 10);
  if(Number.isNaN(y) || y < 1900 || y > 2100){
    if(d) r.tahun = d.getFullYear();
    else r.tahun = new Date().getFullYear();
  }else{
    r.tahun = y;
  }

  return r;
}

function fmtTanggalDisplay(tanggal){
  // tampilkan DD/MM/YYYY untuk UI (apapun format internal)
  const d = parseAnyDate(tanggal);
  if(!d) return "";
  return `${_pad2(d.getDate())}/${_pad2(d.getMonth()+1)}/${d.getFullYear()}`;
}


function optPelatihan(){
  // 1) Prioritas: master pelatihan (bila ada)
  let items = Array.isArray(state.masters.pelatihan) ? state.masters.pelatihan : [];
  items = items.map(x => (typeof x === "string" ? { nama:x } : x)).filter(x=>normStr(x?.nama));

  // 2) Fallback: ambil unik dari master peserta
  if(!items.length){const u = uniq(state.masters.peserta || [], "jenis_pelatihan"); items = u.map(j=>({ nama:j })); }
  return [`<option value="">Semua</option>`].concat(items.map(p=>`<option value="${p.nama}">${p.nama}</option>`)).join("");
}


function optTestTypes(){
  const items = ["PreTest","PostTest","Final","OJT","Presentasi"];
  return [`<option value="">Semua</option>`].concat(items.map(t=>`<option value="${t}">${t}</option>`)).join("");
}

function materiKeyFromRow(r){
  // key utama: kode jika ada, jika tidak pakai nama
  const k = normStr(r?.materi_kode) || normStr(r?.materi_nama);
  return k;
}

// =====================
// SORTING MATERI: AG -> SU -> MD -> KL (rapi untuk transkrip)
// =====================
function materiPrefixRank(kode){
  const k = normStr(kode).toUpperCase();

  // ✅ UPDATE: semua prefix Teknis diletakkan paling atas
  if(k.startsWith("AG")) return 1; // AGRO teknis
  if(k.startsWith("MI")) return 1; // MILL teknis
  if(k.startsWith("AD")) return 1; // ADMN teknis

  // urutan berikutnya
  if(k.startsWith("SU")) return 2;
  if(k.startsWith("MD")) return 3;
  if(k.startsWith("KL")) return 4;

  // lain-lain taruh paling akhir
  return 99;
}


function cmpMateriKodeTranskrip(a, b){
  const ak = normStr(a?.materi_kode).toUpperCase();
  const bk = normStr(b?.materi_kode).toUpperCase();

  const ra = materiPrefixRank(ak);
  const rb = materiPrefixRank(bk);
  if(ra !== rb) return ra - rb;

  // jika prefix sama → urutkan berdasarkan kode (AG001, AG002, dst)
  const c = ak.localeCompare(bk, "id", { numeric:true, sensitivity:"base" });
  if(c !== 0) return c;

  // fallback: nama materi
  const an = normStr(a?.materi_nama);
  const bn = normStr(b?.materi_nama);
  return an.localeCompare(bn, "id", { numeric:true, sensitivity:"base" });
}


function optMateriFromNilai(nilaiRows){
  // ambil dari data nilai yang sudah ter-filter (lebih relevan)
  const map = new Map(); // key -> label
  for(const r of (nilaiRows||[])){
    const key = materiKeyFromRow(r);
    if(!key) continue;
    const kode = normStr(r.materi_kode);
    const nama = normStr(r.materi_nama);
    const label = kode && nama ? `${kode} - ${nama}` : (nama || kode || key);
    if(!map.has(key)) map.set(key, label);
  }

  // fallback bila nilai kosong, ambil dari master materi
  if(map.size === 0){
    for(const m of (state.masters.materi||[])){
      const key = normStr(m.kode) || normStr(m.nama);
      if(!key) continue;
      const label = (m.kode && m.nama) ? `${m.kode} - ${m.nama}` : (m.nama || m.kode);
      if(!map.has(key)) map.set(key, label);
    }
  }

  const arr = [...map.entries()]
    .map(([value,label])=>({value,label}))
    .sort((a,b)=>a.label.localeCompare(b.label));

  return [`<option value="">Semua</option>`]
    .concat(arr.map(x=>`<option value="${x.value}">${x.label}</option>`))
    .join("");
}

// ---------- Dashboard ----------
async function renderDashboard(){
  // filter & compute highlights
  const allNilai = await dbAll("nilai");
  const f = state.filters;

    const filtered = allNilai.filter(r=>{
    const rNik = normNik(r.nik);
    const uNik = normNik(state.user?.username);

    if(state.user.role !== "admin" && rNik !== uNik) return false;

    if(f.tahun && String(normStr(r.tahun)) !== String(normStr(f.tahun))) return false;
    if(f.jenis && normStr(r.jenis_pelatihan) !== normStr(f.jenis)) return false;

    if(state.user.role === "admin" && f.nik && rNik !== normNik(f.nik)) return false;

    // ✅ baru
    if(f.test && normStr(r.test_type) !== normStr(f.test)) return false;
    if(f.materi){
      const mk = materiKeyFromRow(r);
      if(mk !== normStr(f.materi)) return false;
    }

    return true;
  });


  // group by peserta for ranking
  const byPeserta = new Map();
  for(const r of filtered){
    const key = r.nik;
    if(!byPeserta.has(key)) byPeserta.set(key, { nik:r.nik, nama:r.nama, jenis:r.jenis_pelatihan, total:0, n:0 });
    const o = byPeserta.get(key);
    if(typeof r.nilai==="number"){ o.total+=r.nilai; o.n+=1; }
  }
  const pesertaArr = [...byPeserta.values()].map(o=>({ ...o, avg: o.n? (o.total/o.n):0 }))
    .sort((a,b)=>b.avg-a.avg);

  const top3 = pesertaArr.slice(0,3);
  const low3 = pesertaArr.slice(-3).reverse();

  // materi avg
  const byMateri = new Map();
  for(const r of filtered){
    const key = r.materi_kode || r.materi_nama;
    if(!key) continue;
    if(!byMateri.has(key)) byMateri.set(key, { kode:r.materi_kode, nama:r.materi_nama, total:0, n:0 });
    const o = byMateri.get(key);
    if(typeof r.nilai==="number"){ o.total+=r.nilai; o.n+=1; }
  }
  const materiArr = [...byMateri.values()].map(o=>({ ...o, avg: o.n? (o.total/o.n):0 }))
    .sort((a,b)=>b.avg-a.avg);
    function pickMateriHighLow(arr){
    const n = arr.length;

    // default
    let nHigh = 0, nLow = 0;

    if(n >= 6){ nHigh = 3; nLow = 3; }
    else if(n === 5){ nHigh = 2; nLow = 3; }
    else if(n === 4){ nHigh = 2; nLow = 2; }
    else if(n === 3){ nHigh = 1; nLow = 2; }
    else if(n === 2){ nHigh = 1; nLow = 1; }
    else if(n === 1){ nHigh = 1; nLow = 1; }
    else { nHigh = 0; nLow = 0; }

    const high = arr.slice(0, nHigh);
    const low = nLow ? arr.slice(Math.max(0, n - nLow)).reverse() : [];
    return { high, low };
  }

  const { high: materiHighList, low: materiLowList } = pickMateriHighLow(materiArr);


  // failed (based on predikat rules, simplified)
  const gagal = pesertaArr.filter(p=>p.avg<70);

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

      <div class="col-12 ${state.user.role!=="admin" ? "d-none":""}">
        <label class="form-label small mb-1">NIK Peserta</label>
        <input id="fNik" class="form-control form-control-sm" placeholder="Kosongkan untuk semua">
      </div>
    </div>

    </div>

    <div class="row g-3">
      <div class="col-12 col-md-4">
        <div class="card border-0 shadow-soft kpi-card kpi-click" id="cardKpiNilai" role="button" tabindex="0">
          <div class="card-body">
            <div class="text-muted small">Jumlah Data Nilai (filtered)</div>
            <div class="big">${filtered.length}</div>
            <div class="text-muted small">Jumlah Peserta</div>
            <div class="fw-semibold">${pesertaArr.length}</div>
            <div class="small text-muted mt-2"><i class="bi bi-table"></i> Klik untuk lihat tabel</div>
          </div>
        </div>
      </div>

      <div class="col-12 col-md-4">
        <div class="card border-0 shadow-soft kpi-card kpi-click" id="cardKpiQueue" role="button" tabindex="0">
          <div class="card-body">
            <div class="text-muted small">Antrian Belum Terkirim</div>
            <div class="big" id="kQueue">0</div>
            <div class="text-muted small">Klik untuk lihat detail antrian.</div>
            <div class="small text-muted mt-2"><i class="bi bi-cloud-arrow-up"></i> Sync saat online</div>
          </div>
        </div>
      </div>

      <div class="col-12 col-md-4">
        <div class="card border-0 shadow-soft kpi-card kpi-click" id="cardKpiFailed" role="button" tabindex="0">
          <div class="card-body">
            <div class="text-muted small">Gagal Sync</div>
            <div class="big" id="kFailed">0</div>
            <div class="text-muted small">Klik untuk lihat log gagal sync.</div>
            <div class="small text-muted mt-2"><i class="bi bi-exclamation-triangle"></i> Periksa error & coba ulang input</div>
          </div>
        </div>
      </div>

      <div class="col-12">
        <div class="card border-0 shadow-soft">
          <div class="card-body">
            <div class="fw-semibold mb-2">3 Peserta Terbaik</div>
            ${top3.length ? top3.map((p,i)=>`
              <div class="d-flex justify-content-between border-bottom py-2">
                <div><span class="badge text-bg-primary me-2">#${i+1}</span>${p.nama} <span class="text-muted">(${p.nik})</span></div>
                <div class="fw-bold">${p.avg.toFixed(1)}</div>
              </div>`).join("") : `<div class="text-muted">Belum ada data.</div>`}
          </div>
        </div>
      </div>

      <div class="col-12 col-md-6">
        <div class="card border-0 shadow-soft">
          <div class="card-body">
            <div class="fw-semibold mb-2">Materi Rata-rata Tertinggi</div>
${
  materiHighList.length
  ? materiHighList.map((m,i)=>`
    <div class="d-flex justify-content-between ${i < materiHighList.length-1 ? "border-bottom" : ""} py-2">
      <div class="small">${m.kode||""} ${m.nama||""}</div>
      <div class="fw-bold">${m.avg.toFixed(1)}</div>
    </div>
  `).join("")
  : `<div class="text-muted">—</div>`
}
<hr>
<div class="fw-semibold mb-2">Materi Rata-rata Terendah</div>
${
  materiLowList.length
  ? materiLowList.map((m,i)=>`
    <div class="d-flex justify-content-between ${i < materiLowList.length-1 ? "border-bottom" : ""} py-2">
      <div class="small">${m.kode||""} ${m.nama||""}</div>
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
            ${low3.length ? low3.map((p,i)=>`
              <div class="d-flex justify-content-between border-bottom py-2">
                <div><span class="badge text-bg-warning me-2">#${i+1}</span>${p.nama} <span class="text-muted">(${p.nik})</span></div>
                <div class="fw-bold">${p.avg.toFixed(1)}</div>
              </div>`).join("") : `<div class="text-muted">—</div>`}
            <hr>
            <div class="fw-semibold mb-2">Peserta Gagal / Tidak Lulus</div>
            ${gagal.length ? gagal.slice(0,5).map(p=>`<div class="small">${p.nama} (${p.nik}) • ${p.avg.toFixed(1)}</div>`).join("") : `<div class="text-muted small">Tidak ada (nilai avg<70).</div>`}
          </div>
        </div>
      </div>
    </div>
  `);

  $("#fTahun").value = String(state.filters.tahun);
  $("#fJenis").value = state.filters.jenis;
  if($("#fNik")) $("#fNik").value = state.filters.nik;
    if($("#fTest")) $("#fTest").value = state.filters.test;
  if($("#fMateri")) $("#fMateri").value = state.filters.materi;


  $("#fTahun").addEventListener("change", (e)=>{ state.filters.tahun = e.target.value; renderDashboard(); });
  $("#fJenis").addEventListener("change", (e)=>{ state.filters.jenis = e.target.value; renderDashboard(); });
    if($("#fTest")) $("#fTest").addEventListener("change",(e)=>{ state.filters.test = e.target.value; renderDashboard(); });
  if($("#fMateri")) $("#fMateri").addEventListener("change",(e)=>{ state.filters.materi = e.target.value; renderDashboard(); });
  if($("#fNik")) $("#fNik").addEventListener("input",(e)=>{ state.filters.nik = e.target.value.trim(); });
  if($("#fNik")) $("#fNik").addEventListener("change",()=>renderDashboard());

  const q = await dbAll("queue");
  $("#kQueue").textContent = String(q.length);

    // ==== KPI counts (queue + failed) ====
  const failedRows = await dbAll("failed");
  const q2 = await dbAll("queue");
  $("#kQueue").textContent = String(q2.length);
  $("#kFailed").textContent = String(failedRows.length);

  // ==== CLICK HANDLERS (open modal tables) ====
  const bindClick = (el, fn)=>{
    if(!el) return;
    el.addEventListener("click", fn);
    el.addEventListener("keydown",(e)=>{
      if(e.key === "Enter" || e.key === " "){
        e.preventDefault();
        fn();
      }
    });
  };

  // 1) filtered nilai -> modal
  bindClick(document.getElementById("cardKpiNilai"), ()=>{
    const cols = [
      { key:"tanggal", label:"Tanggal" },
      { key:"tahun", label:"Tahun" },
      { key:"jenis_pelatihan", label:"Jenis Pelatihan" },
      { key:"nik", label:"NIK" },
      { key:"nama", label:"Nama" },
      { key:"test_type", label:"Test" },
      { key:"materi_kode", label:"Materi Kode" },
      { key:"materi_nama", label:"Materi Nama" },
      { key:"nilai", label:"Nilai" }
    ];

    const rowsForModal = filtered.map(r=>({
      tanggal: fmtTanggalDisplay(r.tanggal),
      tahun: r.tahun,
      jenis_pelatihan: r.jenis_pelatihan,
      nik: r.nik,
      nama: r.nama,
      test_type: r.test_type,
      materi_kode: r.materi_kode,
      materi_nama: r.materi_nama,
      nilai: r.nilai
    }));

    const meta = `Filtered: Tahun=${state.filters.tahun || "Semua"}, Jenis=${state.filters.jenis || "Semua"}, Test=${state.filters.test || "Semua"}, Materi=${state.filters.materi || "Semua"}`
      + (state.user.role==="admin" ? `, NIK=${state.filters.nik || "Semua"}` : "");

    const fn = `nilai_filtered_${state.filters.tahun || "all"}_${Date.now()}.xlsx`;

    showDataTableModal({
      title: "Data Nilai (Filtered)",
      meta,
      filename: fn,
      columns: cols,
      rows: rowsForModal
    });
  });

  // 2) queue -> modal
  bindClick(document.getElementById("cardKpiQueue"), async ()=>{
    const qRows = await dbAll("queue");
    const cols = [
      { key:"qid", label:"QID" },
      { key:"type", label:"Type" },
      { key:"created_at", label:"Created At" },
      { key:"payload_json", label:"Payload" }
    ];

    const rowsForModal = qRows
      .sort((a,b)=> String(b.created_at||"").localeCompare(String(a.created_at||"")))
      .map(x=>({
        qid: x.qid,
        type: x.type,
        created_at: x.created_at,
        payload_json: JSON.stringify(x.payload || {})
      }));

    showDataTableModal({
      title: "Antrian Belum Terkirim (Queue)",
      meta: `Total queue: ${qRows.length}`,
      filename: `queue_${Date.now()}.xlsx`,
      columns: cols,
      rows: rowsForModal
    });
  });

  // 3) failed -> modal
  bindClick(document.getElementById("cardKpiFailed"), async ()=>{
    const fRows = await dbAll("failed");
    const cols = [
      { key:"fid", label:"FID" },
      { key:"qid", label:"QID Asal" },
      { key:"type", label:"Type" },
      { key:"created_at", label:"Created At" },
      { key:"failed_at", label:"Failed At" },
      { key:"error", label:"Error" },
      { key:"payload_json", label:"Payload" }
    ];

    const rowsForModal = fRows
      .sort((a,b)=> String(b.failed_at||"").localeCompare(String(a.failed_at||"")))
      .map(x=>({
        fid: x.fid,
        qid: x.qid,
        type: x.type,
        created_at: x.created_at,
        failed_at: x.failed_at,
        error: x.error,
        payload_json: x.payload_json
      }));

    showDataTableModal({
      title: "Gagal Sync (Log)",
      meta: `Total gagal: ${fRows.length}`,
      filename: `gagal_sync_${Date.now()}.xlsx`,
      columns: cols,
      rows: rowsForModal
    });
  });
}

// ---------- Nilai list ----------
async function renderNilaiList(){
  const allNilai = await dbAll("nilai");
  const f = state.filters;

  // ✅ base filter untuk transkrip (abaikan filter test)
  const buildBaseFilteredForTranscript = ()=> filterBaseForTranscript(allNilai, state.filters);


  let rows = allNilai.filter(r=>{
    const rNik = normNik(r.nik);
    const uNik = normNik(state.user?.username);

    if(state.user.role !== "admin" && rNik !== uNik) return false;

    if(f.tahun && String(normStr(r.tahun)) !== String(normStr(f.tahun))) return false;
    if(f.jenis && normStr(r.jenis_pelatihan) !== normStr(f.jenis)) return false;

    if(f.test && normStr(r.test_type) !== normStr(f.test)) return false;

    if(f.materi){
      const mk = materiKeyFromRow(r);
      if(mk !== normStr(f.materi)) return false;
    }

    if(state.user.role === "admin" && f.nik && rNik !== normNik(f.nik)) return false;

    return true;
    }).sort((a,b)=> ( (b.tanggal_ts||0) - (a.tanggal_ts||0) ) || ((b.tanggal||"").localeCompare(a.tanggal||"")));


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
        <div class="col-6 ${state.user.role!=="admin" ? "d-none":""}">
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

  function ensureTrxPreviewModal(){
    if(document.getElementById("modalTrxPreview")) return;

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
            <div class="mt-2 d-none">
              <button class="btn btn-outline-secondary btn-sm" id="btnDbgToggle">
                <i class="bi bi-bug"></i> Debug (CCTV)
              </button>
              <div id="dbgBox" class="d-none mt-2 border rounded p-2 bg-light small" style="max-height:220px; overflow:auto;"></div>
            </div>

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

  $("#fTahun").value = String(state.filters.tahun);
  $("#fJenis").value = state.filters.jenis;
    if($("#fTest")) $("#fTest").value = state.filters.test;
  if($("#fMateri")) $("#fMateri").value = state.filters.materi;
  if($("#fNik")) $("#fNik").value = state.filters.nik;
  $("#fLimit").value = String(limit);
  $("#trxDate").valueAsDate = new Date();

  const renderTable = ()=>{
    const tbody = $("#tblNilai");
    tbody.innerHTML = "";
    const show = rows.slice(0, limit);
    for(const r of show){
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td class="small">${fmtTanggalDisplay(r.tanggal)}</td>
        <td>${r.nik}</td>
        <td>${r.nama||""}</td>
        <td class="small">${r.jenis_pelatihan||""}</td>
        <td>${r.test_type||""}</td>
        <td class="small">${(r.materi_kode||"")} ${(r.materi_nama||"")}</td>
        <td class="text-end fw-semibold">${(r.nilai??"")}</td>`;
      tbody.appendChild(tr);
    }
    $("#rowCount").textContent = String(rows.length);
  };

  renderTable();

  $("#fTahun").addEventListener("change",(e)=>{ state.filters.tahun=e.target.value; renderNilaiList(); });
  $("#fJenis").addEventListener("change",(e)=>{ state.filters.jenis=e.target.value; renderNilaiList(); });
    if($("#fTest")) $("#fTest").addEventListener("change",(e)=>{ state.filters.test = e.target.value; renderNilaiList(); });
  if($("#fMateri")) $("#fMateri").addEventListener("change",(e)=>{ state.filters.materi = e.target.value; renderNilaiList(); });
  if($("#fNik")) $("#fNik").addEventListener("change",(e)=>{ state.filters.nik=e.target.value.trim(); renderNilaiList(); });
  $("#fLimit").addEventListener("change",(e)=>{ limit = parseInt(e.target.value,10); renderTable(); });

  $("#btnExportXlsx").addEventListener("click", ()=>{
    const data = rows.map(r=>({
      tahun:r.tahun, jenis_pelatihan:r.jenis_pelatihan, nik:r.nik, nama:r.nama, test_type:r.test_type,
      materi_kode:r.materi_kode, materi_nama:r.materi_nama, nilai:r.nilai, tanggal:r.tanggal
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "nilai");
    XLSX.writeFile(wb, `nilai_${state.filters.tahun||"all"}.xlsx`);
  });

$("#btnPreviewTrx").addEventListener("click", ()=>{
  ensureTrxPreviewModal();

  const m = new bootstrap.Modal(document.getElementById("modalTrxPreview"));
  const trxDate = $("#trxDate").value; // ambil dari input halaman
  $("#trxDate2").value = trxDate || new Date().toISOString().slice(0,10);

  // sumber: Final rows dari rows (sudah terfilter oleh tahun/jenis/test/materi/nik)
  const baseFiltered = buildBaseFilteredForTranscript();
  const finalRowsAll = baseFiltered.filter(r => normStr(r.test_type) === "Final");

  // list NIK yang punya Final
  const nikSet = [...new Set(finalRowsAll.map(r => normNik(r.nik)).filter(Boolean))];

  // kalau bukan admin: kunci ke user login
  let nikOptions = nikSet;
  if(state.user.role !== "admin"){
    nikOptions = [normNik(state.user.username)];
  }

  const sel = $("#trxNikPick");
  sel.innerHTML = nikOptions.map(n => {
    const p = (state.masters.peserta||[]).find(x=>normNik(x.nik)===normNik(n)) || {};
    const nm = p.nama || (finalRowsAll.find(x=>normNik(x.nik)===n)?.nama) || "";
    return `<option value="${n}">${n} - ${nm}</option>`;
  }).join("");
  let __lastTrxData = null;

  // default pilih pertama
  if(!sel.value && nikOptions[0]) sel.value = nikOptions[0];

  const renderPreview = ()=>{
    const nik = sel.value;
    const data = buildTranscriptData({ nik, finalRowsAll });
    __lastTrxData = data;

    const meta = $("#trxMeta");
    const body = $("#trxTblBody");
    body.innerHTML = "";

    if(!data){
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
          <div><b>Tanggal Presentasi</b>: ${fmtTanggalIndoLong(p.tanggal_presentasi||"")}</div>
          <div><b>Judul Makalah</b>: ${normStr(p.judul_presentasi||"")}</div>
        </div>
        <div class="col-12 col-md-6">
          <div><b>Lokasi OJT</b>: ${normStr(p.lokasi_ojt||"")}</div>
          <div><b>Unit</b>: ${normStr(p.unit||"")}</div>
          <div><b>Region</b>: ${normStr(p.region||"")}</div>
          <div><b>Jenis Pelatihan</b>: ${data.jenisLabel}</div>
        </div>
      </div>
    `;

    data.lines.forEach((ln, i)=>{
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${i+1}</td>
        <td>${ln.materi_kode}</td>
        <td>${ln.materi_nama}</td>
        <td class="text-end">${ln.rerata_kelas==null ? "" : r1(ln.rerata_kelas)}</td>
        <td class="text-end fw-semibold">${r1(ln.poin)}</td>
        <td class="text-end fw-semibold">${r1(ln.nilai)}</td>
      `;
      body.appendChild(tr);
    });

    $("#trxTotal").textContent = r1(data.totalWeighted);
    const total = data.totalWeighted;
    const status = getLulusStatus(total);
    const pred = pickPredikatFromMaster(total);

    const badges = $("#trxBadges");
    badges.innerHTML = `
      <div class="d-flex flex-wrap gap-2 align-items-center">
        <span class="badge ${status.badge}">DINYATAKAN: ${status.label}</span>
        <span class="badge text-bg-primary">PREDIKAT: ${pred}</span>
        <span class="badge text-bg-dark">TOTAL: ${r1(total)}</span>
      </div>
    `;

    $("#trxFootNote").textContent = `Tanggal Transkrip: ${fmtTanggalIndoLong($("#trxDate2").value)}`;
    if (typeof renderDbgBox === "function") renderDbgBox();

  };

  $("#btnTrxRender").onclick = renderPreview;
  $("#trxNikPick").onchange = renderPreview;

  // tombol PDF dari preview: tinggal trigger export PDF biasa (pakai filter NIK)
  $("#btnPdfFromPreview").onclick = async ()=>{
  const keepDate = $("#trxDate2").value; // tanggal yg dipilih di preview
  state.filters.nik = sel.value;

  await renderNilaiList();

  // set tanggal transkrip di halaman list (agar PDF pakai tanggal yang sama)
  const trxDateEl = document.getElementById("trxDate");
  if(trxDateEl && keepDate) trxDateEl.value = keepDate;

  document.getElementById("btnExportPdf").click();
};

  $("#btnXlsxFromPreview").onclick = ()=>{
  if(!__lastTrxData) return toast("Data transkrip belum ada.");

  const d = __lastTrxData;
  const total = d.totalWeighted;
  const status = getLulusStatus(total);
  const pred = pickPredikatFromMaster(total);

  // Sheet 1: Ringkasan
  const ringkas = [{
    nik: d.nik,
    nama: d.peserta?.nama || "",
    jenis_pelatihan: d.jenisLabel,
    tanggal_transkrip: $("#trxDate2").value,
    lokasi_ojt: d.peserta?.lokasi_ojt || "",
    unit: d.peserta?.unit || "",
    region: d.peserta?.region || "",
    tanggal_presentasi: d.peserta?.tanggal_presentasi || "",
    judul_makalah: d.peserta?.judul_presentasi || "",
    total_nilai: Math.round(total*10)/10,
    dinyatakan: status.label,
    predikat: pred
  }];

  // Sheet 2: Detail
  const detail = (d.lines||[]).map((x, i)=>({
    no: i+1,
    materi_kode: x.materi_kode,
    materi_nama: x.materi_nama,
    rerata_kelas: x.rerata_kelas==null ? "" : Math.round(x.rerata_kelas*10)/10,
    poin: Math.round((x.poin||0)*10)/10,
    nilai: Math.round((x.nilai||0)*10)/10
  }));

  const ws1 = XLSX.utils.json_to_sheet(ringkas);
  const ws2 = XLSX.utils.json_to_sheet(detail);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, "Ringkasan");
  XLSX.utils.book_append_sheet(wb, ws2, "Detail");

  XLSX.writeFile(wb, `transkrip_${d.nik}_${String(d.tahun||state.filters.tahun)}.xlsx`);
};

  renderPreview();
  m.show();
});


  $("#btnExportPdf").addEventListener("click", async ()=>{
  const trxDate = $("#trxDate").value;
  if(!rows.length) return toast("Tidak ada data untuk dibuat transkrip.");

  // 1) Ambil data Final saja (source untuk transkrip) + abaikan filter test
  const baseFiltered = filterBaseForTranscript(allNilai, state.filters);
  const finalRowsAll = baseFiltered.filter(r => normStr(r.test_type) === "Final");
  if(!finalRowsAll.length) return toast("Tidak ada data Final untuk dibuat transkrip.");

  // 2) Group by peserta
  const grouped = new Map();
  for(const r of finalRowsAll){
    const k = normNik(r.nik);
    if(!k) continue;
    if(!grouped.has(k)) grouped.set(k, []);
    grouped.get(k).push(r);
  }
  if(!grouped.size) return toast("Tidak ada data Final untuk dibuat transkrip.");

  // 3) Rerata map sekali untuk semua peserta
  const avgMap = buildAvgMap(finalRowsAll);

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit:"pt", format:"a4" });

  let first = true;
  for(const [nik, list] of grouped){
    if(!first) doc.addPage();
    first = false;

    const peserta = (state.masters.peserta||[]).find(p => normNik(p.nik) === normNik(nik)) || {};

    const tahun = list[0]?.tahun ?? peserta.tahun ?? state.filters.tahun ?? "";
    const jenisRaw = list[0]?.jenis_pelatihan || peserta.jenis_pelatihan || "-";
    const jenisLabel = jenisDenganTahun(jenisRaw, tahun);

    // Header
    doc.setFont("helvetica","bold");
    doc.setFontSize(14);
    doc.text("KARYAMAS PLANTATION", 40, 40);
    doc.setFontSize(12);
    doc.text("SERIANG TRAINING CENTER", 40, 58);
    doc.setFontSize(16);
    doc.text("TRANSKRIP NILAI", 40, 86);

    doc.setFontSize(11);
    doc.setFont("helvetica","normal");
    const left = 40;
    const right = 320;

    doc.text(`Nama : ${peserta.nama || list[0]?.nama || ""}`, left, 120);
    doc.text(`NIK  : ${nik}`, left, 136);

    doc.text(`Tanggal Presentasi : ${fmtTanggalIndoLong(peserta.tanggal_presentasi || "")}`, left, 152);
    doc.text(`Judul Makalah      : ${normStr(peserta.judul_presentasi || "")}`, left, 168);

    doc.text(`Lokasi OJT : ${normStr(peserta.lokasi_ojt || "")}`, right, 120);
    doc.text(`Unit       : ${normStr(peserta.unit || "")}`, right, 136);
    doc.text(`Region     : ${normStr(peserta.region || "")}`, right, 152);

    doc.setFont("helvetica","bold");
    doc.text(`${jenisLabel}`, left, 192);

    // Tabel
    const bobotObj = getBobotByJenis(jenisRaw);
    const listSorted = list.slice().sort(cmpMateriKodeTranskrip);

        const rowsTbl = [];
    let totalNilaiWeighted = 0;

    // ✅ HITUNG JUMLAH MATERI PER KATEGORI (per peserta)
    const catCount = {};
    for(const rr of listSorted){
      const kk = getKategoriMateri(rr.materi_kode, rr.materi_nama);
      const kNorm = normStr(kk);
      if(!kNorm) continue;
      catCount[kNorm] = (catCount[kNorm] || 0) + 1;
    }

    for(let idx=0; idx<listSorted.length; idx++){
      const r = listSorted[idx];
      const mKey = materiKeyFromRow(r);

      const rk = getRerataKelasFromMap(avgMap, jenisRaw, tahun, mKey);

      const poin = (typeof r.nilai === "number") ? r.nilai : parseFloat(r.nilai);
      const poinOk = Number.isFinite(poin) ? poin : 0;

      const kategori = getKategoriMateri(r.materi_kode, r.materi_nama);
      const bobotPct = getBobotPercentForKategori(bobotObj, kategori);

      const nKat = catCount[normStr(kategori)] || 0;

      // ✅ rumus baru: poin * bobotPct / nKat / 100
      const nilaiWeighted = (nKat > 0 && bobotPct > 0)
        ? (poinOk * bobotPct / nKat / 100)
        : 0;

      totalNilaiWeighted += (Number.isFinite(nilaiWeighted) ? nilaiWeighted : 0);

      rowsTbl.push([
        String(idx+1),
        r.materi_kode || "",
        r.materi_nama || "",
        rk==null ? "" : r1(rk),
        r1(poinOk),
        r1(nilaiWeighted)
      ]);
    }

    doc.autoTable({
      startY: 206,
      head: [[ "No", "Kode", "Jenis Materi", "Rerata Kelas", "Poin", "Nilai" ]],
      body: rowsTbl,
      styles: { fontSize: 9, cellPadding: 3 },
      headStyles: { fillColor: [11,58,103] },
      margin: { left: 40, right: 40 }
    });

    const y = doc.lastAutoTable.finalY + 18;

    const total = totalNilaiWeighted;
    const pred = pickPredikatFromMaster(total);

    doc.setFont("helvetica","bold");
    doc.text(
      `Total Nilai: ${r1(total)}     Dinyatakan: ${total>=70 ? "LULUS":"TIDAK LULUS"}     Predikat: ${pred}`,
      40,
      y
    );

    doc.setFont("helvetica","normal");
    doc.setFontSize(9);
    doc.text(`Tanggal Transkrip: ${fmtTanggalIndoLong(trxDate)}`, 40, y+18);
    doc.text("Dokumen ini dicetak secara komputerisasi", 40, 800);
  }

  doc.save(`transkrip_${state.filters.tahun||"all"}.pdf`);
});

}

// ---------- Input Nilai (Admin) ----------
async function renderInputNilai(){
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
                <datalist id="dlJenis">${optPelatihan().replaceAll("<option","<option").replace('<option value="">Semua</option>','')}</datalist>
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
                <button class="btn btn-primary" type="submit"><i class="bi bi-save"></i> Simpan (Lokal)</button>
                <button class="btn btn-outline-secondary" type="button" id="btnClear">Clear</button>
              </div>
            </form>
          </div>
        </div>
      </div>
    </div>
  `);

  $("#inTanggal").valueAsDate = new Date();

  // build datalists
  const dlNik = $("#dlNik");
  dlNik.innerHTML = state.masters.peserta.map(p=>`<option value="${p.nik}">${p.nama}</option>`).join("");
  $("#dlMateriKode").innerHTML = state.masters.materi.map(m=>`<option value="${m.kode}">${m.nama}</option>`).join("");
  $("#dlMateriNama").innerHTML = state.masters.materi.map(m=>`<option value="${m.nama}">${m.kode}</option>`).join("");

  $("#inNik").addEventListener("input", ()=>{
    const nik = $("#inNik").value.trim();
    const p = state.masters.peserta.find(x=>String(x.nik)===String(nik));
    $("#inNama").value = p?.nama || "";
    if(p?.jenis_pelatihan && !$("#inJenis").value) $("#inJenis").value = p.jenis_pelatihan;
  });

  $("#inKode").addEventListener("input", ()=>{
    const kode = $("#inKode").value.trim();
    const m = state.masters.materi.find(x=>String(x.kode)===String(kode));
    if(m) $("#inMateri").value = m.nama;
  });
  $("#inMateri").addEventListener("input", ()=>{
    const nama = $("#inMateri").value.trim();
    const m = state.masters.materi.find(x=>String(x.nama)===String(nama));
    if(m) $("#inKode").value = m.kode;
  });

  $("#formNilai").addEventListener("submit", async (e)=>{
    e.preventDefault();
      let row = {
      id: "N"+Date.now()+"_"+Math.random().toString(16).slice(2),
      tahun: parseInt($("#inTahun").value,10),
      jenis_pelatihan: $("#inJenis").value.trim(),
      test_type: $("#inTest").value,
      nik: $("#inNik").value.trim(),
      nama: $("#inNama").value.trim(),
      tanggal: $("#inTanggal").value, // dari input date: YYYY-MM-DD
      materi_kode: $("#inKode").value.trim(),
      materi_nama: $("#inMateri").value.trim(),
      nilai: parseFloat($("#inNilai").value),
    };

    row = normalizeNilaiRow(row);

    await dbPut("nilai", row);
    await enqueue("upsert_nilai", row);
    toast("Nilai disimpan ke lokal & masuk antrian sync.");
    $("#inNilai").value="";
  });

  $("#btnClear").addEventListener("click", ()=>{
    $("#formNilai").reset();
    $("#inTanggal").valueAsDate = new Date();
  });

  $("#btnImport").addEventListener("click", async ()=>{
  const btn = $("#btnImport");
    runBusy(btn, async ()=>{
      const f = $("#fileXlsx").files[0];
      if(!f) return toast("Pilih file xlsx dulu.");

      // baca file
      progressStart(100, "Membaca file Excel…");
      const buf = await f.arrayBuffer();

      const wb = XLSX.read(buf, { type:"array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows2 = XLSX.utils.sheet_to_json(ws, { header: 1, defval:"" });
      const headerRaw2 = (rows2[0]||[]).map(h=>String(h||""));

      const norm2 = (s)=>{
        s = String(s||"").trim().toLowerCase().replace(/\s+/g,"_").replace(/[^a-z0-9_]/g,"");
        if(s==="nik") return "nik";
        if(s==="nama") return "nama";
        if(s==="tahun") return "tahun";
        if(s==="jenis_pelatihan") return "jenis_pelatihan";
        if(s==="test_type") return "test_type";
        if(s==="materi_kode") return "materi_kode";
        if(s==="materi_nama") return "materi_nama";
        if(s==="nilai") return "nilai";
        if(s==="tanggal") return "tanggal";
        return s;
      };
      const header2 = headerRaw2.map(norm2);

      const data = [];
      for(let i=1;i<rows2.length;i++){
        const arr = rows2[i];
        if(!arr || arr.every(v=>String(v).trim()==="")) continue;
        const obj={};
        header2.forEach((k,idx)=>{ if(k) obj[k]=arr[idx]; });
        data.push(obj);
      }

      const total = data.length;
      progressStart(total, `Import ${total} baris…`);

      let n=0;
      for(const r of data){
          let row = {
          id: "N"+Date.now()+"_"+Math.random().toString(16).slice(2),

          // boleh kosong; nanti dinormalisasi dari tanggal kalau perlu
          tahun: r.tahun,

          jenis_pelatihan: r.jenis_pelatihan,
          test_type: r.test_type,
          nik: r.nik,
          nama: r.nama,

          // penting: jangan paksa toString dulu, biarkan normalizer deteksi Date/number/string
          tanggal: r.tanggal,

          materi_kode: r.materi_kode,
          materi_nama: r.materi_nama,
          nilai: r.nilai,
        };

        row = normalizeNilaiRow(row);

        await dbPut("nilai", row);
        await enqueue("upsert_nilai", row);
        n++;
        if(n % 10 === 0 || n === total){
          progressSet(n, total, `Mengimpor… (${n}/${total})`);
        }
      }

      await rebuildPelatihanCache();

      progressDone("Import selesai.");
      toast(`Import selesai: ${n} baris (masuk antrian sync).`);
    }, { busyText:"Import…" }).catch(e=>toast(e.message));
  });

}

// ---------- Master ----------
async function renderMaster(){
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
            <!-- ✅ BARU: MASTER BOBOT -->
      <div class="col-12">
        <div class="card border-0 shadow-soft">
          <div class="card-body">
            <div class="fw-semibold mb-2">Upload Master Bobot Penilaian (.xlsx)</div>
            <p class="small text-muted mb-2">
              Kolom: jenis_pelatihan, managerial, teknis, support, ojt, presentasi
              <br><span class="text-muted">(angka persen, contoh: 20, 60, 10, 5, 5)</span>
            </p>
            <input type="file" id="upBobot" class="form-control" accept=".xlsx">
            <button class="btn btn-outline-primary btn-sm mt-2" id="btnUpBobot"><i class="bi bi-upload"></i> Upload ke GAS</button>
          </div>
        </div>
      </div>
      <!-- ✅ END BARU -->
    </div>
  `);

  $("#mTahun").value = String(state.filters.tahun);
  $("#mJenis").value = state.filters.jenis;

  $("#btnPullMaster").addEventListener("click", async ()=>{
  const btn = $("#btnPullMaster");
    runBusy(btn, async ()=>{
      const tahun = $("#mTahun").value;
      const jenis = $("#mJenis").value;

      progressStart(100, "Menarik master dari Google Sheet…");
      const res = await gasCall("pull_master", { tahun, jenis });

      await applyMastersFromServer(res.masters);

      progressDone("Master siap.");
      toast("Master ditarik & disimpan ke offline.");
    }, { busyText:"Menarik…" }).catch(e=>toast(e.message));
  });

  async function uploadXlsx(file){
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type:"array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval:"" });
    if(!rows.length) return [];

    const headerRaw = rows[0].map(h=>String(h||""));
    const norm = (s)=>{
      s = String(s||"").trim().toLowerCase();
      s = s.replace(/\s+/g,"_");
      s = s.replace(/[^a-z0-9_]/g,"");

      // ✅ tambah mapping bobot
      if(s === "jenis") return "jenis_pelatihan";
      if(s === "jenis_pelatihan") return "jenis_pelatihan";
      if(s === "managerial") return "managerial";
      if(s === "teknis") return "teknis";
      if(s === "support") return "support";
      if(s === "ojt") return "ojt";
      if(s === "presentasi") return "presentasi";

      if(s.startsWith("kategori")) return "kategori";
      if(s==="nik") return "nik";
      if(s==="nama") return "nama";
      if(s==="jenis_pelatihan") return "jenis_pelatihan";
      if(s==="lokasi_ojt") return "lokasi_ojt";
      if(s==="tanggal_presentasi") return "tanggal_presentasi";
      if(s==="judul_presentasi") return "judul_presentasi";
      return s;
    };
    const header = headerRaw.map(norm);

    const out = [];
    for(let i=1;i<rows.length;i++){
      const arr = rows[i];
      if(!arr || arr.every(v=>String(v).trim()==="")) continue;
      const obj = {};
      header.forEach((k,idx)=>{ if(k) obj[k] = arr[idx]; });
      out.push(obj);
    }
    return out;
  }

  $("#btnUpPeserta").addEventListener("click", async ()=>{
  const btn = $("#btnUpPeserta");
    runBusy(btn, async ()=>{
      const f = $("#upPeserta").files[0];
      if(!f) return toast("Pilih file dulu.");

      const rows = await uploadXlsx(f);
      const total = rows.length;

      progressStart(total, `Menyiapkan ${total} baris master peserta…`);
      let n=0;

      for(const r of rows){
        await enqueue("upsert_master_peserta", r);
        n++;
        if(n % 20 === 0 || n === total){
          progressSet(n, total, `Masuk antrian… (${n}/${total})`);
        }
      }

      // update state + save
      state.masters.peserta = (state.masters.peserta || []).concat(rows.map(r=>({
        nik: normStr(r.nik),
        nama: normStr(r.nama),
        jenis_pelatihan: normStr(r.jenis_pelatihan),
        lokasi_ojt: normStr(r.lokasi_ojt),
        unit: normStr(r.unit),
        region: normStr(r.region),
        group: normStr(r.group),

        // PENTING: simpan RAW agar parseAnyDate bisa baca number serial Excel / Date / string ISO
        tanggal_presentasi: (r.tanggal_presentasi ?? ""),

        judul_presentasi: normStr(r.judul_presentasi),

        // tahun boleh string/number, biarkan apa adanya dulu
        tahun: (r.tahun ?? "")
      })));


      await saveMastersToDB();
      await rebuildPelatihanCache();
      await syncUsersFromMasterPeserta();

      progressDone("Upload master peserta selesai.");
      toast(`Master peserta masuk antrian sync: ${total} baris.`);
      if(navigator.onLine) syncQueue().catch(()=>{});
    }, { busyText:"Memproses…" }).catch(e=>toast(e.message));
  });


  $("#btnUpMateri").addEventListener("click", async ()=>{
  const btn = $("#btnUpMateri");
    runBusy(btn, async ()=>{
      const f = $("#upMateri").files[0];
      if(!f) return toast("Pilih file dulu.");

      const rows = await uploadXlsx(f);
      const total = rows.length;

      progressStart(total, `Menyiapkan ${total} baris master materi…`);
      let n=0;

      for(const r of rows){
        await enqueue("upsert_master_materi", r);
        n++;
        if(n % 20 === 0 || n === total){
          progressSet(n, total, `Masuk antrian… (${n}/${total})`);
        }
      }

      state.masters.materi = state.masters.materi.concat(rows.map(r=>({
        kode: (r.kode||"").toString(),
        nama: (r.nama||"").toString(),
        kategori: (r.kategori||"").toString(),
      })));

      await saveMastersToDB();

      progressDone("Upload master materi selesai.");
      toast(`Master materi masuk antrian sync: ${total} baris.`);
      if(navigator.onLine) syncQueue().catch(()=>{});
    }, { busyText:"Memproses…" }).catch(e=>toast(e.message));
  });

    // ✅ BARU: Upload Master Bobot
  const btnUpBobot = $("#btnUpBobot");
  if(btnUpBobot){
    btnUpBobot.addEventListener("click", async ()=>{
      runBusy(btnUpBobot, async ()=>{
        const f = $("#upBobot").files[0];
        if(!f) return toast("Pilih file dulu.");

        const rows = await uploadXlsx(f);
        const cleaned = (rows||[])
          .map(r=>({
            jenis_pelatihan: normStr(r.jenis_pelatihan || r.jenis || ""),
            managerial: numVal(r.managerial),
            teknis: numVal(r.teknis),
            support: numVal(r.support),
            ojt: numVal(r.ojt),
            presentasi: numVal(r.presentasi),
          }))
          .filter(r=>r.jenis_pelatihan);

        const total = cleaned.length;
        if(!total) return toast("Tidak ada baris valid. Pastikan kolom jenis_pelatihan terisi.");

        progressStart(total, `Menyiapkan ${total} baris master bobot…`);
        let n=0;

        for(const r of cleaned){
          // masuk antrian sync
          await enqueue("upsert_master_bobot", r);
          n++;
          if(n % 20 === 0 || n === total){
            progressSet(n, total, `Masuk antrian… (${n}/${total})`);
          }
        }

        // merge ke state.masters.bobot (hindari duplikat per jenis_pelatihan)
        const map = new Map();
        for(const x of (state.masters.bobot || [])){
          const k = normKey(normalizeJenisForBobot(x.jenis_pelatihan || ""));
          if(k) map.set(k, x);
        }
        for(const x of cleaned){
          const k = normKey(normalizeJenisForBobot(x.jenis_pelatihan || ""));
          if(!k) continue;
          map.set(k, x);
        }
        state.masters.bobot = [...map.values()];

        await saveMastersToDB();

        progressDone("Upload master bobot selesai.");
        toast(`Master bobot masuk antrian sync: ${total} baris.`);
        if(navigator.onLine) syncQueue().catch(()=>{});
      }, { busyText:"Memproses…" }).catch(e=>toast(e.message));
    });
  }


}

// ---------- Setting ----------
async function renderSettingAdmin(){
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
              <div class="d-flex flex-wrap gap-2">
                <button class="btn btn-outline-danger btn-sm" id="btnWipe">
                  <i class="bi bi-trash"></i> Hapus IndexedDB
                </button>

                <!-- ✅ BARU: clear log failed sync -->
                <button class="btn btn-outline-secondary btn-sm" id="btnClearFailed" title="Hapus log gagal sync">
                  <i class="bi bi-x-circle"></i> Clear Failed
                </button>
              </div>

              <div class="small text-muted mt-2">
                *Clear Failed hanya menghapus log store <code>failed</code> (tidak menghapus data nilai/master).
              </div>
          </div>
        </div>
      </div>
    </div>
  `);

  $("#aTahun").value = String(state.filters.tahun);
  $("#aJenis").value = state.filters.jenis;

  const _el_btnPullNilai = $("#btnPullNilai");
  if(_el_btnPullNilai) _el_btnPullNilai.addEventListener("click", async ()=>{
    runBusy(_el_btnPullNilai, async ()=>{
      const tahun = $("#aTahun").value;
      const jenis = $("#aJenis").value;

      // ✅ AUTO PULL MASTER bila master masih kosong
      const mastersEmpty =
        !(state.masters.peserta?.length) &&
        !(state.masters.materi?.length) &&
        !(state.masters.bobot?.length) &&
        !(state.masters.predikat?.length);

      if(mastersEmpty){
        progressStart(1, "Master kosong. Menarik master dulu…");
        // pilih salah satu: public master / master spesifik tahun+jenis
        // 1) kalau GAS Anda punya pull_public_master:
        const mres = await gasCall("pull_public_master", {});
        await applyMastersFromServer(mres.masters);
        progressDone("Master siap.");
      }

      // ✅ tarik nilai
      const res = await gasCall("pull_nilai", { tahun, jenis });
      const total = res.rows?.length || 0;
      progressStart(total, "Menarik nilai dari Google Sheet…");

      let n=0;
      for(const r of (res.rows||[])){
        const row = normalizeNilaiRow(r);
        await dbPut("nilai", row);
        n++;
        if(n % 10 === 0 || n === total){
          progressSet(n, total, `Menyimpan ke offline… (${n}/${total})`);
        }
      }

      await rebuildPelatihanCache();
      progressDone("Tarik nilai selesai.");
      toast(`Tarik nilai selesai: ${n} baris.`);
    }, { busyText:"Menarik…" }).catch(e=>toast(e.message));
  });



  const _el_btnReset = $("#btnReset");
  if(_el_btnReset) _el_btnReset.addEventListener("click", async ()=>{
    runBusy(_el_btnReset, async ()=>{
      const username = $("#resetUser").value.trim();
      if(!username) return toast("Isi username/NIK.");
      const newHash = await sha256Hex(username === "admin" ? "123456" : username);

      // local update
      const u = await dbGet("users", username);
      if(u){
        u.pass_hash = newHash;
        u.must_change = (username !== "admin");
        await dbPut("users", u);
      }

      await gasCall("admin_reset_password", { username, pass_hash: newHash, must_change: (username !== "admin") });
      toast("Reset password berhasil.");
    }, { busyText:"Reset…" }).catch(e=>toast(e.message));
  });



    const btnAdmChangePass = $("#btnAdmChangePass");
  if(btnAdmChangePass){
    btnAdmChangePass.addEventListener("click", async ()=>{
      runBusy(btnAdmChangePass, async ()=>{
        const np = $("#admNewPass").value.trim();
        if(np.length < 6) return toast("Minimal 6 karakter.");
        await changePassword(np);
        $("#admNewPass").value = "";
        toast("Password berhasil diubah.");
      }, { busyText:"Menyimpan…" }).catch(e=>toast(e.message));
    });

  }
  const _el_btnWipe = $("#btnWipe");
  if(_el_btnWipe) _el_btnWipe.addEventListener("click", async ()=>{
    if(!confirm("Yakin hapus semua data lokal?")) return;
    const db = await dbPromise;
    await db.clear("masters");
    await db.clear("nilai");
    await db.clear("queue");
    toast("Data lokal dihapus. Silakan tarik master & nilai dari Google Sheet.");
  });
    // ✅ BARU: Clear store "failed"
  const _el_btnClearFailed = $("#btnClearFailed");
  if(_el_btnClearFailed) _el_btnClearFailed.addEventListener("click", async ()=>{
    runBusy(_el_btnClearFailed, async ()=>{
      if(!confirm("Yakin hapus semua log Gagal Sync (failed)?")) return;

      await dbClear("failed"); // ✅ sesuai permintaan: dbClear("failed")
      toast("Log Gagal Sync (failed) sudah dikosongkan.");

      // optional: kalau sedang ada elemen KPI failed di layar (mis. dashboard), update cepat bila ada
      const kf = document.getElementById("kFailed");
      if(kf) kf.textContent = "0";
    }, { busyText:"Menghapus…" }).catch(e=>toast(e.message));
  });
}

async function renderSettingUser(){
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

  $("#btnPullMy").addEventListener("click", async ()=>{
  const btn = $("#btnPullMy");
    runBusy(btn, async ()=>{
      const res = await gasCall("pull_user_bundle", { username: state.user.username });

      // masters
      await applyMastersFromServer(res.masters, { rebuildCache:false, syncUsers:true });

      // nilai (progress)
      const total = res.nilai?.length || 0;
      progressStart(total, "Menarik nilai Anda…");

      let n=0;
      for(const r of (res.nilai||[])){
        const row = normalizeNilaiRow(r);
        await dbPut("nilai", row);
        n++;
        if(n % 10 === 0 || n === total){
          progressSet(n, total, `Menyimpan ke offline… (${n}/${total})`);
        }
      }

      progressDone("Tarik data selesai.");
      toast("Data berhasil ditarik.");

      state.filters.nik = "";
      const all = await dbAll("nilai");
      const mine = all.filter(x => normNik(x.nik) === normNik(state.user.username));
      const latestYear = mine.map(x=>parseInt(x.tahun,10)).filter(n=>!Number.isNaN(n)).sort((a,b)=>b-a)[0];
      if(latestYear) state.filters.tahun = latestYear;

      await rebuildPelatihanCache();
      renderView("nilai");
    }, { busyText:"Menarik…" }).catch(e=>toast(e.message));
  });

  $("#btnChangePass").addEventListener("click", async ()=>{
  const btn = $("#btnChangePass");
    runBusy(btn, async ()=>{
      const np = $("#newPass").value.trim();
      if(np.length < 6) return toast("Minimal 6 karakter.");
      await changePassword(np);
      $("#newPass").value = "";
    }, { busyText:"Menyimpan…" }).catch(e=>toast(e.message));
  });

}

// ---------- First login password change ----------
async function promptChangePasswordIfNeeded(){
  if(!state.user?.must_change) return;
  const modalHtml = `
  <div class="modal fade" id="modalPwd" tabindex="-1">
    <div class="modal-dialog">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title">Ganti Password</h5>
        </div>
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
  const m = new bootstrap.Modal($("#modalPwd"), { backdrop:"static", keyboard:false });
  m.show();
  $("#btnPwdSave").addEventListener("click", async ()=>{
  const btn = $("#btnPwdSave");
    runBusy(btn, async ()=>{
      const a = $("#pwdNew1").value.trim();
      const b = $("#pwdNew2").value.trim();
      if(a.length < 6) return toast("Minimal 6 karakter.");
      if(a!==b) return toast("Password tidak sama.");
      await changePassword(a);
      m.hide();
      $("#modalPwd").remove();
    }, { busyText:"Menyimpan…" }).catch(e=>toast(e.message));
  });

}

// ---------- Public pull for login screen ----------
async function pullPublicMasters(){
  try{
    const res = await gasCall("pull_public_master", {});
    await applyMastersFromServer(res.masters);

    toast("Master publik ditarik. Sekarang user bisa login offline.");
  }catch(e){
    toast(e.message);
  }
}

// ---------- Boot ----------
async function boot(){
    // ✅ Chrome Android sering gagal untuk ESM+SW+IndexedDB jika dibuka dari file://
  if(location.protocol === "file:"){
    // tampilkan toast bila UI sudah siap (bootstrap toast ada)
    try{ toast("Aplikasi dibuka dari file://. Di Chrome HP harus lewat https atau http://localhost (tidak bisa dari File Manager)."); }catch(e){}
    console.warn("Running from file:// is not supported on Chrome Android for module+SW.");
    // tetap lanjut, tapi user sudah dapat instruksi jelas
  }
  try{ localStorage.removeItem("GAS_URL"); }catch(e){}
  // register SW
  if("serviceWorker" in navigator){
    try{ await navigator.serviceWorker.register("./sw.js"); }catch(e){ console.warn(e); }
  }

  setNetBadge();

  await ensureDefaultUsers();
  await loadMastersFromDB();
  dbg("masters.fromDB", {
    peserta: (state.masters.peserta||[]).length,
    materi: (state.masters.materi||[]).length,
    pelatihan: (state.masters.pelatihan||[]).length,
    bobot: (state.masters.bobot||[]).length,
    predikat: (state.masters.predikat||[]).length,
  });

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
