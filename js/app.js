// Karyamas Transkrip Nilai (Offline-first)
// Frontend: Bootstrap 5 + IndexedDB (idb) + XLSX + jsPDF
// Backend: Google Apps Script (see Code.gs)

const GAS_URL = "https://script.google.com/macros/s/AKfycbz25x3uRYzzACTp5pwPf4zCpx0Atf2ihqbN7G7IiTbaipkUyRf1-34bgAKvR6-CodfF/exec"; // HARD-CODE: ganti dengan URL WebApp GAS Anda

import { openDB } from "https://cdn.jsdelivr.net/npm/idb@8/+esm";

const $ = (sel, root=document)=>root.querySelector(sel);
const $$ = (sel, root=document)=>[...root.querySelectorAll(sel)];

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

function setNetBadge(){
  const online = navigator.onLine;
  const badge = $("#netBadge");
  if(!badge) return;
  badge.textContent = online ? "Online" : "Offline";
  badge.classList.toggle("badge-online", online);
  badge.classList.toggle("badge-offline", !online);
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
const dbPromise = openDB("karyamas_transkrip_db", 2, {
  upgrade(db){
    // Buat object store secara aman (untuk migrasi dari versi lama)
    if(!db.objectStoreNames.contains("masters")){
      db.createObjectStore("masters", { keyPath:"key" });
    }
    if(!db.objectStoreNames.contains("nilai")){
      const nilai = db.createObjectStore("nilai", { keyPath:"id" });
      try { nilai.createIndex("by_tahun", "tahun"); } catch(e) {}
      try { nilai.createIndex("by_nik", "nik"); } catch(e) {}
      try { nilai.createIndex("by_jenis", "jenis_pelatihan"); } catch(e) {}
    }else{
      // pastikan index ada
      const tx = db.transaction("nilai", "versionchange");
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
    if(!db.objectStoreNames.contains("settings")){
      db.createObjectStore("settings", { keyPath:"key" });
    }
  }
});

async function dbGet(store, key){ return (await dbPromise).get(store, key); }
async function dbPut(store, val){ return (await dbPromise).put(store, val); }
async function dbDel(store, key){ return (await dbPromise).delete(store, key); }
async function dbAll(store){ return (await dbPromise).getAll(store); }

async function loadMastersFromDB(){
  const rows = await dbAll("masters");
  const map = Object.fromEntries(rows.map(r=>[r.key, r.data]));
  state.masters.peserta = map.peserta || [];
  state.masters.materi = map.materi || [];
  state.masters.pelatihan = map.pelatihan || [];
  state.masters.bobot = map.bobot || [];
  state.masters.predikat = map.predikat || [];
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
  const enc = new TextEncoder().encode(str);
  const buf = await crypto.subtle.digest("SHA-256", enc);
  return [...new Uint8Array(buf)].map(b=>b.toString(16).padStart(2,"0")).join("");
}

// ---------- GAS helpers ----------
async function gasCall(action, payload={}){
  if(!GAS_URL || GAS_URL.includes("PASTE_YOUR_GAS_WEBAPP_URL")){
    throw new Error("GAS_URL belum diisi. Isi hardcode di js/app.js.");
  }
  // Gunakan JSONP (doGet) untuk menghindari masalah CORS pada WebApp GAS
  return await gasJsonp(action, payload);
}

function gasJsonp(action, payload){
  return new Promise((resolve, reject)=>{
    const cb = "cb_"+Date.now()+"_"+Math.random().toString(16).slice(2);
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
      clearTimeout(timeout);
      try{ delete window[cb]; }catch(e){ window[cb]=undefined; }
      script.remove();
    }

    const params = new URLSearchParams();
    params.set("action", action);
    params.set("payload", encodeURIComponent(JSON.stringify(payload||{})));
    params.set("callback", cb);

    const url = GAS_URL + (GAS_URL.includes("?") ? "&" : "?") + params.toString();
    const script = document.createElement("script");
    script.src = url;
    script.onerror = ()=>{
      cleanup();
      reject(new Error("Gagal memuat GAS (JSONP). Pastikan URL WebApp benar & akses publik."));
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
  for(const it of items){
    try{
      await gasCall(it.type, it.payload);
      await db.delete("queue", it.qid);
      done++;
      if(done % 5 === 0 || done === total){
        progressSet(done, total, `Sync… (${done}/${total})`);
      }
    }catch(e){
      console.warn("queue sync failed", it, e);
      progressDone("Sync berhenti (ada error).");
      break; // stop on first error
    }
  }

  await refreshQueueBadge();
  $("#syncState").textContent = "Online";
  progressDone("Sync selesai.");
  toast("Sync selesai.");
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


function normStr(v){
  return String(v ?? "").trim();
}
function normNik(v){
  // NIK sering kebawa spasi, angka, atau format lain
  return normStr(v).replace(/\s+/g,"");
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

  // DD/MM/YYYY or DD/MM/YYYY HH:mm:ss
  const m1 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if(m1){
    const dd = Number(m1[1]);
    const mm = Number(m1[2]);
    const yy = Number(m1[3]);
    const hh = Number(m1[4] || 0);
    const mi = Number(m1[5] || 0);
    const ss = Number(m1[6] || 0);
    const d = new Date(yy, mm-1, dd, hh, mi, ss);
    if(!Number.isNaN(d.getTime())) return d;
  }

  // Last resort: Date.parse
  const d3 = new Date(s);
  if(!Number.isNaN(d3.getTime())) return d3;

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
}

// ---------- Nilai list ----------
async function renderNilaiList(){
  const allNilai = await dbAll("nilai");
  const f = state.filters;

  // ✅ base filter untuk transkrip (abaikan filter test)
//    supaya preview/PDF tetap bisa ambil Final walau user sedang filter PreTest/PostTest
function buildBaseFilteredForTranscript(){
  return allNilai.filter(r=>{
    const rNik = normNik(r.nik);
    const uNik = normNik(state.user?.username);

    if(state.user.role !== "admin" && rNik !== uNik) return false;

    if(f.tahun && String(normStr(r.tahun)) !== String(normStr(f.tahun))) return false;
    if(f.jenis && normStr(r.jenis_pelatihan) !== normStr(f.jenis)) return false;

    if(f.materi){
      const mk = materiKeyFromRow(r);
      if(mk !== normStr(f.materi)) return false;
    }

    if(state.user.role === "admin" && f.nik && rNik !== normNik(f.nik)) return false;

    return true;
  });
}


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

  function r1(n){
    const x = (typeof n==="number") ? n : parseFloat(n);
    if(!Number.isFinite(x)) return "";
    return (Math.round(x*10)/10).toFixed(1);
  }

  function getKategoriMateri(materiKode, materiNama){
    const kode = normStr(materiKode);
    const nama = normStr(materiNama);
    let m = null;
    if(kode){
      m = (state.masters.materi||[]).find(x => normStr(x.kode) === kode) || null;
    }
    if(!m && nama){
      m = (state.masters.materi||[]).find(x => normStr(x.nama) === nama) || null;
    }
    let kat = normStr(m?.kategori);
    if(!kat) return "";
    const k = kat.toLowerCase();
    if(k.includes("man")) return "Managerial";
    if(k.includes("tek")) return "Teknis";
    if(k.includes("sup")) return "Support";
    if(k.includes("ojt")) return "OJT";
    if(k.includes("pres")) return "Presentasi";
    return kat;
  }

  function getBobotByJenis(jenis){
    const j = normStr(jenis);
    const b = (state.masters.bobot||[]).find(x => normStr(x.jenis_pelatihan) === j) || null;
    return {
      managerial: parseFloat(b?.managerial) || 0,
      teknis: parseFloat(b?.teknis) || 0,
      support: parseFloat(b?.support) || 0,
      ojt: parseFloat(b?.ojt) || 0,
      presentasi: parseFloat(b?.presentasi) || 0
    };
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

  function jenisDenganTahun(jenis, tahun){
    const j = normStr(jenis);
    const y = String(normStr(tahun));
    if(!j) return y ? `Tahun ${y}` : "-";
    if(!y) return j;
    if(/tahun\s+\d{4}/i.test(j)) return j;
    return `${j} Tahun ${y}`;
  }

  // hitung rerata kelas per materi dari Final rows yg sedang terfilter
  function buildAvgMap(finalRows){
    const avgMap = new Map(); // key -> {total,n}
    for(const r of finalRows){
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

  // bangun data transkrip untuk 1 peserta
  function buildTranscriptData({ nik, finalRowsAll }){
    const mine = finalRowsAll.filter(r => normNik(r.nik) === normNik(nik));
    if(!mine.length) return null;

    const peserta = (state.masters.peserta||[]).find(p => String(p.nik) === String(nik)) || {};
    const tahun = mine[0]?.tahun ?? peserta.tahun ?? state.filters.tahun ?? "";
    const jenisRaw = mine[0]?.jenis_pelatihan || peserta.jenis_pelatihan || "-";
    const jenisLabel = jenisDenganTahun(jenisRaw, tahun);

    const avgMap = buildAvgMap(finalRowsAll);
    const bobotObj = getBobotByJenis(jenisRaw);

    const listSorted = mine.slice().sort((a,b)=>String(a.materi_kode||"").localeCompare(String(b.materi_kode||"")));

    const lines = [];
    let totalWeighted = 0;

    for(const r of listSorted){
      const mKey = materiKeyFromRow(r);
      const rk = getRerataKelasFromMap(avgMap, jenisRaw, tahun, mKey);

      const poin = (typeof r.nilai==="number") ? r.nilai : parseFloat(r.nilai);
      const poinOk = Number.isFinite(poin) ? poin : 0;

      const kategori = getKategoriMateri(r.materi_kode, r.materi_nama);
      const bobotPct = getBobotPercentForKategori(bobotObj, kategori);

      const nilaiWeighted = poinOk * (bobotPct / 100);
      if(Number.isFinite(nilaiWeighted)) totalWeighted += nilaiWeighted;

      lines.push({
        materi_kode: r.materi_kode || "",
        materi_nama: r.materi_nama || "",
        rerata_kelas: rk,
        poin: poinOk,
        nilai: nilaiWeighted
      });
    }

    return {
      nik,
      peserta,
      tahun,
      jenisRaw,
      jenisLabel,
      lines,
      totalWeighted
    };
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

  function fmtTanggalIndoLong(v){
  // output: "17 Desember 2025"
  const d = parseAnyDate(v); // pakai helper yang sudah kita buat sebelumnya
  if(!d) return "";
  const bulan = [
    "Januari","Februari","Maret","April","Mei","Juni",
    "Juli","Agustus","September","Oktober","November","Desember"
  ];
  return `${String(d.getDate()).padStart(2,"0")} ${bulan[d.getMonth()]} ${d.getFullYear()}`;
}

function pickPredikatFromMaster(v){
  const n = (typeof v==="number") ? v : parseFloat(v);
  if(!Number.isFinite(n)) return "-";

  const rules = (state.masters.predikat||[])
    .map(x=>({
      nama: normStr(x.nama),
      min: parseFloat(x.min),
      max: parseFloat(x.max)
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
  if(!Number.isFinite(n)) return { lulus:false, label:"TIDAK", badge:"text-bg-danger" };
  const ok = n >= 70;
  return ok
    ? { lulus:true, label:"LULUS", badge:"text-bg-success" }
    : { lulus:false, label:"TIDAK", badge:"text-bg-danger" };
}

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
    const p = (state.masters.peserta||[]).find(x=>String(x.nik)===String(n)) || {};
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

  // ==============================
  // 1) Ambil data Final saja (source untuk transkrip)
  // ==============================
  const finalRowsAll = rows.filter(r => normStr(r.test_type) === "Final");
  if(!finalRowsAll.length) return toast("Tidak ada data Final untuk dibuat transkrip.");

  // ==============================
  // 2) Helper: format 1 angka belakang koma
  // ==============================
  const r1 = (n)=>{
    const x = (typeof n==="number") ? n : parseFloat(n);
    if(!Number.isFinite(x)) return "";
    return (Math.round(x * 10) / 10).toFixed(1);
  };

  // ==============================
  // 3) Helper: Ambil kategori materi (Managerial/Teknis/Support/OJT/Presentasi)
  // ==============================
  function getKategoriMateri(materiKode, materiNama){
    const kode = normStr(materiKode);
    const nama = normStr(materiNama);

    // cari by kode (prioritas), lalu by nama
    let m = null;
    if(kode){
      m = (state.masters.materi||[]).find(x => normStr(x.kode) === kode) || null;
    }
    if(!m && nama){
      m = (state.masters.materi||[]).find(x => normStr(x.nama) === nama) || null;
    }

    let kat = normStr(m?.kategori);
    // normalisasi ejaan umum
    // (di master Anda disarankan: Managerial/Teknis/Support/OJT/Presentasi)
    if(!kat) return "";
    const k = kat.toLowerCase();
    if(k.includes("man")) return "Managerial";
    if(k.includes("tek")) return "Teknis";
    if(k.includes("sup")) return "Support";
    if(k.includes("ojt")) return "OJT";
    if(k.includes("pres")) return "Presentasi";
    return kat;
  }

  // ==============================
  // 4) Helper: bobot per jenis_pelatihan
  // ==============================
  function getBobotByJenis(jenis){
    const j = normStr(jenis);
    const b = (state.masters.bobot||[]).find(x => normStr(x.jenis_pelatihan) === j) || null;
    // return object % (0-100)
    return {
      managerial: parseFloat(b?.managerial) || 0,
      teknis: parseFloat(b?.teknis) || 0,
      support: parseFloat(b?.support) || 0,
      ojt: parseFloat(b?.ojt) || 0,
      presentasi: parseFloat(b?.presentasi) || 0
    };
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

  // ==============================
  // 5) RERATA KELAS: hitung per (jenis_pelatihan + tahun + materiKey)
  //    sumber: nilai Final semua peserta
  // ==============================
  const avgMap = new Map(); // key -> {total,n}
  for(const r of finalRowsAll){
    const jenis = normStr(r.jenis_pelatihan);
    const tahun = String(normStr(r.tahun));
    const mkey = materiKeyFromRow(r);
    if(!jenis || !tahun || !mkey) continue;

    const key = `${jenis}||${tahun}||${mkey}`;
    if(!avgMap.has(key)) avgMap.set(key, { total:0, n:0 });

    const v = (typeof r.nilai === "number") ? r.nilai : parseFloat(r.nilai);
    if(Number.isFinite(v)){
      const o = avgMap.get(key);
      o.total += v;
      o.n += 1;
    }
  }

  function getRerataKelas(jenis, tahun, materiKey){
    const key = `${normStr(jenis)}||${String(normStr(tahun))}||${normStr(materiKey)}`;
    const o = avgMap.get(key);
    if(!o || !o.n) return "";
    return (o.total / o.n);
  }

  // ==============================
  // 6) Predikat: pakai master_predikat jika ada, fallback ke rule lama
  // ==============================
  function pickPredikatFromMaster(v){
    const n = (typeof v==="number") ? v : parseFloat(v);
    if(!Number.isFinite(n)) return "-";

    const rules = (state.masters.predikat||[])
      .map(x=>({
        nama: normStr(x.nama),
        min: parseFloat(x.min),
        max: parseFloat(x.max)
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

  // ==============================
  // 7) Label jenis pelatihan + tahun: ".... Tahun 2025"
  // ==============================
  function jenisDenganTahun(jenis, tahun){
    const j = normStr(jenis);
    const y = String(normStr(tahun));
    if(!j) return y ? `Tahun ${y}` : "-";
    if(!y) return j;
    if(/tahun\s+\d{4}/i.test(j)) return j; // sudah ada "Tahun 2025"
    return `${j} Tahun ${y}`;
  }

  // ==============================
  // 8) Group by peserta (1 PDF berisi banyak transkrip)
  // ==============================
  const grouped = new Map();
  for(const r of finalRowsAll){
    if(!grouped.has(r.nik)) grouped.set(r.nik, []);
    grouped.get(r.nik).push(r);
  }
  if(!grouped.size) return toast("Tidak ada data Final untuk dibuat transkrip.");

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit:"pt", format:"a4" });

  let first=true;
  for(const [nik, list] of grouped){
    if(!first) doc.addPage();
    first=false;

    // peserta master
    const peserta = (state.masters.peserta||[]).find(p=>String(p.nik)===String(nik)) || {};

    // tentukan tahun & jenis batch (ambil dari list dulu, fallback ke master)
    const tahun = list[0]?.tahun ?? peserta.tahun ?? state.filters.tahun ?? "";
    const jenisRaw = list[0]?.jenis_pelatihan || peserta.jenis_pelatihan || "-";
    const jenisLabel = jenisDenganTahun(jenisRaw, tahun);

    // header
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

    // Data kiri
    doc.text(`Nama : ${peserta.nama || list[0]?.nama || ""}`, left, 120);
    doc.text(`NIK  : ${nik}`, left, 136);

    // Tanggal Presentasi & Judul Makalah
    doc.text(`Tanggal Presentasi : ${fmtTanggalIndoLong(peserta.tanggal_presentasi || "")}`, left, 152);
    doc.text(`Judul Makalah      : ${normStr(peserta.judul_presentasi || "")}`, left, 168);

    // Data kanan
    doc.text(`Lokasi OJT : ${normStr(peserta.lokasi_ojt || "")}`, right, 120);
    doc.text(`Unit       : ${normStr(peserta.unit || "")}`, right, 136);
    doc.text(`Region     : ${normStr(peserta.region || "")}`, right, 152);

    // Jenis pelatihan (bold)
    doc.setFont("helvetica","bold");
    doc.text(`${jenisLabel}`, left, 192);

    // ==============================
    // 9) Susun tabel: Rerata Kelas, Poin (nilai asli), Nilai (poin*bobot)
    // ==============================
    const bobotObj = getBobotByJenis(jenisRaw);

    // urut materi
    const listSorted = list
      .slice()
      .sort((a,b)=>String(a.materi_kode||"").localeCompare(String(b.materi_kode||"")));

    const rowsTbl = [];
    let totalNilaiWeighted = 0;

    for(let idx=0; idx<listSorted.length; idx++){
      const r = listSorted[idx];
      const mKey = materiKeyFromRow(r);

      // Rerata kelas per materi (Final, batch sama)
      const rk = getRerataKelas(jenisRaw, tahun, mKey);

      // Poin = nilai asli final
      const poin = (typeof r.nilai === "number") ? r.nilai : parseFloat(r.nilai);
      const poinOk = Number.isFinite(poin) ? poin : 0;

      // kategori & bobot
      const kategori = getKategoriMateri(r.materi_kode, r.materi_nama);
      const bobotPct = getBobotPercentForKategori(bobotObj, kategori);

      // Nilai = poin * bobot%
      const nilaiWeighted = poinOk * (bobotPct / 100);
      totalNilaiWeighted += (Number.isFinite(nilaiWeighted) ? nilaiWeighted : 0);

      rowsTbl.push([
        String(idx+1),
        r.materi_kode || "",
        r.materi_nama || "",
        r1(rk === "" ? NaN : rk),
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

    // ==============================
    // 10) Total Nilai = SUM nilaiWeighted
    // ==============================
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


  function average(arr){ return arr.length ? arr.reduce((a,b)=>a+b,0)/arr.length : 0; }
  function pickPredikat(v){
    // default rule from sample: Sangat Memuaskan >90; Memuaskan >80-90; Baik ≥76-80; Kurang ≥70-<76; Sangat Kurang <70
    if(v>90) return "Sangat Memuaskan";
    if(v>80) return "Memuaskan";
    if(v>=76) return "Baik";
    if(v>=70) return "Kurang";
    return "Sangat Kurang";
  }
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

      state.masters = res.masters;
      await saveMastersToDB();
      await rebuildPelatihanCache();
      await syncUsersFromMasterPeserta();

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
      state.masters.peserta = state.masters.peserta.concat(rows.map(r=>({
        nik: (r.nik||"").toString(),
        nama: (r.nama||"").toString(),
        jenis_pelatihan: (r.jenis_pelatihan||"").toString(),
        lokasi_ojt: (r.lokasi_ojt||"").toString(),
        unit: (r.unit||"").toString(),
        region: (r.region||"").toString(),
        group: (r.group||"").toString(),
        tanggal_presentasi: (r.tanggal_presentasi||"").toString(),
        judul_presentasi: (r.judul_presentasi||"").toString(),
        tahun: (r.tahun||"").toString(),
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
            <button class="btn btn-outline-danger btn-sm" id="btnWipe"><i class="bi bi-trash"></i> Hapus IndexedDB</button>
          </div>
        </div>
      </div>
    </div>
  `);

  $("#setGas").value = localStorage.getItem("GAS_URL") || "";
  $("#aTahun").value = String(state.filters.tahun);
  $("#aJenis").value = state.filters.jenis;

  const _el_btnSaveGas = $("#btnSaveGas");
  if(_el_btnSaveGas) _el_btnSaveGas.addEventListener("click", ()=>{
    localStorage.setItem("GAS_URL", $("#setGas").value.trim());
    toast("GAS_URL disimpan.");
  });

  const _el_btnPullNilai = $("#btnPullNilai");
  if(_el_btnPullNilai) _el_btnPullNilai.addEventListener("click", async ()=>{
    runBusy(_el_btnPullNilai, async ()=>{
      const res = await gasCall("pull_nilai", { tahun: $("#aTahun").value, jenis: $("#aJenis").value });
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
      state.masters = res.masters;
      await saveMastersToDB();
      await syncUsersFromMasterPeserta();

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
    state.masters = res.masters;
    await saveMastersToDB();
    await rebuildPelatihanCache();
    await syncUsersFromMasterPeserta();
    toast("Master publik ditarik. Sekarang user bisa login offline.");
  }catch(e){
    toast(e.message);
  }
}

// ---------- Boot ----------
async function boot(){
  // register SW
  if("serviceWorker" in navigator){
    try{ await navigator.serviceWorker.register("./sw.js"); }catch(e){ console.warn(e); }
  }

  setNetBadge();
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
