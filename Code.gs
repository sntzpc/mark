/**
 * Karyamas Transkrip Nilai - Google Apps Script backend
 * Deploy as WebApp (Execute as: Me, Access: Anyone with the link)
 *
 * REQUIRED SHEETS:
 * - users         : username | role | name | pass_hash | must_change | updated_at
 * - master_peserta: nik | nama | jenis_pelatihan | lokasi_ojt | unit | region | group | tanggal_presentasi | judul_presentasi | tahun
 * - master_materi : kode | nama | kategori
 * - master_pelatihan: nama | tahun
 * - master_bobot  : jenis_pelatihan | managerial | teknis | support | ojt | presentasi
 * - master_predikat: nama | min | max
 * - nilai         : id | tahun | jenis_pelatihan | nik | nama | test_type | materi_kode | materi_nama | nilai | tanggal | rerata_kelas | poin | updated_at
 */

const SPREADSHEET_ID = "PASTE_SPREADSHEET_ID_HERE";

function doGet(e){
  try{
    const action = (e.parameter && e.parameter.action) || "";
    const payloadStr = (e.parameter && e.parameter.payload) ? decodeURIComponent(e.parameter.payload) : "{}";
    const callback = (e.parameter && e.parameter.callback) || "";
    const payload = JSON.parse(payloadStr || "{}");
    const out = route_(action, payload);
    const resp = { ok:true, ...out };
    if(callback){
      return ContentService
        .createTextOutput(callback + "(" + JSON.stringify(resp) + ");")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(JSON.stringify(resp))
      .setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    const callback = (e.parameter && e.parameter.callback) || "";
    const resp = { ok:false, error:String(err) };
    if(callback){
      return ContentService
        .createTextOutput(callback + "(" + JSON.stringify(resp) + ");")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(JSON.stringify(resp))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents || "{}");
    const action = body.action;
    const payload = body.payload || {};
    const out = route_(action, payload);
    return ContentService.createTextOutput(JSON.stringify({ ok:true, ...out })).setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ ok:false, error:String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

function ss_(){ return SpreadsheetApp.openById(SPREADSHEET_ID); }
function sh_(name){ 
  const s = ss_();
  let sh = s.getSheetByName(name);
  if(!sh){ sh = s.insertSheet(name); }
  return sh;
}

function ensureHeaders_(sheetName, headers){
  const sh = sh_(sheetName);
  const needCols = headers.length;

  const currentLastCol = sh.getLastColumn();
  if(currentLastCol < needCols){
    if(currentLastCol === 0){
      sh.insertColumnsAfter(1, needCols-1);
    }else{
      sh.insertColumnsAfter(currentLastCol, needCols-currentLastCol);
    }
  }

  const current = sh.getRange(1,1,1,needCols).getValues()[0].map(v=>String(v||"").trim());
  const expected = headers.map(v=>String(v||"").trim());

  const allEmpty = current.every(v=>v==="");
  const match = current.every((v,i)=>v===expected[i]);

  if(allEmpty || !match){
    sh.getRange(1,1,1,needCols).setValues([expected]);
  }
  return sh;
}

function route_(action, p){
  switch(action){
    case "login": return login_(p);
    case "change_password": return changePassword_(p);
    case "admin_reset_password": return adminReset_(p);
    case "pull_public_master": return pullPublicMaster_();
    case "pull_master": return pullMaster_(p);
    case "upload_master_peserta": return uploadMasterPeserta_(p);
    case "upload_master_materi": return uploadMasterMateri_(p);
    case "upsert_master_peserta": return upsertMasterPeserta_(p);
    case "upsert_master_materi": return upsertMasterMateri_(p);
    case "pull_nilai": return pullNilai_(p);
    case "upsert_nilai": return upsertNilai_(p);
    case "pull_user_bundle": return pullUserBundle_(p);
    default: throw new Error("Unknown action: "+action);
  }
}

function login_(p){
  const username = String(p.username||"").trim();
  const pass_hash = String(p.pass_hash||"").trim();
  if(!username || !pass_hash) throw new Error("missing credentials");

  ensureHeaders_("users", ["username","role","name","pass_hash","must_change","updated_at"]);
  const sh = sh_("users");
  const values = sh.getDataRange().getValues();
  const idxU = values[0].indexOf("username");
  const idxH = values[0].indexOf("pass_hash");

  // if not exists, auto-create user from master_peserta with default hash (sha256 of NIK done on client)
  let rowIndex = -1;
  for(let i=1;i<values.length;i++){
    if(String(values[i][idxU]).trim()===username){ rowIndex=i; break; }
  }
  if(rowIndex===-1){
    if(username==="admin"){
      // create admin default
      sh.appendRow([username,"admin","Administrator", pass_hash, false, new Date()]);
      rowIndex = sh.getLastRow()-1;
    }else{
      // try master_peserta
      ensureHeaders_("master_peserta", ["nik","nama","jenis_pelatihan","lokasi_ojt","unit","region","group","tanggal_presentasi","judul_presentasi","tahun"]);
      const mp = sh_("master_peserta").getDataRange().getValues();
      const h = mp[0];
      const nikIdx = h.indexOf("nik");
      const namaIdx = h.indexOf("nama");
      let foundName = "";
      for(let i=1;i<mp.length;i++){
        if(String(mp[i][nikIdx]).trim()===username){ foundName = String(mp[i][namaIdx]||""); break; }
      }
      sh.appendRow([username,"user", foundName || ("Peserta "+username), pass_hash, true, new Date()]);
      rowIndex = sh.getLastRow()-1;
    }
    // refresh
  }

  const row = sh.getRange(rowIndex+1, 1, 1, values[0].length).getValues()[0];
  const user = {
    username: String(row[idxU]).trim(),
    role: String(row[values[0].indexOf("role")]||"user"),
    name: String(row[values[0].indexOf("name")]||""),
    pass_hash: String(row[idxH]||""),
    must_change: Boolean(row[values[0].indexOf("must_change")])
  };
  if(user.pass_hash !== pass_hash) throw new Error("invalid password");
  return { user };
}

function changePassword_(p){
  const username = String(p.username||"").trim();
  const pass_hash = String(p.pass_hash||"").trim();
  ensureHeaders_("users", ["username","role","name","pass_hash","must_change","updated_at"]);
  const sh = sh_("users");
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const idxU = h.indexOf("username");
  const idxH = h.indexOf("pass_hash");
  const idxM = h.indexOf("must_change");
  const idxT = h.indexOf("updated_at");
  for(let i=1;i<values.length;i++){
    if(String(values[i][idxU]).trim()===username){
      sh.getRange(i+1, idxH+1).setValue(pass_hash);
      sh.getRange(i+1, idxM+1).setValue(false);
      sh.getRange(i+1, idxT+1).setValue(new Date());
      return { updated:true };
    }
  }
  throw new Error("user not found");
}

function adminReset_(p){
  const username = String(p.username||"").trim();
  const pass_hash = String(p.pass_hash||"").trim();
  const must_change = Boolean(p.must_change);
  ensureHeaders_("users", ["username","role","name","pass_hash","must_change","updated_at"]);
  const sh = sh_("users");
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const idxU = h.indexOf("username");
  const idxH = h.indexOf("pass_hash");
  const idxM = h.indexOf("must_change");
  const idxT = h.indexOf("updated_at");
  for(let i=1;i<values.length;i++){
    if(String(values[i][idxU]).trim()===username){
      sh.getRange(i+1, idxH+1).setValue(pass_hash);
      sh.getRange(i+1, idxM+1).setValue(must_change);
      sh.getRange(i+1, idxT+1).setValue(new Date());
      return { reset:true };
    }
  }
  sh.appendRow([username, username==="admin"?"admin":"user", "", pass_hash, must_change, new Date()]);
  return { reset:true, created:true };
}

function pullPublicMaster_(){
  // minimal masters for login screen: peserta + pelatihan + materi + bobot + predikat
  return { masters: readMasters_({ tahun:"", jenis:"" }) };
}

function pullMaster_(p){
  return { masters: readMasters_(p) };
}

function pullUserBundle_(p){
  const username = String(p.username||"").trim();
  const masters = readMasters_({ tahun:"", jenis:"" });
  // nilai filtered by nik
  ensureHeaders_("nilai", ["id","tahun","jenis_pelatihan","nik","nama","test_type","materi_kode","materi_nama","nilai","tanggal","rerata_kelas","poin","updated_at"]);
  const sh = sh_("nilai");
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const idxNik = h.indexOf("nik");
  const rows = [];
  for(let i=1;i<values.length;i++){
    if(String(values[i][idxNik]).trim()===username){
      rows.push(rowToObj_(h, values[i]));
    }
  }
  return { masters, nilai: rows };
}

function readMasters_(p){
  const tahun = String(p.tahun||"").trim();
  const jenis = String(p.jenis||"").trim();

  ensureHeaders_("master_peserta", ["nik","nama","jenis_pelatihan","lokasi_ojt","unit","region","group","tanggal_presentasi","judul_presentasi","tahun"]);
  ensureHeaders_("master_materi", ["kode","nama","kategori"]);
  ensureHeaders_("master_pelatihan", ["nama","tahun"]);
  ensureHeaders_("master_bobot", ["jenis_pelatihan","managerial","teknis","support","ojt","presentasi"]);
  ensureHeaders_("master_predikat", ["nama","min","max"]);

  const peserta = filterSheet_("master_peserta", (row)=> {
    if(tahun && String(row.tahun) !== tahun) return false;
    if(jenis && String(row.jenis_pelatihan) !== jenis) return false;
    return true;
  });

  return {
    peserta,
    materi: readAll_("master_materi"),
    pelatihan: readAll_("master_pelatihan"),
    bobot: readAll_("master_bobot"),
    predikat: readAll_("master_predikat")
  };
}

function readAll_(sheetName){
  const sh = sh_(sheetName);
  const values = sh.getDataRange().getValues();
  if(values.length<2) return [];
  const h = values[0];
  const out = [];
  for(let i=1;i<values.length;i++){
    if(values[i].every(v=>String(v).trim()==="")) continue;
    out.push(rowToObj_(h, values[i]));
  }
  return out;
}

function filterSheet_(sheetName, fn){
  const all = readAll_(sheetName);
  return all.filter(fn);
}

function uploadMasterPeserta_(p){
  const rows = p.rows || [];
  ensureHeaders_("master_peserta", ["nik","nama","jenis_pelatihan","lokasi_ojt","unit","region","group","tanggal_presentasi","judul_presentasi","tahun"]);
  const sh = sh_("master_peserta");
  const h = sh.getRange(1,1,1,10).getValues()[0];
  // append
  for(const r of rows){
    sh.appendRow([
      r.nik||"", r.nama||"", r.jenis_pelatihan||"", r.lokasi_ojt||"", r.unit||"", r.region||"", r.group||"",
      r.tanggal_presentasi||"", r.judul_presentasi||"", r.tahun||""
    ]);
  }
  return { inserted: rows.length };
}

function uploadMasterMateri_(p){
  const rows = p.rows || [];
  ensureHeaders_("master_materi", ["kode","nama","kategori"]);
  const sh = sh_("master_materi");
  for(const r of rows){
    sh.appendRow([r.kode||"", r.nama||"", r.kategori||""]);
  }
  return { inserted: rows.length };
}

function pullNilai_(p){
  const tahun = String(p.tahun||"").trim();
  const jenis = String(p.jenis||"").trim();
  ensureHeaders_("nilai", ["id","tahun","jenis_pelatihan","nik","nama","test_type","materi_kode","materi_nama","nilai","tanggal","rerata_kelas","poin","updated_at"]);
  const sh = sh_("nilai");
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const out = [];
  for(let i=1;i<values.length;i++){
    const o = rowToObj_(h, values[i]);
    if(tahun && String(o.tahun)!==tahun) continue;
    if(jenis && String(o.jenis_pelatihan)!==jenis) continue;
    out.push(o);
  }
  return { rows: out };
}

function upsertNilai_(p){
  ensureHeaders_("nilai", ["id","tahun","jenis_pelatihan","nik","nama","test_type","materi_kode","materi_nama","nilai","tanggal","rerata_kelas","poin","updated_at"]);
  const sh = sh_("nilai");
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const idIdx = h.indexOf("id");

  // find row by id
  for(let i=1;i<values.length;i++){
    if(String(values[i][idIdx]).trim()===String(p.id).trim()){
      writeNilaiRow_(sh, h, i+1, p);
      return { updated:true };
    }
  }
  // append
  const row = objToRow_(h, p);
  sh.appendRow(row);
  return { inserted:true };
}

function writeNilaiRow_(sh, h, rowIndex, p){
  const row = objToRow_(h, p);
  sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
}

function objToRow_(headers, p){
  const now = new Date();
  const obj = Object.assign({}, p);
  obj.updated_at = now;
  return headers.map(k => (obj[k]!==undefined ? obj[k] : ""));
}

function rowToObj_(headers, row){
  const o = {};
  headers.forEach((k,i)=> o[k]=row[i]);
  return o;
}


function upsertMasterPeserta_(r){
  ensureHeaders_("master_peserta", ["nik","nama","jenis_pelatihan","lokasi_ojt","unit","region","group","tanggal_presentasi","judul_presentasi","tahun"]);
  const sh = sh_("master_peserta");
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const nikIdx = h.indexOf("nik");
  const nik = String(r.nik||"").trim();
  if(!nik) throw new Error("nik kosong");
  for(let i=1;i<values.length;i++){
    if(String(values[i][nikIdx]).trim()===nik){
      sh.getRange(i+1, 1, 1, h.length).setValues([[
        r.nik||"", r.nama||"", r.jenis_pelatihan||"", r.lokasi_ojt||"", r.unit||"", r.region||"", r.group||"",
        r.tanggal_presentasi||"", r.judul_presentasi||"", r.tahun||""
      ]]);
      return { updated:true };
    }
  }
  sh.appendRow([
    r.nik||"", r.nama||"", r.jenis_pelatihan||"", r.lokasi_ojt||"", r.unit||"", r.region||"", r.group||"",
    r.tanggal_presentasi||"", r.judul_presentasi||"", r.tahun||""
  ]);
  return { inserted:true };
}

function upsertMasterMateri_(r){
  ensureHeaders_("master_materi", ["kode","nama","kategori"]);
  const sh = sh_("master_materi");
  const values = sh.getDataRange().getValues();
  const h = values[0];
  const kodeIdx = h.indexOf("kode");
  const kode = String(r.kode||"").trim();
  if(!kode) throw new Error("kode kosong");
  for(let i=1;i<values.length;i++){
    if(String(values[i][kodeIdx]).trim()===kode){
      sh.getRange(i+1, 1, 1, h.length).setValues([[ r.kode||"", r.nama||"", r.kategori||"" ]]);
      return { updated:true };
    }
  }
  sh.appendRow([ r.kode||"", r.nama||"", r.kategori||"" ]);
  return { inserted:true };
}
