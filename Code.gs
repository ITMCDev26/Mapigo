/**
 * MAPIGO backend: Google Apps Script + Google Sheets (database) + Google Drive (poster images).
 *
 * SETUP (about 5 minutes)
 * 1. Create a new Google Sheet. Open Extensions > Apps Script and paste this whole file.
 * 2. Change ADMIN_EMAIL and ADMIN_PASSWORD below (use a strong password, 8+ characters).
 * 3. Pick the function "setup" in the toolbar and click Run. Approve the permissions.
 *    This creates the tabs (Users, Ads, Places, Runs, Settings, Sessions) and your admin account.
 * 4. Delete the password text from ADMIN_PASSWORD (the account already exists).
 * 5. Deploy > New deployment > type "Web app". Execute as: Me. Who has access: Anyone. Deploy.
 * 6. Copy the Web app URL (ends in /exec) into API_URL in mapigo.html AND admin.html.
 * After changing this script later, use Deploy > Manage deployments > Edit > New version.
 */
const ADMIN_EMAIL = "admin@example.com";
const ADMIN_PASSWORD = "change-this-password";
const DRIVE_FOLDER = "MAPIGO uploads";
const TZ = "Asia/Manila";

const HEAD = {
  Users:   ["id","email","name","city","photo","passHash","salt","role","showBoard","active","adj","score","streak","stage","state","created"],
  Ads:     ["id","advertiser","title","desc","cta","image","link","seconds","start","end","order","active","views","clicks"],
  Places:  ["id","name","cat","lat","lng","desc","image","active"],
  Runs:    ["id","userId","name","km","link","pts","status","created"],
  Settings:["key","value"],
  Sessions:["token","userId","exp"],
  Events:  ["id","title","date","time","mult","desc","image","active"],
  QR:      ["id","type","km","label","placeId","design","active","scans","created"],
  RunLog:  ["id","userId","name","km","startedAt","finishedAt","seconds","pts","mult","event","day","status","startQr","finishQr"],
  Joins:   ["id","eventId","userId","name","created"]
};
const MIN_PACE = 150;          // seconds per km: the fastest finish we accept (2:30 per km)
const MAX_SECS = 6 * 3600;     // a run must be finished within 6 hours of the start scan
const MAX_RUNS_DAY = 3;        // QR runs that can earn points per person per day
const PHOTO_FOLDER = "MAPIGO profile photos";
const SETTING_KEYS = ["runPoints","siteUrl","cities","houseTitle","houseText","houseButton","houseLink","houseImage","defaultSeconds"];

/* ---------- one-time setup ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEAD).forEach(n => {
    const s = ss.getSheetByName(n) || ss.insertSheet(n);
    if (s.getLastRow() === 0) { s.appendRow(HEAD[n]); s.setFrozenRows(1); }
    else {   // upgrade older sheets: add any missing columns at the end
      const cur = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0], add = HEAD[n].filter(c => cur.indexOf(c) < 0);
      if (add.length) s.getRange(1, 1, 1, cur.length + add.length).setValues([cur.concat(add)]);
    }
  });
  const d = ss.getSheetByName("Sheet1");
  if (d && d.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(d);
  if (!find("Users", "email", ADMIN_EMAIL)) createUser(ADMIN_EMAIL, ADMIN_PASSWORD, "Admin", "admin", false);
  if (!all("Places").length) {
    [["mono1","Wayfinding Monolith 1","Landmark",14.2475,121.0480,"QR station and pilot starting point."],
     ["mall","Southwoods Mall","Mall",14.2498,121.0505,"Shopping, dining and services."],
     ["park","Township Park","Park",14.2510,121.0475,"Open green space with jogging loop access."]
    ].forEach(p => put("Places", {id:p[0],name:p[1],cat:p[2],lat:p[3],lng:p[4],desc:p[5],image:"",active:true}));
  }
  if (!all("Settings").length) {
    put("Settings",{key:"houseTitle",value:"Your ad could be next on MAPIGO"});
    put("Settings",{key:"houseText",value:"Reach everyone who scans in Southwoods. Digital banners start at PHP 250 per face, per day."});
    put("Settings",{key:"houseButton",value:"Advertise with us"});
    put("Settings",{key:"defaultSeconds",value:"8"});
  }
  if (!find("Settings", "key", "runPoints")) put("Settings",{key:"runPoints",value:"1:10,3:30,5:50,10:100"});
  if (!find("Settings", "key", "cities")) {
    put("Settings",{key:"cities",value:"Biñan,Carmona,Manila,Quezon City,Cebu,Other"});
  }
}

/* ---------- sheet helpers ---------- */
const hdr = n => { const s = sh(n); return s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0]; };
const sh = n => SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n);
const on = v => v === true || String(v).toUpperCase() === "TRUE" || v === 1 || v === "1";
const today = () => Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd");
const newId = () => Utilities.getUuid().slice(0, 8);
function all(n) {
  const v = sh(n).getDataRange().getValues(), h = v.shift();
  return v.map((r, i) => {
    const o = {_row: i + 2};
    h.forEach((k, j) => { let x = r[j]; if (x instanceof Date) x = Utilities.formatDate(x, TZ, "yyyy-MM-dd"); o[k] = x; });
    return o;
  });
}
const find = (n, k, v) => all(n).find(r => String(r[k]).toLowerCase() === String(v).toLowerCase());
function put(n, o) {
  const s = sh(n), h = hdr(n), key = n === "Settings" ? "key" : n === "Sessions" ? "token" : "id";
  const ex = all(n).find(r => String(r[key]) === String(o[key]));
  const row = h.map(k => {
    let x = o[k] !== undefined ? o[k] : (ex && ex[k] !== undefined ? ex[k] : "");
    if (typeof x === "string" && /^[=+\-@]/.test(x)) x = "'" + x;   // block spreadsheet formula injection
    return x;
  });
  if (ex) s.getRange(ex._row, 1, 1, h.length).setValues([row]); else s.appendRow(row);
}
function del(n, k, v) { const r = all(n).find(r => String(r[k]) === String(v)); if (r) sh(n).deleteRow(r._row); }

/* ---------- validation helpers ---------- */
const str = (s, max) => String(s == null ? "" : s).trim().slice(0, max);
const clampN = (v, lo, hi, def) => { v = Number(v); return isFinite(v) ? Math.max(lo, Math.min(hi, v)) : def; };
const url = s => /^https:\/\/[^\s'"()<>]+$/.test(String(s || "")) ? String(s) : "";
const linkUrl = s => /^https?:\/\/[^\s'"()<>]+$/.test(String(s || "")) ? String(s) : "";

/* ---------- auth ---------- */
const sha = s => Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s).map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, "0")).join("");
function hash(pw, salt) { let h = pw + salt; for (let i = 0; i < 50; i++) h = sha(h + salt); return h; }
function createUser(email, pw, name, role, showBoard, city) {
  const salt = Utilities.getUuid(), id = newId();
  put("Users", {id, email: String(email).toLowerCase(), name, city: city || "", photo: "", passHash: hash(pw, salt), salt, role, showBoard: !!showBoard,
    active: true, adj: 0, score: 0, streak: 0, stage: 0, state: "", created: new Date().toISOString()});
  return id;
}
function session(uid) { const t = Utilities.getUuid() + Utilities.getUuid(); put("Sessions", {token: t, userId: uid, exp: Date.now() + 30 * 864e5}); return t; }
function who(tok) {
  if (!tok) return null;
  const s = find("Sessions", "token", tok);
  if (!s || Number(s.exp) < Date.now()) return null;
  const u = find("Users", "id", s.userId);
  return u && on(u.active) ? u : null;
}
const need = d => { const u = who(d.token); if (!u) throw Error("Please log in again."); return u; };
const admin = d => { const u = need(d); if (u.role !== "admin") throw Error("Admins only."); return u; };
const pub = u => ({id: u.id, name: u.name, email: u.email, role: u.role, city: u.city || "", photo: url(u.photo), showBoard: on(u.showBoard)});
function throttle(k) {
  const c = CacheService.getScriptCache(), n = Number(c.get(k) || 0);
  if (n >= 8) throw Error("Too many attempts. Try again in 10 minutes.");
  c.put(k, n + 1, 600);
}
const parse = s => { try { return s ? JSON.parse(s) : null; } catch (e) { return null; } };

function runPts() {
  const m = {};
  String((find("Settings", "key", "runPoints") || {}).value || "1:10,3:30,5:50,10:100").split(",").forEach(p => { const a = p.split(":").map(Number); if (a[0] > 0 && a[1] >= 0) m[a[0]] = a[1]; });
  return m;
}
function runInfo(u) {
  const r = all("RunLog").filter(x => x.userId === u.id && x.status === "running").sort((a, b) => Number(b.startedAt) - Number(a.startedAt))[0];
  return {run: r ? {km: Number(r.km), startedAt: Number(r.startedAt)} : null, now: Date.now()};
}
function evList(tok) {
  const me = who(tok), t = today(), J = all("Joins");
  return all("Events").filter(e => on(e.active) && e.date >= t).sort((a, b) => String(a.date).localeCompare(String(b.date))).slice(0, 3)
    .map(e => ({id: e.id, title: e.title, date: e.date, time: e.time, mult: Number(e.mult) || 2, desc: e.desc, image: e.image,
      count: J.filter(j => j.eventId === e.id).length, joined: !!me && J.some(j => j.eventId === e.id && j.userId === me.id)}));
}

/* ---------- actions ---------- */
const A = {
  boot(d) {
    const t = today(), settings = {};
    all("Settings").forEach(r => settings[r.key] = r.value);
    return {
      settings,
      ads: all("Ads").filter(a => on(a.active) && (!a.start || a.start <= t) && (!a.end || a.end >= t))
        .sort((x, y) => (Number(x.order) || 0) - (Number(y.order) || 0))
        .map(a => ({id: a.id, advertiser: a.advertiser, title: a.title, desc: a.desc, cta: a.cta, image: a.image, link: a.link, seconds: Number(a.seconds) || 0})),
      events: evList(d.token),
      places: all("Places").filter(p => on(p.active))
        .map(p => ({id: p.id, name: p.name, cat: p.cat, lat: Number(p.lat), lng: Number(p.lng), desc: p.desc, image: p.image}))
    };
  },
  register(d) {
    const email = str(d.email, 120).toLowerCase(), name = str(d.name, 24), pw = String(d.password || ""), city = str(d.city, 30);
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw Error("Enter a valid email address.");
    if (name.length < 2) throw Error("Display name needs at least 2 characters.");
    if (city.length < 2) throw Error("Choose your city.");
    if (pw.length < 8) throw Error("Password needs at least 8 characters.");
    if (find("Users", "email", email)) throw Error("That email is already registered.");
    const id = createUser(email, pw, name, "user", !!d.showBoard, city);
    return {token: session(id), user: pub(find("Users", "id", id)), state: null};
  },
  login(d) {
    const email = str(d.email, 120).toLowerCase(), k = "login:" + email;
    throttle(k);
    const u = find("Users", "email", email);
    if (!u || hash(String(d.password || ""), u.salt) !== u.passHash) throw Error("Wrong email or password.");
    if (!on(u.active)) throw Error("This account is disabled. Contact the township office.");
    CacheService.getScriptCache().remove(k);
    return Object.assign({token: session(u.id), user: pub(u), state: parse(u.state)}, runInfo(u));
  },
  me(d) { const u = need(d); return Object.assign({user: pub(u), state: parse(u.state)}, runInfo(u)); },
  logout(d) { if (d.token) del("Sessions", "token", d.token); },
  sync(d) {
    const u = need(d), p = d.pub || {}, json = JSON.stringify(d.state || {});
    if (json.length > 45000) throw Error("Progress data too large.");
    const o = {id: u.id, state: json,
      score: Math.floor(clampN(p.score, 0, 1e6, 0)) + (Number(u.adj) || 0),
      streak: Math.floor(clampN(p.streak, 0, 9999, 0)), stage: Math.floor(clampN(p.stage, 0, 5, 0))};
    if (d.showBoard !== undefined) o.showBoard = !!d.showBoard;
    put("Users", o);
    return {score: o.score};
  },
  events(d) { return {events: evList(d.token)}; },
  join(d) {
    const u = need(d), e = find("Events", "id", d.eventId);
    if (!e || !on(e.active) || e.date < today()) throw Error("This Run Day is not open for joining.");
    put("Joins", {id: e.id + ":" + u.id, eventId: e.id, userId: u.id, name: u.name, created: new Date().toISOString()});
    return {events: evList(d.token)};
  },
  qrinfo(d) {
    const q = find("QR", "id", str(d.code, 20));
    if (!q || !on(q.active)) throw Error("This QR code is not active.");
    if (q.type === "station") put("QR", {id: q.id, scans: (Number(q.scans) || 0) + 1});
    return {qr: {type: q.type, km: Number(q.km) || 0, label: q.label, placeId: q.placeId}};
  },
  scan(d) {
    const u = need(d), q = find("QR", "id", str(d.code, 20));
    if (!q || !on(q.active)) throw Error("This QR code is not active.");
    const km = Number(q.km), now = Date.now();
    put("QR", {id: q.id, scans: (Number(q.scans) || 0) + 1});
    if (q.type === "start") {
      all("RunLog").filter(r => r.userId === u.id && r.status === "running").forEach(r => put("RunLog", {id: r.id, status: "abandoned"}));
      put("RunLog", {id: newId(), userId: u.id, name: u.name, km, startedAt: now, status: "running", startQr: q.id, day: today()});
      return {kind: "start", km, startedAt: now, now, label: q.label};
    }
    if (q.type === "finish") {
      const r = all("RunLog").filter(x => x.userId === u.id && x.status === "running").sort((a, b) => Number(b.startedAt) - Number(a.startedAt))[0];
      if (!r) throw Error("Scan the START QR first, then run, then scan the FINISH QR.");
      if (Number(r.km) !== km) throw Error("You started a " + r.km + " km run, but this is the " + km + " km finish line.");
      const secs = Math.round((now - Number(r.startedAt)) / 1000);
      if (secs > MAX_SECS) { put("RunLog", {id: r.id, status: "expired"}); throw Error("That run timed out after 6 hours. Scan the START QR again."); }
      if (secs < km * MIN_PACE) throw Error("Too fast! A " + km + " km run takes at least " + Math.ceil(km * MIN_PACE / 60) + " minutes. Keep going and scan the finish QR when you arrive.");
      if (all("RunLog").filter(x => x.userId === u.id && x.day === today() && x.status === "finished").length >= MAX_RUNS_DAY)
        throw Error("You reached today's limit of " + MAX_RUNS_DAY + " scored runs. Come back tomorrow!");
      const ev = all("Events").find(e => on(e.active) && e.date === today());
      const joined = !!ev && all("Joins").some(j => j.eventId === ev.id && j.userId === u.id);
      const mult = joined ? Math.max(1, Number(ev.mult) || 2) : 1, base = runPts()[km] !== undefined ? runPts()[km] : Math.round(km * 10), pts = base * mult;
      put("RunLog", {id: r.id, finishedAt: now, seconds: secs, pts, mult, event: joined ? ev.title : "", day: today(), status: "finished", finishQr: q.id});
      return {kind: "finish", km, seconds: secs, base, mult, pts, event: joined ? ev.title : "", runId: r.id, now};
    }
    throw Error("This QR is not a running QR.");
  },
  cancelrun(d) {
    const u = need(d);
    all("RunLog").filter(r => r.userId === u.id && r.status === "running").forEach(r => put("RunLog", {id: r.id, status: "abandoned"}));
  },
  profile(d) {
    const u = need(d), o = {id: u.id};
    if (d.name !== undefined) { const n = str(d.name, 24); if (n.length < 2) throw Error("Display name needs at least 2 characters."); o.name = n; }
    if (d.city !== undefined) { const c = str(d.city, 30); if (c.length < 2) throw Error("Choose your city."); o.city = c; }
    if (d.showBoard !== undefined) o.showBoard = !!d.showBoard;
    put("Users", o);
    return {user: pub(find("Users", "id", u.id))};
  },
  photo(d) {
    const u = need(d);
    if (d.remove) { put("Users", {id: u.id, photo: ""}); return {photo: ""}; }
    const m = String(d.mime || "");
    if (!/^image\/(jpeg|png|webp)$/.test(m)) throw Error("Use a JPG, PNG or WebP image.");
    const bytes = Utilities.base64Decode(String(d.b64 || ""));
    if (bytes.length > 400 * 1024) throw Error("Photo is too large.");
    const it = DriveApp.getFoldersByName(PHOTO_FOLDER), fo = it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
    const f = fo.createFile(Utilities.newBlob(bytes, m, "avatar-" + u.id + "-" + Date.now() + ".jpg"));
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const old = /\/d\/([\w-]+)/.exec(String(u.photo || ""));
    if (old) { try { DriveApp.getFileById(old[1]).setTrashed(true); } catch (e) {} }
    const photo = "https://lh3.googleusercontent.com/d/" + f.getId() + "=s256-c";
    put("Users", {id: u.id, photo});
    return {photo};
  },
  run(d) {
    const u = need(d), km = clampN(d.km, 0.5, 42, 0);
    if (!km) throw Error("Distance must be between 0.5 and 42 km.");
    put("Runs", {id: newId(), userId: u.id, name: u.name, km, link: linkUrl(d.link), pts: Math.round(clampN(d.pts, 0, 500, 0)), status: "pending", created: new Date().toISOString()});
  },
  board(d) {
    const me = who(d.token), city = str(d.city, 30).toLowerCase();
    const everyone = all("Users").filter(u => on(u.active) && on(u.showBoard))
      .sort((a, b) => (Number(b.score) || 0) - (Number(a.score) || 0) || String(a.created).localeCompare(String(b.created)));
    const row = (u, i) => ({rank: i + 1, name: u.name, city: u.city || "", photo: url(u.photo), score: Number(u.score) || 0,
      streak: Number(u.streak) || 0, stage: Number(u.stage) || 0, me: !!(me && me.id === u.id)});
    const rows = city ? everyone.filter(u => String(u.city || "").toLowerCase() === city) : everyone;
    const champs = {};
    everyone.forEach(u => { const c = u.city || "Other"; if (!champs[c]) champs[c] = Object.assign(row(u, 0), {city: c}); });
    const idx = me ? rows.findIndex(u => u.id === me.id) : -1;
    return {
      rows: rows.slice(0, 25).map(row), champs: Object.keys(champs).map(k => champs[k]),
      myRank: idx >= 0 ? idx + 1 : null, total: rows.length,
      mine: me ? {showBoard: on(me.showBoard), score: Number(me.score) || 0} : null
    };
  },
  track(d) {
    const a = find("Ads", "id", d.id);
    if (!a) return;
    const f = d.t === "click" ? "clicks" : "views";
    put("Ads", {id: a.id, [f]: (Number(a[f]) || 0) + 1});
  },

  /* ----- admin only ----- */
  list(d) {
    admin(d);
    if (["Ads","Places","Users","Runs","Settings","Events","QR","RunLog"].indexOf(d.sheet) < 0) throw Error("Unknown sheet.");
    let rows = all(d.sheet).map(r => { delete r._row; return r; });
    if (d.sheet === "Users") rows = rows.map(r => { delete r.passHash; delete r.salt; delete r.state; return r; });
    if (d.sheet === "Events") { const J = all("Joins"); rows = rows.map(r => { r.joined = J.filter(j => j.eventId === r.id).length; return r; }).sort((a, b) => String(b.date).localeCompare(String(a.date))); }
    if (d.sheet === "RunLog") rows = rows.filter(r => r.status !== "abandoned").sort((a, b) => (Number(b.startedAt) || 0) - (Number(a.startedAt) || 0)).slice(0, 300);
    if (d.sheet === "Runs") rows.sort((a, b) => (a.status === "pending" ? 0 : 1) - (b.status === "pending" ? 0 : 1) || String(b.created).localeCompare(String(a.created)));
    return {rows};
  },
  save(d) {
    const me = admin(d), n = d.sheet, r = d.row || {};
    if (n === "Ads") {
      const o = {id: r.id || newId(), advertiser: str(r.advertiser, 80), title: str(r.title, 100), desc: str(r.desc, 160), cta: str(r.cta, 24), image: url(r.image), link: linkUrl(r.link),
        seconds: clampN(r.seconds, 3, 60, 8), start: str(r.start, 10), end: str(r.end, 10), order: clampN(r.order, 0, 999, 1), active: !!r.active};
      if (!o.image) throw Error("Upload a poster image first.");
      if (!r.id) { o.views = 0; o.clicks = 0; }
      put("Ads", o); return {id: o.id};
    }
    if (n === "Places") {
      const lat = Number(r.lat), lng = Number(r.lng);
      if (!isFinite(lat) || !isFinite(lng) || !str(r.name, 80)) throw Error("Name, latitude and longitude are required.");
      const o = {id: r.id || newId(), name: str(r.name, 80), cat: str(r.cat, 30) || "Other", lat, lng, desc: str(r.desc, 400), image: url(r.image), active: !!r.active};
      put("Places", o); return {id: o.id};
    }
    if (n === "Events") {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(str(r.date, 10)) || !str(r.title, 80)) throw Error("Title and date are required.");
      const o = {id: r.id || newId(), title: str(r.title, 80), date: str(r.date, 10), time: str(r.time, 60), mult: clampN(r.mult, 1, 10, 2), desc: str(r.desc, 400), image: url(r.image), active: !!r.active};
      put("Events", o); return {id: o.id};
    }
    if (n === "QR") {
      const type = ["start","finish","station"].indexOf(r.type) >= 0 ? r.type : "";
      if (!type) throw Error("Choose a QR type.");
      const km = type === "station" ? 0 : clampN(r.km, 0.5, 100, 0);
      if (type !== "station" && !km) throw Error("Choose the distance.");
      if (type === "station" && !str(r.placeId, 40)) throw Error("Choose the map location for this station QR.");
      const o = {id: r.id || Utilities.getUuid().replace(/-/g, "").slice(0, 10), type, km, label: str(r.label, 80), placeId: str(r.placeId, 40), design: str(r.design, 1500), active: r.active === undefined ? true : !!r.active};
      if (!r.id) { o.scans = 0; o.created = new Date().toISOString(); }
      put("QR", o); return {id: o.id};
    }
    if (n === "Settings") {
      if (SETTING_KEYS.indexOf(r.key) < 0) throw Error("Unknown setting.");
      put("Settings", {key: r.key, value: r.key === "houseImage" ? url(r.value) : r.key === "siteUrl" ? linkUrl(r.value) : r.key === "runPoints" ? str(r.value, 80) : r.key === "houseLink" ? linkUrl(r.value) || str(r.value, 40) : str(r.value, 500)});
      return {};
    }
    if (n === "Users") {
      const t = find("Users", "id", r.id);
      if (!t) throw Error("User not found.");
      const adj = clampN(r.adj, -100000, 100000, 0);
      const o = {id: t.id, name: str(r.name, 24) || t.name, role: r.role === "admin" ? "admin" : "user", showBoard: !!r.showBoard, active: !!r.active,
        adj, score: Math.max(0, (Number(t.score) || 0) - (Number(t.adj) || 0) + adj)};
      if (r.removePhoto) o.photo = "";
      if (t.id === me.id && (o.role !== "admin" || !o.active)) throw Error("You cannot remove your own admin access.");
      put("Users", o); return {};
    }
    throw Error("Unknown sheet.");
  },
  remove(d) {
    admin(d);
    if (["Ads","Places","Runs","Events","QR","RunLog"].indexOf(d.sheet) < 0) throw Error("Cannot delete from this sheet.");
    del(d.sheet, "id", d.id);
  },
  decide(d) {
    admin(d);
    const r = find("Runs", "id", d.id);
    if (!r || r.status !== "pending") throw Error("This run was already reviewed.");
    put("Runs", {id: r.id, status: d.ok ? "approved" : "rejected"});
    if (!d.ok) {
      const u = find("Users", "id", r.userId), p = Number(r.pts) || 0;
      if (u) put("Users", {id: u.id, adj: (Number(u.adj) || 0) - p, score: Math.max(0, (Number(u.score) || 0) - p)});
    }
  },
  rejectrun(d) {
    admin(d);
    const r = find("RunLog", "id", d.id);
    if (!r || r.status !== "finished") throw Error("Only finished runs can be rejected.");
    put("RunLog", {id: r.id, status: "rejected"});
    const u = find("Users", "id", r.userId), p = Number(r.pts) || 0;
    if (u) put("Users", {id: u.id, adj: (Number(u.adj) || 0) - p, score: Math.max(0, (Number(u.score) || 0) - p)});
  },
  upload(d) {
    admin(d);
    const m = String(d.mime || "");
    if (!/^image\/(jpeg|png|webp|gif)$/.test(m)) throw Error("Images only (JPG, PNG, WebP, GIF).");
    const bytes = Utilities.base64Decode(String(d.b64 || ""));
    if (bytes.length > 6 * 1024 * 1024) throw Error("Image is larger than 6 MB.");
    const it = DriveApp.getFoldersByName(DRIVE_FOLDER), fo = it.hasNext() ? it.next() : DriveApp.createFolder(DRIVE_FOLDER);
    const f = fo.createFile(Utilities.newBlob(bytes, m, "mapigo-" + Date.now() + "." + m.split("/")[1].replace("jpeg", "jpg")));
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return {url: "https://lh3.googleusercontent.com/d/" + f.getId()};
  }
};

/* ---------- web app entry points ---------- */
function doGet() { return ContentService.createTextOutput("MAPIGO API is running."); }
function doPost(e) {
  const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse(e.postData.contents), f = A[d.a];
    if (!f) throw Error("Unknown action.");
    lock.waitLock(20000);
    return out(Object.assign({ok: true}, f.call(A, d) || {}));
  } catch (err) {
    return out({ok: false, error: String(err.message || err)});
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
