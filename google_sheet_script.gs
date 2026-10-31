/**
 * مزامنة بالاتجاهين بين مكتب التوظيف الذكي و Google Sheet.
 *
 * التبويبات (تُنشأ تلقائياً عند أول اتصال):
 *  - المرشحون: المتقدمون الذين لم يُعيَّنوا بعد.
 *  - الموظفين: من تم تعيينهم (نفس أعمدة المرشحين + تاريخ التعيين).
 *  - الوظائف: الوظائف والأقسام والعدد المطلوب، وهي مصدر لوحة المعلومات في التطبيق.
 *  - لوحة المعلومات: ملخص يتحدّث تلقائياً بالمعادلات من التبويبات الثلاثة (للعرض فقط).
 *  - المستخدمون: حسابات الدخول (كلمات المرور محفوظة مشفّرة في أعمدة مخفية).
 *
 * الحسابات والصلاحيات:
 *  - لا تُقرأ البيانات ولا تُعدَّل إلا بعد تسجيل الدخول في التطبيق.
 *  - أول من يفتح التطبيق بعد نشر الكود ينشئ حساب المدير، فأنشئه أنت فوراً.
 *  - المدير ينشئ الحسابات من صفحة "المستخدمون" في التطبيق ويحدد اسم المستخدم وكلمة المرور والصلاحية:
 *      مدير: كل شيء + إدارة المستخدمين | موارد بشرية: المرشحون والوظائف والتعيين | مشاهد: لوحة المعلومات فقط.
 *  - يمكن أيضاً تغيير "الصلاحية" (مدير / موارد بشرية / مشاهد) أو "مفعّل" (نعم / لا) من تبويب المستخدمين مباشرة.
 *
 * طريقة الإعداد (مرة واحدة):
 * 1) افتح https://sheets.new لإنشاء شيت جديد (أو افتح الشيت الحالي).
 * 2) من القائمة: Extensions (الإضافات) ← Apps Script.
 * 3) احذف الكود الموجود والصق هذا الملف كاملاً، ثم احفظ.
 * 4) Deploy (نشر) ← New deployment ← اختر النوع Web app:
 *      Execute as: Me
 *      Who has access: Anyone
 *    ثم Deploy، ووافق على الصلاحيات.
 * 5) انسخ رابط Web app (ينتهي بـ /exec) والصقه في التطبيق:
 *    إعدادات الوظيفة ← رابط Google Sheet ← حفظ الرابط.
 * إذا عدّلت هذا الكود لاحقاً: Deploy ← Manage deployments ← Edit ← Version: New version ← Deploy.
 * الشيت القديم (تبويب المرشحون مع عمود "تم التعيين") يُحوَّل تلقائياً: من كتب له "نعم" ينتقل لتبويب الموظفين.
 *
 * التعديل في الشيت (يظهر في التطبيق ولوحة المعلومات خلال ~20 ثانية):
 *  - المرشحون/الموظفين: عدّل أي خانة، أو أضف صفاً جديداً واترك id فارغاً (يُملأ تلقائياً)، أو احذف الصف كاملاً.
 *  - للتعيين من الشيت: قص الصف (Ctrl+X) من المرشحين والصقه في أول صف فارغ في الموظفين. وللإلغاء العكس.
 *  - أعمدة "القسم" و"درجة القبول" و"القرار" في المرشحين والموظفين يحسبها التطبيق، فلا تعدّلها.
 *  - عمود "الوظيفة" يجب أن يطابق اسم وظيفة في تبويب الوظائف.
 *  - الوظائف: عدّل القسم أو العدد المطلوب أو الحدود، أو أضف وظيفة بصف جديد.
 *    أعمدة "تم تعيينهم" و"المتبقي" و"الحالة" معادلات تلقائية، فلا تعدّلها.
 *    لا تُحذف وظيفة عليها مرشحون أو موظفون (تعود تلقائياً).
 *  - لوحة المعلومات: لإعادة بنائها احذف التبويب، ويُنشأ من جديد عند الاتصال التالي.
 *
 * ربط Excel بالشيت (يتحدّث تلقائياً، للقراءة فقط):
 *   Excel ← Data (بيانات) ← From Web (من ويب) ← الصق "رابط Excel" من صفحة المستخدمون في التطبيق ← OK ← Load (تحميل).
 *   ثم Data ← Queries & Connections ← كليك يمين على الاستعلام ← Properties:
 *   فعّل "Refresh every 1 minutes" و"Refresh data when opening the file".
 *   التعديل في Excel لا يرجع للشيت؛ عدّل في التطبيق أو في Google Sheet.
 * الموقع نفسه على نفس الرابط (لأي أحد):
 *   في Apps Script: + ← HTML ← سمّه index ← الصق محتوى site/index.html كاملاً ← احفظ ← انشر إصداراً جديداً.
 *   بعدها رابط /exec بدون إضافات يفتح التطبيق، ومتصل بهذا الشيت تلقائياً.
 */

const CAND_SHEET = "المرشحون";
const EMP_SHEET = "الموظفين";
const JOB_SHEET = "الوظائف";
const DASH_SHEET = "لوحة المعلومات";
const USER_SHEET = "المستخدمون";

const USER_KEYS = ["username", "name", "role", "active", "created", "last_login", "salt", "hash"];
const USER_HEADERS = ["اسم المستخدم", "الاسم", "الصلاحية", "مفعّل", "تاريخ الإنشاء", "آخر دخول", "salt", "hash"];
const ROLE_AR = { admin: "مدير", hr: "موارد بشرية", viewer: "مشاهد" };
const SESSION_SECONDS = 6 * 60 * 60; // الجلسة تنتهي بعد 6 ساعات بدون استخدام

// المرشحون والموظفون بنفس الأعمدة حتى يمكن نقل الصف بين التبويبين بالقص واللصق
// الأعمدة: A id, B التاريخ, C الوظيفة, D القسم, E الاسم ... L القرار ... O آخر تعديل
const PEOPLE_KEYS = ["id", "date", "job", "department", "name", "contact", "experience", "skill",
  "interview", "expected_salary", "score", "decision", "notes", "hired_at", "updated"];
const PEOPLE_HEADERS = ["id", "التاريخ", "الوظيفة", "القسم", "الاسم", "التواصل", "الخبرة (سنوات)", "المهارة (0-10)",
  "المقابلة (0-10)", "الراتب المتوقع", "درجة القبول", "القرار", "ملاحظات", "تاريخ التعيين", "آخر تعديل"];

// الأعمدة: A id, B الوظيفة, C القسم, D المطلوب, E تم تعيينهم (معادلة), F المتبقي (معادلة), G الحالة (معادلة)
const JOB_KEYS = ["id", "title", "department", "headcount", "hired", "remaining", "status",
  "budget", "hire", "shortlist", "minInterview", "minSkill", "updated"];
const JOB_HEADERS = ["id", "الوظيفة", "القسم", "العدد المطلوب", "تم تعيينهم", "المتبقي", "الحالة",
  "الميزانية الشهرية", "حد التوظيف", "حد القائمة المختصرة", "أدنى درجة مقابلة", "أدنى درجة مهارة", "آخر تعديل"];

const TABLES = {};
TABLES[CAND_SHEET] = { keys: PEOPLE_KEYS, headers: PEOPLE_HEADERS, label: "name" };
TABLES[EMP_SHEET] = { keys: PEOPLE_KEYS, headers: PEOPLE_HEADERS, label: "name" };
TABLES[JOB_SHEET] = { keys: JOB_KEYS, headers: JOB_HEADERS, label: "title" };

// أسماء الأعمدة المعروفة لكل تبويب (الحالية والقديمة) لتحويل شيت قديم بدون فقد بيانات
Object.keys(TABLES).forEach((n) => {
  const t = TABLES[n];
  t.aliases = { "تم التعيين": "hired" };
  t.headers.forEach((h, i) => (t.aliases[h] = t.keys[i]));
});

const isYes_ = (v) => /^(نعم|yes|true|1|✓)$/i.test(String(v).trim());
const blank_ = (v) => (v === null || v === undefined ? "" : v);

function headerRow_(sh) {
  const n = sh.getLastColumn();
  return n ? sh.getRange(1, 1, 1, n).getValues()[0].map((h) => String(h).trim()) : [];
}

function sameHeaders_(sh, headers) {
  const row = headerRow_(sh);
  return headers.every((h, i) => row[i] === h);
}

function dateOnly_(v) {
  const t = v instanceof Date ? v : new Date(Number(v));
  if (!v || isNaN(t.getTime())) return "";
  return Utilities.formatDate(t, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd");
}

// ينشئ التبويبات الناقصة ويصلح رؤوس الأعمدة. يعمل بسرعة إذا كان كل شيء جاهزاً.
function setup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ready = () => Object.keys(TABLES).every((n) => {
    const sh = ss.getSheetByName(n);
    return sh && sameHeaders_(sh, TABLES[n].headers);
  }) && ss.getSheetByName(DASH_SHEET) && ss.getSheetByName(USER_SHEET);
  if (ready()) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (ready()) return;
    const toEmployees = [];
    Object.keys(TABLES).forEach((name) => {
      const t = TABLES[name];
      let sh = ss.getSheetByName(name);
      if (!sh) { sh = ss.insertSheet(name); sh.setRightToLeft(true); }
      if (sameHeaders_(sh, t.headers)) return;
      const old = headerRow_(sh);
      const known = old.length && old.every((h) => h === "" || t.aliases[h]);
      if (known && sh.getLastRow() > 1) {
        // تخطيط قديم: نعيد ترتيب الأعمدة حسب أسمائها
        const values = sh.getRange(2, 1, sh.getLastRow() - 1, old.length).getValues();
        const objs = values.filter((v) => String(v[0]).trim()).map((v) => {
          const o = {};
          old.forEach((h, i) => { if (h) o[t.aliases[h]] = v[i]; });
          return o;
        });
        let keep = objs;
        if (name === CAND_SHEET) {
          keep = objs.filter((o) => !isYes_(o.hired));
          objs.filter((o) => isYes_(o.hired)).forEach((o) => toEmployees.push(o));
        }
        sh.clearContents();
        if (keep.length) sh.getRange(2, 1, keep.length, t.keys.length).setValues(keep.map((o) => t.keys.map((k) => blank_(o[k]))));
      }
      // رؤوس غير معروفة (أحد غيّر اسم عمود): نعيد الأسماء فقط ولا نحرّك البيانات
      sh.getRange(1, 1, 1, t.headers.length).setValues([t.headers]).setFontWeight("bold");
      sh.setFrozenRows(1);
    });
    if (toEmployees.length) {
      const emp = ss.getSheetByName(EMP_SHEET);
      toEmployees.forEach((o) => {
        if (!o.hired_at) o.hired_at = dateOnly_(o.updated);
        emp.appendRow(PEOPLE_KEYS.map((k) => blank_(o[k])));
      });
    }
    jobFormulas_(ss.getSheetByName(JOB_SHEET));
    if (!ss.getSheetByName(DASH_SHEET)) buildDashboard_(ss);
    if (!ss.getSheetByName(USER_SHEET)) {
      const us = ss.insertSheet(USER_SHEET);
      us.setRightToLeft(true);
      us.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]).setFontWeight("bold");
      us.setFrozenRows(1);
      us.hideColumns(USER_KEYS.indexOf("salt") + 1, 2);
    }
  } finally {
    lock.releaseLock();
  }
}

// أعمدة "تم تعيينهم" و"المتبقي" و"الحالة" في الوظائف: تُحسب من تبويب الموظفين
function jobFormulas_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const E = "'" + EMP_SHEET + "'";
  const f = [];
  for (let r = 2; r <= last; r++) {
    f.push([
      `=IF(B${r}="","",COUNTIF(${E}!C:C,B${r}))`,
      `=IF(B${r}="","",MAX(0,N(D${r})-E${r}))`,
      `=IF(B${r}="","",IF(E${r}>=N(D${r}),"✓ مكتمل","غير مكتمل"))`,
    ]);
  }
  sh.getRange(2, JOB_KEYS.indexOf("hired") + 1, f.length, 3).setFormulas(f);
}

function buildDashboard_(ss) {
  const sh = ss.insertSheet(DASH_SHEET, 0);
  sh.setRightToLeft(true);
  const J = "'" + JOB_SHEET + "'", E = "'" + EMP_SHEET + "'", C = "'" + CAND_SHEET + "'";
  sh.getRange("A1").setValue("لوحة المعلومات — تتحدّث تلقائياً من تبويبات الوظائف والموظفين والمرشحين (لا تعدّلها)")
    .setFontWeight("bold").setFontSize(13);
  const kpis = [
    ["المؤشر", "القيمة"],
    ["عدد الأقسام", `=IFERROR(COUNTUNIQUE(FILTER(${J}!C2:C,${J}!B2:B<>"",${J}!C2:C<>"")),0)`],
    ["عدد الوظائف", `=COUNTIF(${J}!B2:B,"?*")`],
    ["وظائف مكتملة الاحتياج", `=COUNTIF(${J}!G2:G,"✓ مكتمل")`],
    ["العدد المطلوب", `=SUM(${J}!D2:D)`],
    ["تم تعيينهم من الاحتياج", `=SUM(${J}!D2:D)-SUM(${J}!F2:F)`],
    ["شاغر متبقٍ", `=SUM(${J}!F2:F)`],
    ["نسبة اكتمال الاحتياج", `=IFERROR((SUM(${J}!D2:D)-SUM(${J}!F2:F))/SUM(${J}!D2:D),1)`],
    ["إجمالي الموظفين", `=COUNTIF(${E}!A2:A,"?*")`],
    ["المرشحون (لم يُعيَّنوا)", `=COUNTIF(${C}!A2:A,"?*")`],
  ];
  sh.getRange(3, 1, kpis.length, 2).setValues(kpis);
  sh.getRange("B10").setNumberFormat("0%");
  sh.getRange("A3:B3").setFontWeight("bold").setBackground("#eef2f7");

  sh.getRange("D3").setFormula(`=QUERY(${J}!A2:G,"select C, count(B), sum(D), sum(E), sum(F) where B is not null group by C ` +
    `label C 'القسم', count(B) 'الوظائف', sum(D) 'المطلوب', sum(E) 'تم تعيينهم', sum(F) 'المتبقي'",0)`);
  sh.getRange("J3").setFormula(`=IFERROR(QUERY(${C}!A2:L,"select L, count(A) where A is not null group by L ` +
    `label L 'قرار المرشحين', count(A) 'العدد'",0),"لا يوجد مرشحون")`);
  sh.getRange("D3:H3").setFontWeight("bold").setBackground("#eef2f7");
  sh.getRange("J3:K3").setFontWeight("bold").setBackground("#eef2f7");
  sh.setColumnWidth(1, 190);
  sh.setColumnWidth(4, 160);
  sh.setColumnWidth(10, 150);

  const chart = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange("D3:D40"))
    .addRange(sh.getRange("F3:G40"))
    .setNumHeaders(1)
    .setOption("title", "الاحتياج مقابل التعيين حسب القسم")
    .setOption("legend", { position: "bottom" })
    .setPosition(15, 4, 0, 0)
    .build();
  sh.insertChart(chart);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function readRows_(name) {
  const t = TABLES[name];
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  const last = sh.getLastRow();
  const values = last > 1 ? sh.getRange(2, 1, last - 1, t.keys.length).getValues() : [];
  return values
    .filter((v) => String(v[0]).trim())
    .map((v) => {
      const row = {};
      t.keys.forEach((k, i) => (row[k] = v[i] instanceof Date ? v[i].toISOString() : v[i]));
      return row;
    });
}

// كل الأشخاص: hired = نعم لمن في تبويب الموظفين
function people_() {
  const out = [];
  [CAND_SHEET, EMP_SHEET].forEach((name) => readRows_(name).forEach((r) => {
    r.hired = name === EMP_SHEET ? "نعم" : "لا";
    out.push(r);
  }));
  return out;
}

// نسخة CSV لـ Excel (Data ← From Web ← الرابط مع ?action=csv)، بدون عمودي id و"آخر تعديل"
function csv_() {
  const cell = (x) => {
    const s = String(blank_(x));
    return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const keys = PEOPLE_KEYS.filter((k) => k !== "id" && k !== "updated");
  const head = keys.map((k) => PEOPLE_HEADERS[PEOPLE_KEYS.indexOf(k)]).concat(["الحالة"]);
  const rows = [head].concat(people_().map((r) => keys.map((k) => r[k]).concat([r.hired === "نعم" ? "موظف" : "مرشح"])));
  // BOM حتى يقرأ Excel العربية بشكل صحيح
  return ContentService.createTextOutput("﻿" + rows.map((r) => r.map(cell).join(",")).join("\r\n"))
    .setMimeType(ContentService.MimeType.CSV);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === "status") {
    setup_();
    return json_({ ok: true, hasUsers: users_().length > 0 });
  }
  if (p.action === "csv") {
    const key = PropertiesService.getScriptProperties().getProperty("CSV_KEY");
    if (!key || p.key !== key) return ContentService.createTextOutput("غير مصرح: استخدم رابط Excel من صفحة المستخدمون في التطبيق.");
    setup_();
    return csv_();
  }
  // بدون action: يعرض الموقع نفسه (ملف index في مشروع Apps Script) لأي أحد يفتح الرابط
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("مكتب التوظيف الذكي")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function ids_(sh) {
  const last = sh.getLastRow();
  return last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map((r) => String(r[0])) : [];
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------- الحسابات ----------
const norm_ = (u) => String(u || "").trim().toLowerCase();
const roleKey_ = (v) => {
  const s = String(v).trim();
  return Object.keys(ROLE_AR).find((k) => k === s.toLowerCase() || ROLE_AR[k] === s) || "viewer";
};

function hash_(salt, password) {
  let h = salt + "|" + password;
  for (let i = 0; i < 300; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8));
  }
  return h;
}

function users_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET);
  const last = sh ? sh.getLastRow() : 0;
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, USER_KEYS.length).getValues()
    .map((v, i) => {
      const u = { row: i + 2 };
      USER_KEYS.forEach((k, j) => (u[k] = v[j]));
      u.username = norm_(u.username);
      u.role = roleKey_(u.role);
      u.active = String(u.active).trim() !== "لا";
      return u;
    })
    .filter((u) => u.username);
}

const findUser_ = (username) => users_().find((u) => u.username === norm_(username)) || null;
const public_ = (u) => ({ username: u.username, name: u.name || u.username, role: u.role });
const iso_ = (d) => (d instanceof Date ? d.toISOString() : d || "");

function checkNew_(username, password, needPassword) {
  if (!/^\S{3,40}$/.test(username)) throw new Error("اسم المستخدم من 3 إلى 40 حرفاً بدون مسافات.");
  if ((needPassword || password) && String(password || "").length < 8) throw new Error("كلمة المرور 8 أحرف على الأقل.");
}

function setPassword_(sh, row, password) {
  const salt = Utilities.getUuid();
  sh.getRange(row, USER_KEYS.indexOf("salt") + 1, 1, 2).setValues([[salt, hash_(salt, password)]]);
}

function addUser_(u) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET);
  const username = norm_(u.username);
  checkNew_(username, u.password, true);
  if (findUser_(username)) throw new Error("اسم المستخدم مستخدم من قبل.");
  const salt = Utilities.getUuid();
  sh.appendRow([username, String(u.name || "").trim() || username, ROLE_AR[roleKey_(u.role)], "نعم", new Date(), "", salt, hash_(salt, String(u.password))]);
}

function newToken_(username) {
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "");
  CacheService.getScriptCache().put("t:" + token, username, SESSION_SECONDS);
  return token;
}

// الجلسة صالحة ما دام الحساب موجوداً ومفعّلاً (الإيقاف أو الحذف يسري فوراً)
function session_(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const username = cache.get("t:" + token);
  if (!username) return null;
  const u = findUser_(username);
  if (!u || !u.active) return null;
  cache.put("t:" + token, username, SESSION_SECONDS);
  return u;
}

function login_(b) {
  const username = norm_(b.username);
  const cache = CacheService.getScriptCache(), fk = "f:" + username;
  const fails = Number(cache.get(fk) || 0);
  if (fails >= 5) throw new Error("محاولات دخول كثيرة. حاول بعد 15 دقيقة.");
  const u = findUser_(username);
  if (!u || hash_(String(u.salt), String(b.password || "")) !== u.hash) {
    cache.put(fk, String(fails + 1), 900);
    throw new Error("اسم المستخدم أو كلمة المرور غير صحيحة.");
  }
  if (!u.active) throw new Error("هذا الحساب موقوف. تواصل مع المدير.");
  cache.remove(fk);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET).getRange(u.row, USER_KEYS.indexOf("last_login") + 1).setValue(new Date());
  return { ok: true, token: newToken_(u.username), user: public_(u) };
}

function activeAdmins_(list) {
  return list.filter((u) => u.role === "admin" && u.active).length;
}

function saveUser_(b, me) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET);
  const u = b.user || {};
  if (b.create) { addUser_(u); return { ok: true }; }
  const list = users_();
  const cur = list.find((x) => x.username === norm_(u.username));
  if (!cur) throw new Error("المستخدم غير موجود.");
  const next = {
    name: u.name !== undefined ? String(u.name).trim() || cur.username : cur.name,
    role: u.role !== undefined ? roleKey_(u.role) : cur.role,
    active: u.active !== undefined ? !!u.active : cur.active,
  };
  if (cur.username === me.username && (next.role !== "admin" || !next.active)) throw new Error("لا يمكنك إيقاف حسابك أو تغيير صلاحيته.");
  const after = list.map((x) => (x === cur ? Object.assign({}, x, next) : x));
  if (!activeAdmins_(after)) throw new Error("يجب أن يبقى مدير واحد مفعّل على الأقل.");
  if (u.password) { checkNew_(cur.username, u.password, true); setPassword_(sh, cur.row, String(u.password)); }
  sh.getRange(cur.row, 2, 1, 3).setValues([[next.name, ROLE_AR[next.role], next.active ? "نعم" : "لا"]]);
  return { ok: true };
}

function deleteUser_(b, me) {
  const list = users_();
  const cur = list.find((x) => x.username === norm_(b.username));
  if (!cur) return { ok: true };
  if (cur.username === me.username) throw new Error("لا يمكنك حذف حسابك.");
  if (!activeAdmins_(list.filter((x) => x !== cur))) throw new Error("يجب أن يبقى مدير واحد مفعّل على الأقل.");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET).deleteRow(cur.row);
  return { ok: true };
}

// الشخص يُكتب في تبويب الموظفين إذا hired = نعم، وإلا في المرشحين، ويُحذف من التبويب الآخر
function write_(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const del = body.action === "delete";
  if (body.kind === "jobs") {
    const sh = ss.getSheetByName(JOB_SHEET);
    const ids = ids_(sh);
    (body.rows || []).forEach((row) => {
      const at = ids.indexOf(String(row.id));
      if (del) {
        if (at >= 0) { sh.deleteRow(at + 2); ids.splice(at, 1); }
        return;
      }
      const values = JOB_KEYS.map((k) => blank_(row[k]));
      if (at >= 0) sh.getRange(at + 2, 1, 1, values.length).setValues([values]);
      else { sh.appendRow(values); ids.push(String(row.id)); }
    });
    jobFormulas_(sh);
    return { ok: true };
  }
  const tabs = [ss.getSheetByName(CAND_SHEET), ss.getSheetByName(EMP_SHEET)];
  const idx = tabs.map(ids_);
  (body.rows || []).forEach((row) => {
    const target = del ? -1 : isYes_(row.hired) ? 1 : 0;
    tabs.forEach((sh, t) => {
      const ids = idx[t];
      const at = ids.indexOf(String(row.id));
      if (t === target) {
        const values = PEOPLE_KEYS.map((k) => blank_(row[k]));
        if (at >= 0) sh.getRange(at + 2, 1, 1, values.length).setValues([values]);
        else { sh.appendRow(values); ids.push(String(row.id)); }
      } else if (at >= 0) {
        sh.deleteRow(at + 2); ids.splice(at, 1);
      }
    });
  });
  return { ok: true };
}

// كل الطلبات من التطبيق: { action, token, ... }. البيانات لا تُقرأ ولا تُعدَّل بدون جلسة صالحة
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: "طلب غير صالح." }); }
  try {
    setup_();
    const a = body.action;
    if (a === "login") return json_(login_(body));
    if (a === "setup") {
      return json_(withLock_(() => {
        if (users_().length) throw new Error("تم إنشاء حساب المدير من قبل. سجّل الدخول.");
        addUser_({ username: body.username, name: body.name, password: body.password, role: "admin" });
        return login_(body);
      }));
    }
    const me = session_(body.token);
    if (!me) return json_({ ok: false, auth: true, error: "انتهت الجلسة. سجّل الدخول من جديد." });
    const deny = () => json_({ ok: false, error: "صلاحيتك لا تسمح بهذا الإجراء." });

    if (a === "me") return json_({ ok: true, user: public_(me) });
    if (a === "logout") { CacheService.getScriptCache().remove("t:" + body.token); return json_({ ok: true }); }
    if (a === "password") {
      if (hash_(String(me.salt), String(body.old || "")) !== me.hash) throw new Error("كلمة المرور الحالية غير صحيحة.");
      checkNew_(me.username, body.password, true);
      withLock_(() => setPassword_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET), me.row, String(body.password)));
      return json_({ ok: true });
    }
    if (a === "list") {
      return json_({ ok: true, rows: people_(), jobs: readRows_(JOB_SHEET), user: public_(me),
        sheetUrl: me.role === "admin" ? SpreadsheetApp.getActiveSpreadsheet().getUrl() : "" });
    }
    if (a === "upsert" || a === "delete") {
      if (me.role === "viewer") return deny();
      return json_(withLock_(() => write_(body)));
    }

    if (me.role !== "admin") return deny();
    if (a === "users") {
      return json_({ ok: true, users: users_().map((u) => Object.assign(public_(u),
        { active: u.active, created: iso_(u.created), last_login: iso_(u.last_login) })) });
    }
    if (a === "user_save") return json_(withLock_(() => saveUser_(body, me)));
    if (a === "user_delete") return json_(withLock_(() => deleteUser_(body, me)));
    if (a === "csv_key") {
      const props = PropertiesService.getScriptProperties();
      let key = props.getProperty("CSV_KEY");
      if (!key || body.reset) { key = Utilities.getUuid().replace(/-/g, ""); props.setProperty("CSV_KEY", key); }
      return json_({ ok: true, key: key });
    }
    return json_({ ok: false, error: "إجراء غير معروف." });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

// عند تعديلك اليدوي في الشيت: يسجّل وقت التعديل، ويعطي الصف الجديد معرّفاً
function onEdit(e) {
  const range = e.range;
  const sh = range.getSheet();
  const name = sh.getName();
  const t = TABLES[name];
  if (!t) return;
  const updCol = t.keys.indexOf("updated") + 1;
  const labelCol = t.keys.indexOf(t.label) + 1;
  const first = Math.max(2, range.getRow());
  const lastRow = range.getRow() + range.getNumRows() - 1;
  if (lastRow < first) return;
  // تعديل عمود "آخر تعديل" نفسه لا يُحتسب
  if (range.getColumn() === updCol && range.getNumColumns() === 1) return;

  const n = lastRow - first + 1;
  const data = sh.getRange(first, 1, n, t.keys.length).getValues();
  const now = Date.now();
  const hiredCol = PEOPLE_KEYS.indexOf("hired_at") + 1;
  data.forEach((v, i) => {
    const r = first + i;
    if (v.every((x) => x === "" || x === null)) return;
    if (!String(v[0]).trim() && String(v[labelCol - 1]).trim()) {
      sh.getRange(r, 1).setValue("s" + now.toString(36) + Math.random().toString(36).slice(2, 7));
    }
    if (name === EMP_SHEET && !String(v[hiredCol - 1]).trim()) sh.getRange(r, hiredCol).setValue(dateOnly_(now));
    sh.getRange(r, updCol).setValue(now);
  });
  if (name === JOB_SHEET) jobFormulas_(sh);
}
