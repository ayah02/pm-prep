/**
 * PM Prep – PMP practice and mock-exam system (PMI Examination Content Outline, July 2026)
 * Student project – CIS 386, Group G5, Al Yamamah University
 *
 * Runs as a Google Apps Script web app bound to the "PM Prep database" Google Sheet.
 * Answer keys never leave the server until a question is answered (practice) or the exam is submitted.
 */

var CONFIG = {
  APP_NAME: 'PM Prep',
  FOOTER: 'Student project – CIS 386 Project Management, Group G5 – Al Yamamah University',
  LOGO_URL: '',            // optional: paste a public image URL of your logo here
  PASS_MARK: 0.70,         // readiness / pass line (report Table 7.1)
  DOMAIN_FLAG: 0.65,       // domain accuracy flag (report Table 7.1)
  MODES: {
    // 1.33 minutes per item, as on the real PMP exam (240 minutes for 180 items)
    course: { label: 'Course mock exam', total: 60, minutes: 80, blueprint: { 1: 20, 2: 25, 3: 15 }, useExamSet: true },
    full:   { label: 'Full-length PMP simulation', total: 180, minutes: 240, blueprint: { 1: 59, 2: 74, 3: 47 } },
    quick:  { label: 'Quick quiz', total: 20, minutes: 27, blueprint: { 1: 7, 2: 8, 3: 5 } }
  },
  // PMP exam domains and their weight on the real exam (PMI Examination Content Outline, July 2026)
  DOMAINS: { 1: 'People', 2: 'Process', 3: 'Business Environment' },
  WEIGHTS: { 1: 0.33, 2: 0.41, 3: 0.26 },
  APPROACHES: ['Predictive', 'Agile', 'Hybrid', 'Any'],   // delivery approach of each item ('Any' = approach-neutral)
  LEVELS: ['Easy', 'Medium', 'Hard'],  // difficulty levels used in the bank, practice filter, results and dashboard
  PMP_TASKS: {
    'I.1': 'Develop a common vision', 'I.2': 'Manage conflicts', 'I.3': 'Lead the project team', 'I.4': 'Engage stakeholders',
    'I.5': 'Align stakeholder expectations', 'I.6': 'Manage stakeholder expectations', 'I.7': 'Help ensure knowledge transfer', 'I.8': 'Plan and manage communication',
    'II.1': 'Develop an integrated project management plan and plan delivery', 'II.2': 'Develop and manage project scope', 'II.3': 'Help ensure value-based delivery',
    'II.4': 'Plan and manage resources', 'II.5': 'Plan and manage procurement', 'II.6': 'Plan and manage finance', 'II.7': 'Plan and optimize quality of products/deliverables',
    'II.8': 'Plan and manage schedule', 'II.9': 'Evaluate project status', 'II.10': 'Manage project closure',
    'III.1': 'Define and establish project governance', 'III.2': 'Plan and manage project compliance', 'III.3': 'Manage and control changes',
    'III.4': 'Remove impediments and manage issues', 'III.5': 'Plan and manage risk', 'III.6': 'Continuous improvement', 'III.7': 'Support organizational change',
    'III.8': 'Evaluate external business environment changes'
  }
};

var SHEETS = { Q: 'Questions', R: 'Responses', A: 'Answers', S: 'Sessions', P: 'Practice', F: 'Flags', U: 'Students' };

/* ------------------------------------------------------------------ web app entry */

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) === 'admin' ? 'Admin' : 'Index';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle(CONFIG.APP_NAME + (page === 'Admin' ? ' – Instructor dashboard' : ''))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ------------------------------------------------------------------ JSON API (standalone site, e.g. GitHub Pages)
 * The public site POSTs {fn, args} as text/plain (no CORS preflight, no Google cookies),
 * so it works for visitors who are signed in to several Google accounts.
 */
var API_ = {
  getConfig: getConfig, getPracticeSet: getPracticeSet, checkPractice: checkPractice,
  startExam: startExam, submitExam: submitExam, getProgress: getProgress,
  getAttemptResult: getAttemptResult, flagQuestion: flagQuestion,
  adminLogin: adminLogin, adminStats: adminStats, adminExport: adminExport, adminAddQuestions: adminAddQuestions,
  studentLogin: studentLogin
};

function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var fn = API_.hasOwnProperty(req.fn) ? API_[req.fn] : null;
    if (!fn) throw new Error('Unknown request.');
    var args = Array.isArray(req.args) ? req.args.slice(0, 6) : [];
    out = { ok: true, result: fn.apply(null, args) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ helpers */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet "' + name + '". Do not rename the sheets.');
  return sh;
}

function table_(name) {
  var values = sheet_(name).getDataRange().getValues();
  var head = values.shift();
  var idx = {};
  head.forEach(function (h, i) { idx[String(h).trim()] = i; });
  return { head: head, idx: idx, rows: values };
}

function append_(name, rowsToAdd) {
  if (!rowsToAdd.length) return;
  var sh = sheet_(name);
  var start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
}

function iso_(d) {
  if (!(d instanceof Date)) return d === '' || d == null ? '' : String(d);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function cleanId_(id) {
  id = String(id || '');
  if (!/^[A-Za-z0-9-]{8,48}$/.test(id)) throw new Error('Invalid session. Please reload the page.');
  return id;
}

function cleanText_(s, max) { return String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max); }

function shuffle_(arr) {
  for (var i = arr.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var t = arr[i]; arr[i] = arr[j]; arr[j] = t;
  }
  return arr;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ------------------------------------------------------------------ question bank */

function pmpDomainOf_(task) { return { I: 'People', II: 'Process', III: 'Business Environment' }[String(task).split('.')[0]] || 'Process'; }


var QCACHE_ = null;

function questions_() {
  if (QCACHE_) return QCACHE_;
  var t = table_(SHEETS.Q), I = t.idx, list = [];
  t.rows.forEach(function (r) {
    var id = String(r[I.QuestionID] || '').trim();
    if (!id) return;
    var active = r[I.Active];
    if (active === false || String(active).toUpperCase() === 'FALSE') return;
    var key = String(r[I.Correct_Option] || '').trim().toUpperCase();
    if ('ABCD'.indexOf(key) < 0 || key.length !== 1) return;
    var eco = String(r[I.ECO_Task] || '');
    // PMP task code, e.g. 'III.5' (from the PMP_Task column, else the start of ECO_Task)
    var pmpTask = (I.PMP_Task !== undefined && String(r[I.PMP_Task] || '').trim()) || eco.split(' ')[0];
    list.push({
      id: id,
      status: String(r[I.Status] || 'Bank'),
      domain: Number(r[I.Domain]) || 0,
      eco: eco,
      lecture: Number(r[I.Lecture]) || 0,
      topic: String(r[I.Lecture_Topic] || ''),
      difficulty: String(r[I.Difficulty] || ''),
      pmpTask: pmpTask, pmp: pmpDomainOf_(pmpTask),
      approach: I.Approach !== undefined ? String(r[I.Approach] || 'Any') : 'Any',
      itemType: I.Item_Type !== undefined ? String(r[I.Item_Type] || '') : '',
      text: String(r[I.Question] || ''),
      options: [r[I.Option_A], r[I.Option_B], r[I.Option_C], r[I.Option_D]].map(String),
      key: key,
      explanation: String(r[I.Explanation] || ''),
      taught: String(r[I.In_Lectures] || '')
    });
  });
  QCACHE_ = list;
  return list;
}

function byId_() {
  var m = {};
  questions_().forEach(function (q) { m[q.id] = q; });
  return m;
}

function publicQ_(q) {   // what the browser may see before answering
  return { id: q.id, domain: q.domain, lecture: q.lecture, topic: q.topic, difficulty: q.difficulty, pmpTask: q.pmpTask,
           approach: q.approach, itemType: q.itemType, text: q.text, options: q.options };
}

function lectureNames_() {   // 'lecture' = PMP task number (1–26); its name is the task, e.g. 'I.2 Manage conflicts'
  var out = {};
  questions_().forEach(function (q) { if (!out[q.lecture]) out[q.lecture] = q.topic; });
  return out;
}

/* ------------------------------------------------------------------ public API: config */

function getConfig() {
  var qs = questions_(), names = lectureNames_(), lectures = {};
  qs.forEach(function (q) {
    if (!lectures[q.lecture]) lectures[q.lecture] = { lecture: q.lecture, domain: q.domain, name: names[q.lecture], count: 0 };
    lectures[q.lecture].count++;
  });
  // what the bank holds: domain x level counts, and the level mix of the fixed course exam set
  var bank = {}, examLv = {}, bankAp = {}, examAp = {};
  qs.forEach(function (q) {
    bankAp[q.approach] = (bankAp[q.approach] || 0) + 1;
    if (q.status === 'Exam') examAp[q.approach] = (examAp[q.approach] || 0) + 1;
    bank[q.domain] = bank[q.domain] || {}; bank[q.domain][q.difficulty] = (bank[q.domain][q.difficulty] || 0) + 1;
    if (q.status === 'Exam') examLv[q.difficulty] = (examLv[q.difficulty] || 0) + 1;
  });
  var modes = {};
  Object.keys(CONFIG.MODES).forEach(function (k) {
    var m = CONFIG.MODES[k];
    modes[k] = { label: m.label, total: m.total, minutes: m.minutes, blueprint: m.blueprint,
                 levelMix: m.useExamSet ? examLv : null, approachMix: m.useExamSet ? examAp : null };
  });
  return {
    appName: CONFIG.APP_NAME, footer: CONFIG.FOOTER, logoUrl: CONFIG.LOGO_URL,
    passMark: CONFIG.PASS_MARK, domains: CONFIG.DOMAINS, modes: modes,
    lectures: Object.keys(lectures).map(function (k) { return lectures[k]; }).sort(function (a, b) { return a.lecture - b.lecture; }),
    bankSize: qs.length,
    levels: CONFIG.LEVELS,
    weights: CONFIG.WEIGHTS,
    approaches: CONFIG.APPROACHES.map(function (a) { return { approach: a, bank: bankAp[a] || 0 }; }),
    bankMatrix: [1, 2, 3].map(function (d) { return { domain: d, counts: CONFIG.LEVELS.map(function (l) { return (bank[d] && bank[d][l]) || 0; }) }; })
  };
}

/* ------------------------------------------------------------------ practice */

function getPracticeSet(studentId, filters) {
  studentId = who_(studentId);
  filters = filters || {};
  var count = Math.min(Math.max(Number(filters.count) || 10, 5), 50);
  var pool = questions_().filter(function (q) {
    if (filters.domain && q.domain !== Number(filters.domain)) return false;
    if (filters.lecture && q.lecture !== Number(filters.lecture)) return false;
    if (filters.difficulty && q.difficulty !== filters.difficulty) return false;
    if (filters.approach && q.approach !== filters.approach) return false;
    if (filters.scenarioOnly && q.itemType && q.itemType !== 'Scenario' && q.itemType !== 'Calculation') return false;
    return true;
  });
  if (filters.weak) {
    var weak = weakLectures_(studentId);
    if (weak.length) {
      var set = {}; weak.forEach(function (w) { set[w.lecture] = true; });
      var focused = pool.filter(function (q) { return set[q.lecture]; });
      if (focused.length >= 5) pool = focused;
    }
  }
  return shuffle_(pool.slice()).slice(0, count).map(publicQ_);
}

function checkPractice(studentId, questionId, selected) {
  studentId = who_(studentId);
  selected = String(selected || '').toUpperCase();
  if ('ABCD'.indexOf(selected) < 0 || selected.length !== 1) throw new Error('Choose an option first.');
  var q = byId_()[questionId];
  if (!q) throw new Error('Question not found.');
  var ok = selected === q.key;
  withLock_(function () { append_(SHEETS.P, [[new Date(), studentId, q.id, selected, ok ? 1 : 0]]); });
  return { correct: ok, key: q.key, explanation: q.explanation };
}

/* ------------------------------------------------------------------ exams */

function pickExam_(mode) {
  var m = CONFIG.MODES[mode];
  var all = questions_(), chosen = [];
  [1, 2, 3].forEach(function (d) {
    var need = m.blueprint[d] || 0;
    var dom = all.filter(function (q) { return q.domain === d; });
    var picked = [];
    if (m.useExamSet) picked = shuffle_(dom.filter(function (q) { return q.status === 'Exam'; })).slice(0, need);
    if (picked.length < need) {
      var used = {}; picked.forEach(function (q) { used[q.id] = true; });
      var rest = shuffle_(dom.filter(function (q) { return !used[q.id]; }));
      // prefer reserves first when topping up the course exam
      rest.sort(function (a, b) { return (a.status === 'Reserve' ? 0 : 1) - (b.status === 'Reserve' ? 0 : 1); });
      picked = picked.concat(rest.slice(0, need - picked.length));
    }
    chosen = chosen.concat(picked);
  });
  return shuffle_(chosen);
}

function startExam(studentId, nickname, mode) {
  studentId = who_(studentId);
  if (!CONFIG.MODES[mode]) throw new Error('Unknown exam mode.');
  var m = CONFIG.MODES[mode];
  var qs = pickExam_(mode);
  var now = new Date();
  var responseId = 'R' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMMddHHmmss') +
                   Math.floor(Math.random() * 9000 + 1000);
  withLock_(function () {
    append_(SHEETS.S, [[responseId, studentId, mode, now, qs.map(function (q) { return q.id; }).join(','), false]]);
  });
  return { responseId: responseId, mode: mode, label: m.label, minutes: m.minutes,
           startedAt: now.getTime(), questions: qs.map(publicQ_), nickname: cleanText_(nickname, 30) };
}

function findSession_(responseId) {
  var t = table_(SHEETS.S), I = t.idx;
  for (var i = t.rows.length - 1; i >= 0; i--) {
    if (String(t.rows[i][I.ResponseID]) === responseId) return { row: i + 2, r: t.rows[i], I: I };
  }
  return null;
}

function submitExam(studentId, responseId, answers, nickname) {
  studentId = who_(studentId);
  responseId = cleanId_(responseId);
  answers = answers || {};
  return withLock_(function () {
    var s = findSession_(responseId);
    if (!s) throw new Error('Exam session not found.');
    if (String(s.r[s.I.StudentID]) !== studentId) throw new Error('This exam belongs to another session.');
    var already = s.r[s.I.Submitted] === true || String(s.r[s.I.Submitted]).toUpperCase() === 'TRUE';
    if (already) return resultFor_(responseId);

    var mode = String(s.r[s.I.Mode]);
    var ids = String(s.r[s.I.QuestionIDs]).split(',');
    var map = byId_();
    var start = s.r[s.I.StartTime] instanceof Date ? s.r[s.I.StartTime] : new Date(s.r[s.I.StartTime]);
    var end = new Date();
    var duration = Math.round((end - start) / 600) / 100;   // minutes, 2 dp
    var correct = 0, answered = 0, ansRows = [];
    ids.forEach(function (id) {
      var q = map[id];
      var sel = String(answers[id] || '').toUpperCase();
      if ('ABCD'.indexOf(sel) < 0 || sel.length !== 1) sel = '';
      var ok = q && sel !== '' && sel === q.key ? 1 : 0;
      if (sel) answered++;
      correct += ok;
      ansRows.push([responseId, id, sel, ok]);
    });
    var total = ids.length;
    // report Table H.2: fewer than half answered, or under 10 seconds per question, is not valid
    var valid = answered >= Math.ceil(total / 2) && duration >= total * 10 / 60;
    append_(SHEETS.R, [[responseId, start, end, duration, 'Real', valid, mode,
                        cleanText_(nickname, 30), studentId, correct, total, Math.round(correct / total * 1000) / 1000]]);
    append_(SHEETS.A, ansRows);
    sheet_(SHEETS.S).getRange(s.row, s.I.Submitted + 1).setValue(true);
    return resultFor_(responseId);
  });
}

/* accuracy lists by difficulty level and by domain x level, in a fixed order */
function levelList_(lvl) {
  return CONFIG.LEVELS.filter(function (l) { return lvl[l] && lvl[l].n; })
    .map(function (l) { return { level: l, correct: lvl[l].c, total: lvl[l].n }; });
}
function addTo_(m, k, ok) { m[k] = m[k] || { c: 0, n: 0 }; m[k].n++; if (ok) m[k].c++; }
function approachList_(ap) {
  return CONFIG.APPROACHES.filter(function (a) { return ap[a] && ap[a].n; })
    .map(function (a) { return { approach: a, correct: ap[a].c, total: ap[a].n }; });
}

function resultFor_(responseId) {
  var R = table_(SHEETS.R), A = table_(SHEETS.A), map = byId_(), names = lectureNames_();
  var resp = null;
  R.rows.forEach(function (r) { if (String(r[R.idx.ResponseID]) === responseId) resp = r; });
  if (!resp) throw new Error('Result not found.');
  var review = [], dom = {}, lec = {}, lvl = {}, pm = {};
  A.rows.forEach(function (r) {
    if (String(r[A.idx.ResponseID]) !== responseId) return;
    var q = map[String(r[A.idx.QuestionID])];
    if (!q) return;
    var sel = String(r[A.idx.SelectedOption] || ''), ok = Number(r[A.idx.IsCorrect]) === 1;
    review.push({ id: q.id, text: q.text, options: q.options, key: q.key, selected: sel, correct: ok,
                  explanation: q.explanation, domain: q.domain, lecture: q.lecture, difficulty: q.difficulty, pmpTask: q.pmpTask, approach: q.approach });
    addTo_(lvl, q.difficulty, ok); addTo_(pm, q.approach, ok);
    dom[q.domain] = dom[q.domain] || { c: 0, n: 0 };
    lec[q.lecture] = lec[q.lecture] || { c: 0, n: 0 };
    dom[q.domain].n++; lec[q.lecture].n++;
    if (ok) { dom[q.domain].c++; lec[q.lecture].c++; }
  });
  var I = R.idx;
  return {
    responseId: responseId, mode: String(resp[I.Mode]),
    score: Number(resp[I.Score]), total: Number(resp[I.Total]), pct: Number(resp[I.ScorePct]),
    passed: Number(resp[I.ScorePct]) >= CONFIG.PASS_MARK, duration: Number(resp[I.DurationMin]),
    valid: resp[I.IsValid] === true || String(resp[I.IsValid]).toUpperCase() === 'TRUE',
    byDomain: Object.keys(dom).map(function (d) { return { domain: Number(d), name: CONFIG.DOMAINS[d], correct: dom[d].c, total: dom[d].n }; }),
    byLevel: levelList_(lvl),
    byApproach: approachList_(pm),
    byLecture: Object.keys(lec).map(function (l) { return { lecture: Number(l), name: names[l], correct: lec[l].c, total: lec[l].n }; })
                 .sort(function (a, b) { return a.correct / a.total - b.correct / b.total; }),
    review: review
  };
}

/* ------------------------------------------------------------------ progress */

/* every answer a student gave (practice + mock exams) as {q, ok} */
function studentAnswers_(studentId) {
  var map = byId_(), list = [];
  function add(qid, ok) { var q = map[qid]; if (q) list.push({ q: q, ok: ok }); }
  var P = table_(SHEETS.P);
  P.rows.forEach(function (r) { if (String(r[P.idx.StudentID]) === studentId) add(String(r[P.idx.QuestionID]), Number(r[P.idx.IsCorrect]) === 1); });
  var R = table_(SHEETS.R), mine = {};
  R.rows.forEach(function (r) { if (String(r[R.idx.StudentID]) === studentId) mine[String(r[R.idx.ResponseID])] = true; });
  var A = table_(SHEETS.A);
  A.rows.forEach(function (r) {
    if (mine[String(r[A.idx.ResponseID])] && String(r[A.idx.SelectedOption])) add(String(r[A.idx.QuestionID]), Number(r[A.idx.IsCorrect]) === 1);
  });
  return list;
}

function lectureStats_(studentId, list) {
  var stats = {};
  (list || studentAnswers_(studentId)).forEach(function (x) {
    var q = x.q;
    stats[q.lecture] = stats[q.lecture] || { lecture: q.lecture, domain: q.domain, c: 0, n: 0 };
    stats[q.lecture].n++; if (x.ok) stats[q.lecture].c++;
  });
  return stats;
}

function weakLectures_(studentId, list) {
  var s = lectureStats_(studentId, list);
  return Object.keys(s).map(function (k) { return s[k]; })
    .filter(function (x) { return x.n >= 3; })
    .sort(function (a, b) { return a.c / a.n - b.c / b.n; })
    .slice(0, 3);
}

function getProgress(studentId) {
  studentId = who_(studentId);
  var names = lectureNames_(), answers = studentAnswers_(studentId), s = lectureStats_(studentId, answers), lvl = {};
  var pm = {};
  answers.forEach(function (x) { addTo_(lvl, x.q.difficulty, x.ok); addTo_(pm, x.q.approach, x.ok); });
  var R = table_(SHEETS.R), I = R.idx, attempts = [];
  R.rows.forEach(function (r) {
    if (String(r[I.StudentID]) !== studentId) return;
    attempts.push({ responseId: String(r[I.ResponseID]), date: iso_(r[I.CompletionTime]), mode: String(r[I.Mode]),
                    score: Number(r[I.Score]), total: Number(r[I.Total]), pct: Number(r[I.ScorePct]), duration: Number(r[I.DurationMin]) });
  });
  var dom = {};
  Object.keys(s).forEach(function (k) {
    var x = s[k];
    dom[x.domain] = dom[x.domain] || { c: 0, n: 0 };
    dom[x.domain].c += x.c; dom[x.domain].n += x.n;
  });
  var answered = 0, correct = 0;
  Object.keys(s).forEach(function (k) { answered += s[k].n; correct += s[k].c; });
  return {
    answered: answered, correct: correct,
    attempts: attempts.reverse(),
    byDomain: Object.keys(dom).map(function (d) { return { domain: Number(d), name: CONFIG.DOMAINS[d], correct: dom[d].c, total: dom[d].n }; }),
    byLevel: levelList_(lvl),
    byApproach: approachList_(pm),
    byLecture: Object.keys(s).map(function (k) { return { lecture: s[k].lecture, name: names[s[k].lecture], correct: s[k].c, total: s[k].n }; })
                 .sort(function (a, b) { return a.lecture - b.lecture; }),
    weak: weakLectures_(studentId, answers).map(function (w) { return { lecture: w.lecture, name: names[w.lecture], correct: w.c, total: w.n }; })
  };
}

function getAttemptResult(studentId, responseId) {
  studentId = who_(studentId); responseId = cleanId_(responseId);
  var s = findSession_(responseId);
  if (!s || String(s.r[s.I.StudentID]) !== studentId) throw new Error('Result not found.');
  return resultFor_(responseId);
}

/* ------------------------------------------------------------------ flags */

function flagQuestion(studentId, questionId, reason, note) {
  studentId = who_(studentId);
  if (!byId_()[questionId]) throw new Error('Question not found.');
  var reasons = ['Unclear wording', 'Two possible answers', 'Wrong answer key', 'Typo', 'Other'];
  reason = reasons.indexOf(reason) >= 0 ? reason : 'Other';
  withLock_(function () { append_(SHEETS.F, [[new Date(), studentId, questionId, reason, cleanText_(note, 300), 'Open']]); });
  return true;
}

/* ------------------------------------------------------------------ student sign-in (university ID + PIN)
 * PINs are stored only as salted SHA-256 hashes. A signed token (HMAC, 30 days) identifies the student
 * on every request, so the browser never sends the PIN again after signing in.
 */
var STUDENT_HEAD_ = ['StudentID', 'Nickname', 'PinHash', 'Salt', 'CreatedAt', 'LastLogin'];

function studentsSheet_() {
  var sh = ss_().getSheetByName(SHEETS.U);
  if (!sh) {
    sh = ss_().insertSheet(SHEETS.U);
    sh.getRange(1, 1, 1, STUDENT_HEAD_.length).setValues([STUDENT_HEAD_]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function hex_(bytes) {
  return bytes.map(function (b) { var h = (b < 0 ? b + 256 : b).toString(16); return h.length === 1 ? '0' + h : h; }).join('');
}

function pinHash_(salt, pin) {
  return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + ':' + String(pin), Utilities.Charset.UTF_8));
}

function tokenSecret_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('TOKEN_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('TOKEN_SECRET', s); }
  return s;
}

function sign_(text) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(text, tokenSecret_())).replace(/=+$/, '');
}

function makeToken_(studentId) {
  var body = studentId + '.' + (Date.now() + 30 * 24 * 3600 * 1000);
  return body + '.' + sign_(body);
}

function who_(token) {
  var parts = String(token || '').split('.');
  if (parts.length !== 3 || !/^[1-9]\d{5,11}$/.test(parts[0])) throw new Error('Please sign in again.');
  var body = parts[0] + '.' + parts[1];
  if (sign_(body) !== parts[2] || Number(parts[1]) < Date.now()) throw new Error('Please sign in again.');
  return parts[0];
}

function studentLogin(studentId, pin, nickname, create) {
  var id = String(studentId == null ? '' : studentId).replace(/\s/g, '');
  if (!/^[1-9]\d{5,11}$/.test(id)) throw new Error('BAD_ID');
  pin = String(pin == null ? '' : pin);
  if (!/^\d{4,8}$/.test(pin)) throw new Error('BAD_PIN');
  var cache = CacheService.getScriptCache(), failKey = 'fail_' + id;
  var fails = Number(cache.get(failKey) || 0);
  if (fails >= 5) throw new Error('LOCKED');
  return withLock_(function () {
    var sh = studentsSheet_();
    var values = sh.getDataRange().getValues(), head = values[0], I = {};
    head.forEach(function (h, i) { I[String(h).trim()] = i; });
    var rowNo = -1, row = null;
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][I.StudentID]).replace(/^'/, '') === id) { rowNo = i + 1; row = values[i]; break; }
    }
    var now = new Date();
    if (row && String(row[I.PinHash])) {
      if (pinHash_(row[I.Salt], pin) !== String(row[I.PinHash])) {
        cache.put(failKey, String(fails + 1), 15 * 60);
        Utilities.sleep(400);
        throw new Error('WRONG_PIN');
      }
      cache.remove(failKey);
      sh.getRange(rowNo, I.LastLogin + 1).setValue(now);
      var nick = String(row[I.Nickname] || '');
      if (nickname && cleanText_(nickname, 30) !== nick) { nick = cleanText_(nickname, 30); sh.getRange(rowNo, I.Nickname + 1).setValue(nick); }
      return { token: makeToken_(id), studentId: id, nickname: nick };
    }
    if (!create) return { needsRegister: true };
    var salt = Utilities.getUuid(), nk = cleanText_(nickname, 30);
    if (row) {   // the instructor cleared this PIN so the student can choose a new one
      sh.getRange(rowNo, 1, 1, STUDENT_HEAD_.length).setValues([["'" + id, nk || String(row[I.Nickname] || ''), pinHash_(salt, pin), salt, row[I.CreatedAt] || now, now]]);
    } else {
      append_(SHEETS.U, [["'" + id, nk, pinHash_(salt, pin), salt, now, now]]);
    }
    return { token: makeToken_(id), studentId: id, nickname: nk, created: true };
  });
}

/* ------------------------------------------------------------------ instructor dashboard */

function adminLogin(password) {
  var real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!real) throw new Error('No admin password is set yet. Add ADMIN_PASSWORD in Project Settings → Script properties.');
  Utilities.sleep(400);
  if (String(password) !== real) throw new Error('Wrong password.');
  var token = Utilities.getUuid();
  CacheService.getScriptCache().put('admin_' + token, '1', 6 * 3600);
  return token;
}

function checkAdmin_(token) {
  if (!token || !CacheService.getScriptCache().get('admin_' + token)) throw new Error('Please log in again.');
}

function median_(a) {
  if (!a.length) return null;
  var s = a.slice().sort(function (x, y) { return x - y; }), m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
}

function adminStats(token, filters) {
  checkAdmin_(token);
  filters = filters || {};
  var mode = filters.mode || 'course';
  var from = filters.from ? new Date(filters.from + 'T00:00:00') : null;
  var to = filters.to ? new Date(filters.to + 'T23:59:59') : null;
  var validOnly = filters.validOnly !== false;
  var map = byId_(), names = lectureNames_();

  var R = table_(SHEETS.R), I = R.idx, keep = {}, scores = [], durations = [], totalRows = 0, students = {};
  R.rows.forEach(function (r) {
    if (String(r[I.Mode]) !== mode) return;
    var end = r[I.CompletionTime] instanceof Date ? r[I.CompletionTime] : new Date(r[I.CompletionTime]);
    if (from && end < from) return;
    if (to && end > to) return;
    totalRows++;
    var valid = r[I.IsValid] === true || String(r[I.IsValid]).toUpperCase() === 'TRUE';
    if (validOnly && !valid) return;
    keep[String(r[I.ResponseID])] = true;
    var sid = String(r[I.StudentID]).replace(/^'/, '');
    var st = students[sid] = students[sid] || { studentId: sid, nickname: '', attempts: 0, sum: 0, best: 0, last: null, lastDate: '' , lastT: 0 };
    st.attempts++; st.sum += Number(r[I.ScorePct]); st.best = Math.max(st.best, Number(r[I.ScorePct]));
    if (end.getTime() >= st.lastT) { st.lastT = end.getTime(); st.last = Number(r[I.ScorePct]); st.lastDate = iso_(end); if (r[I.Nickname]) st.nickname = String(r[I.Nickname]); }
    scores.push(Number(r[I.ScorePct]));
    durations.push(Number(r[I.DurationMin]));
  });

  var A = table_(SHEETS.A), dom = {}, lec = {}, qs = {}, lvl = {}, dl = {}, pm = {};
  A.rows.forEach(function (r) {
    if (!keep[String(r[A.idx.ResponseID])]) return;
    var q = map[String(r[A.idx.QuestionID])]; if (!q) return;
    var sel = String(r[A.idx.SelectedOption] || ''), ok = Number(r[A.idx.IsCorrect]) === 1;
    var x = qs[q.id] = qs[q.id] || { id: q.id, text: q.text, domain: q.domain, lecture: q.lecture, difficulty: q.difficulty, approach: q.approach, pmpTask: q.pmpTask, key: q.key, seen: 0, attempts: 0, correct: 0, picks: { A: 0, B: 0, C: 0, D: 0 } };
    x.seen++;
    if (!sel) return;                       // blanks: wrong in the score, left out of accuracy (Table H.2)
    x.attempts++; x.picks[sel]++; if (ok) x.correct++;
    dom[q.domain] = dom[q.domain] || { c: 0, n: 0 }; dom[q.domain].n++; if (ok) dom[q.domain].c++;
    lec[q.lecture] = lec[q.lecture] || { c: 0, n: 0 }; lec[q.lecture].n++; if (ok) lec[q.lecture].c++;
    addTo_(lvl, q.difficulty, ok); addTo_(pm, q.approach, ok);
    dl[q.domain] = dl[q.domain] || {}; addTo_(dl[q.domain], q.difficulty, ok);
  });

  var lecList = Object.keys(lec).map(function (l) { return { lecture: Number(l), name: names[l], accuracy: lec[l].c / lec[l].n, attempts: lec[l].n }; })
                  .sort(function (a, b) { return a.accuracy - b.accuracy; });
  lecList.forEach(function (x, i) { x.flag = i < 3; });

  var qList = Object.keys(qs).map(function (k) {
    var x = qs[k], wrong = ['A', 'B', 'C', 'D'].filter(function (o) { return o !== x.key; })
      .sort(function (a, b) { return x.picks[b] - x.picks[a]; })[0];
    var p = x.attempts ? x.correct / x.attempts : null;
    var share = x.attempts ? x.picks[wrong] / x.attempts : 0;
    return { id: x.id, text: x.text, domain: x.domain, lecture: x.lecture, difficulty: x.difficulty, approach: x.approach, pmpTask: x.pmpTask, key: x.key, attempts: x.attempts,
             p: p, topDistractor: x.picks[wrong] ? wrong : '', distractorShare: share,
             flagP: p !== null && (p < 0.30 || p > 0.90), flagDistractor: share > 0.40 };
  }).sort(function (a, b) { return (a.p === null ? 2 : a.p) - (b.p === null ? 2 : b.p); });

  var hist = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
  scores.forEach(function (s) { hist[Math.min(9, Math.floor(s * 10))]++; });

  var F = table_(SHEETS.F), flags = {};
  F.rows.forEach(function (r) {
    if (String(r[F.idx.Status]) === 'Closed') return;
    var id = String(r[F.idx.QuestionID]);
    flags[id] = flags[id] || { id: id, count: 0, reasons: {}, notes: [] };
    flags[id].count++;
    var rs = String(r[F.idx.Reason]); flags[id].reasons[rs] = (flags[id].reasons[rs] || 0) + 1;
    if (r[F.idx.Note]) flags[id].notes.push(String(r[F.idx.Note]));
  });

  return {
    mode: mode, modeLabel: CONFIG.MODES[mode] ? CONFIG.MODES[mode].label : mode,
    responses: scores.length, allResponses: totalRows,
    avgScore: scores.length ? scores.reduce(function (a, b) { return a + b; }, 0) / scores.length : null,
    readiness: scores.length ? scores.filter(function (s) { return s >= CONFIG.PASS_MARK; }).length / scores.length : null,
    medianMinutes: median_(durations),
    passMark: CONFIG.PASS_MARK, domainFlag: CONFIG.DOMAIN_FLAG,
    byDomain: [1, 2, 3].map(function (d) { var x = dom[d]; return { domain: d, name: CONFIG.DOMAINS[d], accuracy: x ? x.c / x.n : null, attempts: x ? x.n : 0 }; }),
    weights: CONFIG.WEIGHTS,
    byApproach: CONFIG.APPROACHES.map(function (a) { var x = pm[a]; return { approach: a, accuracy: x ? x.c / x.n : null, attempts: x ? x.n : 0 }; }),
    levels: CONFIG.LEVELS,
    byLevel: CONFIG.LEVELS.map(function (l) { var x = lvl[l]; return { level: l, accuracy: x ? x.c / x.n : null, attempts: x ? x.n : 0 }; }),
    domainLevel: [1, 2, 3].map(function (d) {
      return { domain: d, name: CONFIG.DOMAINS[d], cells: CONFIG.LEVELS.map(function (l) {
        var x = dl[d] && dl[d][l]; return { level: l, accuracy: x ? x.c / x.n : null, attempts: x ? x.n : 0 }; }) };
    }),
    byLecture: lecList, questions: qList, histogram: hist,
    flags: Object.keys(flags).map(function (k) { return flags[k]; }).sort(function (a, b) { return b.count - a.count; }),
    students: Object.keys(students).map(function (k) {
      var x = students[k];
      return { studentId: x.studentId, nickname: x.nickname, attempts: x.attempts, avg: x.sum / x.attempts,
               best: x.best, last: x.last, lastDate: x.lastDate, ready: x.best >= CONFIG.PASS_MARK };
    }).sort(function (a, b) { return b.best - a.best; }),
    registered: Math.max(0, studentsSheet_().getLastRow() - 1),
    sheetUrl: ss_().getUrl()
  };
}

function adminExport(token, which) {
  checkAdmin_(token);
  var name = which === 'answers' ? SHEETS.A : which === 'students' ? SHEETS.U : SHEETS.R;
  var values = (name === SHEETS.U ? studentsSheet_() : sheet_(name)).getDataRange().getValues();
  if (name === SHEETS.U) values = values.map(function (r) { return [r[0], r[1], r[4], r[5]]; });   // never export PIN hashes
  return values.map(function (r) { return r.map(function (v) { return v instanceof Date ? iso_(v) : v; }); });
}

/* ------------------------------------------------------------------ admin: add questions from the Excel template */

var ITEM_TYPES_ = ['Scenario', 'Calculation', 'Knowledge'];

function taskNumber_(code) {             // I.1 -> 1 ... I.8 -> 8, II.1 -> 9 ... II.10 -> 18, III.1 -> 19 ... III.8 -> 26
  var p = String(code).split('.');
  return { I: 0, II: 8, III: 18 }[p[0]] + Number(p[1]);
}

function pick_(list, v, def) {           // case-insensitive match against a list of allowed values
  v = String(v == null ? '' : v).trim();
  if (!v) return def;
  for (var i = 0; i < list.length; i++) if (list[i].toLowerCase() === v.toLowerCase()) return list[i];
  return null;
}

function safeCell_(v) {                  // stop text from being read as a formula
  v = String(v == null ? '' : v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

/**
 * rows: objects with the template columns (QuestionID, Status, PMP_Task, Approach, Item_Type, Difficulty,
 * Question, Option_A..Option_D, Correct_Option, Explanation, Source). Valid rows are appended to the Questions
 * sheet; invalid rows are returned with the reason. dryRun = check only, add nothing.
 * The fixed Exam set is never changed from here.
 */
function adminAddQuestions(token, rows, dryRun) {
  checkAdmin_(token);
  if (!Array.isArray(rows) || !rows.length) throw new Error('The file has no question rows.');
  if (rows.length > 300) throw new Error('Upload at most 300 questions at a time.');
  return withLock_(function () {
    var sh = sheet_(SHEETS.Q), T = table_(SHEETS.Q), head = T.head.map(function (h) { return String(h).trim(); });
    var existing = {}, maxA = 0;
    T.rows.forEach(function (r) {
      var id = String(r[T.idx.QuestionID] || '').trim();
      if (!id) return;
      existing[id.toUpperCase()] = true;
      var m = /^A(\d+)$/i.exec(id); if (m) maxA = Math.max(maxA, Number(m[1]));
    });
    var out = [], added = [], errors = [], stamp = String(iso_(new Date())).slice(0, 10);
    rows.forEach(function (x, i) {
      var line = Number(x._row) || (i + 2), why = [];
      function g(k) { return cleanText_(x[k], 1500); }
      var code = g('PMP_Task').split(' ')[0].toUpperCase().replace(/^(I{1,3})\./, '$1.');
      if (!CONFIG.PMP_TASKS[code]) why.push('PMP_Task must be a task code such as I.2, II.6 or III.5');
      var status = pick_(['Bank', 'Reserve'], g('Status'), 'Bank');
      if (!status) why.push(String(g('Status')).toLowerCase() === 'exam'
        ? 'Status Exam is not allowed here (the 60-question exam set keeps its 20/25/15 blueprint); use Bank or Reserve'
        : 'Status must be Bank or Reserve');
      var approach = pick_(CONFIG.APPROACHES, g('Approach'), 'Any'); if (!approach) why.push('Approach must be Predictive, Agile, Hybrid or Any');
      var type = pick_(ITEM_TYPES_, g('Item_Type'), 'Scenario'); if (!type) why.push('Item_Type must be Scenario, Calculation or Knowledge');
      var level = pick_(CONFIG.LEVELS, g('Difficulty'), null); if (!level) why.push('Difficulty must be Easy, Medium or Hard');
      var text = g('Question'); if (text.length < 10) why.push('Question text is missing or too short');
      var opts = ['Option_A', 'Option_B', 'Option_C', 'Option_D'].map(function (k) { return cleanText_(x[k], 400); });
      if (opts.some(function (o) { return !o; })) why.push('All four options (A–D) are needed');
      else {
        var seen = {}; opts.forEach(function (o) { seen[o.toLowerCase()] = true; });
        if (Object.keys(seen).length < 4) why.push('The four options must be different');
      }
      var key = g('Correct_Option').toUpperCase(); if (!/^[ABCD]$/.test(key)) why.push('Correct_Option must be A, B, C or D');
      var id = g('QuestionID').toUpperCase();
      if (id) {
        if (!/^[A-Z0-9_-]{1,20}$/.test(id)) why.push('QuestionID may use only letters, digits, - and _');
        else if (existing[id]) why.push('QuestionID ' + id + ' already exists');
      }
      if (why.length) { errors.push({ row: line, reasons: why }); return; }
      if (!id) { do { maxA++; id = 'A' + ('00' + maxA).slice(-3); } while (existing[id]); }
      existing[id] = true;
      var d = { I: 1, II: 2, III: 3 }[code.split('.')[0]], task = code + ' ' + CONFIG.PMP_TASKS[code];
      var val = {
        QuestionID: id, Active: true, Status: status, Domain: d, Domain_Name: CONFIG.DOMAINS[d], ECO_Task: task,
        Lecture: taskNumber_(code), Lecture_Topic: task, Difficulty: level, Question: text,
        Option_A: opts[0], Option_B: opts[1], Option_C: opts[2], Option_D: opts[3], Correct_Option: key,
        Explanation: g('Explanation'), Source: g('Source') || ('Added by instructor ' + stamp), In_Lectures: '',
        Approach: approach, Item_Type: type, PMP_Domain: CONFIG.DOMAINS[d], PMP_Task: code
      };
      out.push(head.map(function (h) { var v = val.hasOwnProperty(h) ? val[h] : ''; return typeof v === 'string' ? safeCell_(v) : v; }));
      added.push(id);
    });
    if (out.length && !dryRun) { sh.getRange(sh.getLastRow() + 1, 1, out.length, head.length).setValues(out); QCACHE_ = null; }
    var iq = head.indexOf('Question'), it = head.indexOf('PMP_Task'), il = head.indexOf('Difficulty'), ik = head.indexOf('Correct_Option');
    return { dryRun: !!dryRun, added: dryRun ? 0 : added.length, ready: added.length, ids: added, errors: errors,
             preview: out.slice(0, 8).map(function (r, i) { return { id: added[i], task: it >= 0 ? r[it] : '', level: il >= 0 ? r[il] : '', key: ik >= 0 ? r[ik] : '', text: iq >= 0 ? String(r[iq]).slice(0, 140) : '' }; }),
             bankSize: T.rows.length + (dryRun ? 0 : out.length) };
  });
}

