/**
 * PM Prep – PMP-style practice and mock-exam system
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
    course: { label: 'Course mock exam', total: 60, minutes: 75, blueprint: { 1: 30, 2: 14, 3: 16 }, useExamSet: true },
    full:   { label: 'Full-length practice exam', total: 150, minutes: 180, blueprint: { 1: 74, 2: 35, 3: 41 } },
    quick:  { label: 'Quick quiz', total: 20, minutes: 25, blueprint: { 1: 10, 2: 4, 3: 6 } }
  },
  DOMAINS: { 1: 'Fundamentals & Core Concepts', 2: 'Predictive, Plan-Based', 3: 'Agile Frameworks' },
  LEVELS: ['Easy', 'Medium', 'Hard'],  // difficulty levels used in the bank, practice filter, results and dashboard
  // PMP exam domains, PMI Examination Content Outline (July 2026);
  // every question is tagged with the PMP domain and task it trains (PMP_MAP_ below).
  PMP: [{ key: 'People', weight: 0.33 }, { key: 'Process', weight: 0.41 }, { key: 'Business Environment', weight: 0.26 }],
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
  adminLogin: adminLogin, adminStats: adminStats, adminExport: adminExport,
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

var PMP_MAP_ = {   // PMP ECO (July 2026) task per question; the domain is the roman numeral (I People, II Process, III Business Environment)
  'I.2': 'Q29 B062 B063 B064 P067 S012',
  'I.3': 'Q02 Q27 B005 B010 B032 B034 B035 B036 B040 B060 B115 B130 B146 P022 P023 P024 P034 P057 P066 S002 S013 S014 S045 S055 S067',
  'I.4': 'Q12 R01 B022 B052',
  'I.5': 'Q08 Q13 Q54 B075 B123 P008 S004 S019',
  'I.8': 'Q20 Q23 B047 B048 P015 S015 S022',
  'II.1': 'Q01 Q03 Q18 Q32 Q45 Q48 Q58 B003 B011 B012 B016 B018 B029 B058 B059 B065 B067 B073 B088 B113 B127 B137 B140 B141 B142 B143 B147 B148 B158 P007 S001 S024 S029 S044 S047 S048 S054 S064',
  'II.2': 'Q10 Q31 Q35 Q36 Q46 Q49 Q50 Q51 Q52 Q59 B008 B024 B068 B069 B083 B107 B108 B117 B118 B119 B120 B122 B124 B125 B126 B128 B129 B145 B149 B150 B155 B157 P009 P010 P042 P047 P056 P059 P061 P062 P068 S010 S028 S050 S051 S053 S057 S060 S068 S071 S072',
  'II.3': 'Q05 Q47 B001 B049 B050 B099 B105 B116 B135 B136 B139 B151 B152 B153 B156 P003 P029 P030 P058 S020 S042 S046 S069',
  'II.4': 'Q07 Q15 Q17 R02 B019 B025 B037 B070 P011 P035 P044 S009',
  'II.5': 'B004 B031 B057 B112 P018 P019 P032 P039 P040 P049 P055 S003 S008 S017 S030',
  'II.6': 'Q41 Q42 Q43 Q44 R03 B015 B076 B093 B094 B095 B096 B097 B098 B100 B101 B106 B133 P036 P050 P051 P052 P053 P054 S038 S039 S040 S041 S070',
  'II.7': 'B013 B026 B056 B084 B085 B131 B138 B154 P031 P033 P041 P048 S021 S036 S037 S056 S061',
  'II.8': 'Q09 Q33 Q34 Q37 Q38 Q39 Q40 R04 B017 B077 B078 B079 B080 B081 B082 B086 B087 B089 B090 P043 P045 P046 P069 S031 S032 S033 S034 S035',
  'II.9': 'Q25 Q53 Q56 Q57 Q60 R06 B033 B071 B092 B103 B114 B121 B132 P063 P065 P070 S043 S052 S058 S059 S062 S063 S066',
  'II.10': 'B023 B144 P004 P005 P016',
  'III.1': 'Q16 Q19 B002 B030 B038 B039 B111 P020 S011 S023',
  'III.2': 'B006 B007 B027 B072 B110 P001 P028 P038 S026',
  'III.3': 'Q04 Q22 Q24 B051 B053 P017 P025 P026 P027 S027',
  'III.4': 'Q26 Q30 R05 B066 P021',
  'III.5': 'Q06 Q11 Q21 B014 B020 B021 B028 B041 B042 B043 B044 B045 B046 B054 B074 B091 B102 B104 P002 P012 P013 P014 P037 P060 S005 S006 S007 S016 S018 S073',
  'III.6': 'Q28 Q55 B055 B061 B109 B134 P006 P064 S025 S065',
  'III.7': 'Q14 S049',
  'III.8': 'B009'
};

var PMP_BY_ID_ = null;
function pmpTaskOf_(id) {
  if (!PMP_BY_ID_) {
    PMP_BY_ID_ = {};
    Object.keys(PMP_MAP_).forEach(function (t) { PMP_MAP_[t].split(' ').forEach(function (q) { PMP_BY_ID_[q] = t; }); });
  }
  return PMP_BY_ID_[id] || 'II.1';
}
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
    // a PMP_Task column in the sheet (e.g. 'III.5') overrides the built-in map
    var pmpTask = (I.PMP_Task !== undefined && String(r[I.PMP_Task] || '').trim()) || pmpTaskOf_(id);
    list.push({
      id: id,
      status: String(r[I.Status] || 'Bank'),
      domain: Number(r[I.Domain]) || 0,
      eco: eco,
      lecture: Number(r[I.Lecture]) || 0,
      topic: String(r[I.Lecture_Topic] || ''),
      difficulty: String(r[I.Difficulty] || ''),
      pmpTask: pmpTask, pmp: pmpDomainOf_(pmpTask),
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
  return { id: q.id, domain: q.domain, lecture: q.lecture, topic: q.topic, difficulty: q.difficulty, pmp: q.pmp, pmpTask: q.pmpTask,
           text: q.text, options: q.options, taught: q.taught };
}

function lectureNames_() {
  var counts = {};
  questions_().forEach(function (q) {
    var name = q.eco.replace(/^\s*\d+\.\d+\s*/, '');
    counts[q.lecture] = counts[q.lecture] || {};
    counts[q.lecture][name] = (counts[q.lecture][name] || 0) + 1;
  });
  var out = {};
  Object.keys(counts).forEach(function (l) {
    var best = Object.keys(counts[l]).sort(function (a, b) { return counts[l][b] - counts[l][a]; })[0];
    out[l] = best;
  });
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
  var bank = {}, examLv = {}, bankPmp = {}, examPmp = {};
  qs.forEach(function (q) {
    bankPmp[q.pmp] = (bankPmp[q.pmp] || 0) + 1;
    if (q.status === 'Exam') examPmp[q.pmp] = (examPmp[q.pmp] || 0) + 1;
    bank[q.domain] = bank[q.domain] || {}; bank[q.domain][q.difficulty] = (bank[q.domain][q.difficulty] || 0) + 1;
    if (q.status === 'Exam') examLv[q.difficulty] = (examLv[q.difficulty] || 0) + 1;
  });
  var modes = {};
  Object.keys(CONFIG.MODES).forEach(function (k) {
    var m = CONFIG.MODES[k];
    modes[k] = { label: m.label, total: m.total, minutes: m.minutes, blueprint: m.blueprint,
                 levelMix: m.useExamSet ? examLv : null, pmpMix: m.useExamSet ? examPmp : null };
  });
  return {
    appName: CONFIG.APP_NAME, footer: CONFIG.FOOTER, logoUrl: CONFIG.LOGO_URL,
    passMark: CONFIG.PASS_MARK, domains: CONFIG.DOMAINS, modes: modes,
    lectures: Object.keys(lectures).map(function (k) { return lectures[k]; }).sort(function (a, b) { return a.lecture - b.lecture; }),
    bankSize: qs.length,
    levels: CONFIG.LEVELS,
    pmp: CONFIG.PMP.map(function (p) { return { pmp: p.key, weight: p.weight, bank: bankPmp[p.key] || 0 }; }),
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
    if (filters.pmp && q.pmp !== filters.pmp) return false;
    if (filters.taughtOnly && q.taught !== 'Yes') return false;
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
function pmpList_(pm) {
  return CONFIG.PMP.filter(function (p) { return pm[p.key] && pm[p.key].n; })
    .map(function (p) { return { pmp: p.key, weight: p.weight, correct: pm[p.key].c, total: pm[p.key].n }; });
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
                  explanation: q.explanation, domain: q.domain, lecture: q.lecture, difficulty: q.difficulty, pmp: q.pmp, pmpTask: q.pmpTask });
    addTo_(lvl, q.difficulty, ok); addTo_(pm, q.pmp, ok);
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
    byPmp: pmpList_(pm),
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
  answers.forEach(function (x) { addTo_(lvl, x.q.difficulty, x.ok); addTo_(pm, x.q.pmp, x.ok); });
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
    byPmp: pmpList_(pm),
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

  var A = table_(SHEETS.A), dom = {}, lec = {}, qs = {}, lvl = {}, dl = {}, pm = {}, pt = {};
  A.rows.forEach(function (r) {
    if (!keep[String(r[A.idx.ResponseID])]) return;
    var q = map[String(r[A.idx.QuestionID])]; if (!q) return;
    var sel = String(r[A.idx.SelectedOption] || ''), ok = Number(r[A.idx.IsCorrect]) === 1;
    var x = qs[q.id] = qs[q.id] || { id: q.id, text: q.text, domain: q.domain, lecture: q.lecture, difficulty: q.difficulty, pmp: q.pmp, pmpTask: q.pmpTask, key: q.key, seen: 0, attempts: 0, correct: 0, picks: { A: 0, B: 0, C: 0, D: 0 } };
    x.seen++;
    if (!sel) return;                       // blanks: wrong in the score, left out of accuracy (Table H.2)
    x.attempts++; x.picks[sel]++; if (ok) x.correct++;
    dom[q.domain] = dom[q.domain] || { c: 0, n: 0 }; dom[q.domain].n++; if (ok) dom[q.domain].c++;
    lec[q.lecture] = lec[q.lecture] || { c: 0, n: 0 }; lec[q.lecture].n++; if (ok) lec[q.lecture].c++;
    addTo_(lvl, q.difficulty, ok); addTo_(pm, q.pmp, ok); addTo_(pt, q.pmpTask, ok);
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
    return { id: x.id, text: x.text, domain: x.domain, lecture: x.lecture, difficulty: x.difficulty, pmp: x.pmp, pmpTask: x.pmpTask, key: x.key, attempts: x.attempts,
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
    byPmp: CONFIG.PMP.map(function (p) { var x = pm[p.key]; return { pmp: p.key, weight: p.weight, accuracy: x ? x.c / x.n : null, attempts: x ? x.n : 0 }; }),
    pmpTasks: Object.keys(pt).map(function (t) { return { task: t, name: CONFIG.PMP_TASKS[t] || t, pmp: pmpDomainOf_(t), accuracy: pt[t].c / pt[t].n, attempts: pt[t].n }; })
                .sort(function (a, b) { return a.accuracy - b.accuracy; }),
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
