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
  DOMAINS: { 1: 'Fundamentals & Core Concepts', 2: 'Predictive, Plan-Based', 3: 'Agile Frameworks' }
};

var SHEETS = { Q: 'Questions', R: 'Responses', A: 'Answers', S: 'Sessions', P: 'Practice', F: 'Flags' };

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
  adminLogin: adminLogin, adminStats: adminStats, adminExport: adminExport
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
    list.push({
      id: id,
      status: String(r[I.Status] || 'Bank'),
      domain: Number(r[I.Domain]) || 0,
      eco: eco,
      lecture: Number(r[I.Lecture]) || 0,
      topic: String(r[I.Lecture_Topic] || ''),
      difficulty: String(r[I.Difficulty] || ''),
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
  return { id: q.id, domain: q.domain, lecture: q.lecture, topic: q.topic, difficulty: q.difficulty,
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
  var modes = {};
  Object.keys(CONFIG.MODES).forEach(function (k) {
    var m = CONFIG.MODES[k];
    modes[k] = { label: m.label, total: m.total, minutes: m.minutes, blueprint: m.blueprint };
  });
  return {
    appName: CONFIG.APP_NAME, footer: CONFIG.FOOTER, logoUrl: CONFIG.LOGO_URL,
    passMark: CONFIG.PASS_MARK, domains: CONFIG.DOMAINS, modes: modes,
    lectures: Object.keys(lectures).map(function (k) { return lectures[k]; }).sort(function (a, b) { return a.lecture - b.lecture; }),
    bankSize: qs.length
  };
}

/* ------------------------------------------------------------------ practice */

function getPracticeSet(studentId, filters) {
  filters = filters || {};
  var count = Math.min(Math.max(Number(filters.count) || 10, 5), 50);
  var pool = questions_().filter(function (q) {
    if (filters.domain && q.domain !== Number(filters.domain)) return false;
    if (filters.lecture && q.lecture !== Number(filters.lecture)) return false;
    if (filters.difficulty && q.difficulty !== filters.difficulty) return false;
    if (filters.taughtOnly && q.taught !== 'Yes') return false;
    return true;
  });
  if (filters.weak) {
    var weak = weakLectures_(cleanId_(studentId));
    if (weak.length) {
      var set = {}; weak.forEach(function (w) { set[w.lecture] = true; });
      var focused = pool.filter(function (q) { return set[q.lecture]; });
      if (focused.length >= 5) pool = focused;
    }
  }
  return shuffle_(pool.slice()).slice(0, count).map(publicQ_);
}

function checkPractice(studentId, questionId, selected) {
  studentId = cleanId_(studentId);
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
  studentId = cleanId_(studentId);
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
  studentId = cleanId_(studentId);
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

function resultFor_(responseId) {
  var R = table_(SHEETS.R), A = table_(SHEETS.A), map = byId_(), names = lectureNames_();
  var resp = null;
  R.rows.forEach(function (r) { if (String(r[R.idx.ResponseID]) === responseId) resp = r; });
  if (!resp) throw new Error('Result not found.');
  var review = [], dom = {}, lec = {};
  A.rows.forEach(function (r) {
    if (String(r[A.idx.ResponseID]) !== responseId) return;
    var q = map[String(r[A.idx.QuestionID])];
    if (!q) return;
    var sel = String(r[A.idx.SelectedOption] || ''), ok = Number(r[A.idx.IsCorrect]) === 1;
    review.push({ id: q.id, text: q.text, options: q.options, key: q.key, selected: sel, correct: ok,
                  explanation: q.explanation, domain: q.domain, lecture: q.lecture });
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
    byLecture: Object.keys(lec).map(function (l) { return { lecture: Number(l), name: names[l], correct: lec[l].c, total: lec[l].n }; })
                 .sort(function (a, b) { return a.correct / a.total - b.correct / b.total; }),
    review: review
  };
}

/* ------------------------------------------------------------------ progress */

function lectureStats_(studentId) {
  var map = byId_(), stats = {};
  function add(qid, ok) {
    var q = map[qid]; if (!q) return;
    stats[q.lecture] = stats[q.lecture] || { lecture: q.lecture, domain: q.domain, c: 0, n: 0 };
    stats[q.lecture].n++; if (ok) stats[q.lecture].c++;
  }
  var P = table_(SHEETS.P);
  P.rows.forEach(function (r) { if (String(r[P.idx.StudentID]) === studentId) add(String(r[P.idx.QuestionID]), Number(r[P.idx.IsCorrect]) === 1); });
  var R = table_(SHEETS.R), mine = {};
  R.rows.forEach(function (r) { if (String(r[R.idx.StudentID]) === studentId) mine[String(r[R.idx.ResponseID])] = true; });
  var A = table_(SHEETS.A);
  A.rows.forEach(function (r) {
    if (mine[String(r[A.idx.ResponseID])] && String(r[A.idx.SelectedOption])) add(String(r[A.idx.QuestionID]), Number(r[A.idx.IsCorrect]) === 1);
  });
  return stats;
}

function weakLectures_(studentId) {
  var s = lectureStats_(studentId);
  return Object.keys(s).map(function (k) { return s[k]; })
    .filter(function (x) { return x.n >= 3; })
    .sort(function (a, b) { return a.c / a.n - b.c / b.n; })
    .slice(0, 3);
}

function getProgress(studentId) {
  studentId = cleanId_(studentId);
  var names = lectureNames_(), s = lectureStats_(studentId);
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
    byLecture: Object.keys(s).map(function (k) { return { lecture: s[k].lecture, name: names[s[k].lecture], correct: s[k].c, total: s[k].n }; })
                 .sort(function (a, b) { return a.lecture - b.lecture; }),
    weak: weakLectures_(studentId).map(function (w) { return { lecture: w.lecture, name: names[w.lecture], correct: w.c, total: w.n }; })
  };
}

function getAttemptResult(studentId, responseId) {
  studentId = cleanId_(studentId); responseId = cleanId_(responseId);
  var s = findSession_(responseId);
  if (!s || String(s.r[s.I.StudentID]) !== studentId) throw new Error('Result not found.');
  return resultFor_(responseId);
}

/* ------------------------------------------------------------------ flags */

function flagQuestion(studentId, questionId, reason, note) {
  studentId = cleanId_(studentId);
  if (!byId_()[questionId]) throw new Error('Question not found.');
  var reasons = ['Unclear wording', 'Two possible answers', 'Wrong answer key', 'Typo', 'Other'];
  reason = reasons.indexOf(reason) >= 0 ? reason : 'Other';
  withLock_(function () { append_(SHEETS.F, [[new Date(), studentId, questionId, reason, cleanText_(note, 300), 'Open']]); });
  return true;
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

  var R = table_(SHEETS.R), I = R.idx, keep = {}, scores = [], durations = [], totalRows = 0;
  R.rows.forEach(function (r) {
    if (String(r[I.Mode]) !== mode) return;
    var end = r[I.CompletionTime] instanceof Date ? r[I.CompletionTime] : new Date(r[I.CompletionTime]);
    if (from && end < from) return;
    if (to && end > to) return;
    totalRows++;
    var valid = r[I.IsValid] === true || String(r[I.IsValid]).toUpperCase() === 'TRUE';
    if (validOnly && !valid) return;
    keep[String(r[I.ResponseID])] = true;
    scores.push(Number(r[I.ScorePct]));
    durations.push(Number(r[I.DurationMin]));
  });

  var A = table_(SHEETS.A), dom = {}, lec = {}, qs = {};
  A.rows.forEach(function (r) {
    if (!keep[String(r[A.idx.ResponseID])]) return;
    var q = map[String(r[A.idx.QuestionID])]; if (!q) return;
    var sel = String(r[A.idx.SelectedOption] || ''), ok = Number(r[A.idx.IsCorrect]) === 1;
    var x = qs[q.id] = qs[q.id] || { id: q.id, text: q.text, domain: q.domain, lecture: q.lecture, key: q.key, seen: 0, attempts: 0, correct: 0, picks: { A: 0, B: 0, C: 0, D: 0 } };
    x.seen++;
    if (!sel) return;                       // blanks: wrong in the score, left out of accuracy (Table H.2)
    x.attempts++; x.picks[sel]++; if (ok) x.correct++;
    dom[q.domain] = dom[q.domain] || { c: 0, n: 0 }; dom[q.domain].n++; if (ok) dom[q.domain].c++;
    lec[q.lecture] = lec[q.lecture] || { c: 0, n: 0 }; lec[q.lecture].n++; if (ok) lec[q.lecture].c++;
  });

  var lecList = Object.keys(lec).map(function (l) { return { lecture: Number(l), name: names[l], accuracy: lec[l].c / lec[l].n, attempts: lec[l].n }; })
                  .sort(function (a, b) { return a.accuracy - b.accuracy; });
  lecList.forEach(function (x, i) { x.flag = i < 3; });

  var qList = Object.keys(qs).map(function (k) {
    var x = qs[k], wrong = ['A', 'B', 'C', 'D'].filter(function (o) { return o !== x.key; })
      .sort(function (a, b) { return x.picks[b] - x.picks[a]; })[0];
    var p = x.attempts ? x.correct / x.attempts : null;
    var share = x.attempts ? x.picks[wrong] / x.attempts : 0;
    return { id: x.id, text: x.text, domain: x.domain, lecture: x.lecture, key: x.key, attempts: x.attempts,
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
    byLecture: lecList, questions: qList, histogram: hist,
    flags: Object.keys(flags).map(function (k) { return flags[k]; }).sort(function (a, b) { return b.count - a.count; }),
    sheetUrl: ss_().getUrl()
  };
}

function adminExport(token, which) {
  checkAdmin_(token);
  var name = which === 'answers' ? SHEETS.A : SHEETS.R;
  var values = sheet_(name).getDataRange().getValues();
  return values.map(function (r) { return r.map(function (v) { return v instanceof Date ? iso_(v) : v; }); });
}
