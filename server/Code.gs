/* ============================================================
   음표 테트리스 — 순위표 서버 (Google Apps Script)
   · 로그인 확인은 생각 오락실 통합 계정 서버에 물어봐요.
   · '기록' 시트에 닉네임·주마다 최고 점수 한 줄씩 저장해요.
   · 순위표: 이번 주 / 전체, 닉네임마다 최고 점수 하나
   설치: 새 구글 스프레드시트 → 확장 프로그램 → Apps Script → 이 코드를 붙여 넣고 저장
         배포 → 새 배포 → 유형: 웹 앱 (실행: 나 / 액세스 권한: 모든 사용자) → 배포
         나온 웹 앱 주소(…/exec)를 게임 index.html 의 RANK_API 에 넣어요.
   기록 지우기: 시트에서 그 줄을 지우면 돼요. (순위표는 1분 안에 바뀌어요)
   ============================================================ */
var HUB_API = 'https://script.google.com/macros/s/AKfycbzZlmUD9M4rfo4Y4eEYgNjZH0N8SppUA1Otuo9yMEOqcLfklw8XbS5ZmXj1D3KZY4EG/exec';
var SHEET = '기록';
var HEAD = ['닉네임', '주(월요일)', '최고 점수', '완성 마디', '레벨', '올린 시각'];
var TOP = 30;
var MAX_PER_MEASURE = 300;          // 한 번에 3줄을 맞추면 한 줄에 300점 — 그보다 많을 수 없어요
var MEASURES_PER_LEVEL = 6;

function doGet() { return out({ ok: true, msg: '음표 테트리스 순위표 서버가 켜져 있어요.' }); }

function doPost(e) {
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (b.action === 'board') return out(board(b.nick));
    if (b.action === 'submit') {
      var who = verify(b.nick, b.pin);
      if (!who.ok) return out(who);
      return out(submit(who.nick, b));
    }
    return out({ ok: false, error: '알 수 없는 요청이에요.' });
  } catch (err) {
    return out({ ok: false, error: '서버 오류: ' + err, retry: true });
  }
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- 통합 계정 확인 (6시간 동안 기억) ---------- */
function verify(nick, pin) {
  nick = String(nick || '').trim(); pin = String(pin || '').trim();
  if (!nick || !/^\d{4}$/.test(pin)) return { ok: false, code: 'PIN', error: '로그인 정보가 없어요.' };
  var cache = CacheService.getScriptCache();
  var key = 'v:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, nick + '|' + pin));
  var hit = cache.get(key);
  if (hit) return { ok: true, nick: hit };
  var r = UrlFetchApp.fetch(HUB_API, {
    method: 'post', contentType: 'text/plain;charset=utf-8', muteHttpExceptions: true, followRedirects: true,
    payload: JSON.stringify({ action: 'me', nick: nick, pin: pin })
  });
  var j = {};
  try { j = JSON.parse(r.getContentText()); } catch (e) { return { ok: false, error: '생각 오락실 서버에 연결하지 못했어요.', retry: true }; }
  if (!j.ok) return { ok: false, code: j.code || 'PIN', error: j.error || '생각 오락실에서 다시 로그인해 주세요.' };
  cache.put(key, j.nick || nick, 21600);
  return { ok: true, nick: j.nick || nick };
}

/* ---------- 날짜: 이번 주 월요일 (한국 시간) ---------- */
function weekKey(d) {
  d = d || new Date();
  var tz = 'Asia/Seoul';
  var dow = +Utilities.formatDate(d, tz, 'u');          // 1=월 … 7=일
  var mon = new Date(d.getTime() - (dow - 1) * 86400000);
  return Utilities.formatDate(mon, tz, 'yyyy-MM-dd');
}

/* ---------- 시트 ---------- */
function sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');
  }
  return sh;
}
function rows(sh) {
  var n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, HEAD.length).getDisplayValues();
}

/* ---------- 기록 올리기 ---------- */
function submit(nick, b) {
  var score = Math.floor(+b.score || 0), measures = Math.floor(+b.measures || 0), level = Math.floor(+b.level || 0);
  // 너무 이상한 점수는 받지 않아요
  if (score <= 0 || score % 100 !== 0 || measures <= 0 || measures > 3000 ||
      score > measures * MAX_PER_MEASURE || level !== 1 + Math.floor(measures / MEASURES_PER_LEVEL))
    return { ok: false, error: '기록을 확인할 수 없어요.' };
  var wk = weekKey();
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var best = false;
  try {
    var sh = sheet(), data = rows(sh), row = 0;
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === nick && data[i][1] === wk) { row = i + 2; break; }
    }
    var vals = [[nick, wk, score, measures, level, new Date()]];
    if (!row) { sh.getRange(sh.getLastRow() + 1, 1, 1, HEAD.length).setValues(vals); best = true; }
    else if (score > +data[row - 2][2]) { sh.getRange(row, 1, 1, HEAD.length).setValues(vals); best = true; }
  } finally { lock.releaseLock(); }
  CacheService.getScriptCache().remove('board');
  var bd = board(nick);
  bd.best = best;
  return bd;
}

/* ---------- 순위표 ---------- */
function board(nick) {
  var cache = CacheService.getScriptCache();
  var all, hit = cache.get('board');
  if (hit) all = JSON.parse(hit);
  else {
    var wk = weekKey(), week = {}, ever = {};
    rows(sheet()).forEach(function (r) {
      var o = { n: r[0], s: +r[2] || 0, m: +r[3] || 0 };
      if (!ever[o.n] || o.s > ever[o.n].s) ever[o.n] = o;
      if (r[1] === wk && (!week[o.n] || o.s > week[o.n].s)) week[o.n] = o;
    });
    var sort = function (m) { return Object.keys(m).map(function (k) { return m[k]; }).sort(function (a, b) { return b.s - a.s || b.m - a.m; }); };
    all = { week: sort(week), ever: sort(ever) };
    cache.put('board', JSON.stringify(all), 60);
  }
  var res = { ok: true, week: all.week.slice(0, TOP), ever: all.ever.slice(0, TOP) };
  if (nick) {
    ['week', 'ever'].forEach(function (k) {
      for (var i = 0; i < all[k].length; i++) if (all[k][i].n === nick) { res['me_' + k] = { rank: i + 1, s: all[k][i].s, m: all[k][i].m }; break; }
    });
  }
  return res;
}
