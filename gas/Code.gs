/**
 * JARVIS 操作盤（第1段階：読み取り専用）
 * 土台：スタンドアロンの GAS ウェブアプリ
 * 使う前に、一度だけ setup() を実行してください。
 */
const TZ = 'Asia/Tokyo';
const WD = ['日', '月', '火', '水', '木', '金', '土'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('JARVIS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 接続先は、スクリプト プロパティ（OS_SS_ID＝基幹スプレッドシート、PORTAL_SS_ID＝システム一覧）で管理する。
 * 最初に Setup.gs の setup() を一度だけ実行する（IDはこのファイルには書かない）。
 */
function getSS_() {
  const id = PropertiesService.getScriptProperties().getProperty('OS_SS_ID');
  if (!id) throw new Error('先に setup() を実行してください');
  return SpreadsheetApp.openById(id);
}

function fmt_(date, pattern) {
  return Utilities.formatDate(date, TZ, pattern);
}

function addDays_(date, n) {
  return new Date(date.getTime() + n * 86400000);
}

/** 画面が必要とするデータをまとめて返す */
function getDashboard() {
  const now = new Date();
  const out = {
    today: null, tomorrow: null, week: [], lunch: null, staff: null,
    calendar: [], weather: null, error: '',
    board: null, supporters: null, games: [],
    status: { os: false, calendar: false, weather: false, board: false, games: false }
  };

  try {
    const ss = getSS_();
    out.today = dayInfo_(ss, now);
    out.tomorrow = dayInfo_(ss, addDays_(now, 1));
    for (let i = 0; i < 8 && out.week.length < 5; i++) {
      const d = addDays_(now, i);
      const inf = dayInfo_(ss, d);
      if (inf && inf.event) {
        out.week.push({
          label: fmt_(d, 'M/d') + '(' + WD[+fmt_(d, 'u') % 7] + ')',
          event: inf.event
        });
      }
    }
    out.lunch = lunchInfo_(ss, now);
    out.staff = staffInfo_(ss, now);
    out.status.os = true;
    try { out.board = boardInfo_(ss, now); out.status.board = !!out.board; } catch (e) { out.board = null; }
    try { out.supporters = supporterInfo_(ss, now); } catch (e) { out.supporters = null; }
  } catch (e) {
    out.error = String((e && e.message) || e);
  }

  try {
    out.games = gamesInfo_();
    out.status.games = true;
  } catch (e) {
    out.games = [];
  }

  try {
    out.calendar = calendarInfo_(now);
    out.status.calendar = true;
  } catch (e) {
    out.calendar = [];
  }

  try {
    out.weather = weatherInfo_();
    out.status.weather = !!out.weather;
  } catch (e) {
    out.weather = null;
  }

  return out;
}

/** 「◯月行事予定_職員用」から、その日の行事と学年別下校時刻を読む */
function dayInfo_(ss, date) {
  const m = +fmt_(date, 'M');
  const d = +fmt_(date, 'd');
  const sh = ss.getSheetByName(m + '月行事予定_職員用');
  if (!sh) return null;

  const last = Math.min(sh.getLastRow(), 45);
  const colA = sh.getRange(1, 1, last, 1).getDisplayValues();
  let row = -1;
  for (let r = 4; r < last; r++) {
    if (String(colA[r][0]).trim() === String(d)) { row = r + 1; break; }
  }
  if (row < 0) return null;

  const vals = sh.getRange(row, 1, 1, 10).getDisplayValues()[0];

  // 下校時刻（D〜I列＝1〜6年）。結合セルは左端にしか値がないので補う
  const tr = sh.getRange(row, 4, 1, 6);
  const times = tr.getDisplayValues()[0].map(String);
  tr.getMergedRanges().forEach(function (mr) {
    const c0 = mr.getColumn();
    const n = mr.getNumColumns();
    const val = String(sh.getRange(mr.getRow(), c0).getDisplayValue());
    for (let c = c0; c < c0 + n; c++) {
      const idx = c - 4;
      if (idx >= 0 && idx < 6) times[idx] = val;
    }
  });

  return {
    month: m,
    day: d,
    weekday: String(vals[1]),
    event: String(vals[2]).trim(),
    times: times,
    lunchMark: String(vals[9]).trim()
  };
}

/** 「給食人数」から、その日の人数（合計欄だけ）を読む。氏名の列は使わない */
function lunchInfo_(ss, date) {
  const sh = ss.getSheetByName('給食人数');
  if (!sh) return null;
  const m = +fmt_(date, 'M');
  const d = +fmt_(date, 'd');
  if (String(sh.getRange(1, 1).getDisplayValue()).indexOf(m + '月') !== 0) return null;

  const lastCol = sh.getLastColumn();
  const head = sh.getRange(2, 1, 1, lastCol).getDisplayValues()[0]
    .map(function (s) { return String(s).replace(/\s/g, ''); });
  const col = {
    milk: head.indexOf('牛乳数'),
    meals: head.indexOf('給食数'),
    large: head.indexOf('給食大'),
    small: head.indexOf('給食小')
  };
  if (col.meals < 0 || col.large < 0 || col.small < 0) return null;

  const last = Math.min(sh.getLastRow(), 40);
  const colA = sh.getRange(1, 1, last, 1).getDisplayValues();
  for (let r = 2; r < last; r++) {
    if (String(colA[r][0]).trim() === String(d)) {
      const v = sh.getRange(r + 1, 1, 1, lastCol).getDisplayValues()[0];
      const num = function (i) {
        const x = parseInt(v[i], 10);
        return isNaN(x) ? null : x;
      };
      const meals = num(col.meals);
      if (meals === null) return null;
      return {
        kids: num(col.small),
        adults: num(col.large),
        meals: meals,
        milk: col.milk >= 0 ? num(col.milk) : null
      };
    }
  }
  return null;
}

/** 「今週の予定」の出張・年休欄から、種類だけを返す（氏名は返さない） */
function staffInfo_(ss, date) {
  const sh = ss.getSheetByName('今週の予定');
  if (!sh) return null;
  const label = fmt_(date, 'M/d').replace(/^0/, '') + '(';
  const rows = sh.getRange(1, 1, 12, 8).getDisplayValues();
  for (let r = 0; r < rows.length; r++) {
    if (String(rows[r][0]).indexOf(label) === 0) {
      const text = String(rows[r][7]);
      const kinds = ['出張', '年休', '職免', '割振', '特休', '振替'].filter(function (k) {
        return text.indexOf(k) >= 0;
      });
      return { kinds: kinds };
    }
  }
  return null;
}

/** 自分のGoogleカレンダーの今日の予定 */
function calendarInfo_(now) {
  const evs = CalendarApp.getDefaultCalendar().getEventsForDay(now);
  return evs.slice(0, 6).map(function (e) {
    return {
      time: e.isAllDayEvent() ? '終日' : fmt_(e.getStartTime(), 'HH:mm'),
      title: e.getTitle()
    };
  });
}

/** 天気（Open-Meteo：キー不要の無料サービス。奈良市付近） */
function weatherInfo_() {
  const url = 'https://api.open-meteo.com/v1/forecast?latitude=34.685&longitude=135.805' +
    '&current=temperature_2m,weather_code,wind_speed_10m&wind_speed_unit=ms' +
    '&daily=precipitation_probability_max&timezone=Asia%2FTokyo&forecast_days=1';
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;
  const j = JSON.parse(res.getContentText());
  const c = j.current;
  const code = c.weather_code;
  let text = '曇り';
  if (code === 0) text = '快晴';
  else if (code <= 2) text = '晴れ';
  else if (code === 3) text = '曇り';
  else if (code === 45 || code === 48) text = '霧';
  else if (code >= 51 && code <= 67) text = '雨';
  else if (code >= 71 && code <= 77) text = '雪';
  else if (code >= 80 && code <= 82) text = 'にわか雨';
  else if (code >= 95) text = '雷雨';
  return {
    temp: Math.round(c.temperature_2m),
    text: text,
    pop: j.daily && j.daily.precipitation_probability_max
      ? j.daily.precipitation_probability_max[0] : null,
    wind: Math.round(c.wind_speed_10m)
  };
}


/* ============ メール・チャット（ログインしている本人の分だけ） ============ */

function shorten_(t, n) {
  t = String(t || '').replace(/\s+/g, ' ').trim();
  return t.length > n ? t.slice(0, n) + '…' : t;
}

function fromName_(raw) {
  raw = String(raw || '');
  const m = raw.match(/^\s*"?([^"<]*?)"?\s*<([^>]+)>/);
  if (m) return (m[1].trim() || m[2]);
  return raw.trim();
}

/** Gmail の未読（読み取り専用）と、Google Chat の新着をまとめて返す。片方が失敗しても、もう片方は返す */
function getInbox() {
  const out = {
    mail: { ok: false, unread: 0, items: [], err: '' },
    chat: { ok: false, count: 0, items: [], err: '' }
  };

  try {
    out.mail.unread = Gmail.Users.Labels.get('me', 'INBOX').messagesUnread || 0;
    const list = Gmail.Users.Messages.list('me', { q: 'is:unread in:inbox', maxResults: 8 });
    (list.messages || []).forEach(function (m) {
      const g = Gmail.Users.Messages.get('me', m.id, {
        format: 'metadata', metadataHeaders: ['From', 'Subject']
      });
      const h = {};
      ((g.payload && g.payload.headers) || []).forEach(function (x) { h[x.name] = x.value; });
      out.mail.items.push({
        from: shorten_(fromName_(h.From), 24),
        subject: shorten_(h.Subject || '（件名なし）', 60),
        time: fmt_(new Date(+g.internalDate), 'M/d HH:mm')
      });
    });
    out.mail.ok = true;
  } catch (e) {
    out.mail.err = shorten_((e && e.message) || e, 120);
  }

  try {
    const since = new Date(Date.now() - 3 * 86400000).toISOString();
    const spaces = (Chat.Spaces.list({ pageSize: 30 }).spaces) || [];
    const found = [];
    spaces.forEach(function (sp) {
      let after = since;
      try {
        const id = sp.name.split('/')[1];
        const rs = Chat.Users.Spaces.getSpaceReadState('users/me/spaces/' + id + '/spaceReadState');
        if (rs && rs.lastReadTime) after = rs.lastReadTime;
      } catch (e) { /* 既読位置が取れない部屋は、直近3日分を見る */ }
      let res;
      try {
        res = Chat.Spaces.Messages.list(sp.name, {
          pageSize: 10, orderBy: 'createTime desc', filter: 'createTime > "' + after + '"'
        });
      } catch (e) { return; }
      (res.messages || []).forEach(function (m) {
        found.push({
          t: m.createTime,
          room: shorten_(sp.displayName || '個別チャット', 20),
          sender: shorten_((m.sender && (m.sender.displayName || '')) || '', 16) || '（送信者）',
          text: shorten_(m.text || m.formattedText || '（本文なし）', 60)
        });
      });
    });
    found.sort(function (a, b) { return a.t < b.t ? 1 : -1; });
    out.chat.count = found.length;
    out.chat.items = found.slice(0, 10).map(function (x) {
      return { room: x.room, sender: x.sender, text: x.text, time: fmt_(new Date(x.t), 'M/d HH:mm') };
    });
    out.chat.ok = true;
  } catch (e) {
    out.chat.err = shorten_((e && e.message) || e, 120);
  }

  return out;
}


/* ============ 掲示板（入力用シート）・サポーター・学習ゲーム ============ */

const POST_SEP_ = '\u0001'; // 掲示板が「1つの投稿」を見分ける区切り文字（入力画面と同じ約束）

function stripHtml_(t) {
  return String(t || '')
    .replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}

function splitPosts_(cell) {
  return String(cell || '').split(POST_SEP_).map(function (x) {
    return stripHtml_(x).replace(/https?:\/\/\S+/g, '（リンク）').replace(/\s+/g, ' ').trim();
  }).filter(function (x) { return x; });
}

/** 「入力用」シートで、その日の行を探す（B列：日付。日付値でも「M/d(曜)」の文字でも一致させる） */
function findNoticeRow_(sh, date) {
  const last = sh.getLastRow();
  if (last < 1) return -1;
  const vals = sh.getRange(1, 2, last, 1).getValues();
  const key = fmt_(date, 'yyyy-MM-dd');
  const label = fmt_(date, 'M/d') + '(';
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i][0];
    if (v instanceof Date) {
      if (fmt_(v, 'yyyy-MM-dd') === key) return i + 1;
    } else if (String(v).trim().indexOf(label) === 0) {
      return i + 1;
    }
  }
  return -1;
}

function boardInfo_(ss, date) {
  const sh = ss.getSheetByName('入力用');
  if (!sh) return null;
  const row = findNoticeRow_(sh, date);
  if (row < 0) return { renraku: [], iinkai: [], gakunen: [] };
  const v = sh.getRange(row, 1, 1, 8).getDisplayValues()[0];
  return { renraku: splitPosts_(v[4]), iinkai: splitPosts_(v[6]), gakunen: splitPosts_(v[7]) };
}

/** サポーター月間シフト：その日の出勤人数と名前（名前はズーム時だけ画面に出る） */
function supporterInfo_(ss, date) {
  const sh = ss.getSheetByName('サポーター月間シフト');
  if (!sh) return null;
  const m = +fmt_(date, 'M');
  const d = +fmt_(date, 'd');
  if (String(sh.getRange(1, 1).getDisplayValue()).indexOf(m + '月') < 0) return null;
  const lastCol = sh.getLastColumn();
  const names = sh.getRange(3, 1, 1, lastCol).getDisplayValues()[0]
    .map(function (x) { return String(x).replace(/\s/g, ''); });
  const lastRow = Math.min(sh.getLastRow(), 45);
  const colA = sh.getRange(1, 1, lastRow, 1).getDisplayValues();
  for (let r = 3; r < lastRow; r++) {
    if (String(colA[r][0]).trim() === d + '日') {
      const v = sh.getRange(r + 1, 1, 1, lastCol).getDisplayValues()[0];
      const on = [];
      for (let c = 2; c < lastCol; c++) {
        const x = String(v[c]).trim();
        if (x && !/欠|休|×|✕/.test(x) && names[c]) on.push(names[c]);
      }
      return { count: on.length, names: on };
    }
  }
  return null;
}

/** システム一覧（ポータル）から、カテゴリ「学習ゲーム」の運用中のものを自動で拾う */
function gamesInfo_() {
  const id = PropertiesService.getScriptProperties().getProperty('PORTAL_SS_ID');
  if (!id) return [];
  const sh = SpreadsheetApp.openById(id).getSheetByName('システム一覧');
  if (!sh) return [];
  const vals = sh.getDataRange().getDisplayValues();
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (String(r[2]).trim() !== '学習ゲーム') continue;
    if (!String(r[1]).trim()) continue;
    const state = String(r[5]).trim();
    if (state === '廃止' || state === '休止中') continue;
    const url = [r[6], r[7]].map(function (u) { return String(u).trim(); })
      .filter(function (u) { return /^https:\/\//.test(u); })[0] || '';
    out.push({ name: String(r[1]).trim(), desc: String(r[3]).trim(), state: state, url: url });
  }
  return out;
}

/* ============ 連絡事項の書き込み（確認つき・取り消しできる） ============ */

function noticeDate_(opt) {
  const now = new Date();
  if (opt && opt.md) {
    const p = String(opt.md).split('/');
    const mo = +p[0], dd = +p[1];
    if (!(mo >= 1 && mo <= 12 && dd >= 1 && dd <= 31)) throw new Error('日付を読み取れませんでした。');
    let d = new Date(+fmt_(now, 'yyyy'), mo - 1, dd, 12);
    if (d.getTime() < now.getTime() - 180 * 86400000) d = new Date(+fmt_(now, 'yyyy') + 1, mo - 1, dd, 12);
    return d;
  }
  return addDays_(now, (opt && opt.offset) || 0);
}

function cleanNotice_(t) {
  t = String(t || '').replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!t) throw new Error('連絡する内容が空です。');
  if (t.length > 300) throw new Error('長すぎます（300字まで）。');
  return t.replace(/&/g, '&amp;').replace(/</g, '＜').replace(/>/g, '＞');
}

function previewNotice(opt) {
  cleanNotice_(opt && opt.text);
  const d = noticeDate_(opt);
  const sh = getSS_().getSheetByName('入力用');
  if (!sh) throw new Error('「入力用」シートがありません。');
  const row = findNoticeRow_(sh, d);
  if (row < 0) throw new Error(fmt_(d, 'M月d日') + 'の行が「入力用」シートにありません。');
  return { label: fmt_(d, 'M月d日') + '(' + WD[+fmt_(d, 'u') % 7] + ')', row: row };
}

function commitNotice(opt) {
  const text = cleanNotice_(opt && opt.text);
  const d = noticeDate_(opt);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('ほかの人が編集中です。少し待ってからもう一度お願いします。');
  try {
    const ss = getSS_();
    const sh = ss.getSheetByName('入力用');
    const row = findNoticeRow_(sh, d);
    if (row < 0) throw new Error(fmt_(d, 'M月d日') + 'の行が見つかりません。');
    const cell = sh.getRange(row, 5); // E列：連絡事項（入力画面と同じ）
    const before = cell.getDisplayValue().toString();
    const after = (before.length ? before + '\n' : '') + POST_SEP_ + text;
    cell.setValue(after);
    touchMarker_(ss);
    const label = fmt_(d, 'M月d日') + '(' + WD[+fmt_(d, 'u') % 7] + ')';
    setUndo_({ kind: 'notice', row: row, col: 5, before: before, after: after, label: label, at: Date.now() });
    addLog_('連絡事項を追加：' + label);
    return { label: label };
  } finally {
    lock.releaseLock();
  }
}

function touchMarker_(ss) {
  try {
    let sh = ss.getSheetByName('更新合図');
    if (!sh) { sh = ss.insertSheet('更新合図'); sh.hideSheet(); }
    sh.getRange('A1').setValue(Date.now());
  } catch (e) { /* 合図に失敗しても、書き込みは成功扱い */ }
}

function setUndo_(o) { PropertiesService.getUserProperties().setProperty('jarvis_undo', JSON.stringify(o)); }

function undoLast() {
  const up = PropertiesService.getUserProperties();
  const raw = up.getProperty('jarvis_undo');
  if (!raw) return '取り消せる操作がありません。';
  const u = JSON.parse(raw);
  if (Date.now() - u.at > 6 * 3600000) { up.deleteProperty('jarvis_undo'); return '時間がたっているので、取り消せません。直前の操作だけ取り消せます。'; }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('ほかの人が編集中です。');
  try {
    const ss = getSS_();
    const cell = ss.getSheetByName('入力用').getRange(u.row, u.col);
    if (cell.getDisplayValue().toString() !== u.after) {
      throw new Error('そのあとに内容が書き換えられているので、自動では戻せません。');
    }
    if (u.before) cell.setValue(u.before); else cell.clearContent();
    touchMarker_(ss);
    up.deleteProperty('jarvis_undo');
    addLog_('取り消し：' + u.label);
    return u.label + 'の連絡事項の追加を取り消しました。';
  } finally {
    lock.releaseLock();
  }
}

function addLog_(what) {
  const up = PropertiesService.getUserProperties();
  let a = [];
  try { a = JSON.parse(up.getProperty('jarvis_log') || '[]'); } catch (e) { a = []; }
  a.push({ t: new Date().toISOString(), w: what });
  up.setProperty('jarvis_log', JSON.stringify(a.slice(-20)));
}

/* ============ ToDo（ジャービス専用。ログインした本人のアカウントの中だけに保存） ============ */

const TODO_KEY_ = 'jarvis_todo';
const TODO_KEEP_DAYS_ = 7;

function loadTodos_() {
  let a = [];
  try { a = JSON.parse(PropertiesService.getUserProperties().getProperty(TODO_KEY_) || '[]'); } catch (e) { a = []; }
  const now = Date.now();
  return a.filter(function (t) { return !(t.done && t.at && now - t.at > TODO_KEEP_DAYS_ * 86400000); });
}

function saveTodos_(a) {
  const json = JSON.stringify(a);
  if (json.length > 8500) throw new Error('ToDoが多すぎます。終わったものを消してください。');
  PropertiesService.getUserProperties().setProperty(TODO_KEY_, json);
}

function todoOut_(a) {
  return a.slice().sort(function (x, y) {
    if (x.done !== y.done) return x.done ? 1 : -1;
    const dx = x.due || '9999', dy = y.due || '9999';
    return dx < dy ? -1 : dx > dy ? 1 : 0;
  }).map(function (t) { return { id: t.id, title: t.title, due: t.due || '', done: !!t.done }; });
}

function getTodos() { return todoOut_(loadTodos_()); }

function addTodo(title, due) {
  title = String(title || '').trim().slice(0, 80);
  if (!title) throw new Error('内容が空です。');
  if (due && !/^\d{4}-\d{2}-\d{2}$/.test(due)) due = '';
  const a = loadTodos_();
  a.unshift({ id: Utilities.getUuid().slice(0, 8), title: title, due: due || '', done: false, at: Date.now() });
  saveTodos_(a);
  return todoOut_(a);
}

function toggleTodo(id) {
  const a = loadTodos_();
  a.forEach(function (t) { if (t.id === id) { t.done = !t.done; t.at = Date.now(); } });
  saveTodos_(a);
  return todoOut_(a);
}

function deleteTodo(id) {
  const a = loadTodos_().filter(function (t) { return t.id !== id; });
  saveTodos_(a);
  return todoOut_(a);
}
