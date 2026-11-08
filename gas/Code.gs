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

/** 一度だけ手動で実行。接続先のシートIDをスクリプト プロパティに保存する */
function setup() {
  PropertiesService.getScriptProperties()
    .setProperty('OS_SS_ID', '1IanPrcXeLpW37Ytw7mIJ6YxibubXavJygAxjaGM0VbI');
}

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
    status: { os: false, calendar: false, weather: false }
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
  } catch (e) {
    out.error = String((e && e.message) || e);
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
