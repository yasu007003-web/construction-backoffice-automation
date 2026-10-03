/**
 * 日程調整（返信メールから日時を読み取り、承認後にカレンダーへ反映）
 *
 * Gmailの「日程連絡」ラベルのメールをAIが読み、日程の候補を抜き出して
 * 「日程調整（確認待ち）」シートに書き出す。
 * 人が確認状況を「承認」にすると、Googleカレンダーに予定を作り、
 * 案件管理シートの工事予定日を更新する。
 *
 * 役割分担
 *   AI  … メールに書かれた日付・時間帯・場所を読み取るだけ
 *   GAS … 相対的な日付の計算、所要時間の決定、カレンダーの重複確認、予定の作成
 *   人  … 候補が複数あるときの選択、最終承認
 *
 * 日時の計算はAIにさせない。
 * AIは「10月3日の午前」までを返し、9:00〜12:00に変換するのはGASが行う。
 */

// ===== 設定 =====
const SCHED_CONFIG = {
  TARGET_LABEL: '日程連絡',        // 読み取り対象のラベル
  DONE_LABEL: '日程処理済み',      // 処理後に付けるラベル

  SHEET_NAME: '日程調整（確認待ち）',
  CASE_SHEET: '案件管理',

  // 案件管理シートの見出し名
  H_CASE_NO: '案件番号',
  H_PLAN_DATE: '工事予定日',
  H_UPDATED: '最終更新日',

  // カレンダー。空欄なら既定のカレンダーを使う
  CALENDAR_ID: '',

  // 所要時間の固定ルール
  AM_START: 9,  AM_END: 12,       // 午前：9:00〜12:00
  PM_START: 13, PM_END: 17,       // 午後：13:00〜17:00
  DEFAULT_HOURS: 2,               // 時刻の指定があるときは、そこから2時間

  MAX_MESSAGES: 20,
  MAX_BODY_LENGTH: 3000,
  MODEL: 'gpt-4.1-mini',

  TRIGGER_VALUE: '承認'
};

const SCHED_HEADERS = [
  '処理日時', '受信日時', '件名', '案件番号', '顧客名',
  '開始日時', '終了日時', '場所', '候補', '判定',
  '確認状況', '反映結果', '元メール', 'メールID'
];


/**
 * メールから日程を読み取る。これを実行する。
 */
function extractSchedules() {
  const label = GmailApp.getUserLabelByName(SCHED_CONFIG.TARGET_LABEL);
  if (!label) {
    throw new Error('「' + SCHED_CONFIG.TARGET_LABEL + '」ラベルが見つかりません。Gmailで作成してください。');
  }
  const doneLabel = getOrCreateLabel_(SCHED_CONFIG.DONE_LABEL);

  const sheet = getOrCreateSchedSheet_();
  const threads = label.getThreads(0, SCHED_CONFIG.MAX_MESSAGES);

  let added = 0, skipped = 0, errors = 0;

  threads.forEach(function(thread) {
    // 処理済みのものは飛ばす
    const labels = thread.getLabels().map(function(l) { return l.getName(); });
    if (labels.indexOf(SCHED_CONFIG.DONE_LABEL) >= 0) { skipped++; return; }

    const msg = thread.getMessages()[thread.getMessageCount() - 1];  // 最新の1通
    const received = msg.getDate();

    try {
      const judged = askAiForSchedule_(
        msg.getSubject(), msg.getPlainBody(), received);

      const rows = buildScheduleRows_(judged, msg, received, thread);
      rows.forEach(function(row) { sheet.appendRow(row); added++; });

      thread.addLabel(doneLabel);
    } catch (err) {
      Logger.log('エラー：' + msg.getSubject() + '／' + err.message);
      errors++;
    }
  });

  setupSchedValidation_(sheet);
  Logger.log('追加：' + added + '行／飛ばした：' + skipped + '件／エラー：' + errors + '件');
}


/**
 * AIに日付・時間帯・場所だけを読み取らせる。時刻の計算はさせない。
 */
function askAiForSchedule_(subject, body, received) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY が設定されていません');

  const base = Utilities.formatDate(received, 'Asia/Tokyo', 'yyyy/MM/dd（E）');
  const text = String(body).slice(0, SCHED_CONFIG.MAX_BODY_LENGTH);

  const prompt =
    '次のメールから、作業の日程に関する情報を読み取ってください。\n\n' +
    '【メールを受け取った日】' + base + '\n' +
    '「来週」「再来週」などの表現は、この日を基準に考えてください。\n\n' +
    '【件名】' + subject + '\n' +
    '【本文】\n' + text + '\n\n' +
    '【ルール】\n' +
    '1. 日付の候補をすべて配列で返す。「3日か4日」のように複数あれば、両方返す。\n' +
    '2. 各候補は date（yyyy-MM-dd）と period を持つ。\n' +
    '   period は "午前" "午後" "時刻" "終日" のどれか。\n' +
    '   時刻の指定がある場合は period を "時刻" にし、time に "14:00" の形で書く。\n' +
    '3. 日付が特定できない場合（例：「再来週以降でしたらいつでも」）は、\n' +
    '   candidates を空の配列にして、reason に理由を書く。推測で日付を作らない。\n' +
    '4. 開始時刻や終了時刻を自分で計算しない。書かれていることだけを返す。\n' +
    '5. 案件番号は「2026-018」の形。書かれていなければ null。\n' +
    '6. 場所は、作業を行う住所や建物名。書かれていなければ null。\n' +
    '   差出人の署名欄にある住所を、作業場所と取り違えない。\n\n' +
    '【出力形式】JSONのみ。前置きやコードブロックの記号は付けない。\n' +
    '{"caseNo":"案件番号 または null",' +
    '"customer":"顧客名 または null",' +
    '"place":"場所 または null",' +
    '"candidates":[{"date":"yyyy-MM-dd","period":"午前|午後|時刻|終日","time":"HH:mm または null"}],' +
    '"reason":"候補が空の場合の理由"}';

  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify({
      model: SCHED_CONFIG.MODEL,
      messages: [
        { role: 'system', content: '書かれていることだけを読み取る。日時の計算や推測はしない。' },
        { role: 'user', content: prompt }
      ],
      temperature: 0
    }),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) throw new Error('APIエラー ' + res.getResponseCode());

  let out = JSON.parse(res.getContentText()).choices[0].message.content;
  out = out.replace(/```json|```/g, '').trim();
  Logger.log('AI判定：' + out);
  return JSON.parse(out);
}


/**
 * AIの結果を、シートに書く1行に組み立てる
 */
function buildScheduleRows_(judged, msg, received, thread) {
  const now = new Date();
  const caseNo = cleanValue_(judged.caseNo);
  const customer = cleanValue_(judged.customer);
  const place = cleanValue_(judged.place);
  const candidates = judged.candidates || [];

  const base = [
    Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
    Utilities.formatDate(received, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
    msg.getSubject(), caseNo, customer
  ];

  // 候補がない場合
  if (candidates.length === 0) {
    return [base.concat([
      '', '', place, '',
      '日時を特定できず',
      '未確認',
      cleanValue_(judged.reason) || 'メールから日付を読み取れませんでした',
      '=HYPERLINK("' + thread.getPermalink() + '","メールを開く")',
      msg.getId()
    ])];
  }

  // 候補の一覧を文字にする
  const candText = candidates.map(function(c) {
    return c.date + ' ' + (c.period === '時刻' ? (c.time || '') : c.period);
  }).join(' / ');

  // 先頭の候補で開始・終了を決める（複数ある場合は人が選び直す）
  const range = toTimeRange_(candidates[0]);

  let judge, note;
  if (!range.start) {
    judge = '時間帯が不明';
    note = '日付は分かりましたが、時間帯を決められませんでした';
  } else if (candidates.length > 1) {
    judge = '候補が複数';
    note = '候補から1つ選び、開始日時を書き換えてから承認してください';
  } else {
    const conflicts = findConflicts_(range.start, range.end);
    if (conflicts.length > 0) {
      judge = '予定が重なる';
      note = '既存の予定：' + conflicts.join('、');
    } else {
      judge = '確認待ち';
      note = '';
    }
  }

  return [base.concat([
    range.start ? Utilities.formatDate(range.start, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : '',
    range.end ? Utilities.formatDate(range.end, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : '',
    place, candText, judge, '未確認', note,
    '=HYPERLINK("' + thread.getPermalink() + '","メールを開く")',
    msg.getId()
  ])];
}


/**
 * 候補（日付＋時間帯）から、開始と終了の時刻を決める。
 * ここがAIではなくGASの仕事。
 */
function toTimeRange_(cand) {
  if (!cand || !cand.date) return { start: null, end: null };

  const parts = String(cand.date).split('-');
  if (parts.length !== 3) return { start: null, end: null };

  const y = Number(parts[0]), m = Number(parts[1]) - 1, d = Number(parts[2]);
  const C = SCHED_CONFIG;

  if (cand.period === '午前') {
    return { start: new Date(y, m, d, C.AM_START), end: new Date(y, m, d, C.AM_END) };
  }
  if (cand.period === '午後') {
    return { start: new Date(y, m, d, C.PM_START), end: new Date(y, m, d, C.PM_END) };
  }
  if (cand.period === '時刻' && cand.time) {
    const t = String(cand.time).split(':');
    const hh = Number(t[0]), mm = Number(t[1] || 0);
    if (isNaN(hh)) return { start: null, end: null };
    const s = new Date(y, m, d, hh, mm);
    return { start: s, end: new Date(s.getTime() + C.DEFAULT_HOURS * 60 * 60 * 1000) };
  }
  if (cand.period === '終日') {
    return { start: new Date(y, m, d, C.AM_START), end: new Date(y, m, d, C.PM_END) };
  }
  return { start: null, end: null };
}


/**
 * カレンダーに重なる予定がないか調べる
 */
function findConflicts_(start, end) {
  const cal = SCHED_CONFIG.CALENDAR_ID
            ? CalendarApp.getCalendarById(SCHED_CONFIG.CALENDAR_ID)
            : CalendarApp.getDefaultCalendar();
  if (!cal) return [];

  return cal.getEvents(start, end).map(function(e) {
    return e.getTitle() + '（' +
           Utilities.formatDate(e.getStartTime(), 'Asia/Tokyo', 'HH:mm') + '〜' +
           Utilities.formatDate(e.getEndTime(), 'Asia/Tokyo', 'HH:mm') + '）';
  });
}


/**
 * 確認状況が「承認」になったときに動く。
 * ※他の onEdit がある場合は、そちらから onEditSchedule(e) を呼ぶこと
 */
function onEditSchedule(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() !== SCHED_CONFIG.SHEET_NAME) return;
  if (String(e.value).trim() !== SCHED_CONFIG.TRIGGER_VALUE) return;
  applyScheduleFromRow(e.range.getRow());
}


/**
 * 手動実行用。選んでいる行をカレンダーに反映する。
 */
function applyScheduleForSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== SCHED_CONFIG.SHEET_NAME) {
    SpreadsheetApp.getUi().alert(
      '「' + SCHED_CONFIG.SHEET_NAME + '」シートで行を選んでから実行してください。');
    return;
  }
  applyScheduleFromRow(sheet.getActiveRange().getRow());
}


/**
 * 指定行をカレンダーに反映する
 */
function applyScheduleFromRow(rowNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SCHED_CONFIG.SHEET_NAME);
  if (rowNum < 2) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(function(h) { return String(h).trim(); });
  const idx = {};
  headers.forEach(function(h, i) { idx[h] = i + 1; });

  const get = function(name) {
    return idx[name] ? sheet.getRange(rowNum, idx[name]).getValue() : '';
  };
  const setResult = function(msg, ok) {
    if (!idx['反映結果']) { Logger.log(msg); return; }
    const cell = sheet.getRange(rowNum, idx['反映結果']);
    cell.setValue(msg);
    cell.setBackground(ok ? '#d9ead3' : '#f4cccc');
  };

  const start = toDateValue_(get('開始日時'));
  const end = toDateValue_(get('終了日時'));

  if (!start || !end) {
    setResult('開始日時または終了日時が入っていません', false);
    return;
  }
  if (end <= start) {
    setResult('終了日時が開始日時より前になっています', false);
    return;
  }

  // 承認の時点でもう一度、重なりを確認する
  const conflicts = findConflicts_(start, end);
  if (conflicts.length > 0) {
    setResult('予定が重なるため作成しませんでした：' + conflicts.join('、'), false);
    Logger.log('重複のため中止：' + conflicts.join('、'));
    return;
  }

  const caseNo = String(get('案件番号')).trim();
  const customer = String(get('顧客名')).trim();
  const place = String(get('場所')).trim();

  const title = (caseNo ? caseNo + ' ' : '') + (customer || '作業');

  try {
    const cal = SCHED_CONFIG.CALENDAR_ID
              ? CalendarApp.getCalendarById(SCHED_CONFIG.CALENDAR_ID)
              : CalendarApp.getDefaultCalendar();

    const event = cal.createEvent(title, start, end, {
      location: (place && place !== '不明') ? place : '',
      description: '件名：' + get('件名') + '\n' +
                   'この予定は日程調整シートから自動で作成されました。'
    });

    let msg = 'カレンダーに登録（' +
              Utilities.formatDate(start, 'Asia/Tokyo', 'M/d HH:mm') + '〜' +
              Utilities.formatDate(end, 'Asia/Tokyo', 'HH:mm') + '）';

    // 案件管理シートの工事予定日も更新する
    if (caseNo) {
      const updated = updatePlanDate_(ss, caseNo, start);
      msg += updated ? '／工事予定日を更新' : '／案件管理シートに該当なし';
    }

    setResult(msg, true);
    Logger.log(title + '：' + msg + '｜' + event.getId());
  } catch (err) {
    setResult('カレンダー登録に失敗：' + err.message, false);
  }
}


/**
 * 案件管理シートの工事予定日を更新する
 */
function updatePlanDate_(ss, caseNo, date) {
  const sheet = ss.getSheetByName(SCHED_CONFIG.CASE_SHEET);
  if (!sheet) return false;

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(h) { return String(h).trim(); });
  const cNo = headers.indexOf(SCHED_CONFIG.H_CASE_NO);
  const cPlan = headers.indexOf(SCHED_CONFIG.H_PLAN_DATE);
  const cUpd = headers.indexOf(SCHED_CONFIG.H_UPDATED);
  if (cNo < 0 || cPlan < 0) return false;

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][cNo]).trim() !== caseNo) continue;
    sheet.getRange(i + 1, cPlan + 1)
         .setValue(Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd'));
    if (cUpd >= 0) {
      sheet.getRange(i + 1, cUpd + 1)
           .setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'));
    }
    return true;
  }
  return false;
}


// ===== 補助 =====

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function cleanValue_(v) {
  if (v === null || v === undefined) return '';
  const s = String(v).trim();
  return (s === 'null' || s === 'undefined') ? '' : s;
}

function toDateValue_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const d = new Date(String(v).replace(/-/g, '/'));
  return isNaN(d.getTime()) ? null : d;
}

function getOrCreateSchedSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SCHED_CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SCHED_CONFIG.SHEET_NAME);
    sheet.getRange(1, 1, 1, SCHED_HEADERS.length)
         .setValues([SCHED_HEADERS])
         .setFontWeight('bold')
         .setBackground('#e8eaed');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 確認状況の列にプルダウンを付ける
 */
function setupSchedValidation_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(function(h) { return String(h).trim(); });
  const col = headers.indexOf('確認状況') + 1;
  if (col < 1) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未確認', '承認', '保留', '対象外'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, col, last - 1, 1).setDataValidation(rule);
}


/**
 * テスト用：日程処理済みラベルを外して、もう一度読み込めるようにする
 */
function resetScheduleLabel() {
  const done = GmailApp.getUserLabelByName(SCHED_CONFIG.DONE_LABEL);
  if (!done) { Logger.log('ラベルが見つかりません'); return; }

  const threads = done.getThreads();
  threads.forEach(function(t) { done.removeFromThread(t); });
  Logger.log('ラベルを外しました：' + threads.length + '件');
}
function debugSelection() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  Logger.log('シート名：[' + sheet.getName() + ']');
  Logger.log('選択行：' + sheet.getActiveRange().getRow());
}
