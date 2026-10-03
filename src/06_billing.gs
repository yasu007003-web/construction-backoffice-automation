/**
 * 請求対象案件の抽出
 *
 * 工事が完了しているのに、まだ請求していない案件を抜き出す。
 *
 * 第1話の対応待ち一覧はステータスだけを見ていたので、
 * 人がステータスを更新し忘れると検出できなかった。
 * 今回は案件書類フォルダに請求書PDFがあるかどうかで判定するので、
 * ステータスの更新漏れがあっても拾える。
 *
 * 金額は、第5話で作った「発注照合」シートの結果を使い回す。
 * PDFを読み直すと時間がかかるため、一度読んだ結果を再利用する。
 */

// ===== 設定 =====
const BILL_CONFIG = {
  ROOT_FOLDER_ID: '1P0WJDVvQMWcJ_UGvVE7-8fe86_ST3qGr',

  CASE_SHEET: '案件管理',
  MATCH_SHEET: '発注照合',      // 第5話で作ったシート（金額の取得元）
  RESULT_SHEET: '請求対象',

  // 案件管理シートの見出し名
  H_CASE_NO: '案件番号',
  H_CUSTOMER: '顧客名',
  H_WORK: '工事内容',
  H_STATUS: 'ステータス',
  H_DONE_DATE: '完了日',        // ← 案件管理シートに追加する列
  H_STAFF: '担当者',

  // 請求書とみなすファイル名のキーワード
  INVOICE_WORD: '請求',

  // 判定の区切り（工事完了からの日数）
  URGENT_DAYS: 30,              // これ以上たっていたら「至急」
  DUE_DAYS: 7,                  // これ以上たっていたら「要請求」

  // 請求の対象とみなすステータス
  DONE_STATUSES: ['完了', '請求済']
};

const BILL_HEADERS = [
  '案件番号', '顧客名', '工事内容', '担当者',
  '判定', '完了日', '経過日数', '請求予定額', '備考', 'フォルダ'
];


/**
 * メインの処理。これを実行する。
 */
function extractBillingTargets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const caseSheet = ss.getSheetByName(BILL_CONFIG.CASE_SHEET);
  if (!caseSheet) {
    throw new Error('「' + BILL_CONFIG.CASE_SHEET + '」シートが見つかりません。');
  }

  const values = caseSheet.getDataRange().getValues();
  const headers = values[0].map(function(h) { return String(h).trim(); });

  const col = {};
  [BILL_CONFIG.H_CASE_NO, BILL_CONFIG.H_CUSTOMER, BILL_CONFIG.H_WORK,
   BILL_CONFIG.H_STATUS, BILL_CONFIG.H_DONE_DATE, BILL_CONFIG.H_STAFF
  ].forEach(function(name) { col[name] = headers.indexOf(name); });

  if (col[BILL_CONFIG.H_CASE_NO] < 0) {
    throw new Error('「' + BILL_CONFIG.H_CASE_NO + '」の列が見つかりません。');
  }
  if (col[BILL_CONFIG.H_DONE_DATE] < 0) {
    throw new Error('「' + BILL_CONFIG.H_DONE_DATE +
                    '」の列が案件管理シートにありません。列を追加してください。');
  }

  // 請求書があるかどうかを、フォルダ名ごとにまとめて調べる
  const invoiceMap = readInvoiceFolders_();

  // 第5話の照合結果から金額を引く
  const amountMap = readAmounts_(ss);

  const today = startOfDay_(new Date());
  const results = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const caseNo = String(row[col[BILL_CONFIG.H_CASE_NO]]).trim();
    if (!caseNo) continue;

    const status = col[BILL_CONFIG.H_STATUS] >= 0
                 ? String(row[col[BILL_CONFIG.H_STATUS]]).trim() : '';
    const doneRaw = row[col[BILL_CONFIG.H_DONE_DATE]];
    const isDoneStatus = BILL_CONFIG.DONE_STATUSES.indexOf(status) >= 0;
    const doneDate = toDate_(doneRaw);

    // 工事が終わっていない案件は、そもそも請求の対象外なので一覧に出さない
    if (!doneDate && !isDoneStatus) continue;

    const info = invoiceMap[caseNo] || { hasInvoice: false, url: '' };
    const amount = amountMap[caseNo];

    const r = {
      caseNo: caseNo,
      customer: col[BILL_CONFIG.H_CUSTOMER] >= 0
              ? String(row[col[BILL_CONFIG.H_CUSTOMER]]).trim() : '',
      work: col[BILL_CONFIG.H_WORK] >= 0
          ? String(row[col[BILL_CONFIG.H_WORK]]).trim() : '',
      staff: col[BILL_CONFIG.H_STAFF] >= 0
           ? String(row[col[BILL_CONFIG.H_STAFF]]).trim() : '',
      doneDate: doneDate,
      days: '',
      amount: (amount === undefined || amount === null) ? '' : amount,
      note: '',
      url: info.url
    };

    // --- 判定 ---
    if (info.hasInvoice) {
      r.judge = '請求済';
      if (doneDate) r.days = daysBetween_(doneDate, today);
    } else if (!doneDate) {
      r.judge = '完了日が未入力';
      r.note = 'ステータスは「' + status + '」ですが、完了日が空欄です';
    } else {
      const days = daysBetween_(doneDate, today);
      r.days = days;
      if (days >= BILL_CONFIG.URGENT_DAYS) {
        r.judge = '至急';
        r.note = '完了から' + days + '日たっています';
      } else if (days >= BILL_CONFIG.DUE_DAYS) {
        r.judge = '要請求';
      } else {
        r.judge = '様子見';
        r.note = '完了から' + days + '日';
      }
    }

    if (!info.hasInvoice && r.amount === '' && r.judge !== '完了日が未入力') {
      r.note = (r.note ? r.note + '／' : '') + '金額が分かりません（発注照合を先に実行）';
    }

    results.push(r);
    Logger.log(caseNo + '：' + r.judge +
               (r.days !== '' ? '｜完了から' + r.days + '日' : '') +
               (r.amount !== '' ? '｜' + r.amount + '円' : ''));
  }

  if (results.length === 0) {
    throw new Error('請求対象になりうる案件が1件もありませんでした。' +
                    '完了日またはステータスを確認してください。');
  }

  // 至急 → 要請求 → 完了日が未入力 → 様子見 → 請求済 の順に並べる
  const order = { '至急': 1, '要請求': 2, '完了日が未入力': 3, '様子見': 4, '請求済': 5 };
  results.sort(function(a, b) {
    const d = (order[a.judge] || 9) - (order[b.judge] || 9);
    if (d !== 0) return d;
    return (b.days || 0) - (a.days || 0);
  });

  writeBillingResults_(ss, results);
  Logger.log('抽出した案件：' + results.length + '件');
}


/**
 * 案件フォルダを1回だけ走査して、請求書の有無を調べる
 */
function readInvoiceFolders_() {
  const map = {};
  const root = DriveApp.getFolderById(BILL_CONFIG.ROOT_FOLDER_ID);
  const folders = root.getFolders();

  while (folders.hasNext()) {
    const folder = folders.next();
    const name = String(folder.getName()).trim();
    if (!/^\d{4}-\d+$/.test(name)) continue;

    let hasInvoice = false;
    const files = folder.getFiles();
    while (files.hasNext()) {
      if (files.next().getName().indexOf(BILL_CONFIG.INVOICE_WORD) >= 0) {
        hasInvoice = true;
        break;
      }
    }
    map[name] = { hasInvoice: hasInvoice, url: folder.getUrl() };
  }

  Logger.log('確認した案件フォルダ：' + Object.keys(map).length + '件');
  return map;
}


/**
 * 第5話の「発注照合」シートから、案件ごとの金額を読む。
 * 発注金額があればそちらを優先する（正式な契約額のため）。
 */
function readAmounts_(ss) {
  const map = {};
  const sheet = ss.getSheetByName(BILL_CONFIG.MATCH_SHEET);
  if (!sheet) {
    Logger.log('「' + BILL_CONFIG.MATCH_SHEET + '」シートがないため、金額は空欄になります');
    return map;
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(h) { return String(h).trim(); });
  const cNo = headers.indexOf('案件番号');
  const cQuote = headers.indexOf('見積金額');
  const cOrder = headers.indexOf('発注金額');
  if (cNo < 0) return map;

  for (let i = 1; i < values.length; i++) {
    const no = String(values[i][cNo]).trim();
    if (!no) continue;
    const order = cOrder >= 0 ? Number(values[i][cOrder]) : NaN;
    const quote = cQuote >= 0 ? Number(values[i][cQuote]) : NaN;
    if (!isNaN(order) && order > 0) map[no] = order;
    else if (!isNaN(quote) && quote > 0) map[no] = quote;
  }
  return map;
}


/**
 * 結果をシートに書き出す
 */
function writeBillingResults_(ss, results) {
  let sheet = ss.getSheetByName(BILL_CONFIG.RESULT_SHEET);
  if (!sheet) sheet = ss.insertSheet(BILL_CONFIG.RESULT_SHEET);
  else sheet.clear();

  sheet.getRange(1, 1, 1, BILL_HEADERS.length)
       .setValues([BILL_HEADERS])
       .setFontWeight('bold')
       .setBackground('#e8eaed');

  const rows = results.map(function(r) {
    return [
      r.caseNo, r.customer, r.work, r.staff, r.judge,
      r.doneDate ? Utilities.formatDate(r.doneDate, 'Asia/Tokyo', 'yyyy/MM/dd') : '',
      r.days, r.amount, r.note,
      r.url ? '=HYPERLINK("' + r.url + '","フォルダを開く")' : ''
    ];
  });
  sheet.getRange(2, 1, rows.length, BILL_HEADERS.length).setValues(rows);

  const colors = {
    '至急': '#f4cccc',
    '要請求': '#fce5cd',
    '完了日が未入力': '#fff2cc',
    '様子見': '#ffffff',
    '請求済': '#d9ead3'
  };
  results.forEach(function(r, i) {
    const c = colors[r.judge];
    if (c) sheet.getRange(i + 2, 1, 1, BILL_HEADERS.length).setBackground(c);
  });

  sheet.getRange(2, 8, rows.length, 1).setNumberFormat('#,##0');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, BILL_HEADERS.length);

  // 合計を出す（請求済を除く）
  let total = 0;
  results.forEach(function(r) {
    if (r.judge !== '請求済' && typeof r.amount === 'number') total += r.amount;
  });

  sheet.getRange(rows.length + 3, 1)
       .setValue('未請求の合計：' + total.toLocaleString() + ' 円' +
                 '（' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ' 時点）')
       .setFontWeight('bold');
}


/**
 * 値を日付に変換する。日付でなければ null。
 */
function toDate_(value) {
  if (!value) return null;
  if (value instanceof Date) return startOfDay_(value);

  const s = String(value).trim();
  if (!s) return null;
  const d = new Date(s.replace(/-/g, '/'));
  return isNaN(d.getTime()) ? null : startOfDay_(d);
}


/**
 * 時刻を切り捨てて、日付だけにする
 */
function startOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}


/**
 * 2つの日付の差（日数）
 */
function daysBetween_(from, to) {
  return Math.floor((to.getTime() - from.getTime()) / (24 * 60 * 60 * 1000));
}
