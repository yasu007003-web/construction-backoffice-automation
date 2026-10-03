/**
 * 発注書と見積金額の照合
 *
 * 案件書類フォルダの各案件について、
 *   ・見積書PDF（自分で作ったもの）
 *   ・発注書PDF（お客様から届いたもの）
 * の金額を照合し、結果を「発注照合」シートに出力する。
 *
 * PDFの文字は、Googleドライブの変換機能（OCRを含む）でテキストにしてから、
 * AIに「発注金額はいくらか」だけを判定させる。
 *
 * 役割分担
 *   ドライブ … PDFをテキストにする
 *   AI      … そのテキストから発注金額を1つ取り出す
 *   GAS     … 見積金額と引き算して判定する（計算はAIにさせない）
 *   人      … 差額の原因を判断する
 */

// ===== 設定 =====
const MATCH_CONFIG = {
  ROOT_FOLDER_ID: '1P0WJDVvQMWcJ_UGvVE7-8fe86_ST3qGr',

  CASE_SHEET: '案件管理',
  RESULT_SHEET: '発注照合',

  // ファイル名でどちらの書類かを見分けるキーワード
  QUOTE_WORD: '見積',
  ORDER_WORD: '発注',

  // 案件管理シートの見出し名
  H_CASE_NO: '案件番号',
  H_CUSTOMER: '顧客名',
  H_STATUS: 'ステータス',

  MODEL: 'gpt-4.1-mini',
  MAX_TEXT: 3000,          // AIに渡す本文の最大文字数
  OCR_LANGUAGE: 'ja'
};

const MATCH_HEADERS = [
  '案件番号', '顧客名', '判定', '見積金額', '発注金額', '差額', '備考', 'フォルダ'
];


/**
 * メインの処理。これを実行する。
 */
function checkOrderAmounts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const root = DriveApp.getFolderById(MATCH_CONFIG.ROOT_FOLDER_ID);

  // 案件管理シートから顧客名を引けるようにしておく
  const customerMap = readCustomers_(ss);

  const results = [];
  const folders = root.getFolders();

  while (folders.hasNext()) {
    const folder = folders.next();
    const caseNo = String(folder.getName()).trim();

    // 案件番号の形（2026-018）でないフォルダは飛ばす
    if (!/^\d{4}-\d+$/.test(caseNo)) continue;

    results.push(checkOneCase_(folder, caseNo, customerMap[caseNo] || ''));
  }

  if (results.length === 0) {
    throw new Error('案件フォルダが1つも見つかりませんでした。ROOT_FOLDER_IDを確認してください。');
  }

  results.sort(function(a, b) { return a.caseNo < b.caseNo ? -1 : 1; });
  writeMatchResults_(ss, results);

  Logger.log('照合した案件：' + results.length + '件');
}


/**
 * 1案件分を照合する
 */
function checkOneCase_(folder, caseNo, customer) {
  const base = {
    caseNo: caseNo, customer: customer,
    quote: '', order: '', diff: '', note: '', url: folder.getUrl()
  };

  // フォルダの中から見積書と発注書を探す
  let quoteFile = null, orderFile = null;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName();
    if (name.indexOf(MATCH_CONFIG.ORDER_WORD) >= 0) orderFile = f;
    else if (name.indexOf(MATCH_CONFIG.QUOTE_WORD) >= 0) quoteFile = f;
  }

  if (!quoteFile && !orderFile) {
    base.judge = '対象外';
    base.note = '見積書も発注書もありません';
    Logger.log(caseNo + '：対象外（両方なし）');
    return base;
  }
  if (!orderFile) {
    base.judge = '発注書なし';
    base.note = '発注書が届いていません';
    Logger.log(caseNo + '：発注書なし');
    return base;
  }

  // --- 発注書から金額を読む ---
  let orderAmount = null;
  try {
    const text = extractPdfText_(orderFile);
    if (!text || text.replace(/\s/g, '').length < 10) {
      base.judge = '読み取り失敗';
      base.note = '発注書から文字が取り出せませんでした（画像のみのPDFの可能性）';
      Logger.log(caseNo + '：文字が取り出せず');
      return base;
    }
    orderAmount = askAiForAmount_(text);
  } catch (err) {
    base.judge = '読み取り失敗';
    base.note = err.message;
    Logger.log(caseNo + '：読み取り失敗 ' + err.message);
    return base;
  }

  if (orderAmount === null) {
    base.judge = '読み取り失敗';
    base.note = '発注金額を特定できませんでした';
    Logger.log(caseNo + '：金額を特定できず');
    return base;
  }
  base.order = orderAmount;

  if (!quoteFile) {
    base.judge = '見積書なし';
    base.note = '発注書だけ届いています。見積書を作ったか確認してください';
    Logger.log(caseNo + '：見積書なし（発注 ' + orderAmount + '）');
    return base;
  }

  // --- 見積書から金額を読む ---
  let quoteAmount = null;
  try {
    const qtext = extractPdfText_(quoteFile);
    quoteAmount = askAiForAmount_(qtext);
  } catch (err) {
    base.judge = '読み取り失敗';
    base.note = '見積書の読み取りに失敗：' + err.message;
    return base;
  }

  if (quoteAmount === null) {
    base.judge = '読み取り失敗';
    base.note = '見積金額を特定できませんでした';
    return base;
  }
  base.quote = quoteAmount;

  // --- 引き算はGASがやる ---
  const diff = orderAmount - quoteAmount;
  base.diff = diff;

  if (diff === 0) {
    base.judge = '一致';
  } else {
    base.judge = '金額違い';
    base.note = (diff < 0 ? '発注が見積より' + Math.abs(diff).toLocaleString() + '円 少ない'
                          : '発注が見積より' + diff.toLocaleString() + '円 多い') +
                '／原因は人が確認してください';
  }

  Logger.log(caseNo + '：' + base.judge +
             '｜見積 ' + quoteAmount + '／発注 ' + orderAmount + '／差 ' + diff);
  return base;
}


/**
 * PDFの文字を取り出す。
 * GAS自体にPDFを読む機能はないので、いったんGoogleドキュメントに変換する。
 * この変換にはOCR（画像から文字を読む処理）が含まれる。
 */
function extractPdfText_(file) {
  let docId = null;
  try {
    const resource = {
      title: '_tmp_' + file.getName(),
          mimeType: 'application/pdf'
    };
        Utilities.sleep(10000);   // OCRの回数制限を避けるため3秒待つ
    const inserted = Drive.Files.insert(resource, file.getBlob(), {
      convert: true
    });
    docId = inserted.id;
    const text = DocumentApp.openById(docId).getBody().getText();
    return text;
  } finally {
    // 変換用の一時ファイルは必ず消す
    if (docId) {
      try { DriveApp.getFileById(docId).setTrashed(true); } catch (e) {}
    }
  }
}


/**
 * AIに「金額はいくらか」だけを判定させる。計算はさせない。
 */
function askAiForAmount_(text) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY が設定されていません');

  const body = text.slice(0, MATCH_CONFIG.MAX_TEXT);

  const prompt =
    '次は、見積書または発注書をテキストにしたものです。\n' +
    'この書類の「最終的な金額」を1つだけ取り出してください。\n\n' +
    '【ルール】\n' +
    '1. 合計・発注金額・御見積金額のように、いちばん最後に出てくる総額を選ぶ。\n' +
    '   小計や消費税額、明細1行ごとの金額は選ばない。\n' +
    '2. 数字だけを半角で返す。カンマ・円・¥は付けない。\n' +
    '3. 金額が見つからない、または判断できない場合は null を返す。推測しない。\n' +
    '4. 計算はしない。書かれている数字をそのまま返す。\n\n' +
    '【出力形式】JSONのみ。前置きやコードブロックの記号は付けない。\n' +
    '{"amount": 数字 または null}\n\n' +
    '【書類のテキスト】\n' + body;

  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify({
      model: MATCH_CONFIG.MODEL,
      messages: [
        { role: 'system', content: '書かれている数字をそのまま返す。計算も推測もしない。' },
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

  const amount = JSON.parse(out).amount;
  if (amount === null || amount === undefined || amount === '') return null;

  const n = Number(String(amount).replace(/[,¥円\s]/g, ''));
  return isNaN(n) ? null : n;
}


/**
 * 案件管理シートから案件番号→顧客名の対応を作る
 */
function readCustomers_(ss) {
  const map = {};
  const sheet = ss.getSheetByName(MATCH_CONFIG.CASE_SHEET);
  if (!sheet) return map;

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(h) { return String(h).trim(); });
  const cNo = headers.indexOf(MATCH_CONFIG.H_CASE_NO);
  const cCus = headers.indexOf(MATCH_CONFIG.H_CUSTOMER);
  if (cNo < 0 || cCus < 0) return map;

  for (let i = 1; i < values.length; i++) {
    const no = String(values[i][cNo]).trim();
    if (no) map[no] = String(values[i][cCus]).trim();
  }
  return map;
}


/**
 * 結果をシートに書き出す
 */
function writeMatchResults_(ss, results) {
  let sheet = ss.getSheetByName(MATCH_CONFIG.RESULT_SHEET);
  if (!sheet) sheet = ss.insertSheet(MATCH_CONFIG.RESULT_SHEET);
  else sheet.clear();

  sheet.getRange(1, 1, 1, MATCH_HEADERS.length)
       .setValues([MATCH_HEADERS])
       .setFontWeight('bold')
       .setBackground('#e8eaed');

  const rows = results.map(function(r) {
    return [
      r.caseNo, r.customer, r.judge,
      r.quote === '' ? '' : r.quote,
      r.order === '' ? '' : r.order,
      r.diff === '' ? '' : r.diff,
      r.note,
      r.url ? '=HYPERLINK("' + r.url + '","フォルダを開く")' : ''
    ];
  });
  sheet.getRange(2, 1, rows.length, MATCH_HEADERS.length).setValues(rows);

  const colors = {
    '一致': '#d9ead3',
    '金額違い': '#f4cccc',
    '見積書なし': '#f4cccc',
    '読み取り失敗': '#fce5cd',
    '発注書なし': '#fff2cc',
    '対象外': '#f3f3f3'
  };
  results.forEach(function(r, i) {
    const color = colors[r.judge];
    if (color) sheet.getRange(i + 2, 1, 1, MATCH_HEADERS.length).setBackground(color);
  });

  sheet.getRange(2, 4, rows.length, 3).setNumberFormat('#,##0');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, MATCH_HEADERS.length);
  sheet.getRange(rows.length + 3, 1)
       .setValue('最終照合：' +
                 Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
}


/**
 * 動作確認用。1つのPDFから取り出せた文字をログに出すだけ。
 * 読み取りがうまくいかないとき、そもそも文字が取れているのかを確認する。
 */
function debugPdfText() {
  const root = DriveApp.getFolderById(MATCH_CONFIG.ROOT_FOLDER_ID);
  const folders = root.getFolders();

  while (folders.hasNext()) {
    const folder = folders.next();
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.getName().indexOf(MATCH_CONFIG.ORDER_WORD) < 0) continue;

      Logger.log('■ ' + folder.getName() + ' / ' + f.getName());
      try {
        const t = extractPdfText_(f);
        Logger.log('取り出せた文字数：' + t.length);
        Logger.log(t.slice(0, 500));
      } catch (e) {
        Logger.log('失敗：' + e.message);
      }
    }
  }
}
