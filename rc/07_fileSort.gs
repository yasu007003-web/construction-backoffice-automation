/**
 * 未整理ファイルの自動振り分け
 *
 * 「未整理」フォルダに置かれたPDFの中身を読み、
 * 案件番号と書類の種類を判定して、
 * 「2026-018_見積書.pdf」の形に名前を変えて案件フォルダへ移動する。
 *
 * 第3話では「ファイル名にキーワードがなければ拾えない」で終わった。
 * 今回は中身を読むことで、名前を揃える側を作る。
 *
 * 役割分担
 *   ドライブ … PDFをテキストにする（第5話と同じ方法）
 *   AI      … 案件番号と書類の種類を判定するだけ
 *   GAS     … 名前を変えて移動する、重複を検出する
 *   人      … 判定できなかったものを見る
 *
 * 判定できないものは動かさない。推測で移動すると書類が行方不明になるため。
 */

// ===== 設定 =====
const SORT_CONFIG = {
  // 未整理ファイルを置くフォルダのID
  INBOX_FOLDER_ID: '1vGxp2BQyR5c9vm2fHJLB_b8k0_yFFIOB',

  // 案件書類フォルダのID（第3話から使っているもの）
  ROOT_FOLDER_ID: '1P0WJDVvQMWcJ_UGvVE7-8fe86_ST3qGr',

  RESULT_SHEET: 'ファイル整理',
  CASE_SHEET: '案件管理',

  // 振り分け先の書類の種類（AIにはこの中から選ばせる）
  DOC_TYPES: ['依頼書', '見積書', '発注書', '報告書', '請求書'],

  MODEL: 'gpt-4.1-mini',
  MAX_TEXT: 2000
};

const SORT_HEADERS = [
  '元のファイル名', '判定', '案件番号', '書類の種類', '新しいファイル名', '備考', '場所'
];


/**
 * メインの処理。これを実行する。
 */
function sortUnsortedFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let inbox;
  try {
    inbox = DriveApp.getFolderById(SORT_CONFIG.INBOX_FOLDER_ID);
  } catch (e) {
    throw new Error('未整理フォルダを開けません。INBOX_FOLDER_ID を確認してください。');
  }
  const root = DriveApp.getFolderById(SORT_CONFIG.ROOT_FOLDER_ID);

  // 案件フォルダの一覧を先に作る
  const caseFolders = {};
  const folders = root.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    caseFolders[String(f.getName()).trim()] = f;
  }
  Logger.log('案件フォルダ：' + Object.keys(caseFolders).length + '件');

  // 案件管理シートにある案件番号（実在するかの確認に使う）
  const knownCases = readKnownCases_(ss);

  // 未整理フォルダのファイルを先に配列にする
  // （移動しながら回すと、途中で取りこぼすため）
  const targets = [];
  const files = inbox.getFiles();
  while (files.hasNext()) targets.push(files.next());

  if (targets.length === 0) {
    throw new Error('未整理フォルダにファイルがありません。');
  }
  Logger.log('未整理ファイル：' + targets.length + '件');

  const results = [];
  targets.forEach(function(file) {
    results.push(sortOneFile_(file, caseFolders, knownCases, inbox));
  });

  writeSortResults_(ss, results);

  const moved = results.filter(function(r) { return r.judge === '振り分け済'; }).length;
  Logger.log('振り分けた：' + moved + '件／残した：' + (results.length - moved) + '件');
}


/**
 * 1ファイル分を処理する
 */
function sortOneFile_(file, caseFolders, knownCases, inbox) {
  const original = file.getName();
  const r = {
    original: original, judge: '', caseNo: '', docType: '',
    newName: '', note: '', place: '未整理フォルダ'
  };

  // PDF以外は触らない
  if (file.getMimeType() !== 'application/pdf') {
    r.judge = '対象外';
    r.note = 'PDF以外のファイルです';
    Logger.log(original + '：対象外（PDFではない）');
    return r;
  }

  // --- 中身を読む ---
  let text = '';
  try {
    text = extractPdfTextForSort_(file);
  } catch (err) {
    r.judge = '読み取り失敗';
    r.note = err.message;
    Logger.log(original + '：読み取り失敗 ' + err.message);
    return r;
  }

  if (!text || text.replace(/\s/g, '').length < 10) {
    r.judge = '読み取り失敗';
    r.note = '文字が取り出せませんでした';
    return r;
  }

  // --- AIに判定させる ---
  let judged;
  try {
    judged = askAiForFileType_(text);
  } catch (err) {
    r.judge = 'AI判定に失敗';
    r.note = err.message;
    return r;
  }

  r.caseNo = judged.caseNo || '';
  r.docType = judged.docType || '';

  // --- 判定できなかったものは動かさない ---
  if (!r.caseNo) {
    r.judge = '案件番号が不明';
    r.note = judged.reason || '書類に案件番号の記載がありません';
    Logger.log(original + '：案件番号が不明');
    return r;
  }
  if (!r.docType) {
    r.judge = '書類の種類が不明';
    r.note = judged.reason || '書類の種類を特定できませんでした';
    Logger.log(original + '：書類の種類が不明');
    return r;
  }
  if (knownCases.length > 0 && knownCases.indexOf(r.caseNo) < 0) {
    r.judge = '案件が見つからない';
    r.note = '案件管理シートに「' + r.caseNo + '」がありません';
    Logger.log(original + '：案件管理シートにない案件番号 ' + r.caseNo);
    return r;
  }

  const folder = caseFolders[r.caseNo];
  if (!folder) {
    r.judge = 'フォルダがない';
    r.note = '「' + r.caseNo + '」のフォルダがドライブにありません';
    Logger.log(original + '：フォルダなし ' + r.caseNo);
    return r;
  }

  // --- 新しい名前を決める ---
  r.newName = r.caseNo + '_' + r.docType + '.pdf';

  // --- 同じ名前のファイルが既にある場合は動かさない ---
  const exists = folder.getFilesByName(r.newName);
  if (exists.hasNext()) {
    r.judge = '重複の可能性';
    r.note = '同じ名前のファイルが既にあります（旧版か重複か、人が確認してください）';
    Logger.log(original + '：重複 ' + r.newName);
    return r;
  }

  // --- 名前を変えて移動する ---
  try {
    file.setName(r.newName);
    file.moveTo(folder);
    r.judge = '振り分け済';
    r.place = r.caseNo;
    if (original !== r.newName) {
      r.note = '「' + original + '」から名前を変えました';
    }
    Logger.log(original + ' → ' + r.caseNo + '/' + r.newName);
  } catch (err) {
    file.setName(original);   // 失敗したら名前を戻す
    r.judge = '移動に失敗';
    r.note = err.message;
  }
  return r;
}


/**
 * PDFの文字を取り出す（第5話と同じ方法）
 * ocr は指定しない。指定すると回数制限に引っかかるため。
 */
function extractPdfTextForSort_(file) {
  let docId = null;
  try {
    const inserted = Drive.Files.insert(
      { title: '_tmp_' + file.getName(), mimeType: 'application/pdf' },
      file.getBlob(),
      { convert: true }
    );
    docId = inserted.id;
    return DocumentApp.openById(docId).getBody().getText();
  } finally {
    if (docId) {
      try { DriveApp.getFileById(docId).setTrashed(true); } catch (e) {}
    }
  }
}


/**
 * AIに「案件番号」と「書類の種類」だけを判定させる
 */
function askAiForFileType_(text) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY が設定されていません');

  const body = text.slice(0, SORT_CONFIG.MAX_TEXT);
  const types = SORT_CONFIG.DOC_TYPES.join('、');

  const prompt =
    '次は、業務書類をテキストにしたものです。\n' +
    'この書類の「案件番号」と「書類の種類」を判定してください。\n\n' +
    '【ルール】\n' +
    '1. 案件番号は「2026-018」のような、4桁の数字とハイフンと数字の形。\n' +
    '   書類の中にこの形の番号がなければ null にする。\n' +
    '   日付や電話番号を案件番号として返さない。\n' +
    '2. 書類の種類は、次のどれか1つを選ぶ：' + types + '\n' +
    '   どれにも当てはまらない、または判断できない場合は null にする。\n' +
    '3. 推測で埋めない。書かれていないものは null にする。\n' +
    '4. null にした場合は、reason にその理由を一文で書く。\n\n' +
    '【出力形式】JSONのみ。前置きやコードブロックの記号は付けない。\n' +
    '{"caseNo":"案件番号 または null","docType":"書類の種類 または null","reason":"nullにした理由"}\n\n' +
    '【書類のテキスト】\n' + body;

  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify({
      model: SORT_CONFIG.MODEL,
      messages: [
        { role: 'system', content: '書かれている内容だけで判定する。推測しない。' },
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

  const parsed = JSON.parse(out);
  const caseNo = (parsed.caseNo && parsed.caseNo !== 'null')
               ? String(parsed.caseNo).trim() : '';
  let docType = (parsed.docType && parsed.docType !== 'null')
              ? String(parsed.docType).trim() : '';

  // 一覧にない種類が返ってきたら採用しない
  if (docType && SORT_CONFIG.DOC_TYPES.indexOf(docType) < 0) {
    Logger.log('一覧にない書類の種類が返りました：' + docType);
    docType = '';
  }
  // 案件番号の形を確かめる
  if (caseNo && !/^\d{4}-\d+$/.test(caseNo)) {
    Logger.log('案件番号の形ではありません：' + caseNo);
    return { caseNo: '', docType: docType, reason: '案件番号の形ではない値が返りました' };
  }

  return { caseNo: caseNo, docType: docType, reason: parsed.reason || '' };
}


/**
 * 案件管理シートにある案件番号の一覧を読む
 */
function readKnownCases_(ss) {
  const list = [];
  const sheet = ss.getSheetByName(SORT_CONFIG.CASE_SHEET);
  if (!sheet) return list;

  const values = sheet.getDataRange().getValues();
  const col = values[0].map(function(h) { return String(h).trim(); }).indexOf('案件番号');
  if (col < 0) return list;

  for (let i = 1; i < values.length; i++) {
    const no = String(values[i][col]).trim();
    if (no) list.push(no);
  }
  return list;
}


/**
 * 結果をシートに書き出す
 */
function writeSortResults_(ss, results) {
  let sheet = ss.getSheetByName(SORT_CONFIG.RESULT_SHEET);
  if (!sheet) sheet = ss.insertSheet(SORT_CONFIG.RESULT_SHEET);
  else sheet.clear();

  sheet.getRange(1, 1, 1, SORT_HEADERS.length)
       .setValues([SORT_HEADERS])
       .setFontWeight('bold')
       .setBackground('#e8eaed');

  const rows = results.map(function(r) {
    return [r.original, r.judge, r.caseNo, r.docType, r.newName, r.note, r.place];
  });
  sheet.getRange(2, 1, rows.length, SORT_HEADERS.length).setValues(rows);

  const colors = {
    '振り分け済': '#d9ead3',
    '重複の可能性': '#fce5cd',
    '案件番号が不明': '#fff2cc',
    '書類の種類が不明': '#fff2cc',
    '案件が見つからない': '#f4cccc',
    'フォルダがない': '#f4cccc',
    '読み取り失敗': '#f4cccc',
    'AI判定に失敗': '#f4cccc',
    '移動に失敗': '#f4cccc',
    '対象外': '#f3f3f3'
  };
  results.forEach(function(r, i) {
    const c = colors[r.judge];
    if (c) sheet.getRange(i + 2, 1, 1, SORT_HEADERS.length).setBackground(c);
  });

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, SORT_HEADERS.length);

  const moved = results.filter(function(r) { return r.judge === '振り分け済'; }).length;
  sheet.getRange(rows.length + 3, 1)
       .setValue('振り分け：' + moved + '件／未整理のまま：' + (results.length - moved) + '件' +
                 '（' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '）')
       .setFontWeight('bold');
}


/**
 * 動作確認用。未整理フォルダのファイル名と、取り出せた文字をログに出す。
 */
function debugInbox() {
  const inbox = DriveApp.getFolderById(SORT_CONFIG.INBOX_FOLDER_ID);
  const files = inbox.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    Logger.log('■ ' + f.getName() + '（' + f.getMimeType() + '）');
    try {
      const t = extractPdfTextForSort_(f);
      Logger.log('文字数：' + t.length);
      Logger.log(t.slice(0, 300));
    } catch (e) {
      Logger.log('失敗：' + e.message);
    }
  }
}
