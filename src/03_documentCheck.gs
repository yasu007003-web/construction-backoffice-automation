/**
 * 不足書類チェック
 *
 * 案件管理シートの各案件について、Googleドライブの案件フォルダに
 * 必要な書類が揃っているかを照合し、結果を「不足書類チェック」シートに出力する。
 *
 * 判定に使うのはファイル名だけ。ファイルの中身は読まない。
 */

// ===== 設定 =====
const DOC_CONFIG = {
  // ドライブの「案件書類」フォルダのID（URLの /folders/ より後ろ）
  ROOT_FOLDER_ID: '1P0WJDVvQMWcJ_UGvVE7-8fe86_ST3qGr',

  CASE_SHEET_NAME: '案件管理',          // 読み込み元のシート
  RESULT_SHEET_NAME: '不足書類チェック',  // 書き出し先のシート

  // チェック対象の案件番号。空の配列 [] にすると案件管理シートの全件を対象にする
  TARGET_CASES: ['2026-004', '2026-006', '2026-011', '2026-012', '2026-013', '2026-016'],

  // 案件管理シートの見出し名（シート側と合っていない場合はここを直す）
  HEADER_CASE_NO: '案件番号',
  HEADER_CUSTOMER: '顧客名',
  HEADER_STATUS: 'ステータス'
};

// ステータスごとに必要な書類
const REQUIRED_DOCS = {
  '見積中': ['依頼書'],
  '受注済': ['依頼書', '見積書', '発注書'],
  '施工中': ['依頼書', '見積書', '発注書'],
  '完了':   ['依頼書', '見積書', '発注書', '報告書'],
  '請求済': ['依頼書', '見積書', '発注書', '報告書', '請求書']
};

// 書類の種類を見分けるキーワード（ファイル名にこの文字が入っていれば、その書類とみなす）
const DOC_KEYWORDS = {
  '依頼書': '依頼',
  '見積書': '見積',
  '発注書': '発注',
  '報告書': '報告',
  '請求書': '請求'
};

const RESULT_HEADERS = [
  '案件番号', '顧客名', 'ステータス', '判定',
  '不足書類', '重複の可能性', '分類不能ファイル', 'ファイル数', 'フォルダ'
];


/**
 * メインの処理。この関数を実行する。
 */
function checkMissingDocuments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const caseSheet = ss.getSheetByName(DOC_CONFIG.CASE_SHEET_NAME);

  if (!caseSheet) {
    throw new Error('「' + DOC_CONFIG.CASE_SHEET_NAME + '」シートが見つかりません。シート名を確認してください。');
  }

  // --- 案件管理シートを読む ---
  const values = caseSheet.getDataRange().getValues();
  if (values.length < 2) {
    throw new Error('案件管理シートにデータがありません。');
  }

  const headers = values[0].map(function(h) { return String(h).trim(); });
  const colCaseNo   = headers.indexOf(DOC_CONFIG.HEADER_CASE_NO);
  const colCustomer = headers.indexOf(DOC_CONFIG.HEADER_CUSTOMER);
  const colStatus   = headers.indexOf(DOC_CONFIG.HEADER_STATUS);

  if (colCaseNo < 0 || colStatus < 0) {
    throw new Error('見出しが見つかりません。1行目に「' + DOC_CONFIG.HEADER_CASE_NO +
                    '」「' + DOC_CONFIG.HEADER_STATUS + '」があるか確認してください。');
  }

  // --- ドライブのルートフォルダを開く ---
  let rootFolder;
  try {
    rootFolder = DriveApp.getFolderById(DOC_CONFIG.ROOT_FOLDER_ID);
  } catch (e) {
    throw new Error('案件書類フォルダを開けません。ROOT_FOLDER_ID が正しいか確認してください。詳細：' + e.message);
  }
  Logger.log('対象フォルダ：' + rootFolder.getName());

  // --- 案件フォルダの一覧を先にまとめて取得する（1件ずつ探すと遅いため） ---
  const folderMap = {};
  const folders = rootFolder.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    const name = String(f.getName()).trim();  // 末尾の空白対策
    if (folderMap[name]) {
      Logger.log('※同じ名前のフォルダが複数あります：' + name);
    }
    folderMap[name] = f;
  }
  Logger.log('案件フォルダ数：' + Object.keys(folderMap).length);

  // --- 1行ずつ判定する ---
  const results = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const caseNo = String(row[colCaseNo]).trim();
    if (!caseNo) continue;

    // 対象を絞る設定になっていれば、それ以外は飛ばす
    if (DOC_CONFIG.TARGET_CASES.length > 0 &&
        DOC_CONFIG.TARGET_CASES.indexOf(caseNo) < 0) {
      continue;
    }

    const customer = colCustomer >= 0 ? String(row[colCustomer]).trim() : '';
    const status = String(row[colStatus]).trim();

    results.push(checkOneCase(caseNo, customer, status, folderMap[caseNo]));
  }

  if (results.length === 0) {
    throw new Error('対象の案件が1件も見つかりませんでした。案件番号やTARGET_CASESの設定を確認してください。');
  }

  writeResults(ss, results);
  Logger.log('チェックした案件：' + results.length + '件');
}


/**
 * 1案件分を判定する
 */
function checkOneCase(caseNo, customer, status, folder) {
  const required = REQUIRED_DOCS[status];

  // フォルダがない場合
  if (!folder) {
    Logger.log(caseNo + '：フォルダが見つかりません');
    return {
      caseNo: caseNo, customer: customer, status: status,
      judge: 'フォルダなし',
      missing: required ? required.join('、') : '',
      duplicated: '', unknown: '', fileCount: 0, url: ''
    };
  }

  // ステータスが想定外の場合
  if (!required) {
    Logger.log(caseNo + '：ステータス「' + status + '」は判定の対象外です');
    return {
      caseNo: caseNo, customer: customer, status: status,
      judge: 'ステータス未定義',
      missing: '', duplicated: '', unknown: '',
      fileCount: 0, url: folder.getUrl()
    };
  }

  // --- フォルダ直下のファイル名を集める ---
  const fileNames = [];
  const files = folder.getFiles();
  while (files.hasNext()) {
    fileNames.push(files.next().getName());
  }

  // --- ファイル名を書類の種類に振り分ける ---
  const found = {};        // 種類ごとの件数
  const unknownFiles = []; // どの種類にも当てはまらないファイル

  fileNames.forEach(function(name) {
    let matched = false;
    for (const docType in DOC_KEYWORDS) {
      if (name.indexOf(DOC_KEYWORDS[docType]) >= 0) {
        found[docType] = (found[docType] || 0) + 1;
        matched = true;
        break;  // 1ファイルは1種類とみなす
      }
    }
    if (!matched) unknownFiles.push(name);
  });

  // --- 不足と重複を出す ---
  const missing = required.filter(function(docType) { return !found[docType]; });

  const duplicated = [];
  for (const docType in found) {
    if (found[docType] >= 2) {
      duplicated.push(docType + '（' + found[docType] + '件）');
    }
  }

  // --- 判定 ---
  let judge;
  if (missing.length > 0) {
    judge = '不足あり';
  } else if (duplicated.length > 0 || unknownFiles.length > 0) {
    judge = '要確認';
  } else {
    judge = 'OK';
  }

  Logger.log(caseNo + '：' + judge +
             '｜ファイル' + fileNames.length + '件' +
             '｜不足[' + missing.join(',') + ']' +
             '｜重複[' + duplicated.join(',') + ']' +
             '｜分類不能[' + unknownFiles.join(',') + ']');

  return {
    caseNo: caseNo, customer: customer, status: status,
    judge: judge,
    missing: missing.join('、'),
    duplicated: duplicated.join('、'),
    unknown: unknownFiles.join('、'),
    fileCount: fileNames.length,
    url: folder.getUrl()
  };
}


/**
 * 結果をシートに書き出す
 */
function writeResults(ss, results) {
  let sheet = ss.getSheetByName(DOC_CONFIG.RESULT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DOC_CONFIG.RESULT_SHEET_NAME);
  } else {
    sheet.clear();
  }

  // 見出し
  sheet.getRange(1, 1, 1, RESULT_HEADERS.length)
       .setValues([RESULT_HEADERS])
       .setFontWeight('bold')
       .setBackground('#e8eaed');

  // 本文
  const rows = results.map(function(r) {
    return [
      r.caseNo, r.customer, r.status, r.judge,
      r.missing, r.duplicated, r.unknown, r.fileCount,
      r.url ? '=HYPERLINK("' + r.url + '","フォルダを開く")' : ''
    ];
  });

  sheet.getRange(2, 1, rows.length, RESULT_HEADERS.length).setValues(rows);

  // 判定に応じて色を付ける
  const colors = {
    '不足あり': '#f4cccc',
    'フォルダなし': '#f4cccc',
    '要確認': '#fce5cd',
    'ステータス未定義': '#fff2cc',
    'OK': '#d9ead3'
  };

  results.forEach(function(r, i) {
    const color = colors[r.judge];
    if (color) {
      sheet.getRange(i + 2, 1, 1, RESULT_HEADERS.length).setBackground(color);
    }
  });

  // 見た目を整える
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, RESULT_HEADERS.length);
  sheet.getRange(1, 1, rows.length + 1, RESULT_HEADERS.length)
       .setVerticalAlignment('middle');

  // 実行日時を末尾に記録
  sheet.getRange(rows.length + 3, 1)
       .setValue('最終チェック：' +
                 Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
}


/**
 * 動作確認用。フォルダの中身をログに出すだけの関数。
 * うまく判定されないときに、ファイル名がどう見えているかを確認する。
 */
function debugListFolders() {
  const root = DriveApp.getFolderById(DOC_CONFIG.ROOT_FOLDER_ID);
  Logger.log('ルート：' + root.getName());

  const folders = root.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    Logger.log('[' + f.getName() + ']');
    const files = f.getFiles();
    while (files.hasNext()) {
      Logger.log('    ' + files.next().getName());
    }
  }
}
// ===== 毎朝のメール通知 =====

const MAIL_SETTINGS = {
  // 宛先。空欄にすると自分（スクリプトの所有者）宛てに届く
  TO: '',
  // true にすると、不足も要確認もない日はメールを送らない
  SKIP_WHEN_ALL_OK: true
};

/**
 * 不足書類チェックを実行して、結果をメールで送る。
 * 毎朝のトリガーからはこの関数を呼ぶ。
 */
function sendMissingDocMail() {
  checkMissingDocuments();  // まずチェックを実行してシートを更新する

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName(DOC_CONFIG.RESULT_SHEET_NAME);
  const values = sheet.getDataRange().getValues();

  // 見出しの位置を調べる
  const headers = values[0];
  const idx = {};
  headers.forEach(function(h, i) { idx[String(h).trim()] = i; });

  const problems = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const caseNo = String(row[idx['案件番号']] || '').trim();
    if (!caseNo) continue;

    const judge = String(row[idx['判定']] || '').trim();
    if (judge === 'OK') continue;

    problems.push({
      caseNo: caseNo,
      customer: String(row[idx['顧客名']] || ''),
      status: String(row[idx['ステータス']] || ''),
      judge: judge,
      missing: String(row[idx['不足書類']] || ''),
      duplicated: String(row[idx['重複の可能性']] || ''),
      unknown: String(row[idx['分類不能ファイル']] || '')
    });
  }

  if (problems.length === 0 && MAIL_SETTINGS.SKIP_WHEN_ALL_OK) {
    Logger.log('不足なし。メールは送りません。');
    return;
  }

  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d');
  const subject = '【書類チェック】' + today + ' 要対応 ' + problems.length + '件';

  let body = '【案件書類チェック】' + today + ' 時点\n\n';

  if (problems.length === 0) {
    body += 'すべての案件で書類が揃っています。\n';
  } else {
    // 判定ごとにまとめる
    const groups = {};
    problems.forEach(function(p) {
      if (!groups[p.judge]) groups[p.judge] = [];
      groups[p.judge].push(p);
    });

    ['不足あり', 'フォルダなし', '要確認', 'ステータス未定義'].forEach(function(judge) {
      const list = groups[judge];
      if (!list) return;

      body += '■ ' + judge + '（' + list.length + '件）\n';
      list.forEach(function(p) {
        body += '　' + p.caseNo + ' ' + p.customer + '（' + p.status + '）\n';
        if (p.missing)    body += '　　不足：' + p.missing + '\n';
        if (p.duplicated) body += '　　重複：' + p.duplicated + '\n';
        if (p.unknown)    body += '　　分類不能：' + p.unknown + '\n';
      });
      body += '\n';
    });
  }

  body += '――――――――――――\n';
  body += '詳細はスプレッドシートの「' + DOC_CONFIG.RESULT_SHEET_NAME + '」シートをご覧ください。\n';
  body += SpreadsheetApp.getActiveSpreadsheet().getUrl() + '\n';

  const to = MAIL_SETTINGS.TO || Session.getActiveUser().getEmail();
  GmailApp.sendEmail(to, subject, body);

  Logger.log('メールを送信しました：' + to + '（要対応 ' + problems.length + '件）');
}
