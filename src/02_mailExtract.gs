/**
 * メールからの案件情報抽出（完成版）
 *
 * できること
 *  1. Gmailの「案件依頼」ラベルのメールをAIで読み取り、確認待ちシートに書き出す
 *  2. 元メールを開くリンクを付ける
 *  3. 添付ファイルをGoogleドライブに保存する
 *  4. お客様への返信を「下書き」で作成する（送信はしない）
 *  5. 緊急度「高」の依頼をLINEで通知する（LINE未設定ならメールで通知）
 *  6. 確認状況を「承認」にすると、案件管理シートへ自動で転記する
 */

// ===== 設定 =====
const MAIL_CONFIG = {
  TARGET_LABEL: '案件依頼',             // 読み取り対象のラベル
  DONE_LABEL: '処理済み',               // 処理後に付けるラベル
  SHEET_NAME: '抽出結果（確認待ち）',    // 書き出し先のシート
  CASE_SHEET_NAME: '案件管理',           // 転記先のシート
  MODEL: 'gpt-4.1-mini',                // 使用するモデル
  MAX_MESSAGES: 20,                     // 1回の実行で処理する最大件数
  MAX_BODY_LENGTH: 4000,                // AIに渡す本文の最大文字数

  ENABLE_ATTACHMENT_SAVE: true,         // 添付ファイルを保存する
  ENABLE_REPLY_DRAFT: true,             // 返信の下書きを作る
  ENABLE_URGENT_NOTIFY: true,           // 緊急の依頼を通知する

  ATTACHMENT_FOLDER_NAME: '案件依頼_添付ファイル', // 添付ファイルの保存先フォルダ
  NEW_CASE_STATUS: '見積中',            // 新規登録時のステータス

  // 返信下書きの末尾に付ける署名（自分の内容に書き換えてください）
  SIGNATURE:
    '――――――――――――\n' +
    '〇〇設備工業\n' +
    '担当：〇〇\n' +
    'TEL：000-0000-0000\n' +
    '――――――――――――'
};

const EXTRACT_HEADERS = [
  '処理日時', '受信日時', '件名', '判定', '顧客名', '連絡者', '電話番号',
  '作業場所', '物件名', '工事内容', '希望日', '希望日原文', '緊急度',
  '見積希望', '関連案件番号', '要確認事項', '確認状況', '転記結果',
  '案件番号', '元メール', '添付ファイル', '返信下書き', 'メールID'
];

// 見出し名から列番号を引けるようにする（例：COL['確認状況'] → 17）
const COL = {};
EXTRACT_HEADERS.forEach(function(h, i) { COL[h] = i + 1; });

const EXTRACT_PROMPT = `あなたは設備工事会社の事務担当です。
お客様から届いたメールを読み、工事案件の情報を抜き出し、返信文の案を作ってJSON形式で出力してください。

# 判定
メールを次のいずれかに分類してください。
- 新規依頼：新しい工事・点検・見積もりの依頼
- 既存案件の変更：すでにある案件の日程変更・内容変更・キャンセル
- 案件以外：支払い連絡、挨拶、お礼など、工事の依頼ではないもの

# 抽出のルール
- メールに書かれていないことは推測で埋めないでください。
- 情報がない項目は、空欄にせず必ず「不明」と書いてください。
  ただし「関連案件番号」だけは、該当がなければ空欄（""）にしてください。
- 署名欄の住所は会社の所在地であり、作業場所とは限りません。
  本文で作業場所として示されていない限り、署名の住所を作業場所にしないでください。
- 「>」で始まる引用部分は過去のやりとりです。今回の依頼内容として扱わないでください。
- 1通のメールに複数の現場・工事が含まれる場合は、案件ごとに分けて出力してください。
- 「至急」「今日中」「水漏れ」など急を要する内容は緊急度を「高」、
  日程に余裕があるものは「中」、時期未定や相談段階のものは「低」としてください。
- 足りない情報や、確認したほうがよい点を「要確認事項」に書いてください。

# 見積希望のルール
- 見積もりを希望する記載があれば「あり」、不要と明記されていれば「なし」としてください。
- 見積もりについてメールで触れていない場合は「不明」としてください。推測で「なし」にしないでください。

# 希望日のルール（重要）
- 「希望日」には、YYYY/MM/DD 形式の日付を1つだけ書いてください。
  「以降」「または」「頃」などの言葉を付け加えてはいけません。
- 日付はメールの受信日を基準に計算してください。
- 日付が1つに決まらない場合（候補が複数、期間の指定、「〜以降」など）は「不明」とし、
  候補となる日付や期間は「要確認事項」に書いてください。
  例：要確認事項「希望日の候補：2026/09/30（水）または2026/10/01（木）」
- メールに書かれた表現は、そのまま「希望日原文」に残してください。

# 既存案件の変更のルール
- 日程変更の場合は、変更後の日付を「希望日」に書いてください。
  変更前の日付は「要確認事項」に書いてください。
  例：要確認事項「変更前の日程：2026/09/26」
- 内容変更の場合は、変更後の内容を「工事内容」に書いてください。
- 「関連案件番号」は、メール本文に案件番号が明記されている場合のみ書いてください。推測しないでください。

# 返信文案のルール
- お客様へ送る返信メールの本文の案を「返信文案」に書いてください。
- 冒頭は「顧客名＋様」の宛名から始め、依頼へのお礼と、依頼内容を1〜2行で簡潔に復唱してください。
- お客様に確認が必要な不足情報（住所、希望日時、連絡先など）があれば、丁寧に質問してください。
  社内向けのメモ（変更前の日程など）は書かないでください。
- 金額や日程を確定させる表現は使わず、「担当者より改めてご連絡いたします」としてください。
- 緊急度が「高」の場合は、「内容を確認のうえ、至急ご連絡いたします」という趣旨を入れてください。
- 既存案件の変更の場合は、変更のご依頼を受け付けたことと、確認後に改めて連絡することを書いてください。
- 署名は書かないでください（別途付けます）。
- 案件以外の場合は空文字（""）にしてください。

# 出力形式
次の形式のJSONのみを出力してください。説明文は不要です。
{
  "判定": "新規依頼／既存案件の変更／案件以外",
  "返信文案": "",
  "案件": [
    {
      "顧客名": "",
      "連絡者": "",
      "電話番号": "",
      "作業場所": "",
      "物件名": "",
      "工事内容": "",
      "希望日": "",
      "希望日原文": "",
      "緊急度": "高／中／低",
      "見積希望": "あり／なし／不明",
      "関連案件番号": "",
      "要確認事項": ""
    }
  ]
}
案件以外の場合は「案件」を空の配列にしてください。`;


// =====================================================
// メイン処理：メールを読み取ってシートに書き出す
// =====================================================
function extractFromMails() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('スクリプトプロパティに OPENAI_API_KEY が設定されていません');
  }

  const targetLabel = GmailApp.getUserLabelByName(MAIL_CONFIG.TARGET_LABEL);
  if (!targetLabel) {
    throw new Error('Gmailに「' + MAIL_CONFIG.TARGET_LABEL + '」ラベルが見つかりません');
  }

  let doneLabel = GmailApp.getUserLabelByName(MAIL_CONFIG.DONE_LABEL);
  if (!doneLabel) {
    doneLabel = GmailApp.createLabel(MAIL_CONFIG.DONE_LABEL);
  }

  const sheet = getExtractSheet();
  const threads = targetLabel.getThreads(0, 100);

  let processed = 0;
  let rowsAdded = 0;
  let errors = 0;
  let urgentCount = 0;

  for (let i = 0; i < threads.length; i++) {
    if (processed >= MAIL_CONFIG.MAX_MESSAGES) break;

    const thread = threads[i];

    // 処理済みラベルが付いていればスキップ
    const labelNames = thread.getLabels().map(function(l) { return l.getName(); });
    if (labelNames.indexOf(MAIL_CONFIG.DONE_LABEL) >= 0) continue;

    // スレッド内の最新のメールを対象にする
    const messages = thread.getMessages();
    const message = messages[messages.length - 1];

    try {
      const result = callOpenAIExtract(apiKey, message);
      const hantei = result['判定'] || '不明';
      const cases = Array.isArray(result['案件']) ? result['案件'] : [];

      // 添付ファイルの保存
      let attachCell = 'なし';
      if (MAIL_CONFIG.ENABLE_ATTACHMENT_SAVE) {
        try {
          attachCell = saveAttachments(message);
        } catch (e) {
          attachCell = '保存エラー';
          Logger.log('添付保存エラー：' + message.getSubject() + ' ／ ' + e.message);
        }
      }

      // 返信下書きの作成
      let draftCell = '対象外';
      if (MAIL_CONFIG.ENABLE_REPLY_DRAFT && hantei !== '案件以外') {
        try {
          draftCell = createReplyDraft(message, result['返信文案']);
        } catch (e) {
          draftCell = '作成エラー';
          Logger.log('下書き作成エラー：' + message.getSubject() + ' ／ ' + e.message);
        }
      }

      // シートへ書き出し
      const rows = buildExtractRows(message, thread, hantei, cases, attachCell, draftCell);
      if (rows.length > 0) {
        const startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rows.length, EXTRACT_HEADERS.length).setValues(rows);
        colorExtractRows(sheet, startRow, rows);
        rowsAdded += rows.length;
      }

      // 緊急の依頼を通知
      if (MAIL_CONFIG.ENABLE_URGENT_NOTIFY) {
        const urgentCases = cases.filter(function(c) { return c['緊急度'] === '高'; });
        if (urgentCases.length > 0) {
          try {
            notifyUrgent(message, thread, urgentCases);
            urgentCount++;
          } catch (e) {
            Logger.log('緊急通知エラー：' + message.getSubject() + ' ／ ' + e.message);
          }
        }
      }

      thread.addLabel(doneLabel);
      processed++;

    } catch (e) {
      errors++;
      Logger.log('エラー：' + message.getSubject() + ' ／ ' + e.message);
    }

    Utilities.sleep(500);
  }

  const summary =
    '処理したメール：' + processed + '件 ／ 登録行数：' + rowsAdded + '行 ／ ' +
    '緊急通知：' + urgentCount + '件 ／ エラー：' + errors + '件';
  Logger.log(summary);

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(summary, 'メール抽出', 5);
  } catch (e) {
    // 画面がない実行（トリガー）では何もしない
  }
}


// =====================================================
// OpenAI APIでメールから案件情報を抽出する
// =====================================================
function callOpenAIExtract(apiKey, message) {
  const date = message.getDate();
  const youbi = ['日', '月', '火', '水', '木', '金', '土'][date.getDay()];
  const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd') + '（' + youbi + '）';

  let body = message.getPlainBody() || '';
  if (body.length > MAIL_CONFIG.MAX_BODY_LENGTH) {
    body = body.substring(0, MAIL_CONFIG.MAX_BODY_LENGTH);
  }

  const userContent =
    '受信日：' + dateStr + '\n' +
    '件名：' + message.getSubject() + '\n\n' +
    '本文：\n' + body;

  const payload = {
    model: MAIL_CONFIG.MODEL,
    temperature: 0,
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: EXTRACT_PROMPT },
      { role: 'user', content: userContent }
    ]
  };

  const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code !== 200) {
    throw new Error('OpenAI APIエラー（' + code + '）：' + text.substring(0, 300));
  }

  const content = JSON.parse(text).choices[0].message.content;
  const cleaned = content.replace(/`{3}(?:json)?/g, '').trim();
  return JSON.parse(cleaned);
}


// =====================================================
// AIの結果をシートの行データに変換する
// =====================================================
function buildExtractRows(message, thread, hantei, cases, attachCell, draftCell) {
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  const received = Utilities.formatDate(message.getDate(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  const mailLink = '=HYPERLINK("' + thread.getPermalink() + '","メールを開く")';

  function baseRow() {
    const r = new Array(EXTRACT_HEADERS.length).fill('');
    setCell(r, '処理日時', now);
    setCell(r, '受信日時', received);
    setCell(r, '件名', message.getSubject());
    setCell(r, '判定', hantei);
    setCell(r, '元メール', mailLink);
    setCell(r, '添付ファイル', attachCell);
    setCell(r, '返信下書き', draftCell);
    setCell(r, 'メールID', message.getId());
    return r;
  }

  // 案件以外、または案件が取れなかった場合は1行だけ記録
  if (hantei === '案件以外' || cases.length === 0) {
    const r = baseRow();
    setCell(r, '要確認事項', hantei === '案件以外' ? '' : 'AIが案件情報を抽出できませんでした');
    setCell(r, '確認状況', hantei === '案件以外' ? '確認不要' : '未確認');
    return [r];
  }

  const fields = [
    '顧客名', '連絡者', '電話番号', '作業場所', '物件名', '工事内容',
    '希望日', '希望日原文', '緊急度', '見積希望', '関連案件番号', '要確認事項'
  ];

  return cases.map(function(c) {
    const r = baseRow();
    fields.forEach(function(f) {
      setCell(r, f, extractVal(c[f]));
    });
    setCell(r, '確認状況', '未確認');
    return r;
  });
}

function setCell(row, header, value) {
  row[COL[header] - 1] = value;
}

function extractVal(value) {
  if (value === undefined || value === null) return '';
  return String(value);
}


// =====================================================
// 添付ファイルをGoogleドライブに保存する
// =====================================================
function saveAttachments(message) {
  const attachments = message.getAttachments({ includeInlineImages: false });
  if (attachments.length === 0) return 'なし';

  const root = getOrCreateFolder(MAIL_CONFIG.ATTACHMENT_FOLDER_NAME);
  const dateStr = Utilities.formatDate(message.getDate(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
  const safeSubject = message.getSubject().replace(/[\\\/:*?"<>|]/g, '').substring(0, 40);
  const folder = root.createFolder(dateStr + '_' + safeSubject);

  attachments.forEach(function(a) {
    folder.createFile(a.copyBlob());
  });

  return '=HYPERLINK("' + folder.getUrl() + '","添付' + attachments.length + '件")';
}

function getOrCreateFolder(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}


// =====================================================
// 返信の下書きを作る（送信はしない）
// =====================================================
function createReplyDraft(message, replyText) {
  if (!replyText || String(replyText).trim() === '') return '文案なし';
  const body = String(replyText).trim() + '\n\n' + MAIL_CONFIG.SIGNATURE;
  message.createDraftReply(body);
  return '作成済み';
}


// =====================================================
// 緊急の依頼を通知する（LINE → だめならメール）
// =====================================================
function notifyUrgent(message, thread, urgentCases) {
  const lines = ['【緊急の案件依頼】', '件名：' + message.getSubject()];

  urgentCases.forEach(function(c, i) {
    lines.push('');
    if (urgentCases.length > 1) lines.push('＜' + (i + 1) + '件目＞');
    lines.push('顧客：' + extractVal(c['顧客名']));
    lines.push('場所：' + extractVal(c['作業場所']));
    lines.push('内容：' + extractVal(c['工事内容']));
    lines.push('希望：' + extractVal(c['希望日原文']));
    lines.push('電話：' + extractVal(c['電話番号']));
  });

  lines.push('');
  lines.push('メール：' + thread.getPermalink());

  const text = lines.join('\n');

  // LINEで送れなければメールで送る
  if (sendLineMessage(text)) return;

  GmailApp.sendEmail(
    Session.getEffectiveUser().getEmail(),
    '【緊急】案件依頼：' + message.getSubject(),
    text
  );
}

function sendLineMessage(text) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const userId = props.getProperty('LINE_USER_ID');

  if (!token || !userId) {
    Logger.log('LINEの設定がないため、メールで通知します');
    return false;
  }

  const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({
      to: userId,
      messages: [{ type: 'text', text: text }]
    }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    Logger.log('LINE送信エラー（' + response.getResponseCode() + '）：' + response.getContentText());
    return false;
  }
  return true;
}

// LINE通知のテスト（メニューから実行）
function testLineNotify() {
  const ok = sendLineMessage('【テスト】案件依頼の緊急通知テストです');
  const msg = ok ? 'LINEに送信しました' : 'LINEに送信できませんでした（実行ログを確認してください）';
  Logger.log(msg);
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'LINE通知テスト', 5);
  } catch (e) {}
}


// =====================================================
// 承認した行を案件管理シートへ転記する
// =====================================================

// 確認状況を「承認」にした瞬間に自動で転記する
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== MAIL_CONFIG.SHEET_NAME) return;

  const firstCol = e.range.getColumn();
  const lastCol = firstCol + e.range.getNumColumns() - 1;
  if (COL['確認状況'] < firstCol || COL['確認状況'] > lastCol) return;

  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();

  for (let r = startRow; r < startRow + numRows; r++) {
    if (r < 2) continue;
    try {
      transferRow(sheet, r);
    } catch (err) {
      markTransfer(sheet, r, 'エラー：' + err.message, false);
    }
  }
}

// 承認済みでまだ転記していない行をまとめて転記する（メニューから実行）
function transferAllApproved() {
  const sheet = getExtractSheet();
  const lastRow = sheet.getLastRow();
  let count = 0;

  for (let r = 2; r <= lastRow; r++) {
    try {
      if (transferRow(sheet, r)) count++;
    } catch (err) {
      markTransfer(sheet, r, 'エラー：' + err.message, false);
    }
  }

  const msg = '転記しました：' + count + '件';
  Logger.log(msg);
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '案件管理へ転記', 5);
  } catch (e) {}
}

// 1行を転記する。転記できたらtrueを返す
function transferRow(sheet, rowNum) {
  const values = sheet.getRange(rowNum, 1, 1, EXTRACT_HEADERS.length).getValues()[0];
  const get = function(h) { return values[COL[h] - 1]; };

  if (get('確認状況') !== '承認') return false;
  if (String(get('転記結果')) !== '') return false;  // 転記済み

  const caseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIL_CONFIG.CASE_SHEET_NAME);
  if (!caseSheet) {
    markTransfer(sheet, rowNum, 'エラー：「' + MAIL_CONFIG.CASE_SHEET_NAME + '」シートがありません', false);
    return false;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const hantei = get('判定');
  const kibouDate = parseDateValue(get('希望日'));

  // ----- 新規依頼：案件番号を採番して追加 -----
  if (hantei === '新規依頼') {
    const caseNo = nextCaseNumber(caseSheet, today);

    const bikou = [
      '作業場所：' + get('作業場所'),
      '物件名：' + get('物件名'),
      '連絡者：' + get('連絡者'),
      '電話：' + get('電話番号'),
      '要確認：' + get('要確認事項')
    ].join('／');

    const newRow = [
      caseNo,                      // 案件番号
      get('顧客名'),               // 顧客名
      get('工事内容'),             // 工事内容
      today,                       // 受注日（受付日）
      kibouDate || '',             // 工事予定日
      MAIL_CONFIG.NEW_CASE_STATUS, // ステータス
      '',                          // 担当者（未割り当て）
      today,                       // 最終更新日
      '',                          // 金額
      bikou                        // 備考
    ];

    const target = caseSheet.getLastRow() + 1;
    caseSheet.getRange(target, 1).setNumberFormat('@');
    caseSheet.getRange(target, 1, 1, newRow.length).setValues([newRow]);

    sheet.getRange(rowNum, COL['案件番号']).setValue(caseNo);
    markTransfer(sheet, rowNum, '新規登録', true);
    return true;
  }

  // ----- 既存案件の変更：該当する案件を更新 -----
  if (hantei === '既存案件の変更') {
    const relNo = String(get('関連案件番号')).trim();
    if (!relNo) {
      markTransfer(sheet, rowNum, 'エラー：関連案件番号がありません', false);
      return false;
    }

    const caseRow = findCaseRow(caseSheet, relNo);
    if (caseRow === -1) {
      markTransfer(sheet, rowNum, 'エラー：' + relNo + ' が案件管理にありません', false);
      return false;
    }

    const changes = [];

    if (kibouDate) {
      const oldDate = caseSheet.getRange(caseRow, 5).getValue();
      caseSheet.getRange(caseRow, 5).setValue(kibouDate);
      changes.push('工事予定日 ' + fmtMd(oldDate) + '→' + fmtMd(kibouDate));
    }

    caseSheet.getRange(caseRow, 8).setValue(today);  // 最終更新日

    const oldBikou = caseSheet.getRange(caseRow, 10).getValue();
    const memo = '【' + fmtMd(today) + ' 変更依頼】' + get('要確認事項');
    caseSheet.getRange(caseRow, 10).setValue(oldBikou ? oldBikou + '／' + memo : memo);

    sheet.getRange(rowNum, COL['案件番号']).setValue(relNo);
    const resultText = changes.length > 0
      ? '更新：' + changes.join('、')
      : '更新：備考に追記（日付の変更なし）';
    markTransfer(sheet, rowNum, resultText, true);
    return true;
  }

  // ----- それ以外は転記しない -----
  markTransfer(sheet, rowNum, '対象外（' + hantei + '）', false);
  return false;
}

// 転記結果を書き込み、成功は緑・失敗は赤で表示
function markTransfer(sheet, rowNum, text, success) {
  const cell = sheet.getRange(rowNum, COL['転記結果']);
  cell.setValue(text);
  cell.setBackground(success ? '#c8e6c9' : '#ffcdd2');
}

// 次の案件番号を作る（例：2026-016）
function nextCaseNumber(caseSheet, today) {
  const year = today.getFullYear();
  const lastRow = caseSheet.getLastRow();
  let max = 0;

  if (lastRow >= 2) {
    const nums = caseSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    nums.forEach(function(row) {
      const m = String(row[0]).match(/^(\d{4})-(\d+)$/);
      if (m && Number(m[1]) === year) {
        max = Math.max(max, Number(m[2]));
      }
    });
  }

  return year + '-' + ('000' + (max + 1)).slice(-3);
}

// 案件番号から案件管理シートの行番号を探す（なければ-1）
function findCaseRow(caseSheet, caseNo) {
  const lastRow = caseSheet.getLastRow();
  if (lastRow < 2) return -1;

  const nums = caseSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < nums.length; i++) {
    if (String(nums[i][0]).trim() === caseNo) return i + 2;
  }
  return -1;
}

// セルの値を日付に変換する（日付でなければnull）
function parseDateValue(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const m = String(value).match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return null;
}

// 日付を M/d 形式にする
function fmtMd(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '未設定';
  return Utilities.formatDate(date, 'Asia/Tokyo', 'M/d');
}


// =====================================================
// 書き出し先のシートを用意する
// =====================================================
function getExtractSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MAIL_CONFIG.SHEET_NAME);

  if (sheet) {
    // 列構成が古い場合は止める
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (header.length < EXTRACT_HEADERS.length || header[COL['転記結果'] - 1] !== '転記結果') {
      throw new Error('「' + MAIL_CONFIG.SHEET_NAME + '」シートの列構成が古いです。シートを削除してから再実行してください');
    }
    return sheet;
  }

  sheet = ss.insertSheet(MAIL_CONFIG.SHEET_NAME);

  sheet.getRange(1, 1, 1, EXTRACT_HEADERS.length)
    .setValues([EXTRACT_HEADERS])
    .setFontWeight('bold')
    .setBackground('#e8eaf6');
  sheet.setFrozenRows(1);

  // 電話番号・案件番号が勝手に数値や日付に変換されないよう文字列扱いにする
  const maxRows = sheet.getMaxRows();
  ['電話番号', '関連案件番号', '案件番号'].forEach(function(h) {
    sheet.getRange(1, COL[h], maxRows, 1).setNumberFormat('@');
  });

  // 確認状況をプルダウンにする
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未確認', '承認', '却下', '確認不要'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, COL['確認状況'], 999, 1).setDataValidation(rule);

  // 列幅
  const widths = [
    120, 120, 180, 110, 160, 100, 120, 220, 140, 220, 100, 160,
    70, 80, 100, 260, 90, 200, 90, 100, 100, 90, 140
  ];
  widths.forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });

  return sheet;
}


// =====================================================
// 判定や緊急度に応じて行に色をつける
// =====================================================
function colorExtractRows(sheet, startRow, rows) {
  rows.forEach(function(row, i) {
    const hantei = row[COL['判定'] - 1];
    const kinkyu = row[COL['緊急度'] - 1];
    let color = null;

    if (hantei === '案件以外') {
      color = '#eeeeee';   // グレー
    } else if (hantei === '既存案件の変更') {
      color = '#e3f2fd';   // 水色
    } else if (kinkyu === '高') {
      color = '#ffcdd2';   // 赤
    }

    if (color) {
      sheet.getRange(startRow + i, 1, 1, EXTRACT_HEADERS.length).setBackground(color);
    }
  });
}


// =====================================================
// テスト用：処理済みラベルを外して、もう一度読み込めるようにする
// =====================================================
function resetDoneLabels() {
  const doneLabel = GmailApp.getUserLabelByName(MAIL_CONFIG.DONE_LABEL);
  const targetLabel = GmailApp.getUserLabelByName(MAIL_CONFIG.TARGET_LABEL);
  if (!doneLabel || !targetLabel) return;

  const threads = targetLabel.getThreads(0, 100);
  threads.forEach(function(thread) {
    thread.removeLabel(doneLabel);
  });
  Logger.log('処理済みラベルを外しました：' + threads.length + '件');
}


// =====================================================
// メニュー（対応待ち一覧とメール抽出をまとめた版）
// =====================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('業務支援')
    .addItem('対応待ち一覧を作成', 'createTodoList')
    .addItem('一覧をメールで送信', 'sendTodoMail')
    .addItem('作成＋メール送信', 'dailyTask')
    .addSeparator()
    .addItem('メールから案件を抽出', 'extractFromMails')
    .addItem('承認済みを案件管理へ転記', 'transferAllApproved')
    .addSeparator()
    .addItem('LINE通知のテスト', 'testLineNotify')
    .addToUi();
}
