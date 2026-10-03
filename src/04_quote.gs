/**
 * 見積書作成 v3
 *
 * 抽出結果（確認待ち）シートの行を選んで実行すると、
 * 工事内容をAIに読ませて単価マスタの項目に振り分け、見積書シートに書き込み、
 * PDF化して案件書類フォルダに保存する。
 *
 * 役割分担
 *   AI  … 工事内容が単価マスタのどの項目にあたるか、数量はいくつかを判定するだけ
 *   GAS … マスタから単価を引く／シートに書き込む／PDFにして保存する
 *   人  … 単価の最終決定、値引きの判断、承認
 *
 * 金額の計算は見積書シートの数式が行う。AIにもGASにも計算させない。
 *
 * v3で直したところ
 *   1. PDF化のとき、明細の見出し（濃紺・白文字）まで白くしていた
 *      → 色を退避するのは入力欄だけにした
 *   2. 入力欄の青い文字がそのままPDFに出ていた
 *      → PDF化の間だけ黒にする
 *   3. 数量が「1.」と表示されていた → テンプレート側の表示形式を修正
 *   4. 自社情報・有効期限・挨拶文・御見積金額欄を追加（テンプレートv2）
 */

// ===== 設定 =====
const QUOTE_CONFIG = {
  EXTRACT_SHEET: '抽出結果（確認待ち）',
  MASTER_SHEET: '単価マスタ',
  QUOTE_SHEET: '見積書',
  CASE_SHEET: '案件管理',

  ROOT_FOLDER_ID: '1P0WJDVvQMWcJ_UGvVE7-8fe86_ST3qGr',

  // 抽出結果シートの見出し名
  H_STATUS: '確認状況',
  H_CASE_NO: '案件番号',
  H_CUSTOMER: '顧客名',
  H_WORK: '工事内容',
  H_RESULT: '転記結果',

  // 案件管理シートの見出し名
  C_CASE_NO: '案件番号',
  C_CUSTOMER: '顧客名',
  C_WORK: '工事内容',
  C_STATUS: 'ステータス',
  C_UPDATED: '最終更新日',
  NEW_STATUS: '見積中',

  TRIGGER_VALUE: '見積作成',

  // 見積書シート（テンプレートv2の位置）
  ROW_TO: 3, ROW_SUBJECT: 4, ROW_CASE_NO: 5, ROW_DATE: 6,
  ROW_QUOTE_NO: 7, ROW_VALID: 8,
  DETAIL_FIRST: 14, DETAIL_LAST: 21,
  ROW_DISCOUNT: 24, ROW_NOTE: 29,
  COL_ITEM: 2, COL_QTY: 3, COL_UNIT: 4, COL_PRICE: 5,

  VALID_DAYS: 30,             // 見積の有効期限（発行日からの日数）

  PDF_RANGE: 'A1:F31',        // PDFに出す範囲（この外の注記は出力しない）
  INPUT_RANGES: ['B3:D8', 'B14:E21', 'F24:F25', 'B29:F30'],  // 色を退避する範囲

  MODEL: 'gpt-4.1-mini'
};

// 顧客名にこれらが含まれていれば「御中」、なければ「様」
const CORP_WORDS = ['株式会社', '有限会社', '合同会社', '（株）', '(株)', '㈱',
                    '（有）', '(有)', '組合', 'センター', 'サービス', '商事',
                    '工業', '建設', '不動産', '管理', 'ビル'];


/**
 * 確認状況が「見積作成」になったときに動く。
 * ※メール抽出側にも onEdit がある場合は、そちらから onEditQuote(e) を呼ぶこと
 */
function onEditQuote(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() !== QUOTE_CONFIG.EXTRACT_SHEET) return;
  if (String(e.value).trim() !== QUOTE_CONFIG.TRIGGER_VALUE) return;
  createQuoteFromRow(e.range.getRow());
}


/**
 * 手動実行用。抽出結果シートで選んでいる行の見積書を作る。
 */
function createQuoteForSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== QUOTE_CONFIG.EXTRACT_SHEET) {
    SpreadsheetApp.getUi().alert(
      '「' + QUOTE_CONFIG.EXTRACT_SHEET + '」シートで、見積書を作りたい行を選んでから実行してください。');
    return;
  }
  createQuoteFromRow(sheet.getActiveRange().getRow());
}


/**
 * 指定行から見積書を作る本体
 */
function createQuoteFromRow(rowNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(QUOTE_CONFIG.EXTRACT_SHEET);

  if (rowNum < 2) {
    Logger.log('見出し行が選ばれています。データの行を選んでください。');
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(function(h) { return String(h).trim(); });
  const idx = {};
  headers.forEach(function(h, i) { idx[h] = i + 1; });

  const get = function(name) {
    return idx[name] ? String(sheet.getRange(rowNum, idx[name]).getValue()).trim() : '';
  };

  const customer = get(QUOTE_CONFIG.H_CUSTOMER);
  const work = get(QUOTE_CONFIG.H_WORK);
  let caseNo = get(QUOTE_CONFIG.H_CASE_NO);
  const isNewCase = !caseNo;

  if (!work) {
    setResult_(sheet, rowNum, idx, '工事内容が空です');
    return;
  }

  const master = readMaster_(ss);
  if (master.length === 0) {
    setResult_(sheet, rowNum, idx, '単価マスタが空です');
    return;
  }

  let items;
  try {
    items = askAiForItems_(work, master);
  } catch (err) {
    setResult_(sheet, rowNum, idx, 'AI判定に失敗：' + err.message);
    return;
  }

  const lines = [];
  const notes = [];

  items.forEach(function(it) {
    const name = String(it.item || '').trim();
    const qty = it.qty;

    if (!name || name === '該当なし' || name === '不明') {
      notes.push(it.reason || work);
      return;
    }
    const m = findMaster_(master, name);
    if (!m) {
      notes.push('単価マスタにない項目：' + name);
      return;
    }
    const q = (qty === null || qty === undefined || qty === '') ? '' : qty;
    lines.push({ item: m.name, qty: q, unit: m.unit, price: m.price });
    if (q === '') notes.push('数量が不明：' + m.name);
  });

  if (lines.length === 0) {
    setResult_(sheet, rowNum, idx, '見積書を作れません（' + notes.join(' / ') + '）');
    Logger.log('項目を特定できず：' + notes.join(' / '));
    return;
  }

  if (isNewCase) caseNo = issueCaseNumber_(ss);

  writeQuoteSheet_(ss, {
    caseNo: caseNo, customer: customer, work: work, lines: lines, notes: notes
  });

  let fileUrl = '';
  try {
    fileUrl = saveQuoteAsPdf_(ss, caseNo);
  } catch (err) {
    setResult_(sheet, rowNum, idx, 'PDF保存に失敗：' + err.message);
    return;
  }

  if (idx[QUOTE_CONFIG.H_CASE_NO]) {
    sheet.getRange(rowNum, idx[QUOTE_CONFIG.H_CASE_NO]).setValue(caseNo);
  }
  if (isNewCase) registerCase_(ss, caseNo, customer, work);

  const msg = '見積書を作成（' + lines.length + '項目）' +
              (notes.length ? '／要確認：' + notes.join(' / ') : '');
  setResult_(sheet, rowNum, idx, msg);
  Logger.log(caseNo + '：' + msg + '｜' + fileUrl);
}


/**
 * 単価マスタを読む
 */
function readMaster_(ss) {
  const sheet = ss.getSheetByName(QUOTE_CONFIG.MASTER_SHEET);
  if (!sheet) throw new Error('「' + QUOTE_CONFIG.MASTER_SHEET + '」シートが見つかりません。');

  const values = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < values.length; i++) {
    const name = String(values[i][0]).trim();
    if (!name || name.indexOf('※') === 0) continue;
    list.push({
      name: name,
      unit: String(values[i][1]).trim(),
      price: Number(values[i][2]) || 0
    });
  }
  return list;
}


/**
 * AIに「単価マスタのどの項目か」「数量はいくつか」だけを判定させる。
 */
function askAiForItems_(work, master) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY が設定されていません');

  const menu = master.map(function(m) {
    return '- ' + m.name + ' ／ 単位：' + m.unit;
  }).join('\n');

  const prompt =
    '次の依頼内容を読み、作業項目一覧のどれにあたるかと、数量を判定してください。\n\n' +
    '【依頼内容】\n' + work + '\n\n' +
    '【作業項目一覧】\n' + menu + '\n\n' +
    '【ルール】\n' +
     '1. 作業項目は、必ず一覧の名称をそのまま使う。言い換えない。\n' +
    '   item には名称だけを書く。単位・括弧書き・説明を付け足さない。\n' +
    '   （正しい例："エアコン洗浄" ／ 誤った例："エアコン洗浄（単位：台）"）\n' +
    '2. 一覧にまったく該当する作業がない場合：\n' +
    '   item を "該当なし"、reason を「一覧にない作業：〇〇」の形で書く。\n' +
    '   似ている項目に無理に当てはめてはいけない。\n' +
    '3. 候補は一覧にあるが、どれか決められない場合（例：洗浄か交換か書かれていない）：\n' +
    '   item を "該当なし"、reason を「判断できない：〇〇が不明（候補：A、B）」の形で書く。\n' +
    '   候補には一覧の名称をそのまま入れる。推測で決めない。\n' +
    '   ※2と3は人の対応が変わるので、必ず書き分けること。\n' +
    '4. 数量が書かれていない場合は qty を null にする。1と決めつけない。\n' +
    '5. 1つの依頼に複数の作業が含まれる場合は、配列で複数返す。\n' +
    '6. 金額・単価・値引きは判定しない。出力に含めない。\n\n' +
    '【出力形式】JSONのみ。前置きやコードブロックの記号は付けない。\n' +
    '{"items":[{"item":"作業項目名 または 該当なし","qty":数量 または null,"reason":"該当なしの場合の理由"}]}';

  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify({
      model: QUOTE_CONFIG.MODEL,
      messages: [
        { role: 'system', content: '正確に判定し、指定のJSONだけを返す。推測で埋めない。' },
        { role: 'user', content: prompt }
      ],
      temperature: 0
    }),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) throw new Error('APIエラー ' + res.getResponseCode());

  let text = JSON.parse(res.getContentText()).choices[0].message.content;
  text = text.replace(/```json|```/g, '').trim();

  Logger.log('AI判定：' + text);
  return (JSON.parse(text).items || []);
}


/**
 * 見積書シートに書き込む（数式には触らない）
 */
function writeQuoteSheet_(ss, data) {
  const sheet = ss.getSheetByName(QUOTE_CONFIG.QUOTE_SHEET);
  if (!sheet) throw new Error('「' + QUOTE_CONFIG.QUOTE_SHEET + '」シートが見つかりません。');

  const C = QUOTE_CONFIG;
  const today = new Date();

  sheet.getRange(C.DETAIL_FIRST, C.COL_ITEM,
                 C.DETAIL_LAST - C.DETAIL_FIRST + 1, 4).clearContent();

  const limit = new Date(today.getTime() + C.VALID_DAYS * 24 * 60 * 60 * 1000);

  sheet.getRange(C.ROW_TO, 2).setValue(honorific_(data.customer));
  sheet.getRange(C.ROW_SUBJECT, 2).setValue(data.work);
  sheet.getRange(C.ROW_CASE_NO, 2).setValue(data.caseNo);
  sheet.getRange(C.ROW_DATE, 2)
       .setValue(Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd'));
  sheet.getRange(C.ROW_QUOTE_NO, 2).setValue('Q' + data.caseNo);
  sheet.getRange(C.ROW_VALID, 2)
       .setValue(Utilities.formatDate(limit, 'Asia/Tokyo', 'yyyy/MM/dd') + ' まで');

  data.lines.forEach(function(line, i) {
    const r = C.DETAIL_FIRST + i;
    if (r > C.DETAIL_LAST) return;
    sheet.getRange(r, C.COL_ITEM).setValue(line.item);
    sheet.getRange(r, C.COL_QTY).setValue(line.qty);
    sheet.getRange(r, C.COL_UNIT).setValue(line.unit);
    sheet.getRange(r, C.COL_PRICE).setValue(line.price);
  });

  sheet.getRange(C.ROW_DISCOUNT, 6).setValue(0);   // 値引きは必ず0に戻す
  sheet.getRange(C.ROW_NOTE, 2).setValue(
    data.notes.length ? '【要確認】' + data.notes.join(' / ') : '');

  SpreadsheetApp.flush();
}


/**
 * 顧客名に敬称を付ける。会社らしい語が含まれていれば「御中」、なければ「様」。
 */
function honorific_(name) {
  if (!name) return '';
  const isCorp = CORP_WORDS.some(function(w) { return name.indexOf(w) >= 0; });
  return name + (isCorp ? ' 御中' : ' 様');
}


/**
 * 見積書シートをPDFにして案件フォルダに保存する。
 * 入力欄の黄色と青文字は、PDF化の間だけ白地・黒文字に退避する。
 * 明細の見出し（濃紺・白文字）には触らない。
 */
function saveQuoteAsPdf_(ss, caseNo) {
  const sheet = ss.getSheetByName(QUOTE_CONFIG.QUOTE_SHEET);

  const ranges = QUOTE_CONFIG.INPUT_RANGES.map(function(a1) {
    return sheet.getRange(a1);
  });
  const savedBg = ranges.map(function(r) { return r.getBackgrounds(); });
  const savedFc = ranges.map(function(r) { return r.getFontColors(); });

  ranges.forEach(function(r) {
    r.setBackground('#ffffff');
    r.setFontColor('#000000');
  });
  SpreadsheetApp.flush();

  let blob;
  try {
    const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
                '/export?format=pdf' +
                '&gid=' + sheet.getSheetId() +
                '&range=' + encodeURIComponent(QUOTE_CONFIG.PDF_RANGE) +
                '&portrait=true&size=A4&fitw=true' +
                '&gridlines=false&printtitle=false&sheetnames=false' +
                '&pagenum=false&fzr=false' +
                '&top_margin=0.6&bottom_margin=0.6&left_margin=0.6&right_margin=0.6';

    blob = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
    }).getBlob();
  } finally {
    // 失敗しても必ず元に戻す
    ranges.forEach(function(r, i) {
      r.setBackgrounds(savedBg[i]);
      r.setFontColors(savedFc[i]);
    });
    SpreadsheetApp.flush();
  }

  const root = DriveApp.getFolderById(QUOTE_CONFIG.ROOT_FOLDER_ID);

  let folder = null;
  const folders = root.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    if (String(f.getName()).trim() === caseNo) { folder = f; break; }
  }
  if (!folder) {
    folder = root.createFolder(caseNo);
    Logger.log(caseNo + '：フォルダを新規作成しました');
  }

  const fileName = caseNo + '_見積書.pdf';
  const olds = folder.getFilesByName(fileName);
  while (olds.hasNext()) olds.next().setTrashed(true);

  return folder.createFile(blob.setName(fileName)).getUrl();
}


/**
 * 案件番号を採番する。
 * 案件管理シート・抽出結果シート・ドライブのフォルダ名を全部見て、最大値の次を返す。
 */
function issueCaseNumber_(ss) {
  const year = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy');
  const re = new RegExp('^' + year + '-(\\d+)$');
  let max = 0;

  const scan = function(text) {
    const m = String(text).trim().match(re);
    if (m) max = Math.max(max, Number(m[1]));
  };

  const caseSheet = ss.getSheetByName(QUOTE_CONFIG.CASE_SHEET);
  if (caseSheet) {
    caseSheet.getDataRange().getValues().forEach(function(row) { scan(row[0]); });
  }

  const exSheet = ss.getSheetByName(QUOTE_CONFIG.EXTRACT_SHEET);
  if (exSheet) {
    const values = exSheet.getDataRange().getValues();
    const col = values[0].map(function(h) { return String(h).trim(); })
                         .indexOf(QUOTE_CONFIG.H_CASE_NO);
    if (col >= 0) {
      for (let i = 1; i < values.length; i++) scan(values[i][col]);
    }
  }

  try {
    const folders = DriveApp.getFolderById(QUOTE_CONFIG.ROOT_FOLDER_ID).getFolders();
    while (folders.hasNext()) scan(folders.next().getName());
  } catch (e) {
    Logger.log('フォルダの確認に失敗しました：' + e.message);
  }

  return year + '-' + ('000' + (max + 1)).slice(-3);
}


/**
 * 案件管理シートに新しい案件を登録する（ステータスは見積中）
 */
function registerCase_(ss, caseNo, customer, work) {
  const sheet = ss.getSheetByName(QUOTE_CONFIG.CASE_SHEET);
  if (!sheet) { Logger.log('案件管理シートがないため登録を飛ばしました'); return; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(function(h) { return String(h).trim(); });
  const row = sheet.getLastRow() + 1;

  const put = function(name, value) {
    const c = headers.indexOf(name);
    if (c >= 0) sheet.getRange(row, c + 1).setValue(value);
  };

  put(QUOTE_CONFIG.C_CASE_NO, caseNo);
  put(QUOTE_CONFIG.C_CUSTOMER, customer);
  put(QUOTE_CONFIG.C_WORK, work);
  put(QUOTE_CONFIG.C_STATUS, QUOTE_CONFIG.NEW_STATUS);
  put(QUOTE_CONFIG.C_UPDATED,
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'));

  Logger.log(caseNo + '：案件管理シートに登録しました');
}


/**
 * 転記結果の列にメッセージを書く
 */
function setResult_(sheet, rowNum, idx, msg) {
  const col = idx[QUOTE_CONFIG.H_RESULT];
  if (!col) { Logger.log(msg); return; }

  const cell = sheet.getRange(rowNum, col);
  cell.setValue(msg);
  cell.setBackground(msg.indexOf('見積書を作成') === 0 ? '#d9ead3' : '#f4cccc');
}
/**
 * 単価マスタから項目を探す。
 * 完全一致で見つからなければ、マスタの名称を含んでいるかで探す。
 * AIが「エアコン洗浄（単位：台）」のように余計な語を付けても拾えるようにするため。
 */
function findMaster_(master, name) {
  const exact = master.filter(function(x) { return x.name === name; })[0];
  if (exact) return exact;

  const hits = master.filter(function(x) { return name.indexOf(x.name) >= 0; });
  if (hits.length === 0) return null;

  // 複数当たったときは、より長い名称を優先する
  hits.sort(function(a, b) { return b.name.length - a.name.length; });
  Logger.log('名称がずれていたため部分一致で照合：「' + name + '」→「' + hits[0].name + '」');
  return hits[0];
}
