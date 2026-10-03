/**
 * 対応待ち一覧を作成する（担当者別・緊急度色分け）
 */
function createTodoList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('案件管理');
  
  if (!sourceSheet) {
    throw new Error('「案件管理」シートが見つかりません');
  }
  
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('データがありません');
  }
  
  const data = sourceSheet.getRange(2, 1, lastRow - 1, 10).getValues();
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // 担当者ごとに格納する箱
  const byTantou = {};
  
  function addItem(tantou, kubun, item) {
    if (!byTantou[tantou]) {
      byTantou[tantou] = {
        chakushu: [],
        seikyu: [],
        houchi: [],
        mitsumori: []
      };
    }
    byTantou[tantou][kubun].push(item);
  }
  
  data.forEach(function(row) {
    const bangou = row[0];
    const kokyaku = row[1];
    const koji = row[2];
    const yoteibi = row[4];
    const status = row[5];
    const tantou = row[6] || '担当者未設定';
    const koushinbi = row[7];
    const kingaku = row[8];
    
    if (!bangou) return;
    
    const keikaNissu = koushinbi ? daysBetween(koushinbi, today) : 0;
    
    // ルール①：着手漏れ（工事予定日が3日以内 かつ 受注済）
    if (status === '受注済' && yoteibi) {
      const nokoriNissu = daysBetween(today, yoteibi);
      if (nokoriNissu >= 0 && nokoriNissu <= 3) {
        addItem(tantou, 'chakushu', {
          bangou: bangou,
          kokyaku: kokyaku,
          koji: koji,
          detail: formatDate(yoteibi) + '（' + nokoriDisplay(nokoriNissu) + '）',
          kingaku: kingaku,
          sort: nokoriNissu
        });
      }
    }
    
    // ルール②：放置案件（14日以上動きなし・完了/請求済以外）
    if (status !== '完了' && status !== '請求済' && keikaNissu >= 14) {
      addItem(tantou, 'houchi', {
        bangou: bangou,
        kokyaku: kokyaku,
        koji: koji,
        detail: status + '／最終更新 ' + formatDate(koushinbi) + '（' + keikaNissu + '日経過）',
        kingaku: kingaku,
        sort: -keikaNissu
      });
    }
    
    // ルール③：請求漏れ（完了のまま）
    if (status === '完了') {
      addItem(tantou, 'seikyu', {
        bangou: bangou,
        kokyaku: kokyaku,
        koji: koji,
        detail: '完了 ' + formatDate(koushinbi) + '（' + keikaNissu + '日経過）',
        kingaku: kingaku,
        sort: -keikaNissu
      });
    }
    
    // ルール④：見積放置（見積中・10日以上経過）
    if (status === '見積中' && keikaNissu >= 10) {
      addItem(tantou, 'mitsumori', {
        bangou: bangou,
        kokyaku: kokyaku,
        koji: koji,
        detail: '見積中／最終更新 ' + formatDate(koushinbi) + '（' + keikaNissu + '日経過）',
        kingaku: kingaku,
        sort: -keikaNissu
      });
    }
  });
  
  // 出力シートを準備
  let outSheet = ss.getSheetByName('対応待ち一覧');
  if (outSheet) {
    outSheet.clear();
  } else {
    outSheet = ss.insertSheet('対応待ち一覧');
  }
  
  const output = [];
  output.push(['【対応待ち一覧】' + formatDate(today) + ' 時点', '', '', '', '']);
  output.push(['', '', '', '', '']);
  
  const tantouList = Object.keys(byTantou).sort();
  let grandTotal = 0;
  
  tantouList.forEach(function(tantou) {
    const box = byTantou[tantou];
    
    box.chakushu.sort(function(a, b) { return a.sort - b.sort; });
    box.seikyu.sort(function(a, b) { return a.sort - b.sort; });
    box.houchi.sort(function(a, b) { return a.sort - b.sort; });
    box.mitsumori.sort(function(a, b) { return a.sort - b.sort; });
    
    const count = box.chakushu.length + box.seikyu.length + box.houchi.length + box.mitsumori.length;
    grandTotal += count;
    
    output.push(['●《' + tantou + '》　' + count + '件', '', '', '', '']);
    output.push(['', '', '', '', '']);
    
    addSection(output, '　■ 着手漏れ（工事予定日が近い）', box.chakushu);
    addSection(output, '　■ 請求漏れ（完了済・未請求）', box.seikyu);
    addSection(output, '　■ 放置案件（14日以上動きなし）', box.houchi);
    addSection(output, '　■ 見積放置（10日以上返答なし）', box.mitsumori);
    
    output.push(['', '', '', '', '']);
  });
  
  output.push(['全体合計 ' + grandTotal + ' 件', '', '', '', '']);
  
  outSheet.getRange(1, 1, output.length, 5).setValues(output);
  formatSheet(outSheet, output);
  
  Logger.log('対応待ち一覧を作成しました');
}

/**
 * セクションを出力配列に追加
 */
function addSection(output, title, items) {
  output.push([title, '', '', '', '']);
  
  if (items.length === 0) {
    output.push(['　該当なし', '', '', '', '']);
  } else {
    output.push(['案件番号', '顧客名', '工事内容', '状況', '担当']);
    items.forEach(function(item) {
      output.push([item.bangou, item.kokyaku, item.koji, item.detail, '']);
    });
  }
  
  output.push(['', '', '', '', '']);
}

/**
 * 2つの日付の差（日数）を返す
 */
function daysBetween(from, to) {
  const d1 = new Date(from);
  const d2 = new Date(to);
  d1.setHours(0, 0, 0, 0);
  d2.setHours(0, 0, 0, 0);
  return Math.round((d2 - d1) / (1000 * 60 * 60 * 24));
}

/**
 * 日付を M/d 形式にする
 */
function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'M/d');
}

/**
 * 残り日数の表示文言
 */
function nokoriDisplay(days) {
  if (days === 0) return '本日';
  if (days === 1) return '明日';
  return days + '日後';
}

/**
 * 出力シートの書式を整える
 */
function formatSheet(sheet, output) {
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 280);
  sheet.setColumnWidth(5, 80);
  
  sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  
  let currentSection = '';
  
  for (let i = 0; i < output.length; i++) {
    const cellValue = output[i][0];
    const rowNum = i + 1;
    
    // 担当者見出し
    if (typeof cellValue === 'string' && cellValue.indexOf('●') === 0) {
      sheet.getRange(rowNum, 1, 1, 5)
        .setBackground('#c5cae9')
        .setFontWeight('bold')
        .setFontSize(12);
      continue;
    }
    
    // 区分見出し
    if (typeof cellValue === 'string' && cellValue.indexOf('■') >= 0) {
      sheet.getRange(rowNum, 1, 1, 5)
        .setBackground('#e8eaf6')
        .setFontWeight('bold');
      
      if (cellValue.indexOf('着手漏れ') >= 0) currentSection = 'chakushu';
      else if (cellValue.indexOf('請求漏れ') >= 0) currentSection = 'seikyu';
      else if (cellValue.indexOf('放置案件') >= 0) currentSection = 'houchi';
      else if (cellValue.indexOf('見積放置') >= 0) currentSection = 'mitsumori';
      continue;
    }
    
    // 列見出し
    if (cellValue === '案件番号') {
      sheet.getRange(rowNum, 1, 1, 5)
        .setBackground('#f5f5f5')
        .setFontWeight('bold');
      continue;
    }
    
    // データ行の色付け
    const detail = output[i][3];
    if (typeof detail === 'string' && detail !== '') {
      const color = getUrgencyColor(currentSection, detail);
      if (color) {
        sheet.getRange(rowNum, 1, 1, 5).setBackground(color);
      }
    }
  }
  
  sheet.setFrozenRows(1);
}

/**
 * 緊急度に応じた色を返す
 */
function getUrgencyColor(section, detail) {
  const keikaMatch = detail.match(/（(\d+)日経過）/);
  const keika = keikaMatch ? parseInt(keikaMatch[1]) : null;
  
  if (section === 'chakushu') {
    if (detail.indexOf('本日') >= 0 || detail.indexOf('明日') >= 0) {
      return '#ffcdd2';  // 赤
    }
    return '#ffe0b2';    // オレンジ
  }
  
  if (section === 'seikyu' || section === 'houchi') {
    if (keika === null) return null;
    if (keika >= 30) return '#ffcdd2';  // 赤
    if (keika >= 14) return '#ffe0b2';  // オレンジ
    return '#fff9c4';                    // 黄色
  }
  
  if (section === 'mitsumori') {
    if (keika !== null && keika >= 20) return '#ffe0b2';
    return '#fff9c4';
  }
  
  return null;
}

/**
 * 対応待ち一覧をメールで送る
 */
function sendTodoMail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('対応待ち一覧');
  
  if (!sheet) {
    throw new Error('「対応待ち一覧」シートが見つかりません。先に一覧を作成してください');
  }
  
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 5).getValues();
  
  const today = new Date();
  const dateStr = Utilities.formatDate(today, 'Asia/Tokyo', 'M月d日');
  
  let body = '';
  let htmlBody = '<div style="font-family: sans-serif; font-size: 14px;">';
  
  data.forEach(function(row) {
    const c1 = row[0];
    
    if (c1 === '' || c1 === null) {
      body += '\n';
      htmlBody += '<br>';
      return;
    }
    
    const text = String(c1);
    
    if (text.indexOf('●') === 0) {
      body += '\n' + text + '\n';
      htmlBody += '<div style="background:#c5cae9; padding:6px; font-weight:bold; margin-top:12px;">' + text + '</div>';
      return;
    }
    
    if (text.indexOf('■') >= 0) {
      body += text + '\n';
      htmlBody += '<div style="background:#e8eaf6; padding:4px; font-weight:bold; margin-top:8px;">' + text + '</div>';
      return;
    }
    
    if (text === '案件番号') return;
    
    if (text.indexOf('該当なし') >= 0) {
      body += '　該当なし\n';
      htmlBody += '<div style="color:#888; padding-left:16px;">該当なし</div>';
      return;
    }
    
    if (text.indexOf('【対応待ち一覧】') === 0) {
      body += text + '\n';
      htmlBody += '<div style="font-size:16px; font-weight:bold;">' + text + '</div>';
      return;
    }
    
    if (text.indexOf('合計') >= 0) {
      body += '\n' + text + '\n';
      htmlBody += '<div style="margin-top:12px; font-weight:bold;">' + text + '</div>';
      return;
    }
    
    const line = text + '　' + row[1] + '　' + row[2] + '　' + row[3];
    body += '　' + line + '\n';
    htmlBody += '<div style="padding-left:16px;">' + line + '</div>';
  });
  
  htmlBody += '<br><div style="color:#888; font-size:12px;">このメールは自動送信されています。</div>';
  htmlBody += '</div>';
  
  // 送信先（他の人に送る場合はここを書き換え）
  const to = Session.getActiveUser().getEmail();
  
  GmailApp.sendEmail(
    to,
    '【対応待ち一覧】' + dateStr,
    body,
    { htmlBody: htmlBody }
  );
  
  Logger.log('メールを送信しました: ' + to);
}

/**
 * 一覧作成とメール送信をまとめて実行（トリガー用）
 */
function dailyTask() {
  createTodoList();
  sendTodoMail();
}
