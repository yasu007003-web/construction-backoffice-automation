/**
 * 朝のまとめメール v2
 *
 * これまで別々に作ってきた3つの仕組みの結果を、1通のメールにまとめる。
 *   ・対応待ち一覧（第1話）
 *   ・不足書類チェック（第3話）
 *   ・請求対象の抽出（第6話）
 *
 * 仕組みごとに章を分けるのではなく、緊急度で並べ直す。
 * 作る側は仕組みごとに分かれているが、
 * 受け取る側が知りたいのは「今日まず何をするか」なので。
 *
 * v2で直したところ
 *   同じ案件が複数の仕組みから挙がると、メールに何度も出てきてしまっていた。
 *   案件ごとに1つにまとめ、理由を下に並べる形にした。
 *   その案件がどこに載るかは、いちばん急ぐ理由で決める。
 *
 * 発注照合（第5話）は入れていない。
 * PDFを1件ずつ変換するため時間がかかり、毎朝動かす処理には向かないため。
 */

// ===== 設定 =====
const DIGEST_CONFIG = {
  TODO_SHEET: '対応待ち一覧',
  DOC_SHEET: '不足書類チェック',
  BILL_SHEET: '請求対象',

  TO: '',                    // 空欄なら自分宛て

  URGENT_DAYS: 30            // 「至急」とみなす経過日数
};

// 緊急度（数が小さいほど急ぐ）
const RANK = { urgent: 1, soon: 2, check: 3 };


/**
 * 毎朝のトリガーからはこれを呼ぶ。
 */
function dailyDigest() {
  createTodoList();
  checkMissingDocuments();
  extractBillingTargets();

  sendMorningDigest();
}


/**
 * 3つのシートを読んで、1通のメールにまとめて送る
 */
function sendMorningDigest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // すべての指摘を、いったん1つの配列に集める
  const items = [];
  collectTodoItems_(ss, items);
  collectDocItems_(ss, items);
  const billTotal = collectBillItems_(ss, items);

  // 案件ごとにまとめる
  const box = groupByCase_(items);

  const today = new Date();
  const dateStr = Utilities.formatDate(today, 'Asia/Tokyo', 'M月d日');

  const subject = '【朝の確認】' + dateStr +
                  '　至急' + box.urgent.length + '件' +
                  '／要対応' + box.soon.length + '件';

  const text = buildDigestText_(box, billTotal, dateStr);
  const html = buildDigestHtml_(box, billTotal, dateStr);

  const to = DIGEST_CONFIG.TO || Session.getActiveUser().getEmail();
  GmailApp.sendEmail(to, subject, text, { htmlBody: html });

  Logger.log('朝のまとめを送信：' + to +
             '（指摘' + items.length + '件 → 案件' +
             (box.urgent.length + box.soon.length + box.check.length) + '件）');
}


/**
 * 同じ案件の指摘を1つにまとめる。
 * その案件がどこに載るかは、いちばん急ぐ理由で決める。
 */
function groupByCase_(items) {
  const map = {};

  items.forEach(function(it) {
    if (!map[it.caseNo]) {
      map[it.caseNo] = {
        caseNo: it.caseNo,
        customer: it.customer || '',
        work: it.work || '',
        staff: it.staff || '',
        rank: RANK[it.level],
        reasons: []
      };
    }
    const g = map[it.caseNo];

    // 空のところだけ埋める（先に入った値は上書きしない）
    if (!g.customer && it.customer) g.customer = it.customer;
    if (!g.work && it.work) g.work = it.work;
    if (!g.staff && it.staff) g.staff = it.staff;

    // いちばん急ぐ理由の位置に載せる
    g.rank = Math.min(g.rank, RANK[it.level]);

    g.reasons.push({
      rank: RANK[it.level],
      label: it.label,
      detail: it.detail || ''
    });
  });

  const box = { urgent: [], soon: [], check: [] };

  Object.keys(map).forEach(function(k) {
    const g = map[k];
    // 理由は急ぐ順に並べる
    g.reasons.sort(function(a, b) { return a.rank - b.rank; });

    if (g.rank === RANK.urgent) box.urgent.push(g);
    else if (g.rank === RANK.soon) box.soon.push(g);
    else box.check.push(g);
  });

  // 案件番号の順に並べる
  const byNo = function(a, b) { return a.caseNo < b.caseNo ? -1 : 1; };
  box.urgent.sort(byNo);
  box.soon.sort(byNo);
  box.check.sort(byNo);

  return box;
}


/**
 * 対応待ち一覧シートを読む。
 * 第1話のシートは人が読む形で作ってあるので、見出しを手がかりに拾う。
 */
function collectTodoItems_(ss, items) {
  const sheet = ss.getSheetByName(DIGEST_CONFIG.TODO_SHEET);
  if (!sheet) { Logger.log('対応待ち一覧シートがありません'); return; }

  const last = sheet.getLastRow();
  if (last < 2) return;

  const data = sheet.getRange(1, 1, last, 5).getValues();
  let section = '';
  let staff = '';

  data.forEach(function(row) {
    const c1 = String(row[0] || '').trim();
    if (!c1) return;

    if (c1.indexOf('●') === 0) {
      staff = c1.replace(/[●《》]/g, '').replace(/\s*\d+件$/, '').trim();
      return;
    }
    if (c1.indexOf('■') >= 0) {
      if (c1.indexOf('着手漏れ') >= 0) section = 'chakushu';
      else if (c1.indexOf('請求漏れ') >= 0) section = 'seikyu';
      else if (c1.indexOf('放置案件') >= 0) section = 'houchi';
      else if (c1.indexOf('見積放置') >= 0) section = 'mitsumori';
      else section = '';
      return;
    }
    if (c1 === '案件番号' || c1.indexOf('該当なし') >= 0) return;
    if (c1.indexOf('【') === 0 || c1.indexOf('合計') >= 0) return;

    const detail = String(row[3] || '');
    const base = {
      caseNo: c1,
      customer: String(row[1] || ''),
      work: String(row[2] || ''),
      staff: staff,
      detail: detail
    };

    const m = detail.match(/（(\d+)日経過）/);
    const days = m ? Number(m[1]) : null;

    if (section === 'chakushu') {
      const near = (detail.indexOf('本日') >= 0 || detail.indexOf('明日') >= 0);
      items.push(merge_(base, near ? '工事予定日が近い' : '着手漏れ',
                        near ? 'urgent' : 'soon'));
    } else if (section === 'seikyu' || section === 'houchi') {
      const label = (section === 'seikyu') ? '請求漏れ' : '放置案件';
      const level = (days !== null && days >= DIGEST_CONFIG.URGENT_DAYS)
                  ? 'urgent' : 'soon';
      items.push(merge_(base, label, level));
    } else if (section === 'mitsumori') {
      items.push(merge_(base, '見積放置', 'check'));
    }
  });
}


/**
 * 不足書類チェックシートを読む
 */
function collectDocItems_(ss, items) {
  const rows = readTable_(ss, DIGEST_CONFIG.DOC_SHEET);
  if (!rows) { Logger.log('不足書類チェックシートがありません'); return; }

  rows.forEach(function(r) {
    const judge = String(r['判定'] || '').trim();
    const caseNo = String(r['案件番号'] || '').trim();
    if (!caseNo) return;

    const base = {
      caseNo: caseNo,
      customer: String(r['顧客名'] || ''),
      work: '',
      staff: ''
    };

    if (judge === '不足あり') {
      base.detail = '不足：' + (r['不足書類'] || '');
      items.push(merge_(base, '書類が足りない', 'soon'));
    } else if (judge === 'フォルダなし') {
      base.detail = '不足：' + (r['不足書類'] || '');
      items.push(merge_(base, 'フォルダなし', 'soon'));
    } else if (judge === '要確認') {
      const parts = [];
      if (r['重複の可能性']) parts.push('重複：' + r['重複の可能性']);
      if (r['分類不能ファイル']) parts.push('分類不能：' + r['分類不能ファイル']);
      base.detail = parts.join('／');
      items.push(merge_(base, '書類の要確認', 'check'));
    }
  });
}


/**
 * 請求対象シートを読む。未請求の合計額を返す。
 */
function collectBillItems_(ss, items) {
  const rows = readTable_(ss, DIGEST_CONFIG.BILL_SHEET);
  if (!rows) { Logger.log('請求対象シートがありません'); return 0; }

  let total = 0;

  rows.forEach(function(r) {
    const judge = String(r['判定'] || '').trim();
    const caseNo = String(r['案件番号'] || '').trim();
    if (!caseNo) return;

    const amount = Number(r['請求予定額']);
    if (!isNaN(amount) && amount > 0 && judge !== '請求済') total += amount;

    const base = {
      caseNo: caseNo,
      customer: String(r['顧客名'] || ''),
      work: String(r['工事内容'] || ''),
      staff: String(r['担当者'] || '')
    };

    const money = (!isNaN(amount) && amount > 0)
                ? '／' + amount.toLocaleString() + '円' : '';

    if (judge === '至急') {
      base.detail = '完了から' + r['経過日数'] + '日' + money;
      items.push(merge_(base, '請求がまだ', 'urgent'));
    } else if (judge === '要請求') {
      base.detail = '完了から' + r['経過日数'] + '日' + money;
      items.push(merge_(base, '請求がまだ', 'soon'));
    } else if (judge === '完了日が未入力') {
      base.detail = String(r['備考'] || '');
      items.push(merge_(base, '完了日が未入力', 'check'));
    }
  });

  return total;
}


/**
 * 指摘を1件分の形にする
 */
function merge_(base, label, level) {
  return {
    caseNo: base.caseNo,
    customer: base.customer || '',
    work: base.work || '',
    staff: base.staff || '',
    detail: base.detail || '',
    label: label,
    level: level
  };
}


/**
 * 見出し行のあるシートを、見出し名でひける形にして読む
 */
function readTable_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return null;

  const last = sheet.getLastRow();
  if (last < 2) return [];

  const values = sheet.getRange(1, 1, last, sheet.getLastColumn()).getValues();
  const headers = values[0].map(function(h) { return String(h).trim(); });

  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const obj = {};
    headers.forEach(function(h, j) { if (h) obj[h] = values[i][j]; });
    rows.push(obj);
  }
  return rows;
}


/**
 * 文字だけのメール本文
 */
function buildDigestText_(box, billTotal, dateStr) {
  let s = '【朝の確認】' + dateStr + '\n';
  s += '━━━━━━━━━━━━━━\n\n';

  s += section_('■ 今日やること（' + box.urgent.length + '件）', box.urgent);
  s += section_('■ 今週中に（' + box.soon.length + '件）', box.soon);
  s += section_('■ 手がすいたら（' + box.check.length + '件）', box.check);

  if (billTotal > 0) {
    s += '━━━━━━━━━━━━━━\n';
    s += '未請求の合計：' + billTotal.toLocaleString() + ' 円\n\n';
  }

  s += 'このメールは自動送信されています。\n';
  return s;

  function section_(title, groups) {
    let t = title + '\n';
    if (groups.length === 0) return t + '　なし\n\n';

    groups.forEach(function(g) {
      t += '　' + g.caseNo + ' ' + g.customer;
      if (g.work) t += '（' + g.work + '）';
      t += '\n';
      g.reasons.forEach(function(r) {
        t += '　　・' + r.label + (r.detail ? '：' + r.detail : '') + '\n';
      });
    });
    return t + '\n';
  }
}


/**
 * HTML版のメール本文
 */
function buildDigestHtml_(box, billTotal, dateStr) {
  let h = '<div style="font-family:sans-serif; font-size:14px; line-height:1.7;">';
  h += '<div style="font-size:18px; font-weight:bold; margin-bottom:12px;">' +
       '【朝の確認】' + dateStr + '</div>';

  h += sec_('今日やること', box.urgent, '#c62828', '#ffebee');
  h += sec_('今週中に', box.soon, '#ef6c00', '#fff3e0');
  h += sec_('手がすいたら', box.check, '#f9a825', '#fffde7');

  if (billTotal > 0) {
    h += '<div style="margin-top:20px; padding:10px; background:#e8eaf6; font-weight:bold;">' +
         '未請求の合計：' + billTotal.toLocaleString() + ' 円</div>';
  }

  h += '<div style="color:#888; font-size:12px; margin-top:16px;">' +
       'このメールは自動送信されています。</div></div>';
  return h;

  function sec_(title, groups, color, bg) {
    let t = '<div style="margin-top:18px; padding:6px 10px; background:' + bg +
            '; border-left:4px solid ' + color + '; font-weight:bold;">' +
            title + '（' + groups.length + '件）</div>';

    if (groups.length === 0) {
      return t + '<div style="padding:6px 14px; color:#888;">なし</div>';
    }

    t += '<div style="padding:4px 0;">';
    groups.forEach(function(g) {
      t += '<div style="padding:8px 14px; border-bottom:1px solid #eee;">';
      t += '<div><b>' + g.caseNo + '</b> ' + g.customer +
           (g.work ? '<span style="color:#666;">（' + g.work + '）</span>' : '') +
           '</div>';

      g.reasons.forEach(function(r) {
        const c = (r.rank === RANK.urgent) ? '#c62828'
                : (r.rank === RANK.soon) ? '#ef6c00' : '#f9a825';
        t += '<div style="padding-left:14px; font-size:13px; color:#555;">' +
             '<span style="color:' + c + ';">●</span> ' +
             '<b style="color:' + c + ';">' + r.label + '</b>' +
             (r.detail ? '　' + r.detail : '') + '</div>';
      });
      t += '</div>';
    });
    return t + '</div>';
  }
}
