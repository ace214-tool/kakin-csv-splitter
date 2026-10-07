// ===== 課金CSV分割ツール用 シート取得 GAS（Webアプリとしてデプロイ） =====
// デプロイ設定：「次のユーザーとして実行」= 自分 ／「アクセスできるユーザー」= 全員
//
// 返すのは CONFIG.COLUMNS の4列だけ（金額・決済番号・案件名などは返さない）

const CONFIG = {
  SPREADSHEET_ID: '1-LQGmLhgBIPgb0eBeaZSs4h5fFJuk9DeEEW_Woa5hb0', // MEO請求管理
  HEADER_ROW: 1,                                            // 見出し行
  COLUMNS: ['顧客番号', '対象月', '課金日数', 'プラン'],     // 見出し名で列を探す（列の位置が変わってもOK）
  NUMBER_COLUMN: '課金日数',                                // 数値として返す列
  MONTH_SHEET: /^(\d{4})\.(\d{1,2})$/                       // 「2026.9」形式のシート名
};

// 例） …/exec             → 最新月シート
//      …/exec?sheet=2026.8 → 指定シート
function doGet(e) {
  let out;
  try {
    out = buildData(e && e.parameter && e.parameter.sheet);
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildData(sheetName) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const names = ss.getSheets().map(s => s.getName());

  // シート指定がなければ「YYYY.M」形式のうち一番新しい月を選ぶ
  let target = sheetName ? String(sheetName) : '';
  if (!target) {
    let bestKey = -1;
    names.forEach(n => {
      const m = n.match(CONFIG.MONTH_SHEET);
      if (m) {
        const key = Number(m[1]) * 100 + Number(m[2]);
        if (key > bestKey) { bestKey = key; target = n; }
      }
    });
    if (!target) throw new Error('「2026.9」形式の月シートが見つかりません');
  }

  const sh = ss.getSheetByName(target);
  if (!sh) throw new Error('シート「' + target + '」が見つかりません');

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= CONFIG.HEADER_ROW || lastCol < 1) throw new Error('シート「' + target + '」にデータがありません');

  // 見出し名（改行・空白を除いて比較）から列番号を探す
  const header = sh.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0]
    .map(h => String(h).replace(/[\r\n\s]/g, ''));
  const colIdx = CONFIG.COLUMNS.map(name => header.indexOf(name));
  const missing = CONFIG.COLUMNS.filter((name, i) => colIdx[i] < 0);
  if (missing.length) {
    throw new Error('シート「' + target + '」に必要な列がありません: ' + missing.join(' / '));
  }

  // 必要な4列だけ読む
  const n = lastRow - CONFIG.HEADER_ROW;
  const cols = CONFIG.COLUMNS.map((name, i) => {
    const range = sh.getRange(CONFIG.HEADER_ROW + 1, colIdx[i] + 1, n, 1);
    const disp = range.getDisplayValues();
    if (name !== CONFIG.NUMBER_COLUMN) return disp.map(r => String(r[0]).trim());
    const raw = range.getValues();
    return disp.map((r, k) => (typeof raw[k][0] === 'number') ? raw[k][0] : String(r[0]).trim());
  });

  const rows = [CONFIG.COLUMNS.slice()];
  for (let r = 0; r < n; r++) {
    const row = cols.map(c => c[r]);
    if (row.some(v => v !== '')) rows.push(row);
  }

  return {
    ok: true,
    spreadsheetName: ss.getName(),
    sheets: names,
    sheet: target,
    rows: rows,
    generatedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
  };
}
