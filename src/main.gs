/***********************
 * mt. inn 清掃管理 v4.0
 *  - 清掃指示シート生成
 *  - ステータス入力制御 & ログ
 *  - ステータス色分け（N〜Q）
 *  - スタッフマスタ & プルダウン
 *  - 部屋割りボード（10×4=40部屋）
 *  - 日時表示：MM/dd HH:mm
 *  - PWA（api.gs + pwa.html）
 *
 * ※ doGet は api.gs に移動（PWA HTML配信）
 ***********************/

/** メニュー追加 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('清掃管理')
    .addItem('清掃指示シート生成', 'generateCleaningSheet')
    .addItem('スタッフ初期設定', 'initializeStaffList')
    .addToUi();
}

/**
 * 清掃指示シート生成
 */
function generateCleaningSheet() {
  const FOLDER_NAME = 'neppan_csv';
  const tz = 'Asia/Tokyo';
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyyMMdd'); // 例: 20251118

  /***** 1) 今日の日付が入ったCSVファイルを探す *****/
  const folderIter = DriveApp.getFoldersByName(FOLDER_NAME);
  if (!folderIter.hasNext()) {
    throw new Error('フォルダ「' + FOLDER_NAME + '」が見つかりません。');
  }
  const folder = folderIter.next();
  const files = folder.getFiles();

  let todayFile = null;
  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName(); // 例: businessContactList-20251118-....
    const m = String(name).match(/(\d{8})/);
    if (!m) continue;
    const dateStr = m[1];
    if (dateStr === todayStr) {
      if (!todayFile || f.getLastUpdated().getTime() > todayFile.getLastUpdated().getTime()) {
        todayFile = f;
      }
    }
  }

  if (!todayFile) {
    throw new Error('今日の日付（' + todayStr + '）のCSVファイルがフォルダ「' + FOLDER_NAME + '」に見つかりません。');
  }

  /***** 2) CSVを読み込み（ねっぱん: Shift_JIS） *****/
  const blob = todayFile.getBlob();
  const rawCsvString = blob.getDataAsString('Shift_JIS');

  // CSVクリーニング: 不正な引用符を除去してparseCsvのエラーを防止
  const csvString = cleanCsvText(rawCsvString);

  let rows;
  try {
    rows = Utilities.parseCsv(csvString);
  } catch (e) {
    // parseCsv失敗時はフォールバック: 行単位で分割しカンマ区切り
    Logger.log('parseCsv失敗、フォールバックパーサー使用: ' + e.message);
    rows = fallbackParseCsv(csvString);
  }
  if (!rows || rows.length === 0) {
    throw new Error('CSVの中身が空です。');
  }

  /***** 3) ヘッダー解析 *****/
  const header = rows[0];
  Logger.log('ヘッダー行: ' + JSON.stringify(header));
  Logger.log('使用ファイル名: ' + todayFile.getName());

  const roomIdx   = findColumnIndex(header, ['部屋', '部屋名', '客室', 'ルーム', 'Room', '部屋番号']);
  const nameIdx   = findColumnIndex(header, ['利用者氏名', '氏名', 'お名前', '名前', '宿泊者', '代表者']);
  const planIdx   = findColumnIndex(header, ['商品プラン', 'プラン', '商品名', 'プラン名', 'Plan']);
  const paxIdx    = findColumnIndex(header, ['人数', '人員', '名様', '総人数', '合計人数']);
  const adultIdx  = findColumnIndex(header, ['大人', 'おとな', '成人', '大人人数']);
  const childIdx  = findColumnIndex(header, ['子供', '子ども', 'こども', '小人', '児童', 'お子様']);
  const nightsIdx = findColumnIndex(header, ['泊数']);
  const ciIdx     = findColumnIndex(header, ['チェックイン日', 'チェックイン', 'IN', '到着日']);
  const coIdx     = findColumnIndex(header, ['チェックアウト日', 'チェックアウト', 'OUT', '出発日']);
  const stayIdx   = findColumnIndex(header, ['利用日', '宿泊日', '対象日']);

  if (roomIdx === -1 || nameIdx === -1) {
    throw new Error('部屋／氏名の列が特定できませんでした。ヘッダー: ' + JSON.stringify(header));
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('清掃指示');
  if (!sheet) {
    sheet = ss.insertSheet('清掃指示');
  } else {
    sheet.clear();
  }

  // 条件付き書式は全削除（色はスクリプトで制御）
  sheet.setConditionalFormatRules([]);

  /***** 4) ヘッダー行（A〜Q） *****/
  const outHeaders = [
    '泊数（例：1/3）',      // A
    '部屋番号',              // B
    '氏名',                  // C
    'プラン名',              // D
    '人数',                  // E
    '大人',                  // F
    '子供',                  // G
    '本日CI(H)',             // H
    '連泊(I)',               // I
    '3日清掃(J)',            // J
    '今朝CO(K)',             // K
    'フル清掃(L)',           // L
    'アメニティのみ(M)',      // M
    'ステータス(N)',          // N（在室/清掃含む）
    '清掃担当(O)',            // O
    '点検担当(P)',            // P
    'ステータス更新日時(Q)'   // Q
  ];
  sheet.getRange(1, 1, 1, outHeaders.length).setValues([outHeaders]);

  const todayMidnight = getTodayTokyoMidnight();

  const out   = [];
  const bgA   = [];
  const bgCtoG= [];
  const bgB   = [];

  /***** 5) データ行生成 *****/
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length === 0) continue;

    const rawRoom   = row[roomIdx] || '';
    if (!rawRoom) continue;

    const guestName = row[nameIdx] || '';
    const planName  = (planIdx   >= 0) ? (row[planIdx]   || '') : '';
    const pax       = (paxIdx    >= 0) ? (row[paxIdx]    || '') : '';
    const adults    = (adultIdx  >= 0) ? (row[adultIdx]  || '') : '';
    const childs    = (childIdx  >= 0) ? (row[childIdx]  || '') : '';
    const ciStr     = (ciIdx     >= 0) ? (row[ciIdx]     || '') : '';
    const coStr     = (coIdx     >= 0) ? (row[coIdx]     || '') : '';
    const stayStr   = (stayIdx   >= 0) ? (row[stayIdx]   || '') : '';

    const roomNo = extractRoomNumber(rawRoom);

    // --- A列（泊数）をX列（24列目、インデックス23）から直接取得して「1 | 1」を「1/1」に変換 ---
    const X_COLUMN_INDEX = 23; // X列 = 24列目 = インデックス23（0から始まるため）
    let nightsDisplay = '';
    if (row.length > X_COLUMN_INDEX) {
      const xColumnValue = String(row[X_COLUMN_INDEX] || '').trim();
      if (xColumnValue) {
        // 「1 | 1」を「1/1」に変換
        nightsDisplay = xColumnValue.replace(/\s*\|\s*/g, '/');
      }
    }

    // 清掃ロジック用にn/dを抽出
    let { n, d } = parseNightInfo(nightsDisplay);

    // --- 清掃ロジック ---
    const isCheckin       = (n === 1);
    const isStay          = (n >= 2);
    const isThirdCleaning = (isStay && n % 3 === 0);
    const isCheckoutToday = (d > 0 && n === d);
    const isFullCleaning  = isThirdCleaning || isCheckoutToday;
    const isAmenityOnly   = false;

    const line = [
      nightsDisplay,    // A
      roomNo,           // B
      guestName,        // C
      planName,         // D
      pax,              // E
      adults,           // F
      childs,           // G
      isCheckin,        // H
      isStay,           // I
      isThirdCleaning,  // J
      isCheckoutToday,  // K
      isFullCleaning,   // L
      isAmenityOnly,    // M
      '',               // N ステータス
      '',               // O 清掃担当
      '',               // P 点検担当
      ''                // Q 更新日時
    ];
    out.push(line);

    // --- 背景色決定（A, C〜G, B列ピンク） ---
    const is314 = (roomNo === '314');

    let shade = null;
    if (!is314) {
      if (isCheckin) {
        shade = '#CCFFCC'; // 黄緑
      } else if (isThirdCleaning) {
        shade = '#CCFFFF'; // 水色（3泊ごとの連泊清掃日のみ）
      }
    }
    bgA.push([shade]);

    const rowBgCtoG = [];
    for (let i = 0; i < 5; i++) rowBgCtoG.push(shade);
    bgCtoG.push(rowBgCtoG);

    let bColor = null;
    if (!is314) {
      if (isFullCleaning) {
        bColor = '#FFCCCC'; // ピンク
      } else if (isAmenityOnly) {
        bColor = '#E0CCFF'; // 将来用
      }
    }
    bgB.push([bColor]);
  }

  if (out.length === 0) return;

  /***** 6) 清掃指示シートへ書き込み *****/
  sheet.getRange(2, 1, out.length, out[0].length).setValues(out);
  sheet.getRange(2, 2, out.length, 1).setNumberFormat('@'); // B列=部屋番号を文字列に

  // 背景色（従来ロジック）
  sheet.getRange(2, 1, out.length, 1).setBackgrounds(bgA);        // A
  sheet.getRange(2, 3, out.length, 5).setBackgrounds(bgCtoG);     // C〜G
  sheet.getRange(2, 2, out.length, 1).setBackgrounds(bgB);        // B

  // 列幅
  const widths = [80, 60, 120, 220, 60, 60, 60, 70, 70, 80, 80, 80, 110, 120, 120, 120, 150];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // H〜M列を非表示（ロジック列）
  sheet.hideColumns(8, 6); // 8列目(H)から6列分→H〜M

  // Q列 日付形式（日時：MM/dd HH:mm）
  sheet.getRange(2, 17, out.length, 1).setNumberFormat('MM/dd HH:mm');

  // ステータス列(N)にプルダウンを設定
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['在室', 'チェックアウト', '連泊鍵預け中', '清掃中', '清掃完了', '清掃点検済'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 14, out.length, 1).setDataValidation(statusRule);

  // スタッフ一覧シートを用意して、O,P列にプルダウン
  ensureStaffSheetAndValidation(ss, out.length);

  // ステータス列の背景を一旦クリア（白）
  sheet.getRange(2, 14, out.length, 4).setBackground(null);

  // 部屋割りボードを更新
  updateRoomBoardSheet(ss);
}

/********************************
 * onEdit: ステータス入力制御 & ログ + 色付け + ボード更新
 ********************************/
function onEdit(e) {
  const sheet = e.range.getSheet();
  const ss = sheet.getParent();
  if (sheet.getName() !== '清掃指示') return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row === 1) return; // ヘッダー行は無視

  const STATUS_COL  = 14; // N列 ステータス
  const CLEANER_COL = 15; // O列 清掃担当
  const CHECKER_COL = 16; // P列 点検担当
  const TIME_COL    = 17; // Q列

  if (![STATUS_COL, CLEANER_COL, CHECKER_COL].includes(col)) return;

  const status  = sheet.getRange(row, STATUS_COL).getValue();
  const cleaner = sheet.getRange(row, CLEANER_COL).getValue();
  const checker = sheet.getRange(row, CHECKER_COL).getValue();

  // 清掃完了 → 清掃担当必須
  if (status === '清掃完了' && !cleaner) {
    if (col === STATUS_COL || col === CLEANER_COL) {
      e.range.setValue(e.oldValue || '');
    }
    ss.toast('ステータスを「清掃完了」にするには、O列に清掃担当者名を入力してください。', '清掃ステータス', 5);
    updateStatusRowColor(sheet, row); // 色戻し
    updateRoomBoardSheet(ss);
    return;
  }

  // 清掃点検済 → 点検担当必須
  if (status === '清掃点検済' && !checker) {
    if (col === STATUS_COL || col === CHECKER_COL) {
      e.range.setValue(e.oldValue || '');
    }
    ss.toast('ステータスを「清掃点検済」にするには、P列に点検担当者名を入力してください。', '清掃ステータス', 5);
    updateStatusRowColor(sheet, row);
    updateRoomBoardSheet(ss);
    return;
  }

  // 要件を満たした場合は日時記録＆ログ
  const needLog = (status === '清掃完了' || status === '清掃点検済');
  if (needLog) {
    const now = new Date();
    sheet.getRange(row, TIME_COL).setValue(now);
    sheet.getRange(row, TIME_COL).setNumberFormat('MM/dd HH:mm');

    const roomNo = sheet.getRange(row, 2).getValue(); // B列
    const guest  = sheet.getRange(row, 3).getValue(); // C列

    logCleaningStatus(ss, now, roomNo, guest, status, cleaner, checker);
  }

  // ステータス色付け（N〜Q）
  updateStatusRowColor(sheet, row);

  // 部屋割りボードを更新
  updateRoomBoardSheet(ss);
}

/****************************************
 * N〜Q列のステータス色付け
 ****************************************/
function updateStatusRowColor(sheet, row) {
  const STATUS_COL = 14; // N
  const RANGE_COLS = 4;  // N〜Q
  const status = sheet.getRange(row, STATUS_COL).getValue();

  let color = null;
  if (status === '在室')             color = '#CCFFCC';
  else if (status === 'チェックアウト') color = '#DDDDDD';
  else if (status === '連泊鍵預け中')   color = '#CCE5FF';
  else if (status === '清掃中')       color = '#FFF2CC';
  else if (status === '清掃完了')     color = '#FFD699';
  else if (status === '清掃点検済')   color = '#99E699';
  else color = null; // クリア

  sheet.getRange(row, STATUS_COL, 1, RANGE_COLS).setBackground(color);
}

/****************************************
 * スタッフ一覧シート & プルダウン設定
 ****************************************/
function ensureStaffSheetAndValidation(ss, dataRowCount) {
  let staffSheet = ss.getSheetByName('スタッフ一覧');
  if (!staffSheet) {
    staffSheet = ss.insertSheet('スタッフ一覧');
    staffSheet.getRange(1, 1).setValue('清掃担当');
    staffSheet.getRange(1, 2).setValue('点検担当');
    // 名前入力用の行を少し確保（空のままでOK）
    staffSheet.getRange(2, 1, 30, 2).clearContent();
  }

  const sheet = ss.getSheetByName('清掃指示');
  if (!sheet) return;

  const cleanerRange = staffSheet.getRange('A2:A100');
  const checkerRange = staffSheet.getRange('B2:B100');

  const cleanerRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(cleanerRange, true)
    .setAllowInvalid(false)
    .build();
  const checkerRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(checkerRange, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, 15, dataRowCount, 1).setDataValidation(cleanerRule); // O列
  sheet.getRange(2, 16, dataRowCount, 1).setDataValidation(checkerRule); // P列
}

/****************************************
 * 部屋割りボード更新（10×4=40部屋）
 ****************************************/
function updateRoomBoardSheet(ss) {
  const boardName = '部屋割り';
  let board = ss.getSheetByName(boardName);
  if (!board) {
    board = ss.insertSheet(boardName);
  } else {
    board.clear();
  }

  const sheet = ss.getSheetByName('清掃指示');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // データなし

  const dataRows = lastRow - 1;
  const statusRange = sheet.getRange(2, 14, dataRows, 1).getValues(); // N列
  const roomRange   = sheet.getRange(2, 2,  dataRows, 1).getValues(); // B列（部屋番号）

  // 部屋→ステータスの辞書化（文字列キー）
  const roomStatusMap = {};
  for (let i = 0; i < dataRows; i++) {
    const room = roomRange[i][0];
    const st   = statusRange[i][0];
    if (room) {
      const key = String(room);
      roomStatusMap[key] = st;
    }
  }

  // レイアウト（10×4=40部屋）
  // 3F前半: 301〜310 （部屋番号行:2 / ステータス行:3）
  // 3F後半: 311〜320 （部屋番号行:4 / ステータス行:5）
  // 4F前半: 401〜410 （部屋番号行:6 / ステータス行:7）
  // 4F後半: 411〜420 （部屋番号行:8 / ステータス行:9）
  const layout = [
    { startRow: 2, rooms: range(301, 310) },
    { startRow: 4, rooms: range(311, 320) },
    { startRow: 6, rooms: range(401, 410) },
    { startRow: 8, rooms: range(411, 420) },
  ];

  // タイトル
  board.getRange('A1').setValue('部屋割り（ステータスボード）');

  // 表の作成：部屋番号 & ステータス文字
  layout.forEach(layer => {
    layer.rooms.forEach((room, idx) => {
      const roomCol = 2 + idx; // B〜K
      const roomKey = String(room);
      const status  = roomStatusMap[roomKey] || '';

      // 部屋番号
      board.getRange(layer.startRow, roomCol).setValue(room);

      // ステータス（部屋番号の1行下）
      board.getRange(layer.startRow + 1, roomCol).setValue(status);
    });
  });

  // 見た目調整（スマホ視認性重視）
  board.setColumnWidths(2, 10, 80);   // B〜K
  board.setRowHeights(2, 1, 24);     // 3F前半 部屋番号
  board.setRowHeights(3, 1, 40);     // 3F前半 ステータス
  board.setRowHeights(4, 1, 24);     // 3F後半 部屋番号
  board.setRowHeights(5, 1, 40);     // 3F後半 ステータス
  board.setRowHeights(6, 1, 24);     // 4F前半 部屋番号
  board.setRowHeights(7, 1, 40);     // 4F前半 ステータス
  board.setRowHeights(8, 1, 24);     // 4F後半 部屋番号
  board.setRowHeights(9, 1, 40);     // 4F後半 ステータス

  board.getRange('B2:K9').setHorizontalAlignment('center');
  board.getRange('B2:K9').setVerticalAlignment('middle');
  board.getRange('B2:K9').setFontSize(14);
  board.getRange('B3:K3').setFontSize(16);
  board.getRange('B5:K5').setFontSize(16);
  board.getRange('B7:K7').setFontSize(16);
  board.getRange('B9:K9').setFontSize(16);

  board.getRange('B3:K3').setWrap(true);
  board.getRange('B5:K5').setWrap(true);
  board.getRange('B7:K7').setWrap(true);
  board.getRange('B9:K9').setWrap(true);

  // ステータス色塗り（GASで直接）
  const colorMap = {
    '在室':        '#CCFFCC',
    'チェックアウト': '#DDDDDD',
    '連泊鍵預け中':   '#CCE5FF',
    '清掃中':      '#FFF2CC',
    '清掃完了':    '#FFD699',
    '清掃点検済':  '#99E699',
  };

  layout.forEach(layer => {
    layer.rooms.forEach((room, idx) => {
      const roomCol = 2 + idx;
      const roomKey = String(room);
      const status  = roomStatusMap[roomKey];
      const color   = colorMap[status] || null;
      board.getRange(layer.startRow + 1, roomCol).setBackground(color);
      });
    });
}

/****************************************
 * 清掃ログシートに1行追加（日時フォーマット）
 ****************************************/
function logCleaningStatus(ss, datetime, roomNo, guest, status, cleaner, checker) {
  let logSheet = ss.getSheetByName('清掃ログ');
  if (!logSheet) {
    logSheet = ss.insertSheet('清掃ログ');
    logSheet.getRange(1, 1, 1, 6).setValues([['日時', '部屋番号', '氏名', 'ステータス', '清掃担当', '点検担当']]);
    logSheet.getRange('A:A').setNumberFormat('MM/dd HH:mm');
  }
  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow + 1, 1, 1, 6).setValues([[
    datetime,
    roomNo,
    guest,
    status,
    cleaner,
    checker
  ]]);
  logSheet.getRange(lastRow + 1, 1).setNumberFormat('MM/dd HH:mm');
}

/********************************
 * 共通ユーティリティ
 ********************************/

/** ヘッダーから候補文字を含む列を探す */
function findColumnIndex(headerRow, keywords) {
  for (let i = 0; i < headerRow.length; i++) {
    const h = String(headerRow[i] || '');
    for (let k = 0; k < keywords.length; k++) {
      if (h.indexOf(keywords[k]) !== -1) return i;
    }
  }
  return -1;
}

/** 部屋番号から先頭3桁を取り出す（301S6 → 301） */
function extractRoomNumber(raw) {
  const m = String(raw).match(/\d{3}/);
  return m ? m[0] : String(raw);
}

/** 泊数 "n｜d" または "n/d" を {n, d} に変換 */
function parseNightInfo(str) {
  const s = String(str || '').trim();
  if (!s) return { n: 0, d: 0 };
  // 「1 | 1」形式を「1/1」に変換（念のため、ここでも変換）
  const normalized = s.replace(/\s*\|\s*/g, '/');
  const m = normalized.match(/(\d+)\D+(\d+)/);
  if (m) {
    const n = parseInt(m[1], 10);
    const d = parseInt(m[2], 10);
    return { n: isNaN(n) ? 0 : n, d: isNaN(d) ? 0 : d };
  }
  const digits = normalized.replace(/[^\d]/g, '');
  if (digits) {
    const only = parseInt(digits, 10);
    if (!isNaN(only) && only > 0) return { n: 1, d: only };
  }
  return { n: 0, d: 0 };
}

/** "yyyy/MM/dd" などを Date(0時) にパース（失敗時 null） */
function parseDateSafe(str) {
  const s = String(str || '').trim();
  if (!s) return null;
  const normalized = s.replace(/[.]/g, '/').replace(/-/g, '/');
  const d = new Date(normalized);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/** 今日(Asia/Tokyo)の0時 */
function getTodayTokyoMidnight() {
  const tz = 'Asia/Tokyo';
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd');
  return parseDateSafe(todayStr);
}

/** 日数差（to - from） */
function calcNights(fromDate, toDate) {
  const msPerDay = 24 * 60 * 60 * 1000;
  return Math.round((toDate.getTime() - fromDate.getTime()) / msPerDay);
}

/** 数列生成ユーティリティ */
function range(a, b) {
  return Array.from({ length: b - a + 1 }, (_, i) => a + i);
}

/**
 * CSVテキストのクリーニング
 * - 不正なダブルクォート（閉じていない引用符）を除去
 * - BOMを除去
 * - NULLバイトを除去
 */
function cleanCsvText(text) {
  if (!text) return '';

  // BOM除去
  let cleaned = text.replace(/^\uFEFF/, '');

  // NULLバイト除去
  cleaned = cleaned.replace(/\0/g, '');

  // 各行ごとにダブルクォートの数をチェックし、奇数なら末尾の引用符を除去
  const lines = cleaned.split('\n');
  const fixedLines = lines.map(function(line) {
    const quoteCount = (line.match(/"/g) || []).length;
    if (quoteCount % 2 !== 0) {
      // 奇数 = 閉じていない引用符 → 全てのダブルクォートを除去
      return line.replace(/"/g, '');
    }
    return line;
  });

  return fixedLines.join('\n');
}

/**
 * フォールバックCSVパーサー
 * Utilities.parseCsv()が失敗した場合に使用
 */
function fallbackParseCsv(text) {
  if (!text) return [];
  const lines = text.split(/\r?\n/);
  const result = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    // シンプルなカンマ分割（ダブルクォート内のカンマは非対応だが緊急用）
    const cells = line.split(',').map(function(cell) {
      return cell.replace(/^"/, '').replace(/"$/, '').trim();
    });
    result.push(cells);
  }
  return result;
}
