/***********************
 * mt. inn 清掃管理 PWA API
 * - doGet: PWA HTML配信
 * - getRoomsData: 部屋一覧取得
 * - getStaffData: スタッフ一覧取得
 * - updateRoomStatus: ステータス更新
 * - getSummaryData: 進捗サマリー取得
 ***********************/

/** スプレッドシートID */
var SS_ID = '1q0EPF2Uuhbb215ziy61aoVe7yntDmBjfplYYa5XYgyc';

/** スプレッドシート取得（Web App用にopenById） */
function getSpreadsheet_() {
  return SpreadsheetApp.openById(SS_ID);
}

/**
 * Web App エントリーポイント
 * PWAのHTMLを配信する
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('pwa')
    .setTitle('mt. inn 清掃管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 部屋一覧データ取得
 * 清掃指示シートから全部屋のデータを返す
 */
function getRoomsData() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('清掃指示');
  if (!sheet) return { status: 'error', message: '清掃指示シートが見つかりません。朝のシート生成を実行してください。' };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'ok', rooms: [] };

  var data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
  var rooms = [];
  var seenRooms = {};

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var roomNo = String(row[1]).trim();

    // 3桁の部屋番号のみ対象（レストラン等を除外）
    if (!roomNo.match(/^\d{3}$/)) continue;

    // 同じ部屋番号の重複行はスキップ（最初の行を採用）
    if (seenRooms[roomNo]) continue;
    seenRooms[roomNo] = true;

    var updatedAt = '';
    if (row[16]) {
      try {
        updatedAt = Utilities.formatDate(new Date(row[16]), 'Asia/Tokyo', 'MM/dd HH:mm');
      } catch (ex) {
        updatedAt = String(row[16]);
      }
    }

    rooms.push({
      row: i + 2,
      nights: String(row[0] || ''),
      roomNo: roomNo,
      guest: String(row[2] || ''),
      plan: String(row[3] || ''),
      pax: String(row[4] || ''),
      adults: String(row[5] || ''),
      children: String(row[6] || ''),
      isCheckin: Boolean(row[7]),
      isStay: Boolean(row[8]),
      isThirdCleaning: Boolean(row[9]),
      isCheckout: Boolean(row[10]),
      isFullCleaning: Boolean(row[11]),
      isAmenityOnly: Boolean(row[12]),
      status: String(row[13] || ''),
      cleaner: String(row[14] || ''),
      checker: String(row[15] || ''),
      updatedAt: updatedAt
    });
  }

  return { status: 'ok', rooms: rooms };
}

/**
 * スタッフ一覧取得
 * スタッフ一覧シートから清掃担当・点検担当の名前を返す
 */
function getStaffData() {
  var ss = getSpreadsheet_();
  var staffSheet = ss.getSheetByName('スタッフ一覧');
  if (!staffSheet) return { status: 'ok', cleaners: [], checkers: [] };

  var lastRow = staffSheet.getLastRow();
  if (lastRow < 2) return { status: 'ok', cleaners: [], checkers: [] };

  var data = staffSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var cleaners = [];
  var checkers = [];

  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) cleaners.push(String(data[i][0]).trim());
    if (data[i][1]) checkers.push(String(data[i][1]).trim());
  }

  return { status: 'ok', cleaners: cleaners, checkers: checkers };
}

/**
 * スタッフ一覧を初期化
 * メニューから実行してスタッフ名を一括登録する
 * ※ スプレッドシートで直接追加・削除も可能
 */
function initializeStaffList() {
  var STAFF_NAMES = [
    '村松 俊伊',
    '三本木 伸一',
    '安田 高志',
    '日下部 愛菜',
    '高山 博登',
    '長谷川 栄夫',
    '遠藤 良男',
    '國島 宏之',
    '渡辺 直美',
    '山内 啓子',
    '古川 タミ子',
    '大内 みきこ',
    '三瓶 ヒナ子',
    '村山 久美子',
    '渡部 祝子',
    '橋本 孝',
    '遠藤 詩織',
    '芦野 由美子',
    '三浦 勝幸',
    '渡邊 光子',
    '佐原 ゆい',
    '佐原 あかね',
    '渡辺 幸子',
    '本間 嬉美',
    '吉村 亮太',
    '大橋 福子',
    '渡辺 和子'
  ];

  var ss = getSpreadsheet_();
  var staffSheet = ss.getSheetByName('スタッフ一覧');
  if (!staffSheet) {
    staffSheet = ss.insertSheet('スタッフ一覧');
  }

  // ヘッダー
  staffSheet.getRange(1, 1).setValue('清掃担当');
  staffSheet.getRange(1, 2).setValue('点検担当');

  // 既存データをクリア
  if (staffSheet.getLastRow() > 1) {
    staffSheet.getRange(2, 1, staffSheet.getLastRow() - 1, 2).clearContent();
  }

  // 全スタッフを清掃担当・点検担当の両方に登録
  var values = STAFF_NAMES.map(function(name) { return [name, name]; });
  staffSheet.getRange(2, 1, values.length, 2).setValues(values);

  // 列幅調整
  staffSheet.setColumnWidth(1, 150);
  staffSheet.setColumnWidth(2, 150);

  // プルダウンも再設定
  var cleaningSheet = ss.getSheetByName('清掃指示');
  if (cleaningSheet && cleaningSheet.getLastRow() > 1) {
    var dataRowCount = cleaningSheet.getLastRow() - 1;
    ensureStaffSheetAndValidation(ss, dataRowCount);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    STAFF_NAMES.length + '名のスタッフを登録しました。\nスプレッドシートで追加・削除できます。',
    'スタッフ初期設定完了', 5
  );
}

/**
 * ステータス更新
 * PWAからのステータス変更をスプレッドシートに反映
 * 既存のonEditと同じ処理（色更新、ログ、ボード更新）を実行
 */
function updateRoomStatus(params) {
  var row = params.row;
  var roomNo = params.roomNo;
  var newStatus = params.newStatus;
  var cleaner = params.cleaner || '';
  var checker = params.checker || '';

  // バリデーション
  if (newStatus === '清掃完了' && !cleaner) {
    return { status: 'error', message: '「清掃完了」にするには清掃担当の入力が必要です' };
  }
  if (newStatus === '清掃点検済' && !checker) {
    return { status: 'error', message: '「清掃点検済」にするには点検担当の入力が必要です' };
  }

  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('清掃指示');
  if (!sheet) return { status: 'error', message: '清掃指示シートが見つかりません' };

  // 行の部屋番号を確認（行ずれ防止）
  var currentRoom = String(sheet.getRange(row, 2).getValue()).trim();
  if (currentRoom !== roomNo) {
    return { status: 'error', message: '部屋番号が一致しません（' + currentRoom + ' ≠ ' + roomNo + '）。ページを更新してください。' };
  }

  // ステータス書き込み
  sheet.getRange(row, 14).setValue(newStatus);

  // 担当者書き込み
  if (cleaner) sheet.getRange(row, 15).setValue(cleaner);
  if (checker) sheet.getRange(row, 16).setValue(checker);

  // タイムスタンプ
  var now = new Date();
  sheet.getRange(row, 17).setValue(now);
  sheet.getRange(row, 17).setNumberFormat('MM/dd HH:mm');

  // ステータス色更新（既存関数を再利用）
  updateStatusRowColor(sheet, row);

  // 清掃ログ記録
  if (newStatus === '清掃完了' || newStatus === '清掃点検済') {
    var guest = sheet.getRange(row, 3).getValue();
    logCleaningStatus(ss, now, roomNo, guest, newStatus, cleaner, checker);
  }

  // 部屋割りボード更新（既存関数を再利用）
  updateRoomBoardSheet(ss);

  return {
    status: 'ok',
    message: roomNo + '号室を「' + newStatus + '」に更新しました',
    updatedAt: Utilities.formatDate(now, 'Asia/Tokyo', 'MM/dd HH:mm')
  };
}

/**
 * 進捗サマリー取得
 */
function getSummaryData() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('清掃指示');
  if (!sheet) return { status: 'error', message: '清掃指示シートが見つかりません' };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'ok', total: 0 };

  var statuses = sheet.getRange(2, 14, lastRow - 1, 1).getValues();
  var roomNos = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var checkouts = sheet.getRange(2, 11, lastRow - 1, 1).getValues();

  var total = 0, occupied = 0, checkout = 0, stayKey = 0;
  var cleaning = 0, cleaned = 0, inspected = 0, noStatus = 0;
  var needsCleaning = 0;

  for (var i = 0; i < roomNos.length; i++) {
    var rn = String(roomNos[i][0]).trim();
    if (!rn.match(/^\d{3}$/)) continue;
    total++;

    var st = String(statuses[i][0]);
    if (st === '在室') occupied++;
    else if (st === 'チェックアウト') { checkout++; needsCleaning++; }
    else if (st === '連泊鍵預け中') { stayKey++; needsCleaning++; }
    else if (st === '清掃中') cleaning++;
    else if (st === '清掃完了') cleaned++;
    else if (st === '清掃点検済') inspected++;
    else noStatus++;
  }

  return {
    status: 'ok',
    total: total,
    occupied: occupied,
    checkout: checkout,
    stayKey: stayKey,
    cleaning: cleaning,
    cleaned: cleaned,
    inspected: inspected,
    noStatus: noStatus,
    needsCleaning: needsCleaning
  };
}
