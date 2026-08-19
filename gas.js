// ============================================================
//  DRC - Google Apps Script
//  シート①: DRC         日々のトレード記録
//  シート②: TAG_MASTER  タグマスター
// ============================================================

// ── 設定 ──────────────────────────────────────────────────
const SPREADSHEET_ID  = 'YOUR_SPREADSHEET_ID';   // ★要設定
const DRIVE_FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID';  // ★要設定
const SHEET_DRC       = 'DRC';
const SHEET_TAG       = 'TAG_MASTER';
let _runtimeFolderId  = ''; // HTML側から動的に上書き可能

// ── エントリーポイント ─────────────────────────────────────

function doGet(e) {
  const action = e.parameter.action || '';
  try {
    if (action === 'ping')          return jsonResponse(ping());
    if (action === 'checkFolder')   return jsonResponse(checkFolder(e.parameter.folderId));
    if (action === 'getTagMaster')  return jsonResponse(getTagMaster());
    if (action === 'getDRC')        return jsonResponse(getDRC(e.parameter));
    if (action === 'getImages')      return jsonResponse(getImages(e.parameter));
    return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function doPost(e) {
  try {
    // FormData形式とJSON形式の両方に対応
    var body;
    if (e.parameters && e.parameters.payload) {
      body = JSON.parse(e.parameters.payload[0]);
    } else {
      body = JSON.parse(e.postData.contents);
    }
    var action = body.action || '';
    if (action === 'saveTagMaster') return jsonResponse(saveTagMaster(body.master));
    if (action === 'saveDRC')       return jsonResponse(saveDRC(body.data));
    if (action === 'saveImage')     return jsonResponse(saveImage(body));
    return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  接続確認・設定確認
// ============================================================

/**
 * 接続テスト用ping
 * 戻り値: { status:'ok', message:'connected', spreadsheetName, timestamp }
 */
function ping() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    status: 'ok',
    message: 'connected',
    spreadsheetName: ss.getName(),
    timestamp: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')
  };
}

/**
 * DriveフォルダIDの確認
 * 戻り値: { status:'ok', folderName, folderUrl }
 */
function checkFolder(folderId) {
  if (!folderId) return { status: 'error', message: 'フォルダIDが指定されていません' };
  try {
    const folder = DriveApp.getFolderById(folderId);
    return {
      status: 'ok',
      folderName: folder.getName(),
      folderUrl: folder.getUrl()
    };
  } catch(e) {
    return { status: 'error', message: 'フォルダが見つかりません。IDを確認してください。' };
  }
}

// ============================================================
//  TAG_MASTER
// ============================================================

/**
 * TAG_MASTERシートからタグマスターを読み込みJSONで返す
 * 戻り値: { status:'ok', master: [ {id, name, tags:[]} ] }
 */
function getTagMaster() {
  const sheet = getOrCreateSheet(SHEET_TAG);
  const rows  = sheet.getDataRange().getValues();

  // ヘッダー行がなければデフォルトを書き込んで返す
  if (rows.length <= 1 || (rows.length === 1 && rows[0][0] === '')) {
    const defaultMaster = getDefaultTagMaster();
    writeTagMaster(sheet, defaultMaster);
    return { status: 'ok', master: defaultMaster };
  }

  // 1行目はヘッダー（category_id / category_name / tag）
  const dataRows = rows.slice(1).filter(r => r[0] !== '');
  const map = {};
  const order = [];
  dataRows.forEach(([id, name, tag]) => {
    if (!map[id]) {
      map[id] = { id: String(id), name: String(name), tags: [] };
      order.push(String(id));
    }
    if (tag !== '') map[id].tags.push(String(tag));
  });
  const master = order.map(id => map[id]);
  return { status: 'ok', master };
}

/**
 * TAG_MASTERシートをmasterで上書き保存
 * master: [ {id, name, tags:[]} ]
 */
function saveTagMaster(master) {
  const sheet = getOrCreateSheet(SHEET_TAG);
  writeTagMaster(sheet, master);
  return { status: 'ok' };
}

function writeTagMaster(sheet, master) {
  sheet.clearContents();
  // ヘッダー
  sheet.appendRow(['category_id', 'category_name', 'tag']);
  // データ（1行1タグ）
  master.forEach(cat => {
    if (cat.tags.length === 0) {
      // タグがないカテゴリも行を残す
      sheet.appendRow([cat.id, cat.name, '']);
    } else {
      cat.tags.forEach(tag => {
        sheet.appendRow([cat.id, cat.name, tag]);
      });
    }
  });
}

function getDefaultTagMaster() {
  return [
    { id: 'emotion', name: '感情系',         tags: ['リベンジトレード','FOMO発動','冷静','焦り','自信過剰'] },
    { id: 'action',  name: '行動系',         tags: ['ルール遵守','ルール違反','損切り遅延','オーバーサイズ','早切り'] },
    { id: 'setup',   name: 'セットアップ系', tags: ['A+セットアップ','プレイブック通り','裁量','セットアップなし'] },
    { id: 'market',  name: '相場環境系',     tags: ['トレンド相場','レンジ相場','薄商い','イベント相場'] },
  ];
}

// ============================================================
//  DRC 保存
// ============================================================

/**
 * DRC1件を保存する
 * - 日付が既存なら上書き（重複防止）
 * - 画像はGoogleドライブに保存しURLをシートに記録
 */
function saveDRC(data) {
  // HTML側から渡されたfolderIdがあればそちらを優先
  if (data.folderId) {
    _runtimeFolderId = data.folderId;
  }
  const sheet = getOrCreateSheet(SHEET_DRC);
  ensureDRCHeader(sheet);

  // 既存行チェック（日付重複）
  const allValues = sheet.getDataRange().getValues();
  let existingRow = -1;
  for (let i = 1; i < allValues.length; i++) {
    const cellDate = Utilities.formatDate(new Date(allValues[i][0]), 'Asia/Tokyo', 'yyyy-MM-dd');
    if (cellDate === data.date) { existingRow = i + 1; break; }
  }

  // 画像をDriveに保存してURLを取得
  const imageUrls = saveTradeImages(data);

  // 行データを組み立て
  const row = buildDRCRow(data, imageUrls);

  if (existingRow > 0) {
    // 上書き
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    // 新規追加
    sheet.appendRow(row);
  }

  return { status: 'ok', message: existingRow > 0 ? '上書き保存しました' : '新規保存しました' };
}

/**
 * 画像1枚をDriveに保存してスプレッドシートのURL列を更新
 */
function saveImage(body) {
  var date     = body.date     || '';
  var tradeIdx = body.tradeIdx || 0;
  var base64   = body.base64   || '';
  var mimeType = body.type     || 'image/png';
  var name     = body.name     || 'chart.png';

  if (body.folderId) _runtimeFolderId = body.folderId;

  if (!date || !base64) return { status: 'error', message: '日付または画像データがありません' };

  // Driveに保存
  var folder = getOrCreateDRCFolder(date);
  var blob   = Utilities.newBlob(
    Utilities.base64Decode(base64),
    mimeType,
    'trade' + (tradeIdx+1) + '_' + name
  );
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var fileUrl = file.getUrl();

  // スプレッドシートの該当行・画像列にURLを追記
  var sheet   = getOrCreateSheet(SHEET_DRC);
  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0];
  var imgColNames = ['画像①','画像②','画像③','画像④','画像⑤'];
  var colIdx  = headers.indexOf(imgColNames[tradeIdx]);
  if (colIdx < 0) return { status: 'error', message: '画像列が見つかりません' };

  for (var i = 1; i < rows.length; i++) {
    var rawDate = rows[i][0];
    var rowDate = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(rawDate).substring(0, 10);
    if (rowDate === date) {
      var cell    = sheet.getRange(i+1, colIdx+1);
      var current = cell.getValue();
      cell.setValue(current ? current + '
' + fileUrl : fileUrl);
      break;
    }
  }

  return { status: 'ok', fileUrl: fileUrl };
}

/**
 * トレード画像をDriveに保存しURLを返す
 * 戻り値: [ [url1, url2, ...], [url1], ... ]  (トレード①〜⑤)
 */
function saveTradeImages(data) {
  const folder = getOrCreateDRCFolder(data.date);
  const result = [];

  const trades = data.trades || [];
  for (let i = 0; i < 5; i++) {
    const trade = trades[i] || {};
    const images = trade.images || [];
    const urls = [];
    images.forEach((img, j) => {
      if (!img.base64) return;
      try {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(img.base64),
          img.type || 'image/png',
          `trade${i+1}_chart${j+1}_${img.name || 'chart.png'}`
        );
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        urls.push(file.getUrl());
      } catch(err) {
        urls.push('ERROR: ' + err.message);
      }
    });
    result.push(urls.join('\n'));
  }
  return result;
}

/**
 * 日付フォルダを取得または作成
 * DRC/ → 2026-08/ → 2026-08-14/
 */
function getOrCreateDRCFolder(dateStr) {
  const fid  = _runtimeFolderId || DRIVE_FOLDER_ID;
  const root = DriveApp.getFolderById(fid);
  const ym     = dateStr.substring(0, 7); // "2026-08"
  const ymFolder = getOrCreateSubFolder(root, ym);
  return getOrCreateSubFolder(ymFolder, dateStr);
}

function getOrCreateSubFolder(parent, name) {
  const iter = parent.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parent.createFolder(name);
}

/**
 * DRCシートの1行を組み立て
 */
function buildDRCRow(data, imageUrls) {
  const trades = data.trades || [];
  const t = i => trades[i] || {};

  return [
    // 基本情報
    data.date,
    data.grade_overall || '',
    data.pnl           || 0,
    // コンディション
    data.mood          || '',
    data.sleep         || '',
    data.condition_note|| '',
    // FOMO
    data.fomo_anxiety  ? 'TRUE' : 'FALSE',
    data.fomo_sns      ? 'TRUE' : 'FALSE',
    data.fomo_regret   ? 'TRUE' : 'FALSE',
    data.fomo_urge     ? 'TRUE' : 'FALSE',
    data.fomo_breath   ? 'TRUE' : 'FALSE',
    data.fomo_note     || '',
    // プロセス目標
    data.goal          || '',
    // 前場
    data.am_grade      || '',
    data.am_setup      || '',
    data.am_playbook   || '',
    data.am_size       || '',
    data.am_comment    || '',
    // 後場
    data.pm_grade      || '',
    data.pm_setup      || '',
    data.pm_playbook   || '',
    data.pm_size       || '',
    data.pm_comment    || '',
    // トレード①〜⑤
    t(0).ticker || '', t(0).pnl || '', t(0).analysis || '', t(0).chartNote || '', t(0).tags || '', imageUrls[0] || '',
    t(1).ticker || '', t(1).pnl || '', t(1).analysis || '', t(1).chartNote || '', t(1).tags || '', imageUrls[1] || '',
    t(2).ticker || '', t(2).pnl || '', t(2).analysis || '', t(2).chartNote || '', t(2).tags || '', imageUrls[2] || '',
    t(3).ticker || '', t(3).pnl || '', t(3).analysis || '', t(3).chartNote || '', t(3).tags || '', imageUrls[3] || '',
    t(4).ticker || '', t(4).pnl || '', t(4).analysis || '', t(4).chartNote || '', t(4).tags || '', imageUrls[4] || '',
    // プレイバック
    data.insight       || '',
    data.good          || '',
    data.bad           || '',
    data.emotion       || '',
    data.action        || '',
    // マントラ
    data.mantra_today     || '',
    data.mantra_tomorrow  || '',
    // 保存日時
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
  ];
}

/**
 * DRCシートのヘッダーを初回のみ書き込む
 */
function ensureDRCHeader(sheet) {
  if (sheet.getLastRow() > 0) return;
  const headers = [
    // 基本
    '日付','総合評価','P&L',
    // コンディション
    '朝の気分','睡眠の質','コンディションメモ',
    // FOMO
    'FOMO_心配不安','FOMO_SNS','FOMO_手放した銘柄','FOMO_衝動','FOMO_深呼吸','FOMO自由記述',
    // 目標
    'プロセス目標',
    // 前場
    '前場評価','前場セットアップ','前場PB裁量','前場サイズ','前場コメント',
    // 後場
    '後場評価','後場セットアップ','後場PB裁量','後場サイズ','後場コメント',
    // トレード①〜⑤
    '銘柄①','損益①','分析①','チャート分析①','タグ①','画像①',
    '銘柄②','損益②','分析②','チャート分析②','タグ②','画像②',
    '銘柄③','損益③','分析③','チャート分析③','タグ③','画像③',
    '銘柄④','損益④','分析④','チャート分析④','タグ④','画像④',
    '銘柄⑤','損益⑤','分析⑤','チャート分析⑤','タグ⑤','画像⑤',
    // プレイバック
    '気づき','良かった点','悪かった点','感情の動き','今日から対応',
    // マントラ
    '今日の言葉','明日への言葉',
    // 管理
    '保存日時',
  ];
  sheet.appendRow(headers);
  // ヘッダー行を固定・装飾
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a1a2e').setFontColor('#c8a96e').setFontWeight('bold');
  sheet.setFrozenRows(1);
}

// ============================================================
//  DRC 検索・取得
// ============================================================

/**
 * DRCを検索して返す
 * パラメータ: dateFrom, dateTo, grade, pnlMin, pnlMax, tag, limit
 */
function getDRC(params) {
  const sheet = getOrCreateSheet(SHEET_DRC);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { status: 'ok', records: [] };

  const headers = rows[0];
  const dataRows = rows.slice(1);

  // フィルタ条件
  const dateFrom = params.dateFrom || '';
  const dateTo   = params.dateTo   || '';
  const grade    = params.grade    || '';   // カンマ区切り複数可
  const tag      = params.tag      || '';
  const pnlMin   = params.pnlMin   !== undefined ? Number(params.pnlMin) : null;
  const pnlMax   = params.pnlMax   !== undefined ? Number(params.pnlMax) : null;
  const limit    = Number(params.limit) || 100;

  const grades = grade ? grade.split(',').map(g => g.trim()) : [];

  const records = dataRows
    .filter(row => {
      // スプレッドシートの日付セルはDateオブジェクトの場合があるため文字列化
      var rawDate = row[0];
      var rowDate;
      if (rawDate instanceof Date) {
        rowDate = Utilities.formatDate(rawDate, 'Asia/Tokyo', 'yyyy-MM-dd');
      } else {
        rowDate = String(rawDate).substring(0, 10);
      }
      if (dateFrom && rowDate < dateFrom) return false;
      if (dateTo   && rowDate > dateTo)   return false;
      if (grades.length && !grades.includes(String(row[1]))) return false;
      const pnl = Number(row[2]);
      if (pnlMin !== null && pnl < pnlMin) return false;
      if (pnlMax !== null && pnl > pnlMax) return false;
      if (tag) {
        // トレード①〜⑤のタグ列をすべて検索
        const tagCols = [28, 34, 40, 46, 52]; // タグ①〜⑤の列インデックス（0始まり）
        const allTags = tagCols.map(c => String(row[c] || '')).join(',');
        if (!allTags.includes(tag)) return false;
      }
      return true;
    })
    .slice(-limit)
    .reverse(); // 新しい順

  // オブジェクト形式に変換（日付セルはyyyy-MM-dd文字列に統一）
  const result = records.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      var val = row[i];
      if (val instanceof Date) {
        obj[h] = Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd');
      } else {
        obj[h] = val;
      }
    });
    return obj;
  });

  return { status: 'ok', records: result };
}

// ============================================================
//  画像取得（Base64変換）
// ============================================================

/**
 * スプレッドシートの画像URL列からDriveファイルを読み込みBase64で返す
 * パラメータ: date（対象日付）
 * 戻り値: { status:'ok', images: [ {tradeIdx, imgIdx, base64, mimeType, name}, ... ] }
 */
function getImages(params) {
  var date = params.date || '';
  if (!date) return { status: 'error', message: '日付が指定されていません' };

  var sheet = getOrCreateSheet(SHEET_DRC);
  var rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { status: 'ok', images: [] };

  var headers = rows[0];

  // 対象行を探す
  var targetRow = null;
  for (var i = 1; i < rows.length; i++) {
    var rawDate = rows[i][0];
    var rowDate = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(rawDate).substring(0, 10);
    if (rowDate === date) { targetRow = rows[i]; break; }
  }
  if (!targetRow) return { status: 'ok', images: [] };

  // ヘッダーからオブジェクト化
  var obj = {};
  headers.forEach(function(h, i) { obj[h] = targetRow[i]; });

  var imageKeys = ['画像①','画像②','画像③','画像④','画像⑤'];
  var result = [];

  imageKeys.forEach(function(key, tradeIdx) {
    var urlStr = String(obj[key] || '');
    if (!urlStr) return;
    var urls = urlStr.split('
').map(function(u){ return u.trim(); }).filter(Boolean);
    urls.forEach(function(url, imgIdx) {
      try {
        var fileId = extractFileId(url);
        if (!fileId) return;
        var file = DriveApp.getFileById(fileId);
        var blob = file.getBlob();
        var base64 = Utilities.base64Encode(blob.getBytes());
        result.push({
          tradeIdx: tradeIdx,
          imgIdx:   imgIdx,
          base64:   base64,
          mimeType: blob.getContentType(),
          name:     file.getName()
        });
      } catch(e) {
        // ファイルアクセス失敗はスキップ
      }
    });
  });

  return { status: 'ok', images: result };
}

/**
 * Drive URLからファイルIDを抽出
 */
function extractFileId(url) {
  // 形式1: /file/d/FILE_ID/
  var m = url.match(/\/file\/d\/([^\/\?]+)/);
  if (m) return m[1];
  // 形式2: ?id=FILE_ID
  m = url.match(/[?&]id=([^&]+)/);
  if (m) return m[1];
  return null;
}

// ============================================================
//  ユーティリティ
// ============================================================

function getOrCreateSheet(name) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(name);
  return sheet || ss.insertSheet(name);
}
