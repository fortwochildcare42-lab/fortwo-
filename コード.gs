// コード.gs
// ※ パスワード・フォルダIDはスクリプトプロパティで管理（このファイルに秘密情報なし）

/**
 * @OnlyCurrentDoc
 */

// ========================================
// スクリプトプロパティから設定を読み込む
// ========================================

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    SPREADSHEET_ID: props.getProperty('SPREADSHEET_ID'),
    PASSWORDS: {
      [props.getProperty('PW_FORTWO')]: null,
      [props.getProperty('PW_SHOGU')]:  '小学生',
      [props.getProperty('PW_CHUGU')]:  '中学生',
      [props.getProperty('PW_KOGU')]:   '高校生',
    },
    FOLDERS: {
      '小学生': props.getProperty('FOLDER_SHOGAKUSEI'),
      '中学生': props.getProperty('FOLDER_CHUGAKUSEI'),
      '高校生': props.getProperty('FOLDER_KOKOKSEI'),
    }
  };
}

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('For Two English Garden 🇬🇧')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * パスワード検証（サーバー側で行うため、フロントに秘密情報は一切渡さない）
 * @param {string} rawInput - ユーザーが入力したパスワード
 * @returns {object} { success, isMaster, categories, currentCategory, error }
 */
function verifyPassword(rawInput) {
  if (!rawInput) {
    return { success: false, error: 'パスワードを入力してください' };
  }

  const config = getConfig();

  // 全角数字を半角に変換
  const normalized = String(rawInput).trim().replace(/[０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });

  // マスターキー判定
  const masterKey = Object.keys(config.PASSWORDS).find(function(k) {
    return config.PASSWORDS[k] === null;
  });

  if (normalized === masterKey) {
    return {
      success: true,
      isMaster: true,
      categories: Object.keys(config.FOLDERS),
      currentCategory: Object.keys(config.FOLDERS)[0]
    };
  }

  // 通常パスワード判定
  const matched = config.PASSWORDS[normalized];
  if (matched !== undefined) {
    return {
      success: true,
      isMaster: false,
      categories: [matched],
      currentCategory: matched
    };
  }

  return { success: false, error: 'パスワードが違います' };
}

/**
 * 教材一覧取得
 * @param {string} category - 取得するカテゴリ名（例: '小学生'）
 * @returns {Array} 教材データの配列
 */
function getMaterials(category) {
  const config = getConfig();
  var allData = [];

  // 1. Google DriveフォルダからPDFを取得
  var folderId = config.FOLDERS[category];
  if (folderId) {
    try {
      var folder = DriveApp.getFolderById(folderId);
      var files = folder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        allData.push({
          name: file.getName().replace(/\.pdf$/i, ''),  // 拡張子を表示から除去
          url: file.getUrl(),
          category: category,
          isWeb: false,
          updated: Utilities.formatDate(file.getLastUpdated(), 'JST', 'yyyy/MM/dd')
        });
      }
    } catch (e) {
      console.error('Drive Error [' + category + ']: ' + e);
    }
  }

  // 2. スプレッドシートからWeb教材を取得
  try {
    var ss = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('教材') || ss.getSheetByName('シート') || ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][0] || '').trim();
      var url  = String(data[i][1] || '').trim();
      var categoryRaw = String(data[i][2] || '').trim();
      if (!name || !url) continue;

      if (categoryRaw.includes(category)) {
        allData.push({
          name: name,
          url: url,
          category: category,
          isWeb: true,
          updated: 'Web教材'
        });
      }
    }
  } catch (e) {
    console.error('Sheet Error: ' + e);
  }

  // 名前順ソート
  allData.sort(function(a, b) {
    return (a.name || '').localeCompare(b.name || '', 'ja');
  });

  return allData;
}
