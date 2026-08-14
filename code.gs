/**
 * 開発要件:
 * - アプリ名: パシャッと変身！レシピモン
 * - DB: Google Spreadsheets (コンテナバインド)
 * - Storage: Google Drive (画像保存用)
 * - 機能: 料理ナビ、削除、編集更新、経験値、Gemini API、写真更新、カテゴリ検索
 * - 更新: ★通信クラッシュ(Uncaught js)対策。スプレッドシートの「日付」や「undefined」によるエラーを完全防止
 * - 更新: ★高速化。CacheServiceによるキャッシュ、初期表示の通信を1回に集約、
 *         画像の再アップロード回避（同じ写真なら既存のDriveファイルを使い回す）
 */

const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

const SHEET_RECIPE = 'Recipes';
const SHEET_USER = 'UserData';
const DRIVE_FOLDER_NAME = 'RecipeMon_Images';
const RECIPE_CATEGORIES = ["主食", "主菜", "副菜", "汁系", "スイーツ系", "ソース系", "その他"];

// --- キャッシュ設定（高速化用） ---
const CACHE_KEY_LIST = 'recipe_list_v1';
const CACHE_KEY_SETUP = 'schema_ready_v1';
const CACHE_KEY_IMG_PREFIX = 'recipe_img_v1_';
const CACHE_TTL = 21600; // 6時間（CacheServiceの上限）
const CACHE_CHUNK_SIZE = 30000; // 1値あたり100KB制限に対し、日本語(3byte)でも収まるサイズ
const CACHE_MAX_CHUNKS = 200;

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheetRecipe = ss.getSheetByName(SHEET_RECIPE);
  if (!sheetRecipe) {
    sheetRecipe = ss.insertSheet(SHEET_RECIPE);
    sheetRecipe.appendRow(['id', 'name', 'servings', 'ingredients', 'steps', 'image', 'category', 'created_at']);
    sheetRecipe.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#FF9F1C').setFontColor('white');
  } else {
    const lastCol = sheetRecipe.getLastColumn();
    if (lastCol < 8) {
      sheetRecipe.insertColumnAfter(6);
      sheetRecipe.getRange(1, 7).setValue('category').setFontWeight('bold').setBackground('#FF9F1C').setFontColor('white');
    }
  }

  // 「3-4」等の人数が日付に自動変換されるのを防ぐ（表示崩れ・パース失敗の予防）
  try {
    sheetRecipe.getRange(2, 3, sheetRecipe.getMaxRows() - 1, 1).setNumberFormat('@');
  } catch (e) {
    console.error("servings列の書式設定に失敗: " + e);
  }

  let sheetUser = ss.getSheetByName(SHEET_USER);
  if (!sheetUser) {
    sheetUser = ss.insertSheet(SHEET_USER);
    sheetUser.appendRow(['key', 'value']); 
    sheetUser.appendRow(['current_xp', '0']);
    sheetUser.appendRow(['current_level', '1']);
  }
}

function doGet() {
  ensureSetup_();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('パシャッと変身！レシピモン')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- シート初期化（毎回やると遅いのでキャッシュで間引く） ---
function ensureSetup_() {
  try {
    const cache = CacheService.getScriptCache();
    if (cache.get(CACHE_KEY_SETUP)) return;
    setupSheet();
    cache.put(CACHE_KEY_SETUP, '1', CACHE_TTL);
  } catch (e) {
    console.error("ensureSetup_ error: " + e);
    setupSheet();
  }
}

// --- 安全なJSONパース関数（エラーによるクラッシュ防止） ---
function safeParseJSON(str, defaultVal) {
  if (!str) return defaultVal;
  try {
    return JSON.parse(String(str));
  } catch (e) {
    console.error("JSON Parse Error:", e);
    return defaultVal;
  }
}

// --- キャッシュヘルパー（1値100KB制限があるので分割して保存する） ---
function cachePut_(key, value, ttl) {
  try {
    if (value === null || value === undefined) return;
    const str = String(value);
    const chunks = [];
    for (let i = 0; i < str.length; i += CACHE_CHUNK_SIZE) {
      chunks.push(str.substring(i, i + CACHE_CHUNK_SIZE));
    }
    if (chunks.length === 0 || chunks.length > CACHE_MAX_CHUNKS) return; // 大きすぎるものは諦める
    const payload = {};
    chunks.forEach((c, i) => { payload[key + '_' + i] = c; });
    payload[key + '_meta'] = String(chunks.length);
    CacheService.getScriptCache().putAll(payload, ttl || CACHE_TTL);
  } catch (e) {
    console.error("cachePut_ error: " + e);
  }
}

function cacheGet_(key) {
  try {
    const cache = CacheService.getScriptCache();
    const meta = cache.get(key + '_meta');
    if (!meta) return null;
    const count = parseInt(meta, 10);
    if (!count || count < 1) return null;

    const keys = [];
    for (let i = 0; i < count; i++) keys.push(key + '_' + i);
    const parts = cache.getAll(keys);

    let result = '';
    for (let i = 0; i < count; i++) {
      const part = parts[key + '_' + i];
      if (part === null || part === undefined) return null; // 一部でも消えていたら無効
      result += part;
    }
    return result;
  } catch (e) {
    console.error("cacheGet_ error: " + e);
    return null;
  }
}

function cacheRemove_(key) {
  try {
    const cache = CacheService.getScriptCache();
    const meta = cache.get(key + '_meta');
    const keys = [key + '_meta'];
    if (meta) {
      const count = parseInt(meta, 10) || 0;
      for (let i = 0; i < count; i++) keys.push(key + '_' + i);
    }
    cache.removeAll(keys);
  } catch (e) {
    console.error("cacheRemove_ error: " + e);
  }
}

function invalidateRecipeCache_() {
  cacheRemove_(CACHE_KEY_LIST);
}

// --- レベル計算ロジック ---
function getLevelThreshold(level) {
  return Number(level) * Number(level) * 300;
}

// --- 初期表示用: ステータスとレシピ一覧をまとめて返す（通信回数を半分に） ---
// forceRefresh: スプレッドシートを直接編集した場合など、キャッシュを捨てて読み直す
function getInitialData(forceRefresh) {
  if (forceRefresh) invalidateRecipeCache_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    status: getUserStatus(ss),
    recipes: getRecipeList(ss)
  };
}

// --- ユーザーデータ操作 ---
function getUserStatus(ssOpt) {
  const ss = ssOpt || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USER);
  const data = sheet.getDataRange().getValues();

  let xp = 0;
  let level = 1;

  data.forEach(row => {
    if (row[0] === 'current_xp') xp = parseInt(row[1]) || 0;
    if (row[0] === 'current_level') level = parseInt(row[1]) || 1;
  });

  const currentLevelBaseXp = (level === 1) ? 0 : getLevelThreshold(level - 1);
  const nextLevelThreshold = getLevelThreshold(level);

  // ★ すべての値をNumberまたはStringに変換して返す（不正な型によるクラッシュ防止）
  return { 
    xp: Number(xp), 
    level: Number(level), 
    title: String(getRankTitle(level)),
    baseXp: Number(currentLevelBaseXp),
    nextXp: Number(nextLevelThreshold)
  };
}

function addExperience(amount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USER);
  const data = sheet.getDataRange().getValues();
  
  let currentXp = 0;
  let currentLevel = 1;
  let xpRowIndex = -1;
  let levelRowIndex = -1;

  data.forEach((row, i) => {
    if (row[0] === 'current_xp') { currentXp = parseInt(row[1]) || 0; xpRowIndex = i + 1; }
    if (row[0] === 'current_level') { currentLevel = parseInt(row[1]) || 1; levelRowIndex = i + 1; }
  });

  if (xpRowIndex === -1) { sheet.appendRow(['current_xp', 0]); xpRowIndex = sheet.getLastRow(); }
  if (levelRowIndex === -1) { sheet.appendRow(['current_level', 1]); levelRowIndex = sheet.getLastRow(); }

  let newXp = currentXp + amount;
  let newLevel = currentLevel;
  let isLevelUp = false;

  while (newLevel < 20 && newXp >= getLevelThreshold(newLevel)) {
    newLevel++;
    isLevelUp = true;
  }

  sheet.getRange(xpRowIndex, 2).setValue(newXp);
  if (isLevelUp) {
    sheet.getRange(levelRowIndex, 2).setValue(newLevel);
  }

  const currentLevelBaseXp = (newLevel === 1) ? 0 : getLevelThreshold(newLevel - 1);
  const nextLevelThreshold = getLevelThreshold(newLevel);

  return {
    success: true,
    isLevelUp: Boolean(isLevelUp),
    newLevel: Number(newLevel),
    newXp: Number(newXp),
    newTitle: String(getRankTitle(newLevel)),
    gainedXp: Number(amount),
    baseXp: Number(currentLevelBaseXp),
    nextXp: Number(nextLevelThreshold)
  };
}

function getRankTitle(level) {
  const titles = [
    "見習いコック🥚", "皿洗い見習い🧽", "下ごしらえ係🥔", "キッチンの新人🔰", "タマネギの涙💧",
    "火加減の勉強中🔥", "目玉焼き名人🍳", "家庭の料理番🏠", "包丁使いの達人🔪", "キッチン戦士🛡️",
    "街の定食屋さん🍚", "創作料理の開拓者🎨", "レシピハンター📜", "三ツ星シェフ⭐️", "味の魔術師🧙‍♂️",
    "鉄人予備軍🦾", "究極の味覚王👅", "伝説の料理人👑", "食卓の神様👼", "レシピモンマスター🦖"
  ];
  return titles[Math.min(level - 1, titles.length - 1)];
}

// --- 画像処理ヘルパー ---
function saveImageToDrive(base64Data) {
  if (!base64Data) return "";
  const match = base64Data.match(/^data:(.+);base64,(.+)$/);
  if (!match) return base64Data; 
  const contentType = match[1];
  const blobData = Utilities.base64Decode(match[2]);
  const blob = Utilities.newBlob(blobData, contentType, "recipe_img_" + Utilities.getUuid());
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  let folder;
  if (folders.hasNext()) folder = folders.next(); else folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  return folder.createFile(blob).getId(); 
}

function getImageFromDrive(fileId) {
  try {
    if (!fileId) return "";
    const id = String(fileId);
    if (id.indexOf("data:") === 0) return id;

    // 2回目以降はDriveへのアクセス＆Base64変換を丸ごとスキップ
    const cached = cacheGet_(CACHE_KEY_IMG_PREFIX + id);
    if (cached) return cached;

    const file = DriveApp.getFileById(id);
    const blob = file.getBlob();
    const base64 = "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
    cachePut_(CACHE_KEY_IMG_PREFIX + id, base64, CACHE_TTL);
    return base64;
  } catch (e) {
    console.error("画像読み込みエラー: " + e.toString());
    return "";
  }
}

// --- Gemini API ---
function analyzeImage(base64Image) {
  if (!GEMINI_API_KEY) return { error: "APIキー未設定", message: "スクリプトプロパティに GEMINI_API_KEY を設定してね！" };

  try {
    const match = base64Image.match(/^data:(.+);base64,(.+)$/);
    if (!match) throw new Error("画像データの形式が正しくありません");
    const mimeType = match[1];
    const rawBase64 = match[2];

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${GEMINI_API_KEY}`;
    
    const promptText = `
      あなたはプロの料理アシスタントです。提供された料理の画像（またはレシピのスクショ）を分析し、以下の情報をJSON形式で抽出してください。
      
      出力フォーマット:
      {
        "name": "料理名",
        "servings": "何人分（数字のみ）",
        "category": "カテゴリ（'主食', '主菜', '副菜', '汁系', 'スイーツ系', 'ソース系', 'その他' から1つ選択）",
        "ingredients": [ { "name": "材料名", "amount": "分量" } ],
        "steps": [ "手順1", "手順2" ]
      }
      
      ルール:
      - 言語は日本語。
      - カテゴリは料理の内容から推測し、必ず指定の7つの中から選ぶこと。
        - 主食: ご飯もの、麺類、パンなど
        - 主菜: 肉・魚・卵がメインのおかず
        - 副菜: 野菜、豆腐、海藻などの小さなおかず
        - 汁系: スープ、味噌汁
        - スイーツ系: お菓子、デザート
        - ソース系: ドレッシング、たれ、ジャム、ペーストなど
        - その他: 上記に当てはまらないもの
      - JSON以外のテキストは含めないこと。
    `;

    const payload = {
      contents: [{ parts: [{ text: promptText }, { inline_data: { mime_type: mimeType, data: rawBase64 } }] }],
      generationConfig: { response_mime_type: "application/json" }
    };
    const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = JSON.parse(response.getContentText());
    if (responseCode !== 200) throw new Error(`Gemini API Error: ${responseBody.error?.message || 'Unknown error'}`);

    const recipeData = JSON.parse(responseBody.candidates[0].content.parts[0].text);
    recipeData.image = base64Image;
    if (recipeData.ingredients && Array.isArray(recipeData.ingredients)) recipeData.ingredients.forEach(ing => ing.checked = false); else recipeData.ingredients = [];
    if (!recipeData.steps || !Array.isArray(recipeData.steps)) recipeData.steps = ["手順を読み取れませんでした。編集ボタンから追加してください。"];
    
    if (!RECIPE_CATEGORIES.includes(recipeData.category)) {
      recipeData.category = "主菜";
    }

    return recipeData;
  } catch (e) {
    console.error("Error in analyzeImage: " + e.toString());
    return { error: "解析失敗", message: "読み取りに失敗しちゃった💦", details: e.toString() };
  }
}

// --- レシピ保存 ---
function saveRecipe(recipeData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RECIPE);

  let isUpdate = false;
  let targetId = recipeData.id;
  const category = recipeData.category || "主菜";

  // ★写真が変わっていないときは既存のDriveファイルを使い回す（毎回アップロードすると激遅・容量も無駄）
  let imageId = recipeData.imageId ? String(recipeData.imageId) : "";
  if (recipeData.image) {
    const newId = saveImageToDrive(recipeData.image);
    if (newId) {
      const oldId = imageId;
      imageId = newId;
      if (oldId && oldId !== newId && oldId.indexOf("data:") !== 0) {
        cacheRemove_(CACHE_KEY_IMG_PREFIX + oldId);
        try { DriveApp.getFileById(oldId).setTrashed(true); } catch (e) {}
      }
      // 手元にある画像をそのままキャッシュしておく（次回の読み込みでDriveアクセスが不要になる）
      if (newId.indexOf("data:") !== 0) cachePut_(CACHE_KEY_IMG_PREFIX + newId, recipeData.image, CACHE_TTL);
    }
  }

  if (targetId) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == targetId) {
        // 画像の指定が無い場合は既存の値を維持する（消してしまわないように）
        if (!imageId && data[i][5]) imageId = String(data[i][5]);
        // 1行まとめて書き込む（セル単位のsetValueより速い）
        sheet.getRange(i + 1, 2, 1, 6).setValues([[
          recipeData.name,
          recipeData.servings,
          JSON.stringify(recipeData.ingredients),
          JSON.stringify(recipeData.steps),
          imageId,
          category
        ]]);
        isUpdate = true;
        break;
      }
    }
  }

  // 旧バージョンの埋め込み画像(data:)はIDとして返さない（シート側では維持される）
  const returnedImageId = (imageId && imageId.indexOf("data:") === 0) ? "" : String(imageId);

  if (!isUpdate) {
    targetId = Utilities.getUuid();
    const timestamp = new Date();
    sheet.appendRow([
      targetId,
      recipeData.name,
      recipeData.servings,
      JSON.stringify(recipeData.ingredients),
      JSON.stringify(recipeData.steps),
      imageId,
      category,
      timestamp
    ]);
    invalidateRecipeCache_();
    const xpResult = addExperience(50);
    return { id: String(targetId), imageId: returnedImageId, ...xpResult, isUpdate: false };
  }

  invalidateRecipeCache_();
  return { success: true, id: String(targetId), imageId: returnedImageId, isUpdate: true };
}

function deleteRecipe(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RECIPE);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      const imageId = data[i][5] != null ? String(data[i][5]) : "";
      sheet.deleteRow(i + 1);
      invalidateRecipeCache_();
      if (imageId && imageId.indexOf("data:") !== 0) {
        cacheRemove_(CACHE_KEY_IMG_PREFIX + imageId);
        try { DriveApp.getFileById(imageId).setTrashed(true); } catch (e) {}
      }
      return { success: true };
    }
  }
  return { success: false };
}

function completeCooking() { return { success: true, ...addExperience(150) }; }
function completeQuest(questData) { return { success: true, ...addExperience(100) }; }

// --- リスト取得（キャッシュ優先で高速化） ---
function getRecipeList(ssOpt) {
  const cached = cacheGet_(CACHE_KEY_LIST);
  if (cached) {
    const parsed = safeParseJSON(cached, null);
    if (parsed && Array.isArray(parsed)) return parsed;
  }

  const list = readRecipeList_(ssOpt);
  cachePut_(CACHE_KEY_LIST, JSON.stringify(list), CACHE_TTL);
  return list;
}

function readRecipeList_(ssOpt) {
  const ss = ssOpt || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RECIPE);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  // 必要な範囲だけ読む（getDataRangeより無駄が少ない）
  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const rows = data.reverse();
  const list = [];

  for (const row of rows) {
    // ★重要: スプレッドシート上で「3-4」等がDate(日付)に変換されてしまうと
    // フロントに送る際にクラッシュ(Uncaught js)するため、すべて明示的にStringに変換します。
    const idStr = row[0] != null ? String(row[0]) : "";
    if (!idStr) continue;

    const imageRaw = row[5] != null ? String(row[5]) : "";
    const isInlineImage = imageRaw.indexOf("data:") === 0; // 旧バージョンで直接埋め込まれた画像

    list.push({
      id: idStr,
      name: row[1] != null ? String(row[1]) : "",
      servings: row[2] != null ? String(row[2]) : "",  // 日付オブジェクトを強制的に文字列化
      ingredients: safeParseJSON(row[3] != null ? String(row[3]) : "", []),
      steps: safeParseJSON(row[4] != null ? String(row[4]) : "", []),
      imageId: isInlineImage ? "" : imageRaw,
      image: isInlineImage ? imageRaw : "",
      category: row[6] != null ? String(row[6]) : "主菜"
    });
  }

  return list;
}

function getRecipeImage(imageId) {
  return getImageFromDrive(imageId);
}

function _authorize() {
  try { UrlFetchApp.fetch("https://www.google.com"); DriveApp.getRootFolder(); console.log("✅ 外部アクセス権限の承認が完了しました！"); } catch (e) {}
}
