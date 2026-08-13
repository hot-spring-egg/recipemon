/**
 * 開発要件:
 * - アプリ名: パシャッと変身！レシピモン
 * - DB: Google Spreadsheets (コンテナバインド)
 * - Storage: Google Drive (画像保存用)
 * - 機能: 料理ナビ、削除、編集更新、経験値、Gemini API、写真更新、カテゴリ検索
 * - 更新: ★通信クラッシュ(Uncaught js)対策。スプレッドシートの「日付」や「undefined」によるエラーを完全防止
 */

const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

const SHEET_RECIPE = 'Recipes';
const SHEET_USER = 'UserData';
const DRIVE_FOLDER_NAME = 'RecipeMon_Images';
const RECIPE_CATEGORIES = ["主食", "主菜", "副菜", "汁系", "スイーツ系", "ソース系", "その他"];

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

  let sheetUser = ss.getSheetByName(SHEET_USER);
  if (!sheetUser) {
    sheetUser = ss.insertSheet(SHEET_USER);
    sheetUser.appendRow(['key', 'value']); 
    sheetUser.appendRow(['current_xp', '0']);
    sheetUser.appendRow(['current_level', '1']);
  }
}

function doGet() {
  setupSheet();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('パシャッと変身！レシピモン')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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

// --- レベル計算ロジック ---
function getLevelThreshold(level) {
  return Number(level) * Number(level) * 300;
}

// --- ユーザーデータ操作 ---
function getUserStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
    if (fileId.startsWith("data:")) return fileId;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    return "data:" + blob.getContentType() + ";base64," + base64;
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
  let sheet = ss.getSheetByName(SHEET_RECIPE);
  const imageId = saveImageToDrive(recipeData.image);
  
  let isUpdate = false;
  let targetId = recipeData.id;
  const category = recipeData.category || "主菜";

  if (targetId) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == targetId) {
        sheet.getRange(i + 1, 2).setValue(recipeData.name);
        sheet.getRange(i + 1, 3).setValue(recipeData.servings);
        sheet.getRange(i + 1, 4).setValue(JSON.stringify(recipeData.ingredients));
        sheet.getRange(i + 1, 5).setValue(JSON.stringify(recipeData.steps));
        sheet.getRange(i + 1, 6).setValue(imageId);
        sheet.getRange(i + 1, 7).setValue(category);
        isUpdate = true;
        break;
      }
    }
  }

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
    const xpResult = addExperience(50);
    return { id: String(targetId), ...xpResult, isUpdate: false };
  }
  return { success: true, id: String(targetId), isUpdate: true };
}

function deleteRecipe(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RECIPE);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      const imageId = data[i][5];
      sheet.deleteRow(i + 1);
      if (imageId && !imageId.startsWith("data:")) {
        try { DriveApp.getFileById(imageId).setTrashed(true); } catch (e) {}
      }
      return { success: true };
    }
  }
  return { success: false };
}

function completeCooking() { return { success: true, ...addExperience(150) }; }
function completeQuest(questData) { return { success: true, ...addExperience(100) }; }

// --- リスト取得 ---
function getRecipeList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RECIPE);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const rows = data.slice(1).reverse();
  const list = [];
  
  for (const row of rows) {
    // ★重要: スプレッドシート上で「3-4」等がDate(日付)に変換されてしまうと
    // フロントに送る際にクラッシュ(Uncaught js)するため、すべて明示的にStringに変換します。
    const idStr = row[0] != null ? String(row[0]) : "";
    if (!idStr) continue;

    list.push({
      id: idStr,
      name: row[1] != null ? String(row[1]) : "",
      servings: row[2] != null ? String(row[2]) : "",  // 日付オブジェクトを強制的に文字列化
      ingredients: safeParseJSON(row[3] != null ? String(row[3]) : "", []),
      steps: safeParseJSON(row[4] != null ? String(row[4]) : "", []),
      imageId: row[5] != null ? String(row[5]) : "",
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
