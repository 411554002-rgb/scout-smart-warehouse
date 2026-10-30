/**
 * =============================================================================
 * 童軍倉庫智慧物聯網管理系統 (Scout Smart Warehouse & Asset Management)
 * 後端核心：Google Apps Script (GAS)
 * 
 * 整合功能：
 * 1. LINE Bot 雙向互動：自然語言查庫存、智慧器材推薦、快捷借出/歸還更動試算表
 * 2. Google 試算表即時連動：自動計算借出數/現存可用數、狀態警告更新、新增借還紀錄
 * 3. ESP32 (Wokwi) 終端 API：支援硬體終端透過 HTTP Webhook 同步更新試算表
 * 4. Google Gemini AI 智慧大腦：當非指令查詢時，回答童軍器材使用指引與露營規劃
 * 
 * 安全提醒：
 * 依專案繳交規定，金鑰與網址切勿直接寫死於程式碼中！
 * 請在 Apps Script 專案設定 ->「指令碼屬性 (Script Properties)」中設定：
 * - LINE_CHANNEL_ACCESS_TOKEN : LINE Messaging API 長效 Token
 * - GEMINI_API_KEY            : Google AI Studio 申請的 Gemini API Key
 * - SPREADSHEET_ID            : 試算表 ID (例如: 1narSASC-lVSBgzPSDh_T18MuwizGU7UtDv_GOXfzT4M)
 * =============================================================================
 */

// 讀取指令碼屬性（環境變數）
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const LINE_CHANNEL_ACCESS_TOKEN = SCRIPT_PROPERTIES.getProperty("LINE_CHANNEL_ACCESS_TOKEN");
const GEMINI_API_KEY = SCRIPT_PROPERTIES.getProperty("GEMINI_API_KEY");
// 若未在屬性設定，則預設使用當前試算表
const TARGET_SPREADSHEET_ID = SCRIPT_PROPERTIES.getProperty("SPREADSHEET_ID") || "1narSASC-lVSBgzPSDh_T18MuwizGU7UtDv_GOXfzT4M";

// 可用 Gemini 模型
const CANDIDATE_MODELS = ["gemini-1.5-flash", "gemini-2.0-flash"];

/**
 * 取得目標試算表物件
 */
function getTargetSpreadsheet() {
  if (TARGET_SPREADSHEET_ID) {
    try {
      return SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    } catch (e) {
      console.warn("無法透過 ID 開啟試算表，嘗試取得當前容器試算表: " + e.toString());
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * 取得庫存主工作表（自動辨識包含器材標題的工作表）
 */
function getInventorySheet() {
  const ss = getTargetSpreadsheet();
  if (!ss) return null;
  
  // 優先尋找名稱包含「器材」或「庫存」或第一張表
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const sName = sheets[i].getName();
    if (sName.indexOf("器材") !== -1 || sName.indexOf("庫存") !== -1 || sName.indexOf("清單") !== -1) {
      return sheets[i];
    }
  }
  return sheets[0];
}

/**
 * 取得或建立「借還紀錄」工作表
 */
function getLogSheet() {
  const ss = getTargetSpreadsheet();
  if (!ss) return null;
  let logSheet = ss.getSheetByName("借還紀錄");
  if (!logSheet) {
    logSheet = ss.getSheetByName("表單回應 1") || ss.getSheetByName("表單回應");
  }
  if (!logSheet) {
    logSheet = ss.insertSheet("借還紀錄");
    logSheet.appendRow(["時間戳記", "借用人 / 小隊名稱", "動作類型", "選擇器材或組合包/櫃位", "數量", "備註"]);
  }
  return logSheet;
}

/**
 * 處理 GET 請求：可用於瀏覽器無痕測試、健康檢查或 ESP32 HTTP 快速測試
 * 支援格式：?action=query&code=A1-2 或 ?action=borrow&code=A1-2&user=Scout
 */
function doGet(e) {
  // 純文字直接回傳，確保任何外部連線都能獲得標準 200 回應
  if (!e || !e.parameter || Object.keys(e.parameter).length === 0) {
    return ContentService.createTextOutput("童軍倉庫雲端管理 API 運作中 (Ready)！");
  }

  const action = e.parameter.action;
  const code = e.parameter.code || e.parameter.item;
  const user = e.parameter.user || "Wokwi-ESP32";
  const qty = parseInt(e.parameter.qty || "1", 10);

  if (action === "query" && code) {
    const result = searchInventory(code);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } else if ((action === "borrow" || action === "return") && code) {
    const isBorrow = (action === "borrow");
    const result = handleInventoryChange(code, isBorrow, qty, user);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput("童軍倉庫雲端管理 API 運作正常 (Ready)");
}

/**
 * 處理 POST 請求（LINE Webhook 或 ESP32 JSON 回傳）
 */
function doPost(e) {
  try {
    // ⭐️ 關鍵：LINE Developers 按「Verify」時，通常是空 payload 或空 events
    // 必須秒回純文字 "OK"，不能觸發任何需要尚未授權之試算表的操作！
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput("OK");
    }

    let postData;
    try {
      postData = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return ContentService.createTextOutput("OK");
    }

    // 1. 若為 ESP32 的自訂 API 請求
    if (postData.source === "ESP32_TERMINAL") {
      const code = postData.code;
      const isBorrow = (postData.action === "borrow");
      const user = postData.user || "實體終端機操作";
      const res = handleInventoryChange(code, isBorrow, 1, user);
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. 若為 LINE Messaging API Webhook
    const events = postData.events;
    // ⭐️ 關鍵：Verify 測試時 events 通常是空陣列 []，立即回傳 OK，確保 HTTP 200！
    if (!events || events.length === 0) {
      return ContentService.createTextOutput("OK");
    }

    for (let i = 0; i < events.length; i++) {
      const event = events[i];
      if (event.type === "message" && event.message.type === "text") {
        const replyToken = event.replyToken;
        const userMsg = event.message.text.trim();
        const userId = event.source.userId;
        const userName = getUserDisplayName(userId) || "童軍夥伴";

        // 處理使用者訊息（指令或問答）
        const replyText = processUserMessage(userMsg, userName, userId);
        replyLineMessage(replyToken, replyText);
      }
    }

    return ContentService.createTextOutput("OK");

  } catch (error) {
    console.error("doPost 處理異常: " + error.toString());
    // 即使發生錯誤也回傳 200 OK 給 LINE，避免 LINE 平台判斷 Webhook 伺服器掛掉
    return ContentService.createTextOutput("OK");
  }
}

/**
 * ⭐️ 專用授權觸發函式：請在 Apps Script 手動選擇此函式執行一次！
 * 本函式強制呼叫試算表與網路連線，100% 觸發 Google 彈出完整授權視窗！
 */
function testAuthorizeAll() {
  console.log("正在檢測專案權限...");
  const ss = getTargetSpreadsheet();
  if (ss) {
    console.log("✅ 成功存取試算表: " + ss.getName());
  } else {
    console.warn("⚠️ 未能開啟試算表，請檢查 SPREADSHEET_ID。");
  }
  
  // 測試連線外部網路（確保 UrlFetchApp 獲得授權）
  try {
    const testFetch = UrlFetchApp.fetch("https://www.google.com", { muteHttpExceptions: true });
    console.log("✅ 外部網路請求權限 (UrlFetchApp) 正常！代碼: " + testFetch.getResponseCode());
  } catch (netErr) {
    console.warn("網路測試失敗: " + netErr.toString());
  }
  
  console.log("🎉 所有必要權限已全部順利授權完畢！");
}

/**
 * 核心訊息邏輯：判斷是借還指令、庫存查詢、還是 AI 顧問問答
 */
function processUserMessage(text, userName, userId) {
  const normText = text.trim();

  // 1. 說明指令
  if (normText === "/help" || normText === "說明" || normText === "功能" || normText === "選單") {
    return "⛺ **【童軍倉庫智慧管家】使用指引**\n\n" +
           "📦 **【查詢庫存與器材】**\n" +
           "• 查 虎頭鉗 (查單項品名)\n" +
           "• 查 繩 / 查 爐 / 查 刀 (關鍵字模糊搜尋)\n" +
           "• 查 A1-2 (查詢某特定箱子裡的所有器材)\n" +
           "• A區清單 / B區清單 / C區清單 / D區清單 (看分區所有物品)\n" +
           "• 所有器材 (查看庫存統計總覽)\n\n" +
           "🔄 **【登記借出 / 歸還】**\n" +
           "• 借 虎頭鉗  /  借 A1-2\n" +
           "• 還 虎頭鉗  /  還 A1-2\n" +
           "• 借出 營釘 5 (指定數量)\n\n" +
           "⚠️ **【特別查詢】**\n" +
           "• 缺貨清單 (查看誰被借完了)\n" +
           "• 直接詢問露營、童軍技能，AI 將提供專業建議！";
  }

  // 2. 查看所有物品 / 總覽
  if (normText === "所有器材" || normText === "全部器材" || normText === "器材清單" || normText === "庫存總覽") {
    return getAllInventorySummary();
  }

  // 3. 查看特定分區 (例如: A區、A區清單、B櫃、C區)
  const zoneMatch = normText.match(/^([A-Da-d])(?:區|櫃)?(?:清單)?$/);
  if (zoneMatch) {
    const zoneChar = zoneMatch[1].toUpperCase();
    return getZoneInventory(zoneChar);
  }

  // 4. 缺貨清單查詢
  if (normText === "缺貨清單" || normText === "誰被借完了" || normText === "已全數借出") {
    return getOutOfStockSummary();
  }

  // 5. 借出指令匹配 (支援：借 虎頭鉗、借出 A1-2、借 營釘 5)
  const borrowMatch = normText.match(/^(借|借出|領取)\s*([A-Za-z0-9_\-\u4e00-\u9fa5]+)(?:\s+(\d+))?$/);
  if (borrowMatch) {
    const itemKeyword = borrowMatch[2];
    const qty = borrowMatch[3] ? parseInt(borrowMatch[3], 10) : 1;
    const res = handleInventoryChange(itemKeyword, true, qty, userName);
    return res.message;
  }

  // 6. 歸還指令匹配 (支援：還 虎頭鉗、歸還 A1-2、還 營釘 5)
  const returnMatch = normText.match(/^(還|歸還|放回)\s*([A-Za-z0-9_\-\u4e00-\u9fa5]+)(?:\s+(\d+))?$/);
  if (returnMatch) {
    const itemKeyword = returnMatch[2];
    const qty = returnMatch[3] ? parseInt(returnMatch[3], 10) : 1;
    const res = handleInventoryChange(itemKeyword, false, qty, userName);
    return res.message;
  }

  // 7. 庫存查詢指令匹配 (支援：查 虎頭鉗、查詢 A1-2、A1-2、庫存 營燈)
  const queryMatch = normText.match(/^(查|查詢|庫存|找)\s*([A-Za-z0-9_\-\u4e00-\u9fa5]+)$/);
  if (queryMatch) {
    const keyword = queryMatch[2];
    return formatQueryResult(searchInventory(keyword));
  }

  // 8. 若直接輸入櫃位格式 (如 A1-2, B3-4, a1-2)
  if (/^[A-Za-z]\d+-\d+$/i.test(normText) || /^[A-Za-z]\d+$/i.test(normText)) {
    return formatQueryResult(searchInventory(normText));
  }

  // 9. 其餘訊息交付 Google Gemini 智慧解答 (結合童軍庫存情境)
  return askGeminiWithContext(userId, userName, normText);
}

/**
 * 統計全庫房總覽與分類摘要
 */
function getAllInventorySummary() {
  const sheet = getInventorySheet();
  if (!sheet) return "找不到庫存表。";
  const data = sheet.getDataRange().getValues();

  let totalItems = 0;
  let totalAvailableQty = 0;
  let totalLentQty = 0;
  let zoneCount = { "A": 0, "B": 0, "C": 0, "D": 0, "其他": 0 };

  for (let r = 1; r < data.length; r++) {
    const name = (data[r][2] || "").toString().trim();
    const loc = (data[r][1] || "").toString().trim();
    if (!name && !loc) continue;

    totalItems++;
    const firstChar = loc.charAt(0).toUpperCase();
    if (zoneCount.hasOwnProperty(firstChar)) {
      zoneCount[firstChar]++;
    } else {
      zoneCount["其他"]++;
    }
  }

  return `📊 **【童軍庫房器材總體現況】**\n\n` +
         `• 建檔器材品項總數：**${totalItems}** 項\n` +
         `────────────\n` +
         `🏢 **【各分區器材數量】**\n` +
         `• **A 櫃 (五金與繩索工具)**：${zoneCount["A"]} 項\n` +
         `• **B 櫃 (炊事與露營裝備)**：${zoneCount["B"]} 項\n` +
         `• **C 櫃 (廚房與餐飲食器)**：${zoneCount["C"]} 項\n` +
         `• **D 櫃 (活動材料與工藝品)**：${zoneCount["D"]} 項\n` +
         `────────────\n` +
         `💡 **如何查看各區所有物品？**\n` +
         `請直接在對話框輸入：\n` +
         `👉 **A區** (列出A櫃所有裝備)\n` +
         `👉 **B區** (列出B櫃所有裝備)\n` +
         `👉 **C區** 或 **D區**`;
}

/**
 * 列出特定分區的所有器材
 */
function getZoneInventory(zoneChar) {
  const sheet = getInventorySheet();
  if (!sheet) return "找不到庫存表。";
  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let r = 1; r < data.length; r++) {
    const loc = (data[r][1] || "").toString().trim().toUpperCase();
    const name = (data[r][2] || "").toString().trim();
    if (!name) continue;

    if (loc.startsWith(zoneChar)) {
      let qty = data[r][3] || 1;
      let unit = data[r][4] || "個";
      list.push(`• \`${loc}\` ${name} (${qty}${unit})`);
    }
  }

  if (list.length === 0) {
    return `🔍 ${zoneChar} 區目前尚無建檔器材。`;
  }

  const showList = list.slice(0, 30);
  let reply = `📦 **【${zoneChar} 區器材清單】(共 ${list.length} 項)**\n\n` + showList.join("\n");
  if (list.length > 30) {
    reply += `\n\n*(清單較長，已顯示前 30 項。可輸入特定櫃位如「${zoneChar}1-2」查看特定箱子)*`;
  }
  return reply;
}

/**
 * 試算表智慧搜尋器材（支援多種表頭格式與自動尋找表頭列）
 */
function searchInventory(keyword) {
  const sheet = getInventorySheet();
  if (!sheet) return { success: false, message: "找不到庫存工作表。" };

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: false, message: "試算表無資料。" };

  const normKey = keyword.toString().trim().toUpperCase();
  const matched = [];

  // 1. 動態尋找表頭列 (可能在第 1 列，也可能在中間)
  let headerRowIdx = 0;
  let idCol = -1, locCol = -1, nameCol = -1, totalCol = -1, lentCol = -1, availCol = -1, unitCol = -1, statusCol = -1;

  for (let r = 0; r < Math.min(data.length, 30); r++) {
    const rowStr = data[r].join(",");
    // 優先找完整表頭 (包含「現存可用數」或「器材唯一ID」)
    if (rowStr.indexOf("現存可用數") !== -1 || rowStr.indexOf("器材唯一ID") !== -1 || rowStr.indexOf("建檔總數量") !== -1) {
      headerRowIdx = r;
      break;
    }
    // 次選：簡易表頭 (包含「編號」與「名稱」)
    if (nameCol === -1 && (rowStr.indexOf("名稱") !== -1 || rowStr.indexOf("器材") !== -1)) {
      headerRowIdx = r;
    }
  }

  const header = data[headerRowIdx];
  for (let c = 0; c < header.length; c++) {
    const title = (header[c] || "").toString().trim();
    if (title.indexOf("唯一ID") !== -1) idCol = c;
    else if (title.indexOf("櫃位") !== -1 || title === "編號") locCol = c;
    else if (title.indexOf("名稱") !== -1 || title === "器材") nameCol = c;
    else if (title.indexOf("建檔總數量") !== -1 || title === "總數量" || title === "數量") totalCol = c;
    else if (title.indexOf("借出數") !== -1) lentCol = c;
    else if (title.indexOf("可用數") !== -1) availCol = c;
    else if (title.indexOf("單位") !== -1) unitCol = c;
    else if (title.indexOf("狀態") !== -1) statusCol = c;
  }

  // 預設欄位容錯 (若為簡易表格，自動分配借出數與可用數欄位)
  if (locCol === -1) locCol = 1;
  if (nameCol === -1) nameCol = 2;
  if (totalCol === -1) totalCol = 3;
  if (unitCol === -1) unitCol = 4;
  if (lentCol === -1) lentCol = 5;
  if (availCol === -1) availCol = 6;
  if (statusCol === -1) statusCol = 7;

  for (let r = headerRowIdx + 1; r < data.length; r++) {
    const row = data[r];
    const id = idCol !== -1 ? (row[idCol] || "").toString().trim() : "";
    const loc = (row[locCol] || "").toString().trim();
    const name = (row[nameCol] || "").toString().trim();

    if (!name && !loc) continue;

    // 比對 ID、櫃位或器材名稱
    if (id.toUpperCase().indexOf(normKey) !== -1 ||
        loc.toUpperCase().indexOf(normKey) !== -1 ||
        name.indexOf(keyword.trim()) !== -1) {
      
      let total = parseInt(row[totalCol], 10);
      if (isNaN(total)) total = 1;

      let rawLent = (lentCol !== -1) ? row[lentCol] : null;
      let lent = 0;
      if (rawLent !== null && rawLent !== "" && typeof rawLent !== "boolean" && !isNaN(parseInt(rawLent, 10))) {
        lent = parseInt(rawLent, 10);
      }

      let rawAvail = (availCol !== -1) ? row[availCol] : null;
      let available;
      if (rawAvail !== null && rawAvail !== "" && typeof rawAvail !== "boolean" && !isNaN(parseInt(rawAvail, 10))) {
        available = parseInt(rawAvail, 10);
      } else {
        available = Math.max(0, total - lent);
      }

      let unit = (unitCol !== -1 && row[unitCol] && typeof row[unitCol] === "string") ? row[unitCol].trim() : "個";
      let status = (statusCol !== -1 && row[statusCol] && typeof row[statusCol] === "string") ? row[statusCol].trim() : (available > 0 ? "🟢 正常" : "🔴 已全數借出");

      matched.push({
        rowIdx: r + 1,
        idCol: idCol,
        locCol: locCol,
        totalCol: totalCol,
        lentCol: lentCol,
        availCol: availCol,
        statusCol: statusCol,
        id: id,
        location: loc,
        name: name,
        total: total,
        lent: lent,
        available: available,
        unit: unit,
        status: status
      });
    }
  }

  return { success: true, keyword: keyword, count: matched.length, items: matched };
}

/**
 * 格式化查詢文字回傳
 */
function formatQueryResult(searchResult) {
  if (!searchResult.success) return searchResult.message;
  if (searchResult.count === 0) {
    return `🔍 查無與「${searchResult.keyword}」相關的器材。\n請確認品名或櫃位編號（如 A1-2、營燈、瓦斯爐）。`;
  }

  let text = `📦【器材庫存查詢結果】(共 ${searchResult.count} 筆)\n`;
  const maxShow = Math.min(searchResult.count, 5);

  for (let i = 0; i < maxShow; i++) {
    const it = searchResult.items[i];
    text += `\n🏷️ ${it.name}\n` +
            `• 櫃位編號：${it.location}\n` +
            `• 現存可用：${it.available} ${it.unit} (總量 ${it.total})\n` +
            `• 借出狀態：${it.status}\n` +
            `────────────`;
  }

  if (searchResult.count > 5) {
    text += `\n還有 ${searchResult.count - 5} 筆未列出，請縮小關鍵字精確查詢。`;
  }
  return text;
}

/**
 * 處理借出 / 歸還更動試算表的核心業務邏輯
 */
function handleInventoryChange(keyword, isBorrow, qty, user) {
  const searchRes = searchInventory(keyword);
  if (!searchRes.success || searchRes.count === 0) {
    return {
      success: false,
      message: `❌ 找不到器材「${keyword}」，無法進行${isBorrow ? "借出" : "歸還"}登記。`
    };
  }

  // 取得第一筆符合的器材
  const item = searchRes.items[0];
  const sheet = getInventorySheet();
  const rowIdx = item.rowIdx;

  let currentTotal = item.total;
  let currentLent = item.lent;
  let currentAvail = item.available;

  let newLent = currentLent;
  let newAvail = currentAvail;

  if (isBorrow) {
    if (currentAvail < qty) {
      return {
        success: false,
        message: `🔴【無法借出】\n器材「${item.name}」現存可用數僅剩 ${currentAvail} ${item.unit}，不足欲借出的 ${qty} ${item.unit}。`
      };
    }
    newLent += qty;
    newAvail -= qty;
  } else {
    // 歸還
    if (currentLent < qty && currentLent === 0) {
      return {
        success: false,
        message: `⚠️【提醒】器材「${item.name}」目前在庫量為滿額 (${currentTotal} ${item.unit})，無人借出，無需重複歸還。`
      };
    }
    newLent = Math.max(0, currentLent - qty);
    newAvail = Math.min(currentTotal, currentAvail + qty);
  }

  // ⭐️ 核心寫入更新：若有專屬借出/可用欄位則更新；若只有單一「數量」欄位，則直接更新「數量」！
  if (item.availCol !== -1) {
    sheet.getRange(rowIdx, item.availCol + 1).setValue(newAvail);
  } else if (item.totalCol !== -1) {
    // 試算表只有單一「數量」欄位時，借出直接扣減該數量！
    sheet.getRange(rowIdx, item.totalCol + 1).setValue(newAvail);
  }

  if (item.lentCol !== -1) {
    sheet.getRange(rowIdx, item.lentCol + 1).setValue(newLent);
  }
  
  const newStatus = (newAvail <= 0) ? "🔴 已全數借出" : "🟢 正常";
  if (item.statusCol !== -1) {
    sheet.getRange(rowIdx, item.statusCol + 1).setValue(newStatus);
  }

  // 強制立即將所有修改同步寫入 Google 試算表實體儲存！
  SpreadsheetApp.flush();

  // 寫入「借還紀錄」工作表
  const logSheet = getLogSheet();
  if (logSheet) {
    const timestamp = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
    const actionText = isBorrow ? "借出" : "歸還";
    logSheet.appendRow([timestamp, user, actionText, `${item.location}_${item.name}`, qty, "LINE Bot 自動登記"]);
    SpreadsheetApp.flush();
  }

  const actionMsg = isBorrow ? "借出成功 ✅" : "歸還成功 🎉";
  return {
    success: true,
    action: isBorrow ? "borrow" : "return",
    itemName: item.name,
    location: item.location,
    avail: newAvail,
    lent: newLent,
    message: `📋【${actionMsg}】\n\n` +
             `• 器材名稱：${item.name}\n` +
             `• 櫃位位置：${item.location}\n` +
             `• 登記人員：${user}\n` +
             `• 異動數量：${qty} ${item.unit}\n` +
             `• 最新可用：${newAvail} ${item.unit} (借出中: ${newLent})\n` +
             `• 狀態：${newStatus}`
  };
}

/**
 * 統計全數借出清單
 */
function getOutOfStockSummary() {
  const sheet = getInventorySheet();
  if (!sheet) return "找不到庫存表。";
  const data = sheet.getDataRange().getValues();
  const outList = [];

  for (let r = 1; r < data.length; r++) {
    const name = data[r][2];
    const loc = data[r][1];
    const avail = data[r][5];
    const status = (data[r][7] || "").toString();

    if (name && (avail === 0 || status.indexOf("借出") !== -1)) {
      outList.push(`• [${loc}] ${name}`);
    }
  }

  if (outList.length === 0) {
    return "🟢 **太棒了！目前所有童軍器材皆在庫存中，沒有缺貨！**";
  }
  return `🔴 **【目前全數借出的器材】(共 ${outList.length} 項)**\n\n` + outList.slice(0, 15).join("\n") +
         (outList.length > 15 ? "\n...(其餘省略)" : "");
}

/**
 * 結合童軍情境與 Gemini AI 智慧顧問
 */
function askGeminiWithContext(userId, userName, userMessage) {
  if (!GEMINI_API_KEY) {
    return "⚠️ 尚未在指令碼屬性中設定 GEMINI_API_KEY。";
  }

  const promptContext = `你是童軍團器材裝備庫的智慧管家 AI。
提問者是童軍夥伴：${userName}。
如果使用者是在諮詢露營裝備清單、童軍工程架設、炊事用具或安全注意事項，請以親切、專業、結構清晰的繁體中文回答。
若對方想借還器材，請引導他使用「借 [器材名]」或「查 [器材名]」指令。`;

  try {
    const payload = {
      systemInstruction: { parts: [{ text: promptContext }] },
      contents: [{ role: "user", parts: [{ text: userMessage }] }],
      generationConfig: { temperature: 0.7, maxOutputTokens: 2000 }
    };

    for (let m = 0; m < CANDIDATE_MODELS.length; m++) {
      const model = CANDIDATE_MODELS[m];
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${GEMINI_API_KEY}`;
      const res = UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      if (res.getResponseCode() === 200) {
        const json = JSON.parse(res.getContentText());
        return json.candidates?.[0]?.content?.parts?.[0]?.text || "收到您的訊息！";
      }
    }
    return "抱歉，AI 思考忙碌中，請稍候重試或輸入「/help」查看庫存操作指令。";
  } catch (err) {
    return "系統處理異常: " + err.toString();
  }
}

/**
 * 取得 LINE 使用者暱稱
 */
function getUserDisplayName(userId) {
  if (!LINE_CHANNEL_ACCESS_TOKEN || !userId) return null;
  try {
    const url = "https://api.line.me/v2/bot/profile/" + userId;
    const res = UrlFetchApp.fetch(url, {
      headers: { "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() === 200) {
      return JSON.parse(res.getContentText()).displayName;
    }
  } catch (e) {}
  return null;
}

/**
 * 回覆訊息給 LINE 使用者 (附加底部快捷操作按鈕 Quick Reply)
 */
function replyLineMessage(replyToken, text) {
  if (!LINE_CHANNEL_ACCESS_TOKEN) {
    console.error("尚未配置 LINE_CHANNEL_ACCESS_TOKEN");
    return;
  }

  // 底部捷徑按鈕 (Quick Reply)
  const quickReplyItems = {
    items: [
      {
        type: "action",
        action: { type: "message", label: "📊 庫存總覽", text: "所有器材" }
      },
      {
        type: "action",
        action: { type: "message", label: "🔨 A區(五金)", text: "A區" }
      },
      {
        type: "action",
        action: { type: "message", label: "⛺ B區(露營)", text: "B區" }
      },
      {
        type: "action",
        action: { type: "message", label: "🍳 C區(餐廚)", text: "C區" }
      },
      {
        type: "action",
        action: { type: "message", label: "⚠️ 缺貨清單", text: "缺貨清單" }
      },
      {
        type: "action",
        action: { type: "message", label: "📖 功能選單", text: "說明" }
      }
    ]
  };

  const messages = [{
    type: "text",
    text: text.substring(0, 4000),
    quickReply: quickReplyItems
  }];

  const url = "https://api.line.me/v2/bot/message/reply";
  UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify({ replyToken: replyToken, messages: messages }),
    muteHttpExceptions: true
  });
}
