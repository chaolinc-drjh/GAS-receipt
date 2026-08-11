// =========================================================================
// 1. Web App 進入點 (Main Entry Point)
// =========================================================================

/**
 * Web App 應用程式進入點
 * 負責載入 HTML 範本、設定頁面標題與 Meta 標籤
 * 
 * @param {Object} e - HTTP GET 請求事件物件
 * @returns {HtmlService.HtmlOutput} 渲染後的 HTML 頁面物件
 */
function doGet(e) {
  let IMAGEID = "1DExNWkliQNVaANWYDFJ7pkDfNeqYb5WP";
  try {
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG');
    if (configSheet) {
      const val = configSheet.getRange('D2').getValue().toString().trim();
      if (val) IMAGEID = val;
    }
  } catch (err) {}
  
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('雲端領據管理系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl(`https://drive.google.com/uc?id=${IMAGEID}&export=download&format=png`);
}

// =========================================================================
// 2. 身份驗證與使用者管理 (Authentication & User Management)
// =========================================================================

/**
 * 驗證目前登入使用者的身分與 Email 網域權限
 * 讀取 CONFIG 工作表對照允許網域，若驗證通過則同時取得使用者詳細資料
 * 
 * @returns {Object} 包含 email, isValid(是否合法), userInfo(使用者資訊), allowedDomains(允許網域清單)
 */
function verifyUser() {
  try {
    const email = Session.getActiveUser().getEmail();
    let isValid = false;
    let userInfo = { name: '未知', phone: '', voip: '' };
    let allowedDomains = []; // 用來記錄允許的網域清單

    if (email) {
      const domain = email.split('@')[1].toLowerCase();

      // ★ 讀取 CONFIG 資料表的 A2:A5
      const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG');
      if (configSheet) {
        const domainValues = configSheet.getRange('A2:A5').getValues();
        // 轉換為一維陣列，濾除空白、統一轉為小寫，並移除可能帶有的 '@' 符號
        allowedDomains = domainValues
          .map(row => {
            let d = row[0].toString().trim().toLowerCase();
            if (d.startsWith('@')) d = d.substring(1);
            return d;
          })
          .filter(d => d !== '');
      }

      // ★ 判斷登入者的網域是否在清單中
      isValid = allowedDomains.includes(domain);

      if (isValid) userInfo = getUserInfo(email);
    }
    
    // 動態讀取前端所需的設定變數 (請款單位、匯入帳號等)
    let appConfig = {
      defaultDept: '雲林縣政府',
      accounts: []
    };
    
    if (isValid) {
      const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG');
      if (configSheet) {
        const deptVal = configSheet.getRange('E2').getValue().toString().trim();
        if (deptVal) appConfig.defaultDept = deptVal;
        
        const accNames = configSheet.getRange('F2:F10').getValues();
        const accNums = configSheet.getRange('G2:G10').getValues();
        for (let i = 0; i < accNames.length; i++) {
          let n = accNames[i][0].toString().trim();
          let m = accNums[i][0].toString().trim();
          if (n) appConfig.accounts.push({ name: n, num: m });
        }
      }
    }

    // 將 allowedDomains 與 appConfig 傳回前端
    return { email: email, isValid: isValid, userInfo: userInfo, allowedDomains: allowedDomains, appConfig: appConfig };
  } catch (error) {
    return { email: '未知', isValid: false, userInfo: {}, error: error.toString() };
  }
}

/**
 * 依據 Email 查詢 USER 工作表，取得使用者個人資料與權限角色
 * 
 * @param {string} email - 使用者的 Email 地址
 * @returns {Object} 使用者詳細資料
 */
function getUserInfo(email) {
  const id = email.split('@')[0].toLowerCase();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USER');
  const defaultInfo = { 
    name: id, 
    phone: '', 
    voip: '', 
    isRegistered: false, 
    isAdmin: false, 
    isSuperAdmin: false 
  };
  
  if (!sheet) return defaultInfo;
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const rowId = data[i][0].toString().trim().toLowerCase();
    
    // 比對帳號 (帳號 ID 或 完整 Email)
    if (rowId === id || rowId === email.toLowerCase()) {
      const role = data[i][4] ? data[i][4].toString().trim().toUpperCase() : ''; // 取出 E 欄角色標記
      
      const isSuperAdmin = (role === 'SUPER_ADMIN');
      const isAdmin = isSuperAdmin || (role === 'ADMIN');
      
      return {
        name: data[i][1] ? data[i][1].toString().trim() : id,
        phone: data[i][2] ? data[i][2].toString().trim() : '',
        voip: data[i][3] ? data[i][3].toString().trim() : '',
        isRegistered: (role !== ''),
        isAdmin: isAdmin,
        isSuperAdmin: isSuperAdmin
      };
    }
  }
  
  return defaultInfo;
}

/**
 * 更新 USER 工作表中指定使用者的電話與 VOIP 網路電話號碼
 * 
 * @param {string} email - 使用者的 Email
 * @param {string} [phone] - 聯絡電話/手機號碼
 * @param {string} [voip] - VOIP 網路電話號碼
 */
function updateUserPhone(email, phone, voip) {
  const id = email.split('@')[0];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USER');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === id.toLowerCase()) {
      if (phone !== undefined) sheet.getRange(i + 1, 3).setValue(phone);
      if (voip !== undefined) sheet.getRange(i + 1, 4).setValue(voip);
      return;
    }
  }
}


// =========================================================================
// 3. 領據核心業務邏輯 (Receipt Core CRUD)
// =========================================================================

/**
 * 讀取所有領據紀錄清單，並解析備註中的歷史入帳紀錄與判斷編輯權限
 * 
 * @returns {Array<Object>} 領據物件陣列（依據建立時間反向排序，最新資料在最前）
 */
function getReceiptsList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('領據紀錄');
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const data = sheet.getRange(1, 1, lastRow, 15).getValues();
  const currentUserEmail = Session.getActiveUser().getEmail();
  const today = new Date();
  today.setHours(0,0,0,0);
  
  const result = [];
  
  for (let i = data.length - 1; i >= 1; i--) {
    let row = data[i];
    let dateObj = new Date(row[1]); 
    let timeDiff = today.getTime() - dateObj.getTime();
    let daysDiff = Math.floor(timeDiff / (1000 * 3600 * 24));
    
    let isCreator = (row[13] === currentUserEmail);
    let status = row[6].toString();
    let isEditable = isCreator && (daysDiff <= 7) && (status !== '作廢');
    
    let amount = parseInt(row[4], 10) || 0;
    let remarksStr = row[9].toString();
    
    // ★ 從備註中解析總入帳金額
    let depositedAmount = 0;
    let matches = remarksStr.match(/\[入帳\][^:]+:\s*\$?([0-9,]+)/g);
    if (matches) {
      matches.forEach(m => {
        let val = m.split(':')[1].replace(/[^0-9]/g, '');
        depositedAmount += parseInt(val, 10);
      });
    }

    result.push({
      receiptNum: row[0].toString(),
      date: Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd'),
      department: row[2].toString(),
      projectName: row[3].toString(),
      amount: amount,
      depositedAmount: depositedAmount, // ★ 傳遞給前端
      amountZh: row[5].toString(),
      status: status,
      phone: row[7].toString(),
      voip: row[8].toString(),
      remarks: remarksStr,
      accountType: row[10].toString(),
      accountNum: row[11].toString(),
      creator: row[12].toString(),
      creatorEmail: row[13] ? row[13].toString() : '',
      creatorConfirmed: row[14] ? row[14].toString() : '',
      isEditable: isEditable
    });
  }
  return result;
}

/**
 * 提交領據表單資料（支援「新增 mode: add」與「編輯 mode: edit」）
 * 使用 LockService 防止併發寫入衝突
 * 
 * @param {Object} data - 前端傳入的表單資料（包含 mode, receiptNum, department, projectName, amount, phone, voip, remarks, accountType, isVoid 等）
 * @returns {Object} `{ success: boolean, receiptNum: string, message: string }` 執行結果
 */
function submitReceipt(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const email = Session.getActiveUser().getEmail();
    updateUserPhone(email, data.phone, data.voip);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('領據紀錄');
    if (!sheet) throw new Error("找不到名為「領據紀錄」的工作表。");

    let accountNum = '';
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG');
    if (configSheet) {
      const accNames = configSheet.getRange('F2:F10').getValues();
      const accNums = configSheet.getRange('G2:G10').getValues();
      for (let i = 0; i < accNames.length; i++) {
        if (accNames[i][0].toString().trim() === data.accountType) {
          accountNum = accNums[i][0].toString().trim();
          break;
        }
      }
    }

    const amountZh = convertToChineseAmount(data.amount);

    if (data.mode === 'add') {
      const today = new Date();
      const dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      const creatorName = getUserInfo(email).name; 
      const minguoYear = today.getFullYear() - 1911;
      const receiptNum = generateReceiptNumber(sheet, minguoYear);
      const status = "未入帳"; 

      const rowData = [
        receiptNum, dateString, data.department, data.projectName,
        data.amount, amountZh, status, data.phone || '', data.voip || '',
        data.remarks || '', data.accountType, accountNum, creatorName, email
      ];
      sheet.appendRow(rowData);
      return { success: true, receiptNum: receiptNum, message: '領據建立成功！' };
      
    } else if (data.mode === 'edit') {
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) throw new Error("試算表內尚無領據資料！");
      
      const idValues = sheet.getRange(1, 1, lastRow, 1).getValues(); 
      let rowIndex = -1;
      const targetNum = String(data.receiptNum).trim(); 
      
      for (let i = 1; i < idValues.length; i++) {
        if (String(idValues[i][0]).trim() === targetNum) {
          rowIndex = i + 1; 
          break;
        }
      }
      
      if (rowIndex === -1) {
        let sample = idValues.slice(1, 4).map(r => r[0]).join(', ');
        throw new Error(`找不到編號 [${targetNum}]。現有範例：${sample}`);
      }

      const currentStatus = sheet.getRange(rowIndex, 7).getValue();
      const status = data.isVoid ? "作廢" : currentStatus;

      sheet.getRange(rowIndex, 3).setValue(data.department);
      sheet.getRange(rowIndex, 4).setValue(data.projectName);
      sheet.getRange(rowIndex, 5).setValue(data.amount);
      sheet.getRange(rowIndex, 6).setValue(amountZh);
      sheet.getRange(rowIndex, 7).setValue(status);
      sheet.getRange(rowIndex, 8).setValue(data.phone || '');
      sheet.getRange(rowIndex, 9).setValue(data.voip || '');
      
      // ★ 編輯模式保留舊有入帳紀錄，疊加新的備註
      let oldRemarks = sheet.getRange(rowIndex, 10).getValue().toString();
      let depositRecords = oldRemarks.match(/\[入帳\][\s\S]*/); // 抓出以前的入帳字眼
      let finalRemarks = data.remarks || '';
      if (depositRecords) finalRemarks += "\n" + depositRecords[0];
      
      sheet.getRange(rowIndex, 10).setValue(finalRemarks.trim());
      sheet.getRange(rowIndex, 11).setValue(data.accountType);
      sheet.getRange(rowIndex, 12).setValue(accountNum);

      return { success: true, receiptNum: targetNum, message: '領據修改成功！' };
    }
    
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 將指定領據號碼的紀錄標記為「作廢」，並追加作廢原因
 * 
 * @param {string} receiptNum - 領據編號
 * @param {string} remark - 作廢原因或備註
 * @returns {Object} `{ success: boolean, message?: string }` 執行結果
 */
function voidReceiptRecord(receiptNum, remark) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('領據紀錄');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("無資料");
    
    const idValues = sheet.getRange(1, 1, lastRow, 1).getValues();
    
    for (let i = 1; i < idValues.length; i++) {
      if (String(idValues[i][0]).trim() === String(receiptNum).trim()) {
        sheet.getRange(i + 1, 7).setValue("作廢");
        let oldRemarks = sheet.getRange(i + 1, 10).getValue().toString();
        sheet.getRange(i + 1, 10).setValue(oldRemarks ? oldRemarks + "\n" + remark : remark);
        return { success: true };
      }
    }
    throw new Error("找不到該領據");
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 登記經費入帳紀錄
 * 將入帳日期與金額附加至備註、更新狀態為「已入帳」，並透過 Email 自動發送通知給製表人
 * 
 * @param {string} receiptNum - 領據編號
 * @param {string} depositDate - 入帳日期 (例如: '2026-08-10')
 * @param {string|number} depositAmountStr - 入帳金額字串或數字
 * @returns {Object} `{ success: boolean, message: string }` 執行結果與提示訊息
 */
function addDepositRecord(receiptNum, depositDate, depositAmountStr) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('領據紀錄');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("無資料");
    
    const idValues = sheet.getRange(1, 1, lastRow, 1).getValues();
    let rowIndex = -1;
    for (let i = 1; i < idValues.length; i++) {
      if (String(idValues[i][0]).trim() === String(receiptNum).trim()) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) throw new Error("找不到該領據");

    // 取出信件需要的各項資料
    const amount = parseInt(sheet.getRange(rowIndex, 5).getValue(), 10);
    const projectName = sheet.getRange(rowIndex, 4).getValue();   // 計畫名稱 (D欄)
    const creatorName = sheet.getRange(rowIndex, 13).getValue();  // 製表人姓名 (M欄)
    const creatorEmail = sheet.getRange(rowIndex, 14).getValue(); // 製表人 Email (N欄)
    let currentRemarks = sheet.getRange(rowIndex, 10).getValue().toString();
    
    const addVal = parseInt(depositAmountStr.toString().replace(/[^0-9]/g, ''), 10);
    if (isNaN(addVal) || addVal <= 0) throw new Error("入帳金額無效");

    // 格式化日期並拼接入帳字串
    const formattedDate = depositDate.replace(/-/g, '/');
    let newRecord = `\n[入帳] ${formattedDate}: $${addVal.toLocaleString()}`;
    let newRemarks = currentRemarks ? currentRemarks.trim() + newRecord : newRecord.trim();
    
    // 重新計算總額
    let deposited = 0;
    let matches = newRemarks.match(/\[入帳\][^:]+:\s*\$?([0-9,]+)/g);
    if (matches) {
      matches.forEach(m => {
        let val = m.split(':')[1].replace(/[^0-9]/g, '');
        deposited += parseInt(val, 10);
      });
    }

    // 更新試算表
    sheet.getRange(rowIndex, 10).setValue(newRemarks);
    sheet.getRange(rowIndex, 7).setValue("已入帳"); 
    // ★ 新增：每次出納有新入帳，就清空製表人確認欄，以便再次通知
    sheet.getRange(rowIndex, 15).clearContent(); 

    // ★ 發送自動化 Email 給製表人
    if (creatorEmail) {
      const remaining = Math.max(0, amount - deposited);
      // 使用 HTML 代碼：&#128994; (綠色圓) 與 &#128993; (黃色圓)
      const statusText = deposited >= amount ? '&#128994; 已結清' : '&#128993; 部分入帳';
      const emailSubject = `[系統通知] 領據 ${receiptNum} 經費入帳通知`;
      
      const emailBody = `
        <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; color: #333; line-height: 1.6;">
          <h2 style="color: #198754;">&#128176; 經費入帳通知</h2>
          <p>您好，<strong>${creatorName}</strong>：</p>
          <p>您所製作的領據目前已有經費入帳，系統為您更新明細如下：</p>
          
          <table style="border-collapse: collapse; width: 100%; max-width: 550px; margin-top: 15px; border: 1px solid #dee2e6;">
            <tr><td style="padding: 10px; border: 1px solid #dee2e6; background: #f8f9fa; font-weight: bold; width: 35%;">領據編號</td><td style="padding: 10px; border: 1px solid #dee2e6;">${receiptNum}</td></tr>
            <tr><td style="padding: 10px; border: 1px solid #dee2e6; background: #f8f9fa; font-weight: bold;">計畫名稱</td><td style="padding: 10px; border: 1px solid #dee2e6;">${projectName}</td></tr>
            <tr><td style="padding: 10px; border: 1px solid #dee2e6; background: #f8f9fa; font-weight: bold;">本次入帳日期</td><td style="padding: 10px; border: 1px solid #dee2e6;">${formattedDate}</td></tr>
            <tr><td style="padding: 10px; border: 1px solid #dee2e6; background: #e8f5e9; font-weight: bold; color: #198754;">本次入帳金額</td><td style="padding: 10px; border: 1px solid #dee2e6; font-weight: bold; color: #198754; font-size: 1.1em;">$${addVal.toLocaleString()}</td></tr>
            <tr><td style="padding: 10px; border: 1px solid #dee2e6; background: #f8f9fa; font-weight: bold;">目前總入帳 / 總額</td><td style="padding: 10px; border: 1px solid #dee2e6;">$${deposited.toLocaleString()} / $${amount.toLocaleString()}</td></tr>
            <tr><td style="padding: 10px; border: 1px solid #dee2e6; background: ${remaining > 0 ? '#fdf0f1' : '#e8f5e9'}; font-weight: bold; color: ${remaining > 0 ? '#dc3545' : '#198754'};">待入帳餘額</td><td style="padding: 10px; border: 1px solid #dee2e6; font-weight: bold; color: ${remaining > 0 ? '#dc3545' : '#198754'};">$${remaining.toLocaleString()}</td></tr>
            <tr><td style="padding: 10px; border: 1px solid #dee2e6; background: #f8f9fa; font-weight: bold;">當前狀態</td><td style="padding: 10px; border: 1px solid #dee2e6; font-weight: bold;">${statusText}</td></tr>
          </table>
          
          <p style="margin-top: 25px; font-size: 0.85em; color: #6c757d; border-top: 1px solid #eee; padding-top: 10px;">
            ※ 此為系統自動發送之信件，請勿直接回覆。若有疑問請洽出納人員。<br>
            ※ 可登入 <a href="${ScriptApp.getService().getUrl()}" style="color: #0d6efd;">雲端領據管理系統</a> 查看詳細紀錄。
          </p>
        </div>
      `;
      
      // 群組信箱寄信
      try {
        const aliases = GmailApp.getAliases();
        let targetAlias = "noreply@drjh.ylc.edu.tw"; // fallback
        const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG');
        if (configSheet) {
          const val = configSheet.getRange('C2').getValue().toString().trim();
          if (val) targetAlias = val;
        }
        
        if (aliases.includes(targetAlias)) {
          GmailApp.sendEmail(creatorEmail, emailSubject, "", {
            htmlBody: emailBody,
            name: "領據通知中心",
            from: targetAlias
          });
        } else {
          GmailApp.sendEmail(creatorEmail, emailSubject, "", {
            htmlBody: emailBody,
            name: "資訊組"
          });
          console.warn("未設定或無權限使用 noreply 別名，已改用預設帳號發送。");
        }
      } catch (mailError) {
        console.error("Email 發送失敗: " + mailError);
      }
    }

    return { success: true, message: '入帳登記成功！系統已發送 Email 通知製表人。' };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ★ 新增：製表人確認已讀入帳通知
 */
function confirmReceiptDeposit(receiptNums) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('領據紀錄');
    const data = sheet.getDataRange().getValues();
    // 取得當下時間
    const timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

    for (let i = 1; i < data.length; i++) {
      if (receiptNums.includes(data[i][0].toString().trim())) {
        sheet.getRange(i + 1, 15).setValue(timeStr); // 將時間寫入第 O 欄
      }
    }
    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}



// =========================================================================
// 4. 檔案與資料管理 (Document & Data Management)
// =========================================================================

/**
 * 根據指定領據編號，複製 Google Docs 範本進行動態取代，匯出包含「收據聯」與「存根聯」的 PDF 文件
 * 
 * @param {string} receiptNum - 領據編號
 * @returns {Object} `{ success: boolean, fileName: string, base64: string }` PDF 檔名與 Base64 數據供前端下載
 */
function generatePDF(receiptNum) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('領據紀錄');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("無資料");
  
  const values = sheet.getRange(1, 1, lastRow, 14).getValues();
  let record = null;
  
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(receiptNum).trim()) {
      record = values[i];
      break;
    }
  }
  
  if (!record) throw new Error("找不到領據資料");

  // ★ 從 CONFIG 資料表動態讀取範本 ID
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG');
  let TEMPLATE_ID = '';
    if (configSheet) {
    TEMPLATE_ID = configSheet.getRange('B2').getValue().toString().trim();
  }
    if (!TEMPLATE_ID) {
    throw new Error("系統找不到列印範本 ID！請確認「CONFIG」資料表的 B2 儲存格已填入正確的 Google Doc ID。");
  }

  const tempFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(`暫存處理_${receiptNum}`);
  const tempDoc = DocumentApp.openById(tempFile.getId());
  const tempBody = tempDoc.getBody();
  
  const dateObj = new Date(record[1]);
  const rocYear = dateObj.getFullYear() - 1911;
  const mm = ('0' + (dateObj.getMonth() + 1)).slice(-2);
  const dd = ('0' + dateObj.getDate()).slice(-2);
  const dateStr = `中 華 民 國 ${rocYear} 年 ${mm} 月 ${dd} 日`;
  const amountWithComma = record[4].toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");

  /**
   * 內部輔助函式：文字過長時自動調小字型大小並替換佔位符
   * 
   * @param {string} placeholder - 被替換的樣板文字，如 '{{請款單位}}'
   * @param {string} replacementText - 欲填入的實際內容
   */
  function replaceAndResize(placeholder, replacementText) {
    const length = replacementText.length;
    let multiplier = 1.0;
    if (length > 16) {
      const steps = Math.ceil((length - 16) / 10);
      multiplier = Math.pow(0.9, steps);
    }
    let found = tempBody.findText(placeholder);
    while (found) {
      let textElement = found.getElement().asText();
      let start = found.getStartOffset();
      let end = found.getEndOffsetInclusive();
      let originalSize = textElement.getFontSize(start);
      if (!originalSize) {
        let parentAttrs = textElement.getParent().getAttributes();
        originalSize = parentAttrs[DocumentApp.Attribute.FONT_SIZE] || 14;
      }
      let newSize = Math.max(1, Math.round(originalSize * multiplier));
      textElement.deleteText(start, end);
      if (replacementText !== "") {
        textElement.insertText(start, replacementText);
        textElement.setFontSize(start, start + length - 1, newSize);
      }
      found = tempBody.findText(placeholder);
    }
  }

  replaceAndResize('{{請款單位}}', record[2].toString());
  replaceAndResize('{{計畫名稱}}', record[3].toString());
  replaceAndResize('{{金額大寫}}', record[5].toString());

  tempBody.replaceText('{{領據號碼}}', record[0].toString());
  tempBody.replaceText('{{日期}}', dateStr);
  tempBody.replaceText('{{建檔日期}}', dateStr);
  tempBody.replaceText('{{金額}}', amountWithComma);
  tempBody.replaceText('{{狀態}}', record[6].toString());
  tempBody.replaceText('{{連絡電話}}', record[7].toString());
  tempBody.replaceText('{{網路電話}}', record[8].toString());
  // PDF 只印出除去了入帳紀錄的純備註
  let pureRemarks = record[9].toString().replace(/\[入帳\][\s\S]*/, '').trim();
  tempBody.replaceText('{{備註}}', pureRemarks);
  
  tempBody.replaceText('{{帳戶}}', record[10].toString());
  tempBody.replaceText('{{帳號}}', record[11].toString());
  tempBody.replaceText('{{製表人}}', record[12].toString());
  
  const cleanDoc = DocumentApp.create(`${record[3]}`);
  const cleanBody = cleanDoc.getBody();
  
  cleanBody.setMarginTop(tempBody.getMarginTop());
  cleanBody.setMarginBottom(tempBody.getMarginBottom());
  cleanBody.setMarginLeft(tempBody.getMarginLeft());
  cleanBody.setMarginRight(tempBody.getMarginRight());
  cleanBody.setPageHeight(tempBody.getPageHeight());
  cleanBody.setPageWidth(tempBody.getPageWidth());

  var pageStyle = {};
  pageStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#FFFFFF';
  cleanBody.setAttributes(pageStyle);

  const elements = [];
  for (let i = 0; i < tempBody.getNumChildren(); i++) {
    elements.push(tempBody.getChild(i).copy());
  }

  for (let i = 0; i < elements.length; i++) {
    let el = elements[i].copy();
    let type = el.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) cleanBody.appendParagraph(el);
    else if (type === DocumentApp.ElementType.TABLE) cleanBody.appendTable(el);
    else if (type === DocumentApp.ElementType.LIST_ITEM) cleanBody.appendListItem(el);
  }
  
  if (cleanBody.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH && cleanBody.getChild(0).getText().trim() === '') {
    cleanBody.removeChild(cleanBody.getChild(0));
  }

  let match1 = cleanBody.findText('{{類型}}');
  if (match1) {
    let textEl = match1.getElement().asText();
    textEl.deleteText(match1.getStartOffset(), match1.getEndOffsetInclusive());
    textEl.insertText(match1.getStartOffset(), '(收據聯)');
  }

  cleanBody.appendPageBreak();

  for (let i = 0; i < elements.length; i++) {
    let el = elements[i].copy();
    let type = el.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) cleanBody.appendParagraph(el);
    else if (type === DocumentApp.ElementType.TABLE) cleanBody.appendTable(el);
    else if (type === DocumentApp.ElementType.LIST_ITEM) cleanBody.appendListItem(el);
  }

  cleanBody.replaceText('{{類型}}', '(存根聯)');

  tempDoc.saveAndClose();
  cleanDoc.saveAndClose();
  
  const cleanFile = DriveApp.getFileById(cleanDoc.getId());
  const pdfBlob = cleanFile.getAs('application/pdf');
  const base64Data = Utilities.base64Encode(pdfBlob.getBytes());
  
  tempFile.setTrashed(true);
  cleanFile.setTrashed(true);
  
  return {
    success: true,
    fileName: `領據_${receiptNum}.pdf`,
    base64: base64Data
  };
}

/**
 * 將指定單號（含）以前的所有歷史領據資料搬移至「封存」工作表，並從「領據紀錄」刪除
 * 
 * @param {string} targetReceiptNum - 目標領據單號（該單號及其之前的資料都會被封存）
 * @returns {Object} `{ success: boolean, message: string }` 執行結果訊息
 */
function archiveRecords(targetReceiptNum) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName('領據紀錄');
    let archiveSheet = ss.getSheetByName('封存');
    
    // 如果沒有封存資料表，系統自動建立一個，並複製標題列
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet('封存');
      const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn());
      headers.copyTo(archiveSheet.getRange(1, 1));
    }
    
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < 2) throw new Error("沒有資料可封存");
    
    // 找出目標單號所在的列數 (絕對列號)
    const idValues = sourceSheet.getRange(1, 1, lastRow, 1).getValues();
    let targetRowIndex = -1;
    for (let i = 1; i < idValues.length; i++) {
      if (String(idValues[i][0]).trim() === String(targetReceiptNum).trim()) {
        targetRowIndex = i + 1;
        break;
      }
    }
    
    if (targetRowIndex === -1) throw new Error("找不到目標領據號碼：" + targetReceiptNum);
    
    // 計算要搬移的列數 (從第 2 列一直到目標列)
    const numRowsToArchive = targetRowIndex - 1; 
    const numCols = sourceSheet.getLastColumn();
    
    const sourceRange = sourceSheet.getRange(2, 1, numRowsToArchive, numCols);
    const archiveLastRow = archiveSheet.getLastRow();
    const targetRange = archiveSheet.getRange(archiveLastRow + 1, 1);
    
    // 將資料(包含格式)複製到封存表，然後從原表刪除
    sourceRange.copyTo(targetRange);
    sourceSheet.deleteRows(2, numRowsToArchive);
    
    return { success: true, message: `成功將 ${numRowsToArchive} 筆舊資料搬移至「封存」資料表！` };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}


// =========================================================================
// 5. 通用工具與格式轉換函式 (Utility Helpers)
// =========================================================================

/**
 * 自動產生領據編號（格式：民國年 + 3位數流水號，如 115001）
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 領據紀錄工作表物件
 * @param {number} minguoYear - 當前民國年份 (如 115)
 * @returns {string} 格式化後的領據編號
 */
function generateReceiptNumber(sheet, minguoYear) {
  var lastRow = sheet.getLastRow();
  var prefix = minguoYear.toString();
  if (lastRow <= 1) return prefix + '001';
  var lastNumber = sheet.getRange(lastRow, 1).getValue().toString();
  if (lastNumber.indexOf(prefix) === 0 && lastNumber.length === prefix.length + 3) {
    var sequence = parseInt(lastNumber.slice(-3), 10) + 1;
    return prefix + ('000' + sequence).slice(-3);
  } else {
    return prefix + '001';
  }
}

/**
 * 將阿拉伯數字金額轉換為中文大寫金額（如：10500 -> 壹萬零伍佰元整）
 * 
 * @param {number|string} num - 阿拉伯數字金額
 * @returns {string} 中文大寫金額字串
 */
function convertToChineseAmount(num) {
  if (!num || isNaN(num) || num == 0) return "零元整";
  var digit = ['零', '壹', '貳', '參', '肆', '伍', '陸', '柒', '捌', '玖'];
  var unit = [['元', '萬', '億'], ['', '拾', '佰', '仟']];
  var numStr = Math.abs(num).toString();
  var res = '';
  var zeroCount = 0;
  for (var i = 0; i < numStr.length; i++) {
    var p = numStr.length - i - 1;
    var q = Math.floor(p / 4);
    var m = p % 4;
    var n = parseInt(numStr.charAt(i));
    if (n == 0) { zeroCount++; } else {
      if (zeroCount > 0) res += digit[0];
      zeroCount = 0;
      res += digit[n] + unit[1][m];
    }
    if (m == 0 && (zeroCount < 4 || q == 0) ) res += unit[0][q];
  }
  return res + '整';
}


/**
 * ★ 用來將其他 HTML 檔案(如 JS 或 CSS) 引入到主網頁的函式
 */
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

/**
 ******** 額外工作區 **********
 *
 * 
 */


/**
 * 測試可用此 mail 寄信
 */

function testNoreplyEmail() {

  // ★ 從 CONFIG 資料表動態讀取email
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CONFIG');
  let TEMPLATE_MAIL = '';
    if (configSheet) {
    TEMPLATE_MAIL = configSheet.getRange('C2').getValue().toString().trim();
  }
    if (!TEMPLATE_MAIL) {
    throw new Error("系統找不到email！請確認「CONFIG」資料表的 C2 儲存格已填入正確的 Google Mail。");
  }

  const targetAlias = TEMPLATE_MAIL;
  const myEmail = Session.getActiveUser().getEmail(); // 系統會自動抓您的信箱當收件人
  
  // 1. 抓取目前帳號被授權的所有別名
  const aliases = GmailApp.getAliases();
  console.log("您目前擁有的別名權限有：", aliases);
  
  // 2. 判斷是否有 noreply 的權限並發信
  if (aliases.includes(targetAlias)) {
    try {
      GmailApp.sendEmail(myEmail, "【系統測試】別名發信測試", "恭喜！這封信是由 noreply 別名成功發出的測試信件！", {
        from: targetAlias,
        name: "測試系統"
      });
      console.log("✅ 測試成功！信件已使用 " + targetAlias + " 寄出，請去信箱檢查。");
    } catch (e) {
      console.error("❌ 發信失敗，錯誤訊息：" + e.message);
    }
  } else {
    console.error("❌ 權限異常：您的帳號無法使用 " + targetAlias + " 發信。請確認 Workspace 後台設定。");
  }
}


/**
 * ★ 首次部署專用：一鍵完成所有權限授權 (在 GAS 編輯器手動執行此函式一次即可)
 */
function setupPermissions() {
  // 1. 觸發 SpreadsheetApp & Session 權限
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const email = Session.getActiveUser().getEmail();
  
  // 2. 觸發 GmailApp & ScriptApp 權限
  const aliases = GmailApp.getAliases();
  const webAppUrl = ScriptApp.getService().getUrl();
  
  // 3. 觸發 DriveApp & DocumentApp 權限
  const tempDoc = DocumentApp.create('權限測試暫存檔');
  const tempFile = DriveApp.getFileById(tempDoc.getId());
  tempFile.setTrashed(true); // 測試完自動刪除

  Logger.log("✅ 所有 API 權限授權成功！");
}

