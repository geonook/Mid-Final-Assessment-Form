function batchDownloadDocsToPDF() {
  // 設定資料夾ID（與主程式輸出資料夾一致）
  const folderId = '1KSyHsy1wUcrT82OjkAMmPFaJmwe-uosi';
  
  try {
    // 取得資料夾
    const folder = DriveApp.getFolderById(folderId);
    
    // 取得資料夾内所有Google文件
    const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);
    
    let processedCount = 0;
    let errorCount = 0;
    
    // 遍歷每個Google文件
    while (files.hasNext()) {
      const file = files.next();
      
      try {
        // 取得文件名稱
        const fileName = file.getName();
        
        // 檢查是否已存在同名PDF
        const existingPDFs = folder.getFilesByName(fileName + '.pdf');
        if (existingPDFs.hasNext()) {
          console.log(`跳過 ${fileName}，PDF已存在`);
          continue;
        }
        
        // 將Google文件匯出為PDF blob
        const pdfBlob = file.getAs(MimeType.PDF);
        
        // 設定PDF檔名
        pdfBlob.setName(fileName + '.pdf');
        
        // 將PDF存放到同一資料夾
        const pdfFile = folder.createFile(pdfBlob);
        
        console.log(`成功轉換: ${fileName} -> ${fileName}.pdf`);
        processedCount++;
        
        // 避免超過執行時間限制，每處理10個檔案休息一下
        if (processedCount % 10 === 0) {
          Utilities.sleep(1000);
        }
        
      } catch (error) {
        console.error(`處理 ${file.getName()} 時發生錯誤:`, error.toString());
        errorCount++;
      }
    }
    
    // 顯示處理結果
    console.log(`批次處理完成！`);
    console.log(`成功處理: ${processedCount} 個檔案`);
    console.log(`發生錯誤: ${errorCount} 個檔案`);
    
    // 可選：發送完成通知email
    // sendCompletionEmail(processedCount, errorCount);
    
  } catch (error) {
    console.error('批次處理發生錯誤:', error.toString());
  }
}

// 可選：發送完成通知的函數
function sendCompletionEmail(successCount, errorCount) {
  const email = Session.getActiveUser().getEmail();
  const subject = 'Google文件PDF轉換完成通知';
  const body = `
批次PDF轉換已完成：

✅ 成功處理: ${successCount} 個檔案
❌ 發生錯誤: ${errorCount} 個檔案

請查看Google Drive確認結果。
  `;
  
  GmailApp.sendEmail(email, subject, body);
}

// 輔助函數：取得資料夾ID
function getFolderIdByName(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    const folder = folders.next();
    console.log(`資料夾 "${folderName}" 的ID: ${folder.getId()}`);
    return folder.getId();
  } else {
    console.log(`找不到名為 "${folderName}" 的資料夾`);
    return null;
  }
}

// 輔助函數：列出資料夾內容
function listFolderContents(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  
  console.log(`資料夾 "${folder.getName()}" 內容:`);
  while (files.hasNext()) {
    const file = files.next();
    console.log(`- ${file.getName()} (${file.getMimeType()})`);
  }
}