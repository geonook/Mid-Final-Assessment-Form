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

/**
 * 批次轉換所有子資料夾內的 Google Docs 為 PDF
 * 用於處理 GradeBand 子資料夾結構（G1_LTs, G1_ITs, G2_LTs...）
 * PDF 儲存在與原始 Docs 相同的子資料夾內
 */
function convertAllSubfoldersDocsToPDF() {
  // 設定輸出資料夾 ID（與主程式一致）
  const parentFolderId = '1KSyHsy1wUcrT82OjkAMmPFaJmwe-uosi';

  try {
    console.log('========================================');
    console.log('批次轉換所有子資料夾 Google Docs 為 PDF');
    console.log('========================================');

    // 取得父資料夾
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    console.log(`父資料夾: ${parentFolder.getName()}`);
    console.log('');

    // 遍歷所有子資料夾
    const subfolders = parentFolder.getFolders();
    const results = [];
    let totalProcessed = 0;
    let totalErrors = 0;

    while (subfolders.hasNext()) {
      const subfolder = subfolders.next();
      const folderName = subfolder.getName();

      // 跳過 Merged 資料夾
      if (folderName === 'Merged') {
        console.log(`跳過 "${folderName}" 資料夾`);
        console.log('');
        continue;
      }

      console.log(`處理資料夾: ${folderName}`);

      // 轉換此資料夾內的所有 Docs
      const result = convertSingleFolderDocsToPDF(subfolder);
      results.push({
        folderName: folderName,
        processed: result.processedCount,
        errors: result.errorCount
      });

      totalProcessed += result.processedCount;
      totalErrors += result.errorCount;

      console.log(`  完成: ${result.processedCount} 個檔案，錯誤: ${result.errorCount} 個`);
      console.log('');

      // 每處理一個資料夾休息 1 秒
      Utilities.sleep(1000);
    }

    // 顯示總結果
    console.log('========================================');
    console.log('批次轉換完成！');
    console.log('========================================');
    console.log(`總共處理: ${totalProcessed} 個檔案`);
    console.log(`總共錯誤: ${totalErrors} 個檔案`);
    console.log('');

    // 顯示各資料夾詳細結果
    console.log('各資料夾處理結果:');
    results.forEach(result => {
      console.log(`  ${result.folderName}: ${result.processed} 個成功, ${result.errors} 個失敗`);
    });

    return {
      totalProcessed: totalProcessed,
      totalErrors: totalErrors,
      results: results
    };

  } catch (error) {
    console.error('批次轉換發生錯誤:', error.toString());
    throw error;
  }
}

/**
 * 轉換單一資料夾內的所有 Google Docs 為 PDF
 * @param {Folder} folder - Google Drive 資料夾物件
 * @return {Object} 處理結果 {processedCount, errorCount}
 */
function convertSingleFolderDocsToPDF(folder) {
  let processedCount = 0;
  let errorCount = 0;

  try {
    // 取得資料夾內所有 Google Docs
    const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);

    // 遍歷每個 Google Docs
    while (files.hasNext()) {
      const file = files.next();

      try {
        // 取得檔名
        const fileName = file.getName();

        // 檢查是否已存在同名 PDF
        const pdfFileName = fileName + '.pdf';
        const existingPDFs = folder.getFilesByName(pdfFileName);

        if (existingPDFs.hasNext()) {
          console.log(`    跳過 ${fileName}（PDF 已存在）`);
          continue;
        }

        // 將 Google Docs 匯出為 PDF blob
        const pdfBlob = file.getAs(MimeType.PDF);
        pdfBlob.setName(pdfFileName);

        // 將 PDF 儲存到同一資料夾
        folder.createFile(pdfBlob);

        console.log(`    ✓ 轉換成功: ${fileName}`);
        processedCount++;

        // 每處理 10 個檔案休息 1 秒
        if (processedCount % 10 === 0) {
          Utilities.sleep(1000);
        }

      } catch (error) {
        console.error(`    ✗ 轉換失敗: ${file.getName()} - ${error.toString()}`);
        errorCount++;
      }
    }

  } catch (error) {
    console.error(`  資料夾處理錯誤: ${error.toString()}`);
    errorCount++;
  }

  return {
    processedCount: processedCount,
    errorCount: errorCount
  };
}