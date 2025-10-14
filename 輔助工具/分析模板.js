/**
 * 分析模板結構的輔助函數
 * 用於檢查模板的頁面尺寸、表格數量、表格結構等
 */
function analyzeTemplateStructure() {
  const templateId = '1D2hSZNI8MQzD_OIeCdEvpqp4EWfO2mrjTCHAQZyx6MM';

  try {
    const doc = DocumentApp.openById(templateId);
    const body = doc.getBody();

    let report = '=== 模板結構分析報告 ===\n\n';

    // 1. 頁面尺寸
    report += '【頁面設定】\n';
    report += `頁面寬度: ${body.getPageWidth()} points\n`;
    report += `頁面高度: ${body.getPageHeight()} points\n`;
    report += `上邊距: ${body.getMarginTop()} points\n`;
    report += `下邊距: ${body.getMarginBottom()} points\n`;
    report += `左邊距: ${body.getMarginLeft()} points\n`;
    report += `右邊距: ${body.getMarginRight()} points\n`;

    // 判斷頁面方向
    const width = body.getPageWidth();
    const height = body.getPageHeight();
    if (width === 595 && height === 842) {
      report += '頁面類型: A4 直式 (Portrait)\n';
    } else if (width === 842 && height === 595) {
      report += '頁面類型: A4 橫式 (Landscape)\n';
    } else {
      report += `頁面類型: 自訂 (${width} x ${height})\n`;
    }

    // 2. 表格分析
    const tables = body.getTables();
    report += `\n【表格分析】\n`;
    report += `表格總數: ${tables.length}\n\n`;

    tables.forEach((table, index) => {
      report += `--- 表格 ${index + 1} ---\n`;
      report += `行數: ${table.getNumRows()}\n`;

      const firstRow = table.getRow(0);
      report += `欄數: ${firstRow.getNumCells()}\n`;

      // 讀取第一行（標題行）的內容
      report += '標題行內容: ';
      const headers = [];
      for (let col = 0; col < firstRow.getNumCells(); col++) {
        headers.push(firstRow.getCell(col).getText().trim());
      }
      report += headers.join(' | ') + '\n';

      // 顯示欄寬
      report += '欄寬設定: ';
      const widths = [];
      for (let col = 0; col < firstRow.getNumCells(); col++) {
        try {
          const width = table.getColumnWidth(col);
          widths.push(`${width}pt`);
        } catch (e) {
          widths.push('未設定');
        }
      }
      report += widths.join(', ') + '\n\n';
    });

    // 3. 文件元素統計
    report += '【文件元素統計】\n';
    const numChildren = body.getNumChildren();
    report += `總元素數: ${numChildren}\n`;

    let paragraphCount = 0;
    let tableCount = 0;
    let pageBreakCount = 0;

    for (let i = 0; i < numChildren; i++) {
      const element = body.getChild(i);
      const type = element.getType();

      if (type === DocumentApp.ElementType.PARAGRAPH) {
        paragraphCount++;
      } else if (type === DocumentApp.ElementType.TABLE) {
        tableCount++;
      } else if (type === DocumentApp.ElementType.PAGE_BREAK) {
        pageBreakCount++;
      }
    }

    report += `段落數: ${paragraphCount}\n`;
    report += `表格數: ${tableCount}\n`;
    report += `分頁符號數: ${pageBreakCount}\n`;

    // 4. 佔位符檢查
    report += '\n【佔位符檢查】\n';
    const bodyText = body.getText();
    const placeholders = ['{{Level}}', '{{ClassName}}', '{{Tea}}', '{{IT}}', '{{Count}}', '{{Students}}'];

    placeholders.forEach(placeholder => {
      const found = bodyText.includes(placeholder);
      report += `${placeholder}: ${found ? '✓ 找到' : '✗ 未找到'}\n`;
    });

    // 5. 學生表格識別
    report += '\n【學生表格識別建議】\n';
    let studentTableIndex = -1;

    tables.forEach((table, index) => {
      const firstRow = table.getRow(0);
      const firstCell = firstRow.getCell(0).getText().trim();

      // 檢查是否為學生表格
      if (firstCell.includes('Student ID') || firstCell.includes('ID') ||
          firstRow.getNumCells() === 6) {
        studentTableIndex = index;
        report += `可能的學生表格: 表格 ${index + 1} (索引 ${index})\n`;
        report += `理由: `;
        if (firstCell.includes('Student ID')) report += '包含 "Student ID" ';
        if (firstRow.getNumCells() === 6) report += '有 6 欄 ';
        report += '\n';
      }
    });

    if (studentTableIndex === -1) {
      report += '⚠️ 無法自動識別學生表格，請手動確認\n';
    }

    // 輸出報告
    Logger.log(report);
    console.log(report);

    return report;

  } catch (e) {
    const errorMsg = `❌ 分析失敗: ${e.message}`;
    Logger.log(errorMsg);
    console.error(errorMsg);
    return errorMsg;
  }
}

/**
 * 從選單執行分析
 */
function runTemplateAnalysis() {
  try {
    const report = analyzeTemplateStructure();
    SpreadsheetApp.getUi().alert('模板結構分析\n\n' + report);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ 分析失敗: ' + e.message);
  }
}

/**
 * 添加選單項目
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('班級報告')
    .addItem('分析模板結構', 'runTemplateAnalysis')
    .addToUi();
}
