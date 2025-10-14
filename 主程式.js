/**
 * 2526 Fall Midterm 班級報告生成器
 * 簡化版：專注於核心功能 - 填充模板的學生名單表格
 *
 * 功能：
 * - 從 Google Sheets 讀取學生和班級資料
 * - 複製 A4 橫式模板（2 頁：考試指導 + 班級報告）
 * - 填充右側學生名單表格（6 欄，按學號排序）
 * - 生成獨立的班級報告檔案
 */

// ============================================
// 配置區
// ============================================

const CONFIG = {
  // Google Drive 資源 ID
  templateId: '1D2hSZNI8MQzD_OIeCdEvpqp4EWfO2mrjTCHAQZyx6MM',
  outputFolderId: '1KSyHsy1wUcrT82OjkAMmPFaJmwe-uosi',

  // 表格設定
  studentTableIndex: 1,  // 學生名單表格（右側）
  columnWidths: [90, 100, 140, 140, 80, 120],  // 6 欄寬度

  // 字體大小設定（根據學生人數）
  fontSize: {
    large: 11,   // 1-12 人
    medium: 10,  // 13-18 人
    small: 9     // 19-21 人
  },

  // 檔案命名
  semester: '2526_Fall_Midterm',

  // 批次處理設定
  delayMs: 1000  // 每個檔案間延遲（毫秒）
};

// ============================================
// 主函數
// ============================================

/**
 * 生成所有班級的報告
 * 入口函數：從 Google Sheets 讀取資料並批次生成報告
 */
function generateClassReports() {
  try {
    // 讀取資料
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const studentsSheet = spreadsheet.getSheetByName('Students');
    const classSheet = spreadsheet.getSheetByName('Class');

    if (!studentsSheet || !classSheet) {
      throw new Error('找不到 Students 或 Class 工作表');
    }

    const studentsData = studentsSheet.getDataRange().getValues();
    const classData = classSheet.getDataRange().getValues();

    // 處理資料
    const studentIndexes = getStudentIndexes(studentsData[0]);
    const classIndexes = getClassIndexes(classData[0]);
    const groupedStudents = groupStudentsByClass(studentsData, studentIndexes);

    console.log(`開始處理 ${classData.length - 1} 個班級...`);

    // 生成報告
    const results = [];
    const errors = [];

    for (let i = 1; i < classData.length; i++) {
      const classInfo = {
        name: classData[i][classIndexes.ClassName],
        teacher: classData[i][classIndexes.Tea]
      };

      if (!classInfo.name) continue;

      const students = groupedStudents[classInfo.name];
      if (!students || students.length === 0) {
        console.warn(`班級 ${classInfo.name} 沒有學生資料`);
        errors.push(`${classInfo.name}-${classInfo.teacher} (無學生資料)`);
        continue;
      }

      try {
        const progress = Math.round((i / (classData.length - 1)) * 100);
        console.log(`處理 ${classInfo.name} (${classInfo.teacher})... [${i}/${classData.length - 1}] (${progress}%)`);

        const file = generateSingleReport(classInfo, students, studentIndexes);
        results.push({
          name: classInfo.name,
          teacher: classInfo.teacher,
          url: file.getUrl()
        });

        // 延遲避免 API 限制
        if (i < classData.length - 1) {
          Utilities.sleep(CONFIG.delayMs);
        }

      } catch (e) {
        console.error(`${classInfo.name} 生成失敗: ${e.message}`);
        errors.push(`${classInfo.name}-${classInfo.teacher} (${e.message})`);
      }
    }

    // 生成結果報告
    return formatResults(results, errors);

  } catch (e) {
    console.error(`執行失敗: ${e.message}`);
    throw e;
  }
}

/**
 * 生成單一班級報告
 */
function generateSingleReport(classInfo, students, studentIndexes) {
  // 複製模板
  const template = DriveApp.getFileById(CONFIG.templateId);
  const folder = DriveApp.getFolderById(CONFIG.outputFolderId);

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const fileName = `${classInfo.name}_${classInfo.teacher}_${CONFIG.semester}_${timestamp}`;
  const newFile = template.makeCopy(fileName, folder);

  console.log(`模板複製完成: ${fileName}`);

  // 開啟文件並填充學生表格
  const doc = DocumentApp.openById(newFile.getId());
  const body = doc.getBody();

  fillStudentTable(body, students, studentIndexes);

  doc.saveAndClose();
  console.log(`報告生成完成: ${fileName}`);

  return newFile;
}

/**
 * 填充學生名單表格
 * 核心功能：排序 → 定位 → 清除 → 填充 → 格式化
 */
function fillStudentTable(body, students, studentIndexes) {
  // 步驟 1: 按學號排序
  const sortedStudents = students.slice().sort((a, b) => {
    const idA = String(a[studentIndexes.ID] || '');
    const idB = String(b[studentIndexes.ID] || '');
    return idA.localeCompare(idB, undefined, { numeric: true });
  });

  console.log(`學生已排序，共 ${sortedStudents.length} 人`);

  // 驗證學生人數
  if (sortedStudents.length > 21) {
    console.warn(`⚠️ 警告: 學生人數 ${sortedStudents.length} 超過 21 人，可能超過 2 頁限制`);
  }

  // 步驟 2: 定位學生表格
  const tables = body.getTables();
  if (tables.length < 2) {
    throw new Error(`模板表格數量不足：需要至少 2 個表格，目前只有 ${tables.length} 個`);
  }

  const studentTable = tables[CONFIG.studentTableIndex];

  // 步驟 3: 清除現有資料行（保留標題行）
  while (studentTable.getNumRows() > 1) {
    studentTable.removeRow(1);
  }

  // 步驟 4: 計算字體大小
  const fontSize = calculateFontSize(sortedStudents.length);

  // 步驟 5: 填充學生資料
  sortedStudents.forEach(student => {
    const row = studentTable.appendTableRow();

    // 填入 6 欄資料
    row.getCell(0).setText(String(student[studentIndexes.ID] || ''));
    row.getCell(1).setText(String(student[studentIndexes["Home Room"]] || ''));
    row.getCell(2).setText(String(student[studentIndexes["Chinese Name"]] || ''));
    row.getCell(3).setText(String(student[studentIndexes["English Name"]] || ''));
    row.getCell(4).setText('☐');  // Present 勾選框
    row.getCell(5).setText('☐');  // Signed Paper Returned 勾選框

    // 格式化行
    formatTableRow(row, fontSize);
  });

  // 步驟 6: 格式化標題行
  formatTableRow(studentTable.getRow(0), fontSize, true);

  // 步驟 7: 設定欄寬
  for (let col = 0; col < CONFIG.columnWidths.length; col++) {
    studentTable.setColumnWidth(col, CONFIG.columnWidths[col]);
  }

  // 步驟 8: 設定表格邊框
  studentTable.setAttributes({
    [DocumentApp.Attribute.BORDER_WIDTH]: 1,
    [DocumentApp.Attribute.BORDER_COLOR]: '#000000'
  });

  console.log(`表格填充完成: ${sortedStudents.length} 筆資料`);
}

// ============================================
// 輔助函數
// ============================================

/**
 * 格式化表格行
 */
function formatTableRow(row, fontSize, isHeader = false) {
  for (let col = 0; col < row.getNumCells(); col++) {
    const cell = row.getCell(col);

    // 對齊方式
    cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    const para = cell.getChild(0).asParagraph();
    para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    para.setLineSpacing(1.0);

    // 字體設定
    const text = cell.editAsText();
    text.setFontSize(fontSize);
    text.setFontFamily('Arial');

    // 標題行樣式
    if (isHeader) {
      text.setBold(true);
      cell.setBackgroundColor('#E8E8E8');
    }

    // 邊框
    cell.setAttributes({
      [DocumentApp.Attribute.BORDER_WIDTH]: 1,
      [DocumentApp.Attribute.BORDER_COLOR]: '#000000'
    });
  }
}

/**
 * 根據學生人數計算字體大小
 */
function calculateFontSize(studentCount) {
  if (studentCount > 18) return CONFIG.fontSize.small;
  if (studentCount > 12) return CONFIG.fontSize.medium;
  return CONFIG.fontSize.large;
}

/**
 * 取得學生欄位索引
 */
function getStudentIndexes(headers) {
  const fields = ["English Class", "ID", "Home Room", "Chinese Name", "English Name"];
  const indexes = {};

  fields.forEach(field => {
    const index = headers.indexOf(field);
    if (index === -1) {
      throw new Error(`Students 工作表缺少欄位: ${field}`);
    }
    indexes[field] = index;
  });

  return indexes;
}

/**
 * 取得班級欄位索引
 */
function getClassIndexes(headers) {
  const fields = ["ClassName", "Tea"];
  const indexes = {};

  fields.forEach(field => {
    const index = headers.indexOf(field);
    if (index === -1) {
      throw new Error(`Class 工作表缺少欄位: ${field}`);
    }
    indexes[field] = index;
  });

  return indexes;
}

/**
 * 依 English Class 分組學生
 */
function groupStudentsByClass(studentsData, indexes) {
  const grouped = {};

  for (let i = 1; i < studentsData.length; i++) {
    const className = studentsData[i][indexes["English Class"]];
    if (!className) continue;

    if (!grouped[className]) {
      grouped[className] = [];
    }
    grouped[className].push(studentsData[i]);
  }

  return grouped;
}

/**
 * 格式化結果報告
 */
function formatResults(results, errors) {
  let message = `✅ 班級報告生成完成！\n\n`;
  message += `📊 成功: ${results.length} 個班級\n`;

  if (errors.length > 0) {
    message += `❌ 失敗: ${errors.length} 個班級\n`;
  }

  if (results.length > 0) {
    message += `\n📁 生成的檔案：\n`;
    results.forEach((item, i) => {
      message += `${i + 1}. ${item.name} (${item.teacher})\n`;
    });
    message += `\n🔗 資料夾: https://drive.google.com/drive/folders/${CONFIG.outputFolderId}`;
  }

  if (errors.length > 0) {
    message += `\n\n⚠️ 失敗的班級：\n`;
    errors.forEach((err, i) => {
      message += `${i + 1}. ${err}\n`;
    });
  }

  return message;
}

// ============================================
// UI 選單
// ============================================

/**
 * 創建自訂選單
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('班級報告')
    .addItem('生成 2526 Fall Midterm 報告', 'runReportGeneration')
    .addItem('🧪 測試單一班級', 'testSingleClass')
    .addToUi();
}

/**
 * 從選單執行
 */
function runReportGeneration() {
  try {
    const result = generateClassReports();
    SpreadsheetApp.getUi().alert(result);
  } catch (e) {
    SpreadsheetApp.getUi().alert(`❌ 執行失敗: ${e.message}\n\n請檢查日誌獲取詳細資訊`);
    console.error('執行錯誤:', e);
  }
}

/**
 * 測試模式：只生成第一個班級的報告
 * 用於快速測試和驗證功能
 */
function testSingleClass() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const studentsSheet = spreadsheet.getSheetByName('Students');
    const classSheet = spreadsheet.getSheetByName('Class');

    if (!studentsSheet || !classSheet) {
      throw new Error('找不到 Students 或 Class 工作表');
    }

    const studentsData = studentsSheet.getDataRange().getValues();
    const classData = classSheet.getDataRange().getValues();

    // 處理資料
    const studentIndexes = getStudentIndexes(studentsData[0]);
    const classIndexes = getClassIndexes(classData[0]);
    const groupedStudents = groupStudentsByClass(studentsData, studentIndexes);

    // 只處理第一個班級
    if (classData.length < 2) {
      throw new Error('Class 工作表沒有資料');
    }

    const classInfo = {
      name: classData[1][classIndexes.ClassName],
      teacher: classData[1][classIndexes.Tea]
    };

    const students = groupedStudents[classInfo.name];
    if (!students || students.length === 0) {
      throw new Error(`班級 ${classInfo.name} 沒有學生資料`);
    }

    console.log(`🧪 測試模式：生成 ${classInfo.name} (${classInfo.teacher}) 的報告`);
    const file = generateSingleReport(classInfo, students, studentIndexes);

    const message = `✅ 測試成功！\n\n班級: ${classInfo.name}\n老師: ${classInfo.teacher}\n學生數: ${students.length}\n\n🔗 查看檔案:\n${file.getUrl()}`;
    SpreadsheetApp.getUi().alert(message);

  } catch (e) {
    SpreadsheetApp.getUi().alert(`❌ 測試失敗: ${e.message}\n\n請檢查日誌獲取詳細資訊`);
    console.error('測試錯誤:', e);
  }
}
