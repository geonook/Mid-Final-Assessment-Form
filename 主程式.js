/**
 * 2526 Fall Midterm v2 班級報告生成器
 * 功能升級：佔位符替換 + GradeBand 子資料夾分類
 *
 * 功能：
 * - 從 Google Sheets 讀取學生和班級資料（含完整考試資訊）
 * - 複製 A4 橫式模板（2 頁：考試指導 + 班級報告）
 * - 第一頁：替換所有佔位符 {{...}} 為對應資料
 * - 第二頁：填充學生名單表格（6 欄，按學號排序）
 * - 依 GradeBand 分類儲存到子資料夾
 * - 生成獨立的班級報告檔案
 */

// ============================================
// 配置區
// ============================================

const CONFIG = {
  // Google Sheets 試算表 ID（用於 Apps Script 編輯器直接執行）
  spreadsheetId: '1bo3xsXw0u8Wwbo6ALvPe9idKDVCjhtAVKAKfH8azhJE',

  // Google Drive 資源 ID
  templateId: '1D2hSZNI8MQzD_OIeCdEvpqp4EWfO2mrjTCHAQZyx6MM',
  outputFolderId: '1KSyHsy1wUcrT82OjkAMmPFaJmwe-uosi',

  // 表格設定
  studentTableIndex: 2,  // 學生名單表格（第三個表格，tables[2]）
  columnWidths: [75, 75, 95, 100, 85, 155],  // 6 欄寬度（總計 585pt ≈ 20.6cm）
  // [Student ID, Homeroom, Chinese Name, English Name, Present, Signed Paper]
  // 欄寬足以讓標題完整顯示不需換行

  // 字體大小設定（固定 10pt）
  fontSize: {
    large: 10,   // 固定 10pt
    medium: 10,  // 固定 10pt
    small: 10    // 固定 10pt
  },

  // 檔案命名
  semester: '2526_Fall_Midterm_v2',

  // 批次處理設定
  delayMs: 1000,  // 每個檔案間延遲（毫秒）

  // 佔位符欄位映射（Class 工作表欄位 → 模板佔位符）
  placeholderFields: [
    'GradeBand',    // {{GradeBand}} - 年級段
    'Duration',     // {{Duration}} - 考試時長
    'Periods',      // {{Periods}} - 節次
    'Self-Study',   // {{Self-Study}} - 自習時間
    'Preparation',  // {{Preparation}} - 準備時間
    'ExamTime',     // {{ExamTime}} - 考試時間
    'Level',        // {{Level}} - 等級
    'Classroom',    // {{Classroom}} - 教室
    'Proctor',      // {{Proctor}} - 監考教師
    'Subject',      // {{Subject}} - 科目
    'ClassName',    // {{ClassName}} - 班級名稱
    'Teacher',      // {{Teacher}} - 授課教師
    'Count',        // {{Count}} - 試卷總數
    'Students'      // {{Students}} - 總學生數
  ]
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

    // 排序班級：先按 GradeBand，再按 ClassName 字母順序
    const sortedClassList = sortClassDataByGradeBandAndName(classData, classIndexes);

    console.log(`開始處理 ${sortedClassList.length} 個班級（已按 GradeBand 和 ClassName 排序）...`);

    // 生成報告
    const results = [];
    const errors = [];

    for (let i = 0; i < sortedClassList.length; i++) {
      const classInfo = sortedClassList[i];

      const students = groupedStudents[classInfo.ClassName];
      if (!students || students.length === 0) {
        console.warn(`班級 ${classInfo.ClassName} 沒有學生資料`);
        errors.push(`${classInfo.ClassName}-${classInfo.Teacher} (無學生資料)`);
        continue;
      }

      try {
        const progress = Math.round(((i + 1) / sortedClassList.length) * 100);
        console.log(`處理 ${classInfo.ClassName} (${classInfo.Teacher})... [${i + 1}/${sortedClassList.length}] (${progress}%)`);

        const file = generateSingleReport(classInfo, students, studentIndexes);
        results.push({
          name: classInfo.ClassName,
          teacher: classInfo.Teacher,
          gradeBand: classInfo.GradeBand,
          url: file.getUrl()
        });

        // 延遲避免 API 限制
        if (i < sortedClassList.length - 1) {
          Utilities.sleep(CONFIG.delayMs);
        }

      } catch (e) {
        console.error(`${classInfo.ClassName} 生成失敗: ${e.message}`);
        errors.push(`${classInfo.ClassName}-${classInfo.Teacher} (${e.message})`);
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
 * v2: 支援佔位符替換 + GradeBand 子資料夾
 *
 * @param {Object} classData - 完整的班級資料物件（包含所有欄位）
 * @param {Array} students - 該班級的學生陣列
 * @param {Object} studentIndexes - 學生欄位索引
 * @return {File} 生成的檔案物件
 */
function generateSingleReport(classData, students, studentIndexes) {
  // 取得或建立 GradeBand 子資料夾
  const targetFolder = getOrCreateSubfolder(CONFIG.outputFolderId, classData.GradeBand);

  // 複製模板
  const template = DriveApp.getFileById(CONFIG.templateId);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const fileName = `${classData.ClassName}_${classData.Teacher}_${CONFIG.semester}_${timestamp}`;
  const newFile = template.makeCopy(fileName, targetFolder);

  console.log(`模板複製完成: ${fileName}`);

  // ============================================
  // 階段 1: 填充學生表格（第二頁）- 先執行
  // 此時文件結構完好，表格操作不會失敗
  // ============================================
  let doc = DocumentApp.openById(newFile.getId());
  let body = doc.getBody();
  fillStudentTable(body, students, studentIndexes);
  doc.saveAndClose();
  console.log(`階段 1 完成：學生表格填充完成`);

  // ============================================
  // 階段 2: 替換佔位符（第一頁）- 後執行
  // 即使破壞文件結構，學生表格已完成不受影響
  // ============================================
  doc = DocumentApp.openById(newFile.getId());
  body = doc.getBody();
  replacePlaceholders(body, classData);
  doc.saveAndClose();
  console.log(`階段 2 完成：佔位符替換完成`);

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
  console.log(`表格初始行數: ${studentTable.getNumRows()}`);

  // 步驟 3: 計算字體大小
  const fontSize = calculateFontSize(sortedStudents.length);

  // 步驟 3.5: 先設定欄寬（在填充前，確保生效）
  const totalWidth = CONFIG.columnWidths.reduce((a, b) => a + b, 0);
  console.log(`設定欄寬: [${CONFIG.columnWidths.join(', ')}] (總計: ${totalWidth} pt)`);
  for (let col = 0; col < CONFIG.columnWidths.length; col++) {
    studentTable.setColumnWidth(col, CONFIG.columnWidths[col]);
  }

  // 步驟 4: 填充學生資料（Solution C: 直接覆寫現有行，避免 appendTableRow 的 0-cell bug）
  sortedStudents.forEach((student, index) => {
    const rowIndex = index + 1; // 跳過標題行（索引 0）
    let row;

    if (rowIndex < studentTable.getNumRows()) {
      // 使用現有行（模板預設有 24 行，不創建新行避免 cell 數量為 0 的問題）
      row = studentTable.getRow(rowIndex);
    } else {
      // 如果學生數超過模板預設行數，創建新行
      console.log(`創建新行 ${rowIndex}（超過模板預設行數）`);
      row = studentTable.insertTableRow(rowIndex);

      // 手動確保有 6 個 cells（解決 appendTableRow 創建 0 cell 的問題）
      while (row.getNumCells() < 6) {
        row.appendTableCell('');
      }
    }

    // 填入 6 欄資料
    const values = [
      String(student[studentIndexes.ID] || ''),
      String(student[studentIndexes["Home Room"]] || ''),
      String(student[studentIndexes["Chinese Name"]] || ''),
      String(student[studentIndexes["English Name"]] || ''),
      '☐',  // Present 勾選框
      '☐'   // Signed Paper Returned 勾選框
    ];

    for (let col = 0; col < 6; col++) {
      const cell = row.getCell(col);
      cell.setText(values[col]);
    }

    // 格式化行
    formatTableRow(row, fontSize);
  });

  // 步驟 5: 清除多餘的行（如果學生數少於模板預設行數）
  const targetRows = sortedStudents.length + 1; // +1 for header
  while (studentTable.getNumRows() > targetRows) {
    const lastRowIndex = studentTable.getNumRows() - 1;
    studentTable.removeRow(lastRowIndex);
  }

  // 步驟 6: 格式化標題行
  formatTableRow(studentTable.getRow(0), fontSize, true);

  // 步驟 7: 設定表格邊框（欄寬已在步驟 3.5 設定）
  studentTable.setAttributes({
    [DocumentApp.Attribute.BORDER_WIDTH]: 1,
    [DocumentApp.Attribute.BORDER_COLOR]: '#000000'
  });

  console.log(`表格填充完成: ${sortedStudents.length} 筆資料，最終行數: ${studentTable.getNumRows()}`);
}

// ============================================
// 輔助函數
// ============================================

/**
 * 替換文件中的所有佔位符
 * 使用 findText() 搜尋並替換所有 {{FieldName}} 格式的佔位符
 * 支援有格式的文字（如黃色背景標記）
 *
 * @param {Body} body - Google Docs 文件主體
 * @param {Object} classData - 班級資料物件（包含所有欄位）
 */
function replacePlaceholders(body, classData) {
  console.log('開始替換佔位符...');

  CONFIG.placeholderFields.forEach(field => {
    const placeholder = `{{${field}}}`;
    let value = String(classData[field] || '');

    // 確保值不為 undefined 或 null
    if (value === 'undefined' || value === 'null') {
      value = '';
    }

    // 跳脫正則表達式特殊字元（大括號）
    const pattern = '\\{\\{' + field + '\\}\\}';

    try {
      let searchResult = body.findText(pattern);
      let replacedCount = 0;

      // 逐一找到所有匹配的佔位符並替換
      while (searchResult !== null) {
        const foundElement = searchResult.getElement();
        const start = searchResult.getStartOffset();
        const end = searchResult.getEndOffsetInclusive();

        // 刪除找到的佔位符文字，並插入替換值
        const textElement = foundElement.asText();
        textElement.deleteText(start, end);
        textElement.insertText(start, value);

        replacedCount++;

        // 繼續搜尋下一個匹配項
        searchResult = body.findText(pattern, searchResult);
      }

      if (replacedCount > 0) {
        console.log(`  替換 ${placeholder} → ${value} (${replacedCount} 處)`);
      } else {
        console.log(`  未找到 ${placeholder}`);
      }

    } catch (e) {
      console.warn(`  ⚠️ 替換 ${placeholder} 時發生錯誤: ${e.message}`);
    }
  });

  console.log('佔位符替換完成');
}

/**
 * 取得或建立 GradeBand 子資料夾
 * 清理 GradeBand 名稱中的特殊字元（空格 → 底線，移除撇號）
 *
 * @param {String} parentFolderId - 父資料夾 ID
 * @param {String} gradeBand - GradeBand 值（例如：G3 IT's）
 * @return {Folder} 子資料夾物件
 */
function getOrCreateSubfolder(parentFolderId, gradeBand) {
  const parentFolder = DriveApp.getFolderById(parentFolderId);

  // 清理 GradeBand 名稱：空格 → 底線，移除撇號和其他特殊字元
  const cleanFolderName = String(gradeBand)
    .replace(/\s+/g, '_')      // 空格 → 底線
    .replace(/['''`]/g, '')    // 移除撇號
    .replace(/[^\w\-]/g, '_'); // 其他特殊字元 → 底線

  console.log(`  檢查子資料夾: ${cleanFolderName}`);

  // 檢查子資料夾是否已存在
  const folders = parentFolder.getFoldersByName(cleanFolderName);

  if (folders.hasNext()) {
    const folder = folders.next();
    console.log(`  ✓ 使用現有子資料夾: ${cleanFolderName}`);
    return folder;
  } else {
    const newFolder = parentFolder.createFolder(cleanFolderName);
    console.log(`  ✓ 建立新子資料夾: ${cleanFolderName}`);
    return newFolder;
  }
}

/**
 * 格式化表格行
 */
function formatTableRow(row, fontSize, isHeader = false) {
  for (let col = 0; col < row.getNumCells(); col++) {
    const cell = row.getCell(col);

    // 對齊方式
    cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);

    // 確保 cell 有段落元素（如果是空的，先創建一個）
    if (cell.getNumChildren() === 0) {
      cell.appendParagraph('');
    }

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
 * v2: 擴充支援完整考試資訊欄位
 */
function getClassIndexes(headers) {
  const fields = [
    "ClassName",     // 班級名稱
    "Grade",         // 年級
    "Teacher",       // 授課教師（原 Tea 欄位，保持向後相容）
    "Level",         // 等級
    "Classroom",     // 教室
    "GradeBand",     // 年級段
    "Duration",      // 考試時長
    "Periods",       // 節次
    "Self-Study",    // 自習時間
    "Preparation",   // 準備時間
    "ExamTime",      // 考試時間
    "Proctor",       // 監考教師
    "Subject",       // 科目
    "Count",         // 試卷總數
    "Students"       // 總學生數
  ];

  const indexes = {};

  fields.forEach(field => {
    let index = headers.indexOf(field);

    // 向後相容：Tea 欄位可能被改名為 Teacher
    if (field === "Teacher" && index === -1) {
      index = headers.indexOf("Tea");
    }

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
 * v2: 新增 GradeBand 分類顯示
 */
function formatResults(results, errors) {
  let message = `✅ 班級報告生成完成！(v2)\n\n`;
  message += `📊 成功: ${results.length} 個班級\n`;

  if (errors.length > 0) {
    message += `❌ 失敗: ${errors.length} 個班級\n`;
  }

  if (results.length > 0) {
    message += `\n📁 生成的檔案（依 GradeBand 分類）：\n`;

    // 依 GradeBand 分組顯示
    const groupedByGradeBand = {};
    results.forEach(item => {
      const gb = item.gradeBand || 'Unknown';
      if (!groupedByGradeBand[gb]) {
        groupedByGradeBand[gb] = [];
      }
      groupedByGradeBand[gb].push(item);
    });

    Object.keys(groupedByGradeBand).sort().forEach(gradeBand => {
      message += `\n  📂 ${gradeBand}:\n`;
      groupedByGradeBand[gradeBand].forEach(item => {
        message += `    • ${item.name} (${item.teacher})\n`;
      });
    });

    message += `\n🔗 主資料夾: https://drive.google.com/drive/folders/${CONFIG.outputFolderId}`;
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
    .addItem('步驟 1: 生成所有 Google Docs', 'runReportGeneration')
    .addItem('步驟 2: 合併為 PDF（按 GradeBand）', 'runMergeDocsToPDF')
    .addSeparator()
    .addItem('🚀 一鍵執行（Docs + PDF）', 'runGenerateAndMergePDF')
    .addSeparator()
    .addItem('🧪 測試單一班級（v2）', 'testSingleClass')
    .addToUi();
}

/**
 * 從選單執行 - 生成所有 Google Docs
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
 * 從選單執行 - 合併 Docs 為 PDF（按 GradeBand）
 */
function runMergeDocsToPDF() {
  try {
    const result = mergeDocsToPDFByGradeBand();
    SpreadsheetApp.getUi().alert(result);
  } catch (e) {
    SpreadsheetApp.getUi().alert(`❌ 執行失敗: ${e.message}\n\n請檢查日誌獲取詳細資訊`);
    console.error('執行錯誤:', e);
  }
}

/**
 * 從選單執行 - 一鍵執行（生成 Docs + 合併 PDF）
 */
function runGenerateAndMergePDF() {
  try {
    const result = generateAndMergePDFReports();
    SpreadsheetApp.getUi().alert(result);
  } catch (e) {
    SpreadsheetApp.getUi().alert(`❌ 執行失敗: ${e.message}\n\n請檢查日誌獲取詳細資訊`);
    console.error('執行錯誤:', e);
  }
}

/**
 * 測試模式：只生成第一個班級的報告
 * v2: 支援完整資料結構和佔位符替換測試
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

    // 建立完整的班級資料物件
    const classRow = classData[1];
    const classInfo = {
      ClassName: classRow[classIndexes.ClassName],
      Grade: classRow[classIndexes.Grade],
      Teacher: classRow[classIndexes.Teacher],
      Level: classRow[classIndexes.Level],
      Classroom: classRow[classIndexes.Classroom],
      GradeBand: classRow[classIndexes.GradeBand],
      Duration: classRow[classIndexes.Duration],
      Periods: classRow[classIndexes.Periods],
      'Self-Study': classRow[classIndexes['Self-Study']],
      Preparation: classRow[classIndexes.Preparation],
      ExamTime: classRow[classIndexes.ExamTime],
      Proctor: classRow[classIndexes.Proctor],
      Subject: classRow[classIndexes.Subject],
      Count: classRow[classIndexes.Count],
      Students: classRow[classIndexes.Students]
    };

    const students = groupedStudents[classInfo.ClassName];
    if (!students || students.length === 0) {
      throw new Error(`班級 ${classInfo.ClassName} 沒有學生資料`);
    }

    console.log(`🧪 測試模式 v2：生成 ${classInfo.ClassName} (${classInfo.Teacher}) 的報告`);
    const file = generateSingleReport(classInfo, students, studentIndexes);

    const message = `✅ 測試成功！(v2)\n\n班級: ${classInfo.ClassName}\n老師: ${classInfo.Teacher}\nGradeBand: ${classInfo.GradeBand}\n學生數: ${students.length}\n\n🔗 查看檔案:\n${file.getUrl()}`;
    SpreadsheetApp.getUi().alert(message);

  } catch (e) {
    SpreadsheetApp.getUi().alert(`❌ 測試失敗: ${e.message}\n\n請檢查日誌獲取詳細資訊`);
    console.error('測試錯誤:', e);
  }
}

// ============================================
// 模板分析工具
// ============================================

/**
 * 分析模板結構 - 檢查表格數量和位置
 * 用於診斷 studentTableIndex 設定
 */
function ANALYZE_TEMPLATE() {
  try {
    console.log('========================================');
    console.log('🔍 模板結構分析');
    console.log('========================================');

    const doc = DocumentApp.openById(CONFIG.templateId);
    const body = doc.getBody();
    const tables = body.getTables();

    console.log(`\n📊 表格總數: ${tables.length}\n`);

    tables.forEach((table, index) => {
      console.log(`--- 表格 ${index} (tables[${index}]) ---`);
      console.log(`  行數: ${table.getNumRows()}`);

      const firstRow = table.getRow(0);
      console.log(`  欄數: ${firstRow.getNumCells()}`);

      // 讀取第一行內容
      const headers = [];
      for (let col = 0; col < Math.min(firstRow.getNumCells(), 6); col++) {
        const text = firstRow.getCell(col).getText().trim();
        headers.push(text.substring(0, 20)); // 只顯示前20字元
      }
      console.log(`  第一行: ${headers.join(' | ')}`);

      // 讀取第一列（左側）的內容（前5行）
      const leftColumn = [];
      for (let row = 0; row < Math.min(table.getNumRows(), 5); row++) {
        const text = table.getRow(row).getCell(0).getText().trim();
        leftColumn.push(text.substring(0, 30));
      }
      console.log(`  第一列內容:\n    ${leftColumn.join('\n    ')}`);
      console.log('');
    });

    console.log('========================================');
    console.log('💡 判斷建議:');
    console.log('========================================');

    // 尋找學生表格
    let studentTableFound = false;
    tables.forEach((table, index) => {
      const firstRow = table.getRow(0);
      const firstCell = firstRow.getCell(0).getText().trim();
      const numCols = firstRow.getNumCells();

      if (firstCell.includes('Student') || firstCell.includes('ID') || numCols === 6) {
        console.log(`✅ 表格 ${index} 可能是學生表格:`);
        if (firstCell.includes('Student') || firstCell.includes('ID')) {
          console.log(`   - 第一格包含 "Student" 或 "ID"`);
        }
        if (numCols === 6) {
          console.log(`   - 有 6 欄（符合學生表格結構）`);
        }
        console.log(`   → 建議設定: CONFIG.studentTableIndex = ${index}`);
        console.log('');
        studentTableFound = true;
      }
    });

    if (!studentTableFound) {
      console.log('⚠️ 無法自動識別學生表格，請手動檢查');
    }

    console.log('========================================');
    console.log(`\n目前設定: CONFIG.studentTableIndex = ${CONFIG.studentTableIndex}`);
    console.log('========================================');

  } catch (e) {
    console.error('❌ 分析失敗:', e.message);
    console.error(e.stack);
  }
}

// ============================================
// Apps Script 編輯器直接執行函數
// ============================================

/**
 * 🚀 快速測試入口 - 在 Apps Script 編輯器直接執行
 *
 * 使用方法：
 * 1. 打開 Apps Script 編輯器
 * 2. 在頂部函數選擇器選擇「RUN」
 * 3. 點擊「執行」按鈕 ▶
 * 4. 查看執行日誌（View → Logs 或 Ctrl+Enter / Cmd+Enter）
 *
 * 功能：
 * - 使用 CONFIG.spreadsheetId 讀取試算表資料
 * - 生成第一個班級的報告（快速測試用）
 * - 驗證佔位符替換、學生名單填入、GradeBand 子資料夾等功能
 *
 * 執行結果會顯示在日誌中，包含：
 * - 班級資訊
 * - 生成的檔案連結
 * - 儲存位置
 */
function RUN() {
  try {
    console.log('========================================');
    console.log('🚀 RUN() - Apps Script 編輯器測試模式');
    console.log('========================================');

    // 使用 CONFIG 中的試算表 ID
    console.log(`📊 讀取試算表: ${CONFIG.spreadsheetId}`);
    const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
    const studentsSheet = spreadsheet.getSheetByName('Students');
    const classSheet = spreadsheet.getSheetByName('Class');

    if (!studentsSheet || !classSheet) {
      throw new Error('找不到 Students 或 Class 工作表');
    }

    // 讀取資料
    const studentsData = studentsSheet.getDataRange().getValues();
    const classData = classSheet.getDataRange().getValues();

    console.log(`✓ Students 工作表: ${studentsData.length - 1} 筆學生資料`);
    console.log(`✓ Class 工作表: ${classData.length - 1} 個班級`);

    // 處理索引
    const studentIndexes = getStudentIndexes(studentsData[0]);
    const classIndexes = getClassIndexes(classData[0]);
    const groupedStudents = groupStudentsByClass(studentsData, studentIndexes);

    // 只處理第一個班級
    if (classData.length < 2) {
      throw new Error('Class 工作表沒有資料');
    }

    const classRow = classData[1];
    const classInfo = {
      ClassName: classRow[classIndexes.ClassName],
      Grade: classRow[classIndexes.Grade],
      Teacher: classRow[classIndexes.Teacher],
      Level: classRow[classIndexes.Level],
      Classroom: classRow[classIndexes.Classroom],
      GradeBand: classRow[classIndexes.GradeBand],
      Duration: classRow[classIndexes.Duration],
      Periods: classRow[classIndexes.Periods],
      'Self-Study': classRow[classIndexes['Self-Study']],
      Preparation: classRow[classIndexes.Preparation],
      ExamTime: classRow[classIndexes.ExamTime],
      Proctor: classRow[classIndexes.Proctor],
      Subject: classRow[classIndexes.Subject],
      Count: classRow[classIndexes.Count],
      Students: classRow[classIndexes.Students]
    };

    const students = groupedStudents[classInfo.ClassName];

    if (!students || students.length === 0) {
      throw new Error(`班級 ${classInfo.ClassName} 沒有學生資料`);
    }

    console.log('');
    console.log('📋 班級資訊:');
    console.log(`  班級: ${classInfo.ClassName}`);
    console.log(`  老師: ${classInfo.Teacher}`);
    console.log(`  GradeBand: ${classInfo.GradeBand}`);
    console.log(`  等級: ${classInfo.Level}`);
    console.log(`  教室: ${classInfo.Classroom}`);
    console.log(`  科目: ${classInfo.Subject}`);
    console.log(`  監考教師: ${classInfo.Proctor}`);
    console.log(`  學生數: ${students.length}`);
    console.log('');

    console.log('🔄 開始生成報告...');

    // 生成報告
    const file = generateSingleReport(classInfo, students, studentIndexes);

    console.log('');
    console.log('========================================');
    console.log('✅ 測試完成！');
    console.log('========================================');
    console.log(`📄 檔案名稱: ${file.getName()}`);
    console.log(`🔗 檔案連結: ${file.getUrl()}`);
    console.log(`📂 儲存位置: 輸出資料夾/${classInfo.GradeBand}/`);
    console.log('');
    console.log('💡 提示: 請開啟檔案檢查：');
    console.log('   • 第一頁的 {{...}} 佔位符是否都被正確替換');
    console.log('   • 第二頁的學生名單是否正確填入');
    console.log('   • 檔案是否儲存在正確的 GradeBand 子資料夾');
    console.log('========================================');

  } catch (e) {
    console.error('');
    console.error('========================================');
    console.error('❌ 執行失敗');
    console.error('========================================');
    console.error(`錯誤訊息: ${e.message}`);
    console.error(`錯誤堆疊: ${e.stack}`);
    console.error('========================================');
    throw e;
  }
}

// ============================================
// Apps Script 編輯器直接執行函數（無需 Google Sheets）
// ============================================

/**
 * 🚀 快速批次執行入口 - 在 Apps Script 編輯器直接執行
 *
 * 使用方法：
 * 1. 打開 Apps Script 編輯器
 * 2. 在頂部函數選擇器選擇「RUN_FULL_BATCH」
 * 3. 點擊「執行」按鈕 ▶
 * 4. 查看執行日誌（View → Logs 或 Ctrl+Enter / Cmd+Enter）
 *
 * 功能：
 * - 階段 1: 生成所有 168 個班級的 Google Docs（約 5-8 分鐘）
 * - 階段 2: 按 GradeBand 合併為 PDF（約 5-7 分鐘）
 * - 總計約 10-15 分鐘
 *
 * 執行結果會顯示在日誌中，包含：
 * - 每個階段的進度
 * - 成功/失敗統計
 * - 生成的 PDF 檔案清單
 * - 輸出資料夾連結
 */
function RUN_FULL_BATCH() {
  try {
    console.log('========================================');
    console.log('🚀 RUN_FULL_BATCH() - 完整批次執行');
    console.log('========================================');
    console.log('');
    console.log('⚠️ 預計執行時間：10-15 分鐘');
    console.log('⚠️ 請勿關閉此視窗');
    console.log('');

    // 階段 1: 生成所有 Google Docs
    console.log('========================================');
    console.log('階段 1: 生成所有班級的 Google Docs 檔案');
    console.log('========================================');

    const docsResult = generateClassReports();
    console.log('');
    console.log(docsResult);
    console.log('');

    console.log('✅ 階段 1 完成，休息 5 秒後繼續...');
    console.log('');
    Utilities.sleep(5000);

    // 階段 2: 合併為 PDF
    console.log('========================================');
    console.log('階段 2: 按 GradeBand 合併為 PDF');
    console.log('========================================');

    const pdfResult = mergeDocsToPDFByGradeBand();
    console.log('');
    console.log(pdfResult);
    console.log('');

    // 最終報告
    console.log('========================================');
    console.log('✅ 全部完成！');
    console.log('========================================');
    console.log('');
    console.log('📁 輸出資料夾:');
    console.log(`https://drive.google.com/drive/folders/${CONFIG.outputFolderId}`);
    console.log('');
    console.log('💡 請開啟輸出資料夾檢查：');
    console.log('   • 每個 GradeBand 子資料夾都有對應的 PDF');
    console.log('   • PDF 檔案包含該 GradeBand 的所有班級');
    console.log('   • 班級按字母順序排列');
    console.log('========================================');

  } catch (e) {
    console.error('');
    console.error('========================================');
    console.error('❌ 執行失敗');
    console.error('========================================');
    console.error(`錯誤訊息: ${e.message}`);
    console.error(`錯誤堆疊: ${e.stack}`);
    console.error('========================================');
    throw e;
  }
}

/**
 * 📄 只生成 Google Docs - 在 Apps Script 編輯器直接執行
 *
 * 使用方法：
 * 1. 打開 Apps Script 編輯器
 * 2. 在頂部函數選擇器選擇「RUN_DOCS_ONLY」
 * 3. 點擊「執行」按鈕 ▶
 *
 * 功能：
 * - 生成所有 168 個班級的 Google Docs
 * - 不執行 PDF 合併（可稍後手動執行）
 * - 執行時間約 5-8 分鐘
 */
function RUN_DOCS_ONLY() {
  try {
    console.log('========================================');
    console.log('📄 RUN_DOCS_ONLY() - 只生成 Google Docs');
    console.log('========================================');
    console.log('');

    const result = generateClassReports();

    console.log('');
    console.log(result);
    console.log('');
    console.log('========================================');
    console.log('✅ Google Docs 生成完成！');
    console.log('========================================');
    console.log('');
    console.log('💡 下一步：');
    console.log('   執行 RUN_PDF_ONLY() 合併為 PDF');
    console.log('========================================');

  } catch (e) {
    console.error('');
    console.error('========================================');
    console.error('❌ 執行失敗');
    console.error('========================================');
    console.error(`錯誤訊息: ${e.message}`);
    console.error(`錯誤堆疊: ${e.stack}`);
    console.error('========================================');
    throw e;
  }
}

/**
 * 📑 只合併 PDF - 在 Apps Script 編輯器直接執行
 *
 * 使用方法：
 * 1. 確保已經執行過 RUN_DOCS_ONLY() 或 generateClassReports()
 * 2. 打開 Apps Script 編輯器
 * 3. 在頂部函數選擇器選擇「RUN_PDF_ONLY」
 * 4. 點擊「執行」按鈕 ▶
 *
 * 功能：
 * - 讀取輸出資料夾中的 Google Docs
 * - 按 GradeBand 合併為 PDF
 * - 執行時間約 5-7 分鐘
 */
function RUN_PDF_ONLY() {
  try {
    console.log('========================================');
    console.log('📑 RUN_PDF_ONLY() - 只合併 PDF');
    console.log('========================================');
    console.log('');

    const result = mergeDocsToPDFByGradeBand();

    console.log('');
    console.log(result);
    console.log('');
    console.log('========================================');
    console.log('✅ PDF 合併完成！');
    console.log('========================================');
    console.log('');
    console.log('📁 輸出資料夾:');
    console.log(`https://drive.google.com/drive/folders/${CONFIG.outputFolderId}`);
    console.log('========================================');

  } catch (e) {
    console.error('');
    console.error('========================================');
    console.error('❌ 執行失敗');
    console.error('========================================');
    console.error(`錯誤訊息: ${e.message}`);
    console.error(`錯誤堆疊: ${e.stack}`);
    console.error('========================================');
    throw e;
  }
}

// ============================================
// PDF 合併功能（按 GradeBand 分組）
// ============================================

/**
 * 按 GradeBand 和 ClassName 排序班級資料
 * @param {Array} classData - 原始班級資料陣列（包含標題行）
 * @param {Object} classIndexes - 班級欄位索引
 * @return {Array} 排序後的班級資料物件陣列
 */
function sortClassDataByGradeBandAndName(classData, classIndexes) {
  // 1. 提取所有班級資料（跳過標題行）
  const classList = [];
  for (let i = 1; i < classData.length; i++) {
    const classRow = classData[i];
    if (!classRow[classIndexes.ClassName]) continue;

    classList.push({
      ClassName: classRow[classIndexes.ClassName],
      Grade: classRow[classIndexes.Grade],
      Teacher: classRow[classIndexes.Teacher],
      Level: classRow[classIndexes.Level],
      Classroom: classRow[classIndexes.Classroom],
      GradeBand: classRow[classIndexes.GradeBand],
      Duration: classRow[classIndexes.Duration],
      Periods: classRow[classIndexes.Periods],
      'Self-Study': classRow[classIndexes['Self-Study']],
      Preparation: classRow[classIndexes.Preparation],
      ExamTime: classRow[classIndexes.ExamTime],
      Proctor: classRow[classIndexes.Proctor],
      Subject: classRow[classIndexes.Subject],
      Count: classRow[classIndexes.Count],
      Students: classRow[classIndexes.Students]
    });
  }

  // 2. 排序：先按 GradeBand，再按 ClassName 字母順序
  classList.sort((a, b) => {
    // 先比較 GradeBand
    const gradeBandCompare = a.GradeBand.localeCompare(b.GradeBand);
    if (gradeBandCompare !== 0) return gradeBandCompare;

    // GradeBand 相同，再比較 ClassName 字母順序
    return a.ClassName.localeCompare(b.ClassName);
  });

  console.log(`班級已排序：先按 GradeBand，再按 ClassName 字母順序`);
  return classList;
}

/**
 * 將多個 Google Docs 合併為單一 PDF
 * @param {Array} docsList - Google Docs File 物件陣列
 * @param {String} gradeBandName - GradeBand 名稱（資料夾名稱，例如 "G1_LTs"）
 * @param {String} originalGradeBand - 原始 GradeBand 名稱（例如 "G1 LT's"）
 * @param {Folder} targetFolder - 目標資料夾物件
 * @return {File} 生成的 PDF File 物件
 */
function mergeDocsIntoPDF(docsList, gradeBandName, originalGradeBand, targetFolder) {
  console.log(`  開始合併 ${docsList.length} 個文件...`);

  // 1. 創建臨時合併文件
  const mergedDoc = DocumentApp.create(`${gradeBandName}_Merged_Temp`);
  const mergedBody = mergedDoc.getBody();

  // 2. 清空預設內容
  mergedBody.clear();

  // 3. 遍歷每個 Docs，複製內容
  docsList.forEach((file, index) => {
    console.log(`    複製 ${index + 1}/${docsList.length}: ${file.getName()}`);

    try {
      // 3a. 開啟來源文件
      const sourceDoc = DocumentApp.openById(file.getId());
      const sourceBody = sourceDoc.getBody();

      // 3b. 複製所有元素到合併文件
      const numElements = sourceBody.getNumChildren();
      for (let i = 0; i < numElements; i++) {
        const element = sourceBody.getChild(i);
        const elementType = element.getType();

        // 複製不同類型的元素
        if (elementType === DocumentApp.ElementType.PARAGRAPH) {
          const para = element.asParagraph().copy();
          mergedBody.appendParagraph(para);
        } else if (elementType === DocumentApp.ElementType.TABLE) {
          const table = element.asTable().copy();
          mergedBody.appendTable(table);
        } else if (elementType === DocumentApp.ElementType.LIST_ITEM) {
          const listItem = element.asListItem().copy();
          mergedBody.appendListItem(listItem);
        } else if (elementType === DocumentApp.ElementType.PAGE_BREAK) {
          // 跳過原有的分頁符號，稍後統一插入
          continue;
        }
        // 其他元素類型可以擴充
      }

      // 3c. 插入分頁符號（除了最後一個文件）
      if (index < docsList.length - 1) {
        mergedBody.appendPageBreak();
      }

      // 3d. 每複製 5 個文件休息 1 秒，避免超時
      if ((index + 1) % 5 === 0) {
        Utilities.sleep(1000);
      }

    } catch (e) {
      console.error(`    ⚠️ 複製 ${file.getName()} 時發生錯誤: ${e.message}`);
      // 繼續處理下一個文件
    }
  });

  // 4. 儲存並關閉臨時文件
  mergedDoc.saveAndClose();
  console.log(`  文件合併完成，開始匯出 PDF...`);

  // 5. 匯出為 PDF
  const mergedDocFile = DriveApp.getFileById(mergedDoc.getId());
  const pdfBlob = mergedDocFile.getAs(MimeType.PDF);

  // 6. 設定 PDF 檔名（使用原始 GradeBand 名稱，保留撇號等特殊字元）
  const pdfFileName = `${originalGradeBand}_2526_Fall_Midterm.pdf`;
  pdfBlob.setName(pdfFileName);

  // 7. 儲存 PDF 到目標資料夾
  const pdfFile = targetFolder.createFile(pdfBlob);
  console.log(`  ✅ PDF 已儲存: ${pdfFileName}`);

  // 8. 刪除臨時 Docs
  mergedDocFile.setTrashed(true);

  return pdfFile;
}

/**
 * 按 GradeBand 將 Google Docs 合併為 PDF
 * 讀取輸出資料夾中的所有 GradeBand 子資料夾，將每個子資料夾中的 Docs 合併為單一 PDF
 */
function mergeDocsToPDFByGradeBand() {
  try {
    console.log('========================================');
    console.log('📄 按 GradeBand 合併為 PDF');
    console.log('========================================');

    // 1. 取得輸出資料夾
    const parentFolder = DriveApp.getFolderById(CONFIG.outputFolderId);
    console.log(`輸出資料夾: ${parentFolder.getName()}`);
    console.log('');

    // 2. 遍歷所有 GradeBand 子資料夾
    const subfolders = parentFolder.getFolders();
    const results = [];
    const errors = [];

    // 先統計總數
    const subfoldersArray = [];
    while (subfolders.hasNext()) {
      subfoldersArray.push(subfolders.next());
    }

    console.log(`找到 ${subfoldersArray.length} 個 GradeBand 子資料夾`);
    console.log('');

    // 3. 處理每個子資料夾
    subfoldersArray.forEach((subfolder, folderIndex) => {
      const gradeBandFolderName = subfolder.getName();  // 例如 "G1_LTs"

      try {
        console.log(`[${folderIndex + 1}/${subfoldersArray.length}] 處理 ${gradeBandFolderName}...`);

        // 3a. 取得該子資料夾中所有 Google Docs（排除 PDF）
        const allFiles = subfolder.getFiles();
        const docsList = [];

        while (allFiles.hasNext()) {
          const file = allFiles.next();
          if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
            docsList.push(file);
          }
        }

        if (docsList.length === 0) {
          console.warn(`  ⚠️ ${gradeBandFolderName} 沒有 Google Docs 檔案，跳過`);
          errors.push(`${gradeBandFolderName} (無檔案)`);
          return;
        }

        console.log(`  找到 ${docsList.length} 個班級文件`);

        // 3b. 按檔名字母順序排序
        // 檔名格式：ClassName_Teacher_2526_Fall_Midterm_v2_timestamp
        // 排序會自動依 ClassName 排序（因為 ClassName 在檔名開頭）
        docsList.sort((a, b) => {
          return a.getName().localeCompare(b.getName());
        });

        console.log(`  文件已按檔名字母順序排序`);

        // 3c. 反推原始 GradeBand 名稱（從資料夾名稱轉回）
        // 例如 "G1_LTs" → "G1 LT's"
        const originalGradeBand = gradeBandFolderName
          .replace(/_/g, ' ')  // 底線 → 空格
          .replace(/LTs/g, "LT's")  // 特殊處理：LTs → LT's
          .replace(/ITs/g, "IT's"); // 特殊處理：ITs → IT's

        // 3d. 合併所有 Docs 為單一 PDF
        const pdfFile = mergeDocsIntoPDF(docsList, gradeBandFolderName, originalGradeBand, subfolder);

        results.push({
          gradeBand: originalGradeBand,
          folderName: gradeBandFolderName,
          classCount: docsList.length,
          pdfUrl: pdfFile.getUrl(),
          pdfName: pdfFile.getName()
        });

        console.log(`✅ ${gradeBandFolderName} 完成 (${docsList.length} 個班級)`);

      } catch (e) {
        console.error(`❌ ${gradeBandFolderName} 處理失敗: ${e.message}`);
        errors.push(`${gradeBandFolderName} (${e.message})`);
      }

      // 3e. 每完成一個 GradeBand，休息 2 秒避免超時
      console.log(`  休息 2 秒...\n`);
      Utilities.sleep(2000);
    });

    // 4. 生成結果報告
    return formatMergeResults(results, errors);

  } catch (e) {
    console.error(`執行失敗: ${e.message}`);
    throw e;
  }
}

/**
 * 格式化合併結果報告
 */
function formatMergeResults(results, errors) {
  let message = `✅ PDF 合併完成！\n\n`;
  message += `📊 成功: ${results.length} 個 GradeBand\n`;

  if (errors.length > 0) {
    message += `❌ 失敗: ${errors.length} 個 GradeBand\n`;
  }

  if (results.length > 0) {
    message += `\n📁 生成的 PDF 檔案：\n`;

    results.forEach(item => {
      message += `\n  📂 ${item.gradeBand} (${item.folderName}):\n`;
      message += `    • ${item.pdfName}\n`;
      message += `    • 包含 ${item.classCount} 個班級\n`;
    });

    message += `\n🔗 主資料夾: https://drive.google.com/drive/folders/${CONFIG.outputFolderId}`;
  }

  if (errors.length > 0) {
    message += `\n\n⚠️ 失敗的 GradeBand：\n`;
    errors.forEach((err, i) => {
      message += `${i + 1}. ${err}\n`;
    });
  }

  return message;
}

/**
 * 主函數：生成所有班級的 Google Docs 並合併為 PDF（按 GradeBand）
 * 自動執行兩階段流程：
 * 1. 生成所有獨立的 Google Docs 檔案
 * 2. 按 GradeBand 合併為 PDF
 */
function generateAndMergePDFReports() {
  try {
    console.log('========================================');
    console.log('🚀 自動兩階段執行');
    console.log('========================================\n');

    // 階段 1: 生成所有 Google Docs
    console.log('========================================');
    console.log('階段 1: 生成所有班級的 Google Docs 檔案');
    console.log('========================================');

    const docsResult = generateClassReports();
    console.log(docsResult);

    console.log('');
    console.log('✅ 階段 1 完成，休息 5 秒後繼續...');
    console.log('');
    Utilities.sleep(5000);

    // 階段 2: 合併為 PDF
    console.log('========================================');
    console.log('階段 2: 按 GradeBand 合併為 PDF');
    console.log('========================================');

    const pdfResult = mergeDocsToPDFByGradeBand();
    console.log(pdfResult);

    // 最終報告
    const finalReport = `✅ 全部完成！

【階段 1: Google Docs 生成】
${docsResult}

【階段 2: PDF 合併】
${pdfResult}
`;

    return finalReport;

  } catch (e) {
    console.error(`執行失敗: ${e.message}`);
    throw e;
  }
}
