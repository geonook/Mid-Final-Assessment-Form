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

    console.log(`開始處理 ${classData.length - 1} 個班級...`);

    // 生成報告
    const results = [];
    const errors = [];

    for (let i = 1; i < classData.length; i++) {
      // 建立完整的班級資料物件（包含所有新欄位）
      const classRow = classData[i];
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

      if (!classInfo.ClassName) continue;

      const students = groupedStudents[classInfo.ClassName];
      if (!students || students.length === 0) {
        console.warn(`班級 ${classInfo.ClassName} 沒有學生資料`);
        errors.push(`${classInfo.ClassName}-${classInfo.Teacher} (無學生資料)`);
        continue;
      }

      try {
        const progress = Math.round((i / (classData.length - 1)) * 100);
        console.log(`處理 ${classInfo.ClassName} (${classInfo.Teacher})... [${i}/${classData.length - 1}] (${progress}%)`);

        const file = generateSingleReport(classInfo, students, studentIndexes);
        results.push({
          name: classInfo.ClassName,
          teacher: classInfo.Teacher,
          gradeBand: classInfo.GradeBand,
          url: file.getUrl()
        });

        // 延遲避免 API 限制
        if (i < classData.length - 1) {
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

  // 開啟文件
  const doc = DocumentApp.openById(newFile.getId());
  const body = doc.getBody();

  // 步驟 1: 替換佔位符（第一頁）
  replacePlaceholders(body, classData);

  // 步驟 2: 填充學生表格（第二頁）
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
 * 替換文件中的所有佔位符
 * 搜尋並替換所有 {{FieldName}} 格式的佔位符
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
    // replaceText() 使用正則表達式，需要將 { 和 } 跳脫為 \{ 和 \}
    const escapedPlaceholder = placeholder.replace(/[{}]/g, '\\$&');

    // 在整個文件中搜尋並替換（支援表格和段落）
    try {
      body.replaceText(escapedPlaceholder, value);
      console.log(`  替換 ${placeholder} → ${value}`);
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
    .addItem('生成 2526 Fall Midterm v2 報告', 'runReportGeneration')
    .addItem('🧪 測試單一班級（v2）', 'testSingleClass')
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
