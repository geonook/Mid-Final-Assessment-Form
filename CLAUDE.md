# CLAUDE.md - Mid-Final Assessment Form Generator

> **Documentation Version**: 2.2-beta
> **Last Updated**: 2025-10-23
> **Project**: 2526 Fall Midterm 班級報告生成器
> **Description**: Google Apps Script project for generating mid-term and final assessment class reports from Google Sheets data, with Google Docs merge functionality (beta - has formatting limitations)
> **Template**: Based on CLAUDE_TEMPLATE.md v1.0.0 by Chang Ho Chien

This file provides essential guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 🚨 CRITICAL RULES - READ FIRST

> **⚠️ RULE ADHERENCE SYSTEM ACTIVE ⚠️**
> **Claude Code must explicitly acknowledge these rules at task start**
> **These rules override all other instructions and must ALWAYS be followed:**

### 🔄 **RULE ACKNOWLEDGMENT REQUIRED**
> **Before starting ANY task, Claude Code must respond with:**
> "✅ CRITICAL RULES ACKNOWLEDGED - I will follow all prohibitions and requirements listed in CLAUDE.md"

### ❌ ABSOLUTE PROHIBITIONS
- **NEVER** create documentation files (.md) unless explicitly requested by user
- **NEVER** use git commands with -i flag (interactive mode not supported)
- **NEVER** use `find`, `grep`, `cat`, `head`, `tail`, `ls` commands → use Read, Grep, Glob tools instead
- **NEVER** create duplicate files (enhanced_, improved_, new_, v2_) → ALWAYS extend existing files
- **NEVER** create multiple implementations of same concept → single source of truth
- **NEVER** copy-paste code blocks → extract into shared utilities/functions
- **NEVER** hardcode values that should be configurable → use CONFIG object

### 📝 MANDATORY REQUIREMENTS
- **COMMIT** after every completed task/phase - no exceptions
- **GITHUB BACKUP** - Push to GitHub after every commit: `git push origin main`
- **READ FILES FIRST** before editing - Edit/Write tools will fail if you didn't read the file first
- **SINGLE SOURCE OF TRUTH** - One authoritative implementation per feature/concept

### 🔍 MANDATORY PRE-TASK COMPLIANCE CHECK
> **STOP: Before starting any task, Claude Code must explicitly verify ALL points:**

**Step 1: Rule Acknowledgment**
- [ ] ✅ I acknowledge all critical rules in CLAUDE.md and will follow them

**Step 2: Task Analysis**
- [ ] Will this take >30 seconds? → If YES, use Task agents not Bash
- [ ] Is this 3+ steps? → If YES, use TodoWrite breakdown first
- [ ] Am I about to use grep/find/cat? → If YES, use proper tools instead

**Step 3: Technical Debt Prevention (MANDATORY SEARCH FIRST)**
- [ ] **SEARCH FIRST**: Use Grep to find existing implementations
- [ ] **CHECK EXISTING**: Read any found files to understand current functionality
- [ ] Does similar functionality already exist? → If YES, extend existing code
- [ ] Am I creating a duplicate? → If YES, consolidate instead
- [ ] Have I read the file before editing? → Required for Edit/Write tools

**Step 4: Session Management**
- [ ] Is this a long/complex task? → If YES, plan context checkpoints
- [ ] Have I been working >1 hour? → If YES, consider /compact or session break

> **⚠️ DO NOT PROCEED until all checkboxes are explicitly verified**

## 🐙 GITHUB SETUP & AUTO-BACKUP

### Current GitHub Status
- **Repository**: (Check with `git remote -v`)
- **Auto-push**: After every commit
- **Branch**: main

### 📋 **GITHUB BACKUP WORKFLOW** (MANDATORY)
```bash
# After every commit, always run:
git push origin main

# This ensures:
# ✅ Remote backup of all changes
# ✅ Version history preservation
# ✅ Disaster recovery protection
```

## Project Overview

Google Apps Script project for generating mid-term and final assessment class reports from Google Sheets data. The system reads student and class information from spreadsheets and generates formatted Google Docs reports with student lists organized by English class.

**Current Version: 2526 Fall Midterm (v2.2-beta - Google Docs Merge Testing)**
- Each class report = 2 pages (proctoring guidelines + student list)
- Students automatically sorted by ID number (numeric)
- Fills existing table in template (6 columns including Present/Signed checkboxes)
- **Fixed font size**: 10pt for all content (consistent formatting)
- **Column widths optimized**: 585pt total (headers display without line breaks)
- **Google Docs merge by GradeBand**: ⚠️ **BETA - Has formatting issues** (see Known Issues below)
- **Apps Script direct execution**: Run from Apps Script editor without Google Sheets
- **Two-stage architecture**: Generate Docs → Merge to Google Docs (or run separately)
- **GradeBand subfolder organization**: Automatic folder creation and file organization
- **Test mode**: Single class testing for quick validation

### ⚠️ Known Issues (v2.2-beta)

**Google Docs Merge Formatting Problem**:
- First class in merged document: ✅ Format correct
- Subsequent classes: ❌ Format breaks (layout distorted, tables misaligned)
- Root cause: Element-by-element copying doesn't preserve complex multi-column layouts
- **Workaround**: Use individual class documents for now, or export to PDF manually
- **Status**: Under investigation - may require Google Docs API or alternative approach

## Development Commands

### Deploy to Google Apps Script
```bash
clasp push              # Push local changes to Google Apps Script
clasp pull              # Pull latest changes from Google Apps Script
clasp open              # Open the script project in browser
clasp logs              # View execution logs
```

### Git Commands
```bash
git status              # Check current status
git add .               # Stage all changes
git commit -m "msg"     # Commit with message
git push origin main    # Push to GitHub (MANDATORY after every commit)
```

## Architecture

### Enhanced Structure (v2.2-beta - October 2025)
**Focus**: Template filling + Google Docs merge (beta) + Direct execution

**Code Metrics**:
- Main script: ~1,350 lines (with Google Docs merge functionality)
- Functions: 25+ functions (core + merge + execution helpers)
- Config-driven: All settings in single CONFIG object
- Template-based: Maintains copy-and-fill approach (no hardcoding)
- Placeholder replacement: Full support for all exam metadata fields
- Two-stage processing: Docs generation separate from merge (⚠️ merge has formatting issues)

### Core Data Flow

**Stage 1: Document Generation**
1. **Data Source**: Google Sheets with two worksheets:
   - `Students` sheet: Student records (ID, Home Room, Chinese Name, English Name, English Class)
   - `Class` sheet: Class metadata (ClassName, Teacher, GradeBand, Duration, Periods, etc.)

2. **Processing**:
   - Sort classes by GradeBand → ClassName
   - Group students by English Class → Sort by ID (numeric)
   - Copy template → Replace placeholders → Fill student table
   - Save to GradeBand subfolder

3. **Output**: Individual Google Docs files (one per class) in GradeBand subfolders

**Stage 2: Google Docs Merge (Optional - ⚠️ Beta)**
1. **Input**: Google Docs files organized in GradeBand subfolders
2. **Processing**:
   - For each GradeBand subfolder
   - Copy first Doc as base using `makeCopy()`
   - Append remaining Docs element-by-element
   - Save to Merged folder
3. **Output**: One merged Google Doc per GradeBand in Merged folder
4. **⚠️ Known Issue**: Formatting breaks for classes after the first one

### Project Structure (Google Apps Script)

```
Mid-Final Assessment Form-template/
├── CLAUDE.md              # This file - Essential rules for Claude Code
├── README.md              # User-facing documentation
├── .gitignore             # Git ignore patterns
├── .clasp.json            # Clasp configuration
├── appsscript.json        # Apps Script project configuration
├── 主程式.js              # Main script (PRODUCTION)
└── 輔助工具/              # Helper tools
    ├── 分析模板.js        # Template structure analyzer
    └── 轉檔.js            # PDF batch converter
```

**Note**: Google Apps Script has flat file structure requirements. All `.js` files must be at root or in single subfolder level.

### Script Files

**主程式.js** - Main script (PRODUCTION - v2.2-beta)
- **Purpose**: Generate 2526 Fall Midterm class reports + Google Docs merge by GradeBand (⚠️ beta)
- **Main functions**:
  - `generateClassReports()`: Stage 1 - Generate all Google Docs (✅ stable)
  - `mergeDocsByGradeBand()`: Stage 2 - Merge Docs to Google Docs (⚠️ has formatting issues)
  - `generateAndMergeDocsReports()`: One-click execution (both stages)
- **Apps Script Direct Execution** (NEW):
  - `RUN_FULL_BATCH()`: Complete batch (Docs + Merge) - **Recommended**
  - `RUN_DOCS_ONLY()`: Generate Docs only
  - `RUN_MERGE_ONLY()`: Merge to Google Docs only (requires Docs already generated)
- **Test function**: `testSingleClass()` - Test with single class
- **Entry points**:
  - **Google Sheets Menu**: "班級報告" → Choose execution mode
  - **Apps Script Editor**: Run `RUN_FULL_BATCH()` function directly
- **Core Logic**:
  - Reads Students and Class sheets
  - Groups students by English Class
  - Sorts classes by GradeBand → ClassName (alphabetically)
  - Sorts students by ID (numeric)
  - Copies template document
  - Replaces all placeholders ({{GradeBand}}, {{Duration}}, etc.)
  - Fills student table (tables[2]) with 6 columns
  - Applies fixed font sizing (10pt for all content)
  - Saves to GradeBand subfolder
  - (Optional) Merges all Docs per GradeBand to single Google Doc (⚠️ formatting issues)
- **Key functions**:
  - **Document Generation**:
    - `generateSingleReport()`: Creates individual class document
    - `replacePlaceholders()`: Replaces {{...}} with class data
    - `fillStudentTable()`: Core table filling logic (sort → locate → fill → format)
    - `formatTableRow()`: Unified cell formatting
    - `getOrCreateSubfolder()`: GradeBand subfolder management
  - **Google Docs Merge** (⚠️ Beta):
    - `sortClassDataByGradeBandAndName()`: Two-level class sorting
    - `mergeDocsToGoogleDocs()`: Core merge logic (makeCopy + element appending)
    - `getOrCreateMergedFolder()`: Creates/retrieves Merged folder
    - `formatMergeResults()`: Results formatting and reporting
  - **Utilities**:
    - `calculateFontSize()`: Returns fixed 10pt
    - `getStudentIndexes()`, `getClassIndexes()`: Column mapping
    - `groupStudentsByClass()`: Student grouping by English Class
- **Configuration**: Single CONFIG object
  ```javascript
  const CONFIG = {
    spreadsheetId: '1bo3xsXw0u8Wwbo6ALvPe9idKDVCjhtAVKAKfH8azhJE',  // For Apps Script execution
    templateId: '1D2hSZNI8MQzD_OIeCdEvpqp4EWfO2mrjTCHAQZyx6MM',
    outputFolderId: '1KSyHsy1wUcrT82OjkAMmPFaJmwe-uosi',
    studentTableIndex: 2,  // Third table (tables[2]) - student list table
    columnWidths: [75, 75, 95, 100, 85, 155],  // Total: 585pt (~20.6cm)
    fontSize: { large: 10, medium: 10, small: 10 },  // Fixed 10pt
    semester: '2526_Fall_Midterm_v2',
    delayMs: 1000,
    placeholderFields: [
      'GradeBand', 'Duration', 'Periods', 'Self-Study', 'Preparation',
      'ExamTime', 'Level', 'Classroom', 'Proctor', 'Subject',
      'ClassName', 'Teacher', 'Count', 'Students'
    ]
  };
  ```

**輔助工具/分析模板.js** - Template analyzer (HELPER TOOL)
- **Purpose**: Debug template structure
- **Main function**: `analyzeTemplateStructure()`
- **Reports**: Page dimensions, table count/structure, placeholder presence, column widths
- **Use when**: Template structure changes or debugging table access

**輔助工具/轉檔.js** - PDF converter (HELPER TOOL - UPDATED)
- **Purpose**: Batch convert Google Docs to individual PDF files
- **Main functions**:
  - `convertAllSubfoldersDocsToPDF()`: Convert all GradeBand subfolders (✅ RECOMMENDED)
  - `batchDownloadDocsToPDF()`: Convert single folder (legacy)
  - `convertSingleFolderDocsToPDF(folder)`: Shared conversion logic
- **Features**:
  - Recursive subfolder processing
  - Automatic skip of Merged folder
  - Duplicate checking (skip existing PDFs)
  - Progress logging and statistics
  - Rate limiting (every 10 files)
- **Output**: PDFs saved in same subfolder as source Docs
- **Usage**: Run `convertAllSubfoldersDocsToPDF()` from Apps Script editor
- **Execution time**: ~6-9 minutes for 168 classes

### Common Patterns

**Data Sheet Validation**
All scripts check for required columns and throw errors if missing:
- Students sheet: "English Class", "ID", "Home Room", "Chinese Name", "English Name"
- Class sheet: "ClassName", "Teacher", "GradeBand", and all placeholder fields

**Fixed Font Sizing (v2.1)**
Font size is now fixed at 10pt for all content:
- All student table content: 10pt
- Headers: 10pt bold
- Consistent formatting regardless of class size
- Column widths optimized (585pt) to display headers without line breaks

**Class Sorting (NEW)**
Classes are sorted by two levels before processing:
1. **GradeBand**: Primary sort (e.g., "G1 LT's", "G1 IT's", "G2 LT's")
2. **ClassName**: Secondary sort (alphabetical within each GradeBand)
- Ensures consistent ordering for PDF merge
- Implemented in `sortClassDataByGradeBandAndName()`

**GradeBand Subfolder Organization (NEW)**
Files automatically organized by GradeBand:
- Creates subfolder per GradeBand (e.g., "G1_LTs", "G2_ITs")
- Folder naming: Sanitizes apostrophes and spaces for filesystem compatibility
- Implemented in `getOrCreateSubfolder()`

**Table Formatting**
Consistent approach across all generators:
- Column widths set explicitly for each field
- Headers: bold, centered, gray background (#E8E8E8)
- Content: centered horizontally and vertically
- Border: 1pt black

**Error Handling**
Best practices demonstrated:
- Try-catch blocks around file operations
- Detailed console logging with progress indicators
- User-friendly error messages
- Graceful degradation (continues processing other classes on failure)
- Timeout prevention with strategic `Utilities.sleep()` calls

**PDF Merge Architecture (NEW)**
Element-by-element copying to preserve template formatting:
1. Create temporary merged document
2. For each class Doc:
   - Open source document
   - Copy all elements (paragraphs, tables, lists) using `.copy()`
   - Append to merged document body
   - Insert page break between classes
3. Export merged Doc to PDF
4. Delete temporary Doc
- Maintains all template formatting and structure
- No hardcoded content or dynamic structure creation
- Strategic rest periods (every 5 docs, between GradeBands) to prevent timeout

### Menu Integration

Scripts include `onOpen()` function to add custom menu to Google Sheets:
```javascript
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('班級報告')
    .addItem('步驟 1: 生成所有 Google Docs', 'runReportGeneration')
    .addItem('步驟 2: 合併為 Google Docs（按 GradeBand）', 'runMergeDocs')
    .addSeparator()
    .addItem('🚀 一鍵執行（Docs + 合併）', 'runGenerateAndMergeDocs')
    .addSeparator()
    .addItem('🧪 測試單一班級（v2）', 'testSingleClass')
    .addToUi();
}
```

### Execution Methods

**Method 1: Google Sheets Menu** (User-friendly)
1. Open Google Sheets with Students and Class data
2. Click **班級報告** menu
3. Choose execution mode:
   - **🚀 一鍵執行（Docs + 合併）**: Full automated process (recommended, ⚠️ merge has formatting issues)
   - **步驟 1**: Generate Docs only (if you want to review before merge)
   - **步驟 2**: Merge to Google Docs only (after Docs generated, ⚠️ formatting issues)
   - **🧪 測試**: Test with single class first

**Method 2: Apps Script Editor Direct Execution** (Developer mode)
1. Open Apps Script project: Extensions → Apps Script
2. Select function from dropdown:
   - **`RUN_FULL_BATCH()`**: Complete batch (Docs + Merge) - **Recommended**
   - **`RUN_DOCS_ONLY()`**: Generate Docs only
   - **`RUN_MERGE_ONLY()`**: Merge to Google Docs only (⚠️ formatting issues)
3. Click Run ▶
4. Monitor console logs in real-time
5. **Advantages**:
   - No need to open Google Sheets
   - Direct access to execution logs
   - Faster for testing and debugging
   - Can run from any Google account with script access

## Configuration

### Hard-coded IDs (Update for different deployments)
- Template Document ID: `1D2hSZNI8MQzD_OIeCdEvpqp4EWfO2mrjTCHAQZyx6MM`
- Output Folder ID: `1KSyHsy1wUcrT82OjkAMmPFaJmwe-uosi`
- PDF Conversion Folder: `1KSyHsy1wUcrT82OjkAMmPFaJmwe-uosi` (same as output)

### Page Dimensions (in points)
- A4 Landscape: 842 x 595 (used in this project)

### Margins
- Template default: Pre-configured in template document

## Usage Guide

### Quick Start (Recommended)
**Apps Script Direct Execution** - Fastest method:
1. Open Apps Script: Extensions → Apps Script
2. Select `RUN_DOCS_ONLY()` from function dropdown (⚠️ Skip merge due to formatting issues)
3. Click Run ▶
4. Wait for completion (~8-10 minutes for 168 classes)
5. Check output folder for individual class Docs in GradeBand subfolders

### Manual Two-Stage Process
If you want to try the merge feature (⚠️ has formatting issues):
1. **Stage 1**: Run `RUN_DOCS_ONLY()` or use Sheets menu "步驟 1"
2. Review generated Docs in GradeBand subfolders
3. **Stage 2**: Run `RUN_MERGE_ONLY()` or use Sheets menu "步驟 2" (⚠️ formatting issues)

### Testing
Before full batch, test with single class:
- **Google Sheets**: Menu → "🧪 測試單一班級（v2）"
- Validates template filling logic with first class only
- Shows detailed result with file link

### Output Structure
```
Output Folder (CONFIG.outputFolderId)
├── G1_LTs/
│   ├── A1_TeacherName_2526_Fall_Midterm_v2_20251023_143052.docx
│   ├── A2_TeacherName_2526_Fall_Midterm_v2_20251023_143055.docx
│   └── ... (all individual class Docs)
├── G1_ITs/
│   └── ... (individual class Docs)
├── G2_LTs/
│   └── ... (individual class Docs)
└── Merged/  ← ⚠️ Beta: Contains merged Docs (formatting issues)
    ├── G1 LT's_2526_Fall_Midterm_Merged.docx
    ├── G1 IT's_2526_Fall_Midterm_Merged.docx
    └── G2 LT's_2526_Fall_Midterm_Merged.docx
```

### PDF Conversion (Optional)
If you need individual PDF files for each class:

1. **Open Apps Script editor**: Extensions → Apps Script
2. **Select function**: `convertAllSubfoldersDocsToPDF` from dropdown
3. **Click Run** ▶
4. **Wait for completion**: ~6-9 minutes for 168 classes
5. **Check results**: PDFs saved in same subfolders as Docs

**Output structure after PDF conversion**:
```
Output Folder (CONFIG.outputFolderId)
├── G1_LTs/
│   ├── A1_TeacherName_xxx.docx         ← Google Docs
│   ├── A1_TeacherName_xxx.pdf          ← PDF
│   ├── A2_TeacherName_xxx.docx
│   ├── A2_TeacherName_xxx.pdf
│   └── ...
└── ...
```

**Features**:
- ✅ Automatic subfolder processing
- ✅ Skip existing PDFs (no duplicates)
- ✅ Detailed progress logging
- ✅ Skip Merged folder automatically

### Troubleshooting Tools
- **Template debugging**: Run `輔助工具/分析模板.js` → `analyzeTemplateStructure()`
- **Single folder PDF conversion**: Run `輔助工具/轉檔.js` → `batchDownloadDocsToPDF()`

## Important Implementation Details (2526 Fall Midterm)

### Template Structure
**Page 1**: Exam proctoring guidelines with placeholders
- Title: "SY25-26 ID Midterm & Final Exam Proctoring Guidelines"
- Contains 14 placeholders: {{GradeBand}}, {{Duration}}, {{Periods}}, etc.
- All placeholders replaced with class-specific data using `replacePlaceholders()`

**Page 2**: Three tables
- First table (tables[0]): Exam information header
- Second table (tables[1]): Class information (left side)
- Third table (tables[2]): Student list (right side) - **THIS IS WHAT GETS FILLED**

### Student Sorting
Students are automatically sorted by ID number (numeric sort) before filling the table:
```javascript
const sortedStudents = students.slice().sort((a, b) => {
  const idA = String(a[studentIndexes.ID] || '');
  const idB = String(b[studentIndexes.ID] || '');
  return idA.localeCompare(idB, undefined, { numeric: true });
});
```

### Placeholder Replacement (NEW in v2)
All 14 placeholders on Page 1 are replaced with class-specific data:
```javascript
function replacePlaceholders(body, classData) {
  CONFIG.placeholderFields.forEach(field => {
    const placeholder = `{{${field}}}`;
    const value = String(classData[field] || '');
    body.replaceText(placeholder, value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  });
}
```
- **Regex escaping**: Special characters properly escaped to prevent regex errors
- **Empty handling**: Missing data replaced with empty string (no "undefined")
- **Example replacements**: {{GradeBand}} → "G1 LT's", {{Duration}} → "50 mins", {{Periods}} → "3-4"

### Template Table Filling (Not Creating New Tables)
The script locates and fills the existing table in the template (tables[2]):
1. **Locate**: Access 3rd table (index 2) in document body
2. **Fill**: Overwrite existing rows with student data (Solution C approach)
3. **Format**: Apply fixed 10pt font, alignment, borders
4. **Set widths**: Apply column widths [75, 75, 95, 100, 85, 155] = 585pt total

**6 Columns**:
| Column | Field | Width | Notes |
|--------|-------|-------|-------|
| 0 | Student ID | 75pt | From Students sheet |
| 1 | Homeroom | 75pt | From Students sheet |
| 2 | Chinese Name | 95pt | From Students sheet |
| 3 | English Name | 100pt | From Students sheet |
| 4 | Present | 85pt | Checkbox ☐ (static) |
| 5 | Signed Paper Returned | 155pt | Checkbox ☐ (static) |

**Total width**: 585pt (~20.6cm) - optimized to fit A4 landscape with headers displaying fully

### Font Size Logic (Fixed 10pt)
Font size is now fixed regardless of student count:
- **All students**: 10pt (consistent formatting)
- **Headers**: 10pt bold
- **Column widths optimized**: Headers display without line breaks at 10pt

### File Naming Conventions

**Individual Docs**:
- Format: `{ClassName}_{TeacherName}_2526_Fall_Midterm_v2_{timestamp}`
- Timestamp: `yyyyMMdd_HHmmss` (includes seconds to prevent overwrite)
- Example: `A1_JohnDoe_2526_Fall_Midterm_v2_20251023_143052`
- Saved to: `{OutputFolder}/{GradeBand_Sanitized}/`

**Merged PDFs**:
- Format: `{GradeBand}_2526_Fall_Midterm.pdf`
- GradeBand preserved with original formatting (apostrophes restored)
- Example: `G1 LT's_2526_Fall_Midterm.pdf`
- Saved to: `{OutputFolder}/{GradeBand_Sanitized}/`

### GradeBand Folder Naming
- **Sanitization**: Apostrophes and spaces replaced for filesystem compatibility
- **Examples**:
  - "G1 LT's" → Folder: "G1_LTs", PDF: "G1 LT's_2526_Fall_Midterm.pdf"
  - "G2 IT's" → Folder: "G2_ITs", PDF: "G2 IT's_2526_Fall_Midterm.pdf"
- **Reverse mapping**: Folder names converted back to original format for PDF naming

### Progress Tracking
Console logs show detailed progress during batch processing:
```javascript
// Document generation
const progress = Math.round((i / sortedClassList.length) * 100);
console.log(`處理 ${classInfo.ClassName}... [${i + 1}/${sortedClassList.length}] (${progress}%)`);

// Google Docs merge
console.log(`處理 GradeBand ${folderIndex + 1}/${subfoldersArray.length}: ${gradeBandFolderName}`);
console.log(`  合併 ${docsList.length} 個文件...`);
```

### Execution Time Estimates
Based on 168 classes across 6 GradeBands:
- **Docs generation only**: ~8-10 minutes (✅ recommended)
- **Google Docs merge only**: ~2-3 minutes (⚠️ formatting issues)
- **Full batch (Docs + Merge)**: ~10-15 minutes (⚠️ merge has issues)
- **Test mode (single class)**: ~5-10 seconds

### Test Mode
Use `testSingleClass()` to test with first class only:
- Faster testing without generating all reports (~5-10 seconds)
- Validates template filling logic and placeholder replacement
- Shows detailed result message with file link
- Access via menu: "🧪 測試單一班級（v2）"

## Known Limitations

- **Google Docs merge formatting issues** (v2.2-beta): First class correct, subsequent classes have distorted layout
- Apps Script execution time limit: 6 minutes for custom functions (mitigated by strategic sleep() calls)
- Rate limiting required for batch operations (1-second delay between files)
- Hard-coded resource IDs require manual updates for different environments
- Template capacity: Optimized for reasonable class sizes with fixed 10pt font
- Google Apps Script flat file structure (cannot organize into src/ folders)
- Element-by-element copying doesn't preserve complex multi-column layouts
- Large batches (>200 classes) may require splitting into multiple runs

## Timezone

Project timezone: `Asia/Taipei` (configured in appsscript.json)

## 🚨 TECHNICAL DEBT PREVENTION

### ❌ WRONG APPROACH (Creates Technical Debt):
```bash
# Creating new file without searching first
Write(file_path="新程式.js", content="...")
```

### ✅ CORRECT APPROACH (Prevents Technical Debt):
```bash
# 1. SEARCH FIRST
Grep(pattern="function.*generate", glob="*.js")
# 2. READ EXISTING FILES
Read(file_path="主程式.js")
# 3. EXTEND EXISTING FUNCTIONALITY
Edit(file_path="主程式.js", old_string="...", new_string="...")
```

## 🧹 DEBT PREVENTION WORKFLOW

### Before Creating ANY New File:
1. **🔍 Search First** - Use Grep/Glob to find existing implementations
2. **📋 Analyze Existing** - Read and understand current patterns
3. **🤔 Decision Tree**: Can extend existing? → DO IT | Must create new? → Document why
4. **✅ Follow Patterns** - Use established project patterns
5. **📈 Validate** - Ensure no duplication or technical debt

## 🎯 DEVELOPMENT WORKFLOW

### Standard Task Flow:
1. **Read CLAUDE.md** - Verify rules compliance
2. **Search existing code** - Use Grep/Glob before creating
3. **TodoWrite for complex tasks** - Break down 3+ step tasks
4. **Read before edit** - Always read files before using Edit/Write
5. **Commit frequently** - After each completed feature/fix
6. **Push to GitHub** - `git push origin main` after every commit
7. **Test changes** - Use test mode for validation

### Quick Reference:
```bash
# Development cycle
clasp pull                    # Get latest from Apps Script
# Make changes locally
clasp push                    # Deploy to Apps Script
# Test in Google Sheets
git add . && git commit -m "msg"  # Commit
git push origin main          # Backup to GitHub
```

---

**⚠️ Prevention is better than consolidation - build clean from the start.**
**🎯 Focus on single source of truth and extending existing functionality.**
**📈 Each task should maintain clean architecture and prevent technical debt.**

---

**🎯 Template adapted from CLAUDE_TEMPLATE.md v1.0.0 by Chang Ho Chien**
**📺 Original Tutorial: https://youtu.be/8Q1bRZaHH24**
