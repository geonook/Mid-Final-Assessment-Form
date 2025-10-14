# CLAUDE.md - Mid-Final Assessment Form Generator

> **Documentation Version**: 2.0
> **Last Updated**: 2025-10-13
> **Project**: 2526 Fall Midterm 班級報告生成器
> **Description**: Google Apps Script project for generating mid-term and final assessment class reports from Google Sheets data
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

**Current Version: 2526 Fall Midterm (v2.0 - Simplified)**
- Each class report = 2 pages (proctoring guidelines + student list)
- Students automatically sorted by ID number (numeric)
- Fills existing table in template (6 columns including Present/Signed checkboxes)
- Supports up to 21 students per class
- Dynamic font sizing (9-11pt based on class size)
- **Code reduction**: 450 → 414 lines (8% reduction)
- **Test mode**: Single class testing for quick validation

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

### Simplified Structure (v2.0 - October 2025)
**Focus**: Single-purpose template filling system - maximum simplicity, minimum complexity

**Code Metrics**:
- Main script: 414 lines (optimized with improvements)
- Functions: 10 functions (including test mode)
- Config-driven: All settings in single CONFIG object
- Zero page formatting logic (template pre-configured)
- Zero placeholder replacement (template has no placeholders on Page 1)

### Core Data Flow
1. **Data Source**: Google Sheets with two worksheets:
   - `Students` sheet: Student records (ID, Home Room, Chinese Name, English Name, English Class)
   - `Class` sheet: Class metadata (ClassName, Tea)

2. **Processing**: Group students by English Class → Sort by ID (numeric) → Fill template table

3. **Output**: Google Docs files (one per class) saved to Drive folder

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

**主程式.js** - Main script (PRODUCTION - CURRENT)
- **Purpose**: Generate 2526 Fall Midterm class reports
- **Main function**: `generateClassReports()`
- **Test function**: `testSingleClass()` - NEW: Test with single class
- **Entry point**: Menu → "班級報告" → "生成 2526 Fall Midterm 報告" or "🧪 測試單一班級"
- **Core Logic**:
  - Reads Students and Class sheets
  - Groups students by English Class
  - Sorts students by ID (numeric)
  - Copies template document
  - Fills student table (tables[1]) with 6 columns
  - Applies dynamic font sizing (9-11pt based on class size)
  - Validates student count (warns if >21)
- **Key functions**:
  - `generateSingleReport()`: Creates individual class document
  - `fillStudentTable()`: Core table filling logic (sort → locate → clear → fill → format)
  - `formatTableRow()`: Unified cell formatting
  - `calculateFontSize()`: Font sizing for 21-student max
  - `testSingleClass()`: Test mode for first class only
- **Configuration**: Single CONFIG object
  ```javascript
  const CONFIG = {
    templateId: '1D2hSZNI8MQzD_OIeCdEvpqp4EWfO2mrjTCHAQZyx6MM',
    outputFolderId: '1KSyHsy1wUcrT82OjkAMmPFaJmwe-uosi',
    studentTableIndex: 1,  // Right-side student table
    columnWidths: [90, 100, 140, 140, 80, 120],
    fontSize: { large: 11, medium: 10, small: 9 },
    semester: '2526_Fall_Midterm',
    delayMs: 1000
  };
  ```

**輔助工具/分析模板.js** - Template analyzer (HELPER TOOL)
- **Purpose**: Debug template structure
- **Main function**: `analyzeTemplateStructure()`
- **Reports**: Page dimensions, table count/structure, placeholder presence, column widths
- **Use when**: Template structure changes or debugging table access

**輔助工具/轉檔.js** - PDF converter (HELPER TOOL)
- **Purpose**: Batch convert generated docs to PDF
- **Main function**: `batchDownloadDocsToPDF()`
- **Features**: Duplicate checking, progress logging, rate limiting
- **Folder ID**: `1KSyHsy1wUcrT82OjkAMmPFaJmwe-uosi` (same as output folder)

### Common Patterns

**Data Sheet Validation**
All scripts check for required columns and throw errors if missing:
- Students sheet: "English Class", "ID", "Home Room", "Chinese Name", "English Name"
- Class sheet: "ClassName", "Tea"

**Dynamic Font Sizing**
Font sizes automatically adjust based on student count to fit content on page:
- 19-21 students: 9pt (tight fit, ensures 2-page limit)
- 13-18 students: 10pt (balanced)
- 1-12 students: 11pt (comfortable reading)

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
- Student count validation (warns if >21)

### Menu Integration

Scripts include `onOpen()` function to add custom menu to Google Sheets:
```javascript
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('班級報告')
    .addItem('生成 2526 Fall Midterm 報告', 'runReportGeneration')
    .addItem('🧪 測試單一班級', 'testSingleClass')
    .addToUi();
}
```

## Configuration

### Hard-coded IDs (Update for different deployments)
- Template Document ID: `1D2hSZNI8MQzD_OIeCdEvpqp4EWfO2mrjTCHAQZyx6MM`
- Output Folder ID: `1KSyHsy1wUcrT82OjkAMmPFaJmwe-uosi`
- PDF Conversion Folder: `1KSyHsy1wUcrT82OjkAMmPFaJmwe-uosi` (same as output)

### Page Dimensions (in points)
- A4 Landscape: 842 x 595 (used in this project)

### Margins
- Template default: Pre-configured in template document

## Script Selection Guide

- **2526 Fall Midterm (Current)**: Use `主程式.js` (PRODUCTION - ONLY OPTION)
  - Generates individual files per class
  - Students sorted by ID automatically
  - Fills template's existing 6-column table
  - Run from Google Sheets menu: "班級報告" → "生成 2526 Fall Midterm 報告"
  - **Test mode**: "🧪 測試單一班級" for quick validation
- **PDF output**: First generate docs, then run `輔助工具/轉檔.js`
- **Template debugging**: Use `輔助工具/分析模板.js` to inspect template structure

## Important Implementation Details (2526 Fall Midterm)

### Template Structure
**Page 1**: Exam proctoring guidelines (fixed content, NO placeholders)
- Title: "SY25-26 ID Midterm & Final Exam Proctoring Guidelines"
- Content remains unchanged in generated reports

**Page 2**: Two tables side-by-side
- Left table (tables[0]): Class information (not modified by script)
- Right table (tables[1]): Student list - **THIS IS WHAT GETS FILLED**

### Student Sorting
Students are automatically sorted by ID number (numeric sort) before filling the table:
```javascript
const sortedStudents = students.slice().sort((a, b) => {
  const idA = String(a[studentIndexes.ID] || '');
  const idB = String(b[studentIndexes.ID] || '');
  return idA.localeCompare(idB, undefined, { numeric: true });
});
```

### Student Count Validation
Script validates student count and warns if exceeding template capacity:
```javascript
if (sortedStudents.length > 21) {
  console.warn(`⚠️ 警告: 學生人數 ${sortedStudents.length} 超過 21 人，可能超過 2 頁限制`);
}
```

### Template Table Filling (Not Creating New Tables)
The script locates and fills the existing table in the template (tables[1]):
1. **Locate**: Access 2nd table (index 1) in document body
2. **Clear**: Remove all rows except header row
3. **Fill**: Append new rows for each sorted student
4. **Format**: Apply font sizing, alignment, borders
5. **Set widths**: Apply column widths [90, 100, 140, 140, 80, 120]

**6 Columns**:
| Column | Field | Width | Notes |
|--------|-------|-------|-------|
| 0 | Student ID | 90pt | From Students sheet |
| 1 | Homeroom | 100pt | From Students sheet |
| 2 | Chinese Name | 140pt | From Students sheet |
| 3 | English Name | 140pt | From Students sheet |
| 4 | Present | 80pt | Checkbox ☐ (static) |
| 5 | Signed Paper Returned | 120pt | Checkbox ☐ (static) |

### Font Size Logic (Max 21 Students)
Dynamic sizing ensures content fits within 2-page limit:
- **19-21 students**: 9pt (tight fit)
- **13-18 students**: 10pt (balanced)
- **1-12 students**: 11pt (comfortable reading)

### File Naming Convention
Format: `{ClassName}_{TeacherName}_2526_Fall_Midterm_{timestamp}`
- Timestamp format: `yyyyMMdd_HHmmss` (includes seconds to prevent overwrite)
- Example: `A1_TeacherName_2526_Fall_Midterm_20251015_143052`

### Progress Tracking
Console logs show progress during batch processing:
```javascript
const progress = Math.round((i / (classData.length - 1)) * 100);
console.log(`處理 ${classInfo.name}... [${i}/${classData.length - 1}] (${progress}%)`);
```

### Test Mode (NEW in v2.0)
Use `testSingleClass()` to test with first class only:
- Faster testing without generating all reports
- Validates template filling logic
- Shows detailed result message with file link
- Access via menu: "🧪 測試單一班級"

## Known Limitations

- Apps Script execution time limit: 6 minutes for custom functions
- Rate limiting required for batch operations (1-second delay between files)
- Hard-coded resource IDs require manual updates for different environments
- Maximum 21 students per class (template constraint)
- Google Apps Script flat file structure (cannot organize into src/ folders)

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
