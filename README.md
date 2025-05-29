# 🈺 English-to-Chinese Track Changes Automation

A powerful Python script that **automatically transfers tracked changes** from an **English Word document** (`.docx`) to its aligned **Chinese translation**, preserving Microsoft Word’s **Track Changes** format. It supports insertions, deletions, replacements, and formatting indicators with intelligent paragraph alignment and summary logging.

---

## 📌 Overview

This tool automates the process of syncing revision history from an English document to its corresponding Chinese version.

### 🔽 Input:

* `edited_en.docx` — English document with **Track Changes** enabled.
* `original_cn.docx` — Chinese version of the same document (unaltered).

### 🔼 Output:

* `original_cn_with_tracked_changes.docx` — Chinese document with equivalent **tracked changes** applied and visible in Microsoft Word.

---

## ✅ Features

* 🔁 **Fully Automated**
  No manual intervention needed — the script handles everything from extraction to change application.

* 📝 **Track Changes Preserved**
  Changes (insert, delete, replace, format) are mirrored in the Chinese file using Word’s native Track Changes feature.

* 🌐 **Context-Aware Paragraph Matching**
  Uses `difflib` to align translated edits with matching Chinese paragraphs, ensuring precision.

* 🧠 **Bold Marker Detection**
  Recognizes bold text changes and annotates them in the Chinese output using `[bold]` markers.

* 💬 **Dual Input Modes**
  Accepts both command-line arguments and interactive user prompts for flexible usage.

* 📊 **Change Summary Report**
  After processing, displays a count of each type of change applied or skipped.

---

## ⚙️ Requirements

* **Python**: 3.11 or 3.12

  > ⚠️ Python 3.13+ is **not supported** due to dependency compatibility issues.

* **Microsoft Word** (desktop version) — required for COM automation via `win32com`.

### 📦 Install Dependencies

```bash
pip install pywin32 googletrans==4.0.0-rc1
```

---

## 🚀 Usage

### ▶️ Option 1: Interactive Input Mode

```bash
python main.py
```

You will be prompted to enter the file paths:

```
Enter path to English Word document (.docx): C:\path\to\English_document.docx
Enter path to Chinese Word document (.docx): C:\path\to\Chinese_document.docx
```

### 📁 Output:

```
C:\path\to\Chinese_document_with_tracked_changes.docx
```

---

## 📊 Output Example

After running the script, you’ll receive a summary like the following:

```
📊 Change Summary:
  - insert: 4
  - delete: 2
  - replace: 1
  - format: 1
  - bold: 1
  - skipped: 0
```

## ⚠️ Notes & Recommendations

* ✅ **Close all Microsoft Word windows** before running the script — open Word instances may interfere with automation.
* ❌ Avoid Python 3.13 or newer — known compatibility issues with COM and `googletrans`.
* 📌 **Translation is used only for alignment** — the output is not meant to provide fluent Chinese translation.
* ♻️ Results are **reproducible** when documents have matching structure and context.
* 🛑 **Error-resilient** — segments that fail translation or matching are skipped without halting the script.


