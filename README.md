# 🈺 English-to-Chinese Track Changes Automation

A Python tool that **automatically transfers tracked changes** from an **English Word document** (`.docx`) to its corresponding **Chinese translation**, preserving Microsoft Word’s **Track Changes** format. It handles insertions, deletions, replacements, and even formatting notices.

---

## 📌 Overview

This script automates the transfer of tracked changes between aligned bilingual `.docx` files:

* **Input**:

  * `edited_en.docx` — English document with **tracked changes**
  * `original_cn.docx` — Original Chinese version

* **Output**:

  * `original_cn_with_tracked_changes.docx` — Chinese document with equivalent tracked changes applied, aligned to context.

---

## ✅ Features

* 🔁 **Fully Automated**: No manual editing required — extract, align, translate, and apply changes automatically.
* 📝 **Track Changes Preserved**: Insertions, deletions, replacements, and formatting markers applied as tracked changes in the output.
* 🌐 **Context-Aware Matching**: Uses `difflib` to match translated paragraphs to Chinese paragraphs for accurate positioning.
* 🧠 **Bold Marker Detection**: Marks bold-formatted segments by adding `[bold]` annotations in the Chinese translation.
* 🧮 **Change Summary Logging**: At the end, a summary of applied changes is displayed (insert/delete/replace/format/skipped).
* 💬 **Command-Line and Interactive Support**: Accepts CLI arguments or prompts user input for file paths.

---

## ⚙️ Requirements

* **Python**: 3.11 or 3.12

  > ⚠️ Python 3.13+ is **not supported** due to library compatibility issues.

* **Microsoft Word** installed (script uses Word automation via `win32com`)

### 📦 Install Dependencies

```bash
pip install pywin32 googletrans==4.0.0-rc1

## 🚀 How to Use

###Input required file path

```bash
python main.py
# Prompts:
# Enter path to English Word document (.docx): C:\path\to\English_document.docx (demo) 
# Enter path to Chinese Word document (.docx): C:\path\to\Chinese_document.docx (demo)

###Output
Chinese_document_with_tracked_changes.docx

## 📂 File Structure

| File Name                               | Description                                       |
| --------------------------------------- | ------------------------------------------------- |
| `main.py`                               | Main script that performs extraction and transfer |
| `English_document.docx`                        | English `.docx` with tracked changes              |
| `Chinese_document.docx`                      | Chinese `.docx` with no changes                   |
| `Chinese_document_with_tracked_changes.docx` | Output with mirrored tracked changes              |


## 📊 Output Example

At the end of processing, you'll see a summary:

📊 Change Summary:
  - insert: 4
  - delete: 2
  - replace: 1
  - format: 1
  - bold: 1
  - skipped: 0


## ⚠️ Notes & Recommendations

* ✅ **Close all Word windows** before running — COM automation may fail if Word is already open.
* ❌ Do not use Python 3.13+ — compatibility issues with COM and `googletrans`.
* 📄 Translation is used **only for paragraph alignment** — not for producing fluent Chinese output.
* 🧪 For reproducibility: results are deterministic for the same document pair.
* 📉 If translation fails, the script reports and skips problematic segments without crashing.


