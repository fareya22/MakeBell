import os
import difflib
import win32com.client as win32
from googletrans import Translator

translator = Translator()

change_count = {
    'insert': 0,
    'delete': 0,
    'replace': 0,
    'format': 0,
    'bold': 0,
    'skipped': 0
}


def extract_changes_from_word(path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(path)
    changes = []

    revisions = list(doc.Revisions)
    skip_next = False

    for i, rev in enumerate(revisions):
        if skip_next:
            skip_next = False
            continue

        try:
            rev_type = rev.Type
            rev_text = rev.Range.Text.strip()
            context = rev.Range.Paragraphs(1).Range.Text.strip()
            formatted = rev.Range.Font.Bold

            # Handle replacements (delete + insert in same context)
            if rev_type == 2 and i + 1 < len(revisions):
                next_rev = revisions[i + 1]
                if next_rev.Type == 1:
                    if context == next_rev.Range.Paragraphs(1).Range.Text.strip():
                        changes.append({
                            'type': 'replace',
                            'text_deleted': rev_text,
                            'text_inserted': next_rev.Range.Text.strip(),
                            'context': context,
                            'bold': bool(next_rev.Range.Font.Bold)
                        })
                        skip_next = True
                        continue

            # Regular revision types
            if rev_type == 1:
                changes.append({'type': 'insert', 'text': rev_text, 'context': context, 'bold': bool(formatted)})
            elif rev_type == 2:
                changes.append({'type': 'delete', 'text': rev_text, 'context': context, 'bold': bool(formatted)})
            elif rev_type in (3, 4, 5, 6):
                changes.append({'type': 'format', 'text': rev_text, 'context': context, 'bold': bool(formatted)})

        except Exception as e:
            print(f"⚠️ Error on revision {i}: {e}")

    doc.Close(False)
    word.Quit()
    return changes


def translate_text(text):
    try:
        return translator.translate(text, src='en', dest='zh-cn').text
    except Exception as e:
       # print(f"❌ Couldn't translate '{text}': {e}")
        return None


def find_best_match(target, paragraph_list):
    best_score = 0
    best_match = None
    for para in paragraph_list:
        score = difflib.SequenceMatcher(None, target, para).ratio()
        if score > best_score:
            best_score = score
            best_match = para
    return best_match


def apply_changes_to_chinese(chinese_doc_path, changes):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(chinese_doc_path)
    doc.TrackRevisions = True

    paras = [p.Range.Text.strip() for p in doc.Paragraphs]

    for change in changes:
        if 'type' not in change:
            change_count['skipped'] += 1
            continue

        zh_context = translate_text(change.get('context'))
        if change['type'] == 'replace':
            zh_old = translate_text(change.get('text_deleted'))
            zh_new = translate_text(change.get('text_inserted'))
            if not all([zh_old, zh_new, zh_context]):
                change_count['skipped'] += 1
                continue
        else:
            zh_text = translate_text(change.get('text'))
            if not zh_text or not zh_context:
                change_count['skipped'] += 1
                continue

        best_para = find_best_match(zh_context, paras)
        if not best_para:
            change_count['skipped'] += 1
            continue

        for p in doc.Paragraphs:
            if best_para.strip() == p.Range.Text.strip():
                rng = p.Range

                if change['type'] == 'delete':
                    start = rng.Text.find(zh_text)
                    if start >= 0:
                        del_range = rng.Duplicate
                        del_range.SetRange(rng.Start + start, rng.Start + start + len(zh_text))
                        del_range.Delete()
                        change_count['delete'] += 1

                elif change['type'] == 'insert':
                    rng.InsertAfter(zh_text)
                    change_count['insert'] += 1

                elif change['type'] == 'replace':
                    start = rng.Text.find(zh_old)
                    if start >= 0:
                        rep_range = rng.Duplicate
                        rep_range.SetRange(rng.Start + start, rng.Start + start + len(zh_old))
                        rep_range.Text = zh_new
                        change_count['replace'] += 1
                    else:
                        change_count['skipped'] += 1

                elif change['type'] == 'format':
                    formatted_text = f"[BOLD:{zh_text}]" if change.get('bold') else f"[FORMATTED:{zh_text}]"
                    rng.InsertAfter(formatted_text)
                    change_count['bold' if change.get('bold') else 'format'] += 1

                break

   # Summary
    print("\n📊 Change Detection SumUp:")
    for k, v in change_count.items():
        print(f"  - {k}: {v}")

    output_path = chinese_doc_path.replace('.docx', '_with_tracked_changes.docx')
    doc.SaveAs(output_path)
    doc.Close()
    word.Quit()
    return output_path


def main():
    print("📂 Please provide full path to the English (edited) .docx file:")
    english_doc = input("English file path: ").strip('"')

    print("📂 Please provide full path to the Chinese (original) .docx file:")
    chinese_doc = input("Chinese file path: ").strip('"')

    if not os.path.isfile(english_doc) or not os.path.isfile(chinese_doc):
        print("❌ One or both file paths are invalid.")
        return

    print("\n📥 Extracting changes from English document...")
    changes = extract_changes_from_word(english_doc)

    print("\n🌐 Applying tracked changes to Chinese document. Please wait until I provide you the updated Chinese Document...")
    result_path = apply_changes_to_chinese(chinese_doc, changes)

    print(f"\nDone✅ ! See the output in: {result_path}")


if __name__ == "__main__":
    main()
