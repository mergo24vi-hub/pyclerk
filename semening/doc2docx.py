import os
import win32com.client

def convert_doc_to_docx(doc_path):
    # Перевірка розширення
    if not doc_path.lower().endswith(".doc") or doc_path.lower().endswith(".docx"):
        raise ValueError("Файл має мати розширення .doc")

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    abs_path = os.path.abspath(doc_path)
    doc = word.Documents.Open(abs_path)

    # Створення шляху для .docx
    new_path = os.path.splitext(abs_path)[0] + ".docx"

    # Збереження у форматі docx (FileFormat = 16)
    doc.SaveAs(new_path, FileFormat=16)
    doc.Close()
    word.Quit()
    os.remove(abs_path)

    print("Конвертовано:", new_path)
    return new_path

#convert_doc_to_docx(r'd:\apps\pyclerk\semening\2Акт зміни якісного стану 1420 Е2.doc')