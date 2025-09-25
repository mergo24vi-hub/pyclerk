import win32com.client

def merge_word_documents(doc1_path, doc2_path, output_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Word працює у фоновому режимі
    
    # Відкриваємо перший документ
    doc1 = word.Documents.Open(doc1_path)
    end = doc1.Content.End
    insert_pos = max(0, end - 1)  # не виходимо за межі
    doc1.Range(insert_pos, insert_pos).InsertFile(doc2_path)

    # Зберігаємо об’єднаний документ
    doc1.SaveAs(output_path)
    doc1.Close()
    # Закриваємо Word
    word.Quit()

if __name__=='__main__':
    # Виклик функції
    merge_word_documents(
        "D:\apps\pyclerk\neoapps\babenko.docx",
        "C:\\Users\\mergo24vi\\Documents\\нова особова\\tmp\\БУБЛИК Сергій Миколайович 2.docx",
        "C:\\Users\\mergo24vi\\Documents\\нова особова\\tmp\\БУБЛИК Сергій Миколайович повний.docx"
    )