import pandas as pd
from docx import Document
import os

# === Налаштування ===
excel_file = 'personal/шаблон. сзч.xlsx'  # шлях до твого Excel-файлу
template_file = 'personal/formatted_template.docx'
output_dir = 'generated'
os.makedirs(output_dir, exist_ok=True)

# === Зчитування Excel ===
df = pd.read_excel(excel_file, nrows=3)

# === Обробка кожного рядка ===
for index, row in df.iterrows():
    # Відкриваємо шаблон
    doc = Document(template_file)

    # Словник тегів і значень
    replacements = {
        '«Звання»': row.get('Звання', ''),
        '«ПІБ»': row.get('ПІБ', ''),
        '«Посада»': row.get('Посада', ''),
        '«ДатаНар»': row.get('ДатаНар', ''),
        '«МісцеНар»': row.get('МісцеНар', ''),
        '«Призов»': row.get('Призов', ''),
        '«СімейнийСтан»': row.get('СімейнийСтан', ''),
        '«Дружина»': row.get('Дружина', ''),
        '«Діти»': row.get('Діти', ''),
        '«Батько»': row.get('Батько', ''),
        '«Мати»': row.get('Мати', ''),
        '«Освіта»': row.get('Освіта', ''),
        '«Адреса»': row.get('Адреса', ''),
        '«Телефон»': str(row.get('Телефон', '')),
        '«ІПН»': str(row.get('ІПН', '')),
        '«ГрупаКрові»': row.get('ГрупаКрові', '')
    }

    # Заміна в тексті документа
    for paragraph in doc.paragraphs:
        for key, val in replacements.items():
            if key in paragraph.text:
                for run in paragraph.runs:
                    run.text = run.text.replace(key, str(val))

    # Заміна в таблицях (якщо є)
    for table in doc.tables:
        for row_table in table.rows:
            for cell in row_table.cells:
                for paragraph in cell.paragraphs:
                    for key, val in replacements.items():
                        if key in paragraph.text:
                            for run in paragraph.runs:
                                run.text = run.text.replace(key, str(val))

    # Формуємо ім’я файлу
    pib_safe = str(row.get('ПІБ', 'невідомий')).replace(' ', '_')
    filename = f"анкета_{pib_safe}.docx"
    doc.save(os.path.join(output_dir, filename))

print("✅ Генерація завершена!")
