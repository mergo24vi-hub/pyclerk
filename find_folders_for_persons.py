import pandas as pd
from process_folders import read_descriptions
# Заміни 'your_file.xlsx' на шлях до твого Excel-файлу
file_path = 'personal/шаблон.сзч.xlsx'

# Зчитуємо Excel-файл
df = pd.read_excel(file_path)#, nrows=15

folders = read_descriptions(); print(folders) #; print(folders[0].split()[0])

# Створюємо нову колонку для назв папок (вставимо в перший стовпець пізніше)
folder_names = []

# Опрацювання кожного рядка
for value in df['ПІБ']:
    if pd.isna(value):
        folder_names.append(None)
    else:
        surname = value.split()[0].lower()
        eqv = next((folder for folder in folders if folder.lower().startswith(surname)), None)
        folder_names.append(eqv)

# Вставляємо список назв папок у перший стовпець DataFrame
df.insert(0, 'Папка', folder_names)

# Перезаписуємо Excel-файл із новою колонкою
df.to_excel(file_path, index=False)