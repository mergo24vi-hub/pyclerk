import os
import re
from docx import Document
from parse_fields import parse_fields

path = r"G:\.shortcut-targets-by-id\1Pcnp8gnqT8NS3Zl5AOanpcBmZLHuuv5I\РО робоча\ОСОБИСТІ ПАПКИ\.Не штатні\.СЗЧ\\"

def parse_employee_form_from_4col_table(file_path: str) -> dict:
    doc = Document(file_path)
    data = {}

    # Мапінг ключових фраз на ідентифікатори полів
    field_keywords = {
        "військове звання": "rank",
        "піб": "full_name",
        "посада": "position",
        "дата народження": "birth_date",
        "місце народження": "birth_place",
        "місце проживання": "residence",
        "іпн": "tax_id",
        "телефон": "phone",
        "група крові": "blood_type",
        "паспорт": "passport",
        "військовий квиток": "military_id",
        "освіта": "education",
        "призваний": "enlistment",
        "участь в ато": "ato_participation",
        "оос": "ato_participation",
        "убд": "combatant_status",
        "посвідчення водія": "driver_license",
        "категорія": "driver_license",
        "сімейний стан": "marital_status",
        "діти": "children",
        "батьки": "parents",
        "брати": "siblings",
        "сестри": "siblings",
        "нагороди": "awards",
        "поранення": "injuries",
        "травмування": "injuries",
        "віросповідання": "religion",
        "розміри": "sizes",
        "одягу": "sizes",
        "алкоголю": "attitude_substances",
        "наркотичних": "attitude_substances",
        "судимості": "convictions",
        "досвід робіт": "work_experience",
        "захоплення": "interests",
        "інтереси": "interests",
        "стан здоров’я": "health",
        "алергія": "health",
    }

    # Функція для нормалізації ключів
    def normalize_key(text):
        text = text.strip().lower()
        for keyword, field in field_keywords.items():
            if keyword in text:
                return field
        return None

    # Обробка всіх таблиць у документі
    for table in doc.tables:
        for row in table.rows:
            if len(row.cells) >= 4:
                key_raw = row.cells[2].text.strip()
                val_raw = row.cells[3].text.strip()
                key = normalize_key(key_raw)
                if key and val_raw:
                    if key in data:
                        data[key] += " " + val_raw
                    else:
                        data[key] = val_raw

    return data

def read_descriptions(folder_path):
    descript_path = os.path.join(folder_path, "descript.ion")

    if not os.path.isfile(descript_path):
        print("❌ Файл descript.ion не знайдено.")
        return

    with open(descript_path, "r", encoding="cp1251", errors="ignore") as f:
        lines = f.readlines()

    print("📁 Папки з описом 'стара анкета':\n")

    for line in lines:
        line = line.strip()
        if not line:
            continue

        folder_name = ""
        description = ""

        # ВАРІАНТ 1: назва в лапках
        match = re.match(r'^"(.+?)"\s+(.*)', line)
        if match:
            folder_name = match.group(1).strip()
            description = match.group(2).strip()
        else:
            # ВАРІАНТ 2: без лапок
            parts = line.split(maxsplit=1)
            if len(parts) == 2:
                folder_name, description = parts[0].strip(), parts[1].strip()
            else:
                continue  # рядок без опису — пропускаємо

        # Перевірка опису
        if "стара анкета" in description.lower():
            folder_path_full = os.path.join(folder_path, folder_name)
            exists = os.path.isdir(folder_path_full)
            status = "✅ Є" if exists else "❌ Немає"
            print(f"{status} | 📌 {folder_name} → {description}")

            if exists:
                # Пошук файлів у підпапках з назвою "стара анкета"
                for root, dirs, files in os.walk(folder_path_full):
                    for file in files:
                        if "стара анкета" in file.lower():
                            full_path = os.path.join(root, file)
                            #print(f"   🔍 Знайдено файл: {full_path}")
                            print(parse_fields(full_path))
                            break

# Приклад виклику
f1 = r"G:\.shortcut-targets-by-id\1Pcnp8gnqT8NS3Zl5AOanpcBmZLHuuv5I\РО робоча\ОСОБИСТІ ПАПКИ\.Не штатні\.СЗЧ\\ГОРШКОВ Роман Валерійович 2024.11.15\стара анкета СДД ГОРШКОВ Роман Валерійович.docx"
f2 = r"G:\.shortcut-targets-by-id\1Pcnp8gnqT8NS3Zl5AOanpcBmZLHuuv5I\РО робоча\ОСОБИСТІ ПАПКИ\.Не штатні\.СЗЧ\\Соловян\стара анкета Особова_справа СОЛОВЯН.docx"
f3 = r"G:\.shortcut-targets-by-id\1Pcnp8gnqT8NS3Zl5AOanpcBmZLHuuv5I\РО робоча\ОСОБИСТІ ПАПКИ\.Не штатні\.СЗЧ\\Сіненко\стара анкета Особова_справа Сіненко.docx"

if __name__ == "__main__":
    #rez = parse_employee_form(r"G:\.shortcut-targets-by-id\1Pcnp8gnqT8NS3Zl5AOanpcBmZLHuuv5I\РО робоча\ОСОБИСТІ ПАПКИ\.Не штатні\.СЗЧ\\ГОРШКОВ Роман Валерійович 2024.11.15\стара анкета СДД ГОРШКОВ Роман Валерійович.docx")
    read_descriptions(path.strip())
