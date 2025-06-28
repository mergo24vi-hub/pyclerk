from docx import Document
import re

def parse_fields(file_path: str) -> dict:
    doc = Document(file_path)
    data = {}

    # Мапінг ключових фраз на уніфіковані назви полів
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

    list_fields = {"children", "siblings", "awards"}

    # Функція для нормалізації ключів
    def normalize_key(text):
        clean_text = re.sub(r"[^\wа-яіїєґ]", "", text.strip().lower())
        for keyword, field in field_keywords.items():
            if re.sub(r"[^\wа-яіїєґ]", "", keyword.lower()) in clean_text:
                return field
        return None

    # Основний обхід таблиць
    for table in doc.tables:
        for row in table.rows:
            if len(row.cells) >= 4:
                raw_key = row.cells[2].text.strip()
                raw_val = row.cells[3].text.strip().replace('\n', ' ').strip()
                key = normalize_key(raw_key)

                if key and raw_val:
                    if key in list_fields:
                        data.setdefault(key, []).append(raw_val)
                    elif key in data:
                        data[key] += " " + raw_val
                    else:
                        data[key] = raw_val
                elif not key and raw_key:
                    pass; #print(f"[WARN] Невідомий ключ: '{raw_key}'")

    return data

f1 = r"G:\.shortcut-targets-by-id\1Pcnp8gnqT8NS3Zl5AOanpcBmZLHuuv5I\РО робоча\ОСОБИСТІ ПАПКИ\.Не штатні\.СЗЧ\\ГОРШКОВ Роман Валерійович 2024.11.15\стара анкета СДД ГОРШКОВ Роман Валерійович.docx"
f2 = r"G:\.shortcut-targets-by-id\1Pcnp8gnqT8NS3Zl5AOanpcBmZLHuuv5I\РО робоча\ОСОБИСТІ ПАПКИ\.Не штатні\.СЗЧ\\Соловян\стара анкета Особова_справа СОЛОВЯН.docx"
f3 = r"G:\.shortcut-targets-by-id\1Pcnp8gnqT8NS3Zl5AOanpcBmZLHuuv5I\РО робоча\ОСОБИСТІ ПАПКИ\.Не штатні\.СЗЧ\\Сіненко\стара анкета Особова_справа Сіненко.docx"

if __name__=='__main__':
    rez = parse_fields(f2); print(rez)