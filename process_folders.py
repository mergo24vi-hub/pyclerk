import os
import re
path = r"G:\.shortcut-targets-by-id\1Pcnp8gnqT8NS3Zl5AOanpcBmZLHuuv5I\РО робоча\ОСОБИСТІ ПАПКИ\.Не штатні\.СЗЧ\\"

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


# Приклад виклику
if __name__ == "__main__":
    read_descriptions(path.strip())
