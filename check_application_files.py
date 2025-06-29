import pandas as pd
from docx import Document
import os
from pathlib import Path
from datetime import datetime

# === Налаштування ===
root_folder = 'generated'
root_folder = r'G:\.shortcut-targets-by-id\1Pcnp8gnqT8NS3Zl5AOanpcBmZLHuuv5I\РО робоча\ОСОБИСТІ ПАПКИ\.Не штатні\.СЗЧ\\'

for item in os.listdir(root_folder):
        full_path = os.path.join(root_folder, item)
        if os.path.isdir(full_path):
            found = False
            for filename in os.listdir(full_path):
                if filename.lower().endswith('.docx') and filename.lower().startswith('сдд'):
                    # print(f'✅ Файл знайдено в папці: {full_path} -> {filename}')
                    found = True
                    break
            if not found:
                print(f'❌ У папці {full_path} немає файлів .docx, що починаються з "сдд"')

print("✅ Перевірка завершена!")

