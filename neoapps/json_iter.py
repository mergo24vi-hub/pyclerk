import json5

# Відкрити файл для читання
with open("scannedstuff.json", "r", encoding="utf-8") as f:
    data = json5.load(f)   # Перетворює JSON у Python-об'єкт (dict / list)

for row in data:
    print(row['ПІБ'])