import csv
filename = "jsonstuff.csv"
from mydocxtpl import render2doc
from find_scans import find
from append_wrd import merge_word_documents
from pathlib import Path

with open(filename, newline='', encoding='utf-8-sig') as csvfile:
    reader = csv.DictReader(csvfile, delimiter=';')# Якщо у вас роздільник кома - залишаємо ','; якщо крапка з комою - змініть на ';'
    for row in reader:
        #print(row); break
        # Наприклад, якщо є колонка "Прізвище Ім'я По-батькові"
        if row.get("file"):
            pib = row.get("піб")
            p = pib.split(' ')[0].lower()
            scans = find(p)
            if not scans:
                continue
            print(Path(scans).resolve())
            #print(row.get("photo"), p)
            render2doc(row)
            in1 = str(Path(row.get("file")+".docx").resolve())
            in2 = str(Path(scans).resolve())
            out = str(Path(row.get("піб")+".docx").resolve())
            print(in1, in2, out)
            merge_word_documents(in1,in2,out)

            break

