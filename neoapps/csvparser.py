import csv
from mydocxtpl import render2doc
from find_scans import findscan, findava
from append_wrd import merge_word_documents
from pathlib import Path

prefix = "ua"
filename = "latinos.csv" if prefix == "es" else "slaves.csv"
with open(filename, newline='', encoding='utf-8-sig') as csvfile:
    reader = csv.DictReader(csvfile, delimiter=';')# Якщо у вас роздільник кома - залишаємо ','; якщо крапка з комою - змініть на ';'
    i=0
    for row in reader:
        #print(row); break
        pib = row.get("піб")

        p = pib.split(',' if prefix=="es" else " ")[0].lower()
        row['file']= p
        scans = findscan(prefix, p)
        ava = findava(prefix, p)
        print(scans, ava)
        if not scans or not ava:
            break
        render2doc(prefix, row)
        i = i + 1

        in1 = str(Path(prefix+"tmp/"+row.get("file")+".docx").resolve())
        in2 = str(Path(scans).resolve())
        out = str(Path(prefix+'out/'+row.get("піб")+".docx").resolve())
        print(in1, in2, out)
        merge_word_documents(in1,in2,out)
        if i>2:
            break
        continue



