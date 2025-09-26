import csv
from mydocxtpl import render2doc
from find_scans import finduascan
from append_wrd import merge_word_documents
from pathlib import Path
filename = "soldiers06.csv"
with open(filename, newline='', encoding='utf-8-sig') as csvfile:
    reader = csv.DictReader(csvfile, delimiter=';')# Якщо у вас роздільник кома - залишаємо ','; якщо крапка з комою - змініть на ';'
    i=0
    for row in reader:
        #print(row); break
        pib = row.get("піб")

        p = pib.split(' ')[0].lower()
        row['file']= p
        scans = finduascan(p)
        if not scans:
            continue
        #print(row.get("photo"), p)
        print(row['file'])
        render2doc(row)
        in1 = str(Path(row.get("file")+".docx").resolve())
        in2 = str(Path(scans).resolve())
        out = str(Path('out/'+row.get("піб")+".docx").resolve())
        print(in1, in2, out)
        merge_word_documents(in1,in2,out)
        i=i+1
        if i>50:
            break


