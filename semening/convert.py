import subprocess
import os
libreoffice_path = r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
from pathlib import Path

def convert_doc_to_txt(input_path, output_dir=None):
    if output_dir is None:
        output_dir = os.path.dirname(input_path)

    result = subprocess.run([
        libreoffice_path, "--headless", "--convert-to", "txt:Text", "--outdir", output_dir, input_path
    ], capture_output=True, text=True)

    if result.returncode != 0:
        print("Error:", result.stderr)
    else:
        print("Converted:", os.path.join(output_dir, os.path.splitext(os.path.basename(input_path))[0] + ".txt"))

    with open("test.txt", "r", encoding="windows-1251") as f:
        text = f.read()
        print(text)

convert_doc_to_txt(r'D:\apps\pyclerk\semening\test.docx')