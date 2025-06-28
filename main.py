# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
print('attempt to pdf')
import pymupdf

def mypdf():
    # Use a breakpoint in the code line below to debug your script.
    # PyMuPDF

    doc = pymupdf.open("babenko.pdf")

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(dpi=200)
        pix.pil_save(f"page_{page_num + 1}.png")


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    mypdf()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
