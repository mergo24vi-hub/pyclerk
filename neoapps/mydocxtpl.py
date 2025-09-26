from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm


def render2doc(context):
    doc = DocxTemplate("docxtpl.docx")
    context['photo'] = InlineImage(doc, 'uaavas/' + context['file'] + ".jpg", width=Mm(50))
    doc.render(context)
    doc.save(context['file']+".docx")


if __name__=='__main__':
    context = {
        "піб": "Бабенко Олена Ігорівна",
        "народження": "12.05.1992",
        "убд": "Так",
        "file": "бабенко"
    }
    render2doc(context)