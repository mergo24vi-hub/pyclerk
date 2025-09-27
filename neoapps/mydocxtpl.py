from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm


def render2doc(prefix, context):
    doc = DocxTemplate("docxtpl.docx")
    context['photo'] = InlineImage(doc,prefix+ "avas/" + context['file'] + ".jpg", width=Mm(50))
    doc.render(context)
    doc.save(prefix +"tmp/"+context['file']+".docx")


if __name__=='__main__':
    context = {
        "піб": "Бабенко Олена Ігорівна",
        "народження": "12.05.1992",
        "убд": "Так",
        "file": "бабенко"
    }
