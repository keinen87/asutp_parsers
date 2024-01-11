from docxtpl import DocxTemplate
import uuid


def create_doc_final(path, context, filename, exists=False):
    doc = DocxTemplate(path)
    doc.render(context)
    extension = '.docx'
    if exists:
        filename += uuid.uuid4().hex
    doc.save(filename+extension)