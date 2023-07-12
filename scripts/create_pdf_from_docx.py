import win32com.client
from pathlib import Path


def create_pdf_from_docx(docx_input, pdf_output):
    word = win32com.client.Dispatch("Word.Application")
    wdFormatPDF = 17

    docx_filepath = Path(docx_input).resolve()
    pdf_filepath = Path(pdf_output).resolve()
    doc = word.Documents.Open(str(docx_filepath))
    doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
    doc.Close(0)
