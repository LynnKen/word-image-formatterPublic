import comtypes.client
import comtypes
import os

def convert_doc_to_docx(doc_path):
    comtypes.CoInitialize()  # ← זו השורה החשובה

    # המרה לנתיב מוחלט
    abs_doc_path = os.path.abspath(doc_path)
    
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    
    try:
        doc = word.Documents.Open(abs_doc_path)
        new_path = abs_doc_path + "x"
        doc.SaveAs(new_path, FileFormat=16)
        doc.Close()
        return new_path
    finally:
        word.Quit()