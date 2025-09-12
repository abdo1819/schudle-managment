"""
Module for document conversion tasks.
"""

import os
import pythoncom
import win32com.client

def convert_to_pdf_and_open(docx_path):
    """Convert DOCX to PDF and open the PDF."""
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.visible = False
        doc_path_abs = os.path.abspath(docx_path)
        pdf_path_abs = os.path.splitext(doc_path_abs)[0] + ".pdf"
        
        print(f"\n🔄 Converting {os.path.basename(docx_path)} to PDF...")
        doc = word.Documents.Open(doc_path_abs)
        doc.SaveAs(pdf_path_abs, FileFormat=17)  # 17 is the PDF format
        doc.Close()
        print(f"📄 PDF created: {os.path.basename(pdf_path_abs)}")
        
        # Open the PDF file
        print(f"🚀 Opening {os.path.basename(pdf_path_abs)}...")
        os.startfile(pdf_path_abs)
        
    except Exception as e:
        print(f"❌ PDF conversion/opening failed: {e}")
    finally:
        if word:
            word.Quit()
        pythoncom.CoUninitialize()
