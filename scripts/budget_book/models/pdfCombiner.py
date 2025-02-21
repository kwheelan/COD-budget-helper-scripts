import PyPDF2
import os
import time
import win32com.client

class PDFTool():

    def convert_doc_to_pdf(doc_path, pdf_output_path):
        word = win32com.client.Dispatch("Word.Application")
        
        try:
            # Open the Word document
            doc = word.Documents.Open(doc_path)
            time.sleep(1)  # Let the document fully open, if necessary

            # Export to PDF using ExportAsFixedFormat
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_output_path,
                ExportFormat=17,  # wdExportFormatPDF = 17
                OpenAfterExport=False,
                OptimizeFor=0,  # wdExportOptimizeForPrint = 0
                Item=7,         # wdExportDocumentContent = 7
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=1,  # wdExportCreateWordBookmarks = 1
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=True
            )
            
            # Close the document
            doc.Close()
        except Exception as e:
            print("An error occurred while converting the document:", e)
        finally:
            # Quit the Word application
            word.Quit()

    def combine_pdfs(pdf_list, output_path):
        merger = PyPDF2.PdfMerger()

        for pdf in pdf_list:
            merger.append(pdf)

        merger.write(output_path)
        merger.close()

    def search(filepath, dept):
        for file in os.listdir(filepath):
            if file.startswith(dept) and '.docx' in file:
                return os.path.join(filepath, file)

    @staticmethod
    def find_doc(filepath, dept):
        res = PDFTool.search(filepath, dept)
        if res:
            return res
        return PDFTool.search(os.path.dirname(filepath), dept)
    
    @staticmethod
    def find_and_convert(filepath, dept, pdf_output_path):
        file = PDFTool.find_doc(filepath, dept)
        PDFTool.convert_doc_to_pdf(file, pdf_output_path)
