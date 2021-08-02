import os
from PyPDF2 import PdfFileMerger


# merge PDFs into one file
def merge_pdfs(pdfs):
    merger = PdfFileMerger()

    # append merger with pdfs
    for pdf in pdfs:
        merger.append(pdf)

    merger.write("Destination/Неймтеги.pdf")
    merger.close()

    # remove temporary pdfs
    for pdf in pdfs:
        os.remove(os.path.join(os.getcwd(), pdf))
