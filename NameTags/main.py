import os
from PDFsGenerator import generate_pdfs
from PDFsMerger import merge_pdfs


def generate_nametags(table_src, doc_name):
    # generate temp pdfs and get their names
    filenames = generate_pdfs(table_src=table_src, doc_name=doc_name)

    # merge all pdf files into one result
    merge_pdfs(filenames)

    path = os.getcwd()
    path = os.path.realpath(path)
    os.startfile(path + "\\Destination\Неймтеги.pdf")