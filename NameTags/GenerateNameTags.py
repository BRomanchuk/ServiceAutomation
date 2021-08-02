import os
import win32com.client as win32  # pip install pywin32


# get name of table in link
def get_table_name(src):
    length = len(src)
    start_index, end_index = 0, 0
    for i in range(length - 1, 0, -1):
        if src[i] == '.':
            end_index = i
            break
    for i in range(end_index - 1, 0, -1):
        if src[i] == '\\':
            start_index = i + 1
            break
    table_name = src[start_index : end_index]
    print(start_index, end_index)
    return table_name


def generate_pdfs(table_src, doc_name):
    working_directory = os.getcwd()
    source_name = get_table_name(table_src) + '.xlsx'
    destination_folder = os.path.join(working_directory, 'TempPDFs')

    sql_statement = "SELECT * FROM [" + source_name +"$]"

    # Create a Word application instance
    wordApp = win32.Dispatch('Word.Application')
    wordApp.Visible = True

    # Open Word Template + Open Data Source
    # sourceDoc = wordApp.Documents.Open(os.path.join(working_directory, 'Service.docx'))
    sourceDoc = wordApp.Documents.Open(os.path.join(working_directory, doc_name))
    mail_merge = sourceDoc.MailMerge
    mail_merge.OpenDataSource(
        os.path.join(working_directory, source_name),
        sql_statement
    )

    record_count = mail_merge.DataSource.RecordCount

    # Perform Mail Merge
    filenames = []
    for i in range(1, record_count + 1):
        mail_merge.DataSource.ActiveRecord = i
        mail_merge.DataSource.FirstRecord = i
        mail_merge.DataSource.LastRecord = i

        mail_merge.Destination = 0
        mail_merge.Execute(False)

        # get record value
        base_name = mail_merge.DataSource.DataFields('name1').Value

        filenames.append("TempPDFs/" + base_name + ".pdf")

        targetDoc = wordApp.ActiveDocument

        # Save Files in Word Doc and PDF
        targetDoc.SaveAs2(os.path.join(destination_folder, base_name + '.docx'), 16)
        targetDoc.ExportAsFixedFormat(os.path.join(destination_folder, base_name), 17)
        # Close target file

        targetDoc.Close(False)

    sourceDoc.MailMerge.MainDocumentType = -1

    # close active document and quit Word application
    wordApp.ActiveDocument.Close(False)
    wordApp.Visible = False

    return filenames