import pandas as pd
import tabula
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

def pdf_to_excel(pdf_file, excel_file):
    df = tabula.read_pdf(pdf_file, pages='all')
    writer = pd.ExcelWriter(excel_file, engine='openpyxl')
    for i, page in enumerate(df):
        sheet_name = f'Page {i+1}'
        page.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        for j, column in enumerate(page.columns):
            max_length = max(page[column].astype(str).map(len).max(), len(column))
            column_letter = get_column_letter(j+1)
            worksheet.column_dimensions[column_letter].width = max_length+2
            for cell in worksheet[column_letter]:
                cell.alignment = Alignment(wrap_text=True)

    writer.save()
    writer.close()

pdf_file = input("Enter the name of the PDF file: ")
excel_file = input("Enter the name of the Excel file: ")
pdf_to_excel(pdf_file, excel_file)
