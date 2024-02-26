import tkinter as tk
from tkinter import filedialog
import win32com.client

def excel_to_pdf():
    excel_file = filedialog.askopenfilename(title="Select Excel file", filetypes=(("Excel files", "*.xlsx;*.xls"), ("all files", "*.*")))
    if not excel_file:
        return
    
    pdf_file = filedialog.asksaveasfilename(title="Save PDF file as", defaultextension=".pdf", filetypes=(("PDF files", "*.pdf"), ("all files", "*.*")))
    if not pdf_file:
        return
    
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(excel_file)
    
    # 모든 시트를 선택
    ws = wb.Worksheets
    ws_names = [sheet.Name for sheet in ws]
    
    # PDF로 저장
    wb.WorkSheets(ws_names).Select()
    wb.ActiveSheet.ExportAsFixedFormat(0, pdf_file)
    
    wb.Close()
    excel.Quit()
    
    result_label.config(text="Conversion completed successfully!")

# GUI 생성
root = tk.Tk()
root.title("Excel to PDF Converter")

convert_button = tk.Button(root, text="Convert to PDF", command=excel_to_pdf)
convert_button.pack(pady=10)

result_label = tk.Label(root, text="")
result_label.pack()

root.mainloop()
