import openpyxl
import os

def main():
    
    workbook=openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.cell(10,10,'wdnmd')
    workbook.save('out.xlsx')
main()