import xlsxwriter

workbook  = xlsxwriter.Workbook('filename.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write(0, 0, 'Hello Excel')

workbook.close()