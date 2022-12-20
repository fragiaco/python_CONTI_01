from xlsxwriter import *
from xlsxwriter.utility import xl_rowcol_to_cell
import pandas as pd
import sqlite3
conn = sqlite3.connect('database_conti')


# Create a cursor instance
c = conn.cursor()

query = "SELECT * FROM TABLE_Conti"  # query to collect recors
df = pd.read_sql(query, conn)


startrowval = 2 # index starts from zero

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, startrow=startrowval, index=False, sheet_name='Dati')

# Get the xlsxwriter objects from the dataframe writer object.
workbook = writer.book
worksheet = writer.sheets['Dati']
worksheet.set_zoom(120)
#Set header formating
header_format = workbook.add_format({
        "valign": "vcenter", #Vertical align	'valign'
        "align": "center", #	Horizontal align	'align'
        "bg_color": "#951F06",
         "bold": True,
        'font_color': '#FFFFFF'
    })

#add title
title = "Conti Convento"
#merge cells
format = workbook.add_format()
format.set_font_size(25)
format.set_font_color("#333333")
#
subheader = "Utilizzare il Filtro per ricercare i dati"
worksheet.merge_range('A1:I1', title, format)
worksheet.merge_range('A2:I2', subheader)
worksheet.set_row(2, 20) # Set the header row height to 20
# puting it all together
# Write the column headers with the defined format.
for col_num, value in enumerate(df.columns.values):#https://xlsxwriter.readthedocs.io/working_with_data.html
    #print(col_num, value)
    worksheet.write(startrowval, col_num, value, header_format)
# # Get the dimensions of the dataframe.
# (max_row, max_col) = df.shape
# row = next(df.iterrows())[1]
# for row_num, data in enumerate(df.row):
#     worksheet.write(row_num, 0, data)


# Adjust the column width.
worksheet.set_column('A:A', 3)
worksheet.set_column('D:D', 12)
worksheet.set_column('E:E', 18)
worksheet.set_column('F:F', 25)
worksheet.set_column('G:G', 6)


# Add a number format for cells with money.
money_fmt = workbook.add_format({'num_format': '#,##0.00'})

worksheet.set_column('G:G', 15, money_fmt)
#worksheet.conditional_format(1, max_col, max_row, max_col,
# #                              {'type': '3_color_scale'})
#worksheet.set_column(1, 1, 18, money_fmt)

#worksheet.set_column('G:G', 15, money_fmt)
# # Add a bold format to use to highlight cells.
# bold = workbook.add_format({'bold': True})
#
# # Add a number format for cells with money.
# money = workbook.add_format({'num_format': '$#,##0'})
#
# # data_cols=['Anno', 'Mese', 'Entrate_Uscite', 'Categoria', 'Voce', 'Euro']
# worksheet.write(1, 9, 'Hello', bold)  # Writes a string
# worksheet.write('J3', 'Item', bold)
# worksheet.write('H42', 'Total', bold)
# worksheet.write('H43', '=SUM(H2:H40)', money)
# worksheet.set_column(1, 1, 15)
#
#
#
# cell_format_red = workbook.add_format({'bold': True, 'font_color': 'red'})
# cell_format_black = workbook.add_format({'bold': True, 'font_color': 'black'})
#
# worksheet.write('H42', 'Total', cell_format_red)
# #Formats can also be passed to the worksheet set_row() and set_column() methods
# # to define the default formatting properties for a row or column:
#
# worksheet.set_row(0, 18, cell_format_black)
# worksheet.set_column('B:C', 15, cell_format_red)
#
# for row in range(0, 5):
#     worksheet.write(row, 0, 'Hello')
#
# # Get the dimensions of the dataframe.
# (max_row, max_col) = df.shape
#
# # Write a total using a formula.
#
#
#
#
# # Apply a conditional format to the required cell range.
# # worksheet.conditional_format(1, max_col, max_row, max_col,
# #                              {'type': '3_color_scale'})
#
# # Close the Pandas Excel writer and output the Excel file.
writer.close()

# Commit changes
conn.commit()

# Close our connection
conn.close()