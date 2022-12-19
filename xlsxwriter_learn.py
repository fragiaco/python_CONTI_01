from xlsxwriter import *
import pandas as pd
import sqlite3
conn = sqlite3.connect('database_conti')


# Create a cursor instance
c = conn.cursor()

query = "SELECT * FROM TABLE_Conti"  # query to collect recors
df = pd.read_sql(query, conn)
# Create a Pandas dataframe from the data.
#df = pd.DataFrame({'Data': [10, 20, 30, 20, 15, 30, 45]})

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Dati')

# Get the xlsxwriter objects from the dataframe writer object.
workbook  = writer.book
worksheet = writer.sheets['Dati']

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Add a number format for cells with money.
money = workbook.add_format({'num_format': '$#,##0'})

# data_cols=['Anno', 'Mese', 'Entrate_Uscite', 'Categoria', 'Voce', 'Euro']
worksheet.write(1, 9, 'Hello', bold)  # Writes a string
worksheet.write('J3', 'Item', bold)
worksheet.write('H42', 'Total', bold)
worksheet.write('H43', '=SUM(H2:H40)', money)
worksheet.set_column(1, 1, 15)



cell_format_red = workbook.add_format({'bold': True, 'font_color': 'red'})
cell_format_black = workbook.add_format({'bold': True, 'font_color': 'black'})

worksheet.write('H42', 'Total', cell_format_red)
#Formats can also be passed to the worksheet set_row() and set_column() methods
# to define the default formatting properties for a row or column:

worksheet.set_row(0, 18, cell_format_black)
worksheet.set_column('B:C', 15, cell_format_red)


# Get the dimensions of the dataframe.
(max_row, max_col) = df.shape

# Write a total using a formula.




# Apply a conditional format to the required cell range.
# worksheet.conditional_format(1, max_col, max_row, max_col,
#                              {'type': '3_color_scale'})

# Close the Pandas Excel writer and output the Excel file.
writer.close()

# Commit changes
conn.commit()

# Close our connection
conn.close()