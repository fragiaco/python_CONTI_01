from xlsxwriter import *
from xlsxwriter.utility import xl_rowcol_to_cell
import pandas as pd
import sqlite3


conn = sqlite3.connect('database_conti')

# Create a cursor instance
c = conn.cursor()

query = "SELECT * FROM TABLE_Conti"  # query to collect recors
df = pd.read_sql(query, conn)
# print(df.dtypes)
# df['Anno'] = df['Anno'].astype(int)
# #pd.to_datetime(df['Anno'],format="%Y/%m/%d")
# print(df.dtypes)


startrowval = 2 # index starts from zero

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, startrow=startrowval, index=False, sheet_name='Dati')

# Get the xlsxwriter objects from the dataframe writer object.
workbook = writer.book
worksheet = writer.sheets['Dati']
worksheet.set_zoom(160)
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
worksheet.merge_range('A1:G1', title, format)
worksheet.merge_range('A2:G2', subheader)
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
worksheet.set_column('E:E', 17)
worksheet.set_column('F:F', 25)



# Add a number format for cells with money.
money_fmt = workbook.add_format({'num_format': '#,##0.00'})

worksheet.set_column('G:G', 9, money_fmt)

# Light red fill with dark red text.
value_uscite = workbook.add_format({'bg_color':   '#FFC7CE', 'font_color': '#9C0006'})
worksheet.conditional_format('D28:D41', {'type':     'text', 'criteria': 'containing', 'value':    'Entrate', 'format': value_uscite})
# # add borders
# worksheet.conditional_format('A4:H27', {'type':  'formula','criteria': '=$H4<1000','format':   full_border})

# # data_cols=['Anno', 'Mese', 'Entrate_Uscite', 'Categoria', 'Voce', 'Euro']

# # Apply a conditional format to the required cell range.
# # worksheet.conditional_format(1, max_col, max_row, max_col,
# #                              {'type': '3_color_scale'})


# Get the dimensions of the dataframe.
(max_row, max_col) = df.shape
print('max_row is:=', max_row) #max_row is:= 41
print('max_col is:=', max_col) #max_row is:= 7
# Set the autofilter.
worksheet.autofilter(2, 1, max_row, max_col - 1)



# # Close the Pandas Excel writer and output the Excel file.
writer.close()


# Commit changes
conn.commit()

# Close our connection
conn.close()