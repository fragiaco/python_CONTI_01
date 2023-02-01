#leggo i dati dal database e ricavo anche i nomi colonna
import sqlite3


my_data=[]
try:
    conn = sqlite3.connect('database_messe_orizzontale')
    cur = conn.cursor()
    print('connesso')

    r_set=cur.execute("SELECT * FROM TABLE_Messe ")
    my_data.append(list(map(lambda  x: x[0], cur.description))) # add the column names
    for row in r_set.fetchall():
        my_data.append(row) # adding one row



except:
    pass
for record in my_data:
    print(record)

import tkinter  as tk
from tkinter import *
my_w = tk.Tk()
my_w.geometry("400x250")

# for record in my_data:
#         e = tk.Label(my_w, width=200, fg='blue',text=record,anchor='w')
#         e.pack( side = TOP)




conn.close()
my_w.mainloop()




# from reportlab.lib.units import inch
# from reportlab.lib.pagesizes import letter
# from reportlab.platypus import SimpleDocTemplate
# from reportlab.platypus.tables import Table,TableStyle,colors
# #from my_table_data import my_data # import the data
# my_path='C:\\Users\giaco\\PycharmProjects\\python_CONTI_01\\my_pdf.pdf' # change path, file name
#
# my_doc=SimpleDocTemplate(my_path,pagesize=letter)
# c_width=[0.4*inch,1.5*inch,1*inch,1*inch,1*inch]
# t=Table(my_data,rowHeights=20,repeatRows=1,colWidths=c_width)
# t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.lightgreen),
# ('FONTSIZE',(0,0),(-1,-1),10)]))
# elements=[]
# elements.append(t)
# my_doc.build(elements)



# from reportlab.pdfgen import canvas
# c = canvas.Canvas("my_pdf.pdf")
# c.drawString(100,750,"Welcome to Reportlab!")
# #it starts at the bottom left of the page, so for this example,
# # we told it to draw the string
# # 100 points from the left margin and
# # 750 points from the bottom of the page
# c.save()