from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import reportlab
import sqlite3
import pandas as pd
import os, sys, subprocess


from tkinter import ttk
#importare prima ttkwidgets
from ttkwidgets.autocomplete import AutocompleteEntry
from ttkwidgets.autocomplete import AutocompleteCombobox


from openpyxl.styles import Font, Alignment
from openpyxl.styles import Side, Border

from openpyxl import styles, formatting
import pandas as pd
import numpy as np

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle, Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image

from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.worksheet import Worksheet

# Leggi il file xlsx e trasformalo in dataframe impostando i nomi colonna
from openpyxl.worksheet.worksheet import Worksheet

from xlsxwriter.utility import xl_rowcol_to_cell
from pandastable import Table, TableModel, config

######################################################################
######## FUNZIONE CONNESSIONE AL DATABASE 'database_conti' ###########
######################################################################
def connessione():
    conn = sqlite3.connect('database_messe_orizzontale')

    cur = conn.cursor()
    try:
        cur.execute('''CREATE TABLE TABLE_Messe(ID integer not null PRIMARY KEY ,
                                                Anno integer not null ,
                                                Mese TEXT not null ,
                                                Nome_Celebrante TEXT not null ,
                                                Ad_Mentem integer not null ,
                                                Binate integer,
                                                Binate_Concelebrate integer,
                                                Trinate integer,
                                                Suffragi_Comunitari integer,
                                                Suffragi_Personali integer,
                                                Devozione integer,
                                                Benefattori integer,
                                                Pro_Populo integer)''')

    except:
        pass

    try:
        cur.execute('''CREATE TABLE TABLE_Celebranti(ID integer not null PRIMARY KEY ,
                                                        Celebranti TEXT not null )''')
    except:
        pass

    try:
        cur.execute('''CREATE TABLE TABLE_Suffragi(ID integer not null PRIMARY KEY ,
                                                    Anno integer not null ,
                                                    Mese TEXT not null ,
                                                    Suffragi TEXT not null )''')
    except:
        pass

    print(conn)
    print('Sei connesso al database_conti')
    conn.commit()
    conn.close()


connessione()

############################
######## TKINTER ###########
############################

root = Tk()

# DEFINISCO le dimensioni della GUI e il TITOLO nella barra superiore
height = 950  # altezza
width = 1680  # larghezza
top = 0
left = (root.winfo_screenwidth() - width) / 2
geometry = ("{}x{}+{}+{}".format(width, height, int(left), int(top)))
root.geometry(geometry)
root.resizable(0, 0)
root.title('Registro Messe')

foreground_Bianco = '#ffeddb'
background_Blu = 'blue'
# Label title
title = Label(root, text='Registro Messe', font=('verdana', 40, 'bold'), bg=background_Blu, fg=foreground_Bianco)
title.pack(side=TOP, fill=X)

###################################
######## TKINTER frames ###########
###################################

# Frame Combo - Top
Frame_combo = Frame(root, bd='4', bg=background_Blu, relief=RIDGE)
Frame_combo.place(x=5, y=73, width=1670, height=60)

# Frame Treeview - treeview right Frame
Frame_tree = Frame(root, bd='4', bg=background_Blu, relief=RIDGE)
Frame_tree.place(x=5, y=132, width=1570, height=335)
Frame_tree_Buttons= Frame(root, bd='4', bg=background_Blu, relief=RIDGE)
Frame_tree_Buttons.place(x=1575, y=132, width=100, height=335)


#
Frame_pandastable = Frame(root, bd='4', bg=background_Blu, relief=RIDGE)
Frame_pandastable.place(x=5, y=467, width=1670, height=480)

# Frame Update - update right Frame
Frame_Suffragi = Frame(Frame_pandastable, bd='4', bg=background_Blu, relief=RIDGE)
Frame_Suffragi.place(x=0, y=0, width=470, height=473)


###############################################
######## TKINTER ENTRY Frame_combo ############
###############################################
Entry_Anno_combo_IntVar                 = IntVar()
Entry_Mese_combo_StringVar              = StringVar()
Entry_Nome_Celebrante_combo_StringVar   = StringVar()
Entry_Ad_Mentem_combo_IntVar            = IntVar()
Entry_Binate_combo_IntVar               = IntVar()
Entry_Binate_Conc_combo_IntVar          = IntVar()
Entry_Trinate_combo_IntVar              = IntVar()
Entry_Suffragi_Comunitari_combo_IntVar  = IntVar()
Entry_Suffragi_Personali_combo_IntVar   = IntVar()
Entry_Devozione_combo_IntVar            = IntVar()
Entry_Benefattori_combo_IntVar          = IntVar()
Entry_Pro_Populo_combo_IntVar           = IntVar()


Label_TOTALE_Numero_Messe= Label(Frame_combo, text=' ', font=('verdana', 8, 'bold'),
                                                bg=background_Blu, fg=foreground_Bianco)
Label_TOTALE_Numero_Messe.grid(row=2, column=12, columnspan=2, padx=40, pady=10)


def trace_when_Entry_widget_is_updated(self, *args):
    try:
        Label_TOTALE_Numero_Messe.config(text=' ', font=('verdana', 16, 'bold'), bg=background_Blu, fg=foreground_Bianco)
        value = Entry_Ad_Mentem_combo_IntVar.get()+\
                Entry_Binate_combo_IntVar.get()+\
                Entry_Binate_Conc_combo_IntVar.get()+\
                Entry_Trinate_combo_IntVar.get()+\
                Entry_Suffragi_Comunitari_combo_IntVar.get()+\
                Entry_Suffragi_Personali_combo_IntVar.get()+\
                Entry_Devozione_combo_IntVar.get()+\
                Entry_Benefattori_combo_IntVar.get()+\
                Entry_Pro_Populo_combo_IntVar.get()

        text = "{}".format(value) if value else " "
        Label_TOTALE_Numero_Messe.config(text=text)

    except:
        pass


Entry_Anno_combo_IntVar.trace_variable('w', trace_when_Entry_widget_is_updated)
Entry_Mese_combo_StringVar.trace_variable('w', trace_when_Entry_widget_is_updated)
Entry_Nome_Celebrante_combo_StringVar.trace_variable('w', trace_when_Entry_widget_is_updated)
Entry_Ad_Mentem_combo_IntVar.trace_variable('w', trace_when_Entry_widget_is_updated)
Entry_Binate_combo_IntVar.trace_variable('w', trace_when_Entry_widget_is_updated)
Entry_Binate_Conc_combo_IntVar.trace_variable('w', trace_when_Entry_widget_is_updated)
Entry_Trinate_combo_IntVar.trace_variable('w', trace_when_Entry_widget_is_updated)
Entry_Suffragi_Comunitari_combo_IntVar.trace_variable('w', trace_when_Entry_widget_is_updated)
Entry_Suffragi_Personali_combo_IntVar.trace_variable('w', trace_when_Entry_widget_is_updated)
Entry_Devozione_combo_IntVar.trace_variable('w', trace_when_Entry_widget_is_updated)
Entry_Benefattori_combo_IntVar.trace_variable('w', trace_when_Entry_widget_is_updated)
Entry_Pro_Populo_combo_IntVar.trace_variable('w', trace_when_Entry_widget_is_updated)

# List Anni
Anni = [2020,2021,2022,2023,2024,2025,2026,2027,2028,2029,2030]

# Dropbox Anno
Entry_Anno_combo = ttk.Combobox(Frame_combo, font=("Helvetica", 10), values=Anni, textvariable=Entry_Anno_combo_IntVar)
Entry_Anno_combo.current(4)
Entry_Anno_combo.grid(row=2, column=0)
Entry_Anno_combo['state'] = 'readonly'


# List Mesi
Mesi = ["gennaio",
        "febbraio",
        "marzo",
        "aprile",
        "maggio",
        "giugno",
        "luglio",
        "agosto",
        "settembre",
        "ottobre",
        "novembre",
        "dicembre",
        ]



# Dropbox Mesi
#Entry_Mese_combo = ttk.Combobox(Frame_combo, font=("Helvetica", 10), values=Mesi, textvariable=Entry_Mese_combo_StringVar)
Entry_Mese_combo = AutocompleteCombobox(Frame_combo, font=("Helvetica", 10), completevalues=Mesi, textvariable=Entry_Mese_combo_StringVar)

Entry_Mese_combo.current(0)
Entry_Mese_combo.grid(row=2, column=1)
#Entry_Mese_combo['state'] = 'readonly'

conn = sqlite3.connect('database_messe_orizzontale')
cur = conn.cursor()
query = "SELECT DISTINCT (Celebranti) as Celebranti FROM TABLE_Celebranti"
my_Data = cur.execute(query)
Nomi_Celebranti = [r for r, in my_Data]
Entry_Nome_Celebrante_combo = ttk.Combobox(Frame_combo, font=("Helvetica", 10), values=Nomi_Celebranti, textvariable=Entry_Nome_Celebrante_combo_StringVar)
Entry_Nome_Celebrante_combo.current(1)
Entry_Nome_Celebrante_combo.grid(row=2, column=2)
# Commit changes
conn.commit()
# Close our connection
conn.close()

Entry_Ad_Mentem_combo = Spinbox(Frame_combo, from_=0, to=31, wrap=True, width=11, font=("Helvetica", 12, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Ad_Mentem_combo_IntVar)
Entry_Ad_Mentem_combo.grid\
    (row=2, column=3)
Entry_Binate_combo = Spinbox(Frame_combo,from_=0, to=31,wrap=True, width=10, font=("Helvetica", 12, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Binate_combo_IntVar)
Entry_Binate_combo.grid\
    (row=2, column=4)
Entry_Binate_Conc_combo = Spinbox(Frame_combo,from_=0, to=31,wrap=True, width=11, font=("Helvetica", 12, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Binate_Conc_combo_IntVar)
Entry_Binate_Conc_combo.grid\
    (row=2, column=5)
Entry_Trinate_combo = Spinbox(Frame_combo,from_=0, to=31,wrap=True, width=10, font=("Helvetica", 12, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Trinate_combo_IntVar)
Entry_Trinate_combo.grid\
    (row=2, column=6)
Entry_Suffragi_Comunitari_combo = Spinbox(Frame_combo, from_=0, to=31,wrap=True, width=11,font=("Helvetica", 12, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Suffragi_Comunitari_combo_IntVar)
Entry_Suffragi_Comunitari_combo.grid\
    (row=2, column=7)
Entry_Suffragi_Personali_combo = Spinbox(Frame_combo, from_=0, to=31,wrap=True, width=10,font=("Helvetica", 12, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Suffragi_Personali_combo_IntVar)
Entry_Suffragi_Personali_combo.grid\
    (row=2, column=8)
Entry_Devozione_combo = Spinbox(Frame_combo,from_=0, to=31,wrap=True, width=11, font=("Helvetica", 12, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Devozione_combo_IntVar)
Entry_Devozione_combo.grid\
    (row=2, column=9)
Entry_Benefattori_combo = Spinbox(Frame_combo, from_=0, to=31,wrap=True, width=11,font=("Helvetica", 12, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Benefattori_combo_IntVar)
Entry_Benefattori_combo.grid\
    (row=2, column=10)
Entry_Pro_Populo_combo = Spinbox(Frame_combo,from_=0, to=31,wrap=True, width=11, font=("Helvetica", 12, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Pro_Populo_combo_IntVar)
Entry_Pro_Populo_combo.grid\
    (row=2, column=11)







############################
####### TREEVIEW ###########
############################



# Add some style
style = ttk.Style()
# Pick a theme
style.theme_use("default")

# Configure our treeview colors
style.configure("Treeview",
                background="#D3D3D3",
                foreground="black",
                rowheight=30,
                fieldbackground="#D3D3D3",
                font=('Calibri', 12)
                )

# Headings
style.configure("Treeview.Heading",
                font=('Calibri', 12, 'bold')
                )

# Change selected color
style.map('Treeview',
          background=[('selected', 'blue')]
          )

# Treeview Scrollbar
tree_scroll = Scrollbar(Frame_tree)
tree_scroll.pack(side=RIGHT, fill=Y)

# Create Treeview
my_tree = ttk.Treeview(Frame_tree, yscrollcommand=tree_scroll.set, selectmode="extended")
# Pack to the screen
my_tree.pack()

# Configure the scrollbar
tree_scroll.config(command=my_tree.yview)

# Define Our Columns
my_tree['columns'] = ("ID", "Anno", "Mese", "Nome_Celebrante", "Ad_Mentem", "Binate", "Binate_Concelebrate", "Trinate", "Suffragi_Comunitari", "Suffragi_Personali", "Devozione", "Benefattori", "Pro_Populo")





# Formate Our Columns
my_tree.column("#0", width=0, stretch=NO)
my_tree.column("ID", anchor=W, width=70)
my_tree.column("Anno", anchor=W, width=120)
my_tree.column("Mese", anchor=W, width=120)
my_tree.column("Nome_Celebrante", anchor=W, width=165)
my_tree.column("Ad_Mentem", anchor=W, width=120)
my_tree.column("Binate", anchor=W, width=120)
my_tree.column("Binate_Concelebrate", anchor=W, width=120)
my_tree.column("Trinate", anchor=W, width=120)
my_tree.column("Suffragi_Comunitari", anchor=W, width=120)
my_tree.column("Suffragi_Personali", anchor=W, width=120)
my_tree.column("Devozione", anchor=W, width=120)
my_tree.column("Benefattori", anchor=W, width=120)
my_tree.column("Pro_Populo", anchor=W, width=120)


# Create Headings
my_tree.heading("#0", text="", anchor=W)
my_tree.heading("ID", text="Id", anchor=W)
my_tree.heading("Anno", text="Anno", anchor=W)
my_tree.heading("Mese", text="Mese", anchor=W)
my_tree.heading("Nome_Celebrante", text="Celebrante", anchor=W)
my_tree.heading("Ad_Mentem", text="Ad_Ment", anchor=W)
my_tree.heading("Binate", text="Binate", anchor=W)
my_tree.heading("Binate_Concelebrate", text="Bin_Concel", anchor=W)
my_tree.heading("Trinate", text="Trinate", anchor=W)
my_tree.heading("Suffragi_Comunitari", text="Suffr_Com", anchor=W)
my_tree.heading("Suffragi_Personali", text="Suffr_Pers", anchor=W)
my_tree.heading("Devozione", text="Devozione", anchor=W)
my_tree.heading("Benefattori", text="Benefattori", anchor=W)
my_tree.heading("Pro_Populo", text="Pro_Populo", anchor=W)

def on_double_click(event):
    region_clicked = my_tree.identify_region(event.x, event.y)
    #print(region_clicked) >cell oppure >header

    # numero di colonna della riga su cui faccio doppio click
    column = my_tree.identify_column(event.x)
    #print(column) # esempio #4
    # numero colonna senza # davanti: numero intero
    #la prima colonna del treeview è = 1
    # sottraggo -1 perchè nelle touple primo valore è 0
    # [1:] inizia dal secondo carattere (dunque salta #)
    column_index = int(column[1:])-1                    #NUMERO COLONNA
    #print(column_index)

    #mi da l'ID della riga su cui faccio doppio click
    selected_iid=my_tree.focus()                        #ID RIGA
    #print(selected_iid) #esempio 16

    selected_values = my_tree.item(selected_iid) #('11', '2024', 'febbraio', 'fra Giacomo Rotunno', '0', '0', '0', '0', '0', '0', '0', '0', '0')

    if column_index == '0':
        selected_text = selected_values.get('text')
        #print(selected_text)
    else:
        selected_text = selected_values.get('values')[column_index]
        #print(selected_text)

    #posizione e dimensioni cella selezionata
    column_box = my_tree.bbox(selected_iid, column) #print(column)  esempio '#4'
    #print(column_box) #(1, 112, 70, 30) (x_position, y_position, Width, Height)

    entry_edit = ttk.Entry(Frame_tree, width=column_box[2])

    #salvo valori di colonna e riga in variabili
    #Record column index and item iid
    entry_edit.editing_column_index = column_index
    #print(entry_edit.editing_column_index) #     esempio 1
    entry_edit.editing_item_iid = selected_iid
    #print(entry_edit.editing_item_id)         #esempio 9


    entry_edit.place(x = column_box[0],
                     y = column_box[1],
                     w = column_box[2],
                     h = column_box[3])

    # mi inserisce il valore della casella selezionata
    entry_edit.insert(0, selected_text)
    #seleziona il testo
    entry_edit.select_range(0, END)
    #place the focus on the widget
    entry_edit.focus()

    #event.widget reference the entry widget
    def on_focus_out(event):
        event.widget.destroy()

    #ricorda che funziona solo se premi ENTER
    def on_enter_pressed(event):
        # event.widget reference the entry widget
        new_text = event.widget.get() #salva in new_text il testo modificato
        # print('°°°°°°')
        # print(new_text)
        #We also want to know the item ID
        selected_iid = event.widget.editing_item_iid #numero di riga
        # print('#############')
        # print(selected_iid)
        column_index = event.widget.editing_column_index #numero di colonna. la prima colonna è 0
        # print('#############')
        # print(column_index)
        if column_index == 0: # l'ID non si deve cambiare
             pass
        else:
            current_values = my_tree.item(selected_iid).get("values")
            #print(current_values) # [13, 2024, 'gennaio', 'Fra Alberto Dos Santos', 11, 0, 0, 0, 0, 0, 0, 0, 0]
            # {'text': '', 'image': '', 'values': [15, 2024, 'gennaio', 'Ospite', 28, 0, 0, 0, 0, 0, 0, 0, 0], 'open': 0, 'tags': ['oddrow']}

            #print(current_values)
            current_values[column_index]=new_text
            #print(current_values)
            my_tree.item(selected_iid, values=current_values)


        event.widget.destroy()

        # Update the database
        # Create a database or connect to one that exists
        conn = sqlite3.connect('database_messe_orizzontale')
        #
        # Create a cursor instance
        c = conn.cursor()
        print(conn)

        c.execute("""   UPDATE TABLE_Messe 
                            SET
                            Anno = :Anno,
                            Mese = :Mese,
                            Nome_Celebrante = :Nome_Celebrante,
                            Ad_Mentem = :Ad_Mentem,
                            Binate = :Binate,
                            Binate_Concelebrate = :Binate_Concelebrate,
                            Trinate = :Trinate,
                            Suffragi_Comunitari = :Suffragi_Comunitari,
                            Suffragi_Personali = :Suffragi_Personali,
                            Devozione = :Devozione,
                            Benefattori = :Benefattori,
                            Pro_Populo = :Pro_Populo


             		        WHERE oid =""" + selected_iid,
                  {
                      'Anno': current_values[1],
                      'Mese': current_values[2],
                      'Nome_Celebrante': current_values[3],
                      'Ad_Mentem': current_values[4],
                      'Binate': current_values[5],
                      'Binate_Concelebrate': current_values[6],
                      'Trinate': current_values[7],
                      'Suffragi_Comunitari': current_values[8],
                      'Suffragi_Personali': current_values[8],
                      'Devozione': current_values[10],
                      'Benefattori': current_values[11],
                      'Pro_Populo': current_values[12]
                  })

        #    Commit changes
        conn.commit()
        #
        #         # Close our connection
        conn.close()
        # Add a little message box for fun
        messagebox.showinfo("Updated!", "Riga aggiornata!")

    #when I click outside I want the widget to disappear
    entry_edit.bind("<FocusOut>", on_focus_out)

    #When I click enter UPDATE tree
    entry_edit.bind("<Return>", on_enter_pressed)




my_tree.bind("<Double-1>", on_double_click)


############################
######## SQLITE3 ###########
############################

# Insert into TABLE_Conti
def submit():
    conn = sqlite3.connect('database_messe_orizzontale')
    cur = conn.cursor()

    #dati presi dalla combo di inserimento (non update)
    dati = [(Entry_Anno_combo_IntVar.get(),
             Entry_Mese_combo_StringVar.get(),
             Entry_Nome_Celebrante_combo_StringVar.get(),
             Entry_Ad_Mentem_combo_IntVar.get(),
             Entry_Binate_combo_IntVar.get(),
             Entry_Binate_Conc_combo_IntVar.get(),
             Entry_Trinate_combo_IntVar.get(),
             Entry_Suffragi_Comunitari_combo_IntVar.get(),
             Entry_Suffragi_Personali_combo_IntVar.get(),
             Entry_Devozione_combo_IntVar.get(),
             Entry_Benefattori_combo_IntVar.get(),
             Entry_Pro_Populo_combo_IntVar.get()
             )]


    cur.executemany(
        'INSERT INTO TABLE_Messe (Anno, Mese, Nome_Celebrante, Ad_Mentem, Binate, Binate_Concelebrate, Trinate, Suffragi_Comunitari, Suffragi_Personali, Devozione, Benefattori, Pro_Populo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)', dati)



    conn.commit()
    # Close our connection
    conn.close()

def query_database():
    # Clear the Treeview
    for record in my_tree.get_children():
        my_tree.delete(record)

    # Create a database or connect to one that exists
    conn = sqlite3.connect('database_messe_orizzontale')

    # Create a cursor instance
    c = conn.cursor()

    sql_select_query = """select * from TABLE_Messe where Anno = ? and Mese = ? order by ID DESC """
    c.execute(sql_select_query, (2024, 'gennaio',))
    records = c.fetchall()
    # for x in records:
    #     print(x)
    # print('##################################################')

    # c.execute("SELECT * FROM TABLE_Messe ORDER BY Anno, (CASE Mese\
    #                                                             WHEN 'gennaio' THEN 1\
    #                                                             WHEN 'febbraio' THEN 2\
    #                                                             WHEN 'marzo' THEN 3\
    #                                                             WHEN 'aprile' THEN 4\
    #                                                             WHEN 'maggio' THEN 5\
    #                                                             WHEN 'giugno' THEN 6\
    #                                                             WHEN 'luglio' THEN 7\
    #                                                             WHEN 'agosto' THEN 8\
    #                                                             WHEN 'settembre' THEN 9\
    #                                                             WHEN 'ottobre' THEN 10\
    #                                                             WHEN 'novembre' THEN 11\
    #                                                             WHEN 'dicembre' THEN 12\
    #                                                             END);")
    c.execute("SELECT * FROM TABLE_Messe")
    records = c.fetchall()
    # for x in records:
    #     print(x)


    # for record in records:
    #     print(record)
    # record[0] = id key

    # COLORI RIGHE pari e dispari
    count = 0
    # Create striped row tags
    my_tree.tag_configure('oddrow', background="white")
    my_tree.tag_configure('evenrow', background="lightblue")

    for record in records:
        if count % 2 == 0:
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7], record[8], record[9], record[10], record[11], record[12]),
                           tags=('evenrow'))
        else:
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7],record[8], record[9], record[10], record[11], record[12]),
                           tags = ('oddrow'))
        count += 1

    # Al termine del processo la prima riga risulta evidenziata
    child_id = my_tree.get_children()[0]  # la prima riga dall'alto del treeview
    my_tree.focus(child_id)  # evidenziata
    my_tree.selection_set(child_id)

    # Commit changes
    conn.commit()

    # Close our connection
    conn.close()

def query_database_BY_DATE():
    # Clear the Treeview
    for record in my_tree.get_children():
        my_tree.delete(record)

    # Create a database or connect to one that exists
    conn = sqlite3.connect('database_messe_orizzontale')

    # Create a cursor instance
    c = conn.cursor()

    # sql_select_query = """select * from TABLE_Messe where Anno = ? and Mese = ? order by ID DESC """
    # c.execute(sql_select_query, (2024, 'gennaio',))
    # records = c.fetchall()
    # for x in records:
    #     print(x)
    # print('##################################################')

    c.execute("SELECT * FROM TABLE_Messe ORDER BY Anno, (CASE Mese\
                                                                WHEN 'gennaio' THEN 1\
                                                                WHEN 'febbraio' THEN 2\
                                                                WHEN 'marzo' THEN 3\
                                                                WHEN 'aprile' THEN 4\
                                                                WHEN 'maggio' THEN 5\
                                                                WHEN 'giugno' THEN 6\
                                                                WHEN 'luglio' THEN 7\
                                                                WHEN 'agosto' THEN 8\
                                                                WHEN 'settembre' THEN 9\
                                                                WHEN 'ottobre' THEN 10\
                                                                WHEN 'novembre' THEN 11\
                                                                WHEN 'dicembre' THEN 12\
                                                                END), Nome_Celebrante;")

    records = c.fetchall()
    # for x in records:
    #     print(x)


    # for record in records:
    #     print(record)
    # record[0] = id key

    # COLORI RIGHE pari e dispari
    #count = 0
    # Create striped row tags
    my_tree.tag_configure('white', background="pink")
    my_tree.tag_configure('blue', background="salmon")
    my_tree.tag_configure('yellow', background="lightyellow")
    my_tree.tag_configure('violet', background="khaki")

    for record in records:
        if record[2] == 'gennaio':
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7], record[8], record[9], record[10], record[11], record[12]),
                           tags=('white'))
        elif record[2] == 'febbraio':
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7],record[8], record[9], record[10], record[11], record[12]),
                           tags=('blue'))
        elif record[2] == 'marzo':
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(
                           record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7],
                           record[8], record[9], record[10], record[11], record[12]),
                           tags=('yellow'))
        elif record[2] == 'aprile':
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(
                           record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7],
                           record[8], record[9], record[10], record[11], record[12]),
                           tags=('violet'))
        elif record[2] == 'maggio':
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(
                               record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7],
                               record[8], record[9], record[10], record[11], record[12]),
                           tags=('white'))
        elif record[2] == 'giugno':
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(
                               record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7],
                               record[8], record[9], record[10], record[11], record[12]),
                           tags=('blue'))
        elif record[2] == 'luglio':
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(
                           record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7],
                           record[8], record[9], record[10], record[11], record[12]),
                           tags=('yellow'))
        elif record[2] == 'agosto':
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(
                           record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7],
                           record[8], record[9], record[10], record[11], record[12]),
                           tags=('violet'))
        elif record[2] == 'settembre':
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(
                               record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7],
                               record[8], record[9], record[10], record[11], record[12]),
                           tags=('white'))
        elif record[2] == 'ottobre':
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(
                               record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7],
                               record[8], record[9], record[10], record[11], record[12]),
                           tags=('blue'))

        elif record[2] == 'novembre':
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(
                               record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7],
                               record[8], record[9], record[10], record[11], record[12]),
                           tags=('yellow'))
        else:
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7],record[8], record[9], record[10], record[11], record[12]),
                           tags = ('violet'))
        #count += 1

    # Al termine del processo la prima riga risulta evidenziata
    child_id = my_tree.get_children()[0]  # la prima riga dall'alto del treeview
    my_tree.focus(child_id)  # evidenziata
    my_tree.selection_set(child_id)

    # Commit changes
    conn.commit()

    # Close our connection
    conn.close()



#######################
def remove_one():
    # my_tree.focus() restituisce l'ID della riga selezionata
    row_id = my_tree.focus()
    my_tree.delete(row_id)

    # Create a database or connect to one that exists
    conn = sqlite3.connect('database_messe_orizzontale')

    # Create a cursor instance
    c = conn.cursor()

    # Delete From Database
    c.execute("DELETE from TABLE_Messe WHERE oid =" + row_id)

    # Commit changes
    conn.commit()

    # Close our connection
    conn.close()

    # Add a little message box for fun
    messagebox.showinfo("Deleted", "Riga Cancellata!")

##############################################################
##########TOP WINDOW CELEBRANTI###############################
##############################################################

def Top_W_Celebranti():
    top = Toplevel()
    top.geometry("380x500")
    top.title("Celebranti")

    Frame_top_tree = Frame(top, bd='4', bg=background_Blu, relief=RIDGE)
    Frame_top_tree.pack()


    Label_Celebranti = Label(top, text='Celebranti:', font=('verdana', 8, 'bold'), bg=background_Blu, fg=foreground_Bianco)
    Label_Celebranti.pack()

    Entry_Celebranti_StringVar = StringVar()

    Entry_Celebranti = Entry(top, bd=5, textvariable=Entry_Celebranti_StringVar)
    Entry_Celebranti.pack()

    

    ############################
    ####### TREEVIEW ###########
    ############################

    # Add some style
    style = ttk.Style()
    # Pick a theme
    style.theme_use("default")

    # Configure our treeview colors
    style.configure("Treeview",
                    background="#D3D3D3",
                    foreground="black",
                    rowheight=30,
                    fieldbackground="#D3D3D3",
                    font=('Calibri', 12)
                    )

    # Headings
    style.configure("Treeview.Heading",
                    font=('Calibri', 12, 'bold')
                    )

    # Change selected color
    style.map('Treeview',
              background=[('selected', 'blue')]
              )

    # Treeview Scrollbar
    tree_scroll = Scrollbar(Frame_top_tree)
    tree_scroll.pack(side=RIGHT, fill=Y)

    # Create Treeview
    my_tree = ttk.Treeview(Frame_top_tree, yscrollcommand=tree_scroll.set, selectmode="extended")
    # Pack to the screen
    my_tree.pack()

    # Configure the scrollbar
    tree_scroll.config(command=my_tree.yview)

    # Define Our Columns
    my_tree['columns'] = ("ID", "Celebranti")

    # Formate Our Columns
    my_tree.column("#0", width=0, stretch=NO)
    my_tree.column("ID", anchor=W, width=70)
    my_tree.column("Celebranti", anchor=W, width=200)


    # Create Headings
    my_tree.heading("#0", text="", anchor=W)
    my_tree.heading("ID", text="Id", anchor=W)
    my_tree.heading("Celebranti", text="Celebranti", anchor=W)



    def query_database():
        # Clear the Treeview
        for record in my_tree.get_children():
            my_tree.delete(record)

        # Create a database or connect to one that exists
        conn = sqlite3.connect('database_messe_orizzontale')

        # Create a cursor instance
        c = conn.cursor()

        c.execute("SELECT * FROM TABLE_Celebranti;")
        records = c.fetchall()
        # for x in records:
        #     print('Table Messe')
        #     print(x)

        # COLORI RIGHE pari e dispari
        count = 0
        # Create striped row tags
        my_tree.tag_configure('oddrow', background="white")
        my_tree.tag_configure('evenrow', background="lightblue")

        for record in records:
            if count % 2 == 0:
                my_tree.insert(parent='', index=0, iid=record[0], text='',
                               values=(
                               record[0],
                               record[1]),
                               tags=('evenrow'))
            else:
                my_tree.insert(parent='', index=0, iid=record[0], text='',
                               values=(
                               record[0],
                               record[1]),
                               tags=('oddrow'))
            count += 1

        # Al termine del processo la prima riga risulta evidenziata
        child_id = my_tree.get_children()[0]  # la prima riga dall'alto del treeview
        my_tree.focus(child_id)  # evidenziata
        my_tree.selection_set(child_id)

        # Commit changes
        conn.commit()

        # Close our connection
        conn.close()

    def remove_one():
        # my_tree.focus() restituisce l'ID della riga selezionata
        row_id = my_tree.focus()
        my_tree.delete(row_id)

        # Create a database or connect to one that exists
        conn = sqlite3.connect('database_messe_orizzontale')

        # Create a cursor instance
        c = conn.cursor()

        # Delete From Database
        c.execute("DELETE from TABLE_Celebranti WHERE oid =" + row_id)

        # Commit changes
        conn.commit()

        # Close our connection
        conn.close()


    def submit():
        conn = sqlite3.connect('database_messe_orizzontale')
        cur = conn.cursor()
        dati = [Entry_Celebranti_StringVar.get()]
        cur.execute('INSERT INTO TABLE_Celebranti (Celebranti) VALUES (?)', dati)
        #cur.execute("insert into TABLE_Celebranti (Celebranti) values ('?'), dati")
        #cur.execute('''INSERT INTO TABLE_Celebranti (Celebranti) VALUES ("Entry_Celebranti_StringVar.get()")''')
        conn.commit()
        # Close our connection
        conn.close()

    B_add_celebranti = Button(top, text='aggiungi', width=10, command=lambda: [submit(), query_database()]).pack(side=TOP, pady=20)
    B_delete_celebranti = Button(top, text='cancella', width=10, command=remove_one).pack(side=TOP, pady=20)

    query_database()
    top.mainloop()


##############################################################
############  WINDOW Suffragi  ###############################
##############################################################

def Suffragi_Comunitari():


        Label_Titolo_Suffragi= Label(Frame_Suffragi, text='Inserire i suffragi comunitari:', font=('verdana', 12, 'bold'), bg=background_Blu, fg=foreground_Bianco)
        Label_Titolo_Suffragi.grid(column=0,row=0, columnspan=2, sticky="W")

        # List Anni
        Anni = [2020, 2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030]

        Label_Anno_Suffragi = Label(Frame_Suffragi, text='Anno:', font=('verdana', 12, 'bold'), bg=background_Blu, fg=foreground_Bianco)
        Label_Anno_Suffragi.grid(row=1, column=0, sticky="W",pady=5)

        # Dropbox Anno
        Entry_Anno_Suffragi_StringVar = StringVar()
        Entry_Anno_Suffragi = ttk.Combobox(Frame_Suffragi, font=("Helvetica", 10), values=Anni,
                                        textvariable=Entry_Anno_Suffragi_StringVar)
        #Entry_Anno_Suffragi.current(4)
        Entry_Anno_Suffragi.grid(row=1, column=1, sticky="W",pady=5)
        #Entry_Anno_combo['state'] = 'readonly'

        # List Mesi
        Mesi = ["gennaio",
                "febbraio",
                "marzo",
                "aprile",
                "maggio",
                "giugno",
                "luglio",
                "agosto",
                "settembre",
                "ottobre",
                "novembre",
                "dicembre",
                ]

        Label_Mese_Suffragi= Label(Frame_Suffragi, text='Mese:', font=('verdana', 12, 'bold'), bg=background_Blu, fg=foreground_Bianco)
        Label_Mese_Suffragi.grid(column=0,row=2, sticky="W", pady=5)

        # Dropbox Mesi
        Entry_Mese_Suffragi_StringVar = StringVar()
        Entry_Mese_Suffragi = ttk.Combobox(Frame_Suffragi, font=("Helvetica", 10), values=Mesi, textvariable=Entry_Mese_Suffragi_StringVar)


        #Entry_Mese_Suffragi.current(0)
        Entry_Mese_Suffragi.grid(row=2, column=1,  sticky="W", pady=5)


        Label_Suffragi = Label(Frame_Suffragi, text='Suffragi:', font=('verdana', 12, 'bold'), bg=background_Blu,
                                 fg=foreground_Bianco)
        Label_Suffragi.grid(column=0,row=3, sticky="W", pady=5)
        Entry_Suffragi_StringVar = StringVar()
        Entry_Suffragi = Entry(Frame_Suffragi, bd=5, textvariable=Entry_Suffragi_StringVar)
        Entry_Suffragi.grid(column=1, row=3, sticky="W", pady=5)

        Frame_suffragi_tree = Frame(Frame_Suffragi, bd='4', bg=background_Blu, relief=RIDGE)
        Frame_suffragi_tree.grid(column=0, row=4, columnspan=4, pady=12, padx=2)


        ############################
        ####### TREEVIEW ###########
        ############################

        # Add some style
        style = ttk.Style()
        # Pick a theme
        style.theme_use("default")

        # Configure our treeview colors
        style.configure("Treeview",
                        background="#D3D3D3",
                        foreground="black",
                        rowheight=30,
                        fieldbackground="#D3D3D3",
                        font=('Calibri', 12)
                        )

        # Headings
        style.configure("Treeview.Heading",
                        font=('Calibri', 12, 'bold')
                        )

        # Change selected color
        style.map('Treeview',
                  background=[('selected', 'blue')]
                  )

        # Treeview Scrollbar
        tree_scroll = Scrollbar(Frame_suffragi_tree)
        tree_scroll.pack(side=RIGHT, fill=Y)

        # Create Treeview
        my_tree = ttk.Treeview(Frame_suffragi_tree, yscrollcommand=tree_scroll.set, selectmode="extended", height=9)
        # Pack to the screen
        my_tree.pack()

        # Configure the scrollbar
        tree_scroll.config(command=my_tree.yview)

        # Define Our Columns
        my_tree['columns'] = ("ID", "Anno", "Mese", "Suffragi")

        # Formate Our Columns
        my_tree.column("#0", width=0, stretch=NO)
        my_tree.column("ID", anchor=W, width=40)
        my_tree.column("Anno", anchor=W, width=70)
        my_tree.column("Mese", anchor=W, width=70)
        my_tree.column("Suffragi", anchor=W, width=250)

        # Create Headings
        my_tree.heading("#0", text="", anchor=W)
        my_tree.heading("ID", text="Id", anchor=W)
        my_tree.heading("Anno", text="Anno", anchor=W)
        my_tree.heading("Mese", text="Mese", anchor=W)
        my_tree.heading("Suffragi", text="Suffragi", anchor=W)

        def query_suffragi_database():
            # Clear the Treeview
            for record in my_tree.get_children():
                my_tree.delete(record)

            # Create a database or connect to one that exists
            conn = sqlite3.connect('database_messe_orizzontale')

            # Create a cursor instance
            c = conn.cursor()

            c.execute("SELECT * FROM TABLE_Suffragi;")
            records = c.fetchall()
            for x in records:
                print('TABLE_Suffragi')
                print(x)
        #
            # COLORI RIGHE pari e dispari
            count = 0
            # Create striped row tags
            my_tree.tag_configure('oddrow', background="white")
            my_tree.tag_configure('evenrow', background="lightblue")

            for record in records:
                if count % 2 == 0:
                    my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=  (
                                       record[0],
                                       record[1],
                                       record[2],
                                       record[3]
                                            ),
                                   tags=('evenrow'))
                else:
                    my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=  (
                                       record[0],
                                       record[1],
                                       record[2],
                                       record[3]
                                            ),
                                   tags=('oddrow'))
                count += 1

            # Al termine del processo la prima riga risulta evidenziata
            child_id = my_tree.get_children()[0]  # la prima riga dall'alto del treeview
            my_tree.focus(child_id)  # evidenziata
            my_tree.selection_set(child_id)

            # Commit changes
            conn.commit()

            # Close our connection
            conn.close()

        def query_Suffragi_database_BY_DATE():
            # Clear the Treeview
            for record in my_tree.get_children():
                my_tree.delete(record)

            # Create a database or connect to one that exists
            conn = sqlite3.connect('database_messe_orizzontale')

            # Create a cursor instance
            c = conn.cursor()

            # sql_select_query = """select * from TABLE_Messe where Anno = ? and Mese = ? order by ID DESC """
            # c.execute(sql_select_query, (2024, 'gennaio',))
            # records = c.fetchall()
            # for x in records:
            #     print(x)
            # print('##################################################')

            c.execute("SELECT * FROM TABLE_Suffragi ORDER BY Anno, (CASE Mese\
                                                                        WHEN 'gennaio' THEN 1\
                                                                        WHEN 'febbraio' THEN 2\
                                                                        WHEN 'marzo' THEN 3\
                                                                        WHEN 'aprile' THEN 4\
                                                                        WHEN 'maggio' THEN 5\
                                                                        WHEN 'giugno' THEN 6\
                                                                        WHEN 'luglio' THEN 7\
                                                                        WHEN 'agosto' THEN 8\
                                                                        WHEN 'settembre' THEN 9\
                                                                        WHEN 'ottobre' THEN 10\
                                                                        WHEN 'novembre' THEN 11\
                                                                        WHEN 'dicembre' THEN 12\
                                                                        END), Suffragi;")

            records = c.fetchall()
            # for x in records:
            #     print(x)

            # for record in records:
            #     print(record)
            # record[0] = id key

            # COLORI RIGHE pari e dispari
            # count = 0
            # Create striped row tags
            my_tree.tag_configure('white', background="pink")
            my_tree.tag_configure('blue', background="salmon")
            my_tree.tag_configure('yellow', background="lightyellow")
            my_tree.tag_configure('violet', background="khaki")

            for record in records:
                if record[2] == 'gennaio':
                    my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=(record[0], record[1], record[2], record[3]),
                                   tags=('white'))
                elif record[2] == 'febbraio':
                    my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=(record[0], record[1], record[2], record[3]),
                                   tags=('blue'))
                elif record[2] == 'marzo':
                    my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=(
                                       record[0], record[1], record[2], record[3]),
                                   tags=('yellow'))
                elif record[2] == 'aprile':
                    my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=(
                                       record[0], record[1], record[2], record[3]),
                                   tags=('violet'))
                elif record[2] == 'maggio':
                    my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=(
                                       record[0], record[1], record[2], record[3]),
                                   tags=('white'))
                elif record[2] == 'giugno':
                    my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=(
                                       record[0], record[1], record[2], record[3]),
                                   tags=('blue'))
                elif record[2] == 'luglio':
                    my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=(
                                       record[0], record[1], record[2], record[3]),
                                   tags=('yellow'))
                elif record[2] == 'agosto':
                    my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=(
                                       record[0], record[1], record[2], record[3]),
                                   tags=('violet'))
                elif record[2] == 'settembre':
                    my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=(
                                       record[0], record[1], record[2], record[3]),
                                   tags=('white'))
                elif record[2] == 'ottobre':
                    my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=(
                                       record[0], record[1], record[2], record[3]),
                                   tags=('blue'))

                elif record[2] == 'novembre':
                    my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=(
                                       record[0], record[1], record[2], record[3]),
                                   tags=('yellow'))
                else:
                    my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=(record[0], record[1], record[2], record[3]),
                                   tags=('violet'))
                # count += 1

            # Al termine del processo la prima riga risulta evidenziata
            child_id = my_tree.get_children()[0]  # la prima riga dall'alto del treeview
            my_tree.focus(child_id)  # evidenziata
            my_tree.selection_set(child_id)

            # Commit changes
            conn.commit()

            # Close our connection
            conn.close()

        def remove_one():
            # my_tree.focus() restituisce l'ID della riga selezionata
            row_id = my_tree.focus()
            my_tree.delete(row_id)

            # Create a database or connect to one that exists
            conn = sqlite3.connect('database_messe_orizzontale')

            # Create a cursor instance
            c = conn.cursor()

            # Delete From Database
            c.execute("DELETE from TABLE_Suffragi WHERE oid =" + row_id)

            # Commit changes
            conn.commit()

            # Close our connection
            conn.close()
    #
        def submit():
            conn = sqlite3.connect('database_messe_orizzontale')
            cur = conn.cursor()
            dati = [Entry_Anno_Suffragi.get(), Entry_Mese_Suffragi.get(), Entry_Suffragi_StringVar.get()]
            dati_suffragi = [Entry_Suffragi_StringVar.get()]
            cur.execute('INSERT INTO TABLE_Suffragi (Anno, Mese, Suffragi) VALUES (?,?,?)', dati)

            # cur.execute("insert into TABLE_Celebranti (Celebranti) values ('?'), dati")
            # cur.execute('''INSERT INTO TABLE_Celebranti (Celebranti) VALUES ("Entry_Celebranti_StringVar.get()")''')
            conn.commit()
            # Close our connection
            conn.close()


        B_add_Suffragi = Button(Frame_Suffragi, text='aggiungi', width=10, command=lambda: [submit(), query_suffragi_database()]).grid(row=0,column=3, sticky=W, pady=5)
        B_delete_Suffragi = Button(Frame_Suffragi, text='cancella', width=10, command=remove_one).grid(row=1,column=3, sticky=W)
        B_Tree_Suffragi_sort_by_ID = Button(Frame_Suffragi, text='Sort ID', width=10, command=query_suffragi_database).grid(row=2,column=3, sticky=W)
        B_Tree_Suffragi_sort_by_Date = Button(Frame_Suffragi, text='Sort Data', width=10, command=query_Suffragi_database_BY_DATE).grid(row=3,column=3, sticky=W)
        query_suffragi_database()



B_add = Button(Frame_tree_Buttons, text='aggiungi', width=10, command=lambda: [submit(), query_database()]).pack(side=TOP, pady=20)
#B_excel = Button(Frame_tree_Buttons, text='Filtro_excel', width=10, command=sqlite3_to_excel).pack(side=TOP, pady=20)
#B_update = Button(Frame_tree_Buttons, text='aggiorna', width=10, command='').pack(side=TOP, pady=20)
B_delete = Button(Frame_tree_Buttons, text='cancella', width=10, command=remove_one).pack(side=TOP, pady=20)
B_Nomi_Celebranti = Button(Frame_tree_Buttons, text='Celebranti', width=10, command=Top_W_Celebranti).pack(side=BOTTOM, pady=20)
B_Tree_sort_by_ID=Button(Frame_tree_Buttons, text='Sort ID', width=10, command=query_database).pack(side=TOP, pady=20)
B_Tree_sort_by_Date=Button(Frame_tree_Buttons, text='Sort Data', width=10, command=lambda: [query_database_BY_DATE()]).pack(side=TOP, pady=20)

Suffragi_Comunitari()
query_database()
root.mainloop()