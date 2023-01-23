from tkinter import *
from tkinter import ttk
from tkinter import messagebox

import sqlite3
import pandas as pd
import os, sys, subprocess

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
                                                Binate integer not null ,
                                                Binate_Concelebrate integer not null ,
                                                Trinate integer not null ,
                                                Suffragi_Comunitari integer not null ,
                                                Suffragi_Personali integer not null ,
                                                Devozione integer not null ,
                                                Benefattori integer not null ,
                                                Pro_Populo integer not null ,
                                                Numero integer not null )''')



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

# Frame Combo - left side Frame
Frame_combo = Frame(root, bd='4', bg=background_Blu, relief=RIDGE)
Frame_combo.place(x=50, y=75, width=1600, height=250)
# Frame IN Combo - left side Frame
# Frame_button_in_combo = Frame(Frame_combo, bd='4', bg=background_Blu, relief=RIDGE)
# Frame_button_in_combo.place(y=400, width=445, height=42)


# Frame Treeview - treeview right Frame
Frame_tree = Frame(root, bd='4', bg=background_Blu, relief=RIDGE)
Frame_tree.place(x=50, y=375, width=1600, height=200)
# Frame IN Treeview - left side Frame
# Frame_button_in_Treeview = Frame(Frame_tree, bd='4', bg=background_Blu, relief=RIDGE)
#Frame_button_in_Treeview.place(y=400, width=704, height=42)

# Frame Update - update right Frame
Frame_update = Frame(root, bd='4', bg=background_Blu, relief=RIDGE)
Frame_update.place(x=50, y=675, width=1600, height=200)


###############################################
######## TKINTER LABELS Frame_combo ###########
###############################################
# Frame IN Update - left side Frame
# Frame_button_in_update = Frame(Frame_update, bd='4', bg=background_Blu, relief=RIDGE)
# Frame_button_in_update.place(y=400, width=445, height=42)
#
# # Frame_bottom_left
# Frame_bottom_left = Frame(root, bd='4', bg=background_Blu, relief=RIDGE) #azzurro fiordaliso
# Frame_bottom_left.place(x=20, y=528, width=450, height=415)
#
# # Frame Tabella - Tabella bottom Frame
# Frame_tabella = Frame(root, bd='4', bg=background_Blu, relief=RIDGE)
# Frame_tabella.place(x=475, y=528, width=730, height=415)
#
# # Frame_bottom_right
# Frame_bottom_right = Frame(root, bd='4', bg=background_Blu, relief=RIDGE)
# Frame_bottom_right.place(x=1210, y=528, width=450, height=415)


#Labels in Frame Combo_insert
Label_combo_title = Label(Frame_combo, text='Inserisci Dati:', font=('verdana', 15, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_combo_title.grid(row=0, columnspan=2, padx=10, pady=10, sticky='w')

# Label_riga_vuota = Label(Frame_combo, text='', font=('verdana', 5, 'bold'), bg=background_Blu, fg=foreground_Bianco)
# Label_riga_vuota.grid(row=1)

Label_combo_anno = Label(Frame_combo, text='Anno', font=('verdana', 10, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_combo_anno.grid(row=1, column=0, columnspan=1, padx=10, pady=10, sticky='w')

Label_combo_mese = Label(Frame_combo, text='Mese', font=('verdana', 10, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_combo_mese.grid               (row=1, column=1, columnspan=1, padx=10, pady=10, sticky='w')

Label_combo_Nome_celebrante = Label(Frame_combo, text='Celebrante', font=('verdana', 10, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_combo_Nome_celebrante.grid    (row=1, column=2, padx=10, pady=10, sticky='w')

Label_combo_Ad_Mentem = Label(Frame_combo, text='Ad_Mentem', font=('verdana', 10, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_combo_Ad_Mentem.grid          (row=1, column=3, padx=10, pady=10, sticky='w')

Label_combo_Binate = Label(Frame_combo, text='Binate', font=('verdana', 10, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_combo_Binate.grid             (row=1, column=4, padx=10, pady=10, sticky='w')

Label_combo_Binate_Concelebrate = Label(Frame_combo, text='Binate_Conc.', font=('verdana', 10, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_combo_Binate_Concelebrate.grid(row=1, column=5, padx=10, pady=20, sticky='w')

Label_combo_Trinate = Label(Frame_combo, text='Trinate', font=('verdana', 10, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_combo_Trinate.grid            (row=1, column=6, padx=10, pady=20, sticky='w')

Label_combo_Suffragi_Comunitari = Label(Frame_combo, text='Suffr_Comuntà', font=('verdana', 10, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_combo_Suffragi_Comunitari.grid(row=1, column=7, padx=10, pady=20, sticky='w')

Label_combo_Suffragi_Personali = Label(Frame_combo, text='Suffr_Pers.', font=('verdana', 10, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_combo_Suffragi_Personali.grid(row=1, column=8, padx=10, pady=20, sticky='w')

Label_combo_Devozione = Label(Frame_combo, text='Devozione', font=('verdana', 10, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_combo_Devozione.grid          (row=1, column=9, padx=9, pady=20, sticky='w')

Label_combo_Benefattori = Label(Frame_combo, text='Benefattori', font=('verdana', 10, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_combo_Benefattori.grid        (row=1, column=10, padx=10, pady=20, sticky='w')

Label_combo_Pro_Populo = Label(Frame_combo, text='Pro_Populo', font=('verdana', 10, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_combo_Pro_Populo.grid         (row=1, column=11, padx=10, pady=20, sticky='w')

#LABEL TOTALE MESSE
Label_Riga_Vuota = Label(Frame_combo, text='', font=('verdana', 5, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_Riga_Vuota.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky='w')
Label_Totale_Messe_Celebrate = Label(Frame_combo, text='Totale Messe celebrate', font=('verdana', 13, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_Totale_Messe_Celebrate.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky='w')
Label_STRING_Totale_Messe_Celebrate = Label(Frame_combo, text='Totale Messe celebrate', font=('verdana', 13, 'bold'), bg=background_Blu, fg=foreground_Bianco)
Label_STRING_Totale_Messe_Celebrate.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky='w')






###############################################
######## TKINTER ENTRY Frame_combo ############
###############################################
Entry_Anno_combo_IntVar         = IntVar()
Entry_Mese_combo_StringVar      = StringVar()
Entry_Nome_Celebrante_combo_StringVar = StringVar()
Entry_Ad_Mentem_combo_IntVar    = IntVar()
Entry_Binate_combo_IntVar       = IntVar()
Entry_Binate_Conc_combo_IntVar  = IntVar()
Entry_Trinate_combo_IntVar      = IntVar()
Entry_Suffragi_Comunitari_combo_IntVar = IntVar()
Entry_Suffragi_Personali_combo_IntVar = IntVar()
Entry_Devozione_combo_IntVar    = IntVar()
Entry_Benefattori_combo_IntVar  = IntVar()
Entry_Pro_Populo_combo_IntVar   = IntVar()

# Dropbox Anno
Entry_Anno_combo = Entry(Frame_combo, font=("Helvetica", 8, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Anno_combo_IntVar)
Entry_Anno_combo.grid\
    (row=2, columnspan=1, column=0)

# Entry_Mese_combo = Entry(Frame_combo, font=("Helvetica", 8, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Mese_combo_StringVar)
# Entry_Mese_combo.grid\
#     (row=2, columnspan=1, column=1)

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
Entry_Mese_combo = ttk.Combobox(Frame_combo, font=("Helvetica", 12), values=Mesi, textvariable=Entry_Mese_combo_StringVar)
Entry_Mese_combo.current(0)
Entry_Mese_combo.grid(row=2, columnspan=1, column=1)
Entry_Mese_combo['state'] = 'readonly'

Entry_Nome_Celebrante_combo = Entry(Frame_combo, font=("Helvetica", 8, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Nome_Celebrante_combo_StringVar)
Entry_Nome_Celebrante_combo.grid\
    (row=2, column=2)
Entry_Ad_Mentem_combo = Entry(Frame_combo, font=("Helvetica", 8, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Ad_Mentem_combo_IntVar)
Entry_Ad_Mentem_combo.grid\
    (row=2, column=3)
Entry_Binate_combo = Entry(Frame_combo, font=("Helvetica", 8, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Binate_combo_IntVar)
Entry_Binate_combo.grid\
    (row=2, column=4)
Entry_Binate_Conc_combo = Entry(Frame_combo, font=("Helvetica", 8, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Binate_Conc_combo_IntVar)
Entry_Binate_Conc_combo.grid\
    (row=2, column=5)
Entry_Trinate_combo = Entry(Frame_combo, font=("Helvetica", 8, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Trinate_combo_IntVar)
Entry_Trinate_combo.grid\
    (row=2, column=6)
Entry_Suffragi_Comunitari_combo = Entry(Frame_combo, font=("Helvetica", 8, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Suffragi_Comunitari_combo_IntVar)
Entry_Suffragi_Comunitari_combo.grid\
    (row=2, column=7)
Entry_Suffragi_Personali_combo = Entry(Frame_combo, font=("Helvetica", 8, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Suffragi_Personali_combo_IntVar)
Entry_Suffragi_Personali_combo.grid\
    (row=2, column=8)
Entry_Devozione_combo = Entry(Frame_combo, font=("Helvetica", 8, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Devozione_combo_IntVar)
Entry_Devozione_combo.grid\
    (row=2, column=9)
Entry_Benefattori_combo = Entry(Frame_combo, font=("Helvetica", 8, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Benefattori_combo_IntVar)
Entry_Benefattori_combo.grid\
    (row=2, column=10)
Entry_Pro_Populo_combo = Entry(Frame_combo, font=("Helvetica", 8, 'bold'), bd=5, relief=GROOVE, textvariable=Entry_Pro_Populo_combo_IntVar)
Entry_Pro_Populo_combo.grid\
    (row=2, column=11)




#######################################################
# ID integer not null PRIMARY KEY ,
#                                                 Anno integer not null ,
#                                                 Mese TEXT not null ,
#                                                 Nome_Celebrante TEXT not null ,
#                                                 Ad_Mentem integer not null ,
#                                                 Binate integer not null ,
#                                                 Binate_Concelebrate integer not null ,
#                                                 Trinate integer not null ,
#                                                 Suffragi_Comunitari integer not null ,
#                                                 Suffragi_Personali integer not null ,
#                                                 Devozione integer not null ,
#                                                 Benefattori integer not null ,
#                                                 Pro_Populo integer not null ,
#                                                 Numero integer


root.mainloop()
