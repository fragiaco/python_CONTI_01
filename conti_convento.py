from tkinter         import *
from tkinter         import ttk
from Origine         import *
import sqlite3

conn = sqlite3.connect('database_conti')


cur = conn.cursor()
try:
    cur.execute('''CREATE TABLE TABLE_Conti(ID integer not null PRIMARY KEY ,
                                            Anno TEXT not null ,
                                            Mese TEXT not null ,
                                            Entrate_Uscite TEXT not null ,
                                            Categoria TEXT not null ,
                                            Voce TEXT not null ,
                                            Euro real not null )''')
except:
    pass


print(conn)
print('Sei connesso al database_conti')


conn.commit()

root = Tk()

# root.title("Programma gestione conti Convento Offida")
# root.geometry('1680x950+0+0')

height = 950 # altezza
width = 1680 # larghezza
# top = (root.winfo_screenheight() - height) / 2
top = 0
left = (root.winfo_screenwidth() - width) / 2
geometry = ("{}x{}+{}+{}".format(width, height, int(left), int(top)))
root.geometry(geometry)
root.resizable(0, 0)
root.title('')

#Label title
title = Label(root, text='Database Entrate Uscite', font=('verdana', 40, 'bold'), bg='blue', fg='white')
title.pack(side=TOP, fill=X)
###########################
# Frame 1 - left side Frame
Frame1 = Frame(root, bd='4', bg='blue', relief=RIDGE)
Frame1.place(x=20, y=85, width=550, height=850)
# Frame 1 - bottom side Frame
Frame1in = Frame(Frame1, bd='4', bg='blue', relief=RIDGE)
Frame1in.place(x=15, y=768, width=500, height=60)
# Frame 2 - right side Frame
Frame2 = Frame(root, bd='4', bg='blue', relief=RIDGE)
Frame2.place(x=590, y=85, width=1070, height=850)
# Frame 2in - treeview right Frame
Frame2in_tree = Frame(Frame2, bd='4', bg='blue', relief=RIDGE)
Frame2in_tree.place(x=15, y=15, width=1015, height=335)
# Frame 2in - bottom side right Frame
Frame2in_bottom = Frame(Frame2, bd='4', bg='blue', relief=RIDGE)
Frame2in_bottom.place(x=15, y=360, width=1015, height=470)



# define mylabel
mylabel=Label(Frame2)


def submit():

    conn = sqlite3.connect('database_conti')
    cur = conn.cursor()

    dati = [(anno_combo.get(), mesi_combo.get(), my_combo.get(), categoria_combo.get(), voce_combo.get(), euro.get())]

    cur.executemany('INSERT INTO TABLE_Conti (Anno, Mese, Entrate_Uscite, Categoria, Voce, Euro) VALUES(?, ?, ?, ? ,? ,?)', dati)
    conn.commit()





##########################
#Frame1 Labels
Frame1_title = Label(Frame1, text='Inserisci Dati:', font=('verdana', 20, 'bold'), bg='blue', fg='white')
Frame1_title.grid(row=0, columnspan=2, padx=20, pady=10, sticky='w')

Frame1_anno = Label(Frame1, text='Anno', font=('verdana', 15, 'bold'), bg='blue', fg='white')
Frame1_anno.grid(row=2, padx=20, pady=10, sticky='w')

Frame1_mese = Label(Frame1, text='Mese', font=('verdana', 15, 'bold'), bg='blue', fg='white')
Frame1_mese.grid(row=4, padx=20, pady=25, sticky='w')

Frame1_EntrateUscite = Label(Frame1, text='Entrate/Uscite', font=('verdana', 15, 'bold'), bg='blue', fg='white')
Frame1_EntrateUscite.grid(row=6, padx=25, pady=20, sticky='w')

Frame1_Categoria = Label(Frame1, text='Categoria', font=('verdana', 15, 'bold'), bg='blue', fg='white')
Frame1_Categoria.grid(row=8, padx=25, pady=20, sticky='w')

Frame1_VoceSpesa = Label(Frame1, text='Voce di spesa', font=('verdana', 15, 'bold'), bg='blue', fg='white')
Frame1_VoceSpesa.grid(row=10, padx=25, pady=20, sticky='w')

Frame1_Euro = Label(Frame1, text='Euro', font=('verdana', 15, 'bold'), bg='blue', fg='white')
Frame1_Euro.grid(row=12, padx=25, pady=20, sticky='w')
#######################

def pick_Categoria(e):
    if my_combo.get() == "Entrate":
        categoria_combo.config(values=Categorie_Entrate)
        categoria_combo.current(0)

    if my_combo.get() == "Uscite":
        categoria_combo.config(values=Categorie_Uscite)
        categoria_combo.current(0)


def pick_Voce(e):
    # VOCI ENTRATA
    if categoria_combo.get() == "Collette_Chiesa":
        voce_combo.config(values=Collette_Chiesa)
        voce_combo.current(0)

    if categoria_combo.get() == "Congrua":
        voce_combo.config(values=Congrua)
        voce_combo.current(0)

    if categoria_combo.get() == "Interessi":
        voce_combo.config(values=Interessi)
        voce_combo.current(0)

    if categoria_combo.get() == "Messe celebrate":
        voce_combo.config(values=Messe_celebrate)
        voce_combo.current(0)

    if categoria_combo.get() == "Offerte":
        voce_combo.config(values=Offerte)
        voce_combo.current(0)

    if categoria_combo.get() == "Pensioni":
        voce_combo.config(values=Pensioni)
        voce_combo.current(0)

    if categoria_combo.get() == "Servizi_religiosi":
        voce_combo.config(values=Servizi_religiosi)
        voce_combo.current(0)

    # VOCI USCITA

    if categoria_combo.get() == "Acquisti_Chiesa":
        voce_combo.config(values=Acquisti_Chiesa)
        voce_combo.current(0)

    if categoria_combo.get() == "Acquisti_Convento":
        voce_combo.config(values=Acquisti_Convento)
        voce_combo.current(0)

    if categoria_combo.get() == "Acquisti_Orto_Animali":
        voce_combo.config(values=Acquisti_Orto_Animali)
        voce_combo.current(0)

    if categoria_combo.get() == "Cultura":
        voce_combo.config(values=Cultura)
        voce_combo.current(0)

    if categoria_combo.get() == "Curia_provinciale":
        voce_combo.config(values=Curia_provinciale)
        voce_combo.current(0)

    if categoria_combo.get() == "Domestici":
        voce_combo.config(values=Domestici)
        voce_combo.current(0)

    if categoria_combo.get() == "Elargizioni":
        voce_combo.config(values=Elargizioni)
        voce_combo.current(0)

    if categoria_combo.get() == "Utenze":
        voce_combo.config(values=Utenze)
        voce_combo.current(0)

    if categoria_combo.get() == "Ferie_Viaggi":
        voce_combo.config(values=Ferie_Viaggi)
        voce_combo.current(0)

    if categoria_combo.get() == "Igiene":
        voce_combo.config(values=Igiene)
        voce_combo.current(0)

    if categoria_combo.get() == "Imposte":
        voce_combo.config(values=Imposte)
        voce_combo.current(0)

    if categoria_combo.get() == "Lavori_Impianti":
        voce_combo.config(values=Lavori_Impianti)
        voce_combo.current(0)

    if categoria_combo.get() == "Posta_Cancelleria":
        voce_combo.config(values=Posta_Cancelleria)
        voce_combo.current(0)

    if categoria_combo.get() == "Salute":
        voce_combo.config(values=Salute)
        voce_combo.current(0)

    if categoria_combo.get() == "Veicoli_motore":
        voce_combo.config(values=Veicoli_motore)
        voce_combo.current(0)

    if categoria_combo.get() == "Vestiario":
        voce_combo.config(values=Vestiario)
        voce_combo.current(0)

    if categoria_combo.get() == "Vitto":
        voce_combo.config(values=Vitto)
        voce_combo.current(0)

    if categoria_combo.get() == "Eccedenza_Cassa":
        voce_combo.config(values=Eccedenza_Cassa)
        voce_combo.current(0)


# Dropbox Anno
anno_combo = ttk.Combobox(Frame1, font=("Helvetica", 15), values=Anni)
anno_combo.current(0)
anno_combo.grid(row=2, column=1)
# Dropbox Mesi
mesi_combo = ttk.Combobox(Frame1, font=("Helvetica", 15), values=Mesi)
mesi_combo.current(0)
mesi_combo.grid(row=4, column=1)
# Dropbox Entrate_Uscite
my_combo = ttk.Combobox(Frame1, font=("Helvetica", 15), values=Entrate_Uscite)
#my_combo.current(0)
my_combo.grid(row=6, column=1)

# Bind the ComboBox
my_combo.bind("<<ComboboxSelected>>", pick_Categoria)

# Categoria ComboBox
categoria_combo = ttk.Combobox(Frame1, font=("Helvetica", 15), values=[""])
categoria_combo.current(0)
categoria_combo.grid(row=8, column=1)

# Bind the ComboBox
categoria_combo.bind("<<ComboboxSelected>>", pick_Voce)

# Voce Entrata_Spesa ComboBox Combo Box
voce_combo = ttk.Combobox(Frame1, font=("Helvetica", 15), values=[""])
voce_combo.current(0)
voce_combo.grid(row=10, column=1)

# euro ENTRY
euro = Entry(Frame1, font=("Helvetica", 15, 'bold'), bd=5, relief=GROOVE)
euro.grid(row=12, column=1)



# Add some style
style = ttk.Style()
#Pick a theme
style.theme_use("default")
# Configure our treeview colors

style.configure("Treeview",
	background="#D3D3D3",
	foreground="black",
	rowheight=30,
	fieldbackground="#D3D3D3",
    font=('Calibri', 12)
	)

#Headings
style.configure("Treeview.Heading", font=('Calibri', 12,'bold'))

# Change selected color
style.map('Treeview',
	background=[('selected', 'blue')])

# Create Treeview Frame
# tree_frame = Frame(Frame2)
#Frame2in_tree.pack(pady=10)

# Treeview Scrollbar
tree_scroll = Scrollbar(Frame2in_tree)
tree_scroll.pack(side=RIGHT, fill=Y)

# Create Treeview
my_tree = ttk.Treeview(Frame2in_tree, yscrollcommand=tree_scroll.set, selectmode="extended")
# Pack to the screen
my_tree.pack()

#Configure the scrollbar
tree_scroll.config(command=my_tree.yview)

# Define Our Columns
my_tree['columns'] = ("Id", "Anno", "Mese", "Entrate_Uscite", "Categoria", "Voce", "Euro")

# Formate Our Columns
my_tree.column("#0", width=0, stretch=NO)
my_tree.column("Id", anchor=CENTER, width=50)
my_tree.column("Anno", anchor=CENTER, width=80)
my_tree.column("Mese", anchor=CENTER, width=80)
my_tree.column("Entrate_Uscite", anchor=CENTER, width=160)
my_tree.column("Categoria", anchor=W, width=200)
my_tree.column("Voce", anchor=W, width=250)
my_tree.column("Euro", anchor=W, width=160)


# Create Headings
my_tree.heading("#0", text="", anchor=W)
my_tree.heading("Id", text="Id", anchor=CENTER)
my_tree.heading("Anno", text="Anno", anchor=CENTER)
my_tree.heading("Mese", text="Mese", anchor=CENTER)
my_tree.heading("Entrate_Uscite", text="Entrate_Uscite", anchor=CENTER)
my_tree.heading("Categoria", text="Categoria", anchor=W)
my_tree.heading("Voce", text="Voce", anchor=W)
my_tree.heading("Euro", text="Euro", anchor=W)


def query_database():

    # Clear the Treeview
    for record in my_tree.get_children():
        my_tree.delete(record)

    # Create a database or connect to one that exists
    conn = sqlite3.connect('database_conti')

    # Create a cursor instance
    c = conn.cursor()


    c.execute("SELECT * FROM TABLE_Conti")
    records = c.fetchall()



    for record in records:
        print(record)


    for record in records:
        if record[0] % 2 == 0:
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(record[0], record[1], record[2], record[3], record[4], record[5], record[6]),
                           tags=('evenrow'))
        else:
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(record[0], record[1], record[2], record[3], record[4], record[5], record[6]),
                           tags=('oddrow'))

    child_id = my_tree.get_children()[0]  # for instance the last element in tuple
    my_tree.focus(child_id)
    my_tree.selection_set(child_id)

    # Commit changes
    conn.commit()

    # Close our connection
    conn.close()

#######################################



# B_add = Button(Frame1in, text='add', width=10, command=lambda:[submit()]).grid(row=0, column=0, padx=20, pady=15)

B_add = Button(Frame1in, text='add', width=10, command=lambda:[submit(), query_database()]).grid(row=0, column=0, padx=20, pady=15)
B_update = Button(Frame1in, text='update', width=10).grid(row=0, column=1, padx=20, pady=15)
B_delete = Button(Frame1in, text='delete', width=10).grid(row=0, column=2, padx=20, pady=15)
B_clear = Button(Frame1in, text='clear', width=10).grid(row=0, column=3, padx=20, pady=15)

################treeviw


# Create striped row tags
my_tree.tag_configure('oddrow', background="white")
my_tree.tag_configure('evenrow', background="lightblue")
##############################

#  Record  Boxes
# data_frame = LabelFrame(root, text="Record")
# data_frame.pack(fill="x", expand="yes", padx=20)
# Frame 2in - bottom side right Frame
# Frame2in_bottom = Frame(Frame2, bd='4', bg='blue', relief=RIDGE)
# Frame2in_bottom.place(x=15, y=360, width=1015, height=470)

# Frame1_title = Label(Frame1, text='Inserisci Dati:', font=('verdana', 20, 'bold'), bg='blue', fg='white')
# Frame1_title.grid(row=0, columnspan=2, padx=20, pady=10, sticky='w')

Id_label = Label(Frame2in_bottom, text="Id", font=('verdana', 15, 'bold'), bg='blue', fg='white')
Id_label.grid(row=0, column=0, padx=10, pady=10, sticky='w')
Mese_entry = Entry(Frame2in_bottom)
Mese_entry.grid(row=0, column=1, padx=10, pady=10)

Anno_label = Label(Frame2in_bottom, text="Anno", font=('verdana', 15, 'bold'), bg='blue', fg='white')
Anno_label.grid(row=1, column=0, padx=10, pady=10, sticky='w')
Anno_entry = Entry(Frame2in_bottom)
Anno_entry.grid(row=1, column=1, padx=10, pady=10)

Mese_label = Label(Frame2in_bottom, text="Mese", font=('verdana', 15, 'bold'), bg='blue', fg='white')
Mese_label.grid(row=2, column=0, padx=10, pady=10, sticky='w')
Mese_entry = Entry(Frame2in_bottom)
Mese_entry.grid(row=2, column=1, padx=10, pady=10)

Entrate_Uscite_label = Label(Frame2in_bottom, text="Entrate_Uscite", font=('verdana', 15, 'bold'), bg='blue', fg='white')
Entrate_Uscite_label.grid(row=3, column=0, padx=10, pady=10, sticky='w')
Entrate_Uscite_entry = Entry(Frame2in_bottom)
Entrate_Uscite_entry.grid(row=3, column=1, padx=10, pady=10)
#
Categoria_label = Label(Frame2in_bottom, text="Categoria", font=('verdana', 15, 'bold'), bg='blue', fg='white')
Categoria_label.grid(row=4, column=0, padx=10, pady=10, sticky='w')
Categorie_Entrate_entry = Entry(Frame2in_bottom)
Categorie_Entrate_entry.grid(row=4, column=1, padx=10, pady=10)
#
Voce_label = Label(Frame2in_bottom, text="Voce", font=('verdana', 15, 'bold'), bg='blue', fg='white')
Voce_label.grid(row=5, column=0, padx=10, pady=10, sticky='w')
Voce_label_entry = Entry(Frame2in_bottom)
Voce_label_entry.grid(row=5, column=1, padx=10, pady=10)
#
Euro_label = Label(Frame2in_bottom, text="Euro", font=('verdana', 15, 'bold'), bg='blue', fg='white')
Euro_label.grid(row=6, column=0, padx=10, pady=10, sticky='w')
Euro_label_entry = Entry(Frame2in_bottom)
Euro_label_entry.grid(row=6, column=1, padx=10, pady=10)




# Bind the treeview
# my_tree.bind("<ButtonRelease-1>", select_record)





query_database()
conn.close()
root.mainloop()