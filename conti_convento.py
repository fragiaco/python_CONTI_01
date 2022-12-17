from tkinter         import *
from tkinter         import ttk
from tkinter         import messagebox
#import calculator
from Origine         import *
import sqlite3
import pandas as pd
import os, sys, subprocess


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
#FRame calcolatrice
Frame_calc = Frame(Frame1, bd='4', bg='light blue', relief=RIDGE)
Frame_calc.grid(column=0, row=15, columnspan=2, padx=65, pady=70)


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
Frame2in_bottom.place(x=15, y=360, width=1015, height=480)

#Frame update botton
Frame_update_botton = Frame(Frame2in_bottom, bd='4', bg='blue', relief=RIDGE)
Frame_update_botton.place(x=15, y=405, width=970, height=60)



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

# Python program to create a simple GUI
# calculator using Tkinter

# import everything from tkinter module
#from tkinter import *

# globally declare the expression variable
expression = ""


# Function to update expression
# in the text entry box
def press(num):
	# point out the global expression variable
	global expression

	# concatenation of string
	expression = expression + str(num)

	# update the expression by using set method
	equation.set(expression)


# Function to evaluate the final expression
def equalpress():
	# Try and except statement is used
	# for handling the errors like zero
	# division error etc.

	# Put that code inside the try block
	# which may generate the error
	try:

		global expression

		# eval function evaluate the expression
		# and str function convert the result
		# into string
		total = str(eval(expression))

		equation.set(total)

		# initialize the expression variable
		# by empty string
		expression = ""

	# if error is generate then handle
	# by the except block
	except:

		equation.set(" error ")
		expression = ""


# Function to clear the contents
# of text entry box
def clear():
	global expression
	expression = ""
	equation.set("")


# Driver code
# if __name__ == "__main__":
	# create a GUI window
	# gui = Tk()
    #
	# # set the background colour of GUI window
	# gui.configure(background="light green")
    #
	# # set the title of GUI window
	# gui.title("Simple Calculator")
    #
	# # set the configuration of GUI window
	# gui.geometry("270x150")

	# StringVar() is the variable class
	# we create an instance of this class
equation = StringVar()

	# create the text entry box for
	# showing the expression .
expression_field = Entry(Frame_calc, textvariable=equation)

	# grid method is used for placing
	# the widgets at respective positions
	# in table like structure .
# expression_field.grid(columnspan=4, ipadx=70)

	# create a Buttons and place at a particular
	# location inside the root window .
	# when user press the button, the command or
	# function affiliated to that button is executed .
button1 = Button(Frame_calc, text=' 1 ', fg='black', bg='light grey',
				command=lambda: press(1), height=1, width=7)
button1.grid(row=2, column=0, pady=2)

button2 = Button(Frame_calc, text=' 2 ', fg='black', bg='light grey',
				command=lambda: press(2), height=1, width=7)
button2.grid(row=2, column=1)

button3 = Button(Frame_calc, text=' 3 ', fg='black', bg='light grey',
				command=lambda: press(3), height=1, width=7)
button3.grid(row=2, column=2)

button4 = Button(Frame_calc, text=' 4 ', fg='black', bg='light grey',
				command=lambda: press(4), height=1, width=7)
button4.grid(row=3, column=0, pady=2)

button5 = Button(Frame_calc, text=' 5 ', fg='black', bg='light grey',
				command=lambda: press(5), height=1, width=7)
button5.grid(row=3, column=1)

button6 = Button(Frame_calc, text=' 6 ', fg='black', bg='light grey',
				command=lambda: press(6), height=1, width=7)
button6.grid(row=3, column=2)

button7 = Button(Frame_calc, text=' 7 ', fg='black', bg='light grey',
				command=lambda: press(7), height=1, width=7)
button7.grid(row=4, column=0,pady=2)

button8 = Button(Frame_calc, text=' 8 ', fg='black', bg='light grey',
				command=lambda: press(8), height=1, width=7)
button8.grid(row=4, column=1)

button9 = Button(Frame_calc, text=' 9 ', fg='black', bg='light grey',
				command=lambda: press(9), height=1, width=7)
button9.grid(row=4, column=2)

button0 = Button(Frame_calc, text=' 0 ', fg='black', bg='light grey',
				command=lambda: press(0), height=1, width=7)
button0.grid(row=5, column=0, pady=2)

plus = Button(Frame_calc, text=' + ', fg='black', bg='light grey',
				command=lambda: press("+"), height=1, width=7)
plus.grid(row=2, column=3)

minus = Button(Frame_calc, text=' - ', fg='black', bg='light grey',
				command=lambda: press("-"), height=1, width=7)
minus.grid(row=3, column=3)

multiply = Button(Frame_calc, text=' * ', fg='black', bg='light grey',
				command=lambda: press("*"), height=1, width=7)
multiply.grid(row=4, column=3)

divide = Button(Frame_calc, text=' / ', fg='black', bg='light grey',
				command=lambda: press("/"), height=1, width=7)
divide.grid(row=5, column=3)

equal = Button(Frame_calc, text=' = ', fg='black', bg='light grey',
			command=equalpress, height=1, width=7)
equal.grid(row=5, column=2)

clear = Button(Frame_calc, text='Clear', fg='black', bg='light grey',
			command=clear, height=1, width=7)
clear.grid(row=5, column=1)

Decimal= Button(Frame_calc, text='.', fg='black', bg='light grey',
				command=lambda: press('.'), height=1, width=7)
Decimal.grid(row=6, column=0)
	# start the GUI
	#gui.mainloop()

#########################

def pick_Categoria(e):
    if my_combo.get() == "Entrate":
        categoria_combo.config(values=Categorie_Entrate)
        categoria_combo.set('click!')
        categoria_combo['state'] = 'readonly'

    if my_combo.get() == "Uscite":
        categoria_combo.config(values=Categorie_Uscite)
        categoria_combo.set('click!')
        categoria_combo['state'] = 'readonly'


def pick_Voce(e):
    # VOCI ENTRATA
    if categoria_combo.get() == "Collette_Chiesa":
        voce_combo.config(values=Collette_Chiesa)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Congrua":
        voce_combo.config(values=Congrua)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Interessi":
        voce_combo.config(values=Interessi)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Messe celebrate":
        voce_combo.config(values=Messe_celebrate)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Offerte":
        voce_combo.config(values=Offerte)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Pensioni":
        voce_combo.config(values=Pensioni)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Servizi_religiosi":
        voce_combo.config(values=Servizi_religiosi)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    # VOCI USCITA

    if categoria_combo.get() == "Acquisti_Chiesa":
        voce_combo.config(values=Acquisti_Chiesa)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Acquisti_Convento":
        voce_combo.config(values=Acquisti_Convento)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Acquisti_Orto_Animali":
        voce_combo.config(values=Acquisti_Orto_Animali)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Cultura":
        voce_combo.config(values=Cultura)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Curia_provinciale":
        voce_combo.config(values=Curia_provinciale)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Domestici":
        voce_combo.config(values=Domestici)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Elargizioni":
        voce_combo.config(values=Elargizioni)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Utenze":
        voce_combo.config(values=Utenze)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Ferie_Viaggi":
        voce_combo.config(values=Ferie_Viaggi)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Igiene":
        voce_combo.config(values=Igiene)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Imposte":
        voce_combo.config(values=Imposte)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Lavori_Impianti":
        voce_combo.config(values=Lavori_Impianti)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Posta_Cancelleria":
        voce_combo.config(values=Posta_Cancelleria)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Salute":
        voce_combo.config(values=Salute)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Veicoli_motore":
        voce_combo.config(values=Veicoli_motore)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Vestiario":
        voce_combo.config(values=Vestiario)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Vitto":
        voce_combo.config(values=Vitto)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'


    if categoria_combo.get() == "Eccedenza_Cassa":
        voce_combo.config(values=Eccedenza_Cassa)
        voce_combo.set('click!')
        voce_combo['state'] = 'readonly'



# Dropbox Anno
anno_combo = ttk.Combobox(Frame1, font=("Helvetica", 15), values=Anni)
anno_combo.current(0)
anno_combo.grid(row=2, column=1)
anno_combo['state'] = 'readonly'

# Dropbox Mesi
mesi_combo = ttk.Combobox(Frame1, font=("Helvetica", 15), values=Mesi)
mesi_combo.current(0)
mesi_combo.grid(row=4, column=1)
mesi_combo['state'] = 'readonly'

# Dropbox Entrate_Uscite
my_combo = ttk.Combobox(Frame1, font=("Helvetica", 15), values=Entrate_Uscite)
my_combo.set('click!')
my_combo.grid(row=6, column=1)
my_combo['state'] = 'readonly'

# Bind the ComboBox
my_combo.bind("<<ComboboxSelected>>", pick_Categoria)

# Categoria ComboBox
categoria_combo = ttk.Combobox(Frame1, font=("Helvetica", 15), values=[""])
categoria_combo.current(0)
categoria_combo.grid(row=8, column=1)
categoria_combo['state'] = 'readonly'

# Bind the ComboBox
categoria_combo.bind("<<ComboboxSelected>>", pick_Voce)

# Voce Entrata_Spesa ComboBox Combo Box
voce_combo = ttk.Combobox(Frame1, font=("Helvetica", 15), values=[""])
voce_combo.current(0)
voce_combo.grid(row=10, column=1)
voce_combo['state'] = 'readonly'

# euro ENTRY
# euro = Entry(Frame1, font=("Helvetica", 15, 'bold'), bd=5, relief=GROOVE)
euro = Entry(Frame1, font=("Helvetica", 15, 'bold'), bd=5, relief=GROOVE, textvariable=equation)
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


    count=0

    for record in records:
        print(record)

    #record[0] = id key

    for record in records:
        if count % 2 == 0:
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(record[0], record[1], record[2], record[3], record[4], record[5], record[6]),
                           tags=('evenrow'))
        else:
            my_tree.insert(parent='', index=0, iid=record[0], text='',
                           values=(record[0], record[1], record[2], record[3], record[4], record[5], record[6]),
                           tags=('oddrow'))

        count +=1
    child_id = my_tree.get_children()[0]  # la prima riga dall'alto del treeview
    my_tree.focus(child_id) #evidenziata
    #print(my_tree.focus(child_id)) #stampa l'anno
    my_tree.selection_set(child_id)#oppure questo
#   ID Anno Mese  Entrate_Uscite  Categoria  Voce  Euro
    sql = '''
    SELECT * FROM TABLE_Conti;
    '''

    # df = pd.read_sql_query(sql,conn)
    # print(df)
    # print(df.groupby(['Voce']).count())

    # Commit changes
    conn.commit()


    # Close our connection
    conn.close()

#######################################
def sqlite3_to_excel():

    # Create a database or connect to one that exists
    conn = sqlite3.connect('database_conti')

    # Create a cursor instance
    c = conn.cursor()


    query="SELECT * FROM TABLE_Conti" # query to collect recors
    df = pd.read_sql(query, conn, index_col='ID') # create dataframe
    df.to_excel('database_conti.xlsx') # create excel file

    if sys.platform == "win32":
        os.startfile('database_conti.xlsx')
    else:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, 'database_conti.xlsx'])

    # Commit changes
    conn.commit()

    # Close our connection
    conn.close()
################treeviw


# Create striped row tags
my_tree.tag_configure('oddrow', background="white")
my_tree.tag_configure('evenrow', background="lightblue")


##############################
anno_stringvar = StringVar()
mese_stringvar = StringVar()
entrate_uscite_stringvar = StringVar()
categoria_stringvar = StringVar()
voce_stringvar = StringVar()
euro_stringvar = StringVar()



Frame2_bottom_title = Label(Frame2in_bottom, text='Selezionare sopra la riga da correggere', font=('verdana', 20, 'bold'), bg='blue', fg='white')
Frame2_bottom_title.grid(row=0, columnspan=2, padx=20, pady=10, sticky='w')

Id_label = Label(Frame2in_bottom, text="Id", font=('verdana', 15, 'bold'), bg='blue', fg='white')
Id_label.grid(row=1, column=0, padx=10, pady=10, sticky='w')
Id_entry = Entry(Frame2in_bottom, font=('verdana', 15, 'bold'), bg='blue', fg='white', width=17)
Id_entry.grid(row=1, column=1)

Anno_label = Label(Frame2in_bottom, text="Anno", font=('verdana', 15, 'bold'), bg='blue', fg='white')
Anno_label.grid(row=2, column=0, padx=10, pady=10, sticky='w')
Anno_entry = Entry(Frame2in_bottom, font=('verdana', 15, 'bold'), bg='blue', fg='white', textvariable=anno_stringvar)
# Anno_entry.grid(row=2, column=1, padx=10, pady=10)

Mese_label = Label(Frame2in_bottom, text="Mese", font=('verdana', 15, 'bold'), bg='blue', fg='white')
Mese_label.grid(row=3, column=0, padx=10, pady=10, sticky='w')
Mese_entry = Entry(Frame2in_bottom, font=('verdana', 15, 'bold'), bg='blue', fg='white', textvariable=mese_stringvar)
#Mese_entry.grid(row=3, column=1, padx=10, pady=10)

Entrate_Uscite_label = Label(Frame2in_bottom, text="Entrate_Uscite", font=('verdana', 15, 'bold'), bg='blue', fg='white')
Entrate_Uscite_label.grid(row=4, column=0, padx=10, pady=10, sticky='w')
Entrate_Uscite_entry = Entry(Frame2in_bottom, font=('verdana', 15, 'bold'), bg='blue', fg='white', textvariable=entrate_uscite_stringvar)
#Entrate_Uscite_entry.grid(row=4, column=1, padx=10, pady=10)
#
Categoria_label = Label(Frame2in_bottom, text="Categoria", font=('verdana', 15, 'bold'), bg='blue', fg='white')
Categoria_label.grid(row=5, column=0, padx=10, pady=10, sticky='w')
Categorie_Entrate_entry = Entry(Frame2in_bottom, font=('verdana', 15, 'bold'), bg='blue', fg='white', textvariable=categoria_stringvar)
#Categorie_Entrate_entry.grid(row=5, column=1, padx=10, pady=10)
#
Voce_label = Label(Frame2in_bottom, text="Voce", font=('verdana', 15, 'bold'), bg='blue', fg='white')
Voce_label.grid(row=6, column=0, padx=10, pady=10, sticky='w')
Voce_entry = Entry(Frame2in_bottom, font=('verdana', 15, 'bold'), bg='blue', fg='white', textvariable=voce_stringvar)
#Voce_entry.grid(row=6, column=1, padx=10, pady=10)
#
Euro_label = Label(Frame2in_bottom, text="Euro", font=('verdana', 15, 'bold'), bg='blue', fg='white')
Euro_label.grid(row=7, column=0, padx=10, pady=10, sticky='w')
Euro_entry = Entry(Frame2in_bottom, font=('verdana', 15, 'bold'), bg='blue', fg='white', textvariable=euro_stringvar)
#Euro_entry.grid(row=7, column=1, padx=10, pady=10)



def select_record(e):
# Clear entry boxes
    Id_entry.delete(0, END)
    Anno_entry.delete(0, END)
    Mese_entry.delete(0, END)
    Entrate_Uscite_entry.delete(0, END)
    Categorie_Entrate_entry.delete(0, END)
    Voce_entry.delete(0, END)
    Euro_entry.delete(0, END)

# Grab record Number
    selected = my_tree.focus() #focus restituisce l'ID key
    # print(selected) #esempio 38
# Grab record values
    values = my_tree.item(selected, 'values')
    # print(values) #esempio ('38', '2022', 'gennaio', 'Entrate', 'Messe_celebrate', '', '39.0')

# outpus to entry boxes
    Id_entry.insert(0, values[0]) #0 penso significa all'inizio
    Anno_entry.insert(0, values[1])
    Mese_entry.insert(0, values[2])
    Entrate_Uscite_entry.insert(0, values[3])
    Categorie_Entrate_entry.insert(0, values[4])
    Voce_entry.insert(0, values[5])
    Euro_entry.insert(0, values[6])

    #print(Anno_entry.get())



# Bind the treeview
my_tree.bind("<ButtonRelease-1>", select_record)

#######################
def remove_one():
    #x = my_tree.selection()[0] #restituisce l'Id key
    x = my_tree.focus()
    my_tree.delete(x)

	# Create a database or connect to one that exists
    conn = sqlite3.connect('database_conti')

	# Create a cursor instance
    c = conn.cursor()

	# Delete From Database
    c.execute("DELETE from TABLE_Conti WHERE oid=" + Id_entry.get())

    # Commit changes
    conn.commit()

    # Close our connection
    conn.close()


# Add a little message box for fun
    messagebox.showinfo("Deleted!", "Your Record Has Been Deleted!")
#######################
######
    # Update record
def update_record():
    # Grab the record number
    print('update')
    selected = my_tree.focus()
    print(selected)
    # Update record
    my_tree.item(selected, text="", values=(Id_entry.get(), Anno_entry.get(), Mese_entry.get(), Entrate_Uscite_entry.get(), Categorie_Entrate_entry.get(), Voce_entry.get(), Euro_entry.get()))


    # Update the database
    # Create a database or connect to one that exists
    conn = sqlite3.connect('database_conti')
#
     # Create a cursor instance
    c = conn.cursor()

#
    c.execute("""UPDATE TABLE_Conti SET
    		Anno = :Anno,
    		Mese = :Mese,
    		Entrate_Uscite = :Entrate_Uscite,
    		Categoria = :Categoria,
    		Voce = :Voce,
    		Euro = :Euro

     		WHERE oid = :oid""",
           {
                'Anno': Anno_entry.get(),
                'Mese': Mese_entry.get(),
                'Entrate_Uscite': Entrate_Uscite_entry.get(),
                'Categoria': Categorie_Entrate_entry.get(),
                'Voce': Voce_entry.get(),
                'Euro': Euro_entry.get(),
                'oid': Id_entry.get()
           })
#
#    Commit changes
    conn.commit()
#
#         # Close our connection
    conn.close()
# Add a little message box for fun
    messagebox.showinfo("Deleted!", "Your Record Has Been Updated!")



#         # Clear entry boxes
    Id_entry.delete(0, END)
    Anno_entry.delete(0, END)
    Mese_entry.delete(0, END)
    Entrate_Uscite_entry.delete(0, END)
    Categorie_Entrate_entry.delete(0, END)
    Voce_entry.delete(0, END)
    Euro_entry.delete(0, END)

######
def pick_Categoria_update(e):
    if my_combo_update.get() == "Entrate":
        categoria_combo_update.config(values=Categorie_Entrate)
        #categoria_combo_update.set(entrate_uscite_stringvar.get())
        categoria_combo_update.current(0)

    if my_combo_update.get() == "Uscite":
        categoria_combo_update.config(values=Categorie_Uscite)
        categoria_combo_update.current(0)


def pick_Voce_update(e):
    # VOCI ENTRATA
    if categoria_combo_update.get() == "Collette_Chiesa":
        voce_combo_update.config(values=Collette_Chiesa)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Congrua":
        voce_combo_update.config(values=Congrua)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Interessi":
        voce_combo_update.config(values=Interessi)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Messe celebrate":
        voce_combo_update.config(values=Messe_celebrate)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Offerte":
        voce_combo_update.config(values=Offerte)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Pensioni":
        voce_combo_update.config(values=Pensioni)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Servizi_religiosi":
        voce_combo_update.config(values=Servizi_religiosi)
        voce_combo_update.current(0)

    # VOCI USCITA

    if categoria_combo_update.get() == "Acquisti_Chiesa":
        voce_combo_update.config(values=Acquisti_Chiesa)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Acquisti_Convento":
        voce_combo_update.config(values=Acquisti_Convento)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Acquisti_Orto_Animali":
        voce_combo_update.config(values=Acquisti_Orto_Animali)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Cultura":
        voce_combo_update.config(values=Cultura)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Curia_provinciale":
        voce_combo_update.config(values=Curia_provinciale)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Domestici":
        voce_combo_update.config(values=Domestici)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Elargizioni":
        voce_combo_update.config(values=Elargizioni)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Utenze":
        voce_combo_update.config(values=Utenze)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Ferie_Viaggi":
        voce_combo_update.config(values=Ferie_Viaggi)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == " ":
        voce_combo_update.config(values=Igiene)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Imposte":
        voce_combo_update.config(values=Imposte)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Lavori_Impianti":
        voce_combo_update.config(values=Lavori_Impianti)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Posta_Cancelleria":
        voce_combo_update.config(values=Posta_Cancelleria)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Salute":
        voce_combo_update.config(values=Salute)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Veicoli_motore":
        voce_combo_update.config(values=Veicoli_motore)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Vestiario":
        voce_combo_update.config(values=Vestiario)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Vitto":
        voce_combo_update.config(values=Vitto)
        voce_combo_update.current(0)

    if categoria_combo_update.get() == "Eccedenza_Cassa":
        voce_combo_update.config(values=Eccedenza_Cassa)
        voce_combo_update.current(0)

# Dropbox Anno
anno_combo_update = ttk.Combobox(Frame2in_bottom, font=("Helvetica", 15), values=Anni, textvariable=anno_stringvar)
anno_combo_update.set(anno_stringvar.get())
# anno_combo_update = ttk.Combobox(Frame2in_bottom, font=("Helvetica", 15), values=Anni, textvariable=new_Anno_entry)
# anno_combo_update.set(new_Anno_entry)
anno_combo_update.grid(row=2, column=1)
# Dropbox Mesi
mesi_combo_update = ttk.Combobox(Frame2in_bottom, font=("Helvetica", 15), values=Mesi, textvariable=mese_stringvar)
mesi_combo_update.set(mese_stringvar.get())
mesi_combo_update.grid(row=3, column=1)
# Dropbox Entrate_Uscite
my_combo_update = ttk.Combobox(Frame2in_bottom, font=("Helvetica", 15), values=Entrate_Uscite, textvariable=entrate_uscite_stringvar)
#my_combo.current(0)
my_combo_update.grid(row=4, column=1)

# Bind the ComboBox
my_combo_update.bind("<<ComboboxSelected>>", pick_Categoria_update)

# Categoria ComboBox
categoria_combo_update = ttk.Combobox(Frame2in_bottom, font=("Helvetica", 15), values=[""], textvariable=categoria_stringvar)
categoria_combo_update.current(0)
categoria_combo_update.grid(row=5, column=1)

# Bind the ComboBox
categoria_combo_update.bind("<<ComboboxSelected>>", pick_Voce_update)

# Voce Entrata_Spesa ComboBox Combo Box
voce_combo_update = ttk.Combobox(Frame2in_bottom, font=("Helvetica", 15), values=[""], textvariable=voce_stringvar)
voce_combo_update.current(0)
voce_combo_update.grid(row=6, column=1)

# euro ENTRY
euro_update = Entry(Frame2in_bottom, font=("Helvetica", 15, 'bold'), bd=5, relief=GROOVE, textvariable=euro_stringvar)
euro_update.grid(row=7, column=1)

####


# B_add = Button(Frame1in, text='add', width=10, command=lambda:[submit()]).grid(row=0, column=0, padx=20, pady=15)

B_add = Button(Frame1in, text='add', width=10, command=lambda:[submit(), query_database()]).grid(row=0, column=0, padx=20, pady=15)
B_update = Button(Frame_update_botton, text='update', width=10, command=update_record).grid(row=0, column=1, padx=10, pady=15)
B_delete = Button(Frame1in, text='delete', width=10, command=remove_one).grid(row=0, column=2, padx=20, pady=15)
B_excel = Button(Frame1in, text='excel', width=10, command=sqlite3_to_excel).grid(row=0, column=1, padx=20, pady=15)

#####


query_database()

conn.close()

root.mainloop()

