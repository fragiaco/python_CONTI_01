##############################################################
##########TOP WINDOW CELEBRANTI###############################
##############################################################
from tkinter import *
from tkinter import ttk
import sqlite3



def Top_W_Celebranti():
    top = Toplevel()
    top.geometry("380x500")
    top.title("Celebranti")

    Frame_top_tree = Frame(top, bd='4', bg='blue', relief=RIDGE)
    Frame_top_tree.pack()

    Label_Celebranti = Label(top, text='Celebranti:', font=('verdana', 8, 'bold'), bg='blue',
                             fg='white')
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
        # cur.execute("insert into TABLE_Celebranti (Celebranti) values ('?'), dati")
        # cur.execute('''INSERT INTO TABLE_Celebranti (Celebranti) VALUES ("Entry_Celebranti_StringVar.get()")''')
        conn.commit()
        # Close our connection
        conn.close()

    B_add_celebranti = Button(top, text='aggiungi', width=10, command=lambda: [submit(), query_database()]).pack(
        side=TOP, pady=20)
    B_delete_celebranti = Button(top, text='cancella', width=10, command=remove_one).pack(side=TOP, pady=20)

    query_database()
    top.mainloop()

Top_W_Celebranti()