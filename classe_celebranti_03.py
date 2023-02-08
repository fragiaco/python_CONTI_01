##############################################################
##########TOP WINDOW CELEBRANTI###############################
##############################################################

import tkinter as tk
from tkinter import ttk
import sqlite3



class SettingsWindows:
    def __init__(self):
        self.win = tk.Toplevel() # simile a root = tk.Tk()
        self.win.geometry("380x500")
        self.win.title("Celebranti")

        self.frame = tk.Frame(self.win)
        self.frame.pack(padx=5, pady=5)

        self.frame_tree = tk.Frame(self.win)
        self.frame_tree.pack(padx=5, pady=5 )

        self.label = tk.Label(self.frame, text='Celebranti:', font=('verdana', 8, 'bold'), fg='black')
        self.label.pack(padx=5, pady=5)

        self.Entry_Celebranti_StringVar = tk.StringVar()
        self.Entry_Celebranti = tk.Entry(self.frame, bd=5, textvariable=self.Entry_Celebranti_StringVar)
        self.Entry_Celebranti.pack()

        self.label = tk.Label(self.frame_tree, text='Celebranti:', font=('verdana', 8, 'bold'), fg='black')
        self.label.pack(padx=5, pady=5)



        ############################
        ####### TREEVIEW ###########
        ############################

        # Add some style
        self.style = ttk.Style()
        # Pick a theme
        self.style.theme_use("default")

        # Configure our treeview colors
        self.style.configure("Treeview",
                        background="#D3D3D3",
                        foreground="black",
                        rowheight=30,
                        fieldbackground="#D3D3D3",
                        font=('Calibri', 12)
                        )

        # Headings
        self.style.configure("Treeview.Heading",
                        font=('Calibri', 12, 'bold')
                        )

        # Change selected color
        self.style.map('Treeview',
                  background=[('selected', 'blue')]
                  )

        # Treeview Scrollbar
        self.tree_scroll = ttk.Scrollbar(self.frame_tree)
        self.tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Create Treeview
        self.my_tree = ttk.Treeview(self.frame_tree, yscrollcommand=self.tree_scroll.set, selectmode="extended")
        # Pack to the screen
        self.my_tree.pack()

        # Configure the scrollbar
        self.tree_scroll.config(command=self.my_tree.yview)

        # Define Our Columns
        self.my_tree['columns'] = ("ID", "Celebranti")

        # Formate Our Columns
        self.my_tree.column("#0", width=0, stretch=tk.NO)
        self.my_tree.column("ID", anchor=tk.W, width=70)
        self.my_tree.column("Celebranti", anchor=tk.W, width=200)

        # Create Headings
        self.my_tree.heading("#0", text="", anchor=tk.W)
        self.my_tree.heading("ID", text="Id", anchor=tk.W)
        self.my_tree.heading("Celebranti", text="Celebranti", anchor=tk.W)

        def query_database(self):
            # Clear the Treeview
            for record in self.my_tree.get_children():
                self.my_tree.delete(record)

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
            self.my_tree.tag_configure('oddrow', background="white")
            self.my_tree.tag_configure('evenrow', background="lightblue")

            for record in records:
                if count % 2 == 0:
                    self.my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=(
                                       record[0],
                                       record[1]),
                                   tags=('evenrow'))
                else:
                    self.my_tree.insert(parent='', index=0, iid=record[0], text='',
                                   values=(
                                       record[0],
                                       record[1]),
                                   tags=('oddrow'))
                count += 1

            # Al termine del processo la prima riga risulta evidenziata
            child_id = self.my_tree.get_children()[0]  # la prima riga dall'alto del treeview
            self.my_tree.focus(child_id)  # evidenziata
            self.my_tree.selection_set(child_id)

            # Commit changes
            conn.commit()

            # Close our connection
            conn.close()

        query_database(self)

        def remove_one(self):
            # my_tree.focus() restituisce l'ID della riga selezionata
            row_id = self.my_tree.focus()
            self.my_tree.delete(row_id)

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

        def submit(self):
            conn = sqlite3.connect('database_messe_orizzontale')
            cur = conn.cursor()
            dati = [self.Entry_Celebranti_StringVar.get()]
            cur.execute('INSERT INTO TABLE_Celebranti (Celebranti) VALUES (?)', dati)
            # cur.execute("insert into TABLE_Celebranti (Celebranti) values ('?'), dati")
            # cur.execute('''INSERT INTO TABLE_Celebranti (Celebranti) VALUES ("Entry_Celebranti_StringVar.get()")''')
            conn.commit()
            # Close our connection
            conn.close()

        self.B_add_celebranti = tk.Button(self.frame, text='aggiungi', width=10, command=lambda: [query_database(self), submit(self), query_database(self)]).pack(
            side=tk.TOP, pady=20)
        self.B_delete_celebranti = tk.Button(self.frame, text='cancella', width=10, command=lambda: [remove_one(self)]).pack(side=tk.TOP, pady=20)



class Settings_super(SettingsWindows):
    def __init__(self):
        super().__init__()
        pass



class MainWindow:
    def __init__(self, master):
        self.master = master

        self.frame = tk.Frame(self.master, width=300, height=200, background='lightgrey')
        self.frame.pack()

        self.botton = tk.Button(self.frame, text='Click_me', command=self.c1)
        self.botton.place(x=50, y=50)

        self.botton2 = tk.Button(self.frame, text='?', command= lambda:[''])
        self.botton2.place(x=50, y=100)

        self.flag = 0

    def c1(self):
        self.settings = SettingsWindows()
        self.settings.win.mainloop()






root = tk.Tk()
window = MainWindow(root)
root.mainloop()