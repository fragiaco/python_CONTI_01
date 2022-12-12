
from tkinter        import *
from tkinter        import ttk
from tkinter        import messagebox

from OrigineSearch  import*

import mysql.connector
from sqlserver_config import dbConfig
#import pypyodbc as pyo

#con=pyo.connect(**dbConfig)

con = (mysql.connector.connect(**dbConfig))
print(con)
cursor = con.cursor()



#lass Mywallet():
#   def __init__(self):
#       self.con = mysql.connector.connect(**dbConfig)
#       self.cursor = con.cursor()
#       print('SEI CONNESSO AL DATABASE mywallet')
#       print(con)

#   def __del__(self):
#       self.con.close()

#   def view(self):
#       self.cursor.execute('SELECT * FROM transazione')
#       rows = self.cursor.fetchall()
#       return rows

#   def insert(self,Anno,Mese,EntrateUscite,Categoria,VoceSpesa,EventCommento,Euro):
#       sql=('INSERT INTO transazione(Anno, Mese, EntrateUscite, Categoria, VoceSpesa, EventCommento, Euro)VALUES(?, ?, ?, ?, ?, ?, ?')
#       values=[Anno, Mese, EntrateUscite, Categoria, VoceSpesa, EventCommento, Euro]
#       self.cursor.execute(sql,values)
#       self.con.commit()
#       messagebox.showinfo(title='Book Database', message='new book added')

#   def update(self, ID, Anno, Mese, EntrateUscite, Categoria, VoceSpesa, EventCommento, Euro):
#       tsql = 'UPDATE TableWallet SET  Anno = ?, Mese = ?, EntrateUscite = ?, Categoria = ?, VoceSpesa = ?, EventCommento = ?, Euro = ? WHERE ID=?'
#       self.cursor.execute(tsql, [Anno, Mese, EntrateUscite, Categoria, VoceSpesa, EventCommento, Euro, ID])
#       self.con.commit()
#       messagebox.showinfo(title="Book Database", message="Book Updated")

#   def delete(self, id):
#       delquery = 'DELETE FROM books WHERE ID = ?'
#       self.cursor.execute(delquery, [ID])
#       self.con.commit()
#       messagebox.showinfo(title="Book Database", message="Book Deleted")

#b= Mywallet()

#ef get_selected_row(event):
#   global selected_tuple
#   index = list_bx.curselection()[0]
#   selected_tuple = list_bx.get(index)
#   #title_entry.delete(0, 'end')
#   #title_entry.insert('end', selected_tuple[1])
#   #author_entry.delete(0, 'end')
#   #author_entry.insert('end', selected_tuple[2])
#   #isbn_entry.delete(0, 'end')
#   #isbn_entry.insert('end', selected_tuple[3])



class Conti:
    def __init__(self, root):
        self.root = root
        self.root.title("Programma gestione conti Convento Offida")
        self.root.geometry('1680x1050+0+0')

        self.title = Label(self.root, text='Database Entrate Uscite', font=('verdana', 40, 'bold'), bg='blue', fg='white')
        self.title.pack(side=TOP, fill=X)

        self.ID=StringVar()
        self.Anno=StringVar()
        self.Mese=StringVar()
        self.EntrateUscite=StringVar()
        self.Categorie=StringVar()
        self.Voce_Spesa=StringVar()
        self.Event_Commento=StringVar()
        self.Euro=StringVar()





        #Frame 1 - left side Frame
        Frame1 = Frame(self.root, bd='4', bg='blue', relief=RIDGE)
        Frame1.place(x=20, y=85, width=550, height=950)

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

        Frame1_Commento = Label(Frame1, text='Eventuale commento', font=('verdana', 15, 'bold'), bg='blue', fg='white')
        Frame1_Commento.grid(row=11, padx=25, pady=20, sticky='w')

        Frame1_Euro = Label(Frame1, text='Euro', font=('verdana', 15, 'bold'), bg='blue', fg='white')
        Frame1_Euro.grid(row=12 , padx=25, pady=20, sticky='w')


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



        # Dropbox
        anno_combo = ttk.Combobox(Frame1, font=("Helvetica", 15), values=Anni)
        anno_combo.current(0)
        anno_combo.grid(row=2, column=1)


        mesi_combo = ttk.Combobox(Frame1, font=("Helvetica", 15), values=Mesi)
        mesi_combo.current(0)
        mesi_combo.grid(row=4, column=1)


        my_combo = ttk.Combobox(Frame1, font=("Helvetica", 15), values=Entrate_Uscite)
        my_combo.current(0)
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

        # Eventuale_Commento ENTRY
        euro = Entry(Frame1, font=("Helvetica", 15, 'bold'), bd=5, relief=GROOVE)
        euro.grid(row=11, column=1)

        #euro ENTRY
        euro = Entry(Frame1, font=("Helvetica", 15, 'bold'), bd=5, relief=GROOVE)
        euro.grid(row=12, column=1)

        #Frame 1 - bottom side Frame
        Frame1in=Frame(Frame1, bd='4', bg='blue', relief=RIDGE)
        Frame1in.place(x=15, y=850, width=500, height=60)

        B_add = Button(Frame1in, text='add', width=10, command=self.Add_data).grid(row=0, column=0, padx=20, pady=15)
        B_update = Button(Frame1in, text='update', width=10).grid(row=0, column=1, padx=20, pady=15)
        B_delete = Button(Frame1in, text='delete', width=10).grid(row=0, column=2, padx=20, pady=15)
        B_clear = Button(Frame1in, text='clear', width=10).grid(row=0, column=3, padx=20, pady=15)

        # Frame 2 - right side Frame
        Frame2 = Frame(self.root, bd='4', bg='blue', relief=RIDGE)
        Frame2.place(x=590, y=85, width=1070, height=950)

        L_Search=Label(Frame2, text='Search by', bg='blue', fg='white', font=('verdana', 12, 'bold'))
        L_Search.grid(row=0, column=0, padx=20, pady=10, sticky='w')

        B_Search = Button(Frame2, text='Search', width=10).grid(row=0, column=1, padx=20, pady=15)
        B_Clear_Search = Button(Frame2, text='Clear', width=10).grid(row=0, column=2, padx=20, pady=15)
        B_Show_all = Button(Frame2, text='Show All', width=10, command=self.Display_data).grid(row=0, column=3, padx=20, pady=15)


        Frame2_anno = Label(Frame2, text='Anno', font=('verdana', 7, 'bold'), bg='blue', fg='white')
        Frame2_anno.grid(row=1, column=0, padx=20, pady=10, sticky='w')

        Frame2_mese = Label(Frame2, text='Mese', font=('verdana', 7, 'bold'), bg='blue', fg='white')
        Frame2_mese.grid(row=1, column=1, padx=20, pady=25, sticky='w')

        Frame2_EntrateUscite = Label(Frame2, text='Entrate/Uscite', font=('verdana', 7, 'bold'), bg='blue', fg='white')
        Frame2_EntrateUscite.grid(row=1, column=2, padx=25, pady=20, sticky='w')

        Frame2_Categoria = Label(Frame2, text='Categoria', font=('verdana', 7, 'bold'), bg='blue', fg='white')
        Frame2_Categoria.grid(row=1, column=3)

        Frame2_VoceSpesa = Label(Frame2, text='Voce di spesa', font=('verdana', 7, 'bold'), bg='blue', fg='white')
        Frame2_VoceSpesa.grid(row=1, column=4, padx=25, pady=20, sticky='w')



        def pick_Categoria(e):
            if my_combo_search.get() == "Entrate":
                categoria_combo_search.config(values=Categorie_Entrate_Search)
                categoria_combo_search.current(0)


            if my_combo_search.get() == "Uscite":
                categoria_combo_search.config(values=Categorie_Uscite_Search)
                categoria_combo_search.current(0)

            if my_combo_search.get() == "ALL":
                categoria_combo_search.config(values=All_Data_Search)
                categoria_combo_search.current(0)

        def pick_Voce(e):
            # VOCI ENTRATA
            if categoria_combo_search.get() == "Collette_Chiesa":
                voce_combo_search.config(values=Collette_Chiesa_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Congrua":
                voce_combo_search.config(values=Congrua_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Interessi":
                voce_combo_search.config(values=Interessi_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Messe celebrate":
                voce_combo_search.config(values=Messe_celebrate_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Offerte":
                voce_combo_search.config(values=Offerte_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Pensioni":
                voce_combo_search.config(values=Pensioni_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Servizi_religiosi":
                voce_combo_search.config(values=Servizi_religiosi_Search)
                voce_combo_search.current(0)


            if categoria_combo_search.get() == "ALL":
                voce_combo_search.config(values=All_Data_Search)
                voce_combo_search.current(0)

            # VOCI USCITA

            if categoria_combo_search.get() == "Acquisti_Chiesa":
                voce_combo_search.config(values=Acquisti_Chiesa_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Acquisti_Convento":
                voce_combo_search.config(values=Acquisti_Convento_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Acquisti_Orto_Animali":
                voce_combo_search.config(values=Acquisti_Orto_Animali_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Cultura":
                voce_combo_search.config(values=Cultura_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Curia_provinciale":
                voce_combo_search.config(values=Curia_provinciale_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Domestici":
                voce_combo_search.config(values=Domestici_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Elargizioni":
                voce_combo_search.config(values=Elargizioni_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Utenze":
                voce_combo_search.config(values=Utenze_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Ferie_Viaggi":
                voce_combo_search.config(values=Ferie_Viaggi_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Igiene":
                voce_combo_search.config(values=Igiene_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Imposte":
                voce_combo_search.config(values=Imposte_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Lavori_Impianti":
                voce_combo_search.config(values=Lavori_Impianti_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Posta_Cancelleria":
                voce_combo_search.config(values=Posta_Cancelleria_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Salute":
                voce_combo_search.config(values=Salute_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Veicoli_motore":
                voce_combo_search.config(values=Veicoli_motore_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Vestiario":
                voce_combo_search.config(values=Vestiario_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Vitto":
                voce_combo_search.config(values=Vitto_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "Eccedenza_Cassa":
                voce_combo_search.config(values=Eccedenza_Cassa_Search)
                voce_combo_search.current(0)

            if categoria_combo_search.get() == "ALL":
                voce_combo_search.config(values=All_Data_Search)
                voce_combo_search.current(0)

        # Dropbox Search
        anno_combo_search = ttk.Combobox(Frame2, font=("Helvetica", 8), values=Anno_Search)
        anno_combo_search.current(0)
        anno_combo_search.grid(row=2, column=0, padx=5)


        mesi_combo_search = ttk.Combobox(Frame2, font=("Helvetica", 8), values=Mesi_Search)
        mesi_combo_search.current(0)
        mesi_combo_search.grid(row=2, column=1, padx=2)

        my_combo_search = ttk.Combobox(Frame2, font=("Helvetica", 8), values=Entrate_Uscite_Search)
        my_combo_search.current(0)
        my_combo_search.grid(row=2, column=2, padx=2)

        # Bind the ComboBox
        my_combo_search.bind("<<ComboboxSelected>>", pick_Categoria)

        # Categoria ComboBox
        categoria_combo_search = ttk.Combobox(Frame2, font=("Helvetica", 8), values=[""])
        categoria_combo_search.current(0)
        categoria_combo_search.grid(row=2, column=3, padx=2)

        # Bind the ComboBox
        categoria_combo_search.bind("<<ComboboxSelected>>", pick_Voce)

        # Voce Entrata_Spesa ComboBox Combo Box
        voce_combo_search = ttk.Combobox(Frame2, font=("Helvetica", 8), values=[""])
        voce_combo_search.current(0)
        voce_combo_search.grid(row=2, column=4,padx=2)

        # T_Frame - Tableview Frame
        T_Frame = Frame(Frame2, bd='4', bg='blue', relief=RIDGE)
        T_Frame.place(x=5, y=170, width=1050, height=650)

        scroll_x=Scrollbar(T_Frame, orient=HORIZONTAL)
        scroll_y=Scrollbar(T_Frame, orient=VERTICAL)

        self.Emp=ttk.Treeview(T_Frame, columns=('ID', 'Anno', 'Mese', 'Entrate/Uscite', 'Categoria', 'Voce_Spesa','Event_Commento', 'Euro'), xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)
        scroll_x.pack(side=BOTTOM, fill=X)
        scroll_y.pack(side=RIGHT, fill=Y)

        scroll_x.config(command=self.Emp.xview)
        scroll_y.config(command=self.Emp.yview)
        self.Emp.heading('ID', text="ID")
        self.Emp.heading('Anno', text="ANNO")
        self.Emp.heading('Mese', text="MESE")
        self.Emp.heading('Entrate/Uscite', text="ENTRATE/USCITE")
        self.Emp.heading('Categoria', text="CATEGORIA")
        self.Emp.heading('Voce_Spesa', text="VOCE SPESA")
        self.Emp.heading('Event_Commento', text="COMMENTO")
        self.Emp.heading('Euro', text="EURO")
        self.Emp['show']='headings'
        self.Emp.column('ID', width=50)
        self.Emp.column('Anno', width=50)
        self.Emp.column('Mese', width=50)
        self.Emp.column('Entrate/Uscite', width=50)
        self.Emp.column('Categoria', width=50)
        self.Emp.column('Voce_Spesa', width=50)
        self.Emp.column('Event_Commento', width=50)
        self.Emp.column('Euro', width=50)


        self.Emp.pack(fill=BOTH, expand=1)



    def Display_data(self):
        con = mysql.connector.connect(**dbConfig)
        cur = con.cursor()
        cur.execute('Select * from transazione')
        rows = cur.fetchall()
        if len(rows) != 0:
            self.Emp.delete(*self.Emp.get_children())
            for row in rows:
                self.Emp.insert('', END, values=row)
            con.commit()
        con.close()

    def Add_data(self, Anno, Mese, EntrateUscite, Categorie, Voce_Spesa, Event_Commento, Euro):
        con = mysql.connector.connect(**dbConfig)
        cur = con.cursor()
        cur.execute ('insert into transazione(Anno, Mese, EntrateUscite, Categorie, Voce_Spesa, Event_Commento, Euro) values (%s,%s,%s,%s,%s,%s,%s)',\
                    (self.Anno.get(),\
                     self.Mese.get(),\
                     self.EntrateUscite.get(),\
                     self.Categorie.get(),\
                     self.Voce_Spesa.get(),\
                     self.Event_Commento.get(),\
                     self.Euro.get()))

        def get_selected_row(event):
            global selected_tuple
            index = self.Emp.curselection()[0]
            selected_tuple = self.Emp.get(index)
            title_entry.delete(0, 'end')
            title_entry.insert('end', selected_tuple[1])
            author_entry.delete(0, 'end')
            author_entry.insert('end', selected_tuple[2])
            isbn_entry.delete(0, 'end')
            isbn_entry.insert('end', selected_tuple[3])
            print(index)
            print(selected_tuple)

        con.commit()
        con.close()

        #self.ID=StringVar()
        #self.Anno=StringVar()
        #self.Mese=StringVar()
        #self.EntrateUscite=StringVar()
        #self.Categorie=StringVar()
        #self.Voce_Spesa=StringVar()
        #self.Event_Commento=StringVar()
        #self.Euro=StringVar()




root = Tk()
obj = Conti(root)
root.mainloop()
