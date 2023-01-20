import pandas as pd
import sqlite3

# CONNESSIONE A SQLITE3
conn = sqlite3.connect('database_conti')
cur = conn.cursor()

# MI ASSIOCURO DI ESSERE CONNESSO
# print(conn)
# print('Sei connesso al database_conti')

# CREO DATAFRAME df_database_conti
df_database_conti = pd.read_sql("select * from TABLE_Conti", conn)
# print(df_database_conti.info(verbose=True))


# COMMIT e CLOSE
conn.commit()
conn.close()

# LIST TITOLI DATAFRAMES 12 MESI
list_df_conti_mese_entrate =   ['df_database_conti_entrate_gennaio',
                                                         'df_database_conti_entrate_febbraio',
                                                         'df_database_conti_entrate_marzo',
                                                         'df_database_conti_aprile',
                                                         'df_database_conti_maggio',
                                                         'df_database_conti_giugno',
                                                         'df_database_conti_luglio',
                                                         'df_database_conti_agosto',
                                                         'df_database_conti_settembre',
                                                         'df_database_conti_ottobre',
                                                         'df_database_conti_novembre',
                                                         'df_database_conti_dicembre'
                                                         ]

list_df_conti_mese_uscite = ['df_database_conti_uscite_gennaio',
                                                      'df_database_conti_uscite_febbraio',
                                                      'df_database_conti_uscite_marzo',
                                                      'df_database_conti_uscite_aprile',
                                                      'df_database_conti_uscite_maggio',
                                                      'df_database_conti_uscite_giugno',
                                                      'df_database_conti_uscite_luglio',
                                                      'df_database_conti_uscite_agosto',
                                                      'df_database_conti_uscite_settembre',
                                                      'df_database_conti_uscite_ottobre',
                                                      'df_database_conti_uscite_novembre',
                                                      'df_database_conti_uscite_dicembre'
                                                      ]

list_mese = ['gennaio',
                             'febbraio',
                             'marzo',
                             'aprile',
                             'maggio',
                             'giugno',
                             'luglio',
                             'agosto',
                             'settembre',
                             'ottobre',
                             'novembre',
                             'dicembre'
                             ]

########### imposto anno ##############
# anno = Report.anno_report_func(anno_report_Stringvar)
# B_report = Button(Frame_excell_botton, text='report', width=10, command= lambda: print(Report.anno_report_func(anno_report_Stringvar))).grid(row=0, column=2, padx=20, pady=15)

anno= 2020

i = 0
for x in range (12):

                        list_df_conti_mese_entrate[i]= df_database_conti.loc[
                            (df_database_conti['Anno'] == anno) &
                            (df_database_conti['Mese'] == list_mese[i]) &
                            (df_database_conti['Entrate_Uscite'] == 'Entrate')]
                        #print(list_df_conti_mese_entrate[i].empty)


                        list_df_conti_mese_uscite[i] = df_database_conti.loc[
                            (df_database_conti['Anno'] == anno) &
                            (df_database_conti['Mese'] == list_mese[i]) &
                            (df_database_conti['Entrate_Uscite'] == 'Uscite')]
                        #print(list_df_conti_camerino_mese_uscite[i].head())

                        i += 1

i=0


for x in range (12):
                        list_mese = ['gennaio',
                                     'febbraio',
                                     'marzo',
                                     'aprile',
                                     'maggio',
                                     'giugno',
                                     'luglio',
                                     'agosto',
                                     'settembre',
                                     'ottobre',
                                     'novembre',
                                     'dicembre'
                                     ]
                        dataframe_empty_list = [(i, anno, list_mese[i], 'Entrate', 'vuoto', 'vuoto', 0)]
                        if  list_df_conti_mese_entrate[i].empty:
                                list_df_conti_mese_entrate[i] = pd.DataFrame\
                                    (dataframe_empty_list, columns = ['index', 'Anno', 'Mese', 'Entrate_Uscite', 'Categoria', 'Voce', 'Euro'])


                        dataframe_empty_list = [(i, anno, list_mese[i], 'Uscite', 'vuoto', 'vuoto', 0)]
                        if  list_df_conti_mese_uscite[i].empty:
                                list_df_conti_mese_uscite[i] = pd.DataFrame\
                                    (dataframe_empty_list, columns = ['index', 'Anno', 'Mese', 'Entrate_Uscite', 'Categoria', 'Voce', 'Euro'])

                        print(list_df_conti_mese_entrate[i].to_markdown())
                        print('')
                        print(list_df_conti_mese_uscite[i].to_markdown())
                        print('')

                        #print(list_df_conti_mese_uscite[i].info())
                        i +=1


list_df_conti_mese_entrate = [  list_df_conti_mese_entrate[0], #df_conti_camerino_pivot_entrate_gennaio
                                                            list_df_conti_mese_entrate[1],
                                                            list_df_conti_mese_entrate[2],
                                                            list_df_conti_mese_entrate[3],
                                                            list_df_conti_mese_entrate[4],
                                                            list_df_conti_mese_entrate[5],
                                                            list_df_conti_mese_entrate[6],
                                                            list_df_conti_mese_entrate[7],
                                                            list_df_conti_mese_entrate[8],
                                                            list_df_conti_mese_entrate[9],
                                                            list_df_conti_mese_entrate[10],
                                                            list_df_conti_mese_entrate[11]
                                                        ]

list_df_conti_mese_uscite = [   list_df_conti_mese_uscite[0], #df_conti_camerino_pivot_uscite_gennaio
                                                            list_df_conti_mese_uscite[1],
                                                            list_df_conti_mese_uscite[2],
                                                            list_df_conti_mese_uscite[3],
                                                            list_df_conti_mese_uscite[4],
                                                            list_df_conti_mese_uscite[5],
                                                            list_df_conti_mese_uscite[6],
                                                            list_df_conti_mese_uscite[7],
                                                            list_df_conti_mese_uscite[8],
                                                            list_df_conti_mese_uscite[9],
                                                            list_df_conti_mese_uscite[10],
                                                            list_df_conti_mese_uscite[11]
                                                        ]

# print(list_df_conti_mese_entrate[0].columns)
print('##########################################')
print(list_df_conti_mese_entrate[0].to_markdown())
print('')
print(list_df_conti_mese_uscite[0].to_markdown())
print(df_database_conti)
new_df_database_conti = df_database_conti[df_database_conti.Anno.notnull() &
                                            (df_database_conti.Mese.notnull())]
#df[df.origin.notnull()]
# newdf = df[(df.origin == "JFK") & (df.carrier == "B6")]

#df = pd.DataFrame({'col1': [1, 2], 'col2': [3, 4]})