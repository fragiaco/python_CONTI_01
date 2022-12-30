
import pandas as pd
import numpy as np
from openpyxl.styles import Font, Alignment
from openpyxl import load_workbook

df_conti_camerino_modified = pd.read_excel('conti_camerino_modified.xlsx', index_col=0)

df_conti_camerino_modified['Entrate_Uscite']= df_conti_camerino_modified['Categoria'].apply\
    (lambda x: 'Entrate'    if  x=='Vendite varie' or
                                x=='Salute' or
                                x=='Curia' or
                                x=='Collette-Chiesa' or
                                x=='Congrua' or
                                x=='Interessi' or
                                x=='Messe celebrate' or
                                x=='Offerte' or
                                x=='Pensioni' or
                                x=='Predicazione' or
                                x=='Servizi religiosi' or
                                x=='Stipendi' or
                                x=='Sussidi' or
                                x=='Rimbosi' or
                                x=='Vendite varie' or
                                x=='Eccedenza Cassa'
                            else 'Uscite')


df_conti_camerino_modified = df_conti_camerino_modified [['Anno', 'Mese', 'Entrate_Uscite', 'Categoria', 'Voce','Euro']]

df_conti_camerino_modified["Mese"] = df_conti_camerino_modified["Mese"].astype("category")
df_conti_camerino_modified["Mese"] = df_conti_camerino_modified["Mese"].cat.set_categories(["gennaio","febbraio","marzo","aprile","maggio","giugno","luglio","agosto","settembre","ottobre","novembre","dicembre"])



df_conti_camerino_pivot_entrate_gennaio = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'gennaio') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Entrate')]


df_conti_camerino_pivot_entrate_febbraio = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'febbraio') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Entrate')]

df_conti_camerino_pivot_entrate_marzo = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'marzo') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Entrate')]

df_conti_camerino_pivot_entrate_aprile = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'aprile') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Entrate')]

df_conti_camerino_pivot_entrate_maggio = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'maggio') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Entrate')]

df_conti_camerino_pivot_entrate_giugno = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'giugno') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Entrate')]

df_conti_camerino_pivot_entrate_luglio = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'luglio') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Entrate')]

df_conti_camerino_pivot_entrate_agosto = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'agosto') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Entrate')]

df_conti_camerino_pivot_entrate_settembre = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'settembre') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Entrate')]

df_conti_camerino_pivot_entrate_ottobre = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'ottobre') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Entrate')]

df_conti_camerino_pivot_entrate_novembre = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'novembre') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Entrate')]

df_conti_camerino_pivot_entrate_dicembre = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'dicembre') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Entrate')]

########


df_conti_camerino_pivot_uscite_gennaio = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'gennaio') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Uscite')]

df_conti_camerino_pivot_uscite_febbraio = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'febbraio') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Uscite')]

df_conti_camerino_pivot_uscite_marzo = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'marzo') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Uscite')]

df_conti_camerino_pivot_uscite_aprile = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'aprile') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Uscite')]

df_conti_camerino_pivot_uscite_maggio = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'maggio') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Uscite')]

df_conti_camerino_pivot_uscite_giugno = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'giugno') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Uscite')]

df_conti_camerino_pivot_uscite_luglio = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'luglio') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Uscite')]

df_conti_camerino_pivot_uscite_agosto = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'agosto') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Uscite')]

df_conti_camerino_pivot_uscite_settembre = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'settembre') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Uscite')]

df_conti_camerino_pivot_uscite_ottobre = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'ottobre') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Uscite')]

df_conti_camerino_pivot_uscite_novembre = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'novembre') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Uscite')]

df_conti_camerino_pivot_uscite_dicembre = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'dicembre') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Uscite')]


pivot_gennaio_entrate = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_entrate_gennaio,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_febbraio_entrate = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_entrate_febbraio,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_marzo_entrate = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_entrate_marzo,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_aprile_entrate = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_entrate_aprile,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_maggio_entrate = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_entrate_maggio,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_giugno_entrate = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_entrate_giugno,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_luglio_entrate = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_entrate_luglio,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_agosto_entrate = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_entrate_agosto,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_settembre_entrate = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_entrate_settembre,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_ottobre_entrate = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_entrate_ottobre,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_novembre_entrate = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_entrate_novembre,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_dicembre_entrate = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_entrate_dicembre,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

#############################


pivot_gennaio_uscite = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_uscite_gennaio,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_febbraio_uscite = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_uscite_febbraio,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_marzo_uscite = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_uscite_marzo,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_aprile_uscite = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_uscite_aprile,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_maggio_uscite = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_uscite_maggio,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_giugno_uscite = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_uscite_giugno,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_luglio_uscite = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_uscite_luglio,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_agosto_uscite = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_uscite_agosto,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_settembre_uscite = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_uscite_settembre,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_ottobre_uscite = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_uscite_ottobre,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_novembre_uscite = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_uscite_novembre,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)

pivot_dicembre_uscite = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_uscite_dicembre,
                               values='Euro',
                               index=['Entrate_Uscite', 'Categoria','Voce'],
                               aggfunc='sum',
                               margins=True,
                               margins_name= 'TOTALE',
                               fill_value=0),2)


# with pd.ExcelWriter("conti_camerino_multiple.xlsx") as writer:
#     pivot_gennaio_uscite.to_excel(writer, sheet_name='gennaio_uscite')
#     pivot_febbraio_uscite.to_excel(writer, sheet_name='gennaio_uscite')
#     pivot_marzo_uscite.to_excel(writer, sheet_name='gennaio_uscite')
#     pivot_aprile_uscite.to_excel(writer, sheet_name='gennaio_uscite')
#     pivot_maggio_uscite.to_excel(writer, sheet_name='gennaio_uscite')
#     pivot_giugno_uscite.to_excel(writer, sheet_name='gennaio_uscite')
#     pivot_luglio_uscite.to_excel(writer, sheet_name='gennaio_uscite')
#     pivot_agosto_uscite.to_excel(writer, sheet_name='gennaio_uscite')
#     pivot_settembre_uscite.to_excel(writer, sheet_name='gennaio_uscite')
#     pivot_ottobre_uscite.to_excel(writer, sheet_name='gennaio_uscite')
#     pivot_novembre_uscite.to_excel(writer, sheet_name='gennaio_uscite')
#     pivot_dicembre_uscite.to_excel(writer, sheet_name='gennaio_uscite')
#
#
#
#     pivot_gennaio_entrate.to_excel(writer, sheet_name='gennaio_entrate')


with pd.ExcelWriter("conti_camerino_multiple.xlsx",
                    mode="a",
                    engine="openpyxl",
                    if_sheet_exists="overlay",
                    ) as writer:
                    pivot_gennaio_entrate.to_excel(writer, sheet_name="Gennaio", startrow=5)
                    pivot_gennaio_uscite.to_excel(writer, sheet_name="Gennaio", startrow=(len(pivot_gennaio_entrate)+10))

with pd.ExcelWriter("conti_camerino_multiple.xlsx",
                    mode="a",
                    engine="openpyxl",
                    if_sheet_exists="overlay",
                    ) as writer:
                    pivot_febbraio_entrate.to_excel(writer, sheet_name="Febbraio", startrow=5)
                    pivot_febbraio_uscite.to_excel(writer, sheet_name="Febbraio", startrow=(len(pivot_gennaio_entrate) + 10))

with pd.ExcelWriter("conti_camerino_multiple.xlsx",
                    mode="a",
                    engine="openpyxl",
                    if_sheet_exists="overlay",
                    ) as writer:
                    pivot_marzo_entrate.to_excel(writer, sheet_name="Marzo", startrow=5)
                    pivot_marzo_uscite.to_excel(writer, sheet_name="Marzo", startrow=(len(pivot_gennaio_entrate) + 10))


with pd.ExcelWriter("conti_camerino_multiple.xlsx",
                    mode="a",
                    engine="openpyxl",
                    if_sheet_exists="overlay",
                    ) as writer:
                    pivot_aprile_entrate.to_excel(writer, sheet_name="Aprile", startrow=5)
                    pivot_aprile_uscite.to_excel(writer, sheet_name="Aprile", startrow=(len(pivot_gennaio_entrate) + 10))


with pd.ExcelWriter("conti_camerino_multiple.xlsx",
                    mode="a",
                    engine="openpyxl",
                    if_sheet_exists="overlay",
                    ) as writer:
                    pivot_maggio_entrate.to_excel(writer, sheet_name="Maggio", startrow=5)
                    pivot_maggio_uscite.to_excel(writer, sheet_name="Maggio", startrow=(len(pivot_gennaio_entrate) + 10))


with pd.ExcelWriter("conti_camerino_multiple.xlsx",
                    mode="a",
                    engine="openpyxl",
                    if_sheet_exists="overlay",
                    ) as writer:
                    pivot_giugno_entrate.to_excel(writer, sheet_name="Giugno", startrow=5)
                    pivot_giugno_uscite.to_excel(writer, sheet_name="Giugno", startrow=(len(pivot_gennaio_entrate) + 10))


with pd.ExcelWriter("conti_camerino_multiple.xlsx",
                    mode="a",
                    engine="openpyxl",
                    if_sheet_exists="overlay",
                    ) as writer:
                    pivot_luglio_entrate.to_excel(writer, sheet_name="Luglio", startrow=5)
                    pivot_luglio_uscite.to_excel(writer, sheet_name="Luglio", startrow=(len(pivot_gennaio_entrate) + 10))


with pd.ExcelWriter("conti_camerino_multiple.xlsx",
                    mode="a",
                    engine="openpyxl",
                    if_sheet_exists="overlay",
                    ) as writer:
                    pivot_agosto_entrate.to_excel(writer, sheet_name="Agosto", startrow=5)
                    pivot_agosto_uscite.to_excel(writer, sheet_name="Agosto", startrow=(len(pivot_gennaio_entrate) + 10))


with pd.ExcelWriter("conti_camerino_multiple.xlsx",
                    mode="a",
                    engine="openpyxl",
                    if_sheet_exists="overlay",
                    ) as writer:
                    pivot_settembre_entrate.to_excel(writer, sheet_name="Settembre", startrow=5)
                    pivot_settembre_uscite.to_excel(writer, sheet_name="Settembre", startrow=(len(pivot_gennaio_entrate) + 10))


with pd.ExcelWriter("conti_camerino_multiple.xlsx",
                    mode="a",
                    engine="openpyxl",
                    if_sheet_exists="overlay",
                    ) as writer:
                    pivot_ottobre_entrate.to_excel(writer, sheet_name="Ottobre", startrow=5)
                    pivot_ottobre_uscite.to_excel(writer, sheet_name="Ottobre", startrow=(len(pivot_gennaio_entrate) + 10))


with pd.ExcelWriter("conti_camerino_multiple.xlsx",
                    mode="a",
                    engine="openpyxl",
                    if_sheet_exists="overlay",
                    ) as writer:
                    pivot_novembre_entrate.to_excel(writer, sheet_name="Novembre", startrow=5)
                    pivot_novembre_uscite.to_excel(writer, sheet_name="Novembre", startrow=(len(pivot_gennaio_entrate) + 10))


with pd.ExcelWriter("conti_camerino_multiple.xlsx",
                    mode="a",
                    engine="openpyxl",
                    if_sheet_exists="overlay",
                    ) as writer:
                    pivot_dicembre_entrate.to_excel(writer, sheet_name="Dicembre", startrow=5)
                    pivot_dicembre_uscite.to_excel(writer, sheet_name="Dicembre", startrow=(len(pivot_gennaio_entrate) + 10))



wb = load_workbook(filename = "conti_camerino_multiple.xlsx")
# ws_entrate = wb['gennaio_entrate']
# ws_uscite = wb['gennaio_uscite']

ws_gennaio = wb['Gennaio']
ws_febbraio = wb['Febbraio']
ws_marzo = wb['Marzo']
ws_aprile = wb['Aprile']
ws_maggio = wb['Maggio']
ws_giugno = wb['Giugno']
ws_luglio = wb['Luglio']
ws_agosto = wb['Agosto']
ws_settembre = wb['Settembre']
ws_ottobre = wb['Ottobre']
ws_novembre = wb['Novembre']
ws_dicembre = wb['Dicembre']


list_ws_mese = [ws_gennaio,
                ws_febbraio,
                ws_marzo,
                ws_aprile,
                ws_maggio,
                ws_giugno,
                ws_luglio,
                ws_agosto,
                ws_settembre,
                ws_ottobre,
                ws_novembre,
                ws_dicembre]

i = 0 # contatore list_mese
list_mese = ['Conto del mese di Gennaio',
             'Conto del mese di Febbraio',
             'Conto del mese di Marzo',
             'Conto del mese di Aprile',
             'Conto del mese di Maggio',
             'Conto del mese di Giugno',
             'Conto del mese di Luglio',
             'Conto del mese di Agosto',
             'Conto del mese di Settembre',
             'Conto del mese di Ottobre',
             'Conto del mese di Novembre',
             'Conto del mese di Dicembre',
             ]

# set the height of the row
for sheet in list_ws_mese:
    sheet.row_dimensions[1].height = 70
#ws_gennaio.row_dimensions[1].height = 70
# set the width of the column
    sheet.column_dimensions['A'].width = 15
    sheet.column_dimensions['B'].width = 20
    sheet.column_dimensions['C'].width = 30
    sheet.column_dimensions['D'].width = 15
#merge cells
    sheet.merge_cells('A1:D1')
    top_left_cell = sheet['A1']

    top_left_cell.value = list_mese[i]
    i += 1
    top_left_cell.font=Font(name='Calibri',
                            size=25,
                            bold=True,
                            italic=True,
                            vertAlign='none',
                            underline='single',
                            strike=False,
                            color='a81a1a')

    top_left_cell.alignment = Alignment(horizontal="center", vertical="center")


for sheet in list_ws_mese:
    for row in sheet[7:sheet.max_row]:  # skip the header
        #print(row) #(<Cell 'multiple'.A7>, <Cell 'multiple'.B7>, <Cell 'multiple'.C7>, <Cell 'multiple'.D7>)
        cell = row[3] #il quarto valore della tuple
        #print (cell)# <Cell 'multiple'.D7>
        cell.number_format = '#,##0.00€'
        cell.alignment = Alignment(horizontal="right")
        cell.font=Font(bold=True)

    for row in sheet[7:sheet.max_row]:  # skip the header
        cell = row[2]  #il terzo valore della tuple
        cell.alignment = Alignment(horizontal="right")

    for row in sheet[7:sheet.max_row]:  # skip the header
        cell = row[1]  #il secondo valore della tuple
        cell.alignment = Alignment(horizontal="center", vertical="center")


    list=[]
    # Enumerate the cells in the second row
    for row in sheet.rows:
        for cell in row:
            if (cell.value == ("Categoria") or
                cell.value == ("Entrate") or
                cell.value == ("Euro") or
                cell.value == ("Uscite") or
                cell.value == ('TOTALE') or
                cell.value == ("Voce")):
                print('trovato')
                print(cell)
                list.append(cell)
    print(list)

    for cell in list:
        cell.font = Font(size=15, color='a81a1a', bold=True)



    list=[]
    # Enumerate the cells in the second row
    for row in sheet.rows:
        for cell in row:
            if cell.value == ("Entrate_Uscite"):
                list.append(cell)
    for cell in list:
        cell.font = Font(size=1)

    # for cell in list:
    #     cell.font = Font(size=1)
    #     print(cell)
    #     print(cell.coordinate)
    #     print(cell.row)
    #     print(cell.column)

    list=[]

    for row in sheet.rows:
        for cell in row:
            if cell.value == ("TOTALE"):
                list.append(cell)
    for cell in list:
        sheet.cell(cell.row, column=4).font = Font(size=15, color='a81a1a', bold=True)


wb.save("conti_camerino_styled.xlsx")