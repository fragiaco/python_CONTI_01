# provo ad importare conti_camerino
from xlsxwriter import *
import pandas as pd
#leggo e creo dataframe sensa indici colonna
df_conti_camerino=pd.read_excel('conti_camerino_da_importare.xlsx', header=None)
#verifico
#print(df_conti_camerino.head())
print() # riga vuota

#imposto gli indici colonna
df_conti_camerino.columns = ['Anno', 'Mese', 'Categoria', 'Voce','Euro']
#verifico
#print(df_conti_camerino.head())
print()

#salvo il file con un nome diverso
df_conti_camerino.to_excel('conti_camerino_modified.xlsx')
# Leggo il nuovo file, creo nuovo dataframe senza colonna indice
df_conti_camerino_modified = pd.read_excel('conti_camerino_modified.xlsx', index_col=0)

# Stampo le prime 5 righe. Potrei anche usare .head()
#print(df_conti_camerino_modified.iloc[0:5])

#print(df_conti_camerino_modified.head())
#print(df_conti_camerino_modified.info())

#seleziona solo le colonne desiderate
#wanted_columns = df_conti_camerino_modified[['Anno', 'Mese']]
#salva il fle come hai fatto prima con un nome diverso .to_excel
# Leggo il nuovo file, creo nuovo dataframe senza colonna indice



#pulizia delle righe con valori nulli
# print(df_conti_camerino_modified.isnull().sum())
# print(df_conti_camerino_modified.dropna(inplace=True))
#print(df_conti_camerino_modified.isnull().sum())



# print(df_conti_camerino_modified.loc
#       [df_conti_camerino_modified['Categoria']=='Sussidi'])



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
#print(df_conti_camerino_modified.head())

#df_conti_camerino_modified.to_excel('conti_camerino_modified_excel.xlsx')

########PIVOT
import numpy as np
#df_conti_camerino_modified['Anno'] = df_conti_camerino_modified['Anno'].astype(str)
df_conti_camerino_modified["Mese"] = df_conti_camerino_modified["Mese"].astype("category")
df_conti_camerino_modified["Mese"] = df_conti_camerino_modified["Mese"].cat.set_categories(["gennaio","febbraio","marzo","aprile","maggio","giugno","luglio","agosto","settembre","ottobre","novembre","dicembre"])

# Example 1 - Using loc[] with multiple conditions
                                                                    # df2=df.loc[(df['Discount'] >= 1000) & (df['Discount'] <= 2000)]
#
# # Example 2
# df2=df.loc[(df['Discount'] >= 1200) | (df['Fee'] >= 23000 )]
# print(df2)



# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('conti_camerino_multiple.xlsx', engine='xlsxwriter')





df_conti_camerino_pivot_tabellone_anno_entrate = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'gennaio') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Entrate')]

df_conti_camerino_pivot_tabellone_anno_uscite = df_conti_camerino_modified.loc[
                                        (df_conti_camerino_modified['Anno'] == 2015) &
                                        (df_conti_camerino_modified['Mese'] == 'gennaio') &
                                        (df_conti_camerino_modified['Entrate_Uscite'] == 'Uscite')]




print(df_conti_camerino_pivot_tabellone_anno_entrate.head())

pivot_gennaio_entrate = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_tabellone_anno_entrate,
                               values='Euro',
                               index=['Entrate_Uscite','Categoria','Voce'],
                               columns='Mese',
                               fill_value=0,
                               margins=True),2)

pivot_gennaio_uscite = np.round(pd.pivot_table
                            (df_conti_camerino_pivot_tabellone_anno_uscite,
                               values='Euro',
                               index=['Entrate_Uscite','Categoria','Voce'],
                               columns='Mese',
                               fill_value=0,
                               margins=True),2)

print(pivot_gennaio_entrate)
print(pivot_gennaio_uscite)

# Write each dataframe to a different worksheet.
df_conti_camerino_pivot_tabellone_anno_entrate.to_excel(writer, sheet_name='gennaio_entrate')
df_conti_camerino_pivot_tabellone_anno_uscite.to_excel(writer, sheet_name='gennaio_uscite')
#df_conti_camerino_multiple.to_excel(writer, sheet_name='multiple')
# with pd.ExcelWriter('conti_camerino_multiple.xlsx', engine='openpyxl', mode='a') as writer:
#     df_conti_camerino_pivot_tabellone_anno_uscite.to_excel(writer, sheet_name='gennaio_uscite')


# Write to Multiple Sheets
# with pd.ExcelWriter('Courses.xlsx') as writer:
#     df.to_excel(writer, sheet_name='Technologies')
#     df2.to_excel(writer, sheet_name='Schedule')
#
# # Append DataFrame to existing excel file
# with pd.ExcelWriter('Courses.xlsx',mode='a') as writer:
#     df.to_excel(writer, sheet_name='Technologies')




print (df_conti_camerino_pivot_tabellone_anno_uscite.shape)
# Close the Pandas Excel writer and output the Excel file.
writer.close()

pivot_gennaio_uscite.to_excel('conti_camerino_modified_excel.xlsx')