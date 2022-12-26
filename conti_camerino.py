# provo ad importare conti_camerino

import pandas as pd
#leggo e creo dataframe sensa indici colonna
df_conti_camerino=pd.read_excel('conti_camerino_da_importare.xlsx', header=None)
#verifico
print(df_conti_camerino.info())
print(df_conti_camerino.head())
print() # riga vuota

#imposto gli indici colonna
df_conti_camerino.columns = ['Anno', 'Mese', 'Categoria', 'Voce','Euro']
#verifico
print(df_conti_camerino.head())
print()

#salvo il file con un nome diverso
df_conti_camerino.to_excel('conti_camerino_modified.xlsx')
# Leggo il nuovo file, creo nuovo dataframe senza colonna indice
df_conti_camerino_modified = pd.read_excel('conti_camerino_modified.xlsx', index_col=0)

# Stampo le prime 5 righe. Potrei anche usare .head()
print(df_conti_camerino_modified.iloc[0:5])

#print(df_conti_camerino_modified.head())
print(df_conti_camerino_modified.info())

#seleziona solo le colonne desiderate
#wanted_columns = df_conti_camerino_modified[['Anno', 'Mese']]
#salva il fle come hai fatto prima con un nome diverso .to_excel
# Leggo il nuovo file, creo nuovo dataframe senza colonna indice

#Creo una colonna Entrate-Uscite e confronto con questi dati
'''
['Curia',
'Collette-Chiesa',
'Congrua',
'Interessi',
'Messe celebrate',
'Offerte',
'Pensioni',
'Predicazione',
'Servizi religiosi',
'Stipendi',
'Sussidi',
'Rimborsi',
'Vendite varie',
'Eccedenza Cassa']
'''


print(df_conti_camerino_modified.isnull().sum())
print(df_conti_camerino_modified.dropna(inplace=True))
print(df_conti_camerino_modified.isnull().sum())

# print(df_conti_camerino_modified.loc
#       [df_conti_camerino_modified['Categoria']=='Sussidi'])

print(df_conti_camerino_modified.loc
        [(df_conti_camerino_modified['Categoria']=='Curia') |
        (df_conti_camerino_modified['Categoria']=='Collette-Chiesa') |
        (df_conti_camerino_modified['Categoria']=='Congrua') |
        (df_conti_camerino_modified['Categoria']=='Interessi') |
        (df_conti_camerino_modified['Categoria']=='Messe celebrate') |
        (df_conti_camerino_modified['Categoria']=='Offerte') |
        (df_conti_camerino_modified['Categoria']=='Pensioni') |
        (df_conti_camerino_modified['Categoria']=='Predicazione') |
        (df_conti_camerino_modified['Categoria']=='Servizi religiosi') |
        (df_conti_camerino_modified['Categoria']=='Stipendi') |
        (df_conti_camerino_modified['Categoria']=='Rimborsi') |
        (df_conti_camerino_modified['Categoria']=='Sussidi') |
        (df_conti_camerino_modified['Categoria']=='Vendite varie') |
        (df_conti_camerino_modified['Categoria']=='Eccedenza Cassa')
         ])


print(df_conti_camerino_modified.isnull().sum())
