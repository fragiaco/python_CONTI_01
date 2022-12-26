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
# Leggo il nuovo file, creo nuovo dataframe
df_conti_camerino_modified = pd.read_excel('conti_camerino_modified.xlsx', index_col=0)
# Stampo le prime 5 righe. Potrei anche usare .head()
print(df_conti_camerino_modified.iloc[0:5])

#print(df_conti_camerino_modified.head())
print(df_conti_camerino_modified.info())
