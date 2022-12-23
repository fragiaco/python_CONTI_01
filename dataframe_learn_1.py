import pandas as pd
from openpyxl.workbook import Workbook
import sqlite3

conn = sqlite3.connect('database_conti')
# Create a cursor instance
c = conn.cursor()

query = "SELECT * FROM TABLE_Conti"  # query to collect recors
df = pd.read_sql(query, conn)

# print('#####')
# print(type(df))
# print('#####')
# print(df.head())
# print('#####')
# print(df.info())
# print('#####')
# print(len(df))
# print('#####')
# print(df.shape)
# print(df.index)
# print(df.columns)
# print('#####')
#
# print((df['Anno']).equals(df.Anno)) ### True
# #print(df[['Anno','Mese']])
#
## print(df_Categoria.head())
# print(df_Categoria.iloc[0]) #stampa prima riga
# print(type(df_Categoria.iloc[0])) #Series type
#
# print(df_Categoria.iloc[-1]) #stampa ultima riga
# print(df_Categoria.iloc[[0,1,2,3]]) #stampa prime 4 righe
# print(df_Categoria.iloc[[0,1,2,3]].equals(df_Categoria.iloc[0:4])) #True
# print(df_Categoria[:].head()) #stampa tutte le righe (fermati alle prime 5)
# print(df_Categoria.iloc[0, 2]) #gennaio
# print(df_Categoria.iloc[0, :2]) #primi due valori escluso l'ultimo
# print(df_Categoria.iloc[0:3, [0,3,5]])
# print(df_Categoria.iloc[:, [0,3,5]].head())
# print(df_Categoria.iloc[:, 2].equals(df_Categoria.Mese)) #True
# print(df_Categoria.iloc[:, 2].equals(df_Categoria['Mese'])) #True
#
# print(df_Categoria.loc['Congrua', 'Voce']) #righe in cui compare congrua
# print(df_Categoria.loc['Congrua', 'Voce']) #righe in cui compare congrua e colonna da visualizzare
#
# print(df_Categoria.loc['Pensioni', 'Voce']) #
# print(df_Categoria.loc['Offerte', ['Anno','Voce']]) #
# print(df_Categoria.loc['Offerte', ['Anno','Voce']]) #
# print(df_Categoria.loc[['Offerte', 'Pensioni'], ['Anno','Voce']]) #
# print(df_Categoria.loc[:, ['Anno','Voce']].sort_index(ascending=True, inplace=False)) # Di tutte le categorie stampa le voci


df_Categoria = df.set_index('Categoria')

table=pd.pivot_table(df_Categoria, index=['Anno','Categoria','Voce'],  columns=['Mese'],  values=['Euro'],fill_value=0, aggfunc='sum')
print(table)










# Commit changes
conn.commit()
# Close our connection
conn.close()