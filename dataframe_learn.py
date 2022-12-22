import pandas as pd
from openpyxl.workbook import Workbook
import sqlite3

conn = sqlite3.connect('database_conti')
# Create a cursor instance
c = conn.cursor()

query = "SELECT * FROM TABLE_Conti"  # query to collect recors
df = pd.read_sql(query, conn)

#print dataframe
print(df)
#print Index(['ID', 'Anno', 'Mese', 'Entrate_Uscite', 'Categoria', 'Voce', 'Euro']
print(df.columns)

#print all values in these column|s
print(df['Entrate_Uscite'])
print(df['Entrate_Uscite'][0:3]) #first 3 lines

print(df[['Categoria', 'Voce']])

print(df.iloc[2,2]) #print cell row 2 column 2

#crei un foglio excel con solo i dati che ti servono
wanted_values = df[['Categoria', 'Voce']]
stored = wanted_values.to_excel('Stored_learn_pandas.xlsx', index=None)


df.to_excel('Learn_Dataframe.xlsx')

print(df.loc[df['Voce'] == 'fra Giacomo'])
#     ID  Anno     Mese Entrate_Uscite Categoria         Voce   Euro
# 5    7  2022  gennaio        Entrate   Congrua  fra Giacomo  900.0
# 20  25  2022  gennaio        Entrate   Congrua  fra Giacomo -350.0
# 29  40  2022  gennaio        Entrate   Congrua  fra Giacomo   36.0

print(df.loc[(df['Voce'] == 'fra Giacomo') & (df['Mese']=='agosto')])


# Commit changes
conn.commit()
# Close our connection
conn.close()