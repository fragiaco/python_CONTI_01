import pandas as pd
import sqlite3

conn = sqlite3.connect('database_conti_esperimento')

cur = conn.cursor()
print(conn)
print('Sei connesso al database_conti')



df_conti_camerino_modified = pd.read_excel('Camerino_2012_modified.xlsx')



df_conti_camerino_modified.to_sql('TABLE_Conti',conn,if_exists='replace',index=False)


r_df = pd.read_sql("select * from TABLE_Conti",conn)
print(r_df)

conn.commit()
conn.close()