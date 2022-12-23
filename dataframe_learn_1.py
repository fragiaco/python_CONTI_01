import pandas as pd
from openpyxl.workbook import Workbook
import sqlite3

conn = sqlite3.connect('database_conti')
# Create a cursor instance
c = conn.cursor()

query = "SELECT * FROM TABLE_Conti"  # query to collect recors
df = pd.read_sql(query, conn)

#print dataframe
print(df.drop(['ID', 'Mese', 'Entrate_Uscite'], axis=1, ).head())







# Commit changes
conn.commit()
# Close our connection
conn.close()