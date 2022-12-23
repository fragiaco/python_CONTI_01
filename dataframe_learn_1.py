import pandas as pd
from openpyxl.workbook import Workbook
import sqlite3

conn = sqlite3.connect('database_conti')
# Create a cursor instance
c = conn.cursor()

query = "SELECT * FROM TABLE_Conti"  # query to collect recors
df = pd.read_sql(query, conn)

print('#####')
print(type(df))
print('#####')
print(df.head())
print('#####')
print(df.info())
print('#####')
print(len(df))
print('#####')
print(df.shape)
print(df.index)
print(df.columns)
print('#####')

print((df['Anno']).equals(df.Anno)) ### True
#print(df[['Anno','Mese']])











# Commit changes
conn.commit()
# Close our connection
conn.close()