import pandas as pd
import sqlite3

# Leggo Camerino_2012.xlsx
# Lo trasformo in un dataframe


col_names = ['Anno', 'Mese', 'Categoria', 'Voce', 'Euro']
df_Camerino_2012 = pd.read_excel('Camerino_2012.xlsx', names=col_names)

# Creo la colonna ['Entrate_Uscite'] e scrivo in automatico i valori
df_Camerino_2012['Entrate_Uscite'] = df_Camerino_2012['Categoria'].apply \
        (lambda x: 'Entrate' if x == 'Vendite varie' or
                                x == 'Salute' or
                                x == 'Curia' or
                                x == 'Collette-Chiesa' or
                                x == 'Congrua' or
                                x == 'Interessi' or
                                x == 'Messe celebrate' or
                                x == 'Offerte' or
                                x == 'Pensioni' or
                                x == 'Predicazione' or
                                x == 'Servizi religiosi' or
                                x == 'Stipendi' or
                                x == 'Sussidi' or
                                x == 'Rimbosi' or
                                x == 'Vendite varie' or
                                x == 'Eccedenza Cassa'
        else 'Uscite')

# Assegno alle colonne l'ordine desiderato
df_Camerino_2012 = df_Camerino_2012[['Anno', 'Mese', 'Entrate_Uscite', 'Categoria', 'Voce', 'Euro']]


df_Camerino_2012.to_excel(r'Camerino_2012_modified.xlsx', index=False)

conn = sqlite3.connect('database_conti')

cur = conn.cursor()
try:
    cur.execute('''CREATE TABLE TABLE_Conti(ID integer not null PRIMARY KEY ,
                                            Anno TEXT not null ,
                                            Mese TEXT not null ,
                                            Entrate_Uscite TEXT not null ,
                                            Categoria TEXT not null ,
                                            Voce TEXT not null ,
                                            Euro real not null )''')
except:
    pass

#df_Camerino_2012.to_sql('TABLE_Conti', conn, if_exists='replace', index = False)


print(conn)
print('Sei connesso al database_conti')


conn.commit()
conn.close()