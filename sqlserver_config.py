import mysql.connector

dbConfig = {
    #'driver': 'SQL Server',

    #'Server': 'DESKTOP-O3G7AKN\MSSQLSERVER2019',
    #'Database': 'mywallet',
    #'username': 'DESKTOP-O3G7AKN\giaco',



    'host': 'localhost',
    'database': 'databaseconti',
    'username': 'root',
    'password': 'Stefano71'
            }
mydb = (mysql.connector.connect(**dbConfig))
print(mydb)