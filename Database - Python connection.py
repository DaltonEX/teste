import pyodbc

#database credentials
server='fcatnwsql1.jdnet.deere.com\INST1'
database = 'Database_AppPortaria'
username='João Pedro Pereira de Freitas (GD5CX71)'
password = 'rPDo6lm1K%'

#database connection
connection = pyodbc.connect(f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}")

#retrieve data from table 
cursor = connection.cursor()
cursor.execute(f"INSERT INTO [JDNET\GD5CX71].Portaria (first_name, last_name, pc_code_number) VALUES ('{Nome}','{Sobrenome}', '{barcode}')")

# stores the results in a variable 
# #results = cursor.fetchall()

# Commit and close the transaction
connection.commit()
connection.close()











