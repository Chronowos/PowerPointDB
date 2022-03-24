import pyodbc
import pandas as pd

#SQL Server Connection
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};PORT=1433;SERVER=SERVERIP;UID=admin;PWD=PW;DATABASE=PowerPointDB')
cursor=conn.cursor()

newResult = pd.read_sql("""SELECT * FROM tbl_Employees WHERE Name = 'Leon Negwer'""",conn)

print(newResult)

#Select EmployeeID where Name is Leon Negwer (which is in column Name)
filtered = list(newResult.loc[newResult['Name'] == 'Leon Negwer', 'EmployeeID'])[0]

#Insert into parameterized
cursor.execute("Insert Into tbl_Employees(EmployeeID, Name, PermissionToEdit) Values (?, ?, ?)", '4', 'Willhelm Siegfried', 'True')

conn.commit()







