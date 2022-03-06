from lib2to3.pgen2 import driver
from multiprocessing import connection
from matplotlib.pyplot import title
import pyodbc
import pandas as pd
import os
from datetime import datetime
from plyer import notification

# create SQL Connection
# print(pyodbc.drivers())
connection=pyodbc.connect(driver='{ODBC Driver 17 for SQL Server}', host='localhost', database="AdventureWorks2014", trusted_connection='yes')

# SQL command to read data
sqlQuery="SELECT *   FROM AdventureWorks2014.Sales.SalesOrderHeader"

# Getting the data from sql into pandas dataframe
df=pd.read_sql(sql=sqlQuery,con=connection)

# Export the data on the desktop, puede usar: os.environ["userprofile"] + 
df.to_csv("..\\etl\\files\\" + "SQL_OrderData_" +datetime.now().strftime("%d-%b-%Y %H%M%S") + ".csv", index=False)

# Display notification to user
notification.notify(title="Report Status!!!", 
message=f"Sales data  has been succesfully saved into Excel.\
    \nTotal Rows: {df.shape[0]}\nTotal Columns:  {df.shape[1]}",
    timeout=10)