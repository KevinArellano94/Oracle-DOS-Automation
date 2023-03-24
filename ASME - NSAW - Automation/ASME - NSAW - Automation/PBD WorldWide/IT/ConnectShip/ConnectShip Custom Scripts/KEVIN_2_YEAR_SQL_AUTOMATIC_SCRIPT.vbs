' OPTION EXPLICIT REQUIRES FOR EVERY OBJECT TO BE DECLARED '
Option Explicit

' CONNECTION CREDENTIALS FOR DATABASE '
Dim sql, sqlobjDB, sqlconstring, SQL_CONNECTION
Set sqlobjDB = CreateObject("ADODB.Connection")

' OBJECTS THAT STORES INFORMATION TO CHECK IF THE DATABASE DATA '
Dim sqlInsertIntoFromMainDatabase, sqlDeleteOldInformation

' OBJECTS THAT STORES INFORMATION TO CHECK THE DATABASE DATA THAT IS 2 YEARS OR OLDER '
Dim theDate, previousTwoYears, theDateFormat

sqlconstring = "Provider=SQLOLEDB.1;Data Source=qts-csdev;Initial Catalog=ORACLE_SIM;user id='connectship_test'; password='connectship'"
sqlobjDB.Open sqlconstring

Set SQL_CONNECTION = sqlobjDB

sqlInsertIntoFromMainDatabase = "INSERT INTO [ORACLE_SIM].[dbo].[UPS100F_Archived] SELECT * FROM [ORACLE_SIM].[dbo].[UPS100F] WHERE SHIPDATE < convert(date, GETDATE() - (365 * 2));"
sqlDeleteOldInformation = "DELETE FROM [ORACLE_SIM].[dbo].[UPS100F] WHERE SHIPDATE <  convert(date, GETDATE() - (365 * 2));"

SQL_CONNECTION.Execute (sqlInsertIntoFromMainDatabase)
SQL_CONNECTION.Execute (sqlDeleteOldInformation)

msgbox("Archival process has been completed." & vbNewLine & sqlInsertIntoFromMainDatabase & vbNewLine & sqlDeleteOldInformation)

WScript.Quit