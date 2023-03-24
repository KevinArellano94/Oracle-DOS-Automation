Option Explicit

Dim sql
Dim sqlObjectDatabase
Dim sqlConstructionString ' CREDENTIALS OF THE DATABASE '
Dim sqlConnection
Dim sqlTest
Dim sqlTester
Dim resultsTest
Dim resultsTesting
Set sqlObjectDatabase = CreateObject("ADODB.Connection")

Dim sqlProvider
Dim sqlDataSource
Dim sqlDatabaseName
Dim sqlDatabasePassword

Dim count

Dim Packages,							sqlPackages,								resultsqlPackages

sqlProvider								= "SQLOLEDB.1"
sqlDataSource							= "qts-csdev"
sqlDatabaseName							= "connectship_test"
sqlDatabasePassword					= "connectship"

Packages									= "packages"

sqlConstructionString = "Provider = " & sqlProvider & ";" & "Data Source = " & sqlDataSource & ";" & "Initial Catalog=" & sqlDatabaseName & ";" & "User Id= " & sqlDatabaseName & ";" & "password = " & sqlDatabasePassword & ""

sqlObjectDatabase.Open sqlConstructionString

Set sqlConnection = sqlObjectDatabase

Sub Main()
	
	Call databaseConnectionSubRoutine()
	Call differenceProductsDeclaration()
	
	msgbox("Count: " & count)
	
End Sub

'***********************************************************************************************************************************************'

Sub databaseConnectionSubRoutine()
	
	' TEST FOR CONNECTION ERROR, PERHAPS BAD CREDENTIALS OR SOURCE NOT AVAILABLE '
	sqlTest = "SELECT MIN(MSN) FROM [" & sqlDatabaseName & "].[dbo].[" & Packages & "]"
	sqlConnection.Execute (sqlTest)
	' TEST FOR CONNECTION ERROR, PERHAPS BAD CREDENTIALS OR SOURCE NOT AVAILABLE '
	' CHECK IF CONNECTION EXISTS '
	If (InStr(Err.Description,"Communication link failure") > 0) Then
		writeDebug Err.Description
	Else	
		If (InStr(Err.Description,"Operation is not allowed") > 0) Then
			writeDebug Err.Description
		Else	
			If (InStr(Err.Description,"Closed") > 0) Then
				msgbox("Closed")
			End If
		End If
	End If
	Set resultsTest = sqlObjectDatabase.Execute (sqlTest)
''	msgbox("MSN Tester: " & (resultsTest(0).value))
End Sub

'***********************************************************************************************************************************************'

Sub differenceProductsDeclaration()
	
	' SELECT NUMBER OF ROWS FOR CHECKING '
	sqlPackages = "SELECT COUNT(*) FROM [" & sqlDatabaseName & "].[dbo].[" & Packages & "]"
	' EXECUTING THE SELECT QUERY '
	sqlConnection.Execute (sqlPackages)
	' SET THE VARIABLES FOR CHECKING '
	Set resultsqlPackages = sqlConnection.Execute (sqlPackages)
	
	count = resultsqlPackages(0).value
End Sub

'***********************************************************************************************************************************************'
