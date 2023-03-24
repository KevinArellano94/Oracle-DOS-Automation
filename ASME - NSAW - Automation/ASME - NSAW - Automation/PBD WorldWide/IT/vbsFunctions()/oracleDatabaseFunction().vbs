Option Explicit

Dim oracleConstructionString, oracleObjectDatabase, oracleRecordSet, oracleSQLString
Dim strFirst, strLast, strMSN

Dim oracleProvider, oracleDataSource, oracleDatabaseName, oracleDatabasePassword, oracleLibrary, oracleDatabaseTable

oracleProvider						= "MSDASQL"
oracleDataSource					= "IBMDA720train"
oracleDatabaseName					= "UPSCSTST"
oracleDatabasePassword				= "CSTS031808"
oracleLibrary						= "pbdprddta"
oracleDatabaseTable				= "UPS100F"

Sub Main()
	
	Call oracleDatabaseFunction()
	
End Sub

'***********************************************************************************************************************************************'

Function oracleDatabaseFunction()

	oracleConstructionString = "Provider=" & oracleProvider & ";" & "DSN=" & oracleDataSource & ";" & "uid=" & oracleDatabaseName & ";" & "pwd=" & oracleDatabasePassword & ""
	oracleSQLString = "SELECT (MSN) FROM " & oracleLibrary & "." & oracleDatabaseTable & ""
	'oracleSQLString = "SELECT COUNT(*) MSN FROM " & oracleLibrary & "." & oracleDatabaseTable & ""
	
	Set oracleObjectDatabase = CreateObject("ADODB.Connection")
	oracleObjectDatabase.ConnectionString = oracleConstructionString
	oracleObjectDatabase.Open
	
	Set oracleRecordSet = CreateObject("ADODB.Recordset")
	Set oracleRecordSet.ActiveConnection = oracleObjectDatabase
	
	oracleRecordSet.Source = oracleSQLString
	oracleRecordSet.Open
	
	Do Until oracleRecordSet.EOF
	  oracleSQLString = oracleRecordSet.Fields("MSN").Value
	  msgbox("Oracle MSN: " & oracleSQLString)
	  oracleRecordSet.MoveNext
	Loop
	
	oracleRecordSet.Close
	oracleObjectDatabase.Close

End Function

'***********************************************************************************************************************************************'