' Document Name:        VC_Open.vbs
' Last Modified:	09/11/2017
' Purpose:		Open Database connection to Firebird

Option Explicit

Dim objectDatabase, connectionString					' Connection to Firebird Database

Sub Main()								' Main Subroutine
	Set objectDatabase				= CreateObject( "ADODB.Connection" )
	
									' VeraCore ShippingSync DataBase
	connectionString				= "DSN=Firebird;Uid=SYSDBA;Pwd=masterkey;"
	objectDatabase.Open connectionString
	ScriptDataManager.StoredData("VC_CONNECTION")	= objectDatabase
	
									' Clean Up
	Set objectDatabase				= Nothing	' Clear Variables
End Sub
' ********************************************************************* '