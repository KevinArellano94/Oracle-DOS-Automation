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

Dim Commodity_Contents,				sqlCommodity_Contents,				resultsqlCommodity_Contents
Dim Commodity_ContentsCount,		sqlCommodity_ContentsCount,			resultsqlCommodity_ContentsCount
Dim Commodity_Contents_Archived,	sqlCommodity_Contents_Archived,		resultsqlCommodity_Contents_Archived
Dim Hazmat_Contents,				sqlHazmat_Contents,					resultsqlHazmat_Contents
Dim Hazmat_ContentsCount,			sqlHazmat_ContentsCount,			resultsqlHazmat_ContentsCount
Dim Hazmat_Contents_Archived,		sqlHazmat_Contents_Archived,		resultsqlHazmat_Contents_Archived
Dim Package_Contents,				sqlPackage_Contents,				resultsqlPackage_Contents
Dim Package_ContentsCount,			sqlPackage_ContentsCount,			resultsqlPackage_ContentsCount
Dim Package_Contents_Archived,		sqlPackage_Contents_Archived,		resultsqlPackage_Contents_Archived
Dim Package_Lists,					sqlPackage_Lists,					resultsqlPackage_Lists
Dim Package_ListsCount,				sqlPackage_ListsCount,				resultsqlPackage_ListsCount
Dim Package_Lists_Archived,			sqlPackage_Lists_Archived,			resultsqlPackage_Lists_Archived
Dim Packages,						sqlPackages,						resultsqlPackages
Dim PackagesCount,					sqlPackagesCount,					resultsqlPackagesCount
Dim Packages_Archived,				sqlPackages_Archived,				resultsqlPackages_Archived
Dim PackagesPackageListId,			sqlPackagesPackageListId,			resultsqlPackagesPackageListId
Dim PackagesPackageListId_Archived,	sqlPackagesPackageListId_Archived,	resultsqlPackagesPackageListId_Archived
Dim Packages2,						sqlPackages2,						resultsqlPackages2
Dim Packages2Count,					sqlPackages2Count,					resultsqlPackages2Count
Dim Packages2_Archived,				sqlPackages2_Archived,				resultsqlPackages2_Archived
Dim MSNPackagesCounter,				sqlMSNPackagesCounter,				resultsqlMSNPackagesCounter

' OBJECTS THAT INSERT|DELETE THE DATA '
Dim sqlSelect,						sqlSQLSelect,						resultsqlSelect
Dim sqlInsert,						sqlSQLInsert,						resultsqlInsert
Dim sqlDelete,						sqlSQLDelete,						resultsqlDelete

Commodity_Contents						= "commodity_contents"
Commodity_Contents_Archived				= "commodity_contents_Archived"
Hazmat_Contents							= "hazmat_contents"
Hazmat_Contents_Archived				= "hazmat_contents_Archived"
Package_Contents						= "package_contents"
Package_Contents_Archived				= "package_contents_Archived"
Package_Lists							= "package_lists"
Package_Lists_Archived					= "package_lists_Archived"
Packages								= "packages"
Packages_Archived						= "packages_Archived"
Packages2								= "packages2"
Packages2_Archived						= "packages2_Archived"

sqlProvider								= "SQLOLEDB.1"
sqlDataSource							= "qts-csdev"
sqlDatabaseName							= "connectship_test"
' COMMODITY_CONTENTS SET '
sqlDatabasePassword						= "connectship"

sqlConstructionString = "Provider = " & sqlProvider & ";" & "Data Source = " & sqlDataSource & ";" & "Initial Catalog=" & sqlDatabaseName & ";" & "User Id= " & sqlDatabaseName & ";" & "password = " & sqlDatabasePassword & ""

sqlObjectDatabase.Open sqlConstructionString

Set sqlConnection = sqlObjectDatabase

Call databaseConnectionSubRoutine()

' OBJECTS THAT CHECK THE DATA '
Dim sqlMainDatabaseBefore, sqlBackUpDatabaseBefore
Dim sqlBackUpDatabaseAfter

Dim resultsMainDatabaseBefore, resultsBackUpDatabaseDatabaseBefore
Dim resultsMainDatabaseAfter, resultsBackUpDatabaseDatabaseAfter

' OBJECTS THAT STORES INFORMATION IN A LOG FILE '
Dim UsLogFile, debugOutLogFile, debugOutTextLogFile

' OBJECTS THAT STORES BEFORE AND AFTER ROW COUNT OF BOTH TABLES '
Dim sqlDifferenceBefore
Dim sqlDifferenceAfter

Call differenceProductsDeclaration()

Do Until (resultsqlMSNPackagesCounter(0).value) = (resultsqlPackages(0).value)
	Call differenceProductsDeclaration()
	Call differenceProductsAlgorithm()
	If ISNULL((resultsqlPackages(0).value)) Then
		WScript.Quit
	End If
	Call databaseDataInsertionSubRoutine()
	
	If (resultsqlMSNPackagesCounter(0).value) = (resultsqlPackages(0).value) Then
		Exit Do
	End If
Loop

msgbox("Insertion Completed.")
WScript.Quit

'****************************************************************************************************************************************************************************************************'

Sub differenceProductsDeclaration()
	
	' SELECT NUMBER OF ROWS FOR CHECKING '
	sqlCommodity_Contents				= "SELECT MIN(MSN) FROM [" & sqlDatabaseName & "].[dbo].[" & Commodity_Contents & "]"
	sqlCommodity_ContentsCount			= "SELECT COUNT(*) FROM [" & sqlDatabaseName & "].[dbo].[" & Commodity_Contents & "]"
	sqlCommodity_Contents_Archived		= "SELECT MAX(MSN) FROM [" & sqlDatabaseName & "].[dbo].[" & Commodity_Contents_Archived & "]"
	sqlHazmat_Contents					= "SELECT MIN(MSN) FROM [" & sqlDatabaseName & "].[dbo].[" & Hazmat_Contents & "]"
	sqlHazmat_ContentsCount				= "SELECT COUNT(*) FROM [" & sqlDatabaseName & "].[dbo].[" & Hazmat_Contents & "]"
	sqlHazmat_Contents_Archived			= "SELECT MAX(MSN) FROM [" & sqlDatabaseName & "].[dbo].[" & Hazmat_Contents_Archived & "]"
	sqlPackage_Contents					= "SELECT MIN(MSN) FROM [" & sqlDatabaseName & "].[dbo].[" & Package_Contents & "]"
	sqlPackage_ContentsCount			= "SELECT COUNT(*) FROM [" & sqlDatabaseName & "].[dbo].[" & Package_Contents & "]"
	sqlPackage_Contents_Archived		= "SELECT MAX(MSN) FROM [" & sqlDatabaseName & "].[dbo].[" & Package_Contents_Archived & "]"
	sqlPackage_Lists					= "SELECT MIN(packagelist_id) FROM [" & sqlDatabaseName & "].[dbo].[" & Package_Lists & "]"
	sqlPackage_ListsCount				= "SELECT COUNT(*) FROM [" & sqlDatabaseName & "].[dbo].[" & Package_Lists & "]"
	sqlPackage_Lists_Archived			= "SELECT MAX(packagelist_id) FROM [" & sqlDatabaseName & "].[dbo].[" & Package_Lists_Archived & "]"
	sqlPackages							= "SELECT MIN(MSN) FROM [" & sqlDatabaseName & "].[dbo].[" & Packages & "] WHERE reference_11 < convert(date, GETDATE() - (365 * 2))"
	sqlPackagesCount					= "SELECT COUNT(*) FROM [" & sqlDatabaseName & "].[dbo].[" & Packages & "]"
	sqlPackages_Archived				= "SELECT MAX(MSN) FROM [" & sqlDatabaseName & "].[dbo].[" & Packages_Archived & "]"
	sqlPackagesPackageListId			= "SELECT MIN(packagelist_id) FROM [" & sqlDatabaseName & "].[dbo].[" & Packages & "]"
	sqlPackagesPackageListId_Archived	= "SELECT MAX(packagelist_id) FROM [" & sqlDatabaseName & "].[dbo].[" & Packages_Archived & "]"
	sqlPackages2						= "SELECT MIN(MSN) FROM [" & sqlDatabaseName & "].[dbo].[" & Packages2 & "]"
	sqlPackages2Count					= "SELECT COUNT(*) FROM [" & sqlDatabaseName & "].[dbo].[" & Packages2 & "]"
	sqlPackages2_Archived				= "SELECT MAX(MSN) FROM [" & sqlDatabaseName & "].[dbo].[" & Packages2_Archived & "]"
	sqlMSNPackagesCounter				= "SELECT MAX(MSN) FROM [" & sqlDatabaseName & "].[dbo].[" & Packages & "] WHERE reference_11 < convert(date, GETDATE() - (365 * 2))"
	' EXECUTING THE SELECT QUERY '
	sqlConnection.Execute (sqlCommodity_Contents)
	sqlConnection.Execute (sqlCommodity_ContentsCount)
	sqlConnection.Execute (sqlCommodity_Contents_Archived)
	sqlConnection.Execute (sqlHazmat_Contents)
	sqlConnection.Execute (sqlHazmat_ContentsCount)
	sqlConnection.Execute (sqlHazmat_Contents_Archived)
	sqlConnection.Execute (sqlPackage_Contents)
	sqlConnection.Execute (sqlPackage_ContentsCount)
	sqlConnection.Execute (sqlPackage_Contents_Archived)
	sqlConnection.Execute (sqlPackage_Lists)
	sqlConnection.Execute (sqlPackage_ListsCount)
	sqlConnection.Execute (sqlPackage_Lists_Archived)
	sqlConnection.Execute (sqlPackages)
	sqlConnection.Execute (sqlPackagesCount)
	sqlConnection.Execute (sqlPackages_Archived)
	sqlConnection.Execute (sqlPackagesPackageListId)
	sqlConnection.Execute (sqlPackagesPackageListId_Archived)
	sqlConnection.Execute (sqlPackages2)
	sqlConnection.Execute (sqlPackages2Count)
	sqlConnection.Execute (sqlPackages2_Archived)
	sqlConnection.Execute (sqlMSNPackagesCounter)
	' SET THE VARIABLES FOR CHECKING '
	Set resultsqlCommodity_Contents				= sqlConnection.Execute (sqlCommodity_Contents)
	Set resultsqlCommodity_ContentsCount		= sqlConnection.Execute (sqlCommodity_ContentsCount)
	Set resultsqlCommodity_Contents_Archived	= sqlConnection.Execute (sqlCommodity_Contents_Archived)
	Set resultsqlHazmat_Contents				= sqlConnection.Execute (sqlHazmat_Contents)
	Set resultsqlHazmat_ContentsCount			= sqlConnection.Execute (sqlHazmat_ContentsCount)
	Set resultsqlHazmat_Contents_Archived		= sqlConnection.Execute (sqlHazmat_Contents_Archived)
	Set resultsqlPackage_Contents				= sqlConnection.Execute (sqlPackage_Contents)
	Set resultsqlPackage_ContentsCount			= sqlConnection.Execute (sqlPackage_ContentsCount)
	Set resultsqlPackage_Contents_Archived		= sqlConnection.Execute (sqlPackage_Contents_Archived)
	Set resultsqlPackage_Lists					= sqlConnection.Execute (sqlPackage_Lists)
	Set resultsqlPackage_ListsCount				= sqlConnection.Execute (sqlPackage_ListsCount)
	Set resultsqlPackage_Lists_Archived			= sqlConnection.Execute (sqlPackage_Lists_Archived)
	Set resultsqlPackages						= sqlConnection.Execute (sqlPackages)
	Set resultsqlPackagesCount					= sqlConnection.Execute (sqlPackagesCount)
	Set resultsqlPackages_Archived				= sqlConnection.Execute (sqlPackages_Archived)
	Set resultsqlPackagesPackageListId			= sqlConnection.Execute (sqlPackagesPackageListId)
	Set resultsqlPackagesPackageListId_Archived	= sqlConnection.Execute (sqlPackagesPackageListId_Archived)
	Set resultsqlPackages2						= sqlConnection.Execute (sqlPackages2)
	Set resultsqlPackages2Count					= sqlConnection.Execute (sqlPackages2Count)
	Set resultsqlPackages2_Archived				= sqlConnection.Execute (sqlPackages2_Archived)
	Set resultsqlMSNPackagesCounter				= sqlConnection.Execute (sqlMSNPackagesCounter)
End Sub

'****************************************************************************************************************************************************************************************************'

Sub differenceProductsAlgorithm()
	
	' CHECKING PACKAGES '
	If ((resultsqlPackages(0).value) = (resultsqlPackages_Archived(0).value)) Then
		
	Else
		If ISNULL((resultsqlPackages(0).value)) Then
			UsLogFile = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
			Set debugOutLogFile = CreateObject("Scripting.FileSystemObject")

			' IF ArchivalLogFile.txt EXISTS '
			If (debugOutLogFile.FileExists(UsLogFile & "\Desktop\ArchivalLogFile.txt")) Then
				' OPEN AS IN WRITEABLE FORMAT '
				Set debugOutTextLogFile = debugOutLogFile.OpenTextFile(UsLogFile & "\Desktop\ArchivalLogFile.txt", 8, 2)
			Else
				' CREATE A FILE  AS IN WRITEABLE FORMAT '
				Set debugOutTextLogFile = debugOutLogFile.CreateTextFile(UsLogFile & "\Desktop\ArchivalLogFile.txt", True)
			End If

			debugOutTextLogFile.Write Pd(Month(date()),2) & "/" & Pd(DAY(date()),2) & "/" & YEAR(Date()) & " | " & Pd(Hour(Time()),2) & ":" & Pd(Minute(Time()),2) & ":" & Pd(Second(Time()),2) & " | " & "No data was found for archiving."
			debugOutTextLogFile.WriteLine
			debugOutTextLogFile.Close

			' CLEANUP OBJECTS '
			Set debugOutTextLogFile = Nothing
			Set debugOutLogFile = Nothing
			Exit Sub
		Else
			sqlInsert = "INSERT INTO [" & sqlDatabaseName & "].[dbo].[" & Packages_Archived & "] SELECT * FROM [" & sqlDatabaseName & "].[dbo].[" & Packages & "] WHERE MSN = " & (resultsqlPackages(0).value) & ";"
			sqlConnection.Execute (sqlInsert)
		End If
	End If
	' CHECKING COMMODITY_CONTENTS '
	If ((resultsqlPackages(0).value) = (resultsqlCommodity_Contents(0).value)) Then
		If ((resultsqlCommodity_Contents(0).value) = (resultsqlCommodity_Contents_Archived(0).value)) Then
			
		Else
			sqlInsert = "INSERT INTO [" & sqlDatabaseName & "].[dbo].[" & Commodity_Contents_Archived & "] SELECT * FROM [" & sqlDatabaseName & "].[dbo].[" & Commodity_Contents & "] WHERE MSN = " & (resultsqlPackages(0).value) & ";"
			sqlConnection.Execute (sqlInsert)
		End If
	Else
		
	End If
	' CHECKING HAZMAT_CONTENTS '
	If ((resultsqlPackages(0).value) = (resultsqlHazmat_Contents(0).value)) Then
		If ((resultsqlHazmat_Contents(0).value) = (resultsqlHazmat_Contents_Archived(0).value)) Then
			
		Else
			sqlInsert = "INSERT INTO [" & sqlDatabaseName & "].[dbo].[" & Hazmat_Contents_Archived & "] SELECT * FROM [" & sqlDatabaseName & "].[dbo].[" & Hazmat_Contents & "] WHERE MSN = " & (resultsqlPackages(0).value) & ";"
			sqlConnection.Execute (sqlInsert)
		End If
	Else
		
	End If
	' CHECKING PACKAGE_LISTS '
	If ((resultsqlPackagesPackageListId(0).value) = (resultsqlPackage_Lists(0).value)) Then
		If ((resultsqlPackage_Lists(0).value) = (resultsqlPackage_Lists_Archived(0).value)) Then
			
		Else
			sqlInsert = "INSERT INTO [" & sqlDatabaseName & "].[dbo].[" & Package_Lists_Archived & "] SELECT * FROM [" & sqlDatabaseName & "].[dbo].[" & Package_Lists_Archived & "] WHERE packagelist_id = " & (resultsqlPackagesPackageListId(0).value) & ";"
			sqlConnection.Execute (sqlInsert)
		End If
	Else
		
	End If
	' CHECKING PACKAGE2 '
	If ((resultsqlPackages(0).value) = (resultsqlPackages2(0).value)) Then
		If ((resultsqlPackages2(0).value) = (resultsqlPackages2_Archived(0).value)) Then
			
		Else
			sqlInsert = "INSERT INTO [" & sqlDatabaseName & "].[dbo].[" & Packages2_Archived & "] SELECT * FROM [" & sqlDatabaseName & "].[dbo].[" & Packages2 & "] WHERE MSN = " & (resultsqlPackages(0).value) & ";"
			sqlConnection.Execute (sqlInsert)
		End If
	Else
		
	End If
	
End Sub

'****************************************************************************************************************************************************************************************************'

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

'****************************************************************************************************************************************************************************************************'

Sub databaseDataInsertionSubRoutine()
	
	' UPDATE DATABASE '
	If (1=1) Then
		Dim results, loopCount
		loopCount = 1
		Do Until loopCount = 5
			'MsgBox("Attempting insert " & loopCount)
			
			' IF SQL INSERT DOES NOT ERROR A CHECK WILL BE MADE '
			' IF CHECK DOES NOT FIND DATA THEN SQL WRITTEN TO DEBUG '
			Dim databaseError
			databaseError = ""
			If sqlConnection.Errors.Count > 0 Then
				Dim intcount
				Dim dbCommError
				For intCount = 0 To sqlConnection.Errors.Count - 1
					Set dbCommError = sqlConnection.Errors.Item(intCount)
					writeDebug dbcommerror.Description
					databaseError = databaseError + dbCommError.Description + " ** "
				Next
				writeDebug "SQL-" & sql
				setError(databaseError)
			End If
			Dim checkSQL
			'MsgBox("The information is being written onto the database.")
			sqlSelect = "SELECT MAX(MSN) FROM [" & sqlDatabaseName & "].[dbo].[" & Packages_Archived & "] WHERE MSN = " & (resultsqlPackages(0).value) & ";"
			sqlConnection.Execute (sqlSelect)
			Set resultsqlSelect = sqlConnection.Execute (sqlSelect)
			If ((resultsqlPackages_Archived(0).value) = (resultsqlSelect(0).value)) Then
''				MsgBox("Uh oh. SQL Insert did NOT commit.")
''				MsgBox((resultsqlPackages_Archived(0).value) & " | " & (resultsqlSelect(0).value))
			Else
''				MsgBox("SQL Insert successful")
				If ISNULL(resultsqlCommodity_Contents(0).value) Then
					
				Else
					sqlDelete = "DELETE FROM [" & sqlDatabaseName & "].[dbo].[" & Commodity_Contents & "] WHERE MSN = " & (resultsqlPackages(0).value) & ";"
					sqlConnection.Execute (sqlDelete)
				End If
				If ISNULL(resultsqlHazmat_Contents(0).value) Then
					
				Else
					sqlDelete = "DELETE FROM [" & sqlDatabaseName & "].[dbo].[" & Hazmat_Contents & "] WHERE MSN = " & (resultsqlPackages(0).value) & ";"
					sqlConnection.Execute (sqlDelete)
				End If
				If ISNULL(resultsqlPackage_Contents(0).value) Then
					
				Else
					sqlDelete = "DELETE FROM [" & sqlDatabaseName & "].[dbo].[" & Package_Contents & "] WHERE MSN = " & (resultsqlPackages(0).value) & ";"
					sqlConnection.Execute (sqlDelete)
				End If
				If ISNULL(resultsqlPackage_Lists(0).value) Then
					
				Else
					' DOES NOT DELETE DUE TO REFERENCE "FOREIGN KEY" CONSTRAINT '
					'sqlDelete = "DELETE FROM [" & sqlDatabaseName & "].[dbo].[" & Package_Lists & "] WHERE packagelist_id = " & (resultsqlPackagesPackageListId(0).value) & ";"
					'sqlConnection.Execute (sqlDelete)
				End If
				If ISNULL(resultsqlPackages2(0).value) Then
					
				Else
					sqlDelete = "DELETE FROM [" & sqlDatabaseName & "].[dbo].[" & Packages2 & "] WHERE MSN = " & (resultsqlPackages(0).value) & ";"
					sqlConnection.Execute (sqlDelete)
				End If
				If ISNULL(resultsqlPackages(0).value) Then
					
				Else
					sqlDelete = "DELETE FROM [" & sqlDatabaseName & "].[dbo].[" & Packages & "] WHERE MSN = " & (resultsqlPackages(0).value) & ""
					sqlConnection.Execute (sqlDelete)
				End If
				' ALLOWS TO SEARCH ANY COMPUTER FOR THE FILES '
				UsLogFile = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
				Set debugOutLogFile = CreateObject("Scripting.FileSystemObject")
				' IF ArchivalLogFile.txt EXISTS '
				If (debugOutLogFile.FileExists(UsLogFile & "\Desktop\ArchivalLogFile.txt")) Then
					' OPEN AS IN WRITEABLE FORMAT '
					Set debugOutTextLogFile = debugOutLogFile.OpenTextFile(UsLogFile & "\Desktop\ArchivalLogFile.txt", 8, 2)
				Else
					' CREATE A FILE  AS IN WRITEABLE FORMAT '
					Set debugOutTextLogFile = debugOutLogFile.CreateTextFile(UsLogFile & "\Desktop\ArchivalLogFile.txt", True)
				End If
				debugOutTextLogFile.Write Pd(Month(date()),2) & "/" & Pd(DAY(date()),2) & "/" & YEAR(Date()) & " | " & Pd(Hour(Time()),2) & ":" & Pd(Minute(Time()),2) & ":" & Pd(Second(Time()),2) & " | " & "Archival completed. | Records left from '" & Packages & "': " & (resultsqlPackagesCount(0).value) & " | Records left from '" & Commodity_Contents & "': " & (resultsqlCommodity_ContentsCount(0).value) & " | Records left from '" & package_lists & "': " & (resultsqlPackage_ListsCount(0).value) & " | Records left from '" & Hazmat_Contents & "': " & (resultsqlHazmat_ContentsCount(0).value) & " | Records left from '" & Package_Contents & "': " & (resultsqlPackage_ContentsCount(0).value) & " | Records left from '" & Packages2 & "': " & (resultsqlPackages2Count(0).value)
				debugOutTextLogFile.WriteLine
				debugOutTextLogFile.Close
				Exit Do
			End If
			
		loopCount = loopCount + 1
			If loopCount = 5 Then
''				MsgBox("Well, crap. All inserts failed.")
				' ALLOWS TO SEARCH ANY COMPUTER FOR THE FILES '
				UsLogFile = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
				Set debugOutLogFile = CreateObject("Scripting.FileSystemObject")

				' IF ArchivalLogFile.txt EXISTS '
				If (debugOutLogFile.FileExists(UsLogFile & "\Desktop\ArchivalLogFile.txt")) Then
					' OPEN AS IN WRITEABLE FORMAT '
					Set debugOutTextLogFile = debugOutLogFile.OpenTextFile(UsLogFile & "\Desktop\ArchivalLogFile.txt", 8, 2)
				Else
					' CREATE A FILE  AS IN WRITEABLE FORMAT '
					Set debugOutTextLogFile = debugOutLogFile.CreateTextFile(UsLogFile & "\Desktop\ArchivalLogFile.txt", True)
				End If

				debugOutTextLogFile.Write Pd(Month(date()),2) & "/" & Pd(DAY(date()),2) & "/" & YEAR(Date()) & " | " & Pd(Hour(Time()),2) & ":" & Pd(Minute(Time()),2) & ":" & Pd(Second(Time()),2) & " | " & "Archival process completly failed. | Records left from '" & Packages & "': " & (resultsqlPackagesCount(0).value) & " | Records left from '" & Commodity_Contents & "': " & (resultsqlCommodity_ContentsCount(0).value) & " | Records left from '" & package_lists & "': " & (resultsqlPackage_ListsCount(0).value) & " | Records left from '" & Hazmat_Contents & "': " & (resultsqlHazmat_ContentsCount(0).value) & " | Records left from '" & Package_Contents & "': " & (resultsqlPackage_ContentsCount(0).value) & " | Records left from '" & Packages2 & "': " & (resultsqlPackages2Count(0).value)
				debugOutTextLogFile.WriteLine
				debugOutTextLogFile.Close

				' CLEANUP OBJECTS '
				Set debugOutTextLogFile = Nothing
				Set debugOutLogFile = Nothing
				Exit Do
			End If
		Loop
	End If
End Sub

'****************************************************************************************************************************************************************************************************'

Function pd(n, totalDigits)
	If totalDigits > len(n) Then
		pd = String(totalDigits-len(n),"0") & n
	Else
		pd = n
	End If
End Function

'****************************************************************************************************************************************************************************************************'