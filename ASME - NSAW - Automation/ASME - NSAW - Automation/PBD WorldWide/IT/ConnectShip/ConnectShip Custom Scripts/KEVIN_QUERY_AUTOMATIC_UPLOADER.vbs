' OPTION EXPLICIT REQUIRES FOR EVERY OBJECT TO BE DECLARED '
Option Explicit

' CONNECTION CREDENTIALS FOR DATABASE '
Dim sql, sqlobjDB, sqlconstring, SQL_CONNECTION
Set sqlobjDB = CreateObject("ADODB.Connection")

sqlconstring = "Provider=SQLOLEDB.1;Data Source=qts-csdev;Initial Catalog=ORACLE_SIM;user id='connectship_test'; password='connectship'"
sqlobjDB.Open sqlconstring

Set SQL_CONNECTION = sqlobjDB

' OBJECTS THAT STORES INFORMATION TO VERIFY IF EACH QUERY SHOULD GET INSERTED OR NOT '
Dim Us, debugOut, debugOutText, stringLine, stringLineMSN

' OBJECTS THAT STORES INFORMATION TO GET DOCUMENTED IN LOG FILE '
Dim DataAlreadyExists

' ALLOWS TO SEARCH ANY COMPUTER FOR THE FILES '
Us = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
Set debugOut = CreateObject("Scripting.FileSystemObject")

' IF ScriptInsertFile.txt EXISTS '
If (debugOut.FileExists(Us & "\Desktop\ScriptInsertFile.txt")) Then
	
	' OPEN AS IN READABLE FORMAT '
	Set debugOutText = debugOut.OpenTextFile(Us & "\Desktop\ScriptInsertFile.txt", 1, 2)
	
	' REPEAT PROCESS UNTIL NO MORE DATA EXISTS '
	Do UNTIL debugOutText.AtEndOfStream
		
		' SAVE THE MSN NUMBER '
		stringLineMSN = debugOutText.ReadLine
		' SAVE THE INSERT STATEMENT '
		stringLine = debugOutText.ReadLine
		
		' msgbox("MSN Number: " + stringLineMSN) '
		
		' OBJECTS THAT STORES INFORMATION TO CHECK IF THE DATABASE DATA '
		Dim checkSQL, recordStored
		
		' SELECT FROM TABLE WHERE MSN EQUALS FROM THE ONE READ IN THE FILE '
		checkSQL = "SELECT MSN FROM ORACLE_SIM.dbo.UPS100F WHERE MSN = '" & stringLineMSN & "'"
		
		' CHECK THE DATABASE '
		Set recordStored = SQL_CONNECTION.Execute (checkSQL)
		
		' MsgBox(checkSQL) '
		
		' IF THE RECORD IS FOUND '
		If (recordStored.EOF = False) Then
			' MsgBox("The data already exists for the MSN Number: " + stringLineMSN) '
			
			' IF THE RECORD IS FOUND '
			If InStr(stringLineMSN, stringLineMSN) <> 0 Then
				' MsgBox("It was located." & stringLineMSN) '
	
				' ALLOWS TO SEARCH ANY COMPUTER FOR THE FILES '
				Dim UsDataAlreadyExists, debugOutDataAlreadyExists, debugOutTextDataAlreadyExists
				
				' ALLOWS TO SEARCH ANY COMPUTER FOR THE FILES '
				UsDataAlreadyExists = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
				Set debugOutDataAlreadyExists = CreateObject("Scripting.FileSystemObject")
				
				' IF LogFile.txt EXISTS '
				If (debugOutDataAlreadyExists.FileExists(UsDataAlreadyExists & "\Desktop\LogFile.txt")) Then
					' OPEN AS IN WRITEABLE FORMAT '
					Set debugOutTextDataAlreadyExists = debugOutDataAlreadyExists.OpenTextFile(UsDataAlreadyExists & "\Desktop\LogFile.txt", 8, 2)
				Else
					' CREATE A FILE  AS IN WRITEABLE FORMAT '
					Set debugOutTextDataAlreadyExists = debugOutDataAlreadyExists.CreateTextFile(UsDataAlreadyExists & "\Desktop\LogFile.txt", True)
				End If
				
				debugOutTextDataAlreadyExists.Write Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()) & " | " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now()) & " | " & stringLineMSN & " | " & "The data already exists."
				debugOutTextDataAlreadyExists.WriteLine
				debugOutTextDataAlreadyExists.Close
				
				' CLEANUP OBJECTS '
				Set debugOutTextDataAlreadyExists = Nothing
				Set debugOutDataAlreadyExists = Nothing
				
			Else
				' IF THE RECORD IS NOT FOUND '
				' MsgBox("It was not located.") '
			End If
		
		' IF THE RECORD DOES NOT EXIST '
		ElseIf (recordStored.EOF = True) Then
			
			'******************************************************************************************************'
			Dim recordStoredRetry, loopCount
			loopCount = 1
			Do Until loopCount = 5
				' MsgBox(stringLine) '
				' MsgBox("Attempting insert " & loopCount) '
				
				' EXECUTING THE INSERT QUERY '
				SQL_CONNECTION.Execute stringLine
				
				' OBJECTS THAT STORES INFORMATION OF THE ERROR '
				Dim dbError
				dbError = ""
				If SQL_CONNECTION.Errors.Count > 0 Then
					Dim intcount
					Dim dbCommError
					For intCount = 0 To SQL_CONNECTION.Errors.Count - 1
						Set dbCommError = SQL_CONNECTION.Errors.Item(intCount)
						' msgBox(dbcommerror.Description) '
						dbError = dbError + dbCommError.Description + " ** "
					Next
				End If
				
				' OBJECTS THAT STORES INFORMATION TO CHECK IF THE DATABASE DATA '
				Dim checkSQLRetry
				MsgBox("Delete the database information")
				
				' SELECT FROM TABLE WHERE MSN EQUALS FROM THE ONE READ IN THE FILE '
				checkSQLRetry = "SELECT * FROM ORACLE_SIM.dbo.UPS100F WHERE MSN = '" & stringLineMSN & "'"
				
				' CHECK THE DATABASE '
				Set recordStoredRetry = SQL_CONNECTION.Execute (checkSQLRetry)
				
				' IF THE RECORD IS NOT FOUND '
				If (recordStoredRetry.eof) Then
					' MsgBox("Uh oh. SQL Insert did NOT commit") '
				Else
					' MsgBox("The data was inserted") '
					' ALLOWS TO SEARCH ANY COMPUTER FOR THE FILES '
					Dim UsInsert, debugOutInsert, debugOutTextInsert
					
					' ALLOWS TO SEARCH ANY COMPUTER FOR THE FILES '
					UsInsert = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
					Set debugOutInsert = CreateObject("Scripting.FileSystemObject")
					
					' IF LogFile.txt EXISTS '
					If (debugOutInsert.FileExists(UsInsert & "\Desktop\LogFile.txt")) Then
						' OPEN AS IN WRITEABLE FORMAT '
						Set debugOutTextInsert = debugOutInsert.OpenTextFile(UsInsert & "\Desktop\LogFile.txt", 8, 2)
					Else
						' CREATE A FILE  AS IN WRITEABLE FORMAT '
						Set debugOutTextInsert = debugOutInsert.CreateTextFile(UsInsert & "\Desktop\LogFile.txt", True)
					End If
					
					debugOutTextInsert.Write Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()) & " | " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now()) & " | " & "Nothing was found for importing."
					debugOutTextInsert.WriteLine
					debugOutTextInsert.Close
					
					' CLEANUP OBJECTS '
					Set debugOutTextInsert = Nothing
					Set debugOutInsert = Nothing
				Exit Do
				End If
			
			loopCount = loopCount + 1
			If loopCount = 5 Then
				' MsgBox("Well, crap. This failed too.") '
				
				' ALLOWS TO SEARCH ANY COMPUTER FOR THE FILES '
				Dim UsBackUp, debugOutBackUp, debugOutTextBackUp
				
				' ALLOWS TO SEARCH ANY COMPUTER FOR THE FILES '
				UsBackUp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
				Set debugOutBackUp = CreateObject("Scripting.FileSystemObject")
				
				' IF ScriptInsertFileBackUp.txt EXISTS '
				If (debugOutBackUp.FileExists(UsBackUp & "\Desktop\ScriptInsertFileBackUp.txt")) Then
					' OPEN AS IN WRITEABLE FORMAT '
					Set debugOutTextBackUp = debugOutBackUp.OpenTextFile(UsBackUp & "\Desktop\ScriptInsertFileBackUp.txt", 8, 2)
				Else
					' CREATE A FILE  AS IN WRITEABLE FORMAT '
					Set debugOutTextBackUp = debugOutBackUp.CreateTextFile(UsBackUp & "\Desktop\ScriptInsertFileBackUp.txt", True)
				End If
				debugOutTextBackUp.Write stringLineMSN
				debugOutTextBackUp.WriteLine
				debugOutTextBackUp.Write stringLine & ";"
				debugOutTextBackUp.WriteLine
				debugOutTextBackUp.Close
				
				' CLEANUP OBJECTS '
				Set debugOutTextBackUp = Nothing
				Set debugOutBackUp = Nothing
				
				' ALLOWS TO SEARCH ANY COMPUTER FOR THE FILES '
				Dim UsInsertUnsuccessful, debugOutInsertUnsuccessful, debugOutTextInsertUnsuccessful
				
				' ALLOWS TO SEARCH ANY COMPUTER FOR THE FILES '
				UsInsertUnsuccessful = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
				Set debugOutInsertUnsuccessful = CreateObject("Scripting.FileSystemObject")
				
				' IF LogFile.txt EXISTS '
				If (debugOutInsertUnsuccessful.FileExists(UsInsertUnsuccessful & "\Desktop\LogFile.txt")) Then
					' OPEN AS IN WRITEABLE FORMAT '
					Set debugOutTextInsertUnsuccessful = debugOutInsertUnsuccessful.OpenTextFile(UsInsertUnsuccessful & "\Desktop\LogFile.txt", 8, 2)
				Else
					' CREATE A FILE  AS IN WRITEABLE FORMAT '
					Set debugOutTextInsertUnsuccessful = debugOutInsertUnsuccessful.CreateTextFile(UsInsertUnsuccessful & "\Desktop\LogFile.txt", True)
				End If
				
				debugOutTextInsertUnsuccessful.Write Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()) & " | " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now()) & " | " & stringLineMSN & " | " & "Could not insert the data."
				debugOutTextInsertUnsuccessful.WriteLine
				debugOutTextInsertUnsuccessful.Close
				
				' CLEANUP OBJECTS '
				Set debugOutTextInsertUnsuccessful = Nothing
				Set debugOutInsertUnsuccessful = Nothing
				
				Exit Do
			End If
			Loop
			'******************************************************************************************************'
		Else
			' MsgBox("SQL Insert successful") '
			SQL_CONNECTION.Execute stringLine
		End If
	Loop
	
	' CLOSE THE FILE '
	debugOutText.Close
Else
	' MsgBox ("Nothing was found for importing") '
	
	' ALLOWS TO SEARCH ANY COMPUTER FOR THE FILES '
	Dim UsLogFile, debugOutLogFile, debugOutTextLogFile
	
	' ALLOWS TO SEARCH ANY COMPUTER FOR THE FILES '
	UsLogFile = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
	Set debugOutLogFile = CreateObject("Scripting.FileSystemObject")
	
	' IF LogFile.txt EXISTS '
	If (debugOutLogFile.FileExists(UsLogFile & "\Desktop\LogFile.txt")) Then
		' OPEN AS IN WRITEABLE FORMAT '
		Set debugOutTextLogFile = debugOutLogFile.OpenTextFile(UsLogFile & "\Desktop\LogFile.txt", 8, 2)
	Else
		' CREATE A FILE  AS IN WRITEABLE FORMAT '
		Set debugOutTextLogFile = debugOutLogFile.CreateTextFile(UsLogFile & "\Desktop\LogFile.txt", True)
	End If
	
	debugOutTextLogFile.Write Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()) & " | " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now()) & " | " & "Nothing was found for importing."
	debugOutTextLogFile.WriteLine
	debugOutTextLogFile.Close
	
	' CLEANUP OBJECTS '
	Set debugOutTextLogFile = Nothing
	Set debugOutLogFile = Nothing
End If

' OBJECTS THAT STORES INFORMATION FOR THE DELETION OF THE ScriptInsertFile.txt FILE AFTER INSERT '
Dim UsDeleteScriptInsertFile, debugOutDeleteScriptInsertFile

' ALLOWS TO SEARCH ANY COMPUTER FOR THE FILES '
UsDeleteScriptInsertFile = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
Set debugOutDeleteScriptInsertFile = CreateObject("Scripting.FileSystemObject")

' IF ScriptInsertFile.txt EXISTS '
If (debugOutDeleteScriptInsertFile.FileExists(UsDeleteScriptInsertFile & "\Desktop\ScriptInsertFile.txt")) Then
	debugOutDeleteScriptInsertFile.DeleteFile(UsDeleteScriptInsertFile & "\Desktop\ScriptInsertFile.txt")
Else
	' NOTHING '
End If

WScript.Quit