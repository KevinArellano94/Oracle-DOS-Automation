Option Explicit

Sub Main

	Dim sql, sqlobjDB, sqlconstring, SQL_CONNECTION
	Set sqlobjDB = CreateObject("ADODB.Connection")
	
	sqlconstring = "Provider=SQLOLEDB.1;Data Source=qts-csdev;Initial Catalog=ORACLE_SIM;user id='connectship_test'; password='connectship'"
	sqlobjDB.Open sqlconstring
	
	Set SQL_CONNECTION = sqlobjDB
	
	Dim US,debugfilepath,debugFso,debugOut,debugOutText,insertQuery,stringLine,stringLineMSN,arrServiceList,stringSearchFor
	
	US = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
	Set debugOut = CreateObject("Scripting.FileSystemObject")
	
	If (debugOut.FileExists(US & "\Desktop\ScriptInsertFile.txt")) Then
		Set debugOutText = debugOut.OpenTextFile(US & "\Desktop\ScriptInsertFile.txt", 1, 2)
		
		Do UNTIL debugOutText.AtEndOfStream
			
			stringLineMSN = debugOutText.ReadLine
			stringLine = debugOutText.ReadLine
			stringSearchFor = stringLineMSN
			
			msgbox("MSN Number: " + stringLineMSN)
			
			Dim chksql, rs
				
			chksql = "SELECT MSN FROM ORACLE_SIM.dbo.UPS100F WHERE MSN = '" & stringLineMSN & "'"
			Set rs = SQL_CONNECTION.Execute (chksql)
			
			MsgBox(chksql)
			
			If (rs.EOF = False) Then
				MsgBox("The data already exists for the MSN Number: " + stringLineMSN)
				If InStr(stringLineMSN, stringLineMSN) <> 0 Then
					MsgBox("It was located.")
					
				Else
					MsgBox("It was not located.")
				End If
			ElseIf (rs.EOF = True) Then
				
				'******************************************************************************************************'
				Dim rs2, loopCount
				loopCount = 1
				Do Until loopCount = 5
					msgbox(stringLine)
					MsgBox("Attempting insert " & loopCount)
					SQL_CONNECTION.Execute stringLine
					' IF SQL INSERT DOES NOT ERROR A CHECK WILL BE MADE '
					' IF CHECK DOES NOT FIND DATA THEN SQL WRITTEN TO DEBUG '
					
					Dim dbError
					dbError = ""
					If SQL_CONNECTION.Errors.Count > 0 Then
						Dim intcount
						Dim dbCommError
						For intCount = 0 To SQL_CONNECTION.Errors.Count - 1
							Set dbCommError = SQL_CONNECTION.Errors.Item(intCount)
							msgBox(dbcommerror.Description)
							dbError = dbError + dbCommError.Description + " ** "
						Next
					End If
					
					Dim chksql2
					MsgBox("Delete the database information")
					chksql2 = "SELECT * FROM ORACLE_SIM.dbo.UPS100F WHERE MSN = '" & stringLineMSN & "'"
					Set rs2 = SQL_CONNECTION.Execute (chksql2)
					
					If (rs2.eof) Then
						MsgBox("Uh oh. SQL Insert did NOT commit")
					Else
						MsgBox("The data was inserted")
					Exit Do
					End If
				
				loopCount = loopCount + 1
				If loopCount = 5 Then
					MsgBox("Well, crap. This failed too.")
					Dim US2,debugfilepath2,debugFso2,debugOut2,debugOutText2
					US2 = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
					Set debugOut2 = CreateObject("Scripting.FileSystemObject")
							
					If (debugOut2.FileExists(US2 & "\Desktop\ScriptInsertFileBackUp.txt")) Then
						Set debugOutText2 = debugOut2.OpenTextFile(US2 & "\Desktop\ScriptInsertFileBackUp.txt", 8, 2)
					Else
						Set debugOutText2 = debugOut2.CreateTextFile(US2 & "\Desktop\ScriptInsertFileBackUp.txt", True)
					End If
					debugOutText2.Write stringLineMSN
					debugOutText2.WriteLine
					debugOutText2.Write stringLine & ";"
					debugOutText2.WriteLine
					debugOutText2.Close
					' CLEANUP OBJECTS '
					Set debugOutText2 = Nothing
					Set debugOut2 = Nothing

					Exit Do
				End If
				Loop
				'******************************************************************************************************'
			Else
				MsgBox("SQL Insert successful")
				SQL_CONNECTION.Execute stringLine
			End If
		Loop
		debugOutText.Close
	Else
		msgbox "Nothing was found for importing"
		Dim US4,debugfilepath4,debugFso4,debugOut4,debugOutText4
		US4 = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
		Set debugOut4 = CreateObject("Scripting.FileSystemObject")
				
		If (debugOut4.FileExists(US4 & "\Desktop\LogFile.txt")) Then
			Set debugOutText4 = debugOut4.OpenTextFile(US4 & "\Desktop\LogFile.txt", 8, 2)
		Else
			Set debugOutText4 = debugOut4.CreateTextFile(US4 & "\Desktop\LogFile.txt", True)
		End If
		debugOutText4.Write Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()) & " | " & Hour(Now()) & ":" & Minute(Now()) & ":" & Second(Now()) & " | " & "Nothing was found for importing."
		debugOutText4.WriteLine
		debugOutText4.Close
		' CLEANUP OBJECTS '
		Set debugOutText4 = Nothing
		Set debugOut4 = Nothing
	End If
	
	Dim US3,debugfilepath3,debugFso3,debugOut3,debugOutText3
	US3 = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
	Set debugOut3 = CreateObject("Scripting.FileSystemObject")
	If (debugOut3.FileExists(US3 & "\Desktop\ScriptInsertFile.txt")) Then
		debugOut3.DeleteFile(US3 & "\Desktop\ScriptInsertFile.txt")
	Else
		
	End If
End Sub