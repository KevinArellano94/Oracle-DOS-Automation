Option Explicit

Sub Main

	Dim sql, sqlobjDB, sqlconstring, SQL_CONNECTION
	Set sqlobjDB = CreateObject("ADODB.Connection")
	
	sqlconstring = "Provider=SQLOLEDB.1;Data Source=qts-csdev;Initial Catalog=ORACLE_SIM;user id='connectship_test'; password='connectship'"
	sqlobjDB.Open sqlconstring
	
	Set SQL_CONNECTION = sqlobjDB
	
	Dim MSN, csutils
	MSN = 123456
	
	Dim errorPromptString
	errorPromptString = "The package IS SHIPPED but not logged in Oracle." & vbLf & "Please notify PBD IT if this error continues." & vbLf & "Additional log attempts will be attempted at 6:45pm and 9:45pm EST."
		
	sql =       "INSERT into ORACLE_SIM.dbo.UPS100F"
	sql = sql + " (CS_Workstation, SHIPPER_REFERENCE, CONSIGNEE_REFERENCE,"
	sql = sql + " REF_1, REF_2, REF_3,"
	sql = sql + " SHIPDATE, DIMENSION, TOTAL,"
	sql = sql + " WEIGHT, TTL_FREIGHT, TTL_WEIGHT,"
	sql = sql + " DESCRIPTION, CURRENT_PACKAGE, TOTAL_PACKAGES,"
	sql = sql + " MSN, SERVICE, PACKAGING, COMPANY,"
	sql = sql + " CONTACT, ADDRESS1, ADDRESS2,"
	sql = sql + " ADDRESS3, CITY, STATEPROVINCE,"
	sql = sql + " POSTALCODE, COUNTRYSYMBOL, TRACKING_NUMBER,"
	sql = sql + " TERMS, SHIPPER, ARRIVE_DATE,"
	sql = sql + " COD_AMOUNT, CODE, UPINVOICE, COMMODITY_CLASS, BASE_CHARGE, RESIDENTIAL_CHARGE, FUEL_SURCHARGE, ACCESSORIAL_CHARGE, UPZONE )"
	sql = sql + " values ('TESTCOMPUTER', 'SHIPPER_REFERENCE', 'CONSIGNEE_REFERENCE'"  
	sql = sql + ", '12364567', '01', 'REF3'" 
	sql = sql + ", '02-28-2017', '0x0x0', '12.89'"
	sql = sql + ", '14.5', '10.05', '10.05'"
	sql = sql + ", 'TESTdescript', '1', '1'"
	sql = sql + ", " & "'" & MSN & "', 'UPS_Ground' , 'packaging'"
	sql = sql + ", 'PBD Worldwide', 'TestContact' , '1650 Bluegrass Lakes Pkwy'"
	sql = sql + ", '', '', 'Alpharetta'"
	sql = sql + ", 'GA', '30004', 'UNITED_STATES'"
	sql = sql + ", '123456987ZXC' , 'DDU', 'ALPHA'"
	sql = sql + ", '10/01/18', '12.55', '123BAR456CODE' , 'reference18'"
	sql = sql + ", 'CommClass', '12.55', '3.50', '0.00', '1.00', '6')"

	Dim rs, loopCount
	loopCount = 1
	Do Until loopCount = 5
		MsgBox("Attempting insert " & loopCount)
		SQL_CONNECTION.Execute sql
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
		
		Dim chksql
		MsgBox("Delete the database information")
		chksql = "SELECT * FROM ORACLE_SIM.dbo.UPS100F WHERE MSN = '" & MSN & "'"
		Set rs = SQL_CONNECTION.Execute (chksql)
		
		If (rs.eof) Then
			MsgBox("Uh oh. SQL Insert did NOT commit")
		Else
			MsgBox("SQL Insert successful")
		Exit Do
		End If
	
	loopCount = loopCount + 1
	If loopCount = 5 Then
		MsgBox("Well, crap. All inserts failed.")
		
		Dim US,debugfilepath,debugFso,debugOut,debugOutText,MSN1
		MSN1 = MSN
		US = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%UserProfile%")
		Set debugOut = CreateObject("Scripting.FileSystemObject")
				
		If (debugOut.FileExists(US & "\Desktop\ScriptInsertFile.txt")) Then
			Set debugOutText = debugOut.OpenTextFile(US & "\Desktop\ScriptInsertFile.txt", 8, 2)
		Else
			Set debugOutText = debugOut.CreateTextFile(US & "\Desktop\ScriptInsertFile.txt", True)
		End If
		debugOutText.Write MSN
		debugOutText.WriteLine
		debugOutText.Write sql & ";"
		debugOutText.WriteLine
		debugOutText.Close
		' CLEANUP OBJECTS '
		Set debugOutText = Nothing
		Set debugOut = Nothing
		
		Exit Do
	End If
	Loop
End sub