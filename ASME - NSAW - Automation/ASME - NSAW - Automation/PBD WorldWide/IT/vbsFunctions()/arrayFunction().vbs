' Document Name:	argumentalFunction.vbs
' Last Modified:	07/21/2017
' Purpose:		Mainly focuses array function

Option Explicit

Dim zero
zero = inputbox("This is variable for zero.")

Dim theArray()
	Redim theArray(1, 1, 1, 1)
		theArray(0, 0, 0, 0) = zero
		theArray(0, 0, 0, 1) = "1"
		theArray(0, 0, 1, 0) = "2"
		theArray(0, 0, 1, 1) = "3"
		theArray(0, 1, 0, 0) = "4"
		theArray(0, 1, 0, 1) = "5"
		theArray(0, 1, 1, 0) = "6"
		theArray(0, 1, 1, 1) = "7"
		theArray(1, 0, 0, 0) = "8"
		theArray(1, 0, 0, 1) = "9"
		theArray(1, 0, 1, 0) = "10"
		theArray(1, 0, 1, 1) = "11"
		theArray(1, 1, 0, 0) = "12"
		theArray(1, 1, 0, 1) = "13"
		theArray(1, 1, 1, 0) = "14"
		theArray(1, 1, 1, 1) = "15"
		
	Dim i, j, k, l
	i = 0
	j = 0
	k = 0
	l = 0
	
Sub Main()
	Call arrayFunction()
End Sub
' ********************************************************************* '
Function arrayFunction()
	' FILLING 4 DIMENSIONAL ARRAY '
	For i = 0 to Ubound(theArray, 1) ' UNBOUND OF FIRST DIMENSTION '
		For j = 0 to Ubound(theArray, 2) ' UNBOUND OF SECOND DIMENSTION '
			For k = 0 to Ubound(theArray, 3) ' UNBOUND OF THIRD DIMENSTION '
				For l = 0 to Ubound(theArray, 4) ' UNBOUND OF FOURTH DIMENSTION '
					theArray(i, j, k, l) = "Row: " & i & " | " & "Column: " & j & " | " & "Layer: " & k & " | " & "Dimension: " & l & " | " & theArray(i, j, k, l)
				Next
			Next
		Next
	Next
	' FETCHING VALUES FROM 4 DIMENSIONAL ARRAY '
	For i = 0 to Ubound(theArray, 1) ' UNBOUND OF FIRST DIMENSTION '
		For j = 0 to Ubound(theArray, 2) ' UNBOUND OF SECOND DIMENSTION '
			For k = 0 to Ubound(theArray, 3) ' UNBOUND OF THIRD DIMENSTION '
				For l = 0 to Ubound(theArray, 4) ' UNBOUND OF FOURTH DIMENSTION '
					Call debugLogFileFunction()
				Next
			Next
		Next
	Next
	
	msgbox(UBound(theArray, 4) & " " & LBound(theArray, 4))
	
End Function
' ********************************************************************* '
Function debugLogFileFunction(ErrorMessage)					' Writes a log file for documentation '
	Dim UsLogFile, debugOutLogFile, debugOutTextLogFile			' Objects that stores information in a log file
	
	' CREATES THE WSCRIPT SHELL FOR WRITTING '
	UsLogFile					= CreateObject("WScript.Shell").ExpandEnvironmentStrings("C:\Connectship\failed records")
	Set debugOutLogFile				= CreateObject("Scripting.FileSystemObject")		' Creates File System OBJECT

	If (debugOutLogFile.FileExists(UsLogFile & "\VoidLogFile.txt")) Then	' Checks if VoidLogFile.txt exists
		Set debugOutTextLogFile = debugOutLogFile.OpenTextFile(UsLogFile & "\VoidLogFile.txt", 8, 2)		' Opens in writeable format
	Else
		Set debugOutTextLogFile = debugOutLogFile.CreateTextFile(UsLogFile & "\VoidLogFile.txt", True)		' Creates file in writeable format
	End If
	' WRITES THIS INFORMATION IN THE FILE '
	debugOutTextLogFile.Write Pd(Month(date()),2) & "/" & Pd(DAY(date()),2) & "/" & YEAR(Date()) & " | " _
				& Pd(Hour(Time()),2) & ":" & Pd(Minute(Time()),2) & ":" & Pd(Second(Time()),2) & " | " _
				& ErrorMessage
	debugOutTextLogFile.WriteLine
	debugOutTextLogFile.Close

	Set debugOutTextLogFile				= Nothing	' Clear Variables
	Set debugOutLogFile				= Nothing	' Clear Variables

End Function
' ********************************************************************* '
Function pd(n, totalDigits)						' Converts Time and Date into "MM/DD/YYYY | HH:MM:SS"
	If totalDigits > len(n) Then
		pd					= String(totalDigits-len(n),"0") & n
	Else
		pd					= n
	End If
End Function
' ********************************************************************* '