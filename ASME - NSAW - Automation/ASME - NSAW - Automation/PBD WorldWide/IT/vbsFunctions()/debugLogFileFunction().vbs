' Document Name:	debugLogFileFunction().vbs
' Last Modified:	07/13/2017
' Purpose:		This script serves as a copy/paste function template for file writting.

Option Explicit

Sub Main()
	Call debugLogFileFunction()
End Sub
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