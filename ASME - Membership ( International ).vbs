Option Explicit

Dim id, memberID, nameLast, nameFirst
Dim gender, dateOfBirth, interests, job
Dim address_1, address_2, address_3, address_4, city, stateProvidence, zipPostalCode, country
Dim phoneBusiness, phoneCell, fax, emailAddress
Dim school, degree, degreeType, instatution, dateBegin, dateGraduation, major
Dim instatutionInterest_1, instatutionInterest_2, instatutionInterest_3, instatutionInterest_4, instatutionInterest_5

Call Main

Sub Main()

	Dim fs, objTextFile
	Dim arrStr
	Dim wshShell

	Set fs = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = fs.OpenTextFile("C:\Users\ArellanoK\Downloads\ASME - Memberships - Original.csv")

	Do Until objTextFile.AtEndOfStream
		arrStr = split(objTextFile.ReadLine,",")
		id = arrStr(0)
		
		memberID = arrStr(1)
		nameLast = arrStr(2)
		nameFirst = arrStr(3)
		
		gender = arrStr(4)
		dateOfBirth = arrStr(5)
		interests = arrStr(6)
		
		address_1 = arrStr(7)
		address_2 = arrStr(8)
		address_3 = arrStr(9)
		address_4 = arrStr(10)
		city = arrStr(11)
		stateProvidence = arrStr(12)
		zipPostalCode = arrStr(13)
		country = arrStr(14)
		
		emailAddress = arrStr(15)
		
		school = arrStr(13)
	Loop
	
	Call keystrokes( )

	msgbox("Line " & id & " completed from file.")

	objTextFile.Close
	Set objTextFile = Nothing
	Set fs = Nothing

End Sub

Function keystrokes( )
	Dim wshShell
	Set wshShell = CreateObject("WScript.Shell") 
	wscript.sleep 1000
	
	wshshell.sendkeys "+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}"
	wshshell.sendkeys "+{TAB}" & country
	wshshell.sendkeys "+{TAB}" & stateProvidence
	wshshell.sendkeys "+{TAB}" & city
	wshshell.sendkeys "+{TAB}" & zipPostalCode
''	wshshell.sendkeys "+{TAB}" & address_4
''	wshshell.sendkeys "+{TAB}" & address_3
	wshshell.sendkeys "+{TAB}" & address_2
	wshshell.sendkeys "+{TAB}" & address_1
	
	wshshell.sendkeys "+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}"
	wshshell.sendkeys nameLast
	wshshell.sendkeys "+{TAB}+{TAB}" & nameFirst
	wshshell.sendkeys "+{TAB}+{TAB}Home+{TAB}" 
	
	wscript.sleep 500
	
	wshshell.sendkeys "^{TAB}{TAB}{TAB}" ' & phoneBusiness
	wshshell.sendkeys "{TAB}{TAB}{TAB}" ' & phoneCell
	wshshell.sendkeys "{TAB}{TAB}{TAB}" ' & fax
	wshshell.sendkeys "{TAB}{TAB}" & emailAddress
	wshshell.sendkeys "+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}"
	wshshell.sendkeys "+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}"
	wshshell.sendkeys " "
	
	wscript.sleep 3000
	wshshell.sendkeys "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}"
	wscript.sleep 1000
	wshshell.sendkeys "{RIGHT}"
	
	wshshell.sendkeys "{TAB}{TAB}{TAB}" ' & job
	wshshell.sendkeys "{TAB}{TAB}" & gender
	wshshell.sendkeys "{TAB}{TAB}" & dateOfBirth
	wshshell.sendkeys "^s"
	
	wshshell.sendkeys "%{}"
	wshshell.sendkeys "5"
	wshshell.sendkeys "V"
	wshshell.sendkeys "{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}"
	wshshell.sendkeys "{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}"
	wshshell.sendkeys "{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}"
	wshshell.sendkeys "{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}"
	wshshell.sendkeys "{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}"
	wshshell.sendkeys "{DOWN}{ENTER}"
	
	wshshell.sendkeys "+{TAB}{DOWN}{ENTER}"
	wshshell.sendkeys degree
	wshshell.sendkeys "{TAB}" & degreeType
	wshshell.sendkeys "{TAB}" & instatution & "{TAB}"
	wshshell.sendkeys "{TAB}" & dateBegin
	wshshell.sendkeys "{TAB}" & dateGraduation
	wshshell.sendkeys "{TAB}{TAB}{TAB}"
	wshshell.sendkeys " "
	wshshell.sendkeys "{TAB}" & major
	wshshell.sendkeys "^s"
	
	wshshell.sendkeys "%{}"
	wshshell.sendkeys "5"
	wshshell.sendkeys "V"
	wshshell.sendkeys "{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}"
	wshshell.sendkeys "{ENTER}"
	
	wshshell.sendkeys "{ENTER}"
	wshshell.sendkeys " "
	wshshell.sendkeys "{TAB}+{TAB}"
	wshshell.sendkeys instatutionInterest_1
	wshshell.sendkeys "{TAB}" & instatutionInterest_2
	wshshell.sendkeys "{TAB}" & instatutionInterest_3
	wshshell.sendkeys "{TAB}" & instatutionInterest_4
	wshshell.sendkeys "{TAB}" & instatutionInterest_5
	wshshell.sendkeys "^s"
	
	wscript.sleep 3000
	wshshell.sendkeys "%{}"
	wshshell.sendkeys "5"
	wshshell.sendkeys "V"
	wshshell.sendkeys "{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}"
	wshshell.sendkeys "{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}{UP}"
	wshshell.sendkeys "{ENTER}"
	wshshell.sendkeys "{DOWN}{DOWN}{DOWN}{DOWN}"
	wshshell.sendkeys "{ENTER}"
	
	wshshell.sendkeys "{UP}{LEFT}{LEFT}"
	wshshell.sendkeys "{TAB}{TAB}"
	wshshell.sendkeys "^{HOME}"
	wshshell.sendkeys "MMSTFR"
	wshshell.sendkeys "{TAB} "
	wshshell.sendkeys "{DOWN}{DOWN} "
	wshshell.sendkeys "{DOWN}{DOWN}{DOWN} "
End Function