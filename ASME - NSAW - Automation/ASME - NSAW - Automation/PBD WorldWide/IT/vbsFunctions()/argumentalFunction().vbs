' Document Name:	argumentalFunction.vbs
' Last Modified:	07/21/2017
' Purpose:		Mainly focuses on variable retrival function

Option Explicit

Dim restrictedCountries, restrictedCountriesList
Dim restrictedShippingCountries, restrictedShippingCountriesList
Dim country, countrys

Sub Main()
	Call CountryRestrictionFunction(restrictedCountries)

	If (restrictedCountries = True) Then
		msgbox("Restricted Country: " & country & vbNewLine & "RestrictedCountries Flag = " & restrictedCountries)
	Else
		msgbox("Country Selected: " & country & vbNewLine & "RestrictedCountries Flag = " & restrictedCountries)
	End If

	If (restrictedShippingCountries = True) Then
		msgbox("Restricted Country: " & countrys & vbNewLine & "RestrictedCountries Flag = " & restrictedShippingCountries)
	Else
		msgbox("Country Selected: " & countrys & vbNewLine & "RestrictedCountries Flag = " & restrictedShippingCountries)
	End If
End Sub
' ********************************************************************* '
Function CountryRestrictionFunction(restrictedCountries)
	restrictedCountries = False
	restrictedCountriesList = "" & _
		"Cuba" & vbNewLine & _
		"Iran" & vbNewLine & _
		"North Korea" & vbNewLine & _
		"Burma" & vbNewLine & _
		"Myanmar" & vbNewLine & _
		"Sudan" & vbNewLine & _
		"Syria"
	
	country = inputbox("Pick a country." & vbNewLine & restrictedCountriesList)
	
	If (InStr(restrictedCountriesList, country)) Then
		restrictedCountries = True
	Else
		restrictedCountries = False
	End If
	Call CountryShippingFunction()
End Function
' ********************************************************************* '
Function CountryShippingFunction()
	restrictedShippingCountries = False
	restrictedShippingCountriesList = "" & _
		"UNITED_ARAB_EMIRATES" & vbNewLine & _
		"AFGHANISTAN" & vbNewLine & _
		"BAHRAIN" & vbNewLine & _
		"BURUNDI" & vbNewLine & _
		"IRAN" & vbNewLine & _
		"JORDAN" & vbNewLine & _
		"KUWAIT" & vbNewLine & _
		"LEBANON" & vbNewLine & _
		"OMAN" & vbNewLine & _
		"QATAR" & vbNewLine & _
		"SAUDI_ARABIA" & vbNewLine & _
		"SYRIA" & vbNewLine & _
		"YEMEN" & vbNewLine & _
		"ZAMBIA"
	countrys = inputbox("Pick a country from list." & vbNewLine & restrictedShippingCountriesList)
	
	If (InStr(restrictedShippingCountriesList, countrys)) Then
		restrictedShippingCountries = True
	Else
		restrictedShippingCountries = False
	End If
End Function
' ********************************************************************* '