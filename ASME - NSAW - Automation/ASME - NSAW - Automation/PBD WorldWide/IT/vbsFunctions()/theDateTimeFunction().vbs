Option Explicit

Dim theDate, theTime

Sub Main()
	
	Call theDateTime
	msgbox("Today, " & theDate & ", the time that this program ran was at: " & theTime)
	
End Sub

'***********************************************************************************************************************************************'

Function theDateTime()
	theDate = Pd(Month(Date()),2) & "/" & Pd(Day(Date()),2) & "/" & Year(Date())
	theTime = Pd(Hour(Time()), 2) & ":" & Pd(Minute(Time()), 2) & ":" & Pd(Second(Time()), 2)
End Function

'***********************************************************************************************************************************************'

Function pd(n, totalDigits)
	If totalDigits > len(n) Then
		pd = String(totalDigits-len(n),"0") & n
	Else
		pd = n
	End If
End Function

'***********************************************************************************************************************************************'
