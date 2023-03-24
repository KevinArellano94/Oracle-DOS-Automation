Option Explicit
' DECLARES THE VARIABLES FOR SHOWING THE STATUS WINDOW '
Dim oIE
Dim oShell
Dim sTitle
Dim SWidth
Dim SHeight
Dim sWidthW
Dim sHeightW
Set oIE = CreateObject("InternetExplorer.Application")
Set oShell = CreateObject("WScript.Shell")

sTitle = "Message dialog"

With oIE

	.FullScreen = False
	.AddressBar = False
	.ToolBar = False
	.StatusBar = False
	.Resizable = False

	.Navigate("about:blank")
	Do Until .readyState = 4: wscript.sleep 100: Loop
		
		With .document
			With .ParentWindow
				
				SWidth = 1920
				SHeight = 1046

				SWidthW = (1920 * .25)
				SHeightW = (1046 * .2)

				.resizeto SWidthW, SHeightW
				.moveto (SWidth - SWidthW)/2, (SHeight - SHeightW)/3
			End With
		.WriteLn ("<html>")
		.WriteLn   ("<head>")
		.WriteLn     ("<script language=""vbscript"">")
		.WriteLn       ("Sub ExitButton ()")
		.WriteLn         ("ExitButtonID= LCase(Trim(window.event.srcelement.id))")
		.WriteLn         ("If Left(ExitButtonID, 4)=""xbtn"" Then")
		.WriteLn           ("window.exitbtn.value= Mid(ExitButtonID, 5)")
		.WriteLn         ("End If")
		.WriteLn       ("End Sub")
		.WriteLn       ("Sub NoRefreshKey ()")
		.WriteLn         ("Select Case window.event.keycode")
		.WriteLn           ("Case 82: SuppressKey= window.event.ctrlkey")
		.WriteLn           ("Case 116: SuppressKey= True")
		.WriteLn         ("End Select")
		.WriteLn         ("If SuppressKey Then")
		.WriteLn           ("window.event.keycode= 0")
		.WriteLn           ("window.event.cancelbubble= True")
		.WriteLn           ("window.event.returnvalue= False")
		.WriteLn         ("End If")
		.WriteLn       ("End Sub")
		.WriteLn       ("Sub NoContextMenu ()")
		.WriteLn         ("window.event.cancelbubble= True")
		.WriteLn         ("window.event.returnvalue= False")
		.WriteLn       ("End Sub")
		.WriteLn       ("Set document.onclick= GetRef(""ExitButton"")")
		.WriteLn       ("Set document.onkeydown= GetRef(""NoRefreshKey"")")
		.WriteLn       ("Set document.oncontextmenu= GetRef(""NoContextMenu"")")
		.WriteLn     ("</script>")
		.WriteLn   ("</head>")
		.WriteLn   ("<body>")

		.WriteLn    ("Starting the script....")

		.WriteLn   ("</body>")
		.WriteLn ("</html>")

		With .ParentWindow.document.body
		  .style.backgroundcolor = "white"
		  .scroll="no"
		  .style.Font = "20pt 'Georgia'"
		  .style.borderStyle = "outset"
		  .style.borderWidth = "2px"
		End With

		.Title = sTitle
		oIE.Visible = True
''		WScript.Sleep 100
		oShell.AppActivate sTitle
	End With
End With

Dim WshShell
Set WshShell = CreateObject("WScript.Shell")

Dim msgIENumber
Dim msgIESqlConnection
Dim msgIESqlDataGet
'Call webViwer()
Dim i
i = 0

Sub Main()
	
	Call Time()
	
End Sub

'***********************************************************************************************************************************************'

Function Time()
	
''	wscript.sleep 1000
	MsgIE("Connecting to the Database." & vbNewLine & vbNewLine & vbNewLine & vbNewLine & "Seconds that have passed: " & i)
	
End Function

'***********************************************************************************************************************************************'

Sub MsgIE(sMsg)
  On Error Resume Next ' Just in case the IE window is closed
  If sMsg = "IE_Quit" Then
    oIE.Quit
  Else
    oIE.Document.Body.InnerText = sMsg
    oShell.AppActivate sTitle
  End If
End Sub

'***********************************************************************************************************************************************'
