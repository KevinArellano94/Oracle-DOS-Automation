' Document Name:        VC_Load_Order.vbs
' Last Modified:        09/12/2017
' Purpose:              Load order script to query VeraCore Firebird Database

Option Explicit

Dim databaseConnection, context, currentPackage, tempString             ' 
Dim pickSlipID, sqlQuery, sqlResult                                     ' 

Sub Main()								' Main Subroutine
	Set databaseConnection				= ScriptDataManager.StoredData( "VC_CONNECTION" ) 		' Create SQL Query and Capture Eesponse
	Set context					= ScriptDatamanager.ClientContext
	Set currentPackage				= context( "CURRENT_PACKAGE" )
	
	pickSlipID					= currentPackage( "REFERENCE_1" )
	sqlQuery					= "SELECT * FROM shpimp WHERE packageid = '" & pickSlipID & "'"
	Set sqlResult					= databaseConnection.Execute ( sqlQuery )
	
	If sqlResult.EOF Then
		sqlResult.Close
		tempString = "Pick slip number not found."
		seterror tempString
		MsgBox( "Order not found." )
		Exit Sub
	Else
		' Populate User Interface
		setField( "CONSIGNEE_COUNTRY",			replaceSpace( getRsString( sqlResult( "STOCNTRY" ) ) ),	True,		0 )
		setField( "CONSIGNEE_CONTACT",			getRsString( sqlResult( "STOCONTACT" ) ),		True,		0 )
		setField( "CONSIGNEE_COMPANY",			getRsString( sqlResult( "STONAME" ) ),			True,		0 )
		setField( "CONSIGNEE_ADDRESS_LINE_1",		getRsString( sqlResult( "STOADDR1" ) ),			True,		0 )
		setField( "CONSIGNEE_ADDRESS_LINE_2",		getRsString( sqlResult( "STOADDR2" ) ),			True,		0 )
		setField( "CONSIGNEE_ADDRESS_LINE_3",		getRsString( sqlResult( "STOADDR3" ) ),			True,		0 )
		setField( "CONSIGNEE_CITY",			getRsString( sqlResult( "STOCITY" ) ),			True,		0 )
		setField( "CONSIGNEE_STATEPROVINCE",		getRsString( sqlResult( "STOSTATE" ) ),			True,		0 )
		setField( "CONSIGNEE_POSTALCODE",		getRsString( sqlResult( "STOZIP" ) ),			True,		0 )
		setField( "CONSIGNEE_PHONE",			getRsString( sqlResult( "STOPHONE" ) ),			True,		0 )
		
		setField( "DECLARED_VALUE_AMOUNT",		getRsString( sqlResult( "INSURVAL" ) ),			True,		0 )
		setField( "REFERENCE_2",			getRsString( sqlResult( "FROMSYS" ) ),			True,		0 )
		setField( "RETURN_CONTACT",			"SHIPPING DEPT",					True,		0 )
		setField( "RETURN_ADDRESS_LINE_1",		"3280 SUMMIT RIDGE PKWY",				True,		0 )
		setField( "RETURN_ADDRESS_LINE_2",		"DISTRIBUTION CENTER",					True,		0 )
		setField( "RETURN_CITY",			"Duluth",						True,		0 )
		setField( "RETURN_STATEPROVINCE",		"GA",							True,		0 )
		setField( "RETURN_POSTALCODE",			"30096",						True,		0 )
		setField( "RETURN_PHONE",			"770-442-8633",						True,		0 )
		setField( "RETURN_COUNTRY",			"UNITED_STATES",					True,		0 )
		
		setField( "SHIPPER",				"Duluth",						True,		0 )
		
		' Check residential
		Dim residentualVariable
		residentualVariable				= getRsString( sqlResult( "COMMFLAG" ) )
		If ( residentualVariable = "R" ) Then
			setField( "CONSIGNEE_RESIDENTIAL",	1,							True,		0 )
		End If
		
		' Handle Third Party Billing (Will apply to those clients set up with Client Freight Accounts)
		Dim shipOption
		shipOption					= getRsString( sqlResult( "BILLTYPE" ) )
		If ( shipOption = "1" Or shipOption = "3" Or shipOption = "4" Or shipOption = "7" ) Then
			setfield( "THIRD_PARTY_BILLING",			True,						        	True,	0 )
			setfield( "THIRD_PARTY_BILLING_ACCOUNT",		getRsString( sqlResult( "THIRDPARTYACCT" ) ),	        	True,	0 )
			setField( "THIRD_PARTY_BILLING_COMPANY",		getRsString( sqlResult( "THIRDNAME" ) ),	        	True,	0 )
			setField( "THIRD_PARTY_BILLING_CONTACT",		getRsString( sqlResult( "THIRDCONTACT" ) ),	        	True,	0 )
			setField( "THIRD_PARTY_BILLING_ADDRESS_LINE_1",		getRsString( sqlResult( "THIRDADDR1" ) ),	        	True,	0 )
			setField( "THIRD_PARTY_BILLING_ADDRESS_LINE_2",		getRsString( sqlResult( "THIRDADDR2" ) ),	        	True,	0 )
			setField( "THIRD_PARTY_BILLING_ADDRESS_LINE_3",		getRsString( sqlResult( "THIRDADDR3" ) ),	        	True,	0 )
			setField( "THIRD_PARTY_BILLING_CITY",			getRsString( sqlResult( "THIRDCITY" ) ),	        	True,	0 )
			setField( "THIRD_PARTY_BILLING_STATEPROVINCE",		getRsString( sqlResult( "THIRDSTATE" ) ),	        	True,	0 )
			setField( "THIRD_PARTY_BILLING_COUNTRY",		replaceSpace( getRsString( sqlResult( "THIRDCNTRY" ) ) ),	True,	0 )
			setField( "THIRD_PARTY_BILLING_POSTALCODE",		getRsString( sqlResult( "THIRDZIP" ) ),			        True,	0 )
			setField( "THIRD_PARTY_BILLING_PHONE",			getRsString( sqlResult( "THIRDPHONE" ) ),			True,	0 )
			setField( "TERMS",					"SHIPPER",					        	True,	0 )
			setField( "TERMS_OF_SALE",				"DDP",						        	True,	0 )
		End If
		
		' Weigh Package
		Dim objectWeightMacro
		Set objectWeightMacro				= CreateObject("Progistics.Dictionary")
		objectWeightMacro.Value("NAME")			= "MACRO_WEIGH"
		ScriptDataManager.AddMacro objectWeightMacro
		Set objectWeightMacro				= Nothing
		
		' Handle Carrier/Service
		Dim service, shipMethod
		service = getRsString( sqlResult( "CARRIERCODE" ) )
		
		Dim currentPackageWeight
		currentPackageWeight = currentPackage("WEIGHT")
		
		If inStr( service, "U" ) Then
			Select Case service
				Case "U11"
					shipMethod = "CONNECTSHIP_UPS.UPS.GND"
				Case "U16"
					shipMethod = ""
				Case "U20"
					shipMethod = ""
				Case "U48"
					shipMethod = "CONNECTSHIP_UPS.UPS.STD"
				Case "U01"
					shipMethod = "CONNECTSHIP_UPS.UPS.NDA"
				Case "U60"
					shipMethod = "CONNECTSHIP_UPS.UPS.NAM"
				Case "U61"
					shipMethod = ""
				Case "U07"
					shipMethod = "CONNECTSHIP_UPS.UPS.2DA"
				Case "U35"
					shipMethod = "CONNECTSHIP_UPS.UPS.2AM"
				Case "U36"
					shipMethod = ""
				Case "U45"
					shipMethod = ""
				Case "U21"
					shipMethod = "CONNECTSHIP_UPS.UPS.3DA"
				Case "U24"
					shipMethod = ""
				Case "U02"
					shipMethod = ""
				Case "U26"
					shipMethod = ""
				Case "U44"
					shipMethod = ""
				Case "U43"
					shipMethod = "CONNECTSHIP_UPS.UPS.NDS"
				Case "U46"
					shipMethod = ""
				Case "U54"
					shipMethod = "CONNECTSHIP_UPS.UPS.EPD"
				Case "U08"
					shipMethod = ""
				Case "U25"
					shipMethod = ""
				Case "U49"
					shipMethod = "CONNECTSHIP_UPS.UPS.EXP"
				Case "U63"
					shipMethod = "CONNECTSHIP_UPS.UPS.EXPPLS"
				Case "U64"
					shipMethod = ""
				Case "U65"
					shipMethod = ""
				Case "U66"
					shipMethod = ""
				Case "U67"
					shipMethod = ""
				Case "U68"
					shipMethod = ""
				Case "U69"
					shipMethod = ""
                                Case "US1"
					shipMethod = "CONNECTSHIP_UPS.UPS.SPPS"
			End Select
		ElseIf inStr(service, "R") Then
			shipMethod = "TANDATA_FEDEXFSMS.FEDEX.GND"
		ElseIf inStr(service, "F") Then
			Select Case service
				Case "F14"
					shipMethod = "TANDATA_FEDEXFSMS.FEDEX.EXP"
				Case "F69"
					shipMethod = "TANDATA_FEDEXFSMS.FEDEX.IECO"
				Case "F75"
					shipMethod = "TANDATA_FEDEXFSMS.FEDEX.IFR2"
				Case "F06"
					shipMethod = "TANDATA_FEDEXFSMS.FEDEX.STD"
				Case "F07"
					shipMethod = ""
				Case "F08"
					shipMethod = ""
				Case "F09"
					shipMethod = ""
				Case "F10"
					shipMethod = ""
				Case "F15"
					shipMethod = "TANDATA_FEDEXFSMS.FEDEX.FR1"
				Case "F01"
					shipMethod = "TANDATA_FEDEXFSMS.FEDEX.PRI"
				Case "F02"
					shipMethod = ""
				Case "F03"
					shipMethod = ""
				Case "F04"
					shipMethod = ""
				Case "F05"
					shipMethod = ""
				Case "F18"
					shipMethod = "TANDATA_FEDEXFSMS.FEDEX.FO"
				Case "F19"
					shipMethod = ""
				Case "F60"
					shipMethod = ""
				Case "F61"
					shipMethod = "TANDATA_FEDEXFSMS.FEDEX.IFO"
				Case "F62"
					shipMethod = ""
				Case "F63"
					shipMethod = ""
				Case "F64"
					shipMethod = ""
				Case "F74"
					shipMethod = ""
				Case "F65"
					shipMethod = "TANDATA_FEDEXFSMS.FEDEX.IPRI"
				Case "F71"
					shipMethod = ""
				Case "F72"
					shipMethod = ""
				Case "F11"
					shipMethod = "TANDATA_FEDEXFSMS.FEDEX.2DA"
				Case "F16"
					shipMethod = "TANDATA_FEDEXFSMS.FEDEX.FR2"
				Case "F21"
					shipMethod = ""
				Case "F20"
					shipMethod = ""
				Case "F17"
					shipMethod = "TANDATA_FEDEXFSMS.FEDEX.FR3"
			End Select
		ElseIf inStr(service, "P") Then
			Select Case service
				Case "P01"
					shipMethod = "TANDATA_USPS.USPS.FIRST"
				Case "P02"
					shipMethod = ""
				Case "P32"
					shipMethod = ""
				Case "P41"
					shipMethod = ""
				Case "P42"
					shipMethod = ""
				Case "P43"
					shipMethod = ""
				Case "P44"
					shipMethod = "TANDATA_USPS.USPS.PS_DBMC"
				Case "P45"
					shipMethod = "TANDATA_USPS.USPS.MEDIA_BMC"
				Case "P46"
					shipMethod = "TANDATA_USPS.USPS.LIBR_BMC"
				Case "P47"
					shipMethod = "TANDATA_USPS.USPS.BPM"
				Case "P48"
					shipMethod = "TANDATA_USPS.USPS.BPM_BULK"
				Case "P49"
					shipMethod = ""
				Case "P50"
					shipMethod = ""
				Case "P60"
					shipMethod = ""
				Case "P61"
					shipMethod = ""
				Case "P62"
					shipMethod = ""
				Case "P63"
					shipMethod = ""
				Case "P64"
					shipMethod = ""
				Case "P65"
					shipMethod = ""
				Case "P66"
					shipMethod = ""
				Case "P80"
					shipMethod = ""
				Case "P81"
					shipMethod = ""
				Case "P75"
					shipMethod = ""
				Case "P76"
					shipMethod = ""
				Case "P78"
					shipMethod = "TANDATA_USPS.USPS.I_GXG"
				Case "P79"
					shipMethod = ""
				Case "P67"
					shipMethod = ""
				Case "P68"
					shipMethod = ""
				Case "P69"
					shipMethod = ""
				Case "P70"
					shipMethod = ""
				Case "P71"
					shipMethod = ""
				Case "P72"
					shipMethod = ""
				Case "P73"
					shipMethod = ""
				Case "P74"
					shipMethod = ""
				Case "P03"
					shipMethod = "TANDATA_USPS.USPS.PRIORITY"
				Case "P04"
					shipMethod = ""
				Case "P05"
					shipMethod = ""
				Case "P06"
					shipMethod = ""
			End Select
		ElseIf inStr(service, "C") Then
			Select Case service
				Case "C01"
					shipMethod = "TANDATA_LTL.COUR.COUR"
				Case "C02"
					shipMethod = ""
				Case "C03"
					shipMethod = ""
				Case "C04"
					shipMethod = "CONNECTSHIP_UPSMAILINNOVATIONS.UPS.STD"
				Case "C05"
					shipMethod = "TANDATA_LTL.MULT.MULT"
			End Select
		End If
		
		setField( "PRIMARY_SUBCATEGORY",		shipMethod,		True,		0 )
		sqlResult.Close
		Set sqlResult					= Nothing
	End If
End Sub
' ********************************************************************* '
Function getRsString( objRSField )					' Function to Select Data From SQL Response
    Dim retString
    
    retString							= ""
    If IsNull( objRSField ) = True Then
        retString						= ""
    Else
        retString						= objRSField
        retString						= Trim( retString )
    End If
    
    retString							= Replace( retString, """", "" )
    retString							= Replace( retString, "'", "" )
    
    getRsString = retString
End Function
' ********************************************************************* '
Sub setField( fieldName, fieldValue, suppressScript, fieldIndex )	' Function to Set Fields in User Interface
    Dim macro
    Set macro							= CreateObject( "Progistics.Dictionary" )
    macro.Value( "NAME" )					= "MACRO_SET_FIELD"
    macro.Value( "FIELD_NAME" )					= fieldName
    macro.Value( "FIELD_VALUE" )				= fieldValue
    
    If fieldIndex > 0 Then
        macro.Value( "FIELD_INDEX" )				= fieldIndex
    End If
    
    macro.Value( "SUPPRESS_SCRIPT" )				= suppressScript
    ScriptDataManager.AddMacro macro
    Set macro							= Nothing
End Sub
' ********************************************************************* '
Function replaceSpace( strInput )					' Function to Handle Countries
	Dim Result
	Result							= strInput
	If InStr( strInput, " " ) > 0 Then
		Result						= Replace( strInput, " ", "_" )
	End If
	
	If ( strInput = "USA" ) Then
		Result						= "UNITED_STATES"
	End If
	
	If ( strInput = "US" ) Then
		Result						= "UNITED_STATES"
	End If
	
	replaceSpace						= Result
End Function
' ********************************************************************* '
Sub setError( buf )
	Dim macro
	Set macro						= CreateObject("Progistics.Dictionary")
	macro.Value("NAME")					= "MACRO_SET_PACKAGE_ERROR"
	macro.Value("ERROR_MESSAGE")				= buf
	ScriptDataManager.AddMacro macro
	Set macro						= Nothing
End Sub
' ********************************************************************* '