' Option Explicit '

Dim library
Dim dbConn

library = "pbdprddta." ' GOOD ON 720 '

Dim context
Dim currentPackage
Dim ResidentialFee_UPSGND
Dim ResidentialFee_NonUPSGND
Dim upseq

Sub Main()

	' RESET PRYOR EVENT LABEL FLAG '
	ScriptDataManager.StoredData("EVENTLABEL") = False
	
	Set dbConn = ScriptDataManager.StoredData("PBD_CONNECTION")
	
	On Error Resume Next
	Dim dbconn, connect, sqltest, error_message, library
	library = "pbdprddta." ' GOOD ON 720 '
	
	upseq = 1	'7377'
	
	' LOAD DATA FROM THE CLIENT '
	Set context = ScriptDataManager.ClientContext
	Set currentPackage = context("CURRENT_PACKAGE")
	error_message = package( "ERROR_MESSAGE" )
	
	' TEST FOR CONNECTION ERROR, PERHAPS BAD CREDENTIALS OR SOURCE NOT AVAILABLE '
	Set dbconn = ScriptDataManager.StoredData("PBD_CONNECTION")
	sqltest = "SELECT MIN(MSN) FROM " & library & " UPS100F"
	dbconn.Execute (sqltest)
	
	' TEST FOR CONNECTION ERROR, PERHAPS BAD CREDENTIALS OR SOURCE NOT AVAILABLE '
	' CHECK IF CONNECTION EXISTS '
	If (InStr(Err.Description,"Communication link failure") > 0) Then
		writeDebug Err.Description
	Else
	
		If (InStr(Err.Description,"Operation is not allowed") > 0) Then
			writeDebug Err.Description
		Else
	
			If (InStr(error_message,"Sockets error") > 0) Then
				writeDebug error_message
			
			Else
			
				If (InStr(Err.Description,"Closed") > 0) Then
					writeDebug Err.Description
				End If
			End If
		End If
	End If
	
	On Error Goto 0
	
	' INITIALIZE THE ScriptDataManager.ClientContext FIRST, MAKING CSW SCRIPTING ACCESSIBLE '
	Set context = ScriptDataManager.ClientContext
	
	' RETRIEVE THE CURRENT PACAKGE CONTEXT '
	Dim packageList
	Dim upseq ' VALUE FOR UPS105F (Project 7377) '
	upseq = 1
	ResidentialFee_UPSGND = 2.55
	ResidentialFee_NonUPSGND = 3.00
	If (IsObject(context( "CURRENT_PACKAGE_LIST" ))) Then
		Set packageList = context( "CURRENT_PACKAGE_LIST" )

		' RETRIEVE SHIPMENT DATA FROM THE APPLICATION '
		' msgbox "Running Multiple Package Insert" '
		For Each package In packageList
			Dim COMPUTER					' MACHINE NAME '
			Dim SHIPPER						' SHIPPER SYMBOL '
			Dim SHIPDATE					' SHIPDATE '
			Dim ARRIVE_DATE					' SHIPDATE '
			Dim PRIMARY_SUBCATEGORY			' SERVICE SYMBOL IN Server/Carrier/Service FORMAT '
			Dim PRIMARY_CATEGORY			' SERVICE SYMBOL IN Server/Carrier/Service FORMAT '
			Dim WEIGHT						' PACKAGE WEIGHT '
			Dim DIMENSION					' PACKAGE DIMENSION '
			Dim PACKAGING					' PACKAGING TYPE '
			Dim TERMS						' PAYMENT TERMS '
			Dim CONSIGNEE					' CONSIGNEE '
			Dim CONSIGNEE_COMPANY			' CONSIGNEE COMPANY '
			Dim CONSIGNEE_CONTACT			' CONSIGNEE CONTACT '
			Dim CONSIGNEE_ADDRESS1			' CONSIGNEE ADDRESS1 '
			Dim CONSIGNEE_ADDRESS2			' CONSIGNEE ADDRESS2 '
			Dim CONSIGNEE_CITY				' CONSIGNEE CITY '
			Dim CONSIGNEE_STATEPROVINCE		' CONSIGNEE STATEPROVINCE '
			Dim CONSIGNEE_POSTALCODE		' CONSIGNEE POSTAL CODE '
			Dim CONSIGNEE_COUNTRY			' CONSIGNEE COUNTRY '
			Dim COMMODITY_CLASS				' LTL COMMODITY CLASS '
			Dim COD_AMOUNT					' COD AMOUNT '
			Dim DESCRIPTION					' PACKAGE DESCRIPTION '
			Dim NOFN_SEQUENCE				' CURRENT PACKAGE NUMBER '
			Dim NOFN_TOTAL					' NUMBER OF TOTAL PACKAGES '
			Dim TRACKING_NUMBER				' PACKAGE TRACKING NUMBER '
			Dim PRIMARY_TOTAL				' RATE TOTAL CHARGES '
			Dim PRIMARY_BASE				' RATE BASE CHARGES '
			Dim PRIMARY_SPECIAL				' RATE FUEL SURCHARGE AND OTHER ACCESSORIAL FEES'
			Dim RESIDENTIAL_CHARGE
			Dim FUEL_SURCHARGE				' RATE FUEL SURCHARGE '
			Dim ACCESSORIAL_CHARGE
			Dim BASE_CHARGE
			Dim ZONE						' PACKAGE ZONE'
			Dim MSN							' PACKAGE MSN '
			Dim BAR_CODE
			Dim REFERENCE_1					' REFERENCE 1 '
			Dim REFERENCE_2					' REFERENCE 2 '
			Dim REFERENCE_3					' REFERENCE 3 '
			Dim REFERENCE_5					' REFERENCE 5 '
			Dim REFERENCE_6					' ORIGINAL REFERENCE 1 '
			Dim REFERENCE_7					' REFERENCE 7 '
			Dim REFERENCE_8					' REFERENCE 8 '
			Dim REFERENCE_9					' REFERENCE 9 '
			Dim REFERENCE_12				' REFERENCE 12 '
			Dim REFERENCE_14				' REFERENCE 14  - PO NUMBER - FLAG FOR CASEMATE/WALMART ORDERS AND INVOICES FOR ASQ ORDERS '
			Dim REFERENCE_15	'7377'		' REFERENCE 15 - SELECTED CARRIER '
			Dim REFERENCE_16	'7377'		' REFERENCE 16 - CARTON ID '
			Dim REFERENCE_18				' REFERENCE 18 '
			Dim duns
			Dim duns2
			Dim dunstrack
			Dim packageID
			Dim REF1_PO_NUM
			
			' GLOBAL VARIABLE FOR STORING FLAG TO BE USED IN PRE_SHIP FOR SETTING CUSTOM FIELD VALUE '
			ScriptDataManager.StoredData("setCustomField") = True
			COMPUTER                = package("COMPUTER")				' MACHINE NAME '
			SHIPPER                 = package("SHIPPER")				' SHIPPER SYMBOL '
			SHIPDATE                = package("SHIPDATE")				' SHIPDATE '
			ARRIVE_DATE             = package("ARRIVE_DATE")			' SHIPDATE '
			PRIMARY_SUBCATEGORY     = package("PRIMARY_SUBCATEGORY")	' SERVICE SYMBOL IN Server/Carrier/Service FORMAT '
			PRIMARY_CATEGORY		= package("PRIMARY_CATEGORY")		' SERVICE SYMBOL IN Server/Carrier/Service FORMAT '
			WEIGHT                  = package("WEIGHT")					' PACKAGE WEIGHT '
			DIMENSION               = package("DIMENSION")				' PACAKGE DIMENSION '
			PACKAGING               = package("PACKAGING")				' PACKAGING TYPE '
			TERMS                   = package("TERMS")					' PAYMENT TERMS '
			Set CONSIGNEE           = package("CONSIGNEE")				' CONSIGNEE '
			CONSIGNEE_COMPANY       = CONSIGNEE.Company					' CONSIGNEE COMPANY '
			CONSIGNEE_CONTACT       = CONSIGNEE.Contact					' CONSIGNEE CONTACT '
			CONSIGNEE_ADDRESS1      = CONSIGNEE.Address1				' CONSIGNEE ADDRESS1 '
			CONSIGNEE_ADDRESS2      = CONSIGNEE.Address2				' CONSIGNEE ADDRESS2 '
			CONSIGNEE_CITY          = CONSIGNEE.City					' CONSIGNEE CITY '
			CONSIGNEE_STATEPROVINCE = CONSIGNEE.StateProvince			' CONSIGNEE STATEPROVINCE '
			CONSIGNEE_POSTALCODE    = CONSIGNEE.PostalCode				' CONSIGNEE POSTAL CODE '
			CONSIGNEE_COUNTRY       = CONSIGNEE.CountrySymbol			' CONSIGNEE COUNTRY '
			COMMODITY_CLASS         = package("COMMODITY_CLASS")		' LTL COMMODITY CLASS '
			COD_AMOUNT              = package("COD_AMOUNT")				' COD AMOUNT '
			DESCRIPTION             = package("DESCRIPTION")			' PACAKGE DESCRIPTION '
			NOFN_SEQUENCE           = package("NOFN_SEQUENCE")			' CURRENT PACKAGE NUMBER '
			NOFN_TOTAL              = package("NOFN_TOTAL")				' NUMBER OF TOTAL PACKAGES '
			TRACKING_NUMBER         = package("TRACKING_NUMBER")		' PACKAGE TRACING NUMBER '
			PRIMARY_TOTAL           = package("PRIMARY_TOTAL")			' RATE TOTAL CHARGES '
			PRIMARY_BASE			= package("PRIMARY_BASE")			' RATE BASE CHARGES '
			PRIMARY_SPECIAL			= package("PRIMARY_SPECIAL")		' RATE FUEL SURCHARGE AND OTHER ACCESSORIAL FEES '
			FUEL_SURCHARGE			= package("FUEL_SURCHARGE")			' RATE FUEL SURCHARGE '
			MSN                     = package("MSN")					' PACKAGE MSN '
			ZONE                    = package("ZONE")					' PACKAGE ZONE'
			REFERENCE_1             = package("REFERENCE_1") 			' REFERENCE 1 '
			REFERENCE_2             = package("REFERENCE_2") 			' REFERENCE 2 '
			REFERENCE_5             = package("REFERENCE_5") 			' REFERENCE 5 '
			REFERENCE_6             = package("REFERENCE_6") 			' ORIGINAL REFERENCE 1 '
			REFERENCE_14            = package("REFERENCE_14") 			' REFERENCE 14 - PO NUMBER - FLAG FOR CASEMATE/WALMART ORDERS AND INVOICES FOR ASQ ORDERS '
			REFERENCE_15 			= package("REFERENCE_15")	' 7377 '	' SELECTED CARRIER '
			REFERENCE_16			= package("REFERENCE_16")	' 7377 '	' CARTON ID '
			
			' SWITCH BACK INVOICE NUMBER WITH WALMART PO NUMBER FOR CASEMATE ORDERS WITH A PO NUMBER '
			If ((REFERENCE_2 = "09" Or REFERENCE_2 = "2P") And REFERENCE_14 <> "") Then
				REF1_PO_NUM 	= REFERENCE_1
				REF14_INV_NUM 	= REFERENCE_14
				REFERENCE_1 	= REF14_INV_NUM
				REFERENCE_6 	= REF14_INV_NUM
				REFERENCE_14 	= REF1_PO_NUM
			End If

			' THE FEDEX BARCODE IS OVER 30 CHARACTERS AND CONTAINS MORE INFORMATION THAN THE TRACKING NUMBER '
			' FOR THAT REASON WE ARE REPLACING package("BAR_CODE") WITH package("TRACKING_NUMBER") '
			BAR_CODE		= package("TRACKING_NUMBER")	' LABLE BARCODE '
			REFERENCE_3     = package("REFERENCE_3") 		' REF3 '
			REFERENCE_7     = package("REFERENCE_7") 		' ORIGINAL REF2 '
			REFERENCE_8     = package("REFERENCE_8") 		' ORIGINAL REF4 '
			REFERENCE_9     = package("REFERENCE_9") 		' ORIGINAL REF5 '
			REFERENCE_12    = package("REFERENCE_12") 		' DD Pick ID '
			REFERENCE_18    = package("REFERENCE_18") 		' WORLDEASE '
			
			' SCRUB DATA '
			REFERENCE_7 	= Trim(REFERENCE_7)
			DESCRIPTION 	= Replace(DESCRIPTION  , "'", "")
			DESCRIPTION 	= Mid(DESCRIPTION,1,50)
			
			' TRIM DATA FIELDS '
			COMPUTER                = Mid(COMPUTER,1,10)				' MACHINE NAME '
			SHIPPER                 = Mid(SHIPPER,1,40)					' SHIPPER SYMBOL '
			PRIMARY_SUBCATEGORY     = Mid(PRIMARY_SUBCATEGORY,1,60)		' SERVICE SYMBOL IN Server/Carrier/Service FORMAT '
			PRIMARY_SUBCATEGORY     = UCase(PRIMARY_SUBCATEGORY)
			WEIGHT                  = Mid(WEIGHT,1,15)					' PACKAGE WEIGHT MyWord = UCase("Hello World") '
			DIMENSION               = Mid(DIMENSION,1,15)				' PACKAGE DIMENSIONS '
			PACKAGING               = Mid(PACKAGING,1,30)				' PACKAGING TYPE '
			TERMS                   = Mid(TERMS,1,30)					' PAYMENT TERMS '
			CONSIGNEE_COMPANY       = Mid(CONSIGNEE_COMPANY,1,40)		' CONSIGNEE COMPANY '
			CONSIGNEE_CONTACT       = Mid(CONSIGNEE_CONTACT,1,40)		' CONSIGNEE CONTACT '
			CONSIGNEE_ADDRESS1      = Mid(CONSIGNEE_ADDRESS1,1,40)		' CONSIGNEE ADDRESS1 '
			CONSIGNEE_ADDRESS2      = Mid(CONSIGNEE_ADDRESS2,1,40)		' CONSIGNEE ADDRESS2 '
			CONSIGNEE_CITY          = Mid(CONSIGNEE_CITY,1,25)			' CONSIGNEE CITY '
			CONSIGNEE_STATEPROVINCE = Mid(CONSIGNEE_STATEPROVINCE,1,3)	' CONSIGNEE STATEPROVINCE '
			CONSIGNEE_POSTALCODE    = Mid(CONSIGNEE_POSTALCODE,1,12)	' CONSIGNEE POSTAL CODE '
			CONSIGNEE_COUNTRY       = Mid(CONSIGNEE_COUNTRY,1,60)		' CONSIGNEE COUNTRY '
			COMMODITY_CLASS         = Mid(COMMODITY_CLASS,1,3)			' LTL COMMODITY CLASS '
			COD_AMOUNT              = Mid(COD_AMOUNT,1,15)				' COD AMOUNT '
			DESCRIPTION             = Mid(DESCRIPTION,1,60)				' PACKAGE DESCRIPTION '
			NOFN_SEQUENCE           = Mid(NOFN_SEQUENCE,1,15)			' CURRENT PACKAGE NUMBER '
			NOFN_TOTAL              = Mid(NOFN_TOTAL,1,15)				' NUMBER OF TOTAL PACAKGES '
			TRACKING_NUMBER         = Mid(TRACKING_NUMBER,1,30)			' PACKAGE TRACKING NUMBER '
			PRIMARY_TOTAL           = Mid(PRIMARY_TOTAL,1,15)			' RATE TOTAL CHARGES '
			MSN                     = Mid(MSN,1,30)						' PACKAGE MSN'
			REFERENCE_3             = Mid(REFERENCE_3,1,30) 			' REF3 '
			REFERENCE_6             = Mid(REFERENCE_6,1,30) 			' ORIGINAL REF1 '
			REFERENCE_7             = Mid(REFERENCE_7,1,30) 			' ORIGINAL REF2 '
			REFERENCE_8             = Mid(REFERENCE_8,1,30) 			' ORIGINAL REF4 '
			REFERENCE_9             = Mid(REFERENCE_9,1,30) 			' ORIGINAL REF5 '
			REFERENCE_12            = Mid(REFERENCE_12,1,30) 			' DD PICK ID '
			
			' Scrub DATA '
			REFERENCE_7 	= Mid(REFERENCE_7,1,30)
			DESCRIPTION 	= Replace(DESCRIPTION  , "'", "")
			DESCRIPTION 	= Mid(DESCRIPTION,1,60)
			
			If PRIMARY_SUBCATEGORY = "TANDATA_LTL.AAA.AAA" And CONSIGNEE_COUNTRY = "CANADA"  Then
				PRIMARY_TOTAL = (PRIMARY_TOTAL + 19.00)
			End If
			
			BASE_CHARGE = PRIMARY_BASE
			RESIDENTIAL_CHARGE = 0.00
			
			ACCESSORIAL_CHARGE = PRIMARY_SPECIAL - FUEL_SURCHARGE
			
			' ADDED TO SUBTRACT 10.00 FROM PRIMARY SPECIAL FOR CANADA WORLDEASE SHIPMENTS '
			If PRIMARY_SUBCATEGORY = "TANDATA_UPS.UPS.WSTD" And CONSIGNEE_COUNTRY = "CANADA" AND REFERENCE_18 = "WORLDEASE" Then
				PRIMARY_SPECIAL 	= Csng(PRIMARY_SPECIAL - 10.00)
				PRIMARY_TOTAL 		= Csng( PRIMARY_TOTAL - 10.00)
				ACCESSORIAL_CHARGE 	= 0.00
			ElseIf InStr(PRIMARY_SUBCATEGORY, "CONNECTSHIP_UPSMAILINNOVATIONS") Then
				ACCESSORIAL_CHARGE 	= 0.00
			Else
				ACCESSORIAL_CHARGE 	= PRIMARY_SPECIAL - FUEL_SURCHARGE
				ACCESSORIAL_CHARGE 	= CSng(ACCESSORIAL_CHARGE)
			End If
			
			' ACCESSORIAL_CHARGE = PRIMARY_SPECIAL - FUEL_SURCHARGE '
			
			ACCESSORIAL_CHARGE = CSng(ACCESSORIAL_CHARGE)
			
			If CONSIGNEE.Residential = True and PRIMARY_CATEGORY = "TANDATA_UPS.UPS" And PRIMARY_SUBCATEGORY <> "TANDATA_UPS.UPS.SPPS" And CONSIGNEE_COUNTRY = "UNITED_STATES" Then
				If InStr(PRIMARY_SUBCATEGORY, "GND") > 0 Then
					If (BASE_CHARGE > ResidentialFee_UPSGND) Then
						RESIDENTIAL_CHARGE = ResidentialFee_UPSGND
					End If
				Else			
					If (BASE_CHARGE > ResidentialFee_NonUPSGND) Then
						RESIDENTIAL_CHARGE = ResidentialFee_NonUPSGND
					End If
				End If
				BASE_CHARGE = PRIMARY_BASE - RESIDENTIAL_CHARGE
			End If
			
			' ADDED THIS FOR ASQ DHL base_charge = 0 and fuel_charge = 0 '
			
			If InStr(PRIMARY_SUBCATEGORY, "DHL") and REFERENCE_2 = "28"  Then
				BASE_CHARGE 		= 0
				FUEL_SURCHARGE 		= 0
				ACCESSORIAL_CHARGE 	= 0
				PRIMARY_TOTAL 		= 0
				RESIDENTIAL_CHARGE 	= 0
				REFERENCE_6 		=  REFERENCE_14
				REFERENCE_1			=  REFERENCE_14
			End If
			
			' FSC UPDATED DMD 2/25/04 START '
			Dim newDate
			Dim tempString
			newDate 		= ""
			newDate 		= Year(SHIPDATE)
			tempString 		= Month(SHIPDATE)
			newDate 		= newDate & Right("0" & tempString, 2)
			tempString 		= Day(SHIPDATE)
			newDate 		= newDate & Right("0" & tempString, 2)
			
			If ARRIVE_DATE = "" Then
				ARRIVE_DATE = "00000000"
			Else
				Dim newARRIVE_DATE
				newARRIVE_DATE 		= Year(ARRIVE_DATE)
				tempString 			= Month(ARRIVE_DATE)
				tempString 			= Right("0" & tempString,2)
				newARRIVE_DATE 		= newARRIVE_DATE & tempString
				tempString 			= Day(ARRIVE_DATE)
				tempString 			= Right("0" & tempString,2)
				newARRIVE_DATE 		= newARRIVE_DATE & tempString
				ARRIVE_DATE 		= newARRIVE_DATE
			End If
			
			'********************************************************************************************************************'
			' THE FOLLOWING SECTIONS IS USED FOR CREATING THE ACTUAL TRACKING USED TO TRACK SMARTPOST PACKAGES '
			If PRIMARY_SUBCATEGORY = "TANDATA_MISC.SMPT.SMPT" Then
				'msgbox "Package is Smartpost"'
				packageID	= TRACKING_NUMBER
				If shipper = "ALPHA" Then
					Duns	= "D02901001012"
					duns2	= "9102901001012"
				ElseIf shipper = "DULUTH" Then
					Duns	= "D02901001012"
					duns2	= "9102901001012"
				ElseIf shipper = "TOS" Then
					Duns	= "D02927007742"
					duns2	= "9102927007742"
				End If
				
				dunstrack = Duns2 & packageID
				Dim pos1, pos2, pos3, pos4, pos5, pos6, pos7, pos8, pos9, pos10, pos11, pos12, pos13, pos14, pos15, pos16, pos17, pos18, pos19, pos20, pos21, chkdgtnumodd, chkdgtnumeven, chkdgtnumsum, chkdgtnum
				pos1 = Mid(dunstrack,21,1)
				pos2 = Mid(dunstrack,20,1)
				pos3 = Mid(dunstrack,19,1)
				pos4 = Mid(dunstrack,18,1)
				pos5 = Mid(dunstrack,17,1)
				pos6 = Mid(dunstrack,16,1)
				pos7 = Mid(dunstrack,15,1)
				pos8 = Mid(dunstrack,14,1)
				pos9 = Mid(dunstrack,13,1)
				pos10 = Mid(dunstrack,12,1)
				pos11 = Mid(dunstrack,11,1)
				pos12 = Mid(dunstrack,10,1)
				pos13 = Mid(dunstrack,9,1)
				pos14 = Mid(dunstrack,8,1)
				pos15 = Mid(dunstrack,7,1)
				pos16 = Mid(dunstrack,6,1)
				pos17 = Mid(dunstrack,5,1)
				pos18 = Mid(dunstrack,4,1)
				pos19 = Mid(dunstrack,3,1)
				pos20 = Mid(dunstrack,2,1)
				pos21 = Mid(dunstrack,1,1)
				pos1 = cint(pos1)
				pos2 = cint(pos2)
				pos3 = cint(pos3)
				pos4 = cint(pos4)
				pos5 = cint(pos5)
				pos6 = cint(pos6)
				pos7 = cint(pos7)
				pos8 = cint(pos8)
				pos9 = cint(pos9)
				pos10 = cint(pos10)
				pos11 = cint(pos11)
				pos12 = cint(pos12)
				pos13 = cint(pos13)
				pos14 = cint(pos14)
				pos15 = cint(pos15)
				pos16 = cint(pos16)
				pos17 = cint(pos17)
				pos18 = cint(pos18)
				pos19 = cint(pos19)
				pos20 = cint(pos20)
				pos21= cint(pos21)
				
				' USED TO CACLUATE THE CHECK DIGIT '
				chkdgtnumodd	= (pos1 + pos3 + pos5 + pos7 + pos9 + pos11 + pos13 + pos15 + pos17 + pos19 + pos21) * 3
				chkdgtnumeven	= pos2 + pos4 + pos6 + pos8 + pos10 + pos12 + pos14 + pos16 + pos18 + pos20
				chkdgtnumsum	= chkdgtnumodd + chkdgtnumeven
				chkdgtnum		= 10 - Right(chkdgtnumsum, 1)
				If chkdgtnum = 10 Then
					chkdgtnum	= 0
				End If
				
				Dim TRACKING_NUM, TRACKING_NUMBER2
				TRACKING_NUM		= dunstrack & chkdgtnum
				TRACKING_NUMBER2	= Mid(TRACKING_NUM, 3, 20)
				
				' THIS IS THE ACTUAL NUMBER USED FOR SMARTPOST TRACKING '
				TRACKING_NUMBER		= TRACKING_NUMBER2
			End If
			
			If PRIMARY_SUBCATEGORY = "TANDATA_MISC.SMPT.SMPT" Then
				TRACKING_NUMBER = TRACKING_NUMBER2
			End If
			
			'**************************************************************************************************'
			' WRITE BACK A CUSTOM TRACKING NUMBER FOR UMI SHIPMENTS TO ASSIS WITH FRIEHT BILLING '
			' THE UMI TRACKING NUMBER TO BE WRITTEN BACK TO UPS100F WILL BE 9 - 10 CHARACTERS: REF1 (7 - 8) + REF2 (2) '
			
			If PRIMARY_SUBCATEGORY = "CONNECTSHIP_UPSMAILINNOVATIONS.UPS.ECO" Then
				If(NOFN_TOTAL > 1) Then
					If (SHIPPER = "ALPHA") Then
						TRACKING_NUMBER = "MI" & "001306" & currentPackage("REFERENCE_1") & "." & NOFN_SEQUENCE
					ElseIf (SHIPPER ="DULUTH") Then
						TRACKING_NUMBER = "MI" & "003195" & currentPackage("REFERENCE_1") & "." & NOFN_SEQUENCE
					ElseIf (SHIPPER ="PHILLY") Then
						TRACKING_NUMBER = "MI" & "008601" & currentPackage("REFERENCE_1") & "." & NOFN_SEQUENCE
					ElseIf (SHIPPER ="WMU") Then
						TRACKING_NUMBER = "MI" & "012656" & currentPackage("REFERENCE_1") & "." & NOFN_SEQUENCE
					ElseIf (SHIPPER ="JONES") Then
						TRACKING_NUMBER = "MI" & "007173" & currentPackage("REFERENCE_1") & "." & NOFN_SEQUENCE
					ElseIf (SHIPPER = "WASHDC") Then
						TRACKING_NUMBER = "MI" & "000000" & currentPackage("REFERENCE_1") & "." & NOFN_SEQUENCE
					Else
						TRACKING_NUMBER = currentPackage("REFERENCE_1") & "." &  currentPackage("REFERENCE_2")
					End If
				End If
			End If
			
			If TRACKING_NUMBER = "" Then
				TRACKING_NUMBER = "n/a"
			End If
			
			' UPDATE DATABASE '
			Dim sql
			Dim sql105f
			Dim sql105f2
			If (1=1) Then
				sql =       "INSERT into " & library & "UPS100F"
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
				sql = sql + " values ('" & Left(COMPUTER,10) & "' , '" & REFERENCE_6 & "' , '" & REFERENCE_7
				sql = sql + "' , '" & REFERENCE_3 & "' , '" & REFERENCE_8 & "' , '" & REFERENCE_9
				sql = sql + "' , '" & newDate & "' , '" & DIMENSION & "' , '" & PRIMARY_TOTAL
				sql = sql + "' , '" & WEIGHT & "' , '" & PRIMARY_TOTAL & "'  , '" & WEIGHT
				sql = sql + "' , '" & DESCRIPTION & "' , '" & NOFN_SEQUENCE & "' , '" & NOFN_TOTAL
				sql = sql + "' , '" & MSN & "' , '" & PRIMARY_SUBCATEGORY & "' , '" & PACKAGING
				sql = sql + "' , '" & Replace(CONSIGNEE_COMPANY,"'","|") & "' , '" & Replace(CONSIGNEE_CONTACT,"'","|") & "' , '" & Replace(CONSIGNEE_ADDRESS1,"'","|")
				sql = sql + "' , '" & Replace(CONSIGNEE_ADDRESS2,"'","|") & "' , '" & Replace(CONSIGNEE_ADDRESS2,"'","|") & "' , '" & Replace(CONSIGNEE_CITY,"'","|")
				sql = sql + "' , '" & Left(CONSIGNEE_STATEPROVINCE,3) & "' , '" & CONSIGNEE_POSTALCODE & "' , '" & CONSIGNEE_COUNTRY
				sql = sql + "' , '" & TRACKING_NUMBER & "' , '" & TERMS & "', '" & SHIPPER
				sql = sql + "' , '" & ARRIVE_DATE & "' , '" & COD_AMOUNT & "' , '" & BAR_CODE & "' , '" & REFERENCE_18
				sql = sql + "' , '" & Left(COMMODITY_CLASS,3) & "'," & BASE_CHARGE & "," & RESIDENTIAL_CHARGE & "," & FUEL_SURCHARGE & "," & ACCESSORIAL_CHARGE & ", '" & ZONE & "')"
				writeDebug "SQL multi package-" & sql
				
				' PROJECT 7377 WRITE TO UPS105F '
				If ScriptDataManager.StoredData("AMAZON") = "Y" Then
					sql105f = "INSERT into " & library & "UPS105F"
					sql105f = sql105f + " (UPMSN, UPNAM, UPSEQ, UPVAL, UPSTK#) VALUES ('" & MSN & "', 'SSCC18', '" & upseq & "', '" & REFERENCE_16 & "', '" & TRACKING_NUMBER & "')"
					dbConn.Execute sql105f
					sql105f2 = "INSERT into " & library & "UPS105F"
					sql105f2 = sql105f2 + " (UPMSN, UPNAM, UPSEQ, UPVAL, UPSTK#) VALUES ('" & MSN & "', 'SCAC', '" & upseq & "', '" & REFERENCE_15 & "', '" & TRACKING_NUMBER & "')"
					dbConn.Execute sql105f2
					upseq = upseq + 1
				End If
				
				Dim successFlag
				successFlag = False
				Dim rs, loopCount
				loopCount = 1
				Do Until loopCount = 5
					dbConn.Execute sql
					' IF SQL INSERT DOES NOT ERROR A CHECK WILL BE MADE '
					' IF CHECK DOES NOT FIND DATA THEN SQL WRITTEN TO DEBUG '
					Dim dbError
					dbError = ""
					If dbConn.Errors.Count > 0 Then
						Dim intcount
						Dim dbCommError
						For intCount = 0 To dbConn.Errors.Count - 1
							Set dbCommError = dbConn.Errors.Item(intCount)
							writeDebug dbcommerror.Description
							dbError = dbError + dbCommError.Description + " ** "
						Next
						writeDebug "SQL-" & sql
						setError(dbError)
					End If
					Dim chksql
					chksql = "SELECT * FROM " & library & "UPS100F WHERE MSN = '" & MSN & "'"
					Set rs = dbConn.Execute (chksql)
					If (rs.eof) Then
						writeDebug "SQL Insert did not commit for MSN: " & MSN
						writeDebug sql
					Else
						writedebug "SQL Insert committed for MSN: " & MSN
						successFlag = True
						Exit Do
					End If
					
					loopCount = loopCount + 1
					If loopCount = 5 Then
						Dim record
						record =	Left(COMPUTER,10) & "," & REFERENCE_6 & "," & REFERENCE_7
						record = record + "," & REFERENCE_3 & "," & REFERENCE_8 & "," & REFERENCE_9
						record = record + "," & newDate & "," & DIMENSION & "," & PRIMARY_TOTAL
						record = record + "," & WEIGHT & "," & PRIMARY_TOTAL & "  , " & WEIGHT
						record = record + "," & DESCRIPTION & "," & NOFN_SEQUENCE & "," & NOFN_TOTAL
						record = record + "," & MSN & "," & PRIMARY_SUBCATEGORY & "," & PACKAGING
						record = record + "," & Replace(CONSIGNEE_COMPANY,"","|") & "," & Replace(CONSIGNEE_CONTACT,"","|") & "," & Replace(CONSIGNEE_ADDRESS1,"","|")
						record = record + "," & Replace(CONSIGNEE_ADDRESS2,"","|") & "," & Replace(CONSIGNEE_ADDRESS2,"","|") & "," & Replace(CONSIGNEE_CITY,"","|")
						record = record + "," & Left(CONSIGNEE_STATEPROVINCE,3) & "," & CONSIGNEE_POSTALCODE & "," & CONSIGNEE_COUNTRY
						record = record + "," & TRACKING_NUMBER & "," & TERMS & ", " & SHIPPER
						record = record + "," & ARRIVE_DATE & "," & COD_AMOUNT & "," & BAR_CODE & "' , '" & REFERENCE_18
						record = record + "," & Left(COMMODITY_CLASS,3) & "," & BASE_CHARGE & "," & RESIDENTIAL_CHARGE & "," & FUEL_SURCHARGE & "," & ACCESSORIAL_CHARGE & "," & ZONE
						writeDebug record
						Exit Do
					End If
				Loop
				If (successFlag = False) Then
					writeInsertFile sql
				End If
			End If
			
			' RESET THE DOCUMENT ASSIGNMENTS OVERRIDEN IN THE PRE_SHIP SCRIPT '
			Dim macro, sc
			sc = currentPackage("PRIMARY_CATEGORY")
			If PRIMARY_SUBCATEGORY = "TANDATA_USPS.USPS.SPCL" And REFERENCE_2 = "44" Then
				Set macro = CreateObject("Progistics.Dictionary")
				macro.Value("NAME")						= "MACRO_MODIFY_DOCUMENT_ATTRIBUTES"
				macro.Value("CATEGORY")					= sc
				macro.Value("DOCUMENT_FORMAT")			= "CUSTOM_TANDATA_USPS_LABEL.MMS"
				macro.Value("PRINT")					= True
				macro.Value("VIEW_DOCUMENT_MANAGER")	= False
				ScriptDataManager.AddMacro macro
				Set Macro = Nothing
			Else
				Set macro = CreateObject("Progistics.Dictionary")
				macro.Value("NAME")						= "MACRO_MODIFY_DOCUMENT_ATTRIBUTES"
				macro.Value("CATEGORY")					= sc
				macro.Value("DOCUMENT_FORMAT")			= "CUSTOM_TANDATA_USPS_LABEL_RTRN_ADDR.MMS"
				macro.Value("PRINT")					= True
				macro.Value("VIEW_DOCUMENT_MANAGER")	= False
				ScriptDataManager.AddMacro macro
				Set Macro = Nothing
			End If
		Next
	
	Else
		Dim package
		' msgbox "Running single package insert" '
		Set package = context( "CURRENT_PACKAGE" )
		COMPUTER                = package("COMPUTER")				' MACHINE NAME '
		SHIPPER                 = package("SHIPPER")				' SHIPPER SYMBOL '
		SHIPDATE                = package("SHIPDATE")				' SHIPDATE '
		ARRIVE_DATE             = package("ARRIVE_DATE")			' SHIPDATE '
		PRIMARY_SUBCATEGORY     = package("PRIMARY_SUBCATEGORY")	' SERVICE SYMBOL IN Server/Carrier/Service FORMAT '
		PRIMARY_CATEGORY		= package("PRIMARY_CATEGORY")		' SERVICE SYMBOL IN Server/Carrier/Service FORMAT '
		WEIGHT                  = package("WEIGHT")					' PACKAGE WEIGHT '
		DIMENSION               = package("DIMENSION")				' PACAKGE DIMENSION '
		PACKAGING               = package("PACKAGING")				' PACKAGING TYPE '
		TERMS                   = package("TERMS")					' PAYMENT TERMS '
		Set CONSIGNEE           = package("CONSIGNEE")				' CONSIGNEE '
		CONSIGNEE_COMPANY       = CONSIGNEE.Company					' CONSIGNEE COMPANY '
		CONSIGNEE_CONTACT       = CONSIGNEE.Contact					' CONSIGNEE CONTACT '
		CONSIGNEE_ADDRESS1      = CONSIGNEE.Address1				' CONSIGNEE ADDRESS1 '
		CONSIGNEE_ADDRESS2      = CONSIGNEE.Address2				' CONSIGNEE ADDRESS2 '
		CONSIGNEE_CITY          = CONSIGNEE.City					' CONSIGNEE CITY '
		CONSIGNEE_STATEPROVINCE = CONSIGNEE.StateProvince			' CONSIGNEE STATEPROVINCE '
		CONSIGNEE_POSTALCODE    = CONSIGNEE.PostalCode				' CONSIGNEE POSTAL CODE '
		CONSIGNEE_COUNTRY       = CONSIGNEE.CountrySymbol			' CONSIGNEE COUNTRY '
		COMMODITY_CLASS         = package("COMMODITY_CLASS")		' LTL COMMODITY CLASS '
		COD_AMOUNT              = package("COD_AMOUNT")				' COD AMOUNT '
		DESCRIPTION             = package("DESCRIPTION")			' PACAKGE DESCRIPTION '
		NOFN_SEQUENCE           = package("NOFN_SEQUENCE")			' CURRENT PACKAGE NUMBER '
		NOFN_TOTAL              = package("NOFN_TOTAL")				' NUMBER OF TOTAL PACKAGES '
		TRACKING_NUMBER         = package("TRACKING_NUMBER")		' PACKAGE TRACING NUMBER '
		PRIMARY_TOTAL           = package("PRIMARY_TOTAL")			' RATE TOTAL CHARGES '
		PRIMARY_BASE			= package("PRIMARY_BASE")			' RATE BASE CHARGES '
		PRIMARY_SPECIAL			= package("PRIMARY_SPECIAL")		' RATE FUEL SURCHARGE AND OTHER ACCESSORIAL FEES '
		FUEL_SURCHARGE			= package("FUEL_SURCHARGE")			' RATE FUEL SURCHARGE '
		MSN                     = package("MSN")					' PACKAGE MSN '
		ZONE                    = package("ZONE")					' PACKAGE ZONE'
		REFERENCE_1             = package("REFERENCE_1") 			' REFERENCE 1 '
		REFERENCE_2             = package("REFERENCE_2") 			' REFERENCE 2 '
		REFERENCE_5             = package("REFERENCE_5") 			' REFERENCE 5 '
		REFERENCE_6             = package("REFERENCE_6") 			' ORIGINAL REFERENCE 1 '
		REFERENCE_14            = package("REFERENCE_14") 			' REFERENCE 14 - PO NUMBER - FLAG FOR CASEMATE/WALMART ORDERS AND INVOICES FOR ASQ ORDERS '
		REFERENCE_15 			= package("REFERENCE_15")	' 7377 '	' SELECTED CARRIER '
		REFERENCE_16			= package("REFERENCE_16")	' 7377 '	' CARTON ID '
		
		' SWITCH BACK INVOICE NUMBER WITH WALMART PO NUMBER FOR CASEMATE ORDERS WITH A PO NUMBER '
		If ((REFERENCE_2 = "09" Or REFERENCE_2 = "2P") And REFERENCE_14 <> "") Then
			REF1_PO_NUM		= REFERENCE_1
			REF14_INV_NUM	= REFERENCE_14
			REFERENCE_1		= REF14_INV_NUM
			REFERENCE_6		= REF14_INV_NUM
			REFERENCE_14	= REF1_PO_NUM
		End If
		
		' THE FEDEX BARCODE IS OVER 30 CHARACTERS AND CONTAINS MORE INFORMATION THAN THE TRACKING NUMBER '
		' FOR THAT REASON WE ARE REPLACING package("BAR_CODE") WITH package("TRACKING_NUMBER") '
		BAR_CODE			= package("TRACKING_NUMBER")	' LABLE BARCODE '
		REFERENCE_3         = package("REFERENCE_3") 		' REF3 '
		REFERENCE_7         = package("REFERENCE_7") 		' ORIGINAL REF2 '
		REFERENCE_8         = package("REFERENCE_8") 		' ORIGINAL REF4 '
		REFERENCE_9         = package("REFERENCE_9") 		' ORIGINAL REF5 '
		REFERENCE_12        = package("REFERENCE_12") 		' DD PICK ID '
		REFERENCE_18        = package("REFERENCE_18") 		' WORLDEASE '
		
		' SCRUB DATA '
		REFERENCE_7 = Trim(REFERENCE_7)
		DESCRIPTION = Replace(DESCRIPTION  , "'", "")
		DESCRIPTION = Mid(DESCRIPTION,1,50)
		
		' TRIM DATA FIELDS '
		COMPUTER                = Mid(COMPUTER,1,10)				' MACHINE NAME '
		SHIPPER                 = Mid(SHIPPER,1,40)					' SHIPPER SYMBOL '
		PRIMARY_SUBCATEGORY     = Mid(PRIMARY_SUBCATEGORY,1,60)		' SERVICE SYMBOL IN Server/Carrier/Service FORMAT '
		PRIMARY_SUBCATEGORY     = UCase(PRIMARY_SUBCATEGORY)
		WEIGHT                  = Mid(WEIGHT,1,15)					' PACKAGE WEIGHT MyWord = UCase("Hello World") '
		DIMENSION               = Mid(DIMENSION,1,15)				' PACKAGE DIMENSIONS '
		PACKAGING               = Mid(PACKAGING,1,30)				' PACKAGING TYPE '
		TERMS                   = Mid(TERMS,1,30)					' PAMENT TERMS '
		CONSIGNEE_COMPANY       = Mid(CONSIGNEE_COMPANY,1,40)		' CONSIGNEE COMPANY '
		CONSIGNEE_CONTACT       = Mid(CONSIGNEE_CONTACT,1,40)		' CONSIGNEE CONTACT '
		CONSIGNEE_ADDRESS1      = Mid(CONSIGNEE_ADDRESS1,1,40)		' CONSIGNEE ADDRESS1 '
		CONSIGNEE_ADDRESS2      = Mid(CONSIGNEE_ADDRESS2,1,40)		' CONSIGNEE ADDRESS2 '
		CONSIGNEE_CITY          = Mid(CONSIGNEE_CITY,1,25)			' CONSIGNEE CITY '
		CONSIGNEE_STATEPROVINCE = Mid(CONSIGNEE_STATEPROVINCE,1,3)	' CONSIGNEE STATEPROVINCE '
		CONSIGNEE_POSTALCODE    = Mid(CONSIGNEE_POSTALCODE,1,12)	' CONSIGNEE POSTAL CODE'
		CONSIGNEE_COUNTRY       = Mid(CONSIGNEE_COUNTRY,1,60)		' CONSIGNEE COUNTRY '
		COMMODITY_CLASS         = Mid(COMMODITY_CLASS,1,3)			' LTL COMMODITY CLASS '
		COD_AMOUNT              = Mid(COD_AMOUNT,1,15)				' COD AMOUNT '
		DESCRIPTION             = Mid(DESCRIPTION,1,60)				' PACKAGE DESCRIPTION '
		NOFN_SEQUENCE           = Mid(NOFN_SEQUENCE,1,15)			' CURRENT PACKAGE NUMBER '
		NOFN_TOTAL              = Mid(NOFN_TOTAL,1,15)				' NUMBER OF TOTAL PACKAGES '
		TRACKING_NUMBER         = Mid(TRACKING_NUMBER,1,30)			' PACKAGE TRACKING NUMBER '
		PRIMARY_TOTAL           = Mid(PRIMARY_TOTAL,1,15)			' RATE TOTAL CHARGES '
		MSN                     = Mid(MSN,1,30)						' PACKAGE MSN '
		REFERENCE_3             = Mid(REFERENCE_3,1,30) 			' REF3 '
		REFERENCE_6             = Mid(REFERENCE_6,1,30) 			' ORIGINAL REF1 '
		REFERENCE_7             = Mid(REFERENCE_7,1,30) 			' ORIGINAL REF2 '
		REFERENCE_8             = Mid(REFERENCE_8,1,30) 			' ORIGINAL REF4 '
		REFERENCE_9             = Mid(REFERENCE_9,1,30) 			' ORIGINAL REF5 '
		REFERENCE_12            = Mid(REFERENCE_12,1,30) 			' DD PICK ID '
		
		' SCRUB DATA '
		REFERENCE_7 = Mid(REFERENCE_7,1,30)
		DESCRIPTION = Replace(DESCRIPTION  , "'", "")
		DESCRIPTION = Mid(DESCRIPTION,1,60)
		
		If PRIMARY_SUBCATEGORY = "TANDATA_LTL.AAA.AAA" And CONSIGNEE_COUNTRY = "CANADA" And REFERENCE_18 = "WORLDEASE" Then
			PRIMARY_TOTAL = (PRIMARY_TOTAL + 19.00)
		End If
		
		BASE_CHARGE = PRIMARY_BASE
		RESIDENTIAL_CHARGE = 0.00
		
		' ADDED TO SUBTRACT 10.00 FROM PRIMARY SPECIAL FOR CANADA WorldEase SHIPMENTS '
		If PRIMARY_SUBCATEGORY = "TANDATA_UPS.UPS.WSTD" And CONSIGNEE_COUNTRY = "CANADA" Then
			PRIMARY_SPECIAL = Csng(PRIMARY_SPECIAL - 10.00)
			PRIMARY_TOTAL = Csng(PRIMARY_TOTAL - 10.00)
		End If
		
		ACCESSORIAL_CHARGE = PRIMARY_SPECIAL - FUEL_SURCHARGE
		ACCESSORIAL_CHARGE = CSng(ACCESSORIAL_CHARGE)
		
		If CONSIGNEE.Residential = True and PRIMARY_CATEGORY = "TANDATA_UPS.UPS" And PRIMARY_SUBCATEGORY <> "TANDATA_UPS.UPS.SPPS" And CONSIGNEE_COUNTRY = "UNITED_STATES" Then
			If InStr(PRIMARY_SUBCATEGORY, "GND") > 0 Then
				If (BASE_CHARGE > ResidentialFee_UPSGND) Then
					RESIDENTIAL_CHARGE = ResidentialFee_UPSGND
				End If
			Else
				If (BASE_CHARGE > ResidentialFee_NonUPSGND) Then
					RESIDENTIAL_CHARGE = ResidentialFee_NonUPSGND
				End If
			End If
			BASE_CHARGE = PRIMARY_BASE - RESIDENTIAL_CHARGE
		End If
	End If
	
	' ADDED FOR ASQ DHL PACAKGES '
	If InStr(PRIMARY_SUBCATEGORY, "DHL") and REFERENCE_2 = "28"  Then
		BASE_CHARGE			= 0
		FUEL_SURCHARGE		= 0
		ACCESSORIAL_CHARGE	= 0
		PRIMARY_TOTAL		= 0
		RESIDENTIAL_CHARGE	= 0
		REFERENCE_6			= REFERENCE_14
		REFERENCE_1			= REFERENCE_14
	End If
	
	' FSC UPDATE '
	Dim newDate2
	Dim tempString2
	newDate2 = ""
	newDate2 = Year(SHIPDATE)
	tempString2 = Month(SHIPDATE)
	newDate2 = newDate2 & Right("0" & tempString2, 2)
	tempString2 = Day(SHIPDATE)
	newDate2 = newDate2 & Right("0" & tempString2, 2)
	
	If ARRIVE_DATE = "" Then
		ARRIVE_DATE = "00000000"
	Else
		Dim newARRIVE_DATE2
		newARRIVE_DATE2 = Year(ARRIVE_DATE)
		tempString2 = Month(ARRIVE_DATE)
		tempString2 = Right("0" & tempString2,2)
		newARRIVE_DATE2 = newARRIVE_DATE2 & tempString2
		tempString2 = Day(ARRIVE_DATE)
		tempString2 = Right("0" & tempString2,2)
		newARRIVE_DATE2 = newARRIVE_DATE2 & tempString2
		ARRIVE_DATE = newARRIVE_DATE2
	End If
	
	'********************************************************************************************************************'
    ' THE FOLLOWING SECTION IS USED FOR CREATING THE ACTUAL TRACKING USED TO TRACK SMARTPOST PACKAGES '
	If PRIMARY_SUBCATEGORY = "TANDATA_MISC.SMPT.SMPT" Then
		packageID	= TRACKING_NUMBER
		If shipper = "ALPHA" Then
			Duns	= "D02901001012"
			duns2	= "9102901001012"
		ElseIf shipper = "DULUTH" Then
			Duns	= "D02901001012"
			duns2	= "9102901001012"
		ElseIf shipper = "TOS" Then
			Duns	= "D02927007742"
			duns2	= "9102927007742"
		End If
		
		dunstrack = Duns2 & packageID
		Dim pos1b, pos2b, pos3b, pos4b, pos5b, pos6b, pos7b, pos8b, pos9b, pos10b, pos11b, pos12b, pos13b, pos14b, pos15b, pos16b, pos17b, pos18b, pos19b, pos20b, pos21b, chkdgtnumoddb, chkdgtnumevenb, chkdgtnumsumb, chkdgtnumb
		pos1b = Mid(dunstrack,21,1)
		pos2b = Mid(dunstrack,20,1)
		pos3b = Mid(dunstrack,19,1)
		pos4b = Mid(dunstrack,18,1)
		pos5b = Mid(dunstrack,17,1)
		pos6b = Mid(dunstrack,16,1)
		pos7b = Mid(dunstrack,15,1)
		pos8b = Mid(dunstrack,14,1)
		pos9b = Mid(dunstrack,13,1)
		pos10b = Mid(dunstrack,12,1)
		pos11b = Mid(dunstrack,11,1)
		pos12b = Mid(dunstrack,10,1)
		pos13b = Mid(dunstrack,9,1)
		pos14b = Mid(dunstrack,8,1)
		pos15b = Mid(dunstrack,7,1)
		pos16b = Mid(dunstrack,6,1)
		pos17b = Mid(dunstrack,5,1)
		pos18b = Mid(dunstrack,4,1)
		pos19b = Mid(dunstrack,3,1)
		pos20b = Mid(dunstrack,2,1)
		pos21b = Mid(dunstrack,1,1)
		pos1b = cint(pos1b)
		pos2b = cint(pos2b)
		pos3b = cint(pos3b)
		pos4b = cint(pos4b)
		pos5b = cint(pos5b)
		pos6b = cint(pos6b)
		pos7b = cint(pos7b)
		pos8b = cint(pos8b)
		pos9b = cint(pos9b)
		pos10b = cint(pos10b)
		pos11b = cint(pos11b)
		pos12b = cint(pos12b)
		pos13b = cint(pos13b)
		pos14b = cint(pos14b)
		pos15b = cint(pos15b)
		pos16b = cint(pos16b)
		pos17b = cint(pos17b)
		pos18b = cint(pos18b)
		pos19b = cint(pos19b)
		pos20b = cint(pos20b)
		pos21b= cint(pos21b)
		
		' USED TO CALCULATE THE CHECK DIGIT '
		chkdgtnumoddb	= (pos1b + pos3b + pos5b + pos7b + pos9b + pos11b + pos13b + pos15b+ pos17b + pos19b + pos21b) * 3
		chkdgtnumevenb	= pos2b + pos4b + pos6b + pos8b + pos10b + pos12b + pos14b + pos16b + pos18b + pos20b
		chkdgtnumsumb	= chkdgtnumoddb + chkdgtnumevenb
		chkdgtnumb		= 10 - Right(chkdgtnumsumb, 1)
		If chkdgtnumb = 10 Then
			chkdgtnumb = 0
		End If
		
		Dim TRACKING_NUMb, TRACKING_NUMBER2b
		TRACKING_NUMb = dunstrack & chkdgtnumb
		TRACKING_NUMb = Mid(TRACKING_NUMb, 3, 20)
		TRACKING_NUMBER2b = Mid(TRACKING_NUMb, 3, 20)
		
		' THIS IS THE ACUTAL NUMBER USED FOR SMART POST TRACKING '
		TRACKING_NUMBER = TRACKING_NUMBER2b
	End If
	
	If PRIMARY_SUBCATEGORY = "TANDATA_MISC.SMPT.SMPT" Then
		TRACKING_NUMBER = TRACKING_NUMb
	End If
	
	'**************************************************************************************************'
	
	' WRITE BACK A CUSTOM TRACKING NUMBER FOR UMI SHIPMENTS TO ASSIST WITH FREIGHT BILLING '
	' THE UMI TRACKING NUMBER TO BE WRITTEN BACK TO UPS100F WILL BE 9 - 10 CHARACTERS: REF1 ( 7 - 8 ) + REF2 (2) '
	If PRIMARY_SUBCATEGORY = "CONNECTSHIP_UPSMAILINNOVATIONS.UPS.ECO" Then
		If (SHIPPER = "ALPHA") Then
			TRACKING_NUMBER = "MI" & "001306" & currentPackage("REFERENCE_1")
		ElseIf (SHIPPER ="DULUTH") Then
			TRACKING_NUMBER = "MI" & "003195" & currentPackage("REFERENCE_1")
		ElseIf (SHIPPER ="PHILLY") Then
			TRACKING_NUMBER = "MI" & "008601" & currentPackage("REFERENCE_1")
		ElseIf (SHIPPER ="WMU") Then
			TRACKING_NUMBER = "MI" & "012656" & currentPackage("REFERENCE_1")
		ElseIf (SHIPPER ="JONES") Then
			TRACKING_NUMBER = "MI" & "007173" & currentPackage("REFERENCE_1")
		ElseIf (SHIPPER ="WASHDC") Then
			TRACKING_NUMBER = "MI" & "000000" & currentPackage("REFERENCE_1")
		Else
			TRACKING_NUMBER = currentPackage("REFERENCE_1") & "." &  currentPackage("REFERENCE_2")
		End If
	End If
	
	If TRACKING_NUMBER = "" Then
		TRACKING_NUMBER = "n/a"
	End If
	
	'  UPDATE DATABASE '
    Dim sql2
    If (1=1) Then
		sql2 =       "INSERT into " & library & "UPS100F"
		sql2 = sql2 + " (CS_Workstation, SHIPPER_REFERENCE, CONSIGNEE_REFERENCE,"
		sql2 = sql2 + " REF_1, REF_2, REF_3,"
		sql2 = sql2 + " SHIPDATE, DIMENSION, TOTAL,"
		sql2 = sql2 + " WEIGHT, TTL_FREIGHT, TTL_WEIGHT,"
		sql2 = sql2 + " DESCRIPTION, CURRENT_PACKAGE, TOTAL_PACKAGES,"
		sql2 = sql2 + " MSN, SERVICE, PACKAGING, COMPANY,"
		sql2 = sql2 + " CONTACT, ADDRESS1, ADDRESS2,"
		sql2 = sql2 + " ADDRESS3, CITY, STATEPROVINCE,"
		sql2 = sql2 + " POSTALCODE, COUNTRYSYMBOL, TRACKING_NUMBER,"
		sql2 = sql2 + " TERMS, SHIPPER, ARRIVE_DATE,"
		sql2 = sql2 + " COD_AMOUNT, CODE, UPINVOICE, COMMODITY_CLASS, BASE_CHARGE, RESIDENTIAL_CHARGE, FUEL_SURCHARGE, ACCESSORIAL_CHARGE, UPZONE )"
		sql2 = sql2 + " values ('" & Left(COMPUTER,10) & "' , '" & REFERENCE_6 & "' , '" & REFERENCE_7
		sql2 = sql2 + "' , '" & REFERENCE_3 & "' , '" & REFERENCE_8 & "' , '" & REFERENCE_9
		sql2 = sql2 + "' , '" & newDate2 & "' , '" & DIMENSION & "' , '" & PRIMARY_TOTAL
		sql2 = sql2 + "' , '" & WEIGHT & "' , '" & PRIMARY_TOTAL & "'  , '" & WEIGHT
		sql2 = sql2 + "' , '" & DESCRIPTION & "' , '" & NOFN_SEQUENCE & "' , '" & NOFN_TOTAL
		sql2 = sql2 + "' , '" & MSN & "' , '" & PRIMARY_SUBCATEGORY & "' , '" & PACKAGING
		sql2 = sql2 + "' , '" & Replace(CONSIGNEE_COMPANY,"'","|") & "' , '" & Replace(CONSIGNEE_CONTACT,"'","|") & "' , '" & Replace(CONSIGNEE_ADDRESS1,"'","|")
		sql2 = sql2 + "' , '" & Replace(CONSIGNEE_ADDRESS2,"'","|") & "' , '" & Replace(CONSIGNEE_ADDRESS2,"'","|") & "' , '" & Replace(CONSIGNEE_CITY,"'","|")
		sql2 = sql2 + "' , '" & Left(CONSIGNEE_STATEPROVINCE,3) & "' , '" & CONSIGNEE_POSTALCODE & "' , '" & CONSIGNEE_COUNTRY
		sql2 = sql2 + "' , '" & TRACKING_NUMBER & "' , '" & TERMS & "', '" & SHIPPER
		sql2 = sql2 + "' , '" & ARRIVE_DATE & "' , '" & COD_AMOUNT & "' , '" & BAR_CODE & "' , '" & REFERENCE_18
		sql2 = sql2 + "' , '" & Left(COMMODITY_CLASS,3) & "'," & BASE_CHARGE & "," & RESIDENTIAL_CHARGE & "," & FUEL_SURCHARGE & "," & ACCESSORIAL_CHARGE & ", '" & ZONE & "')"
		writeDebug "SQL single package-" & sql2
		
		' 7377 SINGLE PACKAGE INSERT '
		If ScriptDataManager.StoredData("AMAZON") = "Y" Then
			sql105f = "INSERT into " & library & "UPS105F"
			sql105f = sql105f + " (UPMSN, UPNAM, UPSEQ, UPVAL, UPSTK#) VALUES ('" & MSN & "', 'SSCC18', '" & upseq & "', '" & REFERENCE_16 & "', '" & TRACKING_NUMBER & "')"
			dbConn.Execute sql105f
			sql105f2 = "INSERT into " & library & "UPS105F"
			sql105f2 = sql105f2 + " (UPMSN, UPNAM, UPSEQ, UPVAL, UPSTK#) VALUES ('" & MSN & "', 'SCAC', '" & upseq & "', '" & REFERENCE_15 & "', '" & TRACKING_NUMBER & "')"
			dbConn.Execute sql105f2
			upseq = upseq + 1
		End If
		
		Dim successFlag2
		successFlag2 = False		
		Dim rs22, loopCount2
		loopCount2 = 1
		Do Until loopCount2 = 5
			dbConn.Execute sql2
			' IF SQL INSERT DOES NOT ERROR A CHECK WILL BE MADE '
			' IF CHECK DOES NOT FIND DATA THEN SQL WRITTEN TO DEBUG '
			Dim dbError2
			dbError2 = ""
			If dbConn.Errors.Count > 0 Then
				writeDebug "SQL-" & sql2
				Dim intcount2
				Dim dbCommError2
				For intCount = 0 To dbConn.Errors.Count - 1
					Set dbCommError = dbConn.Errors.Item(intCount)
					writeDebug dbcommerror.Description
					dbError2 = dbError2 + dbCommError.Description + " ** "
				Next
				writeFailedRecord "SQL-" & sql2
				writeDebug "SQL-" & sql2
				setError(dbError2)
			End If
			Dim chksql2b
			chksql2b = "SELECT * FROM " & library & "UPS100F WHERE MSN = '" & MSN & "'"
			Set rs22 = dbConn.Execute (chksql2b)
			If (rs22.eof) Then
				writeDebug "SQL Insert did not commit for MSN: " & MSN
				writeDebug sql2
			Else
				writedebug "SQL Insert committed for MSN: " & MSN
				successFlag2 = True
				Exit Do
			End If
			
			loopCount2 = loopCount2 + 1
			If loopCount2 = 5 Then
				Dim record2
				record2 =	Left(COMPUTER,10) & "," & REFERENCE_6 & "," & REFERENCE_7
				record2 = record2 + "," & REFERENCE_3 & "," & REFERENCE_8 & "," & REFERENCE_9
				record2 = record2 + "," & newDate & "," & DIMENSION & "," & PRIMARY_TOTAL
				record2 = record2 + "," & WEIGHT & "," & PRIMARY_TOTAL & "  , " & WEIGHT
				record2 = record2 + "," & DESCRIPTION & "," & NOFN_SEQUENCE & "," & NOFN_TOTAL
				record2 = record2 + "," & MSN & "," & PRIMARY_SUBCATEGORY & "," & PACKAGING
				record2 = record2 + "," & Replace(CONSIGNEE_COMPANY,"","|") & "," & Replace(CONSIGNEE_CONTACT,"","|") & "," & Replace(CONSIGNEE_ADDRESS1,"","|")
				record2 = record2 + "," & Replace(CONSIGNEE_ADDRESS2,"","|") & "," & Replace(CONSIGNEE_ADDRESS2,"","|") & "," & Replace(CONSIGNEE_CITY,"","|")
				record2 = record2 + "," & Left(CONSIGNEE_STATEPROVINCE,3) & "," & CONSIGNEE_POSTALCODE & "," & CONSIGNEE_COUNTRY
				record2 = record2 + "," & TRACKING_NUMBER & "," & TERMS & ", " & SHIPPER
				record2 = record2 + "," & ARRIVE_DATE & "," & COD_AMOUNT & "," & BAR_CODE & "' , '" & REFERENCE_18
				record2 = record2 + "," & Left(COMMODITY_CLASS,3) & "," & BASE_CHARGE & "," & RESIDENTIAL_CHARGE & "," & FUEL_SURCHARGE & "," & ACCESSORIAL_CHARGE & "," & ZONE
				writeFailedRecord record2
				writeDebug record2
				Exit Do
			End If
			Loop
			If (successFlag2 = False) Then
				writeInsertFile sql2
			End If
		End If
		
		' RESET THE DOCUMENT ASSIGNMENTS OVERRIDEN IN THE PRE_SHIP SCRIPT '
		Dim macro2, sc2
		sc2 = currentPackage("PRIMARY_CATEGORY")
		If PRIMARY_SUBCATEGORY = "TANDATA_USPS.USPS.SPCL" And REFERENCE_2 = "44" Then
			Set macro2 = CreateObject("Progistics.Dictionary")
			macro2.Value("NAME") = "MACRO_MODIFY_DOCUMENT_ATTRIBUTES"
			macro2.Value("CATEGORY") = sc2
			macro2.Value("DOCUMENT_FORMAT") = "CUSTOM_TANDATA_USPS_LABEL.MMS"
			macro2.Value("PRINT") = True
			macro2.Value("VIEW_DOCUMENT_MANAGER") = False
			ScriptDataManager.AddMacro macro2
			Set Macro2 = Nothing
		Else
			Set macro2 = CreateObject("Progistics.Dictionary")
			macro2.Value("NAME") = "MACRO_MODIFY_DOCUMENT_ATTRIBUTES"
			macro2.Value("CATEGORY") = sc2
			macro2.Value("DOCUMENT_FORMAT") = "CUSTOM_TANDATA_USPS_LABEL_RTRN_ADDR.MMS"
			macro2.Value("PRINT") = True
			macro2.Value("VIEW_DOCUMENT_MANAGER") = False
			ScriptDataManager.AddMacro macro2
			Set Macro2 = Nothing
		End If
	End If
  
  'Global variable for storing flag to be used in PRE_SHIP for setting custom field value'
  ScriptDataManager.StoredData("setCustomField") = True
End Sub

'****************************************************************************************************************************************************************************************************'

Sub setError( buf )
	Dim macro
	Set macro = CreateObject("Progistics.Dictionary")
	macro.Value("NAME") = "MACRO_SET_PACKAGE_ERROR"
	macro.Value("ERROR_MESSAGE") = buf
	ScriptDataManager.AddMacro macro
	Set macro = Nothing
End Sub

'****************************************************************************************************************************************************************************************************'

Sub setField( fieldName, fieldValue, suppressScript, fieldIndex )
	Dim macro
	Set macro = CreateObject("Progistics.Dictionary")
	macro.Value("NAME") = "MACRO_SET_FIELD"
	macro.Value("FIELD_NAME") = fieldName
	macro.Value("FIELD_VALUE") = fieldValue
	
	If fieldIndex > 0 Then
		macro.Value("FIELD_INDEX") = fieldIndex
	End If
	
	macro.Value("SUPPRESS_SCRIPT") = suppressScript
	ScriptDataManager.AddMacro macro
	Set macro = Nothing
End Sub

'****************************************************************************************************************************************************************************************************'

Sub DebugLog( buf, pri )
	setField "REFERENCE_20", buf, True, 0
End Sub

'****************************************************************************************************************************************************************************************************'

Function writeDebug(msg)
	Dim debugfilepath,debugFso,debugOut,strHeader,context,currentPackage
	debugfilepath = "C:\SQLDebug.txt"
	Set debugFso = CreateObject("Scripting.FileSystemObject")
	Set debugOut = debugFso.OpenTextFile(debugFilePath, 8, True)
	
	debugOut.WriteLine "Debuging at Postship " & Now()
	debugOut.WriteLine "Insert Query is: " & msg
	debugOut.WriteLine
	
	' CLEANUP OBJECTS '
	
	Set debugOut = Nothing
	Set debugFso = Nothing
End Function

'****************************************************************************************************************************************************************************************************'

Function writeFailedRecord(msg)
	Dim debugfilepath,debugFso,debugOut,strHeader
	Set debugFso = CreateObject("Scripting.FileSystemObject")
	debugfilepath = "C:\CSWFailedRecords.csv"
	If debugFso.FileExists(debugfilepath) = False Then
		Set debugOut = debugFso.OpenTextFile(debugFilePath, 8, True)
		strHeader =  "CS_Workstation,SHIPPER_REFERENCE,CONSIGNEE_REFERENCE,"
		strHeader = strHeader + "REF_1,REF_2,REF_3,"
		strHeader = strHeader + "SHIPDATE,DIMENSION,TOTAL,"
		strHeader = strHeader + "WEIGHT,TTL_FREIGHT,TTL_WEIGHT,"
		strHeader = strHeader + "DESCRIPTION,CURRENT_PACKAGE,TOTAL_PACKAGES,"
		strHeader = strHeader + "MSN,SERVICE,PACKAGING,COMPANY,"
		strHeader = strHeader + "CONTACT,ADDRESS1,ADDRESS2,"
		strHeader = strHeader + "ADDRESS3,CITY,STATEPROVINCE,"
		strHeader = strHeader + "POSTALCODE,COUNTRYSYMBOL,TRACKING_NUMBER,"
		strHeader = strHeader + "TERMS,SHIPPER,ARRIVE_DATE,"
		strHeader = strHeader + "COD_AMOUNT,CODE,UPINVOICE,COMMODITY_CLASS,BASE_CHARGE,RESIDENTIAL_CHARGE,FUEL_SURCHARGE,ACCESSORIAL_CHARGE,UPZONE"
		debugOut.Writeline strHeader
		debugOut.WriteLine msg
	Else
	
	Set debugOut = debugFso.OpenTextFile(debugFilePath, 8, False)
	debugOut.WriteLine msg
	End If
	
	Set debugOut = Nothing
	Set debugFso = Nothing
End Function

'****************************************************************************************************************************************************************************************************'

Sub writeInsertFile(msg)
	Dim debugFilePath
	debugFilePath = "C:\CSW_Failed_Transactions\CSW_FailedInsert.sql"
	' IF FOR ANY REASON THE RECORD FAILS, A RECORD IS INSERTED TO THE FILE BELOW '
	On Error Resume Next
	Dim debugFso,debugOut,strHeader
	Set debugFso = CreateObject("Scripting.FileSystemObject")
	If debugFso.FileExists(debugfilepath) = False Then
		Set debugOut = debugFso.OpenTextFile(debugFilePath, 8, True)
	Else
	Set debugOut = debugFso.OpenTextFile(debugFilePath, 8, False)
	debugOut.WriteLine msg
	End If
	debugOut.WriteLine msg
	
	' DISPLAY A VISUAL PROMPT TO DISTRIBUTION FOR IT NOTIFICATION '
	Dim errorPromptString
	errorPromptString = "The package IS SHIPPED but not logged in Oracle." & vbLf & "Please notify PBD IT if this error continues." & vbLf & "Additional log attempts will be attempted at 6:45pm and 9:45pm EST."
	dim cswUtil
	Set cswUtil = CreateObject("CSWUtils.Utils")
	cswUtil.Inform errorPromptString, "**ORACLE LOG ERROR**", 64
	
	Set csutils = Nothing
	Set debugOut = Nothing
	Set debugFso = Nothing
End Sub

'****************************************************************************************************************************************************************************************************'

sub screenmsg (msg)
	dim cswUtil
	Set cswUtil = CreateObject("CSWUtils.Utils")
	cswUtil.Inform msg, "**misc message** Shipment Log Error", 64
	Set csutils = Nothing
end sub

'****************************************************************************************************************************************************************************************************'