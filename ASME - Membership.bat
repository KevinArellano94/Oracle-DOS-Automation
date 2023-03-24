@if (@CodeSection == @Batch) @then

@echo off
    set SendKeys=CScript //nologo //E:JScript "%~F0"
    cls
    color 0a
    REM :loop
	timeout /t 2 /nobreak >nul
	set "Last_Name="
	set "First_Name="
	set "Middle_Name="
	Set "Gender=M"
	Set "BirthDate="
	set "Email_Address="
	set "DivisionInstitute_1="
	
	
	set "Company="
	set "Address_1=Street Drive Pkwy 123"
	set "Address_2="
	set "Address_3=EEEE"
	set "Address_4=Sector E12 4"
	set "City="
	set "State="
	set "Country="
	set "Postal_Code=00000"
	
	set "Phone_Business="
	set "Cell="
	set "Fax="
	
	set "DivisionInstitute_1=01"
	set "DivisionInstitute_2=14"
	set "DivisionInstitute_3=22"
	set "DivisionInstitute_4=16"
	set "DivisionInstitute_5=21"
	
	set "Job=02"
	set "Title=34"
REM	set "Title=None Specified"
	
	set "Degree=BS"
	set "Degree_Type=D"
	set "Institution=000010000000"
	set "Begin_Date=8.1.2008"
	set "End_Date=5.1.2010"
	set "Major=Mechanical Engineer"
	
	timeout /t 1 /nobreak >nul
	
	%SendKeys% "+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}"
	%SendKeys% "+{TAB}+{TAB}+{TAB}"		REM For US Only
	REM %SendKeys% "+{TAB}%Country%"
	REM %SendKeys% "+{TAB}%City%"
	REM %SendKeys% "+{TAB}%State%"
	%SendKeys% "+{TAB}%Postal_Code%"
	%SendKeys% "+{TAB}"
	
	timeout /t 3 /nobreak >nul
	
	REM %SendKeys% "+{TAB}%Address_4%"
	REM %SendKeys% "+{TAB}%Address_3%"
	%SendKeys% "+{TAB}%Address_2%"
	%SendKeys% "+{TAB}%Address_1%"
	
	%SendKeys% "+{TAB}+{TAB}+{TAB}+{TAB}%Company%"
	%SendKeys% "+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}%Last_Name%"
	%SendKeys% "+{TAB}%Middle_Name%"
	%SendKeys% "+{TAB}%First_Name%"
	%SendKeys% "+{TAB}+{TAB}H"
	
	timeout /t 1 /nobreak >nul
	
	%SendKeys% "^{TAB}{TAB}{TAB}%Phone_Business%"
	%SendKeys% "{TAB}{TAB}{TAB}%Cell%"
	%SendKeys% "{TAB}{TAB}{TAB}%Fax%"
	%SendKeys% "{TAB}{TAB}%Email_Address%"
	
	timeout /t 8 /nobreak >nul
	
	%SendKeys% "{TAB}{TAB}{TAB}%Job%"
	%SendKeys% "{TAB}{TAB}%Gender%"
	%SendKeys% "{TAB}{TAB}%BirthDate%"
	%SendKeys% "^s"
	
	timeout /t 8 /nobreak >nul
	
	%SendKeys% "%Degree%"
	%SendKeys% "{TAB}%Degree_Type%"
	%SendKeys% "{TAB}%Institution%{TAB}"
	%SendKeys% "{TAB}%Begin_Date%"
	%SendKeys% "{TAB}%End_Date%"
	%SendKeys% "{TAB}{TAB}{TAB}"
REM		%SendKeys% " "
	%SendKeys% "{TAB}%Major%"
	%SendKeys% "^s"
	
	timeout /t 4 /nobreak >nul
	
	%SendKeys% " "
	%SendKeys% "{TAB}+{TAB}"
	%SendKeys% "%DivisionInstitute_1%"
	%SendKeys% "{TAB}%DivisionInstitute_2%"
	%SendKeys% "{TAB}%DivisionInstitute_3%"
	%SendKeys% "{TAB}%DivisionInstitute_4%"
	%SendKeys% "{TAB}%DivisionInstitute_5%"
	%SendKeys% "^s"
	
	timeout /t 5 /nobreak >nul
	
	%SendKeys% "{TAB}{TAB}%Title%"
	%SendKeys% "{TAB}%Job%"
	%SendKeys% "^s"
	
	timeout /t 10 /nobreak >nul
	
	%SendKeys% "{TAB}{TAB}E"
	
	timeout /t 3 /nobreak >nul
	
	timeout /t 3 /nobreak >nul
	
	%SendKeys% "MMCM{TAB}"
	
	timeout /t 3 /nobreak >nul
	
	%SendKeys% "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}"
	%SendKeys% "{TAB}"
	%SendKeys% "^s"
    REM goto :loop

@end

var WshShell = WScript.CreateObject("WScript.Shell");
WshShell.SendKeys(WScript.Arguments(0));