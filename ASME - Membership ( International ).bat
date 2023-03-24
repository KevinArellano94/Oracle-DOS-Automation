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
	
	set "DivisionInstitute_2="
	set "DivisionInstitute_3="
	set "DivisionInstitute_4="
	set "DivisionInstitute_5="
	
	set "Job="
	set "Title="
REM	set "Title=None Specified"
	
	set "Degree=BA"
	set "Degree_Type=A"
	set "Institution=000000000000"
	set "Begin_Date=1.1.1994"
	set "End_Date=1.1.1994"
	set "Major=Engineer"
	
	timeout /t 1 /nobreak >nul
	
	%SendKeys% "+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}"
REM		%SendKeys% "+{TAB}+{TAB}+{TAB}"		REM For US Only
	%SendKeys% "+{TAB}%Country%"
	%SendKeys% "+{TAB}%City%"
	%SendKeys% "+{TAB}%State%"
	%SendKeys% "+{TAB}%Postal_Code%"
	
REM		timeout /t 3 /nobreak >nul
	
	%SendKeys% "+{TAB}%Address_4%"
	%SendKeys% "+{TAB}%Address_3%"
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
	%SendKeys% "+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}"
	%SendKeys% "+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}+{TAB}"
	%SendKeys% "+{TAB}+{TAB}+{TAB}"
	%SendKeys% " "
	
	timeout /t 4 /nobreak >nul
	
	%SendKeys% "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}"
	
	timeout /t 3 /nobreak >nul
	
	%SendKeys% "{TAB}{TAB}{TAB}%Job%"
	%SendKeys% "{TAB}{TAB}%Gender%"
	%SendKeys% "{TAB}{TAB}%BirthDate%"
	%SendKeys% "^s"
	
	timeout /t 5 /nobreak >nul
	
	%SendKeys% "%Degree%"
	%SendKeys% "{TAB}%Degree_Type%"
	%SendKeys% "{TAB}%Institution%{TAB}"
	%SendKeys% "{TAB}%Begin_Date%"
	%SendKeys% "{TAB}%End_Date%"
	%SendKeys% "{TAB}{TAB}{TAB}"
	%SendKeys% " "
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
	
REM		timeout /t 5 /nobreak >nul
	
REM		%SendKeys% "{TAB}{TAB}%Title%"
REM		%SendKeys% "{TAB}%Job%"
REM		%SendKeys% "^s"
	
	timeout /t 6 /nobreak >nul
	
	%SendKeys% "{TAB}{TAB}E"
	%SendKeys% "{TAB}"
	
REM		timeout /t 3 /nobreak >nul
	
REM		%SendKeys% "Name"
REM		%SendKeys% "{TAB}Phone"
REM		%SendKeys% "{TAB}Email"
	
	timeout /t 3 /nobreak >nul
	
	%SendKeys% "MMSTFR{TAB}"
	
	timeout /t 3 /nobreak >nul
	
	%SendKeys% "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}"
	%SendKeys% "GSM{TAB}"
REM		timeout /t 1 /nobreak >nul
REM		%SendKeys% "{TAB} "
REM		timeout /t 1 /nobreak >nul
REM		%SendKeys% "{TAB} "
REM		timeout /t 1 /nobreak >nul
REM		%SendKeys% "{TAB} "
REM		timeout /t 1 /nobreak >nul
REM		%SendKeys% "{TAB} "
REM		timeout /t 1 /nobreak >nul
REM		%SendKeys% "{TAB} "
	%SendKeys% "^s"
REM goto :loop

@end

var WshShell = WScript.CreateObject("WScript.Shell");
WshShell.SendKeys(WScript.Arguments(0));