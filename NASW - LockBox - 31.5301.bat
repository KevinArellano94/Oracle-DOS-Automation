@if (@CodeSection == @Batch) @then

@echo off
    set SendKeys=CScript //nologo //E:JScript "%~F0"
    cls
    color 0a
    REM :loop
		timeout /t 2 /nobreak >nul
		set Company="31"
		set Member="NASWAGN"
		
		set PO="L4886683"
		set item_1_quantity="1"
		set item_1="31.5301"
		set Billing="12117435"
		set Shipping="12133410"
		set "Order_By=."
		set Shipping_Method="U09"
		set payment="3013259"
		set cost="35.15"
		
		set Email=""
		set Phone=""
		set First_Name=""
		set Last_Name=""
		set "Address_2="
		set Zip=""
	   
	   %SendKeys% "+{F10}"
        %SendKeys% "+{TAB}"
	   %SendKeys% "+{TAB}"
	   %SendKeys% "+{TAB}"
	   %SendKeys% "%Company%"
	   %SendKeys% "{ENTER}"
	   %SendKeys% "%Billing%"
	   %SendKeys% "{TAB}"
	   %SendKeys% "%Shipping%"
REM	   %SendKeys% "^v"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "%PO%"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "I"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "%Shipping_Method%"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "%payment%"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% ""
	   %SendKeys% "{TAB}"
	   %SendKeys% "N"
	   %SendKeys% "{ENTER}"
	   %SendKeys% "%Order_By%"
REM	   %SendKeys% "%First_Name% %Last_Name%"
	   %SendKeys% "{ENTER}"
	   %SendKeys% "{ENTER}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "%item_1%"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "%item_1_quantity%"
	   %SendKeys% "{TAB}"
	   %SendKeys% "%cost%"
	   %SendKeys% "{ENTER}"
	   %SendKeys% "{ENTER}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "4"
	   %SendKeys% "{ENTER}"
	   %SendKeys% "I"
	   %SendKeys% "{PGUP}"
	   %SendKeys% "{PGDN}"
	   %SendKeys% "{ENTER}"
	   %SendKeys% "{ENTER}"
	   %SendKeys% "{PGDN}"
	   %SendKeys% "{DOWN}"
	   %SendKeys% "{DOWN}"
	   %SendKeys% "{DOWN}"
	   %SendKeys% "{DOWN}"
	   %SendKeys% "{DOWN}"
	   %SendKeys% "{DOWN}"
	   %SendKeys% "{DOWN}"
	   %SendKeys% "{DOWN}"
	   %SendKeys% "{DOWN}"
	   %SendKeys% "{DOWN}"
	   %SendKeys% "{DOWN}"
	   %SendKeys% "+{TAB}"
	   
	   %SendKeys% "CK"
	   %SendKeys% "%payment%"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "1"
	   %SendKeys% "{TAB}"
	   %SendKeys% "%cost%"
	   %SendKeys% "{TAB}"
	   %SendKeys% "{TAB}"
	   %SendKeys% "P"
    REM goto :loop

@end

var WshShell = WScript.CreateObject("WScript.Shell");
WshShell.SendKeys(WScript.Arguments(0));