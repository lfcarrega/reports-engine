#Requires AutoHotkey v2.0

exe_name := "nano_reports.exe"
exe_icon := "main.ico"

source := "main.ahk"
compiler := "..\..\_ahk\compiler\ahk2exe.exe"
c_base := "..\..\_ahk\v2\autohotkey64.exe"

If (FileExist(source) && FileExist(compiler)) {
	If FileExist(exe_icon)
		RunWait compiler " /in `"" source "`" /base " c_base " /icon " exe_icon " /out ..\" exe_name
	Else 
		RunWait compiler " /in `"" source "`" /base " c_base " /out ..\" exe_name
}