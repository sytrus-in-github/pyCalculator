Set objShell = WScript.CreateObject("WScript.Shell")

prompt = "Do you want to create desktop shortcut for:" + vbcrlf + "1. all users" + vbcrlf +"2. current user" + vbcrlf + "0. Don't create desktop shortcut" + vbcrlf
'strInput = UserInput( ":" )
If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
    WScript.StdOut.Write prompt
    strInput = WScript.StdIn.ReadLine
Else
    strInput = InputBox( prompt )
End If

if Left( strInput, 1) = "1" then
	desktopDir = objShell.SpecialFolders("AllUsersDesktop")
elseif Left( strInput, 1) = "2" then
	desktopDir = objShell.SpecialFolders("Desktop")
else
	wscript.quit
end if

Set objShortCut = objShell.CreateShortcut(desktopDir & "\py_calc.lnk")
objShortCut.TargetPath =  replace( WScript.ScriptFullName, WScript.ScriptName, "" ) & "\py_calc.bat"
objShortcut.IconLocation = "%SystemRoot%\system32\calc.exe"
objShortCut.Description = "A python based wrap as a calculator."
objShortCut.Save