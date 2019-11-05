Set fso = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")

sysDrvLetter = wshShell.ExpandEnvironmentStrings("%SystemDrive%")
filespec = sysDrvLetter & "\go.txt"

WScript.Echo "Waiting for " & filespec

'Check every 1 second to see if the file exists
While Not fso.FileExists(filespec)
	Wscript.Sleep 1000
Wend

'When it does exist, delete it and go on
fso.DeleteFile(filespec)
