Set objShell = CreateObject("WScript.Shell")
strArgs = "-WindowStyle Hidden -File """ & objShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents\OMMigrate\Obsidian-Addon-Opener.ps1"" """ & WScript.Arguments(0) & """"
objShell.Run "powershell.exe " & strArgs, 0, False
