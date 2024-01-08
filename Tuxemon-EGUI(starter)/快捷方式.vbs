Set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
set oShellLink = WshShell.CreateShortcut(strDesktop & "/Tuxemon.lnk")
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
StrCurPath = fso.GetFolder(".")
oShellLink.TargetPath = StrCurPath+"\Tuxemon.exe"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = StrCurPath+"\icon\tuxemon.ico, 0"
oShellLink.Description = "快捷方式"
oShellLink.WorkingDirectory = StrCurPath
oShellLink.Save