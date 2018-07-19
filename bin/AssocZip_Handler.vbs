' Обработчик ассоциации Zip для файлов quarantine*.zip, virusinfo_auto_*.zip, ZOO_*.zip

Option Explicit
Dim oShell, AppPath, File, Name, pos, AppName, DefProgID, DefVerb, DefHash, OSVer, IsWin8AndNewer, IsWinXP, IsWin7, oShellApp, curVerb, LaunchCom, bVerbOpenBased

AppName = "AVZ DeQuarantine"

Set oShell     = CreateObject("WScript.Shell")

File = WScript.Arguments(0)

pos = instrrev(File, "\")
if pos <> 0 then
	Name = mid(File, pos + 1)
	if instr(1, Name, "quarantine", 1) <> 0 or StrBeginWith(Name, "virusinfo_auto_") or StrBeginWith(Name, "ZOO_") then
		AppPath = oShell.SpecialFolders("AppData") & "\AVZ DeQuarantine\AVZ - распаковать карантин.cmd"
		oShell.Run "cmd.exe" & " " & "/c" & " " & """" & """" & AppPath & """" & " " & """" & File & """" & """", 1, false
		WScript.Quit
	end if
end if

On Error Resume Next
OSVer = CSng(Replace(oShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion"),".",","))
On Error Goto 0
IsWin8AndNewer = (OSVer > 6.1)
IsWinXP = (OSVer < 6)
IsWin7 = (OSVer >= 6 and OSVer <= 6.1)

' revert to default app class
if IsWin8AndNewer then
	DefHash = WScript.Arguments(1)
	DefVerb = WScript.Arguments(2)
else
	DefProgID = WScript.Arguments(1)
end if

On Error Resume Next
if 0 = Len(DefProgID) and 0 = Len(DefHash) then
	if IsWinXP then
		oShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\Application"
	elseif IsWin7 then
		oShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\ProgID"
	elseif IsWin8AndNewer then
		DefProgID = oShell.RegRead("HKCR\.zip\")
		if DefProgID <> "" then
			curVerb = oShell.RegRead("HKCU\Software\Classes\" & DefProgID & "\shell\")
			if DefVerb <> "" then
				oShell.RegWrite "HKCU\Software\Classes\" & DefProgID & "\shell\", DefVerb, "REG_SZ"
			else
				if curVerb = "" then
					bVerbOpenBased = true
					LaunchCom = oShell.RegRead ("HKCU\Software\Classes\" & DefProgID & "\shell\open\command\")
					oShell.RegDelete ("HKCU\Software\Classes\" & DefProgID & "\shell\open\command\")
					oShell.RegDelete ("HKCU\Software\Classes\" & DefProgID & "\shell\open\")
				else
					LaunchCom = oShell.RegRead ("HKCU\Software\Classes\" & DefProgID & "\shell\AVZ.DeQuarantine\command\")
					oShell.RegDelete ("HKCU\Software\Classes\" & DefProgID & "\shell\AVZ.DeQuarantine\command\")
					oShell.RegDelete ("HKCU\Software\Classes\" & DefProgID & "\shell\AVZ.DeQuarantine\")				
				end if
			end if
		end if
	end if
else
	if IsWinXP then
		if StrBeginWith(DefProgID, "Applications") then
			oShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\Application", Mid(DefProgID, Len("Applications")+2), "REG_SZ"
		else
			oShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\ProgID", DefProgID, "REG_SZ"
		end if
	elseif IsWin7 then
		oShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\ProgID", DefProgID, "REG_SZ"
	elseif IsWin8AndNewer then
		oShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\Hash", DefHash, "REG_SZ"		
	end if
end if
'if Err.Number <> 0 then
'	Msgbox "AVZ.Dequarantine не удалось выполнить запись в " & "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\ProgID" & _
'		". Возможно, запись заблокирована защитным ПО. Файл " & File & " не может быть открыт. Внесите AVZ.Dequarantine в исключения или переустановите программу.", vbCritical
'else
	if IsWinXP then
		Set oShellApp = CreateObject("Shell.Application")
		oShellApp.ShellExecute File, "", "", "", 1
	else
		oShell.Run "rundll32.exe zipfldr.dll,RouteTheCall" & " " & """" & File & """", 1, true
	end if
'end if

WScript.Sleep 2000

' write back own app class
if IsWinXP then
	oShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\Application", "AVZ.DeQuarantine", "REG_SZ"
	oShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\ProgID"
elseif IsWin7 then
	oShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\ProgID", "Applications\AVZ.DeQuarantine", "REG_SZ"
elseif IsWin8AndNewer then
	oShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\Hash"
	if DefProgID <> "" then
		oShell.RegWrite "HKCR\" & DefProgID & "\shell\", "AVZ.DeQuarantine", "REG_SZ"
		if LaunchCom <> "" then
			if bVerbOpenBased then
				oShell.RegWrite "HKCU\Software\Classes\" & DefProgID & "\shell\open\command\", LaunchCom, "REG_EXPAND_SZ"
			else
				oShell.RegWrite "HKCU\Software\Classes\" & DefProgID & "\shell\AVZ.DeQuarantine\command\", LaunchCom, "REG_EXPAND_SZ"
			end if
		else
			oShell.RegWrite "HKCU\Software\Classes\" & DefProgID & "\shell\", curVerb, "REG_SZ"
		end if
	end if
end if

Function StrBeginWith(Text, BeginPart)
    StrBeginWith = (StrComp(Left(Text, Len(BeginPart)), BeginPart, 1) = 0)
End Function
