' Установщик скрипта "AVZ DeQuarantine"

Option Explicit
Dim oShell, oFSO, curPath, AppData, SendTo, InstFolder, AppLink, AppName, AppPath, oFile, oTS, str, StartMenuPrograms, SysRoot, bTwice, DefVerb, DefProgID, UserChoiceHash
Dim StartMenuLink_1, StartMenuLink_2, StartMenuLink_3, StartMenuLink_4, StartMenuLink_5, EditorPath, txtClass, MsgFinish, AssocKey, OSVer, IsWin8AndNewer, IsWinXP, IsWin7, oShellApp

Set oShell     = CreateObject("WScript.Shell")
Set oFSO       = CreateObject("Scripting.FileSystemObject")

AppName = "AVZ DeQuarantine by Dragokas"
SendTo = oShell.SpecialFolders("SendTo")
AppData = oShell.SpecialFolders("AppData")
StartMenuPrograms = oShell.SpecialFolders("Programs")
SysRoot = oShell.ExpandEnvironmentStrings("%SystemRoot%")

'Папка установки
InstFolder = AppData & "\AVZ DeQuarantine"
AppPath = InstFolder & "\" & "AVZ - распаковать карантин.cmd"

On Error Resume Next
OSVer = CSng(Replace(oShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion"),".",","))
On Error Goto 0
IsWin8AndNewer = (OSVer > 6.1)
IsWinXP = (OSVer < 6)
IsWin7 = (OSVer >= 6 and OSVer <= 6.1)
if OSVer = 0 then Msgbox "Неизвестная версия ОС!", vbExclamation

if WScript.Arguments.Count > 1 then
	if StrComp(WScript.Arguments(1), "Twice", 1) = 0 then bTwice = true 'маркер того, что скрипт был перезапущен
end if

if WScript.Arguments.Count > 0 then
	if StrComp(WScript.Arguments(0), "AssocInstall", 1) = 0 then
		InstallZipAssoc
		if not bTwice then msgbox "OK.",,AppName
		WScript.Quit
	elseif StrComp(WScript.Arguments(0), "AssocUnInstall", 1) = 0 then
		UnInstallZipAssoc
		msgbox "OK.",,AppName
		WScript.Quit
	end if
end if

AppLink = SendTo & "\" & "AVZ - распаковать карантин.lnk"

StartMenuLink_1 = StartMenuPrograms & "\AVZ DeQuarantine\AVZ DeQuarantine.lnk"
StartMenuLink_2 = StartMenuPrograms & "\AVZ DeQuarantine\Настройки.lnk"
StartMenuLink_3 = StartMenuPrograms & "\AVZ DeQuarantine\Удалить AVZ DeQuarantine.lnk"
StartMenuLink_4 = StartMenuPrograms & "\AVZ DeQuarantine\quarantine.zip - Создать ассоциацию.lnk"
StartMenuLink_5 = StartMenuPrograms & "\AVZ DeQuarantine\quarantine.zip - Удалить ассоциацию.lnk"

'Получение пути к программе по-умолчанию для редактирования файлов .txt
EditorPath = oShell.ExpandEnvironmentStrings(RemoveArgument(GetDefaultAppString("txt")))
if not oFSO.FileExists(EditorPath) then
	EditorPath = SysRoot & "\notepad.exe"
end if

curPath = oFSO.GetParentFolderName(WScript.ScriptFullname)

'Получение версии скрипта
if oFSO.FileExists(curPath & "\bin\AVZ - распаковать карантин.cmd") then
	Set oFile = oFSO.GetFile(curPath & "\bin\AVZ - распаковать карантин.cmd")
    Set oTS   = oFile.OpenAsTextStream(1)
    Do While Not oTS.AtEndOfStream
        str = oTS.ReadLine()
        if instr(1, str, "AppVersion", 1) <> 0 then
			AppName = AppName & " ver." & mid(str, instr(str, "=") + 1)
			exit Do
		end if
    Loop
	oTS.Close
    set oTS = Nothing: set oFile = Nothing
end if

' UnInstaller
if oFSO.FileExists(InstFolder & "\AVZ - распаковать карантин.cmd") then
	if msgbox("AVZ DeQuarantine уже установлен. Хотите удалить его?", vbYesNo, AppName) = vbYes then
		on error resume next
        oShell.CurrentDirectory = SysRoot

		UnInstallZipAssoc

        if oFSO.FileExists(AppLink) then oFSO.DeleteFile AppLink, true
        if oFSO.FolderExists(StartMenuPrograms & "\AVZ DeQuarantine") then
            oFSO.DeleteFolder StartMenuPrograms & "\AVZ DeQuarantine", true
        end if
        err.Clear
		if oFSO.FolderExists(InstFolder) then oFSO.DeleteFolder InstFolder, true
		if err.Number <> 0 then
			msgbox "Удалите самостоятельно папку: " & InstFolder,, AppName
			oShell.Run "explorer.exe " & """" & InstFolder & """"
		else
			msgbox "AVZ DeQuarantine успешно удален.", 0, AppName
		end if
	end if
	WScript.Quit
end if

'Проверка, что запущен не из архива
if not oFSO.FileExists(curPath & "\bin\AVZ - распаковать карантин.cmd") then
	MsgBox "Сначала нужно распаковать все файлы из архива.",, AppName
	WScript.Quit
end if

'Создаю папку для установки приложения
if not oFSO.FolderExists(InstFolder) then oFSO.CreateFolder InstFolder

'Копирую файлы скрипта
oFSO.CopyFile curPath & "\bin\*", InstFolder & "\", true
oFSO.CopyFile WScript.ScriptFullName, InstFolder & "\AVZ DeQuarantine (Установщик).vbs", true
oShell.Run "cmd.exe /c ""<NUL set /p=>""" & InstFolder & "\7za.exe" & """:Zone.Identifier:$DATA""", 0, false
oShell.Run "cmd.exe /c ""<NUL set /p=>""" & InstFolder & "\AssocZip_Handler.vbs" & """:Zone.Identifier:$DATA""", 0, false
oShell.Run "cmd.exe /c ""<NUL set /p=>""" & InstFolder & "\AVZ - распаковать карантин.cmd" & """:Zone.Identifier:$DATA""", 0, false

'Создание ярлыка в папке SendTo (контекстное меню)
with oShell.CreateShortcut(AppLink)
	.Description        = AppName
	.IconLocation       = InstFolder & "\" & "Зверьки.ico"
	.TargetPath         = AppPath
	.WorkingDirectory   = InstFolder
	.Save
end with

if vbYes =  msgbox ("Желаете создать ярлыки в меню ПУСК ?", vbYesNo, AppName) then
  'Создание ярлыков в меню ПУСК (для настройки + удаления)
  if not oFSO.FolderExists(StartMenuPrograms & "\AVZ DeQuarantine") then
	oFSO.CreateFolder StartMenuPrograms & "\AVZ DeQuarantine"
  end if

  with oShell.CreateShortcut(StartMenuLink_1)
	.Description        = AppName
	.IconLocation       = InstFolder & "\" & "Зверьки.ico"
	.TargetPath         = AppPath
	.WorkingDirectory   = InstFolder
	.Save
  end with
  with oShell.CreateShortcut(StartMenuLink_2)
	.Description        = "Настройка " & AppName
	.IconLocation       = EditorPath
	.TargetPath         = EditorPath
    .Arguments          = """" & AppPath & """"
	.WorkingDirectory   = InstFolder
	.Save
  end with
  with oShell.CreateShortcut(StartMenuLink_3)
	.Description        = "Удаление " & AppName
	.TargetPath         = InstFolder & "\AVZ DeQuarantine (Установщик).vbs"
	.WorkingDirectory   = InstFolder
	.Save
  end with
  with oShell.CreateShortcut(StartMenuLink_4)
	.Description        = "Создать ассоциацию для файлов quarantine*.zip"
	.TargetPath         = "wscript.exe"
	.Arguments          = """" & InstFolder & "\AVZ DeQuarantine (Установщик).vbs" & """" & " " & "AssocInstall"
	.WorkingDirectory   = InstFolder
	.Save
  end with
  with oShell.CreateShortcut(StartMenuLink_5)
	.Description        = "Вернуть стандартную ассоциацию для файлов *quarantine*.zip"
	.TargetPath         = "wscript.exe"
	.TargetPath         = "wscript.exe"
	.Arguments          = """" & InstFolder & "\AVZ DeQuarantine (Установщик).vbs" & """" & " " & "AssocUnInstall"
	.WorkingDirectory   = InstFolder
	.Save
  end with
  MsgFinish = " и в меню ""Пуск"""
end if

  Dim AssocAlsoText

'  if IsWin8AndNewer then
'    AssocAlsoText = vbcrlf & "(потребуется повышение привилегий)"
'  end if

if not IsWin8AndNewer then
  if vbYes = Msgbox ("Хотите создать ассоциацию для файлов *quarantine*.zip ?" & AssocAlsoText, vbYesNo, AppName) then
	InstallZipAssoc
  end if
end if

if vbYes = Msgbox ("Скрипт установлен в контекстное меню ""Отправить"" (SendTo)" & MsgFinish & "." & vbCrLf & vbCrLf & _
	   "Файлы приложения находятся в папке:" & vbCrLf & InstFolder & vbCrLf & vbCrLf & _
       "Желаете настроить работу скрипта по своему вкусу?", vbYesNo, AppName) then

	oShell.Run """" & EditorPath & """" & " " & """" & AppPath & """", 1, false

end if

Set oShell = Nothing: Set oFSO = Nothing

' установка ассоциации для файлов quarantine*.zip
Sub InstallZipAssoc()
	Dim DefProgID, IconRes, pos, DefExeName, curExeName, deqVerb

	On Error Resume Next

	'is already installed
	curExeName = oShell.RegRead ("HKEY_CURRENT_USER\Software\Classes\Applications\AVZ.DeQuarantine\shell\open\command\")
	if (0 <> Len(curExeName)) then 
		UnInstallZipAssoc
	end if

	if IsWin8AndNewer then
		UserChoiceHash = oShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\Hash")
		if not RemoveUserChoiceHash() then
		    if not isAdminRights() then
				if bTwice then 'already re-launched ?
					Msgbox "Возникла ошибка при попытке повысить привилегии. Функция ассоциации не установлена!", vbCritical, AppName
				else
					Set oShellApp = CreateObject("Shell.Application")
					oShellApp.ShellExecute WScript.FullName, """" & WScript.ScriptFullName & """" & " " & "AssocInstall Twice", "", "runas", 1
				end if
				Exit Sub
			end if
		end if
	end if
	
	'read current zip application class if exists
	if IsWinXP then
		DefProgID = oShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\Application")
		if DefProgID <> "" then
			DefProgID = "Applications\" & DefProgID
		else
			DefProgID = oShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\ProgID")
		end if
	else
		DefProgID = oShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\ProgID")
	end if

	'read application's icon
	if DefProgID <> "" then
		IconRes = oShell.RegRead("HKCR\" & DefProgID & "\DefaultIcon\")
		if IconRes = "" then
			DefVerb = oShell.RegRead("HKCR\" & DefProgID & "\shell\")
			if DefVerb = "" then
				DefExeName = oShell.RegRead("HKCR\" & DefProgID & "\shell\open\command\")	
			elseif DefVerb <> "AVZ.DeQuarantine" then
				DefExeName = oShell.RegRead("HKCR\" & DefProgID & "\shell\" & DefVerb & "\command\")
			end if
			if DefExeName <> "" then
				IconRes = GetFilePath(DefExeName) & ",0"
			end if
		end if
	else
		DefProgID = oShell.RegRead("HKCR\.zip\")
		if DefProgID <> "" then
			IconRes = oShell.RegRead("HKCR\" & DefProgID & "\DefaultIcon\")
			if IconRes = "" then
				DefVerb = oShell.RegRead("HKCR\" & DefProgID & "\shell\")
				if DefVerb = "" then
					DefExeName = oShell.RegRead("HKCR\" & DefProgID & "\shell\open\command\")
				elseif DefVerb <> "AVZ.DeQuarantine" then
					DefExeName = oShell.RegRead("HKCR\" & DefProgID & "\shell\" & DefVerb & "\command\")
				end if
				if DefExeName <> "" then
					IconRes = GetFilePath(DefExeName) & ",0"
				end if
			end if
		end if
		DefProgID = ""
	end if

	'try masking under current icon
	if IconRes <> "" then
		oShell.RegWrite "HKEY_CURRENT_USER\Software\Classes\Applications\AVZ.DeQuarantine\DefaultIcon\", IconRes, "REG_EXPAND_SZ"
	end if

	'rewrite app class by own
 	if IsWinXP then
		oShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\Application", "AVZ.DeQuarantine", "REG_SZ"
		oShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\ProgID"

	elseif IsWin7 then
		Err.Clear
		oShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\ProgID", "Applications\AVZ.DeQuarantine", "REG_SZ"

		if Err.Number <> 0 then
			RegKeyResetSecurity "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice"
			Err.Clear
			oShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\ProgID", "Applications\AVZ.DeQuarantine", "REG_SZ"
			if Err.Number <> 0 then
				if isAdminRights() then
					msgbox "Ошибка! Не могу установить ассоциацию. Причина: не хватает прав для записи в реестр."
				else
					if bTwice then
						Msgbox "Возникла ошибка при попытке повысить привилегии. Функция ассоциации не установлена!", vbCritical, AppName
					else
						Set oShellApp = CreateObject("Shell.Application")
						oShellApp.ShellExecute WScript.FullName, """" & WScript.ScriptFullName & """" & " " & "AssocInstall", "", "runas", 1
					end if
				end if
			end if
		end if

	elseif IsWin8AndNewer then
		DefProgID = oShell.RegRead("HKCR\.zip\")
		if DefProgID <> "" then
			DefVerb = oShell.RegRead("HKCR\" & DefProgID & "\shell\")
			DefExeName = oShell.RegRead("HKCU\Software\Classes\" & DefProgID & "\shell\" & DefVerb & "\command\")
			if DefExeName = "" or instr(1, DefExeName, "AVZ DeQuarantine", 1) <> 0 then 'only verb in HKCU
				deqVerb = "open"
			else '"open" verb is already exists
				oShell.RegWrite "HKCU\Software\Classes\" & DefProgID & "\shell\", "AVZ.DeQuarantine", "REG_SZ"
				deqVerb = "AVZ.DeQuarantine"
			end if

			if DefVerb = "AVZ.DeQuarantine" then DefVerb = ""

			oShell.RegWrite "HKCU\Software\Classes\" & DefProgID & "\shell\" & deqVerb & "\command\", _
				"wscript.exe" & " " & """" & InstFolder & "\AssocZip_Handler.vbs" & """" & " " & """" & "%1" & """" & " " & """" & UserChoiceHash & """" & " " & """" & DefVerb & """", _
				"REG_EXPAND_SZ"
		end if
	end if

	'write execution string for own application class including old info (old app class)
	'1.  wscript.exe
	'2.  AssocZip_Handler.vbs -> it directs quarantine*.zip files into new app class execution string, and the rest into the old execution string
	'3.  %1
	'4.  DefProgID (or DefHash - Win8+)
	'5.  Old Verb (Win8+ only)
	if IsWin8AndNewer then DefProgID = UserChoiceHash

	oShell.RegWrite "HKEY_CURRENT_USER\Software\Classes\Applications\AVZ.DeQuarantine\shell\open\command\", _
		"wscript.exe" & " " & """" & InstFolder & "\AssocZip_Handler.vbs" & """" & " " & """" & "%1" & """" & " " & """" & DefProgID & """" & " " & """" & DefVerb & """", _
		"REG_EXPAND_SZ"

End Sub

' удаление ассоциации файлов quarantine*.zip, virusinfo_auto_*.zip, ZOO_*.zip
Sub UnInstallZipAssoc()
	On Error Resume Next
	Dim AssocKey, pos, cnt, NewAssocString, DefProgID, qt1, qt2, CurAssocClass, oldHash, curVerb, CurExeName

	NewAssocString = oShell.RegRead ("HKEY_CURRENT_USER\Software\Classes\Applications\AVZ.DeQuarantine\shell\open\command\")

	'checking requirements
	if IsWinXP then
		CurAssocClass = oShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\Application")
	elseif IsWin7 then
		CurAssocClass = oShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\ProgID")
	end if

	'is it still own app class for zip assoc.?
	if CurAssocClass = "Applications\AVZ.DeQuarantine" or CurAssocClass = "AVZ.DeQuarantine" or IsWin8AndLater then
		'is installed
		if (0 <> Len(NewAssocString)) then
			pos = 0
			Do
				cnt = cnt + 1
				pos = pos + 1
				pos = instr(pos, NewAssocString, """")
				if pos <> 0 then
					Select case Cnt
					case 5
						qt1 = pos
					case 6
						qt2 = pos
						if qt2 - qt1 <> 1 then 'not empty
							DefProgID = mid(NewAssocString, qt1 + 1, qt2 - qt1 - 1)
							if IsWinXP then
								if StrBeginWith(DefProgID, "Applications") then
									oShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\Application", Mid(DefProgID, Len("Applications")+2), "REG_SZ"
								else
									oShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\ProgID", DefProgID, "REG_SZ"
									oShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\Application"
								end if
							elseif IsWin7 then
								oShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\ProgID", DefProgID, "REG_SZ"
'							elseif IsWin8AndLater then
'								UserChoiceHash = DefProgID
							end if
						end if
					case 7
						qt1 = pos
					case 8
						qt2 = pos
						if qt2 - qt1 <> 1 then 'not empty
							DefVerb = mid(NewAssocString, qt1 + 1, qt2 - qt1 - 1)
						end if						
					End Select
				end if
			Loop Until pos = 0
	    end if
		'if last time UserChoice app class wasn't defined
		if 0 = Len(DefProgID) then 
			if IsWinXP then
				oShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\Application"
			elseif IsWin7 then
				oShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\ProgID"
			end if
		end if
	end if

if false then
	if IsWin8AndLater then
		DefProgID = oShell.RegRead("HKCR\.zip\")
		if DefProgID <> "" then
			oShell.RegDelete "HKCU\Software\Classes\" & DefProgID & "\shell\AVZ.DeQuarantine\command\"
			oShell.RegDelete "HKCU\Software\Classes\" & DefProgID & "\shell\AVZ.DeQuarantine\"
			curVerb = oShell.RegRead("HKCU\Software\Classes\" & DefProgID & "\shell\")
			if curVerb = "" then curVerb = "open"
			CurExeName = oShell.RegRead("HKCU\Software\Classes\" & DefProgID & "\shell\" & curVerb & "\command\")
			if CurExeName = "" then
				oShell.RegDelete "HKCU\Software\Classes\" & DefProgID & "\shell\"
				oShell.RegDelete "HKCU\Software\Classes\" & DefProgID & "\"
			end if
			if curVerb = "open" and instr(1,CurExeName,"AVZ DeQuarantine",1) <> 0 then
				oShell.RegDelete "HKCU\Software\Classes\" & DefProgID & "\shell\open\command\"
				oShell.RegDelete "HKCU\Software\Classes\" & DefProgID & "\shell\open\"
				oShell.RegDelete "HKCU\Software\Classes\" & DefProgID & "\shell\"
				oShell.RegDelete "HKCU\Software\Classes\" & DefProgID & "\"
			end if
		end if
		' is zip overwrited ?
		oldHash = oShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\Hash")
		if oldHash = "" and UserChoiceHash <> "" then
			' import old data
			oShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\Hash", UserChoiceHash, "REG_SZ"
		end if
	end if
end if
	
	'removing own app class
	oShell.RegDelete "HKEY_CURRENT_USER\Software\Classes\Applications\AVZ.DeQuarantine\shell\open\command\"
	oShell.RegDelete "HKEY_CURRENT_USER\Software\Classes\Applications\AVZ.DeQuarantine\shell\open\"
	oShell.RegDelete "HKEY_CURRENT_USER\Software\Classes\Applications\AVZ.DeQuarantine\shell\"
	oShell.RegDelete "HKEY_CURRENT_USER\Software\Classes\Applications\AVZ.DeQuarantine\DefaultIcon\"
	oShell.RegDelete "HKEY_CURRENT_USER\Software\Classes\Applications\AVZ.DeQuarantine\"
End Sub

Function RemoveUserChoiceHash()
	On Error Resume Next
	RegKeyResetSecurity "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice"
	oShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.zip\UserChoice\Hash"
	RemoveUserChoiceKey = (Err.Number = 0)
End Function

Function StrBeginWith(Text, BeginPart)
    StrBeginWith = (StrComp(Left(Text, Len(BeginPart)), BeginPart, 1) = 0)
End Function

' получить строку вызова для указанной ассоциации
Function GetDefaultAppString(Extension)
	On Error Resume Next
	Dim txtClass, file
	txtClass = oShell.RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\." & Extension & "\UserChoice\ProgID")
	if 0 <> Len(txtClass) then
		file = oShell.RegRead("HKEY_CLASSES_ROOT\" & txtClass & "\shell\open\command\")
	end if
	if 0 = len(file) then
		txtClass = oShell.RegRead("HKEY_CLASSES_ROOT\." & Extension & "\")
		if 0 <> len(txtClass) then
			file = oShell.RegRead("HKEY_CLASSES_ROOT\" & txtClass & "\shell\open\command\")
		end if
	end if
	GetDefaultAppString = file
End Function

' удалить аргумент из строки
Function RemoveArgument(sPath)
	Dim tmp, pos, ret
	tmp = trim(sPath)
	if left(tmp, 1) = """" then
		pos = instr(2, tmp, """")
        if pos <> 0 then 
			ret = mid(tmp, 2, pos - 2)
		else
			ret = mid(tmp, 2)
		end if
	else
		pos = instr(tmp, " ")
		if pos <> 0 then
			ret = Left(tmp, pos - 1)
		else
			ret = tmp
		end if
	end if
	RemoveArgument = ret
End Function

Sub RegKeyResetSecurity(sKey)
	On Error Resume Next
	Dim oWMI, oStdReg, CurSID, AdmSID, objSD, SID, i

	Const FILE_WRITE_DATA = 2
	Const ACE_ALLOW = 0
	Const ACE_DENY = 1
	Const HKCU = &H80000001

	'set oShell = CreateObject("WScript.Shell")

	Set oStdReg = GetObject("winmgmts:{impersonationLevel=Impersonate, (Security)}!\\.\root\default:StdRegProv") ' grant SeSecurityPrivilege

	Set oWMI = GetObject("winmgmts:\\.\Root\CIMV2")

	CurSID = oWMI.Get("Win32_UserAccount.Domain='" & oShell.ExpandEnvironmentStrings("%UserDomain%") & "',Name='" & oShell.ExpandEnvironmentStrings("%UserName%") & "'").SID
	AdmSID = "S-1-5-32-544"

	'Set objSD = oWMI.Get("Win32_SecurityDescriptor").SpawnInstance_()

	oStdReg.GetSecurityDescriptor HKCU, sKey, objSD

	If Not IsNull(objSD.DACL) Then

		For i = 0 to UBound(objSD.DACL)

			SID = objSD.DACL(i).Trustee.SIDString

			if ((SID = CurSID) or (SID = AdmSID)) then

				if (objSD.DACL(i).AccessMask And FILE_WRITE_DATA) then

					if (objSD.DACL(i).AceType = ACE_DENY) then
					
						objSD.DACL(i).AceType = ACE_ALLOW

					end if
				end if
			end if
		next
	end if

	oStdReg.SetSecurityDescriptor HKCU, sKey, objSD

	set oStdReg = Nothing
	set oWMI = Nothing
End Sub

function UnQuote(byval str)
	if left(str,1) = """" then str = mid(str, 2)
	if right(str,1) = """" then str = left(str, len(str)-1)
	UnQuote = str
end function

function GetFilePath(DirtyLine)
	Dim pos
	pos = instr(1, DirtyLine, ".exe", 1)
	if pos <> 0 then
		GetFilePath = UnQuote(left(DirtyLine, pos+3))
	end if
end function

Function isAdminRights()
    Const KQV = 1, KSV = 2, HKLM = &H80000002
	Dim oReg, strKey, intErrNum, flagAccess
    Set oReg = GetObject("winmgmts:root\default:StdRegProv")
    strKey = "System\CurrentControlSet\Control\Session Manager"
    intErrNum = oReg.CheckAccess(HKLM, strKey, KQV + KSV, flagAccess)
    isAdminRights = flagAccess
    Set oReg = Nothing
End Function