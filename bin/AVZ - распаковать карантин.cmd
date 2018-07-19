@echo off
SetLocal EnableExtensions EnableDelayedExpansion
chcp 866 >nul

set AppVersion=1.20

title AVZ DeQuarantine Script ver. %AppVersion%
set "args=%~1"

:: >>>>>   ���������   <<<<<

:: �������� + / -

:: �������������� ����� ��� ���������� ��� ��������������?
set OverWrite=+

:: ����������� �� ���� ������ ��� ������� � ������������� �������
set AskPassword=+

:: �� ���������� ��������� ����-�����
set OpenLogFile=+

:: �� ���������� ��������� ����� � �������������� �������
set OpenFolder=+

:: ���������� ��������� ���������� � ������ ����� ���������� (��������, .vir)
set AddExtension=

:: �� �������� �� ���������� ����� ������ � ��������� AVZ ��� �� ���������� ����������� � �������� ����������
set NoWarning_LackOfFiles=-

:: �� ��������, ���� ����� ����������� ��� �����
set NoWarning_UnknownFile=-

:: �� ��������, ���� ����� ����� ����������� ������
set NoWarning_UnknownPassword=-

:: �� �������� � ������ �������
set NoWarning_EmptyArchive=-

:: ���������� ���� ��������� ������ ����������� (����� ������� �������� � �������� �����������)
set DoMove=-

:: ������� ��� �������� (����� ���������� ���������� �����)
set RemoveDelays=-






goto main

:Using [ᮮ�饭�� �� �訡��]
  mode 104,41
  echo ��ꥪ�: "%args%"
  echo.
  echo  �訡��: %~1
  echo. _______________________________________________
  echo  AVZ DeQuarantine   by Alex Dragokas   ver. %AppVersion%
  echo. 
  echo  ��ᯠ���騪 ��࠭⨭� AVZ
  echo.
  echo.
  echo  ���ᮡ� �ᯮ�짮�����:
  echo.
  echo. 1. ��१ ���⥪�⭮� ����
  echo  - ���ਬ��, ������� ��⭨� � ����� ��� -^> Shell:SendTo (��� ���⥪�⭮�� ���� "��ࠢ���")
  echo.
  echo  2. ��������� ��ꥪ� �� ��� ��⭨�
  echo.
  echo. ����� ��ꥪ�� ����� �������� / ��ࠢ���� �१ ���⥪�⭮� ����:
  echo.
  echo  - �����    AVZ         (�㤥� �ᯠ������ �� � ������ Quarantine � Infected)
  echo  - �����    Quarantine  (�ᯠ����� ⮫쪮 �� �⮩ �����)
  echo  - �����    Infected    (�ᯠ����� ⮫쪮 �� �⮩ �����)
  echo  - 䠩�     .dta / .dat (�㤥� ��ࠡ�⠭ ⮫쪮 1 䠩�)
  echo  - 䠩�     .ini        (�㤥� ��ࠡ�⠭ ⮫쪮 1 䠩�)
  echo  - ��࠭⨭ .zip        (quarantine*.zip, virusinfo_auto_*.zip, ZOO_*.zip)
  echo.
  echo  3. ����஢���� �⮣� ��⭨�� � ���� �� �����:
  echo.
  echo  - ����� AVZ         (�㤥� �ᯠ������ �� � ������ Quarantine � Infected)
  echo  - ����� Quarantine  (�ᯠ����� ⮫쪮 �� �⮩ �����)
  echo  - ����� Infected    (�ᯠ����� ⮫쪮 �� �⮩ �����)
  echo  - � ����� �冷� � 䠩���� .dta / .dat , .ini ��� .zip
  echo. 
  echo  � ��⥬ ����� ��⭨��.
  echo.
  echo  ��ᯠ������� ��࠭⨭ ��࠭���� � �����⠫�� Unpacked \ ^< �ਣ����쭮� ��� 䠩�� ^>
  echo. "Unpacked" �㤥� ᮧ���� �冷� � ��࠭⨭�묨 䠩����.
  echo.
  echo  �᫨ ��� ��ࠡ�⪨ 㪠�뢠���� �����, � �� ������� 䠩�, "Unpacked" �।���⥫쭮 ��頥���.
  echo.
  echo  � ����� Unpacked ᮧ������ ^^^!^^^!^^^!_AVZ_QFiles_^^^!^^^!^^^!.txt � ᯨ᪮� �ᯮ������� 䠩��� �� ��ࠦ����� ��設�.
  pause >NUL
Exit /B

:main
cd /d "%~dp0"

:: �����প� �����६����� �ᯠ����� ����� 1 䠩�� (��娢�) ��࠭⨭�, ��।������ ���⥪��� ���� ��� ����������
if "%~2" neq "" ((For %%a in (%*) do call "%~fs0" "%%~a")& goto :eof)

set "arc7z=.\7za.exe"
set "WinRAR=.\rar.exe"
set "WinRAR_enabled=n"
set "Q0=Unpacked\^^^!^^^!^^^!_AVZ_QFiles_^^^!^^^!^^^!.txt"
set "Q1=Unpacked\^^^!^^^!^^^!_AVZ_QFiles_^^^!^^^!^^^!_1.txt"
set "Q2=Unpacked\^^^!^^^!^^^!_AVZ_QFiles_^^^!^^^!^^^!_2.txt"
set "Q3=Unpacked\^^^!^^^!^^^!_AVZ_QFiles_^^^!^^^!^^^!_3.txt"
set "qt=""

for /F "delims=#" %%a in ('"prompt #$H#& echo on & for %%A in (1) do rem"') do set "DEL=%%a"

if "%DoMove%"=="+" (set QrCommand=move) else (set QrCommand=copy)

:: ���� ��娢��஢
if not exist "%arc7z%" (
  if exist "%SystemDrive%\Program Files\7-Zip\7z.exe" set "arc7z=%SystemDrive%\Program Files\7-Zip\7z.exe"
  if exist "%SystemDrive%\Program Files\7-Zip\7za.exe" set "arc7z=%SystemDrive%\Program Files\7-Zip\7za.exe"
)
if /i "%WinRAR_enabled%"=="y" (
if not exist "%WinRAR%" (
  if exist "%SystemDrive%\Program Files\WinRAR\WinRAR.exe" set "WinRAR=%SystemDrive%\Program Files\WinRAR\WinRAR.exe"
  if exist "%SystemDrive%\Program Files\WinRAR\rar.exe" set "WinRAR=%SystemDrive%\Program Files\WinRAR\rar.exe"
))

if "%~1" neq "" goto ARGS

:: �᫨ ��⭨� � ����� � ��࠭⨭�묨 䠩����
if exist avz*.dta (call :Unpack . *.ini & exit /B)
if exist bcqr*.dat (call :Unpack . *.ini & exit /B)

:: �᫨ ��⭨� � ����� AVZ
if exist avz.exe for %%f in (Quarantine Infected) do if exist ".\%%~f" call :Subf ".\%%~f\*"
if "%Success%"=="true" exit /B

:: �᫨ *.dta 䠩�� ������ ���㡨�� ��⠫��� - ��ࠡ��뢠�� ��, ����� ⮫쪮 ���������
for /F "tokens=1* delims=[]" %%a in ('dir /b /s /a-d ".\avz*.dta" ".\bcqr*.dat" 2^>NUL ^| find /n /v ""') do for /f "delims=" %%h in ("%%~dpb\.") do if "%%~nxh" neq "Unpacked" set "DTA[%%a]=%%~dpb"
set DTA[ 2>NUL 1>NUL|| goto NOBAT_DTA
set "Prev="
For /F "tokens=1* delims==" %%a in ('set DTA[') do call :ProcDTAFolder "%%~b"
exit /B

:NOBAT_DTA

:: �᫨ ��⭨� ����饭 �冷� � ZIP-��࠭⨭��
if exist *.zip For /F "delims=" %%a in ('dir /b /a ".\*.zip" 2^>NUL') do (call "%~f0" ".\%%~a" & set "ArcGained=true")
if "%ArcGained%"=="true" (
  rem call :ClearArc & 
  exit /B
)

if /i "%WinRAR_enabled%"=="y" (
  if exist *.rar For /F "delims=" %%a in ('dir /b /a ".\*.rar" 2^>NUL') do (call "%~f0" ".\%%~a" & set "ArcGained=true")
)
if "%ArcGained%"=="true" (
    rem call :ClearArc
    exit /B
)

:: �᫨ ��࠭⨭ ��室���� � ����� �� �����⠫����
For /F "delims=" %%a in ('dir /b /s /ad ".\*" 2^>NUL') do (
  if /i "%%~nxa" neq "Unpacked" For /F "delims=" %%b in ('dir /b /a-d "%%~a\*.zip" 2^>NUL') do (
    call "%~f0" "%%~a\%%~b" & set "ArcGained=true"
))
if "%ArcGained%"=="true" (
    rem call :ClearArc
    exit /B
)

if /i "%WinRAR_enabled%"=="y" (
  For /F "delims=" %%a in ('dir /b /s /ad ".\*" 2^>NUL') do (
    if /i "%%~nxa" neq "Unpacked" For /F "delims=" %%b in ('dir /b /a-d "%%~a\*.rar" 2^>NUL') do (
      call "%~f0" "%%~a\%%~b" & set "ArcGained=true"
  ))
)
if "%ArcGained%"=="true" (
    rem call :ClearArc
    exit /B
)

:: �� �१ ���⥪�⭮� ����
if "%~1"=="" (
  if /i "%NoWarning_UnknownFile%" neq "+" call :Using "����୮� �ᯮ������� ��⭨�� !!!" 
  exit /B
)

:ARGS

if not exist "%~1\" (
  <NUL set /p=>"%~f1":Zone.Identifier:$DATA
  if /i "%~x1"==".zip" goto ZIP_FILE
  if /i "%WinRAR_enabled%"=="y" (
  if /i "%~x1"==".rar" goto RAR_FILE
))

:: �᫨ dta / dat / ini 䠩� ����� � ����⢥ ��㬥��
set "Name=%~n1"
if /i "%Name:~,3%"=="avz" (
    if /i "%~x1"==".dta" (
    call :Unpack "%~dp1" "%~n1.ini" "don't remove Unpacked"
  ) else (
    if /i "%~x1"==".ini" call :Unpack "%~dp1" "%~nx1" "don't remove Unpacked"
  )
) else (
if /i "%Name:~,4%"=="bcqr" (
    if /i "%~x1"==".dat" (
    call :Unpack "%~dp1" "%~n1.ini" "don't remove Unpacked"
  ) else (
    if /i "%~x1"==".ini" call :Unpack "%~dp1" "%~nx1" "don't remove Unpacked"
  )
))
if "%Success%"=="true" exit /B

:: �᫨ ����� � ��࠭⨭�묨 䠩���� 㪠���� � ����⢥ ��㬥��
if exist "%~1\avz*.dta" call :Unpack "%~1" "*.ini" && exit /B
if exist "%~1\bcqr*.dat" call :Unpack "%~1" "*.ini" && exit /B

:: �᫨ ����� AVZ 㪠���� � ����⢥ ��㬥��
if exist "%~f1\avz.exe" for %%f in ("%~f1\Quarantine" "%~f1\Infected") do if exist "%%~f" call :Subf "%%~f\*"
if "%Success%"=="true" exit /B

:: �᫨ ZIP ��娢 㪠��� � ����⢥ ��㬥��
:ZIP_TYPE

set "ArcGained=false"

:: �᫨ ��㬥�� - �� �� �����, �ࠧ� ���室�� � ��ࠡ�⪥ ��㬥��, ��� 䠩��
if not exist "%~1\" goto FILES

:: �᫨ *.zip ��娢 ������ ���㡨�� ��⠫���
for /F "delims=" %%a in ('dir /b /s /ad "%~1\*" 2^>NUL') do (
  if /i "%%~nxa" neq "Unpacked" For /F "delims=" %%b in ('dir /b /a-d "%%~a\*.zip" 2^>NUL') do (
    call "%~f0" "%%~a\%%~b" & set "ArcGained=true"
  )
)
:: �᫨ ��娢 � ��୥��� �����
for /F "delims=" %%a in ('dir /b /a-d "%~1\*.zip" 2^>NUL') do (
    call "%~f0" "%~1\%%~a" & set "ArcGained=true"
)
if "%ArcGained%"=="true" (
  rem call :ClearArc
  exit /B
)

:: �᫨ *.rar ��娢 ������ ���㡨�� ��⠫���
if /i "%WinRAR_enabled%"=="y" (
  for /F "delims=" %%a in ('dir /b /s /ad "%~1\*" 2^>NUL') do (
    if /i "%%~nxa" neq "Unpacked" For /F "delims=" %%b in ('dir /b /a-d "%%~a\*.rar" 2^>NUL') do (
      call "%~f0" "%%~a\%%~b" & set "ArcGained=true"  
  ))
  for /F "delims=" %%a in ('dir /b /a-d "%~1\*.rar" 2^>NUL') do call "%~f0" "%~1\%%~a" & set "ArcGained=true"
)

:: �᫨ 㦥 ��ࠡ�⠫� �� ��娢� � �����⠫���� = �� ��室.
if "%ArcGained%"=="true" (
  rem call :ClearArc
  exit /B
)

:DTA
:: �᫨ *.dta 䠩�� ������ ���㡨�� ��⠫��� - ��ࠡ��뢠�� ��, ����� ⮫쪮 ���������
if exist "%~1\" (
  for /F "tokens=1* delims=[]" %%a in ('dir /b /s /a-d "%~1\avz*.dta" "%~1\bcqr*.dat" 2^>NUL ^| find /n /v ""') do for /f "delims=" %%h in ("%%~dpb\.") do if "%%~nxh" neq "Unpacked" set "DTA[%%a]=%%~dpb"
  set DTA[ 2>NUL 1>NUL|| goto FILES
  set "Prev="
  For /F "tokens=1* delims==" %%a in ('set DTA[') do call :ProcDTAFolder "%%~b"
  exit /B
)

:FILES

:: ���� �� ������ ��娢��
if exist "%~1\" (
  rem call :Using "����� ������ ⨯ 䠩�� ���� ��� 䠩��� ��� �ᯠ����� !!!"
  if "%~n1"=="virusinfo_syscure" exit /B
  if "%~n1"=="virusinfo_syscheck" exit /B
  if /i "%NoWarning_EmptyArchive%"=="-" (
    call :Using "��娢 ���⮩ !"
  )
  Exit /B
)

<NUL set /p=>"%~f1":Zone.Identifier:$DATA
if /i "%~x1"==".zip" goto ZIP_FILE
if /i "%WinRAR_enabled%"=="y" (
if /i "%~x1"==".rar" goto RAR_FILE
)
:: �᫨ ��࠭⨭ UVs
set "Arg=%~n1"
if "%Arg:~0,4%"=="ZOO_" exit /b
:: ����, �������⭮� ���७�� 䠩��
if /i "%NoWarning_UnknownFile%" neq "+" call :Using "����� ������ ⨯ 䠩�� !!!"
Exit /B

:ProcDTAFolder
  if "%Prev%" neq "%~1" (
    set "Prev=%~1"
    call "%~f0" "%~1"
  )
Exit /B

:ZIP_FILE

set "ArcSrc=%~f1"
set "ArcDest=%~dp1%~n1"

:: UVs ?
set "isUVs="
set "Arg=%~n1"
if "%Arg:~0,4%"=="ZOO_" set isUVs=true

call :TryUnpackZIP "%ArcSrc%" "%ArcDest%" && if defined isUVs (call :Unpack "%ArcDest%" "" "" UVs) else (call "%~f0" "%ArcDest%")
rem call :ClearArc
Exit /B

:RAR_FILE

set "ArcSrc=%~f1"
set "ArcDest=%~dp1%~n1"

call :TryUnpackRAR "%ArcSrc%" "%ArcDest%" && call "%~f0" "%ArcDest%"
rem call :ClearArc
Exit /B

:ClearArc
  :: �� �㭪�� �㦭� ⮫쪮 � ��砥 �ᯠ����� ��娢��� �� ��⭨�� �ଠ� CBI.
  :: ���� �� ��������� ������ ����� � ᠬ��� ��⭨�� � �� ���.
  Exit /B
  :: 2>NUL del /f "%arc7z%"
  :: 2>NUL del /f "%WinRAR%"
Exit /B

:TryUnpackZIP [ArcSrc] [ArcDest]
  set "ArcSrc=%~1"
  set "ArcDest=%~2"
  ::extrac32.exe /Y /L "%temp%" "%~f0" 7za.exe
  set "ch="
  if exist "%ArcDest%" (
    echo �����, � ������ � ��� �ᯠ������, 㦥 �������:
    echo.
    echo "%ArcDest%"
    echo.
    if /i "%OverWrite%"=="+" (set ch=Y) else (set /p "ch=������� ��? [Enter, Y / N] ")
  )
  if /i "%ch%"=="Y" (rd /S /Q "%ArcDest%"& md "%ArcDest%")
  set "SuccessUnpack=false"
  if "%SuccessUnpack%"=="false" "%arc7z%" x "%ArcSrc%" -o"%ArcDest%" -y -p"virus"    && set "SuccessUnpack=true"
  if "%SuccessUnpack%"=="false" "%arc7z%" x "%ArcSrc%" -o"%ArcDest%" -y -p"infected" && set "SuccessUnpack=true"
  if "%SuccessUnpack%"=="false" "%arc7z%" x "%ArcSrc%" -o"%ArcDest%" -y -p"malware"  && set "SuccessUnpack=true"
  if "%SuccessUnpack%"=="false" "%arc7z%" x "%ArcSrc%" -o"%ArcDest%" -y -p"clean"    && set "SuccessUnpack=true"
  if "%SuccessUnpack%"=="false" "%arc7z%" x "%ArcSrc%" -o"%ArcDest%" -y -p""         && set "SuccessUnpack=true"
  if "%SuccessUnpack%"=="false" if /i "%AskPassword%"=="+" "%arc7z%" x "%ArcSrc%" -o"%ArcDest%" -y && set "SuccessUnpack=true"
:Wrong_PASS_ZIP
  if /i "%SuccessUnpack%"=="false" (
    if /i "%NoWarning_UnknownPassword%"=="-" (
      echo.
      echo ��娢 ����� ��������� ��஫� ��� ���०��� !!!
      echo.
      pause >NUL
    )
  )
Exit /B


:TryUnpackRAR [ArcSrc] [ArcDest]
  set "ArcSrc=%~1"
  set "ArcDest=%~2"
  ::extrac32.exe /Y /L "%temp%" "%~f0" rar.exe
  set "ch="
  if exist "%ArcDest%" (
    echo �����, � ������ � ��� �ᯠ������, 㦥 �������:
    echo.
    echo "%ArcDest%"
    echo.
    if /i "%OverWrite%"=="+" (set ch=Y) else (set /p "ch=������� ��? [Y/N] ")
  )
  if /i "%ch%"=="Y" rd /S /Q "%ArcDest%"
  set "SuccessUnpack=false"
  if "%SuccessUnpack%"=="false" "%WinRAR%" x -o+ -y -p"virus"    "%ArcSrc%" *.* "%ArcDest%\" && set "SuccessUnpack=true"
  if "%SuccessUnpack%"=="false" "%WinRAR%" x -o+ -y -p"infected" "%ArcSrc%" *.* "%ArcDest%\" && set "SuccessUnpack=true"
  if "%SuccessUnpack%"=="false" "%WinRAR%" x -o+ -y -p"malware"  "%ArcSrc%" *.* "%ArcDest%\" && set "SuccessUnpack=true"
  if "%SuccessUnpack%"=="false" "%WinRAR%" x -o+ -y -p"clean"    "%ArcSrc%" *.* "%ArcDest%\" && set "SuccessUnpack=true"
  if "%SuccessUnpack%"=="false" "%WinRAR%" x -o+ -y -p-          "%ArcSrc%" *.* "%ArcDest%\" && set "SuccessUnpack=true"
  if "%SuccessUnpack%"=="false" if /i "%AskPassword%"=="+" "%WinRAR%" x -o+ -y "%ArcSrc%" *.* "%ArcDest%\" && set "SuccessUnpack=true"
:Wrong_PASS_RAR
  if "%SuccessUnpack%"=="false" (
    if /i "%NoWarning_UnknownPassword%"=="-" (
      echo.
      echo ��娢 ����� ��������� ��஫� ��� ���०��� !!!
      echo.
      pause >NUL
    )
  )
Exit /B


:Subf [�����]
  for /f "delims=" %%a in ('dir /b /ad "%~1" 2^>NUL') do call :Unpack "%~1\%%a" *.ini
  set "Success=true"
Exit /B


:Unpack [�����] [��᪠ ��� ini] [����� Unpacked ?] [��࠭⨭ UVs ?]
  
  pushd "%~1" || exit /B 1

  Set Files=0
  set FilesCnt=0
  set FilesCur=0
  Set SuccessCopied=0

  :: ���⪠ �����
  if "%~3"=="" if Exist "Unpacked" rd /s /q "Unpacked"
  if not Exist "Unpacked" md "Unpacked"

  echo.
  echo �⥭�� ����...
  echo.

  if "%~4"=="" goto UnpackAVZ

  :UnpackUVs

  :: UTF-16 => 866
  chcp 866 >NUL
  for %%a in (*.txt) do if "%%~xa"==".txt" (
    cmd /d /a /c type "%%a" > "%%a_866"
    set /a FilesCnt+=1
  )

  for %%a in (*.txt) do if "%%~xa"==".txt" (
    set /a FilesCur+=1

    (set /p src=& set /p src=& set /p src=& set /p src=) < "%%a_866"
    For /F "tokens=2*" %%r in ("!src!") do set "src866=%%s"

    call :Proc "%%~na" "!src866!" UVs
  )
  for %%a in (*.txt) do if "%%~xa"==".txt" del "%%a_866"

  goto UnpackEND

  :UnpackAVZ

  :: 2-� ��ப� ini-䠩�� ������ � ���ᨢ
  chcp 1251 >NUL

  For /F "delims=" %%a in ('dir /b /a-d "%~2" 2^>NUL') do (set /p avz[%%~na]=& set /p avz[%%~na]=)<"%%a"

  :: ������ � ��� �ᯮ������� 䠩��� �� ��ࠦ����� ��⥬�
  For /F "tokens=1-2* delims==" %%a in ('set avz[') do set /a FilesCnt+=1

  :: ��ॢ��� � ����஢�� OEM-866, ��饯��� ���ᨢ, 㤠��� ����� �ࠧ� "Src="
  chcp 866 >NUL
  For /F "tokens=1-2* delims==" %%a in ('set avz[') do for /F "tokens=2 delims=[]" %%n in ("%%a") do (
    For /F "UseBackQ tokens=1-2 delims==" %%s in ("%%n.ini") do if /i "%%s"=="Infected" (set /a FilesCur+=1& call :Proc "%%t" "%%c")
  )

  :UnpackEND

  echo.
  echo.
  echo �����⮢�� ���� ...

  sort "%Q3%" /O "%Q1%"
  set /a FalseCopied=%Files% - %SuccessCopied%
  (
  echo.
  echo ----------------------------------------
  if exist "%Q2%" (type "%Q2%" & echo.)
  echo. ����� ������:  %Files%
  echo  ������:        %FalseCopied%
  ) >> "%Q1%"
  for /F %%? in ('echo ��') do chcp 1251 >nul& CMD.EXE /D /A /C (set /p=��)<NUL > "%Q0%"
  CMD.EXE /D /U /C TYPE "%Q1%" >> "%Q0%"
  chcp 866 >NUL

  del /f "%Q1%" 2>NUL
  del /f "%Q2%" 2>NUL
  del /f "%Q3%" 2>NUL

  if "%Files%" neq "%SuccessCopied%" (
    if /i "%NoWarning_LackOfFiles%" neq "+" (
      echo ------------ 
      echo �������� ^^!^^!^^! ������� 䠩�� �� 㤠���� �ᯠ������ ^^!
      call :Wait 1
    )
  ) else (
    if /i "%NoWarning_LackOfFiles%" neq "-" (
      echo �஢�ઠ ������⢠: �����.
      call :Wait 1
    )
  )

  :: anti shell bug
  if "%OpenFolder%"=="+" call :Wait 1
  if "%OpenLogFile%"=="+" start "" "%Q0%"
  if "%OpenFolder%"=="+" call :Wait 1
  if "%OpenFolder%"=="+" explorer "%cd%\Unpacked"

  popd
  set "Success=true"
Exit /B 0


:Proc [䠩� dta] [��室�� ���� ����࠭⨭������ 䠩��] [isUVs]
  set "NewName=%~nx2"
  set "NewName=%NewName:?=_%"
  set "NewName=%NewName:<=(%"
  set "NewName=%NewName:>=)%"
  set "NewName=%NewName::=%%AddExtension%"
  set /a Progress=%FilesCur% * 100 / %FilesCnt%
  title %Progress%%% - AVZ DeQuarantine Script
  set /p "=%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%DEL%%Progress%%% - ��२���������. "<NUL

  :: �᫨ � 楫���� ����� ⠪�� 䠩� 㦥 ����: ���६��� +1 � ���᪮� ᢮�������
  call :GetEmptyName "Unpacked" "%NewName%" "NewName" SysNum

  :: ������� ��� �㦭� ������ [ . \ Unpacked \ < �ਣ����쭮� ��� 䠩�� > ]
  if exist "%~1" >NUL %QrCommand% /y "%~1" "Unpacked\%NewName%" && set /A SuccessCopied+=1 || (
    if not Exist "Unpacked" (
      md "Unpacked"
      >NUL %QrCommand% /y "%~1" "Unpacked\%NewName%" && set /A SuccessCopied+=1
    )
  )

  if exist "Unpacked\%NewName%" (
    set result=OK
  ) else (
    set result=--
    if exist "%~1" (
      echo �� ���� ᪮��஢��� "%~1" -^> "Unpacked\%NewName%"
      call :chcp 1251
      echo �� ���� ᪮��஢��� "%~1" -^> "Unpacked\%NewName%">> "%Q2%"
      chcp 866 >NUL
    ) else (
      echo ��������� 䠩� ��࠭⨭� "%~1"
      rem echo -^> ����������� ���� ��������� "%~1">> "%Q2%"
      call :chcp 1251
      echo -^> ��������� 䠩� ��࠭⨭� "%~1">> "%Q2%"
      chcp 866 >NUL
    )
  )

  >nul chcp 1251
  >>"%Q3%" echo %result%  %~1	"%~2"	-^>	"%NewName%"
  >nul chcp 866
  set /A Files+=1
Exit /B

:chcp [cp]
  >NUL chcp %~1
exit /b

:Wait
  if "%RemoveDelays%"=="+" exit /b
  set /a packets=%~1+1
  Timeout /? 2>NUL 1>&2 && >NUL Timeout /NOBREAK /T %~1 || ping -n %packets% 127.1 >NUL
exit /b

:GetEmptyName %1-Folder %2-FileName %3-Var.Return %4-out_idx
  if exist "%~1\%~2" (
    if defined "Empty[%~2]" (
      call :CheckKnownName "%~1" "%~2" "%~3" !"Empty[%~2]"! "%~4" || (
        set SystemNum=0
        call :GetEmptyNameFast "%~1" "%~2" "%~3" "%~4"
      )
    ) else (
      set SystemNum=0
      call :GetEmptyNameFast "%~1" "%~2" "%~3" "%~4"
    )
  ) else (
    set "%~3=%~2"
    set "%~4=0"
  )
exit /B

:CheckKnownName %1-Folder %2-FileName %3-Var.Return %4-in_Last_index %5-out_new_index
  set idx=%~4
  set /a idx+=1
  set "NewFileName=%~n2 (%idx%)%~x2"
  if exist "%~1\%NewFileName%" exit /B 1
  set "%~3=%NewFileName%"
  set "!qt!Empty[%~2]!qt!=%idx%"
  set "%~5=%idx%"
exit /B 0

:GetEmptyNameFast %1-Folder %2-FileName %3-Var.Return
  set /a SystemNum+=1
  Set "NewFileName=%~n2 (%SystemNum%)%~x2"
  if exist "%~1\%NewFileName%" (
    goto GetEmptyNameFast
  ) else (
    set "%~3=%NewFileName%"
    set "!qt!Empty[%~2]!qt!=%SystemNum%"
    set "%~4=%SystemNum%"
  )
exit /B
