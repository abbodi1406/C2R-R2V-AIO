<!-- : Begin batch script
@setlocal DisableDelayedExpansion
@set uivr=v20u
@echo off

:: change to 0 to keep Office C2R vNext license (subscription or lifetime)
set vNextOverride=1

:: change to 1 to enable debug mode
set _Debug=0

:: change to 0 to enable debug mode without cleaning or converting
set _Cnvrt=1

:: change to 1 to use VBScript to access WMI
:: automatically enabled if wmic.exe is not installed
set WMI_VBS=0

:: change to 1 to use Windows PowerShell to access WMI
:: automatically enabled if wmic.exe and VBScript are not installed
set WMI_PS=0

:: ##################################################################

set _args=
set _args=%*
if not defined _args goto :NoProgArgs
for %%A in (%*) do (
if /i "%%A"=="-wow" set _rel1=1
if /i "%%A"=="-arm" set _rel2=1
)
:NoProgArgs
set "_cmdf=%~f0"
if exist "%SystemRoot%\Sysnative\cmd.exe" if not defined _rel1 (
setlocal EnableDelayedExpansion
start %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" -wow "
exit /b
)
if exist "%SystemRoot%\SysArm32\cmd.exe" if /i %PROCESSOR_ARCHITECTURE%==AMD64 if not defined _rel2 (
setlocal EnableDelayedExpansion
start %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" -arm "
exit /b
)
set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "SysPath=%SystemRoot%\Sysnative"
set "Path=%SystemRoot%\Sysnative;%SystemRoot%;%SystemRoot%\Sysnative\Wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%Path%"
)
set "_ln============================================================="
set "_err===== ERROR ===="
set "_psc=powershell -nop -c"
if /i "%PROCESSOR_ARCHITECTURE%"=="amd64" set "xBit=x64"&set "xOS=x64"
if /i "%PROCESSOR_ARCHITECTURE%"=="arm64" set "xBit=x86"&set "xOS=A64"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" set "xBit=x86"&set "xOS=x86"
if /i "%PROCESSOR_ARCHITEW6432%"=="amd64" set "xBit=x64"&set "xOS=x64"
if /i "%PROCESSOR_ARCHITEW6432%"=="arm64" set "xBit=x86"&set "xOS=A64"
set winbuild=1
for /f "tokens=6 delims=[]. " %%# in ('ver') do set winbuild=%%#
set _cwmi=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" (
cmd /c "wmic path Win32_ComputerSystem get CreationClassName /value" 2>nul | find /i "ComputerSystem" 1>nul && set _cwmi=1
)
set _pwsh=1
for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" set _pwsh=0
if not exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" set _pwsh=0
2>nul %_psc% $ExecutionContext.SessionState.LanguageMode | find /i "Full" 1>nul || set _pwsh=0

if %_Cnvrt% NEQ 1 set _Debug=1

set _fC2R=TheEnd
set "_Null=1>nul 2>nul"
set "_temp=%SystemRoot%\Temp"
reg.exe query HKU\S-1-5-19 %_Null% || goto :E_Admin

:Passed
if not exist "%_temp%\" mkdir "%_temp%" %_Null%
set "_onat=HKLM\SOFTWARE\Microsoft\Office"
set "_owow=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office"
set "_batf=%~f0"
set "_batp=%_batf:'=''%"
set "_Local=%LocalAppData%"
set "_log=%~dpn0"
set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "_dsk=%%b"
if exist "%PUBLIC%\Desktop\desktop.ini" set "_dsk=%PUBLIC%\Desktop"
set _WSH=1
reg query "HKCU\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled 2>nul | find /i "0x0" 1>nul && (set _WSH=0)
reg query "HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled 2>nul | find /i "0x0" 1>nul && (set _WSH=0)
if %_cwmi% EQU 0 if %WMI_PS% EQU 0 if %_WSH% EQU 1 if exist "%SysPath%\vbscript.dll" set WMI_VBS=1
if %_cwmi% EQU 0 if %WMI_VBS% EQU 0 if %_pwsh% EQU 1 set WMI_PS=1
if %_cwmi% EQU 0 if %WMI_VBS% EQU 0 if %WMI_PS% EQU 0 goto :E_WMI
if %WMI_VBS% NEQ 0 if %WMI_PS% EQU 0 (
if %_WSH% EQU 0 goto :E_WSH
if not exist "%SysPath%\vbscript.dll" goto :E_VBS
set _cwmi=0
)
if %WMI_PS% NEQ 0 (
if %_pwsh% EQU 0 goto :E_PWS
set _cwmi=0
set WMI_VBS=0
)

set "_csq=cscript.exe //NoLogo //Job:WmiQuery "%~nx0?.wsf""
set "_csm=cscript.exe //NoLogo //Job:WmiMethod "%~nx0?.wsf""
set "_csp=cscript.exe //NoLogo //Job:WmiPKey "%~nx0?.wsf""

setlocal EnableDelayedExpansion
copy /y nul "!_work!\#.rw" %_Null% && (
if exist "!_work!\#.rw" del /f /q "!_work!\#.rw"
) || (
set "_log=!_dsk!\%~n0"
)
pushd "!_work!"

if %_Debug% EQU 0 (
  set "_Nul1=1>nul"
  set "_Nul2=2>nul"
  set "_Nul6=2^>nul"
  set "_Nul3=1>nul 2>nul"
  call :Begin
) else (
  set "_Nul1="
  set "_Nul2="
  set "_Nul6="
  set "_Nul3="
  echo.
  echo Running in Debug Mode...
  echo The window will be closed when finished
  echo.
  echo writing debug log to:
  echo "!_log!_Debug.log"
  @echo on
  @prompt $G
  @call :Begin >"!_log!_tmp.log" 2>&1 &cmd /u /c type "!_log!_tmp.log">"!_log!_Debug.log"&del "!_log!_tmp.log"
)
@color 07
@title %ComSpec%
@echo off
@exit /b

:Begin
@color 1F
set "_title=Office Click-to-Run Retail-to-Volume %uivr%"
title %_title%
set "_SLMGR=%SysPath%\slmgr.vbs"
set "_oApp=0ff1ce15-a989-479d-af46-f275c6370663"
if %_Debug% EQU 0 (
set "_cscript=cscript.exe //NoLogo //B"
) else (
set "_cscript=cscript.exe //NoLogo"
)
set _LTS19=0
set _LTS21=0
set _LTS24=0
set "_tag="&set "_ons= 2016"

echo %_ln%
echo Running C2R-R2V %uivr%
echo %_ln%

if %winbuild% LSS 7601 (
set "msg=Windows 7 SP1 is the minimum supported OS"
goto :%_fC2R%
)
sc query ClickToRunSvc %_Nul3%
set error1=%errorlevel%
sc query OfficeSvc %_Nul3%
set error2=%errorlevel%
if %error1% EQU 1060 if %error2% EQU 1060 (
set "msg=Office C2R service is not detected"
goto :%_fC2R%
)

set _Office16=0
for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (
  set _Office16=1
)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (
  set _Office16=1
)
set _Office15=0
for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\15.0\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (
  set _Office15=1
)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\15.0\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (
  set _Office15=1
)
if %_Office16% EQU 0 if %_Office15% EQU 0 (
set "msg=Office C2R InstallPath is not detected"
goto :%_fC2R%
)

:Reg16istry
if %_Office16% EQU 0 goto :Reg15istry
set "_InstallRoot="
set "_ProductIds="
set "_GUID="
set "_Config="
set "_PRIDs="
set "_LicensesPath="
set "_Integrator="
for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\ClickToRun /v InstallPath" %_Nul6%') do (set "_InstallRoot=%%b\root")
if not "%_InstallRoot%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\ClickToRun /v InstallPath" %_Nul6%') do (set "_OSPPVBS=%%b\Office16\OSPP.VBS")
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\ClickToRun /v PackageGUID" %_Nul6%') do (set "_GUID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_ProductIds=%%b")
  set "_Config=%_onat%\ClickToRun\Configuration"
  set "_PRIDs=%_onat%\ClickToRun\ProductReleaseIDs"
) else (
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\ClickToRun /v InstallPath" %_Nul6%') do (set "_InstallRoot=%%b\root")
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\ClickToRun /v InstallPath" %_Nul6%') do (set "_OSPPVBS=%%b\Office16\OSPP.VBS")
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\ClickToRun /v PackageGUID" %_Nul6%') do (set "_GUID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_ProductIds=%%b")
  set "_Config=%_owow%\ClickToRun\Configuration"
  set "_PRIDs=%_owow%\ClickToRun\ProductReleaseIDs"
)
set "_LicensesPath=%_InstallRoot%\Licenses16"
set "_Integrator=%_InstallRoot%\integration\integrator.exe"
for /f "skip=2 tokens=2*" %%a in ('"reg query %_PRIDs% /v ActiveConfiguration" %_Nul6%') do set "_PRIDs=%_PRIDs%\%%b"
if "%_ProductIds%"=="" (
if %_Office15% EQU 0 (set "msg=Office C2R ProductIDs are not detected"&goto :%_fC2R%) else (goto :Reg15istry)
)
if not exist "%_LicensesPath%\ProPlus*.xrm-ms" (
if %_Office15% EQU 0 (set "msg=Office C2R Licenses files are not detected"&goto :%_fC2R%) else (goto :Reg15istry)
)
if not exist "%_Integrator%" (
if %_Office15% EQU 0 (set "msg=Office C2R Licenses Integrator is not detected"&goto :%_fC2R%) else (goto :Reg15istry)
)
if exist "%_LicensesPath%\Word2019VL_KMS_Client_AE*.xrm-ms" (set _LTS19=1&set "_tag=2019"&set "_ons= 2019")
if exist "%_LicensesPath%\Word2021VL_KMS_Client_AE*.xrm-ms" (set _LTS21=1)
if exist "%_LicensesPath%\Word2024VL_KMS_Client_AE*.xrm-ms" (set _LTS24=1)
if %winbuild% LSS 10240 if !_LTS21! EQU 1 (set "_tag=2021"&set "_ons= 2021")
if %_Office15% EQU 0 goto :CheckC2R

:Reg15istry
set "_Install15Root="
set "_Product15Ids="
set "_Con15fig="
set "_PR15IDs="
set "_OSPP15Ready="
set "_Licenses15Path="
for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\15.0\ClickToRun /v InstallPath" %_Nul6%') do (set "_Install15Root=%%b\root")
if not "%_Install15Root%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\15.0\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=%_onat%\15.0\ClickToRun\Configuration /v ProductReleaseIds"
  set "_PR15IDs=%_onat%\15.0\ClickToRun\ProductReleaseIDs"
  set "_OSPP15Ready=%_onat%\15.0\ClickToRun\Configuration"
) else (
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\15.0\ClickToRun /v InstallPath" %_Nul6%') do (set "_Install15Root=%%b\root")
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\15.0\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=%_owow%\15.0\ClickToRun\Configuration /v ProductReleaseIds"
  set "_PR15IDs=%_owow%\15.0\ClickToRun\ProductReleaseIDs"
  set "_OSPP15Ready=%_owow%\15.0\ClickToRun\Configuration"
)
set "_OSPP15ReadT=REG_SZ"
if "%_Product15Ids%"=="" (
reg query %_onat%\15.0\ClickToRun\propertyBag /v productreleaseid %_Nul3% && (
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\15.0\ClickToRun\propertyBag /v productreleaseid" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=%_onat%\15.0\ClickToRun\propertyBag /v productreleaseid"
  set "_OSPP15Ready=%_onat%\15.0\ClickToRun"
  set "_OSPP15ReadT=REG_DWORD"
  )
reg query %_owow%\15.0\ClickToRun\propertyBag /v productreleaseid %_Nul3% && (
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\15.0\ClickToRun\propertyBag /v productreleaseid" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=%_owow%\15.0\ClickToRun\propertyBag /v productreleaseid"
  set "_OSPP15Ready=%_owow%\15.0\ClickToRun"
  set "_OSPP15ReadT=REG_DWORD"
  )
)
set "_Licenses15Path=%_Install15Root%\Licenses"
set _OSPP15VBS=
for %%G in (
"%ProgramFiles%"
"%ProgramW6432%"
"%ProgramFiles(x86)%"
) do if exist "%%~G\Microsoft Office\Office15\OSPP.VBS" (
if not defined _OSPP15VBS set "_OSPP15VBS=%%~G\Microsoft Office\Office15\OSPP.VBS"
)
if "%_Product15Ids%"=="" (
if %_Office16% EQU 0 (set "msg=Office 2013 C2R ProductIDs are not detected"&goto :%_fC2R%) else (goto :CheckC2R)
)
if not exist "%_Licenses15Path%\ProPlus*.xrm-ms" (
if %_Office16% EQU 0 (set "msg=Office 2013 C2R Licenses files are not detected"&goto :%_fC2R%) else (goto :CheckC2R)
)
if %winbuild% LSS 9200 if "%_OSPP15VBS%"=="" (
if %_Office16% EQU 0 (set "msg=Office 2013 C2R Licensing tool OSPP.vbs is not detected"&goto :%_fC2R%) else (goto :CheckC2R)
)

:CheckC2R
set _OMSI=0
if %_Office16% EQU 0 (
for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\16.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\EntityPicker.dll" set _OMSI=1
for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\16.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\EntityPicker.dll" set _OMSI=1
)
if %_Office15% EQU 0 (
for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\15.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\EntityPicker.dll" set _OMSI=1
for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\15.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\EntityPicker.dll" set _OMSI=1
)
if %winbuild% GEQ 9200 (
set _spp=SoftwareLicensingProduct
set _sps=SoftwareLicensingService
set "_vbsi=%_SLMGR% /ilc "
set "_vbsf=%_SLMGR% /ilc "
) else (
set _spp=OfficeSoftwareProtectionProduct
set _sps=OfficeSoftwareProtectionService
set _vbsi="!_OSPP15VBS!" /inslic:
set _vbsf="!_OSPPVBS!" /inslic:
)
set "_wmi="
call :qrSingle %_sps% Version
for /f "tokens=2 delims==" %%# in ('%_qr%') do set _wmi=%%#
if "%_wmi%"=="" (
set "msg=%_sps% WMI version is not detected"
goto :%_fC2R%
)
echo.
echo %_ln%
echo Checking Office Licenses...
echo %_ln%
call :qrQuery %_spp% "ApplicationID='%_oApp%' AND Description LIKE '%%%%KMSCLIENT%%%%'" LicenseFamily
%_qr% %_Nul2% | findstr /I /C:"Office" %_Nul1% && (set _KMS=1) || (set _KMS=0)
call :qrQuery %_spp% "ApplicationID='%_oApp%' AND Description LIKE '%%%%TIMEBASED%%%%'" LicenseFamily
%_qr% %_Nul2% | findstr /I /C:"Office" %_Nul1% && (set _Time=1) || (set _Time=0)
call :qrQuery %_spp% "ApplicationID='%_oApp%' AND Description LIKE '%%%%Trial%%%%'" LicenseFamily
%_qr% %_Nul2% | findstr /I /C:"Office" %_Nul1% && (set _Time=1)
call :qrQuery %_spp% "ApplicationID='%_oApp%' AND Description LIKE '%%%%Grace%%%%'" LicenseFamily
%_qr% %_Nul2% | findstr /I /C:"Office" %_Nul1% && (set _Grace=1) || (set _Grace=0)
call :qrQuery %_spp% "ApplicationID='%_oApp%'" LicenseFamily
%_qr% > "!_temp!\crvchk.txt" 2>&1
call :qrQuery %_spp% "ApplicationID='%_oApp%' AND LicenseFamily LIKE 'Office16O365%%%%'" LicenseFamily
if %_Office16% EQU 1 find /i "Office16MondoVL_KMS_Client" "!_temp!\crvchk.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% || (set _Grace=1)
)
call :qrQuery %_spp% "ApplicationID='%_oApp%' AND LicenseFamily LIKE 'OfficeO365%%%%'" LicenseFamily
if %_Office15% EQU 1 find /i "OfficeMondoVL_KMS_Client" "!_temp!\crvchk.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% || (set _Grace=1)
)
if %_Time% EQU 0 if %_Grace% EQU 0 if %_KMS% EQU 1 (
set "msg=No Conversion or Cleanup required"
goto :%_fC2R%
)

call :subOffice

if %vNextOverride% EQU 1 if %_Cnvrt% EQU 1 (
set sub_o365=0
set sub_proj=0
set sub_vsio=0
if %sub_next% EQU 1 (
  reg delete HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing /f %_Nul3%
  rmdir /s /q "!_Local!\Microsoft\Office\Licenses\" %_Nul3%
  rmdir /s /q "!ProgramData!\Microsoft\Office\Licenses\" %_Nul3%
  )
set sub_next=0
)

set _Retail=0
set "_ocq=ApplicationID='%_oApp%' AND LicenseStatus='1' AND PartialProductKey is not NULL"
call :qrQuery %_spp% "%_ocq%" Description fix
%_qr% %_Nul2% |findstr /V /R "^$" >"!_temp!\crvRetail.txt"
find /i "RETAIL channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
find /i "RETAIL(MAK) channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
find /i "TIMEBASED_SUB channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
set rancopp=0
if %_Cnvrt% EQU 1 if %sub_next% EQU 0 if %_Retail% EQU 0 if %_OMSI% EQU 0 if %_pwsh% equ 1 (
echo.
echo %_ln%
echo Cleaning Current Office Licenses...
echo %_ln%
set rancopp=1
%_Nul3% %_psc% "$f=[IO.File]::ReadAllText('!_batp!') -split ':embdbin\:.*';iex ($f[1])"
title %_title%
)

:R16V
echo.
echo %_ln%
echo Checking installed Office Products...
echo %_ln%
echo.
set _SubID=O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud
set _O16O365=0
set _C16Msg=0&set _C16Vol=0
set _C15Msg=0&set _C15Vol=0
call :qrQuery %_spp% "%_ocq%" LicenseFamily fix
if %_Retail% EQU 1 %_qr% %_Nul2% |findstr /V /R "^$" >"!_temp!\crvRetail.txt"
call :qrQuery %_spp% "ApplicationID='%_oApp%'" LicenseFamily fix
%_qr% %_Nul2% |findstr /V /R "^$" >"!_temp!\crvVolume.txt" 2>&1

if %_Office16% EQU 0 goto :R15V

set _S24ID=ProPlus2024,Standard2024
set _S21ID=ProPlus2021,Standard2021
set _S19ID=ProPlus2019,Standard2019
set _S16ID=Mondo,Standard
set _P24ID=ProjectPro2024,ProjectStd2024
set _P21ID=ProjectPro2021,ProjectStd2021
set _P19ID=ProjectPro2019,ProjectStd2019
set _P16ID=ProjectPro,ProjectStd
set _I24ID=VisioPro2024,VisioStd2024
set _I21ID=VisioPro2021,VisioStd2021
set _I19ID=VisioPro2019,VisioStd2019
set _I16ID=VisioPro,VisioStd
set _A24ID=Excel2024,Outlook2024,PowerPoint2024,Word2024
set _A21ID=Excel2021,Outlook2021,PowerPoint2021,Publisher2021,Word2021
set _A19ID=Excel2019,Outlook2019,PowerPoint2019,Publisher2019,Word2019
set _A16ID=Excel,Outlook,PowerPoint,Publisher,Word
set _E24ID=Access2024,SkypeforBusiness2024
set _E21ID=Access2021,SkypeforBusiness2021
set _E19ID=Access2019,SkypeforBusiness2019
set _E16ID=Access,SkypeforBusiness
set _R24ID=Professional2024,HomeBusiness2024,HomeStudent2024,Home2024
set _R21ID=Professional2021,HomeBusiness2021,HomeStudent2021
set _R19ID=Professional2019,HomeBusiness2019,HomeStudent2019
set _R16ID=Professional,HomeBusiness,HomeStudent,%_SubID%
set _V24ID=%_S24ID%,%_A24ID%,%_E24ID%,%_P24ID%,%_I24ID%
set _V21ID=%_S21ID%,%_A21ID%,%_E21ID%,%_P21ID%,%_I21ID%
set _V19ID=%_S19ID%,%_A19ID%,%_E19ID%,%_P19ID%,%_I19ID%
set _V16ID=%_S16ID%,%_A16ID%,%_E16ID%,%_P16ID%,%_I16ID%
set _RetID=%_R24ID%,%_V24ID%,%_R21ID%,%_V21ID%,%_R19ID%,%_V19ID%,%_R16ID%,%_V16ID%
set _Suites=ProPlus,%_S16ID%,%_R16ID%,%_S19ID%,%_R19ID%,%_S21ID%,%_R21ID%,%_S24ID%,%_R24ID%
set _PrjSKU=%_P16ID%,%_P19ID%,%_P21ID%,%_P24ID%
set _VisSKU=%_I16ID%,%_I19ID%,%_I21ID%,%_I24ID%

echo %_ProductIds%>"!_temp!\crvProductIds.txt"
for %%a in (%_RetID%,ProPlus,OneNote,Publisher2024,Home,Home2019,Home2021) do (
set _%%a=0
)
for %%a in (%_RetID%,OneNote) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProductIds.txt" %_Nul1% && set _%%a=1
)
if !_LTS24! EQU 0 for %%a in (%_V24ID%) do (
set _%%a=0
)
if !_LTS24! EQU 1 for %%a in (%_V24ID%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office24%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0&set _C16Vol=1) || (set _%%a=1)
  )
)
if !_LTS21! EQU 0 for %%a in (%_V21ID%) do (
set _%%a=0
)
if !_LTS21! EQU 1 for %%a in (%_V21ID%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office21%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0&set _C16Vol=1) || (set _%%a=1)
  )
)
if !_LTS19! EQU 0 for %%a in (%_V19ID%) do (
set _%%a=0
)
if !_LTS19! EQU 1 for %%a in (%_V19ID%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office19%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0&set _C16Vol=1) || (set _%%a=1)
  )
)
for %%a in (%_V16ID%,OneNote) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office16%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0&set _C16Vol=1) || (set _%%a=1)
  )
)
reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && (
  find /i "Office16ProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0&set _C16Vol=1) || (set _ProPlus=1)
)
reg query %_PRIDs%\ProPlusVolume.16 %_Nul3% && (
  find /i "Office16ProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0&set _C16Vol=1) || (set _ProPlus=1)
)
if %_Retail% EQU 1 for %%a in (%_RetID%,OneNote) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office16%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office16%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office16%%aR_Sub" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office16%%aR_PIN" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office16%%aE5R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office16%%aEDUR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office16%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office16%%aO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office16%%aCO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office16%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office16%%aXC2RVL_MAKC2R" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office19%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office19%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office19%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office19%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office21%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office21%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office21%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office21%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office24%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office24%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office24%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office24%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  )
)
if %_Retail% EQU 1 reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && (
  find /i "Office16ProPlusR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0)
  find /i "Office16ProPlusR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0)
  find /i "Office16ProPlusMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0)
  find /i "Office16ProPlusVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0)
)
call :qrQuery %_spp% "ApplicationID='%_oApp%' AND LicenseFamily like 'Office16O365%%%%'" LicenseFamily
find /i "Office16MondoVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% && (
  for %%a in (%_SubID%) do set _%%a=0
  )
)
if %sub_o365% EQU 1 (
for %%a in (%_Suites%) do set _%%a=0
echo Microsoft Office is activated with a vNext license.
echo.
)
if %sub_proj% EQU 1 (
for %%a in (%_PrjSKU%) do set _%%a=0
echo Microsoft Project is activated with a vNext license.
echo.
)
if %sub_vsio% EQU 1 (
for %%a in (%_VisSKU%) do set _%%a=0
echo Microsoft Visio is activated with a vNext license.
echo.
)

if %_Cnvrt% NEQ 1 (if %_Office15% EQU 1 (goto :R15V) else (set "msg=Finished"&goto :%_fC2R%))

for %%a in (%_RetID%,ProPlus,OneNote) do if !_%%a! EQU 1 (
set _C16Msg=1&set _C16Vol=1
)
if %_C16Msg% EQU 1 (
echo.
echo %_ln%
echo Installing Office Volume Licenses...
echo %_ln%
echo.
)
if %_C16Msg% EQU 0 goto :endRV16

set "_arr="
for %%# in ("!_LicensesPath!\client-issuance-*.xrm-ms") do (
if %WMI_PS% NEQ 0 (
  if defined _arr (set "_arr=!_arr!;"!_LicensesPath!\%%~nx#"") else (set "_arr="!_LicensesPath!\%%~nx#"")
  ) else (
  %_cscript% %_vbsf%"!_LicensesPath!\%%~nx#"
  )
)
if %WMI_PS% NEQ 0 (
  %_Nul3% %_psc% "$sls='%_sps%'; $f=[IO.File]::ReadAllText('!_batp!') -split ':embdxrm\:.*'; iex ($f[1]); InstallLicenseArr '!_arr!'; InstallLicenseFile '"!_LicensesPath!\pkeyconfig-office.xrm-ms"'"
  ) else (
  %_cscript% %_vbsf%"!_LicensesPath!\pkeyconfig-office.xrm-ms"
  )

set _jump=0
set _DidO365=0
if !_Mondo! EQU 1 (
call :InsLic Mondo
)
if !_O365ProPlus! EQU 1 (
set _DidO365=1
echo O365ProPlus 2016 Suite ^<-^> Mondo 2016 Licenses
call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365Business! EQU 1 if !_DidO365! EQU 0 (
set _DidO365=1
echo O365Business 2016 Suite ^<-^> Mondo 2016 Licenses
call :InsLic O365Business NCHRJ-3VPGW-X73DM-6B36K-3RQ6B
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365SmallBusPrem! EQU 1 if !_DidO365! EQU 0 (
set _DidO365=1
echo O365SmallBusPrem 2016 Suite ^<-^> Mondo 2016 Licenses
call :InsLic O365SmallBusPrem 3FBRX-NFP7C-6JWVK-F2YGK-H499R
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365HomePrem! EQU 1 if !_DidO365! EQU 0 (
set _DidO365=1
echo O365HomePrem 2016 Suite ^<-^> Mondo 2016 Licenses
call :InsLic O365HomePrem 9FNY8-PWWTY-8RY4F-GJMTV-KHGM9
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365EduCloud! EQU 1 if !_DidO365! EQU 0 (
set _DidO365=1
echo O365EduCloud 2016 Suite ^<-^> Mondo 2016 Licenses
call :InsLic O365EduCloud 8843N-BCXXD-Q84H8-R4Q37-T3CPT
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_DidO365! EQU 1 set _jump=1&set _O16O365=1
if !_Mondo! EQU 1 if !_DidO365! EQU 0 (
echo Mondo 2016 Suite
call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
goto :endRV16
)

for %%a in (%_P16ID%,%_I16ID%) do (
  if !_%%a2024! EQU 1 (echo %%a 2024 SKU&call :InsLic %%a2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (echo %%a 2021 SKU&call :InsLic %%a2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (echo %%a 2019 SKU -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (echo %%a 2016 SKU -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
)

if !_jump! EQU 1 goto :endRV16

for %%a in (ProPlus) do (
  if !_%%a2024! EQU 1 (set _jump=1&echo %%a 2024 Suite&call :InsLic ProPlus2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (set _jump=1&echo %%a 2021 Suite&call :InsLic ProPlus2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (set _jump=1&echo %%a 2019 Suite -^> ProPlus%_ons% Licenses&call :InsLic ProPlus%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (set _jump=1&echo %%a 2016 Suite -^> ProPlus%_ons% Licenses&call :InsLic ProPlus%_tag%)
)
if !_jump! EQU 1 goto :endRV16

for %%a in (Professional) do (
  if !_%%a2024! EQU 1 (set _jump=1&echo %%a 2024 Suite -^> ProPlus 2024 Licenses&call :InsLic ProPlus2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (set _jump=1&echo %%a 2021 Suite -^> ProPlus 2021 Licenses&call :InsLic ProPlus2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (set _jump=1&echo %%a 2019 Suite -^> ProPlus%_ons% Licenses&call :InsLic ProPlus%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (set _jump=1&echo %%a 2016 Suite -^> ProPlus%_ons% Licenses&call :InsLic ProPlus%_tag%)
)
if !_jump! EQU 1 goto :endRV16

for %%a in (SkypeforBusiness) do (
  if !_%%a2024! EQU 1 (echo %%a 2024 App&call :InsLic %%a2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (echo %%a 2021 App&call :InsLic %%a2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (echo %%a 2019 App -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (echo %%a 2016 App -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
)

for %%a in (Access) do (
  if !_%%a2024! EQU 1 (echo %%a 2024 App&call :InsLic %%a2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (echo %%a 2021 App&call :InsLic %%a2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (echo %%a 2019 App -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (echo %%a 2016 App -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
)

for %%a in (Standard) do (
  if !_%%a2024! EQU 1 (set _jump=1&echo %%a 2024 Suite&call :InsLic Standard2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (set _jump=1&echo %%a 2021 Suite&call :InsLic Standard2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (set _jump=1&echo %%a 2019 Suite -^> Standard%_ons% Licenses&call :InsLic Standard%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (set _jump=1&echo %%a 2016 Suite -^> Standard%_ons% Licenses&call :InsLic Standard%_tag%)
)
if !_jump! EQU 1 goto :endRV16

for %%a in (HomeBusiness,HomeStudent,Home) do (
  if !_%%a2024! EQU 1 (set _jump=1&echo %%a 2024 Suite -^> Standard 2024 Licenses&call :InsLic Standard2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (set _jump=1&echo %%a 2021 Suite -^> Standard 2021 Licenses&call :InsLic Standard2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (set _jump=1&echo %%a 2019 Suite -^> Standard%_ons% Licenses&call :InsLic Standard%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (set _jump=1&echo %%a 2016 Suite -^> Standard%_ons% Licenses&call :InsLic Standard%_tag%)
)
if !_jump! EQU 1 goto :endRV16

for %%a in (%_A16ID%) do (
  if !_%%a2024! EQU 1 (echo %%a 2024 App&call :InsLic %%a2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (echo %%a 2021 App&call :InsLic %%a2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (echo %%a 2019 App -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (echo %%a 2016 App -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
)
for %%a in (OneNote) do (
  if !_%%a! EQU 1 (echo %%a 2016 App&call :InsLic %%a)
)

:endRV16
if %_Office15% EQU 0 goto :GVLKC2R

:R15V
set _S15ID=Mondo,Standard
set _P15ID=ProjectPro,ProjectStd
set _I15ID=VisioPro,VisioStd
set _A15ID=Excel,Groove,InfoPath,OneNote,Outlook,PowerPoint,Publisher,Word
set _E15ID=Access,Lync
set _V15ID=%_S15ID%,%_A15ID%,%_E15ID%,%_P15ID%,%_I15ID%
set _R15ID=%_V15ID%,SPD,Professional,HomeBusiness,HomeStudent,%_SubID%

echo %_Product15Ids%>"!_temp!\crvProduct15s.txt"
for %%a in (%_R15ID%,ProPlus) do (
set _%%a=0
)
for %%a in (%_R15ID%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProduct15s.txt" %_Nul1% && set _%%a=1
)
for %%a in (%_V15ID%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProduct15s.txt" %_Nul1% && (
  find /i "Office%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0&set _C15Vol=1) || (set _%%a=1)
  )
)
reg query %_PR15IDs%\Active\ProPlusRetail\x-none %_Nul3% && (
  find /i "OfficeProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0&set _C15Vol=1) || (set _ProPlus=1)
)
reg query %_PR15IDs%\Active\ProPlusVolume\x-none %_Nul3% && (
  find /i "OfficeProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0&set _C15Vol=1) || (set _ProPlus=1)
)
if %_Retail% EQU 1 for %%a in (%_R15ID%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProduct15s.txt" %_Nul1% && (
  find /i "Office%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office%%aR_Sub" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office%%aR_PIN" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office%%aO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office%%aCO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  find /i "Office%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0)
  )
)
if %_Retail% EQU 1 reg query %_PR15IDs%\Active\ProPlusRetail\x-none %_Nul3% && (
  find /i "OfficeProPlusR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0)
  find /i "OfficeProPlusR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0)
  find /i "OfficeProPlusMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0)
  find /i "OfficeProPlusVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0)
)
call :qrQuery %_spp% "ApplicationID='%_oApp%' AND LicenseFamily like 'OfficeO365%%%%'" LicenseFamily
find /i "OfficeMondoVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% && (
  for %%a in (%_SubID%) do set _%%a=0
  )
)

if %_Cnvrt% NEQ 1 (set "msg=Finished"&goto :%_fC2R%)

for %%a in (%_R15ID%,ProPlus) do if !_%%a! EQU 1 (
set _C15Msg=1&set _C15Vol=1
)
if %_C15Msg% EQU 1 if %_C16Msg% EQU 0 (
echo.
echo %_ln%
echo Installing Office Volume Licenses...
echo %_ln%
echo.
)
if %_C15Msg% EQU 0 goto :endRV15

set "_arr="
for %%# in ("!_Licenses15Path!\client-issuance-*.xrm-ms") do (
if %WMI_PS% NEQ 0 (
  if defined _arr (set "_arr=!_arr!;"!_Licenses15Path!\%%~nx#"") else (set "_arr="!_Licenses15Path!\%%~nx#"")
  ) else (
  %_cscript% %_vbsi%"!_Licenses15Path!\%%~nx#"
  )
)
if %WMI_PS% NEQ 0 (
  %_Nul3% %_psc% "$sls='%_sps%'; $f=[IO.File]::ReadAllText('!_batp!') -split ':embdxrm\:.*'; iex ($f[1]); InstallLicenseArr '!_arr!'; InstallLicenseFile '"!_Licenses15Path!\pkeyconfig-office.xrm-ms"'"
  ) else (
  %_cscript% %_vbsi%"!_Licenses15Path!\pkeyconfig-office.xrm-ms"
  )

set _jump=0
set _DidO365=0
if !_Mondo! EQU 1 (
call :Ins15Lic Mondo
)
if !_O365ProPlus! EQU 1 if !_O16O365! EQU 0 (
set _DidO365=1
echo O365ProPlus 2013 Suite ^<-^> Mondo 2013 Licenses
call :Ins15Lic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365SmallBusPrem! EQU 1 if !_O16O365! EQU 0 if !_DidO365! EQU 0 (
set _DidO365=1
echo O365SmallBusPrem 2013 Suite ^<-^> Mondo 2013 Licenses
call :Ins15Lic O365SmallBusPrem 3FBRX-NFP7C-6JWVK-F2YGK-H499R
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365HomePrem! EQU 1 if !_O16O365! EQU 0 if !_DidO365! EQU 0 (
set _DidO365=1
echo O365HomePrem 2013 Suite ^<-^> Mondo 2013 Licenses
call :Ins15Lic O365HomePrem 9FNY8-PWWTY-8RY4F-GJMTV-KHGM9
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365Business! EQU 1 if !_O16O365! EQU 0 if !_DidO365! EQU 0 (
set _DidO365=1
echo O365Business 2013 Suite ^<-^> Mondo 2013 Licenses
call :Ins15Lic O365Business MCPBN-CPY7X-3PK9R-P6GTT-H8P8Y
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_DidO365! EQU 1 set _jump=1
if !_Mondo! EQU 1 if !_O16O365! EQU 0 if !_DidO365! EQU 0 (
echo Mondo 2013 Suite
call :Ins15Lic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
goto :endRV15
)

for %%a in (%_P15ID%,%_I15ID%) do (
  if !_%%a! EQU 1 (echo %%a 2013 SKU&call :Ins15Lic %%a)
)

if !_Mondo! EQU 0 if !_DidO365! EQU 0 for %%a in (SPD) do (
  if !_%%a! EQU 1 (set _jump=1&echo SharePoint Designer 2013 App -^> Mondo 2013 Licenses&call :Ins15Lic Mondo)
)
if !_jump! EQU 1 goto :endRV15

for %%a in (ProPlus) do (
  if !_%%a! EQU 1 (set _jump=1&echo %%a 2013 Suite&call :Ins15Lic %%a)
)
if !_jump! EQU 1 goto :endRV15

for %%a in (Professional) do (
  if !_%%a! EQU 1 (set _jump=1&echo %%a 2013 Suite -^> ProPlus 2013 Licenses&call :Ins15Lic ProPlus)
)
if !_jump! EQU 1 goto :endRV15

for %%a in (Lync) do (
  if !_%%a! EQU 1 (echo SkypeforBusiness 2015 App&call :Ins15Lic %%a)
)

for %%a in (Access) do (
  if !_%%a! EQU 1 (echo %%a 2013 App&call :Ins15Lic %%a)
)

for %%a in (Standard) do (
  if !_%%a! EQU 1 (set _jump=1&echo %%a 2013 Suite&call :Ins15Lic %%a)
)
if !_jump! EQU 1 goto :endRV15

for %%a in (HomeBusiness,HomeStudent) do (
  if !_%%a! EQU 1 (set _jump=1&echo %%a 2013 Suite -^> Standard 2013 Licenses&call :Ins15Lic Standard)
)
if !_jump! EQU 1 goto :endRV15

for %%a in (%_A15ID%) do (
  if !_%%a! EQU 1 (echo %%a 2013 App&call :Ins15Lic %%a)
)

:endRV15
goto :GVLKC2R

:InsLic
set "_ID=%1Volume"
set "_patt=%1VL_"
set "_pkey="
set "_kpey="
if not "%2"=="" (
set "_ID=%1Retail"
set "_patt=%1R_"
set "_pkey=PidKey=%2"
set "_kpey=%2"
)
reg delete %_Config% /f /v %_ID%.OSPPReady %_Nul3%
"!_Integrator!" /I /License PRIDName=%_ID%.16 %_pkey% PackageGUID="%_GUID%" PackageRoot="!_InstallRoot!" %_Nul1%

set fallback=0
call :qrQuery %_spp% "ApplicationID='%_oApp%'" LicenseFamily fix
%_qr% %_Nul2% | find /i "%_patt%" %_Nul1% || (set fallback=1)
if %fallback% equ 0 goto :IntOK

set "_lsfs="
for %%# in ("!_LicensesPath!\%_patt%*.xrm-ms") do (
set "_lsfs=!_lsfs! %%~nx#"
)
if defined _kpey (
  for %%# in ("!_LicensesPath!\%1DemoR*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1E5R*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1EDUR*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1MSDNR*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1O365R*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1CO365R*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
)
set "_arr="
for %%# in (!_lsfs!) do (
if %WMI_PS% NEQ 0 (
  if defined _arr (set "_arr=!_arr!;"!_LicensesPath!\%%~nx#"") else (set "_arr="!_LicensesPath!\%%~nx#"")
  ) else (
  %_cscript% %_vbsf%"!_LicensesPath!\%%~nx#"
  )
)
if %WMI_PS% NEQ 0 (
  %_Nul3% %_psc% "$sls='%_sps%'; $f=[IO.File]::ReadAllText('!_batp!') -split ':embdxrm\:.*'; iex ($f[1]); InstallLicenseArr '!_arr!'"
  )
call :qrPKey %_sps% %_wmi% %_kpey%
if defined _kpey %_qr% %_Nul3%

:IntOK
reg add %_Config% /f /v %_ID%.OSPPReady /t REG_SZ /d 1 %_Nul1%
reg query %_Config% /v ProductReleaseIds | findstr /I "%_ID%" %_Nul1%
if %errorlevel% NEQ 0 (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Config% /v ProductReleaseIds') do reg add %_Config% /v ProductReleaseIds /t REG_SZ /d "%%b,%_ID%" /f %_Nul1%
)
exit /b

:Ins15Lic
set "_ID=%1Volume"
set "_patt=%1VL_"
set "_pkey="
if not "%2"=="" (
set "_ID=%1Retail"
set "_patt=%1R_"
set "_pkey=%2"
)
reg delete %_OSPP15Ready% /f /v %_ID%.OSPPReady %_Nul3%
set "_arr="
for %%# in ("!_Licenses15Path!\%_patt%*.xrm-ms") do (
if %WMI_PS% NEQ 0 (
  if defined _arr (set "_arr=!_arr!;"!_Licenses15Path!\%%~nx#"") else (set "_arr="!_Licenses15Path!\%%~nx#"")
  ) else (
  %_cscript% %_vbsi%"!_Licenses15Path!\%%~nx#"
  )
)
if %WMI_PS% NEQ 0 (
  %_Nul3% %_psc% "$sls='%_sps%'; $f=[IO.File]::ReadAllText('!_batp!') -split ':embdxrm\:.*'; iex ($f[1]); InstallLicenseArr '!_arr!'"
  )
call :qrPKey %_sps% %_wmi% %_pkey%
if defined _pkey %_qr% %_Nul3%
reg add %_OSPP15Ready% /f /v %_ID%.OSPPReady /t %_OSPP15ReadT% /d 1 %_Nul1%
reg query %_Con15fig% %_Nul2% | findstr /I "%_ID%" %_Nul1%
if %errorlevel% NEQ 0 (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Con15fig% %_Nul6%') do reg add %_Con15fig% /t REG_SZ /d "%%b,%_ID%" /f %_Nul1%
)
exit /b

:GVLKC2R
set _CtRMsg=0
if %_C16Msg% EQU 1 set _CtRMsg=1
if %_C15Msg% EQU 1 set _CtRMsg=1
if %_CtRMsg% EQU 1 (
echo %_ln%
echo Installing Missing KMS Client Keys...
echo %_ln%
echo.
)
call :qrMethod %_sps% Version %_wmi% RefreshLicenseStatus
if %winbuild% GEQ 9200 %_qr% %_Nul3%
for %%# in (15,16,19,21,24) do call :C2RLoc %%#
if %_Retail% EQU 0 (
for %%# in (15) do if !_Loc%%#! EQU 0 call :C2Runi %%#
)
if %_Retail% EQU 0 if %sub_o365% EQU 0 if %sub_proj% EQU 0 if %sub_vsio% EQU 0 (
for %%# in (16,19,21,24) do if !_Loc%%#! EQU 0 call :C2Runi %%#
)
if %_C15Vol% EQU 1 (
for %%# in (15) do if !_Loc%%#! EQU 1 call :C2Rins %%#
)
if %_C16Vol% EQU 1 (
for %%# in (16,19,21,24) do if !_Loc%%#! EQU 1 call :C2Rins %%#
)
call :qrMethod %_sps% Version %_wmi% RefreshLicenseStatus
if %winbuild% GEQ 9200 %_qr% %_Nul3%
if exist "%SysPath%\spp\store_test\2.0\tokens.dat" (
echo.
echo %_ln%
echo Refreshing Windows Insider Preview Licenses...
echo %_ln%
echo.
if %WMI_PS% NEQ 0 (
  %_Nul3% %_psc% "$sls='%_sps%'; $f=[IO.File]::ReadAllText('!_batp!') -split ':embdxrm\:.*'; iex ($f[1]); ReinstallLicenses"
  if !ERRORLEVEL! NEQ 0 %_Nul3% %_psc% "$sls='%_sps%'; $f=[IO.File]::ReadAllText('!_batp!') -split ':embdxrm\:.*'; iex ($f[1]); ReinstallLicenses"
  ) else (
  %_cscript% %_SLMGR% /rilc
  if !ERRORLEVEL! NEQ 0 %_cscript% %_SLMGR% /rilc
  )
)
set "msg=Finished"&goto :%_fC2R%

:C2Runi
call :qrQuery %_spp% "Name LIKE 'Office %~1%%%%' AND PartialProductKey is not NULL" ID
for /f "tokens=2 delims==" %%# in ('%_qr% %_Nul6%') do (set "aID=%%#"&call :UniKey)
exit /b

:C2Rins
call :qrQuery %_spp% "Description LIKE 'Office %1, VOLUME_KMSCLIENT%%%%' AND PartialProductKey is NULL" ID
for /f "tokens=2 delims==" %%# in ('%_qr% %_Nul6%') do (set "aID=%%#"&call :InsKey)
exit /b

:C2RLoc
set _Loc%1=0
if %1 EQU 19 (
if defined _ProductIds reg query %_Config% /v ProductReleaseIds %_Nul2% | findstr 2019 %_Nul1% && set _Loc%1=1
exit /b
)
if %1 EQU 21 (
if defined _ProductIds reg query %_Config% /v ProductReleaseIds %_Nul2% | findstr 2021 %_Nul1% && set _Loc%1=1
exit /b
)
if %1 EQU 24 (
if defined _ProductIds reg query %_Config% /v ProductReleaseIds %_Nul2% | findstr 2024 %_Nul1% && set _Loc%1=1
exit /b
)

for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\%1.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\EntityPicker.dll" (
set _Loc%1=1
)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\%1.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\EntityPicker.dll" (
set _Loc%1=1
)

if %1 EQU 16 if defined _ProductIds (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Config% /v ProductReleaseIds') do echo %%b>"!_temp!\crvO16.txt"
for %%a in (%_V16ID%,ProjectProX,ProjectStdX,VisioProX,VisioStdX) do (
  findstr /I /C:"%%aVolume" "!_temp!\crvO16.txt" %_Nul1% && set _Loc%1=1
  )
for %%a in (%_V16ID%,%_R16ID%) do (
  findstr /I /C:"%%aRetail" "!_temp!\crvO16.txt" %_Nul1% && set _Loc%1=1
  )
reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && set _Loc%1=1
reg query %_PRIDs%\ProPlusVolume.16 %_Nul3% && set _Loc%1=1
exit /b
)

if %1 EQU 15 if defined _Product15Ids (
set _Loc%1=1
exit /b
)

if exist "%ProgramFiles%\Microsoft Office\Office%1\OSPP.VBS" set _Loc%1=1
if not %xOS%==x86 if exist "%ProgramW6432%\Microsoft Office\Office%1\OSPP.VBS" set _Loc%1=1
if not %xOS%==x86 if exist "%ProgramFiles(x86)%\Microsoft Office\Office%1\OSPP.VBS" set _Loc%1=1
exit /b

:subOffice
set kNext=HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\LicensingNext
set sub_next=0
set sub_o365=0
set sub_proj=0
set sub_vsio=0
set _Identity=0
dir /b /s /a:-d "!_Local!\Microsoft\Office\Licenses\*" %_Nul3% && (set _Identity=1&set sub_next=1)
dir /b /s /a:-d "!ProgramData!\Microsoft\Office\Licenses\*" %_Nul3% && (set _Identity=1&set sub_next=1)
if %_Identity% EQU 0 call :officeSub
exit /b

:officeSub
reg query %kNext% %_Nul3% || exit /b
reg query %kNext% | findstr /i /r ".*retail" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x2" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r ".*retail" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x3" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r ".*volume" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x2" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r ".*volume" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x3" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r "project.*" %_Nul2% | find /i "0x2" %_Nul1% && set sub_proj=1
reg query %kNext% | findstr /i /r "project.*" %_Nul2% | find /i "0x3" %_Nul1% && set sub_proj=1
reg query %kNext% | findstr /i /r "visio.*" %_Nul2% | find /i "0x2" %_Nul1% && set sub_vsio=1
reg query %kNext% | findstr /i /r "visio.*" %_Nul2% | find /i "0x3" %_Nul1% && set sub_vsio=1
if %sub_o365% EQU 1 set sub_next=1
if %sub_proj% EQU 1 set sub_next=1
if %sub_vsio% EQU 1 set sub_next=1
exit /b

:UniKey
call :qrMethod %_spp% ID %aID% UninstallProductKey
%_qr% %_Nul3%
exit /b

:InsKey
set _eof=0
for %%A in (
fceda083-1203-402a-8ec4-3d7ed9f3648c,aaea0dc8-78e1-4343-9f25-b69b83dd1bce,4ab4d849-aabc-43fb-87ee-3aed02518891
f3fb2d68-83dd-4c8b-8f09-08e0d950ac3b,76093b1b-7057-49d7-b970-638ebcbfd873,a3b44174-2451-4cd6-b25f-66638bfb9046
0bc88885-718c-491d-921f-6f214349e79c,fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9,500f6619-ef93-4b75-bcb4-82819998a3ca
e914ea6e-a5fa-4439-a394-a9bb3293ca09
1dc00701-03af-4680-b2af-007ffc758a1f
) do (
if /i "%aID%" EQU "%%A" set _eof=1
)
if %_eof% EQU 1 exit /b
set "_key="
call :qrQuery %_spp% "ID='%aID%'" LicenseFamily
for /f "tokens=2 delims==" %%# in ('%_qr%') do echo %%#
call :keys %aID%
if "%_key%"=="" (echo No associated KMS Client key found&echo.&exit /b)
call :qrPKey %_sps% %_wmi% %_key%
%_qr% %_Nul3%
set ERRORCODE=%ERRORLEVEL%
if %ERRORCODE% NEQ 0 (
cmd /c exit /b %ERRORCODE%
echo Failed: 0x!=ExitCode!
)
echo.
exit /b

:keys
if "%~1"=="" exit /b
goto :%1 %_Nul2%

:: Office 2024
:8d368fc1-9470-4be2-8d66-90e836cbb051
set "_key=XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB" &:: Professional Plus
exit /b

:bbac904f-6a7e-418a-bb4b-24c85da06187
set "_key=V28N4-JG22K-W66P8-VTMGK-H6HGR" &:: Standard
exit /b

:f510af75-8ab7-4426-a236-1bfb95c34ff8
set "_key=FQQ23-N4YCY-73HQ3-FM9WC-76HF4" &:: Project Professional
exit /b

:9f144f27-2ac5-40b9-899d-898c2b8b4f81
set "_key=PD3TT-NTHQQ-VC7CY-MFXK3-G87F8" &:: Project Standard
exit /b

:fa187091-8246-47b1-964f-80a0b1e5d69a
set "_key=B7TN8-FJ8V3-7QYCP-HQPMV-YY89G" &:: Visio Professional
exit /b

:923fa470-aa71-4b8b-b35c-36b79bf9f44b
set "_key=JMMVY-XFNQC-KK4HK-9H7R3-WQQTV" &:: Visio Standard
exit /b

:72e9faa7-ead1-4f3d-9f6e-3abc090a81d7
set "_key=82FTR-NCHR7-W3944-MGRHM-JMCWD" &:: Access
exit /b

:cbbba2c3-0ff5-4558-846a-043ef9d78559
set "_key=F4DYN-89BP2-WQTWJ-GR8YC-CKGJG" &:: Excel
exit /b

:bef3152a-8a04-40f2-a065-340c3f23516d
set "_key=D2F8D-N3Q3B-J28PV-X27HD-RJWB9" &:: Outlook
exit /b

:b63626a4-5f05-4ced-9639-31ba730a127e
set "_key=CW94N-K6GJH-9CTXY-MG2VC-FYCWP" &:: PowerPoint
exit /b

:0002290a-2091-4324-9e53-3cfe28884cde
set "_key=4NKHF-9HBQF-Q3B6C-7YV34-F64P3" &:: Skype for Business
exit /b

:d0eded01-0881-4b37-9738-190400095098
set "_key=MQ84N-7VYDM-FXV7C-6K7CC-VFW9J" &:: Word
exit /b

:: Office 2021
:fbdb3e18-a8ef-4fb3-9183-dffd60bd0984
set "_key=FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH" &:: Professional Plus
exit /b

:080a45c5-9f9f-49eb-b4b0-c3c610a5ebd3
set "_key=KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3" &:: Standard
exit /b

:76881159-155c-43e0-9db7-2d70a9a3a4ca
set "_key=FTNWT-C6WBT-8HMGF-K9PRX-QV9H8" &:: Project Professional
exit /b

:6dd72704-f752-4b71-94c7-11cec6bfc355
set "_key=J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T" &:: Project Standard
exit /b

:fb61ac9a-1688-45d2-8f6b-0674dbffa33c
set "_key=KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4" &:: Visio Professional
exit /b

:72fce797-1884-48dd-a860-b2f6a5efd3ca
set "_key=MJVNY-BYWPY-CWV6J-2RKRT-4M8QG" &:: Visio Standard
exit /b

:1fe429d8-3fa7-4a39-b6f0-03dded42fe14
set "_key=WM8YG-YNGDD-4JHDC-PG3F4-FC4T4" &:: Access
exit /b

:ea71effc-69f1-4925-9991-2f5e319bbc24
set "_key=NWG3X-87C9K-TC7YY-BC2G7-G6RVC" &:: Excel
exit /b

:a5799e4c-f83c-4c6e-9516-dfe9b696150b
set "_key=C9FM6-3N72F-HFJXB-TM3V9-T86R9" &:: Outlook
exit /b

:6e166cc3-495d-438a-89e7-d7c9e6fd4dea
set "_key=TY7XF-NFRBR-KJ44C-G83KF-GX27K" &:: PowerPoint
exit /b

:aa66521f-2370-4ad8-a2bb-c095e3e4338f
set "_key=2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ" &:: Publisher
exit /b

:1f32a9af-1274-48bd-ba1e-1ab7508a23e8
set "_key=HWCXN-K3WBT-WJBKY-R8BD9-XK29P" &:: Skype for Business
exit /b

:abe28aea-625a-43b1-8e30-225eb8fbd9e5
set "_key=TN8H9-M34D3-Y64V9-TR72V-X79KV" &:: Word
exit /b

:: Office 2019
:85dd8b5f-eaa4-4af3-a628-cce9e77c9a03
set "_key=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP" &:: Professional Plus
exit /b

:6912a74b-a5fb-401a-bfdb-2e3ab46f4b02
set "_key=6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK" &:: Standard
exit /b

:2ca2bf3f-949e-446a-82c7-e25a15ec78c4
set "_key=B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B" &:: Project Professional
exit /b

:1777f0e3-7392-4198-97ea-8ae4de6f6381
set "_key=C4F7P-NCP8C-6CQPT-MQHV9-JXD2M" &:: Project Standard
exit /b

:5b5cf08f-b81a-431d-b080-3450d8620565
set "_key=9BGNQ-K37YR-RQHF2-38RQ3-7VCBB" &:: Visio Professional
exit /b

:e06d7df3-aad0-419d-8dfb-0ac37e2bdf39
set "_key=7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2" &:: Visio Standard
exit /b

:9e9bceeb-e736-4f26-88de-763f87dcc485
set "_key=9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT" &:: Access
exit /b

:237854e9-79fc-4497-a0c1-a70969691c6b
set "_key=TMJWT-YYNMB-3BKTF-644FC-RVXBD" &:: Excel
exit /b

:c8f8a301-19f5-4132-96ce-2de9d4adbd33
set "_key=7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK" &:: Outlook
exit /b

:3131fd61-5e4f-4308-8d6d-62be1987c92c
set "_key=RRNCX-C64HY-W2MM7-MCH9G-TJHMQ" &:: PowerPoint
exit /b

:9d3e4cca-e172-46f1-a2f4-1d2107051444
set "_key=G2KWX-3NW6P-PY93R-JXK2T-C9Y9V" &:: Publisher
exit /b

:734c6c6e-b0ba-4298-a891-671772b2bd1b
set "_key=NCJ33-JHBBY-HTK98-MYCV8-HMKHJ" &:: Skype for Business
exit /b

:059834fe-a8ea-4bff-b67b-4d006b5447d3
set "_key=PBX3G-NWMT6-Q7XBW-PYJGG-WXD33" &:: Word
exit /b

:: Office 2016
:829b8110-0e6f-4349-bca4-42803577788d
set "_key=WGT24-HCNMF-FQ7XH-6M8K7-DRTW9" &:: Project Professional C2R-P
exit /b

:cbbaca45-556a-4416-ad03-bda598eaa7c8
set "_key=D8NRQ-JTYM3-7J2DX-646CT-6836M" &:: Project Standard C2R-P
exit /b

:b234abe3-0857-4f9c-b05a-4dc314f85557
set "_key=69WXN-MBYV6-22PQG-3WGHK-RM6XC" &:: Visio Professional C2R-P
exit /b

:361fe620-64f4-41b5-ba77-84f8e079b1f7
set "_key=NY48V-PPYYH-3F4PX-XJRKJ-W4423" &:: Visio Standard C2R-P
exit /b

:e914ea6e-a5fa-4439-a394-a9bb3293ca09
set "_key=DMTCJ-KNRKX-26982-JYCKT-P7KB6" &:: MondoR
exit /b

:9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce
set "_key=HFTND-W9MK4-8B7MJ-B6C4G-XQBR2" &:: Mondo
exit /b

:d450596f-894d-49e0-966a-fd39ed4c4c64
set "_key=XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99" &:: Professional Plus
exit /b

:dedfa23d-6ed1-45a6-85dc-63cae0546de6
set "_key=JNRGM-WHDWX-FJJG3-K47QV-DRTFM" &:: Standard
exit /b

:4f414197-0fc2-4c01-b68a-86cbb9ac254c
set "_key=YG9NW-3K39V-2T3HJ-93F3Q-G83KT" &:: Project Professional
exit /b

:da7ddabc-3fbe-4447-9e01-6ab7440b4cd4
set "_key=GNFHQ-F6YQM-KQDGJ-327XX-KQBVC" &:: Project Standard
exit /b

:6bf301c1-b94a-43e9-ba31-d494598c47fb
set "_key=PD3PC-RHNGV-FXJ29-8JK7D-RJRJK" &:: Visio Professional
exit /b

:aa2a7821-1827-4c2c-8f1d-4513a34dda97
set "_key=7WHWN-4T7MP-G96JF-G33KR-W8GF4" &:: Visio Standard
exit /b

:67c0fc0c-deba-401b-bf8b-9c8ad8395804
set "_key=GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW" &:: Access
exit /b

:c3e65d36-141f-4d2f-a303-a842ee756a29
set "_key=9C2PK-NWTVB-JMPW8-BFT28-7FTBF" &:: Excel
exit /b

:d8cace59-33d2-4ac7-9b1b-9b72339c51c8
set "_key=DR92N-9HTF2-97XKM-XW2WJ-XW3J6" &:: OneNote
exit /b

:ec9d9265-9d1e-4ed0-838a-cdc20f2551a1
set "_key=R69KK-NTPKF-7M3Q4-QYBHW-6MT9B" &:: Outlook
exit /b

:d70b1bba-b893-4544-96e2-b7a318091c33
set "_key=J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6" &:: Powerpoint
exit /b

:041a06cb-c5b8-4772-809f-416d03d16654
set "_key=F47MM-N3XJP-TQXJ9-BP99D-8K837" &:: Publisher
exit /b

:83e04ee1-fa8d-436d-8994-d31a862cab77
set "_key=869NQ-FJ69K-466HW-QYCP2-DDBV6" &:: Skype for Business
exit /b

:bb11badf-d8aa-470e-9311-20eaf80fe5cc
set "_key=WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6" &:: Word
exit /b

:: Office 2013
:1dc00701-03af-4680-b2af-007ffc758a1f
set "_key=CWH2Y-NPYJW-3C7HD-BJQWB-G28JJ" &:: MondoR
exit /b

:dc981c6b-fc8e-420f-aa43-f8f33e5c0923
set "_key=42QTK-RN8M7-J3C4G-BBGYM-88CYV" &:: Mondo
exit /b

:b322da9c-a2e2-4058-9e4e-f59a6970bd69
set "_key=YC7DK-G2NP3-2QQC3-J6H88-GVGXT" &:: Professional Plus
exit /b

:b13afb38-cd79-4ae5-9f7f-eed058d750ca
set "_key=KBKQT-2NMXY-JJWGP-M62JB-92CD4" &:: Standard
exit /b

:4a5d124a-e620-44ba-b6ff-658961b33b9a
set "_key=FN8TT-7WMH6-2D4X9-M337T-2342K" &:: Project Professional
exit /b

:427a28d1-d17c-4abf-b717-32c780ba6f07
set "_key=6NTH3-CW976-3G3Y2-JK3TX-8QHTT" &:: Project Standard
exit /b

:e13ac10e-75d0-4aff-a0cd-764982cf541c
set "_key=C2FG9-N6J68-H8BTJ-BW3QX-RM3B3" &:: Visio Professional
exit /b

:ac4efaf0-f81f-4f61-bdf7-ea32b02ab117
set "_key=J484Y-4NKBF-W2HMG-DBMJC-PGWR7" &:: Visio Standard
exit /b

:6ee7622c-18d8-4005-9fb7-92db644a279b
set "_key=NG2JY-H4JBT-HQXYP-78QH9-4JM2D" &:: Access
exit /b

:f7461d52-7c2b-43b2-8744-ea958e0bd09a
set "_key=VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB" &:: Excel
exit /b

:fb4875ec-0c6b-450f-b82b-ab57d8d1677f
set "_key=H7R7V-WPNXQ-WCYYC-76BGV-VT7GH" &:: Groove
exit /b

:a30b8040-d68a-423f-b0b5-9ce292ea5a8f
set "_key=DKT8B-N7VXH-D963P-Q4PHY-F8894" &:: InfoPath
exit /b

:1b9f11e3-c85c-4e1b-bb29-879ad2c909e3
set "_key=2MG3G-3BNTT-3MFW9-KDQW3-TCK7R" &:: Lync
exit /b

:efe1f3e6-aea2-4144-a208-32aa872b6545
set "_key=TGN6P-8MMBC-37P2F-XHXXK-P34VW" &:: OneNote
exit /b

:771c3afa-50c5-443f-b151-ff2546d863a0
set "_key=QPN8Q-BJBTJ-334K3-93TGY-2PMBT" &:: Outlook
exit /b

:8c762649-97d1-4953-ad27-b7e2c25b972e
set "_key=4NT99-8RJFH-Q2VDH-KYG2C-4RD4F" &:: Powerpoint
exit /b

:00c79ff1-6850-443d-bf61-71cde0de305f
set "_key=PN2WF-29XG2-T9HJ7-JQPJR-FCXK4" &:: Publisher
exit /b

:d9f5b1c6-5386-495a-88f9-9ad6b41ac9b3
set "_key=6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7" &:: Word
exit /b

:embdbin:
function UninstallLicenses($DllPath) {
    $TB = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1).DefineDynamicModule(2).DefineType(0)
    
    [void]$TB.DefinePInvokeMethod('SLOpen', $DllPath, 22, 1, [int], @([IntPtr].MakeByRefType()), 1, 3)
    [void]$TB.DefinePInvokeMethod('SLGetSLIDList', $DllPath, 22, 1, [int],
        @([IntPtr], [int], [Guid].MakeByRefType(), [int], [int].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
    [void]$TB.DefinePInvokeMethod('SLUninstallLicense', $DllPath, 22, 1, [int], @([IntPtr], [IntPtr]), 1, 3)

    $SPPC = $TB.CreateType()
    $Handle = 0
    [void]$SPPC::SLOpen([ref]$Handle)
    $pnReturnIds = 0
    $ppReturnIds = 0

    if (!$SPPC::SLGetSLIDList($Handle, 0, [ref][Guid]"0ff1ce15-a989-479d-af46-f275c6370663", 6, [ref]$pnReturnIds, [ref]$ppReturnIds)) {
        foreach ($i in 0..($pnReturnIds - 1)) {
            [void]$SPPC::SLUninstallLicense($Handle, [Int64]$ppReturnIds + [Int64]16 * $i)
        }    
    }
}

$OSPP = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform" -ErrorAction SilentlyContinue).Path
if ($OSPP) {
    UninstallLicenses ($OSPP + "osppc.dll")
}
UninstallLicenses "sppc.dll"
:embdbin:

:embdxrm:
$wmi = Get-WmiObject $sls
function InstallLicenseFile($Lsc) {
    try {
        $null = $wmi.InstallLicense([IO.File]::ReadAllText($Lsc))
    } catch {
        $host.SetShouldExit($_.Exception.HResult)
    }
}
function InstallLicenseArr($Str) {
    $a = $Str -split ';'
    ForEach ($x in $a) {InstallLicenseFile "$x"}
}
function InstallLicenseDir($Loc) {
    dir $Loc *.xrm-ms -af -s | select -expand FullName | % {InstallLicenseFile "$_"}
}
function ReinstallLicenses() {
    $Oem = "$env:SystemRoot\system32\oem"
    $Spp = "$env:SystemRoot\system32\spp\tokens"
    InstallLicenseDir "$Spp"
    If (Test-Path $Oem) {InstallLicenseDir "$Oem"}
}
:embdxrm:

:E_Admin
echo %_err%
echo This script requires administrator privileges.
echo To do so, right-click on this script and select 'Run as administrator'
goto :E_Exit

:E_PWS
echo %_err%
echo Windows PowerShell is not installed.
echo It is required for this script to work.
goto :E_Exit

:E_VBS
echo %_err%
echo VBScript engine is not installed.
echo It is required for this script to work.
goto :E_Exit

:E_WSH
echo %_err%
echo Windows Script Host is disabled.
echo It is required for this script to work.
goto :E_Exit

:E_WMI
echo %_err%
echo This script require one of these to work:
echo wmic.exe tool
echo VBScript engine
echo Windows PowerShell
goto :E_Exit

:TheEnd
if exist "%_temp%\crv*.txt" del /f /q "%_temp%\crv*.txt"
echo.
echo %_ln%
echo %msg%
echo %_ln%
goto :E_Exit

:E_Exit
if %_Debug% EQU 1 goto :eof
echo.
echo Press any key to exit.
pause >nul
goto :eof

:qrPKey
if %_cwmi% EQU 1 (
set "_qr=wmic path %1 where Version='%2' call InstallProductKey ProductKey="%3""
exit /b
)
if %WMI_VBS% NEQ 0 (
set "_qr=%_csp% %1 "%3""
exit /b
)
set _qr=%_psc% "try {$null=([WMI]'%1=''%2''').InstallProductKey('%3')} catch {$host.SetShouldExit($_.Exception.HResult)}"
exit /b

:qrMethod
if %_cwmi% EQU 1 (
set "_qr=wmic path %1 where %2='%3' call %4"
exit /b
)
if %WMI_VBS% NEQ 0 (
set "_qr=%_csm% "%1.%2='%3'" %4"
exit /b
)
set _qr=%_psc% "try {$null=([WMI]'%1.%2=''%3''').%4()} catch {$host.SetShouldExit($_.Exception.HResult)}"
exit /b

:qrSingle
if %_cwmi% EQU 1 (
set "_qr=wmic path %1 get %2 /value"
exit /b
)
if %WMI_VBS% NEQ 0 (
set "_qr=%_csq% %1 %2"
exit /b
)
set _qr=%_psc% "(([WMISEARCHER]'SELECT %2 FROM %1').Get()).Properties | %% {$_.Name+'='+$_.Value}"
exit /b

:qrQuery
set "_quxt="
set "_quxt=%~4"
if %_cwmi% EQU 1 (
set "_qr=wmic path %1 where "%~2" get %3 /value"
if defined _quxt set "_qr=wmic path %1 where "%~2" get %3"
exit /b
)
if %WMI_VBS% NEQ 0 (
set "_qr=%_csq% %1 "%~2" %3"
exit /b
)
set "_rq=%~2"
set "_rq=%_rq:'=''%"
set _qr=%_psc% "(([WMISEARCHER]'SELECT %3 FROM %1 WHERE %_rq%').Get()).Properties | %% {$_.Name+'='+$_.Value}"
exit /b

----- Begin wsf script --->
<package>
   <job id="WmiQuery">
      <script language="VBScript">
         If WScript.Arguments.Count = 3 Then
            wExc = "Select " & WScript.Arguments.Item(2) & " from " & WScript.Arguments.Item(0) & " where " & WScript.Arguments.Item(1)
            wGet = WScript.Arguments.Item(2)
         Else
            wExc = "Select " & WScript.Arguments.Item(1) & " from " & WScript.Arguments.Item(0)
            wGet = WScript.Arguments.Item(1)
         End If
         Set objCol = GetObject("winmgmts:\\.\root\CIMV2").ExecQuery(wExc,,48)
         For Each objItm in objCol
            For each Prop in objItm.Properties_
               If LCase(Prop.Name) = LCase(wGet) Then
                  WScript.Echo Prop.Name & "=" & Prop.Value
                  Exit For
               End If
            Next
         Next
      </script>
   </job>
   <job id="WmiMethod">
      <script language="VBScript">
         On Error Resume Next
         wPath = WScript.Arguments.Item(0)
         wMethod = WScript.Arguments.Item(1)
         Set objCol = GetObject("winmgmts:\\.\root\CIMV2:" & wPath)
         objCol.ExecMethod_(wMethod)
         WScript.Quit Err.Number
      </script>
   </job>
   <job id="WmiPKey">
      <script language="VBScript">
         On Error Resume Next
         wExc = "SELECT Version FROM " & WScript.Arguments.Item(0)
         wKey = WScript.Arguments.Item(1)
         Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2").ExecQuery(wExc,,48)
         For each colService in objWMIService
            Exit For
         Next
         set objService = colService
         objService.InstallProductKey(wKey)
         WScript.Quit Err.Number
      </script>
   </job>
</package>