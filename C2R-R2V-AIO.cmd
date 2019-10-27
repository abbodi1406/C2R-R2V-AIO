@setlocal DisableDelayedExpansion
@set uivr=v11 AIO
@echo off
:: Licenses used for converting Office 365 ProPlus:
:: set _O365asO2019=0 -> use Office 2016 Mondo (if you want some Office 365 features)
:: set _O365asO2019=1 -> use Office 2019 ProPlus (only for Windows 7 and 8.1)
set _O365asO2019=0

:: set to 1 to enable debug mode
set _Debug=0

:: ##################################################################

:: set to 0 to enable debug mode without cleaning or converting
set _Cnvrt=1

set "SysPath=%SystemRoot%\System32"
if exist "%SystemRoot%\Sysnative\reg.exe" (set "SysPath=%SystemRoot%\Sysnative")
set "Path=%SysPath%;%SystemRoot%;%SysPath%\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
set "_ln============================================================="
set "_err===== ERROR ===="
reg query HKU\S-1-5-19 >nul 2>&1 || (
set "msg=ERROR: right click on the script and 'Run as administrator'"
goto :TheEnd
)

set "param=%~f0"
cmd /v:on /c echo(^^!param^^!| findstr /R "[| ` ~ ! @ %% \^ & ( ) \[ \] { } + = ; ' , |]*^"
if %errorlevel% EQU 0 (
echo.
echo %_err%
echo Disallowed special characters detected in file path name.
echo Make sure the path do not contain the following special characters,
echo ^` ^~ ^! ^@ %% ^^ ^& ^( ^) [ ] { } ^+ ^= ^; ^' ^,
echo.
echo Press any key to exit.
pause >nul
goto :eof
)

if not exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" goto :E_PS

if %_Cnvrt% neq 1 set _Debug=1
set xBit=amd64
if /i %PROCESSOR_ARCHITECTURE%==x86 (if not defined PROCESSOR_ARCHITEW6432 set xBit=x86)
set "_temp=%SystemRoot%\Temp"
set "_log=%~dpn0"
set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "_dsk=%%b"
if exist "%SystemDrive%\Users\Public\Desktop\desktop.ini" set "_dsk=%SystemDrive%\Users\Public\Desktop"
setlocal EnableDelayedExpansion

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
  copy /y nul "!_work!\#.rw" 1>nul 2>nul && (if exist "!_work!\#.rw" del /f /q "!_work!\#.rw") || (set "_log=!_dsk!\%~n0")
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
color 1F
title Office Click-to-Run Retail-to-Volume %uivr%
for /f "tokens=6 delims=[]. " %%# in ('ver') do set winbuild=%%#
set "_SLMGR=%SysPath%\slmgr.vbs"
set "_oApp=0ff1ce15-a989-479d-af46-f275c6370663"
if %_Debug% EQU 0 (
set "_cscript=cscript //Nologo //B"
) else (
set "_cscript=cscript //Nologo"
)

echo %_ln%
echo Running C2R-R2V %uivr%
echo %_ln%

if %winbuild% lss 7601 (
set "msg=Windows 7 SP1 is the minimum supported OS..."
goto :TheEnd
)
sc query ClickToRunSvc %_Nul3%
set error1=%errorlevel%
sc query OfficeSvc %_Nul3%
set error2=%errorlevel%
if %error1% equ 1060 if %error2% equ 1060 (
set "msg=Could not detect Office ClickToRun service..."
goto :TheEnd
)

set _Office19=0
set _Office16=0
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\*.xrm-ms" (
  set _Office16=1&set "_OSPPVBS=%%b\Office16\OSPP.VBS"
)
if exist "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" (
  set _Office16=1&set "_OSPPVBS=%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS"
) else if exist "%ProgramW6432%\Microsoft Office\Office16\OSPP.VBS" (
  set _Office16=1&set "_OSPPVBS=%ProgramW6432%\Microsoft Office\Office16\OSPP.VBS"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" (
  set _Office16=1&set "_OSPPVBS=%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS"
)
set _Office15=0
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses\*.xrm-ms" (
  set _Office15=1
)
if exist "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" (
  set _Office15=1&set "_OSPP15VBS=%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS"
) else if exist "%ProgramW6432%\Microsoft Office\Office15\OSPP.VBS" (
  set _Office15=1&set "_OSPP15VBS=%ProgramW6432%\Microsoft Office\Office15\OSPP.VBS"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" (
  set _Office15=1&set "_OSPP15VBS=%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS"
)
if %_Office16% equ 0 if %_Office15% equ 0 (
set "msg=No installed Office 2013/2016/2019 product detected..."
goto :TheEnd
)
if %_Office16% equ 1 call :Reg16istry
if %_Office15% equ 1 call :Reg15istry
goto :CheckC2R

:Reg16istry
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if not errorlevel 1 (set "_InstallRoot=%%b\root")
if not "%_InstallRoot%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v PackageGUID" %_Nul6%') do if not errorlevel 1 (set "_GUID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do if not errorlevel 1 (set "_ProductIds=%%b")
  set "_Config=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
  set "_PRIDs=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs"
) else (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if not errorlevel 1 (set "_InstallRoot=%%b\root")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun /v PackageGUID" %_Nul6%') do if not errorlevel 1 (set "_GUID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do if not errorlevel 1 (set "_ProductIds=%%b")
  set "_Config=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
  set "_PRIDs=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\ProductReleaseIDs"
)
set "_LicensesPath=%_InstallRoot%\Licenses16"
set "_Integrator=%_InstallRoot%\integration\integrator.exe"
for /f "skip=2 tokens=2*" %%a in ('"reg query %_PRIDs% /v ActiveConfiguration" %_Nul6%') do set "_PRIDs=%_PRIDs%\%%b"
exit /b

:Reg15istry
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do if not errorlevel 1 (set "_Install15Root=%%b\root")
if not "%_Install15Root%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v PackageGUID" %_Nul6%') do if not errorlevel 1 (set "_GU15ID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do if not errorlevel 1 (set "_Product15Ids=%%b")
  set "_Con15fig=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds"
  set "_PR15IDs=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs"
  set "_OSPP15Ready=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration"
) else (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do if not errorlevel 1 (set "_Install15Root=%%b\root")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun /v PackageGUID" %_Nul6%') do if not errorlevel 1 (set "_GU15ID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do if not errorlevel 1 (set "_Product15Ids=%%b")
  set "_Con15fig=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds"
  set "_PR15IDs=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs"
  set "_OSPP15Ready=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun\Configuration"
)
set "_OSPP15ReadT=REG_SZ"
if "%_Product15Ids%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid" %_Nul6%') do if not errorlevel 1 (set "_Product15Ids=%%b")
  set "_Con15fig=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid"
  set "_OSPP15Ready=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun"
  set "_OSPP15ReadT=REG_DWORD"
)
set "_Licenses15Path=%_Install15Root%\Licenses"
set "_Integ15rator=%_Install15Root%\integration\integrator.exe"
exit /b

:CheckC2R
if %_Office16% equ 1 (
if "%_ProductIds%"=="" (
set "msg=Could not detect Office 2016/2019 ProductIDs..."
if %_Office15% equ 0 goto :TheEnd
)
if not exist "!_LicensesPath!\*.xrm-ms" (
set "msg=Could not detect Office 2016/2019 Licenses files..."
if %_Office15% equ 0 goto :TheEnd
)
if not exist "!_Integrator!" (
set "msg=Could not detect Office 2016/2019 Licenses Integrator..."
if %_Office15% equ 0 goto :TheEnd
)
if %winbuild% lss 9200 if not exist "!_OSPPVBS!" (
set "msg=Could not detect Office 2016/2019 Licensing tool {OSPP.vbs}..."
if %_Office15% equ 0 goto :TheEnd
)
if %winbuild% geq 10240 set _O365asO2019=0
if exist "!_LicensesPath!\Word2019VL_KMS_Client_AE*.xrm-ms" (set "_tag=2019") else (set "_tag=")
)
if %_Office15% equ 1 (
if "%_Product15Ids%"=="" (
set "msg=Could not detect Office 2013 ProductIDs..."
if %_Office16% equ 0 goto :TheEnd
)
if not exist "!_Licenses15Path!\*.xrm-ms" (
set "msg=Could not detect Office 2013 Licenses files..."
if %_Office16% equ 0 goto :TheEnd
)
if %winbuild% lss 9200 if not exist "!_OSPP15VBS!" (
set "msg=Could not detect Office 2013 Licensing tool {OSPP.vbs}..."
if %_Office16% equ 0 goto :TheEnd
)
)

:PassedC2R
if %winbuild% geq 9200 (
set _spp=SoftwareLicensingProduct
set _sps=SoftwareLicensingService
set "_vbsi=%_SLMGR% /ilc "
) else (
set _spp=OfficeSoftwareProtectionProduct
set _sps=OfficeSoftwareProtectionService
set _vbsi="!_OSPP15VBS!" /inslic:
)
set "_wmi="
for /f "tokens=2 delims==" %%# in ('"wmic path %_sps% get version /value" %_Nul6%') do if not errorlevel 1 set "_wmi=%%#"
if not defined _wmi (
set "msg=Could not execute %_sps% WMI..."
goto :TheEnd
)
echo.
echo %_ln%
echo Checking Office Licenses...
echo %_ln%
wmic path %_spp% where (ApplicationID='%_oApp%' AND Description like '%%KMSCLIENT%%') get LicenseFamily %_Nul2% | findstr /I /C:"Office" %_Nul1% && (set _KMS=1) || (set _KMS=0)
wmic path %_spp% where (ApplicationID='%_oApp%' AND Description like '%%TIMEBASED%%') get LicenseFamily %_Nul2% | findstr /I /C:"Office" %_Nul1% && (set _Time=1) || (set _Time=0)
wmic path %_spp% where (ApplicationID='%_oApp%' AND Description like '%%Trial%%') get LicenseFamily %_Nul2% | findstr /I /C:"Office" %_Nul1% && (set _Time=1)
wmic path %_spp% where (ApplicationID='%_oApp%' AND Description like '%%Grace%%') get LicenseFamily %_Nul2% | findstr /I /C:"Office" %_Nul1% && (set _Grace=1) || (set _Grace=0)
wmic path %_spp% where (ApplicationID='%_oApp%') get LicenseFamily %_Nul2% | find /i "Office16MondoVL_KMS_Client" %_Nul1% && (
wmic path %_spp% where 'ApplicationID="%_oApp%" AND LicenseFamily like "Office16O365%%"' get LicenseFamily %_Nul2% | find /i "O365" %_Nul1% || (set _Grace=1)
)
wmic path %_spp% where (ApplicationID='%_oApp%') get LicenseFamily %_Nul2% | find /i "OfficeMondoVL_KMS_Client" %_Nul1% && (
wmic path %_spp% where 'ApplicationID="%_oApp%" AND LicenseFamily like "OfficeO365%%"' get LicenseFamily %_Nul2% | find /i "O365" %_Nul1% || (set _Grace=1)
)
if %_Time% equ 0 if %_Grace% equ 0 if %_KMS% equ 1 (
set "msg=No Conversion or Cleanup Required..."
goto :TheEnd
)

set _Retail=0
wmic path %_spp% where (ApplicationID='%_oApp%' AND LicenseStatus='1' AND PartialProductKey IS NOT NULL) get Description %_Nul2% |findstr /V /R "^$" >"!_temp!\crvRetail.txt"
find /i "RETAIL channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
find /i "RETAIL(MAK) channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
find /i "TIMEBASED_SUB channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
set "_copp="
if exist "%SysPath%\msvcr100.dll" (
set _copp=%_temp%
) else if exist "!_InstallRoot!\vfs\System\msvcr100.dll" (
set _copp="!_InstallRoot!\vfs\System"
) else if exist "!_Install15Root!\vfs\System\msvcr100.dll" (
set _copp="!_Install15Root!\vfs\System"
) else if exist "%SystemRoot%\SysWOW64\msvcr100.dll" (
set _copp=%_temp%
set xBit=x86
) else if exist "!_InstallRoot!\vfs\SystemX86\msvcr100.dll" (
set _copp="!_InstallRoot!\vfs\SystemX86"
set xBit=x86
) else if exist "!_Install15Root!\vfs\SystemX86\msvcr100.dll" (
set _copp="!_Install15Root!\vfs\SystemX86"
set xBit=x86
)
if %_Cnvrt% equ 1 if %_Retail% equ 0 if defined _copp (
echo.
echo %_ln%
echo Cleaning Current Office Licenses...
echo %_ln%
pushd %_copp%
%_Nul3% powershell -noprofile -exec bypass -c "$d='!cd!';$f=[io.file]::ReadAllText('%~f0') -split ':%xBit%exe\:.*';iex ($f[1]);X 1;"
if exist cleanospp.exe (
%_Nul3% cleanospp.exe -Licenses
%_Nul3% del /f /q cleanospp.exe
) else (
echo.
echo ERROR: could not extract cleanospp.exe %xBit%
)
popd
)
echo.
echo %_ln%
echo Installing Office Volume Licenses...
echo %_ln%
echo.
set _O16O365=0
if %_Retail% equ 1 wmic path %_spp% where (ApplicationID='%_oApp%' AND LicenseStatus='1' AND PartialProductKey IS NOT NULL) get LicenseFamily %_Nul2% |findstr /V /R "^$" >"!_temp!\crvRetail.txt"
wmic path %_spp% where (ApplicationID='%_oApp%') get LicenseFamily %_Nul2% |findstr /V /R "^$" >"!_temp!\crvVolume.txt" 2>&1

if %_Office16% equ 0 goto :R15V

set _O19Ids=ProPlus2019,ProjectPro2019,VisioPro2019,Standard2019,ProjectStd2019,VisioStd2019,Access2019,SkypeforBusiness2019
set _O16Ids=ProjectPro,VisioPro,Standard,ProjectStd,VisioStd,Access,SkypeforBusiness
set _A19Ids=Excel2019,Outlook2019,PowerPoint2019,Publisher2019,Word2019
set _A16Ids=Excel,Outlook,PowerPoint,Publisher,Word
set _V19Ids=%_O19Ids%,%_A19Ids%
set _V16Ids=Mondo,%_O16Ids%,%_A16Ids%,OneNote
set _R16Ids=%_V16Ids%,Professional,HomeBusiness,HomeStudent,O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud
set _RetIds=%_V19Ids%,Professional2019,HomeBusiness2019,HomeStudent2019,%_R16Ids%

echo %_ProductIds%>"!_temp!\crvProductIds.txt"
for %%a in (%_RetIds%,ProPlus) do (
set _%%a=0
)
for %%a in (%_RetIds%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProductIds.txt" %_Nul1% && set _%%a=1
)
for %%a in (%_V19Ids%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office19%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
for %%a in (%_V16Ids%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office16%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && (
  find /i "Office16ProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
reg query %_PRIDs%\ProPlusVolume.16 %_Nul3% && (
  find /i "Office16ProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
if %_Retail% equ 1 for %%a in (%_RetIds%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office16%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aR_Sub" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aR_PIN" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aE5R_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aEDUR_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aO365R_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aCO365R_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aXC2RVL_MAKC2R" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office19%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office19%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office19%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office19%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  )
)
if %_Retail% equ 1 reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && (
  find /i "Office16ProPlusR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "Office16ProPlusR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "Office16ProPlusMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "Office16ProPlusVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
)

if %_Cnvrt% neq 1 (if %_Office15% equ 1 (goto :R15V) else (set "msg=Finished"&goto :TheEnd))

if !_Mondo! equ 1 (
call :InsLic Mondo
)
if !_O365ProPlus! equ 1 (
  if !_O365asO2019! equ 1 (
  if !_Mondo! equ 0 (
  echo O365ProPlus 2016 Suite -^> ProPlus%_tag% Licenses
  echo.
  call :InsLic ProPlus%_tag%
  )
  ) else (
  echo O365ProPlus 2016 Suite ^<-^> Mondo 2016 Licenses
  echo.
  call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
  if !_Mondo! equ 0 call :InsLic Mondo
  )
)
if !_O365Business! equ 1 if !_O365ProPlus! equ 0 (
set _O365ProPlus=1
echo O365Business 2016 Suite ^<-^> Mondo 2016 Licenses
echo.
call :InsLic O365Business NCHRJ-3VPGW-X73DM-6B36K-3RQ6B
if !_Mondo! equ 0 call :InsLic Mondo
)
if !_O365SmallBusPrem! equ 1 if !_O365Business! equ 0 if !_O365ProPlus! equ 0 (
set _O365ProPlus=1
echo O365SmallBusPrem 2016 Suite ^<-^> Mondo 2016 Licenses
echo.
call :InsLic O365SmallBusPrem 3FBRX-NFP7C-6JWVK-F2YGK-H499R
if !_Mondo! equ 0 call :InsLic Mondo
)
if !_O365HomePrem! equ 1 if !_O365SmallBusPrem! equ 0 if !_O365Business! equ 0 if !_O365ProPlus! equ 0 (
set _O365ProPlus=1
echo O365HomePrem 2016 Suite ^<-^> Mondo 2016 Licenses
echo.
call :InsLic O365HomePrem 9FNY8-PWWTY-8RY4F-GJMTV-KHGM9
if !_Mondo! equ 0 call :InsLic Mondo
)
if !_O365EduCloud! equ 1 if !_O365HomePrem! equ 0 if !_O365SmallBusPrem! equ 0 if !_O365Business! equ 0 if !_O365ProPlus! equ 0 (
set _O365ProPlus=1
echo O365EduCloud 2016 Suite ^<-^> Mondo 2016 Licenses
echo.
call :InsLic O365EduCloud 8843N-BCXXD-Q84H8-R4Q37-T3CPT
if !_Mondo! equ 0 call :InsLic Mondo
)
if !_O365ProPlus! equ 1 set _O16O365=1
if !_Mondo! equ 1 if !_O365ProPlus! equ 0 (
echo Mondo 2016 Suite
echo.
call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if %_Office15% equ 1 (goto :R15V) else (goto :GVLKC2R)
)
if !_ProPlus2019! equ 1 if !_O365ProPlus! equ 0 (
echo ProPlus 2019 Suite
echo.
call :InsLic ProPlus%_tag%
)
if !_ProPlus! equ 1 if !_O365ProPlus! equ 0 if !_ProPlus2019! equ 0 (
echo ProPlus 2016 Suite -^> ProPlus%_tag% Licenses
echo.
call :InsLic ProPlus%_tag%
)
if !_Professional2019! equ 1 if !_O365ProPlus! equ 0 if !_ProPlus2019! equ 0 if !_ProPlus! equ 0 (
echo Professional 2019 Suite -^> ProPlus%_tag% Licenses
echo.
call :InsLic ProPlus%_tag%
)
if !_Professional! equ 1 if !_O365ProPlus! equ 0 if !_ProPlus2019! equ 0 if !_ProPlus! equ 0 if !_Professional2019! equ 0 (
echo Professional 2016 Suite -^> ProPlus%_tag% Licenses
echo.
call :InsLic ProPlus%_tag%
)
if !_Standard2019! equ 1 if !_O365ProPlus! equ 0 if !_ProPlus2019! equ 0 if !_ProPlus! equ 0 if !_Professional2019! equ 0 if !_Professional! equ 0 (
echo Standard 2019 Suite
echo.
call :InsLic Standard2019
)
if !_Standard! equ 1 if !_O365ProPlus! equ 0 if !_ProPlus2019! equ 0 if !_ProPlus! equ 0 if !_Professional2019! equ 0 if !_Professional! equ 0 if !_Standard2019! equ 0 (
echo Standard 2016 Suite -^> Standard%_tag% Licenses
echo.
call :InsLic Standard%_tag%
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a2019! equ 1 (
echo %%a 2019 SKU
echo.
if defined tag (call :InsLic %%a2019) else (call :InsLic %%a)
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a! equ 1 (
if !_%%a2019! equ 0 (
  echo %%a 2016 SKU -^> %%a%_tag% Licenses
  echo.
  call :InsLic %%a%_tag%
  )
)
for %%a in (HomeBusiness2019,HomeStudent2019) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus2019! equ 0 if !_ProPlus! equ 0 if !_Professional2019! equ 0 if !_Professional! equ 0 if !_Standard2019! equ 0 if !_Standard! equ 0 (
  set _Standard2019=1
  echo %%a Suite -^> Standard2019 Licenses
  echo.
  call :InsLic Standard2019
  )
)
for %%a in (HomeBusiness,HomeStudent) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus2019! equ 0 if !_ProPlus! equ 0 if !_Professional2019! equ 0 if !_Professional! equ 0 if !_Standard2019! equ 0 if !_Standard! equ 0 if !_%%a2019! equ 0 (
  set _Standard2019=1
  echo %%a 2016 Suite -^> Standard%_tag% Licenses
  echo.
  call :InsLic Standard%_tag%
  )
)
for %%a in (%_A19Ids%,OneNote) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus2019! equ 0 if !_ProPlus! equ 0 if !_Professional2019! equ 0 if !_Professional! equ 0 if !_Standard2019! equ 0 if !_Standard! equ 0 (
  echo %%a App
  echo.
  call :InsLic %%a
  )
)
for %%a in (%_A16Ids%) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus2019! equ 0 if !_ProPlus! equ 0 if !_Professional2019! equ 0 if !_Professional! equ 0 if !_Standard2019! equ 0 if !_Standard! equ 0 if !_%%a2019! equ 0 (
  echo %%a 2016 App
  echo.
  call :InsLic %%a%_tag%
  )
)
for %%a in (Access2019) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus2019! equ 0 if !_ProPlus! equ 0 if !_Professional2019! equ 0 if !_Professional! equ 0 (
  echo %%a App
  echo.
  call :InsLic %%a
  )
)
for %%a in (Access) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus2019! equ 0 if !_ProPlus! equ 0 if !_Professional2019! equ 0 if !_Professional! equ 0 if !_%%a2019! equ 0 (
  echo %%a 2016 App
  echo.
  call :InsLic %%a%_tag%
  )
)
for %%a in (SkypeforBusiness2019) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus2019! equ 0 if !_ProPlus! equ 0 (
  echo %%a App
  echo.
  call :InsLic %%a
  )
)
for %%a in (SkypeforBusiness) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus2019! equ 0 if !_ProPlus! equ 0 if !_%%a2019! equ 0 (
  echo %%a 2016 App
  echo.
  call :InsLic %%a%_tag%
  )
)
if %_Office15% equ 1 (goto :R15V) else (goto :GVLKC2R)

:R15V
if %_Cnvrt% equ 1 (
for %%# in ("!_Licenses15Path!\client-issuance-*.xrm-ms") do (
%_cscript% %_vbsi%"!_Licenses15Path!\%%~nx#"
)
%_cscript% %_vbsi%"!_Licenses15Path!\pkeyconfig-office.xrm-ms"
)

set _O15Ids=Standard,ProjectPro,VisioPro,ProjectStd,VisioStd,Access,Lync
set _A15Ids=Excel,Groove,InfoPath,OneNote,Outlook,PowerPoint,Publisher,Word
set _R15Ids=SPD,Mondo,%_O15Ids%,%_A15Ids%,Professional,HomeBusiness,HomeStudent,O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem
set _V15Ids=Mondo,%_O15Ids%,%_A15Ids%

echo %_Product15Ids%>"!_temp!\crvProduct15s.txt"
for %%a in (%_R15Ids%,ProPlus) do (
set _%%a=0
)
for %%a in (%_R15Ids%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProduct15s.txt" %_Nul1% && set _%%a=1
)
for %%a in (%_V15Ids%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProduct15s.txt" %_Nul1% && (
  find /i "Office%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
reg query %_PR15IDs%\Active\ProPlusRetail\x-none %_Nul3% && (
  find /i "OfficeProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
reg query %_PR15IDs%\Active\ProPlusVolume\x-none %_Nul3% && (
  find /i "OfficeProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
if %_Retail% equ 1 for %%a in (%_R15Ids%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProduct15s.txt" %_Nul1% && (
  find /i "Office%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aR_Sub" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aR_PIN" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aO365R_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aCO365R_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  )
)
if %_Retail% equ 1 reg query %_PR15IDs%\Active\ProPlusRetail\x-none %_Nul3% && (
  find /i "OfficeProPlusR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "OfficeProPlusR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "OfficeProPlusMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "OfficeProPlusVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
)

if %_Cnvrt% neq 1 (set "msg=Finished"&goto :TheEnd)

if !_Mondo! equ 1 (
call :Ins15Lic Mondo
)
if !_O365ProPlus! equ 1 if !_O16O365! equ 0 (
echo O365ProPlus 2013 Suite ^<-^> Mondo 2013 Licenses
echo.
call :Ins15Lic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if !_Mondo! equ 0 call :Ins15Lic Mondo
)
if !_O365SmallBusPrem! equ 1 if !_O365ProPlus! equ 0 if !_O16O365! equ 0 (
set _O365ProPlus=1
echo O365SmallBusPrem 2013 Suite ^<-^> Mondo 2013 Licenses
echo.
call :Ins15Lic O365SmallBusPrem 3FBRX-NFP7C-6JWVK-F2YGK-H499R
if !_Mondo! equ 0 call :Ins15Lic Mondo
)
if !_O365HomePrem! equ 1 if !_O365SmallBusPrem! equ 0 if !_O365ProPlus! equ 0 if !_O16O365! equ 0 (
set _O365ProPlus=1
echo O365HomePrem 2013 Suite ^<-^> Mondo 2013 Licenses
echo.
call :Ins15Lic O365HomePrem 9FNY8-PWWTY-8RY4F-GJMTV-KHGM9
if !_Mondo! equ 0 call :Ins15Lic Mondo
)
if !_O365Business! equ 1 if !_O365HomePrem! equ 0 if !_O365SmallBusPrem! equ 0 if !_O365ProPlus! equ 0 if !_O16O365! equ 0 (
set _O365ProPlus=1
echo O365Business 2013 Suite ^<-^> Mondo 2013 Licenses
echo.
call :Ins15Lic O365Business MCPBN-CPY7X-3PK9R-P6GTT-H8P8Y
if !_Mondo! equ 0 call :Ins15Lic Mondo
)
if !_Mondo! equ 1 if !_O365ProPlus! equ 0 if !_O16O365! equ 0 (
echo Mondo 2013 Suite
echo.
call :Ins15Lic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
goto :GVLKC2R
)
if !_SPD! equ 1 if !_Mondo! equ 0 if !_O365ProPlus! equ 0 (
echo SharePoint Designer 2013 App -^> Mondo 2013 Licenses
echo.
call :Ins15Lic Mondo
goto :GVLKC2R
)
if !_ProPlus! equ 1 if !_O365ProPlus! equ 0 (
echo ProPlus 2013 Suite
echo.
call :Ins15Lic ProPlus
)
if !_Professional! equ 1 if !_O365ProPlus! equ 0 if !_ProPlus! equ 0 (
echo Professional 2013 Suite -^> ProPlus 2013 Licenses
echo.
call :Ins15Lic ProPlus
)
if !_Standard! equ 1 if !_O365ProPlus! equ 0 if !_ProPlus! equ 0 if !_Professional! equ 0 (
echo Standard 2013 Suite
echo.
call :Ins15Lic Standard
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a! equ 1 (
echo %%a 2013 SKU
echo.
call :Ins15Lic %%a
)
for %%a in (HomeBusiness,HomeStudent) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus! equ 0 if !_Professional! equ 0 if !_Standard! equ 0 (
  set _Standard=1
  echo %%a 2013 Suite -^> Standard 2013 Licenses
  echo.
  call :Ins15Lic Standard
  )
)
for %%a in (%_A15Ids%) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus! equ 0 if !_Professional! equ 0 if !_Standard! equ 0 (
  echo %%a 2013 App
  echo.
  call :Ins15Lic %%a
  )
)
for %%a in (Access) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus! equ 0 if !_Professional! equ 0 (
  echo %%a 2013 App
  echo.
  call :Ins15Lic %%a
  )
)
for %%a in (Lync) do if !_%%a! equ 1 (
if !_O365ProPlus! equ 0 if !_ProPlus! equ 0 (
  echo SkypeforBusiness 2015 App
  echo.
  call :Ins15Lic %%a
  )
)
goto :GVLKC2R

:InsLic
set "_ID=%1Volume"
set "_pkey="
if not "%2"=="" (
set "_ID=%1Retail"
set "_pkey=PidKey=%2"
)
reg delete %_Config% /f /v %_ID%.OSPPReady %_Nul3%
"!_Integrator!" /I /License PRIDName=%_ID%.16 %_pkey% PackageGUID="%_GUID%" PackageRoot="!_InstallRoot!" %_Nul1%
reg add %_Config% /f /v %_ID%.OSPPReady /t REG_SZ /d 1 %_Nul1%
reg query %_Config% /v ProductReleaseIds | findstr /I "%_ID%" %_Nul1%
if %errorlevel% neq 0 (
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
for %%# in ("!_Licenses15Path!\%_patt%*.xrm-ms") do (
%_cscript% %_vbsi%"!_Licenses15Path!\%%~nx#"
)
if defined _pkey wmic path %_sps% where version='%_wmi%' call InstallProductKey ProductKey="%_pkey%" %_Nul3%
reg add %_OSPP15Ready% /f /v %_ID%.OSPPReady /t %_OSPP15ReadT% /d 1 %_Nul1%
reg query %_Con15fig% | findstr /I "%_ID%" %_Nul1%
if %errorlevel% neq 0 (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Con15fig%') do reg add %_Con15fig% /t REG_SZ /d "%%b,%_ID%" /f %_Nul1%
)
exit /b

:GVLKC2R
echo %_ln%
echo Installing Missing KMS Client Keys...
echo %_ln%
echo.
for %%# in (15,16,19) do call :C2RLoc %%#
for %%# in (15,16,19) do if !_Office%%#! equ 0 call :C2Runi %%#
for %%# in (15,16,19) do if !_Office%%#! equ 1 call :C2Rins %%#
if exist "%SysPath%\spp\store_test\2.0\tokens.dat" (
echo.
echo %_ln%
echo Refreshing Windows Insider Preview Licenses...
echo %_ln%
echo.
%_cscript% %_SLMGR% /rilc
)
set "msg=Finished"
goto :TheEnd

:C2Runi
for /f "tokens=2 delims==" %%# in ('"wmic path %_spp% where (Name like 'Office %~1%%' AND PartialProductKey IS NOT NULL) get ID /value" %_Nul6%') do (set "aID=%%#"&call :UniKey)
exit /b

:C2Rins
for /f "tokens=2 delims==" %%# in ('"wmic path %_spp% where (Description like 'Office %1, VOLUME_KMSCLIENT%%' AND PartialProductKey IS NULL) get ID /value" %_Nul6%') do (set "aID=%%#"&call :InsKey)
exit /b

:C2RLoc
set _Office%1=0
if %1 equ 19 (
if defined _ProductIds reg query %_Config% /v ProductReleaseIds %_Nul2% | findstr 2019 %_Nul1% && set _Office%1=1
exit /b
)

for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\%1.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _Office%1=1
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\%1.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set _Office%1=1

if %1 equ 16 if defined _ProductIds (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Config% /v ProductReleaseIds') do echo %%b>"!_temp!\crvO16.txt"
for %%a in (%_R16Ids%) do (
  findstr /I /C:"%%aRetail" "!_temp!\crvO16.txt" %_Nul1% && set _Office%1=1
  )
for %%a in (%_V16Ids%,ProjectProX,ProjectStdX,VisioProX,VisioStdX) do (
  findstr /I /C:"%%aVolume" "!_temp!\crvO16.txt" %_Nul1% && set _Office%1=1
  )
reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && set _Office%1=1
reg query %_PRIDs%\ProPlusVolume.16 %_Nul3% && set _Office%1=1
exit /b
)

if %1 equ 15 if defined _Product15Ids (
set _Office%1=1
exit /b
)

if exist "%ProgramFiles%\Microsoft Office\Office%1\OSPP.VBS" set _Office%1=1
if exist "%ProgramW6432%\Microsoft Office\Office%1\OSPP.VBS" set _Office%1=1
if exist "%ProgramFiles(x86)%\Microsoft Office\Office%1\OSPP.VBS" set _Office%1=1
exit /b

:UniKey
wmic path %_spp% where ID='%aID%' call UninstallProductKey %_Nul3%
exit /b

:InsKey
if /i '%aID%' equ '1dc00701-03af-4680-b2af-007ffc758a1f' exit /b
if /i '%aID%' equ 'e914ea6e-a5fa-4439-a394-a9bb3293ca09' exit /b
if /i '%aID%' equ '0bc88885-718c-491d-921f-6f214349e79c' exit /b
if /i '%aID%' equ 'fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9' exit /b
if /i '%aID%' equ '500f6619-ef93-4b75-bcb4-82819998a3ca' exit /b
set "_key="
for /f "tokens=2 delims==" %%# in ('"wmic path %_spp% where ID='%aID%' get LicenseFamily /value"') do echo %%#
call :keys %aID%
if "%_key%"=="" (echo Could not find matching key&echo.&exit /b)
wmic path %_sps% where version='%_wmi%' call InstallProductKey ProductKey="%_key%" %_Nul3%
set ERRORCODE=%ERRORLEVEL%
if %ERRORCODE% neq 0 (
cmd /c exit /b %ERRORCODE%
echo Failed: 0x!=ExitCode!
)
echo.
exit /b

:keys
if "%~1"=="" exit /b
goto :%1 %_Nul2%

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

:x86exe:
Add-Type -Language CSharp -TypeDefinition @"
 using System.IO; public class BAT85{ public static void Decode(string tmp, string s) { MemoryStream ms=new MemoryStream(); n=0;
 byte[] b85=new byte[255]; string a85="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!#$&()+,-./;=?@[]^_{|}~";
 int[] p85={52200625,614125,7225,85,1}; for(byte i=0;i<85;i++){b85[(byte)a85[i]]=i;} bool k=false;int p=0; foreach(char c in s){
 switch(c){ case'\0':case'\n':case'\r':case'\b':case'\t':case'\xA0':case' ':case':': k=false;break; default: k=true;break; }
 if(k){ n+= b85[(byte)c] * p85[p++]; if(p == 5){ ms.Write(n4b(), 0, 4); n=0; p=0; } } }         if(p>0){ for(int i=0;i<5-p;i++){
 n += 84 * p85[p+i]; } ms.Write(n4b(), 0, p-1); } File.WriteAllBytes(tmp, ms.ToArray()); ms.SetLength(0); }
 private static byte[] n4b(){ return new byte[4]{(byte)(n>>24),(byte)(n>>16),(byte)(n>>8),(byte)n}; } private static long n=0; }
"@; function X([int]$r=1){ $tmp="$r._"; [BAT85]::Decode($tmp, $f[$r+1]); expand $d\$tmp -F:* -R; del $tmp -force }

:x86exe:
::O/bZg00000vK{~c00000EC2ui000000|5a50RR9100000N(o.=0RRIP07L++00000005js0BQgLV{Bz&Zf|pNa4uzdWdQc^MT=k^07P4WfQ;wY20#S?GDBb]000jF
::;nm&MErwD{qEHrEKpp=lE4!StIjVhYGEh-9LZjuPWy,GR62Jlhej|n&_uKW6JX@WuJ4[Z_kYt3O-3H&yT7Xj+o4fURrH}b!/-?A$DH+m5C2L4!4FJpx9Mk~.J=kXZ
::U/CFjr&Rgvoq-h4u$rt9afgky.V1ZJS$m#hvh~{a73J9WY_2l~Yx~|L&G$l}k92Hp8jj}bq(3!5(?aM,uZbN;6eM6P0+iY8xIn]P|NHNMK?Pqe00319Kovt3[14wf
::Gv8x2-Z$F|L5ZeGO8hSfy&OR}5@}vCsQ=|MV#.(700|;!EiMQI=wAwEwmS^a)T42tWFa4HX|+uAA=#L]Hg{?d@eX#1;/l&/?B/Qr@c&!d/p@T?e&DgkS@J|/)Cxv!
::YiYHHOl-=mclY[+rH}t.^hQ_=9cQzpNk7h.3(0R?5@!@Jo=u@rvv(N6tc9_c#IhQ^{@nV1G!ZFl.MPTbqG00FD6It?[hBUDxvbs(5+7}fD^r4,c#+s=]Q)wGfCee8
::xsM;#1g{S&Nb_72^^34Tu0pVq&i90Z_;jS_EOYmnZA$X,49L&_$5OK)wrh5t;k#uCuJbF;P87OqxG66jIDjCB!asve+[Rp=^GRecapW1ufaTT|6tnSgL5Lq,nG17!
::_^H9xIzJik)=_xRS-xF-3g-.;L7IJ^+^b{[v(VP=pbG@9LQ_.E0W}MFiD0p+Lz[T(b0GI&!-^v{a2JC$Gu.6L2tPm(K=}fP6|y)Pb]zI+pb3FDJhB,|ZUN?3ki7H$
::2r^0s0zqI4fENbLbKnzgL)0MN0[4r-{eW}e^/F(e?/d@|NCy!2fmz@$5t{tyKR~UFe7_Lj#$F6c21$R38Bq(ZtAec!MS5PANT=8nI!)Pb$w3O?7v$I7IeBBU!kY-p
::klvN{f,gf7BB/pyCUxy05[+iLC}o&&P.i.[E9rWS;{(za0wz,e?eXya7nNwlDaxU-NjPr]F0(sqn}2DMzi^Jo,Se5.MivntdFL0NCY(!V]A2q6,V0@W24,M+Mjh=8
::F(;^(aEZ},)2FX/lAPFu_vvv99VeC6O3_Qvmpq{igQggsMAzkL8X+yGdmhuD0i9VFc5vX@9Tf!Flv.,b3SDfBbnc9t2TWOtUDE?7lcxWS-IkM-hDLnpL]T&jL+7w9
::u{nv67x?cEdRSkTF^E~mH|6$dgT_hrY0$T4XQNAAJGU6SsX2,Ug?y1^9{AHqv/CLZG;WGi@bYS-;H&3Cv)}@IuTncv@$u;|U6S22o!&Ua=alSCr47d?C&4Z|a!2/(
::U1=8$[tf~bf{EsAS.Vx&=.K7W/L.Tx==RKYYvJ1IY$~RRk.I6A@Mdj@RO/jFhnYYn[50#,^x,8mYIjotRG!)&?mBpJ;.GXOQ{B!(Pwk[(&VeGZCOtyBm[cUflKTF8
::9n4}c(4sBwa|Sh99SG+K-T|2e,G;Jzuv~FFOSdTU@+bRrTAb0KhSNwV#(c6Vh=I4CM0xt5#)q(wNi=@k^M_iVUaG|tJl,wbCPPvD+Ks}?v(3!&8Pei^x)5TNxIq3@
::)HZ.Kf!2,{tQR@CE#6|[qchRl]v!@Ism#ppc0oY5q[J3~ayWfA0Uk;k_!6lWGVo0E/#oOXC-Z2Ca!C&~2d#x98r8ikG33JLeG8(^hKQZ8(5,TI3v3;)-~-FSCK}5b
::no#nRZXg@95UA98Xm5b1LFEYqM0$Zr;;?&(M_$vx{a}n39-eA4fj$eUBGM]V3[$SI,!-7i2]knF=w?j^]wxH0X^k|lOj6?T)IhY[0qYE,h|y.HO9U_V@8ic-1a4(|
::3=0zgh_7j2kcQ4CCLz5{0!RoJ$nnPIhh1obDO/ZuDDoq}HB5xQO$xcTp&iDcSL2UtCn0rQk}~?;^b-8zlqNTYjRsA5W1;#6;WfL8lEC1nqr?hFJfkssE.Je-C1pzF
::id^)GNFx&WJe0~!FXK@e[b/c_08BagZhlgN9e]3ORBXAyI|ztg|8{r]x3H!Fi{#T3a~qi3bcE[FZaO{_q..D24;NJefnrSCnghm](=5RQHK.,hIq5!nB9K5DW]zHw
::rhwb,gvPbM3^(BTVnp$eP2dYzs0^+ZaDH{jRS)NVxAgNZX#UFBykAs|aGy;0jiPW#f}|?jCa7V0e-tY(KbVjp[{{2mF2zOFKOU=USnU8X41MI(-hD64HX5qqioJ;Y
::p}0Zi$48aWXsey}OQ@-6t45zu9DJFoPYOuNX&a,T]xT4~7],Vy@q@lFcwZDyklB^)&9b5hGx(s21oELX[@nVFebC/ZYXpEmJ}Px/LM.2aB[CxQ)fqGBV=UTPq-tLA
::HAWSn&BRYu?X0g@8c3}@P$]pFGRS&n)&Eb6@=d@{2O}u~f)(^7MV36[FX.=F6KZ|0yo=Zv8QMPEy-=xd9wz_FYZ+&mZ~]Yq.-QOZDA~?q4NYSBb!g_hCV~OPV]OOq
::A#7$5D$3Pdll782#-aBb)Z))-eSDp)p?@}m/c7ZBvw;LOs9k5MQHQQQG[L2b{k#hnO^OiR6b.A?No1mqHJC/5BhBzrnRK3GU!hiIZU!12w2pWvrV2IVY!u-{[FN.2
::CG?@Oh-o?+;W4B#;-?tQCtLYdGzD3AFUdDuD(mSc,XZR6bDQ)}h,[J.SI-AFUanfXRkQwoL9~Plr51u!NphusW?i$0c2+d.KY~]lEef1Osp|?N3gA8C.vxw9mvx|H
::ML80+sItx5Ws,x0qjJV)/QFf^g;h=,w=?]mEs$=t;[5CGo!R~8mvoh,,U|&[+qm@W38U|@8a@xJ#4{Cilfa{;,E+zQWn?BN=X@[efk|ZPZ$48]__mOk3@m.=dZSlS
::D$5#k)rj9eRfVM@BZ_~N6b#mY6y1-_k|5Sn#D.neT!zN^,l&#fnU5j/)Vd~tC6nJBxeI&oz;fvUa6V(s#Wko_x8)7mqJ5/+OA),{IPh|XoC6L2vg-VG3JTwli|S#n
::WnDl-Xh+=GEIH3N.{5JHa?nU.hmx,73bvr3_b(?i)K7N]wiSx,R^Q1O?$o!0fR{nTp|=R,;O;N-h.$P)5K^5[4}d4qYKmCvy~FAU,]8)]3+,aLeHZ6IZ-esWs=T7V
::+vl)3PmoOkEp3C1u9t=m=uAx?^q$j#FIv7+n88RPi)4&e3)w;hk9~Pn@;w^((.VUhM@UJ^Mo}Iov=9N|8wZoCvEz;(0yqu~&qE(UGjW9fiYDIM.kuwhl0i[IL08Kc
::D.|@weChO}vuCjUjt^R^4m}V1$87OHXsYOr-TXEab;H!VA=jmkLgGV]HUoPT0fd/;84Q2^ANEl0{hkQ=f9?m_S4Fw.I]OL9.3,DJcy36q)1-Uu1{Xbh@1p[}wD7Eq
::uRH60Jn)r#^;&ZE$u2hpctdBr(OKV^?[4d]oj&5xcWRFAl#6hLH,x6emhNzLtfwf3eg-Qek$#xwd}Udggkb86WfSl.bQU(GA$fAMxi7CVO7bOM4=uT$V4b7cY)t/~
::]~u$uaPVBkYA0j2.ar|c5VPph&YDf#h4LA2snwO1cQNX.@^4kM$g/&|gQQ6fMbbAg9(Yw3q;i}vqgU;Z1tI0q{1)Jjsw2i,n[2l9RH_wT;sds]B5afhs1WHOOM@hX
::L^icEHiKH-vzQdI75=PDIcPuIXYo_9r{,ADI_QIF!1$GXXG?y1{]hA]x@}o?XK;,Neo1WUdswiX2P#$y5;x](k1$0F3T3IzvoL+TlRqp5i($M]vQcG=m!C_unlHSU
::XxFE&TB)y+(Dttd12K9),/lKt!&n$yGm3#SHJZ6H9.#i6mgP2LFM0gvN@THi;z9=dWw;2MS~#X,g.FB/[_##+avx}eCDk8CwYL35u#,gCC|XlPM{V/.2!x8|MWi2A
::_uOOqD!iEms-#2Ccc,u&n3tvI0RP(d;+PCk!aX5a,s2)gPsI2c@fQl,r2NGEZ]d/2{{0hlDzRW(g=Fn1RD#ajPfFgfe|gh#Lm]Z57nccY_FC4?p!jd2tQS)P@7~oR
::$qLD2aBl=1)V?&q57mPKbI;bJ)3eTld/ZE1ef6xj8tJuDbcD_4/Vtl&Na~-rX!v8ol[Ps7cBlPNaRi]{QG25kgfN@BtgVYqkRMYpz-Aj_J)/Xn74Nd$!@8h.3f3iR
::crB1GsH4;$M#KhL1J7e|#{.x,M})g;=6+m{FVFsbj2oXo=ZxOY{c_7nfvp+#B/I^$xo+AO(scNe7oRtDVDasW-2zU&-]OU4.)NRI2vm#8@0p|yUOv[2@;WF-1;d_)
::reT51IjF(D#,5Q-gA2myeo0B+I]67zYU)_=P8IxM,a3fkPxe{36TAeln$qh_3lJeoX~,JiGvkS@=(,SW7UA7-tY=sW-iy}qgBuY)4$Fo,SXCuff&20M=X4xazH(u9
::Bei}DbHzolqGC;$8!2x/pIC,r;!zI9Lw3iLaI9(.I#nd&ILsXn9EqoP1dG_L8|n26Xt{#gXlY~Acutpek,A)X[XW^74(TIm?kn!A&2+jJ^5Ge|uINpNwj]8q8Tab;
::wS=8d9+#3BBHORt0V=g5XQyRFoO89xK1O,p^yJyPuC6M_[NU4Z$Y1!$GX2-JoZB,t.UM(V7E;_}R4Nhp1CNmp4VfzAH1IUW&bf4|6;|^fCz49$,LN&yVWdH2y0=ad
::;)SsZp=tbiMc;SPYwMlT1hA,Fq&1@aDx-U[=FX|}eVi#=73g/cBjxD$a7lbh#WV1uI!hf&K(mOXrX?(z?9OE8{T~53q;2=FNnOg_(ay5BIeVASq[-dcdk1kSEoS$c
::j9qc}b.7KAOIA~aoyyb^.4LNba2S,+ZpZ4M]te2_86N6G+4-D[QQfHJN}4H!YJ[@S/LZaw2LJxBnY6=@O}7C,4yDwvifo6CFg$|w3beFUMEg|NJkR}H.encM@[qpo
::b]+L{BjU(^W_.MNF;HIZo;AF.11sfnM;gp!_WXccf{w4wnB(~!4mA3mmc&=EdeB;o#I,n}&kP1Gr8zgO)Cn4(J{1=W$.2p(y3]ptJvRU{JiN$7VC]hX=oTAx]X/8;
::X=E8RfsVHB5=j18p~ChVW!r=vw#QGiWKDR)run1f{ttDR1OwYy+(&w8Jtb3l+(8M(#ZQR)=(8oKG?o!6AEeA5nU?bNw=bA1&cvwi5,4}rI6QG3=P4ELa;0!djfKHj
::-C~xZ]R~u.g[2u6n,y_R@u?XLX#5&=V[3x4Kb=KG|4y-kwdk}WbZkqn!ib73fCW&kMp9TWS|0oBB=qCD7]~Q1C{TDrq!j8)pJFJxQ[N1v0T7mq8TSf|B|&3ue-C#s
::]qOX.Mg[|hX+Tnt;Ut;qey7Yq&.h~E^dzuP@OgMy,+r8NbB{S@CuHqu3EMV4zA_3EXl~U}UI3eIXZ[5bLrpv&ZZ^+wMX~bQ[Y~tqFNm&#gQgsD_0{]avFMlc,mrW8
::2i)Mv${breta]FkiJ_poz,K5$hDp&-3EWuuPz1cJTPA5.@o.OvB)DpD&Um^}-)k@ZFPT-d|G8^Ei{QH)=BY-GoOr/[D,NnFI3X#x/RThiEtM;q,q@L^Za/e9mn3dT
::dc/}3tr7-om.|Z=e6kDv?+i_ki@Q3Wrp03AXJ-Cvx#aQ/h0P3an5V|hldsoPU3w1JJn9Cs&e5~Sdnk}7I(;V9MzUnNSj^ghvUT#!vL{Qggfl@^kBfQTEY1z~Q+ZH]
::Px=Fvao7KfCFZ_],KpVuVJoFB{P0ujd6|rxN#wHQ4Ka,Rg2mae9n..1?_Qe$R1-iSNKB3HJXfakQ$NdJRi8{l.WVTo^~gI-bg|ue8ffpgwSDwZz.S(ZwPw#_JK7_#
::h2aBvhEyj^p!I[NT&cC?pPe!?yrBFx1)+.,n0|&tyFWDvqE)/;Gm.qjcu7TrfUm^PLoLi.AjoY2-lFEUMuDP7NgViH+StFbGHHds[He-PCd&zwccW&[HZD!|sAk=K
::[?5aeAEo=Mf]!{$x7YTcxx4;2zx6.0MuA!H!CK1PzJAaKcMdtY-{kWhH!?R-jf&#]!QCq61sC)6Fy1{Dtjc;,$e,#uQU8f(FResBv{nnuQ585PcF{RVLwUi42#wL|
::,Z+bUk6d!UM_^)wx9~odBKD&Dp#CS${y=~Htg-_a{lyS@?bTO~48+5xJl|skjA;M9/kaPHPhnQ|_0oZ7XVt;}a.S!1=?cu[Bft8mNJkQWV&fE8Zz]+m2^EBOs.S&g
::Rl9sY&^/b)F{bp2.S,d02[d(hCi05Zc0klLTGX.8HJ]tXv9nFf&M,jl8+UY1@J!,HFODDM7Vr1C{$!8@!{AYU{J^TKu.&^X=qDGug1T5LP5fTN3+plFl6tH)H}2ch
::cCNtf)RlB.lSrxfIFTMp?!_NcyZ0I!9IaiCF{]x!k(2PkeBI,)_fR{.PG[=]!ZW+ms4qG-m)C)[TqPLq;45snC)hMc4uEVGMX9KLQ?xTRr{uKrpMQ_qt[Q0f1;rIU
::j/Fr$V66/z--5{9)CUPO#J0=xr1@,@/T=N7boUQa_f[rH9[?l,+Iyv{]@9Bt0X1pTEj7/9zN],cHLKOT,Kk=wx]^M9L#c=Dd{l#a8^H^p9JAT_tu5{/8r#,XLNb)Z
::^myvpzkW)PZHa,N,R@+XBVwvJZ_1sS88x[t$1jNLG7Pn.]-H=DNY8LkA8Bo08EM4Rm6g(zvbk-S8U@kOD[XJo?P2j{PPb?poAZD/ORG@WaXhy.wEN~i;^w_MQ7dVq
::o/y2GXKDfqWu;59|7paKEuiQZfsK_F]2m4vh$]u8(Q8Gbu}^;80$xAYL0X]F)j6qJ+!0SZ3eUMfVcUl2FHtzS5ssKtfMJHAV_euMq=r$!,-buZx~Lp~?dJDw/D9Br
::{c]U#yDGNxKTUJ@+o8@e&?9xg[Og9ukLoQto6wAc0^raL!|jFo@E)yFG-Q,EdO;o2.ZQC^uHmetY.haA{UkggrdGAV?R9a4w.A_ze6.B)tj?LL@|ZDwP!jHKJ-#hx
::eAyKct(K}~?t7|c19==ZA8H=+Vf52I+fhe-cEdpnsSWhFbrgEgIax|Ctafn689K6zW}sX&Iv4_f^rH$GFxhO~HCzNs2IS[U[DJKXtCam8@1Qza]FY;AgwLSvR)HU8
::&f51J5t7eU|2TtUyp,{ny6DoSy[ewn0+F2.ZrjgdmR{gh|I=H9iC4]5.9O5YssV_b=Ak#MESTY09xn^M.lt_VO{Z7bZ.p?Ed2pO|s-E4Y1v[]&]!zo0VcS9ww4F,o
::m]1/i{.F+b!@aI/bv_tzARO?[_K9MHLPi#6s[#,6TcFTZ8?i@gk_ZQJ-XH$lq9lvG&AXJ.=y;m~fTGEg(ZzAIZj!-K5q7HkKe01&tt_LabaU$hQ5g)d&LUscWIGe9
::PAPT&WKW+3px,ei{J&fynH@e6_d6CP&Oys[o.kA(!Wlki[{9];]dG|Db6I;w5$-?!!JcE^GQ)yq&mlCOl/(E^ddwC_w@DOW.1]|-ayE&6uA1i]t@)4$2zT]Ga;)Dg
::9J3y^Smey_,MC^#SE{xM7!Zhef-aj43h&]aiXvH5ONC/iDROSt=F7,)ySDp6S[Hw43=qEhIM&Roi;^V8503c{RGloxT2PSwNhb4]@Up-G+.GLsLLjSd+-;P6km5Jx
::_kpt}w&Y!KPFq,Ay##yKzaX[agUm|)Mig]xB$U2!l#W5zn&/_,=e5_J3~g?p#f3wYkch8-#]f}E=?,/9iUErT(b-y|7+1SF4.PC]7Y&oFGII9R)_8NckprE|6]#sW
::@!QG8xEZ/TpY@ORLVK|R)TthHv5bnKGood4Hvn^{)cePcy$jsmyUo10K$?3Ln!k5]IC?ZADn{h@ZSwZo0AF$L?jw#;+gMnrQZCh?W=CJ[M.f(&739~u=)p}}j&GVG
::IGNR7i)lv-_g4~2Bze_l5cwp@H0|D$^)GVSHB2Tf?t#Dy!@A[Ao!tCC4o,u)N~i{8a^8D~saKIjGaudgJYI7DXfjh/,BKt}z5DybHJUklG=W[?H/Xs47zp_+{cOC9
::BtDxII)-QxSD6p9cW-$PGKGcEyhf,r/T7-m;#(JrNK+17NdWT)X+{iL)9SvD^g+U(MwWU0q.jX/8./[=)vwU}{/_h?KtVmyoqly8CLgrnE]KML#}Tu~rY^f$5L4[|
::Lu|8=FO9W_9~.W-N!e1X)~tW5oVxK;]&Aak-;aTOwuPOY1P-6(m/d$rc006)@el!~=r6S7UDgUnC7j.VC[r.=oAc^|py}+E#p&d[Mt|{D/hkC-XkT6UXkP-b/-/3|
::cUSC]?gE@K0n=9snUx~.U]dUBuypR/!KVGx+=V5((e36cJHh=Qd+Z.;zWU?_;VxBrCBLRaE}PQm)xuL)smUzKO(00marq$E1D(VI=p?nI9?S$5+kcaMyS-H=SaANj
::J4ez}mo.]!zXDyXx1OgbI^bEb&kL9L!V=n~3GC]#DH}_q+XJF$R=7wJ5~Y)PZfUf-@f,=!q6jO/oe$!v0jRK{B+26{yO2)cRP^=p^b8_zeeZ8nv2&Cs5Sy=$Lspu)
::@!btLrRgXcPk!!WQ,HWXqyJ~Kd3e[gA!wFmR385Rob}oLH6Z}2C.kVrc58C6j]6GqHygyK8{H_C2yor,tyDWZz2@ZLU#M5+zxA=&^}=CC&TD1|8-0bQz;-=Y9IB,s
::Y=~A62)4}lcXIWeMKdo(+hfRr7v7_4!lUWL^s84/H;xSvs!rRqN$1ZZoXLBmUOC1~q91q6kHg#v)i+(J0p|pk7C@0P00bCdV{noI[B#5H83=WOfB}GFiE~IxKyib{
::2p~cr^yB--iZf;GumwN_1LDFU41&i,2b6,#-W.zK8CV4X;@9Z~88kECb8$n/2n.;}7z#o#Ac89kEof_NPyq&S93)R6gupliSq7{pFoXdWVxZ1J2m,i,2Z,x7MfCq7
::-1K-St(!fhPhKv(k}&f}liBvqj)HpZT&V_~#6Tb)1/zoE3#;TaKwYF5B]f2@?,Fl!?[4Lznuq/iA7/abu-fmL5VDBl2wV]~Aag^VLn;MqNb.;EM2NPCy]/2jencJ#
::k9;c]$oG.Ov9~cWu_#hVF,?n4Mtk.rVJ^W]W3j)ji(bKvVu~bINV-SDE~&4HBdM=Mzr/)#OBj/ilE[Ou!u{k?fxfmT?[v1]V@4nh$Yb=ee=Lt7W2j8|V|_,DD8^u9
::jGdgF9.nbcA~p$SleUw,6TVF^o8(v/Jo!HnPi(Lhr1#@aByc42&Kto1=}&gd/$$qDmdr-$MutYVM&G96LpGj^lg.HtneO8~H71?t^-,))Oco}al3kMxWUd-P6gXS;
::FAHU^GPkm[Ws[@xGPoJvnO.tZ-2C1U3~?yw/l7&e^hrA#ml0,4ndoMErCuSH?6gc{s+nvGuiml.4j-]q(]3Vl0C(UJ4P!o^GECmVc0;q&LuY~Z2fqi,2hGFpVEs7$
::5PcYZFnu)C/WHzMj0keVHzIcl_Vtf.!b5~)1n7v$[+GyLy(hg|1]=ZuobMLBXT2aC5$5#XICtVZ{SMr,]/quF@;35EdJk|Ovwwvj&OAv#^mBRA{22eJKA]w8Utuq^
::m+&$5cKOo}^-S?S5=K5uht)MM7J8}_Pl!E454l2(hYUhmM9dO.d4xD)I8r#WJ[OwJN19[cnB$n246.I!kuIi.?0-VSx+?|wD@ui4go.qu(zHj_-p+/lv$2_6!ejd}
::KZeJAjQS3LQlUY_F3Yv#pZF+VCy2[JYUDx&h!YYg|4+^tm_Y;;C?0tSYuMR=ER^ef4s0LX9[qsLnM]2ySQ~-D4wW71H{AIU.C!]Sl]G6DfCmH,6dpJ{fOsMCgW_w9
::hs-O,9~vJxA36_f9w0bG[){s6mIp8oWF9a$LGuT~913|bShd5U2hSW3aY&$gu];QrAhH)GT,S/0=pqD!Rb!y|VvPtkOH[Fu5l3M/1RWtT4UupN08A?Fn1V5JVgkiP
::iwT)&7!x!ma87hifKNash+g1ogk&$fCM..~kjQ3YgoG3lVMt/#5-6)+k@=$UNYq6O@3jMX0Du
:x86exe:

:amd64exe:
Add-Type -Language CSharp -TypeDefinition @"
 using System.IO; public class BAT85{ public static void Decode(string tmp, string s) { MemoryStream ms=new MemoryStream(); n=0;
 byte[] b85=new byte[255]; string a85="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!#$&()+,-./;=?@[]^_{|}~";
 int[] p85={52200625,614125,7225,85,1}; for(byte i=0;i<85;i++){b85[(byte)a85[i]]=i;} bool k=false;int p=0; foreach(char c in s){
 switch(c){ case'\0':case'\n':case'\r':case'\b':case'\t':case'\xA0':case' ':case':': k=false;break; default: k=true;break; }
 if(k){ n+= b85[(byte)c] * p85[p++]; if(p == 5){ ms.Write(n4b(), 0, 4); n=0; p=0; } } }         if(p>0){ for(int i=0;i<5-p;i++){
 n += 84 * p85[p+i]; } ms.Write(n4b(), 0, p-1); } File.WriteAllBytes(tmp, ms.ToArray()); ms.SetLength(0); }
 private static byte[] n4b(){ return new byte[4]{(byte)(n>>24),(byte)(n>>16),(byte)(n>>8),(byte)n}; } private static long n=0; }
"@; function X([int]$r=1){ $tmp="$r._"; [BAT85]::Decode($tmp, $f[$r+1]); expand $d\$tmp -F:* -R; del $tmp -force }

:amd64exe:
::O/bZg00000!XN-u00000EC2ui000000|5a50RR9100000N(o.=0RRIP08Rh]00000005js0BQgLV{Bz&Zf|pNa4uzdWdJ2ilOu2.08U$gfQ;wo2jFoN0Jl?HBq9KS
::008cfr?;tDkd,4b?GCKCnmgo^dROl1myYJ8Y[swTP1jY!@&qo78NeWY-&N09^};1hHr_Y9(.i&X?+U{D4J1V3T.t~y[NrC7gAuv$x6wA}pz^=v0APe[s&8KHtqj}V
::TY5IvrZCKGsl^3=tMEw[L91Gg)i,enb^]KESYEUQ6PnPL[Rm1^]OpKairi,f_k7_#d8/Zycui;2LNP+/?8^.jVvHV/3ZIgA/&Vni[4x?0ATj]|0surq15oWyy7zi^
::HB6K?5D5ivNx_HH2x8P2uomP0e@(w3P+Okihc$G8JOM=jN9Tgi2lQOs7XlGjmMdOv2e^psS;i,uh!i&mG#9}&3zwGp/f5_)@F|O,.]LA#aWZMV;b2^cYm9@+Yf-kN
::Wu3[lzAt^+/I.BTzuf=ct|_e/D!OGB$;aNt-,G.)k?f3LU/N1?L&(=eOt?t-&pU/puAE|c)k@C8lM^/[wLofdS]+ZAzPp5~0klq9+~vBWYsSrg,h#f?$X^Tv!5?2?
::fhb,_4&PYTWFJ8ANKu2(oAEFH)flX?_!t$&/Y7rYGdKP&S84@oa!3=&vr^|phoNH4jJEU{I$3PmiL,MfP3_;8Ff7W)-J32J-n^&{)PmjMw4,i/o$K4a^~7~nS{bN$
::M})t9wB41(W$F@zS8$A3$+ML}kOrh+2(RxSbgA78ec|B=mYMh!X9CN5@A7Rn@1i73R#/FkYxogzoNL6730C,U$O;J1;+ifuaHKli/lk_p194KhQ-($VF~]1VFjfmt
::[k9]?6IsaS+ME5yT7kWzYjNEpSi}Af^xYvU1bW$9/[5.;etIM/TIKXRQKYv_lY9awS,LO(=QBUT?R&ZtBGX1bjp7dcH1~Ko-DQIxqjy4/(v.1Ve437#G+0e1H&{.P
::{+fqhnGR0qQuMl}VA}w?|D?!m;DJ(c{,t,tF#B/Q{;&9RKo$.6l3GDt)qTt$a6SvE4)a;2[WI-x=LPV2OlE#Q]U@J/6al{qOV+$r#zw7qQN#]GwC{&QdW+bCyf{FH
::.l2n+d;}@K8ybp7]M/!z(YZp&g;4Q&,p;5AdZ1YHoMsGz=0V3Q]aSrn,4trDV)@3$D;tr9rKh;H#s_xH&Cpxjfoyto9KFG/j!dZ/R0i^mi{7/ZJNzsVRy;Qh#tDv!
::(j=CM;y_ah1tU}p.0v[/!usOKsV@g1JyAdNkOq.IPy9]f0WV&D#YtUJ1Bg$}w4+YGaIJrOu0PFEdZ)$5yHKwrYx@0YaUz/+yIOZCHn1l8C0#J)E0R$DqE!E!xn@s7
::e4IXnLHPJEo~gEXB3$y0uBchan,nkO8$/+w4D9@r6QtW6-Mv2[#EJ&0]oYbaXz{yz&Jd&pa5xNenY}|y!!k!Y#n,wzV4]hj//YvTi_]ZpAuWp8|5=Y3q3N4$Gv#tP
::on@|8au7l]Z8BC;=J[#iM^hP=AB2OwiveWFDHkbU/~j2eS0[+|(}7!10j53pgZgjt#NAE?8y]$FX-IJ8=f4=^yb@Dq)D7BQip~tEZ4RI~o++yp;n@.&;gB0kxtBzg
::m-MhEUP#PQuK0u}SnWNAwtgGA$z5AJ=oJ1+Y4~wPp9=83gn8G]pDwzHEG{VJt6|7=hf}@SoIxQ^ipf_i/qN;_C7g&2DJ=Xar+?XlK0//sHmvQ0Qh^/;M{i@aICoDL
::Km8)&d~l2;.4kFP/0f)yLKBG-vX+pY.Sc_ql9tg2yov514yG6Ap+j61Ckp_eLzM0D=0kJ-,5]~QxU=_vfJPo^_,n._ug_Si|LYeTe4l.x8I9jUt&S8JWk.S!2]9g$
::o}Y(bs@dEU;{=lRUHz,ZEW058ca&!-0q(NpsuPenCS2H1OVvG9Vk.T~wT67Ec_BF)Kh&4t(W#U8EW-2qPqBc4@N(B&RS/+6AM6r44y65S__v2Qa#cTS!@0}zB01mG
::VMTCckW0]}6]S$VDIL+8cCES4iLPaVr{6x_vMtJ539;[W^~gZ[fOFSRuy4id^2WpxFJh!3IY!zCNy13(eRtFD)5x^4(WK=.gw~adMd8JdzasL]09Kin;$)dl)Kp0I
::9?1TXF2thKtNn=weZ/LE;!oI.hL13uH)Fooj^xEI_xqSr_upw~F@o^[8uK#PNWU7UmC63PD_oQS7KpWM_Ov~;K1$VoJUcj_^MN&0x.wMZy-{9I6!?B#UFC=bQ9rng
::[?J3wubU^g-RcQmP=O@&tfLX4g^$a$]l[]JVL?yLA@36FON&uJt5nU0?zfRI57/J^U37G$Gd#s8]rIj;e_m{lB^[p=lGu$!r8t-e??Vp$v+J{@iX,^^cRxVM7}~3a
::yLgs24M/uqVQ];sq^]3Nk{I~Y5NtGkc~XhuUwi1uqx[(_xwBv(OUimXvV;Q?2^.@taC{V?WE5~RpAwNi&zop~1r(ZubmzscFDa1O#^-EoahfJ7l{Z?Wq{jx&KJ&aQ
::uUabHg(mFBqA-BU5e+y[4^jFxpeTs.)PRPIa~5FN432lDO7qY0gQcLYpyvybh3LKkmq_aVko!;.!A=lll}|h42Mu8Rt+_e98(C2[]whw]/p~d,kAr-xvH2_&9+P7y
::;U[1DwZ/l}Gk!k@S#Ii^b#dvWlP[@|BmLeSPpOz|D.{=x.PUm_&.EPRXVimL;XVb1Co?j#Zay9Kj-iTHW&_$U38d#DqDVD^D.iMK-#M1_=5g#g@mws3yVHV_s8bi9
::as[NqZ0T@;VH&[b4rS$Os[6G09_x3~jzNDi^#/JR#y@g&js)oB)!DzrAdl4Nh~+rf78eHeV+s_EFb0,hl[}/})4YQJ6vv)S0|Pcrb@RnYUIsc;WK7rZnPCY-hU#kT
::teoab0{[Rs$(oD|7JW;vkE3M1o$;-),tB&Q)}o~71NyqxvghU.Q^m=LJPBDJ]o5_B?pE,beT8N+Dkv7VwvZ(Ewa;]|I)?cjQN]eh07rpr$TxVr/}Xe;{O4MA?[5[_
::E,tc}6n}}RYQ)ACwp/Im6r_aA!U7~LU=MByEbV,uZ^XoQ8GlN3hLu1K1nVeMAszQHUt09rQf?nlSlkfIG[z4bX|.)rv5fEwCx.uCwK]]@gFz6^{d-@b5dDm=tF6ol
::pjzd+eQiZ9UxTxj&/CcO&LOm^-ti;Ua1rEu.rZPs75,NBY0~^/vE{|hSzJ!bLAF7-AF8nAuNy4g!~MEkAX,oz7eFXJilezFUr8Radx0#-U)snZh6[F)_QK0b8C@eY
::DzkL[2bbnu;;xKYe)60vkL9hLoJR_kht[b@O_rI,Zj6MI2c7JPVdEI+f.mM8Igiyi{5.y;zY6w[y+qvikG(fQC8JReJBG7Wv7giY!uHS8)U?+!HZ(p2wd1[M]{eiQ
::e/^~|+w?fGvMO6R|0s?^Op)$CpOonNV3o]VGp9ZF8B#N)gnxxxiq?upk!trQD[P;Yj9?c2[MdYlJM=-jourz~I+1@{YTT7eUcT}Ysjj_or2QzILejM7l)~]RpC}.V
::&c)0ah!Uu?Vcl1N#U-hoPT6bTE[cy0lQjmU]Cx3,mFWjUiSfn.iCVc^Khqgi)Gl4bhpi9=nnM5H?e{CPJ.x1y=-q4D&r,{Uj)w^@bvjcUi?t#JY}V^m^@rZ&K;{;d
::+KOXvK^b,6^/l!X7c3caYJ;Ilk3sl[,wr{J)Oy})a0qr43@6$CrzB5{VzGOY^noJ(rP4z&PyAKYBl&LfI3J_~KOt5rZpLC7BQPBZR{xjSub;K039mSkd#1B1/q$Nf
::C4X4bcfzgbPh]UHS.sRk+tLouwh#PJr!6hSnxnk@Q.XEG,H8k_eW5j,8S]KuDS~Z=W~wTEf1W?_@nri#X}_AjSF)i^ndAFUSH.{$w-&{oI.=hm|LgaR8mwK=r$PI#
::eaB!r-Ov{YBF-d2/703ks0(H|55h4[eN;w[7goI[?_8e+POv[y@#@][X1[s!A+v6ap.dH0]bHPJlAagn]o{3+bdO|dz(Q=cTtq|iFM4F@+.zg#N3hItRy]NOo4)e1
::@@6jJKZ?HGXHu^KaVpu|R]t(LR9s+.d]v?$-V2NHb(/)sI+7_rAcD8e.}{uB-n!b99/FZUJwN6e@5Uf8)lKi,,hSj_hc}~Npyi$[hfT.PWdB;oTEeCVN[=RV8g#t)
::.|Vv1J4r!.g#F{0)T}-P$k[sp{nF?#8{25G75Tf/IL~|8JXd-L1;O{)^r@!uO7xxcZMMzRNDpj(3v/Och1Y{ldGa4W.01cu^ol^_]9Y4#?}gWiN2}2mpV|QQS)3HZ
::iSFOfA?xb7^w)xp6o&8=#kYe]{5GHVnNo9Im;!xy;~&lHc2hV+w~I-5;hbOqz#N^oQ13E}W.Bf[[fkroHRZFm(6lVNSHfm?bx&~db-UiJzbLYwKYeRt|KQ6/Q|VN;
::;}?}Q^@vz_,K=i1F0HrNO[1}C-[x(1BiuN9aV}XNex@_cRw}gy1@)x;o/B5VgY$n#Qp,tIfs66&L$hhw3@0NORUAaheZ##6$)a]0jUj59n[Q)aN9^r#V??hA{_Cy{
::0;,n4+NG#$2$Q.gU09LRY#2LY-/711p1ml/uv}Kyt|!6Tb2|?pM7I!s,u^o([ulpx{_jA21DqQUd4^i(C8y#N9I3xnnKyOcHI9_m]Fko?m$D&?Q&X]FLjy0zh}+n1
::JyZq!tv$Anm|-(},uAig2a9_C3Aj&mpff!]Eamr.Bis[0vo8i+zn^izw0aSOo|LWIL#Vct|Gtvb(u8Ct5|vse6jNTt0ibHI$pER2Obx,(vz?5uCN-}}t4XK|?;^CF
::uGWXD+K-j#f|Sm~]D[OkZhA$9P,lUnsnYKhIAtat1;)0kPtoM5Tt^0tF6{X~IW-mT0)Sj1lCQ^9U}I!A9&gEsrm-QXa6{9)pf{C6li=.HwYQZ}/lb]O?9[AmwmIC8
::H[f]cIIV)sE~mG8ySZ(!W~F^mZBc;F/C]{d,=zgFU1/0={wzwi0L^Q,+fsrh)ukEyeem~5=kqoH-dr]YQVLD&njD(,ql!y4]IzWSQ|KqGWM4^=8(u^MiPy-Pk)aF[
::|Bi/z8D~96U85nTpj$9G$vv;6$aYEzIYg[G{!p{A8r)+vfw@F,BOF(X@U#L[V3F1n5X(;thlT/p0!5bz&s[~@B[VbM5Ii;f,dICT$i]oCn_|HB0^gx&6ZZwArGl57
::Mi@TgEpwud+)5=Bu}THjeOf2FBR=N#IWq$F37@5#UCW-S^9xF72iF&P6Md^;Mjd+KG0c-3a)3BI{@mJMLJgvn]$1[fVM,sPHWuuNAwqlWO4vp2R)$p[a|4CkR4W.o
::tL]Qo=T0(a,K]VS4ehP(,[js4oiTZmy)8CE7m8[qQza-ab@mu3bXgDTIimE}iF@Vr09[6[Sx$[w!70UfUfM?^#qRLhtr)TCk|0n/{0~qe1xlij^yAxpB@H+XMJ!U^
::rW|lr8-A!~(HNr;6!yJ9.~Kyks2pJI4C#eW=WW@k!7@!Mpp2K7g@^guDYe+Z!Z[T_HgC]4SBcsHF4qXEix|wV&T#qG{Tom@+y[=8ltQ=v|6MR5Qh@Lil5J4)G~|dE
::RjDRd#m$56r,0q^[o@j-tHHJ#Y)+2FZIb/v2haa=JI9[1JApxo4($UMhZ{H]mE!H#?yIf+ZN?wenC{rOZaLqaBGBppvwnpTVK@|.XH]/P51vBdf.|Gq!3D5CtYM@F
::,XlPRGKYZ2)n~aehy;+i-pWBU{o~cA{7ce+So]1|.PO|juIMBCePJI_T,Rk-+h)4FPi(m82Z}nN^AskHbi}~+D!R27Ge.Hp0FIBfkHI0LQ|.o$@Hv_~cK@]m{IoI1
::b3+ZU!}4eU.XVOL63?7cF!.VF!SBP_Ucym+y^mWNv?i}2-5SgZ]6?Kk[,)qKh;AugFnu9#bJj2xbJoKI!{85b9|{k~]+y4|WzIg^JUne6.Qgm_[E8Nu2d?[2IN|uP
::_Ox1G)+_RiSjKEILS8p!#(8CP@PogX(&7T[9zs7r4BmObPp,f_(wLNim$c)eg8#HbHct2sV=.dUygyT4SAx&w{#)9vy7~SrjuEHs!0@1x0B|bO=JqX-HK5u6V.Bjo
::5Yi&98=y$hSkWMe1iH6t[^j[|;8e+5QRcUIP-(J.zWKW^o5Q1uk&Ml7aE0X&/Y1/1xtv0Iqc?&Rgwu^X=s8kWf@tm=G#mY6w6|t7FKS[pBb5tU8Yto;&]b8|UU-?&
::Ip+y{TMi8Ifg+CDC/mc~Ie@G|)^G(SIGUc);rf43RBA~~LnN.5TpN0A5{V1QEubUB^,MvvLq!DR.4aKf+L02cg+9Pb1?hOHz6;2ZB!WYNhs[(;g2s+Iy1iX6OO(aE
::OGj@0Kq&#x451z(#6u{{otZ&niA1EAVhC6o2#nCn8jAB[9t7-nAiVJ{I]/@OAu2fzk^2Qa?Pxo.@m#z|y4V)7gtOnYrzF|oX!S.KbXi!_7PToyP5dC3(hAoORu@2/
::$+-l5mk~3Hjk?_|ORbD5]ojIQln_l}fns#qKY,V(EQx3q;;1+J5S&25Zk7QQxlv1R1Yx~2II2/U{Mr$/n.MHEU=|6.7euQ~goeDANYo|khAHiKEG-Kj5Fx?lOetwE
::CaTcFG5xJ#TnUvWWkG;#1qnbMmM#+59O{ZB&MpmK96TwxXR{?n,&D;/HAf^VsF$haVrlr[rrb;bI??4|y?NISyYNsn7vaw6/OiuktK3YYB{agtejRj4ZfRxHgf[Q]
::rS)mRK/;8ix-f25flXLzw/;KgL.eUoDIui}u;83FY6y@kj1ZLnlX9$TRp@.T7)sG7Ary,7ZXOxZ+P?IzN.#lgv=-Z+J4[t07(+cu[_-Qa.De4)2_rFz0hP,;RFa?|
::r-^W#-N9}Aw_B.SRHaW7DM17mVN7W?uoM~~QTn)k!jYt5mt;0pD(c38fEYnx0RCeco;dTrsSz}@/Y#psQ#s8V2u!9&DAY&5z#7qVyC95HjcKYHiE!U26c|]KsqNNq
::2a[&tO2D^ZLTc!#ER/)K3i?L_IwmTfUka[A7EXvUO;$2B_Bh2;mm?zRzON{~,.g8|LRvy.09=Aa!|.T+!gIx2&]Ol.n]4w7q+PRqB/qJmOKpaPN/3Zy,HQMabd0Ee
::CDN-ZIAYgpGs474+^^DN4m[|kKG+_uA16GfG4$!CUca8!6@7|O7)]F]T$=-8I2D?V-(G9@{k_6?.-UX)b}Qs8+xi7[gX4B.5tOaH&fLZ&]0MD(41|yv?3P543csaA
::Sdo3ZjQxJxYd9[PC;xHt-SPJNw8p(6lNfn_R9KX/nn$8H)60fGw!O@]aMXj##!-mO19,g!gnxPvG&vO_SE0KGS/K}@TUsa[fr+{lZ$L)X,CMC-9@9A(CvujT1+)kG
::VNGzd.;jNqq)r.|z0sr?WI=00M4jea)H]v)5UR06YmMfzC/Xvnmo,thV4aC]jFCgD8kOUHf(wlXTht[WBp4eGnQXsN!Zjl@lSc&ewUaoSOxvinC8B8HwTpg4]MEWR
::=G9E|HI62&6l@m$!bq9+QwxofrKE~,[h{k[FFV{/tNI3@{_s~=BVWIU1{amo2u&Y?I3gJXAzf3+TH(j66=,z=,{C.mGWmy+7A6{]i.Mu(#JnfDi7V1,AVDXk+gsaG
::1|v6,pyYrw,PPp|?fMzd5/_b^VXf5]I3dUhQB$Uis/6$xyZpug,#$KhxttekRT^j#lj3qnx]of-+r6+i[tlx=r5vdFu8}Wpx.NgtSarVuua=O.goYGGu?mphMyjCQ
::H?btM)hy]}GvvKAyCh{-sp7W;TF;vvc3TAXx578T2gFX4@h91x3XW[F^JyKk(C|#!Qwl]9UMS6NRAs5y?tFpPDL_/eV.lErd6K$dLypX9xkflexe@mT??IuSHL4U(
::8Ek7M&-j1tIW#kXZB,+X-WW7iIkXv=E^j!eQi.dZH&J!y)}8Ek@H1b}|Ei7F#lzeXB1o,JYxVJ~njz+)7]1k.cMX?Uad@(;9G1,9R7+XhT2^v1I;}Z?6s{VIiX,6L
::jHoGX=/{zbMloQd=h.A7B{XcMispTPRTS;mk08_NsX?dVX=~/ZG+r+s;uXdJ8D.FQ]tQV&$0v!pI+-g.g|i(@U@xHsN[=cPp#HG8^-|;xHkkwNf-VY6rI[WMz4Ewe
::(TGxuJTWl#e#7kHsPL1h]-=5yb;-5/sv)8&W8=v}e9qI(^J&8DUwVae#GQBma7(6@Oh}OE[iRPo@ZpDco!!qY_sE3PuFeB9&c{Sj-S{$Be$DD!e)YvUg=j+sHV|Yo
::6pC3|wYJkXw^.0EOx|Qn+?zH{c9,i_iw~0AGk;S_c^89q5rmoFs/Z|.nbDw/{LQwp8+w]srZMKin/G)dAP[2V!M#Mjo7Xf/.Z$1QGWLI]v/5#!e=_f#z)8d/]t$##
::AeG.#M)ZpqCF;/6A2rqxdk-Fx+xrf|SZ$jhz_H.Ip4Uv)]j4H1N-G~WEGZSIV)@Y4ROks.P)U?jRuzGTSy~vZ98n5]g3lMh34RhqgeY016^K[.f,6U+b(O4r71sq;
::z/dMvxvxFEDjW]B6;hILy2]JAJiqy92V/kPaQS[yz,Lv(zKY+.pzLh@]y&VGTCh{A{KiE;=wX3Ozpg(2y{SX~+R2n_yb866b?o7FRNPv#3Zc6;Ag{WLqJg-NE1mHk
::3P^])S72+)]?4A4D~9w8JV7_h!!K~RhlvZ|31?((IEgUMyl}zJQVS.Kg3S8WSo0Gw3V0B{PiL_,Olx)qQA&S2Y5[YQ,piiHqGH9BuK?f&N2^nur?a]5Qh&({s67_L
::$tK}hGm5,eF8Cmv0X2)xFR;1,YHPoFp~ImHf@x.GIr)Wia3Gj4NH1/rL,JSOv_eoxHn@rIMz3vyzn8@tSQ&)xc-YJFHoAqB8?a|+C2/U==)JaY1X-lF^D69j.QRL}
::[D}{OyNC-!Xwpl(h/gN56dsbn=Iy}R0{N}4#trwc(eF{#=z?Ur)YD.rAJL4qt!ia|i/dIQh-fgQeY;FEYHuuUJZ6Hh)#qOkcKLDBo!pP)kCZ.52|ItN{laRQbRt;=
::1Hbk}cSfUhfZGJ+PY)Hjnescj]LK_AG+t9CrNL#QdMJp?ZP8Q9+I#|28;FJvH;tf[skBiGuN4mP4qlv2#bDd@h?]nBddFg34!8du8g3B!WYQfO9KYYcF44zy3;t~x
::_nO=yx}IDm^,[wTcv(wHDwJ/=1ez0xS)HdlfcjGYx[GYN[6-tPPHI2hgT~25V;gjr)ws@o&fF&X9HgJ~Or;KI]s=Kd=TzPWOiswPL##?L@X8tNGMPBKr)CFqGGU^A
::N?cZy9256@o@ViS$BUv2|AyNc2A1Hl$pn0?QiEqj8_I]Mo8JqW7Po64uv~VxF#tk7PD7NKj(xLFMK7DolQ^4MlgVsM+a7X}{TkU&5QV;#-9lu,+LzjTvM9$$C5j3;
::3NIJ95TJvy])ReFLANxEIp+OyaH#YR^Q-16At@5VX=W!4qOp#&uqPuLJ@BL$6.]wev4w3LKJ(hlxD-[$b)F3taJu=zlEe^pkCa2^pN|P9p@Q+O.7jB6_0zoct4JiO
::61&&&{gV4ELU!bssq!zsgwDL(jZuhzpb9j1@{jFsB1a=[tU)e5!}bKn_JQeEH0tK=4wKqfwju5=-/3?c&g)g.(knu,HqHOs2E}8!_G5[yr^-aj_M^dP)klCFpj/EW
::&#;{(THlxW&(njH#)4-!M.4.=OwOM460QRePm]ceg_(pqCt8wzIbdpv8X|IWrlNTF4@v_wLL{z4D]v@w^skk#^L.gdCrI@|&sqkcj+$#Vi-}y+WAWUC+axS2G-iSQ
::AbNvU4+Q;!54ZE=|5vQ(ceU^PhRGYscwpTi#RK#P=o{g0TZ~tP+DCGLy?l}i/V.w2QTctCZ43~@Z[{Wyu;zlFhP@Oyp-bkcaQx]!;1$m_m&3,I(;(6_gWtayS!!IE
::oDD!X!_=,m&VmsGFu@}Q9]yU}H$3r?p21xV5zaAGmBGLafo^B14SPKNW3ZVn$).9T!2{|B(?O,T.us@chTvM{6Z.D|g8SGC#NF2~xB1MOmLahByIKFj!?X&Zz[P;^
::Dwa3tYhVCuq4bqm.hm;GU!rig#8wk)BLGbx7oh?)3uYq)jiBX#fd9}(!Lrym/Q_CR,spm)m=F$}m|;uxP$$Bko&/Uxt_!c)EhruCPPo9?T$rHCmP=kh@Sgi&x=$m[
::kD&qi-3|apE.2U8z,L;O)1^hd?^MiYI18aR2@~hP96UY0o[hJ7nCN(9u@a,Qgy/mw7yPhdC5?Q^NK24B#GQF^Rq#s0.F6(ZZKqAWKmrdPd+Io;$jGc-8XhIqRoZ.^
::bGRRDAA5b7t^#hC4,OSzaS4])o!!yhkKc87#dpU$KspGOml#N^/0&QOV|c5oyQ;fyQ)Nb#@op1b.K4roy0W]nI;|F~I?BAx9bNZ$;zG6,y1Y8Qy1kv^U0;DF.Cqk(
::JG$&u,+iEw,j=[2,r~8{W&pvo#cr8hnO(RRj2$?TV7suL(aT{!vwOCy,)urO-3DNu??TX-@ELKh-Iiarh&Ck)synN[jk~pXjyoQAlkVLFyUKT&cX2zxXW[@;@=J67
::^.=RSch_64c=vo4?9hln6([8{2Hsk{2c8c-TzDpURq(PJo#Dgai]B=V3ghVT$MHJ6V!R&n9$p[vG2Rc+53di/2k#+BGF|}BRbo1Ey3T,I0AK
:amd64exe:

:E_PS
echo %_err%
echo Windows PowerShell is required for this script to work.
echo.
echo Press any key to exit.
if %_Debug% EQU 1 goto :eof
pause >nul
goto :eof

:TheEnd
del /f /q "!_temp!\crv*.txt" >nul 2>&1
echo.
echo %_ln%
echo %msg%
echo %_ln%
echo.
echo Press any key to exit...
if %_Debug% EQU 1 goto :eof
pause >nul
goto :eof
