@echo off

set Online=1
set KMS_IP=172.16.0.2
set KMS_Port=1688
If defined Renewal_Task call :Re_Activate>>"%windir%\Online_KMS_Activation_Script\logfile.txt"&exit
If defined Run_Once call :Re_Activate>>"%windir%\Online_KMS_Activation_Script\logfile.txt"&exit
:Re_Activate


set /a loop=1&set/a max_loop=1
if defined Renewal_Task set /a max_loop=3
if defined Run_Once set /a max_loop=5
:repeat


setlocal EnableExtensions EnableDelayedExpansion
set "servers="
set "servers=%servers% kms.digiboy.i"
set "servers=%servers%r"
set "servers=%servers% kms.mrxn.n"
set "servers=%servers%et"
set "servers=%servers% kms8.MSGuides.c"
set "servers=%servers%om"
set "servers=%servers% kms9.MSGuides.c"
set "servers=%servers%om"
set "servers=%servers% kms.chinancce.c"
set "servers=%servers%om"
set "servers=%servers% kms.library.h"
set "servers=%servers%k"
set "servers=%servers% kms.03k.o"
set "servers=%servers%rg"
set "servers=%servers% kms.digiboy.i"
set "servers=%servers%r"
set n=1&for %%a in (%servers%) do (set server[!n!]=%%a&set /A n+=1)&set /a max_servers=!n!-1
set server_num=1
:server
set /a activation_ok=1
if %server_num% gtr !max_servers! (
	if defined Renewal_Task (echo No KMS server available. Exiting...&exit 1651565635 rem Dummy Numbers To Show Error In Task)
	if defined Run_Once (echo No KMS server available. Exiting...&exit 1651565635 rem Dummy Numbers To Show Error In Task)
	echo No KMS server available, Press any key to Continue & pause>nul & goto:EOF)
	
set KMS_IP=!server[%server_num%]!
echo. &echo Trying with KMS server %KMS_IP% &echo.

cd /d "%~dp0"
IF /I "%PROCESSOR_ARCHITECTURE%" EQU "AMD64" (set xOS=x64) else (set xOS=Win32)
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
if %winbuild% GEQ 9600 (
reg.exe add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoGenTicket /t REG_DWORD /d 1 /f >nul 2>&1
)
if %winbuild% LSS 9200 set win7=1
if %winbuild% LSS 14393 goto :Main

SET "RegKey=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages"
SET "Pattern=Microsoft-Windows-*Edition~31bf3856ad364e35"
SET "EditionPKG=NUL"
FOR /F "TOKENS=8 DELIMS=\" %%A IN ('REG QUERY "%RegKey%" /f "%Pattern%" /k 2^>NUL ^| FIND /I "CurrentVersion"') DO (
  REG QUERY "%RegKey%\%%A" /v "CurrentState" 2>NUL | FIND /I "0x70" 1>NUL && (
    FOR /F "TOKENS=3 DELIMS=-~" %%B IN ('ECHO %%A') DO SET "EditionPKG=%%B"
  )
)
IF /I "%EditionPKG:~-7%"=="Edition" (
SET "EditionID=%EditionPKG:~0,-7%"
) ELSE (
FOR /F "TOKENS=3 DELIMS=: " %%A IN ('DISM /English /Online /Get-CurrentEdition 2^>NUL ^| FIND /I "Current Edition :"') DO SET "EditionID=%%A"
)
FOR /F "TOKENS=2 DELIMS==" %%A IN ('"WMIC PATH SoftwareLicensingProduct WHERE (Name LIKE 'Windows%%' AND PartialProductKey is not NULL) GET LicenseFamily /VALUE"') DO IF NOT ERRORLEVEL 1 SET "EditionWMI=%%A"
IF NOT DEFINED EditionWMI (
IF %winbuild% GEQ 17063 FOR /F "SKIP=2 TOKENS=3 DELIMS= " %%A IN ('REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionId') DO SET "EditionID=%%A"
GOTO :Main
)
FOR %%A IN (Cloud,CloudN) DO (IF /I "%EditionWMI%"=="%%A" GOTO :Main)
SET EditionID=%EditionWMI%

:Main
for %%A in (14,15,16) do call :officeLoc %%A
call :SPP
call :OSPP

del /f /q sppchk.txt >nul 2>&1
del /f /q osppchk.txt >nul 2>&1
echo.
if %activation_ok%==0 (
echo Activation wasn't successfull. Trying an other server.
set /a server_num+=1
goto :server
)
if defined Renewal_Task (echo Exiting...&exit)
if defined Run_Once (schtasks /delete /tn Online_KMS_Activation_Script-Run_Once /f 1>nul 2>nul &del /f /q %windir%\Online_KMS_Activation_Script\Online_KMS_Activation_Script-Run_Once.cmd >nul 2>&1 &echo Exiting...&exit)
Echo.
echo Press any key to Continue.
pause>nul
Goto:EOF

:SPP
set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService
if %loc_off15% equ 0 if %loc_off16% equ 0 (set "aword=No Installed") else (set "aword=No Supported KMS Client")
wmic path %spp% where (Description like '%%KMSCLIENT%%') get Name 2>nul | findstr /i Office 1>nul && (set office15=1) || (if not defined win7 echo.&echo %aword% Office 2013/2016/2019 Product Detected...)
wmic path %spp% where (Description like '%%KMSCLIENT%%') get Name 2>nul | findstr /i Windows 1>nul && (set WinVL=1) || (echo.&echo No Supported KMS Client Windows Detected...)
if not defined office15 if not defined WinVL exit /b
wmic path %spp% where (Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get Name 2>nul | findstr /i Windows 1>nul && (set gvlk=1) || (set gvlk=0)
for /f "tokens=2 delims==" %%A in ('"wmic path %sps% get Version /VALUE"') do set ver=%%A
wmic path %sps% where version='%ver%' call SetKeyManagementServiceMachine MachineName="%KMS_IP%" >nul 2>&1
wmic path %sps% where version='%ver%' call SetKeyManagementServicePort %KMS_Port% >nul 2>&1
for /f "tokens=2 delims==" %%G in ('"wmic path %spp% where (Description like '%%KMSCLIENT%%') get ID /VALUE"') do (set app=%%G&call :sppchk)
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceDnsPublishing 0 >nul 2>&1
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceHostCaching 0 >nul 2>&1
exit /b

:sppchk
wmic path %spp% where ID='%app%' get Name > sppchk.txt
find /i "Office" sppchk.txt 1>nul && (set office=1) || (set office=0)
find /i "Office 15" sppchk.txt 1>nul && (if %loc_off15% equ 0 exit /b)
find /i "Office 16" sppchk.txt 1>nul && (if %loc_off16% equ 0 exit /b)
find /i "Office 19" sppchk.txt 1>nul && (if %loc_off16% equ 0 exit /b)
if %office% equ 0 wmic path %spp% where ID='%app%' get LicenseStatus | findstr "1" >nul 2>&1 && (echo.&call :activate %app%&exit /b)
wmic path %spp% where (PartialProductKey is not NULL) get ID | findstr /i "%app%" >nul 2>&1 && (echo.&call :activate %app%&exit /b)
if %office% equ 1 (call :offchk15&exit /b) else (if %gvlk% equ 1 exit /b)
if %winbuild% LSS 10240 (call :winchk&exit /b)
for %%A in (
b71515d9-89a2-4c60-88c8-656fbcca7f3a,af43f7f0-3b1e-4266-a123-1fdb53f4323b,075aca1f-05d7-42e5-a3ce-e349e7be7078
11a37f09-fb7f-4002-bd84-f3ae71d11e90,43f2ab05-7c87-4d56-b27c-44d0f9a3dabd,2cf5af84-abab-4ff0-83f8-f040fb2576eb
6ae51eeb-c268-4a21-9aae-df74c38b586d,ff808201-fec6-4fd4-ae16-abbddade5706,34260150-69ac-49a3-8a0d-4a403ab55763
4dfd543d-caa6-4f69-a95f-5ddfe2b89567,5fe40dd6-cf1f-4cf2-8729-92121ac2e997,903663f7-d2ab-49c9-8942-14aa9e0a9c72
2cc171ef-db48-4adc-af09-7c574b37f139,5b2add49-b8f4-42e0-a77c-adad4efeeeb1
) do (
if /i '%app%' equ '%%A' exit /b
)
if not defined EditionID (call :winchk&exit /b)
if /i '%app%' equ '0df4f814-3f57-4b8b-9a9d-fddadcd69fac' if /i %EditionID% neq CloudE exit /b
if /i '%app%' equ 'ec868e65-fadf-4759-b23e-93fe37f2cc29' if /i %EditionID% neq ServerRdsh exit /b
if /i '%app%' equ 'e4db50ea-bda1-4566-b047-0ca50abc6f07' if /i %EditionID% neq ServerRdsh exit /b
if /i '%app%' equ 'e0c42288-980c-4788-a014-c080d2e1926e' if /i %EditionID% neq Education exit /b
if /i '%app%' equ '73111121-5638-40f6-bc11-f1d7b0d64300' if /i %EditionID% neq Enterprise exit /b
if /i '%app%' equ '2de67392-b7a7-462a-b1ca-108dd189f588' if /i %EditionID% neq Professional exit /b
if /i '%app%' equ '3f1afc82-f8ac-4f6c-8005-1d233e606eee' if /i %EditionID% neq ProfessionalEducation exit /b
if /i '%app%' equ '82bbc092-bc50-4e16-8e18-b74fc486aec3' if /i %EditionID% neq ProfessionalWorkstation exit /b
if /i '%app%' equ '3c102355-d027-42c6-ad23-2e7ef8a02585' if /i %EditionID% neq EducationN exit /b
if /i '%app%' equ 'e272e3e2-732f-4c65-a8f0-484747d0d947' if /i %EditionID% neq EnterpriseN exit /b
if /i '%app%' equ 'a80b5abf-76ad-428b-b05d-a47d2dffeebf' if /i %EditionID% neq ProfessionalN exit /b
if /i '%app%' equ '5300b18c-2e33-4dc2-8291-47ffcec746dd' if /i %EditionID% neq ProfessionalEducationN exit /b
if /i '%app%' equ '4b1571d3-bafb-4b40-8087-a961be2caf65' if /i %EditionID% neq ProfessionalWorkstationN exit /b
if /i '%app%' equ '58e97c99-f377-4ef1-81d5-4ad5522b5fd8' if /i %EditionID% neq Core exit /b
if /i '%app%' equ 'cd918a57-a41b-4c82-8dce-1a538e221a83' if /i %EditionID% neq CoreSingleLanguage exit /b
call :winchk
exit /b

:OSPP
set spp=OfficeSoftwareProtectionProduct
set sps=OfficeSoftwareProtectionService
if defined win7 (set "aword=2010/2013/2016/2019") else (set "aword=2010")
wmic path %sps% get Version >nul 2>&1 || (echo.&echo No Installed Office %aword% Product Detected...&exit /b)
wmic path %spp% where (Description like '%%KMSCLIENT%%') get Name >nul 2>&1 || (echo.&echo No Supported KMS Client Office %aword% Product Detected...&exit /b)
for /f "tokens=2 delims==" %%A in ('"wmic path %sps% get Version /VALUE" 2^>nul') do set ver=%%A
wmic path %sps% where version='%ver%' call SetKeyManagementServiceMachine MachineName="%KMS_IP%" >nul 2>&1
wmic path %sps% where version='%ver%' call SetKeyManagementServicePort %KMS_Port% >nul 2>&1
for /f "tokens=2 delims==" %%G in ('"wmic path %spp% where (Description like '%%KMSCLIENT%%') get ID /VALUE"') do (set app=%%G&call :osppchk)
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceDnsPublishing 0 >nul 2>&1
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceHostCaching 0 >nul 2>&1
exit /b

:osppchk
wmic path %spp% where ID='%app%' get Name > osppchk.txt
find /i "Office 14" osppchk.txt 1>nul && (set off14=1&if %loc_off14% equ 0 exit /b) || (set off14=0)
find /i "Office 15" osppchk.txt 1>nul && (if %loc_off15% equ 0 exit /b)
find /i "Office 16" osppchk.txt 1>nul && (if %loc_off16% equ 0 exit /b)
find /i "Office 19" osppchk.txt 1>nul && (if %loc_off16% equ 0 exit /b)
set office=0
wmic path %spp% where ID='%app%' get LicenseStatus | findstr "1" >nul 2>&1 && (echo.&call :activate %app%&exit /b)
wmic path %spp% where (PartialProductKey is not NULL) get ID | findstr /i "%app%" >nul 2>&1 && (echo.&call :activate %app%&exit /b)
if %off14% equ 1 (call :offchk14) else (call :offchk15)
exit /b

:winchk
echo.
wmic path %spp% where (LicenseStatus='1' and Description like '%%KMSCLIENT%%') get Name 2>nul | findstr /i "Windows" >nul 2>&1 && (exit /b)
wmic path %spp% where (LicenseStatus='1' and GracePeriodRemaining='0' and PartialProductKey is not NULL) get Name 2>nul | findstr /i "Windows" >nul 2>&1 && (
for /f "tokens=2 delims==" %%x in ('"wmic path %spp% where ID='%app%' get Name /VALUE"') do echo Checking: %%x
echo Product is permanently activated.
exit /b
)
call :insKey %app%
exit /b

:offchk
set ls=0
set ls2=0
for /f "tokens=2 delims==" %%A in ('"wmic path %spp% where (Name like '%%Office%~2%%') get LicenseStatus /VALUE" 2^>nul') do set /a ls=%%A
if "%~4" neq "" (
for /f "tokens=2 delims==" %%A in ('"wmic path %spp% where (Name like '%%Office%~4%%') get LicenseStatus /VALUE" 2^>nul') do set /a ls2=%%A
)
if "%ls2%" equ "1" (
echo Checking: %5
echo Product is permanently activated.
exit /b
)
if "%ls%" equ "1" (
echo Checking: %3
echo Product is permanently activated.
exit /b
)
call :insKey %app%
exit /b

:offchk15
if /i '%app%' equ '0bc88885-718c-491d-921f-6f214349e79c' exit /b
if /i '%app%' equ 'fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9' exit /b
if /i '%app%' equ '500f6619-ef93-4b75-bcb4-82819998a3ca' exit /b
if /i '%app%' equ '85dd8b5f-eaa4-4af3-a628-cce9e77c9a03' (
wmic path %spp% where 'PartialProductKey is not NULL' get ID | findstr /i "0bc88885-718c-491d-921f-6f214349e79c" 1>nul 2>nul && (exit /b)
)
if /i '%app%' equ '2ca2bf3f-949e-446a-82c7-e25a15ec78c4' (
wmic path %spp% where 'PartialProductKey is not NULL' get ID | findstr /i "fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9" 1>nul 2>nul && (exit /b)
)
if /i '%app%' equ '5b5cf08f-b81a-431d-b080-3450d8620565' (
wmic path %spp% where 'PartialProductKey is not NULL' get ID | findstr /i "500f6619-ef93-4b75-bcb4-82819998a3ca" 1>nul 2>nul && (exit /b)
)
if /i '%app%' equ '85dd8b5f-eaa4-4af3-a628-cce9e77c9a03' (
call :offchk "%app%" "19ProPlus2019VL_MAK_AE" "Office ProPlus 2019" "19ProPlus2019XC2RVL_MAKC2R" "Office ProPlus 2019 C2R"
exit /b
)
if /i '%app%' equ '6912a74b-a5fb-401a-bfdb-2e3ab46f4b02' (
call :offchk "%app%" "19Standard2019VL_MAK_AE" "Office Standard 2019"
exit /b
)
if /i '%app%' equ '2ca2bf3f-949e-446a-82c7-e25a15ec78c4' (
call :offchk "%app%" "19ProjectPro2019VL_MAK_AE" "Project Pro 2019" "19ProjectPro2019XC2RVL_MAKC2R" "Project Pro 2019 C2R"
exit /b
)
if /i '%app%' equ '1777f0e3-7392-4198-97ea-8ae4de6f6381' (
call :offchk "%app%" "19ProjectStd2019VL_MAK_AE" "Project Standard 2019"
exit /b
)
if /i '%app%' equ '5b5cf08f-b81a-431d-b080-3450d8620565' (
call :offchk "%app%" "19VisioPro2019VL_MAK_AE" "Visio Pro 2019" "19VisioPro2019XC2RVL_MAKC2R" "Visio Pro 2019 C2R"
exit /b
)
if /i '%app%' equ 'e06d7df3-aad0-419d-8dfb-0ac37e2bdf39' (
call :offchk "%app%" "19VisioStd2019VL_MAK_AE" "Visio Standard 2019"
exit /b
)
if /i '%app%' equ 'd450596f-894d-49e0-966a-fd39ed4c4c64' (
call :offchk "%app%" "16ProPlusVL_MAK" "Office ProPlus 2016"
exit /b
)
if /i '%app%' equ 'dedfa23d-6ed1-45a6-85dc-63cae0546de6' (
call :offchk "%app%" "16StandardVL_MAK" "Office Standard 2016"
exit /b
)
if /i '%app%' equ '4f414197-0fc2-4c01-b68a-86cbb9ac254c' (
call :offchk "%app%" "16ProjectProVL_MAK" "Project Pro 2016"
exit /b
)
if /i '%app%' equ 'da7ddabc-3fbe-4447-9e01-6ab7440b4cd4' (
call :offchk "%app%" "16ProjectStdVL_MAK" "Project Standard 2016"
exit /b
)
if /i '%app%' equ '6bf301c1-b94a-43e9-ba31-d494598c47fb' (
call :offchk "%app%" "16VisioProVL_MAK" "Visio Pro 2016"
exit /b
)
if /i '%app%' equ 'aa2a7821-1827-4c2c-8f1d-4513a34dda97' (
call :offchk "%app%" "16VisioStdVL_MAK" "Visio Standard 2016"
exit /b
)
if /i '%app%' equ '829b8110-0e6f-4349-bca4-42803577788d' (
call :offchk "%app%" "16ProjectProXC2RVL_MAKC2R" "Project Pro 2016 C2R"
exit /b
)
if /i '%app%' equ 'cbbaca45-556a-4416-ad03-bda598eaa7c8' (
call :offchk "%app%" "16ProjectStdXC2RVL_MAKC2R" "Project Standard 2016 C2R"
exit /b
)
if /i '%app%' equ 'b234abe3-0857-4f9c-b05a-4dc314f85557' (
call :offchk "%app%" "16VisioProXC2RVL_MAKC2R" "Visio Pro 2016 C2R"
exit /b
)
if /i '%app%' equ '361fe620-64f4-41b5-ba77-84f8e079b1f7' (
call :offchk "%app%" "16VisioStdXC2RVL_MAKC2R" "Visio Standard 2016 C2R"
exit /b
)
if /i '%app%' equ 'b322da9c-a2e2-4058-9e4e-f59a6970bd69' (
call :offchk "%app%" "ProPlusVL_MAK" "Office ProPlus 2013"
exit /b
)
if /i '%app%' equ 'b13afb38-cd79-4ae5-9f7f-eed058d750ca' (
call :offchk "%app%" "StandardVL_MAK" "Office Standard 2013"
exit /b
)
if /i '%app%' equ '4a5d124a-e620-44ba-b6ff-658961b33b9a' (
call :offchk "%app%" "ProjectProVL_MAK" "Project Pro 2013"
exit /b
)
if /i '%app%' equ '427a28d1-d17c-4abf-b717-32c780ba6f07' (
call :offchk "%app%" "ProjectStdVL_MAK" "Project Standard 2013"
exit /b
)
if /i '%app%' equ 'e13ac10e-75d0-4aff-a0cd-764982cf541c' (
call :offchk "%app%" "VisioProVL_MAK" "Visio Pro 2013"
exit /b
)
if /i '%app%' equ 'ac4efaf0-f81f-4f61-bdf7-ea32b02ab117' (
call :offchk "%app%" "VisioStdVL_MAK" "Visio Standard 2013"
exit /b
)
call :insKey %app%
exit /b

:offchk14
set "vPrem="&set "vPro="
for /f "tokens=2 delims==" %%A in ('"wmic path %spp% where (Name like '%%OfficeVisioPrem-MAK%%') get LicenseStatus /VALUE" 2^>nul') do set vPrem=%%A
for /f "tokens=2 delims==" %%A in ('"wmic path %spp% where (Name like '%%OfficeVisioPro-MAK%%') get LicenseStatus /VALUE" 2^>nul') do set vPro=%%A
if /i '%app%' equ '6f327760-8c5c-417c-9b61-836a98287e0c' (
call :offchk "%app%" "ProPlus-MAK" "Office ProPlus 2010" "ProPlusAcad-MAK" "Office Professional Academic 2010"
exit /b
)
if /i '%app%' equ '9da2a678-fb6b-4e67-ab84-60dd6a9c819a' (
call :offchk "%app%" "Standard-MAK" "Office Standard 2010"
exit /b
)
if /i '%app%' equ 'ea509e87-07a1-4a45-9edc-eba5a39f36af' (
call :offchk "%app%" "SmallBusBasics-MAK" "Office Home and Business 2010"
exit /b
)
if /i '%app%' equ 'df133ff7-bf14-4f95-afe3-7b48e7e331ef' (
call :offchk "%app%" "ProjectPro-MAK" "Project Pro 2010"
exit /b
)
if /i '%app%' equ '5dc7bf61-5ec9-4996-9ccb-df806a2d0efe' (
call :offchk "%app%" "ProjectStd-MAK" "Project Standard 2010"
exit /b
)
if /i '%app%' equ '92236105-bb67-494f-94c7-7f7a607929bd' (
call :offchk "%app%" "VisioPrem-MAK" "Visio Premium 2010" "VisioPro-MAK" "Visio Pro 2010"
exit /b
)
if defined _vPrem exit /b
if /i '%app%' equ 'e558389c-83c3-4b29-adfe-5e4d7f46c358' (
call :offchk "%app%" "VisioPro-MAK" "Visio Pro 2010" "VisioStd-MAK" "Visio Standard 2010"
exit /b
)
if defined _vPro exit /b
if /i '%app%' equ '9ed833ff-4f92-4f36-b370-8683a4f13275' (
call :offchk "%app%" "VisioStd-MAK" "Visio Standard 2010"
exit /b
)
call :insKey %app%
exit /b

:officeLoc
set loc_off%1=0
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\%1.0\Common\InstallRoot /v Path" 2^>nul') do if exist "%%b\OSPP.VBS" set loc_off%1=1
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\%1.0\Common\InstallRoot /v Path" 2^>nul') do if exist "%%b\OSPP.VBS" set loc_off%1=1
if exist "%ProgramFiles%\Microsoft Office\Office%1\OSPP.VBS" set loc_off%1=1
if exist "%ProgramFiles(x86)%\Microsoft Office\Office%1\OSPP.VBS" set loc_off%1=1
exit /b

:insKey
set "ka=echo keys.Add"
(echo edition = "%1"
echo Set keys = CreateObject ^("Scripting.Dictionary"^)
echo.
echo 'Windows 10
%ka% "58e97c99-f377-4ef1-81d5-4ad5522b5fd8", "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99" 'Home
%ka% "7b9e1751-a8da-4f75-9560-5fadfe3d8e38", "3KHY7-WNT83-DGQKR-F7HPR-844BM" 'Home N
%ka% "cd918a57-a41b-4c82-8dce-1a538e221a83", "7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH" 'Home Single Language
%ka% "a9107544-f4a0-4053-a96a-1479abdef912", "PVMJN-6DFY6-9CCP6-7BKTT-D3WVR" 'Home China
%ka% "2de67392-b7a7-462a-b1ca-108dd189f588", "W269N-WFGWX-YVC9B-4J6C9-T83GX" 'Pro
%ka% "a80b5abf-76ad-428b-b05d-a47d2dffeebf", "MH37W-N47XK-V7XM9-C7227-GCQG9" 'Pro N
%ka% "3f1afc82-f8ac-4f6c-8005-1d233e606eee", "6TP4R-GNPTD-KYYHQ-7B7DP-J447Y" 'Pro Education
%ka% "5300b18c-2e33-4dc2-8291-47ffcec746dd", "YVWGF-BXNMC-HTQYQ-CPQ99-66QFC" 'Pro Education N
%ka% "82bbc092-bc50-4e16-8e18-b74fc486aec3", "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J" 'Pro Workstation
%ka% "4b1571d3-bafb-4b40-8087-a961be2caf65", "9FNHH-K3HBT-3W4TD-6383H-6XYWF" 'Pro Workstation N
%ka% "e0c42288-980c-4788-a014-c080d2e1926e", "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2" 'Education
%ka% "3c102355-d027-42c6-ad23-2e7ef8a02585", "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ" 'Education N
%ka% "73111121-5638-40f6-bc11-f1d7b0d64300", "NPPR9-FWDCX-D2C8J-H872K-2YT43" 'Enterprise
%ka% "e272e3e2-732f-4c65-a8f0-484747d0d947", "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4" 'Enterprise N
%ka% "e0b2d383-d112-413f-8a80-97f373a5820c", "YYVX9-NTFWV-6MDM3-9PT4T-4M68B" 'Enterprise G
%ka% "e38454fb-41a4-4f59-a5dc-25080e354730", "44RPN-FTY23-9VTTB-MP9BX-T84FV" 'Enterprise G N
%ka% "7b51a46c-0c04-4e8f-9af4-8496cca90d5e", "WNMTR-4C88C-JK8YV-HQ7T2-76DF9" 'Enterprise 2015 LTSB
%ka% "87b838b7-41b6-4590-8318-5797951d8529", "2F77B-TNFGY-69QQF-B8YKP-D69TJ" 'Enterprise 2015 LTSB N
%ka% "2d5a5a60-3040-48bf-beb0-fcd770c20ce0", "DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ" 'Enterprise 2016 LTSB
%ka% "9f776d83-7156-45b2-8a5c-359b9c9f22a3", "QFFDN-GRT3P-VKWWX-X7T3R-8B639" 'Enterprise 2016 LTSB N
%ka% "32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee", "M7XTQ-FN8P6-TTKYV-9D4CC-J462D" 'Enterprise LTSC 2018
%ka% "7103a333-b8c8-49cc-93ce-d37c09687f92", "92NFX-8DJQP-P6BBQ-THF9C-7CG2H" 'Enterprise LTSC 2018 N
%ka% "e4db50ea-bda1-4566-b047-0ca50abc6f07", "7NBT4-WGBQX-MP4H7-QXFF8-YP3KX" 'Enterprise Remote Server
%ka% "ec868e65-fadf-4759-b23e-93fe37f2cc29", "CPWHC-NT2C7-VYW78-DHDB2-PG3GK" 'Enterprise Remote Sessions
%ka% "0df4f814-3f57-4b8b-9a9d-fddadcd69fac", "NBTWJ-3DR69-3C4V8-C26MC-GQ9M6" 'Lean
echo.
echo 'Windows Server 2019
%ka% "de32eafd-aaee-4662-9444-c1befb41bde2", "N69G4-B89J2-4G8F4-WWYCC-J464C" 'Standard
%ka% "34e1ae55-27f8-4950-8877-7a03be5fb181", "WMDGN-G9PQG-XVVXX-R3X43-63DFG" 'Datacenter
%ka% "034d3cbb-5d4b-4245-b3f8-f84571314078", "WVDHN-86M7X-466P6-VHXV7-YY726" 'Essentials
%ka% "a99cc1f0-7719-4306-9645-294102fbff95", "FDNH6-VW9RW-BXPJ7-4XTYG-239TB" 'Azure Core
%ka% "73e3957c-fc0c-400d-9184-5f7b6f2eb409", "N2KJX-J94YW-TQVFB-DG9YT-724CC" 'Standard ACor
%ka% "90c362e5-0da1-4bfd-b53b-b87d309ade43", "6NMRW-2C8FM-D24W7-TQWMY-CWH2D" 'Datacenter ACor
%ka% "8de8eb62-bbe0-40ac-ac17-f75595071ea3", "GRFBW-QNDC4-6QBHG-CCK3B-2PR88" 'ServerARM64
echo.
echo 'Windows Server 2016
%ka% "8c1c5410-9f39-4805-8c9d-63a07706358f", "WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY" 'Standard
%ka% "21c56779-b449-4d20-adfc-eece0e1ad74b", "CB7KF-BWN84-R7R2Y-793K2-8XDDG" 'Datacenter
%ka% "2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283", "JCKRF-N37P4-C2D82-9YXRT-4M63B" 'Essentials
%ka% "7b4433f4-b1e7-4788-895a-c45378d38253", "QN4C6-GBJD2-FB422-GHWJK-GJG2R" 'Cloud Storage
%ka% "3dbf341b-5f6c-4fa7-b936-699dce9e263f", "VP34G-4NPPG-79JTQ-864T4-R3MQX" 'Azure Core
%ka% "61c5ef22-f14f-4553-a824-c4b31e84b100", "PTXN8-JFHJM-4WC78-MPCBR-9W4KR" 'Standard ACor
%ka% "e49c08e7-da82-42f8-bde2-b570fbcae76c", "2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG" 'Datacenter ACor
%ka% "43d9af6e-5e86-4be8-a797-d072a046896c", "K9FYF-G6NCK-73M32-XMVPY-F9DRR" 'ServerARM64
echo.
echo 'Windows 8.1
%ka% "fe1c3238-432a-43a1-8e25-97e7d1ef10f3", "M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK" 'Core
%ka% "78558a64-dc19-43fe-a0d0-8075b2a370a3", "7B9N3-D94CG-YTVHR-QBPX3-RJP64" 'Core N
%ka% "c72c6a1d-f252-4e7e-bdd1-3fca342acb35", "BB6NG-PQ82V-VRDPW-8XVD2-V8P66" 'Core Single Language
%ka% "db78b74f-ef1c-4892-abfe-1e66b8231df6", "NCTT7-2RGK8-WMHRF-RY7YQ-JTXG3" 'Core China
%ka% "ffee456a-cd87-4390-8e07-16146c672fd0", "XYTND-K6QKT-K2MRH-66RTM-43JKP" 'Core ARM
%ka% "c06b6981-d7fd-4a35-b7b4-054742b7af67", "GCRJD-8NW9H-F2CDX-CCM8D-9D6T9" 'Pro
%ka% "7476d79f-8e48-49b4-ab63-4d0b813a16e4", "HMCNV-VVBFX-7HMBH-CTY9B-B4FXY" 'Pro N
%ka% "096ce63d-4fac-48a9-82a9-61ae9e800e5f", "789NJ-TQK6T-6XTH8-J39CJ-J8D3P" 'Pro with Media Center
%ka% "81671aaf-79d1-4eb1-b004-8cbbe173afea", "MHF9N-XY6XB-WVXMC-BTDCT-MKKG7" 'Enterprise
%ka% "113e705c-fa49-48a4-beea-7dd879b46b14", "TT4HM-HN7YT-62K67-RGRQJ-JFFXW" 'Enterprise N
%ka% "0ab82d54-47f4-4acb-818c-cc5bf0ecb649", "NMMPB-38DD4-R2823-62W8D-VXKJB" 'Embedded Industry Pro
%ka% "cd4e2d9f-5059-4a50-a92d-05d5bb1267c7", "FNFKF-PWTVT-9RC8H-32HB2-JB34X" 'Embedded Industry Enterprise
%ka% "f7e88590-dfc7-4c78-bccb-6f3865b99d1a", "VHXM3-NR6FT-RY6RT-CK882-KW2CJ" 'Embedded Industry Automotive
%ka% "e9942b32-2e55-4197-b0bd-5ff58cba8860", "3PY8R-QHNP9-W7XQD-G6DPH-3J2C9" 'with Bing
%ka% "c6ddecd6-2354-4c19-909b-306a3058484e", "Q6HTR-N24GM-PMJFP-69CD8-2GXKR" 'with Bing N
%ka% "b8f5e3a3-ed33-4608-81e1-37d6c9dcfd9c", "KF37N-VDV38-GRRTV-XH8X6-6F3BB" 'with Bing Single Language
%ka% "ba998212-460a-44db-bfb5-71bf09d1c68b", "R962J-37N87-9VVK2-WJ74P-XTMHR" 'with Bing China
%ka% "e58d87b5-8126-4580-80fb-861b22f79296", "MX3RK-9HNGX-K3QKC-6PJ3F-W8D7B" 'Pro for Students
%ka% "cab491c7-a918-4f60-b502-dab75e334f40", "TNFGH-2R6PB-8XM3K-QYHX2-J4296" 'Pro for Students N
echo.
echo 'Windows Server 2012 R2
%ka% "b3ca044e-a358-4d68-9883-aaa2941aca99", "D2N9P-3P6X9-2R39C-7RTCD-MDVJX" 'Standard
%ka% "00091344-1ea4-4f37-b789-01750ba6988c", "W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9" 'Datacenter
%ka% "21db6ba4-9a7b-4a14-9e29-64a60c59301d", "KNC87-3J2TX-XB4WP-VCPJV-M4FWM" 'Essentials
%ka% "b743a2be-68d4-4dd3-af32-92425b7bb623", "3NPTF-33KPT-GGBPR-YX76B-39KDD" 'Cloud Storage
echo.
echo 'Windows 8
%ka% "c04ed6bf-55c8-4b47-9f8e-5a1f31ceee60", "BN3D2-R7TKB-3YPBD-8DRP2-27GG4" 'Core
%ka% "197390a0-65f6-4a95-bdc4-55d58a3b0253", "8N2M2-HWPGY-7PGT9-HGDD8-GVGGY" 'Core N
%ka% "8860fcd4-a77b-4a20-9045-a150ff11d609", "2WN2H-YGCQR-KFX6K-CD6TF-84YXQ" 'Core Single Language
%ka% "9d5584a2-2d85-419a-982c-a00888bb9ddf", "4K36P-JN4VD-GDC6V-KDT89-DYFKP" 'Core China
%ka% "af35d7b7-5035-4b63-8972-f0b747b9f4dc", "DXHJF-N9KQX-MFPVR-GHGQK-Y7RKV" 'Core ARM
%ka% "a98bcd6d-5343-4603-8afe-5908e4611112", "NG4HW-VH26C-733KW-K6F98-J8CK4" 'Pro
%ka% "ebf245c1-29a8-4daf-9cb1-38dfc608a8c8", "XCVCF-2NXM9-723PB-MHCB7-2RYQQ" 'Pro N
%ka% "a00018a3-f20f-4632-bf7c-8daa5351c914", "GNBB8-YVD74-QJHX6-27H4K-8QHDG" 'Pro with Media Center
%ka% "458e1bec-837a-45f6-b9d5-925ed5d299de", "32JNW-9KQ84-P47T8-D8GGY-CWCK7" 'Enterprise
%ka% "e14997e7-800a-4cf7-ad10-de4b45b578db", "JMNMF-RHW7P-DMY6X-RF3DR-X2BQT" 'Enterprise N
%ka% "10018baf-ce21-4060-80bd-47fe74ed4dab", "RYXVT-BNQG7-VD29F-DBMRY-HT73M" 'Embedded Industry Pro
%ka% "18db1848-12e0-4167-b9d7-da7fcda507db", "NKB3R-R2F8T-3XCDP-7Q2KW-XWYQ2" 'Embedded Industry Enterprise
echo.
echo 'Windows Server 2012
%ka% "f0f5ec41-0d55-4732-af02-440a44a3cf0f", "XC9B7-NBPP2-83J2H-RHMBY-92BT4" 'Standard
%ka% "d3643d60-0c42-412d-a7d6-52e6635327f6", "48HP8-DN98B-MYWDG-T2DCC-8W83P" 'Datacenter
%ka% "7d5486c7-e120-4771-b7f1-7b56c6d3170c", "HM7DN-YVMH3-46JC3-XYTG7-CYQJJ" 'MultiPoint Standard
%ka% "95fd1c83-7df5-494a-be8b-1300e1c9d1cd", "XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G" 'MultiPoint Premium
echo.
echo 'Windows 7
%ka% "b92e9980-b9d5-4821-9c94-140f632f6312", "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4" 'Professional
%ka% "54a09a0d-d57b-4c10-8b69-a842d6590ad5", "MRPKT-YTG23-K7D7T-X2JMM-QY7MG" 'Professional N
%ka% "5a041529-fef8-4d07-b06f-b59b573b32d2", "W82YF-2Q76Y-63HXB-FGJG9-GF7QX" 'Professional E
%ka% "ae2ee509-1b34-41c0-acb7-6d4650168915", "33PXH-7Y6KF-2VJC9-XBBR8-HVTHH" 'Enterprise
%ka% "1cb6d605-11b3-4e14-bb30-da91c8e3983a", "YDRBP-3D83W-TY26F-D46B2-XCKRJ" 'Enterprise N
%ka% "46bbed08-9c7b-48fc-a614-95250573f4ea", "C29WB-22CC8-VJ326-GHFJW-H9DH4" 'Enterprise E
%ka% "db537896-376f-48ae-a492-53d0547773d0", "YBYF6-BHCR3-JPKRB-CDW7B-F9BK4" 'Embedded POSReady 7
%ka% "e1a8296a-db37-44d1-8cce-7bc961d59c54", "XGY72-BRBBT-FF8MH-2GG8H-W7KCW" 'Embedded Standard
%ka% "aa6dd3aa-c2b4-40e2-a544-a6bbb3f5c395", "73KQT-CD9G6-K7TQG-66MRP-CQ22C" 'Embedded ThinPC
echo.
echo 'Windows Server 2008 R2
%ka% "a78b8bd9-8017-4df5-b86a-09f756affa7c", "6TPJF-RBVHG-WBW2R-86QPH-6RTM4" 'Web
%ka% "cda18cf3-c196-46ad-b289-60c072869994", "TT8MH-CG224-D3D7Q-498W2-9QCTX" 'HPC
%ka% "68531fb9-5511-4989-97be-d11a0f55633f", "YC6KT-GKW9T-YTKYR-T4X34-R7VHC" 'Standard
%ka% "7482e61b-c589-4b7f-8ecc-46d455ac3b87", "74YFP-3QFB3-KQT8W-PMXWJ-7M648" 'Datacenter
%ka% "620e2b3d-09e7-42fd-802a-17a13652fe7a", "489J6-VHDMP-X63PK-3K798-CPX3Y" 'Enterprise
%ka% "8a26851c-1c7e-48d3-a687-fbca9b9ac16b", "GT63C-RJFQ3-4GMB6-BRFB9-CB83V" 'Itanium
%ka% "f772515c-0e87-48d5-a676-e6962c3e1195", "736RG-XDKJK-V34PF-BHK87-J6X3K" 'MultiPoint Server
echo.
echo.
echo if keys.Exists^(edition^) then
echo WScript.Echo keys.Item^(edition^)
echo End If
)>"%temp%\key.vbs"
echo.
set "key="
for /f "tokens=2 delims==" %%A in ('"wmic path %spp% where ID='%1' get Name /VALUE"') do echo Installing Key for: %%A
for /f %%A in ('cscript //Nologo "%temp%\key.vbs"') do set "key=%%A"
del /f /q "%temp%\key.vbs" >nul 2>&1
if "%key%" EQU "" (echo Could not find matching KMS Client key&exit /b)
wmic path %sps% where version='%ver%' call InstallProductKey ProductKey="%key%" >nul 2>&1

:activate
wmic path %spp% where ID='%1' call ClearKeyManagementServiceMachine >nul 2>&1
wmic path %spp% where ID='%1' call ClearKeyManagementServicePort >nul 2>&1
for /f "tokens=2 delims==" %%x in ('"wmic path %spp% where ID='%1' get Name /VALUE"') do echo Activating: %%x
wmic path %spp% where ID='%1' call Activate >nul 2>&1
set ERRORCODE=%ERRORLEVEL%
for /f "tokens=2 delims==" %%x in ('"wmic path %spp% where ID='%1' get GracePeriodRemaining /VALUE"') do (set gpr=%%x&set /a gpr2=%%x/1440)
if %gpr% equ 43200 if %office% equ 0 if not defined win7 (echo Windows Core/ProfessionalWMC Activation Successful&echo Remaining Period: 30 days ^(%gpr% minutes^)&exit /b)
if %gpr% equ 64800 (echo Windows Core/ProfessionalWMC Activation Successful&echo Remaining Period: 45 days ^(%gpr% minutes^)&exit /b)
if %gpr% gtr 259200 (echo Windows EnterpriseG/EnterpriseGN Activation Successful&echo Remaining Period: %gpr2% days ^(%gpr% minutes^)&exit /b)
if %gpr% equ 259200 (
echo Product Activation Successful
) else (
call cmd /c exit /b %ERRORCODE%
echo Product Activation Failed: 0x%=ExitCode%
set activation_ok=0
)
echo Remaining Period: %gpr2% days ^(%gpr% minutes^)
exit /b

:UnsupportedVersion

