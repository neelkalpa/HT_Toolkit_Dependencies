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
echo 'Office 2019
%ka% "0bc88885-718c-491d-921f-6f214349e79c", "VQ9DP-NVHPH-T9HJC-J9PDT-KTQRG" 'Professional Plus C2R-P
%ka% "fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9", "XM2V9-DN9HH-QB449-XDGKC-W2RMW" 'Project Professional C2R-P
%ka% "500f6619-ef93-4b75-bcb4-82819998a3ca", "N2CG9-YD3YK-936X4-3WR82-Q3X4H" 'Visio Professional C2R-P
%ka% "85dd8b5f-eaa4-4af3-a628-cce9e77c9a03", "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP" 'Professional Plus
%ka% "6912a74b-a5fb-401a-bfdb-2e3ab46f4b02", "6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK" 'Standard
%ka% "2ca2bf3f-949e-446a-82c7-e25a15ec78c4", "B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B" 'Project Professional
%ka% "1777f0e3-7392-4198-97ea-8ae4de6f6381", "C4F7P-NCP8C-6CQPT-MQHV9-JXD2M" 'Project Standard
%ka% "5b5cf08f-b81a-431d-b080-3450d8620565", "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB" 'Visio Professional
%ka% "e06d7df3-aad0-419d-8dfb-0ac37e2bdf39", "7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2" 'Visio Standard
%ka% "9e9bceeb-e736-4f26-88de-763f87dcc485", "9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT" 'Access
%ka% "237854e9-79fc-4497-a0c1-a70969691c6b", "TMJWT-YYNMB-3BKTF-644FC-RVXBD" 'Excel
%ka% "c8f8a301-19f5-4132-96ce-2de9d4adbd33", "7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK" 'Outlook
%ka% "3131fd61-5e4f-4308-8d6d-62be1987c92c", "RRNCX-C64HY-W2MM7-MCH9G-TJHMQ" 'PowerPoint
%ka% "9d3e4cca-e172-46f1-a2f4-1d2107051444", "G2KWX-3NW6P-PY93R-JXK2T-C9Y9V" 'Publisher
%ka% "734c6c6e-b0ba-4298-a891-671772b2bd1b", "NCJ33-JHBBY-HTK98-MYCV8-HMKHJ" 'Skype for Business
%ka% "059834fe-a8ea-4bff-b67b-4d006b5447d3", "PBX3G-NWMT6-Q7XBW-PYJGG-WXD33" 'Word
echo.
echo 'Office 2016
%ka% "829b8110-0e6f-4349-bca4-42803577788d", "WGT24-HCNMF-FQ7XH-6M8K7-DRTW9" 'Project Professional C2R-P
%ka% "cbbaca45-556a-4416-ad03-bda598eaa7c8", "D8NRQ-JTYM3-7J2DX-646CT-6836M" 'Project Standard C2R-P
%ka% "b234abe3-0857-4f9c-b05a-4dc314f85557", "69WXN-MBYV6-22PQG-3WGHK-RM6XC" 'Visio Professional C2R-P
%ka% "361fe620-64f4-41b5-ba77-84f8e079b1f7", "NY48V-PPYYH-3F4PX-XJRKJ-W4423" 'Visio Standard C2R-P
%ka% "e914ea6e-a5fa-4439-a394-a9bb3293ca09", "DMTCJ-KNRKX-26982-JYCKT-P7KB6" 'MondoR
%ka% "9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce", "HFTND-W9MK4-8B7MJ-B6C4G-XQBR2" 'Mondo
%ka% "d450596f-894d-49e0-966a-fd39ed4c4c64", "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99" 'Professional Plus
%ka% "dedfa23d-6ed1-45a6-85dc-63cae0546de6", "JNRGM-WHDWX-FJJG3-K47QV-DRTFM" 'Standard
%ka% "4f414197-0fc2-4c01-b68a-86cbb9ac254c", "YG9NW-3K39V-2T3HJ-93F3Q-G83KT" 'Project Professional
%ka% "da7ddabc-3fbe-4447-9e01-6ab7440b4cd4", "GNFHQ-F6YQM-KQDGJ-327XX-KQBVC" 'Project Standard
%ka% "6bf301c1-b94a-43e9-ba31-d494598c47fb", "PD3PC-RHNGV-FXJ29-8JK7D-RJRJK" 'Visio Professional
%ka% "aa2a7821-1827-4c2c-8f1d-4513a34dda97", "7WHWN-4T7MP-G96JF-G33KR-W8GF4" 'Visio Standard
%ka% "67c0fc0c-deba-401b-bf8b-9c8ad8395804", "GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW" 'Access
%ka% "c3e65d36-141f-4d2f-a303-a842ee756a29", "9C2PK-NWTVB-JMPW8-BFT28-7FTBF" 'Excel
%ka% "d8cace59-33d2-4ac7-9b1b-9b72339c51c8", "DR92N-9HTF2-97XKM-XW2WJ-XW3J6" 'OneNote
%ka% "ec9d9265-9d1e-4ed0-838a-cdc20f2551a1", "R69KK-NTPKF-7M3Q4-QYBHW-6MT9B" 'Outlook
%ka% "d70b1bba-b893-4544-96e2-b7a318091c33", "J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6" 'Powerpoint
%ka% "041a06cb-c5b8-4772-809f-416d03d16654", "F47MM-N3XJP-TQXJ9-BP99D-8K837" 'Publisher
%ka% "83e04ee1-fa8d-436d-8994-d31a862cab77", "869NQ-FJ69K-466HW-QYCP2-DDBV6" 'Skype for Business
%ka% "bb11badf-d8aa-470e-9311-20eaf80fe5cc", "WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6" 'Word
echo.
echo 'Office 2013
%ka% "dc981c6b-fc8e-420f-aa43-f8f33e5c0923", "42QTK-RN8M7-J3C4G-BBGYM-88CYV" 'Mondo
%ka% "b322da9c-a2e2-4058-9e4e-f59a6970bd69", "YC7DK-G2NP3-2QQC3-J6H88-GVGXT" 'Professional Plus
%ka% "b13afb38-cd79-4ae5-9f7f-eed058d750ca", "KBKQT-2NMXY-JJWGP-M62JB-92CD4" 'Standard
%ka% "4a5d124a-e620-44ba-b6ff-658961b33b9a", "FN8TT-7WMH6-2D4X9-M337T-2342K" 'Project Professional
%ka% "427a28d1-d17c-4abf-b717-32c780ba6f07", "6NTH3-CW976-3G3Y2-JK3TX-8QHTT" 'Project Standard
%ka% "e13ac10e-75d0-4aff-a0cd-764982cf541c", "C2FG9-N6J68-H8BTJ-BW3QX-RM3B3" 'Visio Professional
%ka% "ac4efaf0-f81f-4f61-bdf7-ea32b02ab117", "J484Y-4NKBF-W2HMG-DBMJC-PGWR7" 'Visio Standard
%ka% "6ee7622c-18d8-4005-9fb7-92db644a279b", "NG2JY-H4JBT-HQXYP-78QH9-4JM2D" 'Access
%ka% "f7461d52-7c2b-43b2-8744-ea958e0bd09a", "VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB" 'Excel
%ka% "fb4875ec-0c6b-450f-b82b-ab57d8d1677f", "H7R7V-WPNXQ-WCYYC-76BGV-VT7GH" 'Groove
%ka% "a30b8040-d68a-423f-b0b5-9ce292ea5a8f", "DKT8B-N7VXH-D963P-Q4PHY-F8894" 'InfoPath
%ka% "1b9f11e3-c85c-4e1b-bb29-879ad2c909e3", "2MG3G-3BNTT-3MFW9-KDQW3-TCK7R" 'Lync
%ka% "efe1f3e6-aea2-4144-a208-32aa872b6545", "TGN6P-8MMBC-37P2F-XHXXK-P34VW" 'OneNote
%ka% "771c3afa-50c5-443f-b151-ff2546d863a0", "QPN8Q-BJBTJ-334K3-93TGY-2PMBT" 'Outlook
%ka% "8c762649-97d1-4953-ad27-b7e2c25b972e", "4NT99-8RJFH-Q2VDH-KYG2C-4RD4F" 'Powerpoint
%ka% "00c79ff1-6850-443d-bf61-71cde0de305f", "PN2WF-29XG2-T9HJ7-JQPJR-FCXK4" 'Publisher
%ka% "d9f5b1c6-5386-495a-88f9-9ad6b41ac9b3", "6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7" 'Word
echo.
echo 'Office 2010
%ka% "09ed9640-f020-400a-acd8-d7d867dfd9c2", "YBJTT-JG6MD-V9Q7P-DBKXJ-38W9R" 'Mondo
%ka% "ef3d4e49-a53d-4d81-a2b1-2ca6c2556b2c", "7TC2V-WXF6P-TD7RT-BQRXR-B8K32" 'Mondo2
%ka% "6f327760-8c5c-417c-9b61-836a98287e0c", "VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB" 'Professional Plus
%ka% "9da2a678-fb6b-4e67-ab84-60dd6a9c819a", "V7QKV-4XVVR-XYV4D-F7DFM-8R6BM" 'Standard
%ka% "df133ff7-bf14-4f95-afe3-7b48e7e331ef", "YGX6F-PGV49-PGW3J-9BTGG-VHKC6" 'Project Professional
%ka% "5dc7bf61-5ec9-4996-9ccb-df806a2d0efe", "4HP3K-88W3F-W2K3D-6677X-F9PGB" 'Project Standard
%ka% "92236105-bb67-494f-94c7-7f7a607929bd", "D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ" 'Visio Premium
%ka% "e558389c-83c3-4b29-adfe-5e4d7f46c358", "7MCW8-VRQVK-G677T-PDJCM-Q8TCP" 'Visio Professional
%ka% "9ed833ff-4f92-4f36-b370-8683a4f13275", "767HD-QGMWX-8QTDB-9G3R2-KHFGJ" 'Visio Standard
%ka% "8ce7e872-188c-4b98-9d90-f8f90b7aad02", "V7Y44-9T38C-R2VJK-666HK-T7DDX" 'Access
%ka% "cee5d470-6e3b-4fcc-8c2b-d17428568a9f", "H62QG-HXVKF-PP4HP-66KMR-CW9BM" 'Excel
%ka% "8947d0b8-c33b-43e1-8c56-9b674c052832", "QYYW6-QP4CB-MBV6G-HYMCJ-4T3J4" 'Groove ^(SharePoint Workspace^)
%ka% "ca6b6639-4ad6-40ae-a575-14dee07f6430", "K96W8-67RPQ-62T9Y-J8FQJ-BT37T" 'InfoPath
%ka% "ab586f5c-5256-4632-962f-fefd8b49e6f4", "Q4Y4M-RHWJM-PY37F-MTKWH-D3XHX" 'OneNote
%ka% "ecb7c192-73ab-4ded-acf4-2399b095d0cc", "7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ" 'Outlook
%ka% "45593b1d-dfb1-4e91-bbfb-2d5d0ce2227a", "RC8FX-88JRY-3PF7C-X8P67-P4VTT" 'Powerpoint
%ka% "b50c4f75-599b-43e8-8dcd-1081a7967241", "BFK7F-9MYHM-V68C7-DRQ66-83YTP" 'Publisher
%ka% "2d0882e7-a4e7-423b-8ccc-70d91e0158b1", "HVHB3-C6FV7-KQX9W-YQG79-CRY7T" 'Word
%ka% "ea509e87-07a1-4a45-9edc-eba5a39f36af", "D6QFG-VBYP2-XQHM7-J97RH-VVRCK" 'Home and Business
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

