@echo off
setlocal
echo Starting automatic update...
cd %~dp0
for %%f in ("%~dp0") do (
set dir=%%~sf
)
echo|set /p=Finding latest version... 
powershell -Command "(New-Object Net.WebClient).DownloadFile('http://192.168.1.105/files/spec sheet/version.txt', '%dir%version.txt')"
if not exist %dir%version.txt (
echo Failed, aborting.
goto cleanup
)
echo OK
powershell -Command "(New-Object Net.WebClient).DownloadFile('http://192.168.1.105/files/spec sheet/zipjs.bat', '%dir%zipjs.bat')"
FOR /F "delims=," %%i IN (%dir%version.txt) DO (
set version=%%i
set file='http://192.168.1.105/files/spec sheet/%%i.zip'
)
echo|set /p=Downloading latest version... 
powershell -Command "(New-Object Net.WebClient).DownloadFile(%file%, '%dir%%version%.zip')"
if not exist "%dir%%version%.zip" (
echo Failed, aborting.
goto cleanup
)
echo OK
if not exist "%dir%Temp" mkdir "%dir%Temp"
echo|set /p=Extracting %version%... 
2>nul (
  >>"%dir%SpecSheet.xlsm" (call )
) && (if exist "%dir%SpecSheet.xlsm" del "%dir%SpecSheet.xlsm")
call %dir%zipjs.bat unzip -source "%dir%%version%.zip" -destination "%dir%Temp" -keep yes -force yes
if not exist %dir%Temp (
echo Failed, aborting.
goto cleanup
)
xcopy /D /E /Y /Q "%dir%Temp" "%dir%"
rem #### upload/update archive files
echo Syncing archive...
if not exist "%dir%archive\exclude.txt" (
@echo \saves\>"%dir%archive\exclude.txt"
)
if not exist "%dir%saves\exclude.txt" (
copy /y NUL "%dir%saves\exclude.txt" >NUL
)
if not exist Y:\ (
net use Y: \\192.168.1.105\archive Portland1 /user:tcg\tcg> nul
if not exist Y:\ (
net use Y: \\etdock1\archive Portland1 /user:tcg\tcg> nul
if not exist Y:\ (
echo Failed, aborting.
goto cleanup
)
)
)
xcopy /Y /Q "Y:\exclude.txt" "%dir%archive\exclude.txt"> nul
xcopy /Y /Q "Y:\saves\exclude.txt" "%dir%saves\exclude.txt"> nul
echo|set /p=Uploading archive files...   
xcopy /D /E /Y /Q "%dir%archive" "Y:/" /exclude:%dir%archive\exclude.txt
echo|set /p=Downloading archive files... 
xcopy /D /E /Y /Q "Y:/" "%dir%archive" /exclude:%dir%archive\exclude.txt
echo|set /p=Uploading save files...      
xcopy /D /E /Y /Q "%dir%saves" "Y:/saves" /exclude:%dir%saves\exclude.txt
echo|set /p=Downloading save files...    
xcopy /D /E /Y /Q "Y:/saves" "%dir%saves" /exclude:%dir%saves\exclude.txt
net use Y: /delete /yes> nul
echo Update completed.

:cleanup
echo|set /p=Cleaning up... 
if exist "%dir%zipjs.bat" del "%dir%zipjs.bat"
if exist "%dir%version.txt" del "%dir%version.txt"
if exist "%dir%%version%.zip" del "%dir%%version%.zip"
if exist "%dir%Temp" rmdir /S /Q "%dir%Temp"
if exist "%dir%Temp" (
echo Failed, aborting.
Exit /b
)
echo OK
pause
rem launch specsheet if it's not already open
rem -- if excel is installed
reg query "HKLM\Software\Microsoft\Office\Excel" >nul
if %errorlevel% equ 0 (
2>nul (
  >>"%dir%SpecSheet.xlsm" (call )
) && (if exist "%dir%SpecSheet.xlsm" Start "" "%dir%SpecSheet.xlsm")
)
endlocal
