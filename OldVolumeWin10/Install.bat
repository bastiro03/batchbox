@echo off
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo Administrator privileges are required. Try restarting with elevated privileges...
    goto UACPrompt
) else (
    goto gotAdmin
)

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
title OldVolumeWin10
echo OldVolumeWin10
echo --------------
echo This program activates the old volume slider from Windows 8.1 in Windows 10.
echo.
echo 1: ON
echo 2: OFF
choice /c 12 /n
if %errorlevel%==1 goto on
if %errorlevel%==2 goto off

:on
ON.reg
pause
exit

:off
OFF.reg
pause
exit