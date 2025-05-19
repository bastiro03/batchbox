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
cls
title macOSVMPatcher v4
set header=macOSVMPatcher v4
echo %header%
echo -----------------
echo This program will patch a VirtualBox VM to make installing macOS on it possible.
echo Please close all VirtualBox instances before continuing. 
pause
echo.
echo Enter the name of the VM you want to install macOS on.
set /p vmname
pause
echo.
echo Patching "%vmname%"...
if exist %programfiles%\Oracle\VirtualBox\ cd %programfiles%\Oracle\VirtualBox\ else goto notfound
@echo on
VBoxManage.exe modifyvm "%vmname%" --cpuid-set 00000001 000106e5 00100800 0098e3fd bfebfbff
VBoxManage setextradata "%vmname%" "VBoxInternal/Devices/efi/0/Config/DmiSystemProduct" "iMac19,3"
VBoxManage setextradata "%vmname%" "VBoxInternal/Devices/efi/0/Config/DmiSystemVersion" "1.0"
VBoxManage setextradata "%vmname%" "VBoxInternal/Devices/efi/0/Config/DmiBoardProduct" "Iloveapple"
VBoxManage setextradata "%vmname%" "VBoxInternal/Devices/smc/0/Config/DeviceKey" "ourhardworkbythesewordsguardedpleasedontsteal(c)AppleComputerInc"
VBoxManage setextradata "%vmname%" "VBoxInternal/Devices/smc/0/Config/GetKeyFromRealSMC" 1
@echo off
echo.
echo Patching complete.
echo The program will now exit.
pause
exit /b

:notfound
echo.
echo No VirtualBox installation could be found.
echo The program will now exit.
pause
exit /b