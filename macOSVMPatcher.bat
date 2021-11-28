@echo off
cls
title macOSVMPatcher v3
set header=macOSVMPatcher v3
echo %header%
echo -----------------
echo The program will patch your VM for use with macOS.
echo It is recommended to close all VirtualBox instances before patching. 
pause
cls
echo %header%
echo -----------------
echo Enter the name of the VM you want to run macOS on.
set /p vmname
pause
cls
echo %header%
echo -----------------
echo Patching "%vmname%"...
if exist C:\Program Files\Oracle\VirtualBox\ cd C:\Program Files\Oracle\VirtualBox\ else goto notfound
VBoxManage.exe modifyvm "%vmname%" --cpuid-set 00000001 000106e5 00100800 0098e3fd bfebfbff
VBoxManage setextradata "%vmname%" "VBoxInternal/Devices/efi/0/Config/DmiSystemProduct" "iMac11,3"
VBoxManage setextradata "%vmname%" "VBoxInternal/Devices/efi/0/Config/DmiSystemVersion" "1.0"
VBoxManage setextradata "%vmname%" "VBoxInternal/Devices/efi/0/Config/DmiBoardProduct" "Iloveapple"
VBoxManage setextradata "%vmname%" "VBoxInternal/Devices/smc/0/Config/DeviceKey" "ourhardworkbythesewordsguardedpleasedontsteal(c)AppleComputerInc"
VBoxManage setextradata "%vmname%" "VBoxInternal/Devices/smc/0/Config/GetKeyFromRealSMC" 1
cls
echo %header%
echo -----------------
echo Patching complete.
echo The program will now exit...
pause
exit /b

:notfound
cls
echo %header%
echo -----------------
echo Your VirtualBox installation was not found.
echo The program will now exit...
pause
exit /b