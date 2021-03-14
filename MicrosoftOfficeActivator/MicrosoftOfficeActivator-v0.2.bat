@echo off
title Microsoft Office Activator
goto menu
:menu
cls
echo ============================================================================
echo Microsoft Office Activator - Version 0.2
echo ============================================================================
echo #Currently Supported:
echo - Microsoft Office 2010 Standard
echo - Microsoft Office 2010 Professional Plus
echo.
echo What do you want to do today?
echo A: Start the activation
echo B: Change the options [NOT WORKING]
echo C: Read a troubleshooting guide [NOT WORKING]
echo D: View the Credits
echo E: Exit the activator
choice /n /c ABCDE
goto menu%errorlevel%
:menu1
goto bit
:menu2
goto beta
:menu3
goto beta
:menu4
goto credits
:menu5
goto exit
:bit
cls
echo Is your Office installation 32-Bit or 64-Bit?
echo A: 32-Bit
echo B: 64-Bit
echo C: I don't know, take me back
choice /n /c ABC
goto bit%errorlevel%
:bit1
goto 32bit
:bit2
goto 64bit
:bit3
goto menu
:64bit
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office14" else goto error
goto keys
:32bit
if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office14" else goto error
goto keys
:credits
cls
echo ============================================================================
echo #Official website: http://MSGuides.com
echo #Explanation Guide: bit.ly/kms-server
echo #Contact: msguides.com@gmail.com
echo #Support: donate.msguides.com
echo #This unofficial version was made by bastiro03
echo ============================================================================
echo Would you like to visit the official website?
echo Y=Yes, please
echo A=I would like to activate now
echo N=No, go back
choice /n /c YAN
goto credits%errorlevel%
:credits1
explorer "http://MSGuides.com"
pause >nul
:credits2
goto bit
:credits3
goto menu
:keys
cscript //nologo ospp.vbs /unpkey:8R6BM >nul
cscript //nologo ospp.vbs /unpkey:H3GVB >nul
cscript //nologo ospp.vbs /inpkey:V7QKV-4XVVR-XYV4D-F7DFM-8R6BM >nul
cscript //nologo ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB >nul
set i=1
goto servers
:servers
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto unsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul
cscript //nologo ospp.vbs /act | find /i "successful"
goto unused
:error
echo An error occured. Go to http://MSGuides.com for more information.
echo Return to menu?
echo A: Yes
echo B: No, try again
echo C: No, exit the activator
goto error%errorlevel%
:error1
goto menu
:error2
goto start
:error3
goto exit
:finish
echo The activation was completed. What now?
echo A: Back to the main menu
echo B: Exit the activator
choice /n /c AB
goto finish%errorlevel%
:finish1
goto menu
:finish2
goto exit
:exit
echo Do you really want to exit?
pause
exit
:beta
echo This program is an early beta, so this feature has not been implemented yet.
pause
echo You will now be sent back to the main menu...
goto menu
:unused
echo Connection to the KMS server failed. Trying another one... 
set /a i+=1
if %i%==1 goto finish else goto error