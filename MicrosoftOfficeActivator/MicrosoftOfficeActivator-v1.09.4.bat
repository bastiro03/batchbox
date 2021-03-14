@echo off

:menu
cls
title Microsoft Office Activator - v1.09.4
set header=Microsoft Office Activator - v1.09.4
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo #Currently Supported:
echo - Microsoft Office 2010 Standard
echo - Microsoft Office 2010 Professional Plus
echo - Microsoft Office 2013 Standard
echo - Microsoft Office 2013 Professional Plus
echo.
echo Main Menu:
echo 1: Start 
echo 2: Settings
echo 3: Help
echo 4: Credits
echo 5: Exit
choice /n /c 12345
if %errorlevel%==1 goto method
if %errorlevel%==2 goto settings
if %errorlevel%==3 goto help
if %errorlevel%==4 goto credits
if %errorlevel%==5 goto exit

:method
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo Select activation method:
echo 1: Standard
echo 2: Simple
echo 3: Legacy
echo 4: Back to menu
choice /n /c 1234
if %errorlevel%==1 goto standard
if %errorlevel%==2 goto simple
if %errorlevel%==3 goto legacy
if %errorlevel%==4 goto menu

:standard
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo Activation in progress...
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office14"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office14"
cscript //nologo ospp.vbs /unpkey:8R6BM >nul
cscript //nologo ospp.vbs /unpkey:H3GVB >nul
cscript //nologo ospp.vbs /inpkey:V7QKV-4XVVR-XYV4D-F7DFM-8R6BM >nul
cscript //nologo ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB >nul
set i=1
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto error
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul
cscript //nologo ospp.vbs /act | find /i "successful" >nul
set /a i+=1
goto finish

:simple
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo Activation in progress...
if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office15"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" set folder="%ProgramFiles(x86)%\Microsoft Office\Office15"
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office14"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" set folder="%ProgramFiles(x86)%\Microsoft Office\Office14"
cscript %folder%\ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
cscript %folder%\ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT
cscript %folder%\ospp.vbs /sethst:kms8.msguides.com
cscript %folder%\ospp.vbs /setprt:1688
cscript %folder%\ospp.vbs /act
goto finish

:legacy
@echo off
title Activate Microsoft Office 2010 Volume for FREE!&cls&echo ============================================================================&echo #Project: Activating Microsoft software products for FREE without software&echo ============================================================================&echo.&echo #Supported products:&echo - Microsoft Office 2010 Standard Volume&echo - Microsoft Office 2010 Professional Plus Volume&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office14")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office14")&echo.&echo ============================================================================&echo Activating your Office...&cscript //nologo ospp.vbs /unpkey:8R6BM >nul&cscript //nologo ospp.vbs /unpkey:H3GVB >nul&cscript //nologo ospp.vbs /inpkey:V7QKV-4XVVR-XYV4D-F7DFM-8R6BM >nul&cscript //nologo ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB >nul&set i=1
:server
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto notsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul&echo ============================================================================&echo.&echo.
cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ============================================================================&echo.&echo #My official blog: MSGuides.com&echo.&echo #How it works: bit.ly/kms-server&echo.&echo #Please feel free to contact me at msguides.com@gmail.com if you have any questions or concerns.&echo.&echo #Please consider supporting this project: donate.msguides.com&echo #Your support is helping me keep my servers running everyday!&echo.&echo ============================================================================&choice /n /c YN /m "Would you like to visit my blog [Y,N]?" & if errorlevel 2 exit) || (echo The connection to my KMS server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto server)
explorer "http://MSGuides.com"&goto halt
:notsupported
echo.&echo ============================================================================&echo Sorry! Your version is not supported.
:halt
pause >nul

:credits
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo Credits:
echo ============================================================================
echo #Official website: http://MSGuides.com
echo #How it works: https://docs.microsoft.com/en-us/previous-versions/tn-archive/ff793434(v=technet.10)
echo #Contact: msguides.com@gmail.com
echo #Support: donate.msguides.com
echo #Unofficial version by Basti Ro.
echo ============================================================================
echo 1: Visit the official website
echo 2: Read the guide
echo 3: Back to menu
choice /n /c 1234
if %errorlevel%==1 explorer "http://MSGuides.com"& exit
if %errorlevel%==2 explorer "https://docs.microsoft.com/en-us/previous-versions/tn-archive/ff793434(v=technet.10)"& exit
if %errorlevel%==3 goto menu

:error
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo Activation failed.
echo 1: Back to menu
echo 2: Try again
echo 3: Exit
choice /n /c 123
if %errorlevel%==1 goto menu
if %errorlevel%==2 goto method
if %errorlevel%==3 goto exit

:finish
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo Activation was completed. 
echo 1: Back to menu
echo 2: Exit
choice /n /c 12
if %errorlevel%==1 goto menu
if %errorlevel%==2 goto exit

:exit
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
choice /m "Are you sure you want to quit?"
if %errorlevel%==1 exit
if %errorlevel%==2 goto menu

:beta
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo This program is still in development, and this feature has not been implemented yet.
echo You will now be sent back to the main menu...
pause
goto menu

:change_color
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo Change color:
echo.
echo 1. Dark theme
echo 2. Light theme (don't hurt my eyes)
echo 3. Light theme (hurt my eyes)
echo 4. Yellow
echo 5. Green
echo 6. Red
echo 7. Blue
echo 8. Back to menu
choice /n /c 12345678 
if %errorlevel%==1 color 07
if %errorlevel%==2 color 70
if %errorlevel%==3 color f0
if %errorlevel%==4 color 6
if %errorlevel%==5 color a
if %errorlevel%==6 color c
if %errorlevel%==7 color 3
if %errorlevel%==8 goto menu

:settings
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo Settings:
echo 1: Colors
echo 2: Back to menu
choice /n /c AB
if %errorlevel%==1 goto change_color
if %errorlevel%==2 goto menu

:help
echo Product keys:
echo These are just in case you can't find yours.
echo For legal reasons you should still use the one you actually bought.
echo V7QKV-4XVVR-XYV4D-F7DFM-8R6BM
echo VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
echo YC7DK-G2NP3-2QQC3-J6H88-GVGXT
pause
goto menu