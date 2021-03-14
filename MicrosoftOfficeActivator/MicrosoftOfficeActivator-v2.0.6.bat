@echo off

:boot
title Microsoft Office Activator - v2.0.6
set header=Microsoft Office Activator - v2.0.6
set method=1

:start_languages
setlocal disableDelayedExpansion
FOR /F "tokens=2 delims==" %%a IN ('%wmic_path% os get OSLanguage /Value') DO set OSLanguage=%%a
if "%OSLanguage%"=="1033" set language=en-US& call :set_language_english
if "%OSLanguage%"=="1031" set language=de-DE& call :set_language_german

:set_language_english
set string1=Currently supported Office versions:
set string2=Microsoft Office 2010 Standard
set string3=Microsoft Office 2010 Professional Plus
set string4=Microsoft Office 2013 Standard
set string5=Microsoft Office 2013 Professional Plus

set string6=Main Menu:
set string7=Start
set string8=Settings
set string9=Help
set string10=Credits
set string11=Exit

set string12=Activation in progress...

set string13=Official website
set string14=How it works
set string15=Contact
set string16=Donations
set string17=Unofficial version by Basti Ro.
set string18=Visit the official website
set string19=Read the guide
set string20=Back to Menu

set string21=Activation failed.
set string22=Try again

set string23=Activation was completed. 

set string24=Are you sure you want to quit?

set string25=This program is still in development, and this feature has not been implemented yet.

set string26=Colors
set string27=Activation method

set string28=Change the terminal's colors
set string29=Default
set string30=Light
set string31=Yellow
set string32=Green
set string33=Red
set string34=Blue

set string35=Select activation method
set string36=Default
set string37=Simple
set string38=Legacy

set string39=This program is for troubleshooting use only. Please make sure you legitimately bought the Office suite you want to activate.
set string40=Usage is very simple, just start the program and follow th instructions on the screen.
set string41=Product-Keys (only if the real one isn't usable)

set string42=Language
set string43=Choose the language for the program to use
set string44=German
set string45=English
goto menu

:set_language_german
set string1=Derzeit unterstützte Office-Versionen:
set string2=Microsoft Office 2010 Standard
set string3=Microsoft Office 2010 Professional Plus
set string4=Microsoft Office 2013 Standard
set string5=Microsoft Office 2013 Professional Plus

set string6=Hauptmenü:
set string7=Start
set string8=Einstellungen
set string9=Hilfe
set string10=Credits
set string11=Beenden

set string12=Aktivierung läuft...

set string13=Offizielle Website
set string14=Wie es funktioniert
set string15=Kontakt
set string16=Spenden
set string17=Inoffizielle Version von Basti Ro.
set string18=Die offizielle Website besuchen
set string19=Die Anleitung lesen

set string20=Zurück zum Hauptmenü

set string21=Aktivierung fehlgeschlagen.
set string22=Erneut versuchen

set string23=Aktivierung abgeschlossen.

set string24=Wirklich beenden?

set string25=Dieses Programm befindet sich noch in der Entwicklungsphase, daher wurde diese Funktion noch nicht eingebaut.

set string26=Fensterfarben

set string27=Aktivierungsart

set string28=Ändere die Farben des Fensters
set string29=Standard
set string30=Hell
set string31=Gelb
set string32=Grün
set string33=Rot
set string34=Blau

set string35=Wähle die Aktivierungsart
set string36=Standard
set string37=Simpel
set string38=Legacy (Englisch)

set string39=Dieses Programm dient nur zur Fehlerbehebung. Stellen Sie sicher, dass Sie die Office-Suite, die Sie aktivieren möchten, legal erworben haben.
set string40=Die Verwendung ist sehr einfach, starten Sie einfach das Programm und folgen sie den Anweisungen auf dem Bildschirm.
set string41=Produktschlüssel (wenn der eigene nicht verwendbar ist)

set string42=Sprache
set string43=Wähle die für das Programm zu verwendende Sprache
set string44=Deutsch
set string45=Englisch
goto menu

:menu
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo %string1%
echo %string2%
echo %string3%
echo %string4%
echo %string5%
echo.
echo %string6%
echo 1: %string7%
echo 2: %string8%
echo 3: %string9%
echo 4: %string10%
echo 5: %string11%
choice /n /c 12345
if %errorlevel%==1 goto start
if %errorlevel%==2 goto settings
if %errorlevel%==3 goto help
if %errorlevel%==4 goto credits
if %errorlevel%==5 goto exit

:start
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo %string12%
if %method%==1 call :standard
if %method%==2 call :simple
if %method%==3 goto legacy


:standard
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office14"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office14"
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office15"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office15"
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
echo %string10%
echo ============================================================================
echo %string13%: http://MSGuides.com
echo %string14%: https://docs.microsoft.com/en-us/previous-versions/tn-archive/ff793434(v=technet.10)
echo %string15%: msguides.com@gmail.com
echo %string16%: donate.msguides.com
echo %string17%
echo ============================================================================
echo 1: %string18%
echo 2: %string19%
echo 3: %string20%
choice /n /c 1234
if %errorlevel%==1 explorer "http://MSGuides.com"& exit
if %errorlevel%==2 explorer "https://docs.microsoft.com/en-us/previous-versions/tn-archive/ff793434(v=technet.10)"& exit
if %errorlevel%==3 goto menu

:error
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo %string21%
echo 1: %string20%
echo 2: %string22%
echo 3: %string11%
choice /n /c 123
if %errorlevel%==1 goto menu
if %errorlevel%==2 goto start
if %errorlevel%==3 goto exit

:finish
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo %string23%
echo 1: %string20%
echo 2: %string11%
choice /n /c 12
if %errorlevel%==1 goto menu
if %errorlevel%==2 goto exit

:exit
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo %string24%
choice
if %errorlevel%==1 exit
if %errorlevel%==2 goto menu

:beta
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo %string25%
echo %string20%
pause
goto menu

:settings
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo %string8%:
echo 1: %string26%
echo 2: %string27%
echo 3: %string42%
echo 4: %string20%
choice /n /c 1234
if %errorlevel%==1 goto color
if %errorlevel%==2 goto method
if %errorlevel%==3 goto language
if %errorlevel%==4 goto menu

:color
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo %string28%:
echo.
echo 1. %string29%
echo 2. %string30% (soft)
echo 3. %string30% (hardcore)
echo 4. %string31%
echo 5. %string32%
echo 6. %string33%
echo 7. %string34%
echo 8. %string20%
choice /n /c 12345678 
if %errorlevel%==1 color 07
if %errorlevel%==2 color 70
if %errorlevel%==3 color f0
if %errorlevel%==4 color 6
if %errorlevel%==5 color a
if %errorlevel%==6 color c
if %errorlevel%==7 color 3
if %errorlevel%==8 goto menu

:method
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo %string35%:
echo.
echo 1. %string36%
echo 2. %string37%
echo 3. %string38%
echo 4. %tect20%
choice /n /c 1234
if %errorlevel%==1 set method=1
if %errorlevel%==2 set method=2
if %errorlevel%==3 set method=3
if %errorlevel%==4 goto menu

:language
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo %string43%:
echo.
echo 1: %string44%
echo 2: %string45%
echo 3: %string20%
choice /n /C 123
if %errorlevel%==1 goto german
if %errorlevel%==2 goto english
if %errorlevel%==3 goto menu

:help
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo %string39%
echo %string40%
echo.
echo %string41%:
echo V7QKV-4XVVR-XYV4D-F7DFM-8R6BM
echo VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
echo YC7DK-G2NP3-2QQC3-J6H88-GVGXT
pause
goto menu