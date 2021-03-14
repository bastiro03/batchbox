@echo off
set header=Microsoft Office Activator - Version 0.2
set string26=Choose
set string28=Go back
set string63=Change color:
set string64=Dark theme
set string65=Light theme *please don't hurt my eyes edition*
set string66=Light theme *please hurt my eyes edition*
set string67=Yellow
set string68=Green
set string69=Red
set string70=Blue
goto menu

:menu
cls
echo ============================================================================
echo Microsoft Office Activator - Version 1.0
echo ============================================================================
echo #Currently Supported:
echo - Microsoft Office 2010 Standard
echo - Microsoft Office 2010 Professional Plus
echo - Microsoft Office 2013 Standard
echo - Microsoft Office 2013 Professional Plus
echo.
echo What do you want to do today?
echo A: Start the activation
echo B: Change the options [BETA]
echo C: Read the guide [BETA]
echo D: View the Credits
echo E: Exit the activator
choice /n /c ABCDE
goto menu%errorlevel%

:menu1
goto method
:menu2
goto change_color

:menu3
goto beta

:menu4
goto credits

:menu5
goto exit

:method
cls
echo What method would you like to use?
echo A: Standard (Office 2010)
echo B: Simple (Office 2013) [BETA]
echo C: Back to menu
choice /n /c ABC
goto method%errorlevel%

:method1
echo Standard activation in progress...
call standard

:method2
echo Simple activation in progress...
call simple

:method3
goto menu

:standard
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office14"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office14" else goto error
cscript //nologo ospp.vbs /unpkey:8R6BM >nul
cscript //nologo ospp.vbs /unpkey:H3GVB >nul
cscript //nologo ospp.vbs /inpkey:V7QKV-4XVVR-XYV4D-F7DFM-8R6BM >nul
cscript //nologo ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB >nul
set i=1
if %i%==1 set KMS_Sev=kms7.MSGuides.com
if %i%==2 set KMS_Sev=kms8.MSGuides.com
if %i%==3 set KMS_Sev=kms9.MSGuides.com
if %i%==4 goto unsupported
cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul
cscript //nologo ospp.vbs /act | find /i "successful"
set /a i+=1
if %i%==1 goto finish else goto error

:credits
cls
echo ============================================================================
echo #Official website: http://MSGuides.com
echo #Explanation Guide: https://docs.microsoft.com/en-us/previous-versions/tn-archive/ff793434(v=technet.10)
echo #Contact: msguides.com@gmail.com
echo #Support: donate.msguides.com
echo #Unofficial version by Basti Ro.
echo #Some code snippets ripped from "RiiConnect24 Patcher"
echo ============================================================================
echo Would you like to visit the official website or the guide?
echo T= Visit the official website
echo G= Visit the guide
echo A= No, go to activation
echo N= No, go back
choice /n /c TGAN
goto credits%errorlevel%

:credits1
explorer "http://MSGuides.com"

:credits2
explorer "https://docs.microsoft.com/en-us/previous-versions/tn-archive/ff793434(v=technet.10)"

:credits3
goto method

:credits4
goto menu

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
goto 64bit

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
choice /m "Do you really want to exit?"
if %errorlevel%==2 goto menu
if %errorlevel%==1 exit

:beta
echo This program is still in early development, so this feature has not been implemented yet.
echo You will now be sent back to the main menu...
pause
goto menu

:simple
if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office15"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" set folder="%ProgramFiles(x86)%\Microsoft Office\Office15"
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office14"
if exist "cd /d %ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" set folder="cd /d %ProgramFiles(x86)%\Microsoft Office\Office14"
cscript %folder%\ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
cscript %folder%\ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT
cscript %folder%\ospp.vbs /sethst:kms8.msguides.com
cscript %folder%\ospp.vbs /setprt:1688
cscript %folder%\ospp.vbs /act
goto finish

:change_color
cls
echo %header%
echo -----------------------------------------------------------------------------------------------------------------------------
echo.
echo %string63%
echo.
echo 1. %string64%
echo 2. %string65%
echo 3. %string66%
echo 4. %string67%
echo 5. %string68%
echo 6. %string69%
echo 7. %string70%
echo.
echo E. %string28%
set /p s=%string26%: 
if %s%==1 set tempcolor=07&goto save_color
if %s%==2 set tempcolor=70&goto save_color
if %s%==3 set tempcolor=f0&goto save_color
if %s%==4 set tempcolor=6&goto save_color
if %s%==5 set tempcolor=a&goto save_color
if %s%==6 set tempcolor=c&goto save_color
if %s%==7 set tempcolor=3&goto save_color
if %s%==e goto menu
if %s%==E goto menu

:save_color
if exist "%MainFolder%\background_color.txt" del /q "%MainFolder%\background_color.txt"
color %tempcolor%
echo>>"%MainFolder%\background_color.txt" %tempcolor%
goto change_color