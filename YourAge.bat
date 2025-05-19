@echo off
cls
title YourAge
echo How old are you?
set /p a
if %a%<=3 echo You liar. No one can use a computer at such a young age.
else echo So you are %a% years old. Very interesting.
pause