@echo off
cls
title YourAge: A test program by Bastiro03
echo How old are you?
set /p a
if %a%==1 echo You are only a year old? You must be kidding! &&else echo So you are %a% years old. Very interesting.
pause