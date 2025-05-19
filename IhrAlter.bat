@echo off
cls
title IhrAlter
echo Wie alt bist du?
set /p a
if %a%<=1 echo Du Lügner. Niemand kann in so einem jungen Alter einen PC bedienen.
else echo Du bist also %a% Jahre alt. Sehr interessant.
pause