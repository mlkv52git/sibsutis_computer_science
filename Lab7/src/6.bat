@echo off
:doWhile
ping -n 1 -w 10000 192.168.0.255 >nul
set SPISOK=
for %%i in (*) do set SPISOK=!SPISOK!  %%i
echo %DATE%   %TIME% >>mon.txt
echo %SPISOK%  >> mon.txt 
goto doWhile