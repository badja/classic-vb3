@echo off
if d==%4 goto compile
if k==%4 goto compile
goto error

:compile
if errorlevel 1 goto delete
%583lnk %1.obj %2 %3

:delete
if %4==k goto end
if exist %1.obj del %1.obj
if exist %1.lst del %1.lst
goto end

:error
echo This batch file is designed only to be run by ADI TI-83 Assembler.

:end
echo.
