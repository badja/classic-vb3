@echo off
if d==%4 goto compile
if k==%4 goto compile
goto error

:compile
cd %1
tasm -t80 -i %2 %3.obj %3.lst
goto end

:error
echo This batch file is designed only to be run by ADI TI-83 Assembler.

:end
echo.
