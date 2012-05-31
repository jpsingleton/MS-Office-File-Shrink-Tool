@echo off
REM This script decompresses an office 2007/2010 file and repacks it smaller.
REM James Singleton 2011 - unop.co.uk
REM This work is licensed under a Creative Commons Attribution 3.0 Unported License.

REM decompress file to temp folder
echo Unpacking: %~n1
START "Unpacking..." /MIN /BELOWNORMAL /WAIT 7za x -o"%~dp1%~n1_working" %1 -ssw

REM exit here if decompress fails
IF %ERRORLEVEL% NEQ 0 EXIT

REM recompress file
echo.
echo Repacking: %~n1_shrunk
START "Repacking..." /MIN /BELOWNORMAL /WAIT 7za a -tzip "%~dp1%~n1_shrunk%~x1" "%~dp1%~n1_working\*" -mx9 -ssw

REM remove the temp folder
rmdir "%~dp1%~n1_working" /s /q