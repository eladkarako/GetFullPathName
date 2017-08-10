@echo off
if exist %1 goto FILEEXIST
goto NOFILE

:FILEEXIST
C:\PROGRA~2\MICROS~4\VB98\LINK.EXE /EDIT /SUBSYSTEM:CONSOLE %1
rem "C:\Program Files\Microsoft Visual Studio\vb98\LINK.EXE" /EDIT /SUBSYSTEM:CONSOLE %1
echo.
echo Finish.
goto ENDME

:NOFILE
echo.
echo need argument to reLink (exe file created by Visual Basic 6.+)
goto ENDME

:ENDME
echo Press Any Key to Finish (Quit ReLinker)
echo                                        -----=================-----
echo                                        ---===  Elad Karako  ===---
echo                                        ---===  2008.Israel  ===---
echo                                        -----=================-----
pause >nul
