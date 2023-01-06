
@echo off
echo.Current User is '%USERNAME%'

cd %~dp0

set "filemask=ChilkatAx*.dll"
for %%A in (%filemask%) do regsvr32 %%A || GOTO:EOF

ECHO.&PAUSE&GOTO:EOF

