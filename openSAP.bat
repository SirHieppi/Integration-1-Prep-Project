@ECHO OFF

cd %USERPROFILE%\Downloads

:: dir /o-d /b | findstr ".sap"

FOR /F "eol=| delims=" %%I IN ('DIR "./*.sap" /A-D /B /O-D /TW 2^>nul') DO (
    SET NewestFile=%%I
    GOTO FoundFile
)
ECHO No *.sap file found!
GOTO :EOF

:FoundFile

"%NewestFile%"

:EOF
