REM @ECHO OFF
If NOT EXIST "release" mkdir "release"
SET "RELEASE_FOLDER=release"
SET "VERSION=default"

FOR /F %%i in ('powershell -Command "Get-Content -Path src\main.au3 | Select-String -Pattern `^\s*Global\s*\$sVersion\s*=\s*""(\d+\.\d+\.\d+).*?""$` | % {'$($_.matches.groups[1])'}"') do (
	set VERSION=%%i
)

REM FOR /F %%i IN ('FINDSTR /R "^Global $sVersion = $" "src\main.au3" ^| FINDSTR /R "[0-9]\.[0-9]\.[0-9]"') DO (
REM 	SET version=%%i
REM )

ECHO %version%
IF "%VERSION%" equ "default" GOTO exit-version-not-parsed

GOTO exit-pause

:release_source
REM create source release
SET "RELEASE_SOURCE=FilesAu3-Src"
ECHO "Create source release: %RELEASE_SOURCE%"
powershell -Command "Compress-Archive -Path 'src', 'README.md' -DestinationPath '%RELEASE_FOLDER%\%RELEASE_SOURCE%.zip'"

:release_x86
REM create x86 release
SET "RELEASE_X86=FilesAu3-x86"
ECHO "Create source release: %RELEASE_SOURCE%"
powershell -Command "Compress-Archive -Path 'src', 'README.md' -DestinationPath '%RELEASE_FOLDER%\%RELEASE_SOURCE%.zip'"

:release_x64
REM create x64 release
SET "RELEASE_X64=FilesAu3-x64"
ECHO "Create source release: %RELEASE_SOURCE%"
powershell -Command "Compress-Archive -Path 'src', 'README.md' -DestinationPath '%RELEASE_FOLDER%\%RELEASE_SOURCE%.zip'"

GOTO exit

:exit-version-not-parsed
ECHO Error: Version could not be parsed.

:exit-pause
pause

:exit