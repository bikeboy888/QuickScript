@ECHO OFF

SETLOCAL
PATH=%PATH%;C:\Program Files (x86)\Microsoft Visual Studio 9.0\Common7\IDE

devenv.com QuickScript.sln /build "Release|Pocket PC 2003 (ARMV4)"

DIR "Output\Release\Pocket PC 2003 (ARMV4)"

REM devenv.exe QuickScript.sln /build /project "Release|Pocket PC 2003 (ARMV4)"
