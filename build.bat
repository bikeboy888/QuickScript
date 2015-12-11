@ECHO OFF

SETLOCAL
PATH=%PATH%;C:\Program Files (x86)\Microsoft Visual Studio 9.0\Common7\IDE

RMDIR /s /q Output
devenv.com QuickScript.sln /build "Release|Pocket PC 2003 (ARMV4)"
devenv.com QuickScript.sln /build "Release|Windows Mobile 5.0 Pocket PC SDK (ARMV4I)"

DIR "Output\Release\Pocket PC 2003 (ARMV4)"
DIR "Output\Release\Windows Mobile 5.0 Pocket PC SDK (ARMV4I)"
