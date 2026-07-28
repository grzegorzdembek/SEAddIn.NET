@echo off
color 0C
echo ========================================================
echo DEINSTALATOR DODATKU SOLID EDGE
echo ========================================================
echo.
echo Czyszczenie wpisow w Rejestrze Windows...
echo.

"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" "%~dp0SolidEdgeAdd-In.dll" /unregister

echo.
echo ========================================================
echo DEINSTALACJA ZAKONCZONA!
echo Wpis usuniety. Mozesz bezpiecznie usunac folder z dysku.
echo ========================================================
pause