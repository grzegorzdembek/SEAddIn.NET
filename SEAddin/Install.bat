@echo off
color 0A
echo ========================================================
echo INSTALATOR DODATKU SOLID EDGE
echo ========================================================
echo.
echo Rejestrowanie pliku SolidEdgeAdd-In.dll w systemie...
echo.

"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" "%~dp0SolidEdgeAdd-In.dll" /codebase /tlb

echo.
echo ========================================================
echo INSTALACJA ZAKONCZONA! Sprawdz powyzsze komunikaty.
echo Jesli widzisz "Types registered successfully", jest super.
echo ========================================================
pause