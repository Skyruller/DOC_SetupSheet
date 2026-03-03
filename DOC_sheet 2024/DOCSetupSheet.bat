@echo off
echo Registering DOC SetupSheet plugin for PowerMill...

REM COM-регистрация
C:\WINDOWS\Microsoft.NET\Framework64\v4.0.30319\regasm.exe "C:\Program Files\Autodesk\DOC_sheet 2024\SetupSheet.dll" /register /codebase

REM Добавление в Implemented Categories
reg.exe ADD "HKCR\CLSID\{8C96851C-7A01-4389-8FBF-22C3DC7B09FD}\Implemented Categories\{311b0135-1826-4a8c-98de-f313289f815e}" /reg:64 /f

echo Done! Now enable the plugin in PowerMill: File > Options > Manage Installed Plugins
pause