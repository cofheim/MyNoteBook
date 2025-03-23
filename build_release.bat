@echo off
echo Building Release version...

:: Путь к MSBuild
set MSBUILD="C:\Program Files\Microsoft Visual Studio\2022\Community\Msbuild\Current\Bin\MSBuild.exe"

:: Сборка проекта
%MSBUILD% NBapp.sln /p:Configuration=Release /p:Platform=x64

:: Создание папки для релиза
if not exist "Release" mkdir Release

:: Копирование файлов
echo Copying files...
xcopy "x64\Release\NBapp.exe" "Release\" /Y
xcopy "x64\Release\*.dll" "Release\" /Y

:: Создание папки для данных
if not exist "Release\data" mkdir Release\data

echo Done!
pause 