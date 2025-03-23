@echo off
echo Preparing release files...

rem Create Release directory if it doesn't exist
if not exist "Release" mkdir Release

rem Copy the executable and DLLs
xcopy /Y "x64\Release\NBapp.exe" "Release\"
xcopy /Y "x64\Release\*.dll" "Release\"

rem Create an empty contacts.json file if it doesn't exist
if not exist "Release\contacts.json" (
    echo [] > "Release\contacts.json"
    echo Created empty contacts.json file
)

echo Release preparation completed successfully!
echo.
echo Files in Release directory:
dir "Release"

pause 