@echo off
chcp 65001 > nul

cd /d "%~dp0"

echo 빌드 시작...

dotnet publish -c Release -r win-x64 --self-contained true -p:PublishingSingleFile=true

if errorlevel 1 (
    echo.
    echo 빌드 실패
) else (
    echo.
    echo 빌드 완료
    echo 결과: bin\Release\net8.0-windows\win-x64\publish
)

pause