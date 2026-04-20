@echo off
cd C:\GitHub\TableDataConverter
dotnet publish -c Release -r win-x64 --self-contained false /p:PublishSingleFile=true
pause