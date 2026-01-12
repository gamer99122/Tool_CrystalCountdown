@echo off
echo Publishing...
echo.
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
echo.
echo Done! Files are in: bin\Release\net8.0-windows\win-x64\publish\
echo.
pause
start "" "bin\Release\net8.0-windows\win-x64\publish"
