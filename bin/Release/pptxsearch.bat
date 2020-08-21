@echo off
dir *.pptx /b | findstr "\.pptx$" | dotnet "%~dp0./netcoreapp2.1/PowerpointSearcher/PowerpointSearcher.dll" %1
