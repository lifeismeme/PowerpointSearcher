@echo off
dir /w | findstr "\.pptx$" | dotnet "%~dp0./netcoreapp2.1/PowerpointSearcher/PowerpointSearcher.dll" %1
