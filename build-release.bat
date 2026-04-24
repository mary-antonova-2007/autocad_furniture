@echo off
setlocal

if "%ACAD%"=="" (
  echo Set ACAD to your AutoCAD install folder first.
  echo Example:
  echo   set ACAD=C:\Program Files\Autodesk\AutoCAD 2025
  exit /b 1
)

set MSBUILD=%ProgramFiles%\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe
if not exist "%MSBUILD%" set MSBUILD=%ProgramFiles%\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe
if not exist "%MSBUILD%" set MSBUILD=%ProgramFiles%\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe
if not exist "%MSBUILD%" set MSBUILD=dotnet

if not "%MSBUILD%"=="dotnet" if not exist "%MSBUILD%" (
  echo MSBuild was not found.
  exit /b 1
)

if "%MSBUILD%"=="dotnet" (
  dotnet build AutoCAD_BoardSorter.sln -c Release
) else (
  "%MSBUILD%" AutoCAD_BoardSorter.sln /p:Configuration=Release
)
exit /b %ERRORLEVEL%
