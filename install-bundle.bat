@echo off
setlocal

set CONFIG=Release
set BUNDLE_NAME=AutoCAD_BoardSorter.bundle
set BUNDLE_SOURCE=%~dp0bundle\%BUNDLE_NAME%
set BUNDLE_TARGET=%APPDATA%\Autodesk\ApplicationPlugins\%BUNDLE_NAME%
set BUNDLE_TARGET_COMMON=%PROGRAMDATA%\Autodesk\ApplicationPlugins\%BUNDLE_NAME%
set BUILD_OUTPUT=%~dp0src\AutoCAD_BoardSorter\bin\%CONFIG%
set ACAD_YEAR=2022

call "%~dp0build-release.bat"
if errorlevel 1 exit /b %ERRORLEVEL%

if exist "%BUNDLE_TARGET%" rmdir /s /q "%BUNDLE_TARGET%"
mkdir "%BUNDLE_TARGET%\Contents\Windows\%ACAD_YEAR%"

xcopy "%BUNDLE_SOURCE%\PackageContents.xml" "%BUNDLE_TARGET%\" /y >nul
xcopy "%BUILD_OUTPUT%\AutoCAD_BoardSorter.dll" "%BUNDLE_TARGET%\Contents\Windows\%ACAD_YEAR%\" /y >nul
if exist "%BUILD_OUTPUT%\AutoCAD_BoardSorter.deps.json" xcopy "%BUILD_OUTPUT%\AutoCAD_BoardSorter.deps.json" "%BUNDLE_TARGET%\Contents\Windows\%ACAD_YEAR%\" /y >nul
if exist "%BUILD_OUTPUT%\AutoCAD_BoardSorter.pdb" xcopy "%BUILD_OUTPUT%\AutoCAD_BoardSorter.pdb" "%BUNDLE_TARGET%\Contents\Windows\%ACAD_YEAR%\" /y >nul

if exist "%BUNDLE_TARGET_COMMON%" rmdir /s /q "%BUNDLE_TARGET_COMMON%" 2>nul
mkdir "%BUNDLE_TARGET_COMMON%\Contents\Windows\%ACAD_YEAR%" 2>nul
if exist "%BUNDLE_TARGET_COMMON%\Contents\Windows\%ACAD_YEAR%" (
  xcopy "%BUNDLE_SOURCE%\PackageContents.xml" "%BUNDLE_TARGET_COMMON%\" /y >nul
  xcopy "%BUILD_OUTPUT%\AutoCAD_BoardSorter.dll" "%BUNDLE_TARGET_COMMON%\Contents\Windows\%ACAD_YEAR%\" /y >nul
  if exist "%BUILD_OUTPUT%\AutoCAD_BoardSorter.deps.json" xcopy "%BUILD_OUTPUT%\AutoCAD_BoardSorter.deps.json" "%BUNDLE_TARGET_COMMON%\Contents\Windows\%ACAD_YEAR%\" /y >nul
  if exist "%BUILD_OUTPUT%\AutoCAD_BoardSorter.pdb" xcopy "%BUILD_OUTPUT%\AutoCAD_BoardSorter.pdb" "%BUNDLE_TARGET_COMMON%\Contents\Windows\%ACAD_YEAR%\" /y >nul
)

echo Installed:
echo   %BUNDLE_TARGET%
if exist "%BUNDLE_TARGET_COMMON%\Contents\Windows\%ACAD_YEAR%" echo   %BUNDLE_TARGET_COMMON%
echo Restart AutoCAD and run BDSORT.
