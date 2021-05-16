@echo off
if "%1" == "" goto HELP_MSG
  powershell -NoLogo -NoExit -ExecutionPolicy RemoteSigned -File ./checkReportHelper.ps1 %1
  exit

:HELP_MSG
echo Usage : checkReport folder
echo    ex : checkReport c:/Reports"
