@echo off
rem techpartner launcher for the packaged Windows app: <install>\resources\cli\techpartner.cmd
setlocal
set ELECTRON_RUN_AS_NODE=1
"%~dp0..\..\TechPartner Office.exe" "%~dp0techpartner.cjs" %*
endlocal
