@echo off
REM Start EZFIO.PS1 in an elevated PowerShell interpreter.
REM Here be dragons.
REM First start a standard powershell and use it's Start-Process cmdlet
REM to start *another* powershell, this one as administrator, to interpret
REM the script.  Care must be taken to properly quote the path to the script.

set GO='%cd%\ezfio.ps1'
powershell -Command "$p = new-object System.Diagnostics.ProcessStartInfo 'PowerShell'; $p.Arguments = {-WindowStyle hidden -Command ". %GO%"}; $p.Verb = 'RunAs'; [System.Diagnostics.Process]::Start($p) | out-null;"
