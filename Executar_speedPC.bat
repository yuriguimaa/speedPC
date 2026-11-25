@echo off
rem speedPC - abre a ferramenta (pede administrador automaticamente)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0speedPC.ps1" %*
