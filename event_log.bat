@echo on
powershell.exe  -ExecutionPolicy Bypass -Command "&{Start-Process Powershell -Argumentlist '-ExecutionPolicy Bypass -File ""%~dp0event_log.ps1""' -Verb RunAs}"