@echo on
for %%A IN (*.ps1) Do (SET myfile="%%~dpnxA")
powershell.exe  -ExecutionPolicy Bypass -Command "&{Start-Process Powershell -Argumentlist '-ExecutionPolicy Bypass -File ""%myfile%""' -Verb RunAs}"