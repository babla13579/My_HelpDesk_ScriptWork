@echo off
cls
if not exist C:\Users\%username%\Desktop\Scripts\ mkdir C:\Users\%username%\Desktop\Scripts\
robocopy M:\Shared_SMB\Folder_Path\ C:\Users\%username%\Desktop\Scripts\ /E /MT:127 /s /purge
if exist "C:\Users\%username%\Desktop\Scripts\Update Scripts.bat" ( copy /y "M:\Shared_SMB\Folder_Path\Update Scripts.bat" "C:\Users\%username%\Desktop\Scripts\Update Scripts.bat" )
:: pause
