' Launch Winget Updater in background mode
' This script runs the Python application without showing a console window

' Get the script's directory
Set objFSO = CreateObject("Scripting.FileSystemObject")
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Create Shell object
Set objShell = CreateObject("WScript.Shell")

' Change to the script directory
objShell.CurrentDirectory = strScriptPath

' Explicitly set the LOCALAPPDATA environment variable
localAppData = objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
Set envVars = objShell.Environment("PROCESS")
envVars("LOCALAPPDATA") = localAppData

' Python executable path - using default Python from PATH
strPythonExe = "python"

' Command to run
strCommand = """" & strPythonExe & """ """ & strScriptPath & "\main.py"" --standalone"

' Run the command with window style 0 (hidden)
' 0 = Hidden window
' True = Don't wait for program to finish
objShell.Run strCommand, 0, False

' Clean up
Set objShell = Nothing
Set objFSO = Nothing

