'=========================================================================
' Description: Extension for SendPathToMail to write the file paths
' to a temporary file and afterwards read it out in the main process.
' This is needed for implementation in the windows context menu.
' 
' Author: Constantin Heinzler
' Version: 1.0.2
' License: MIT
'=========================================================================
' ACTUAL SCRIPT PROCESS
'=========================================================================

' get given path from arguments
' expect multiple paths if script is implemented in SendTo directory
Set paths = Wscript.Arguments

Dim fs    
Set fs = CreateObject("Scripting.FileSystemObject")

' static path to temporary cache file
Dim tempDir
tempDir = fs.GetSpecialFolder(2)
Dim tempPath
tempPath = tempDir & "\temp.sptm"
' open the file in append mode
set file = fs.OpenTextFile(tempPath, 8, True)

' loop through available paths and write to file
For Each path In paths
	' write current path to temp file
	file.WriteLine path
Next

' close temp file
file.close

' check if SendPathToMail is already running
' If the cookie file exists -> Script is running and nothing should be done
' Otherwise start the script
' static path to temporary cache file

Dim cookiePath
cookiePath = tempDir & "\cookie.sptm"

' start a main instance
Dim objShell
Set objShell = Wscript.CreateObject("WScript.Shell")

Dim currentDir
currentDir = fs.GetParentFolderName(WScript.ScriptFullName)
' Chr(34) being double quotes to escape the path
realScript = Chr(34) + currentDir + "\sendPathToMail.vbs" + Chr(34)

' use extra cscript call to enable hidden execution
Dim command
command = "cscript " + realScript

' 0 = run hidden, True = wait for task completion
objShell.Run command, 0, True

' clean up
Set objShell = Nothing

WScript.Quit()