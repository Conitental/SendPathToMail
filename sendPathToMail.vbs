'=========================================================================
' Description: VBScript to extend Windows "Send To" context menu to
' resolve the filepath of the selected file and send it using outlook
'
' Author: Constantin Heinzler
' Last Change: 01.12.2018
' License: MIT
'=========================================================================
' ACTUAL SCRIPT PROCESS
'=========================================================================
Set paths = Wscript.Arguments

' loop through all available arguments (->paths)
Dim driveLetter, realDrive, fullPath, mailBody
For Each path In paths
	' create a false loop to simulate modern "continue" statements in VBS
	Do
		driveLetter = stripPath(path)

		' if no drive letter can be retrieved the path is not valid
		If isEmpty(driveLetter) Then
			MsgBox "The following path cannot be validated:" + vbCrLf + path, vbOKOnly, "Path is not valid"
			Exit Do
		End If

		realDrive = getNetDrive(driveLetter)

		' continue if no net use drive could be found (for it is a local drive then)
		If isEmpty(realDrive) Then
			MsgBox "The following path cannot be attached to an email as it is stored on a local drive:" + vbCrLf + path + vbCrLf + "Place the file on a network share and try again.", vbOKOnly, "Detected local file"
			Exit Do
		End If
		fullPath = concatRealPath(realDrive, path)

		mailBody = mailBody + "<br>" + fullPath
	Loop While False
Next

' end script if no path could be retrieved
If isEmpty(mailBody) Then Wscript.Quit

' actually open email using connected links
openMail(mailBody)

'=========================================================================
' FUNCTIONS
'=========================================================================
Sub openMail(pathsToSend)

	Dim outobj, mailobj
	Dim objFileToRead

	' create outlook object
	Set outobj = CreateObject("Outlook.Application")
	Set mailobj = outobj.CreateItem(0)

	' attach parameters
	With mailobj
	.HTMLBody = pathsToSend
	.Display
	End With

	' clear the memory
	Set outobj = Nothing
	Set mailobj = Nothing

End Sub

Function concatRealPath(realDrive, rawPath)
	' strip two characters of the raw path ( e.g.: "C:\Windows\" --> "Windows\")
	Dim nakedPath
	nakedPath = Mid(RawPath, 3)

	' concat and add file:/// for the link to be clickable
	' use html to enable spaces in paths
	Dim fullPath
	fullPath = realDrive + nakedPath
	concatRealPath = "<a href=""file:///" + fullPath + """>" + fullPath + "</a>"
End Function

Function stripPath(path)
	' strip the first characterr of the given path and validate to be a drive assigned letter
	Dim char
	char = Left(path, 1)

	' regex to validate char
	Set re = New RegExp
	re.Pattern = "[a-z]"
	re.IgnoreCase = True
	re.Global = True
	isLetter = re.Test(char)

	' output the char if it is a letter and end function if not
	If isLetter = True Then
	    stripPath = char
	End If
End Function

Function getNetDrive(assignedLetter)
	' run command to get raw wireless output
	Dim netDrives
	netDrives = shellRun("NET USE")

	' regex net use output and find given letter
	Set re = New RegExp
	re.Pattern = assignedLetter + ":.*"
	re.IgnoreCase = True
	re.Global = True
	Set matches = re.Execute(netDrives)

	' get single found line
	Dim driveRaw
	Dim match
	For Each match in matches
	  driveRaw = match.value
	Next

	' match drive assignment until a space appears
	re.Pattern = "\\[^\s]+"
	Set matches = re.Execute(driveRaw)

	For Each match in matches
	  getNetDrive = match.value
	Next
End Function

Function shellRun(sCmd)
    'Run a shell command, returning the output as a string
    Dim oShell
    Set oShell = CreateObject("WScript.Shell")
    
    'run command
    Dim oExec
    Dim oOutput
    Set oExec = oShell.Exec(sCmd)
    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object
    Dim s
    Dim sLine
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend

    shellRun = s
End Function