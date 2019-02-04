'=========================================================================
' Description: VBScript to extend Windows "Send To" context menu to
' resolve the filepath of the selected file and send it using outlook
'
' Author: Constantin Heinzler
' Version: 1.0.1
' License: MIT
'=========================================================================
' ACTUAL SCRIPT PROCESS
'=========================================================================

' create FileSystemObject
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")

' static path to temporary cache file
Dim tempDir
tempDir = fs.GetSpecialFolder(2)
Dim tempPath
tempPath = tempDir & "\temp.sptm"
Dim cookiePath
cookiePath = tempDir & "\cookie.sptm"

' create a temporary cookie file to show that the main script is running
set cookie = fs.OpenTextFile(cookiePath, 8, True)
' write something to the cookie
cookie.WriteLine "I'm alive!"
' close temp file
cookie.close

' wait a short time to give the handler scripts time to finish
WScript.Sleep(2500)

' check if the temp file has been created and quit if not
If Not fs.FileExists(tempPath) Then Wscript.Quit()

' open temp file
Set file = fs.OpenTextFile( tempPath, 1)

Dim paths()

' loop through lines in temp file and add to array
i = 0
Do While file.AtEndOfStream <> True
    line = file.ReadLine
    ReDim Preserve paths(i)
    paths(i) = line
    i=i+1
loop
' close file reading
file.Close

' recheck if the file is still existing and delete it
If fs.FileExists(tempPath) Then fs.DeleteFile tempPath

' declare variable to store paths that are local
Dim localPaths

' loop through all available arguments (->paths)
Dim driveLetter, realDrive, fullPath, mailBody
For Each path In paths
	' create a false loop to simulate modern "continue" statements in VBS
	Do
		driveLetter = isMappedDrive(path)

		' if no drive letter can be retrieved the drive is not mapped and will be transfered directly
		If driveLetter = "" Then
			fullPath = concatRealPath("REAL", path)
			mailBody = mailBody + "<br>" + fullPath
			Exit Do
		End If

		' resolve unc path from driveLetter
		realDrive = getNetDrive(driveLetter)

		' continue if no net use drive could be found (for it is a local drive then)
		If realDrive = "" Then
			' append local path to all local paths and add a new line
			localPaths = localPaths + path + vbCrLf
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

' finally deleting the cookie file to show that the main process is finished
If fs.FileExists(cookiePath) Then fs.DeleteFile cookiePath

' end script process
Wscript.Quit()

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
	' if "REAL" is given as realDrive the path is already valid and does not need to be processed
	If realDrive = "REAL" Then
		fullPath = rawPath
	Else
		' strip two characters of the raw path ( e.g.: "C:\Windows\" --> "Windows\")
		Dim nakedPath
		nakedPath = Mid(RawPath, 3)

		' concat and add file:/// for the link to be clickable
		' use html to enable spaces in paths
		Dim fullPath
		fullPath = realDrive + nakedPath
	End If

	concatRealPath = "<a href=""file:///" + fullPath + """>" + fullPath + "</a>"
End Function

Function isMappedDrive(path)
	isMappedDrive = fs.GetDriveName(path)
End Function

Function getNetDrive(drive)
	Dim letter, share
	letter = Left(drive, 1)
	Set share = fs.GetDrive(letter)
	getNetDrive = share.ShareName
End Function
