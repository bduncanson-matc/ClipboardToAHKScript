Option Explicit
'Declaring variables
Dim objHTML, objFile, objFSO, objWshShell, objOpenFile, strDeskPath, strAHKSetKey, strAHKSetKey2, strExecuteVBScript
Dim strCombinedPath, strFileName, strSendAHK, strFullAHK, strClipText, strAHKSend, strEndAHK, dateLastMod

'ahk string to set alt + ctrl + v to this script
strAHKSetKey = "^!v::"
strAHKSetKey2 = "^!c::"
strEndAHK = "return" & vbNewLine
strSendAHK = "send "
strFileName = "\AHKCopy.ahk"


'using html object to pase the clipboard text
Set objHTML = CreateObject("htmlfile")
strClipText = objHTML.ParentWindow.ClipboardData.GetData("text")
Set objHTML = Nothing

'create desktop path string
Set objWshShell = WScript.CreateObject("WScript.Shell")
strDeskPath = objWshShell.SpecialFolders("Desktop")
Set objWshShell = Nothing

'create the vbscript string variable for a ADK send hardware strokes 
strAHKSend = strSendAHK & strClipText & vbNewLine & strEndAHK 
strExecuteVBScript = strAHKSetKey2 & vbNewLine & "Run " & strDeskPath & "\Clipboard.vbs" & vbNewLine & strEndAHK & vbNewLine




'create path + file string
strCombinedPath = strDeskPath & "\" & strFileName

'create autohotkey full string
strFullAHK = strAHKSetKey & vbNewLine & strAHKSend & strExecuteVBScript

'check for AHK file delete if exists
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(strCombinedPath) Then
	objFSO.DeleteFile(strCombinedPath)
End If

If Not objFSO.FileExists(strCombinedPath) Then 
	objFSO.CreateTextFile(strCombinedPath)
End If

Set objFile = objFSO.GetFile(strCombinedPath)
dateLastMod = objFile.DateLastModified

If DateAdd("s", 2, dateLastMod) > Now Then ' if created in the last two seconds
   Set objOpenFile = objFSO.OpenTextFile(strCombinedPath, 2) 'for writing
   objOpenFile.WriteLine(strFullAHK) 'write the string for sending a string
   objOpenFile.Close
End If

Set objOpenFile = Nothing
Set objFile = Nothing
Set objFSO = Nothing
