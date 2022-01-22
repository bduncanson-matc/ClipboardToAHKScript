Option Explicit
'Declaring variables
Dim objHTML, objFile, objFSO, objWshShell, objOpenFile, strDeskPath, strAHKSetKey, strExecuteVBscript, 
strCombinedPath, strFileName, strAHKSend, strFullAHK, strClipText, strEndAHKSend dateLastMod
'ahk string to set alt + ctrl + v to this script
strAHKSetKey = "^!v::"
strAHKSetKey2 = "^!::"
strEndAHKSend = "return"
strFileName = "AHKCopy.ahk" strAHKSetKey,


'using html object to pase the clipboard text
Set objHTML = CreateObject("htmlfile")
strClipText = objHTML.ParentWindow.ClipboardData.GetData("text")
Set objHTML = Nothing

'create desktop path string
Set objWshShell = WScript.CreateObject("WScript.Shell")
strDeskPath = objWshShell.SpecialFolders("Desktop")
Set objWshShell = Nothing

'create the vbscript string variable for a ADK send hardware strokes 
strAHKSend = "Send " & strClipText
strExecuteVBScript = strAHKSetKey & vbNewLine & "Run " & strDesktopPath & "\Clipboard.vbs" & vbNewLine & strEndAHKSend & vbNewLine




'create path + file string
strCombinedPath = strDeskPath & "\" & strFileName

'create autohotkey full string
strFullAHK = strAHKSetKey & vbNewLine & strAHKSend & vbNewLine & strExecuteVBScript

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
