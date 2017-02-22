Dim windowsShell

Set windowsShell = CreateObject("WScript.Shell")
curFolder = windowsShell.CurrentDirectory

ListSubfolders curFolder

Sub ListSubfolders(currentFolder)
	Dim FSO
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Dim folder
	Set folder = FSO.GetFolder(currentFolder)
	Dim subfolders
	Set subfolders = folder.Subfolders

	Dim subfolder
	Dim string
	For Each subfolder in subfolders
		string = string & subfolder.name
		string = string & vbCrLf
		' ListFilesInFolder subfolder
		MoveFilesToParent subfolder
	Next
	' MsgBox string
End Sub

Sub ListFilesInFolder (currentFolder)
	Dim files
	Set files = currentFolder.Files
	
	Dim file
	Dim string
	For Each file in files
		string = string & file.name
		string = string & vbCrLf
	Next
	MsgBox string	
End Sub

Sub MoveFilesToParent (currentFolder)
	Dim files
	Set files = currentFolder.Files
	Dim FSO
	Set FSO = CreateObject("Scripting.FileSystemObject")
	
	Const FOF_CREATEPROGRESSDLG = &H0&
	targetFolder = FSO.GetParentFolderName(currentFolder)
	Set objShell = CreateObject("Shell.Application")
	Set objFolder = objShell.NameSpace(targetFolder)
	'objFolder.MoveHere "Z:\Videos\Anime\Dragon Ball\Season 01\*.*", FOF_CREATEPROGRESSDLG
	objFolder.MoveHere currentFolder + "\*.*", FOF_CREATEPROGRESSDLG
	
	' Dim file
	' For Each file in files
		' If FSO.GetExtensionName(file) = "mkv" Then
			' FSO.MoveFile currentFolder + "\*.mkv", FSO.GetParentFolderName(currentFolder)
		' End If
	' Next
End Sub



'Sub MoveFilesToParent (currentFolder)
	'Dim files
	'Set files = currentFolder.Files
	'Dim FSO
	'Set FSO = CreateObject("Scripting.FileSystemObject")
	'Dim objFileCopy
	'Dim strFileName
	
	'Dim file
	'For Each file in files	
		'If FSO.GetExtensionName(file) = "mkv" Then
			' Move to the file destination
			'If FSO.FileExists(FSO.GetParentFolderName(currentFolder) & "\" & strFileName) Then
				'WScript.Echo "File (" & strFileName & ") already exists on destination folder (" & FSO.GetParentFolderName(currentFolder) & ")!"
			'Else
				'FSO.MoveFile objFileCopy.Path, FSO.GetParentFolderName(currentFolder)
				'WSCript.Echo "File (" & strFileName & ") was moved to " & FSO.GetParentFolderName(currentFolder) & "!"
			'End If
		'End If
'End Sub


'Set currentDirectory = filesys.GetParentFolderName(WScript.ScriptFullName)
'Set currentDirectory = filesys.GetFolder("& currentDirectory &")


Set windowsShell = Nothing
WScript.Quit