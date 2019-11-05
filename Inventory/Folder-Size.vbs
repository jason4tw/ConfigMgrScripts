Option Explicit

Const CUSTOM_NAMESPACE = 	"ITLocal"
Const DATA_CLASS = 			"Folder_Size"
Dim FOLDERS

FOLDERS = array ("C:\Data","C:\Intel","C:\Drivers","C:\Dell","C:\Dummy","%Temp%","C:\IT")

Const wbemCimtypeString = 8 
Const wbemCimtypeUint32 = 19
Const wbemCimtypeDatetime = 101
Const wbemCimtypeReal32 = 4

Dim g_WshShell
Dim g_fso

Set g_WshShell = WScript.CreateObject("WScript.Shell")
Set g_fso = CreateObject("Scripting.FileSystemObject")

Main

Sub Main
	Dim argsNamed
	Dim namespace, dataClass
	Dim scriptPath
	Dim folderPath
		
	scriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")

	Set argsNamed = WScript.Arguments.Named

	If GetNamespace(namespace) = True Then
		If GetDataClass(dataClass, namespace) Then
			For Each folderPath in FOLDERS
				'WScript.Echo "- " & folderPath & " : " & GetFolderSize(folderPath)
				UpdateDataObject folderPath, GetFolderSize(folderPath), dataClass, namespace
			Next
		End If
	End If
			
End Sub

Function GetNamespace(ByRef localNamespace)
	Dim locator
	Dim rootNamespace, topNamespace
		
	GetNamespace = False

	On Error Resume Next
	Err.Clear

	Set locator = CreateObject("WbemScripting.SWbemLocator")
		
	If Err.Number <> 0 Then
		On Error GoTo 0
		Exit Function	
	End If
	
	Err.Clear
	
	Set localNamespace = locator.ConnectServer(".", "root\" & CUSTOM_NAMESPACE)
		
	If Err.Number = 0 Then
		On Error GoTo 0
		GetNamespace = True
		Exit Function
	End If
	
	Err.Clear
	Set localNamespace = Nothing

	Set rootNamespace = locator.ConnectServer(".", "root")
		
	If Err.Number <> 0 Then
		On Error GoTo 0
		Exit Function	
	End If

	Err.Clear

	Set topNamespace = rootNamespace.Get("__namespace")
		
	If Err.Number <> 0 Then
		On Error GoTo 0
		Exit Function	
	End If
		
	Err.Clear

	Set localNamespace = topNamespace.SpawnInstance_ 
	
	If Err.Number <> 0 Then
		On Error GoTo 0
		Exit Function	
	End If
		
	Err.Clear

	localNamespace.name = CUSTOM_NAMESPACE
	localNamespace.Put_()
		
	If Err.Number <> 0 Then
		On Error GoTo 0
		Exit Function	
	End If
		
	Err.Clear

	Set localNamespace = locator.ConnectServer(".", "root\" & CUSTOM_NAMESPACE)
	
	If Err.Number <> 0 Then
		On Error GoTo 0
		Exit Function	
	End If

	On Error GoTo 0
	GetNamespace = True
	
End Function

Function UpdateDataObject(ByVal path, ByVal mb, ByRef dc, ByRef ns)
	Dim dataObject, objectPath, objectWmiPath
		
	UpdateDataObject = False
	
	If mb = -1 Then
		Exit Function
	End If
	
	objectWmiPath = Replace(path, "\", "\\")
	
	objectPath = "\\.\root\" & CUSTOM_NAMESPACE & ":" & DATA_CLASS & ".Path=""" & objectWmiPath & """"
				
	On Error Resume Next
	Err.Clear
	
	Set dataObject = ns.Get(objectPath)
	
	If Err.Number <> 0 Then
		Err.Clear

		Set dataObject = dc.SpawnInstance_
		
		If Err.Number <> 0 Then
			On Error GoTo 0
			Exit Function
		End If
		
	End If
	
	On Error GoTo 0

	dataObject.Size = mb
	dataObject.Path = path
		
	On Error Resume Next
	Err.Clear

	dataObject.Put_()
	
	If Err.Number = 0 Then
		UpdateDataObject = True
	End If
	
	On Error Goto 0
		
End Function

Function GetDataClass(ByRef dc, ByRef ns)
	
	GetDataClass = False
	
	On Error Resume Next
	Err.Clear
	
	Set dc = ns.Get(DATA_CLASS)
	
	If err.Number = 0 Then
		GetDataClass = True
	Else
		Err.Clear

		Set dc = ns.Get()

		If err.Number <> 0 Then
			GetDataClass = False
		Else
			On Error GoTo 0

			dc.Path_.Class = DATA_CLASS
			dc.Properties_.Add "Size", wbemCimtypeReal32
			dc.Properties_.Add "Path", wbemCimtypeString
			dc.Properties_("Path").Qualifiers_.add "key", True
				
			On Error Resume Next
			Err.Clear
			
			dc.Put_()
			
			If err.Number <> 0 Then
				GetDataClass = False
			Else
				Set dc = ns.Get(DATA_CLASS)
				GetDataClass = True
			End If
		End If
	End If
	
	On Error GoTo 0

End Function

Function GetFolderSize(ByVal path)

	Dim folder
	
	path = g_WshShell.ExpandEnvironmentStrings(path)
	
	On Error Resume Next
	Err.Clear
	
	Set folder = g_fso.GetFolder(path)
	
	If Err.Number = 0 Then
		GetFolderSize = Round(CDbl(folder.Size / 1024 / 1024), 4)
	Else	
		GetFolderSize = -1
	End If
	
	On Error Goto 0
	

End Function
