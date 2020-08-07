Option Explicit

Const CUSTOM_NAMESPACE = 	"ITLocal"
Const DATA_CLASS = 			"Outlook_Configuration"
Dim FOLDERS

Const wbemCimtypeString = 8 
Const wbemCimtypeUint32 = 19
Const wbemCimtypeDatetime = 101
Const wbemCimtypeReal32 = 4
Const wbemCimtypeBoolean = 11

const HKEY_USERS = &H80000003
const HKEY-LOCAL_MACHINE = &H80000002

Dim g_WshShell
Dim g_fso

Set g_WshShell = WScript.CreateObject("WScript.Shell")
Set g_fso = CreateObject("Scripting.FileSystemObject")

Main

Sub Main
	Dim argsNamed
	Dim namespace, dataClass
	Dim registry
	Dim scriptPath
	Dim userProfilesPath, userProfilesFolder, profileFolder
	Dim regCommandLine, returnValue, profilePath
	Dim outlookProfilesKey, outlookProfiles, outlookProfile, outlookConfigs, outlookConfig
		
	scriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")

	Set argsNamed = WScript.Arguments.Named

	Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv") 
	
	returnValue = registry.EnumKey(HKEY_USERS, outlookProfilesKey, outlookProfiles)
	WScript.Echo "****" & returnValue
	
	If returnValue = 0 Then
		
		For Each outlookProfile in outlookProfiles
		
			registry.EnumKey HKEY_USERS, outlookProfilesKey & "\" & outlookProfile, outlookConfigs
			WScript.Echo outlookProfile
			
			For Each outlookConfig in outlookConfigs
				WScript.Echo "..." & outlookConfigs
			
			Next
		
		Next
	End If
				
	'If GetNamespace(namespace) = True Then
	'	If GetDataClass(dataClass, namespace) Then
			

				'UpdateDataObject folderPath, GetFolderSize(folderPath), dataClass, namespace
	'		Next
	'	End If
	'End If
			
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
			dc.Properties_.Add "UserName", wbemCimtypeString
			dc.Properties_.Add "OutlookProfileName", wbemCimtypeString
			dc.Properties_.Add "CachedModeEnabled", wbemCimtypeBoolean
			dc.Properties_("UserName").Qualifiers_.add "key", True
			dc.Properties_("OutlookProfileName").Qualifiers_.add "key", True
				
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

