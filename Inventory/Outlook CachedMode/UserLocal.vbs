Option Explicit

Const CUSTOM_NAMESPACE = 	"UserLocal"
Const DATA_CLASS = 			"Outlook_Configuration"
Dim FOLDERS

Const wbemCimtypeString = 8 
Const wbemCimtypeUint32 = 19
Const wbemCimtypeDatetime = 101
Const wbemCimtypeReal32 = 4
Const wbemCimtypeBoolean = 11

const HKEY_USERS = &H80000003

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
	Dim userProfilesPath, userProfilesFolder
	Dim regCommandLine, returnValue, profilePath
	Dim outlookProfilesKey, outlookProfiles, outlookProfile, outlookConfigs, outlookConfig
		
	scriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")

	Set argsNamed = WScript.Arguments.Named
	
	Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv") 
	
	'registry.EnumKey HKEY_USERS, outlookProfilesKey, outlookProfiles
	
	'For Each outlookProfile in outlookProfiles
	
	'	registry.EnumKey HKEY_USERS, outlookProfilesKey & "\" & outlookProfile, outlookConfigs
		
	'	For Each outlookConfig in outlookConfigs
		
	'	Next
	
	'Next
			
	If GetNamespace(namespace) = True Then
		If GetDataClass(dataClass, namespace) Then
			WScript.Echo "OK"

				'UpdateDataObject folderPath, GetFolderSize(folderPath), dataClass, namespace
	'		Next
		End If
	End If
			
End Sub

Function GetNamespace(ByRef localNamespace)
	Dim locator
	Dim rootNamespace, topNamespace
	Dim securityDescriptor, acl, result
		
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
	
	securityDescriptor = array(1, 0, 4, 129, 132, 0, 0, 0, 148, 0, 0, 0, 0, 0, 0, 0, 20, 0, 0, 0, 2, 0, 112, 0, 5, 0, 0, 0, 0, 0, 20, 0, 8, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 5, 11, 0, 0, 0, 0, 18, 24, 0, 63, 0, 6, 0, 1, 2, 0, 0, 0, 0, 0, 5, 32, 0, 0, 0, 32, 2, 0, 0, 0, 18, 20, 0, 19, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 5, 20, 0, 0, 0, 0, 18, 20, 0, 19, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 5, 19, 0, 0, 0, 0, 18, 20, 0, 19, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 5, 11, 0, 0, 0, 1, 2, 0, 0, 0, 0, 0, 5, 32, 0, 0, 0, 32, 2, 0, 0, 1, 2, 0, 0, 0, 0, 0, 5, 32, 0, 0, 0, 32, 2, 0, 0)
	set acl = localNamespace.get("__systemsecurity=@")
	result = acl.setsd(securityDescriptor)	
	
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
			dc.Properties_.Add "CachedModeEnabled", wbemCimtypeBoolean
			dc.Properties_.Add "CachedModeConfig", wbemCimtypeString
			dc.Properties_.Add "CachedModeConfigValue", wbemCimtypeString
			dc.Properties_("UserName").Qualifiers_.add "key", True
			dc.Properties_("CachedModeConfig").Qualifiers_.add "key", True
				
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

