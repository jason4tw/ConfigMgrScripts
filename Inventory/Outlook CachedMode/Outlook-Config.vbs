Option Explicit

Const CUSTOM_NAMESPACE = 	"UserLocal"
Const DATA_CLASS = 			"Outlook_Configuration"

Const wbemCimtypeString = 8 
Const wbemCimtypeUint32 = 19
Const wbemCimtypeDatetime = 101
Const wbemCimtypeReal32 = 4
Const wbemCimtypeBoolean = 11

const HKCU = &H80000001
const HKLM = &H80000002
const HKU  = &H80000003

Dim g_WshShell
Dim g_fso
Dim g_registry
Dim g_WshNetwork

Set g_WshShell = WScript.CreateObject("WScript.Shell")
Set g_WshNetwork = WScript.CreateObject("WScript.Network")
Set g_fso = CreateObject("Scripting.FileSystemObject")
Set g_registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv") 

Main

Sub Main
	Dim argsNamed
	Dim namespace, dataClass
	Dim scriptPath
	Dim outlookVersion
	
	scriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")

	Set argsNamed = WScript.Arguments.Named

	If GetNamespace(namespace) = True Then
		If GetDataClass(dataClass, namespace) Then
			
			outlookVersion = GetOutlookVersion
			If outlookVersion <> "" Then
				OutlookProfileInfo outlookVersion, dataClass, namespace
			End If
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
		GetNamespace = True
	End If
	
	On Error GoTo 0
	
End Function

Function UpdateDataObject(ByVal version, ByRef config, ByVal value, ByRef dc, ByRef ns)
	Dim dataObject, objectWmiPath, username
		
	UpdateDataObject = False
	
	username = g_WshNetwork.UserName
	
	objectWmiPath = "\\.\root\" & CUSTOM_NAMESPACE & ":" & DATA_CLASS & ".CachedModeConfig=""" & config & """,Username=""" & username & """"
				
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

	dataObject.Username = username
	dataObject.CachedModeEnabled = GetCachedExchangeMode(value)
	dataObject.CachedModeConfig = config
	dataObject.CachedModeConfigValue = ByteArrayToString(value)
		
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
	End If
	
	On Error GoTo 0

End Function

Function ByteArrayToString(ByVal bytes)
	Dim byteString, B, tmp
	
	byteString = ""
	
	If(VarType(bytes) = vbArray Or vbVariant) And (UBound(bytes) > 0) Then
		For Each B in bytes
			If(B < &H10) Then
				tmp = "0" & Hex(B)
			Else
				tmp = Hex(B)
			End If
			
			byteString = byteString & tmp & " "
		Next
	End If
	
	ByteArrayToString = byteString

End Function

Function GetOutlookVersion()
	Dim result, value

	Const curVer = "SOFTWARE\Classes\Outlook.Application\CurVer"

	GetOutlookVersion = ""
  
	result = g_registry.GetStringValue(HKLM, curVer, "", value)
	
	If result = 0  Then
		GetOutlookVersion = value
 	End If
	 
End Function

Function GetCachedExchangeMode(ByVal bytes)
  Dim result

  result = -2
  
  If (VarType(bytes) = vbArray Or vbVariant) Then
	If (UBound(bytes) > 0) Then
		If((bytes(1) And 1) <> 0) Then
			result = vbTrue
		Else
			result = vbFalse
		End If
	End If
  End If
  
  GetCachedExchangeMode = result
End Function

Sub OutlookProfileInfo(ByVal version, ByRef dc, ByRef ns)

	Dim result
	Dim defaultProfile, outlookConfigs, outlookConfig
	Dim value
	
	Const outlookProfilesKey = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
	Const cachedModeConfig = "00036601"
	
	result = g_registry.GetStringValue(HKCU, outlookProfilesKey, "DefaultProfile", defaultProfile)

	If result = 0 Then
		
		result = g_registry.EnumKey(HKCU, outlookProfilesKey & "\" & defaultProfile, outlookConfigs)
		
		If result = 0 and IsArray(outlookConfigs) Then
			For Each outlookConfig in outlookConfigs
		
				result = g_registry.GetBinaryValue(HKCU, outlookProfilesKey & "\" & defaultProfile & "\" & outlookConfig, cachedModeConfig, value)
				
				If result = 0 Then
					WScript.Echo ByteArrayToString(value)
					UpdateDataObject version, outlookConfig, value, dc, ns
				End If
			
			Next
		End If
	End If

End Sub
