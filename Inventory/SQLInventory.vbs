Option Explicit

Const HKLM = &H80000002

Const CUSTOM_NAMESPACE = 	"UHS"
Const DATA_CLASS = 			"SQLInstance"
Dim FOLDERS

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
	Dim wmiNamespace, dataClass
	Dim scriptPath
		
	scriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")

	Set argsNamed = WScript.Arguments.Named

	If GetNamespace(wmiNamespace) = True Then
		If GetDataClass(dataClass, wmiNamespace) Then
		
		If CSI_GetBitness("OS") = 64 Then
			GetSQLInfo 64, dataClass, wmiNamespace
		End If

		GetSQLInfo 32, dataClass, wmiNamespace
		
		End If
	End If
			
End Sub

Sub GetSQLInfo (ByRef architecture, ByRef dc, ByRef ns) 
	Dim sqlComponents, component
	Dim sqlInstances, instance
	Dim context, locator, regProvider

	Set context = CreateObject("WbemScripting.SWbemNamedValueSet")
	context.Add "__ProviderArchitecture", architecture

	Set locator = CreateObject("WbemScripting.SWbemLocator")
	Set regProvider = locator.ConnectServer("", "root\cimv2", "", "",,,, context).Get("StdRegProv")

	sqlComponents = EnumKeys (HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names", context, locator, regProvider)

	If Not IsNull(sqlComponents) Then
		
		For Each component in sqlComponents
		
			sqlInstances = EnumRegValues (HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\" & component, context, locator, regProvider)

			If Not IsNull(sqlInstances) Then
			
				For Each instance in SQLInstances
					UpdateDataObject dc, ns, ReadRegValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\" & component & "\", instance, "String", context, locator, regProvider), component, 64, context, locator, regProvider
				Next
			End If
		Next
	
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

Function UpdateDataObject(ByRef dc, ByRef ns, ByRef instanceName, ByRef componentType, ByRef architecture, ByRef context, ByRef locator, ByRef regProvider)
	Dim dataObject, objectPath
	
	Dim version, patchLevel, collation, tcpPort, edition
	
	version = ReadRegValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & instanceName & "\Setup", "Version", "String", context, locator, regProvider)
	patchLevel = ReadRegValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & instanceName & "\Setup", "PatchLevel", "String", context, locator, regProvider)
	edition = ReadRegValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & instanceName & "\Setup", "Edition", "String", context, locator, regProvider)
	
	If componentType = "SQL" Then
		collation = ReadRegValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & instanceName & "\Setup", "Collation", "String", context, locator, regProvider)
		tcpPort = ReadRegValue(HKLM, "SOFTWARE\Microsoft\Microsoft SQL Server\" & instanceName & "\MSSQLServer\SuperSocketNetLib\Tcp\IPAll", "TcpPort", "String", context, locator, regProvider)
	Else
		collation = ""
		tcpPort = ""	
	End If
	
	WScript.Echo " ++ " & instanceName & ": " & version & ", " & patchLevel & ", " & edition & ", " & collation & ", " & tcpPort
		
	objectPath = "\\.\root\" & CUSTOM_NAMESPACE & ":" & DATA_CLASS & ".InstanceName=""" & instanceName & """,InstanceType=""" & componentType & """"
				
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

	dataObject.InstanceName = instanceName
	dataObject.InstanceType = componentType
	dataObject.Version = version
	dataObject.PatchLevel = patchLevel
	dataObject.Collation = collation
	dataObject.TCPPort = tcpPort
	dataObject.Edition = edition
	dataObject.Architecture = CStr(architecture)
	On Error Resume Next
	Err.Clear

	dataObject.Put_()
	
	If Err.Number = 0 Then
		UpdateDataObject = True
	End If
	
	On Error Goto 0
	
	UpdateDataObject = False
		
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
			dc.Properties_.Add "InstanceName", wbemCimtypeString
			dc.Properties_.Add "InstanceType", wbemCimtypeString
			dc.Properties_.Add "Version", wbemCimtypeString
			dc.Properties_.Add "PatchLevel", wbemCimtypeString
			dc.Properties_.Add "Collation", wbemCimtypeString
			dc.Properties_.Add "TCPPort", wbemCimtypeString
			dc.Properties_.Add "Edition", wbemCimtypeString
			dc.Properties_.Add "Architecture", wbemCimtypeString
			dc.Properties_("InstanceName").Qualifiers_.add "key", True
			dc.Properties_("InstanceType").Qualifiers_.add "key", True
				
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

Function ReadRegValue(ByRef hive, ByRef key, ByRef valuename, ByRef valuetype, ByRef context, ByRef locator, ByRef regProvider)

	Dim inParams, outParams
	
	If valuetype = "DWORD" Then
		Set inParams = regProvider.Methods_("GetDWORDValue").InParameters
	Else
		Set inParams = regProvider.Methods_("GetStringValue").InParameters
	End If
	
	inParams.hDefKey = hive
	inParams.sSubKeyName = key
	inParams.sValueName = valuename
	
	If valuetype = "DWORD" Then
		Set outParams = regProvider.ExecMethod_("GetDWORDValue", inParams, , context)
		ReadRegValue = outParams.uValue
	Else
		Set outParams = regProvider.ExecMethod_("GetStringValue", inParams, , context)
		ReadRegValue = outParams.sValue
	End If
	
	
End Function

Function EnumRegValues(ByRef hive, ByRef key, ByRef context, ByRef locator, ByRef regProvider)

	Dim inParams, outParams
	Dim val
	
	Set inParams = regProvider.Methods_("EnumValues").InParameters
	
	inParams.hDefKey = hive
	inParams.sSubKeyName = key
	
	Set outParams = regProvider.ExecMethod_("EnumValues", inParams, , context)
	
	EnumRegValues = outParams.sNames
	
End Function

Function EnumKeys(ByRef hive, ByRef key, ByRef context, ByRef locator, ByRef regProvider)
	Dim inParams, outParams
	Dim val
	
	Set inParams = regProvider.Methods_("EnumKey").InParameters
	
	inParams.hDefKey = hive
	inParams.sSubKeyName = key
	
	Set outParams = regProvider.ExecMethod_("EnumKey", inParams, , context)
	
	EnumKeys = outParams.sNames

End Function

Function CSI_GetBitness(Target)
'CSI_GetBitness Function - updates at http://csi-windows.com.
' All bitness checks in one function

  Select Case Ucase(Target)
    
    Case "OS", "WINDOWS", "OPERATING SYSTEM"
      CSI_GetBitness = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth
    
    Case "HW", "HARDWARE"
      CSI_GetBitness = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").DataWidth
    
    Case "PROCESS", "PROC"
      'One liner to retrieve process architecture string which will reveal if running in 32bit subsystem
      ProcessArch = CreateObject("WScript.Shell").Environment("Process")("PROCESSOR_ARCHITECTURE")
      If lcase(ProcessArch) = "x86" Then
        CSI_GetBitness = 32
      Else
        If instr(1,"AMD64,IA64",ProcessArch,1) > 0 Then 
          CSI_GetBitness = 64
        Else
          CSI_GetBitness = 0 'unknown processor architecture
        End If
      End If
     Case Else
       CSI_GetBitness = 99999 'unknown Target item (OS, Hardware, Process)
  End Select

End Function
