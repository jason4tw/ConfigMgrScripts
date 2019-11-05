' Version 1.1 - January 9 2018
'
' Version Info:
' 	- Updated the call to scanstate.exe to account for space in path names
'
' Jason Sandys
' ConfigMgrFTW!

Option Explicit

Const CUSTOM_NAMESPACE = 	"ITLocal"
Const DATA_CLASS = 			"USMT_Estimate"

Const wbemCimtypeString = 8 
Const wbemCimtypeUint32 = 19
Const wbemCimtypeDatetime = 101
Const ScanStateOptions = "/uel:90"

Dim g_WshShell
Dim g_fso

Set g_WshShell = WScript.CreateObject("WScript.Shell")
Set g_fso = CreateObject("Scripting.FileSystemObject")

Main

Sub Main
	Dim architecture
	Dim estimateFile
	Dim argsNamed
	Dim namespace
	Dim scriptPath
		
	architecture = "x86"
	estimateFile = "%temp%\usmt-estimate.xml"
	
	scriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "")

	estimateFile = g_WshShell.ExpandEnvironmentStrings(estimateFile)
		
	If g_WshShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%") = "%ProgramFiles(x86)%" Then
		architecture = "x86"
	Else
		architecture = "amd64"
	End If
		
	Set argsNamed = WScript.Arguments.Named

	RunScanState scriptPath & architecture, estimateFile, argsNamed
		
	If GetNamespace(namespace) = True Then
		CreateDataObject ParseOutput(estimateFile), namespace
	End If
			
End Sub

Sub RunScanState (ByVal subfolder, ByVal output, ByVal args)

	Dim scanstateArgs, commandLine
	Dim xmlArgs, xmlArg
		
	If args.Exists("xml") Then
				
		xmlArgs = Split(args("xml"), ",", -1, 1)		
		
		For Each xmlArg In xmlArgs
			scanStateArgs = scanstateArgs & "/i:""" & subfolder & "\" & xmlArg & """ "
		Next
		
	Else
		scanstateArgs = "/i:""" & subfolder & "\MigApp.xml"" /i:""" & subfolder & "\MigDocs.xml"" "
	End If
	
	scanstateArgs = scanstateArgs & "/nocompress"
	commandLine = """" & subfolder & "\scanstate.exe"" ""%temp%\usmtestimate"" /l:""%temp%\scanstate.log"" " & ScanStateOptions & " /p:""" & output & """ " & scanstateArgs
		
	WScript.Echo commandLine
	
	g_WshShell.Run commandLine, 0, True

End Sub

Function ParseOutput(ByRef outputFilename)
		
	Dim output
	Dim xpath, nodes, i
		
	ParseOutput = -1
		
	' Check to make sure the output file exists
	If Not g_fso.FileExists(outputFilename) Then
		Exit Function
	End If
		
	Set output = CreateObject("Msxml2.DOMDocument")
		
	' Load the whole XML config file at once
	output.async = False
	output.load(outputFilename)
	
	' Check the file to make sure it is valid XML
	If output.parseError.errorCode <> 0 Then
		Exit Function
	Else
		' Set our XML query language to XPath
		output.setProperty "SelectionLanguage", "XPath"
	End If
			
	xpath = "/PreMigration/storeSize/size[1]"
	
	Set nodes = output.SelectNodes(xpath)
	
	For i = 0 To nodes.Length - 1
		ParseOutput = nodes.Item(i).Text
	Next
			
End Function

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

Function CreateDataObject(ByVal bytes, ByRef localNamespace)
	Dim dataClass, dataObject
	Dim dt, typeLib
		
	CreateDataObject = False
	
	CreateDataClass(localNamespace)
		
	Set dt = CreateObject("WbemScripting.SWbemDateTime")
	dt.SetVarDate(Now)
		
	Set typeLib = CreateObject("Scriptlet.TypeLib")
			
	On Error Resume Next
		
	Err.Clear
	
	Set dataClass = localNamespace.Get(DATA_CLASS)
		
	Set dataObject = dataClass.SpawnInstance_
		
	dataObject.SizeEstimate = CDbl(bytes / 1024 / 1024)
	dataObject.DateTime = dt
	dataObject.GUID = typeLib.Guid
		
	dataObject.Put_()
		
	On Error Goto 0
		
End Function

Sub CreateDataClass(ByRef localNamespace)
	Dim dataClass
		
	On Error Resume Next
	
	Set dataClass = localNamespace.Get()
		
	dataClass.Path_.Class = DATA_CLASS
	dataClass.Properties_.Add "SizeEstimate", wbemCimtypeUint32
	dataClass.Properties_.Add "DateTime", wbemCimtypeDatetime
	dataClass.Properties_.Add "GUID", wbemCimtypeString
	dataClass.Properties_("GUID").Qualifiers_.add "key", True
		
	dataClass.Put_()
		
	On Error GoTo 0
	
End Sub
