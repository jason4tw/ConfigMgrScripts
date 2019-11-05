Option Explicit

Const MSG_MAIN_BEGIN =						"Beginning Execution at "
Const MSG_MAIN_FINISH =						"Finished Execution at "
Const MSG_DIVIDER =							"----------------------------------------"

Const DEFAULT_EVENTLOG_PREFIX =    			"ConfigMgr Site Assignment Script -- "

Const MSG_ASSIGNSITE_UNABLETORETRIEVE =		"Unable to retrieve current assigned ConfigMgr site: "
Const MSG_ASSIGNSITE_CURRENTDOMAIN =		"The current domain for this system is "
Const MSG_ASSIGNSITE_CURRENTSITE =			"The current ConfigMgr site for this system is "
'Const MSG_ASSIGNSITE_DESIREDSITE =			"The desired ConfigMgr site for this system based on domain is "
Const MSG_ASSIGNSITE_ADSITE =				"The current AD site for this system is "

'Const MSG_ASSIGNSITE_NOMATCH =				"Unable to match the AD site to a desired ConfigMgr site"
Const MSG_ASSIGNSITE_DESIREDSITEMATCH =		"The desired ConfigMgr site matches the current ConfigMgr site "
Const MSG_ASSIGNSITE_DESIREDSITENOMATCH =	" ! The desired ConfigMgr site does not match the current ConfigMgr site, reassigning client to "
Const MSG_ASSIGNSITE_ADSITE_FAIL =			"Unable to retrieve AD Site information: "

Const MSG_ASSIGNSITE_SUCCESS =				"Successfully reassigned the client to "
Const MSG_ASSIGNSITE_FAIL =					"Failed to reassign the client to "

Const MSG_NO_SITECODE_SPECIFIED =			"No site code to assign the client to was specified."

Dim g_fso, g_WshShell, g_logPathandName, g_startTime, g_wshNetwork
Dim g_namedArguments

Set g_fso = CreateObject ("Scripting.FileSystemObject")
Set g_WshShell = WScript.CreateObject("WScript.Shell")
Set g_wshNetwork = WScript.CreateObject("WScript.Network")
Set g_namedArguments = WScript.Arguments.Named

Main

Sub Main
	Dim client
	Dim errorCode

	g_startTime = Now
	
	On Error GoTo 0

	Set client = WScript.CreateObject ("Microsoft.SMS.Client")

	errorCode = Err.Number
	
	On Error Resume Next
	Err.Clear

	If errorCode = 0 Then
		Err.Clear
		
		WriteFirstLogMsg
		
		If g_namedArguments.Exists("SiteCode") Then
			AssignSite(client)
		Else
			WriteLogMsg MSG_NO_SITECODE_SPECIFIED, 1, 1, 1
		End If
		
		WriteFinalLogMsg	
	End If
	
End Sub

Sub WriteFirstLogMsg

	Dim configMgrLogPath
	Dim errorCode, errorMsg
	Dim logfile

	configMgrLogPath = GetCCMLogPath
	
	If configMgrLogPath = "" Then
		configMgrLogPath = g_WshShell.ExpandEnvironmentStrings("%WinDir%") & "\"
	Else
		configMgrLogPath = configMgrLogPath & "Logs\"
	End If

	g_logPathandName = configMgrLogPath & Replace(WScript.ScriptName, ".vbs", "", 1, 1, 1) & ".log"
		
	On Error Resume Next
	
	Set logfile = g_fso.OpenTextFile(g_logPathandName, 8, True)
	
	errorCode = Err.Number
	
	On Error GoTo 0
	
	If errorCode = 0 Then
		logfile.Close

		WriteLogMsg MSG_DIVIDER, 1, 1, 0
		WriteLogMsg MSG_MAIN_BEGIN & g_startTime, 1, 1, 1

	End If
				
End Sub

Sub WriteLogMsg(msg, msgtype, echomsg, eventlog)
	Dim outmsg, theTime, logfile
	
    theTime = Time
	
	outmsg = "<![LOG[" & msg & "]LOG]!><time="
	outmsg = outmsg & """" & DatePart("h", theTime) & ":" & DatePart("n", theTime) & ":" & DatePart("s", theTime) & ".000+0"""
	outmsg = outmsg & " date=""" & Replace(Date, "/", "-") & """"
	outmsg = outmsg & " component=""" & WScript.ScriptName & """ context="""" type=""" & msgtype & """ thread="""" file=""" & WScript.ScriptName & """>"

	On Error Resume Next

	Set logfile = g_fso.OpenTextFile(g_logPathandName, 8, True)
	logfile.WriteLine outmsg
	logfile.Close
	
	On Error GoTo 0
	
	'If echomsg = 1 Then
	'	WScript.Echo msg
	'End If
	
	If eventlog = 1 Then
		g_WshShell.LogEvent 0, DEFAULT_EVENTLOG_PREFIX & msg
	End If
	
End Sub

Sub WriteFinalLogMsg

	Dim registryLocation
	Dim logfile
	Dim logFileSize, maxLogFileSize
	Dim finishTime
	
	finishTime = Now
	
	WriteLogMsg MSG_MAIN_FINISH & finishTime, 1, 1, 0
	WriteLogMsg MSG_DIVIDER, 1, 0, 0
	
	g_WshShell.LogEvent 0, MSG_MAIN_FINISH & g_logPathandName
	
End Sub

Sub AssignSite (ByRef smsClient)

	Dim assignedSite, desiredSite
	Dim adInfo
	Dim currentDomain
	Dim errorCode, errorMsg

	On Error Resume Next
	Err.Clear
	
	assignedSite = smsClient.GetAssignedSite
	
	errorCode = Err.Number
	errorMsg = Err.Description & " (" & Err.Number & ")"
	
	On Error GoTo 0
	
	If errorCode <> 0 Then
		
		WriteLogMsg MSG_ASSIGNSITE_UNABLETORETRIEVE & errorMsg, 3, 1, 1
		
		Exit Sub
	End If
	
	WriteLogMsg MSG_ASSIGNSITE_CURRENTSITE & assignedSite, 1, 1, 1
	
	currentDomain = g_wshNetwork.userDomain
	
	WriteLogMsg MSG_ASSIGNSITE_CURRENTDOMAIN & currentDomain, 1, 1, 1
	
	On Error Resume Next
	Err.Clear
	
	Set adInfo = CreateObject("ADSystemInfo")
	
	errorCode = Err.Number
	msg = Err.Description
	
	On Error GoTo 0
	
	If errorCode <> 0 Then
		WriteLogMsg MSG_ASSIGNSITE_ADSITE_FAIL & msg, 3, 1, 1
		Exit Sub
	End If
	
	WriteLogMsg MSG_ASSIGNSITE_ADSITE & adInfo.SiteName, 1, 1, 1
	
	desiredSite = g_namedArguments.Item("SiteCode")

	If assignedSite = desiredSite Then
	
		WriteLogMsg MSG_ASSIGNSITE_DESIREDSITEMATCH & desiredSite, 1, 1, 1	
	
	Else
		WriteLogMsg MSG_ASSIGNSITE_DESIREDSITENOMATCH & desiredSite, 1, 1, 1
		
		On Error Resume Next
		Err.Clear
		
		smsClient.SetAssignedSite desiredSite, 3
		
		errorCode = Err.Number
		errorMsg = Err.Description & " (" & Err.Number & ")"
		On Error GoTo 0
		
		If errorCode <> 0 Then
			WriteLogMsg MSG_ASSIGNSITE_FAIL & desiredSite & ": " & errorMsg, 3, 1, 1
		
		Else
			WriteLogMsg MSG_ASSIGNSITE_SUCCESS & smsClient.GetAssignedSite, 1, 1, 1
		End If
	
	End If		

End Sub

Function GetCCMLogPath
	Dim namedValueSet, inParams, outParamsx
	Dim locator, services, regProvider
	
	Const HKLM = &h80000002
	
	GetCCMLogPath = ""

	On Error Resume Next
	Err.Clear
	
	Set namedValueSet = CreateObject("WbemScripting.SWbemNamedValueSet")

	If Err.Number <> 0 Then
		Exit Function
	End If
		
	namedValueSet.Add "__ProviderArchitecture", 64

	Set locator = CreateObject("Wbemscripting.SWbemLocator")

	If Err.Number <> 0 Then
		Exit Function
	End If

	Set services = locator.ConnectServer("","root\default","","",,,,namedValueSet)

	If Err.Number <> 0 Then
		Exit Function
	End If

	Set regProvider = services.Get("StdRegProv") 

	If Err.Number <> 0 Then
		Exit Function
	End If

	Set inParams = regProvider.Methods_("GetStringValue").Inparameters

	If Err.Number <> 0 Then
		Exit Function
	End If

	inParams.Hdefkey = HKLM
	inParams.Ssubkeyname = "SOFTWARE\Microsoft\SMS\Client\Configuration\Client Properties"
	inParams.Svaluename = "Local SMS Path"

	Set outParamsx = regProvider.ExecMethod_("GetStringValue", inParams,,namedValueSet)

	If Err.Number <> 0 Then
		Exit Function
	End If

	GetCCMLogPath = outParamsx.SValue

End Function
