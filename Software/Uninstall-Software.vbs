' Uninstall-Software Script
' Version 1.10
' Jason Sandys
' http://home.configmgrftw.com
'
' 1.0:		Orignal Release
' 1.01:		Fixed bug in calling WriteLogMsg on line 155
' 1.02:		Forgot the stupid comma on line 155
' 1.10:		Added exclude string option
'			Added some additional error handling for the command-line execution
'			Added option to skip command-line FixupUninstallString
'			Added option to add parenthesis top the command line

Option Explicit

Const MSG_MAIN_BEGIN =						"Beginning Execution at "
Const MSG_MAIN_FINISH =						"Finished Execution at "
Const MSG_ELAPSED_TIME = 					"Total script execution time is "
Const MSG_DIVIDER =							"-------------------------------------------------------------------"

Const MSG_NO_PRODUCT_STRING_SPECIFIED =		"No product string specified on the command line."
Const MSG_PRODUCT_STRING =					"Searching for products with the following string in their display name: "
Const MSG_EXCLUDE_STRINGS =					"Excluding products with the following strings in their display name: "
Const MSG_64BIT_SEARCH =					"Searching for 64-bit applications..."
Const MSG_32BIT_SEARCH =					"Searching for 32-bit applications..."
Const MSG_HWINV =							"Initiating hardware inventory."
Const MSG_FOUND_PRODUCT =					"  + Found product with display name: "
Const MSG_PRODUCT_UNINSTALLSTRING =			"    - Uninstall string: "
Const MSG_PRODUCT_UNINSTALLSTRINGMOD =		"    - Modified uninstall string: "
Const MSG_RUNNING_UNINSTALL =				"    - Running uninstall"
Const MSG_UNINSTALL_COMPLETE =				"    = Uninstall completed with return code: "
Const MSG_TOTAL_PRODUCT_COUNT =				"Toal number of products matching criteria: "
Const MSG_PRODUCT_COUNT32 =					"Number of 32-bit products matching criteria: "
Const MSG_PRODUCT_COUNT64 =					"Number of 64-bit products matching criteria: "

Const HKEY_LOCAL_MACHINE = &H80000002

Dim g_Named: Set g_Named = WScript.Arguments.Named
Dim g_WshShell : Set g_WshShell = CreateObject("WScript.Shell")
Dim g_fso : Set g_fso = CreateObject ("Scripting.FileSystemObject")

Dim g_logPath : g_logPath = ""
Dim g_echoLog : g_echoLog = False
Dim g_configMgrAgentPresent : g_configMgrAgentPresent = False
Dim g_startTime : g_startTime = Now

Dim productCount32 : productCount32 = 0
Dim productCount64 : productCount64 = 0

Dim productSearchString
Dim productExcludeStrings
Dim doUninstall : doUninstall = False
Dim doHWInv : doHWInv = False
Dim dontModify : dontModify = False
Dim addParens : addParens = False

WriteFirstLogMsg

' If running as a compliance item comment out the following 13 lines
If g_Named.Exists("Product") Then
	productSearchString = g_Named.Item("Product")
Else
	productSearchString = ""
End If

If g_Named.Exists("Exclude") Then
	productExcludeStrings = split(g_Named.Item("Exclude"), ",")
End If

doUninstall = g_Named.Exists("Uninstall")
g_echoLog = Not(g_Named.Exists("Uninstall"))
doHWInv = g_Named.Exists("hwinv")
dontModify = g_named.Exists("DoNotMod")
addParens = g_named.Exists("AddParens")

' If running as a compliance item uncomment the following line (but do not modify)
' g_echoLog = False
'
' If running as a compliance item uncomment the following three lines and modify as needed
' productSearchString = ""
' doUninstall = True for remediation script, False for detection script 
' doHWInv = True or False for remediation script, False for detection script

If productSearchString <> "" Then

	WriteLogMsg MSG_PRODUCT_STRING & productSearchString, 1, g_echoLog

	If Not IsEmpty(productExcludeStrings) Then
		WriteLogMsg MSG_EXCLUDE_STRINGS & Join(productExcludeStrings), 1, g_echoLog
	End If

	If CSI_GetBitness("OS") = 64 Then
		WriteLogMsg MSG_64BIT_SEARCH, 1, g_echoLog	
		productCount64 = UninstallProducts(64, productSearchString, productExcludeStrings, dontModify, addParens, doUninstall)
	End If

	WriteLogMsg MSG_32BIT_SEARCH, 1, g_echoLog	
	productCount32 = UninstallProducts(32, productSearchString, productExcludeStrings, dontModify, addParens, doUninstall)
	
	WriteLogMsg MSG_PRODUCT_COUNT64 & productCount64, 1, g_echoLog	
	WriteLogMsg MSG_PRODUCT_COUNT32 & productCount32, 1, g_echoLog	
	WriteLogMsg MSG_TOTAL_PRODUCT_COUNT & productCount64 + productCount32, 1, g_echoLog	

	If g_configMgrAgentPresent And doHWInv Then
		WriteLogMsg MSG_HWINV, 1, g_echoLog	
		InitiateHardwareInventory
	End If
	
Else
	WriteLogMsg MSG_NO_PRODUCT_STRING_SPECIFIED, 2, g_echoLog	

End If

WScript.Echo productCount64 + productCount32

WriteFinalLogMsg

 Sub InitiateHardwareInventory()  
      On Error Resume Next

      Dim oCPAppletMgr  : Set oCPAppletMgr  = CreateObject("CPApplet.CPAppletMgr")  
      Dim oClientAction : Set oClientAction = Nothing  
      Dim oClientActions : Set oClientActions = oCPAppletMgr.GetClientActions()  
      
	  For Each oClientAction In oClientActions  
           If oClientAction.Name = "Hardware Inventory Collection Cycle" Then  
                oClientAction.PerformAction  
           End If  
      Next  
 End Sub  

Function UninstallProducts(ByRef architecture, ByRef sProductName, ByRef sExcludeStrings, ByRef donotModify, ByRef addParenthesis, ByRef performUninstall)
	
	Dim oContext, oLocator, oRegProvider
	Dim uninstallKeys
	Dim sPath, sKey, sDisplayName, sUninstallString
	Dim returnCode, returnType
	Dim count : count = 0
	Dim isExcluded : isExcluded = False
	Dim exclusion
	
	Set oContext = CreateObject("WbemScripting.SWbemNamedValueSet")
	oContext.Add "__ProviderArchitecture", architecture

	Set oLocator = CreateObject("WbemScripting.SWbemLocator")
	Set oRegProvider = oLocator.ConnectServer("", "root\cimv2", "", "",,,, oContext).Get("StdRegProv") 

	sPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	
	uninstallKeys = EnumKeys(HKEY_LOCAL_MACHINE, sPath, oContext, oLocator, oRegProvider)
	
	If Not IsNull(uninstallKeys) Then
	
		For Each sKey In uninstallKeys
		
			sDisplayName = ReadRegValue(HKEY_LOCAL_MACHINE, sPath & "\" & sKey, "DisplayName", "String", oContext, oLocator, oRegProvider)
			
			If InStr(1, sDisplayName, sProductName, 1) Then

				isExcluded = False

				For Each exclusion in sExcludeStrings
					If InStr(1, sDisplayName, exclusion, 1) Then
						isExcluded = True
					End If
				Next

				If Not isExcluded Then
					WriteLogMsg MSG_FOUND_PRODUCT & sDisplayName, 1, g_echoLog	

					sUninstallString = ReadRegValue(HKEY_LOCAL_MACHINE, sPath & "\" & sKey, "UninstallString", "String", oContext, oLocator, oRegProvider)
					
					WriteLogMsg MSG_PRODUCT_UNINSTALLSTRING & sUninstallString, 1, g_echoLog	

					If addParenthesis Then
						sUninstallString = """" & sUninstallString & """"
					End If
					
					If Not donotModify Then
						FixupUninstallString sUninstallString
					End If

					WriteLogMsg MSG_PRODUCT_UNINSTALLSTRINGMOD & sUninstallString, 1, g_echoLog	

					If performUninstall = true Then
						WriteLogMsg MSG_RUNNING_UNINSTALL, 1, g_echoLog
						
						On Error Resume Next
						Err.Clear
						returnCode = g_WshShell.Run (sUninstallString, 0, true)

						If Err.Number <> 0 Then
							returnCode = Err.Number
						End If

						On Error Goto 0

						If returnCode = 0 Then
							returnType = 0
							count = count + 1
						ElseIf returnCode <> 0 Then
							returnType = 2
						End If

						WriteLogMsg MSG_UNINSTALL_COMPLETE & returnCode, returnType, g_echoLog

					End If
				End If
			End If
		Next
	End If

	UninstallProducts = count

End Function

Sub FixupUninstallString(ByRef uninstallString)

	If InStr(1, uninstallString, "/I", 1) Then
		uninstallString = Replace(uninstallString, "/I", "/X", 1, 1, 1)
	End If
	
	If InStr(1, uninstallString, "/q", 1) = 0 Then
		uninstallString = uninstallString & " /q"
	End If
	
	If InStr(1, uninstallString, "/NoRestart", 1) = 0 Then
		uninstallString = uninstallString & " /NoRestart"
	End If

End Sub

Function CSI_GetBitness(Target)

	Dim ProcessorArch

	Select Case Ucase(Target)
	
		Case "OS", "WINDOWS", "OPERATING SYSTEM"
			CSI_GetBitness = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth
			
		Case "HW", "HARDWARE"
			CSI_GetBitness = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").DataWidth
			
		Case "PROCESS", "PROC"
		
			ProcessorArch = CreateObject("WScript.Shell").Environment("Process")("PROCESSOR_ARCHITECTURE")
			
			If lcase(ProcessArch) = "x86" Then
				CSI_GetBitness = 32
			Else
				If instr(1, "AMD64,IA64", ProcessArch, 1) > 0 Then
					CSI_GetBitness = 64
				Else
					CSI_GetBitness = 0
				End If
			End If
			
		Case Else
			CSI_GetBitness = 99999
	End Select
	
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

Sub WriteLogMsg(msg, msgtype, echomsg)
	Dim outmsg, theTime, logfile
	
	If g_logPath <> "" Then

		theTime = Time
		
		outmsg = "<![LOG[" & msg & "]LOG]!><time="
		outmsg = outmsg & """" & DatePart("h", theTime) & ":" & DatePart("n", theTime) & ":" & DatePart("s", theTime) & ".000+0"""
		outmsg = outmsg & " date=""" & Replace(Date, "/", "-") & """"
		outmsg = outmsg & " component=""" & WScript.ScriptName & """ context="""" type=""" & msgtype & """ thread="""" file=""" & WScript.ScriptName & """>"

		On Error Resume Next

		Set logfile = g_fso.OpenTextFile(g_logPath, 8, True)
		logfile.WriteLine outmsg
		logfile.Close

		If echomsg = 1 or echomsg = True Then
			WScript.Echo msg
		End If
		
		On Error Goto 0
		
	End If	
End Sub

Sub WriteFinalLogMsg

	Dim finishTime, totalSeconds, elapsedHours, elapsedMinutes, elapsedSeconds
	
	finishTime = Now

	totalSeconds = DateDiff ("s", g_startTime, finishTime)
	
	elapsedHours = CInt(totalSeconds / 3600)
	totalSeconds = totalSeconds Mod 3600

	elapsedMinutes = CInt(totalSeconds / 60)
	totalSeconds = totalSeconds Mod 60

	elapsedSeconds = totalSeconds
	
	WriteLogMsg MSG_MAIN_FINISH & finishTime, 1, g_echoLog
	WriteLogMsg MSG_ELAPSED_TIME & ZeroPadNumber(elapsedHours) & ":" & ZeroPadNumber(elapsedMinutes) & ":" & ZeroPadNumber(elapsedSeconds), 1, g_echoLog
	WriteLogMsg MSG_DIVIDER, 1, g_echoLog
	
End Sub

Function ZeroPadNumber (ByRef number)

	Dim numberString

	If number < 10 Then
		numberString = CStr("0" & number)
	Else
		numberString = CStr(number)
	End If

	ZeroPadNumber = numberString

End Function

Sub WriteFirstLogMsg

	Dim logfile
	Dim userEnv
	Dim errorCode

	On Error Resume Next
	Err.Clear
	
	If CSI_GetBitness("OS") = 64 Then
		g_logPath = g_WshShell.RegRead("HKLM\SOFTWARE\Microsoft\SMS\Client\Configuration\Client Properties\Local SMS Path")
	Else
		g_logPath = g_WshShell.RegRead("HKLM\SOFTWARE\Wow6432Node\Microsoft\SMS\Client\Configuration\Client Properties\Local SMS Path")
	End If
	
	errorCode = Err.Number
	On Error Goto 0

	If errorCode <> 0 Then
		g_configMgrAgentPresent = False
	
		Set userEnv = g_WshShell.Environment("Process") 
		g_logPath = userEnv("TEMP")
	Else
			g_configMgrAgentPresent = True

			g_logPath = g_logPath & "\Logs"
	End If

	g_logPath = g_logPath & "\" & WScript.ScriptName & ".log"
	g_logPath = Replace (g_logPath, ".vbs", "", 1, 1, 1)

	WScript.Echo " >> " & g_logPath
	
	On Error Resume Next
	Err.Clear

	Set logfile = g_fso.OpenTextFile(g_logPath, 8, True)
	
	errorCode = Err.Number
	On Error Goto 0
	
	If errorCode <> 0 Then
		g_logPath = ""
	End If

	logfile.Close
	
	WriteLogMsg MSG_DIVIDER, 1, g_echoLog
	WriteLogMsg MSG_MAIN_BEGIN & g_startTime, 1, g_echoLog
				
End Sub

