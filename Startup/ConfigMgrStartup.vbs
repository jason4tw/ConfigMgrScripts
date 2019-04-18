' ConfigMgr Startup Script
' Version 1.83
' Jason Sandys
' https://home.configmgrftw.com
'
' 1.0:		Orignal Release
' 1.01:		Minor corrections on lines 493 and 496
' 1.02:		Added sleep on line 529 to wait for a service to start up fully 
'           Added Err.Clear on line 952 to clear the Error
'			Changed expectedStartMode to newStartMode on 510 and 513
' 1.03:		Changed options to parameters on line 275
' 1.50:		Added CheckCacheDuringStartup function to use WMI instead of UIResource
'				which can't normally be used when script is run as a startup
'			Normalized error messages by changing "error code" to "error"
'			Added error descriptions to log file in addition to error codes
'			Added default true value to CheckLocalAdmin sub to handle Case
'				where no local admin specified in configuration file
'			Added Check client assignment message to beginning of CheckAssignment Sub
'			Added use of multiple PATCH properties
'			Added auto-hotfix ability
' 			Added startup Delay
' 1.51:		Added additional error checking in CheckService and CheckServices functions
'			Fixed bug in AutoHotfix Sub
'			Added WriteErrorFile sub and error checking
'			Added date/time stamps to begining and end of execution and elapsed Time
'			Added set clientError = true for WMI errors to create the error file
'			Added error checking for error files
'			Added last result tracking to the registry to help maintain state and 
'				determine whether we should try to delete the error file
' 1.52:		Added errorcheck to CheckClient Function to check when ccm WMI namespace is
'				found but no SMS_Agent object exists -- maybe the client existed previously
' 1.53:		Added WinPE check to beginning so script simply exits if run in WinPE
'			Added IsAdmin check to determine if user is an admin or not (code copied from
'			http://csi-windows.com/toolkit/csi-IsAdmin)
' 1.54:		Updated WMI Namespace to check
' 1.55:		Added last error message to error log file and updated naming format for Error
'				log file include FQDN
' 1.56:		Added Extra error checking and logging to CheckCache Sub
' 1.57:		Added ability to specify multiple accounts to add to the local admins group, comma separated
' 1.60:		Added ability to specify ccmsetup parameters (like source and BITSPriority) 
'				using the CCMSetupParameter XML element
'			Updated AutoHotfix feature to automatically detect OS Architecture and
'				then select either the i386 or x64 subfolder of the directory specified.
'				This only works is the AgentVersion major value is 5.
' 1.61:		Added undefined constants
' 1.65:		Added auto discovery of local built-in admin group name for non-english systems
'				Updated client version comparison to account for SP1 CU3 which uses a single zero 
'			for the second part of the major version instead of a double zero
'				Modified CCMSetupParameter option to account for ccmsetup parameters that do not
'			have values like /NoCRLCheck
' 1.66:		Added registry value deletion by setting the Delete attribute of the RegistryValueCheck
'				element to True.
' 1.67:		Fixed problem if install path (location of ccmsetup.exe) contains spaces.
' 1.68:		Corrected bug in 1.67
' 1.69:		Fixed issue starting the SMS Agent Host (ccmexec) service if it's not started
' 1.70:		Added support for randomly choosing between multiple install locations
' 1.75:		Fixed issue with AutoHotfix if spaces exist in the msp paths
'			Added fix to disable zone checks when running ccmsetup
'			Added code to check for base client agent version so that only hotfixes can be applied 
'				instead of rerunning ccmsetup (this feature is not complete yet though)
' 1.80		Added pre-requsiste checks to prevent script from performing any checks or installing
'			the client agent
'			Corrected elapsed time calculation for final log message
' 1.81		Updated error handling when checking for admin shares
' 1.82		Updated log file handling when moving the log file to ccm/logs. If a log file for the script
'				already exists in ccm/logs, it will be renamed to include the date and time and then the new
'				log will be copied in.
' 1.83		Update CacheSize checking so that a default value is no longer forced during client agent install, 
'			reinstall, or during any run of the script if the option is ommited from the configuratin file or
'			set to zero.

Option Explicit

Dim g_fso, g_WshShell, g_logPathandName, g_startTime, g_lastErrorMsg

Set g_fso = CreateObject ("Scripting.FileSystemObject")
Set g_WshShell = WScript.CreateObject("WScript.Shell")

Const OPTION_LOCALADMIN =					"LocalAdmin"
Const OPTION_LOCALADMIN_GROUP =				"LocalAdminGroup"
Const OPTION_AGENTVERSION =					"AgentVersion"
Const OPTION_UAGENTVERSION =				"UnpatchedAgentVersion"
Const OPTION_DEFAULT_RUNINTERVAL =			"MinimumInterval"
Const OPTION_CACHESIZE =					"CacheSize"

Const OPTION_INSTALLPATH = 					"ClientLocation"
Const OPTION_SITECODE =						"SiteCode"
Const OPTION_MAXLOGFILE_SIZE =				"MaxLogFile"
Const OPTION_ERROR_LOCATION =				"ErrorLocation"
Const OPTION_AUTOHOTFIX =					"AutoHotfix"
Const OPTION_STARTUPDELAY =					"Delay"
Const OPTION_WMISCRIPT =					"WMIScript"
Const OPTION_WMISCRIPT_ASYNCH =				"WMIScriptAsynch"
Const OPTION_WMISCRIPTOPTIONS =				"WMIScriptOptions"

Const DEFAULT_REGISTRY_LOCATION =			"HKLM\Software\ConfigMgrStartup"
Const DEFAULT_LOCALADMIN_GROUP = 			"Administrators"
Const DEFAULT_AGENTVERSION =				"4.00.6487.2000"
Const DEFAULT_UAGENTVERSION =				"0.00.0000.0000"
Const DEFAULT_RUN_INTERVAL =				12
Const DEFAULT_CACHESIZE =					"0"

Const DEFAULT_REGISTRY_LASTRUN_VALUE =		"Last Run"
Const DEFAULT_REGISTRY_LOGLOCATION_VALUE =	"Log Location"
Const DEFAULT_REGISTRY_LASTRESULT_VALUE =	"Last Execution Result"
Const DEFAULT_CONFIGFILE_PARAMETER =		"config"
Const DEFAULT_MAXLOGFILE_SIZE =				"2048"
Const DEFAULT_EVENTLOG_PREFIX =    			"ConfigMgr StartUp Script -- "

Const DEFAULT_WMISCRIPT_ASYNCH =			"1"

Const MSG_OK =								"...OK"
Const MSG_NOTOK =							"...FAILED"
Const MSG_FOUND = 							"...found"
Const MSG_NOTFOUND =						"...not found"

Const MSG_DEBUG =							"DEBUG: "

Const MSG_MAIN_BEGIN =						"Beginning Execution at "
Const MSG_MAIN_FINISH =						"Finished Execution at "
Const MSG_ELAPSED_TIME = 					"Total script execution time is "
Const MSG_DIVIDER =							"----------------------------------------"

Const MSG_LOGMSG_CLIENTSTATUS =        		"Client Check "
Const MSG_LOGMSG_FILEERROR =        		"Unable to create or update error file at "
Const MSG_LOGMSG_FILEOK =        			"Successfully created or updated error file at "
Const MSG_LOGMSG_FILEDELETEERROR = 			"Unable to remove error file at "
Const MSG_LOGMSG_FILEDELETEOK = 			"Successfully removed error file at "

Const MSG_OPENCONFIG_NOT_SPECIFIED =		"Configuration file not specified on command-line with config switch"
Const MSG_OPENCONFIG_DOESNOTEXIST =			"The configuration file does not exist: "
Const MSG_OPENCONFIG_PARSEERROR	=			"The specified configuration file contains a parsing error: "
Const MSG_OPENCONFIG_OPENED	=				"Opened configuration file: "

Const MSG_LOADOPTIONS_STARTED = 			"Loading Options and Parameters from configuration file"
Const MSG_LOADOPTIONS_OPTIONLOADED =		"Option loaded: "
Const MSG_LOADOPTIONS_PROPERTYLOADED =		"Property loaded: "
Const MSG_LOADOPTIONS_PARAMLOADED =			"Parameter loaded: "

Const MSG_LASTRUN_VERIFYING =				"Verifying Last Run time from Registry: "
Const MSG_LASTRUN_NOLASTRUN =				"No last run time recorded in registry"
Const MSG_LASTRUN_TIME =					"Last run time: "
Const MSG_LASTRUN_TIMENOTOK =				"Existing because last run time was less than expected number of hours ago: "

Const MSG_LASTRESULT_VERIFYING =			"Verifying Last Result from Registry: "
Const MSG_LASTRESULT_NOLASTRESULT =			"No last result recorded in registry"
Const MSG_LASTRESULT_RESULT =				"Last execution result: "
Const MSG_LASTRESULT_FAIL =					"Failed"
Const MSG_LASTRESULT_SUCCEED = 				"Succeeded"

Const MSG_CHECKWMI_ERROR = 					"Error Connecting to WMI: "
Const MSG_CHECKWMI_SUCCESS = 				"Successfully Connected to WMI"

Const MSG_WMISCRIPT_NOTFOUND =				"Could not find WMI Script: "
Const MSG_WMISCRIPT_EXECUTING =				"Executing WMI Script: "
Const MSG_WMISCRIPT_EXECUTINGASYNCH =		"Asynchonously executing WMI Script: "
Const MSG_WMISCRIPT_OPTIONS =				"WMI Script Options: "
Const MSG_WMISCRIPT_ERROR =					"Failed to successfully run the WMI Script: "
Const MSG_WMISCRIPT_SUCCESS =				"Successfully ran the WMI Script."
Const MSG_WMISCRIPT_SUCCESSASYNCH =			"Successfully started the WMI Script, not waiting for results."

Const MSG_CHECKSERVICES_START = 			"START: Service Check..."
Const MSG_CHECKSERVICE_STARTMODE =			"...expected StartMode of "
Const MSG_CHECKSERVICE_STARTMODEOK =		"...set start mode to "
Const MSG_CHECKSERVICE_STARTMODEFAIL =		"...failed to set start mode with error: "
Const MSG_CHECKSERVICE_STATE =				"...expected State of "
Const MSG_CHECKSERVICE_STARTEDOK =			"...started service"
Const MSG_CHECKSERVICE_STARTEDFAIL =		"...could not start service, error: "
Const MSG_CHECKSERVICE_STOPPEDOK =			"...stopped service"
Const MSG_CHECKSERVICE_STOPPEDFAIL =		"...could not stop service, error: "

Const MSG_CHECKADMINSHARE_START = 			"START: Admin Share Check..."
Const MSG_CHECKADMINSHARE_SETSUCCESS =		"...set AutoShareWks registry value, a reboot is required to create the Admin$ share."
Const MSG_CHECKADMINSHARE_SETFAIL =			"...unable to set AutoShareWks registry value: "

Const MSG_CHECKREGISTRY_START =				"START: Registry Check..."
Const MSG_CHECKREGISTRY_EXPECTED =			"...expected value of "
Const MSG_CHECKREGISTRY_ENFORCEOK =			"...successfully updated value"
Const MSG_CHECKREGISTRY_ENFORCEFAIL =		"...failed to update value"
Const MSG_CHECKREGISTRY_DELETE =			"...deleting value"
Const MSG_CHECKREGISTRY_DELETEOK =			"...successfully deleted value"
Const MSG_CHECKREGISTRY_DELETEFAIL =		"...failed to delete value"

Const MSG_CHECKLOCALADMIN_START = 			"START: Local Admin Check..."
Const MSG_CHECKLOCALADMIN_ALREADYMEMBER =	"...user already member"
Const MSG_CHECKLOCALADMIN_ADDMEMBEROK =		"...user add successful"
Const MSG_CHECKLOCALADMIN_ADDMEMBERFAIL = 	"...user add failed with error: "
Const MSG_CHECKLOCALADMIN_FINDLOCALGROUP =  " *Using local admin group named "
Const MSG_CHECKLOCALADMIN_CHECKACCOUNT =	" *Checking local admin group membership for "
Const MSG_CHECKLOCALADMIN_FINISH =			"Finished checking local admin group membership: "

Const MSG_GETLOCALADMINGROUPNAME = "No local admin group name specified, discovering... "

Const MSG_CHECKCLIENT_START =				"START: Checking Client Status..."
Const MSG_CHECKCLIENT_WMINOTFOUND =			" *Cannot connect to ConfigMgr WMI Namespace: "
Const MSG_CHECKCLIENT_VERSION =				"START: Getting ConfigMgr agent version..."
Const MSG_CHECKCLIENT_VERSIONNOTFOUND =		" *Cannot determine ConfigMgr agent version"
Const MSG_CHECKCLIENT_OLDVERSION =			" *Old version of agent found: "
Const MSG_CHECKCLIENT_UNPATCHEDVERSION =	" *Unpatched version of agent found: "
Const MSG_CHECKCLIENT_CCMEXEC =				"START: Checking SMS Agent Host Status..."
Const MSG_CHECKCLIENT_VERSIONFOUND =		" *SMS Agent Host version: "
Const MSG_CHECKCLIENT_MOVEDLOG =			"Moved log file to "
Const MSG_CHECKCLIENT_MOVELOGFAIL =			"Unable to move log file, error: "
Const MSG_CHECKCLIENT_GETLOGDIRECTORY =		"Unable to get agent log directory with error: "
Const MSG_CHECKCLIENT_VERSIONEXPECTED =		" *Expected client version: "
Const MSG_CHECKCLIENT_UVERSIONEXPECTED =	" *Expected unpatched client version: "

Const MSG_CHECKCACHE_START =				"START: Check agent cache..."
Const MSG_CHECKCACHE_CREATEFAIL =			" *Could not create UIResourceManager with error: "
Const MSG_CHECKCACHE_WMIFAIL =				" *Could not retrieve the cache object from WMI with error: "
Const MSG_CHECKCACHE_WMIWRITEFAIL =			" *Could not set the cache size in WMI with error: "
Const MSG_CHECKCACHE_CACHEFAIL =			" *Could not retrieve agent cache size with error: "
Const MSG_CHECKCACHE_SETSIZE =				" *Set cache size to "
Const MSG_CHECKCACHE_SIZEOK =	  			" *Current cache size matches desired size."
Const MSG_CHECKCACHE_SIZEIS =	  			" *Current cache size is "
Const MSG_CHECKCACHE_DESIREDSIZE =			" *Desired cache size is "

Const MSG_PREREQ_SERVICE =					"Prerequisite check: checking for service named "
Const MSG_PREREQ_REGKEY = 					"Prerequisite check: checking for a reg key at "
Const MSG_PREREQ_REGVALUE = 				"Prerequisite check: checking for a reg value at "
Const MSG_PREREQ_PASSED =					" *Passed: "
Const MSG_PREREQ_FAILED =					" *Failed: "
Const MSG_PREREQ_EXISTS =					"item exists"
Const MSG_PREREQ_DOESNOTEXIST =				"item does not exist"
Const MSG_PREREQ_ALLPASSED =				"All prerequisities passed"
Const MSG_PREREQ_ONEFAILED =				"At least one prequisitie failed, exiting ..."

Const MSG_INSTALLCLIENT_START =				"START: Client Install..."
Const MSG_INSTALLCLIENT_PATHCHECK =			" *Checking for ccmsetup in "
Const MSG_INSTALLCLIENT_COMMANDLINE =		" *Initiating client install with command-line: "
Const MSG_INSTALLCLIENT_SUCCESS =			" *Successfully initiated CCMSetup"
Const MSG_INSTALLCLIENT_FAILED =			" *Failed to initiate CCMSetup with error: "
Const MSG_AUTOPATCH_COMMANDLINE =			" *Discovering client hotfixes from: "
Const MSG_AUTOPATCH_DIRERROR =				" *Unable to open the hotfix folder: "
Const MSG_AUTOPATCH_FOUNDHOTFIX =			"  ...Found hotfix: "
Const MSG_HOTFIX_FILEVERIFY =				" *Verifying hotfix accessibility: "
Const MSG_HOTFIX_DUPLICATE =				"  ...Hotfix already added: "
Const MSG_HOTFIX_MULTIPLE =					" *Multiple hotfixes specified, cannot verify accessibility "

Const MSG_OS_ARCHITECTURE =					" *Detected OS Architecture is: "

Const MSG_CHECKASSIGNMENT_START =			"START: Checking client assignment..."
Const MSG_CHECKASSIGNMENT_OK =				" *Client assigned to site " 
Const MSG_CHECKASSIGNMENT_NOTOK =			" *Client not assigned to site, initiaing (re-)install"

If Not InWinPE And IsAdmin Then
	Main
End If

Sub Main

	Dim argsNamed
	Dim WshShell
	Dim xmlConfig
	Dim configOptions
	Dim msiProperties
	Dim ccmsetupParams
	Dim preReqs
	Dim clientError
	Dim lastResult
	
	clientError = False
	g_startTime = Now
	
	'On Error Resume Next
	
	Set argsNamed = WScript.Arguments.Named
	
	Set configOptions = WScript.CreateObject("Scripting.Dictionary")
	Set msiProperties = WScript.CreateObject("Scripting.Dictionary")
	Set ccmsetupParams = WScript.CreateObject("Scripting.Dictionary")

	WriteFirstLogMsg 
	
	If OpenConfig (argsNamed, xmlConfig) Then
	
		WriteLogMsg MSG_LOADOPTIONS_STARTED, 1, 1, 0
		
		LoadOptions xmlConfig, configOptions, 0, "Option", MSG_LOADOPTIONS_OPTIONLOADED
		LoadOptions xmlConfig, msiProperties, 0, "InstallProperty", MSG_LOADOPTIONS_PROPERTYLOADED
		LoadOptions xmlConfig, ccmsetupParams, 1, "CCMSetupParameter", MSG_LOADOPTIONS_PARAMLOADED

		If CheckPreReqs(xmlConfig) Then 

			lastResult = GetLastResult(configOptions)
			
			If LastRunOK(configOptions) = True Then
			
				If configOptions.Exists(OPTION_STARTUPDELAY) Then
					Delay CInt(configOptions.Item(OPTION_STARTUPDELAY))
				End If 
				
				If CheckWMI(configOptions) = True Then
				
					If Not CheckServices(xmlConfig) Or Not CheckAdminShare Or Not CheckRegistry(xmlConfig) Or Not CheckLocalAdmin(configOptions) Then
						clientError = True
					End If
					
					If CheckClient(configOptions) = False Then
						If Not InstallClient(configOptions, msiProperties, ccmsetupParams) Then
							clientError = True
						End If				
					Else If options.Exists(OPTION_CACHESIZE) Then
						CheckCache configOptions
					End If
				Else
					clientError = True

				End If
				
				If configOptions.Exists(OPTION_ERROR_LOCATION) Then
					WriteErrorFile clientError, configOptions.Item(OPTION_ERROR_LOCATION), lastResult												
				End If
			End If
		End If
	End If
	
	WriteFinalLogMsg configOptions
	
End Sub

Function CheckPreReqs (ByRef config)

	Dim returnValue
	Dim wmi, errorCode, errorMsg
	Dim prereqServiceNodes, prereqRegNodes 
	Dim prereqCheckNode, value, condition, service
	
	returnValue = True

	Set prereqServiceNodes = config.documentElement.selectNodes ( "/Startup/PreReq[@Type = 'Service']" )
	
	If prereqServiceNodes.Length > 0 Then

		On Error Resume Next
		Err.Clear
	
		Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

		errorCode = Err.Number
		errorMsg = Err.Description & " (" & Err.Number & ")"
		
		On Error Goto 0
		
		If errorCode <> 0 Then
			WriteLogMsg MSG_CHECKWMI_ERROR & errorMsg, 3, 1, 0
			CheckPreReqs = False
		Else
			For Each prereqCheckNode In prereqServiceNodes
		
				value = prereqCheckNode.text
				condition = prereqCheckNode.GetAttribute("Condition")

				If condition <> "MustExist" And condition <> "MustNotExist" Then
					condition = "MustExist"
				End If

				WriteLogMsg MSG_PREREQ_SERVICE & value & " which " & condition, 1, 1, 0

				On Error Resume Next
				Err.Clear

				Set service = wmi.Get("Win32_Service.Name='" & value & "'")
	
				errorCode = Err.Number
				errorMsg = Err.Description & " (" & Err.Number & ")"

				On Error Goto 0
				
				returnValue = returnValue And EvaluatePreReq(errorCode = 0, condition)
			
			Next
		End If

		Set wmi = Nothing

	End If

	Set prereqRegNodes = config.documentElement.selectNodes ( "/Startup/PreReq[@Type = 'Reg']" )
	
	If prereqRegNodes.Length > 0 Then

		For Each prereqCheckNode In prereqRegNodes
	
			value = prereqCheckNode.text
			condition = prereqCheckNode.GetAttribute("Condition")

			If condition <> "MustExist" And condition <> "MustNotExist" Then
				condition = "MustExist"
			End If

			If Right(value, 1) = "\" Then
				WriteLogMsg MSG_PREREQ_REGKEY & value & " (" & condition & ")", 1, 1, 0
			Else
				WriteLogMsg MSG_PREREQ_REGVALUE & value & " (" & condition & ")", 1, 1, 0
			End If			

			On Error Resume Next
			Err.Clear

			g_WshShell.RegRead(value)

			errorCode = Err.Number
			errorMsg = Err.Description & " (" & Err.Number & ")"

			On Error Goto 0
			
			returnValue = returnValue And EvaluatePreReq(errorCode = 0, condition)
		
		Next

	End If

	If returnValue = True Then
		WriteLogMsg MSG_PREREQ_ALLPASSED, 1, 1, 0
	Else
		WriteLogMsg MSG_PREREQ_ONEFAILED, 3, 1, 0
	End If

	CheckPreReqs = returnValue

End Function

Function EvaluatePreReq(ByVal itemExists, ByVal prereqCondition)

	Dim msg, msgType
	
	EvaluatePreReq = False

	If (itemExists = True And prereqCondition = "MustExist") Or (itemExists = False And prereqCondition = "MustNotExist") Then
		EvaluatePreReq = True
		msgType = 1
		msg = MSG_PREREQ_PASSED
	Else
		EvaluatePreReq = False
		msgType = 3
		msg = MSG_PREREQ_FAILED
	End If

	If itemExists Then
		msg = msg + MSG_PREREQ_EXISTS
	Else
		msg = msg + MSG_PREREQ_DOESNOTEXIST
	End If

	WriteLogMsg msg, msgType, 1, 0

End Function

 Sub WriteErrorFile (ByVal clientErr, ByVal errorLocation, ByVal lastExecutionResult)
	Dim badLog, badLogFileName
	'Dim network
	Dim regKey, fqdn
	Dim errorCode, errorMsg
	Dim registryLocation
	
	'Set network = WScript.CreateObject("WScript.Network")
	
	regKey = "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
	fqdn = g_WshShell.RegRead( regKey & "\hostname" ) & "." & g_WshShell.RegRead( regKey & "\domain" )
	
	'badLogFileName = errorLocation & "\" & network.ComputerName & ".log"
	badLogFileName = errorLocation & "\" & fqdn & ".log"

	registryLocation = DEFAULT_REGISTRY_LOCATION & "\" & DEFAULT_REGISTRY_LASTRESULT_VALUE

	If clientErr = True Then

		WriteLogMsg MSG_LOGMSG_CLIENTSTATUS & MSG_LASTRESULT_FAIL, 1, 1, 0
		g_WshShell.LogEvent 1, MSG_LOGMSG_CLIENTSTATUS & MSG_LASTRESULT_FAIL
		g_WshShell.RegWrite registryLocation, MSG_LASTRESULT_FAIL, "REG_SZ"

		On Error Resume Next
		Err.Clear

		Set badLog = g_fso.OpenTextFile(badLogFileName, 8, True)
	    errorCode = Err.Number
	    errorMsg = Err.Description & " (" & Err.Number & ")"
    
   		On Error Goto 0
    
	    If errorCode <> 0 Then
	    	WriteLogMsg MSG_LOGMSG_FILEERROR & badLogFileName & ": " & errorMsg, 3, 1, 0
	    	Exit Sub
	    End If

		badLog.WriteLine Date & " " & Time
		badLog.WriteLine g_lastErrorMsg
		badLog.Close
		
    	WriteLogMsg MSG_LOGMSG_FILEOK & badLogFileName, 1, 1, 0
		
	Else 
		WriteLogMsg MSG_LOGMSG_CLIENTSTATUS & MSG_LASTRESULT_SUCCEED, 1, 1, 0
		g_WshShell.LogEvent 0, MSG_LOGMSG_CLIENTSTATUS & MSG_LASTRESULT_SUCCEED
		g_WshShell.RegWrite registryLocation, MSG_LASTRESULT_SUCCEED, "REG_SZ"

		If lastExecutionResult = False Then
			On Error Resume Next
			Err.Clear
	
			g_fso.DeleteFile badLogFileName, True
		    errorCode = Err.Number
		    errorMsg = Err.Description & " (" & Err.Number & ")"
	    
	   		On Error Goto 0
	    
		    If errorCode <> 0 Then
		    	WriteLogMsg MSG_LOGMSG_FILEDELETEERROR & badLogFileName & ": " & errorMsg, 2, 1, 0
		    Else
		    	WriteLogMsg MSG_LOGMSG_FILEDELETEOK & badLogFileName, 1, 1, 0
			End If
		End If

	End If				

End Sub

Sub Delay(ByVal delayTime)

	Dim countdown
	
	For countdown = delayTime To 1 Step -1
		WriteLogMsg countdown & "...", 1, 1, 0
		WScript.Sleep(1000)
	Next
	
End Sub

Function GetLastResult(ByRef options)
	Dim registryLocation, lastResult
	Dim errorCode
	
	GetLastResult = True
	
	registryLocation = DEFAULT_REGISTRY_LOCATION & "\" & DEFAULT_REGISTRY_LASTRESULT_VALUE
	
	WriteLogMsg MSG_LASTRESULT_VERIFYING & registryLocation, 1, 1, 0
	
	On Error Resume Next

	Err.Clear
	lastResult = g_WshShell.RegRead(registryLocation)
	
	errorCode = Err.Number
	
	On Error Goto 0
	
	If errorCode <> 0 Then
		WriteLogMsg MSG_LASTRESULT_NOLASTRESULT, 1, 1, 0
		lastResult = MSG_LASTRESULT_SUCCEED
	Else	
		WriteLogMsg MSG_LASTRESULT_RESULT & lastResult, 1, 1, 0
		
	End If
	
	If lastResult = MSG_LASTRESULT_FAIL Then
		GetLastResult = False
	End If		
	
End Function

Function OpenConfig(ByRef args, ByRef config)

	Dim configFilename

	OpenConfig = False
	
	'Check for the proper command line arguments
	If Not ( args.Exists ( DEFAULT_CONFIGFILE_PARAMETER ) ) Then
	 	' Print the proper usage and return
	 	WriteLogMsg MSG_OPENCONFIG_NOT_SPECIFIED, 3, 1, 1
		Exit Function
	End If
	
	configFilename = args.Item ( DEFAULT_CONFIGFILE_PARAMETER )

	' Check to make sure the specified config file exists
	If Not g_fso.FileExists ( configFilename ) Then
		WriteLogMsg MSG_OPENCONFIG_DOESNOTEXIST & configFilename, 3, 1, 1
		Exit Function
	End If

	Set config  = CreateObject ( "Msxml2.DOMDocument" )

	' Load the whole XML config file at once
	config.async = False
	config.load ( configFilename )
	
	' Check the file to make sure it is valid XML
	If config.parseError.errorCode <> 0 Then
		WriteLogMsg MSG_OPENCONFIG_PARSEERROR & config.parseError.reason, 3, 1, 0
		Exit Function
	Else
		' Set our XML query language to XPath
		config.setProperty "SelectionLanguage", "XPath"
		WriteLogMsg MSG_OPENCONFIG_OPENED & configFilename, 1, 1, 1
	End If
	
	OpenConfig = True
	
End Function

Sub LoadOptions(ByRef config, ByRef values, ByVal allowBlankValues, ByVal nodeName, ByVal msg)

	Dim nodes, node, name, value

	Set nodes = config.documentElement.selectNodes ( "/Startup/" & nodeName )

	For Each node In nodes
	
		name = node.getAttribute("Name")
		value = node.text
		
		If allowBlankValues = 1 And Len(value) = 0 Then
			value = ""
		End If
		
		If allowBlankValues = 1 Or Len(value) > 0 Then
		
			values.Add name, value
			WriteLogMsg msg & name & ": '" & value & "'", 1, 1, 0
			
		End If
		
	Next

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
	
	On Error Goto 0
	
	If msgtype = 3 Or msgtype = 2 Then
		g_lastErrorMsg = msg
	End If

	'If echomsg = 1 Then
	'	WScript.Echo msg
	'End If
	
	If eventlog = 1 Then
	 g_WshShell.LogEvent 0, DEFAULT_EVENTLOG_PREFIX & msg
	End If
	
End Sub

Sub WriteFinalLogMsg(ByRef options)

	Dim registryLocation
	Dim logfile
	Dim logFileSize, maxLogFileSize
	Dim finishTime, totalSeconds, elapsedHours, elapsedMinutes, elapsedSeconds
	
	registryLocation = DEFAULT_REGISTRY_LOCATION & "\" & DEFAULT_REGISTRY_LOGLOCATION_VALUE
	
	g_WshShell.RegWrite registryLocation, g_logPathandName, "REG_SZ"

	finishTime = Now

	totalSeconds = DateDiff ("s", g_startTime, finishTime)
	
	elapsedHours = CInt(totalSeconds / 3600)
	totalSeconds = totalSeconds Mod 3600

	elapsedMinutes = CInt(totalSeconds / 60)
	totalSeconds = totalSeconds Mod 60

	elapsedSeconds = totalSeconds
	
	WriteLogMsg MSG_MAIN_FINISH & finishTime, 1, 1, 0
	WriteLogMsg MSG_ELAPSED_TIME & ZeroPadNumber(elapsedHours) & ":" & ZeroPadNumber(elapsedMinutes) & ":" & ZeroPadNumber(elapsedSeconds), 1, 1, 0
	WriteLogMsg MSG_DIVIDER, 1, 0, 0
	
	Set logfile = g_fso.GetFile(g_logPathandName)
	
	logFileSize = logfile.Size / 1024
	
	maxLogFileSize = GetOptionValue(OPTION_MAXLOGFILE_SIZE, DEFAULT_MAXLOGFILE_SIZE, options)
	
	If logFileSize > maxLogFileSize Then
		g_fso.CopyFile g_logPathandName, g_logPathandName & ".old", True
		g_fso.DeleteFile g_logPathandName, True
	End If
	
	g_WshShell.LogEvent 0, MSG_MAIN_FINISH & g_logPathandName
	
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

	Dim registryLocation
	Dim logfile, logpath
	Dim userEnv
	Dim errorCode
	Dim logFileSize, maxLogFileSize

	Set userEnv = g_WshShell.Environment("Process") 

	logpath = userEnv("TEMP")

	registryLocation = DEFAULT_REGISTRY_LOCATION & "\" & DEFAULT_REGISTRY_LOGLOCATION_VALUE
	
	On Error Resume Next
	
	g_logPathandName = g_WshShell.RegRead (registryLocation)
	
	errorCode = Err.Number                               
	On Error Goto 0
	
	If errorCode = 0 Then
		
		On Error Resume Next
		
		Set logfile = g_fso.OpenTextFile(g_logPathandName, 8, True)
		
		errorCode = Err.Number
		
		On Error Goto 0
		
		If errorCode <> 0 Then
			g_logPathandName = logpath & "\" & WScript.ScriptName & ".log"
		Else
			logfile.Close
	
		End If
	Else
		g_logPathandName = logpath & "\" & WScript.ScriptName & ".log"
	End If
	
	WriteLogMsg MSG_DIVIDER, 1, 1, 0
	WriteLogMsg MSG_MAIN_BEGIN & g_startTime, 1, 1, 1
				
End Sub

Function LastRunOK(ByRef options)
	
	Dim registryLocation, lastRunTime, lastRunInterval, minimumRunInterval
	Dim errorCode
	
	registryLocation = DEFAULT_REGISTRY_LOCATION & "\" & DEFAULT_REGISTRY_LASTRUN_VALUE
	
	WriteLogMsg MSG_LASTRUN_VERIFYING & registryLocation, 1, 1, 0
	
	On Error Resume Next

	Err.Clear
	lastRunTime = g_WshShell.RegRead(registryLocation)
	
	errorCode = Err.Number
	
	On Error Goto 0
	
	minimumRunInterval = CInt(GetOptionValue(OPTION_DEFAULT_RUNINTERVAL, DEFAULT_RUN_INTERVAL, options))

	If errorCode <> 0 Then
		WriteLogMsg MSG_LASTRUN_NOLASTRUN, 1, 1, 0
		lastRunInterval = minimumRunInterval + 1
	Else	
		lastRunInterval = DateDiff("h", lastRunTime, Now) 
		
		WriteLogMsg MSG_LASTRUN_TIME & lastRunTime, 1, 1, 0
	
	End If
	
	If lastRunInterval < minimumRunInterval Then
		LastRunOK = False
		WriteLogMsg MSG_LASTRUN_TIMENOTOK & minimumRunInterval, 1, 1, 0
	Else
		g_WshShell.RegWrite registryLocation, Now, "REG_SZ"
		LastRunOK = True
	End If
	
End Function

Function CheckWMI (ByRef options)
	
	Dim errorCode, errorMsg
	Dim fixScript, fixScriptOptions, fixScriptPath
	Dim fixScriptAsynch
	Dim wmiOK
	
	wmiOK = CheckWMIConnectivity(options)
	CheckWMI = wmiOK
	If wmiOK <> True Then
    	
    	fixScript = GetOptionValue(OPTION_WMISCRIPT, "0", options) 
    	fixScriptAsynch = GetOptionValue(OPTION_WMISCRIPT_ASYNCH, DEFAULT_WMISCRIPT_ASYNCH, options)
    	fixScriptOptions = GetOptionValue(OPTION_WMISCRIPTOPTIONS, "", options)
    	
    	If fixScript <> "0" Then
    	
    		fixScriptPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") & fixScript
    		
			If Not g_fso.FileExists(fixScriptPath) Then
        		WriteLogMsg MSG_WMISCRIPT_NOTFOUND & fixScriptPath, 3, 1, 0

    		Else
				If fixScriptAsynch = "0" Then
        			WriteLogMsg MSG_WMISCRIPT_EXECUTING & fixScriptPath, 1, 1, 0
        		Else
        			WriteLogMsg MSG_WMISCRIPT_EXECUTINGASYNCH & fixScriptPath, 1, 1, 0
        		End If
        		
        		If fixScriptOptions <> "" Then
        			fixScriptOptions = Replace(fixScriptOptions, "%logpath%", options.Item(OPTION_ERROR_LOCATION))
        		
        			WriteLogMsg MSG_WMISCRIPT_OPTIONS & fixScriptOptions, 1, 1, 0
        		End If
        		
				On Error Resume Next
				Err.Clear

    			g_WshShell.Run "cscript.exe " & Chr(34) & fixScriptPath & Chr(34) & " " & fixScriptOptions, 0, Eval(fixScriptAsynch = "0")
    			
    			On Error Goto 0
    			
       			If errorCode <> 0 Then
        			WriteLogMsg MSG_WMISCRIPT_ERROR & errorCode, 3, 1, 0
   				ElseIf fixScriptAsynch <> "0" Then
	        		WriteLogMsg MSG_WMISCRIPT_SUCCESSASYNCH, 1, 1, 0
   				Else
	        		WriteLogMsg MSG_WMISCRIPT_SUCCESS, 1, 1, 0
	        		
					If fixScriptAsynch = "0" Then
						CheckWMI = CheckWMIConnectivity(options)
					End If
				
				End If

    		End If
    	
    	End If
    
    End If
    
End Function

Function CheckWMIConnectivity (ByRef options)
    Dim wmi
	Dim errorCode, errorMsg
    
    On Error Resume Next
	Err.Clear

    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    errorCode = Err.Number
    errorMsg = Err.Description & " (" & Err.Number & ")"
    
    If errorCode <> 0 Then
    	CheckWMIConnectivity = False
		WriteLogMsg MSG_CHECKWMI_ERROR & errorMsg, 3, 1, 0
	Else
    	CheckWMIConnectivity = True
 		WriteLogMsg MSG_CHECKWMI_SUCCESS, 1, 1, 0

	End If
	
    Set wmi = Nothing

End Function

Function CheckServices(ByRef config)

	Dim wmi, errorCode, errorMsg
	Dim serviceCheckNodes, serviceToCheck, serviceName, expectedServiceState, expectedServiceStartMode, enforce
	Dim returnCode
	
	CheckServices = True
	
	Set serviceCheckNodes = config.documentElement.selectNodes ( "/Startup/ServiceCheck" )
	
	On Error Resume Next
	Err.Clear

	Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

    errorCode = Err.Number
    errorMsg = Err.Description & " (" & Err.Number & ")"
    
   	On Error Goto 0
    
    If errorCode <> 0 Then
    	WriteLogMsg MSG_CHECKWMI_ERROR & errorMsg, 3, 1, 0
   		CheckServices = False
	Else
		WriteLogMsg MSG_CHECKSERVICES_START, 1, 1, 1
	
		For Each serviceToCheck In serviceCheckNodes
	
			serviceName = 				serviceToCheck.getAttribute("Name")
			expectedServiceState = 		serviceToCheck.getAttribute("State")
			expectedServiceStartMode =	serviceToCheck.getAttribute("StartMode")
			enforce =					serviceToCheck.getAttribute("Enforce")
	
			If Not CheckService(wmi, serviceName, expectedServiceState, expectedServiceStartMode, enforce) Then
				CheckServices = False
			End If
		
		Next
	End If
		
	Set wmi = Nothing

End Function

Function CheckService(ByRef wmi, serviceName, expectedServiceState, expectedServiceStartMode, enforce)
	Dim service
	Dim msg
	Dim serviceStatus
	Dim returnCode, errorCode, errorMsg
	Dim newStartMode
  
	serviceStatus = 1
	CheckService = True

	On Error Resume Next
	Err.Clear

	Set service = wmi.Get("Win32_Service.Name='" & serviceName & "'")
	
    errorCode = Err.Number
    errorMsg = Err.Description & " (" & Err.Number & ")"

	On Error Goto 0
    
    If errorCode <> 0 Then
	
		msg = " *" & serviceName & MSG_NOTFOUND & ": " & errorMsg
		WriteLogMsg msg & MSG_NOTFOUND, 2, 1, 0

		CheckService = False
		Exit Function
		
	End If
		
	msg = " *" & service.Name

	If IsObject(service) Then
		msg = msg & MSG_FOUND & " (" & service.State & "," & service.StartMode & ")"
		
		If service.StartMode <> expectedServiceStartMode Then
			msg = msg & MSG_CHECKSERVICE_STARTMODE & expectedServiceStartMode			

			If enforce = "True" Then
			
			    If expectedServiceStartMode = "Auto" Then
				   newStartMode = "Automatic"
	        	Else
	          		newStartMode = expectedServiceStartMode
	        	End If 
			
				returnCode = service.ChangeStartMode(newStartMode)
				
				If returnCode = 0 Then
					msg = msg & MSG_CHECKSERVICE_STARTMODEOK & newStartMode
				Else
					msg = msg & MSG_CHECKSERVICE_STARTMODEFAIL & returnCode
					serviceStatus = 3
					CheckService = False

				End If
	
			End If	
		End If
					
		If service.State <> expectedServiceState Then
			msg = msg & MSG_CHECKSERVICE_STATE & expectedServiceState
			
			If enforce = "True" Then
				
				If expectedServiceState = "Running" Then
					returnCode = service.StartService()
					
					WScript.Sleep(15000)
					
					If returnCode = 0 Then
						msg = msg &  MSG_CHECKSERVICE_STARTEDOK
					Else
						msg = msg &  MSG_CHECKSERVICE_STARTEDFAIL & returnCode
						serviceStatus = 3
						CheckService = False
					End If

				ElseIf expectedServiceState = "Stopped" Then
					returnCode = service.StopService()
					
					If returnCode = 0 Then
						msg = msg &  MSG_CHECKSERVICE_STOPPEDOK
					Else
						msg = msg &  MSG_CHECKSERVICE_STOPPEDFAIL & returnCode
						serviceStatus = 3
						CheckService = False

					End If
				End If
				
						
			End If	
		End If
		
	Else
		msg = msg & MSG_NOTFOUND
		serviceStatus = 2
	End If
	
	If serviceStatus = 1 Then
		msg = msg & MSG_OK
	Else
		msg = msg & MSG_NOTOK
	
	End If
	
	WriteLogMsg msg, serviceStatus, 1, 0
End Function

Function CheckAdminShare

	Dim wmi, adminShare, adminShareRegValue, errorMsg, errorCode
	Dim msg, status
	
	adminShareRegValue = 1
	status = 1

	WriteLogMsg MSG_CHECKADMINSHARE_START, 1, 1, 1

	Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

	On Error Resume Next
	Err.Clear
		
	Set adminShare = wmi.Get("Win32_Share.Name='ADMIN$'")
	
	errorCode = Err.Number
	
	On Error GoTo 0
	
	msg = " *Admin$"
	
	If IsObject(adminShare) And errorCode = 0 Then
		msg = msg & MSG_FOUND
	Else
		msg = msg &  MSG_NOTFOUND
		status = 3
		
		On Error Resume Next
		Err.Clear
		
		adminShareRegValue = g_WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Services\LanManServer\Parameters\AutoShareWks")
		
		errorCode = Err.Number
		errorMsg = Err.Description & " (" & Err.Number & ")"
		
	   	On Error Goto 0

		If errorCode = 0 And adminShareRegValue = 0 Then
		
			On Error Resume Next
			Err.Clear
			g_WshShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\LanManServer\Parameters\AutoShareWks", 1, "REG_DWORD"
			
			errorCode = Err.Number
			errorMsg = Err.Description & " (" & Err.Number & ")"
			
		   	On Error Goto 0

			If errorCode = 0 Then
				msg = msg & MSG_CHECKADMINSHARE_SETSUCCESS
			Else
				msg = msg & MSG_CHECKADMINSHARE_SETFAIL & errorMsg
				status = 3
			End If
		
		End If
		
	End If
	
	If status = 1 Then
		msg = msg & MSG_OK
		CheckAdminShare = True
	Else
		msg = msg & MSG_NOTOK
		CheckAdminShare = False
	
	End If
	
	WriteLogMsg msg, status, 1, 0
	
	Set wmi = Nothing

End Function

Function CheckRegistry(ByRef config)
	Dim registryCheckNodes, registryValueToCheck
	Dim regKey, regValue, expectedValue, valueType, enforce
	Dim errorCode
	Dim actualValue
	Dim msg, msgOk, msgFail
	Dim regStatus
	Dim deleteValue
	
	CheckRegistry = True

	Set registryCheckNodes = config.documentElement.selectNodes ( "/Startup/RegistryValueCheck" )
	
	WriteLogMsg MSG_CHECKREGISTRY_START, 1, 1, 1

	For Each registryValueToCheck In registryCheckNodes
		regKey = registryValueToCheck.getAttribute("Key")
		regValue = registryValueToCheck.getAttribute("Value")
		expectedValue = registryValueToCheck.getAttribute("Expected")
		enforce = registryValueToCheck.getAttribute("Enforce")
		valueType = registryValueToCheck.getAttribute("Type")
		deleteValue = registryValueToCheck.getAttribute("Delete")
		
		regStatus = 1

		If valueType = "REG_DWORD" Then
			expectedValue = CInt(expectedValue)
		End If
			
		On Error Resume Next
		Err.Clear
		
		actualValue = g_WshShell.RegRead(regKey & "\" & regValue)

		errorCode = Err.Number
		
		On Error Goto 0
		
		msg = " *" & regKey & "\" & regValue
		
		If errorCode <> 0 Then
			msg = msg & MSG_NOTFOUND
			regStatus = 2
		Else
			msg = msg & MSG_FOUND & " (" & actualValue & ")"
			
			If deleteValue = "True" Then
				msg = msg & MSG_CHECKREGISTRY_DELETE
				
				enforce = "True"
			
			ElseIf actualValue <> expectedValue Then
				msg = msg & MSG_CHECKREGISTRY_EXPECTED & expectedValue
				regStatus = 2
				CheckRegistry = False
			Else
				enforce = "False"
			End If
			
		End If
		
		If enforce = "True" Then
				
			On Error Resume Next
			Err.Clear

			If deleteValue = "True" Then
				msgOk = MSG_CHECKREGISTRY_DELETEOK
				msgFail = MSG_CHECKREGISTRY_DELETEFAIL

				g_WshShell.RegDelete regKey & "\" & regValue
			
				errorCode = Err.Number
			Else
				msgOk = MSG_CHECKREGISTRY_ENFORCEOK
				msgFail = MSG_CHECKREGISTRY_ENFORCEFAIL
			
				g_WshShell.RegWrite regKey & "\" & regValue, expectedValue, valueType
			
				errorCode = Err.Number
			End If
			 
			On Error Goto 0
	
			If errorCode = 0 Then
				msg = msg & msgOk
				CheckRegistry = True

			Else
				msg = msg & msgFail
				regStatus = 3
				CheckRegistry = False
			
			End If
		
		End If
		
		If CheckRegistry = True Then
			msg = msg & MSG_OK
		Else
			msg = msg & MSG_NOTOK
		
		End If
		
		WriteLogMsg msg, regStatus, 1, 0
	
	Next


End Function

Function  GetLocalAdminGroupName(ByRef net)
	Dim network
	Dim accounts, account
	Dim wmi
		
	GetLocalAdminGroupName = DEFAULT_LOCALADMIN_GROUP
			
	On Error Resume Next
		
	Set wmi = GetObject("winmgmts:\\" & net.ComputerName & "\root\cimv2")
		
	Set accounts = wmi.ExecQuery("Select * From Win32_Group Where LocalAccount = TRUE And SID = 'S-1-5-32-544'")
	
	For Each account In accounts
		GetLocalAdminGroupName = account.Name
	Next
		
	On Error Goto 0
	
End Function

Function CheckLocalAdmin(ByRef options)
	Dim accountName, accountNames, accounts
	Dim localAdminGroupName, localAdminGroup
	Dim errorCode, errorMsg, msg, status
	Dim network
	
	status = 1
	
	CheckLocalAdmin = True
	
	If options.Exists(OPTION_LOCALADMIN) Then
		WriteLogMsg MSG_CHECKLOCALADMIN_START, 1, 1, 1

		Set network = WScript.CreateObject("WScript.Network")
		
		accounts = options.Item(OPTION_LOCALADMIN)
				
		If options.Exists(OPTION_LOCALADMIN_GROUP) Then
			localAdminGroupName = GetOptionValue(OPTION_LOCALADMIN_GROUP, DEFAULT_LOCALADMIN_GROUP, options)
		Else
			WriteLogMsg MSG_GETLOCALADMINGROUPNAME, 1, 1, 0
			localAdminGroupName = GetLocalAdminGroupName(network)
		End If
				
		msg = MSG_CHECKLOCALADMIN_FINDLOCALGROUP & localAdminGroupName

		On Error Resume Next
		Err.Clear
		
		Set localAdminGroup = GetObject("WinNT://" & network.ComputerName & "/" & localAdminGroupName & ",group")
		
		errorCode = Err.Number
		errorMsg = Err.Description & " (" & Err.Number & ")"
		
		On Error Goto 0
		
		If errorCode = 0 Then

			msg = msg & MSG_FOUND
			WriteLogMsg msg, status, 1, 0

			accountNames = Split(accounts,",", -1, 1)
			
			For Each accountName In accountNames
			
				accountName = Trim(accountName)
			
				msg = MSG_CHECKLOCALADMIN_CHECKACCOUNT & accountName
				
				If localAdminGroup.IsMember("WinNT://" & accountName) Then
					msg = msg & MSG_CHECKLOCALADMIN_ALREADYMEMBER
					WriteLogMsg msg, 1, 1, 0
				Else
					On Error Resume Next
					Err.Clear
					
					localAdminGroup.Add ("WinNT://"& accountName)
					
					errorCode = Err.Number
					errorMsg = Err.Description & " (" & Err.Number & ")"
					
					On Error Goto 0
					
					If errorCode = 0 Then
						msg = msg & MSG_CHECKLOCALADMIN_ADDMEMBEROK
						WriteLogMsg msg, 1, 1, 0
					Else
						msg = msg & MSG_CHECKLOCALADMIN_ADDMEMBERFAIL & errorMsg
						status = 3
						WriteLogMsg msg, 3, 1, 0
					End If
				End If
			Next
			
		Else
			msg = msg & MSG_NOTFOUND
			status = 2
		
		End If
		
		If status = 1 Then
			msg = MSG_CHECKLOCALADMIN_FINISH & MSG_OK
	 		CheckLocalAdmin = True
		Else
			msg = MSG_CHECKLOCALADMIN_FINISH & MSG_NOTOK
	 		CheckLocalAdmin = False
		
		End If
		
		WriteLogMsg msg, status, 1, 0
	End If


End Function

Sub CheckCache(ByRef options)
	Dim uiResManager, cache
	Dim errorCode, errorMsg
	Dim desiredCacheSize
	
	WriteLogMsg MSG_CHECKCACHE_START, 1, 1, 1

	desiredCacheSize = CInt(GetOptionValue(OPTION_CACHESIZE, DEFAULT_CACHESIZE, options))
	WriteLogMsg MSG_CHECKCACHE_DESIREDSIZE & desiredCacheSize, 1, 1, 0

	On Error Resume Next
	Err.Clear
	
	Set uiResManager = CreateObject("UIResource.UIResourceMgr")
	errorCode = Err.Number
	errorMsg = Err.Description & " (" & Err.Number & ")"
	
	On Error Goto 0
	
	If errorCode <> 0 Then
		WriteLogMsg MSG_CHECKCACHE_CREATEFAIL & errorMsg, 2, 1, 0
		Exit Sub
	End If
	
	On Error Resume Next
	Err.Clear

	Set cache = uiResManager.GetCacheInfo
	
	errorCode = Err.Number	
	errorMsg = Err.Description & " (" & Err.Number & ")"
	
	On Error Goto 0

	If errorCode <> 0 Then
		Set uiResManager = Nothing
		WriteLogMsg MSG_CHECKCACHE_CACHEFAIL & errorMsg, 2, 1, 0
		Exit sub
	End If
	
	WriteLogMsg MSG_CHECKCACHE_SIZEIS & cache.TotalSize, 1, 1, 0

	If desiredCacheSize > 0 And cache.TotalSize <> desiredCacheSize Then

		On Error Resume Next
		Err.Clear
		
		cache.TotalSize = desiredCacheSize
		
		errorCode = Err.Number	
		errorMsg = Err.Description & " (" & Err.Number & ")"
		
		On Error Goto 0
	
		If errorCode <> 0 Then
			Set uiResManager = Nothing
			WriteLogMsg MSG_CHECKCACHE_CACHEFAIL & errorMsg, 2, 1, 0
			Exit Sub
		End If		
	
		WriteLogMsg MSG_CHECKCACHE_SETSIZE & desiredCacheSize, 1, 1, 0
	Else
		WriteLogMsg MSG_CHECKCACHE_SIZEOK, 1, 1, 0
	End If

	Set uiResManager = Nothing

End Sub

Sub CheckCacheDuringStartup(ByRef options)
	Dim cache
	Dim errorCode, errorMsg
	Dim cacheSize, desiredCacheSize
	
	WriteLogMsg MSG_CHECKCACHE_START, 1, 1, 1

	On Error Resume Next
	Err.Clear
	
	Set cache = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\ccm\SoftMgmtAgent:CacheConfig='Cache'")
	
	errorCode = Err.Number
	errorMsg = Err.Description & " (" & Err.Number & ")"

	On Error Goto 0

	If errorCode <> 0 Then
	   	WriteLogMsg MSG_CHECKCACHE_WMIFAIL & errorCode, 2, 1, 0
	    Exit Sub
	End If
	
	cacheSize = cache.Size
	
   	WriteLogMsg " *Current cache size is: " & cacheSize, 1, 1, 1

	desiredCacheSize = CInt(GetOptionValue(OPTION_CACHESIZE, DEFAULT_CACHESIZE, options))
	
	If cacheSize <> desiredCacheSize Then
		cache.Size = desiredCacheSize

		On Error Resume Next
		Err.Clear

		cache.Put_
		errorCode = Err.Number
		errorMsg = Err.Description & " (" & Err.Number & ")"
	
		On Error Goto 0
	
		If errorCode <> 0 Then
		   	WriteLogMsg MSG_CHECKCACHE_WMIWRITEFAIL & errorMsg, 2, 1, 0
		    Exit Sub
		End If
	
		WriteLogMsg MSG_CHECKCACHE_SETSIZE & desiredCacheSize, 1, 1, 0
		
	Else
		WriteLogMsg MSG_CHECKCACHE_SIZEOK & cacheSize, 1, 1, 0
	End If

    Set cache = Nothing

End Sub

Function CheckAssignment
	Dim smsClient, siteCode
	Dim errorCode
	
	WriteLogMsg MSG_CHECKASSIGNMENT_START, 1, 1, 1

	On Error Resume Next
	Err.Clear

	Set smsClient = CreateObject ("Microsoft.SMS.Client")
	
	errorCode = Err.Number
	
	siteCode = smsClient.GetAssignedSite
	
	On Error Goto 0
	
	If Len(siteCode) = 0 Or errorCode <> 0 Then
		WriteLogMsg MSG_CHECKASSIGNMENT_NOTOK & siteCode, 1, 1, 0
		CheckAssignment = False
	Else
		WriteLogMsg MSG_CHECKASSIGNMENT_OK & siteCode, 1, 1, 0
		CheckAssignment = True
	End If

	Set smsClient = Nothing

End Function

Function CheckClient(ByRef options)
	Dim wmi, ccmWMI, errorCode, errorMsg
	Dim clientProperties, clientProp
	Dim clientVersion, expectedVersion, expectedUnpatchedversion
	Dim configMgrLogPath, configMgrLogFilePath, configMgrOldLogFilePath
	
	clientVersion = "0"
	configMgrLogPath = ""
	
	WriteLogMsg MSG_CHECKCLIENT_START, 1, 1, 0

	On Error Resume Next
	
	Err.Clear

    Set ccmWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\ccm")
   
    errorCode = Err.Number
    errorMsg = Err.Description & " (" & Err.Number & ")"
    
   	On Error Goto 0
    
    If errorCode <> 0 Then
    	WriteLogMsg MSG_CHECKCLIENT_WMINOTFOUND & errorMsg, 2, 1, 0
    	CheckClient = False
		Exit Function
    End If
    
   	WriteLogMsg MSG_CHECKCLIENT_VERSION, 1, 1, 0
   
   	Set clientProperties = ccmWmi.ExecQuery("Select * from SMS_Client")
   	
   	On Error Resume Next
		
	For Each clientProp In clientProperties
		clientVersion=clientProp.ClientVersion
	Next
	
	On Error Goto 0
	
	If clientVersion = "0" Then
    	WriteLogMsg MSG_CHECKCLIENT_VERSIONNOTFOUND & errorMsg, 2, 1, 0
		CheckClient = False
		Exit Function
	End If
	
	expectedVersion = GetOptionValue(OPTION_AGENTVERSION, DEFAULT_AGENTVERSION, options)
	expectedUnpatchedversion = GetOptionValue(OPTION_UAGENTVERSION, DEFAULT_UAGENTVERSION, options)
	
	WriteLogMsg MSG_CHECKCLIENT_VERSIONEXPECTED & expectedVersion, 1, 1, 0
	
	If CheckClientVersion(expectedVersion, clientVersion) = False Then
		If expectedUnpatchedversion <> DEFAULT_UAGENTVERSION And CheckClientVersion(expectedUnpatchedversion, clientVersion) = False Then
			WriteLogMsg MSG_CHECKCLIENT_OLDVERSION & clientVersion, 2, 1, 0
			CheckClient = False
			Exit Function
		Else
			WriteLogMsg MSG_CHECKCLIENT_UNPATCHEDVERSION & clientVersion, 2, 1, 0
			
			Exit Function
		End If
	Else
	 	WriteLogMsg MSG_CHECKCLIENT_VERSIONFOUND & clientVersion, 1, 1, 0
	End If
		
	WriteLogMsg MSG_CHECKCLIENT_CCMEXEC, 1, 1, 0
	
  	Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

	If CheckService(wmi, "CCMExec", "Running", "Auto", "True") = False Then
	   	CheckClient = False
	   	Exit Function
	End If
		
	CheckClient = True
	
	On Error Resume Next
	Err.Clear
	
	configMgrLogPath = g_WshShell.RegRead("HKLM\SOFTWARE\Wow6432Node\Microsoft\SMS\Client\Configuration\Client Properties\Local SMS Path")
	
	errorCode = Err.Number

	If errorCode <> 0 Then
		Err.Clear
		
		configMgrLogPath = g_WshShell.RegRead("HKLM\SOFTWARE\Microsoft\SMS\Client\Configuration\Client Properties\Local SMS Path")
		errorCode = Err.Number
		errorMsg = Err.Description & " (" & Err.Number & ")"
	End If

	If errorCode = 0 Then
		configMgrLogFilePath = configMgrLogPath & "Logs\" & WScript.ScriptName & ".log"
		
		If configMgrLogFilePath <> g_logPathandName Then
		
			Err.Clear
			
			If g_fso.FileExists(configMgrLogFilePath) Then
				Dim timedate, timedateString
				timedate = Now
				timedateString = Year(timedate) & Right("0" & Month(timedate), 2) & Right("0" & Day(timedate), 2) & Right("0" & Hour(timedate), 2) & Right("0" & Minute(timedate), 2)

				configMgrOldLogFilePath = configMgrLogPath & "Logs\" & WScript.ScriptName & "-" & timedateString & ".log"
				g_fso.MoveFile configMgrLogFilePath, configMgrOldLogFilePath
			End If

			g_fso.MoveFile g_logPathandName, configMgrLogFilePath
			
			errorCode = Err.Number
			
			If errorCode = 0 Then
				g_logPathandName = configMgrLogFilePath
				WriteLogMsg MSG_CHECKCLIENT_MOVEDLOG & configMgrLogFilePath, 1, 1, 0
			Else
				WriteLogMsg MSG_CHECKCLIENT_MOVELOGFAIL & errorMsg, 2, 1, 0
			End If			
		End If
	Else
		WriteLogMsg MSG_CHECKCLIENT_GETLOGDIRECTORY & errorMsg, 3, 1, 0
	End If
	
	On Error Goto 0
	
	CheckClient = CheckAssignment
    
    Set wmi = Nothing
    Set ccmWMI = Nothing
End Function

Function CheckClientVersion(ByVal expectedVersion, ByVal currentVersion)

	Dim currentVersionArray, expectedVersionArray
	Dim versionPartCount
	
	CheckClientVersion = True
	
	currentVersionArray = Split(currentVersion, ".", -1, 1)
	expectedVersionArray = Split(expectedVersion, ".", -1, 1)
	
	For versionPartCount=0 To 3
		If currentVersionArray(versionPartCount) > expectedVersionArray(versionPartCount) Then
			Exit Function
		ElseIf currentVersionArray(versionPartCount) < expectedVersionArray(versionPartCount) Then
		
			CheckClientVersion = False
			Exit Function
		End If
	Next

End Function

Function InstallClient(ByRef options, ByRef properties, ByRef parameters)
	Dim fsp, mp, slp, cacheSize, siteCode
	Dim possiblePaths, possiblePathCount, possiblePathIndexStart, possiblePathIndex, pathFound
	Dim commandLine
	Dim returnCode
	Dim param, paramValue
	Dim prop, propValue
	Dim installPatchProperty
	Dim errorCode, errorMsg, msg
	Dim osArchitecture, expectedClientMajorVersion, hfDirectory
		
	installPatchProperty = ""

	WriteLogMsg MSG_INSTALLCLIENT_START, 1, 1, 0

	If options.Exists(OPTION_INSTALLPATH) And Len(Options.Item(OPTION_INSTALLPATH)) <> 0 Then
	
		cacheSize = GetOptionValue(OPTION_CACHESIZE, DEFAULT_CACHESIZE, options)
		
		pathFound = False
		
		possiblePaths = Split(Options.Item(OPTION_INSTALLPATH), ";", -1, 1)
		
		possiblePathCount = UBound(possiblePaths)
		
		Randomize
		possiblePathIndexStart = CInt(Rnd * possiblePathCount)
		possiblePathIndex = possiblePathIndexStart
		
		While (possiblePathIndex <= possiblePathCount And pathFound = False)

			commandLine = possiblePaths(possiblePathIndex) & "\ccmsetup.exe"
			msg = MSG_INSTALLCLIENT_PATHCHECK & possiblePaths(possiblePathIndex)
			
			If g_fso.FileExists(commandLine) Then
				WriteLogMsg msg & MSG_FOUND, 1, 1, 0
				pathFound = True
			Else
			
				WriteLogMsg msg & MSG_NOTFOUND & " " & commandLine, 2, 1, 0
			
				If possiblePathIndex = possiblePathCount Then
					possiblePathIndex = 0
				Else
					possiblePathIndex = possiblePathIndex + 1
				End If
				
				If possiblePathIndex = possiblePathIndexStart Then
					possiblePathIndex = possiblePathCount + 1
				End If
			
			End If
			
		Wend

		If pathFound = False Then
						
			InstallClient = False
			Exit Function
			
		End If

		commandLine = """" & commandLine & """"

		For Each param In parameters.Keys
			
			If Len(parameters.Item(param)) > 0 Then
				commandLine = commandLine & " /" & param & ":" & parameters.Item(param)
			Else
				commandLine = commandLine & " /" & param
			End If
			
		Next

		If cacheSize > 0 Then
			commandLine = commandLine & " SMSCACHESIZE=" & cacheSize
		End If

		If options.Exists(OPTION_SITECODE) Then
			commandLine = commandLine & " SMSSITECODE=" & options.Item(OPTION_SITECODE)
		End If
		
		For Each prop In properties.Keys
			If InStr(1, prop, "PATCH", 1) Then
				
				If InStr(1, properties.Item(prop), ";", 1) Then
					WriteLogMsg MSG_HOTFIX_MULTIPLE, 1, 1, 0
					installPatchProperty = installPatchProperty & properties.Item(prop)
				Else
				
					msg = MSG_HOTFIX_FILEVERIFY & properties.Item(prop)
	
					If g_fso.FileExists(properties.Item(prop)) Then				
						If Len(installPatchProperty) > 0 Then
							installPatchProperty = installPatchProperty & ";"
						End If
	
						installPatchProperty = installPatchProperty & properties.Item(prop)
	
						WriteLogMsg msg & MSG_FOUND, 1, 1, 0				
					Else
						WriteLogMsg msg & MSG_NOTFOUND & ": " & errorMsg, 2, 1, 0				
					End If
				End If
			Else
				commandLine = commandLine & " " & prop & "=" & properties.Item(prop)
			End If
		Next
		
		If options.Exists(OPTION_AUTOHOTFIX) Then
			If Len(installPatchProperty) > 0 Then
				installPatchProperty = installPatchProperty & ";"
			End If
			
			hfDirectory = Trim(options.Item(OPTION_AUTOHOTFIX))
			
			If Len(hfDirectory) > 0 Then
			
				expectedClientMajorVersion = Left(GetOptionValue(OPTION_AGENTVERSION, DEFAULT_AGENTVERSION, options),1)

				If expectedClientMajorVersion = "5" Then
				
					If Right(hfDirectory, 1) <> "\" Then
						hfDirectory = hfDirectory & "\"
					End If
				
					osArchitecture = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth
					
					WriteLogMsg MSG_OS_ARCHITECTURE & osArchitecture, 1, 1, 0
		
					If osArchitecture = "32" Then
						hfDirectory = hfDirectory & "i386"
					ElseIf osArchitecture = "64" Then
						hfDirectory = hfDirectory & "x64"
					End If
					
				End If
				
				AutoHotfix hfDirectory, installPatchProperty
			
			End If
		End If

		If Len(installPatchProperty) > 0 Then
			commandLine = commandLine & " PATCH=""" & installPatchProperty & """"
		End If
		
		WriteLogMsg MSG_INSTALLCLIENT_COMMANDLINE & commandLine, 1, 1, 0
		
		g_WshShell.Environment("PROCESS").Item("SEE_MASK_NOZONECHECKS") = 1
		returnCode = g_WshShell.Run(commandLine, 0, True)
		'returnCode = 0
		
		g_WshShell.Environment("PROCESS").Remove("SEE_MASK_NOZONECHECKS")
		If returnCode <> 0 Then
		  WriteLogMsg MSG_INSTALLCLIENT_FAILED & returnCode, 3, 1, 0
		  InstallClient = False
		Else
		  WriteLogMsg MSG_INSTALLCLIENT_SUCCESS, 1, 1, 0
		  InstallClient = True
    	End If

	End If
	
End Function

Function GetOptionValue(ByVal optionName, ByVal defaultValue, ByRef options)

	If options.Exists(optionName) Then
		GetOptionValue = options.Item(optionName)
	Else
		GetOptionValue = defaultValue
	End If
	
End Function

Sub AutoHotfix(ByVal hotfixDirectory, ByRef patchProperty)

	Dim errorCode, errorMsg
	Dim hfDir

	WriteLogMsg MSG_AUTOPATCH_COMMANDLINE & hotfixDirectory, 1, 1, 0

    On Error Resume Next
	Err.Clear

	Set hfDir = g_fso.GetFolder(hotfixDirectory)
	errorCode = Err.Number	
	errorMsg = Err.Description & " (" & Err.Number & ")"
	
	On Error Goto 0

	If errorCode <> 0 Then
	   WriteLogMsg MSG_AUTOPATCH_DIRERROR & errorMsg, 2, 1, 0
	   Exit Sub
	End If
	
	FindHotfixes hfDir, patchProperty
	
End Sub

Sub FindHotfixes(ByVal directory, ByRef patchProperty)

	Dim file, subfolder
	
    For Each file In directory.Files
    
    	If LCase(g_fso.GetExtensionName(file)) = "msp" Then
    	
    		If InStr(1, patchProperty, file.Name, 1) Then
    			WriteLogMsg MSG_HOTFIX_DUPLICATE & file.Name, 2, 1, 0
			Else    		
    	
	    		If Len(patchProperty) > 0 Then
	    			patchProperty = patchProperty & ";"
	    		End If
	    		
	    		WriteLogMsg MSG_AUTOPATCH_FOUNDHOTFIX & file.Path, 1, 1, 0
	    		
	    		patchProperty = patchProperty & """" & file.Path & """"
			End If    	
    	End If
    
    Next
    
	For Each subfolder In directory.subfolders
		FindHotfixes subfolder, patchProperty
	Next

End Sub

Function InWinPE
	Dim sysEnv, systemDrive
	Set sysEnv = g_WshShell.Environment("PROCESS")
	systemDrive = sysEnv("SYSTEMDRIVE")
	
	InWinPE = Eval(SystemDrive = "X:")
End Function

Function IsAdmin
 
  On Error Resume Next
  
  CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
  
  If err.number = 0 Then 
  	IsAdmin = True
  Else
  	IsAdmin = False
  	WScript.Echo "User is not a local admin."

  End If
  
  On Error Goto 0
  
End Function