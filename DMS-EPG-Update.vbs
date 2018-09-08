'General script options, dimming and setting public variables
'Generelle Script-Optionen, Erstellen und Setzen von öffentlichen Variablen
Option Explicit
Const ScriptVersion="2017.04.12.16.00"

'Protocol, IP/name and port of DVBViewer Media Server
'Protokoll, IP/Name und Port des DVBViewer Media Server
'Default/Standard: "http://127.0.0.1:8089"
dim ServiceURL
ServiceURL="http://127.0.0.1:8089"

'Name of the DVBViewer Media Server user account. No user name is "".
'Name des DVBViewer Media Server Benutzers. Kein Benutzername ist "".
'Default/Standard: ""
dim ServiceUser
ServiceUser=""

'Password of the DVBViewer Media Server user account. No password is "".
'Passwort des DVBViewer Media Server Benutzers. Kein Passwort ist "".
'Default/Standard: ""
dim ServicePassword
ServicePassword=""

'What language file should be used?
'   Script tries to auto detect language and looks for the appropriate file. 
'   If file does not exist, script tries to use "en.ini" hardcoded. If "en.ini" also does not exist, script ends with an error.
'Welche Sprachdatei soll verwendet werden?
'   Das Script versucht die Sprache selbst festzustellen und sucht eine entsprechende Sprachdatei.
'   Wenn die Datei nicht existiert, wird fix "en.ini" verwendet. Sollte "en.ini" auch nicht existieren, endet das Script mit einem Fehler.
'Example/Beispiel: "de.ini"
'Default/Standard: ""
dim LanguageFile
LanguageFile=""

'LogFile, deactivated is "".
'Log-Datei, deaktiviert ist "".
'Default/Standard: "DMS-EPG-Update.log"
dim LogFile
LogFile="DMS-EPG-Update.log"

'Runs to keep in the logfile.
'   "-1": All entries are kept.
'   "0": Only the last run is kept in the log.
'   Values are rounded up to integer values. "-0,1" becomes "0" etc.
'   Invalid values are handled as "-1".
'Anzahl der Durchläufe, die in der Log-Datei aufbewahrt werden sollen.
'   "-1": Alle Durchläufe werden aufbewahrt.
'   "0": Nur der letzte Lauf wird aufbewahrt.
'   Werte werden auf Integer-Zahlen aufgerunden. "-0,1" wird "0" etc.
'   Ungültige Werte werden als "-1" behandelt.
'Default/Standard: 10
dim RunsToKeepInLog
RunsToKeepInLog=10

'What should the script do after the EPG update?
'Was soll das Skript nach dem EPG-Update tun?
'"Hibernate", "Standby", "Shutdown" or/oder ""
'Default/Standard: "Standby"
dim ActionAfterEPGUpdate
ActionAfterEPGUpdate="Standby"

'If EPG update task is started but queued: How many minutes should the script wait to see if the task can be started?
'Wenn der EPG-Update-Task gestartet wurde, sich aber in der Warteschleife befindet: Wie viele Minuten soll das Script auf den Start des Tasks warten?
'Default/Standard: 15
dim TimeToWaitForEPGUpdateStart
TimeToWaitForEPGUpdateStart=15

'If ActionAfterEPGUpdate<>"": How many minutes should the script wait until standby prerequisites are met?
'Wenn ActionAfterEPGUpdate<>"": Wie viele Minuten soll das Script auf die Erfüllung der Standby-Voraussetzungen warten?
'Default/Standard: 15
dim TimeToWaitForActionAfterEPGUpdatePrerequisites
TimeToWaitForActionAfterEPGUpdatePrerequisites=15

'Time in seconds to wait at start
'Zeit in Sekunden, die beim Start gewartet werden soll
'Default/Standard: 30
dim WaitBeforeStart
WaitBeforeStart=30

'Fill missing EPG entries today and tomorrow
'Fehlende EPG-Einträge heute und morgen ergänzen
'true or/oder false.
'Default/Standard: true
dim CreateEPGEntry
CreateEPGEntry=true

'Delete all EPG entries before starting EPG update
'Alle EPG-Einträge vor dem EPG-Update löschen
'true or/oder false.
'Default/Standard: false
dim ClearEPG
ClearEPG=false


'Internal variables - do not change!
'Interne Variablen - nicht verändern!
dim arrtemp
dim BeginWaitforEPGUpdateStarted, BeginWaitforActionAfterEPGUpdatePrerequisitesMet, BeginWaitforActionAfterEPGUpdatePrerequisitesMetCount
dim EPGUpdateFinished, EPGUpdateQueued
dim fso, fsofile
dim getvaluefrominifiletempstring
dim inifile, iniFileObject, inifiletouse, inifileinfostring, intEqualPos
dim KeepFromLine
dim LogMsgString, logarray, line, LanguageInfoString, LanguageFileObject, LanguageFileDefault, LCIDtoUse, LCIDSplit, LCIDDictionary, LCIDsComputer, LCIDoWMI, LCIDcolOperatingSystems, LCIDoOS, LCIDiOSLang
dim objFSOini, oHTTP, objargs, objinifile
dim ReadyForStandby, RunsFoundSearchText, RunsFound, ReturnString
dim ScriptStartTime, sectionname
dim strFilePath, strSection, strKey, strline, strleftstring
dim Tag, TargetVariableNameTemp, tempargumentname, tsInput
dim w, x, y, z, EPGChannelID, EPGXMLData, EPGXMLDate, a, b, c
dim xmldoc, colnodes, objnode
dim xmldocb, colnodesb, objnodeb
Dim newepgstart, newepgstop
dim epgquerystart, epgqueryend


'Configuring internal variables
EPGUpdateFinished=false
EPGUpdateQueued=false
ReadyForStandby=false
Set LCIDDictionary = CreateObject("Scripting.Dictionary")
LanguageFileDefault="en.ini"
ScriptStartTime=now
SectionName="default"

'Check cscript.exe
If "CSCRIPT.EXE" <> UCase(Right(WScript.Fullname, 11)) Then
	msgbox "Script must be started with cscript.exe, not with wscript.exe.", vbOKOnly+vbCritical, "Error"
	LogMsg("***** Start *****")
	LogMsg("Script start time: " & DatePart("yyyy", ScriptStartTime) & "-" & Right("0" & DatePart("m", ScriptStartTime), 2) & "-" & Right("0" & DatePart("d", ScriptStartTime), 2) & " " & Right("0" & DatePart("h", ScriptStartTime), 2) & ":" & Right("0" & DatePart("n", ScriptStartTime), 2) & ":" & Right("0" & DatePart("s", ScriptStartTime), 2) & ".")
	LogMsg("Script version """ & ScriptVersion & """.")
	LogMsg "Script must be started with cscript.exe, not with wscript.exe."
	LogMsg("***** End *****")
	wscript.quit 1
End If

'Weekday
select case weekday(now, 2)
	case 1
		Tag="Monday"
	case 2
		Tag="Tuesday"
	case 3
		Tag="Wednesday"
	case 4
		Tag="Thursday"
	case 5
		Tag="Friday"
	case 6
		Tag="Saturday"
	case 7
		Tag="Sunday"
end select

'Check arguments for ini file
Set iniFileObject = CreateObject("Scripting.FileSystemObject")
iniFile=""
iniFileToUse=""
Set objArgs = WScript.Arguments
For x = 0 to (objArgs.Count-1)
	if left(objArgs(x),2)="--" then
		TempArgumentName=right(objArgs(x),len(objArgs(x))-2)
	elseif left(objArgs(x),1)="/" or left(objArgs(x),1)="-" then
		TempArgumentName=right(objArgs(x),len(objArgs(x))-1)
	else
		TempArgumentName=objArgs(x)
	end if
	if lcase(left(lcase(tempargumentname), len("ini:")))=lcase("ini:") then
		arrtemp=split(TempArgumentName,":")
		if ubound(arrtemp)=1 then inifile=arrtemp(1)
		if ubound(arrtemp)=2 then inifile=arrtemp(1) & ":" & arrtemp(2)
		if ubound(arrtemp)>2 then
			inifile=""
			inifileInfoString="More than two "":"" passed in file path part of parameter /ini." & vbcrlf & "First, read the file ""readme.txt""," & vbcrlf & "then try ""cscript.exe DMS-EPG-Update.vbs /ini:sample.ini""."
		end if
	end if
	TempArgumentName=""
Next

if inifile="" then
	if inifileInfoString<>"" then
		wscript.echo inifileInfoString
	else
		wscript.echo "Ini file parameter has not been passed to the script or does not contain a file name, exiting." & vbcrlf & "First, read the file ""readme.txt""," & vbcrlf & "then try ""cscript.exe DMS-EPG-Update.vbs /ini:sample.ini""."
	end if
	wscript.quit 1
else
	if iniFileObject.FileExists(inifile) then
		iniFileToUse=iniFile
	else
		wscript.echo "Ini file not found, exiting." & vbcrlf & "First, read the file ""readme.txt""," & vbcrlf & "then try ""cscript.exe DMS-EPG-Update.vbs /ini:sample.ini""."
		wscript.quit 1
	end if
end if
inifile=inifiletouse

'Get language and log file name
LanguageFile=GetValueFromIniFile("LanguageFile", LanguageFile)
LogFile=GetValueFromIniFile("LogFile", LogFile)

'Check language file
Call FillLCIDDictionary
LCIDsComputer = "."
Set LCIDoWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & LCIDsComputer & "\root\cimv2")
Set LCIDcolOperatingSystems = LCIDoWMI.ExecQuery ("Select * from Win32_OperatingSystem")

For Each LCIDoOS in LCIDcolOperatingSystems
	LCIDiOSLang = LCIDoOS.OSLanguage
Next

Set LanguageFileObject = CreateObject("Scripting.FileSystemObject")

if LanguageFile="" then
	'Detecting system language
	If LCIDDictionary.Exists(getlocale) then
		LCIDtoUse=LCIDiOSLang
		LCIDSplit=split(LCIDDictionary.Item(LCIDtoUse),";")
		LanguageInfoString="Locale ID (LCID): " & LCIDtoUse & " (" & LCIDSplit(0) & ", " & LCIDSplit(1) & ", " & LCIDSplit(2) & ")."
	else
		LCIDtoUse=1033
		LCIDSplit=split(LCIDDictionary.Item(LCIDtoUse),";")
		LanguageInfoString="Locale ID " & getlocale & " unknown. Defaulting to LCID 1033 (" & LCIDSplit(0) & ", " & LCIDSplit(1) & ", " & LCIDSplit(2) & ")."
	end if
	'Check files
	if LanguageFileObject.FileExists(lcidsplit(1)) then
		LanguageFile=lcidsplit(1)
	elseif LanguageFileObject.FileExists(lcidsplit(2)) then
		LanguageFile=lcidsplit(2)
	else
		LanguageFile=LanguageFileDefault
		if LanguageFileObject.FileExists(languagefiledefault) then
			LanguageFile=LanguageFileDefault
		else
			LogMsg("***** Start *****")
			LogMsg("Script start time: " & DatePart("yyyy",ScriptStartTime) & "-" & Right("0" & DatePart("m",ScriptStartTime), 2) & "-" & Right("0" & DatePart("d",ScriptStartTime), 2) & " " & Right("0" & DatePart("h",ScriptStartTime), 2) & ":" & Right("0" & DatePart("n",ScriptStartTime), 2) & ":" & Right("0" & DatePart("s",ScriptStartTime), 2) & ".")
			LogMsg("Script version """ & ScriptVersion & """.")
			LogMsg("Default language file """ & languagefiledefault & """ not found, exiting.")
			LogMsg("***** End *****")
			wscript.quit 1
		end if
	end if
else
	if LanguageFileObject.FileExists(languagefile) then
		'do nothing
	else
		LanguageInfoString="Language file """ & languagefile & """ not found, using default value """ & LanguageFileDefault & """."
		if LanguageFileObject.FileExists(languagefiledefault) then
			LanguageFile=LanguageFileDefault
		else
			LogMsg("***** Start *****")
			LogMsg("Script start time: " & DatePart("yyyy",ScriptStartTime) & "-" & Right("0" & DatePart("m",ScriptStartTime), 2) & "-" & Right("0" & DatePart("d",ScriptStartTime), 2) & " " & Right("0" & DatePart("h",ScriptStartTime), 2) & ":" & Right("0" & DatePart("n",ScriptStartTime), 2) & ":" & Right("0" & DatePart("s",ScriptStartTime), 2) & ".")
			LogMsg("Script version """ & ScriptVersion & """.")
			LogMsg(LanguageInfoString)
			LogMsg("Default language file """ & languagefiledefault & """ not found, exiting.")
			LogMsg("***** End *****")
			wscript.quit 1
		end if
	end if
end if

'Get values from ini file
ActionAfterEPGUpdate=GetValueFromIniFile("ActionAfterEPGUpdate", ActionAfterEPGUpdate)
RunsToKeepInLog=GetValueFromIniFile("RunsToKeepInLog", RunsToKeepInLog)
ServiceURL=GetValueFromIniFile("ServiceURL", ServiceURL)
ServiceUser=GetValueFromIniFile("ServiceUser", ServiceUser)
ServicePassword=GetValueFromIniFile("ServicePassword", ServicePassword)
TimeToWaitForEPGUpdateStart=GetValueFromIniFile("TimeToWaitForEPGUpdateStart", TimeToWaitForEPGUpdateStart)
TimeToWaitForActionAfterEPGUpdatePrerequisites=GetValueFromIniFile("TimeToWaitForActionAfterEPGUpdatePrerequisites", TimeToWaitForActionAfterEPGUpdatePrerequisites)
WaitBeforeStart=GetValueFromIniFile("WaitBeforeStart", WaitBeforeStart)
CreateEPGEntry=GetValueFromIniFile("CreateEPGEntry", CreateEPGEntry)
ClearEPG=GetValueFromIniFile("ClearEPG", ClearEPG)


'Check ActionAfterEPGUpdate values
select case lcase(ActionAfterEPGUpdate)
	case "hibernate"
		'ok, do nothing
	case "standby"
		'ok, do nothing
	case "shutdown"
		'ok, do nothing
	case ""
		'ok, do nothing
	case else
		'set default value
		LogMsg(LanguageGetLine1Var(072, ActionAfterEPGUpdate))
		ActionAfterEPGUpdate="Standby"
end select


'Start with prerequisites met
LogMsg("***** " & LanguageGetLine0Var(003) & " *****")
LogMsg(LanguageGetLine1Var(001, DatePart("yyyy",ScriptStartTime) & "-" & Right("0" & DatePart("m",ScriptStartTime), 2) & "-" & Right("0" & DatePart("d",ScriptStartTime), 2) & " " & Right("0" & DatePart("h",ScriptStartTime), 2) & ":" & Right("0" & DatePart("n",ScriptStartTime), 2) & ":" & Right("0" & DatePart("s",ScriptStartTime), 2)))
LogMsg(LanguageGetLine1Var(004, ScriptVersion))

'Display final settings
LogMsg(LanguageGetLine0Var(018))
LogMsg("  LanguageFile=""" & LanguageFile & """")
LogMsg("  LogFile=""" & LogFile & """")
LogMsg("  RunsToKeepInLog=" & RunsToKeepInLog)
LogMsg("  WaitBeforeStart=" & WaitBeforeStart)
LogMsg("  ServiceURL=""" & ServiceURL & """")
LogMsg("  ServiceUser=""" & ServiceUser & """")
LogMsg("  ServicePassword=""" & ServicePassword & """")
LogMsg("  TimeToWaitForEPGUpdateStart=" & TimeToWaitForEPGUpdateStart)
LogMsg("  ActionAfterEPGUpdate=" & ActionAfterEPGUpdate)
LogMsg("  TimeToWaitForActionAfterEPGUpdatePrerequisites=" & TimeToWaitForActionAfterEPGUpdatePrerequisites)
LogMsg("  CreateEPGEntry=" & CreateEPGEntry)
LogMsg("  ClearEPG=" & ClearEPG)

if LanguageInfoString<>"" then LogMsg(LanguageInfoString)
LogMsg(LanguageGetLine1Var(019, WaitBeforeStart))
'wscript.sleep(WaitBeforeStart*1000)


'CleanLogFile
call CleanLogFile

'Test connection to DVBViewer Media Server and authentication
ReturnString=HTTPGet(ServiceURL & "/api/status2.html", "")
If instr(1, ReturnString, "<epgudate>") > 0 then
	'authentication works, we get data
else
	LogMsg(LanguageGetLine0Var(070))
	LogMsg("***** " & LanguageGetLine0Var(011) & " *****")
	wscript.quit 1
end if


'Clear EPG
if ClearEPG=true then
	LogMsg(LanguageGetLine0Var(016))
	HTTPGet ServiceURL & "/api/epgclear.html", ""
	wscript.sleep(5000)
end if


'Start EPG update task
LogMsg(LanguageGetLine0Var(020))
HTTPGet ServiceURL & "/tasks.html", "task=EPGStart&aktion=tasks"
wscript.sleep(5000)


'Wait for EPG update completion
BeginWaitforEPGUpdateStarted=now
EPGUpdateQueued=false
EPGUpdateFinished=false
do until EPGUpdateFinished=true
	ReturnString=HTTPGet(ServiceURL & "/api/status2.html", "")
	if instr(1, ReturnString, "<epgudate>0</epgudate>")>0 then
		'EPG update complete
		LogMsg(LanguageGetLine0Var(021))
		EPGUpdateFinished=true
	elseif instr(1, ReturnString, "<epgudate>1</epgudate>")>0 then
		'EPG update still running
		LogMsg(LanguageGetLine0Var(023))
		wscript.sleep(60000)
	elseif instr(1, ReturnString, "<epgudate>2</epgudate>")>0 then
		'EPG update queued
		EPGUpdateQueued=true
		LogMsg(LanguageGetLine0Var(022))
		LogMsg(LanguageGetLine1Var(024, TimeToWaitForEPGUpdateStart))
		do until datediff("s", BeginWaitforEPGUpdateStarted, now)=>TimeToWaitForEPGUpdateStart*60 _
		OR EPGUpdateQueued=false
			ReturnString=HTTPGet(ServiceURL & "/api/status2.html", "")
			if instr(1, ReturnString, "<epgudate>2</epgudate>")>0 then
				LogMsg(LanguageGetLine0Var(025))
				wscript.sleep(60000)
			else
				EPGUpdateQueued=false
				exit do
			end if
		loop
		if datediff("s", BeginWaitforEPGUpdateStarted, now)=>TimeToWaitForEPGUpdateStart*60 then
			exit do
		end if
	else
		'Unknown <epgudate> value
		LogMsg(LanguageGetLine1Var(071, mid(ReturnString,instr(1, ReturnString, "<epgudate>")+10, instr(1, ReturnString, "</epgudate>")-instr(1, ReturnString, "<epgudate>")-10)))
		LogMsg("***** " & LanguageGetLine0Var(011) & " *****")

		wscript.quit 1
	end if
loop


'End if EPG Update is queued and TimeToWaitForEPGUpdateStart timed out
if EPGUpdateQueued=true then
	LogMsg(LanguageGetLine1Var(026, TimeToWaitForEPGUpdateStart))
	LogMsg(LanguageGetLine0Var(030))
	HTTPGet ServiceURL & "/tasks.html", "task=AutoTimer&aktion=tasks"
	LogMsg("***** " & LanguageGetLine0Var(011) & " *****")
	wscript.quit 1
end if


'Create EPG entries
if CreateEPGEntry=true then
	'HTTPGet ServiceURL & "/api/epgclear.html", "source=4"
	LogMsg(LanguageGetLine0Var(017))
	epgquerystart=now()
	epgquerystart=dateadd("h", hour(epgquerystart)*-1, epgquerystart)
	epgquerystart=dateadd("n", minute(epgquerystart)*-1, epgquerystart)
	epgquerystart=dateadd("s", second(epgquerystart)*-1, epgquerystart)
	epgqueryend=epgquerystart
	epgqueryend=dateadd("d", 1, epgqueryend)
	epgqueryend=dateadd("h", 23, epgqueryend)
	epgqueryend=dateadd("n", 59, epgqueryend)
	epgqueryend=dateadd("s", 59, epgqueryend)
	EPGXMLData="<?xml version=""1.0"" encoding=""ISO-8859-1""?> <!DOCTYPE tv SYSTEM ""xmltv.dtd""><tv source-info-name=""DVBViewer Media Server EPG Update Script"" generator-info-name=""DVBViewer Media Server EPG Update Script"">" &_
		"<programme start=""$start$" & "00"" stop=""$stop$" & "00"" channel=""$EPGChannelID$""><charset>255</charset><title>$title$</title><description>DVBViewer Media Server EPG Update Script</description></programme></tv>"
	Set xmlDoc = CreateObject("Msxml2.DOMDocument")
	xmlDoc.Async = false
	xmlDoc.validateOnParse = false
	xmlDoc.resolveExternals = false
	xmlDoc.preserveWhiteSpace = false
	xmlDoc.Load ServiceURL & "/api/getchannelsxml.html"
	Set colNodes = xmlDoc.selectNodes("/channels/root/group/channel")
	c=0
	a=right("0000" & year(epgquerystart), 4) & right("00" & month(epgquerystart), 2) & right("00" & day(epgquerystart), 2) & right("00" & hour(epgquerystart), 2) & right("00" & minute(epgquerystart), 2) & right("00" & second(epgquerystart), 2)
	b=right("0000" & year(epgqueryend), 4) & right("00" & month(epgqueryend), 2) & right("00" & day(epgqueryend), 2) & right("00" & hour(epgqueryend), 2) & right("00" & minute(epgqueryend), 2) & right("00" & second(epgqueryend), 2)
	For Each objNode in colNodes
		c=c+1
		wscript.stdout.write chr(13) & string(79," ") & chr(13) & left(right("00000" & c, 5) & "/" & right("00000" & colNodes.length, 5) & ", EPGID " & objNode.Attributes.getNamedItem("EPGID").Text & ", " & objNode.Attributes.getNamedItem("name").Text, 79) & chr(13)
		Set xmlDocb = CreateObject("Msxml2.DOMDocument")
		xmlDocb.Async = false
		xmlDocb.validateOnParse = false
		xmlDocb.resolveExternals = false
		xmlDocb.preserveWhiteSpace = false
		xmlDocb.Load ServiceURL & "/api/epg.html?lvl=2&channel=" & objNode.Attributes.getNamedItem("EPGID").Text & "&start=" & cdbl(epgquerystart) & "&end=" & cdbl(epgqueryend)
		Set colNodesb = xmlDocb.selectNodes("/epg/programme")
		if colnodesb.length = 0 then
			HTTPPut ServiceURL & "/cgi-bin/EPGimport", replace(replace(replace(replace(EPGXMLData, "$start$", a), "$stop$", b), "$EPGChannelID$", objNode.Attributes.getNamedItem("EPGID").Text), "$title$", objNode.Attributes.getNamedItem("name").Text)
		end if
	Next
end if
wscript.stdout.write chr(13) & string(79," ") & chr(13)


'Start AutoTimer task
LogMsg(LanguageGetLine0Var(030))
HTTPGet ServiceURL & "/api/tasks.html", "task=AutoTimer"
wscript.sleep(35000)


'Standby
BeginWaitforActionAfterEPGUpdatePrerequisitesMet=now
BeginWaitforActionAfterEPGUpdatePrerequisitesMetCount=0
if ActionAfterEPGUpdate<>"" then
	LogMsg(LanguageGetLine0Var(035))
	do until ReadyForStandby=true _
	OR datediff("s", BeginWaitforActionAfterEPGUpdatePrerequisitesMet, now)=>TimeToWaitForActionAfterEPGUpdatePrerequisites*60
		ReturnString=HTTPGet(ServiceURL&"/api/status2.html", "")
		if instr(1, ReturnString, "<reccount>0</reccount>")>0 _
		AND instr(1, ReturnString, "<streamclientcount>0</streamclientcount>")>0 _
		AND instr(1, ReturnString, "<epgudate>0</epgudate>")>0 then
			ReadyForStandby=true
			'Enter standby
			LogMsg(LanguageGetLine0Var(036))
			HTTPGet ServiceURL & "/tasks.html", "task=" & ActionAfterEPGUpdate & "&aktion=tasks"
		else
			'Prerequisites not met, try again
			if BeginWaitforActionAfterEPGUpdatePrerequisitesMetCount=0 then
				LogMsg(LanguageGetLine1Var(038, TimeToWaitForActionAfterEPGUpdatePrerequisites))
			end if
			LogMsg(LanguageGetLine0Var(037))
			BeginWaitforActionAfterEPGUpdatePrerequisitesMetCount=BeginWaitforActionAfterEPGUpdatePrerequisitesMetCount+1
			wscript.sleep(60000)
		end if
	loop
end if


'TimeToWaitForActionAfterEPGUpdatePrerequisites timed out
if ReadyForStandby=false then
	LogMsg(LanguageGetLine1Var(039, TimeToWaitForActionAfterEPGUpdatePrerequisites))
	LogMsg("***** " & LanguageGetLine0Var(011) & " *****")
	wscript.quit 1
end if


'End
LogMsg("***** " & LanguageGetLine0Var(011) & " *****")
wscript.quit 0




Function HTTPGet(sUrl, sRequest)
	set oHTTP=WScript.CreateObject("MSXML2.ServerXMLHTTP") 
	oHTTP.open "Get", sUrl & "?" & sRequest, false, ServiceUser, ServicePassword
	oHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	on error resume next
	err.clear
	oHTTP.send
	if err.number<>0 then
		LogMsg(languagegetline4var(005, sURL, err.source, err.number, err.description))
		LogMsg("***** " & LanguageGetLine0Var(011) & " *****")
		wscript.quit 1
	end if
	on error goto 0
	HTTPGet = oHTTP.responseText
End Function

Function HTTPPut(sUrl, sRequest)
	set oHTTP=WScript.CreateObject("MSXML2.ServerXMLHTTP") 
	oHTTP.open "Post", sUrl, false, ServiceUser, ServicePassword
	oHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oHTTP.setRequestHeader "Content-Length", len(srequest)
	on error resume next
	err.clear
	oHTTP.send srequest
	if err.number<>0 then
		LogMsg(languagegetline4var(005, sURL, err.source, err.number, err.description))
		LogMsg("***** " & LanguageGetLine0Var(011) & " *****")
		wscript.quit 1
	end if
	on error goto 0
	HTTPPut = oHTTP.responseText
End Function


Sub LogMsg(msg)
	LogMsgString=Right("0" & DatePart("h",time), 2) & ":" & Right("0" & DatePart("n",time), 2) & ":" & Right("0" & DatePart("s",time), 2) & " " & msg
	if len(LogMsgString)>79 then
		wscript.echo left(LogMsgstring,76) & "..."
	else
		wscript.echo LogMsgstring
	end if
	If Len(LogFile)>0 Then
		set fso = CreateObject("Scripting.FileSystemObject")
		set fsofile = fso.OpenTextFile(LogFile, 8, true)
		fsofile.writeline DatePart("yyyy",Date) & "-" & Right("0" & DatePart("m",Date), 2) & "-" & Right("0" & DatePart("d",Date), 2) & " " & LogMsgString
		fsofile.close
		set fsofile = nothing
		Set fso = nothing
	End If
End Sub


Sub CleanLogfile()
	RunsFoundSearchText="***** " & LanguageGetLine0Var(003) & " *****"
	If IsNumeric(runstokeepInLog) Then
		RunsToKeepInLog=int(RunsToKeepInLog)
		if runstokeepInLog>0 then
			LogMsg(LanguageGetLine1Var(059,RunsToKeepInLog))
		elseif runstokeepInLog=-1 then
			LogMsg(LanguageGetLine0Var(060,RunsToKeepInLog))
		else
			LogMsg(LanguageGetLine1Var(061,RunsToKeepInLog))
			LogMsg(LanguageGetLine0Var(062))
			RunsToKeepInLog=-1
		end if
	else
		LogMsg(LanguageGetLine1Var(061,RunsToKeepInLog))
		LogMsg(LanguageGetLine0Var(062))
		RunsToKeepInLog=-1
	end if

	If RunsToKeepInLog>0 then
		set fso = CreateObject("Scripting.FileSystemObject")
		If FSO.FileExists(logFile) Then
			Set tsInput = FSO.OpenTextFile(logfile)
			logarray = Split(tsInput.ReadAll(), vbNewLine)
			tsInput.Close
			RunsFound=0
			For Line = UBound(logarray) To 0 Step -1
				strline = logarray(Line)
				if instr(1, strline, RunsFoundSearchText, vbTextCompare) then
					RunsFound=RunsFound+1
				end if
				if RunsFound=RunsToKeepInLog then
					KeepFromLine=line
					exit for
				end if
			Next
			If KeepFromLine > 0 then
				'delete file
				fso.deletefile(logfile)
				'write new file
				set fso = CreateObject("Scripting.FileSystemObject")
				set fsofile = fso.OpenTextFile(LogFile, 8, true)

				for line=KeepFromLine to ubound(logarray)
					if line=ubound(logarray) then
						if logarray(line)<>"" then
							fsofile.writeline logarray(line)
						end if
					else
						fsofile.writeline logarray(line)
					end if
				next
				fsofile.close
				LogMsg(LanguageGetLine1Var(063,keepfromline))
			else
				LogMsg(LanguageGetLine1Var(064,runstokeepInLog))
			end if
		End If
	else
		If FSO.FileExists(logFile) Then
			LogMsg(LanguageGetLine0Var(065))
			fso.deletefile(logfile)
		else
			LogMsg(LanguageGetLine0Var(066))
		end if
	end if
end sub


Public Function LanguageGetLine0Var(LanguageLineNumber)
	LanguageGetLine0Var=replace(readini(languagefile, "default", right("000" & LanguageLineNumber, 3)), "$CRLF$", vbcrlf,1,-1,1)
End Function


Public Function LanguageGetLine1Var(LanguageLineNumber, LanguageReplaceVar1)
	LanguageGetLine1Var=replace(replace(readini(languagefile, "default", right("000" & LanguageLineNumber, 3)),"$1$",LanguageReplaceVar1), "$CRLF$", vbcrlf, 1,-1 ,1)
End Function


Public Function LanguageGetLine2Var(LanguageLineNumber, LanguageReplaceVar1, LanguageReplaceVar2)
	LanguageGetLine2Var=replace(replace(replace(readini(languagefile, "default", right("000" & LanguageLineNumber, 3)),"$1$",LanguageReplaceVar1), "$2$", LanguageReplaceVar2), "$CRLF$", vbcrlf,1,-1,1)
End Function


Public Function LanguageGetLine3Var(LanguageLineNumber, LanguageReplaceVar1, LanguageReplaceVar2, LanguageReplaceVar3)
	LanguageGetLine3Var=replace(replace(replace(replace(readini(languagefile, "default", right("000" & LanguageLineNumber, 3)),"$1$",LanguageReplaceVar1), "$2$", LanguageReplaceVar2), "$3$", LanguageReplaceVar3), "$CRLF$", vbcrlf,1,-1,1)
End Function


Public Function LanguageGetLine4Var(LanguageLineNumber, LanguageReplaceVar1, LanguageReplaceVar2, LanguageReplaceVar3, LanguageReplaceVar4)
	LanguageGetLine4Var=replace(replace(replace(replace(replace(readini(languagefile, "default", right("000" & LanguageLineNumber, 3)),"$1$",LanguageReplaceVar1), "$2$", LanguageReplaceVar2), "$3$", LanguageReplaceVar3), "$4$", LanguageReplaceVar4), "$CRLF$", vbcrlf,1,-1,1)
End Function


Public Function LanguageGetLine5Var(LanguageLineNumber, LanguageReplaceVar1, LanguageReplaceVar2, LanguageReplaceVar3, LanguageReplaceVar4, LanguageReplaceVar5)
	LanguageGetLine5Var=replace(replace(replace(replace(replace(replace(readini(languagefile, "default", right("000" & LanguageLineNumber, 3)),"$1$",LanguageReplaceVar1), "$2$", LanguageReplaceVar2), "$3$", LanguageReplaceVar3), "$4$", LanguageReplaceVar4), "$5$", LanguageReplaceVar5), "$CRLF$", vbcrlf,1,-1,1)
End Function


Public Function LanguageGetLine6Var(LanguageLineNumber, LanguageReplaceVar1, LanguageReplaceVar2, LanguageReplaceVar3, LanguageReplaceVar4, LanguageReplaceVar5, LanguageReplaceVar6)
	LanguageGetLine6Var=replace(replace(replace(replace(replace(replace(replace(readini(languagefile, "default", right("000" & LanguageLineNumber, 3)),"$1$",LanguageReplaceVar1), "$2$", LanguageReplaceVar2), "$3$", LanguageReplaceVar3), "$4$", LanguageReplaceVar4), "$5$", LanguageReplaceVar5), "$6$", LanguageReplaceVar6), "$CRLF$", vbcrlf,1,-1,1)
End Function


Public Function LanguageGetLine7Var(LanguageLineNumber, LanguageReplaceVar1, LanguageReplaceVar2, LanguageReplaceVar3, LanguageReplaceVar4, LanguageReplaceVar5, LanguageReplaceVar6, LanguageReplaceVar7)
	LanguageGetLine7Var=replace(replace(replace(replace(replace(replace(replace(replace(readini(languagefile, "default", right("000" & LanguageLineNumber, 3)),"$1$",LanguageReplaceVar1), "$2$", LanguageReplaceVar2), "$3$", LanguageReplaceVar3), "$4$", LanguageReplaceVar4), "$5$", LanguageReplaceVar5), "$6$", LanguageReplaceVar6), "$7$", LanguageReplaceVar7), "$CRLF$", vbcrlf,1,-1,1)
End Function


Function ReadIni( myFilePath, mySection, myKey )
    ' This function returns a value read from an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be returned
    '
    ' Returns:
    ' the [string] value for the specified key in the specified section
    '
    ' CAVEAT:     Will return a space if key exists but value is blank
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre, Rob van der Woude, Markus Gruber

    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Set objFSOini = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )
	strLine		= ""

    If objFSOini.FileExists( strFilePath ) Then
        Set objIniFile = objFSOini.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )

            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = LTrim( objIniFile.ReadLine )

                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Left( strLine, intEqualPos - 1 )
                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Mid( strLine, intEqualPos + 1 )
                             ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = LTrim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
		LogMsg(LanguageGetLine1Var(077, strFilePath))
		LogMsg("***** " & LanguageGetLine0Var(011) & " *****")
        Wscript.Quit 1
    End If
End Function

sub FillLCIDDictionary()
	LCIDDictionary.Add 1025, "Arabic - Saudi Arabia;ar-sa.ini;ar.ini"
	LCIDDictionary.Add 1026, "Bulgarian;bg.ini;bg.ini"
	LCIDDictionary.Add 1027, "Catalan;ca.ini;ca.ini"
	LCIDDictionary.Add 1028, "Chinese - Taiwan;zh-tw.ini;zh.ini"
	LCIDDictionary.Add 1029, "Czech;cs.ini;cs.ini"
	LCIDDictionary.Add 1030, "Danish;da.ini;da.ini"
	LCIDDictionary.Add 1031, "German - Germany;de-de.ini;de.ini"
	LCIDDictionary.Add 1032, "Greek;el.ini;el.ini"
	LCIDDictionary.Add 1033, "English - United States;en-us.ini;en.ini"
	LCIDDictionary.Add 1034, "Spanish - Spain;es-es.ini;es.ini"
	LCIDDictionary.Add 1035, "Finnish;fi.ini;fi.ini"
	LCIDDictionary.Add 1036, "French - France;fr-fr.ini;fr.ini"
	LCIDDictionary.Add 1037, "Hebrew;he.ini;he.ini"
	LCIDDictionary.Add 1038, "Hungarian;hu.ini;hu.ini"
	LCIDDictionary.Add 1039, "Icelandic;is.ini;is.ini"
	LCIDDictionary.Add 1040, "Italian - Italy;it-it.ini;it.ini"
	LCIDDictionary.Add 1041, "Japanese;ja.ini;ja.ini"
	LCIDDictionary.Add 1042, "Korean;ko.ini;ko.ini"
	LCIDDictionary.Add 1043, "Dutch - Netherlands;nl-nl.ini;nl.ini"
	LCIDDictionary.Add 1044, "Norwegian - Bokml;no-no.ini;no.ini"
	LCIDDictionary.Add 1045, "Polish;pl.ini;pl.ini"
	LCIDDictionary.Add 1046, "Portuguese - Brazil;pt-br.ini;pt.ini"
	LCIDDictionary.Add 1047, "Raeto-Romance;rm.ini;rm.ini"
	LCIDDictionary.Add 1048, "Romanian - Romania;ro.ini;ro.ini"
	LCIDDictionary.Add 1049, "Russian;ru.ini;ru.ini"
	LCIDDictionary.Add 1050, "Croatian;hr.ini;hr.ini"
	LCIDDictionary.Add 1051, "Slovak;sk.ini;sk.ini"
	LCIDDictionary.Add 1052, "Albanian;sq.ini;sq.ini"
	LCIDDictionary.Add 1053, "Swedish - Sweden;sv-se.ini;sv.ini"
	LCIDDictionary.Add 1054, "Thai;th.ini;th.ini"
	LCIDDictionary.Add 1055, "Turkish;tr.ini;tr.ini"
	LCIDDictionary.Add 1056, "Urdu;ur.ini;ur.ini"
	LCIDDictionary.Add 1057, "Indonesian;id.ini;id.ini"
	LCIDDictionary.Add 1058, "Ukrainian;uk.ini;uk.ini"
	LCIDDictionary.Add 1059, "Belarusian;be.ini;be.ini"
	LCIDDictionary.Add 1060, "Slovenian;sl.ini;sl.ini"
	LCIDDictionary.Add 1061, "Estonian;et.ini;et.ini"
	LCIDDictionary.Add 1062, "Latvian;lv.ini;lv.ini"
	LCIDDictionary.Add 1063, "Lithuanian;lt.ini;lt.ini"
	LCIDDictionary.Add 1065, "Farsi;fa.ini;fa.ini"
	LCIDDictionary.Add 1066, "Vietnamese;vi.ini;vi.ini"
	LCIDDictionary.Add 1067, "Armenian;hy.ini;hy.ini"
	LCIDDictionary.Add 1068, "Azeri - Latin;az-az.ini;az.ini"
	LCIDDictionary.Add 1069, "Basque;eu.ini;eu.ini"
	LCIDDictionary.Add 1070, "Sorbian;sb.ini;sb.ini"
	LCIDDictionary.Add 1071, "Macedonian (FYROM);mk.ini;mk.ini"
	LCIDDictionary.Add 1072, "Southern Sotho;st.ini;st.ini"
	LCIDDictionary.Add 1073, "Tsonga;ts.ini;ts.ini"
	LCIDDictionary.Add 1074, "Setsuana;tn.ini;tn.ini"
	LCIDDictionary.Add 1076, "Xhosa;xh.ini;xh.ini"
	LCIDDictionary.Add 1077, "Zulu;zu.ini;zu.ini"
	LCIDDictionary.Add 1078, "Afrikaans;af.ini;af.ini"
	LCIDDictionary.Add 1080, "Faroese;fo.ini;fo.ini"
	LCIDDictionary.Add 1081, "Hindi;hi.ini;hi.ini"
	LCIDDictionary.Add 1082, "Maltese;mt.ini;mt.ini"
	LCIDDictionary.Add 1084, "Gaelic - Scotland;gd.ini;gd.ini"
	LCIDDictionary.Add 1085, "Yiddish;yi.ini;yi.ini"
	LCIDDictionary.Add 1086, "Malay - Malaysia;ms-my.ini;ms.ini"
	LCIDDictionary.Add 1089, "Swahili;sw.ini;sw.ini"
	LCIDDictionary.Add 1091, "Uzbek  Latin;uz-uz.ini;uz.ini"
	LCIDDictionary.Add 1092, "Tatar;tt.ini;tt.ini"
	LCIDDictionary.Add 1097, "Tamil;ta.ini;ta.ini"
	LCIDDictionary.Add 1102, "Marathi;mr.ini;mr.ini"
	LCIDDictionary.Add 1103, "Sanskrit;sa.ini;sa.ini"
	LCIDDictionary.Add 2049, "Arabic - Iraq;ar-iq.ini;ar.ini"
	LCIDDictionary.Add 2052, "Chinese - China;zh-cn.ini;zh.ini"
	LCIDDictionary.Add 2055, "German - Switzerland;de-ch.ini;de.ini"
	LCIDDictionary.Add 2057, "English - United Kingdom;en-gb.ini;en.ini"
	LCIDDictionary.Add 2058, "Spanish - Mexico;es-mx.ini;es.ini"
	LCIDDictionary.Add 2060, "French - Belgium;fr-be.ini;fr.ini"
	LCIDDictionary.Add 2064, "Italian - Switzerland;it-ch.ini;it.ini"
	LCIDDictionary.Add 2067, "Dutch - Belgium;nl-be.ini;nl.ini"
	LCIDDictionary.Add 2068, "Norwegian - Nynorsk;no-no.ini;no.ini"
	LCIDDictionary.Add 2070, "Portuguese - Portugal;pt-pt.ini;pt.ini"
	LCIDDictionary.Add 2072, "Romanian - Moldova;ro-mo.ini;ro.ini"
	LCIDDictionary.Add 2073, "Russian - Moldova;ru-mo.ini;ru.ini"
	LCIDDictionary.Add 2074, "Serbian - Latin;sr-sp.ini;sr.ini"
	LCIDDictionary.Add 2077, "Swedish - Finland;sv-fi.ini;sv.ini"
	LCIDDictionary.Add 2092, "Azeri - Cyrillic;az-az.ini;az.ini"
	LCIDDictionary.Add 2108, "Gaelic - Ireland;gd-ie.ini;gd.ini"
	LCIDDictionary.Add 2110, "Malay  Brunei;ms-bn.ini;ms.ini"
	LCIDDictionary.Add 2115, "Uzbek - Cyrillic;uz-uz.ini;uz.ini"
	LCIDDictionary.Add 3073, "Arabic - Egypt;ar-eg.ini;ar.ini"
	LCIDDictionary.Add 3076, "Chinese - Hong Kong SAR;zh-hk.ini;zh.ini"
	LCIDDictionary.Add 3079, "German - Austria;de-at.ini;de.ini"
	LCIDDictionary.Add 3081, "English - Australia;en-au.ini;en.ini"
	LCIDDictionary.Add 3084, "French - Canada;fr-ca.ini;fr.ini"
	LCIDDictionary.Add 3098, "Serbian - Cyrillic;sr-sp.ini;sr.ini"
	LCIDDictionary.Add 4097, "Arabic - Libya;ar-ly.ini;ar.ini"
	LCIDDictionary.Add 4100, "Chinese - Singapore;zh-sg.ini;zh.ini"
	LCIDDictionary.Add 4103, "German - Luxembourg;de-lu.ini;de.ini"
	LCIDDictionary.Add 4105, "English - Canada;en-ca.ini;en.ini"
	LCIDDictionary.Add 4106, "Spanish - Guatemala;es-gt.ini;es.ini"
	LCIDDictionary.Add 4108, "French - Switzerland;fr-ch.ini;fr.ini"
	LCIDDictionary.Add 5121, "Arabic - Algeria;ar-dz.ini;ar.ini"
	LCIDDictionary.Add 5124, "Chinese - Macau SAR;zh-mo.ini;zh.ini"
	LCIDDictionary.Add 5127, "German - Liechtenstein;de-li.ini;de.ini"
	LCIDDictionary.Add 5129, "English - New Zealand;en-nz.ini;en.ini"
	LCIDDictionary.Add 5130, "Spanish - Costa Rica;es-cr.ini;es.ini"
	LCIDDictionary.Add 5132, "French - Luxembourg;fr-lu.ini;fr.ini"
	LCIDDictionary.Add 6145, "Arabic - Morocco;ar-ma.ini;ar.ini"
	LCIDDictionary.Add 6153, "English - Ireland;en-ie.ini;en.ini"
	LCIDDictionary.Add 6154, "Spanish - Panama;es-pa.ini;es.ini"
	LCIDDictionary.Add 7169, "Arabic - Tunisia;ar-tn.ini;ar.ini"
	LCIDDictionary.Add 7177, "English - South Africa;en-za.ini;en.ini"
	LCIDDictionary.Add 7178, "Spanish - Dominican Republic;es-do.ini;es.ini"
	LCIDDictionary.Add 8193, "Arabic - Oman;ar-om.ini;ar.ini"
	LCIDDictionary.Add 8201, "English - Jamaica;en-jm.ini;en.ini"
	LCIDDictionary.Add 8202, "Spanish - Venezuela;es-ve.ini;es.ini"
	LCIDDictionary.Add 9217, "Arabic - Yemen;ar-ye.ini;ar.ini"
	LCIDDictionary.Add 9225, "English - Caribbean;en-cb.ini;en.ini"
	LCIDDictionary.Add 9226, "Spanish - Colombia;es-co.ini;es.ini"
	LCIDDictionary.Add 10241, "Arabic - Syria;ar-sy.ini;ar.ini"
	LCIDDictionary.Add 10249, "English - Belize;en-bz.ini;en.ini"
	LCIDDictionary.Add 10250, "Spanish - Peru;es-pe.ini;es.ini"
	LCIDDictionary.Add 11265, "Arabic - Jordan;ar-jo.ini;ar.ini"
	LCIDDictionary.Add 11273, "English - Trinidad;en-tt.ini;en.ini"
	LCIDDictionary.Add 11274, "Spanish - Argentina;es-ar.ini;es.ini"
	LCIDDictionary.Add 12289, "Arabic - Lebanon;ar-lb.ini;ar.ini"
	LCIDDictionary.Add 12298, "Spanish - Ecuador;es-ec.ini;es.ini"
	LCIDDictionary.Add 13313, "Arabic - Kuwait;ar-kw.ini;ar.ini"
	LCIDDictionary.Add 13321, "English - Phillippines;en-ph.ini;en.ini"
	LCIDDictionary.Add 13322, "Spanish - Chile;es-cl.ini;es.ini"
	LCIDDictionary.Add 14337, "Arabic - United Arab Emirates;ar-ae.ini;ar.ini"
	LCIDDictionary.Add 14346, "Spanish - Uruguay;es-uy.ini;es.ini"
	LCIDDictionary.Add 15361, "Arabic - Bahrain;ar-bh.ini;ar.ini"
	LCIDDictionary.Add 15370, "Spanish - Paraguay;es-py.ini;es.ini"
	LCIDDictionary.Add 16385, "Arabic - Qatar;ar-qa.ini;ar.ini"
	LCIDDictionary.Add 16394, "Spanish - Bolivia;es-bo.ini;es.ini"
	LCIDDictionary.Add 17418, "Spanish - El Salvador;es-sv.ini;es.ini"
	LCIDDictionary.Add 18442, "Spanish - Honduras;es-hn.ini;es.ini"
	LCIDDictionary.Add 19466, "Spanish - Nicaragua;es-ni.ini;es.ini"
	LCIDDictionary.Add 20490, "Spanish - Puerto Rico;es-pr.ini;es.ini"
end sub

function GetValueFromIniFile(KeyName, TargetVariableName)
	if sectionname="default" then
		TargetVariableNameTemp=TargetVariableName
	end if
	do until sectionname=""
		GetValueFromIniFileTempString=readini(inifile, SectionName, KeyName)
		select case lcase(GetValueFromIniFileTempString)
			case ""
				'Key not found
				if TargetVariableNameTemp=true then
					GetValueFromIniFile=true
				elseif TargetVariableNameTemp=false then
					GetValueFromIniFile=false
				else
					GetValueFromIniFile=TargetVariableNameTemp
				end if
			case " "
				'Key found but empty
				if TargetVariableNameTemp=true then
					GetValueFromIniFile=true
				elseif TargetVariableNameTemp=false then
					GetValueFromIniFile=false
				else
					GetValueFromIniFile=TargetVariableNameTemp
				end if
			case "true"
				'correctly set boolean value
				GetValueFromIniFile=true
			case "false"
				'correctly set boolean value
				GetValueFromIniFile=false
			case else
				if left(GetValueFromIniFileTempString,1)="""" then GetValueFromIniFileTempString=right(GetValueFromIniFileTempString, len(GetValueFromIniFileTempString)-1)
				if right(GetValueFromIniFileTempString,1)="""" then GetValueFromIniFileTempString=left(GetValueFromIniFileTempString, len(GetValueFromIniFileTempString)-1)
				GetValueFromIniFile=GetValueFromIniFileTempString
		end select
		GetValueFromIniFileTempString=""
		if lcase(sectionname)="default" then
			sectionname=Tag
			TargetVariableNameTemp=GetValueFromIniFile
		else
			sectionname=""
		end if
	loop
	TargetVariableNameTemp=""
	sectionname="default"
end function
