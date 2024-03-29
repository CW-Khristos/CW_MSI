''PME_REMOVAL.VBS
''DESIGNED TO AUTOMATICALLY UNINSTALL NABLE PME SERVICES SILENTLY
''WRITTEN BY : CJ BLEDSOE / CBLEDSOE<@>IPMCOMPUTERS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objEXEC, objHOOK, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE, AGENT_REMOVAL.VBS, REF #2 , REF #68 , REF #69 , FIXES #21 , FIXES #31
strVER = 1
strREPO = "CW_MSI"
strBRCH = "master"
strDIR = vbnullstring
''DEFAULT SUCCESS
errRET = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''ENVIRONMENT VARIABLES
strPD = objWSH.expandenvironmentstrings("%PROGRAMDATA%")
strPF = objWSH.expandenvironmentstrings("%PROGRAMFILES%")
strPF86 = objWSH.expandenvironmentstrings("%PROGRAMFILES(X86)%")
''CHECK 'PERSISTENT' FOLDERS , REF #2 , REF #73
if (not (objFSO.folderexists("c:\temp"))) then
  objFSO.createfolder("c:\temp")
end if
if (not (objFSO.folderexists("C:\IT\"))) then
  objFSO.createfolder("C:\IT\")
end if
if (not (objFSO.folderexists("C:\IT\Scripts\"))) then
  objFSO.createfolder("C:\IT\Scripts\")
end if
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\PME_REMOVAL")) then            ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\PME_REMOVAL", true
  set objLOG = objFSO.createtextfile("C:\temp\PME_REMOVAL")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\PME_REMOVAL", 8)
else                                                          ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\PME_REMOVAL")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\PME_REMOVAL", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                         ''ARGUMENTS WERE PASSED
  ''ARGUMENT OUTPUT DISABLED TO SANITIZE
  'for x = 0 to (wscript.arguments.count - 1)
  '  objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  '  objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  'next 
  if (wscript.arguments.count >= 1) then                      ''SET VARIABLES ACCEPTING ARGUMENTS
  end if
elseif (wscript.arguments.count < 1) then                     ''NOT ENOUGH ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  'call LOGERR(1)
  'call CLEANUP()
end if

''------------
''BEGIN SCRIPT
if (errRET = 0) then
    objOUT.write vbnewline & vbnewline & now & vbtab & " - STARTING PME_REMOVAL" & vbnewline
    objLOG.write vbnewline & vbnewline & now & vbtab & " - STARTING PME_REMOVAL" & vbnewline
	''AUTOMATIC UPDATE, PME_REMOVAL.VBS, REF #2 , REF #69 , REF #68 , FIXES #9
  ''DOWNLOAD CHKAU.VBS SCRIPT, REF #2 , REF #69 , REF #68
  call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/master/chkAU.vbs", "C:\IT\Scripts", "chkAU.vbs")
  ''EXECUTE CHKAU.VBS SCRIPT, REF #69
  objOUT.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : PME_REMOVAL : " & strVER
  objLOG.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : PME_REMOVAL : " & strVER
  intRET = objWSH.run ("cmd.exe /C " & chr(34) & "cscript.exe " & chr(34) & "C:\IT\Scripts\chkAU.vbs" & chr(34) & " " & _
    chr(34) & strREPO & chr(34) & " " & chr(34) & strBRCH & chr(34) & " " & chr(34) & strDIR & chr(34) & " " & _
    chr(34) & wscript.scriptname & chr(34) & " " & chr(34) & strVER & chr(34) & chr(34), 0, true)
  ''CHKAU RETURNED - NO UPDATE FOUND , REF #2 , REF #69 , REF #68
  objOUT.write vbnewline & "errRET='" & intRET & "'"
  objLOG.write vbnewline & "errRET='" & intRET & "'"
  intRET = (intRET - vbObjectError)
  objOUT.write vbnewline & "errRET='" & intRET & "'"
  objLOG.write vbnewline & "errRET='" & intRET & "'"
  if ((intRET = 4) or (intRET = 10) or (intRET = 11) or (intRET = 1) or (intRET = 2147221505) or (intRET = 2147221517)) then
    objOUT.write vbnewline & now & vbtab & vbtab & " - NO UPDATE FOUND : PME_REMOVAL : " & strVER
    objLOG.write vbnewline & now & vbtab & vbtab & " - NO UPDATE FOUND : PME_REMOVAL : " & strVER
    ''DETERMINE OS ARCHITECTURE
    if (GetOSbits = 64) then
      strPF = strPF86
    elseif (GetOSbits = 32) then
      strPF = strPF
    end if
    ''STOP SERVICES
    objOUT.write vbnewline & now & vbtab & vbtab & " - STOPPING PME SERVICES"
    objLOG.write vbnewline & now & vbtab & vbtab & " - STOPPING PME SERVICES"
    call HOOK("sc stop " & chr(34) & "EcosystemAgent" & chr(34))
    call HOOK("sc stop " & chr(34) & "EcosystemAgentMaintenance" & chr(34))
    call HOOK("sc stop " & chr(34) & "PME.Agent.PmeService" & chr(34))
    call HOOK("sc stop " & chr(34) & "SolarWinds.MSP.CacheService" & chr(34))
    call HOOK("sc stop " & chr(34) & "SolarWinds.MSP.RpcServerService" & chr(34))
    ''KILL SERVICE PROCESSES
    objOUT.write vbnewline & now & vbtab & vbtab & " - STOPPING PME PROCESSES"
    objLOG.write vbnewline & now & vbtab & vbtab & " - STOPPING PME PROCESSES"
    call HOOK("taskkill /F /IM PME.Agent.exe /T")
    call HOOK("taskkill /F /IM RequestHandlerAgent.exe /T")
    call HOOK("taskkill /F /IM FileCacheServiceAgent.exe /T")
    call HOOK("taskkill /F /IM SolarWinds.MSP.Ecosystem.WindowsAgent.exe /T")
    call HOOK("taskkill /F /IM SolarWinds.MSP.Ecosystem.WindowsAgentMaint.exe /T")
    '''''''''''''''''''''''
    '' > PME Version 2.0 ''
    '''''''''''''''''''''''
    ''PME AGENT
    if (objFSO.fileexists(strPF & "\MspPlatform\PME\unins000.exe")) then
      objOUT.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING PME AGENT"
      objLOG.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING PME AGENT"
      'objWSH.run chr(34) & strPF & "\MspPlatform\PME\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true
      intRET = objWSH.run (chr(34) & strPF & "\MspPlatform\PME\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true)
      'for intLOOP = 0 to 20
      '  wscript.sleep 6000
      'next
      'call HOOK("taskkill /F /IM unins000.exe")
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING PME DRIECTORIES"
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING PME DRIECTORIES"
      if (objFSO.folderexists(strPD & "\MspPlatform\PME\archives")) then
        objFSO.deletefolder chr(34) & strPD & "\MspPlatform\PME\archives" & chr(34), true
      end if
      if (objFSO.folderexists(strPD & "\MspPlatform\PME")) then
        objFSO.deletefolder chr(34) & strPD & "\MspPlatform\PME" & chr(34), true
      end if
    end if
    ''ECOSYSTEM AGENT
    if (objFSO.fileexists(strPF & "\MspPlatform\Ecosystem Agent\unins000.exe")) then
      objOUT.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING ECOSYSTEM AGENT"
      objLOG.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING ECOSYSTEM AGENT"
      'objWSH.run chr(34) & strPF & "\MspPlatform\Ecosystem Agent\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true
      intRET = objWSH.run (chr(34) & strPF & "\MspPlatform\Ecosystem Agent\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true)
      'for intLOOP = 0 to 20
      '  wscript.sleep 6000
      'next
      'call HOOK("taskkill /F /IM unins000.exe")
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING ECOSYSTEM AGENT DRIECTORIES"
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING ECOSYSTEM AGENT DRIECTORIES"
      if (objFSO.folderexists(strPF & "\MspPlatform\Ecosystem Agent")) then
        objFSO.deletefolder chr(34) & strPF & "\MspPlatform\Ecosystem Agent" & chr(34), true
      end if
      if (objFSO.folderexists(strPD & "\MspPlatform\Ecosystem Agent")) then
        objFSO.deletefolder chr(34) & strPD & "\MspPlatform\Ecosystem Agent" & chr(34), true
      end if
      if (objFSO.folderexists(strPD & "\MspPlatform\EcosystemAgent")) then
        objFSO.deletefolder chr(34) & strPD & "\MspPlatform\EcosystemAgent" & chr(34), true
      end if
      if (objFSO.folderexists(strPD & "\MspPlatform\EcosystemAgentMaintenance")) then
        objFSO.deletefolder chr(34) & strPD & "\MspPlatform\EcosystemAgentMaintenance" & chr(34), true
      end if
    end if
    ''REQUEST HANDLER AGENT
    if (objFSO.fileexists(strPF & "\MspPlatform\RequestHandlerAgent\unins000.exe")) then
      objOUT.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING REQUEST HANDLER AGENT"
      objLOG.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING REQUEST HANDLER AGENT"
      'objWSH.run chr(34) & strPF & "\MspPlatform\RequestHandlerAgent\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true
      intRET = objWSH.run (chr(34) & strPF & "\MspPlatform\RequestHandlerAgent\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true)
      'for intLOOP = 0 to 20
      '  wscript.sleep 6000
      'next
      'call HOOK("taskkill /F /IM unins000.exe")
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING REQUEST HANDLER AGENT DRIECTORIES"
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING REQUEST HANDLER AGENT DRIECTORIES"
      if (objFSO.folderexists(strPF & "\MspPlatform\RequestHandlerAgent")) then
        objFSO.deletefolder chr(34) & strPF & "\MspPlatform\RequestHandlerAgent" & chr(34), true
      end if
      if (objFSO.folderexists(strPD & "\MspPlatform\RequestHandlerAgent")) then
        objFSO.deletefolder chr(34) & strPD & "\MspPlatform\RequestHandlerAgent" & chr(34), true
      end if
    end if
    ''FILE CACHE SERVICE AGENT
    if (objFSO.fileexists(strPF & "\MspPlatform\FileCacheServiceAgent\unins000.exe")) then
      objOUT.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING FILE CACHE SERVICE AGENT"
      objLOG.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING FILE CACHE SERVICE AGENT"
      'objWSH.run chr(34) & strPF & "\MspPlatform\FileCacheServiceAgent\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true
      intRET = objWSH.run (chr(34) & strPF & "\MspPlatform\FileCacheServiceAgent\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true)
      'for intLOOP = 0 to 20
      '  wscript.sleep 6000
      'next
      'call HOOK("taskkill /F /IM unins000.exe")
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING FILE CACHE SERVICE AGENT DRIECTORIES"
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING FILE CACHE SERVICE AGENT DRIECTORIES"
      if (objFSO.folderexists(strPF & "\MspPlatform\FileCacheServiceAgent")) then
        objFSO.deletefolder chr(34) & strPF & "\MspPlatform\FileCacheServiceAgent" & chr(34), true
      end if
      if (objFSO.folderexists(strPD & "\MspPlatform\FileCacheServiceAgent")) then
        objFSO.deletefolder chr(34) & strPD & "\MspPlatform\FileCacheServiceAgent" & chr(34), true
      end if
      if (objFSO.folderexists(strPD & "\MspPlatform\SolarWinds.MSP.CacheService" & chr(34))) then
        objFSO.deletefolder chr(34) & strPD & "\MspPlatform\SolarWinds.MSP.CacheService" & chr(34), true
      end if
    end if
    ''CLEAR PROGRAM FILES / PROGRAM FILES (X86) FOLDER
    if (objFSO.fileexists(strPF & "\MspPlatform")) then
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING PROGRAM FILES\MSPPLATFORM DRIECTORY"
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING PROGRAM FILES\MSPPLATFORM DRIECTORY"
      objFSO.deletefolder chr(34) & strPF & "\MspPlatform" & chr(34), true
    end if
    ''CLEAR PROGRAMDATA\MSPPLATFORM FOLDER
    if (objFSO.folderexists(strPD & "\MspPlatform")) then
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING PROGRAMDATA\MSPPLATFORM DRIECTORY"
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING PROGRAMDATA\MSPPLATFORM DRIECTORY"
      objFSO.deletefolder chr(34) & strPD & "\MspPlatform" & chr(34), true
    end if
    '''''''''''''''''''''''
    '' < PME Version 2.0 ''
    '''''''''''''''''''''''
    ''PME AGENT
    if (objFSO.fileexists(strPF & "\SolarWinds MSP\PME\unins000.exe")) then
      objOUT.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING PME AGENT"
      objLOG.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING PME AGENT"
      'objWSH.run chr(34) & strPF & "\SolarWinds MSP\PME\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true
      intRET = objWSH.run (chr(34) & strPF & "\SolarWinds MSP\PME\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true)
      'for intLOOP = 0 to 20
      '  wscript.sleep 6000
      'next
      'call HOOK("taskkill /F /IM unins000.exe")
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING PME DRIECTORIES"
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING PME DRIECTORIES"
      if (objFSO.folderexists(strPD & "\SolarWinds MSP\PME\archives")) then
        objFSO.deletefolder chr(34) & strPD & "\SolarWinds MSP\PME\archives" & chr(34), true
      end if
      if (objFSO.folderexists(strPD & "\SolarWinds MSP\PME")) then
        objFSO.deletefolder chr(34) & strPD & "\SolarWinds MSP\PME" & chr(34), true
      end if
    end if
    ''ECOSYSTEM AGENT
    if (objFSO.fileexists(strPF & "\SolarWinds MSP\Ecosystem Agent\unins000.exe")) then
      objOUT.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING ECOSYSTEM AGENT"
      objLOG.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING ECOSYSTEM AGENT"
      'objWSH.run chr(34) & strPF & "\SolarWinds MSP\Ecosystem Agent\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true
      intRET = objWSH.run (chr(34) & strPF & "\SolarWinds MSP\Ecosystem Agent\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true)
      'for intLOOP = 0 to 20
      '  wscript.sleep 6000
      'next
      'call HOOK("taskkill /F /IM unins000.exe")
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING ECOSYSTEM AGENT DRIECTORIES"
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING ECOSYSTEM AGENT DRIECTORIES"
      if (objFSO.folderexists(strPF & "\SolarWinds MSP\Ecosystem Agent")) then
        objFSO.deletefolder chr(34) & strPF & "\SolarWinds MSP\Ecosystem Agent" & chr(34), true
      end if
      if (objFSO.folderexists(strPD & "\SolarWinds MSP\Ecosystem Agent")) then
        objFSO.deletefolder chr(34) & strPD & "\SolarWinds MSP\Ecosystem Agent" & chr(34), true
      end if
      if (objFSO.folderexists(strPD & "\SolarWinds MSP\EcosystemAgent")) then
        objFSO.deletefolder chr(34) & strPD & "\SolarWinds MSP\EcosystemAgent" & chr(34), true
      end if
      if (objFSO.folderexists(strPD & "\SolarWinds MSP\EcosystemAgentMaintenance")) then
        objFSO.deletefolder chr(34) & strPD & "\SolarWinds MSP\EcosystemAgentMaintenance" & chr(34), true
      end if
    end if
    ''REQUEST HANDLER AGENT
    if (objFSO.fileexists(strPF & "\SolarWinds MSP\RequestHandlerAgent\unins000.exe")) then
      objOUT.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING REQUEST HANDLER AGENT"
      objLOG.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING REQUEST HANDLER AGENT"
      'objWSH.run chr(34) & strPF & "\SolarWinds MSP\RequestHandlerAgent\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true
      intRET = objWSH.run (chr(34) & strPF & "\SolarWinds MSP\RequestHandlerAgent\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true)
      'for intLOOP = 0 to 20
      '  wscript.sleep 6000
      'next
      'call HOOK("taskkill /F /IM unins000.exe")
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING REQUEST HANDLER AGENT DRIECTORIES"
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING REQUEST HANDLER AGENT DRIECTORIES"
      if (objFSO.folderexists(strPF & "\SolarWinds MSP\RequestHandlerAgent")) then
        objFSO.deletefolder chr(34) & strPF & "\SolarWinds MSP\RequestHandlerAgent" & chr(34), true
      end if
      if (objFSO.folderexists(strPD & "\SolarWinds MSP\RequestHandlerAgent")) then
        objFSO.deletefolder chr(34) & strPD & "\SolarWinds MSP\RequestHandlerAgent" & chr(34), true
      end if
    end if
    ''FILE CACHE SERVICE AGENT
    if (objFSO.fileexists(strPF & "\SolarWinds MSP\CacheService\unins000.exe")) then
      objOUT.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING FILE CACHE SERVICE AGENT"
      objLOG.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING FILE CACHE SERVICE AGENT"
      'objWSH.run chr(34) & strPF & "\SolarWinds MSP\CacheService\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true
      intRET = objWSH.run (chr(34) & strPF & "\SolarWinds MSP\CacheService\unins000.exe" & chr(34) & " /s /qn /silent /verysilent /norestart", 0, true)
      'for intLOOP = 0 to 20
      '  wscript.sleep 6000
      'next
      'call HOOK("taskkill /F /IM unins000.exe")
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING FILE CACHE SERVICE AGENT DRIECTORIES"
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING FILE CACHE SERVICE AGENT DRIECTORIES"
      if (objFSO.folderexists(strPF & "\SolarWinds MSP\FileCacheServiceAgent")) then
        objFSO.deletefolder chr(34) & strPF & "\SolarWinds MSP\FileCacheServiceAgent" & chr(34), true
      end if
      if (objFSO.folderexists(strPD & "\SolarWinds MSP\FileCacheServiceAgent")) then
        objFSO.deletefolder chr(34) & strPD & "\SolarWinds MSP\FileCacheServiceAgent" & chr(34), true
      end if
      if (objFSO.folderexists(strPD & "\SolarWinds MSP\SolarWinds.MSP.CacheService" & chr(34))) then
        objFSO.deletefolder chr(34) & strPD & "\SolarWinds MSP\SolarWinds.MSP.CacheService" & chr(34), true
      end if
    end if
    ''CLEAR PROGRAM FILES / PROGRAM FILES (X86) FOLDER
    if (objFSO.fileexists(strPF & "\SolarWinds MSP")) then
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING PROGRAM FILES\SOLARWINDS MSP DRIECTORY"
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING PROGRAM FILES\SOLARWINDS MSP DRIECTORY"
      objFSO.deletefolder chr(34) & strPF & "\SolarWinds MSP" & chr(34), true
    end if
    ''CLEAR PROGRAMDATA\SOLARWINDS MSP FOLDER
    if (objFSO.folderexists(strPD & "\SolarWinds MSP")) then
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING PROGRAMDATA\SOLARWINDS MSP DRIECTORY"
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING PROGRAMDATA\SOLARWINDS MSP DRIECTORY"
      objFSO.deletefolder chr(34) & strPD & "\SolarWinds MSP" & chr(34), true
    end if
    ''REMOVE SERVICES
    objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING PME SERVICES"
    objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING PME SERVICES"
    call HOOK("sc delete " & chr(34) & "EcosystemAgent" & chr(34))
    call HOOK("sc delete " & chr(34) & "EcosystemAgentMaintenance" & chr(34))
    call HOOK("sc delete " & chr(34) & "PME.Agent.PmeService" & chr(34))
    call HOOK("sc delete " & chr(34) & "SolarWinds.MSP.CacheService" & chr(34))
    call HOOK("sc delete " & chr(34) & "SolarWinds.MSP.RpcServerService" & chr(34))
  end if
elseif (errRET <> 0) then
  call LOGERR(errRET)
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''FUNCTIONS
function GetOSbits()
   if (objWSH.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%") = "AMD64") then
      GetOSbits = 64
   else
      GetOSbits = 32
   end if
end function

''SUB-ROUTINES
sub FILEDL(strURL, strDL, strFILE)                            ''CALL HOOK TO DOWNLOAD FILE FROM URL , 'ERRRET'=11
  strSAV = vbnullstring
  ''SET DOWNLOAD PATH
  strSAV = strDL & "\" & strFILE
  objOUT.write vbnewline & now & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  objLOG.write vbnewline & now & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  ''ADD WINHTTP SECURE CHANNEL TLS REGISTRY KEYS
  call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp" & chr(34) & _
    " /f /v DefaultSecureProtocols /t REG_DWORD /d 0x00000A00 /reg:32")
  call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp" & chr(34) & _
    " /f /v DefaultSecureProtocols /t REG_DWORD /d 0x00000A00 /reg:64")
  ''CHECK IF FILE ALREADY EXISTS
  if objFSO.fileexists(strSAV) then
    ''DELETE FILE FOR OVERWRITE
    objFSO.deletefile(strSAV)
  end if
  ''CREATE HTTP OBJECT
  set objHTTP = createobject("WinHttp.WinHttpRequest.5.1")
  ''DOWNLOAD FROM URL
  objHTTP.open "GET", strURL, false
  objHTTP.send
  if (objHTTP.status = 200) then
    dim objStream
    set objStream = createobject("ADODB.Stream")
    with objStream
      .Type = 1 'adTypeBinary
      .Open
      .Write objHTTP.ResponseBody
      .SaveToFile strSAV
      .Close
    end with
    set objStream = nothing
  end if
  ''CHECK THAT FILE EXISTS
  if objFSO.fileexists(strSAV) then
    objOUT.write vbnewline & vbnewline & now & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
  end if
	set objHTTP = nothing
  if ((err.number <> 0) and (err.number <> 58)) then          ''ERROR RETURNED DURING UPDATE CHECK , 'ERRRET'=11
    call LOGERR(11)
  end if
end sub

sub HOOK(strCMD)                                              ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND , 'ERRRET'=12
  on error resume next
  objOUT.write vbnewline & now & vbtab & vbtab & " - EXECUTING : HOOK(" & strCMD & ")"
  objLOG.write vbnewline & now & vbtab & vbtab & " - EXECUTING : HOOK(" & strCMD & ")"
  set objHOOK = objWSH.exec(strCMD)
  while (not objHOOK.stdout.atendofstream)
    if (instr(1, strCMD, "takeown /F ") = 0) then             ''SUPPRESS 'TAKEOWN' SUCCESS MESSAGES
      strIN = objHOOK.stdout.readline
      if (strIN <> vbnullstring) then
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      end if
    end if
  wend
  wscript.sleep 10
  if (instr(1, strCMD, "takeown /F ") = 0) then               ''SUPPRESS 'TAKEOWN' SUCCESS MESSAGES
    strIN = objHOOK.stdout.readall
    if (strIN <> vbnullstring) then
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
    end if
  end if
  set objHOOK = nothing
  if (err.number <> 0) then                                   ''ERROR RETURNED , 'ERRRET'=12
    call LOGERR(12)
  end if
end sub

sub LOGERR(intSTG)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  errRET = intSTG
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
		err.clear
  end if
  select case intSTG
    case 1                                                    ''NOT ENOUGH ARGUMENTS , 'ERRRET'=1
      objOUT.write vbnewline & vbnewline & now & vbtab & " - NOT ENOUGH ARGUMENTS PASSED"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - NOT ENOUGH ARGUMENTS PASSED"
  end select
end sub

sub CLEANUP()                                                 ''SCRIPT CLEANUP
  on error resume next
  if (errRET = 0) then         															  ''PME_REMOVAL COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - PME_REMOVAL SUCCESSFUL : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - PME_REMOVAL SUCCESSFUL : " & errRET & " : " & now
    err.clear
  elseif (errRET <> 0) then    															  ''PME_REMOVAL FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - PME_REMOVAL FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - PME_REMOVAL FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "PME_REMOVAL", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - PME_REMOVAL COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - PME_REMOVAL COMPLETE" & vbnewline
  objLOG.close
  ''EMPTY OBJECTS
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN ERROR NUMBER
  wscript.quit err.number
end sub