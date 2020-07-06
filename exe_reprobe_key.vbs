''EXE_REPROBE_KEY.VBS
''DESIGNED TO AUTOMATE DOWNLOAD AND INSTALL OF WINDOWS PROBE SOFTWARE
''UTILIZES THE SYSTEM SPECIFIC WINDOWS PROBE EXE INSTALLER WITH CONFIGURED PARAMETERS
''ACCEPTS 5 PARAMETERS , REQUIRES 4 PARAMETERS
''REQUIRED PARAMETER : 'STRKEY' , STRING TO SET ACTIVATION KEY
''REQUIRED PARAMETER : 'STRPRB' , STRING TO SET PROBE TYPE
''REQUIRED PARAMETER : 'STRUSR' , STRING TO SET USER
''REQUIRED PARAMETER : 'STRPWD' , STRING TO SET PASSWORD
''OPTIONAL PARAMETER : 'STRSVR' , STRING TO SET SERVER ADDRESS
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
dim strREPO, strBRCH, strDIR
''VARIABLES ACCEPTING PARAMETERS - CONFIGURES WINDOWS SOFTWARE PROBE EXE
dim strKEY, strSVR
dim strIN, strOUT, strRCMD
dim strPRB, strDMN, strUSR, strPWD
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objEXEC, objHOOK, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE , EXE_REPROBE_KEY.VBS , REF #2 , REF #69 , FIXES #7 , FIXES #13 , FIXES #18
strVER = 5
strREPO = "CW_MSI"
strBRCH = "master"
strDIR = vbnullstring
''DEFAULT SUCCESS
errRET = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''CREATE SCRIPTING OBJECTS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
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
if (objFSO.fileexists("C:\temp\EXE_REPROBE_KEY")) then      ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\EXE_REPROBE_KEY", true
  set objLOG = objFSO.createtextfile("C:\temp\EXE_REPROBE_KEY")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\EXE_REPROBE_KEY", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\EXE_REPROBE_KEY")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\EXE_REPROBE_KEY", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  ''ARGUMENT OUTPUT DISABLED TO SANITIZE
  'for x = 0 to (wscript.arguments.count - 1)
  '  objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  '  objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  'next 
  if (wscript.arguments.count > 3) then                     ''SET REQUIRED VARIABLES ACCEPTING ARGUMENTS
    strKEY = objARG.item(0)                                 ''SET REQUIRED PARAMTERS 'STRKEY' , ACTIVATION KEY
    strPRB = objARG.item(1)                                 ''SET REQUIRED PARAMETER 'STRPRB' , PROBE TYPE - WORKGROUP_WINDOWS / NETWORK_WINDOWS
    if ((lcase(strPRB) = "workgroup") or (lcase(strPRB) = "local")) then
      strPRB = "Workgroup_Windows"
    elseif ((lcase(strPRB) = "network") or (lcase(strPRB) = "domain")) then
      strPRB= "Network_Windows"
    end if
    strUSR = objARG.item(2)                                 ''SET REQUIRED PARAMETER 'STRUSR' , TARGET USER
    if (instr(1, strUSR, "\")) then                         ''INPUT VALIDATION FOR 'STRUSR'
      strDMN = split(strUSR, "\")(0)                        '' 'LOCAL' - PASS 'USERNAME' ONLY; AND 'DOMAIN' - 'DOMAIN\USER' DOMAIN LOGON
      strUSR = split(strUSR, "\")(1)                        ''STRIP WORKGROUP / DOMAIN FROM PASSED VARIABLE TO ENSURE WE HAVE USER NAME ONLY
    elseif (instr(1, strUSR, "\") = 0) then                 ''INPUT VALIDATION FOR 'STRUSR'
      strDMN = "."                                          '' CONVERT 'LOCAL' FOR SCRIPT RE-EXECUTION IN CHKAU.VBS
    end if
    strPWD = objARG.item(3)                                 ''SET REQUIRED PARAMETER 'STRPWD' , USER PASSWORD
    if (wscript.arguments.count = 4) then                   ''NO OPTIONAL ARGUMENT PASSED
      strSVR = "ncentral.cwitsupport.com"                   ''SET OPTIONAL PARAMETER 'STRSVR' , 'DEFAULT' SERVER ADDRESS
    elseif (wscript.arguments.count = 5) then               ''OPTIONAL ARGUMENT PASSED
      if (strSVR = vbnullstring) then                       ''OPTIONAL 'STRSVR' ARGUMENT EMPTY
        strSVR = "ncentral.cwitsupport.com"                 ''SET OPTIONAL PARAMETER 'STRSVR' , 'DEFAULT' SERVER ADDRESS
      elseif (strSVR <> vbnullstring) then                  ''OPTIONAL 'STRSVR' ARGUMENT NOT EMPTY
        strSVR = objARG.item(4)                             ''SET OPTIONAL PARAMETER 'STRSVR' , PASSED SERVER ADDRESS; SEPARATE MULTIPLES WITH ','
      end if
    end if
  else                                                      ''NOT ENOUGH ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
    call LOGERR(1)
  end if
else                                                        ''NOT ENOUGH ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  call LOGERR(1)
end if

''------------
''BEGIN SCRIPT
if (errRET = 0) then                                        ''ARGUMENTS PASSED , CONTINUE SCRIPT
	objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING : EXE_REPROBE_KEY"
	objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING : EXE_REPROBE_KEY"
	''AUTOMATIC UPDATE, EXE_REPROBE_KEY.VBS, REF #2 , REF #68 , REF #69 , FIXES #7
  ''DOWNLOAD CHKAU.VBS SCRIPT, REF #2 , REF #68 , REF #69
  call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/master/chkAU.vbs", "C:\IT\Scripts", "chkAU.vbs")
  ''EXECUTE CHKAU.VBS SCRIPT, REF #69
  objOUT.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : EXE_REPROBE_KEY : " & strVER
  objLOG.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : EXE_REPROBE_KEY : " & strVER
  intRET = objWSH.run ("cmd.exe /C " & chr(34) & "cscript.exe " & chr(34) & "C:\IT\Scripts\chkAU.vbs" & chr(34) & " " & _
    chr(34) & strREPO & chr(34) & " " & chr(34) & strBRCH & chr(34) & " " & chr(34) & strDIR & chr(34) & " " & _
    chr(34) & wscript.scriptname & chr(34) & " " & chr(34) & strVER & chr(34) & " " & _
    chr(34) & strCID & "|" & strCNM & "|" & strPRB & "|" & strDMN & "\" & strUSR & "|" & strPWD & "|" & strSVR & chr(34) & chr(34), 0, true)
  ''CHKAU RETURNED - NO UPDATE FOUND , REF #2 , REF #69 , REF #68
  objOUT.write vbnewline & "errRET='" & intRET & "'"
  objLOG.write vbnewline & "errRET='" & intRET & "'"
  intRET = (intRET - vbObjectError)
  objOUT.write vbnewline & "errRET='" & intRET & "'"
  objLOG.write vbnewline & "errRET='" & intRET & "'"
  if ((intRET = 4) or (intRET = 10) or (intRET = 11) or (intRET = 1) or (intRET = 2147221505) or (intRET = 2147221517)) then
    objOUT.write vbnewline & now & vbtab & vbtab & " - NO UPDATE FOUND : EXE_REPROBE_KEY : " & strVER
    objLOG.write vbnewline & now & vbtab & vbtab & " - NO UPDATE FOUND : EXE_REPROBE_KEY : " & strVER
    ''DISABLED VERIFY NETWORK SETTINGS; WILL BE PASSING NETWORK TYPE AS PARAMETER , REF #71
    ''VERIFY NETWORK WORKGROUP / DOMAIN SETTINGS , REF #7 , FIXES #12
    'set objEXEC = objWSH.exec("net config workstation")
    'while (not objEXEC.stdout.atendofstream)
    '  strIN = objEXEC.stdout.readline
    '  'objOUT.write vbnewline & now & vbtab & vbtab & strIN
    '  'objLOG.write vbnewline & now & vbtab & vbtab & strIN
    '  if ((trim(strIN) <> vbnullstring) and (instr(1, lcase(strIN), "logon domain"))) then
    '    objOUT.write vbnewline & now & vbtab & vbtab & strIN
    '    objLOG.write vbnewline & now & vbtab & vbtab & strIN
    '    strDMN = (split(strIN, " ")(ubound(split(strIN, " "))))
    '    ''HANDLE "\" IN PASSED 'STRUSR'
    '    if (instr(1, lcase(strUSR), "\")) then
    '      strUSR = strDMN & "\" & split(strUSR, "\")(1)
    '    ''HANDLE NO "\" IN PASSED 'STRUSR'
    '    elseif (instr(1, lcase(strUSR), "\") = 0) then
    '      strUSR = strDMN & "\" & strUSR
    '    end if
    '    if (strPRB = "Workgroup_Windows") then
    '      strUSR = split(strUSR, "\")(1)
    '    end if
    '  end if
    '  if (err.number <> 0) then
    '    call LOGERR(2)
    '  end if
    'wend
    'set objEXEC = nothing
    ''DOWNLOAD SVCPERM.VBS SCRIPT TO GRANT USER SERVICE LOGON , 'ERRRET'=2
    objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING SERVICE LOGON SCRIPT : SVCPERM"
    objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING SERVICE LOGON SCRIPT : SVCPERM"
    call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/master/SVCperm.vbs", "C:\IT\Scripts", "SVCperm.vbs")
    if (errRET <> 0) then
      call LOGERR(2)
    end if
    ''EXECUTE SERVICE LOGON SCRIPT : SVCPERM , 'ERRRET'=3
    objOUT.write vbnewline & now & vbtab & vbtab & " - EXECUTING SERVICE LOGON SCRIPT : SVCPERM"
    objLOG.write vbnewline & now & vbtab & vbtab & " - EXECUTING SERVICE LOGON SCRIPT : SVCPERM"
    if ((strDMN <> vbnullstring) and (strDMN <> ".")) then   ''EXECUTE SVCPERM.VBS AT DOMAIN LEVEL
      call HOOK("cscript.exe //nologo " & chr(34) & "c:\IT\Scripts\svcperm.vbs" & chr(34) & " " & chr(34) & strDMN & "\" & strUSR & chr(34) & " " & chr(34) & "domain" & chr(34))
    elseif ((strDMN = vbnullstring) or (strDMN = ".")) then  ''EXECUTE SVCPERM.VBS AT LOCAL LEVEL
      call HOOK("cscript.exe //nologo " & chr(34) & "c:\IT\Scripts\svcperm.vbs" & chr(34) & " " & chr(34) & strUSR & chr(34) & " " & chr(34) & "local" & chr(34))
    end if
    if (errRET <> 0) then
      call LOGERR(3)
    end if
    ''DOWNLOAD WINDOWS PROBE MSI , 'ERRRET'=4
    objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS PROBE SYSTEM-SPECIFIC EXE"
    objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS PROBE SYSTEM-SPECIFIC EXE"
    call FILEDL("https://raw.githubusercontent.com/CW-Khristos/CW_MSI/master/WindowsProbeSetup.exe", "C:\IT", "WindowsSoftwareProbe.exe")
    if (errRET <> 0) then
      call LOGERR(4)
    end if
    ''INSTALL WINDOWS PROBE
    objOUT.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS PROBE"
    objLOG.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS PROBE"
    ''WINDOWS PROBE RE-CONFIGURATION COMMAND, VALIDATED 08/13/2018, PROBE REQUIRES ADMIN USER PRIOR TO RUNNING, FIXES #6 , FIXES #18
    select case lcase(strPRB)
      ''LOCAL ONLY
      case "local_windows"
        strRCMD = chr(34) & "c:\IT\WindowsSoftwareProbe.exe" & chr(34) & " /s /v" & chr(34) & " /qn /norestart /l*v c:\temp\probe_install.log" & _
          " AGENTACTIVATIONKEY=" & strKEY & " SERVERPROTOCOL=HTTPS SERVERPORT=443 SERVERADDRESS=" & strSVR & " PROBETYPE=" & strPRB & _
          " AGENTUSERNAME=\" & chr(34) & strUSR & "\" & chr(34) & " AGENTPASSWORD=\" & chr(34) & strPWD & "\" & chr(34) & " " & chr(34)
      ''WORKGROUP ENVIRONMENT
      case "workgroup_windows"
        if (instr(1, strUSR, "\")) then
          strSUSR = split(strUSR, "\")(1)
        elseif (instr(1, strUSR, "\") = 0) then
          strSUSR = strUSR
        end if
        ''WORKGROUP_WINDOWS - " AGENTUSERNAME=" & chr(34) & split(strUSR, "\")(1) - STRIP RETRIEVED "LOGON DOMAIN" INFORMATION FROM 'STRUSR' PRIOR TO EXECUTING MSIEXEC , FIXES #12
        strRCMD = chr(34) & "c:\IT\WindowsSoftwareProbe.exe" & chr(34) & " /s /v" & chr(34) & " /qn /norestart /l*v c:\temp\probe_install.log" & _
          " AGENTACTIVATIONKEY=" & strKEY & " SERVERPROTOCOL=HTTPS SERVERPORT=443 SERVERADDRESS=" & strSVR & " PROBETYPE=" & strPRB & _
          " AGENTUSERNAME=\" & chr(34) & strSUSR & "\" & chr(34) & " AGENTPASSWORD=\" & chr(34) & strPWD & "\" & chr(34) & " " & chr(34)
      ''DOMAIN ENVIRONMENT
      case "network_windows"
        strRCMD = chr(34) & "c:\IT\WindowsSoftwareProbe.exe" & chr(34) & " /s /v" & chr(34) & " /qn /norestart /l*v c:\temp\probe_install.log" & _
          " AGENTACTIVATIONKEY=" & strKEY & " SERVERPROTOCOL=HTTPS SERVERPORT=443 SERVERADDRESS=" & strSVR & " PROBETYPE=" & strPRB & _
          " AGENTDOMAIN=" & strDMN & " AGENTUSERNAME=\" & chr(34) & strUSR & "\" & chr(34) & " AGENTPASSWORD=\" & chr(34) & strPWD & "\" & chr(34) & " " & chr(34)
    end select
    ''RE-CONFIGURE WINDOWS PROBE , 'ERRRET'=5 , 'STRCMD' OUTPUT DISABLED TO SANITIZE
    call HOOK(strRCMD)
    if (errRET <> 0) then
      call LOGERR(5)
    end if
  end if
elseif (errRET <> 0) then                                   ''NO ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  call LOGERR(errRET)
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub FILEDL(strURL, strDL, strFILE)                          ''CALL HOOK TO DOWNLOAD FILE FROM URL , 'ERRRET'=11
  strSAV = vbnullstring
  ''SET DOWNLOAD PATH
  strSAV = strDL & "\" & strFILE
  objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
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
    objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
    objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
  end if
	set objHTTP = nothing
  if ((err.number <> 0) and (err.number <> 58)) then        ''ERROR RETURNED DURING DOWNLOAD , 'ERRRET'=11
    call LOGERR(11)
  end if
end sub

sub HOOK(strCMD)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND , 'ERRRET'=12
  on error resume next
  objOUT.write vbnewline & now & vbtab & vbtab & " - EXECUTING : HOOK" '& strCMD
  objLOG.write vbnewline & now & vbtab & vbtab & " - EXECUTING : HOOK" '& strCMD
  set objHOOK = objWSH.exec(strCMD)
  if (instr(1, strCMD, "takeown /F ") = 0) then             ''SUPPRESS 'TAKEOWN' SUCCESS MESSAGES
    while (not objHOOK.stdout.atendofstream)
      strIN = objHOOK.stdout.readline
      if (strIN <> vbnullstring) then
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      end if
    wend
    wscript.sleep 10
    strIN = objHOOK.stdout.readall
    if (strIN <> vbnullstring) then
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
    end if
  end if
  set objHOOK = nothing
  if (err.number <> 0) then                                 ''ERROR RETURNED DURING UPDATE CHECK , 'ERRRET'=12
    call LOGERR(12)
  end if
end sub

sub LOGERR(intSTG)                                          ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  errRET = intSTG
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
		err.clear
  end if
  ''CUSTOM ERROR CODES
  select case intSTG
    case 1                                                  '' 'ERRRET'=1 - NOT ENOUGH ARGUMENTS
      objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES CUSTOMER ID, CUSTOMER NAME, DOMAIN, USER, AND PASSWORD"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES CUSTOMER ID, CUSTOMER NAME, DOMAIN, USER, AND PASSWORD"
  end select
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  on error resume next
  if (errRET = 0) then         															''EXE_REPROBE_KEY COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "EXE_REPROBE_KEY SUCCESSFUL : " & errRET & " : " & now
    objLOG.write vbnewline & "EXE_REPROBE_KEY SUCCESSFUL : " & errRET & " : " & now
    err.clear
  elseif (errRET <> 0) then    															''EXE_REPROBE_KEY FAILED
    objOUT.write vbnewline & "EXE_REPROBE_KEY FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & "EXE_REPROBE_KEY FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "EXE_REPROBE_KEY", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - EXE_REPROBE_KEY COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - EXE_REPROBE_KEY COMPLETE" & vbnewline
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