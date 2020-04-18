''REPROBE.VBS
''DESIGNED TO AUTOMATE DOWNLOAD AND INSTALL OF WINDOWS PROBE SOFTWARE
''UTILIZES THE CUSTOMER SPECIFIC WINDOWS PROBE EXE INSTALLER
''ACCEPTS 7 PARAMETERS , REQUIRES 6 PARAMETERS
''REQUIRED PARAMETER : 'STRCID' , STRING TO SET CUSTOMER ID
''REQUIRED PARAMETER : 'STRCNM' , STRING TO SET CUSTOMER NAME
''REQUIRED PARAMETER : 'STRPRB' , STRING TO SET PROBE TYPE
''REQUIRED PARAMETER : 'STRDMN' , STRING TO SET DOMAIN
''REQUIRED PARAMETER : 'STRUSR' , STRING TO SET USER; 'LOCAL' - PASS 'USERNAME' ONLY; AND 'DOMAIN' - 'DOMAIN\USER' DOMAIN LOGON -
''REQUIRED PARAMETER : 'STRPWD' , STRING TO SET PASSWORD
''OPTIONAL PARAMETER : 'STRSVR' , STRING TO SET SERVER ADDRESS
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
''VARIABLES ACCEPTING PARAMETERS - CONFIGURES WINDOWS SOFTWARE PROBE MSI
dim strIN, strOUT, strRCMD
dim strCID, strCNM, strSVR
dim strPRB, strDMN, strUSR, strPWD
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objEXEC, objHOOK, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE , RE-PROBE.VBS , REF #2 , REF #69 , FIXES #7
strVER = 15
strREPO = "CW_MSI"
strBRCH = "dev"
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
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\reprobe")) then              ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\reprobe", true
  set objLOG = objFSO.createtextfile("C:\temp\reprobe")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\reprobe", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\reprobe")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\reprobe", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  if (wscript.arguments.count > 4) then                     ''SET REQUIRED VARIABLES ACCEPTING ARGUMENTS
    strCID = objARG.item(0)                                 ''SET REQUIRED PARAMTERS 'STRCID' , CUSTOMER ID
    strCNM = objARG.item(1)                                 ''SET REQUIRED PARAMETER 'STRCNM' , CUSTOMER NAME
    strPRB = objARG.item(2)                                 ''SET REQUIRED PARAMETER 'STRPRB' , PROBE TYPE - WORKGROUP_WINDOWS / NETWORK_WINDOWS
    if ((lcase(strPRB) = "workgroup") or (lcase(strPRB) = "local")) then
      strPRB = "Workgroup_Windows"
    elseif ((lcase(strPRB) = "network") or (lcase(strPRB) = "domain")) then
      strPRB= "Network_Windows"
    end if
    strUSR = objARG.item(3)                                 ''SET REQUIRED PARAMETER 'STRUSR' , TARGET USER
    if (instr(1, strUSR, "\")) then                         ''INPUT VALIDATION FOR 'STRUSR'
      strDMN = split(strUSR, "\")(0)                        '' 'LOCAL' - PASS 'USERNAME' ONLY; AND 'DOMAIN' - 'DOMAIN\USER' DOMAIN LOGON
      strUSR = split(strUSR, "\")(1)                        ''STRIP WORKGROUP / DOMAIN FROM PASSED VARIABLE TO ENSURE WE HAVE USER NAME ONLY
    elseif (instr(1, strUSR, "\") = 0) then                 ''INPUT VALIDATION FOR 'STRUSR'
      strDMN = "."                                          '' CONVERT 'LOCAL' FOR SCRIPT RE-EXECUTION IN CHKAU.VBS
    end if
    strPWD = objARG.item(4)                                 ''SET REQUIRED PARAMETER 'STRPWD' , USER PASSWORD
    if (wscript.arguments.count = 5) then                   ''NO OPTIONAL ARGUMENT PASSED
      strSVR = "ncentral.cwitsupport.com"                   ''SET OPTIONAL PARAMETER 'STRSVR' , 'DEFAULT' SERVER ADDRESS
    elseif (wscript.arguments.count > 5) then               ''OPTIONAL ARGUMENTS PASSED
      strSVR = objARG.item(5)                               ''SET OPTIONAL PARAMETER 'STRSVR' , PASSED SERVER ADDRESS; SEPARATE MULTIPLES WITH ','
      if (strSVR = vbnullstring) then                       ''OPTIONAL 'STRSVR' ARGUMENT EMPTY
        strSVR = "ncentral.cwitsupport.com"                 ''SET OPTIONAL PARAMETER 'STRSVR' , 'DEFAULT' SERVER ADDRESS
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
	objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING : RE-PROBE"
	objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING : RE-PROBE"
	''AUTOMATIC UPDATE, RE-PROBE.VBS, REF #2 , REF #69 , REF #68 , FIXES #7
  ''DOWNLOAD CHKAU.VBS SCRIPT, REF #2 , REF #69 , REF #68
  call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/chkAU.vbs", "chkAU.vbs")
  ''EXECUTE CHKAU.VBS SCRIPT, REF #69
  objOUT.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : RE-PROBE : " & strVER
  objLOG.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : RE-PROBE : " & strVER
  intRET = objWSH.run ("cmd.exe /C " & chr(34) & "cscript.exe " & chr(34) & "C:\temp\chkAU.vbs" & chr(34) & " " & _
    chr(34) & strREPO & chr(34) & " " & chr(34) & strBRCH & chr(34) & " " & chr(34) & strDIR & chr(34) & " " & _
    chr(34) & wscript.scriptname & chr(34) & " " & chr(34) & strVER & chr(34) & " " & _
    chr(34) & strCID & "|" & strCNM & "|" & strPRB & "|" & strDMN & "\" & strUSR & "|" & strPWD & "|" & strSVR & chr(34), 0, true)
  ''CHKAU RETURNED - NO UPDATE FOUND , REF #2 , REF #69 , REF #68
  intRET = (intRET - vbObjectError)
  if ((intRET = 4) or (intRET = 10) or (intRET = 11) or (intRET = 1)) then
    objOUT.write vbnewline & now & vbtab & vbtab & " - NO UPDATE FOUND : RE-PROBE : " & strVER
    objLOG.write vbnewline & now & vbtab & vbtab & " - NO UPDATE FOUND : RE-PROBE : " & strVER
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
    '  end if
    '  if (err.number <> 0) then
    '    call LOGERR(2)
    '  end if
    'wend
    'set objEXEC = nothing
    ''DOWNLOAD SVCPERM.VBS SCRIPT TO GRANT USER SERVICE LOGON , 'ERRRET'=2
    objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING SERVICE LOGON SCRIPT : SVCPERM"
    objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING SERVICE LOGON SCRIPT : SVCPERM"
    call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/SVCperm.vbs", "SVCperm.vbs")
    if (errRET <> 0) then
      call LOGERR(2)
    end if
    ''EXECUTE SERVICE LOGON SCRIPT : SVCPERM , 'ERRRET'=3
    objOUT.write vbnewline & now & vbtab & vbtab & " - EXECUTING SERVICE LOGON SCRIPT : SVCPERM"
    objLOG.write vbnewline & now & vbtab & vbtab & " - EXECUTING SERVICE LOGON SCRIPT : SVCPERM"
    if ((strDMN <> vbnullstring) and (strDMN <> ".")) then   ''EXECUTE SVCPERM.VBS AT DOMAIN LEVEL
      call HOOK("cscript.exe //nologo " & chr(34) & "c:\temp\svcperm.vbs" & chr(34) & " " & chr(34) & strDMN & "\" & strUSR & chr(34) & " " & chr(34) & "domain" & chr(34))
    elseif ((strDMN = vbnullstring) or (strDMN = ".")) then  ''EXECUTE SVCPERM.VBS AT LOCAL LEVEL
      call HOOK("cscript.exe //nologo " & chr(34) & "c:\temp\svcperm.vbs" & chr(34) & " " & chr(34) & strUSR & chr(34) & " " & chr(34) & "local" & chr(34))
    end if
    if (errRET <> 0) then
      call LOGERR(3)
    end if
    ''DOWNLOAD WINDOWS PROBE MSI , 'ERRRET'=4
    objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS PROBE SYSTEM-SPECIFIC MSI"
    objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS PROBE SYSTEM-SPECIFIC MSI"
    call FILEDL("https://github.com/CW-Khristos/CW_MSI/raw/master/Windows%20Software%20Probe.msi", "windows software probe.msi")
    if (errRET <> 0) then
      call LOGERR(4)
    end if
    ''INSTALL WINDOWS PROBE
    objOUT.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS PROBE"
    objLOG.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS PROBE"
    ''WINDOWS PROBE RE-CONFIGURATION COMMAND, VALIDATED 08/13/2018, PROBE REQUIRES ADMIN USER PRIOR TO RUNNING, FIXES #6
    select case lcase(strPRB)
      ''LOCAL ONLY
      case "local_windows"
        strRCMD = "msiexec /i " & chr(34) & "c:\temp\windows software probe.msi" & chr(34) & " /qn CUSTOMERID=" & strCID & " CUSTOMERNAME=" & chr(34) & strCNM & chr(34) & _
          " SERVERPROTOCOL=" & chr(34) & "HTTPS" & chr(34) & " SERVERPORT=443 SERVERADDRESS=" & chr(34) & strSVR & chr(34) & " PROBETYPE=" & chr(34) & strPRB & chr(34) & _
          " AGENTUSERNAME=" & chr(34) & strUSR & chr(34) & " AGENTPASSWORD=" & chr(34) & strPWD & chr(34) & " /l*v c:\temp\probe_install.log ALLUSERS=2"
      ''WORKGROUP ENVIRONMENT
      case "workgroup_windows"
        ''WORKGROUP_WINDOWS - " AGENTUSERNAME=" & chr(34) & split(strUSR, "\")(1) - STRIP RETRIEVED "LOGON DOMAIN" INFORMATION FROM 'STRUSR' PRIOR TO EXECUTING MSIEXEC , FIXES #12
        if (instr(1, strUSR, "\")) then
          strSUSR = split(strUSR, "\")(1)
        elseif (instr(1, strUSR, "\") = 0) then
          strSUSR = strUSR
        end if
        strRCMD = "msiexec /i " & chr(34) & "c:\temp\windows software probe.msi" & chr(34) & " /qn CUSTOMERID=" & strCID & " CUSTOMERNAME=" & chr(34) & strCNM & chr(34) & _
          " SERVERPROTOCOL=" & chr(34) & "HTTPS" & chr(34) & " SERVERPORT=443 SERVERADDRESS=" & chr(34) & strSVR & chr(34) & " PROBETYPE=" & chr(34) & strPRB & chr(34) & _
          " AGENTUSERNAME=" & chr(34) & strSUSR & chr(34) & " AGENTPASSWORD=" & chr(34) & strPWD & chr(34) & " /l*v c:\temp\probe_install.log ALLUSERS=2"
      ''DOMAIN ENVIRONMENT
      case "network_windows"
        strRCMD = "msiexec /i " & chr(34) & "c:\temp\windows software probe.msi" & chr(34) & " /qn CUSTOMERID=" & strCID & " CUSTOMERNAME=" & chr(34) & strCNM & chr(34) & _
          " SERVERPROTOCOL=" & chr(34) & "HTTPS" & chr(34) & " SERVERPORT=443 SERVERADDRESS=" & chr(34) & strSVR & chr(34) & " PROBETYPE=" & chr(34) & strPRB & chr(34) & _
          " AGENTDOMAIN=" & chr(34) & strDMN & chr(34) & " AGENTUSERNAME=" & chr(34) & strUSR & chr(34) & " AGENTPASSWORD=" & chr(34) & strPWD & chr(34) & " /l*v c:\temp\probe_install.log ALLUSERS=2"
    end select
    ''RE-CONFIGURE WINDOWS PROBE , 'ERRRET'=5
    objOUT.write vbnewline & now & vbtab & vbtab & " - EXECUTING : " & strRCMD
    objLOG.write vbnewline & now & vbtab & vbtab & " - EXECUTING : " & strRCMD
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
sub FILEDL(strURL, strFILE)                                 ''CALL HOOK TO DOWNLOAD FILE FROM URL , 'ERRRET'=11
  strSAV = vbnullstring
  ''SET DOWNLOAD PATH
  strSAV = "C:\temp\" & strFILE
  objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  if objFSO.fileexists(strSAV) then
    objFSO.deletefile(strSAV)
  end if
  ''CREATE HTTP OBJECT
  set objHTTP = createobject( "WinHttp.WinHttpRequest.5.1" )
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
  if (err.number <> 0) then                                 ''ERROR RETURNED , 'ERRRET'=11
    call LOGERR(11)
  end if
end sub

sub HOOK(strCMD)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND , 'ERRRET'=12
  on error resume next
  set objHOOK = objWSH.exec(strCMD)
	while (not objHOOK.stdout.atendofstream)
		strIN = objHOOK.stdout.readline
		if (strIN <> vbnullstring) then
			objOUT.write vbnewline & now & vbtab & vbtab & strIN 
			objLOG.write vbnewline & now & vbtab & vbtab & strIN 
		end if
	wend
	wscript.sleep 10
  strIN = objHOOK.stdout.readall
  if (strIN <> vbnullstring) then
    objOUT.write vbnewline & now & vbtab & vbtab & strIN 
    objLOG.write vbnewline & now & vbtab & vbtab & strIN 
  end if
  set objHOOK = nothing
  if (err.number <> 0) then                                 ''ERROR RETURNED , 'ERRRET'=12
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
  select case intSTG
    case 1                                                  '' 'ERRRET'=1 - NOT ENOUGH ARGUMENTS
      objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES CUSTOMER ID, CUSTOMER NAME, DOMAIN, USER, AND PASSWORD"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES CUSTOMER ID, CUSTOMER NAME, DOMAIN, USER, AND PASSWORD"
  end select
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  on error resume next
  if (errRET = 0) then         															''RE-PROBE COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "RE-PROBE SUCCESSFUL : " & errRET & " : " & now
    objLOG.write vbnewline & "RE-PROBE SUCCESSFUL : " & errRET & " : " & now
    err.clear
  elseif (errRET <> 0) then    															''RE-PROBE FAILED
    objOUT.write vbnewline & "RE-PROBE FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & "RE-PROBE FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "RE-PROBE", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - RE-PROBE COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - RE-PROBE COMPLETE" & vbnewline
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