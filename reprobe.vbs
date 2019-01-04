''REPROBE.VBS
''DESIGNED TO AUTOMATE DOWNLOAD AND INSTALL OF WINDOWS PROBE SOFTWARE
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
dim strIN, strOUT
''VARIABLES ACCEPTING PARAMETERS - CONFIGURES WINDOWS SOFTWARE PROBE MSI
dim strCID, strCNM, strPRB, strDMN, strUSR, strPWD, strRCMD
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objEXEC, objHOOK, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE, RE-PROBE.VBS, REF #2 , FIXES #7
strVER = 5
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
  if (wscript.arguments.count > 4) then                     ''
    strCID = objARG.item(0)
    strCNM = objARG.item(1)
    strPRB = objARG.item(2)
    strDMN = objARG.item(3)
    if (instr(1, strDMN, "\")) then
      strDMN = replace(strDMN, "\", vbnullstring)
    end if
    strUSR = objARG.item(4)
    strPWD = objARG.item(5)
  else                                                      ''NOT ENOUGH ARGUMENTS PASSED, END SCRIPT
    errRET = 1
  end if
else
  errRET = 1
end if

''------------
''BEGIN SCRIPT
if (errRET <> 0) then                                      ''NO ARGUMENTS PASSED, END SCRIPT
  objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES CUSTOMER ID, CUSTOMER NAME, DOMAIN, USER, AND PASSWORD"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES CUSTOMER ID, CUSTOMER NAME, DOMAIN, USER, AND PASSWORD"
  call CLEANUP()
elseif (errRET = 0) then
	objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING : RE-PROBE"
	objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING : RE-PROBE"
	''AUTOMATIC UPDATE, RE-PROBE.VBS, REF #2 , FIXES #7
	call CHKAU()
  ''GRANT USER SERVICE LOGON
  objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING SERVICE LOGON SCRIPT : SVCPERM"
  objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING SERVICE LOGON SCRIPT : SVCPERM"
  call FILEDL("https://github.com/CW-Khristos/CW_MSI/raw/master/SVCperm.vbs", "SVCperm.vbs")
  ''EXECUTE SERVICE LOGON SCRIPT : SVCPERM
  objOUT.write vbnewline & now & vbtab & vbtab & " - EXECUTING SERVICE LOGON SCRIPT : SVCPERM"
  objLOG.write vbnewline & now & vbtab & vbtab & " - EXECUTING SERVICE LOGON SCRIPT : SVCPERM"
  if (strDMN <> vbnullstring) then
    call HOOK("cscript.exe //nologo " & chr(34) & "c:\temp\svcperm.vbs" & chr(34) & " " & chr(34) & strDMN & "\" & strUSR & chr(34))
  elseif (strDMN = vbnullstring) then
    call HOOK("cscript.exe //nologo " & chr(34) & "c:\temp\svcperm.vbs" & chr(34) & " " & chr(34) & strUSR & chr(34))
  end if
	''DOWNLOAD WINDOWS PROBE MSI
	objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS PROBE MSI"
	objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS PROBE MSI"
  call FILEDL("https://github.com/CW-Khristos/CW_MSI/raw/master/Windows%20Software%20Probe.msi", "windows software probe.msi")
  ''INSTALL WINDOWS PROBE
  objOUT.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS PROBE"
  objLOG.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS PROBE"
  ''WINDOWS PROBE RE-CONFIGURATION COMMAND, VALIDATED 08/13/2018, PROBE REQUIRES ADMIN USER PRIOR TO RUNNING, FIXES #6
  select case strPRB
    case "Local_Windows"
      strRCMD = "msiexec /i " & chr(34) & "c:\temp\windows software probe.msi" & chr(34) & " /qn CUSTOMERID=" & strCID & " CUSTOMERNAME=" & chr(34) & strCNM & chr(34) & _
        " SERVERPROTOCOL=" & chr(34) & "HTTPS://" & chr(34) & " SERVERPORT=443 SERVERADDRESS=" & chr(34) & "ilmcw.dyndns.biz" & chr(34) & " PROBETYPE=" & chr(34) & strPRB & chr(34) & _
        " AGENTUSERNAME=" & chr(34) & strUSR & chr(34) & " AGENTPASSWORD=" & chr(34) & strPWD & chr(34) & " /l*v c:\temp\probe_install.log ALLUSERS=2"
    case "Workgroup_Windows"
      strRCMD = "msiexec /i " & chr(34) & "c:\temp\windows software probe.msi" & chr(34) & " /qn CUSTOMERID=" & strCID & " CUSTOMERNAME=" & chr(34) & strCNM & chr(34) & _
        " SERVERPROTOCOL=" & chr(34) & "HTTPS://" & chr(34) & " SERVERPORT=443 SERVERADDRESS=" & chr(34) & "ilmcw.dyndns.biz" & chr(34) & " PROBETYPE=" & chr(34) & strPRB & chr(34) & _
        " AGENTUSERNAME=" & chr(34) & strUSR & chr(34) & " AGENTPASSWORD=" & chr(34) & strPWD & chr(34) & " /l*v c:\temp\probe_install.log ALLUSERS=2"
    case "Network_Windows"
      strRCMD = "msiexec /i " & chr(34) & "c:\temp\windows software probe.msi" & chr(34) & " /qn CUSTOMERID=" & strCID & " CUSTOMERNAME=" & chr(34) & strCNM & chr(34) & _
        " SERVERPROTOCOL=" & chr(34) & "HTTPS://" & chr(34) & " SERVERPORT=443 SERVERADDRESS=" & chr(34) & "ilmcw.dyndns.biz" & chr(34) & " PROBETYPE=" & chr(34) & strPRB & chr(34) & _
        " AGENTDOMAIN=" & chr(34) & strDMN & chr(34) & " AGENTUSERNAME=" & chr(34) & strUSR & chr(34) & " AGENTPASSWORD=" & chr(34) & strPWD & chr(34) & " /l*v c:\temp\probe_install.log ALLUSERS=2"
  end select
  ''RE-CONFIGURE WINDOWS PROBE
	objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING : " & strRCMD
	objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING : " & strRCMD
  call HOOK(strRCMD)
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub CHKAU()																									''CHECK FOR SCRIPT UPDATE, RE-PROBE.VBS, REF #2 , FIXES #7
  ''REMOVE WINDOWS AGENT CACHED VERSION OF SCRIPT
  if (objFSO.fileexists("C:\Program Files (x86)\N-Able Technologies\Windows Agent\cache\" & wscript.scriptname)) then
    objFSO.deletefile "C:\Program Files (x86)\N-Able Technologies\Windows Agent\cache\" & wscript.scriptname, true
  end if
	''ADD WINHTTP SECURE CHANNEL TLS REGISTRY KEYS
	call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp" & chr(34) & _
		" /f /v DefaultSecureProtocols /t REG_DWORD /d 0x00000A00 /reg:32")
	call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp" & chr(34) & _
		" /f /v DefaultSecureProtocols /t REG_DWORD /d 0x00000A00 /reg:64")
	''SCRIPT OBJECT FOR PARSING XML
	set objXML = createobject("Microsoft.XMLDOM")
	''FORCE SYNCHRONOUS
	objXML.async = false
	''LOAD SCRIPT VERSIONS DATABASE XML
	if objXML.load("https://github.com/CW-Khristos/scripts/raw/master/version.xml") then
		set colVER = objXML.documentelement
		for each objSCR in colVER.ChildNodes
			''LOCATE CURRENTLY RUNNING SCRIPT
			if (lcase(objSCR.nodename) = lcase(wscript.scriptname)) then
				''CHECK LATEST VERSION
				if (cint(objSCR.text) > cint(strVER)) then
					objOUT.write vbnewline & now & vbtab & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					objLOG.write vbnewline & now & vbtab & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					''DOWNLOAD LATEST VERSION OF SCRIPT
					call FILEDL("https://github.com/CW-Khristos/CW_MSI/raw/master/reprobe.vbs", wscript.scriptname)
					''RUN LATEST VERSION
					if (wscript.arguments.count > 0) then             ''ARGUMENTS WERE PASSED
						for x = 0 to (wscript.arguments.count - 1)
							strTMP = strTMP & " " & chr(34) & objARG.item(x) & chr(34)
						next
            objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
            objLOG.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
						objWSH.run "cscript.exe //nologo " & chr(34) & "c:\temp\" & wscript.scriptname & chr(34) & strTMP, 0, false
					elseif (wscript.arguments.count = 0) then         ''NO ARGUMENTS WERE PASSED
            objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
            objLOG.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
						objWSH.run "cscript.exe //nologo " & chr(34) & "c:\temp\" & wscript.scriptname & chr(34), 0, false
					end if
					''END SCRIPT
					call CLEANUP()
				end if
			end if
		next
	end if
	set colVER = nothing
	set objXML = nothing
end sub

sub FILEDL(strURL, strFILE)                                 ''CALL HOOK TO DOWNLOAD FILE FROM URL
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
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
    errRET = 2
		err.clear
  end if
end sub

sub HOOK(strCMD)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
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
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
		errRET = 3
		err.clear
  end if
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  if (errRET = 0) then         															''RE-PROBE COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "RE-PROBE SUCCESSFUL : " & NOW
  elseif (errRET <> 0) then    															''RE-PROBE FAILED
    objOUT.write vbnewline & "RE-PROBE FAILURE : " & NOW & " : " & errRET
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
  wscript.quit errRET
end sub