''REAGENT.VBS
''DESIGNED TO AUTOMATE DOWNLOAD AND INSTALL OF WINDOWS AGENT SOFTWARE
''ACCEPTS 3 PARAMETERS, REQUIRES 2 PARAMETERS
''REQUIRED PARAMETER : 'STRCID', STRING TO SET CUSTOMER ID
''REQUIRED PARAMETER : 'STRCNM', STRING TO SET CUSTOMER NAME
''OPTIONAL PARAMETER : 'STRSVR', STRING TO SET SERVER ADDRESS
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
''VARIABLES ACCEPTING PARAMETERS - CONFIGURES WINDOWS AGENT MSI
dim strIN, strOUT, strRCMD
dim strCID, strCNM, strSVR
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objEXEC, objHOOK, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE, RE-AGENT.VBS, REF #2 , FIXES #8
strVER = 6
''DEFAULT SUCCESS
errRET = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\reagent")) then              ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\reagent", true
  set objLOG = objFSO.createtextfile("C:\temp\reagent")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\reagent", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\reagent")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\reagent", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  if (wscript.arguments.count > 1) then                     ''SET VARIABLES ACCEPTING ARGUMENTS
    strCID = objARG.item(0)                                 ''CUSTOMER ID
    strCNM = objARG.item(1)                                 ''CUSTOMER NAME
    if (wscript.arguments.count = 2) then
      strSVR = "ncentral.cwitsupport.com"                   ''SERVER ADDRESS
    elseif (wscript.arguments.count = 3) then
      if (strSVR = vbnullstring) then
        strSVR = "ncentral.cwitsupport.com"                 ''SERVER ADDRESS
      elseif (strSVR <> vbnullstring) then
        strSVR = objARG.item(6)                             ''SERVER ADDRESS
      end if
    end if
  else                                                      ''NOT ENOUGH ARGUMENTS PASSED, END SCRIPT
    errRET = 1
  end if
end if

''------------
''BEGIN SCRIPT
if (errRET <> 0) then                                      ''NO ARGUMENTS PASSED, END SCRIPT
  objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES CUSTOMER ID, CUSTOMER NAME"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES CUSTOMER ID, CUSTOMER NAME"
  call CLEANUP()
elseif (errRET = 0) then                                   ''ARGUMENTS PASSED, CONTINUE SCRIPT
	objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING RE-AGENT"
	objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING RE-AGENT"
	''AUTOMATIC UPDATE, RE-AGENT.VBS, REF #2 , FIXES #8
	call CHKAU()
	''DOWNLOAD WINDOWS AGENT MSI
	objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS AGENT MSI"
  objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS AGENT MSI"
  call FILEDL("https://github.com/CW-Khristos/CW_MSI/raw/master/Windows%20Agent.msi", "windows agent.msi")
  ''INSTALL WINDOWS AGENT
  objOUT.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS AGENT"
  objLOG.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS AGENT"
  ''WINDOWS AGENT RE-CONFIGURATION COMMAND
	strRCMD = "msiexec /i " & chr(34) & "c:\temp\windows agent.msi" & chr(34) & " /qn CUSTOMERID=" & strCID & _
		" CUSTOMERNAME=" & chr(34) & strCNM & chr(34) & " SERVERPROTOCOL=https:// SERVERPORT=443 SERVERADDRESS=" & chr(34) & strSVR & chr(34) & _
    " /l*v c:\temp\agent_install.log ALLUSERS=2"
	''RE-CONFIGURE WINDOWS AGENT
	call HOOK(strRCMD)
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub CHKAU()																									''CHECK FOR SCRIPT UPDATE, RE-AGENT.VBS, REF #2 , FIXES #8
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
					call FILEDL("https://github.com/CW-Khristos/CW_MSI/raw/master/reagent.vbs", wscript.scriptname)
					''RUN LATEST VERSION
					if (wscript.arguments.count > 0) then             ''ARGUMENTS WERE PASSED
						for x = 0 to (wscript.arguments.count - 1)
							strTMP = strTMP & " " & chr(34) & objARG.item(x) & chr(34)
						next
            objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING " & objSCR.nodename & " : " & objSCR.text & vbnewline
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
  objOUT.write vbnewline & now & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  objLOG.write vbnewline & now & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  ''CREATE HTTP OBJECT
  set objHTTP = createobject( "WinHttp.WinHttpRequest.5.1" )
  ''DOWNLOAD FROM URL
  objHTTP.open "GET", strURL, false
  objHTTP.send
  if objFSO.fileexists(strSAV) then
    objFSO.deletefile(strSAV)
  end if
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
  if (err.number <> 0) then
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
    errRET = 3
    objOUT.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
  end if
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  if (errRET = 0) then         															''RE-AGENT COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "RE-AGENT SUCCESSFUL : " & NOW
  elseif (errRET <> 0) then    															''RE-AGENT FAILED
    objOUT.write vbnewline & "RE-AGENT FAILURE : " & NOW & " : " & errRET
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "RE-AGENT", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - RE-AGENT COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - RE-AGENT COMPLETE" & vbnewline
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