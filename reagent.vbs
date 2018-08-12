''REAGENT.VBS
''DESIGNED TO AUTOMATE DOWNLOAD AND INSTALL OF WINDOWS AGENT SOFTWARE
''REQUIRES 2 PARAMETERS; 'STRCID' AND 'STRCNM'
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim retSTOP, strVER
dim objIN, objOUT, objARG, objWSH
dim objFSO, objLOG, objEXEC, objHOOK
dim strIN, strOUT, strCID, strCNM, strRCMD
''VERSION FOR SCRIPT UPDATE, RE-AGENT,#8
strVER = 2
''DEFAULT SUCCESS
retSTOP = 0
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
  if (wscript.arguments.count > 1) then                     ''
    strCID = objARG.item(0)
    strCNM = objARG.item(1)
  else                                                      ''NOT ENOUGH ARGUMENTS PASSED, END SCRIPT
    retSTOP = 1
  end if
end if
if (retSTOP <> 0) then                                      ''NO ARGUMENTS PASSED, END SCRIPT
  objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES CUSTOMER ID, CUSTOMER NAME"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES CUSTOMER ID, CUSTOMER NAME"
  call CLEANUP()
elseif (retSTOP = 0) then
	objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING RE-AGENT"
	objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING RE-AGENT"
	''AUTOMATIC UPDATE, RE-AGENT.VBS,#8
	call CHKAU()
	''DOWNLOAD WINDOWS AGENT MSI
	objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS AGENT MSI"
  objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS AGENT MSI"
  call FILEDL("https://github.com/CW-Khristos/CW_MSI/raw/master/Windows%20Agent.msi", "windows agent.msi")
  ''INSTALL WINDOWS AGENT
  objOUT.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS AGENT"
  objLOG.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS AGENT"
  ''WINDOWS AGENT RE-CONFIGURATION COMMAND
	strRCMD = "msiexec /i " & chr(34) & "c:\temp\windows agent.msi" & chr(34) & " /qb CUSTOMERID=" & strCID & _
		" CUSTOMERNAME=" & chr(34) & strCNM & chr(34) & " SERVERPROTOCOL=https SERVERPORT=443 SERVERADDRESS=ilmcw.dyndns.biz" & _
		" /l*v c:\temp\install.log ALLUSERS=2"
	''RE-CONFIGURE WINDOWS AGENT
	call HOOK(strRCMD)
end if
''END SCRIPT
call CLEANUP()

''SUB-ROUTINES
sub CHKAU()																									''CHECK FOR SCRIPT UPDATE, RE-AGENT,#8
	''ADD WINHTTP SECURE CHANNEL TLS REGISTRY KEYS
	call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp" & chr(34) & _
		" /f /v DefaultSecureProtocols /t REG_DWORD /d 0x00000A00 /reg:32")
	call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp" & chr(34) & _
		" /f /v DefaultSecureProtocols /t REG_DWORD /d 0x00000A00 /reg:64")
	set objXML = createobject("Microsoft.XMLDOM")
	objXML.async = false
	if objXML.load("https://github.com/CW-Khristos/scripts/raw/master/version.xml") then
		set colVER = objXML.documentelement
		for each objSCR in colVER.ChildNodes
			if (lcase(objSCR.nodename) = lcase(wscript.scriptname)) then
				if (cint(objSCR.text) > cint(strVER)) then
					objOUT.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					objLOG.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					if (objFSO.fileexists("C:\Program Files (x86)\N-Able Technologies\Windows Agent\cache\" & wscript.scriptname)) then
						objFSO.deletefile "C:\Program Files (x86)\N-Able Technologies\Windows Agent\cache\" & wscript.scriptname, true
					end if
					call FILEDL("https://github.com/CW-Khristos/CW_MSI/raw/master/reagent.vbs", wscript.scriptname)
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
    set objHTTP = nothing
  end if
  if (err.number <> 0) then
    retSTOP = 2
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
    retSTOP = 3
    objOUT.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
  end if
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
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
  wscript.quit err.number
end sub