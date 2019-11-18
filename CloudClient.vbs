''CLOUDCLIENT.VBS
''DESIGNED TO REMOTELY DOWNLOAD AND INSTALL CW CLOUD CLIENT SOFTWARE
''ACCEPTS 2 PARAMETERS , REQUIRES 2 PARAMETERS
''REQUIRED PARAMETER 'STRUSR' ; USERNAME FOR CLOUD USER TO BE CREATED
''REQUIRED PARAMETER 'STRPWD' ; PASSWORD FOR CLOUD USER TO BE CREATED
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER, strIN, strOUT
''VARIABLES ACCEPTING PARAMETERS
dim strUSR, strPWD
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH
dim objFSO, objLOG, objEXEC, objHOOK
''VERSION FOR SCRIPT UPDATE, CLOUDCLIENT.VBS , REF #2 , FIXES #15 , FIXES #16
strVER = 3
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
if (objFSO.fileexists("C:\temp\CLOUDCLIENT")) then          ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\CLOUDCLIENT", true
  set objLOG = objFSO.createtextfile("C:\temp\CLOUDCLIENT")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\CLOUDCLIENT", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\CLOUDCLIENT")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\CLOUDCLIENT", 8)
end if
''CHECK EXECUTION METHOD OF SCRIPT
strIN = lcase(mid(wscript.fullname, instrrev(wscript.fullname, "\") + 1))
if (strIN <> "cscript.exe") Then
  if (wscript.arguments.count > 0) then                     ''ARGUMENTS WERE PASSED
    for x = 0 to (wscript.arguments.count - 1)
      strTMP = strTMP & " " & chr(34) & objARG.item(x)
    next
    objOUT.write vbnewline & "SCRIPT LAUNCHED VIA EXPLORER, EXECUTING SCRIPT VIA CSCRIPT..."
    objLOG.write vbnewline & "SCRIPT LAUNCHED VIA EXPLORER, EXECUTING SCRIPT VIA CSCRIPT..."
    objWSH.run "cscript.exe //nologo " & chr(34) & Wscript.ScriptFullName & chr(34) & strTMP
    wscript.quit
  end if
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  if (wscript.arguments.count > 1) then                     ''REQUIRED ARGUMENTS PASSED
    strUSR = objARG.item(0)                                 ''SET REQUIRED PARAMETER 'STRUSR' ; USERNAME FOR CLOUD USER TO BE CREATED
    strPWD = objARG.item(1)                                 ''SET REQUIRED PARAMETER 'STRPWD' ; PASSWORD FOR CLOUD USER TO BE CREATED
    if (wscript.arguments.count > 2) then                   ''OPTIONAL ARGUMENTS PASSED
    end if
  end if
else                                                        ''NO ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  call LOGERR(1)
end if

''------------
''BEGIN SCRIPT
if (errRET <> 0) then
elseif (errRET = 0) then
  objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING CLOUDCLIENT"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING CLOUDCLIENT"
  ''CREATE LOCAL CLOUD ACCOUNT
  if (strPWD <> vbnullstring) then
    ''CREATE CLOUD USER
    objOUT.write vbnewline & now & vbtab & vbtab & " - CREATING CLOUD USER"
    objLOG.write vbnewline & now & vbtab & vbtab & " - CREATING CLOUD USER"
    call HOOK("net user " & chr(34) & strUSR & chr(34) & " " & chr(34) & strPWD & chr(34) & _
      "  /add /active:yes /expires:never /passwordchg:yes /passwordreq:yes /Y")
    ''SET PASSWORD TO NEVER EXPIRE
    objOUT.write vbnewline & now & vbtab & vbtab & " - SETTING CLOUD PASSWORD TO NEVER EXPIRE"
    objLOG.write vbnewline & now & vbtab & vbtab & " - SETTING CLOUD PASSWORD TO NEVER EXPIRE"
    call HOOK("wmic useraccount where Name='" & strUSR & "' set PasswordExpires=FALSE")
    ''ADD RMMTECH TO LOCAL ADMINISTRATORS GROUP
    objOUT.write vbnewline & now & vbtab & vbtab & " - ADDING RMMTECH TO LOCAL ADMINISTRATORS GROUP"
    objLOG.write vbnewline & now & vbtab & vbtab & " - ADDING RMMTECH TO LOCAL ADMINISTRATORS GROUP"
    call HOOK("net localgroup " & chr(34) & "Administrators" & chr(34) & " " & chr(34) & strUSR & chr(34) & " /add")
  end if
  ''DOWNLOAD CW CLOUDCLIENT SOFTWARE INSTALLER
  ''http://computerwarriorsitsupport.com/client/CloudClient.exe - SHOULD ALWAYS BE THE LATEST CLIENT
  call FILEDL("http://computerwarriorsitsupport.com/client/CloudClient.exe", "CloudClient.exe")
  ''EXECUTE CW CLOUDCLIENT SOFTWARE INSTALLER
  call HOOK(chr(34) & "C:\temp\CloudClient.exe" & chr(34) & " /s")
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub CHKAU()																					        ''CHECK FOR SCRIPT UPDATE , 'ERRRET'=10 , CLOUDCLIENT.VBS , REF #2 , FIXES #15
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
        objOUT.write vbnewline & now & vbtab & " - CloudClient :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
        objLOG.write vbnewline & now & vbtab & " - CloudClient :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
				if (cint(objSCR.text) > cint(strVER)) then
					objOUT.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					objLOG.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					''DOWNLOAD LATEST VERSION OF SCRIPT
					call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/CloudClient.vbs", wscript.scriptname)
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
          if (err.number <> 0) then
            call LOGERR(10)
          end if
					''END SCRIPT
					call CLEANUP()
				end if
			end if
		next
	end if
	set colVER = nothing
	set objXML = nothing
  if (err.number <> 0) then                                 ''ERROR RETURNED DURING UPDATE CHECK , 'ERRRET'=10
    call LOGERR(10)
  end if
end sub

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
  if (err.number <> 0) then                                 ''ERROR RETURNED DURING UPDATE CHECK , 'ERRRET'=11
    call LOGERR(11)
  end if
end sub

sub HOOK(strCMD)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND , 'ERRRET'=12
  on error resume next
  objOUT.write vbnewline & now & vbtab & vbtab & "EXECUTING : " & strCMD
  objLOG.write vbnewline & now & vbtab & vbtab & "EXECUTING : " & strCMD
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
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
		err.clear
  end if
  select case intSTG
    case 1                                                  '' 'ERRRET'=1 - NOT ENOUGH ARGUMENTS
      objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES USERNAME AND PASSWORD FOR CLOUD USER TO BE CREATED"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES USERNAME AND PASSWORD FOR CLOUD USER TO BE CREATED"
  end select
  errRET = intSTG
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  if (errRET = 0) then                                      ''SCRIPT COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - CLOUDCLIENT COMPLETE : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - CLOUDCLIENT COMPLETE : " & now
    err.clear
  elseif (errRET <> 0) then                                 ''SCRIPT FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - CLOUDCLIENT FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - CLOUDCLIENT FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE , ERROR CODE WILL BE DEFINED RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "CLOUDCLIENT", "fail")
  end if
  ''EMPTY OBJECTS
  set objEXEC = nothing
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT , RETURN ERROR
  wscript.quit err.number
end sub