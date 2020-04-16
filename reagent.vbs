''RE-AGENT.VBS
''DESIGNED TO AUTOMATE DOWNLOAD AND INSTALL OF WINDOWS AGENT SOFTWARE
''UTILIZES THE CUSTOMER SPECIFIC WINDOWS AGENT EXE INSTALLER WITH CUSTOMER ID
''ACCEPTS 2 PARAMETERS , REQUIRES 1 PARAMETERS
''REQUIRED PARAMETER : 'STRCID' , STRING TO SET CUSTOMER ID
''OPTIONAL PARAMETER : 'STRSVR' , STRING TO SET SERVER ADDRESS
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
dim strREPO, strBRCH, strDIR
''VARIABLES ACCEPTING PARAMETERS - CONFIGURES WINDOWS AGENT MSI
dim strIN, strOUT, strRCMD, strCID
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objEXEC, objHOOK, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE , RE-AGENT.VBS , REF #2 , REF #69 , FIXES #8 , FIXES #13
strVER = 13
strREPO = "CW_MSI"
strBRCH = "dev"
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
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\REAGENT")) then              ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\REAGENT", true
  set objLOG = objFSO.createtextfile("C:\temp\REAGENT")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\REAGENT", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\REAGENT")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\REAGENT", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  if (wscript.arguments.count > 0) then                     ''SET REQUIRED VARIABLES ACCEPTING ARGUMENTS
    strCID = objARG.item(0)                                 ''SET REQUIRED PARAMETER 'STRCID' , CUSTOMER ID
  end if
elseif (wscript.arguments.count = 0) then                   ''NOT ENOUGH ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  call LOGERR(1)
end if

''------------
''BEGIN SCRIPT
if (errRET = 0) then                                        ''ARGUMENTS PASSED, CONTINUE SCRIPT
	objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING RE-AGENT"
	objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING RE-AGENT"
	''AUTOMATIC UPDATE, RE-AGENT.VBS, REF #2 , REF #68 , REF #69 , FIXES #8
  ''DOWNLOAD CHKAU.VBS SCRIPT, REF #2 , REF #68 , REF #69
  call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/chkAU.vbs", "chkAU.vbs")
  ''EXECUTE CHKAU.VBS SCRIPT, REF #69
  objOUT.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : RE-AGENT : " & strVER
  objLOG.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : RE-AGENT : " & strVER
  intRET = objWSH.run ("cmd.exe /C " & chr(34) & "cscript.exe " & chr(34) & "C:\temp\chkAU.vbs" & chr(34) & " " & _
    chr(34) & strREPO & chr(34) & " " & chr(34) & strBRCH & chr(34) & " " & chr(34) & strDIR & chr(34) & " " & _
    chr(34) & wscript.scriptname & chr(34) & " " & chr(34) & strVER & chr(34) & " " & _
    chr(34) & strCID & "|" & strSVR & chr(34), 0, true)
  ''CHKAU RETURNED - NO UPDATE FOUND , REF #2 , REF #69 , REF #68
  intRET = (intRET - vbObjectError)
  if ((intRET = 4) or (intRET = 10) or (intRET = 11) or (intRET = 1)) then
    ''DOWNLOAD WINDOWS AGENT MSI , 'ERRRET'=2 , REF #2 , FIXES #13
    objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS AGENT CUSTOMER-SPECIFIC EXE"
    objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS AGENT CUSTOMER-SPECIFIC EXE"
    call FILEDL("http://ncentral.cwitsupport.com/dms/FileDownload?customerID=" & strCID & "&softwareID=101", strCID & "WindowsAgentSetup.exe")
    if (errRET <> 0) then
      call LOGERR(2)
    end if
    ''INSTALL WINDOWS AGENT
    objOUT.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS AGENT"
    objLOG.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS AGENT"
    ''WINDOWS AGENT RE-CONFIGURATION COMMAND , REF #2 , FIXES #13
    strRCMD = "c:\temp\" & strCID & "WindowsAgentSetup.exe -ai"
    objOUT.write vbnewline & now & vbtab & vbtab & " - EXECUTING : " & strRCMD
    objLOG.write vbnewline & now & vbtab & vbtab & " - EXECUTING : " & strRCMD
    call HOOK(strRCMD)
    if (errRET <> 0) then
      call LOGERR(3)
    end if
  end if
elseif (errRET <> 0) then                                   ''NO ARGUMENTS PASSED, END SCRIPT , 'ERRRET'=1
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
      objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES CUSTOMER ID"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES CUSTOMER ID"
  end select
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  on error resume next
  if (errRET = 0) then         															''RE-AGENT COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "RE-AGENT SUCCESSFUL : " & errRET & " : " & now
    err.clear
  elseif (errRET <> 0) then    															''RE-AGENT FAILED
    objOUT.write vbnewline & "RE-AGENT FAILURE : " & errRET & " : " & now
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
  wscript.quit err.number
end sub