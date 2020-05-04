''MSI_REAGENT_KEY.VBS
''DESIGNED TO AUTOMATE DOWNLOAD AND INSTALL OF WINDOWS AGENT SOFTWARE
''UTILIZES THE SYSTEM SPECIFIC WINDOWS AGENT MSI INSTALLER WITH CONFIGURED PARAMETERS
''ACCEPTS 2 PARAMETERS , REQUIRES 1 PARAMETER
''REQUIRED PARAMETER : 'STRKEY' , STRING TO SET ACTIVATION KEY
''OPTIONAL PARAMETER : 'STRSVR' , STRING TO SET SERVER ADDRESS
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
dim strREPO, strBRCH, strDIR
''VARIABLES ACCEPTING PARAMETERS - CONFIGURES WINDOWS AGENT MSI
dim strKEY, strSVR
dim strIN, strOUT, strRCMD
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objEXEC, objHOOK, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE , RE-AGENT.VBS , REF #2 , REF #69 , FIXES #19
strVER = 2
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
''CHECK 'PERSISTENT' FOLDERS , REF #2 , REF #73
if (not (objFSO.folderexists("C:\IT\"))) then
  objFSO.createfolder("C:\IT\")
end if
if (not (objFSO.folderexists("C:\IT\Scripts\"))) then
  objFSO.createfolder("C:\IT\Scripts\")
end if
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\MSI_REAGENT_KEY")) then      ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\MSI_REAGENT_KEY", true
  set objLOG = objFSO.createtextfile("C:\temp\MSI_REAGENT_KEY")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSI_REAGENT_KEY", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\MSI_REAGENT_KEY")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSI_REAGENT_KEY", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  if (wscript.arguments.count >= 1) then                    ''SET REQUIRED VARIABLES ACCEPTING ARGUMENTS
    strKEY = objARG.item(0)                                 ''SET REQUIRED PARAMETER 'STRKEY' , ACTIVATION KEY
    if (wscript.arguments.count = 1) then                   ''NO OPTIONAL ARGUMENTS PASSED
      strSVR = "ncentral.cwitsupport.com"                   ''SET OPTIONAL PARAMETER 'STRSVR' , 'DEFAULT' SERVER ADDRESS
    elseif (wscript.arguments.count > 1) then               ''OPTIONAL ARGUMENTS PASSED
      strSVR = objARG.item(1)                               ''SET OPTIONAL PARAMETER 'STRSVR' , PASSED SERVER ADDRESS ; SEPARATE MULTIPLES WITH ','
      if (strSVR = vbnullstring) then                       ''OPTIONAL 'STRSVR' ARGUMENT EMPTY
        strSVR = "ncentral.cwitsupport.com"                 ''SET OPTIONAL PARAMETER 'STRSVR' , 'DEFAULT' SERVER ADDRESS
      end if
    end if
  end if
elseif (wscript.arguments.count = 0) then                   ''NOT ENOUGH ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  call LOGERR(1)
end if

''------------
''BEGIN SCRIPT
if (errRET = 0) then                                        ''ARGUMENTS PASSED, CONTINUE SCRIPT
	objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING MSI_REAGENT_KEY"
	objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING MSI_REAGENT_KEY"
	''AUTOMATIC UPDATE, MSI_REAGENT_KEY.VBS, REF #2 , REF #68 , REF #69 , FIXES #8
  ''DOWNLOAD CHKAU.VBS SCRIPT, REF #2 , REF #68 , REF #69
  call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/chkAU.vbs", "chkAU.vbs")
  ''EXECUTE CHKAU.VBS SCRIPT, REF #69
  objOUT.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : MSI_REAGENT_KEY : " & strVER
  objLOG.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : MSI_REAGENT_KEY : " & strVER
  intRET = objWSH.run ("cmd.exe /C " & chr(34) & "cscript.exe " & chr(34) & "C:\temp\chkAU.vbs" & chr(34) & " " & _
    chr(34) & strREPO & chr(34) & " " & chr(34) & strBRCH & chr(34) & " " & chr(34) & strDIR & chr(34) & " " & _
    chr(34) & wscript.scriptname & chr(34) & " " & chr(34) & strVER & chr(34) & " " & _
    chr(34) & strKEY & "|" & strSVR & chr(34) & chr(34), 0, true)
  ''CHKAU RETURNED - NO UPDATE FOUND , REF #2 , REF #69 , REF #68
  objOUT.write vbnewline & "errRET='" & intRET & "'"
  objLOG.write vbnewline & "errRET='" & intRET & "'"
  intRET = (intRET - vbObjectError)
  objOUT.write vbnewline & "errRET='" & intRET & "'"
  objLOG.write vbnewline & "errRET='" & intRET & "'"
  if ((intRET = 4) or (intRET = 10) or (intRET = 11) or (intRET = 1) or (intRET = 2147221505) or (intRET = 2147221517)) then
    objOUT.write vbnewline & now & vbtab & vbtab & " - NO UPDATE FOUND : MSI_REAGENT_KEY : " & strVER
    objLOG.write vbnewline & now & vbtab & vbtab & " - NO UPDATE FOUND : MSI_REAGENT_KEY : " & strVER
    ''DOWNLOAD WINDOWS AGENT MSI , 'ERRRET'=2 , REF #2 , FIXES #13
    objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS AGENT SYSTEM-SPECIFIC MSI"
    objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS AGENT SYSTEM-SPECIFIC MSI"
    call FILEDL("https://github.com/CW-Khristos/CW_MSI/raw/master/Windows%20Agent.msi", "windows agent.msi")
    if (errRET <> 0) then
      call LOGERR(2)
    end if
    ''INSTALL WINDOWS AGENT
    objOUT.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS AGENT"
    objLOG.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS AGENT"
    ''WINDOWS AGENT RE-CONFIGURATION COMMAND , REF #2 , FIXES #13 , FIXES #19
    strRCMD = "msiexec /i " & chr(34) & "c:\temp\windows agent.msi" & chr(34) & " /qn" & _
      " AGENTACTIVATIONKEY=" & chr(34) & strKEY & chr(34) & " SERVERPROTOCOL=https:// SERVERPORT=443 SERVERADDRESS=" & chr(34) & strSVR & chr(34) & _
      " /l*v c:\temp\agent_install.log ALLUSERS=2"
    ''RE-CONFIGURE WINDOWS AGENT , 'ERRRET'=3
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
  if ((err.number <> 0) and (err.number <> 58)) then        ''ERROR RETURNED DURING DOWNLOAD , 'ERRRET'=11
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
      objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES ACTIVATION KEY"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES ACTIVATION KEY"
  end select
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  on error resume next
  if (errRET = 0) then         															''MSI_REAGENT_KEY COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "MSI_REAGENT_KEY SUCCESSFUL : " & errRET & " : " & now
    objLOG.write vbnewline & "MSI_REAGENT_KEY SUCCESSFUL : " & errRET & " : " & now
    err.clear
  elseif (errRET <> 0) then    															''MSI_REAGENT_KEY FAILED
    objOUT.write vbnewline & "MSI_REAGENT_KEY FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & "MSI_REAGENT_KEY FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "MSI_REAGENT_KEY", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - MSI_REAGENT_KEY COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - MSI_REAGENT_KEY COMPLETE" & vbnewline
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