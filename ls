'RHC.VBS
'Recent update: 01/10/06: erudakov added LOGON exe for Websense
'Recent update: 01/20/06: tnovak add lpt1: mapping for group PRT_SMEPFS02
'Recent update: 01/22/06: tnovak REMARKED OUT workstation printer mapping FOR 'SMN2K8269A'USE CHANGE TO PRT_SMNFC_PA05
'Recent update: 01/30/06: tnovak CHANGED LPT1: printer for workstation SFHXP4542A printer mapping to SFHPF05
'Recent update: 01/30/06: tnovak Added STANZA for SFH Patient Finanical Service print queues PRT_SFHPF05
'Recent update: 02/01/06: tnovak Added STANZA for SFH Patient Finanical Service print queues PRT_SFHPF09, PRT_SFHPF10, and REMARKED OUT workstation lpt1: mappings
'Recent update: 02/02/06: tnovak REMARKED OUT workstation printer mapping FOR LPT1:
'Recent update: 02/02/06: jpaugys migrated impresss from kreg server to wsmcapps1 server \\wsmcapps1\apps\impress
'Recent update: 02/02/06: jpaugys migrated kreg from kreg server to wsmcapps1 server \\wsmcapps1\apps\kreg
'Recent update: 02/10/06: bmitchell remarked out the statement to email if a vbs file is in the start folder
'Recent update: 02/22/06: PBernacki changed the server drive mapping for App_MSO_Users to RHCAPPS234 from RHCAPPS7
'Recent update: 03/02/06: Jpaugys added share for apollo application
'Recent update: 03/06/06: Jpaugys added shares rhc ressched group
'Recent update: 03/09/06: TNovak added PRT_WHPF03 printer group
'Recent update: 03/27/06: Jpaugys added share for Spiritual Services and Mission
'Recent update: 03/27/06: Jpaugys added printer for RMCXP5558A at \\RMCPRINT\RMCPATREG
'Recent update: 03/27/06: Jpaugys addedshare and mappping for wsmcfs01\shared\WSMC\RMC IRB
'Recent update: 03/27/06: Jpaugys addedshare and mappping for wsmcfs01\shared\WSMC\residenct emr assesment
'Recent update: 03/27/06: Jpaugys consolidated all of the apollo groups with the same access rights
'Recent update: 03/27/06: kpokryfke added areprinter mapping for olrxp3365a to olrphar02
'Recent update: 04/03/06: TNOVAK added printer mapping for group prt_hfmcpf08
'Recent update: 04/10/06: CHALLER edited app_disable_account message
'Recent update: 04/10/06: JEbmeier edited RMCEMRP to RMCFS01
'Recent update: 04/24/06: JFiorella edited app_disable_account message
'Recent update: 04/25/06: JPaugys added share for rhc family medical
'Recent update: 04/25/06: JPaugys added share for smn rehab
'Recent update: 04/10/06: JEbmeier edited RMCAPPS2/RMCACCT to RMCAPPS6; delete printer mapping to RMCAPPS2\CANCER
'Recent update: 05/05/06: TNovak added group 'App_CancerReg_User' drive mapping for P:
'Recent update: 04/10/06: JEbmeier edited Global Groups mapping RMCEMRP to RMCAPPS6
'Recent update: 05/16/06: JPAUGYS ADDED GROUP AND DRIVE FOR SJH Facilities Management
'Recent update: 06/05/06: JPAUGYS ADDED GROUP AND DRIVE FOR WSMC Volunteer Works users
'Recent update: 06/12/06: TNOVAK ADDED GROUP TEST FOR MEMBERS OF 'App_eCareCMS_Users group
'Recent update: 06/29/06: TNOVAK ADDED GROUP TEST FOR MEMBERS OF 'App_eCareCMS_Users group to map the Q: drive to RHCAPPS254\EcareCMS
'Recent update: 06/29/06: MGuy ADDED Drive Map for App_Medicall_Admins members
'Recent update: 07/07/06: MGuy Changed ADSI Install (InstallADSI) & File Registration call
'Recent update: 07/10/06: MGuy Put Back ADSI Install (InstallADSI) & File Registration call
'Recent update: 07/10/06: MGuy Updated files (RICHTX32.OXC & COMCTL32.OCX) in ..\netlogon\bin used by (InstallADSI)
'Recent update: 07/14/06: Jpaugys Changed EPO agent login push
'Recent update: 08/01/06: TNovak added workstation test for default printer, OLRXP4307B, OLRXP3837E,OLRXP3837C,OLRXP3837D
'Recent update: 08/01/06: JEbmeier modified the ERAS group to point to RHCAPPS249 instead of RHCAPPS5, and added mapping w:\ to \\RHCAPPS249\APPS
'Recent update: 08/01/06: JTalbutt added drive map for App_Noah_RMC members
'Recent update: 09/28/06: BMitchell added printer group for PRT_SFHHCLIN01
'Recent Update: 10/23/06: PNielander added drive map for App_Sitemaster_Users group to map the ?: drive to SMNAPPS3\Apps\SiteMaster\LiveData)
'Recent Update: 10/23/06: Jpaugys added I drive for app_wsmc_impac_users group
'Recent Update: 11/13/06: Jpaugys added drive for DMI 
'Recent Update: 11/13/06: Jpaugys added drive for Rehab_Lab_Division
'Recent Update: 12/07/06: Added P:\ Drive mapping for App_ChartOne_RMC on RMCAPPS6
'Recent update: 12/14/06: TNOVAK ADDED GROUP TEST FOR MEMBERS OF 'App_eCareCMS_Users' AND 'APC-3M Encoder' group
'Recent update: 12/15/06: TNOVAK REMOVED GROUP TEST FOR MEMBERS OF 'App_eCareCMS_Users' AND 'APC-3M Encoder' group
'Recent update: 2/12/07: jpaugys added wsmc ftp drive s:
'Recent update: 2/13/07: jpaugys added ip op drive O:
'Recent update: 2/20/07:  jpaugys added printer mapping for PRT_SKOHH15 \\RMCPRINT\SKOHH1
'Recent update: 4/06/07:  kpokryfke replaced log on script with copy from earlier today due to log on errors after 2 pm
'Recent update: 4/09/07:  BMitchell removed printer groups for RMCGAXX printers
'Recent update: 4/17/07:  JFiorella added T:\ Drive Mapping for SJH Chief Residents
'Recent update: 4/25/07:  BMitchell removed printer group for SKOHH03 printer
'Recent update: 5/01/07:  JPaugys rem' install of epo agent 3.5.5.438
'Recent update: 5/08/07:  Jpaugys Created Y: drive for nursing education     \\rhcfs01\shared\rhc\RHC_Nursing_Education_Council
'Recent update: 5/30/07:  BMitchell removed mapping for the I: drive for Softmed, Charlie Moore request.
'Recent update: 5/31/07:   Jpaugys Added q: drive for wsmc payroll conversion   \\rhcraftp101\ftproot$\v_iatric
'Recent update: 7/16/07:   Jpaugys added r: drive for ces seh reports \\rhcraftp101\sehreports$
'Recent update: 7/23/07:   Jpaugys added r: drive for ces wh reports \\rhcraftp101\whreports$
'Recent update: 7/26/07:   Jpaugys added r: drive for ces sfh reports \\rhcraftp101\sfhreports$
'Recent update: 08/17/07: BMitchell added user and admin group for Drive R: mapping \\SFHFS01\Shared\TSAContracts
'Recent update: 09/17/07: Jpaugys added r drive for all hospitals with access to folders on a hospital level and removed each hospital r drive
'Recent update: 10/09/07: Jpaugys added p drive for ces 
'Recent update: 10/25/07: Jpaugys added I drive to map all vendor ftp folders without providing access
'Recent update: 10/25/07: Erudakov added WS test LogonApp.exe execution
'Recent update: 10/30/07: MGUY added p drive for App_CDMP_RMC
'Recent update: 12/11/07: Jpaugys added t drive for app_rmcorschedule
'Recent update: 12/18/07: MGUY added H: drive map for RLICHS01_Home
'Recent update: 1/10/08: JEbmeier removed references to RMCTRANS
'Recent update: 3/04/08: Jpaugys added t drive for wsmc_hr \\rmcfs01\SHARED\wsmc_hr
'Recent update: 4/08/08:   made current epo version 3.6.0.574
'Recent update: 4/21/08:   made current epo version 3.6.0.603
'Recent update: 5/13/08:   addec UDSPRO lnk copy EER
'Recent update: 5/20/08: Changed SFH MRS Server from SFHAPPS4 to SFHAPPS5
'Recent update: 5/21/08: DSagan - Added group "RHCMASTER SHARED" to map users 'G Drive' to "\\rhcmaster\shared"
'Recent update: 5/27/08: Jpaugys added p drive for app_psc_reports \\rhcraftp101\ftproot$\V_PSC\Reports_From_PSC"
'Recent update: 7/22/08: PBernacki added n drive for app_mso_test \\rhcapps234test\apps"
'Recent update: 8/29/08: MGUY adde "APP_HPP_QSDeskTop" to place Physician Portal shortcut on QuickStart User desktops
'Recent update: 09/09/07: BMitchell Removed printer group mappings after PFS moves.
'Recent update: 11/10/08: PBernacki added s: drive for group RHC_RACS_Daily_Operations
'Recent update: 11/12/08:jpaugys added q: drive for group app_systoc_collections
'Recent update: 11/21/08: PBernacki added s: drive for the group App_Pracplus_BOReports_WSMC
'recent update: 12/26/08: jpaugys Added T drive for app_sjhorschedule
'recent update: 1/05/09:  jpaugys removed T drive for app_sjhorschedule
'recent update: 1/05/09:  jpaugys Added T drive for APP_RHC_SCHEDULING
'recent update: 1/26/09:  jpaugys Added v drive for wsmc_vellos_users
'Recent Update: 2/12/09: JFiorella Changed J: Drive mapping for App_MrsWin to RMCAPPS2
'Recent Update: 2/12/09: JFiorella Changed P:\ Drive mapping for App_ChartOne_RMC to RMCAPPS2
'Recent Update: 2/12/09: JFiorella Added P:\ Drive mapping for RMC Medical Library to RMCAPPS2
'Recent update: 3/12/09: JFiorella Changed p drive for App_CDMP_RMC from RMCAPPS6 (was remarked out) to RMCAPPS2
'Recent Update: 3/18/09: JFiorella Added F:\ Drive mapping for ScottCare (RMC Cardiac Rehab) to RMCAPPS2
'Recent Update: 7/08/09: JFiorella Changed Heartbase and EMS from SJHAPPS1 to SJHAPPS68
'Recent Update: 7/08/09: JFiorella Changed MRS Tumor Registry from SMNAPPS3 to SMEAPPS5
'Recent Update: 7/30/09: PBernacki Added J: drive for SEH PCCH for RHCAPPS27\Shared
'Recent Update: 8/10/09: jpaugys Added Z: drive \\rhcraftp101\ftproot$\v_ics
'Recent update: 8/29/09: MGUY added "APP_HPP_QA" to place QA Test Physician Portal shortcut on User desktops
'Recent update: 10/20/09: JFiorella Changed and added mappings drive P for APP_METRIQ_xxx @ SME,SJH,SFH,RMC & WSMC
'Recent update: 10/20/09: RSCHULTZ Remarked all lines for Q: drive mapping for 'bodyworks' global group
'Recent update: 11/17/09: MGUY added "APP_HPPx_DeskTop" to place Physician Portal Extended Timeout shortcut on User desktops


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NOTE: DO NOT CHANGE THE FIRST SECTION OF THE SCRIPT.
'DO NOT CHANGE THE AREA (between the asterisks) 
'It set variables, defines functions, gathers information on the computer, 
'creates objects, installs ADSI, and provides Script Functions.
'ALL MODIFICATIONS SHOULD BE MADE TO THE END of the script.
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Initialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
DIM HASPNAGENT
DIM INSTALLPNAGENT
DIM PNMAYBE
DIM PNPOSSIBLE
Dim WshFS
'Dim WshF
Dim WshFDC
Dim WshFile
'Dim WshFSource
'Dim WshAE
'Dim WshTSc
'Dim WshSW
Dim WshReg
Dim WSHNetwork
Dim WshShell
Dim WshProcessEnvironment
Dim WshLogonServer
Dim WshSysRoot
Dim EnvWinDir
Dim EnvSystem
'Dim IPAddress
'Dim SubnetMask
Dim OS
Dim OSreg
Dim Username
Dim strUserID
Dim UserObj
Dim UseNTServer
'''Dim RLICISShareSvr
Dim DomainName
Dim Port
Dim resdom1
Dim resdom2
DIM collapseme
Dim goaway

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Set Static Variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''RLICISShareSvr="RLICSMS"	'Sets the IS NT Server that drives are mapped to
'''DomainName = "RMC_DCBKP"	'Sets the users domain to specific domain controller for testing purposes
DomainName = "reshealthcare.org"	'Sets the users domain
'UseNTServer = "no"		'Sets the running of this logon script on Windows NT Servers to no. 
UseNTServer = "yes"		'Sets the running of this logon script on Windows NT Servers to Yes. 


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Begin Script
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
Set WshFS = WScript.CreateObject("Scripting.FileSystemObject")
Set WSHNetwork = WScript.CreateObject("WScript.Network")
Set WshShell = WScript.CreateObject("WScript.Shell")
Set WshProcessEnvironment = WshShell.Environment("Process")

PCName = UCase(WsHNetwork.ComputerName)

If (Left(pcname,5) = "RHCTS") Or (Left(pcname,7) = "RHCRATS") Then 
   WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\LogonApp.exe http://10.250.9.180:15880 /COPY /NOPERSIST ",0,true
   WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\LogonApp.exe http://10.40.8.180:15880 /COPY /NOPERSIST ",0,True
Else
   WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\LogonApp.exe http://10.250.9.180:15880 /COPY /DHCP ",0,true
   WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\LogonApp.exe http://10.40.8.180:15880 /COPY /DHCP ",0,true
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Determine OS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
OS = WshProcessEnvironment("OS")
If OS <> "Windows_NT" Then
	OS = "Win32"
	WshLogonServer = "\\RLICDC1"
	WshLogonPath = WshLogonServer & "\netlogon"
	WshInstall = WshLogonPath & "\Bin\"
	EnvWinDir = WshProcessEnvironment("windir")
	WshSys32 = EnvWinDir & "\system\"
End If

If UseNTServer = "yes" Then	
     If OS = "Windows_NT" Then 
          WshLogonServer = WshShell.ExpandEnvironmentStrings("%LOGONSERVER%")
          WshSysRoot = WshShell.ExpandEnvironmentStrings("%SYSTEMROOT%")
          WshLogonPath = WshLogonServer & "\netlogon"
          WshInstall = WshLogonPath & "\Bin\"
          EnvWinDir = WshProcessEnvironment("windir")
          WshSys32 = EnvWinDir & "\system32\"
          EnvSystem = WshProcessEnvironment("systemdrive")
          wshhomeshare = WshProcessEnvironment("homeshare")
          wshhomepath = WshProcessEnvironment("homepath")
          wshhomedata = (wshhomeshare + wshhomepath)
     End If
Else
     If OS = "Windows_NT" Then
          If WshShell.RegRead ("HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions\ProductType") = "WinNT" Then
               WshLogonServer = WshShell.ExpandEnvironmentStrings("%LOGONSERVER%")
               WshSysRoot = WshShell.ExpandEnvironmentStrings("%SYSTEMROOT%")
               WshLogonPath = WshLogonServer & "\netlogon"
               WshInstall = WshLogonPath & "\Bin\"
               EnvWinDir = WshProcessEnvironment("windir")
               WshSys32 = EnvWinDir & "\system32\"
               EnvSystem = WshProcessEnvironment("systemdrive")
          Else
               Dim Box
               Box = MsgBox ("Will NOT run WSH Login Script on a Windows NT Server - Exiting Logon Script!", 64, "Login to " & RHCMASTER)
               WScript.Quit
          End If
     Else
     End If
End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Determine Computer Host Name
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
PCName = WsHNetwork.ComputerName
PCName = UCase(PCName)
'WSCRIPT.ECHO PCName
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Create Recycle Bin if it is not there 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
IF wshfs.FolderExists(envsystem + "\RECYCLER") Then
 ' wscript.quit
Else
    WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\tqcrunas.exe /f %logonserver%\netlogon\createRCB1.cmd", 0, True 
end If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'The following section checks for the existence of the NOLOGINSCRIPT.TXT file 
'in the root ofthe system drive.  If it exists, exit this script.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
IF wshfs.fileexists(envsystem + "\NOLOGINSCRIPT.TXT") then
     wscript.quit
end if


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Init Variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'The following set of code will check the OS and will change a registry key based on that.
'It leaves Windows NT and below alone.
dim OSIS
OSIS = "nonzero"
csn = 0
If OS = "Windows_NT" Then
     wshshell.run "%COMSPEC% /c %logonserver%\netlogon\desktop.cmd ", 0, True 
     OSreg = WshShell.RegRead ("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
     REa = 0
          If OSreg = "" Then
               WshShell.Run "%COMSPEC% /c net time \\rlicdc1 /set /y",0,true
               REa = 3
               OSIS = "UNKNOWN"
          ElseIf OSreg = "Microsoft Windows 2000" Then
               WshShell.Run "%COMSPEC% /c net time \\reshealthcare.org /set /y",0,true
               REa = 1
               OSIS = "2K"
               On Error Resume Next
               strComputer = "."
               Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
               Set colItems = objWMIService.ExecQuery("Select caption from Win32_OperatingSystem",,48)
               For Each objItem in colItems
'                    Wscript.Echo objItem.Caption
                    IF objitem.caption = "Microsoft Windows 2000 Professional" then
                         OSIS = "2KPRO"
                    END if
               Next
'         If csn is made to equal 1, then the csnw.cmd file below will run.
'         csn = 1
          ElseIf OSreg = "Microsoft Windows XP" Then
               WshShell.Run "%COMSPEC% /c net time \\reshealthcare.org /set /y",0,true
               REa = 1
               OSIS = "XP"
          ElseIf Left(OSreg,13) = "Windows Vista" Then
               WshShell.Run "%COMSPEC% /c net time \\reshealthcare.org /set /y",0,true
               REa = 1
               OSIS = "VISTA"     
          ElseIf OSreg = "Microsoft Windows Server 2003" Then
               WshShell.Run "%COMSPEC% /c net time \\reshealthcare.org /set /y",0,true
               REa = 1
               OSIS = "2K3"
          End If
          IF REa = 1 THEN
               If WshShell.RegRead ("HKLM\SYSTEM\CurrentControlSet\Control\Print\Providers\LanMan Print Services\servers\addprinterdrivers") = "1" Then
                    wshshell.run "%COMSPEC% /c %logonserver%\netlogon\printerdriver.cmd ", 0, True 
'                    wshshell.run "%COMSPEC% /c %logonserver%\netlogon\printerdriver.cmd " & envwindir & nullFile, 0, True 
               End If
          End If
End If


'The following section checks for the existence of a file indicating the PN Agent installed.
'If it exists, it sets a variable.  There are two different versions of the PN Agent.
IF OS = "Windows_NT" THEN
     HASPNAGENT = 0
     INSTALLPNAGENT = 0
     IF wshfs.fileexists(envsystem + "\Program Files\Citrix\ICA Client\pn.exe") then
          HASPNAGENT = 1
     elseif wshfs.fileexists(envsystem + "\PROGRAM FILES\CITRIX\PNAGENT\PNAGENT.EXE") then
          HASPNAGENT = 1
     end if
     IF HASPNAGENT = 1 THEN
          INSTALLPNAGENT = 0
     ELSEIF HASPNAGENT = 0 THEN
          INSTALLPNAGENT = 1
     END IF
END IF

PNPOSSIBLE = 0
IF OSIS = "2KPRO" THEN
     PNPOSSIBLE = 1
END IF

IF OSIS = "XP" THEN
     PNPOSSIBLE = 1
END IF
IF PCName = "RMCXP1298A" OR PCName = "RMCAPPS4" OR PCName = "RMC2KCOMMFAX" OR PCName = "RMC2KCOMMPAGE" OR PCName = "RMC2KCOMMADT" OR PCName = "RMC2K5200A" OR PCName = "RMC2K5200B" OR PCName = "RMC2K5200C" OR PCName = "RMC2K5200D" OR PCName = "RMC2K5200E" OR PCName = "RMC2K5200F" OR PCName = "RMC2K5200G" OR PCName = "RMC2K5200H" OR PCName = "RMC2K5200I" OR PCName = "OLR2K8400A" OR PCName = "OLR2K8400B" OR PCName = "OLR2K8400C" OR PCName = "OLR2K8400D" OR PCName = "OLR2K8400E" OR PCName = "SWITCHROOM1" OR PCName = "TELCOMM3" OR PCName = "SMN2K1099A" OR PCName = "SMN2K1099C" OR PCName = "SMN2K1099D" OR PCName = "SMN2K1099B" OR PCName = "RHCAPPS10" THEN
PNPOSSIBLE = 0
END IF

'The next section determines whether or not the device is on a subnet that is supposed 
strComputer = "."
Set objWMIService = GetObject( _ 
    "winmgmts:\\" & strComputer & "\root\cimv2")
Set IPConfigSet = objWMIService.ExecQuery _
    ("Select IPAddress from Win32_NetworkAdapterConfiguration ")
 
For Each IPConfig in IPConfigSet
    If Not IsNull(IPConfig.IPAddress) Then 
        For i=LBound(IPConfig.IPAddress) _
            to UBound(IPConfig.IPAddress)
'                WScript.Echo IPConfig.IPAddress(i)
'                wscript.echo LEN(IPCONFIG.IPADDRESS(i))
                dot1 = instr(Ipconfig.ipaddress(i), ".")
                dot2 = instr(dot1+1,Ipconfig.ipaddress(i), ".",1)
                dot3 = instr(dot2+1,Ipconfig.ipaddress(i), ".",1)
                dot4 = left(ipconfig.ipaddress(i),dot3-1)
          IF dot4 = "10.170.236" OR dot4 = "10.170.237" OR dot4 = "10.170.238" OR dot4 = "10.170.239" OR dot4 = "10.170.240" OR dot4 = "10.170.241" OR dot4 = "10.170.242" OR dot4 = "10.170.243" OR dot4 = "10.170.244" OR dot4 = "10.170.245" OR dot4 = "10.170.246" OR dot4 = "10.170.247" OR dot4 = "10.170.248" OR dot4 = "10.170.249" OR dot4 = "10.170.250" OR dot4 = "10.170.251" OR dot4 = "10.170.252" OR dot4 = "10.170.253" OR dot4 = "10.170.254" OR dot4 = "10.170.255" OR dot4 = "10.40.10" OR dot4 = "10.40.11" OR dot4 = "10.40.232" OR dot4 = "10.40.233" OR dot4 = "10.40.234" OR dot4 = "10.40.235" OR dot4 = "10.40.248" OR dot4 = "10.40.249" OR dot4 = "10.40.250" OR dot4 = "10.40.251" OR dot4 = "10.40.8" OR dot4 = "10.40.9" OR dot4 = "10.60.10" OR dot4 = "10.60.11" OR dot4 = "10.60.152" OR dot4 = "10.60.20" OR dot4 = "10.60.21" OR dot4 = "10.60.22" OR dot4 = "10.60.23" OR dot4 = "10.60.8" OR dot4 = "10.60.9" OR dot4 = "10.90.10" OR dot4 = "10.90.11" OR dot4 = "10.90.12" OR dot4 = "10.90.13" OR dot4 = "10.90.14" OR dot4 = "10.90.15" OR dot4 = "10.90.16" OR dot4 = "10.90.17" OR dot4 = "10.90.18" OR dot4 = "10.90.19" OR dot4 = "10.90.20" OR dot4 = "10.90.21" OR dot4 = "10.90.22" OR dot4 = "10.90.23" OR dot4 = "10.90.8" OR dot4 = "10.90.9" THEN
               PNSUBNET = 1
          END IF
        Next
    End If
Next
PNDOIT = PNPOSSIBLE AND INSTALLPNAGENT AND PNSUBNET
IF PNDOIT = 1 THEN
'     WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\tqcrunas /f %logonserver%\netlogon\pnagent",0,true
'     wscript.quit
END IF

'wscript.echo osis
'wscript.echo pnpossible
'wscript.echo haspnagent
'wscript.echo installpnagent
'wscript.echo pnsubnet

'epo  installer
'IF OS = "Windows_NT" THEN
'     DIM EPOGOOD
'     DIM EPOIS
'     EPOGOOD = "DOESNOTEXIST"
'     EPOIS = "nonzero"
'     IF OSIS = "2K" OR OSIS = "XP" OR OSIS = "2KPRO" OR OSIS = "2K3" THEN
'          EPOIS = WshShell.RegRead ("HKLM\SOFTWARE\Network Associates\ePolicy Orchestrator\Application Plugins\ePOAgent3000\Version")
'          IF EPOIS <> "3.6.0.603" Then
'               EPOGOOD = 0
'               IF PCName = "THERATRAC" THEN
'                    EPOGOOD = 1
'               END IF
'               IF EPOGOOD = 0 THEN
'                    WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\tqcrunas /f %logonserver%\netlogon\mcaf",0,true
'                   WshShell.Run "%comspec% /c %logonserver%\netlogon\needepo.bat", 0, True
'               END IF
'          END IF
'     END IF
'END IF
'end of epo installer


If csn = 1 Then
     wshshell.run "%COMSPEC% /c %logonserver%\netlogon\csnw.cmd ", 0, True 
End If

If OS = "Windows_NT" Then
     wshshell.run "%COMSPEC% /c %logonserver%\netlogon\pcitrix.cmd ", 0, True 
End If
XPSPa = 0
XPSPb = 0
XPSPc = 0

     IF REa = 1 THEN
          If WshShell.RegRead ("HKLM\software\microsoft\windows nt\currentversion\CSDVersion") = "Service Pack 2" Then
               XPSPa = 1
          End If
     End If

     If WshShell.RegRead ("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName") = "Microsoft Windows XP" Then
          XPSPb = 1
     End If

XPSPc = XPSPa and XPSPb

IF XPSPc = 1 then
     wshshell.run "%COMSPEC% /c %logonserver%\netlogon\xpsp2.cmd ", 0, True 
'     wshshell.run "%COMSPEC% /c %logonserver%\netlogon\xpsp2.cmd " & envwindir & nullFile, 0, True 
End If


'The following set of code will check for the existence of a .vbs file in the all users startup folder and send a based on that.
'IF REa = 1 THEN
'     wshshell.run "%COMSPEC% /c %logonserver%\netlogon\vbsstart.cmd ", 0, True 
'End If
'IF REa = 3 THEN
'     wshshell.run "%COMSPEC% /c %logonserver%\netlogon\vbsstaNT.cmd ", 0, True 
'End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next     ' Continue when error found in script
Do
strUserID = WSHNetwork.UserName
Loop Until strUserID <> ""
''Next line will display userid to screen
'WSCRIPT.ECHO STRUSERID

'Get User Object
if REa = 3 then
     adsipath="WinNT://" & "RHCMASTER" & "/" & strUserID
end if
if REa = 1 then
'     Set userobj = GetObject( "LDAP://CN=" & struserid & ";CN=Users;DC=reshealthcare,DC=org" )
''''''''''''''''''''''''''''''''''''''''''''''''''
Dim UserGroups
UserGroups=""
Set objSysInfo = CreateObject("ADSystemInfo")
Set objNetwork = CreateObject("Wscript.Network")

strUserPath = "LDAP://" & objSysInfo.UserName
Set objUser = GetObject(strUserPath)
'set userobj = struserpath
'For Each strGroup in objUser.MemberOf                       @@ERR 07/22/05 begin change
'     strGroupPath = "LDAP://" & strGroup
'     Set objGroup = GetObject(strGroupPath)
'UserGroups=UserGroups & "[" & "CN=" & objgroup.cn & "]"
'Next
If VarType(objUser.MemberOf) = 8 then                       ' object string - just one line
    strGroupPath = "LDAP://" & objUser.MemberOf
    Set objGroup = GetObject(strGroupPath)
    UserGroups=UserGroups & "[" & "CN=" & objgroup.cn & "]"
Else
    For Each strGroup in objUser.MemberOf
      strGroupPath = "LDAP://" & strGroup
      Set objGroup = GetObject(strGroupPath)
      UserGroups=UserGroups & "[" & "CN=" & objgroup.cn & "]"
    Next
End if
'                                                            @@ERR 07/22/05 end change
'The following line will echo the user groups to the screen.
'wscript.echo usergroups
end if

if REa = 3 then
     Set userobj = GetObject(adsipath)
end if

'Init Groups
Dim GroupObj
if REa = 3 then
     For Each GroupObj In UserObj.Groups
          UserGroups=UserGroups & "[" & "CN=" & GroupObj.Name & "]"
     Next 
     'Next line will display user group membership to screen for NT Clients
     'WSCRIPT.ECHO UserGroups
end if


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function:      MapDrive
' Arguments:     Drive Letter (with ":") and Share Name on ShareSvr
' Description:   Maps Share' mapping CANNOT exist already
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub MapDrive( strDrive, strShare )
	Dim Path
	If InStr(strShare,"\") then     
		Path="\\" & strShare
'	Else    
'		Path="\\" & RLICISShareSvr & "\" & strShare
	End If
'	Map drive
	On Error Resume Next
	WSHNetwork.MapNetworkDrive strDrive, Path
	On Error GoTo 0
	Err.Clear   ' Clear the error.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function:      MapDriveDelete
' Arguments:     Drive Letter (with ":") 
' Description:   Deletes the Drive Specified
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub MapDriveDelete( strDrive )
	On Error Resume Next
	WSHNetwork.RemoveNetworkDrive StrDrive, True
	On Error GoTo 0
	Err.Clear   ' Clear the error.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function: Return value of specified environment string
'
' Arguments: Environment variable to find
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetEnv(strName)
	Dim strEnv

	GetEnv = ""
	strName=UCase(strName)
	For Each strEnv In WshProcessEnvironment
		If strName = UCase(Left(strEnv,Len(strName))) Then GetEnv=Right(strEnv,Len(strEnv)-Len(strName)-1)
	Next
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function: Check if user is in Group
'
' Arguments: Group to check membership of
' 
' Note: Requires pre-initialization of UserGroups variant
' Note: The InGroup conditional test is CASE SENSITIVE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InGroup(strGroup)
	InGroup=False
	If InStr(UserGroups,"[" & strGroup & "]") Then
			InGroup=True
	End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Function:      InstallADSI
'' Description:   Checks for and installs ADSI
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function InstallADSI
If OS = "Win32" Then
	If Not WshFS.FileExists(WshSys32 & "adsnt.dll") Then
		WshShell.Run WshInstall & "adsiw95 /q:a /r:n ", 0, True 
	End If
Else
	If Not WshFS.FileExists(WshSys32 & "adsnt.dll") Then
		WshShell.Run WshInstall & "adsi86 /q:a /r:n ", 0, True 
	End If
End IF
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function:      Register
' Description:   Checks for and installs WSH controls
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function Register(WshFile, lSize, sLicKey, sLic)
Dim wshFDC
Dim sSaveL

If wshFS.FileExists(WshSys32 & WshFile) Then
	Set wshF = WshFS.GetFile(WshSys32 & WshFile)
	If wshF.size = lSize Then
		If wshFS.FileExists(WshInstall & WshFile) Then
			Set wshFDC = wshFS.GetFile(WshInstall & WshFile)
			' check the last modified time/date stamp
			If CStr(wshF.DateLastModified) = CStr(wshFDC.DateLastModified) Then
				Exit Function 
				End If
				Else
				' benefit of the doubt the file may be being upgraded
				Exit Function
		End If
	End If
End If

' problem cannot overwrite an existing file
If wshFS.FileExists(WshInstall & WshFile) Then
	WshFS.CopyFile WshInstall & WshFile, WshSys32
	WshShell.Run WshSys32 & "regsvr32 /s " & WshSys32 & WshFile, 0, True 
	sSaveL = ""
	If sLicKey <> "" And sLic <> "" Then
		sSaveL = ", " & AddLicense(sLicKey, sLic)
		End If
End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function:      AddLicense
' Description:   Checks for and installs licensed WSH controls
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function AddLicense(sLicKey, sLic)
Dim wshReg
Dim wshKey
Dim wshHive

Set wshReg = WScript.CreateObject("RegObj.Registry.1")
Set wshHive = wshReg.RegKeyFromString("\HKEY_CLASSES_ROOT\Licenses")
For Each wshKey in wshHive.Subkeys
If wshKey.name = CStr(sLicKey) Then
' AddLicense = "Control already Licensed" 
Exit Function
End If
Next
wshHive.Subkeys.Add CStr(sLicKey)
Set wshHive = wshReg.RegKeyFromString("\HKEY_CLASSES_ROOT\Licenses\" & CStr(sLicKey))
wshHive.Values.Add "", sLic, 1 ' string Default value
AddLicense = "Licensed Control"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'The following section checks for the existence of %windir%\reg.exe
'If it does not exist, it puts it there.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If wshFS.FileExists(envwindir + "\reg.exe") Then
'
else
     WshFS.CopyFile wshlogonpath + "\reg.exe", envwindir + "\"
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Install and register Components
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Register "msvbvm50.dll", 1355776, "", ""
Register "RICHTX32.ocx", 212240, "", ""
Rem Register "Olepro32.DLL", 164112, "", "" 
Register "COMCTL32.ocx", 608448, "", ""
InstallADSI

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Find LogonServer for Win9x Machines
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If OS = "Win32" Then
	Dim ColDrives
	Dim W9xLogonServer
	Set ColDrives = WSHNetwork.EnumNetworkDrives
	For i = 0 To ColDrives.Count -1
		W9xLogonServer = ColDrives(i)
	Next
	W9xLogonServer = UCASE(W9xLogonServer)
	W9xLogonServer = RIGHT(W9xLogonServer,(LEN(W9xLogonServer)-2))
End If
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'*****************************************************************************************
'

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
'USE THE FOLLOWING SECTIONS TO MODIFY THE SCRIPT'S OPERATION
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Workstation Time
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'WshShell.Run "%COMSPEC% /c net time \\reshealthcare.org /set /y",0,true


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'The following set of code will disable the proxy server setting.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
IF OS = "Windows_NT" THEN
     WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\tqcrunas.exe /f %logonserver%\netlogon\DisableProxy", 0, True 
     WshShell.RegWrite "HKCU\Software\microsoft\windows\currentversion\internet settings\proxyenable","0","REG_DWORD" 
If Left(OSreg,13) <> "Windows Vista" Then WshShell.RegWrite "HKLM\Software\microsoft\windows\currentversion\internet settings\proxyenable","0","REG_DWORD" 
END IF

'websense test
'If InGroup("CN=Internet WWW Persist") Then
'   WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\LogonApp.exe http://10.250.9.180:15880 /COPY ",0,true
'   WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\LogonApp.exe http://10.40.8.180:15880 /COPY ",0,true
'End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Drive Map Standards
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'You must delete mapped drives before you can map them to something else.  Use MapDriveDelete function.
'NOTE: The InGroup conditional test is CASE SENSITIVE
'Sample Formats:
'	MapDrive "g:", RLICISShareSvr & "\Shared"		'Use this where you want to define a server by variable
'	MapDrive "p:", "RLICCONTRMGR\CMAPP"			'Use this to map directly to a server and share

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Configure the User Home Drives by Group
'RLICHS01\Home
'SJHHS01\Home
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If InGroup("CN=RLICHS01_Home") Then
	MapDriveDelete "h:"
	MapDrive "h:", "rlichs01\home$\%username%"
End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Configure the Department Shared Drives by Group
'User Home Drives are Mapped to H: in RHCVB.Bat
'NOTE: The InGroup conditional test is CASE SENSITIVE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If InGroup("CN=PWHomeCare") Then
	MapDriveDelete "g:"
	MapDriveDelete "p:"
	MapDrive "g:", "RLICFS01\SHARED"
	MapDrive "p:", "rhcapps207\apps"
End If


If InGroup("CN=Print-Graphics") Then
	MapDriveDelete "m:"
	MapDrive "m:", "RHCAPPS72\GRAPHICS"
End If

If InGroup("CN=OLRFS01 Shared") Then
	MapDriveDelete "g:"
	MapDrive "g:", "OLRFS01\SHARED"
End If

If InGroup("CN=RLICFSWP005 Shared") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RLICFSWP005\SHARED"
End If

If InGroup("CN=RHCMASTER Shared") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RHCMASTER\SHARED"
End If

If InGroup("CN=RLICFS01 Shared") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RLICFS01\SHARED"
End If

If InGroup("CN=RMCFS01 Shared") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=SFHFS01 Shared") Then
	MapDriveDelete "g:"
	MapDrive "g:", "SFHFS01\SHARED"
End If

If InGroup("CN=WHFS01 Shared") Then
	MapDriveDelete "g:"
	MapDrive "g:", "WHFS01\SHARED"
End If

If InGroup("CN=RHCFS01 Shared") Then
        MapDriveDelete "g:"
        MapDrive "g:", "RHCFS01\Shared"
End If

If InGroup("CN=WSMCFS01 Shared") Then
        MapDriveDelete "g:"
        MapDrive "g:", "WSMCFS01\Shared\WSMC"
End If

If InGroup("CN=RMC_ER_DOCS") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RLICFSWP002\ER"
End If

If InGroup("CN=RMC SPD") Then
	MapDriveDelete "s:"
	MapDrive "s:", "RHCFS01\SHARED"
End If

If InGroup("CN=SFH Anesthesiology") Then
        MapDriveDelete "g:"
        MapDrive "g:", "SFHFS01\SHARED"
End If

If InGroup("CN=RHC Development") Then
        MapDriveDelete "g:"
        MapDrive "g:", "RLICFS01\SHARED"
End If

If InGroup("CN=RMC Finance TrndStar") Then
	MapDriveDelete "f:"
	MapDrive "f:", "RMCFS01\SHARED"
End If

If InGroup("CN=RMC Finance WSMC") Then
        MapDriveDelete "f:"
        MapDrive "f:", "RHCFS01\SHARED"
End If

If InGroup("CN=RHC_Premier") Then
	MapDriveDelete "f:"
	MapDrive "f:", "RHCFS01\SHARED\RMC\Premier"
End If  

If InGroup("CN=OLR Nursing Admin") Then
	MapDriveDelete "g:"
	MapDrive "g:", "OLRFS01\SHARED"
End If

If InGroup("CN=RLI") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End IF

If InGroup("CN=OLR Medical Staff") Then
	MapDriveDelete "f:"
	MapDrive "f:", "OLRAPPS1\APPS"
End IF

If InGroup("CN=RMC ClinicalResearch") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\Shared"
End IF

If InGroup("CN=OLR RESQ") Then
	MapDriveDelete "R:"
	MapDrive "R:", "OLRAPPS1\RESQ"
End IF

If InGroup("CN=RMC PA") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\Shared"
End IF

If InGroup("CN=RMC PM") Then
	'Physical Medicine
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\Shared"
End IF

If InGroup("CN=RMC Food & Nutrition") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=RMC Nursing Admin") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=RMC Facility Service") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=OLR Facility Service") Then
	MapDriveDelete "g:"
	MapDrive "g:", "OLRFS01\SHARED"
End If
If InGroup("CN=HomeHealth_Leaders") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RLICFS01\SHARED"
End If

If InGroup("CN=RMC Heart & Lung") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=RMC Biomed Engineer") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=RMC Cancer Center") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=RMC Radiology") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=RHC_RACS_Daily_Operations") Then
        MapDriveDelete "s:"
        MapDrive "s:", "RHCAPPS111\SHARED"
End If

If InGroup("CN=App_PracPlus_BOReports_WSMC") Then
        MapDriveDelete "s:"
        MapDrive "s:", "RHCAPPS111\SHARED"
End If

If InGroup("CN=RMC Administration") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=RMC Cardio Resp") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=RMC Medical Records") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=RMC Family Practice") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=RHC Fitness Center") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=RMC Dialysis") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=RMC Security") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=RMC Medical Staffing") Then 
	MapDriveDelete "g:"	
	MapDrive "g:", "RMCFS01\Shared\Medical_ Staff" 
End If

If InGroup("CN=RMC Communications") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=SEH PCCH") Then
        MapDriveDelete "g:"
        MapDrive "g:", "RHCAPPS27\SHARED"
End If

If InGroup("CN=SEH PCCH Users") Then
        MapDriveDelete "g:"
        MapDrive "g:", "RHCAPPS27\SHARED"
End If

If InGroup("CN=RMC Admitting & Reg") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=RMC PIR") Then
	MapDriveDelete "g:"
	MapDrive "g:", "RMCFS01\SHARED"
End If

If InGroup("CN=RMC Surgery") Then
	MapDriveDelete "J:"
	MapDrive "J:", "RMCFS01\SHARED"
End If

If InGroup("CN=APP_METRIQ_RMC") Then
'	RMC Medical Records Department - Metriq Tumor Registry System
	MapDriveDelete "P:"
	MapDrive "P:", "RMCAPPS2\APPS"
End If

If InGroup("CN=APP_CDMP_RMC") Then
	MapDriveDelete "P:"
	MapDrive "P:", "RMCAPPS2\APPS"
End If

If InGroup("CN=App_ScottCare_RMC") Then
	MapDriveDelete "E:"
	MapDrive "F:", "RMCAPPS2\APPS\ScottCare"
End If
If InGroup("CN=APP_PROSERIES_RMC") Then
	MapDriveDelete "P:"
	MapDrive "P:", "RMCFS01\APPS"
End If

If InGroup("CN=APP_METRIQ_SME") Then
'	SME Medical Records Department - Metriq Tumor Registry System
	MapDriveDelete "P:"
	MapDrive "P:", "SMEAPPS5\APPS"
End If

If InGroup("CN=APP_METRIQ_SJH") Then
'	SJH Medical Records Department - Metriq Tumor Registry System
	MapDriveDelete "P:"
	MapDrive "P:", "SJHAPPS5\APPS"
End If

If InGroup("CN=WSMC FAC_OPS") Then
	MapDriveDelete "G:"
	MapDrive "G:", "wsmcfs01\share"
End If

If InGroup("CN=WSMC 1000LAKE") Then
	MapDriveDelete "G:"
	MapDrive "G:", "wsmcfs01\share"
End If

If InGroup("CN=WSMC Audiology") Then
	MapDriveDelete "G:"
	MapDrive "G:", "rhcfs01\Shared"
End If

If InGroup("CN=WSMC_LAB") Then
	MapDriveDelete "G:"
	MapDrive "G:", "rhcfs01\Shared"
End If

If InGroup("CN=WSMC Impress Users") Then
	MapDriveDelete "I:"
	MapDrive "I:", "wsmcapps1\apps\impress"
End If

If InGroup("CN=App_Superscript_WSMCPlaza_Users") Then
	MapDriveDelete "P:"
	MapDrive "P:", "wsmcapps1\apps"
End If

If InGroup("CN=App_Superscript_WSMCCenter_Users") Then
	MapDriveDelete "V:"
	MapDrive "V:", "wsmcapps1\apps"
End If

If InGroup("CN=App_Superscript_SME_Users") Then
	MapDriveDelete "P:"
	MapDrive "P:", "wsmcapps1\apps"
End If

If InGroup("CN=App_Superscript_SEH_Users") Then
	MapDriveDelete "Q:"
	MapDrive "Q:", "wsmcapps1\apps"
End If

If InGroup("CN=WSMC Central Sterile") Then
	MapDriveDelete "g:"
	MapDrive "g:", "wsmcfs01\Share"
End If

If InGroup("CN=WSMC GI LAB") Then
	MapDriveDelete "g:"
	MapDrive "g:", "wsmcfs01\Share"
End If

If InGroup("CN=APP_METRIQ_WSMC") Then
'	WSMC Medical Records Department - Metriq Tumor Registry System
	MapDriveDelete "p:"
	MapDrive "p:", "wsmcapps1\apps"
End If

If InGroup("CN=WSMC PFS") Then
	MapDriveDelete "g:"
	MapDrive "g:", "wsmcfs01\Share"
End If



If InGroup("CN=WSMC REG_SCAN SHARED") Then
	MapDriveDelete "M:"
	MapDrive "M:", "reg_scan\shared"
End If

If InGroup("CN=WSMC SuperScript Users") Then
	MapDriveDelete "v:"
	MapDrive "v:", "wsmcapps1\apps"
End If

If InGroup("CN=WSMC AGENCY USERS") Then
	MapDriveDelete "k:"
	MapDrive "k:", "wsmcapps1\apps"
End If

If InGroup("CN=WSMC HR SHARED") Then
	MapDriveDelete "S:"
	MapDrive "S:", "HR\SHARED"
End If

If InGroup("CN=WSMC HOUSE MANAGERS") Then
	MapDriveDelete "P:"
	MapDrive "P:", "WSMCFS01\share\WSMC\House Managers\"
End If

If InGroup("CN=WSMC STIX") Then
	MapDriveDelete "s:"
	MapDrive "s:", "WSMCFS01\share\"
End If

If InGroup("CN=WSMC KREG USERS") Then
	MapDriveDelete "k:"
	MapDrive "k:", "wsmcapps1\apps\kreg"
End If

If InGroup("CN=WSMC APOLLO USERS") Then
	MapDriveDelete "L:"
	MapDrive "L:", "WSMCAPPS3\APPS\APOLLO"
End If

If InGroup("CN=IP_OP_THERAPY") Then
	MapDriveDelete "O:"
	MapDrive "O:", "wsmcfs01\share\WSMC\IP OP Therapy"
End If


If InGroup("CN=RHC RESSCHED") Then
	MapDriveDelete "I:"
	MapDrive "I:", "rhcweb01\ressched"
End If

If InGroup("CN=RHC RESSCHED") Then
	MapDriveDelete "v:"
	MapDrive "v:", "rhcfs01\shared\rmc\residency tracking"
End If

If InGroup("CN=SS MISSION") Then
	MapDriveDelete "S:"
	MapDrive "S:", "wsmcfs01\shared\WSMC\Spiritual Services and Mission"
End If

If InGroup("CN=RMC IRB") Then
	MapDriveDelete "I:"
	MapDrive "I:", "wsmcfs01\shared\WSMC\RMC IRB"
End If

If InGroup("CN=Residency EMR") Then
	MapDriveDelete "J:"
	MapDrive "J:", "wsmcfs01\shared\WSMC\Residency EMR Assessment"
End If

If InGroup("CN=RHC Family Medical") Then
	MapDriveDelete "v:"
	MapDrive "v:", "wsmcfs01\Shared\WSMC\Family Medical"
End If

If InGroup("CN=SMN Rehab") Then
	MapDriveDelete "R:"
	MapDrive "R:", "wsmcfs01\Shared\WSMC\rehab"
End If

If InGroup("CN=SJH Facilities Management") Then
	MapDriveDelete "J:"
	MapDrive "J:", "wsmcfs01\Shared\WSMC\SJH Facilities Management" 
End If

If InGroup("CN=WSMC Volunteer Works Users") Then
	MapDriveDelete "v:"
        MapDrive "v:", "wsmcapps1\apps\Volunteer Works"
End If

If InGroup("CN=wsmc payroll conversion") Then
	MapDriveDelete "q:"
	MapDrive "q:", "rhcraftp101\ftproot$\v_iatric"
End If

If InGroup("CN=TSACadmin") Then
	MapDriveDelete "R:"
	MapDrive "R:", "SFHfs01\Shared\TSAContracts"
End If

If InGroup("CN=TSAContracts") Then
	MapDriveDelete "R:"
	MapDrive "R:", "SFHfs01\Shared\TSAContracts"
End If


If InGroup("CN=ces_ftp_drive") Then
	MapDriveDelete "j:"
	MapDrive "j:", "rhcraftp101\ftproot$\v_ces"
End If


' Steve H. This group mapping is temporary until all user's are migrated and shared data
' is moved to RMCFS01. This mapping will need to be below the \\RMCFS01\Shared mapping
If InGroup("CN=RMC EmerMedResProg") Then
	MapDriveDelete "g:"
	MapDriveDelete "t:"
	MapDrive "g:", "RLICFS01\SHARED"
'	MapDrive "t:", "RMCEMRP\HOME"
	MapDrive "J:", "RMCFS01\SHARED"
End If

If InGroup("CN=SFH PHARMACY") Then
	MapDriveDelete "g:"
	MapDrive "g:", "sfhfs01\SHARED"
End If

If InGroup("CN=IS-Support Center") Then
	MapDriveDelete "M:"
	MapDrive "M:", "RLICFS01\SHARED"
End If

If InGroup("CN=SMN Safety Council") Then
	MapDriveDelete "M:"
	MapDrive "M:", "RLICFS01\SHARED"
End If

If InGroup("CN=Nursing Exe Committ") Then
        MapDriveDelete "N:"
        MapDrive "N:", "RLICFS01\SHARED"
End If

If InGroup("CN=RHC_Pharmacy_Automation") Then
        MapDriveDelete "r:"
        MapDrive "r:", "RHCFS01\Shared\RLIC"
End If

If InGroup("CN=SME SuperS Users_NFC") Then
	MapDriveDelete "v:"
	MapDrive "v:", "rmcfs01\shared\sme\NFC"
End If

If InGroup("CN=SME SuperS Users_POB") Then
	MapDriveDelete "v:"
	MapDrive "v:", "rmcfs01\shared\sme\POB"
End If


If InGroup("CN=APP_WSMC_IMPAC_USERS") Then
	MapDriveDelete "I:"
	MapDrive "I:", "impac\impac"
End If


If InGroup("CN=DMI") Then
	MapDriveDelete "M:"
	MapDrive "M:", "wsmcfs01\Share\WSMC\DMI"
End If


If InGroup("CN=Rehab_Lab_Division") Then
	MapDriveDelete "R:"
	MapDrive "R:", "wsmcfs01\Share\WSMC\Rehab Lab Division"
End If


If InGroup("CN=app_ics") Then
	MapDriveDelete "z:"
	MapDrive "z:", "rhcraftp101\ftproot$\V_ICS"
End If



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Configure the Apps by Group
'NOTE: The InGroup conditional test is CASE SENSITIVE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If InGroup("CN=RMC-Angoss") Then
	MapDriveDelete "f:"
	MapDrive "f:", "RMCAPPS1\ANGOSS" 
End If

If InGroup("CN=IS-Operations") Then
'	MapDriveDelete "i:"
	MapDrive "i:", "\\rhcraftp101\ftproot$"
End If

If InGroup("CN=RA_FTP_Admins") Then
'	MapDriveDelete "i:"
	MapDrive "i:", "\\rhcraftp101\ftproot$"
End If

If InGroup("CN=App_Access_It_Universal") Then
	MapDriveDelete "L:"
	MapDrive "L:", "RHCAPPS62A\IMAGES" 
End If

If InGroup("CN=App_SYSTOC") Then
	MapDriveDelete "P:"
	MapDrive "P:", "rhcapps209\systoc72" 
End If

If InGroup("CN=APP_VCMD_HFNRC") Then
	MapDriveDelete "I:"
	MapDrive "I:", "hfnrcapps3\HFNRC" 
End If

If InGroup("CN=RHC Occ Health DB") Then
	MapDriveDelete "o:"
	MapDrive "o:", "RLICWEB01\OHD" 
End If

If InGroup("CN=App_MAS90_SKO") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RHCAPPS2\Apps"
End If

If InGroup("CN=ESachs") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RHCAPPS1\APPS" 
End If

If InGroup("CN=APP_HIIQ") Then
	MapDriveDelete "p:"
	MapDrive "p:", "sfhfs01\APPS" 
End If

If InGroup("CN=Karch") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RHCAPPS1\APPS" 
End If

If InGroup("CN=App_PMM_EDI") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RHCAPPS14\editemp" 
End If

If InGroup("CN=APP_AED") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RHCAPPS5\APPS\AED" 
End If

If InGroup("CN=App_PMM_Admins") Then
	MapDriveDelete "q:"
	MapDrive "q:", "RHCAPPS40\edi850" 
End If

If InGroup("CN=APP_MUSE_STRESS") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RHCEKG1\VOL000" 
End If

If InGroup("CN=App_Noah_RMC") Then
	MapDriveDelete "r:"
	MapDrive "r:", "RMCAPPS6\NOAH35\NOAH System\Database" 
End If

If InGroup("CN=OLR Aims") Then
	MapDriveDelete "p:"
	MapDrive "p:", "OLRFS01\SHARED" 
End If

If InGroup("CN=OLR_SoftMed") Then
	MapDriveDelete "e:"
	MapDrive "e:", "OLRAPPS1\APPS" 
End If

If InGroup("CN=APC-3M Encoder") Then
	'Some NEBO users also need APC-3M Encoder'
'        MapDriveDelete "q:"
'        MapDrive "q:", "RHCAPPS1A\APPS"
        MapDrive "q:", "RHCAPPS1V\APPS"

End If

'If InGroup("CN=BodyWorks") Then
'	'Some NEBO users also need Body Works
'        MapDriveDelete "q:"
'        MapDrive "q:", "RHCAPPS1A\APPS"
'End If

'If InGroup("CN=QS1") Then
'	MapDriveDelete "p:"	
'	MapDrive "p:", "rhcapps201\QS1" 
'End If

If InGroup("CN=App_Xtend_SMN") Then
        MapDriveDelete "n:"
        MapDrive "n:", "SMNAPPS4\Xtend"
End If

If InGroup("CN=App_MediCall_Users") Then
        MapDriveDelete "n:"
        MapDrive "n:", "RHCAPPS243\Xtend"
        MapDriveDelete "w:"
        MapDrive "w:", "RMCAPPS10\Xtend"
        MapDriveDelete "x:"
        MapDrive "x:", "SMEAPPS10\Xtend"   
        If InGroup("CN=App_Xtend_SMN") Then
             MapDriveDelete "y:"
             MapDrive "y:", "SMNAPPS4\Xtend"
        End If
End If

If InGroup("CN=App_MediCall_Admins") Then
        MapDriveDelete "u:"
        MapDrive "u:", "RHCWEB04\Xtend"
        MapDriveDelete "v:"
        MapDrive "v:", "RHCAPPS245\Xtend"
End If

If InGroup("CN=App_Eclipse") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RHCAPPS10\ECLIPSE"
End If

If InGroup("CN=App_FIM_OLR") Then
        MapDriveDelete "p:"
        MapDrive "p:", "OLRFS01\FIM"
End If

If InGroup("CN=App_eCareCMS_Users") Then
'The following section checks for the existence of the EcareCMS.bat file in the root of
'the system drive.  If it exists, it executes.
	If wshfs.fileexists(envsystem + "\ECARECMS.CMD") then
		WshShell.Run "%COMSPEC% /c c:\EcareCMS.CMD ", 0, True
   	End if
End If

If InGroup("CN=App_eCareCMS_Users") Then
	MapDriveDelete "q:"
	MapDrive "q:", "rhcapps254\ecarecms"
End If

If InGroup("CN=App_Ecare_RMC") Then
	MapDriveDelete "p:"	
	MapDrive "p:", "RHCAPPS11\APPS" 
End If

If InGroup("CN=ERAS") Then
	MapDriveDelete "X:"	
	MapDrive "X:", "Rhcapps249\ERASSERV.2K6" 
	MapDrive "W:", "Rhcapps249\APPS"
End If

If InGroup("CN=App_BW_SMNH") Then
	MapDriveDelete "q:"	
	MapDrive "q:", "RMCAPPS1\APPS" 
End If

If InGroup("CN=App_ROCS_SME") Then
	MapDriveDelete "p:"	
	MapDrive "p:", "SMEAPPS5\APPS" 
End If

If InGroup("CN=App_Mrsys_SFH") Then
	MapDriveDelete "p:"	
	MapDrive "p:", "SFHAPPS5\APPS" 
End If

If InGroup("CN=RMC Medical Library") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RMCAPPS2\APPS"
End If

If InGroup("CN=RMC Medical Records") Then
	MapDriveDelete "p:"	
	MapDrive "p:", "RMCAPPS6\APPS" 
End If

If InGroup("CN=App_ChartOne_RMC") Then
	MapDriveDelete "p:"	
	MapDrive "p:", "RMCAPPS2\APPS" 
End If


If InGroup("CN=RMC PIR MedQuest") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RMCAPPS6\APPS"
End If

If InGroup("CN=RMC Registrar Users") Then 
	MapDriveDelete "r:"	
	MapDrive "r:", "RMCAPPS1\RegNet" 
End If

If InGroup("CN=RMC Sentinel") Then
	MapDriveDelete "p:"
	MapDrive "p:", "Rmcapps6\APPS"
End If

If InGroup("CN=App_QMatic_WH") Then
	MapDriveDelete "Q:"
	MapDrive "Q:", "WHAPPS1\APPS"
End If

If InGroup("CN=App_Ecare_SFH") Then
	MapDriveDelete "P:"
	MapDrive "P:", "RHCAPPS11\APPS"
End If

If InGroup("CN=App_Ecare_Admins") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RHCAPPS11\APPS"
End If

If InGroup("CN=App_Ecare_SJH") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RHCAPPS11\APPS"
End If

If InGroup("CN=App_Ecare_SEH") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RHCAPPS11\APPS"
End If

If InGroup("CN=App_Ecare_SMN") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RHCAPPS11\APPS"
End If

If InGroup("CN=App_Ecare_HFMC") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RHCAPPS11\APPS"
End If

If InGroup("CN=App_ECARE_HFNRC") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RHCAPPS11\APPS"
End If

If InGroup("CN=App_Ecare_WH") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RHCAPPS11\APPS"
End If

If InGroup("CN=App_Ecare_WSMC") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RHCAPPS11\APPS"
End If

If InGroup("CN=APP_MWP_OLR") Then
	MapDriveDelete "g:"
	MapDrive "g:", "OLRTRANS\SHARED"
	MapDriveDelete "i:"
	MapDrive "i:", "OLRTRANS\CQTRANS$"
End If

If InGroup("CN=APP_MWP_RMC") Then
	MapDriveDelete "g:"
	MapDriveDelete "i:"
	MapDrive "g:", "RMCFS01\Shared"
End If

If InGroup("CN=App_Resatrac_SFH") Then
	MapDriveDelete "p:"
	MapDrive "p:", "SFHAPPS5\APPS"
End If

If InGroup("CN=App_Kronos_MedRec_SM") Then
        MapDriveDelete "q:"
        MapDrive "q:", "SMNAPPS1\APPS"
End If

If InGroup("CN=App_EDIM_User") Then
        MapDriveDelete "U:"
        MapDrive "U:", "SFHAPPS1\EDIMDATA"
End If

If InGroup("CN=App_EDIM_ADMIN") Then
        MapDriveDelete "U:"
        MapDrive "U:", "SFHAPPS1\EDIMDATA"
End If

If InGroup("CN=APP_KRONOS_MEDREC_S") Then
        MapDriveDelete "T:"
        MapDrive "T:", "SMNAPPS1\APPS"
End If

If InGroup("CN=App_Kronos_SMN") Then
'	MapDriveDelete "p:"
	MapDrive "p:", "SMNAPPS1\APPS"
End If

If InGroup("CN=App_Kronos_SFH") Then
'	MapDriveDelete "t:"
	MapDrive "t:", "sfhapps3\apps"
End If

If InGroup("CN=App_Kronos_WH") Then
'	MapDriveDelete "u:"
	MapDrive "u:", "whapps1\apps"
End If

If InGroup("CN=APP_TIMECARDS_HFMC") Then
'	MapDriveDelete "f:"
	MapDrive "f:", "srva003c\Applications"
End If

If InGroup("CN=App_MSO_Users") Then
'	MapDriveDelete "L:"
	MapDrive "l:", "rhcapps234\apps"
End If

If InGroup("CN=App_MSO_Test") Then
'	MapDriveDelete "N:"
	MapDrive "n:", "rhcapps234test\apps"
End If

If InGroup("CN=APP_METRIQ_SFH") Then
'	SFH Medical Records Department - Metriq Tumor Registry System
	MapDriveDelete "p:"
	MapDrive "p:", "SFHAPPS5\APPS"
End If

If InGroup("CN=RMCEMRP Residents") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RMCAPPS6\APPS"
End If

If InGroup("CN=RMCEMRP Admin") Then
	MapDriveDelete "p:"
	MapDrive "p:", "RMCAPPS6\APPS"
End If

If InGroup("CN=WH Sentinel") Then
	MapDriveDelete "p:"
	MapDrive "p:", "Rmcapps6\APPS"
End If

If InGroup("CN=WH MIDAS") Then
	MapDriveDelete "p:"
	MapDrive "p:", "WHAPPS2\Midshare"
End If

If InGroup("CN=App_GKC20_WH_Admins") Then
	MapDriveDelete "U:"
	MapDrive "U:", "whapps1\apps"
End If

If InGroup("CN=App_GKC20_WH_Users") Then
	MapDriveDelete "U:"
	MapDrive "U:", "whapps1\apps"
End If

If InGroup("CN=App_UDS-PRO_RMC") Then
	MapDriveDelete "E:"
	MapDrive "E:", "RMCAPPS1\APPS"
End If

If InGroup("CN=App_UDSPRO_SJH") Then
	MapDriveDelete "E:"
	MapDrive "E:", "SJHAPPS1\APPS"
End If

If InGroup("CN=App_Security_SJH") Then
	MapDriveDelete "S:"
	MapDrive "S:", "SJHAPPS1\APPS\Security"
End If

If InGroup("CN=App_Anacomp_SFH") Then
	MapDriveDelete "p:"
	MapDrive "p:", "SFHAPPS5\APPS"
End If

If InGroup("CN=App_Anacomp_RHC") Then
	MapDriveDelete "O:"
	MapDrive "O:", "RHCLABSRV04\ANACOMP"
End If

If InGroup("CN=App_ReWin_RMC") Then
        MapDriveDelete "q:"
        MapDrive "q:", "RHCAPPS9A\APPS"
End If

If InGroup("CN=App_Orsos_HFMC_Users") Then
	MapDriveDelete "p:"
	MapDrive "p:", "HFMCAPPS1\PERSE"
End If

If InGroup("CN=App_PHARMACY_RMC") Then
        MapDriveDelete "q:"
        MapDrive "q:", "SMNAPPS3\APPS"
End If

If InGroup("CN=SJH APP_EMS") Then
        MapDriveDelete "I:"
        MapDrive "I:", "SJHAPPS68\APPS"
End If

If InGroup("CN=SJH HeartBase") Then
        MapDriveDelete "S:"
        MapDrive "S:", "SJHAPPS68\APPS\HeartBase"
End If

If InGroup("CN=SJH RHCZETRON") Then
        MapDriveDelete "S:"
        MapDrive "S:", "SJHAPPS1\APPS\ZETRON"
End If

If InGroup("CN=SJH SOFTTEST") Then
        MapDriveDelete "K:"
        MapDrive "K:", "SJHAPPS1\APPS"
End If

If InGroup("CN=App_ROCS_RBH") Then
        MapDriveDelete "F:"
        MapDrive "F:", "rhcfs01\shared\RBH\Apps\DMH"
End If

If InGroup("CN=App_DASA_RBH") Then
        MapDriveDelete "I:"
        MapDrive "I:", "rhcfs01\shared\RBH\Apps"
End If

If InGroup("CN=App_CARE_T&B_RBH") Then
        MapDriveDelete "J:"
        MapDrive "J:", "rhcfs01\shared\RBH\Apps\CU\CARE"
        MapDriveDelete "M:"
        MapDrive "M:", "rhcfs01\shared\RBH\Apps\CU\RBHN"
        MapDriveDelete "N:"
        MapDrive "N:", "rhcfs01\shared\RBH\Apps\CU\TBDATA"
End If

If InGroup("CN=App_CARE_GL_RBH") Then
        MapDriveDelete "K:"
        MapDrive "K:", "rhcfs01\shared\RBH\Apps\CU\GLDATA"
End If

If InGroup("CN=App_CARE_AP_RBH") Then
        MapDriveDelete "L:"
        MapDrive "L:", "rhcfs01\shared\RBH\Apps\CU\APDATA"
End If

If InGroup("CN=App_IMPAC_SJH") Then
	MapDriveDelete "P:"
	MapDrive "P:", "SJHAPPS3\APPS" 
End If

If InGroup("CN=SJH Chief Residents") Then
	MapDriveDelete "T:"
	MapDrive "T:", "WHFS01\shared\sjh\Chief_Residents"
End If

If InGroup("CN=APP_CECC_REP") Then
	MapDriveDelete "L:"
	MapDrive "L:", "RHCAPPS50\CECC_REPORTS"
End If

If InGroup("CN=APP_Leaders_Guide") Then
	MapDriveDelete "S:"
	MapDrive "S:", "RHCFS01\Services$"
End If

If InGroup("CN=APP_LAUNDRY_RHC") Then
        MapDriveDelete "X:"
        MapDrive "X:", "RHCFS01\shared\rhc\laundry\CTXDATA"
End If

If InGroup("CN=RHC_Nurs_Orientation") Then
        MapDriveDelete "X:"
        MapDrive "X:", "RHCFS01\shared\RHC"
End If

If InGroup("CN=App_SURGICALPAK_SFH") Then
        MapDriveDelete "X:"
        MapDrive "X:", "SFHFS01\CTXDATA"
End If

'If InGroup("CN=App_MaintMizer_SME") Then
'        MapDriveDelete "v:"
'        MapDrive "v:", "smnapps3\apps\mt2000"
'End If

If InGroup("CN=App_MedComp_RBH") Then
        MapDriveDelete "I:"
        MapDrive "I:", "RHCAPPS65\medcomp\Caseware2020\Server"
End If

If InGroup("CN=App_ERC_Financial_RBH") Then
	MapDriveDelete "J:"
        MapDrive "J:", "RHCAPPS65\ERC\Financial_Database"
End If 

If PCName = "HFHXP1001B" Then
	MapDriveDelete "y:"
	MapDrive "y:", "HFMCAPPS3\paigext"
End If

If PCName = "HFHXPBADGEPC" Then
	MapDriveDelete "y:"
	MapDrive "y:", "RHCFS01\SHARED\HFH\BADGE"
End If          

If InGroup("CN=App_Sitemaster_Users") Then
        MapDriveDelete "q:"
        MapDrive "q:", "SMNAPPS3\APPS\Sitemaster\LiveData"
End If


If InGroup("CN=RHC Nursing Ed Council") Then
        MapDriveDelete "y:"
        MapDrive "y:", "rhcfs01\shared\rhc\RHC_Nursing_Education_Council"
End If




If InGroup("CN=App_CES_Reports") Then
	MapDriveDelete "R:"
	MapDrive "R:", "rhcapps144\ces$\reports"
End If

If InGroup("CN=FTPDRIVE_USERS") Then
	MapDriveDelete "I:"
	MapDrive "I:", "rhcraftp101\ftproot$"
End If


If InGroup("CN=app_rmcorschedule") Then
	MapDriveDelete "T:"
	MapDrive "T:", "rhcweb01\rmcorschedule$" 
End If


If InGroup("CN=wsmc_hr") Then
	MapDriveDelete "t:"
	MapDrive "t:", "rmcfs01\SHARED\wsmc_hr"
End If




If InGroup("CN=APP_PSC_REPORTS") Then
	MapDriveDelete "P:"
	MapDrive "P:", "rhcraftp101\ftproot$\V_PSC\Reports_From_PSC"
End If


If InGroup("CN=APP_SYSTOC_Collections") Then
	MapDriveDelete "v:"
	MapDrive "v:", "rhcraftp101\ftproot$\V_ICS\systoc"
End If




If InGroup("CN=APP_RHC_SCHEDULING") Then
	MapDriveDelete "T:"
	MapDrive "T:", "rhcweb01\scheduling" 
End If


If InGroup("CN=WSMC_VELLOS_USERS") Then
	MapDriveDelete "v:"
	MapDrive "v:", "wsmcapps3\apps\vellos"
End If




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Configure the Printers by Group
'NOTE: The InGroup conditional test is CASE SENSITIVE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'You can only create printer on a Windows95/98 machine, if the specific Printer
'Driver has already been installed.  Meaning, a printer of the same kind is
'already installed on the machine.

If PCName = "RLICXP3856A" Then 
	PrinterPath = "\\RHCPRINT\RLICIS07"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\RLICIS07"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "RMCXP5605B" Then 
	PrinterPath = "\\RMCPRINT\RMCADMIN05"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
End If

If PCName = "WHXP7050A" Then 
	PrinterPath = "\\WHPRINT\WHCM01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
End If

If PCName = "RLICXP3853A" Then 
	PrinterPath = "\\RHCPRINT\RLICIS03"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
End If

If InGroup("CN=PRT_RMCADMIN05") Then
	PrinterPath = "\\RMCPRINT\RMCADMIN05"		'Uniquely defined variable
	PrinterDriver = "Imagistics im4720/5520 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEADMIT04") Then
	PrinterPath = "\\SMEPRINT\SMEADMIT04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2300 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEADMIT08") Then
	PrinterPath = "\\SMEPRINT\SMEADMIT08"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2420 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMNLAB07") Then
	PrinterPath = "\\SMEPRINT\SMNLAB07"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2420 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRQRM01") Then
        PrinterPath = "\\OLRPRINT2\OLRQRM01"		'Uniquely defined variable
	PrinterDriver = "HP Color LaserJet 4600 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRPACU01") Then
	PrinterPath = "\\OLRPRINT2\OLRPACU01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SEHOPPHA01") Then
	PrinterPath = "\\SEHPRINT2\SEHOPPHA01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMECALLCTR03") Then
	PrinterPath = "\\SMEPRINT\SMECALLCTR03"		'Uniquely defined variable
	PrinterDriver = "Imagistics im3511 PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMESMD01") Then
	PrinterPath = "\\SMEPRINT\SMESMD01"		'Uniquely defined variable
	PrinterDriver = "Generic 75/60BW-1 PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SMEXP2412C" Then
	PrinterPath = "\\SMEPRINT\SMESURG06"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMESURG06"
End If

If PCName = "SMEXP2412E" Then
	PrinterPath = "\\SMEPRINT\SMESURG06"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMESURG06"
End If

If PCName = "SMEXP2153A" Then
	PrinterPath = "\\SMEPRINT\SMESURG06"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMESURG06"
End If

If PCName = "SJHXP3045B" Then
	PrinterPath = "\\SJHPRINT2\SJHADMT2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"		'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SJHPRINT2\SJHADMT2"  
End If

If PCName = "SMEXP3921A" Then
	PrinterPath = "\\SMEPRINT\SMESURG06"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMESURG06"
End If

If InGroup("CN=PRT_SMEQUALITY07") Then
	PrinterPath = "\\SMEPRINT\SMEQUALITY07"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2420 PCL 6"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_ALBPRK01") Then
	PrinterPath = "\\SJHPRINT2\ALBPRK01_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"		'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
	PrinterPath = "\\SJHPRINT2\ALBPRK01_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"		'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
End If

If InGroup("CN=PRT_STONY01") Then
	PrinterPath = "\\SJHPRINT2\STONY01_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver

	PrinterPath = "\\SJHPRINT2\STONY01_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
End If

If InGroup("CN=PRT_RHCAKH01") Then
	PrinterPath = "\\RHCPRINT2\RHCAKH01_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver

	PrinterPath = "\\RHCPRINT2\RHCAKH01_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
End If

If InGroup("CN=PRT_RHCPYT01") Then
	PrinterPath = "\\RHCPRINT2\RHCPYT01_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver

	PrinterPath = "\\RHCPRINT2\RHCPYT01_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
End If

If InGroup("CN=PRT_RHCMCG01") Then
	PrinterPath = "\\RHCPRINT2\8SMICH01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True

	PrinterPath = "\\RHCPRINT2\RHCMCG01_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath

	PrinterPath = "\\RHCPRINT2\RHCMCG01_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
End If

If InGroup("CN=PRT_RHCBL01") Then
	PrinterPath = "\\RHCPRINT2\RHCBL01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True

	PrinterPath = "\\RHCPRINT2\RHCBL01_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath

	PrinterPath = "\\RHCPRINT2\RHCBL01_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
End If

If InGroup("CN=PRT_SMEHR08") Then
	PrinterPath = "\\SMEPRINT\SMEHR08"		'Uniquely defined variable
	PrinterDriver = "HP Color LaserJet 4500"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCMAINT01") Then
	PrinterPath = "\\RMCPRINT\RMCMAINT01"		'Uniquely defined variable
	PrinterDriver = "Generic MFP 20ppm PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RHCNOR01") Then
	PrinterPath = "\\SJHPRINT2\PTRBILL1"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True

	PrinterPath = "\\RHCPRINT2\RHCNOR01_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath

	PrinterPath = "\\RHCPRINT2\RHCNOR01_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
End If

If InGroup("CN=PRT_SFHCAMP01") Then
	PrinterPath = "\\SFHPRINT2\SFHCAMP01_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver

	PrinterPath = "\\SFHPRINT2\SFHCAMP01_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
End If

If InGroup("CN=PRT_SFHBRSC01") Then
        PrinterPath = "\\SFHPRINT2\SFHBRSC01"
        PrinterDriver = "HP LaserJet 4000 Series PCL6"		'Printer Driver Name
        WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
End If

If InGroup("CN=PRT_SFHBRSC02") Then
        PrinterPath = "\\SFHPRINT2\SFHBRSC02"
        PrinterDriver = "HP LaserJet 4000 Series PCL6"		'Printer Driver Name
        WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
        WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_SMNFC_PA01") Then
	PrinterPath = "\\SMEPRINT\SMNFCPA01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5Si/5Si MX PS"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMNFCPA01"
End If

If InGroup("CN=PRT_SMNFC_PA02") Then
	PrinterPath = "\\SMEPRINT\SMNFCPA02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8000 Series PS"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMNFCPA02"
End If

If InGroup("CN=PRT_SMECSCHED02") Then
        PrinterPath = "\\SMEPRINT\SMECSCHED02"
        PrinterDriver = "HP LaserJet 4300 Series PCL 6"		'Printer Driver Name
        WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
End If

If InGroup("CN=PRT_SMNFC_PA05_ND") Then
	PrinterPath = "\\SMEPRINT\SMNFCPA05"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4 Plus"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
       	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMNFCPA05"
End If

If InGroup("CN=PRT_HFHTNPR01") Then
	PrinterPath = "\\HFHPRINT\HFHTNPR01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4000 Series PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If InGroup("CN=PRT_HFMCPF08") Then
	PrinterPath = "\\HFHPRINT\HFMCPF08"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
        MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "HFHPRINT\HFMCPF08"
End If

If InGroup("CN=PRT_HFMC_PA02") Then
	PrinterPath = "\\HFHPRINT\HFMCPA02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4 Plus"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt2:"
	MapDrive "LPT2:", "HFHPRINT\HFMCPA02"
End If

If InGroup("CN=PRT_HFMC_PA01") Then
	PrinterPath = "\\HFHPRINT\HFMCPA01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 6"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "HFHPRINT\HFMCPA01"
End If

If InGroup("CN=PRT_HFHIS07") Then
	PrinterPath = "\\HFHPRINT\HFHIS07"		'Uniquely defined variable
	PrinterDriver = "Imagistics im4511 PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_HFHIS08") Then
	PrinterPath = "\\HFHPRINT\HFHIS08"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=HFMCIS01") Then
	PrinterPath = "\\HFHPRINT\HFMCIS01"		'Uniquely defined variable
	PrinterDriver = "HP Business Inkjet 2250 PCL 5C"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=HFMCIS02") Then
	PrinterPath = "\\HFHPRINT\HFMCIS02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=OLR Medical Records") Then
	PrinterPath = "\\OLRPRINT2\OLRMEDREC"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRLAB06") Then
	PrinterPath = "\\OLRPRINT2\OLRLAB06"		'Uniquely defined variable
	PrinterDriver = "Generic MFP 25ppm f PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRLAB07") Then
	PrinterPath = "\\OLRPRINT2\OLRLAB07"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_MHADM01") Then
	PrinterPath = "\\RHCPRINT2\MHADM01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_MHMEDREC01") Then
	PrinterPath = "\\RHCPRINT2\MHMEDREC01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5L"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_MGPM01") Then
	PrinterPath = "\\RHCPRINT\MGPM01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\MGPM01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "RHCPRINT2\MGPM01"
End If

If InGroup("CN=PRT_SEH_PSY6") Then
	PrinterPath = "\\SEHPRINT2\SEHPSY05"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SEHPRINT2\SEHPSY05"
End If

If InGroup("CN=PRT_RHCMED01") Then
	PrinterPath = "\\RHCPRINT2\RHCMED01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 5e"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_RHCFMC01") Then
	PrinterPath = "\\RHCPRINT2\RHCFMC01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True

	PrinterPath = "\\RHCPRINT2\RHCFMC01_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath

	PrinterPath = "\\RHCPRINT2\RHCFMC01_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RHCFRA01") Then
	PrinterPath = "\\RHCPRINT2\RHCFRA01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True

	PrinterPath = "\\RHCPRINT2\RHCFRA01_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath

	PrinterPath = "\\RHCPRINT2\RHCFRA01_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RHCRSC01") Then
	PrinterPath = "\\RHCPRINT2\RHCRSC01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True

	PrinterPath = "\\RHCPRINT2\RHCRSC01_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath

	PrinterPath = "\\RHCPRINT2\RHCRSC01_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RHCRP01") Then
	PrinterPath = "\\RHCPRINT2\RHCRP01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True

	PrinterPath = "\\RHCPRINT2\RHCRP01_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath

	PrinterPath = "\\RHCPRINT2\RHCRP01_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCFP02") Then
	PrinterPath = "\\RMCPRINT\RMCFP02"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True

	PrinterPath = "\\RMCPRINT\RMCFP02_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath

	PrinterPath = "\\RMCPRINT\RMCFP02_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMNSPSRV01") Then
	PrinterPath = "\\SMEPRINT\SMESPSRV01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WshNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_WSMCTCOM01") Then
	PrinterPath = "\\WSMCPRINT\WSMCTCOM01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WshNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_WSMCMRKT01") Then
	PrinterPath = "\\WSMCPRINT\WSMCMRKT01"		'Uniquely defined variable
	PrinterDriver = "HP Business Inkjet 2300 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_WSMCBD01") Then
	PrinterPath = "\\WSMCPRINT\WSMCBD01"		'Uniquely defined variable
	PrinterDriver = "HP Business Inkjet 2300 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SEHPSY05") Then
	PrinterPath = "\\SEHPRINT2\SEHPSY05"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SEHPRINT2\SEHPSY05"
End If

If InGroup("CN=PRT_WSMCHIS") Then
	PrinterPath = "\\WSMCPRINT\WSMCHISAPPS01"		'Uniquely defined variable
	PrinterDriver = "hp color LaserJet 2550 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_WSMCHIS") Then
	PrinterPath = "\\WSMCPRINT\WSMCHISAPPS02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5Si"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_WSMCHIS") Then
	PrinterPath = "\\WSMCPRINT\WSMCFSVC01"	'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4 Plus"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_WSMCHIS") Then
	PrinterPath = "\\WSMCPRINT\WSMCFSVC02"		'Uniquely defined variable
	PrinterDriver = "hp color LaserJet 2550 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCHAD01") Then
	PrinterPath = "\\RMCPRINT\RMCHAD01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True

	PrinterPath = "\\RMCPRINT\RMCHAD01_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath

	PrinterPath = "\\RMCPRINT\RMCHAD01_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RHCDEVON") Then
	PrinterPath = "\\RHCPRINT2\RHCDEVON"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True

	PrinterPath = "\\RHCPRINT2\RHCDEVON_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath

	PrinterPath = "\\RHCPRINT2\RHCDEVON_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RHCPHYS02") Then
	PrinterPath = "\\RHCPRINT2\RHCPHYS02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
'	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_RHCCOLLSR1") Then
	PrinterPath = "\\RHCPRINT2\RHCCOLLSR1"		'Uniquely defined variable
	PrinterDriver = "HP Color LaserJet 4650 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
'	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_WHPFS01") Then
	PrinterPath = "\\WHPRINT\WHPFS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_RMCLS01") Then
	PrinterPath = "\\RMCPRINT\RMCLS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 5e"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRCHF") Then
	PrinterPath = "\\OLRPRINT2\OLRCHF"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4L"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRDRLG") Then
	PrinterPath = "\\OLRPRINT2\OLRDRLG"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLR2N") Then
	PrinterPath = "\\OLRPRINT2\OLR2N"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_OLR3S") Then
	PrinterPath = "\\OLRPRINT2\OLR3S"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLR3N") Then
	PrinterPath = "\\OLRPRINT2\OLR3N"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLR4S") Then
	PrinterPath = "\\OLRPRINT2\OLR4S"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRPA05") Then
	PrinterPath = "\\HFHPRINT\OLRPA05"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 5e"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "HFHPRINT\OLRPA05"
End If

If InGroup("CN=PRT_OLR4E") Then
	PrinterPath = "\\OLRPRINT2\OLR4E"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	'WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_OLRPA06") Then
	PrinterPath = "\\HFHPRINT\OLRPA06"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	'WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_OLR5E") Then
	PrinterPath = "\\OLRPRINT2\OLR5E"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRCCRM3") Then
	PrinterPath = "\\OLRPRINT2\OLRCCRM3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRCCRM1") Then
	PrinterPath = "\\OLRPRINT2\OLRCCRM1"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRCCRMROOM332") Then
	PrinterPath = "\\OLRPRINT2\OLR_CCRM_ROOM332"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRNADM") Then
	PrinterPath = "\\OLRPRINT2\OLRNADM"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLREKG") Then
	PrinterPath = "\\OLRPRINT2\OLREKG"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRVPPS") Then
	PrinterPath = "\\OLRPRINT2\OLRVPPS"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRIS") Then
	PrinterPath = "\\OLRPRINT2\OLRIS"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRICN") Then
	PrinterPath = "\\OLRPRINT2\OLRICN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRER") Then
	PrinterPath = "\\OLRPRINT2\OLRER"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5M"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRAMB01") Then
	PrinterPath = "\\OLRPRINT2\OLRAMB01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLROLRPATHOLOGY") Then
	PrinterPath = "\\OLRPRINT2\OLRPATHOLOGY"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WshNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_SJHRHAB1") Then
	PrinterPath = "\\SJHPRINT2\SJHRHAB1"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 2200 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If InGroup("CN=PRT_RLICIS01") Then
	PrinterPath = "\\RHCPRINT2\RLICIS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS02") Then
	PrinterPath = "\\RHCPRINT2\RLICIS02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS03") Then
	PrinterPath = "\\RHCPRINT\RLICIS03"
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\RLICIS03"		'Uniquely defined variable
	PrinterDriver = "HP Color LaserJet 5500 PCL6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS04") Then
	PrinterPath = "\\RHCPRINT2\RLICIS04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
'	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_RLICIS05") Then
	PrinterPath = "\\RHCPRINT2\RLICIS05"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS09") Then
	PrinterPath = "\\RHCPRINT2\RLICIS09"		'Uniquely defined variable
	PrinterDriver = "HP C LaserJet 4500-PS"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS10") Then
	PrinterPath = "\\RHCPRINT2\RLICIS10"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "RHCPRINT2\RLICIS10"
End If

If InGroup("CN=PRT_RLICIS11") Then
	PrinterPath = "\\RHCPRINT2\RLICIS11"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "RHCPRINT2\RLICIS11"
End If

If InGroup("CN=PRT_RLICIS12") Then
	PrinterPath = "\\RHCPRINT2\RLICIS12"		'Uniquely defined variable
	PrinterDriver = "HP 2500C Series"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS13") Then
	PrinterPath = "\\RHCPRINT2\RLICIS13"		'Uniquely defined variable
	PrinterDriver = "HP Color LaserJet 4550 PCL 6"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS14") Then
	PrinterPath = "\\RHCPRINT2\RLICIS14"		'Uniquely defined variable
	PrinterDriver = "HP DesignJet 800PS 42 by HP"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS15") Then
	PrinterPath = "\\RHCPRINT2\RLICIS15"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS16") Then
	PrinterPath = "\\RHCPRINT2\RLICIS16"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS17") Then
	PrinterPath = "\\RHCPRINT2\RLICIS17"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS18") Then
	PrinterPath = "\\RHCPRINT\RLICIS18"
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\RLICIS18"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS19") Then
	PrinterPath = "\\RHCPRINT\RLICIS19"
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\RLICIS19"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS20") Then
	PrinterPath = "\\RHCPRINT2\RLICIS20"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 PCL6" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS21") Then
	PrinterPath = "\\RHCPRINT\RLICIS21"
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\RLICIS21"		'Uniquely defined variable
	PrinterDriver = "Imagistics im4510 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS29") Then
	PrinterPath = "\\RHCPRINT2\RLICIS29"		'Uniquely defined variable
	PrinterDriver = "Imagistics im4720 PCL6" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS31") Then
	PrinterPath = "\\RHCPRINT2\RLICIS31"		'Uniquely defined variable
	PrinterDriver = "HP Color LaserJet 4250 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICGRAPHICS04") Then
	PrinterPath = "\\RHCPRINT2\RLICGRAPHICS04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4V"		 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCDay1") Then
	PrinterPath = "\\RMCPRINT\RMCDay1"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 6P"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=RMC2E01") Then
	PrinterPath = "\\RMCPRINT\RMC2E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCCOM01") Then
	PrinterPath = "\\RMCPRINT\RMCCOM01"		'Uniquely defined variable
	PrinterDriver = "HP Color LaserJet 4600 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
'	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_RMCCOM02") Then
	PrinterPath = "\\RMCPRINT\RMCCOM02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
'	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_RMCCOM03") Then
	PrinterPath = "\\RMCPRINT\RMCCOM03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCLAB01") Then
	PrinterPath = "\\RMCPRINT\RMCLAB01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2200 Series PCL" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCPHARM01") Then
	PrinterPath = "\\RMCPRINT\RMCPHARM01"		'Uniquely defined variable
	PrinterDriver = "HP Business Inkjet 2300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCRMCHR03") Then
	PrinterPath = "\\RMCPRINT\RMCHR03"		'Uniquely defined variable
	PrinterDriver = "HP Business Inkjet 2300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS23") Then
	PrinterPath = "\\RHCPRINT2\RLICIS23"		'Uniquely defined variable
	PrinterDriver = "HP Business Inkjet 2300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RLICIS24") Then
	PrinterPath = "\\RHCPRINT2\RLICIS24"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCADM1") Then
	PrinterPath = "\\RMCPRINT\RMCADM1"		'Uniquely defined variable
	PrinterDriver = "Lexmark Optra R Series"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SFHLAB01") Then
	PrinterPath = "\\SFHPRINT2\SFHLAB01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SKOHHCB4") Then
	PrinterPath = "\\RHCPRINT2\SKOHHCB4"		'Uniquely defined variable
	PrinterDriver = "Imagistics im4511 PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SKOHHCB2") Then
	PrinterPath = "\\RHCPRINT2\SKOHHCB2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SKOHHCB1") Then
	PrinterPath = "\\RHCPRINT2\SKOHHCB1"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SKOHH01") Then
	PrinterPath = "\\RHCPRINT2\SKOHH01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SKOHH02") Then
	PrinterPath = "\\RHCPRINT2\SKOHH02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SKOHH15") Then
	PrinterPath = "\\RHCPRINT2\SKOHH15"		'Uniquely defined variable
	PrinterDriver = "HP Color LaserJet 4600 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If


If InGroup("CN=PRT_SMECR01") Then
	PrinterPath = "\\SMEPRINT\SMECR01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2300 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEADMIN09") Then
	PrinterPath = "\\SMEPRINT\SMEADMIN09"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet Business Inkjet 2300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEADMIN10") Then
	PrinterPath = "\\SMEPRINT\SMEADMIN10"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEADMIN11") Then
	PrinterPath = "\\SMEPRINT\SMEADMIN11"		'Uniquely defined variable
	PrinterDriver = "HP Color LaserJet 3500"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEPFS07") Then
	PrinterPath = "\\SMEPRINT\SMEPFS07"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2300 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEDIABCTR01") Then
	PrinterPath = "\\SMEPRINT\SMEDIABCTR01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
        WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_SMEGILAB01") Then
	PrinterPath = "\\SMEPRINT\SMEGILAB01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SME10FLR01") Then
	PrinterPath = "\\SMEPRINT\SME10FLR01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
        WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_SMESPD02") Then
	PrinterPath = "\\SMEPRINT\SMESPD02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEFAM02") Then
	PrinterPath = "\\SMEPRINT\SMEFAM02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEHR05") Then
	PrinterPath = "\\SMEPRINT\SMEHR05"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEHR06") Then
	PrinterPath = "\\SMEPRINT\SMEHR06"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEMEDSTAFF02") Then
	PrinterPath = "\\SMEPRINT\SMEMEDSTAFF02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEMEDREC05") Then
	PrinterPath = "\\SMEPRINT\SMEMEDREC05"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2300 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEMEDREC08") Then
	PrinterPath = "\\SMEPRINT\SMEMEDREC08"		'Uniquely defined variable
	PrinterDriver = "HP Business Inkjet 2300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEMEDREC07") Then
	PrinterPath = "\\SMEPRINT\SMEMEDREC07"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2300 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEPFS03") Then
	PrinterPath = "\\SMEPRINT\SMEPFS03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMEPFS03"
End If

If InGroup("CN=PRT_SMEPFS11") Then
	PrinterPath = "\\SMEPRINT\SMEPFS11"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8150 Series PCL"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMEPFS11"
End If

If InGroup("CN=PRT_SMEPTOT01") Then
	PrinterPath = "\\SMEPRINT\SMEPTOT01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEPTOT02") Then
	PrinterPath = "\\SMEPRINT\SMEPTOT02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEPTOT03") Then
	PrinterPath = "\\SMEPRINT\SMEPTOT03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEPTOT04") Then
	PrinterPath = "\\SMEPRINT\SMEPTOT04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMENURSEAD01") Then
	PrinterPath = "\\SMEPRINT\SMENURSEAD01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEADMIN02") Then
	PrinterPath = "\\SMEPRINT\SMEADMIN02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEADMIN03") Then
	PrinterPath = "\\SMEPRINT\SMEADMIN03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEADMIN05") Then
	PrinterPath = "\\SMEPRINT\SMEADMIN05"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4 Plus"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEADMIN08") Then
	PrinterPath = "\\SMEPRINT\SMEADMIN08"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
'        WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_SMENURSEAD02") Then
	PrinterPath = "\\SMEPRINT\SMENURSEAD02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMERAD05") Then
	PrinterPath = "\\SMEPRINT\SMERAD05"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMERAD06") Then
	PrinterPath = "\\SMEPRINT\SMERAD06"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_SMERAD06PS") Then
	PrinterPath = "\\SMEPRINT\SMERAD06PS"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_SMERAD07") Then
	PrinterPath = "\\SMEPRINT\SMERAD07"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMERAD08") Then
	PrinterPath = "\\SMEPRINT\SMERAD08"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMETRAIN01") Then
	PrinterPath = "\\SMEPRINT\SMETRAIN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMETRAIN03") Then
	PrinterPath = "\\SMEPRINT\SMETRAIN03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMETRAIN04") Then
	PrinterPath = "\\SMEPRINT\SMETRAIN04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2200 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMECALLCTR01") Then
	PrinterPath = "\\SMEPRINT\SMECALLCTR01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4650 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMECALLCTR02") Then
	PrinterPath = "\\SMEPRINT\SMECALLCTR02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEMEDSTAFF03") Then
	PrinterPath = "\\SMEPRINT\SMEMEDSTAFF03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4650 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMNPFS07") Then
	PrinterPath = "\\SMEPRINT\SMNPFS07"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMNPFS07"
End If

If InGroup("CN=PRT_SMNPFS01") Then
	PrinterPath = "\\SMEPRINT\SMNPFS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
        MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMNPFS01"
End If

If InGroup("CN=PRT_SMNPFS04") Then
	PrinterPath = "\\SMEPRINT\SMNPFS04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4P"	        'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMNPFS04"
End If

If InGroup("CN=PRT_SKOHH04") Then
	PrinterPath = "\\RHCPRINT2\SKOHH04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SKOHH05") Then
	PrinterPath = "\\RHCPRINT2\SKOHH05"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SKOHH06_Default") Then
	PrinterPath = "\\RHCPRINT2\SKOHH06"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_SKOHH06") Then
	PrinterPath = "\\RHCPRINT2\SKOHH06"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SKOHH07_Default") Then
	PrinterPath = "\\RHCPRINT2\SKOHH07"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_SKOHH07") Then
	PrinterPath = "\\RHCPRINT2\SKOHH07"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

'If InGroup("CN=PRT_SKOHH08") Then
'	PrinterPath = "\\RHCPRINT2\SKOHH08"		'Uniquely defined variable
'	PrinterDriver = "HP LaserJet 2100 Series PS"	'Printer Driver Name
'	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
'End If

If InGroup("CN=PRT_SKOHH09") Then
	PrinterPath = "\\RHCPRINT2\SKOHH09"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

'If InGroup("CN=PRT_SKOHH10") Then
'	PrinterPath = "\\RHCPRINT2\SKOHH10"		'Uniquely defined variable
'	PrinterDriver = "HP LaserJet 2100 Series PS"	'Printer Driver Name
'	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
'End If

If InGroup("CN=PRT_SKOHH11") Then
	PrinterPath = "\\RHCPRINT2\SKOHH11"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SKOHHPR1") Then
	PrinterPath = "\\RHCPRINT2\SKOHHPR1"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SKOHH14") Then
	PrinterPath = "\\RHCPRINT2\SKOHH14"		'Uniquely defined variable
	PrinterDriver = "HP Color LaserJet 4600 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SKOHH17a") Then
	PrinterPath = "\\RHCPRINT2\SKOHH17"		'Uniquely defined variable
	PrinterDriver = "Imagistics im8130/im6530 PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_SKOHH18a") Then
	PrinterPath = "\\RHCPRINT2\SKOHH18"		'Uniquely defined variable
	PrinterDriver = "Imagistics im8130/im6530 PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SKOHH17b") Then
	PrinterPath = "\\RHCPRINT2\SKOHH17"		'Uniquely defined variable
	PrinterDriver = "Imagistics im8130/im6530 PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SKOHH18b") Then
	PrinterPath = "\\RHCPRINT2\SKOHH18"		'Uniquely defined variable
	PrinterDriver = "Imagistics im8130/im6530 PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_SKOHHPR1_Default") Then
	PrinterPath = "\\RHCPRINT2\SKOHHPR1"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_RHACOLLSR1") Then
	PrinterPath = "\\RHCPRINT2\RHACOLLSR1"		'Uniquely defined variable
	PrinterDriver = "HP Color LaserJet 4650 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCHR02") Then
	PrinterPath = "\\RMCPRINT\RMCHR02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCQIR3") Then
	PrinterPath = "\\RMCPRINT\RMCQIR3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If



If PCName = "WHXP7706A" Then
	PrinterPath = "\\WHPRINT\WHSPD01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
        MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WHPRINT\WHSPD01"
End If

If PCName = "WHXP7167A" Then
	PrinterPath = "\\WHPRINT\WHOR03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
        MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WHPRINT\WHOR03"
End If

If PCName = "WHXP7167B" Then
	PrinterPath = "\\WHPRINT\WHOR03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
        MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WHPRINT\WHOR03"
End If

If PCName = "WHXP7460A" Then
	PrinterPath = "\\WHPRINT\WHOPS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
        MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WHPRINT\WHOPS01"
End If

If PCName = "WHXP7168A" Then
	PrinterPath = "\\WHPRINT\WHRECOV01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
        MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WHPRINT\WHRECOV01"
End If

If PCName = "WHXP7168B" Then
	PrinterPath = "\\WHPRINT\WHRECOV01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
        MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WHPRINT\WHRECOV01"
End If

If PCName = "WH2K7543B" Then
	PrinterPath = "\\WHPRINT\WHCH01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2200 PCL6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_WHHR01") Then
	PrinterPath = "\\WHPRINT\WHHR01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"		'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_SFHHCLIN01") Then
	PrinterPath = "\\SFHPRINT2\SFHHCLIN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath

	PrinterPath = "\\SFHPRINT2\SFHHCLIN04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=WHWIC Printer") Then
	PrinterPath = "\\WHPRINT\WHWIC"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "WH2K7644A" Then
	PrinterPath = "\\WHPRINT\WHPOB307"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WHPRINT\WHPOB307"
End If

If PCName = "WHXP7420A" Then
	PrinterPath = "\\WHPRINT\WHSURG02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
        MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WHPRINT\WHSURG02"
End If

If PCName = "WHXP7463A" Then
	PrinterPath = "\\WHPRINT\WHBIO01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "WHXP7105B" Then
	PrinterPath = "\\WHPRINT\WHOPS02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 PCL6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
        MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WHPRINT\WHOPS02"
End If

If InGroup("CN=PRT_WSMC_WSMCPFS01") Then
	PrinterPath = "\\WSMCPRINT\WSMCPFS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8000 Series PCL"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WSMCPRINT\WSMCPFS01"
End If

If InGroup("CN=PRT_WSMC_WSMCPFS02") Then
	PrinterPath = "\\WSMCPRINT\WSMCPFS02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt2:"
	MapDrive "LPT2:", "WSMCPRINT\WSMCPFS02"
End If

If InGroup("CN=PRT_WSMCTRANS01") Then
	PrinterPath = "\\WSMCPRINT\WSMCTRANS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WshNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_WSMCHH01") Then
	PrinterPath = "\\RHCPRINT2\WSMCHH01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEDIET02") Then
	PrinterPath = "\\SMEPRINT\SMNDIET02"		'Uniquely defined variable
	PrinterDriver = "HP Color LaserJet 3500"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMC2S01") Then
	PrinterPath = "\\RMCPRINT\RMC2S01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMC2W01") Then
	PrinterPath = "\\RMCPRINT\RMC2W01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMC3S01") Then
	PrinterPath = "\\RMCPRINT\RMC3S01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMC4E01") Then
	PrinterPath = "\\RMCPRINT\RMC4E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCADMIN02") Then
	PrinterPath = "\\RMCPRINT\RMCADMIN02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "RMCXP5278A" Then
	PrinterPath = "\\RMCPRINT\RMC4E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "RMCXP6348A" Then
	PrinterPath = "\\RMCPRINT\RMC4E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "RMCXP7909A" Then
	PrinterPath = "\\RMCPRINT\RMC4E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "RMC2K7909C" Then
	PrinterPath = "\\RMCPRINT\RMC4E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "RMCXP7904C" Then
	PrinterPath = "\\RMCPRINT\RMC2E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "RMCXP7942D" Then
	PrinterPath = "\\RMCPRINT\RMCSICU"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "RMCXP5411B" Then
	PrinterPath = "\\RMCPRINT\RMCMICU"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "RMCXP6334A" Then
	PrinterPath = "\\RMCPRINT\RMCPATREG"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "RMCXP2979A" Then
	PrinterPath = "\\RMCPRINT\RMCPATREG"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_RSCPRT01") Then
	PrinterPath = "\\RHCPRINT\RSCPRT01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\RSCPRT01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "RHCPRINT2\RSCPRT01"
End If

If InGroup("CN=PRT_RSCPRT02") Then
	PrinterPath = "\\RHCPRINT\RSCPRT02"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\RSCPRT02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "RHCPRINT2\RSCPRT02"
End If

If InGroup("CN=PRT_RSCPRT03") Then
	PrinterPath = "\\RHCPRINT\RSCPRT03"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\RSCPRT03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "RHCPRINT2\RSCPRT03"
End If

If InGroup("CN=OLR Emergency Room") Then
	PrinterPath = "\\OLRPRINT2\OLRADMIT"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=OLROPRE") Then
	PrinterPath = "\\OLRPRINT2\OLROPREG"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_HFMCLAB02") Then
	PrinterPath = "\\HFHPRINT\HFMCLAB02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_HFMCLAB03") Then
	PrinterPath = "\\HFHPRINT\HFMCLAB03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLR4W") Then
	PrinterPath = "\\OLRPRINT2\OLR4W"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4M Plus"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRMDS01") Then
	PrinterPath = "\\OLRPRINT2\OLRMDS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_OLRCASH01") Then
	PrinterPath = "\\OLRPRINT2\OLRCASH01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4 Plus"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If InGroup("CN=RMC4W01") Then
	PrinterPath = "\\RMCPRINT\RMC4W01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMC4SP01") Then
	PrinterPath = "\\RMCPRINT\RMC4S"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMCCC01") Then
	PrinterPath = "\\RMCPRINT\RMCCC01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCER") Then
	PrinterPath = "\\RMCPRINT\RMCER01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMCEMRP01") Then
	PrinterPath = "\\RMCPRINT\RMCEMRP01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4 Plus"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCHR04") Then
	PrinterPath = "\\RMCPRINT\RMCHR04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4SI"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCPAY01") Then
	PrinterPath = "\\RMCPRINT\RMCPAY01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4SI"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	PrinterPath = "\\RMCPRINT\RMCPAY02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4SI"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMCEMRP02") Then
	PrinterPath = "\\RMCPRINT\RMCEMRP02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8000 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMNDEV01") Then
	PrinterPath = "\\SMEPRINT\SMNDEV01"		'Uniquely defined variable
	PrinterDriver = "HP DeskJet 1600C Printer"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMECASHIER01") Then
	PrinterPath = "\\SMEPRINT\SMECASHIER01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMECASHIER01"
End If

If InGroup("CN=PRT_RMCRAD10") Then
	PrinterPath = "\\RMCPRINT\RMCRAD10"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCRAD13") Then
	PrinterPath = "\\RMCPRINT\RMCRAD13"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCREHAB01") Then
	PrinterPath = "\\RMCPRINT\RMCREHAB01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4000 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "RMCXP5535B" Then
	PrinterPath = "\\RMCPRINT\RMCREHAB01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4000 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "RMCXP9956A" Then
	PrinterPath = "\\RMCPRINT\RMCRAD10"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SJH2K4127B" Then
	PrinterPath = "\\SJHPRINT2\SJHLAB04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8100 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SJH2K3449B" Then
	PrinterPath = "\\SJHPRINT2\SJHADMT2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"		'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
	WshNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SJHPRINT2\SJHADMT2"  
End If


If InGroup("CN=PRT_hme01") Then	
	PrinterPath = "\\rhcapps202\hme01"		'Uniquely defined variable
	PrinterDriver = "Generic / Text Only"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMESMD01") Then
	PrinterPath = "\\SMEPRINT\SMESMD01"		'Uniquely defined variable
	PrinterDriver = "Imagistics im4720 PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver	
End If

If InGroup("CN=RMCIS") Then
	PrinterPath = "\\RHCPRINT\RMCIS"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RMCPRINT\RMCIS"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6" 	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMCIS2") Then
	PrinterPath = "\\RMCPRINT\RMCIS2"		'Uniquely defined variable
	PrinterDriver = "HP Business Inkjet 2200/2250(HPA)" 	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMCNA1") Then
	PrinterPath = "\\RMCPRINT\RMCNA1"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMCNA02") Then
	PrinterPath = "\\RMCPRINT\RMCNA02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMCHR01") Then
	PrinterPath = "\\RMCPRINT\RMCHR01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCHR1700") Then
	PrinterPath = "\\RMCPRINT\RMCHR1700"		'Uniquely defined variable
	PrinterDriver = "HP Business Inkjet 2200\2250(HBA)"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If


' The following printer settings are required for the group to print NEBO.
If InGroup("CN=RMCPA01") Then
	PrinterPath = "\\HFHPRINT\RMCPA01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "HFHPRINT\RMCPA01"
End If

If InGroup("CN=PRT_HFMCPA2") Then
	PrinterPath = "\\HFHPRINT\HFMCPA2"	'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDrive "LPT2:", "HFHPRINT\HFMCPA2"
	MapDrive "LPT1:", "HFHPRINT\HFMCPA2"
End If

If InGroup("CN=RMCPSI02") Then
	PrinterPath = "\\RMCPRINT\RMCPSI02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMCRAD02") Then
	PrinterPath = "\\RMCPRINT\RMCRAD02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMCFP01") Then
	PrinterPath = "\\RMCPRINT\RMCFP01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 6P"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCMR4") Then
	PrinterPath = "\\RMCPRINT\RMCMR4"		'Uniquely defined variable
	PrinterDriver = "Imagistics im8130/im6530 PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMCMRKT1") Then
	PrinterPath = "\\RMCPRINT\RMCMRKT1"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL6" 	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMCPHAR") Then
	PrinterPath = "\\RMCPRINT\RMCPHAR"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 6P/6MP - Standard"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMCPIR01") Then
	PrinterPath = "\\RMCPRINT\RMCPIR01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt2:"
	MapDrive "LPT2:", "RMCPRINT\RMCPIR01"
End If

If InGroup("CN=RMCPI01") Then
	PrinterPath = "\\RHCPRINT2\RMCPI01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PS" 	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_PF05") Then
	PrinterPath = "\\SFHPRINT2\PF05"	'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDrive "LPT1:", "SFHPRINT2\PF05"
End If

If InGroup("CN=RMCPHYS01") Then
	PrinterPath = "\\RHCPRINT2\RMCPHYS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2200 Series PCL 6" 	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=RMCMR1") Then
	PrinterPath = "\\rmcPRINT\rmcmr1"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=SFHTRAIN01") Then
	PrinterPath = "\\SFHPRINT2\SFHTRN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SFHXP2102A" Then
	PrinterPath = "\\SFHPRINT2\SFHCASH01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5"  	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHCASH01"
End If

If InGroup("CN=WHCC01") Then
	PrinterPath = "\\WHPRINT\WHCC01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=Wh Training Printer") Then
	PrinterPath = "\\WHPRINT\WHTRAIN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_WHTRAIN01") Then
	PrinterPath = "\\WHPRINT\WHTRAIN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SFHPF01") Then
	PrinterPath = "\\SFHPRINT2\SFHPF01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5"         	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHPF01"
End If

If InGroup("CN=PRT_SFHPF02") Then
	PrinterPath = "\\SFHPRINT2\SFHPF02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 5e"  	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHPF02"
End If

If InGroup("CN=PRT_SFHPF03") Then
	PrinterPath = "\\SFHPRINT2\SFHPF03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 5e"  	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHPF03"
End If

If InGroup("CN=PRT_SFHPF04") Then
	PrinterPath = "\\SFHPRINT2\SFHPF04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SFHPF05") Then
	PrinterPath = "\\SFHPRINT2\SFHPF05"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHPF05"
	MapDriveDelete "lpt2:"
	MapDrive "LPT2:", "SFHPRINT2\SFHPF05"
End If

If InGroup("CN=PRT_SFHPF09") Then
	PrinterPath = "\\SFHPRINT2\SFHPF09"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHPF09"
	MapDriveDelete "lpt2:"
	MapDrive "LPT2:", "SFHPRINT2\SFHPF09"
End If

If InGroup("CN=PRT_SFHPF10") Then
	PrinterPath = "\\SFHPRINT2\SFHPF10"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHPF10"
	MapDriveDelete "lpt2:"
	MapDrive "LPT2:", "SFHPRINT2\SFHPF10"
End If

If InGroup("CN=PRT_SFHLAB08") Then
	PrinterPath = "\\SFHPRINT2\SFHLAB08"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8150 PCL 6"  	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If	

If InGroup("CN=PRT_SFHCASH01") Then
	PrinterPath = "\\SFHPRINT2\SFHCASH01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\SFHPRINT2\SFHCASH01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5"          	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHCASH01"
End If

If InGroup("CN=PRT_HFNRPFS01") Then
	PrinterPath = "\\RHCPRINT\HFNRPFS01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\HFNRPFS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"          	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "RHCPRINT2\HFNRPFS01"
End If

If InGroup("CN=PRT_SEHSUR01") Then
	PrinterPath = "\\SEHPRINT2\SEHSUR01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL 6" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SEHADM01") Then
	PrinterPath = "\\SEHPRINT2\SEHADM01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4 Plus" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SEHAMB01") Then
	PrinterPath = "\\SEHPRINT2\SEHAMB01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 6" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SEHAMB02") Then
	PrinterPath = "\\SEHPRINT2\SEHAMB02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 6" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMECOMM1") Then
	PrinterPath = "\\SMEPRINT\SMECOMM1"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 6" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMECOMMAND01") Then
	PrinterPath = "\\SMEPRINT\SMEICC01"		'Uniquely defined variable
	PrinterDriver = "Generic MFP 25ppm f PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver	
End If

If InGroup("CN=PRT_SEHFP01") Then
	PrinterPath = "\\SEHPRINT\SEHFP01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\SEHPRINT2\SEHFP01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SEHHIS01") Then
	PrinterPath = "\\SEHPRINT\SEHHIS01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\SEHPRINT2\SEHHIS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL 6" 'Printer Driver Name	
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SEHHIS02") Then
	PrinterPath = "\\SEHPRINT\SEHHIS02"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\SEHPRINT2\SEHHIS02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SEHHIS03") Then
	PrinterPath = "\\SEHPRINT\SEHHIS03"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\SEHPRINT2\SEHHIS03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4 Plus" 'Printer Driver Name			
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SEHGIL01") Then
	PrinterPath = "\\SEHPRINT\SEHGIL01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\SEHPRINT2\SEHGIL01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4 Plus" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SEHPSY01") Then
	PrinterPath = "\\SEHPRINT\SEHPSY01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\SEHPRINT2\SEHPSY01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4 Plus" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SEHPHA01") Then
	PrinterPath = "\\SEHPRINT\SEHPHA01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\SEHPRINT2\SEHPHA01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL 6" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SEHREH01") Then
	PrinterPath = "\\SEHPRINT\SEHREH01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\SEHPRINT2\SEHREH01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4 Plus" 'Printer Driver Name	
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SEHXP4028B" Then
	PrinterPath = "\\SEHPRINT2\SEHEDU01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 6" 'Printer Driver Name	
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SEHXP5881B" Then
	PrinterPath = "\\SEHPRINT2\SEHSPIRITSRV01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2300 Series PCL 6" 'Printer Driver Name	
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_SMESECUR01") Then
	PrinterPath = "\\SEHPRINT\SMESECUR01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\SEHPRINT2\SMESECUR01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL 6" 'Printer Driver Name	
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEHR01") Then
	PrinterPath = "\\SMEPRINT\SMEHR01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2200 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEHR02") Then
	PrinterPath = "\\SMEPRINT\SMEHR02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEHR03") Then
        PrinterPath = "\\SMEPRINT\SMEHR03"
        WSHNetwork.RemovePrinterConnection PrinterPath, True, True
        PrinterPath = "\\SMEPRINT\SMEHR03"
        PrinterDriver = "HP LaserJet 4300 Series PCL 6"		'Printer Driver Name
        WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
End If	

If InGroup("CN=PRT_SMEHR04") Then
	PrinterPath = "\\SMEPRINT\SMEHR04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2300 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEMEDREC03") Then
	PrinterPath = "\\SMEPRINT\SMEMEDREC03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEMEDREC04") Then
        PrinterPath = "\\SMEPRINT\SMEMEDREC04"
        PrinterDriver = "HP LaserJet 4300 PS"		'Printer Driver Name
        WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
End If	

If InGroup("CN=PRT_SMEEOD01") Then
	PrinterPath = "\\SMEPRINT\SMEEOD01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4M"	        'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEFAC01") Then
	PrinterPath = "\\SMEPRINT\SMEFAC01"		'Uniquely defined variable
	PrinterDriver = "HP COLOR LASERJET 4600 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SMEXP2126A" Then
	PrinterPath = "\\SMEPRINT\SMEFAC01"		'Uniquely defined variable
	PrinterDriver = "HP COLOR LASERJET 4600 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SFHXP6120B" Then
	PrinterPath = "\\SFHPRINT2\SFHHR01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5si"               'Printer Driver Name	
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_SMEFAC05") Then
	PrinterPath = "\\SMEPRINT\SMEFAC05"	'Uniquely defined variable
	PrinterDriver = "Generic MFP 35ppm PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEINSUR02") Then
	PrinterPath = "\\SMEPRINT\SMEINSUR02"		'Uniquely defined variable
	PrinterDriver = "HP COLOR LASERJET 4600 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMETRAIN04") Then
	PrinterPath = "\\SMEPRINT\SMETRAIN04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2200 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEOT01") Then
	PrinterPath = "\\SMEPRINT\SMEOT01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEPT01") Then
	PrinterPath = "\\SMEPRINT\SMEPT01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SMEBIOMED01") Then
	PrinterPath = "\\SMEPRINT\SMEBIOMED01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2300 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_vs10") Then
	PrinterPath = "\\rhcPRINT2\vs10"		'Uniquely defined variable
	PrinterDriver = "Imagistics im4510 PS"	        'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_HFNRCCOLOR01") Then
	PrinterPath = "\\RHCPRINT2\HFNRCCOLOR01"		'Uniquely defined variable
	PrinterDriver = "HP COLOR LASERJET 4600 PCL 6"	        'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_HFNRCR01") Then
        PrinterPath = "\\RHCPRINT2\HFNRCR01"		'Uniquely defined variable
	PrinterDriver = "HP LASERJET 4 Plus"	        'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver 
End If

If InGroup("CN=PRT_HFNRCR02") Then
        PrinterPath = "\\RHCPRINT2\HFNRCR02"		        'Uniquely defined variable
	PrinterDriver = "HP LASERJET 4050 Series PCL 6"	        'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver 
End If      

If InGroup("CN=PRT_HFNRCR03") Then
        PrinterPath = "\\RHCPRINT2\HFNRCR03"		        'Uniquely defined variable
	PrinterDriver = "HP Business Inkjet 2250 PCL 5C"        'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver 
End If      

If InGroup("CN=PRT_HFNRCR04") Then
        PrinterPath = "\\RHCPRINT2\HFNRCR04"		'Uniquely defined variable
	PrinterDriver = "HP LASERJET 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver 
End If

If InGroup("CN=PRT_HFNRCRCOL2") Then
	PrinterPath = "\\RHCPRINT2\HFNRCRCOL2"		'Uniquely defined variable
	PrinterDriver = "HP COLOR LASERJET 3700 PS"	        'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If  

If InGroup("CN=PRT_RMCMRS01") Then
	PrinterPath = "\\RMCPRINT\RMCMRS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"	'Printer Driver Name
       WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCMRKT2") Then
	PrinterPath = "\\RMCPRINT\RMCMRKT2"		'Uniquely defined variable
	PrinterDriver = "Imagistics im3510 PS"		'Printer Driver Name
       WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCPA04P") Then
	PrinterPath = "\\HFHPRINT\RMCPA04P"		'Uniquely defined variable
	PrinterDriver = "Imagistics im3510 PCL6"		'Printer Driver Name
       WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "HFHPRINT\RMCPA04P"
	MapDriveDelete "lpt2:"
	MapDrive "LPT2:", "HFHPRINT\RMCPA04P"
End If

If InGroup("CN=PRT_PCBADM01") Then
	PrinterPath = "\\RHCPRINT2\PCBADM01"		'Uniquely defined variable
	PrinterDriver = "Imagistics fx2080 PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_PCBCAC01") Then
	PrinterPath = "\\RHCPRINT\PCBCAC01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCBCAC01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 PCL6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_PCBFIN01") Then
	PrinterPath = "\\RHCPRINT\PCBFIN01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCBFIN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "RHCPRINT2\PCBFIN01"
End If

If InGroup("CN=PRT_PCBTRN01") Then
	PrinterPath = "\\RHCPRINT\PCBTRN01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCBTRN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_PCM1401") Then
	PrinterPath = "\\RHCPRINT\PCM1401"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCM1401"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_PCM1402") Then
	PrinterPath = "\\RHCPRINT\PCM1402"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCM1402"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_PCM1701") Then
	PrinterPath = "\\RHCPRINT\PCM1701"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCM1701"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 6P"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_PCM1901") Then
	PrinterPath = "\\RHCPRINT2\PCM1901"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 Ps"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_PCW4501") Then
	PrinterPath = "\\RHCPRINT\PCW4501"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCW4501"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_PCW4502") Then
	PrinterPath = "\\RHCPRINT\PCW4502"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCW4502"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_PCB1801") Then
	PrinterPath = "\\RHCPRINT2\PCB1801"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SFHCMHC01") Then
	PrinterPath = "\\SFHPRINT2\SFHCMHC01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHCMHC01"
End If

If InGroup("CN=PRT_SFHMEDREC03") Then
	PrinterPath = "\\SFHPRINT2\SFHMEDREC03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If InGroup("CN=PRT_SFHMEDREC04") Then
	PrinterPath = "\\SFHPRINT2\SFHMEDREC04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "SFHXP3355A" Then
	PrinterPath = "\\SFHPRINT2\SFHPF02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 5e"  	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHPF02"
End If

If PCName = "SFHXP2117A" Then
	PrinterPath = "\\SFHPRINT2\SFHPF02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 5e"  	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHPF02"
End If

If PCName = "SFHXP6209A" Then
	PrinterPath = "\\SFHPRINT2\SFHPF02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 5e"  	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHPF02"
End If

If PCName = "SFHXP2608A" Then
	PrinterPath = "\\SFHPRINT2\SFHPF01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5"         	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHPF01"
End If

If PCName = "SFHXP2608A" Then
	PrinterPath = "\\SFHPRINT2\SFHCASH01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5"          	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHCASH01"	
End If

If InGroup("CN=PRT_HFNRC_BO01") Then
	PrinterPath = "\\RHCPRINT2\HFNRCBO01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"          	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "RHCPRINT2\HFNRCBO01"
End If

If InGroup("CN=PRT_HFNRC_BO02") Then
	PrinterPath = "\\RHCPRINT\HFNRCBO02"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\HFNRCBO02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"          	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "RHCPRINT2\HFNRCBO02"
End If

If InGroup("CN=PRT_RMCHM01") Then
	PrinterPath = "\\RMCPRINT\RMCHM01"		'Uniquely defined variable
	PrinterDriver = "HP Business Inkjet 2300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	
End If

If InGroup("CN=PRT_RMCME02") Then
	PrinterPath = "\\RMCPRINT\RMCME02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2300 Series PCL 5e"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	
End If

If InGroup("CN=PRT_SMEPACU1") Then
	PrinterPath = "\\SMEPRINT\SMEPACU1"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2300 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver	
End If

If InGroup("CN=PRT_SMEADMIN01") Then
	PrinterPath = "\\SMEPRINT\SMEadmin01"		'Uniquely defined variable
	PrinterDriver = "XEROX Phaser 8400DP PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver	
End If

If InGroup("CN=PRT_RMCRAD09") Then
	PrinterPath = "\\RMCPRINT\RMCRAD09"		'Uniquely defined variable
	PrinterDriver = "Lexmark Optra S 1650"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver	
End If

If InGroup("CN=PRT_RMCRMCCOM02") Then
	PrinterPath = "\\RMCPRINT\RMCCOM02"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver	
End If

If InGroup("CN=PRT_HFHCSI01") Then
	PrinterPath = "\\HFHPRINT\HFHCSI01"		'Uniquely defined variable
	PrinterDriver = "HP Color Laserjet 4600 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_HFHCSI02") Then
	PrinterPath = "\\HFHPRINT\HFHCSI02"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_HFHPUR002P") Then
	PrinterPath = "\\HFHPRINT\PUR002P"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5N"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_HFHXRLAS1_5") Then
	PrinterPath = "\\HFHPRINT\XRLAS1_5"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 PS"	'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_OLRADMIN01") Then
	PrinterPath = "\\OLRPRINT2\OLRADMIN01"		'Uniquely defined variable
	PrinterDriver = "Imagistics im3510 PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SJHQA03") Then
	PrinterPath = "\\SJHPRINT2\SJHQA03"		'Uniquely defined variable
	PrinterDriver = "Imagistics cm4530 Series PSL3"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SMN2K8216A" Then
	PrinterPath = "\\SMEPRINT\SMNFCPA02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8000 Series PS"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMNFCPA02"
End If

If PCName = "SMN2K8267A" Then
	PrinterPath = "\\SMEPRINT\SMNFCPA02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8000 Series PS"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMNFCPA02"
End If

If PCName = "SMN2K8513A" Then
	PrinterPath = "\\SMEPRINT\SMNFCPA02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8000 Series PS"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMNFCPA02"
End If

If PCName = "RMCXP5268B" Then
	PrinterPath = "\\RMCPRINT\RMCCASH02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "RMCPRINT\RMCCASH02"
End If

If PCName = "HHLT283" Then
	PrinterPath = "\\RMCPRINT\RMCHH01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "HHLT285" Then
	PrinterPath = "\\RMCPRINT\RMCHH01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "HHLT293" Then
	PrinterPath = "\\RMCPRINT\RMCHH01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "HHLT314" Then
	PrinterPath = "\\RMCPRINT\RMCHH01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "OLRXP3134B" Then
	PrinterPath = "\\OLRPRINT2\OLROPREG01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "OLRXP3151A" Then
	PrinterPath = "\\OLRPRINT2\OLRCASH01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4 Plus"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "OLRXP4307B" Then
	PrinterPath = "\\OLRPRINT2\OLRCATH01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "OLRXP3837E" Then
	PrinterPath = "\\OLRPRINT2\OLRCATH01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "OLRXP3837C" Then
	PrinterPath = "\\OLRPRINT2\OLRCATH01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "OLRXP3837D" Then
	PrinterPath = "\\OLRPRINT2\OLRCATH01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "OLRXP3833A" Then
	PrinterPath = "\\OLRPRINT2\OLR4S"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "OLRPRINT2\OLR4S"	
End If

If PCName = "OLRXP3751A" Then
	PrinterPath = "\\OLRPRINT2\OLR5E"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "OLRPRINT2\OLR5E"		
End If

If PCName = "OLRXP3751E" Then
	PrinterPath = "\\OLRPRINT2\OLR5E"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "OLRPRINT2\OLR5E"	
End If

If PCName = "OLRXP3751F" Then
	PrinterPath = "\\OLRPRINT2\OLR5E"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "OLRPRINT2\OLR5E"	
End If

If PCName = "OLRXP4185A" Then
	PrinterPath = "\\OLRPRINT2\OLR2N"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "OLRPRINT2\OLR2N"	
End If

If PCName = "OLRXP8476A" Then
	PrinterPath = "\\OLRPRINT2\OLRCHF"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4L"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_WHPF03") Then
	PrinterPath = "\\WHPRINT\WHPF03"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4000 Series PCL"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WHPRINT\WHPF03"
	MapDriveDelete "lpt2:"
	MapDrive "LPT2:", "WHPRINT\WHPF03"
End If

If InGroup("CN=PRT_WHPF02") Then
	PrinterPath = "\\WHPRINT\WHPF02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt2:"
	MapDrive "LPT2:", "WHPRINT\WHPF02"
End If

If InGroup("CN=PRT_WHPF01") Then
	PrinterPath = "\\WHPRINT\WHPF01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WHPRINT\WHPF01"
End If

If PCName = "WH2K7264C" Then
	PrinterPath = "\\WHPRINT\WHPF01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WHPRINT\WHPF01"
End If

If InGroup("CN=PRT_RMCPRPMG02") Then
	PrinterPath = "\\RMCPRINT\RMCPRPMG02"		'Uniquely defined variable
	PrinterDriver = "HP Business Inkjet 2600 PCL 5C"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_RMCPRPMG03") Then
	PrinterPath = "\\RMCPRINT\RMCPRPMG03"		'Uniquely defined variable
	PrinterDriver = "Imagistics im4510 PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SMEXP2110A" Then
	PrinterPath = "\\SMEPRINT\SMEADMIN09"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet Business Inkjet 2300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "RHAXP2801A" Then
	PrinterPath = "\\RHCPRINT2\RHCHME01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "RHAXP2801B" Then
	PrinterPath = "\\RHCPRINT2\RHCHME01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "RHAXP2805A" Then
	PrinterPath = "\\RHCPRINT2\RHCHME01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "RHAXP2808A" Then
	PrinterPath = "\\RHCPRINT2\RHCHME01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "RHAXP2818A" Then
	PrinterPath = "\\RHCPRINT2\RHCHME01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "RHAXP2819A" Then
	PrinterPath = "\\RHCPRINT2\RHCHME01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "RHAXP2820A" Then
	PrinterPath = "\\RHCPRINT2\RHCHME01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "SFHXP5848A" Then
	PrinterPath = "\\SFHPRINT2\SFH4S02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5M"		 'Printer Driver Name	
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP5810A" Then
	PrinterPath = "\\SFHPRINT2\SFHSOR01"	'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2200 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SMEXP2858A" Then
	PrinterPath = "\\SMEPRINT\SMEFAM02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WshNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SMEXP2474A" Then
	PrinterPath = "\\SMEPRINT\SMEFP04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2420 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WshNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SJHXP5762A" Then
	PrinterPath = "\\SJHPRINT2\SJHOPSG3"
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\SJHPRINT2\SJHOPSG3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4100 PCL 6"		'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
End If

If PCName = "SFHXP3025A" Then
	PrinterPath = "\\SFHPRINT2\SFHSPD01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SFHXP6387A" Then
	PrinterPath = "\\SFHPRINT2\SFHRADT01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP6104A" Then
	PrinterPath = "\\SFHPRINT2\SFHRADT01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2868A" Then
	PrinterPath = "\\SFHPRINT2\SFHRADT01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2230A" Then
	PrinterPath = "\\SFHPRINT2\SFHINFC01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2565A" Then
	PrinterPath = "\\SFHPRINT2\SFHRADT01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2564A" Then
	PrinterPath = "\\SFHPRINT2\SFHRADT01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2058A" Then
	PrinterPath = "\\SFHPRINT2\SFHRADT01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2548A" Then
	PrinterPath = "\\SFHPRINT2\SFHMRI01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2547A" Then
	PrinterPath = "\\SFHPRINT2\SFHMRI01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2547B" Then
	PrinterPath = "\\SFHPRINT2\SFHMRI01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2547C" Then
	PrinterPath = "\\SFHPRINT2\SFHMRI01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2370A" Then
	PrinterPath = "\\SFHPRINT2\SFHRAD03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP6375A" Then
	PrinterPath = "\\SFHPRINT2\SFHMRI01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP6377A" Then
	PrinterPath = "\\SFHPRINT2\SFHCTS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL" 	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2558A" Then
	PrinterPath = "\\SFHPRINT2\SFHCTS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL" 	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2554A" Then
	PrinterPath = "\\SFHPRINT2\SFHCTS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL" 	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2610A" Then
	PrinterPath = "\\SFHPRINT2\SFHCS01"
	PrinterDriver = "HP LaserJet 4000 Series PCL"  	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2609A" Then
	PrinterPath = "\\SFHPRINT2\SFHCS01"
	PrinterDriver = "HP LaserJet 4000 Series PCL"  	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2557A" Then
	PrinterPath = "\\SFHPRINT2\SFHXRAY01"
	PrinterDriver = "HP LaserJet 4000 Series PCL"  	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2557B" Then
	PrinterPath = "\\SFHPRINT2\SFHXRAY01"
	PrinterDriver = "HP LaserJet 4000 Series PCL"  	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2532A" Then
	PrinterPath = "\\SFHPRINT2\SFHXRAY01"
	PrinterDriver = "HP LaserJet 4000 Series PCL"  	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2791A" Then
	PrinterPath = "\\SFHPRINT2\SFHXRAY01"
	PrinterDriver = "HP LaserJet 4000 Series PCL"  	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SEHXP4465A" Then
	PrinterPath = "\\SEHPRINT2\SEHPSY05"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 PCL 6" 	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SEHPRINT2\SEHPSY05"
End If

If PCName = "SFHXP2751A" Then
	PrinterPath = "\\SFHPRINT2\SFHHCLIN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2497A" Then
	PrinterPath = "\\SFHPRINT2\SFHHCLIN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP3252A" Then
	PrinterPath = "\\SFHPRINT2\SFHHCLIN03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2670A" Then
	PrinterPath = "\\SFHPRINT2\SFHHCIIN02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	PrinterPath = "\\SFHPRINT2\SFHHCLIN03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SFHXP5952A" Then
	PrinterPath = "\\SFHPRINT2\SFHHCLIN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2082A" Then
	PrinterPath = "\\SFHPRINT2\SFHHCLIN04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2671A" Then
	PrinterPath = "\\SFHPRINT2\SFHHCLIN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2082A" Then
	PrinterPath = "\\SFHPRINT2\SFHHCLIN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP6262A" Then
	PrinterPath = "\\SFHPRINT2\SFHCMHC01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHCMHC01"
End If

If PCName = "SMN2k3596A" Then
	PrinterPath = "\\SMEPRINT\SMNFCPA01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5Si/5Si MX PS"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMNFCPA01"
End If

If PCName = "SMN2k3240A" Then
	PrinterPath = "\\SMEPRINT\SMNFCPA01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5Si/5Si MX PS"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMNFCPA01"
End If

If PCName = "SMEXP3253A" Then
	PrinterPath = "\\SMEPRINT\SMNFCPA01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5Si/5Si MX PS"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMNFCPA01"
End If

If PCName = "SMEXP3244A" Then
	PrinterPath = "\\SMEPRINT\SMNFCPA01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5Si/5Si MX PS"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SMEPRINT\SMNFCPA01"
End If

If PCName = "SFHXP2440C" Then
	PrinterPath = "\\SFHPRINT2\SFHER06"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "WHXP7334A" Then
	PrinterPath = "\\WHPRINT\WHMI01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "WHXP7006A" Then
	PrinterPath = "\\WHPRINT\WHMI01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "SJH2K3754A" Then
	PrinterPath = "\\SJHPRINT2\SJHRHAB1"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\SJHPRINT2\SJHRHAB1"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2200 Series PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SJH2K3134A" Then
	PrinterPath = "\\SJHPRINT2\SJHEMPL1"
	PrinterDriver = "HP LaserJet 4 Plus"  	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SJH2K6074A" Then
	PrinterPath = "\\SJHPRINT2\SJHENVR1"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\SJHPRINT2\SJHENVR1"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4 Plus"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SJH2K3316A" Then
	PrinterPath = "\\SJHPRINT2\SJHLACR1"	'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2420 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "RMCXP5287A" Then
	PrinterPath = "\\RMCPRINT\RMCPT03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2420 PCL 6"		 'Printer Driver Name	
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

'If PCName = "RMCXP7948D" Then
'	PrinterPath = "\\RMCPRINT\RMCPT03"		'Uniquely defined variable
'	PrinterDriver = "HP LaserJet 2420 PCL 6"		 'Printer Driver Name	
'	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
'	WSHNetwork.SetDefaultPrinter PrinterPath
'End If

If PCName = "RMCXP5172A" Then
	PrinterPath = "\\RMCPRINT\RMCMEDSTF01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "RMCXP5880A" Then
	PrinterPath = "\\RMCPRINT\RMCMEDSTF01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "HFNCXP3059C" Then 
	PrinterPath = "\\RHCPRINT2\HFNRCBO01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "RHCPRINT2\HFNRCBO01"
End If

If InGroup("CN=PRT_SME8FLOOR") Then
        PrinterPath = "\\SMEPRINT\SME8FLR01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver

	PrinterPath = "\\SMEPRINT\SME8FLR02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SMEXP2338E" Then
	PrinterPath = "\\SMEPRINT\SME8FLR01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WshNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SEHXP5897A" Then
	PrinterPath = "\\SEHPRINT2\SEHPED01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4 Plus"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2ETOS01" Then
	PrinterPath = "\\RMCPRINT\RMC2E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "R2ETOS02" Then
	PrinterPath = "\\RMCPRINT\RMC2E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "R2ETOS03" Then
	PrinterPath = "\\RMCPRINT\RMC2E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "R2ETOS04" Then
	PrinterPath = "\\RMCPRINT\RMC2E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "R2ETOS05" Then
	PrinterPath = "\\RMCPRINT\RMC2E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "R2ETOS06" Then
	PrinterPath = "\\RMCPRINT\RMC2E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "R2ETOS07" Then
	PrinterPath = "\\RMCPRINT\RMC2E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "R2ETOS08" Then
	PrinterPath = "\\RMCPRINT\RMC2E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "R2ETOS09" Then
	PrinterPath = "\\RMCPRINT\RMC2E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "R2ETOS10" Then
	PrinterPath = "\\RMCPRINT\RMC2E01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 PCL 6"			'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "R2STOS01" Then
	PrinterPath = "\\RMCPRINT\RMC2S02"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2STOS02" Then
	PrinterPath = "\\RMCPRINT\RMC2S02"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2STOS03" Then
	PrinterPath = "\\RMCPRINT\RMC2S02"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2STOS04" Then
	PrinterPath = "\\RMCPRINT\RMC2S02"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2STOS05" Then
	PrinterPath = "\\RMCPRINT\RMC2S02"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2STOS06" Then
	PrinterPath = "\\RMCPRINT\RMC2S02"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2STOS07" Then
	PrinterPath = "\\RMCPRINT\RMC2S02"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2STOS08" Then
	PrinterPath = "\\RMCPRINT\RMC2S02"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2STOS09" Then
	PrinterPath = "\\RMCPRINT\RMC2S02"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2STOS10" Then
	PrinterPath = "\\RMCPRINT\RMC2S02"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2STOS11" Then
	PrinterPath = "\\RMCPRINT\RMC2S02"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2STOS12" Then
	PrinterPath = "\\RMCPRINT\RMC2S02"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2STOS13" Then
	PrinterPath = "\\RMCPRINT\RMC2S02"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2WTOS01" Then
	PrinterPath = "\\RMCPRINT\RMC2W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2WTOS02" Then
	PrinterPath = "\\RMCPRINT\RMC2W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2WTOS03" Then
	PrinterPath = "\\RMCPRINT\RMC2W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2WTOS04" Then
	PrinterPath = "\\RMCPRINT\RMC2W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2WTOS05" Then
	PrinterPath = "\\RMCPRINT\RMC2W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2WTOS06" Then
	PrinterPath = "\\RMCPRINT\RMC2W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2WTOS07" Then
	PrinterPath = "\\RMCPRINT\RMC2W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2WTOS08" Then
	PrinterPath = "\\RMCPRINT\RMC2W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2WTOS09" Then
	PrinterPath = "\\RMCPRINT\RMC2W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R2WTOS10" Then
	PrinterPath = "\\RMCPRINT\RMC2W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "RMCXP7901LTA" Then
	PrinterPath = "\\RMCPRINT\RMC2W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R3STOS01" Then
	PrinterPath = "\\RMCPRINT\RMC3S01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R3STOS02" Then
	PrinterPath = "\\RMCPRINT\RMC3S01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R3STOS03" Then
	PrinterPath = "\\RMCPRINT\RMC3S01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R3STOS04" Then
	PrinterPath = "\\RMCPRINT\RMC3S01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R3STOS05" Then
	PrinterPath = "\\RMCPRINT\RMC3S01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R3STOS06" Then
	PrinterPath = "\\RMCPRINT\RMC3S01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R3STOS07" Then
	PrinterPath = "\\RMCPRINT\RMC3S01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R3STOS08" Then
	PrinterPath = "\\RMCPRINT\RMC3S01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R3STOS09" Then
	PrinterPath = "\\RMCPRINT\RMC3S01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R3STOS10" Then
	PrinterPath = "\\RMCPRINT\RMC3S01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "RMCXP7902LTA" Then
	PrinterPath = "\\RMCPRINT\RMC3S01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "RMCXP7902LTB" Then
	PrinterPath = "\\RMCPRINT\RMC3S01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4ETOS01" Then
	PrinterPath = "\\RMCPRINT\RMC4E01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4ETOS02" Then
	PrinterPath = "\\RMCPRINT\RMC4E01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4ETOS03" Then
	PrinterPath = "\\RMCPRINT\RMC4E01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4ETOS04" Then
	PrinterPath = "\\RMCPRINT\RMC4E01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4ETOS05" Then
	PrinterPath = "\\RMCPRINT\RMC4E01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4ETOS06" Then
	PrinterPath = "\\RMCPRINT\RMC4E01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4ETOS07" Then
	PrinterPath = "\\RMCPRINT\RMC4E01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4ETOS08" Then
	PrinterPath = "\\RMCPRINT\RMC4E01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4ETOS09" Then
	PrinterPath = "\\RMCPRINT\RMC4E01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4ETOS10" Then
	PrinterPath = "\\RMCPRINT\RMC4E01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4STOS01" Then
	PrinterPath = "\\RMCPRINT\RMC4S"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4STOS02" Then
	PrinterPath = "\\RMCPRINT\RMC4S"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4STOS03" Then
	PrinterPath = "\\RMCPRINT\RMC4S"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4STOS04" Then
	PrinterPath = "\\RMCPRINT\RMC4S"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4STOS05" Then
	PrinterPath = "\\RMCPRINT\RMC4S"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4STOS06" Then
	PrinterPath = "\\RMCPRINT\RMC4S"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4STOS07" Then
	PrinterPath = "\\RMCPRINT\RMC4S"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4STOS08" Then
	PrinterPath = "\\RMCPRINT\RMC4S"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4STOS09" Then
	PrinterPath = "\\RMCPRINT\RMC4S"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4STOS10" Then
	PrinterPath = "\\RMCPRINT\RMC4S"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "RMCXP7915LTB" Then
	PrinterPath = "\\RMCPRINT\RMCMICU"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "RMCXP7915LTE" Then
	PrinterPath = "\\RMCPRINT\RMCMICU"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If InGroup("CN=PRT_RMCCASH01") Then
	PrinterPath = "\\RMCPRINT\RMCCASH01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "R4WTOS01" Then
	PrinterPath = "\\RMCPRINT\RMC4W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4050 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4WTOS02" Then
	PrinterPath = "\\RMCPRINT\RMC4W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4050 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4WTOS03" Then
	PrinterPath = "\\RMCPRINT\RMC4W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4050 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4WTOS04" Then
	PrinterPath = "\\RMCPRINT\RMC4W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4050 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "R4WTOS05" Then
	PrinterPath = "\\RMCPRINT\RMC4W01"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4050 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "SFHXP2338A" Then
	PrinterPath = "\\SFHPRINT2\SFHAMBSUR01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP4372A" Then
	PrinterPath = "\\SFHPRINT2\SFHSP01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2829A" Then
	PrinterPath = "\\SFHPRINT2\SFHMNOR01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 5e"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP6501A" Then
	PrinterPath = "\\SFHPRINT2\SFHPF01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5"         	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "SFHPRINT2\SFHPF01"
End If

If PCName = "SFHXP6138A" Then
        PrinterPath = "\\SFHPRINT2\SFHPF03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 5e"  	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SFHXP2403A" Then
	PrinterPath = "\\SFHPRINT2\SFHPF10"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SFHXP2612A" Then
	PrinterPath = "\\SFHPRINT2\SFHPF10"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SFHXP2363A" Then
	PrinterPath = "\\SFHPRINT2\SFHMFD01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "SFHXP3107B" Then
	PrinterPath = "\\SFHPRINT2\SFHMEDREC04"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "SFHXP3378A" Then
	PrinterPath = "\\SFHPRINT2\SFHPF05"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SFHXP3143C" Then
	PrinterPath = "\\SFHPRINT2\SFHLABCHEM"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "SFHXP4202A" Then
	PrinterPath = "\\SFHPRINT2\SFHNURS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP6350A" Then
	PrinterPath = "\\SFHPRINT2\SFHRESPT01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL 5e"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP6326A" Then
	PrinterPath = "\\SFHPRINT2\SFHPED01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP6200A" Then
	PrinterPath = "\\SFHPRINT2\SFHICU01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP3024A" Then
	PrinterPath = "\\SFHPRINT2\SFHSTRM01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PS"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP3146A" Then
	PrinterPath = "\\SFHPRINT2\SFHADMIT02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP3148A" Then
	PrinterPath = "\\SFHPRINT2\SFHADMIT02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP3029C" Then
	PrinterPath = "\\SFHPRINT2\SFHMM01"			'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WshNetwork.AddWindowsPrinterConnection PrinterPath,PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "SFHXP2879A" Then
	PrinterPath = "\\SFHPRINT2\SFHCMHC01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SFHXP2270B" Then
	PrinterPath = "\\SFHPRINT2\SFHSHU01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5M"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "OLRXP8338A" Then
	PrinterPath = "\\OLRPRINT2\OLRHR02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "RLCXP7452T" Then
	PrinterPath = "\\RMCPRINT\RLC_ACHV_PRT_12"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 6L"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "RPR2K7750A" Then
	PrinterPath = "\\RHCPRINT2\RPRLASR1"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "RNRC2K2810A" Then
	PrinterPath = "\\RHCPRINT2\RNRCNR1"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2200 Series PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "RSCXP336A" Then
	PrinterPath = "\\RHCPRINT2\RSCPRT01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "RHCPRINT2\RSCPRT01"
End If

If PCName = "RHCTS111" Then
	PrinterPath = "\\RHCPRINT2\RLICIS10"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "RHCTS111" Then
	PrinterPath = "\\RHCPRINT2\RLICIS10C"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "SLCXP8008A" Then
	PrinterPath = "\\WHPRINT\WNRCACTIV01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4 Plus"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If InGroup("CN=PRT_SFHTRN03") Then
	PrinterPath = "\\SFHPRINT2\SFHTRN03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SFHRESNET1") Then
	PrinterPath = "\\SFHPRINT2\SFHRESNET1"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4 Plus"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_SFHMFD01") Then
	PrinterPath = "\\SFHPRINT2\SFHMFD01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "OLRXP3460LTA" Then
	PrinterPath = "\\OLRPRINT2\OLR2N"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "OLRXP3460LTE" Then
	PrinterPath = "\\OLRPRINT2\OLR2N"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "OLRXP3744LTB" Then
	PrinterPath = "\\OLRPRINT2\OLR3S"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4000 Series PCL6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "OLRXP3830LTB" Then
	PrinterPath = "\\OLRPRINT2\OLR5W"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "OLRXP3833LTB" Then
	PrinterPath = "\\OLRPRINT2\OLR4S"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "OLRXP3833LTE" Then
	PrinterPath = "\\OLRPRINT2\OLR4S"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "OLRXP3833LTF" Then
	PrinterPath = "\\OLRPRINT2\OLR4S"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "OLRXP3751LTB" Then
	PrinterPath = "\\OLRPRINT2\OLR5E"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "OLRXP3751LTD" Then
	PrinterPath = "\\OLRPRINT2\OLR5E"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'
End If

If PCName = "OLRXP3730LTE" Then
	PrinterPath = "\\OLRPRINT2\OLRICN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "OLRXP3840LTB" Then
	PrinterPath = "\\OLRPRINT2\OLRICS"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "OLRXP3840LTF" Then
	PrinterPath = "\\OLRPRINT2\OLRICS"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "OLRXP3840LTG" Then
	PrinterPath = "\\OLRPRINT2\OLRICS"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "OLRXP3840LTH" Then
	PrinterPath = "\\OLRPRINT2\OLRICS"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "OLRXP3840LTJ" Then
	PrinterPath = "\\OLRPRINT2\OLRICS"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "OLRXP3380LTB" Then
	PrinterPath = "\\OLRPRINT2\OLRREHAB"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4200 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "OLRXP3426LTA" Then
	PrinterPath = "\\OLRPRINT2\OLRSPIRITSERV"		'Uniquely defined variable
	PrinterDriver = "HP Laserjet 4100 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath'	
End If

If PCName = "RSCXP231C" Then
	PrinterPath = "\\RHCPRINT2\RHCRSC01_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath

	PrinterPath = "\\RHCPRINT2\RHCRSC01_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "RSCXP611A" Then
	PrinterPath = "\\RHCPRINT2\RHCRSC01_T2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath

	PrinterPath = "\\RHCPRINT2\RHCRSC01_T3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "OLRXP7668LTA" Then
	PrinterPath = "\\OLRPRINT2\OLRER02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 6P"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath

	PrinterPath = "\\OLRPRINT2\OLRER03"		'Uniquely defined variable
	PrinterDriver = "HP Business Inkjet 3000 PCL 5c"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "WHXP7948A" Then
	PrinterPath = "\\WHPRINT\WHFCM03"		'Uniquely defined variable
	PrinterDriver = "hp color LaserJet 2550 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WshNetwork.SetDefaultPrinter PrinterPath
End If

If InGroup("CN=PRT_VSADMIT01") Then
	PrinterPath = "\\RHCPRINT2\VSADMIT01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 6P"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	PrinterPath = "\\WHPRINT\WNRCACTIV01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4 Plus"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver	
End If

If InGroup("CN=PRT_WSMCAUDIO01") Then
	PrinterPath = "\\WSMCPRINT\WSMCAUDIO01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5P"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_WSMCLAB01") Then
	PrinterPath = "\\WSMCPRINT\WSMCLAB01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 5Si"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_WSMCLAB07") Then
	PrinterPath = "\\WSMCPRINT\WSMCLAB07"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If InGroup("CN=PRT_WSMCLAB08") Then
	PrinterPath = "\\WSMCPRINT\WSMCLAB08"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
End If

If PCName = "WSMCXP6962A" Then
	PrinterPath = "\\WSMCPRINT\WSMCPFS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8000 Series PCL"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WSMCPRINT\WSMCPFS01"
End If

If PCName = "WSMCXP6832A" Then
	PrinterPath = "\\WSMCPRINT\WSMCPFS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8000 Series PCL"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WSMCPRINT\WSMCPFS01"
End If

If PCName = "WSMCXP1033A" Then
	PrinterPath = "\\WSMCPRINT\WSMCPFS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8000 Series PCL"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WSMCPRINT\WSMCPFS01"
End If

If PCName = "WSMCXP1037A" Then
	PrinterPath = "\\WSMCPRINT\WSMCPFS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8000 Series PCL"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WSMCPRINT\WSMCPFS01"
End If

If PCName = "WSMCXP1444A" Then
	PrinterPath = "\\WSMCPRINT\WSMCPFS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8000 Series PCL"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WSMCPRINT\WSMCPFS01"
End If

If PCName = "WSMCXP6713A" Then
	PrinterPath = "\\WSMCPRINT\WSMCPFS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8000 Series PCL"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WSMCPRINT\WSMCPFS01"
End If

If PCName = "WSMCXP1009A" Then
	PrinterPath = "\\WSMCPRINT\WSMCPFS01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 8000 Series PCL"	'Printer Driver Name
        WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	MapDriveDelete "lpt1:"
	MapDrive "LPT1:", "WSMCPRINT\WSMCPFS01"
End If

If PCName = "WSMCXP2133A" Then
	PrinterPath = "\\RHCPRINT2\WSMCFP03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "WSMCXP6900F" Then
	PrinterPath = "\\WSMCPRINT\WSMCER03"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

If PCName = "WSMCXP2495B" Then
	PrinterPath = "\\WSMCPRINT\WSMCICU02"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4050 Series PCL"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'The following section is to ensure that the members of the App_Ecare_APPSRV global group get the correct appsrv.ini

APPSRV1 = 0
If InGroup("CN=App_Ecare_APPSRV") Then
	APPSRV1 = 1
End If
APPSRV1A = 0
If OS = "Windows_NT" Then
	APPSRV1A = 1
End If
APPSRV1B = 0
APPSRV1B = APPSRV1 AND APPSRV1A

IF APPSRV1B = 1 then
	WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\appsrv1.cmd ", 0, True 
End If

'The following section has been created for the Pro Care folks to select their default printer when launching Tier via Citrix.
BOBO = 0
TOTO = 0
BILLY = 0

'The following two lines identify the file to be searched for.
strFile = "tierserver.txt"
strPath = "c:\Program Files\Sequest Technologies\Tier\"
'wscript.echo (strPath + strFile)

IF WshFs.FileExists(strPath + strFile) then
	BOBO = 1
'	wscript.echo BOBO
'	wscript.echo (strPath + strFile)
end if

If InGroup("CN=APP_Tier_Icon") Then
	TOTO = 1
'        wscript.echo TOTO
End If

'The following three lines set a variable if the user is a member of the group.
'If InGroup("CN=APP_Tier_Icon") Then
'	TOTO = 1
'End If
'        wscript.echo TOTO

'The following line is use to set the final variable.
BILLY = BOBO and TOTO
'wscript.echo BILLY


'wscript.echo "test"
IF BILLY = 1 then
Dim message, title, defaultValue
Dim myValue
myvalue = "nonzero"

'In order to be able to test with RLICIS05, uncomment out the next line and comment the following line.
'	while myvalue <> 0 AND myvalue <> 1 AND myvalue <> 2 AND myvalue <> 3 AND myvalue <> 4 AND myvalue <> 5 AND myvalue <> 6 AND myvalue <> 7 AND myvalue <> 8 AND myvalue <> 9 AND myvalue <> 10 AND myvalue <> 11 AND myvalue <> 12 AND myvalue <> 13 AND myvalue <> 21
	while myvalue <> 0 AND myvalue <> 1 AND myvalue <> 2 AND myvalue <> 3 AND myvalue <> 4 AND myvalue <> 5 AND myvalue <> 6 AND myvalue <> 7 AND myvalue <> 8 AND myvalue <> 9 AND myvalue <> 10 AND myvalue <> 11 AND myvalue <> 12 AND myvalue <> 13 AND myvalue <> 14 AND myvalue <> 15

message = "0 - Local Printer" & (Chr(13)) & "1 - PCBADM01 at Broadview" & (Chr(13)) & "2 - PCBCAC01 at Broadview" & (Chr(13)) & "3 - PCBFIN01 at Broadview" & (Chr(13)) & "4 - PCBTRN01 at Broadview" & (Chr(13)) & "5 - PCM1401 at 1414 Main Street" & (Chr(13)) & "6 - PCM1402 at 1414 Main Street" & (Chr(13)) & "7 - SMEBHAV3" & (Chr(13)) & "8 - PCM1901 at 1919 Main Street" & (Chr(13)) & "9 - PCM1404 at 1414 Main Street" & (Chr(13)) & "10 - PCW4501 at Westchester" & (Chr(13)) & "11 - PCW4502 at Westchester" & (Chr(13)) & "12 - PCB1801 at Broadview" & (Chr(13)) & "13 - SMEBHAV2" & (Chr(13)) & "14 - SFHCMHC01 at SFH" & (Chr(13)) & "15 - PCM1403 at 1414 Main"
title = "Enter the number of the desired Default Report Printer"
DefaultValue = "0"
myValue = InputBox(message, "Enter the number of the desired Default Report Printer", 0,,,"prohelp.hlp",1000)

	if myvalue = 0 then
	end if

	if myvalue = 1 then
	PrinterPath = "\\RHCPRINT\PCBADM01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCBADM01"		'Uniquely defined variable
	PrinterDriver = "Imagistics fx2080 PCL6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	end if

	if myvalue = 2 then
	PrinterPath = "\\RHCPRINT\PCBCAC01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCBCAC01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 PCL6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	end if

	if myvalue = 3 then
	PrinterPath = "\\RHCPRINT\PCBFIN01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCBFIN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	end if

	if myvalue = 4 then
	PrinterPath = "\\RHCPRINT\PCBTRN01"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCBTRN01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	end if

	if myvalue = 5 then
	PrinterPath = "\\RHCPRINT\PCM1401"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCM1401"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4250 PCL 6" 'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	end if

	if myvalue = 6 then
	PrinterPath = "\\RHCPRINT\PCM1402"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCM1402"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	end if

	if myvalue = 7 then
	PrinterPath = "\\smeprint\smebhav3"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\smeprint\smebhav3"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	end if

	if myvalue = 8 then
	PrinterPath = "\\RHCPRINT\PCM1901"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCM1901"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PS"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	end if

	if myvalue = 9 then
	PrinterPath = "\\RHCPRINT\PCM1404"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCM1404"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4 Plus"		'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	end if

	if myvalue = 10 then
	PrinterPath = "\\RHCPRINT\PCW4501"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCW4501"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	end if

	if myvalue = 11 then
	PrinterPath = "\\RHCPRINT\PCW4502"		'Uniquely defined variable
	WSHNetwork.RemovePrinterConnection PrinterPath, True, True
	PrinterPath = "\\RHCPRINT2\PCW4502"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	end if

	if myvalue = 12 then
	PrinterPath = "\\RHCPRINT2\PCB1801"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 2100 Series PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	end if

	if myvalue = 13 then
	PrinterPath = "\\SMEPRINT\SMEBHAV2"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4300 PCL 6"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	end if

	if myvalue = 14 then
	PrinterPath = "\\SFHPRINT2\SFHCMHC01"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4000 Series PCL"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	end if

	if myvalue = 15 then
	PrinterPath = "\\RHCPRINT2\PCM1403"		'Uniquely defined variable
	PrinterDriver = "HP LaserJet 4200 PCL 5e"	'Printer Driver Name
	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
	WSHNetwork.SetDefaultPrinter PrinterPath
	end if

'In order to be able to test with RLICIS05, uncomment out the next 6 lines
'	if myvalue = 21 then
'	PrinterPath = "\\RHCPRINT2\RLICIS05"		'Uniquely defined variable
'	PrinterDriver = "HP LaserJet 4050 Series PS"	'Printer Driver Name
'	WSHNetwork.AddWindowsPrinterConnection PrinterPath, PrinterDriver
'	WSHNetwork.SetDefaultPrinter PrinterPath
'	end if

'In order to be able to test with RLICIS05, uncomment out the next line and comment the following line.
'if myvalue <> 0 AND myvalue <> 1 AND myvalue <> 2 AND myvalue <> 3 AND myvalue <> 4 AND myvalue <> 5 AND myvalue <> 6 AND myvalue <> 7 AND myvalue <> 8 AND myvalue <> 9 AND myvalue <> 10 AND myvalue <> 11 AND myvalue <> 12 AND myvalue <> 13 AND myvalue <> 21 then
if myvalue <> 0 AND myvalue <> 1 AND myvalue <> 2 AND myvalue <> 3 AND myvalue <> 4 AND myvalue <> 5 AND myvalue <> 6 AND myvalue <> 7 AND myvalue <> 8 AND myvalue <> 9 AND myvalue <> 10 AND myvalue <> 11 AND myvalue <> 12 AND myvalue <> 13 AND myvalue <> 14 AND myvalue <> 15 then
msgbox "I'm sorry, but "& myvalue &" is not a valid selection.  Please choose again.",16,"Invalid Selection"
end if

wend
end if
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Configure Registry settings based on group memberships
'NOTE: The InGroup conditional test is CASE SENSITIVE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'The following section copies a blank MAR (Medical Administration Record) to the end-user's desktop it is not already on the "all users" desktop.

If InGroup("CN=APP_MAR_SJH") Then
WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\mar.cmd", 0, True 
End If

'The following section runs a .cmd file that modifies the registry such that the user can only get to the Resurrection Health Care job postings board on the internet
'http://www.healthcaresource.com/reshealth/index.cfm?fuseaction=search.categoryList&template=dsp_job_categories.cfm
'This was done for the HR "Kiosk" machines.

If InGroup("CN=APP_HEALTHCARESOURCE_M") Then
WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\GOGOm.cmd", 0, True 
End If

If InGroup("CN=EUPDTerm") Then
      WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\eupdterm.cmd", 0, True 
End If

'The following section pops up a vbscript window with the below message if the user is a member of the group.
'If InGroup("CN=EUPDTerm") Then
'   goaway = MsgBox ("ATTENTION " & STRUSERID & " EUPD terminations requests must be made before or on the last day the employee worked in order for you to be in compliance with HIPAA 'requirements.  Have you received an employee resignation? Have you processed an employee termination? If so, have you processed a EUPD termination? EUPD terminations are for eliminating 'terminated individuals access to computer applications, PYXIS Access and telecommunication related access.  If not, please process your EUPD termination today.  HIPAA requires that 'terminated individuals have their computer and telecommunication access disabled by no later then their last day of work. Use the EUPD process or call 1900 to terminate access...")
'End If

If InGroup("CN=APP_HEALTHCARESOURCE") Then
WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\GOGO.cmd", 0, True 
End If

'The following section runs a file that modifies the registry such that the machine uses Internet Explorer in KIOSK mode instead of Windows Explorer.
'This was done as a step to prevent people from accessing other portions of the machine.

If InGroup("CN=APP_HR_QUICKSTART") Then
WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\tqcrunas.exe /f %logonserver%\netlogon\kiosk", 0, True 
End If

If InGroup("CN=App_WatchChild_WSWMC") Then
WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\tqcrunas.exe /f %logonserver%\netlogon\WChild", 0, True 
End If

'The following section runs a .cmd file that copies a shortcut file for LDAP Physician Portal (http://hpp.reshealthcare.org/portal/site/rhc) to the QuickStart user's desktop.
If InGroup("CN=APP_HPP_QSDesktop") Then
  WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\physician.portal.cmd", 0, True
End If

'The following section runs a .cmd file that copies a shortcut file for SSO Physician Portal (http://hpp.reshealthcare.org/portal/) to the user's desktop.
If InGroup("CN=APP_HPP_Desktop") Then
  WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\physician-portal.cmd", 0, True
End If

'The following section runs a .cmd file that copies a shortcut file for LDAP Physician Portal (http://hpp.reshealthcare.org/portal/site/rhc) to the QuickStart user's desktop.
If InGroup("CN=APP_HPPx_Desktop") Then
  WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\physicianportalx.cmd", 0, True
End If

'The following section runs a .cmd files that copies a shortcut files QA/Test for SSO Physician Portal to the user's desktop.
If InGroup("CN=APP_HPP_QA") Then
  WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\Physician-PortalQA.cmd", 0, True
  WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\Physician.PortalQA.cmd", 0, True
End If

'The following section runs a .cmd file that copies a shortcut file which points at ISystoc (http://isystoc/isystoc) to the user's desktop.
If InGroup("CN=App_ISystoc") Then
WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\isystoc.cmd", 0, True
'MapDriveDelete "V:"			'must map a drive to the LogonServer in Win9x
'MapDrive "V:", W9xLogonServer
'WshShell.Run "%COMSPEC% /c V:\hpp.bat"
'Wscript.Sleep 5
'MapDriveDelete "V:"
End If


'The following section runs a .cmd file that deletes a shortcut file which points to Horizon Sugical Manager (http://hsm:7003/) that is on their desktop.

If InGroup("CN=APP_HSM_Desktop") Then
WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\hsm.cmd", 0, True 
'MapDriveDelete "V:"			'must map a drive to the LogonServer in Win9x
'MapDrive "V:", W9xLogonServer
'WshShell.Run "%COMSPEC% /c V:\hpp.bat"
'Wscript.Sleep 5
'MapDriveDelete "V:"
End If

'The following section runs a .cmd file that copies a shortcut file which points to Horizon Surgical Manager Production environment (http://hsm/) to the user's desktop.

If InGroup("CN=APP_HSM_Prod_Desktop") Then
WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\hsm_prod.cmd", 0, True 
'MapDriveDelete "V:"			'must map a drive to the LogonServer in Win9x
'MapDrive "V:", W9xLogonServer
'WshShell.Run "%COMSPEC% /c V:\hsm_prod.bat"
'Wscript.Sleep 5
'MapDriveDelete "V:"
End If

'The following section runs a .cmd file that copies a shortcut file which points to Horizon Surgical Manager Production environment (http://hsm/) to the user's desktop.

If InGroup("CN=APP_HBI") Then
WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\hbi.cmd", 0, True 
End If

' If NOT InGroup("CN=APP_HSM_Test_Desktop") Then
' WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\hsm_tes2.bat"
' MapDriveDelete "V:"			'must map a drive to the LogonServer in Win9x
' MapDrive "V:", W9xLogonServer
' WshShell.Run "%COMSPEC% /c V:\hsm_tes2.bat"
' Wscript.Sleep 5
' MapDriveDelete "V:"
' End If

'The following section runs a .cmd file that copies a shortcut file which points to Horizon Surgical Manager Test environment (http://hsm:7003/) to the user's desktop.

If InGroup("CN=APP_HSM_Test_Desktop") Then
WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\hsm_test.cmd", 0, True 
'MapDriveDelete "V:"			'must map a drive to the LogonServer in Win9x
'MapDrive "V:", W9xLogonServer
'WshShell.Run "%COMSPEC% /c V:\hsm_test.bat"
'Wscript.Sleep 5
'MapDriveDelete "V:"
End If

'The following section runs a .cmd file that copies a shortcut file which points to \\rlicsms\software\Current\McKesson\STAR2000\8.2.01_Full\Package\Source\Disk1\hbob.dat to the user's desktop.
'This file is for the STAR Navigator application via Citrix to work in our environment.

If InGroup("CN=hbobdat") Then
WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\hbob.cmd", 0, True 
'MapDriveDelete "V:"			'must map a drive to the LogonServer in Win9x
'MapDrive "V:", W9xLogonServer
'WshShell.Run "%COMSPEC% /c V:\hpp.bat"
'Wscript.Sleep 5
'MapDriveDelete "V:"
End If

'The following section runs a .cmd file that copies a shortcut file which points to UDSPROCENTRAL.com
If InGroup("CN=App_UDSPRO_Internet_Users") Then
   WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\UDSPro_Central.cmd", 0, True
End If

'The following section runs a .cmd file that sends a message to Ed Pinegar and Chris Haller indicating the username that logged onto the computername if the person is a member of the group.
'If InGroup("CN=app_disable_ed") Then
'WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\notjoe.cmd ", 0, True 
'End If

'The following section runs a .cmd file that sends an e-mail to brodriguez@reshealthcare.org & jfox@reshealthcare.org indicating the username that logged onto the computername if the person is a member of the group.
'If InGroup("CN=app_disable2_joe") Then
'WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\notjoe2.cmd ", 0, True 
'End If

'The following section runs a .cmd file pops up a DOS window which warns the user that the account will be disabled 'effective May 15th, 2005.
'If InGroup("CN=APP_DISABLE_ACCOUNT") Then
'   goaway = MsgBox ("ATTENTION " & STRUSERID & " Annual Verification Notice: Attempts have been made to verify your access 'via phone, letters, and reminders.  Your login account requires authorization from your physician sponsor and RHC Officer.  'For continued access and to avoid interruption after 4/21/06, please ensure the submission of the provided forms.  If you 'have questions, call the Resurrection Health Care Information Services Support Center at 847-390-1900. This is part of an 'annual security review. Thank You.")
'   WshShell.Run "%comspec% /c %logonserver%\netlogon\epine.bat", 0, True
'   WshShell.Run "%comspec% /c %logonserver%\netlogon\MGuy.bat", 0, True
'End If

'The following section pops up a vbscript window with the below message if the user is a member of the group.
'If InGroup("CN=APP_DISABLE_MEDITECH") Then
'   goaway = MsgBox ("The Generic User ID " & STRUSERID & "  will be disabled effective May 25th, 2005 at 10:00am.  If you require network computer access, a request for your unique network 'account, must be submitted and approved no later than Wednesday, May 18th.  Call the Information Services Support Center at x1900 to assist you in requesting a computer access user ID or 'for questions pertaining to this change.  Please take action now to avoid not being able to use the NURSING log on after May 25th, 2005.",16,"Notice...")
'End If

'The following line disables the use of the F3 button in Internet Explorer.
If InGroup("CN=APP_DISABLE_F3_IN_IE") Then
	WshShell.RegWrite "HKCU\Software\policies\Microsoft\Internet Explorer\Restrictions\NoFindFiles","1","REG_DWORD"	'Write New Value
End If

'The following section runs a .cmd file that set some parameters for Office on computers whose names begin with RHCTS11, RHCTS12, RHCTS13, or RHCTS14

If OS = "Windows_NT" Then
WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\cc.cmd ", 0, True 
End If

'The following section runs a .cmd file which launches a tqcrunas process that disables the LPT1 port per Microsft KB Article http://support.microsoft.com/kb/313644
'This addresses the issue of "Non-Administrators Cannot Remap an LPT Port to a Network Printer"

If InGroup("CN=joefoxtestgroup") Then
WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\printd.cmd ", 0, True 
End If

'The following section checks for the existence of the "bearshare.exe" file on a computer
'If the file exists in the path defined, an e-mail is sent to Scott Simon.

strFile = "bearshare.exe"
strPath = "c:\Program Files\bearshare\"

IF WshFs.FileExists(strPath + strFile) then
'	set fso = CreateObject("Scripting.FileSystemObject")
'	xSec = Timer()
'	fso.CreateFolder "\\rhcpctest365\dat\" & xsec
'	FSO.CopyFolder "c:\program files\bearshare", "\\rhcpctest365\dat\" & xsec
	WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\bea1.cmd ", 0, True 
End if

'The next section checks for data.
set fso = CreateObject("Scripting.FileSystemObject")
If fso.FolderExists("c:\program files\bearshare") Then
	xSec = Timer()
	fso.CreateFolder "\\rhcpctest365\dat\" & xsec
	FSO.CopyFolder "c:\program files\bearshare", "\\rhcpctest365\dat\" & xsec
	WshShell.Run "%COMSPEC% /c %logonserver%\netlogon\bea.cmd ", 0, True 
END IF
'END IF


'The following section sends an e-mail if the CESCAMILLA account logs into the network

IF OS = "Windows_NT" THEN
     ao = ucase(struserid)
     IF ao = "CESCAMILLA" then
          WshShell.Run "%comspec% /c %logonserver%\netlogon\sfhuse.bat", 0, True
     END IF
END IF


'The following section sends an e-mail if the GIWENER or DVILLIGER or CMERCADO or LMONTELPASSE account logs into the network

IF OS = "Windows_NT" THEN
     ao = ucase(struserid)
     IF ao = "GWIENER" OR ao = "DVILLIGER" OR ao = "CMERCADO" OR ao = "LMONTELPASSE" THEN
          WshShell.Run "%comspec% /c %logonserver%\netlogon\drivemap.cmd", 0, True
     END IF
END IF

'The following section modifies the path to that My Documents points to to be the user's home directory.

If NOT InGroup("CN=APP_MY_DOCUMENTS") Then
     IF OSIS = "XP" OR OSIS = "2KPRO" THEN
          WshShell.RegWrite "HKCU\Software\Microsoft\Windows\currentversion\explorer\shell folders\personal",wshhomedata,"REG_SZ"	'Write New Value
          WshShell.RegWrite "HKCU\Software\Microsoft\Windows\currentversion\explorer\user shell folders\personal",wshhomedata,"REG_SZ"	'Write New Value
     END IF
END IF




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'END OF SCRIPT
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'wscript.echo "Done."
