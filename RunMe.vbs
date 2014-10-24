Option Explicit

'Declarations
Dim ObjShell
Dim OWSH
Dim WalletAdr 'Self Explanitory
Dim StrFileName
Dim Filesys
Dim objFSO
Dim ObjFSO2 ' Avoiding redefinition
Dim Appdata2 ' Appdata is s system variable so it doesn't like anything else /sad
Dim strfolderpath2 'Documentation check path
Dim ObjFile 
Dim wshSystemEnv
Dim i ' File check variable
Dim Arch
Dim ArchPath
Dim Windir2
Dim OCLPath 'Where is OpenCL Located?
Dim OCL
Dim R 'Value entered for menu
Dim Q 'Quit variable
Dim R2 ' Second Menu Variable
Dim strMenu
Dim rc
Dim Cuda
Dim MSVCRx86
Dim CudaPath
Dim MSVCRx86path
Dim Electrumpresent
dim objWMIService 
Dim objItem 
Dim colItems 
Dim strComputer
Dim compModel
Dim strKeyPath
Dim strValueName
Dim strValue
Dim oReg
Dim AMD
Dim Threads
Dim CPU
Dim ThreadsPath
Dim CPUPath
Dim CPUMine
dim rc2
dim intanswer
dim objFileToWrite
dim S ' Yet another sleep variable
dim objFileToRead
dim strnumcpu 'number of host threads
dim intnumcpu
dim finalnumcpu
dim hidden
dim hiddenpath
dim inthidden
dim R3
dim SkeinPool
dim QubitPool
dim GroestlPool
dim scryptpool
dim ScryptPath
dim groestlpath
dim skeinpath
dim qubitpath
dim Strfolderpath3
dim User2
dim credentials
dim xmg
dim credentialspath
dim xmgpool
dim xmgpath
Set Filesys = CreateObject("Scripting.FileSystemObject") 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objShell = CreateObject("Wscript.Shell")
Set wshSystemEnv = objShell.Environment("SYSTEM")
dim UseAVX
dim AVXPath

objShell.CurrentDirectory = ".\Scripts" ' Change working directory 

ThreadsPath = ".\Threads.conf"
CPUPath = ".\CPUMine.conf"
HiddenPath = ".\Hidden.conf"
CredentialsPath = ".\credentials.conf"
AVXPath = ".\credentials.conf"
XMGPath = ".\xmg.conf"

'Begin Documentation Run Check
'Appdata2=objShell.ExpandEnvironmentStrings("%Appdata%")
'Strfolderpath2 = Appdata2 & "\Electrum-MYR" 
'User2=objShell.ExpandEnvironmentStrings("%USERPROFILE%")
'Strfolderpath3 = User2 & "\Desktop" 
'Set objShell = CreateObject("Wscript.Shell")
' If Not objFSO.FileExists(User2 & "\desktop" & "\Electrum-MyrWallet.exe") then
 'objFSO.CopyFile "Electrum-MyrWallet.exe", (User2 & "\desktop\")
'End If
 'If Not objFSO.FolderExists(strfolderpath2) then
  'Electrumpresent =0
  'CreateObject("WScript.Shell").Run "http://myriadplatform.org/mining-setup/"
  'CreateObject("WScript.Shell").Run ".\Electrum-MyrWallet.exe"
'One time only, Set Threads VAR to 1 less core than host system maximum to try to keep system fluidity
 'End If

'set objShell = CreateObject("Wscript.Shell")
 'If objFSO.FolderExists(strfolderpath2) then
  'Electrumpresent =1
'End if
'End Documentation Check

'Are we on X86 or X64?
ArchPath= "C:\Program Files (X86)"
If objFSO.FolderExists(ArchPath) then
Arch="X64"
End If

If Not objFSO.FolderExists(ArchPath) then
Arch="X86"
End If

'Open CL Check
'Windir2=objShell.ExpandEnvironmentStrings("%WinDir%")
'OCLPath= Windir2 & "\System32\OpenCL.DLL"
'If objFSO.FileExists(OCLPath) then
'OCL=1
'End If
'If Not objFSO.fileExists(OCLPath) then
'OCL=0
'End If

'Nvidia Check
'Windir2=objShell.ExpandEnvironmentStrings("%WinDir%")
'CudaPath= Windir2 & "\System32\NVCuda.DLL"
'If objFSO.FileExists(CudaPath) then
'Cuda=1
'End If
'If Not objFSO.fileExists(CudaPath) then
'Cuda=0
'End If

'Visual Studio 2010 runtime check x86

'Windir2=objShell.ExpandEnvironmentStrings("%WinDir%")
'MSVCRx86Path= Windir2 & "\System32\MSVCR100.DLL"
'If objFSO.FileExists(MSVCRx86path) then
'MSVCRx86=1
'Else MSVCRx86=0
'End If

' AMD or Intel Check
const HKEY_LOCAL_MACHINE = &H80000002
strKeyPath = "SYSTEM\ControlSet001\Services\intelppm"
strValueName = "Start"
strValue = "1"
strComputer = "."
' WMI connection to Root CIM and get the computer type
set objWMIService = GetObject("winmgmts:\\" _
& strComputer & "\root\cimv2")
set colItems = objWMIService.ExecQuery(_
"Select Manufacturer from Win32_Processor")
'Loop through the results and store the type in compModel
for each objItem in colItems
compModel = objItem.Manufacturer
next
'Get a registry object
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
'Check the computer type. If the processor is an Intel, then re-enable the driver
if compModel = "GenuineIntel" then
AMD=0
else
AMD=1
end if


'Did we create the folder but not the wallet aka did our setup get borked, WE MAKE DAMN SURE THE WALLET EXISTS
'If Not objFSO.fileExists(strfolderpath2 & "\wallets\default_wallet") then
'i=0
'elseif objFSO.fileExists(strfolderpath2 & "\wallets\default_wallet") then
'i=1
'end if
'If i=0 and Electrumpresent=1 then
'CreateObject("WScript.Shell").Run "Electrum-MyrWallet.exe"
'end if

'Sleep loop
'Do until i=1
'If Not objFSO.fileExists(strfolderpath2 & "\wallets\default_wallet") then
'i=0
'elseif objFSO.fileExists(strfolderpath2 & "\wallets\default_wallet") then
'i=1
'end if
'loop

'Parse Wallet with a horrendously coded hack
'If i=1 then
'StrFileName = Appdata2 & "\Electrum-MYR\wallets\default_wallet"
'Set objFile = objFSO.OpenTextFile(StrFileName) 
'Set ObjFso2 = CreateObject("Scripting.FileSystemObject")
'ObjFile.Skip(19)  
'WalletADR = ObjFile.Read(34)
'End If

'Menu code
strMenu="Please type Enter your selection (1-2)" & VbCrLf &_
 "1 Change Credentials - DO ON FIRST RUN" & VbCrLf &_
 "2 Advanced (Opens poweruser options)" & VbCrLf &_
 "3 Change Pool - Default is NoncePool" & VbCrLf &_
 "4 Mine!" 
 rc=InputBox(strMenu,"Menu",1)
 If IsNumeric(rc) Then
     Select Case rc
         Case 1 
			 R=7
			 S=1
         Case 2
			 R=6
			 Case 3 
			 R2=6
			 Case 4
			 R=1
     End Select
 End If

If R=7 Then

Credentials = InputBox("-u poolworker -p worker password", _
    "e.g. -u aceoyame.test -p password")
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(CredentialsPath,2,true)
objFileToWrite.WriteLine(Credentials)
objFileToWrite.Close
End if



'Poweruser Menu code

If R=6 Then

strMenu="Please type Enter your selection (1-5)" & VbCrLf &_
 "1 Number of CPU Threads you wish to use" & VbCrLf &_
 "2 Launch Hidden?" & VbCrLf &_
 "3 Kill Silent Miners" & VbCrLf &_
 "4 Use AVX"

 rc=InputBox(strMenu,"Menu",1)
 If IsNumeric(rc2) Then
     Select Case rc
         Case 1 
             R2=2
         Case 2
             R2=4
		 Case 3
			 R2=5
	     Case 4
		     R2=7
     End Select
 End If
End if 
 'Begin Poweruser Menu Actions
 
 If R2=1 Then
intAnswer = _
    Msgbox("Open the CPU Miner?", _
        vbYesNo, "Mine on CPU?")
		
	End If

If intAnswer = vbYes Then
CPUMine=1
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(CPUPath,2,true)
objFileToWrite.WriteLine(CPUMine)
objFileToWrite.Close
Set objFileToWrite = Nothing
wscript.echo("Please restart the script")
End IF

If intAnswer = vbNo Then
CPUMine=0
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(CPUPath,2,true)
objFileToWrite.WriteLine(CPUMine)
objFileToWrite.Close
Set objFileToWrite = Nothing
wscript.echo("Please restart the script")
End If

If R2=2 Then

Threads = InputBox("Please enter the # of CPU threads to use", _
    "CPU threads")
	
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(ThreadsPath,2,true)
objFileToWrite.WriteLine(Threads)
objFileToWrite.Close

End if

If R2=3 Then
wscript.echo("Launching instructions to set up merged mining")
CreateObject("WScript.Shell").Run "http://myriad.p2pool.geek.nz/merge"
End If

If R2=4 Then
intHidden = _
    Msgbox("Launch Hidden?", _
        vbYesNo, "Yes")
		
	End If

If intHidden = vbYes Then
Hidden=1
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(HiddenPath,2,true)
objFileToWrite.WriteLine(Hidden)
objFileToWrite.Close
Set objFileToWrite = Nothing
wscript.echo("Please restart the script")
End IF

If intHidden = vbNo Then
Hidden=0	
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(HiddenPath,2,true)
objFileToWrite.WriteLine(Hidden)
objFileToWrite.Close
Set objFileToWrite = Nothing
wscript.echo("Please restart the script")
End If

If R2=5 Then
objshell.Run ".\KillSilent.bat"
Wscript.quit
End If

If R2=6 Then

strMenu="Select which pool to change" & VbCrLf &_
 "1 XMG"


 rc=InputBox(strMenu,"Menu",1)
 If IsNumeric(rc2) Then
     Select Case rc
         Case 1 
             R3=1
     End Select
 End If
 End If
 
 If R3=1 Then
XMGPool = InputBox("Enter your XMG Pool - URL ONLY", _
    "Stratum+tcp://Pool.URL:Port")
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(XMGpath,2,true)
objFileToWrite.WriteLine(XMGpool)
objFileToWrite.Close
End If


If R2=7 Then
UseAVX = InputBox("Use AVX? CPUZ can identify CPU support", _
    "AVX? 1 for yes 0 for no")
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(AVXpath,2,true)
objFileToWrite.WriteLine(UseAVX)
objFileToWrite.Close
End If
'Read Variables from files

Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(".\XMG.conf",1)
Set ObjFso2 = CreateObject("Scripting.FileSystemObject")
SkeinPool = objFileToRead.ReadAll()
objFileToRead.Close
Set objFileToRead = Nothing

Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(".\CPUMine.conf",1)
Set ObjFso2 = CreateObject("Scripting.FileSystemObject")
CPUMine = objFileToRead.ReadAll()
objFileToRead.Close
Set objFileToRead = Nothing

Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(".\Threads.conf",1)
Set ObjFso2 = CreateObject("Scripting.FileSystemObject")
Threads = objFileToRead.ReadAll()
objFileToRead.Close
Set objFileToRead = Nothing

Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(".\Hidden.conf",1)
Set ObjFso2 = CreateObject("Scripting.FileSystemObject")
Hidden = objFileToRead.ReadAll()
objFileToRead.Close
Set objFileToRead = Nothing

Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(".\credentials.conf",1)
Set ObjFso2 = CreateObject("Scripting.FileSystemObject")
Credentials = objFileToRead.ReadAll()
objFileToRead.Close
Set objFileToRead = Nothing

Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(".\AVX.conf",1)
Set ObjFso2 = CreateObject("Scripting.FileSystemObject")
UseAVX = objFileToRead.ReadAll()
objFileToRead.Close
Set objFileToRead = Nothing

'XMG Mining - CPU 
While R=1 do
If Hidden=1 Then
objshell.Run ".\XMGhidden.vbs" 
End If
If Hidden=0 Then
Set ObjShell = CreateObject("WScript.Shell") 
objshell.Run chr(34) & ".\XMG.vbs" & Chr(34), 0
Set ObjShell = Nothing
End If
wscript.quit
loop
wend

