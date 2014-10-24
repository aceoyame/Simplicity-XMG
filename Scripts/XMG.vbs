Option Explicit

'Declarations
Dim ObjShell
Dim OWSH
Dim WalletAdr 'Self Explanitory
Dim StrFileName
Dim Filesys
Dim objFSO
Dim ObjFSO2 ' Avoiding the certain redefinition
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
dim xmgpool
dim xmgpath
dim useavx
dim avxpath
dim credentials
dim cmdline
dim cmdlinepath

Set Filesys = CreateObject("Scripting.FileSystemObject") 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objShell = CreateObject("Wscript.Shell")
Set wshSystemEnv = objShell.Environment("SYSTEM")
ThreadsPath = ".\Threads.conf"
CPUPath = ".\CPUMine.conf"
HiddenPath = ".\Hidden.conf"
XMGPath = ".\XMG.conf"
AvxPath = ".\AVX.conf"
cmdlinepath = ".\cmdline.bat"


'Are we on X86 or X64?
ArchPath= "C:\Program Files (X86)"
If objFSO.FolderExists(ArchPath) then
Arch="X64"
End If

If Not objFSO.FolderExists(ArchPath) then
Arch="X86"
End If

'Open CL Check
Windir2=objShell.ExpandEnvironmentStrings("%WinDir%")
OCLPath= Windir2 & "\System32\OpenCL.DLL"
If objFSO.FileExists(OCLPath) then
OCL=1
End If
If Not objFSO.fileExists(OCLPath) then
OCL=0
End If

'Nvidia Check
Windir2=objShell.ExpandEnvironmentStrings("%WinDir%")
CudaPath= Windir2 & "\System32\NVCuda.DLL"
If objFSO.FileExists(CudaPath) then
Cuda=1
End If
If Not objFSO.fileExists(CudaPath) then
Cuda=0
End If

'Visual Studio 2010 runtime check x86

Windir2=objShell.ExpandEnvironmentStrings("%WinDir%")
MSVCRx86Path= Windir2 & "\System32\MSVCR100.DLL"
If objFSO.FileExists(MSVCRx86path) then
MSVCRx86=1
Else MSVCRx86=0
End If

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

'Read Variables from files
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(".\CPUMine.conf",1)
Set ObjFso2 = CreateObject("Scripting.FileSystemObject")
CPUMine = objFileToRead.ReadAll()
objFileToRead.Close
Set objFileToRead = Nothing

Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(".\Credentials.conf",1)
Set ObjFso2 = CreateObject("Scripting.FileSystemObject")
Credentials = objFileToRead.ReadAll()
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

Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(".\XMG.conf",1)
Set ObjFso2 = CreateObject("Scripting.FileSystemObject")
XMGPool = objFileToRead.ReadAll()
objFileToRead.Close
Set objFileToRead = Nothing

Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(".\AVX.conf",1)
Set ObjFso2 = CreateObject("Scripting.FileSystemObject")
UseAVX = objFileToRead.ReadAll()
objFileToRead.Close
Set objFileToRead = Nothing

'Qubit Mining - CPU and GPU
If Arch="X86" Then
objShell.CurrentDirectory = ".\XMG\32"
end if
If CPUMine=1 and Arch="X86" Then
cmdline=("Minerd.EXE" & " " & "-a m7mhash -o" & " " &  XMGPool & " " & credentials & "-t" & " " & threads)
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cmdlinePath,2,true)
objFileToWrite.WriteLine(cmdline)
objFileToWrite.Close
objshell.run cmdlinepath
End If
If Arch="X64" Then
objShell.CurrentDirectory = ".\XMG\64\"
End If
If Arch="X86" Then
objShell.CurrentDirectory = ".\XMG\x86\"
cmdline=("Minerd.EXE" & " " & "-a m7mhash -o" & " " &  XMGPool & " " & credentials & "-t" & " " & threads)
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cmdlinePath,2,true)
objFileToWrite.WriteLine(cmdline)
objFileToWrite.Close
End If
If CPUmine=1 and Arch="X64" and UseAVX=0 Then
objShell.CurrentDirectory = ".\Generic\"
cmdline=("Minerd.EXE" & " " & "-a m7mhash -o" & " " &  XMGPool & " " & credentials & "-t" & " " & threads)
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cmdlinePath,2,true)
objFileToWrite.WriteLine(cmdline)
objFileToWrite.Close
objshell.run cmdlinepath
End If
If CPUmine=1 and Arch="X64" and UseAVX=1 Then
objShell.CurrentDirectory = ".\AVX\"
cmdline=("Minerd.EXE" & " " & "-a m7mhash -o" & " " &  XMGPool & " " & credentials & "-t" & " " & threads)
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cmdlinePath,2,true)
objFileToWrite.WriteLine(cmdline)
objFileToWrite.Close
objshell.run cmdlinepath
End If
If AMD=1 And Arch="X64"  And CpuMine=1 Then
objShell.CurrentDirectory = ".\AMD\"
cmdline=("Minerd.EXE" & " " & "-a m7mhash -o" & " " &  XMGPool & " " & credentials & "-t" & " " & threads)
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cmdlinePath,2,true)
objFileToWrite.WriteLine(cmdline)
objFileToWrite.Close
objshell.run cmdlinepath
End IF
If AMD=1 And Arch="X86"Then
cmdline=("Minerd.EXE" & " " & "-a m7mhash -o" & " " &  XMGPool & " " & credentials & "-t" & " " & threads)
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cmdlinePath,2,true)
objFileToWrite.WriteLine(cmdline)
objFileToWrite.Close
objshell.run cmdlinepath
End If

