Attribute VB_Name = "Module4"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Type PROCESSENTRY32
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long
   th32DefaultHeapID As Long
   th32ModuleID As Long
   cntThreads As Long
   th32ParentProcessID As Long
   pcPriClassBase As Long
   dwFlags As Long
   szExeFile As String * 260
End Type
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type
Private Const PROCESS_ALL_ACCESS = 0
Private Const TH32CS_SNAPPROCESS As Long = 2&

Private Const WINNT As Integer = 2
Private Const WIN98 As Integer = 1

Public KillAppReturn As Boolean

Public Function ShellLnk(ByVal LnkFile As String) As Long
    ShellLnk = ShellExecute(0&, vbNullString, LnkFile, vbNullString, vbNullString, vbNormalFocus)
End Function

Public Function getVersion() As Integer
 Dim udtOSInfo As OSVERSIONINFO
 Dim intRetVal As Integer
        
   With udtOSInfo
       .dwOSVersionInfoSize = 148
       .szCSDVersion = Space$(128)
   End With
   
   intRetVal = GetVersionExA(udtOSInfo)
 
   getVersion = udtOSInfo.dwPlatformId
End Function

Public Function Killapp(myName As String)
Select Case getVersion()
Case WIN98
Killapp9X (myName)
Case WINNT
KillappNT (myName)
End Select
End Function

Private Function KillappNT(myName As String)
   Dim uProcess As PROCESSENTRY32
   Dim rProcessFound As Long
   Dim hSnapShot As Long
   Dim szExename As String
   Dim exitCode As Long
   Dim myProcess As Long
   Dim AppKill As Boolean
   Dim appCount As Integer
   Dim I As Integer
   On Local Error GoTo Finish
   appCount = 0
   uProcess.dwSize = Len(uProcess)
   hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
   rProcessFound = ProcessFirst(hSnapShot, uProcess)
   Do While rProcessFound
       I = InStr(1, uProcess.szExeFile, Chr(0))
       szExename = LCase$(Left$(uProcess.szExeFile, I - 1))
       If Right$(szExename, Len(myName)) = LCase$(myName) Then
           KillAppReturn = True
           appCount = appCount + 1
           myProcess = OpenProcess(1&, -1&, uProcess.th32ProcessID)
           AppKill = TerminateProcess(myProcess, 0&)
           Call CloseHandle(myProcess)
       End If
       rProcessFound = ProcessNext(hSnapShot, uProcess)
   Loop
   Call CloseHandle(hSnapShot)
Finish:
KillAppReturn = False
End Function

Private Function Killapp9X(myName As String)
   Dim uProcess As PROCESSENTRY32
   Dim rProcessFound As Long
   Dim hSnapShot As Long
   Dim szExename As String
   Dim exitCode As Long
   Dim myProcess As Long
   Dim AppKill As Boolean
   Dim appCount As Integer
   Dim I As Integer
   On Local Error GoTo Finish
   appCount = 0
   uProcess.dwSize = Len(uProcess)
   hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
   rProcessFound = ProcessFirst(hSnapShot, uProcess)
   Do While rProcessFound
       I = InStr(1, uProcess.szExeFile, Chr(0))
       szExename = LCase$(Left$(uProcess.szExeFile, I - 1))
       If Right$(szExename, Len(myName)) = LCase$(myName) Then
           KillAppReturn = True
           appCount = appCount + 1
           myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
           AppKill = TerminateProcess(myProcess, exitCode)
           Call CloseHandle(myProcess)
       End If
       rProcessFound = ProcessNext(hSnapShot, uProcess)
   Loop
   Call CloseHandle(hSnapShot)
Finish:
KillAppReturn = False
End Function


