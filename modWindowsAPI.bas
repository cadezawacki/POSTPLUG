
Attribute VB_Name = "modWindowsAPI"
Option Explicit

Private Const MAX_PATH As Integer = 260
Private Const TH32CS_SNAPPROCESS As Long = 2&

Public Type ProcessInfo
    processID As Long
    processExists As Boolean
    processExeName As String
End Type

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As LongPtr
    th32moduleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

#If VBA7 Then

    ' VBA7 is true if this code runs in the VBA7-Environment (Excel/Office 2010 and above).
    ' This branch of code will run in both 32-bit and 64-bit VBA because
    ' PtrSafe keyword and LongPtr data type are available in both.

    Public Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
    
    Private Declare PtrSafe Function CreateToolHelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" _
        (ByVal lFlags As Long, ByVal lProcessID As Long) As LongPtr
    
    Private Declare PtrSafe Function ProcessFirst Lib "kernel32" Alias "Process32First" _
        (ByVal hSnapShot As LongPtr, uProcess As PROCESSENTRY32) As Long
    
    Private Declare PtrSafe Function ProcessNext Lib "kernel32" Alias "Process32Next" _
        (ByVal hSnapShot As LongPtr, uProcess As PROCESSENTRY32) As Long
    
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" _
        (ByVal hObject As LongPtr) As Long
    
    Private Declare PtrSafe Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpbuffer As String, nsize As Long) As Long
        
    Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
        (ByVal lpbuffer As String, nsize As Long) As Long
    
#Else
    
    ' This branch of code will be used if run with Excel versions older than 2010.
    ' There was only 32-bit Excel available.
    
    Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    
    Private Declare Function CreateToolHelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" _
        (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
    
    Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" _
        (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
        
    Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" _
        (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
    
    Private Declare Function CloseHandle Lib "kernel32" _
        (ByVal hObject As Long) As Long
    
    Private Declare Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpbuffer As String, nsize As Long) As Long
        
    Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
        (ByVal lpbuffer As String, nsize As Long) As Long

#End If
        
Public Function CMS_GetUserLoginName() As String
    ' Returns the network login name
    Dim lngLen As Long, lngX As Long
    Static userLoginName As String
    If Len(userLoginName) = 0 Then
        userLoginName = String$(254, 0)
        lngLen = 255
        lngX = apiGetUserName(userLoginName, lngLen)
        If (lngX > 0) Then
            userLoginName = Left$(userLoginName, lngLen - 1)
        Else
            userLoginName = vbNullString
        End If
    End If
    CMS_GetUserLoginName = LCase(userLoginName)
End Function

Public Function CMS_GetUserMachineName() As String
    Dim lRet As Long
    Dim lMaxLen As Long
    Static ssMachineName As String
    If Len(ssMachineName) = 0 Then
        lMaxLen = 100
        ssMachineName = String$(lMaxLen, vbNullChar)
        lRet = GetComputerName(ssMachineName, lMaxLen)
        ssMachineName = Left$(ssMachineName, lMaxLen)
    End If
    CMS_GetUserMachineName = UCase(ssMachineName)
End Function

Public Function CMS_GetUserDomainName() As String
Attribute CMS_GetUserDomainName.VB_Description = "Returns user's DomainName"
    Dim WshNetwork As Object
    Set WshNetwork = CreateObject("WScript.Network")
    CMS_GetUserDomainName = UCase(WshNetwork.UserDomain)
End Function

' ---------------------------------------------------------------------------------------------------------------------
' GetProcessInfo function uses CreateToolHelpSnapshot to take a snapshot of all processes in the system.
' Using TH32CS_SNAPPROCESS includes all processes in the system in the snapshot.
' Then it walks through the list of processes recorded in the snapshot using ProcessFirst and ProcessNext
' and looks for process with ID equal to processID (which is Excel PID saved previously to windows registry as session)
' ---------------------------------------------------------------------------------------------------------------------
Public Function GetProcessInfo(processID As Long, procInfo As ProcessInfo) As Boolean
    Dim uProcess As PROCESSENTRY32
    Dim uProcessSize As Long
    Dim rProcessFound As Long
    Dim hSnapShot As LongPtr
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim iFound As Integer

    On Error GoTo ErrHandler
    
    GetProcessInfo = False

    procInfo.processID = processID
    procInfo.processExists = False
    procInfo.processExeName = ""
    uProcessSize = LenB(uProcess) ' Use LenB to determine the actual number of bytes required to hold a given variable in memory.
    uProcess.dwSize = uProcessSize
    hSnapShot = CreateToolHelpSnapshot(TH32CS_SNAPPROCESS, 0) ' Takes a snapshot of all processes in the system
    rProcessFound = ProcessFirst(hSnapShot, uProcess) ' Retrieves information about the first process encountered in a system snapshot.

    Do While rProcessFound
       iFound = VBA.Strings.InStr(1, uProcess.szExeFile, Chr(0)) - 1
       If iFound > 0 Then
           szExename = LCase$(Left$(uProcess.szExeFile, iFound))
           If (uProcess.th32ProcessID = processID) Then
               procInfo.processExists = True
               procInfo.processExeName = szExename
               Call CloseHandle(hSnapShot)
               GetProcessInfo = True
               Exit Function
           End If
       End If
       rProcessFound = ProcessNext(hSnapShot, uProcess) ' Retrieves information about the next process recorded in a system snapshot.
    Loop
    
    Call CloseHandle(hSnapShot)
    
    Exit Function
    
ErrHandler:
    Debug.Print "Process handling failed: " & Err.Description, vbExclamation
    
End Function







