Attribute VB_Name = "mApis"

Public Enum SYSTEM_INFORMATION_CLASS
        SystemBasicInformation
        SystemHandleInformation
End Enum

Public Declare Function ZwQuerySystemInformation Lib "ntdll.dll" ( _
ByVal SystemInformationClass As SYSTEM_INFORMATION_CLASS, _
ByVal pSystemInformation As Long, _
ByVal SystemInformationLength As Long, _
ByRef ReturnLength As Long) As Long

Public Type SYSTEM_HANDLE_TABLE_ENTRY_INFO
        UniqueProcessId         As Integer
        CreatorBackTraceIndex   As Integer
        ObjectTypeIndex         As Byte
        HandleAttributes        As Byte
        HandleValue             As Integer
        pObject                 As Long
        GrantedAccess           As Long
End Type

Public Type SYSTEM_HANDLE_INFORMATION
        NumberOfHandles         As Long
        Handles(1 To 1)         As SYSTEM_HANDLE_TABLE_ENTRY_INFO
End Type

Public Const STATUS_INFO_LENGTH_MISMATCH = &HC0000004
Public Const STATUS_ACCESS_DENIED = &HC0000022

Public Declare Function ZwWriteVirtualMemory Lib "ntdll.dll" ( _
ByVal ProcessHandle As Long, _
ByVal BaseAddress As Long, _
ByVal pBuffer As Long, _
ByVal NumberOfBytesToWrite As Long, _
ByRef NumberOfBytesWritten As Long) As Long

Public Declare Function ZwOpenProcess Lib "ntdll.dll" ( _
ByRef ProcessHandle As Long, _
ByVal AccessMask As Long, _
ByRef ObjectAttributes As OBJECT_ATTRIBUTES, _
ByRef ClientId As CLIENT_ID) As Long

Public Type OBJECT_ATTRIBUTES
        Length              As Long
        RootDirectory       As Long
        ObjectName          As Long
        Attributes          As Long
        SecurityDescriptor  As Long
        SecurityQualityOfService As Long
End Type

Public Type CLIENT_ID
        UniqueProcess       As Long
        UniqueThread        As Long
End Type

Public Const PROCESS_QUERY_INFORMATION      As Long = &H400
Public Const STATUS_INVALID_CID             As Long = &HC000000B

Public Declare Function ZwClose Lib "ntdll.dll" ( _
ByVal ObjectHandle As Long) As Long

Public Const ZwGetCurrentProcess            As Long = -1
Public Const ZwGetCurrentThread             As Long = -2
Public Const ZwCurrentProcess               As Long = ZwGetCurrentProcess
Public Const ZwCurrentThread                As Long = ZwGetCurrentThread

Public Declare Function ZwCreateJobObject Lib "ntdll.dll" ( _
ByRef JobHandle As Long, _
ByVal DesiredAccess As Long, _
ByRef ObjectAttributes As OBJECT_ATTRIBUTES) As Long

Public Declare Function ZwAssignProcessToJobObject Lib "ntdll.dll" ( _
ByVal JobHandle As Long, _
ByVal ProcessHandle As Long) As Long

Public Declare Function ZwTerminateJobObject Lib "ntdll.dll" ( _
ByVal JobHandle As Long, _
ByVal ExitStatus As Long) As Long

Public Const OBJ_INHERIT = &H2
Public Const STANDARD_RIGHTS_REQUIRED       As Long = &HF0000
Public Const SYNCHRONIZE                    As Long = &H100000
Public Const JOB_OBJECT_ALL_ACCESS          As Long = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H1F
Public Const PROCESS_DUP_HANDLE             As Long = &H40
Public Const PROCESS_ALL_ACCESS             As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Public Const THREAD_ALL_ACCESS              As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF)
Public Const OB_TYPE_PROCESS                As Long = &H5

Public Type PROCESS_BASIC_INFORMATION
        ExitStatus          As Long
        PebBaseAddress      As Long
        AffinityMask        As Long
        BasePriority        As Long
        UniqueProcessId     As Long
        InheritedFromUniqueProcessId As Long
End Type

Public Declare Function ZwDuplicateObject Lib "ntdll.dll" ( _
ByVal SourceProcessHandle As Long, _
ByVal SourceHandle As Long, _
ByVal TargetProcessHandle As Long, _
ByRef TargetHandle As Long, _
ByVal DesiredAccess As Long, _
ByVal HandleAttributes As Long, _
ByVal Options As Long) As Long

Public Const DUPLICATE_CLOSE_SOURCE = &H1
Public Const DUPLICATE_SAME_ACCESS = &H2
Public Const DUPLICATE_SAME_ATTRIBUTES = &H4

Public Declare Function ZwQueryInformationProcess Lib "ntdll.dll" ( _
ByVal ProcessHandle As Long, _
ByVal ProcessInformationClass As PROCESSINFOCLASS, _
ByVal ProcessInformation As Long, _
ByVal ProcessInformationLength As Long, _
ByRef ReturnLength As Long) As Long

Public Enum PROCESSINFOCLASS
        ProcessBasicInformation
End Enum

Public Const STATUS_SUCCESS                 As Long = &H0
Public Const STATUS_INVALID_PARAMETER       As Long = &HC000000D

Public Declare Function EnumProcesses Lib "psapi.dll" ( _
ByVal lpidProcess As Long, _
ByVal cb As Long, _
ByRef cbNeeded As Long) As Long

Public Declare Function Api_GetProcessImageFileName Lib "psapi.dll" Alias "GetProcessImageFileNameA" ( _
ByVal hProcess As Long, _
ByVal lpImageFileName As String, _
ByVal nSize As Long) As Long

Public Declare Function EnumProcessModules Lib "psapi.dll" ( _
ByVal hProcess As Long, _
ByRef lphModule As Long, _
ByVal cb As Long, _
ByRef lpcbNeeded As Long) As Long

Public Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" ( _
ByVal hProcess As Long, _
ByVal hModule As Long, _
ByVal lpFileName As String, _
ByVal nSize As Long) As Long

Public Declare Function ZwTerminateProcess Lib "ntdll.dll" ( _
ByVal ProcessHandle As Long, _
ByVal ExitStatus As Long) As Long

Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Public Type SECURITY_ATTRIBUTES
        nLength             As Long
        lpSecurityDescriptor As Long
        bInheritHandle      As Long
End Type

Public Type a_my
         name               As String
         Pid                As Long
         tid                As Long
         Handle             As Long
End Type

Public Declare Function AssignProcessToJobObject Lib "Kernel32" ( _
ByVal hJob As Long, _
ByVal hProcess As Long) As Long

Public Declare Function TerminateJobObject Lib "Kernel32" ( _
ByVal hJob As Long, _
ByVal hProcess As Long) As Long

Public Declare Function CreateJobObject Lib "kernel32.dll" Alias "CreateJobObjectA" ( _
lpJobAttributes As SECURITY_ATTRIBUTES, _
lpName As String) As Long

Public Declare Function TerminateProcess Lib "kernel32.dll" ( _
ByVal hProcess As Long, _
ByVal uExitCode As Long) As Long

Public Declare Function WriteProcessMemory Lib "Kernel32" ( _
ByVal hProcess As Long, _
ByVal lpBaseAddress As Long, _
ByVal lpBuffer As Long, _
ByVal nSize As Long, _
lpNumberOfBytesWritten As Long) As Long

Public Declare Function API_CreateRemoteThread Lib "Kernel32" Alias "CreateRemoteThread" ( _
ByVal hProcess As Long, _
lpThreadAttributes As SECURITY_ATTRIBUTES, _
ByVal dwStackSize As Long, _
lpStartAddress As Long, _
lpParameter As Any, _
ByVal dwCreationFlags As Long, _
lpThreadId As Long) As Long

Public Declare Function API_GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" ( _
ByVal lpModuleName As String) As Long

Public Declare Function API_GetProcAddress Lib "Kernel32" Alias "GetProcAddress" ( _
ByVal hModule As Long, _
ByVal lpProcName As String) As Long

Public Declare Function OpenProcess0 Lib "Kernel32" Alias "OpenProcess" ( _
ByVal dwDesiredAccess As Long, _
ByVal bInheritHandle As Long, _
ByVal dwProcessId As Long) As Long

Public Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" ( _
ByVal lpModuleName As String) As Long

Public Declare Function GetProcAddress Lib "Kernel32" ( _
ByVal hModule As Long, _
ByVal lpProcName As String) As Long

Public Declare Function NtUnmapViewOfSection Lib "ntdll.dll" ( _
ByVal ProcessHandle As Long, _
ByVal BaseAddress As Long) As Long

Public Declare Function OpenThread Lib "Kernel32" ( _
ByVal h As Long, _
ByVal a As Boolean, _
ByVal b As Long) As Long

Public Declare Function TerminateThread Lib "Kernel32" ( _
ByVal a As Long, _
ByVal b As Long) As Long

Public Declare Function PostThreadMessageA Lib "user32" ( _
ByVal idThread As Long, _
ByVal msg As Long, _
ByVal wParam As Long, _
ByVal lParam As Long) As Long

Public Declare Function ZwOpenThread Lib "ntdll.dll" ( _
ByRef ThreadHandle As Long, _
ByVal AccessMask As Long, _
ByRef ObjectAttributes As OBJECT_ATTRIBUTES, _
ByRef ClientId As CLIENT_ID) As Long

Public Declare Function ZwTerminateThread Lib "ntdll.dll" ( _
ByVal ThreadHandle As Long, _
ByVal ExitStatus As Long) As Long
Public Function NT_SUCCESS(ByVal Status As Long) As Boolean
          NT_SUCCESS = (Status >= 0)
End Function

Public Sub CopyMemory(ByVal Dest As Long, ByVal Src As Long, ByVal cch As Long)
Dim Written As Long
        Call ZwWriteVirtualMemory(ZwCurrentProcess, Dest, Src, cch, Written)
End Sub

Public Function IsItemInArray(ByVal dwItem, ByRef dwArray() As Long) As Boolean
Dim Index As Long
        For Index = LBound(dwArray) To UBound(dwArray)
                If (dwItem = dwArray(Index)) Then IsItemInArray = True: Exit Function
        Next
        IsItemInArray = False
End Function

Public Sub AddItemToArray(ByVal dwItem As Long, ByRef dwArray() As Long)
On Error GoTo ErrHdl

        If (IsItemInArray(dwItem, dwArray)) Then Exit Sub
        
        ReDim Preserve dwArray(UBound(dwArray) + 1)
        dwArray(UBound(dwArray)) = dwItem
ErrHdl:
        
End Sub
