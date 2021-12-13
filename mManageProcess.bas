Attribute VB_Name = "mManageProcess"
'Thanks to the coder of this functions, it was taken from a VBGOOD
Option Explicit
Public Function OpenProcess(ByVal dwDesiredAccess As Long, ByVal bInhert As Boolean, ByVal ProcessId As Long) As Long
        Dim st As Long
        Dim cid As CLIENT_ID
        Dim OA As OBJECT_ATTRIBUTES
        Dim NumOfHandle As Long
        Dim pbi As PROCESS_BASIC_INFORMATION
        Dim I As Long
        Dim hProcessToDup As Long, hProcessCur As Long, hProcessToRet As Long
       Dim bytBuf() As Byte
        Dim arySize As Long: arySize = &H20000
        Do
                ReDim bytBuf(arySize)
                st = ZwQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), arySize, 0&)
                If (Not NT_SUCCESS(st)) Then
                        If (st <> STATUS_INFO_LENGTH_MISMATCH) Then
                                Erase bytBuf
                                Exit Function
                        End If
                Else
                        Exit Do
                End If

                arySize = arySize * 2
                ReDim bytBuf(arySize)
        Loop
        NumOfHandle = 0
        Call CopyMemory(VarPtr(NumOfHandle), VarPtr(bytBuf(0)), Len(NumOfHandle))
        Dim h_info() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
        ReDim h_info(NumOfHandle)
        Call CopyMemory(VarPtr(h_info(0)), VarPtr(bytBuf(0)) + Len(NumOfHandle), Len(h_info(0)) * NumOfHandle)
        For I = LBound(h_info) To UBound(h_info)
                With h_info(I)
                        If (.ObjectTypeIndex = OB_TYPE_PROCESS) Then
                                cid.UniqueProcess = .UniqueProcessId
                                st = ZwOpenProcess(hProcessToDup, PROCESS_DUP_HANDLE, OA, cid)
                                If (NT_SUCCESS(st)) Then
                                        st = ZwDuplicateObject(hProcessToDup, .HandleValue, ZwCurrentProcess, hProcessCur, PROCESS_ALL_ACCESS, 0, DUPLICATE_SAME_ATTRIBUTES)
                                        If (NT_SUCCESS(st)) Then
                                                st = ZwQueryInformationProcess(hProcessCur, ProcessBasicInformation, VarPtr(pbi), Len(pbi), 0)
                                                If (NT_SUCCESS(st)) Then
                                                        If (pbi.UniqueProcessId = ProcessId) Then
                                                                st = ZwDuplicateObject(hProcessToDup, .HandleValue, ZwGetCurrentProcess, hProcessToRet, dwDesiredAccess, OBJ_INHERIT, DUPLICATE_SAME_ATTRIBUTES)
                                                                If (NT_SUCCESS(st)) Then OpenProcess = hProcessToRet
                                                        End If
                                                End If
                                        End If
                                        st = ZwClose(hProcessCur)
                                End If
                                st = ZwClose(hProcessToDup)
                        End If
                End With
        Next
       
End Function
Public Function LzOpenProcess(ByVal dwDesiredAccess As Long, ByVal ProcessId As Long) As Long
          Dim st As Long
          Dim cid As CLIENT_ID
          Dim OA As OBJECT_ATTRIBUTES
          Dim NumOfHandle As Long
          Dim pbi As PROCESS_BASIC_INFORMATION
          Dim I As Long
          Dim hProcessToDup As Long, hProcessCur As Long, hProcessToRet As Long
          OA.Length = Len(OA)
        cid.UniqueProcess = ProcessId + 1
        st = ZwOpenProcess(hProcessToRet, dwDesiredAccess, OA, cid)
        If (NT_SUCCESS(st)) Then LzOpenProcess = hProcessToRet: Exit Function
        st = 0

          Dim bytBuf() As Byte
          Dim arySize As Long: arySize = 1
          Do
                  ReDim bytBuf(arySize)
                  st = ZwQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), arySize, 0&)
                  If (Not NT_SUCCESS(st)) Then
                          If (st <> STATUS_INFO_LENGTH_MISMATCH) Then
                                  Erase bytBuf
                                  Exit Function
                          End If
                  Else
                          Exit Do
                  End If
  
                  arySize = arySize * 2
                  ReDim bytBuf(arySize)
          Loop
          NumOfHandle = 0
          Call CopyMemory(VarPtr(NumOfHandle), VarPtr(bytBuf(0)), Len(NumOfHandle))
          Dim h_info() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
          ReDim h_info(NumOfHandle)
          Call CopyMemory(VarPtr(h_info(0)), VarPtr(bytBuf(0)) + Len(NumOfHandle), Len(h_info(0)) * NumOfHandle)
          
          For I = LBound(h_info) To UBound(h_info)
                  With h_info(I)
                          If (.ObjectTypeIndex = OB_TYPE_PROCESS) Then
                                  cid.UniqueProcess = .UniqueProcessId
                                  st = ZwOpenProcess(hProcessToDup, PROCESS_DUP_HANDLE, OA, cid)
                                  If (NT_SUCCESS(st)) Then
                                          st = ZwDuplicateObject(hProcessToDup, .HandleValue, ZwGetCurrentProcess, hProcessCur, PROCESS_ALL_ACCESS, 0, DUPLICATE_SAME_ATTRIBUTES)
                                          If (NT_SUCCESS(st)) Then
                                                  st = ZwQueryInformationProcess(hProcessCur, ProcessBasicInformation, VarPtr(pbi), Len(pbi), 0)
                                                  If (NT_SUCCESS(st)) Then
                                                          If (pbi.UniqueProcessId = ProcessId) Then
                                                                  st = ZwDuplicateObject(hProcessToDup, .HandleValue, ZwGetCurrentProcess, hProcessToRet, dwDesiredAccess, 0, DUPLICATE_SAME_ATTRIBUTES)
                                                                  If (NT_SUCCESS(st)) Then LzOpenProcess = hProcessToRet
                                                          End If
                                                  End If
                                          End If
                                          st = ZwClose(hProcessCur)
                                  End If
                                  st = ZwClose(hProcessToDup)
                          End If
                  End With
          Next
          Erase h_info
End Function

Public Function MyTerminateProcess(ByVal hProcess As Long, ByVal ExitStatus As Long) As Boolean
        Dim st As Long
        Dim hJob As Long
        Dim OA As OBJECT_ATTRIBUTES
        MyTerminateProcess = False
        OA.Length = Len(OA)
        st = ZwCreateJobObject(hJob, JOB_OBJECT_ALL_ACCESS, OA)
        If (NT_SUCCESS(st)) Then
                st = ZwAssignProcessToJobObject(hJob, hProcess)
                If (NT_SUCCESS(st)) Then
                        st = ZwTerminateJobObject(hJob, ExitStatus)
                        If (NT_SUCCESS(st)) Then MyTerminateProcess = True
                End If
                ZwClose (hJob)
        End If
        If (Not MyTerminateProcess) Then MyTerminateProcess = NT_SUCCESS(ZwTerminateProcess(hProcess, ExitStatus))
End Function



