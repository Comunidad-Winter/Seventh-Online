Attribute VB_Name = "Admin"


Option Explicit

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type

Public Apuestas As tAPuestas

Public NPCs As Long
Public DebugSocket As Boolean

Public ReiniciarServer As Long

Public tInicioServer As Long
Public EstadisticasWeb As New clsEstadisticasIPC

Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloUserPuedeAtacar As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long '[Gonzalo]
Public IntervaloUserPuedeUsar As Long
Public IntervaloFlechasCazadores As Long

Public MinutosWs As Long
Public Puerto As Integer

Public MAXPASOS As Long

Public BootDelBackUp As Byte

Public IpList As New Collection
Public ClientsCommandsQueue As Byte

Public Type TCPESStats
    BytesEnviados As Double
    BytesRecibidos As Double
    BytesEnviadosXSEG As Long
    BytesRecibidosXSEG As Long
    BytesEnviadosXSEGMax As Long
    BytesRecibidosXSEGMax As Long
    BytesEnviadosXSEGCuando As Date
    BytesRecibidosXSEGCuando As Date
End Type

Public TCPESStats As TCPESStats

'Public ResetThread As New clsThreading

Function VersionOK(ByVal Ver As String) As Boolean
VersionOK = (Ver = ULTIMAVERSION)
End Function


Public Function ValidarLoginMSG(ByVal N As Integer) As Integer
On Error Resume Next
Dim AuxInteger As Integer
Dim AuxInteger2 As Integer
AuxInteger = SD(N)
AuxInteger2 = SDM(N)
ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function


Sub ReSpawnOrigPosNpcs()
On Error Resume Next

Dim i As Integer
Dim MiNPC As npc
   
For i = 1 To LastNPC
   'OJO
   If Npclist(i).flags.NPCActive Then
        
        If InMapBounds(Npclist(i).Orig.Map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)
        End If
        
        'tildada por sugerencia de yind
        'If Npclist(i).Contadores.TiempoExistencia > 0 Then
        '        Call MuereNpc(i, 0)
        'End If
   End If
   
Next i

End Sub

Sub WorldSave()
On Error Resume Next
'Call LogTarea("Sub WorldSave")

Dim loopX As Integer
Dim Porc As Long

Call SendData(SendTarget.ToAll, 0, 0, "||Servidor> Iniciando WorldSave" & FONTTYPE_SERVER)

#If SeguridadAlkon Then
    Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
#End If

Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales

Dim j As Integer, k As Integer

For j = 1 To NumMaps
    If MapInfo(j).BackUp = 1 Then k = k + 1
Next j

FrmStat.ProgressBar1.min = 0
FrmStat.ProgressBar1.max = k
FrmStat.ProgressBar1.value = 0

For loopX = 1 To NumMaps
    'DoEvents
    
    If MapInfo(loopX).BackUp = 1 Then
    
            Call GrabarMapa(loopX, App.Path & "\WorldBackUp\Mapa" & loopX)
            FrmStat.ProgressBar1.value = FrmStat.ProgressBar1.value + 1
    End If

Next loopX

FrmStat.Visible = False

If FileExist(DatPath & "\bkNpc.dat", vbNormal) Then Kill (DatPath & "bkNpc.dat")
If FileExist(DatPath & "\bkNPCs-HOSTILES.dat", vbNormal) Then Kill (DatPath & "bkNPCs-HOSTILES.dat")

For loopX = 1 To LastNPC
    If Npclist(loopX).flags.BackUp = 1 Then
            Call BackUPnPc(loopX)
    End If
Next

Call SendData(SendTarget.ToAll, 0, 0, "||Servidor> WorldSave ha conclu�do" & FONTTYPE_SERVER)

End Sub

Public Sub PurgarPenas()
Dim i As Integer
For i = 1 To LastUser
    If UserList(i).flags.UserLogged Then
    
        If UserList(i).Counters.Pena > 0 Then
                
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.Map, Libertad.X, Libertad.Y, True)
                    Call SendData(SendTarget.ToIndex, i, 0, "||Has sido liberado!" & FONTTYPE_INFO)
                End If
                
        End If
        
    End If
Next i
End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = "")
        
        UserList(UserIndex).Counters.Pena = Minutos
       
        
        Call WarpUserChar(UserIndex, Prision.Map, Prision.X, Prision.Y, True)
        
        If GmName = "" Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has sido encarcelado, deberas permanecer en la carcel " & Minutos & " minutos." & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & GmName & " te ha encarcelado, deberas permanecer en la carcel " & Minutos & " minutos." & FONTTYPE_INFO)
        End If
        
End Sub


Public Sub BorrarUsuario(ByVal UserName As String)
On Error Resume Next
If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
    Kill CharPath & UCase$(UserName) & ".chr"
End If
End Sub

Public Function BANCheck(ByVal name As String) As Boolean

BANCheck = (Val(GetVar(App.Path & "\charfile\" & name & ".chr", "FLAGS", "Ban")) = 1)

End Function

Public Function PersonajeExiste(ByVal name As String) As Boolean

PersonajeExiste = FileExist(CharPath & UCase$(name) & ".chr", vbNormal)

End Function
Public Function MD5ok(ByVal md5formateado As String) As Boolean
Dim i As Integer

If MD5ClientesActivado = 1 Then
    For i = 0 To UBound(MD5s)
        If (md5formateado = MD5s(i)) Then
            MD5ok = True
            Exit Function
        End If
    Next i
    MD5ok = False
Else
    MD5ok = True
End If

End Function

Public Sub MD5sCarga()
Dim LoopC As Integer

MD5ClientesActivado = Val(GetVar(IniPath & "Server.ini", "MD5Hush", "Activado"))

If MD5ClientesActivado = 1 Then
    ReDim MD5s(Val(GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptados")))
    For LoopC = 0 To UBound(MD5s)
        MD5s(LoopC) = GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptado" & (LoopC + 1))
        MD5s(LoopC) = txtOffset(hexMd52Asc(MD5s(LoopC)), 55)
    Next LoopC
End If

End Sub


Public Function UnBan(ByVal name As String) As Boolean
'Unban the character
Call WriteVar(App.Path & "\charfile\" & name & ".chr", "FLAGS", "Ban", "0")

'Remove it from the banned people database
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", name, "BannedBy", "NOBODY")
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", name, "Reason", "NO REASON")
End Function

Public Function CheckHD(ByVal hd As String) As Boolean
'***************************************************
'Author: Nahuel Casas (Zagen)
'Last Modify Date: 07/12/2009
' 07/12/2009: Zagen - Agreg� la funcion de agregar los digitos de un Serial Baneado.
'***************************************************
Open App.Path & "\DAT\BanHds.dat" For Input As #1
Dim Linea As String, Total As String
Do Until EOF(1)
Line Input #1, Linea
Total = Total + Linea + vbCrLf
Loop
Close #1
Dim Ret As String
If InStr(1, Total, hd) Then
CheckHD = True
End If
End Function

Public Sub BanIpAgrega(ByVal ip As String)
BanIps.Add ip

Call BanIpGuardar
End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long
Dim Dale As Boolean
Dim LoopC As Long

Dale = True
LoopC = 1
Do While LoopC <= BanIps.Count And Dale
    Dale = (BanIps.Item(LoopC) <> ip)
    LoopC = LoopC + 1
Loop

If Dale Then
    BanIpBuscar = 0
Else
    BanIpBuscar = LoopC - 1
End If
End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean

On Error Resume Next

Dim N As Long

N = BanIpBuscar(ip)
If N > 0 Then
    BanIps.Remove N
    BanIpGuardar
    BanIpQuita = True
Else
    BanIpQuita = False
End If

End Function

Public Sub BanIpGuardar()
Dim ArchivoBanIp As String
Dim ArchN As Long
Dim LoopC As Long

ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

ArchN = FreeFile()
Open ArchivoBanIp For Output As #ArchN

For LoopC = 1 To BanIps.Count
    Print #ArchN, BanIps.Item(LoopC)
Next LoopC

Close #ArchN

End Sub

Public Sub BanIpCargar()
Dim ArchN As Long
Dim Tmp As String
Dim ArchivoBanIp As String

ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

Do While BanIps.Count > 0
    BanIps.Remove 1
Loop

ArchN = FreeFile()
Open ArchivoBanIp For Input As #ArchN

Do While Not EOF(ArchN)
    Line Input #ArchN, Tmp
    BanIps.Add Tmp
Loop

Close #ArchN

End Sub

Public Sub ActualizaEstadisticasWeb()

Static Andando As Boolean
Static contador As Long
Dim Tmp As Boolean

contador = contador + 1

If contador >= 10 Then
    contador = 0
    Tmp = EstadisticasWeb.EstadisticasAndando()
    
    If Andando = False And Tmp = True Then
        Call InicializaEstadisticas
    End If
    
    Andando = Tmp
End If

End Sub

Public Sub ActualizaStatsES()

Static TUlt As Single
Dim Transcurrido As Single

Transcurrido = Timer - TUlt

If Transcurrido >= 5 Then
    TUlt = Timer
    With TCPESStats
        .BytesEnviadosXSEG = CLng(.BytesEnviados / Transcurrido)
        .BytesRecibidosXSEG = CLng(.BytesRecibidos / Transcurrido)
        .BytesEnviados = 0
        .BytesRecibidos = 0
        
        If .BytesEnviadosXSEG > .BytesEnviadosXSEGMax Then
            .BytesEnviadosXSEGMax = .BytesEnviadosXSEG
            .BytesEnviadosXSEGCuando = CDate(Now)
        End If
        
        If .BytesRecibidosXSEG > .BytesRecibidosXSEGMax Then
            .BytesRecibidosXSEGMax = .BytesRecibidosXSEG
            .BytesRecibidosXSEGCuando = CDate(Now)
        End If
        
        If frmEstadisticas.Visible Then
            Call frmEstadisticas.ActualizaStats
        End If
    End With
End If

End Sub


Public Function UserDarPrivilegioLevel(ByVal name As String) As Long
If EsAdministrador(name) Then
    UserDarPrivilegioLevel = 4
ElseIf EsDios(name) Then
    UserDarPrivilegioLevel = 3
ElseIf EsSemiDios(name) Then
    UserDarPrivilegioLevel = 2
ElseIf EsVIP(name) Then
    UserDarPrivilegioLevel = 1
Else
    UserDarPrivilegioLevel = 0
End If
End Function

