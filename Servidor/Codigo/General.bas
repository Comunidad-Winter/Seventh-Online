Attribute VB_Name = "General"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

'Global ANpc As Long
'Global Anpc_host As Long

Option Explicit

Global LeerNPCs As New clsIniReader
Global LeerNPCsHostiles As New clsIniReader
Global LeerClan As New clsIniReader

Public Function PonerPuntos(Numero As Long) As String
Dim i As Integer
Dim Cifra As String
 
Cifra = str(Numero)
Cifra = Right$(Cifra, Len(Cifra) - 1)
For i = 0 To 4
    If Len(Cifra) - 3 * i >= 3 Then
        If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
            PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos
        End If
    Else
        If Len(Cifra) - 3 * i > 0 Then
            PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos
        End If
        Exit For
    End If
Next
 
PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)
 
End Function

Sub DarCuerpoDesnudo(ByVal userindex As Integer, Optional ByVal Mimetizado As Boolean = False)

Select Case UCase$(UserList(userindex).Raza)
    Case "HUMANO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 21
                    Else
                        UserList(userindex).Char.Body = 21
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 39
                    Else
                        UserList(userindex).Char.Body = 39
                    End If
      End Select
    Case "ELFO OSCURO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 32
                    Else
                        UserList(userindex).Char.Body = 32
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 40
                    Else
                        UserList(userindex).Char.Body = 40
                    End If
      End Select
    Case "ENANO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 53
                    Else
                        UserList(userindex).Char.Body = 53
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 60
                    Else
                        UserList(userindex).Char.Body = 60
                    End If
      End Select
    Case "GNOMO"
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 53
                    Else
                        UserList(userindex).Char.Body = 53
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 60
                    Else
                        UserList(userindex).Char.Body = 60
                    End If
      End Select
    Case Else
      Select Case UCase$(UserList(userindex).Genero)
                Case "HOMBRE"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 21
                    Else
                        UserList(userindex).Char.Body = 21
                    End If
                Case "MUJER"
                    If Mimetizado Then
                        UserList(userindex).CharMimetizado.Body = 39
                    Else
                        UserList(userindex).Char.Body = 39
                    End If
      End Select
    
End Select

UserList(userindex).flags.Desnudo = 1

End Sub


Sub Bloquear(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Map As Integer, ByVal X As Integer, ByVal Y As Integer, b As Byte)
'b=1 bloquea el tile en (x,y)
'b=0 desbloquea el tile indicado

Call SendData(sndRoute, sndIndex, sndMap, "BQ" & X & "," & Y & "," & b)

End Sub


Function HayAgua(Map As Integer, X As Integer, Y As Integer) As Boolean

If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
    If MapData(Map, X, Y).Graphic(1) >= 1505 And _
       MapData(Map, X, Y).Graphic(1) <= 1520 And _
       MapData(Map, X, Y).Graphic(2) = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If

End Function

Sub LimpiarMundoEntero()
Call SendData(SendTarget.toall, 0, 0, "||Servidor> Limpiando Mundo." & FONTTYPE_SERVER)
Call SendData(SendTarget.toall, 0, 0, "||Servidor> Entrega de premios a clanes." & FONTTYPE_SERVER)
Dim MapaActual As Long
Dim Y As Long
Dim X As Long
 
For MapaActual = 1 To NumMaps
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If MapData(MapaActual, X, Y).OBJInfo.ObjIndex = 378 Then Exit For
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex > 0 And MapData(MapaActual, X, Y).Blocked = 0 Then
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex <> 555 Then
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex <> 674 Then
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex <> 554 Then
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex <> 162 Then
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex <> 168 Then
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex <> 804 Then
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex <> 805 Then
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex <> 806 Then
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex <> 807 Then
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex <> 808 Then
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex <> 566 Then
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex <> 569 Then
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex <> 573 Then
                If MapData(MapaActual, X, Y).OBJInfo.ObjIndex <> 570 Then
                If ItemNoEsDeMapa(MapData(MapaActual, X, Y).OBJInfo.ObjIndex) Then Call EraseObj(SendTarget.ToMap, 0, MapaActual, 10000, MapaActual, X, Y)
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            End If
        Next X
    Next Y
Next MapaActual
 
LimpiezaTimerMinutos = TimerCleanWorld
 
Call SendData(SendTarget.toall, 0, 0, "||Servidor> Limpieza del mundo realizada." & FONTTYPE_SERVER)
Call SendData(SendTarget.toall, 0, 0, "||Servidor> Entrega de premios realizada." & FONTTYPE_SERVER)
End Sub

Sub EnviarSpawnList(ByVal userindex As Integer)
Dim k As Integer, SD As String
SD = "SPL" & UBound(SpawnList) & ","

For k = 1 To UBound(SpawnList)
    SD = SD & SpawnList(k).NpcName & ","
Next k

Call SendData(SendTarget.toindex, userindex, 0, SD)
End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
#If UsarQueSocket = 0 Then

Obj.AddressFamily = AF_INET
Obj.Protocol = IPPROTO_IP
Obj.SocketType = SOCK_STREAM
Obj.Binary = False
Obj.Blocking = False
Obj.BufferSize = 1024
Obj.LocalPort = Port
Obj.backlog = 5
Obj.listen

#End If
End Sub




Sub Main()
On Error Resume Next
Dim f As Date

ChDir App.Path
ChDrive App.Path

Call BanIpCargar

Prision.Map = 66
Libertad.Map = 1

Prision.X = 55
Prision.Y = 55
Libertad.X = 50
Libertad.Y = 50


LastBackup = Format(Now, "Short Time")
Minutos = Format(Now, "Short Time")



ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer
ReDim Parties(1 To MAX_PARTIES) As clsParty
ReDim Guilds(1 To MAX_GUILDS) As clsClan



IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"



LevelSkill(1).LevelValue = 3
LevelSkill(2).LevelValue = 5
LevelSkill(3).LevelValue = 7
LevelSkill(4).LevelValue = 10
LevelSkill(5).LevelValue = 13
LevelSkill(6).LevelValue = 15
LevelSkill(7).LevelValue = 17
LevelSkill(8).LevelValue = 20
LevelSkill(9).LevelValue = 23
LevelSkill(10).LevelValue = 25
LevelSkill(11).LevelValue = 27
LevelSkill(12).LevelValue = 30
LevelSkill(13).LevelValue = 33
LevelSkill(14).LevelValue = 35
LevelSkill(15).LevelValue = 37
LevelSkill(16).LevelValue = 40
LevelSkill(17).LevelValue = 43
LevelSkill(18).LevelValue = 45
LevelSkill(19).LevelValue = 47
LevelSkill(20).LevelValue = 50
LevelSkill(21).LevelValue = 53
LevelSkill(22).LevelValue = 55
LevelSkill(23).LevelValue = 57
LevelSkill(24).LevelValue = 60
LevelSkill(25).LevelValue = 63
LevelSkill(26).LevelValue = 65
LevelSkill(27).LevelValue = 67
LevelSkill(28).LevelValue = 70
LevelSkill(29).LevelValue = 73
LevelSkill(30).LevelValue = 75
LevelSkill(31).LevelValue = 77
LevelSkill(32).LevelValue = 80
LevelSkill(33).LevelValue = 83
LevelSkill(34).LevelValue = 85
LevelSkill(35).LevelValue = 87
LevelSkill(36).LevelValue = 90
LevelSkill(37).LevelValue = 93
LevelSkill(38).LevelValue = 95
LevelSkill(39).LevelValue = 97
LevelSkill(40).LevelValue = 100
LevelSkill(41).LevelValue = 100
LevelSkill(42).LevelValue = 100
LevelSkill(43).LevelValue = 100
LevelSkill(44).LevelValue = 100
LevelSkill(45).LevelValue = 100
LevelSkill(46).LevelValue = 100
LevelSkill(47).LevelValue = 100
LevelSkill(48).LevelValue = 100
LevelSkill(49).LevelValue = 100
LevelSkill(50).LevelValue = 100


ListaRazas(1) = "Humano"
ListaRazas(2) = "Elfo"
ListaRazas(3) = "Elfo Oscuro"
ListaRazas(4) = "Gnomo"
ListaRazas(5) = "Enano"

Torneo_Clases_Validas(1) = "Guerrero"
Torneo_Clases_Validas(2) = "Mago"
Torneo_Clases_Validas(3) = "Paladin"
Torneo_Clases_Validas(4) = "Clerigo"
Torneo_Clases_Validas(5) = "Bardo"
Torneo_Clases_Validas(6) = "Asesino"
Torneo_Clases_Validas(7) = "Druida"
Torneo_Clases_Validas(8) = "Cazador"
 
Torneo_Alineacion_Validas(1) = "Criminal"
Torneo_Alineacion_Validas(2) = "Ciudadano"
Torneo_Alineacion_Validas(3) = "Armada Caos"
Torneo_Alineacion_Validas(4) = "Armada Real"


ListaClases(1) = "Mago"
ListaClases(2) = "Clerigo"
ListaClases(3) = "Guerrero"
ListaClases(4) = "Asesino"
ListaClases(5) = "Ladron"
ListaClases(6) = "Bardo"
ListaClases(7) = "Druida"
ListaClases(8) = "Bandido"
ListaClases(9) = "Paladin"
ListaClases(10) = "Cazador"
ListaClases(11) = "Pescador"
ListaClases(12) = "Herrero"
ListaClases(13) = "Leñador"
ListaClases(14) = "Minero"
ListaClases(15) = "Carpintero"
ListaClases(16) = "Sastre"
ListaClases(17) = "Pirata"

SkillsNames(1) = "Suerte"
SkillsNames(2) = "Magia"
SkillsNames(3) = "Robar"
SkillsNames(4) = "Tacticas de combate"
SkillsNames(5) = "Combate con armas"
SkillsNames(6) = "Meditar"
SkillsNames(7) = "Apuñalar"
SkillsNames(8) = "Ocultarse"
SkillsNames(9) = "Supervivencia"
SkillsNames(10) = "Talar arboles"
SkillsNames(11) = "Comercio"
SkillsNames(12) = "Defensa con escudos"
SkillsNames(13) = "Pesca"
SkillsNames(14) = "Mineria"
SkillsNames(15) = "Carpinteria"
SkillsNames(16) = "Herreria"
SkillsNames(17) = "Liderazgo"
SkillsNames(18) = "Domar animales"
SkillsNames(19) = "Armas de proyectiles"
SkillsNames(20) = "Wresterling"
SkillsNames(21) = "Navegacion"
SkillsNames(22) = "Equitacion"


frmCargando.Show

'Call PlayWaveAPI(App.Path & "\wav\harp3.wav")

frmMain.caption = frmMain.caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
IniPath = App.Path & "\"
CharPath = App.Path & "\Charfile\"

'Bordes del mapa
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)
DoEvents

frmCargando.Label1(2).caption = "Iniciando Arrays..."

Call LoadGuildsDB


Call CargarSpawnList
'¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
frmCargando.Label1(2).caption = "Cargando Server.ini"

MaxUsers = 0
Call LoadSini
Call CargaApuestas

'*************************************************
Call CargaNpcsDat
'*************************************************

frmCargando.Label1(2).caption = "Cargando Obj.Dat"
'Call LoadOBJData
Call LoadOBJData
    
frmCargando.Label1(2).caption = "Cargando Hechizos.Dat"
Call CargarHechizos
    
    
Call LoadArmasHerreria
Call LoadArmadurasHerreria
Call LoadObjCarpintero
Call LoadQuests

If BootDelBackUp Then
    
    frmCargando.Label1(2).caption = "Cargando BackUp"
    Call CargarBackUp
Else
    frmCargando.Label1(2).caption = "Cargando Mapas"
    Call LoadMapData
End If


Call SonidosMapas.LoadSoundMapInfo


'Comentado porque hay worldsave en ese mapa!
'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Dim LoopC As Integer

'Resetea las conexiones de los usuarios
For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
Next LoopC

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

With frmMain
    .LimpiezaTimer.Enabled = True
    .LimpiezaTimer.Enabled = True
    .Timer1.Enabled = True
    If ClientsCommandsQueue <> 0 Then
        .CmdExec.Enabled = True
    Else
        .CmdExec.Enabled = False
    End If
    .GameTimer.Enabled = True
    .FX.Enabled = True
    .Auditoria.Enabled = True
    .LimpiezaTimer.Enabled = True
    .TIMER_AI.Enabled = True
    .npcataca.Enabled = True
End With

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Configuracion de los sockets

Call SecurityIp.InitIpTables(1000)

#If UsarQueSocket = 1 Then

Call IniciaWsApi(frmMain.hWnd)
SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 0 Then

frmCargando.Label1(2).caption = "Configurando Sockets"

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Binary = False
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

Call ConfigListeningSocket(frmMain.Socket1, Puerto)

#ElseIf UsarQueSocket = 2 Then

frmMain.Serv.Iniciar Puerto

#ElseIf UsarQueSocket = 3 Then

frmMain.TCPServ.Encolar True
frmMain.TCPServ.IniciarTabla 1009
frmMain.TCPServ.SetQueueLim 51200
frmMain.TCPServ.Iniciar Puerto

#End If

If frmMain.Visible Then frmMain.txStatus.caption = "Escuchando conexiones entrantes ..."
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿




Unload frmCargando


'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
Close #N

'Ocultar
If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

tInicioServer = GetTickCount() And &H7FFFFFFF
Call InicializaEstadisticas

End Sub

Function FileExist(ByVal file As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************
    FileExist = Dir$(file, FileType) <> ""
End Function

Function ReadField(ByVal pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'All these functions are much faster using the "$" sign
'after the function. This happens for a simple reason:
'The functions return a variant without the $ sign. And
'variants are very slow, you should never use them.

'*****************************************************************
'Devuelve el string del campo
'*****************************************************************
Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String
  
Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = mid$(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = pos Then
            ReadField = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i

FieldNum = FieldNum + 1
If FieldNum = pos Then
    ReadField = mid$(Text, LastPos + 1)
End If

End Function
Public Function Tilde(Data As String) As String
 
Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(Data), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U")
 
End Function

Function MapaValido(ByVal Map As Integer) As Boolean
MapaValido = Map >= 1 And Map <= NumMaps
End Function

Sub MostrarNumUsers()
Call SendData(toall, 0, 0, "ON" & NumUsers)
frmMain.CantUsuarios.caption = "Numero de usuarios jugando: " & NumUsers

End Sub


Public Sub LogCriticEvent(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoReal(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoCaos(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogIndex(ByVal index As Integer, ByVal Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\" & index & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogError(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\errores.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogStatic(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogTarea(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile(1) ' obtenemos un canal
Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & Desc
Close #nfile

Exit Sub

errhandler:


End Sub


Public Sub LogClanes(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\IP.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub


Public Sub LogDesarrollo(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\desarrollo.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & str
Close #nfile

End Sub



Public Sub LogGM(Nombre As String, texto As String, VIP As Boolean)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
If VIP Then
    Open App.Path & "\logs\VIPs\" & Nombre & ".log" For Append Shared As #nfile
Else
    Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
End If
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub




Public Sub LogAsesinato(texto As String)
On Error GoTo errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub LogUSER(Nombre As String, texto As String, Consejero As Boolean)
On Error GoTo errhandler
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
If VIP Then
Open App.Path & "\logs\muertes\logsconsejeros" & Nombre & ".log" For Append Shared As #nfile
Else
Open App.Path & "\logs\muertes\logsusers" & Nombre & ".log" For Append Shared As #nfile
End If
Print #nfile, Date & " " & Time & " " & texto
Close #nfile
Exit Sub
errhandler:
End Sub
Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:


End Sub
Public Sub LogHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogCheating(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CH.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogCriticalHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time & " " & texto
Print #nfile, ""
Close #nfile

Exit Sub

errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
Dim Arg As String
Dim i As Integer


For i = 1 To 33

Arg = ReadField(i, cad, 44)

If Arg = "" Then Exit Function

Next i

ValidInputNP = True

End Function


Sub Restart()


'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

If frmMain.Visible Then frmMain.txStatus.caption = "Reiniciando."

Dim LoopC As Integer
  
#If UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    frmMain.Socket1.Startup
      
    frmMain.Socket2(0).Cleanup
    frmMain.Socket2(0).Startup

#ElseIf UsarQueSocket = 1 Then

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 2 Then

#End If

For LoopC = 1 To MaxUsers
    Call CloseSocket(LoopC)
Next

ReDim UserList(1 To MaxUsers)

For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
Next LoopC

LastUser = 0
NumUsers = 0

ReDim Npclist(1 To MAXNPCS) As npc 'NPCS
ReDim CharList(1 To MAXCHARS) As Integer

Call LoadSini
Call LoadOBJData

Call LoadMapData

Call CargarHechizos

#If UsarQueSocket = 0 Then

'*****************Setup socket
frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = 1024

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

'Escucha
frmMain.Socket1.LocalPort = Val(Puerto)
frmMain.Socket1.listen

#ElseIf UsarQueSocket = 1 Then

#ElseIf UsarQueSocket = 2 Then

#End If

If frmMain.Visible Then frmMain.txStatus.caption = "Escuchando conexiones entrantes ..."

'Log it
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & Time & " servidor reiniciado."
Close #N

'Ocultar

If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

  
End Sub


Public Function Intemperie(ByVal userindex As Integer) As Boolean
    
    If MapInfo(UserList(userindex).pos.Map).Zona <> "DUNGEON" Then
        If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger <> 1 And _
           MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger <> 2 And _
           MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger <> 4 Then Intemperie = True
    Else
        Intemperie = False
    End If
    
End Function

Public Sub TiempoInvocacion(ByVal userindex As Integer)
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(userindex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
           If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(userindex).MascotasIndex(i), 0)
        End If
    End If
Next i
End Sub

Public Sub EfectoFrio(ByVal userindex As Integer)

Dim modifi As Integer

If UserList(userindex).Counters.Frio < IntervaloFrio Then
  UserList(userindex).Counters.Frio = UserList(userindex).Counters.Frio + 1
Else
  If MapInfo(UserList(userindex).pos.Map).Terreno = Nieve Then
    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muriendo de frio, abrigate o moriras!!." & FONTTYPE_INFO)
    modifi = Porcentaje(UserList(userindex).Stats.MaxHP, 5)
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - modifi
    If UserList(userindex).Stats.MinHP < 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Has muerto de frio!!." & FONTTYPE_INFO)
            UserList(userindex).Stats.MinHP = 0
            Call UserDie(userindex)
    End If
    Call SendData(SendTarget.toindex, userindex, 0, "ASH" & UserList(userindex).Stats.MinHP)
  Else
    modifi = Porcentaje(UserList(userindex).Stats.MaxSta, 5)
    Call QuitarSta(userindex, modifi)
    Call SendData(SendTarget.toindex, userindex, 0, "ASS" & UserList(userindex).Stats.MinSta)
    'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Has perdido stamina, si no te abrigas rapido perderas toda!!." & FONTTYPE_INFO)
  End If
  
  UserList(userindex).Counters.Frio = 0
  
  
End If

End Sub

Public Sub EfectoMimetismo(ByVal userindex As Integer)

If UserList(userindex).Counters.Mimetismo < IntervaloInvisible Then
    UserList(userindex).Counters.Mimetismo = UserList(userindex).Counters.Mimetismo + 1
Else
    'restore old char
    Call SendData(SendTarget.toindex, userindex, 0, "||Recuperas tu apariencia normal." & FONTTYPE_INFO)
    
    UserList(userindex).Char.Body = UserList(userindex).CharMimetizado.Body
    UserList(userindex).Char.Head = UserList(userindex).CharMimetizado.Head
    UserList(userindex).Char.CascoAnim = UserList(userindex).CharMimetizado.CascoAnim
    UserList(userindex).Char.ShieldAnim = UserList(userindex).CharMimetizado.ShieldAnim
    UserList(userindex).Char.WeaponAnim = UserList(userindex).CharMimetizado.WeaponAnim
        
    
    UserList(userindex).Counters.Mimetismo = 0
    UserList(userindex).flags.Mimetizado = 0
    Call ChangeUserChar(SendTarget.ToMap, userindex, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
End If
            
End Sub



Public Sub EfectoInvisibilidad(ByVal userindex As Integer)

If UserList(userindex).Counters.Invisibilidad < IntervaloInvisible Then
    UserList(userindex).Counters.Invisibilidad = UserList(userindex).Counters.Invisibilidad + 1
Else
    UserList(userindex).Counters.Invisibilidad = 0
    UserList(userindex).flags.Invisible = 0
    If UserList(userindex).flags.Oculto = 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFO)
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
    End If
End If

End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
    Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
Else
    Npclist(NpcIndex).flags.Paralizado = 0
    Npclist(NpcIndex).flags.Inmovilizado = 0
End If

End Sub

Public Sub EfectoCegueEstu(ByVal userindex As Integer)

If UserList(userindex).Counters.Ceguera > 0 Then
    UserList(userindex).Counters.Ceguera = UserList(userindex).Counters.Ceguera - 1
Else
    If UserList(userindex).flags.Ceguera = 1 Then
        UserList(userindex).flags.Ceguera = 0
        Call SendData(SendTarget.toindex, userindex, 0, "NSEGUE")
    End If
    If UserList(userindex).flags.Estupidez = 1 Then
        UserList(userindex).flags.Estupidez = 0
        Call SendData(SendTarget.toindex, userindex, 0, "NESTUP")
    End If

End If


End Sub


Public Sub EfectoParalisisUser(ByVal userindex As Integer)

If UserList(userindex).Counters.Paralisis > 0 Then
    UserList(userindex).Counters.Paralisis = UserList(userindex).Counters.Paralisis - 1
Else
    UserList(userindex).flags.Paralizado = 0
    'UserList(UserIndex).Flags.AdministrativeParalisis = 0
    Call SendData(SendTarget.toindex, userindex, 0, "PARADOK")
End If

End Sub

Public Sub RecStamina(userindex As Integer, EnviarStats As Boolean, Intervalo As Integer)

If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = 1 And _
   MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = 2 And _
   MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = 4 Then Exit Sub


Dim massta As Integer
If UserList(userindex).Stats.MinSta < UserList(userindex).Stats.MaxSta Then
   If UserList(userindex).Counters.STACounter < Intervalo Then
       UserList(userindex).Counters.STACounter = UserList(userindex).Counters.STACounter + 1
   Else
       EnviarStats = True
       UserList(userindex).Counters.STACounter = 0
       massta = RandomNumber(1, Porcentaje(UserList(userindex).Stats.MaxSta, 5))
       UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta + massta
       If UserList(userindex).Stats.MinSta > UserList(userindex).Stats.MaxSta Then
            UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MaxSta
        End If
    End If
End If

End Sub

Public Sub EfectoVeneno(userindex As Integer, EnviarStats As Boolean)
Dim N As Integer

If UserList(userindex).Counters.Veneno < IntervaloVeneno Then
  UserList(userindex).Counters.Veneno = UserList(userindex).Counters.Veneno + 1
Else
  Call SendData(SendTarget.toindex, userindex, 0, "||Estas envenenado, si no te curas moriras." & FONTTYPE_VENENO)
  UserList(userindex).Counters.Veneno = 0
  N = RandomNumber(1, 5)
  UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - N
  If UserList(userindex).Stats.MinHP < 1 Then Call UserDie(userindex)
  Call SendData(SendTarget.toindex, userindex, 0, "ASH" & UserList(userindex).Stats.MinHP)
End If

End Sub

Public Sub DuracionPociones(userindex As Integer)

'Controla la duracion de las pociones
If UserList(userindex).flags.DuracionEfecto > 0 Then
   UserList(userindex).flags.DuracionEfecto = UserList(userindex).flags.DuracionEfecto - 1
   If UserList(userindex).flags.DuracionEfecto = 0 Then
        UserList(userindex).flags.TomoPocion = False
        UserList(userindex).flags.TipoPocion = 0
        'volvemos los atributos al estado normal
        Dim loopX As Integer
        For loopX = 1 To NUMATRIBUTOS
              UserList(userindex).Stats.UserAtributos(loopX) = UserList(userindex).Stats.UserAtributosBackUP(loopX)
              Call SendData(toindex, userindex, UserList(userindex).pos.Map, "PZ" & UserList(userindex).Stats.UserAtributos(Fuerza))
              Call SendData(toindex, userindex, UserList(userindex).pos.Map, "PX" & UserList(userindex).Stats.UserAtributos(Agilidad))

        Next
   End If
End If

End Sub

Public Sub Sanar(userindex As Integer, EnviarStats As Boolean, Intervalo As Integer)

If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = 1 And _
   MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = 2 And _
   MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = 4 Then Exit Sub
       

Dim mashit As Integer
'con el paso del tiempo va sanando....pero muy lentamente ;-)
If UserList(userindex).Stats.MinHP < UserList(userindex).Stats.MaxHP Then
   If UserList(userindex).Counters.HPCounter < Intervalo Then
      UserList(userindex).Counters.HPCounter = UserList(userindex).Counters.HPCounter + 1
   Else
      mashit = RandomNumber(2, Porcentaje(UserList(userindex).Stats.MaxSta, 5))
      
      UserList(userindex).Counters.HPCounter = 0
      UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP + mashit
      If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
      Call SendData(SendTarget.toindex, userindex, 0, "||Has sanado." & FONTTYPE_INFO)
      EnviarStats = True
    End If
End If

End Sub

Public Sub CargaNpcsDat()
'Dim NpcFile As String
'
'NpcFile = DatPath & "NPCs.dat"
'ANpc = INICarga(NpcFile)
'Call INIConf(ANpc, 0, "", 0)
'
'NpcFile = DatPath & "NPCs-HOSTILES.dat"
'Anpc_host = INICarga(NpcFile)
'Call INIConf(Anpc_host, 0, "", 0)

Dim npcfile As String

npcfile = DatPath & "NPCs.dat"
Call LeerNPCs.Initialize(npcfile)

npcfile = DatPath & "NPCs-HOSTILES.dat"
Call LeerNPCsHostiles.Initialize(npcfile)

End Sub
Public Sub CargarClan()
Dim clanfile As String
clanfile = App.Path & "\guilds\guildsinfo.inf"
Call LeerClan.Initialize(clanfile)
End Sub

Public Sub DescargaNpcsDat()
'If ANpc <> 0 Then Call INIDescarga(ANpc)
'If Anpc_host <> 0 Then Call INIDescarga(Anpc_host)

End Sub

Sub PasarSegundo()


    Dim i As Integer
    For i = 1 To LastUser
    
    If UserList(i).Counters.LentejaTiempo > 0 Then
    UserList(i).Counters.LentejaTiempo = UserList(i).Counters.LentejaTiempo - 1
    If UserList(i).Counters.LentejaTiempo <= 0 Then
            Call SendData(SendTarget.toindex, i, 0, "KKQ")
      End If
    End If
    
        If UserList(i).Counters.EntreTiempo > 0 Then
    UserList(i).Counters.EntreTiempo = UserList(i).Counters.EntreTiempo - 1
    If UserList(i).Counters.EntreTiempo <= 0 Then
            Call ReComenzarTouchDown
      End If
    End If
    
    If UserList(i).Counters.MuereEnTD > 0 Then
    UserList(i).Counters.MuereEnTD = UserList(i).Counters.MuereEnTD - 1
    If UserList(i).Counters.MuereEnTD <= 0 Then
    
    If UserList(i).flags.TeamTD = 1 Then
    Call RevivirUsuario(i)
    UserList(i).Stats.MinHP = UserList(i).Stats.MaxHP
    UserList(i).Stats.MinMAN = UserList(i).Stats.MaxMan
    UserList(i).Char.Body = 320
    Call WarpUserChar(i, 120, 57, 29)
    End If
    
    If UserList(i).flags.TeamTD = 2 Then
    Call RevivirUsuario(i)
    UserList(i).Stats.MinHP = UserList(i).Stats.MaxHP
    UserList(i).Stats.MinMAN = UserList(i).Stats.MaxMan
    UserList(i).Char.Body = 322
    Call WarpUserChar(i, 120, 24, 72)
    End If
    
      End If
    End If
    
    
    If UserList(i).flags.ActivoGema Then
            UserList(i).flags.TimeGema = UserList(i).flags.TimeGema - 1
            If UserList(i).flags.TimeGema <= 0 Then
            UserList(i).flags.ActivoGema = 0
            UserList(i).flags.GemaActivada = ""
            UserList(i).flags.TimeGema = 0
            SendData SendTarget.toindex, i, 0, "||El efecto de la Gema ha terminado." & "~255~0~0~1~0"
    End If
    End If
    
    If UserList(i).Counters.TiraItem > 0 Then
    UserList(i).Counters.TiraItem = UserList(i).Counters.TiraItem - 1
      End If
      
    If UserList(i).Counters.TimeComandos > 0 Then
    UserList(i).Counters.TimeComandos = UserList(i).Counters.TimeComandos - 1
      End If
    
        'Cerrar usuario
        If UserList(i).Counters.Saliendo Then
            UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
            If UserList(i).Counters.Salir <= 0 Then
                'If NumUsers <> 0 Then NumUsers = NumUsers - 1

                Call SendData(SendTarget.toindex, i, 0, "||Gracias por jugar SeventhAO" & FONTTYPE_INFO)
                Call SendData(SendTarget.toindex, i, 0, "MEJUI")
                
                Call CloseSocket(i)
                Exit Sub
            End If
         End If
         
                    If UserList(i).flags.Muerto Then
       
            If UserList(i).flags.TimeRevivir > 0 Then
   
             UserList(i).flags.TimeRevivir = UserList(i).flags.TimeRevivir - 1
             
        End If
   
    End If
         
    Next i

  
     If CuentaRegresiva > 0 Then
        If CuentaRegresiva > 1 Then
            Call SendData(SendTarget.toall, 0, 0, "||Contando..." & CuentaRegresiva - 1 & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.toall, 0, 0, "||YA!!!!!!!!!" & "~255~0~0~1~0")
        End If
        CuentaRegresiva = CuentaRegresiva - 1
    End If
    
         If CuentaArena > 0 Then
        If CuentaArena > 1 Then
            Call SendData(SendTarget.ToTorneo, 0, 0, "||Contando..." & CuentaArena - 1 & FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToTorneo, 0, 0, "||YA!!!!!!!!!" & "~255~0~0~1~0")
        End If
        CuentaArena = CuentaArena - 1
    End If
    
    If CuentaTorneo > 0 Then
        If CuentaTorneo > 1 Then
            Call SendData(SendTarget.toall, 0, 0, "||Inscripciones al Torneo abiertas en ... " & CuentaTorneo - 1 & "~255~255~255~1~0")
        Else
            Call SendData(SendTarget.toall, 0, 0, "||Inscripciones abiertas." & "~255~0~0~1~0")
                     Hay_Torneo = True
                     UsuariosEnTorneo = 0
        End If
        CuentaTorneo = CuentaTorneo - 1
    End If
End Sub
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
    'WorldSave
    Call DoBackUp

    'commit experiencias
    Call mdParty.ActualizaExperiencias

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

End Sub

 
Sub GuardarUsuarios()
    haciendoBK = True
    
    Call SendData(SendTarget.toall, 0, 0, "BKW")
    Call SendData(SendTarget.toall, 0, 0, "||Servidor> Grabando Personajes" & FONTTYPE_SERVER)
    
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            Call SaveUser(i, CharPath & UCase$(UserList(i).name) & ".chr")
        End If
    Next i
    
    Call SendData(SendTarget.toall, 0, 0, "||Servidor> Personajes Grabados" & FONTTYPE_SERVER)
    Call SendData(SendTarget.toall, 0, 0, "BKW")

    haciendoBK = False
End Sub


Sub InicializaEstadisticas()
Dim Ta As Long
Ta = GetTickCount() And &H7FFFFFFF

Call EstadisticasWeb.Inicializa(frmMain.hWnd)
Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)

End Sub


Function ZonaCura(ByVal userindex As Integer) As Boolean
' Autor: Joan Calderón - SaturoS.
'Codigo: Sacerdotes automaticos.
Dim X As Integer, Y As Integer
For Y = UserList(userindex).pos.Y - MinYBorder + 1 To UserList(userindex).pos.Y + MinYBorder - 1
        For X = UserList(userindex).pos.X - MinXBorder + 1 To UserList(userindex).pos.X + MinXBorder - 1
       
            If MapData(UserList(userindex).pos.Map, X, Y).NpcIndex > 0 Then
                If Npclist(MapData(UserList(userindex).pos.Map, X, Y).NpcIndex).NPCtype = 1 Then
                    If Distancia(UserList(userindex).pos, Npclist(MapData(UserList(userindex).pos.Map, X, Y).NpcIndex).pos) < 10 Then
                        ZonaCura = True
                        Exit Function
                    End If
                End If
            End If
           
        Next X
Next Y
ZonaCura = False
End Function


 
'[Loopzer]
'Anti-Cheats Lac(Loopzer Anti-Cheats)
Public Sub LoadAntiCheat()
    Dim i As Integer
 
        Lac_Camina = CLng(Val(GetVar$(App.Path & "\Anti Cheat.ini", "INTERVALOS", "Caminar")))
    Lac_Lanzar = CLng(Val(GetVar$(App.Path & "\Anti Cheat.ini", "INTERVALOS", "Lanzar")))
    Lac_Usar = CLng(Val(GetVar$(App.Path & "\Anti Cheat.ini", "INTERVALOS", "Usar")))
    Lac_Tirar = CLng(Val(GetVar$(App.Path & "\Anti Cheat.ini", "INTERVALOS", "Tirar")))
    Lac_Pociones = CLng(Val(GetVar$(App.Path & "\Anti Cheat.ini", "INTERVALOS", "Pociones")))
    Lac_Pegar = CLng(Val(GetVar$(App.Path & "\Anti Cheat.ini", "INTERVALOS", "Pegar")))
 
    For i = 1 To MaxUsers
        ResetearLac i
    Next
   
End Sub
Public Sub ResetearLac(userindex As Integer)
With UserList(userindex).Lac
    .LCaminar.init Lac_Camina
    .LPociones.init Lac_Pociones
    .LUsar.init Lac_Usar
    .LPegar.init Lac_Pegar
    .LLanzar.init Lac_Lanzar
    .LTirar.init Lac_Tirar
End With
 
End Sub
Public Sub CargaLac(userindex As Integer)
With UserList(userindex).Lac
    Set .LCaminar = New Cls_InterGTC
    Set .LLanzar = New Cls_InterGTC
    Set .LPegar = New Cls_InterGTC
    Set .LPociones = New Cls_InterGTC
    Set .LTirar = New Cls_InterGTC
    Set .LUsar = New Cls_InterGTC
 
    .LCaminar.init Lac_Camina
    .LPociones.init Lac_Pociones
    .LUsar.init Lac_Usar
    .LPegar.init Lac_Pegar
    .LLanzar.init Lac_Lanzar
    .LTirar.init Lac_Tirar
End With
 
End Sub
Public Sub DescargaLac(userindex As Integer)
Exit Sub
With UserList(userindex).Lac
    Set .LCaminar = Nothing
    Set .LLanzar = Nothing
    Set .LPegar = Nothing
    Set .LPociones = Nothing
    Set .LTirar = Nothing
    Set .LUsar = Nothing
End With
End Sub
'[/Loopzer]

Public Sub DragObjects(ByVal userindex As Integer)
Dim tmpUserObj As UserOBJ
 
    With UserList(userindex)
 
        tmpUserObj = .Invent.Object(ObjSlot1)
        .Invent.Object(ObjSlot1) = .Invent.Object(ObjSlot2)
        .Invent.Object(ObjSlot2) = tmpUserObj
 
        Call UpdateUserInv(False, userindex, ObjSlot1)
        Call UpdateUserInv(False, userindex, ObjSlot2)
    End With
End Sub

Public Sub ReyFeer()
    If ReyON = 1 Then Exit Sub
    Dim Feer As WorldPos
    Dim Rey As Integer
    Rey = 937 'Cambiar el 937 (REY)
 
    Feer.Map = 106 ' Cambiar por el numero del mapa
    Feer.X = 50 ' Cambiar por la X del mapa
    Feer.Y = 38 'Cambiar por la Y del mapa
    
    'Posiciones Guardias
    Dim Guardia1 As WorldPos
    Dim Guardia2 As WorldPos
    Dim Guardia3 As WorldPos
    Dim Guardia4 As WorldPos
    
    Dim Guardia As Integer
    Guardia = 938
    
    Guardia1.Map = 106
    Guardia1.X = 51
    Guardia1.Y = 37
    
    Guardia2.Map = 106
    Guardia2.X = 51
    Guardia2.Y = 36
    
    Guardia3.Map = 106
    Guardia3.X = 50
    Guardia3.Y = 36

    Guardia4.Map = 106
    Guardia4.X = 51
    Guardia4.Y = 36
    '/Posiciones Guardias
    
    If ReyON = 0 Then
    Call SendData(toall, 0, 0, "||El espiritu del rey reaparecio en las profundidades del castillo hundido.." & FONTTYPE_INFO)
    Call SpawnNpc(Rey, Feer, True, False)
    Npclist(Rey).Aura = 20248
    Call SpawnNpc(Guardia, Guardia1, True, False)
    Call SpawnNpc(Guardia, Guardia2, True, False)
    Call SpawnNpc(Guardia, Guardia3, True, False)
    Call SpawnNpc(Guardia, Guardia4, True, False)
    ReyON = 1
    End If
End Sub

Public Sub TerminaDuelin(ByVal Feer As Integer, ByVal Agus As Integer)

If UserList(Agus).flags.Muerto = 1 Then
UserList(Feer).flags.RondasDuelo = UserList(Feer).flags.RondasDuelo + 1
UserList(Feer).Stats.MinHP = UserList(Feer).Stats.MaxHP
UserList(Feer).Stats.MinMAN = UserList(Feer).Stats.MaxMan

Call RevivirUsuario(Agus)
UserList(Agus).Stats.MinHP = UserList(Agus).Stats.MaxHP
UserList(Agus).Stats.MinMAN = UserList(Agus).Stats.MaxMan

Call WarpUserChar(Feer, 12, 27, 46, False)
Call WarpUserChar(Agus, 12, 40, 55, False)
End If

If UserList(Feer).flags.Muerto = 1 Then
Call RevivirUsuario(Feer)
UserList(Feer).Stats.MinHP = UserList(Feer).Stats.MaxHP
UserList(Feer).Stats.MinMAN = UserList(Feer).Stats.MaxMan

UserList(Agus).flags.RondasDuelo = UserList(Agus).flags.RondasDuelo + 1
UserList(Agus).Stats.MinHP = UserList(Agus).Stats.MaxHP
UserList(Agus).Stats.MinMAN = UserList(Agus).Stats.MaxMan

Call WarpUserChar(Feer, 12, 27, 46, False)
Call WarpUserChar(Agus, 12, 40, 55, False)
End If

If UserList(Feer).flags.RondasDuelo >= 2 Then
        'Reset Duelo Usuario Perdedor
        UserList(Agus).flags.EnDuelo = False
        UserList(Agus).flags.DueliandoContra = ""
        UserList(Agus).flags.LeMandaronDuelo = False
        UserList(Agus).flags.UltimoEnMandarDuelo = ""
        UserList(Agus).flags.RondasDuelo = 0
        'Reset Duelo Usuario Perdedor
        'Set Usuario Ganador
        UserList(Feer).flags.EnDuelo = False
        UserList(Feer).flags.DueliandoContra = ""
        UserList(Feer).flags.RondasDuelo = 0
        'Set Usuario Ganador
        'Set Todo
        'UserList(Feer).DuelosGanados = UserList(Feer).DuelosGanados + 1
        'UserList(Agus).DuelosPerdidos =UserList(Agus).DuelosPerdidos + 1
        SendData SendTarget.toall, Agus, 0, "||Duelos: " & UserList(Feer).name & " venció en duelo a " & UserList(Agus).name & "." & "~255~255~255~0~1"
        WarpUserChar Agus, PosUserDuelo1.Map, PosUserDuelo1.X, PosUserDuelo1.Y, True 'No jodan, esta al revez porque a mi se me canta la chota ~ Feer~
        WarpUserChar Feer, PosUserDuelo2.Map, PosUserDuelo2.X, PosUserDuelo2.Y, True 'No jodan, esta al revez porque a mi se me canta la chota ~ Feer~
        UserList(Feer).Stats.DuelosGanados = UserList(Feer).Stats.DuelosGanados + 1
        UserList(Agus).Stats.DuelosPerdidos = UserList(Agus).Stats.DuelosPerdidos + 1
End If

If UserList(Agus).flags.RondasDuelo >= 2 Then
        'Reset Duelo Usuario Perdedor
        UserList(Feer).flags.EnDuelo = False
        UserList(Feer).flags.DueliandoContra = ""
        UserList(Feer).flags.LeMandaronDuelo = False
        UserList(Feer).flags.UltimoEnMandarDuelo = ""
        UserList(Feer).flags.RondasDuelo = 0
        'Reset Duelo Usuario Perdedor
        
        'Set Usuario Ganador
        UserList(Agus).flags.EnDuelo = False
        UserList(Agus).flags.DueliandoContra = ""
        UserList(Agus).flags.RondasDuelo = 0
        'Set Usuario Ganador
        'Set Todo
        
        'UserList(Agus).DuelosGanados = UserList(Agus).DuelosGanados + 1
        'UserList(Feer).DuelosPerdidos =UserList(Feer).DuelosPerdidos + 1
        SendData SendTarget.toall, Feer, 0, "||Duelos: " & UserList(Agus).name & " venció en duelo a " & UserList(Feer).name & "." & "~255~255~255~0~1"
        WarpUserChar Feer, PosUserDuelo1.Map, PosUserDuelo1.X, PosUserDuelo1.Y, True 'No jodan, esta al revez porque a mi se me canta la chota ~ Feer~
        WarpUserChar Agus, PosUserDuelo2.Map, PosUserDuelo2.X, PosUserDuelo2.Y, True 'No jodan, esta al revez porque a mi se me canta la chota ~ Feer~
        UserList(Agus).Stats.DuelosGanados = UserList(Agus).Stats.DuelosGanados + 1
        UserList(Feer).Stats.DuelosPerdidos = UserList(Feer).Stats.DuelosPerdidos + 1
End If

End Sub
