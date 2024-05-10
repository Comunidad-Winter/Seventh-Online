Attribute VB_Name = "ES"
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

Option Explicit

Public Sub CargarSpawnList()
    Dim N As Integer, LoopC As Integer
    N = Val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(N) As tCriaturasEntrenador
    For LoopC = 1 To N
        SpawnList(LoopC).NpcIndex = Val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & LoopC))
        SpawnList(LoopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & LoopC)
    Next LoopC
    
End Sub

Function EsAdministrador(ByVal name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String
 
NumWizs = Val(GetVar(IniPath & "Server.ini", "INIT", "Administradores"))
 
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Administradores", "Administrador" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(name) = NomB Then
        EsAdministrador = True
        Exit Function
    End If
Next WizNum
EsAdministrador = False
 
End Function

Function EsDios(ByVal name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = Val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(name) = NomB Then
        EsDios = True
        Exit Function
    End If
Next WizNum
EsDios = False
End Function

Function EsSemiDios(ByVal name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = Val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(name) = NomB Then
        EsSemiDios = True
        Exit Function
    End If
Next WizNum
EsSemiDios = False

End Function

Function EsVIP(ByVal name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = Val(GetVar(IniPath & "vip.ini", "INIT", "VIP"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "vip.ini", "VIP", "VIP" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(name) = NomB Then
        EsVIP = True
        Exit Function
    End If
Next WizNum
EsVIP = False
End Function

Function EsRolesMaster(ByVal name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = Val(GetVar(IniPath & "Server.ini", "INIT", "RolesMasters"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "RolesMasters", "RM" & WizNum))
    If Left(NomB, 1) = "*" Or Left(NomB, 1) = "+" Then NomB = Right(NomB, Len(NomB) - 1)
    If UCase$(name) = NomB Then
        EsRolesMaster = True
        Exit Function
    End If
Next WizNum
EsRolesMaster = False
End Function



Public Sub CargarHechizos()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer Hechizos.dat se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

On Error GoTo errhandler

If frmMain.Visible Then frmMain.txStatus.caption = "Cargando Hechizos."

Dim Hechizo As Integer
Dim Leer As New clsIniReader

Call Leer.Initialize(DatPath & "Hechizos.dat")

'obtiene el numero de hechizos
NumeroHechizos = Val(Leer.GetValue("INIT", "NumeroHechizos"))
ReDim Hechizos(1 To NumeroHechizos) As tHechizo

frmCargando.cargar.min = 0
frmCargando.cargar.max = NumeroHechizos
frmCargando.cargar.value = 0

'Llena la lista
For Hechizo = 1 To NumeroHechizos

    Hechizos(Hechizo).Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
    Hechizos(Hechizo).Desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
    Hechizos(Hechizo).PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
    
    Hechizos(Hechizo).HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
    Hechizos(Hechizo).TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
    Hechizos(Hechizo).PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
    
    Hechizos(Hechizo).Tipo = Val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
    Hechizos(Hechizo).WAV = Val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
    Hechizos(Hechizo).FXgrh = Val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
    
    Hechizos(Hechizo).loops = Val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
    
    Hechizos(Hechizo).Resis = Val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
    
    Hechizos(Hechizo).SubeHP = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
    Hechizos(Hechizo).MinHP = Val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
    Hechizos(Hechizo).MaxHP = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
    
    Hechizos(Hechizo).SubeMana = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
    Hechizos(Hechizo).MiMana = Val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
    Hechizos(Hechizo).MaMana = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
    
    Hechizos(Hechizo).SubeSta = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
    Hechizos(Hechizo).MinSta = Val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
    Hechizos(Hechizo).MaxSta = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
    
    Hechizos(Hechizo).Lenteja = Val(Leer.GetValue("Hechizo" & Hechizo, "Lenteja"))
    Hechizos(Hechizo).ActivaVIP = Val(Leer.GetValue("Hechizo" & Hechizo, "ActivaVIP"))
    
    Hechizos(Hechizo).SubeAgilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
    Hechizos(Hechizo).MinAgilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
    Hechizos(Hechizo).MaxAgilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
    
    Hechizos(Hechizo).SubeFuerza = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
    Hechizos(Hechizo).MinFuerza = Val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
    Hechizos(Hechizo).MaxFuerza = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
    
    Hechizos(Hechizo).SubeCarisma = Val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
    Hechizos(Hechizo).MinCarisma = Val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
    Hechizos(Hechizo).MaxCarisma = Val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
    
    
    Hechizos(Hechizo).Invisibilidad = Val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
    Hechizos(Hechizo).Paraliza = Val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
    Hechizos(Hechizo).Inmoviliza = Val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
    Hechizos(Hechizo).RemoverParalisis = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
    Hechizos(Hechizo).RemoverEstupidez = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
    Hechizos(Hechizo).RemoverEstupidez = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
    Hechizos(Hechizo).RemueveInvisibilidadParcial = Val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
    
    
    Hechizos(Hechizo).CuraVeneno = Val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
    Hechizos(Hechizo).Envenena = Val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
    Hechizos(Hechizo).Maldicion = Val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
    Hechizos(Hechizo).RemoverMaldicion = Val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
    Hechizos(Hechizo).Bendicion = Val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
    Hechizos(Hechizo).Revivir = Val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
    Hechizos(Hechizo).ExclusivoClase = UCase$(Leer.GetValue("Hechizo" & Hechizo, "ExclusivoClase"))
    Hechizos(Hechizo).ProhibidoClase = UCase$(Leer.GetValue("Hechizo" & Hechizo, "ProhibidoClase"))
    
    Hechizos(Hechizo).Ceguera = Val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
    Hechizos(Hechizo).Estupidez = Val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
    
    Hechizos(Hechizo).Invoca = Val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
    Hechizos(Hechizo).NumNpc = Val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
    Hechizos(Hechizo).Cant = Val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
    Hechizos(Hechizo).Mimetiza = Val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
    
    
    Hechizos(Hechizo).Materializa = Val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
    Hechizos(Hechizo).ItemIndex = Val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
    
    Hechizos(Hechizo).MinSkill = Val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
    Hechizos(Hechizo).ManaRequerido = Val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
    
    'Barrin 30/9/03
    Hechizos(Hechizo).StaRequerido = Val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
    
    Hechizos(Hechizo).Target = Val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
    frmCargando.cargar.value = frmCargando.cargar.value + 1
    
    Hechizos(Hechizo).NeedStaff = Val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
    Hechizos(Hechizo).StaffAffected = CBool(Val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
    
Next Hechizo

Set Leer = Nothing
Exit Sub

errhandler:
 MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.Description
 
End Sub

Public Sub DarPremioCastillos()
On Error GoTo handler
Dim LoopC As Integer
    For LoopC = 1 To LastUser
        If UserList(LoopC).GuildIndex <> 0 Then
            If Guilds(UserList(LoopC).GuildIndex).GuildName = CastilloNorte Then
                UserList(LoopC).Stats.PuntosTorneo = UserList(LoopC).Stats.PuntosTorneo + 7
                Call SendData(SendTarget.toindex, (LoopC), 0, "||Has Recibido 7 puntos de torneo por mantener el Castillo Norte." & FONTTYPE_GUILD)
                Call EnviarPuntos(LoopC)
                Call SendData(SendTarget.toindex, 0, 0, "TW")
            End If
            If Guilds(UserList(LoopC).GuildIndex).GuildName = CastilloSur Then
                UserList(LoopC).Stats.PuntosTorneo = UserList(LoopC).Stats.PuntosTorneo + 8
                Call SendData(SendTarget.toindex, (LoopC), 0, "||Has Recibido 8 puntos de torneo por mantener el Castillo Sur." & FONTTYPE_GUILD)
                Call EnviarPuntos(LoopC)
                Call SendData(SendTarget.toindex, 0, 0, "TW")
            End If
            If Guilds(UserList(LoopC).GuildIndex).GuildName = CastilloEste Then
                UserList(LoopC).Stats.PuntosTorneo = UserList(LoopC).Stats.PuntosTorneo + 8
                Call SendData(SendTarget.toindex, (LoopC), 0, "||Has Recibido 8 puntos de torneo por mantener el Castillo Este." & FONTTYPE_GUILD)
                Call EnviarPuntos(LoopC)
                Call SendData(SendTarget.toindex, 0, 0, "TW")
            End If
            If Guilds(UserList(LoopC).GuildIndex).GuildName = CastilloOeste Then
                UserList(LoopC).Stats.PuntosTorneo = UserList(LoopC).Stats.PuntosTorneo + 7
                Call SendData(SendTarget.toindex, (LoopC), 0, "||Has Recibido 7 puntos de torneo por mantener el Castillo Oeste." & FONTTYPE_GUILD)
                Call EnviarPuntos(LoopC)
                Call SendData(SendTarget.toindex, 0, 0, "TW")
            End If
        End If
    Next LoopC
Exit Sub
handler:
Call LogError("Error en DarPremioCastillos.")
End Sub

Public Sub DoBackUp()
'Call LogTarea("Sub DoBackUp")
haciendoBK = True
Dim i As Integer



' Lo saco porque elimina elementales y mascotas - Maraxus
''''''''''''''lo pongo aca x sugernecia del yind
'For i = 1 To LastNPC
'    If Npclist(i).flags.NPCActive Then
'        If Npclist(i).Contadores.TiempoExistencia > 0 Then
'            Call MuereNpc(i, 0)
'        End If
'    End If
'Next i
'''''''''''/'lo pongo aca x sugernecia del yind



Call SendData(SendTarget.toall, 0, 0, "BKW")


Call WorldSave
Call modGuilds.v_RutinaElecciones
Call ResetCentinelaInfo     'Reseteamos al centinela


Call SendData(SendTarget.toall, 0, 0, "BKW")

'Call EstadisticasWeb.Informar(EVENTO_NUEVO_CLAN, 0)

haciendoBK = False

'Log
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
Print #nfile, Date & " " & Time
Close #nfile
End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByVal MAPFILE As String)
On Error Resume Next
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim TempInt As Integer
    Dim LoopC As Long
    
    If FileExist(MAPFILE & ".map", vbNormal) Then
        Kill MAPFILE & ".map"
    End If
    
    If FileExist(MAPFILE & ".inf", vbNormal) Then
        Kill MAPFILE & ".inf"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open MAPFILE & ".Map" For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open MAPFILE & ".Inf" For Binary As FreeFileInf
    Seek FreeFileInf, 1
    'map Header
            
    Put FreeFileMap, , MapInfo(Map).MapVersion
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
                ByFlags = 0
                
                If MapData(Map, X, Y).Blocked Then ByFlags = ByFlags Or 1
                If MapData(Map, X, Y).Graphic(2) Then ByFlags = ByFlags Or 2
                If MapData(Map, X, Y).Graphic(3) Then ByFlags = ByFlags Or 4
                If MapData(Map, X, Y).Graphic(4) Then ByFlags = ByFlags Or 8
                If MapData(Map, X, Y).trigger Then ByFlags = ByFlags Or 16
                
                Put FreeFileMap, , ByFlags
                
                Put FreeFileMap, , MapData(Map, X, Y).Graphic(1)
                
                For LoopC = 2 To 4
                    If MapData(Map, X, Y).Graphic(LoopC) Then _
                        Put FreeFileMap, , MapData(Map, X, Y).Graphic(LoopC)
                Next LoopC
                
                If MapData(Map, X, Y).trigger Then _
                    Put FreeFileMap, , CInt(MapData(Map, X, Y).trigger)
                
                '.inf file
                
                ByFlags = 0
                
                If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
                   If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otFogata Then
                        MapData(Map, X, Y).OBJInfo.ObjIndex = 0
                        MapData(Map, X, Y).OBJInfo.Amount = 0
                    End If
                End If
    
                If MapData(Map, X, Y).TileExit.Map Then ByFlags = ByFlags Or 1
                If MapData(Map, X, Y).NpcIndex Then ByFlags = ByFlags Or 2
                If MapData(Map, X, Y).OBJInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
                Put FreeFileInf, , ByFlags
                
                If MapData(Map, X, Y).TileExit.Map Then
                    Put FreeFileInf, , MapData(Map, X, Y).TileExit.Map
                    Put FreeFileInf, , MapData(Map, X, Y).TileExit.X
                    Put FreeFileInf, , MapData(Map, X, Y).TileExit.Y
                End If
                
                If MapData(Map, X, Y).NpcIndex Then _
                    Put FreeFileInf, , Npclist(MapData(Map, X, Y).NpcIndex).Numero
                
                If MapData(Map, X, Y).OBJInfo.ObjIndex Then
                    Put FreeFileInf, , MapData(Map, X, Y).OBJInfo.ObjIndex
                    Put FreeFileInf, , MapData(Map, X, Y).OBJInfo.Amount
                End If
            
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf

    'write .dat file
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Name", MapInfo(Map).name)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "MusicNum", MapInfo(Map).Music)
    Call WriteVar(MAPFILE & ".dat", "mapa" & Map, "MagiaSinefecto", MapInfo(Map).MagiaSinEfecto)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "StartPos", MapInfo(Map).StartPos.Map & "-" & MapInfo(Map).StartPos.X & "-" & MapInfo(Map).StartPos.Y)

    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Terreno", MapInfo(Map).Terreno)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Zona", MapInfo(Map).Zona)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Restringir", MapInfo(Map).Restringir)
    Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "BackUp", str(MapInfo(Map).BackUp))

    If MapInfo(Map).Pk Then
        Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Pk", "0")
    Else
        Call WriteVar(MAPFILE & ".dat", "Mapa" & Map, "Pk", "1")
    End If

End Sub
Sub LoadArmasHerreria()

Dim N As Integer, lc As Integer

N = Val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

ReDim Preserve ArmasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmasHerrero(lc) = Val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
Next lc

End Sub

Sub LoadArmadurasHerreria()

Dim N As Integer, lc As Integer

N = Val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

ReDim Preserve ArmadurasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmadurasHerrero(lc) = Val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
Next lc

End Sub

Sub LoadObjCarpintero()

Dim N As Integer, lc As Integer

N = Val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))

ReDim Preserve ObjCarpintero(1 To N) As Integer

For lc = 1 To N
    ObjCarpintero(lc) = Val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
Next lc

End Sub



Sub LoadOBJData()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer desde el OBJ.DAT se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

'Call LogTarea("Sub LoadOBJData")

On Error GoTo errhandler

If frmMain.Visible Then frmMain.txStatus.caption = "Cargando base de datos de los objetos."

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim Object As Integer
Dim Leer As New clsIniReader

Call Leer.Initialize(DatPath & "Obj.dat")

'obtiene el numero de obj
NumObjDatas = Val(Leer.GetValue("INIT", "NumObjs"))

frmCargando.cargar.min = 0
frmCargando.cargar.max = NumObjDatas
frmCargando.cargar.value = 0


ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
'Llena la lista
For Object = 1 To NumObjDatas
        
    ObjData(Object).name = Leer.GetValue("OBJ" & Object, "Name")
    
    ObjData(Object).GrhIndex = Val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
    If ObjData(Object).GrhIndex = 0 Then
        ObjData(Object).GrhIndex = ObjData(Object).GrhIndex
    End If
    
    ObjData(Object).OBJType = Val(Leer.GetValue("OBJ" & Object, "ObjType"))
    
    ObjData(Object).Newbie = Val(Leer.GetValue("OBJ" & Object, "Newbie"))
    ObjData(Object).Aura = Val(Leer.GetValue("OBJ" & Object, "CreaAura"))
    
    Select Case ObjData(Object).OBJType
        Case eOBJType.otArmadura
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
        
        Case eOBJType.otESCUDO
            ObjData(Object).ShieldAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otCASCO
            ObjData(Object).CascoAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otWeapon
            ObjData(Object).WeaponAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).Apuñala = Val(Leer.GetValue("OBJ" & Object, "Apuñala"))
            ObjData(Object).DosManos = Val(Leer.GetValue("OBJ" & Object, "DosManos"))
            ObjData(Object).Envenena = Val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).Paraliza = Val(Leer.GetValue("OBJ" & Object, "Paraliza"))
            ObjData(Object).MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).proyectil = Val(Leer.GetValue("OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = Val(Leer.GetValue("OBJ" & Object, "Municiones"))
            ObjData(Object).StaffPower = Val(Leer.GetValue("OBJ" & Object, "StaffPower"))
            ObjData(Object).StaffDamageBonus = Val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
            ObjData(Object).Refuerzo = Val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
            
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otHerramientas
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
        
        Case eOBJType.otInstrumentos
            ObjData(Object).Snd1 = Val(Leer.GetValue("OBJ" & Object, "SND1"))
            ObjData(Object).Snd2 = Val(Leer.GetValue("OBJ" & Object, "SND2"))
            ObjData(Object).Snd3 = Val(Leer.GetValue("OBJ" & Object, "SND3"))
        
        Case eOBJType.otMinerales
            ObjData(Object).MinSkill = Val(Leer.GetValue("OBJ" & Object, "MinSkill"))
        
        Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
            ObjData(Object).IndexAbierta = Val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
            ObjData(Object).IndexCerrada = Val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
            ObjData(Object).IndexCerradaLlave = Val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
        
        Case otPociones
            ObjData(Object).TipoPocion = Val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
            ObjData(Object).MaxModificador = Val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
            ObjData(Object).MinModificador = Val(Leer.GetValue("OBJ" & Object, "MinModificador"))
            ObjData(Object).DuracionEfecto = Val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
            
        Case eOBJType.otMonturas
            ObjData(Object).MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
        
        Case eOBJType.otBarcos
            ObjData(Object).MinSkill = Val(Leer.GetValue("OBJ" & Object, "MinSkill"))
            ObjData(Object).MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
        
        Case eOBJType.otFlechas
            ObjData(Object).MaxHIT = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).Envenena = Val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).Paraliza = Val(Leer.GetValue("OBJ" & Object, "Paraliza"))
    End Select
    
    ObjData(Object).Ropaje = Val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
    ObjData(Object).HechizoIndex = Val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
    
    ObjData(Object).LingoteIndex = Val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
    
    ObjData(Object).MineralIndex = Val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
    
    ObjData(Object).MaxHP = Val(Leer.GetValue("OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = Val(Leer.GetValue("OBJ" & Object, "MinHP"))
    
    ObjData(Object).Mujer = Val(Leer.GetValue("OBJ" & Object, "Mujer"))
    ObjData(Object).Hombre = Val(Leer.GetValue("OBJ" & Object, "Hombre"))
    
    ObjData(Object).MinDef = Val(Leer.GetValue("OBJ" & Object, "MINDEF"))
    ObjData(Object).MaxDef = Val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
    
    ObjData(Object).RazaEnana = Val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
    
    ObjData(Object).Valor = Val(Leer.GetValue("OBJ" & Object, "Valor"))
    ObjData(Object).Skill = Val(Leer.GetValue("OBJ" & Object, "Skill"))
    ObjData(Object).SkillM = Val(Leer.GetValue("OBJ" & Object, "SkillM"))
    
    ObjData(Object).Crucial = Val(Leer.GetValue("OBJ" & Object, "Crucial"))
    
    ObjData(Object).Cerrada = Val(Leer.GetValue("OBJ" & Object, "abierta"))
    If ObjData(Object).Cerrada = 1 Then
        ObjData(Object).Llave = Val(Leer.GetValue("OBJ" & Object, "Llave"))
        ObjData(Object).clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))
    End If
    
    'Puertas y llaves
    ObjData(Object).clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))
    
    ObjData(Object).texto = Leer.GetValue("OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = Val(Leer.GetValue("OBJ" & Object, "VGrande"))
    
    ObjData(Object).Agarrable = Val(Leer.GetValue("OBJ" & Object, "Agarrable"))
    ObjData(Object).ForoID = Leer.GetValue("OBJ" & Object, "ID")
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        ObjData(Object).ClaseProhibida(i) = Leer.GetValue("OBJ" & Object, "CP" & i)
    Next i
    
    ObjData(Object).DefensaMagicaMax = Val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
    ObjData(Object).DefensaMagicaMin = Val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
    
    ObjData(Object).SkCarpinteria = Val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
    
    If ObjData(Object).SkCarpinteria > 0 Then _
        ObjData(Object).Madera = Val(Leer.GetValue("OBJ" & Object, "Madera"))
    
    'Bebidas
    ObjData(Object).MinSta = Val(Leer.GetValue("OBJ" & Object, "MinST"))
    
    ObjData(Object).NoSeCae = Val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
    ObjData(Object).SoloVIP = Val(Leer.GetValue("OBJ" & Object, "SoloVIP"))
    
    frmCargando.cargar.value = frmCargando.cargar.value + 1
Next Object

Set Leer = Nothing

Exit Sub

errhandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description


End Sub

Sub LoadUserStats(ByVal userindex As Integer, ByRef UserFile As clsIniReader)

Dim LoopC As Integer


For LoopC = 1 To NUMATRIBUTOS
  UserList(userindex).Stats.UserAtributos(LoopC) = CInt(UserFile.GetValue("ATRIBUTOS", "AT" & LoopC))
  UserList(userindex).Stats.UserAtributosBackUP(LoopC) = UserList(userindex).Stats.UserAtributos(LoopC)
Next LoopC

For LoopC = 1 To NUMSKILLS
  UserList(userindex).Stats.UserSkills(LoopC) = CInt(UserFile.GetValue("SKILLS", "SK" & LoopC))
Next LoopC

For LoopC = 1 To MAXUSERHECHIZOS
  UserList(userindex).Stats.UserHechizos(LoopC) = CInt(UserFile.GetValue("Hechizos", "H" & LoopC))
Next LoopC

Dim tmpStr As String
For LoopC = 1 To MAXUSERQUESTS
    tmpStr = UserFile.GetValue("Quests", "Q" & LoopC)
    UserList(userindex).Stats.UserQuests(LoopC).QuestIndex = CInt(ReadField(1, tmpStr, Asc("-")))
    UserList(userindex).Stats.UserQuests(LoopC).NPCsKilled = CInt(ReadField(2, tmpStr, Asc("-")))
Next LoopC
 
UserList(userindex).Stats.UserQuestsDone = UserFile.GetValue("Quests", "UserQuestsDone")

UserList(userindex).Stats.PuntosTorneo = CLng(UserFile.GetValue("STATS", "PuntosTorneo"))
UserList(userindex).Stats.PuntosDonacion = CLng(UserFile.GetValue("STATS", "PuntosDonacion"))
UserList(userindex).Stats.PuntosVIP = CLng(UserFile.GetValue("STATS", "PuntosVIP"))
UserList(userindex).Stats.RetosGanados = CLng(UserFile.GetValue("STATS", "RetosGanados"))
UserList(userindex).Stats.RetosPerdidos = CLng(UserFile.GetValue("STATS", "RetosPerdidos"))
UserList(userindex).Stats.DuelosGanados = CLng(UserFile.GetValue("STATS", "DuelosGanados"))
UserList(userindex).Stats.DuelosPerdidos = CLng(UserFile.GetValue("STATS", "DuelosPerdidos"))
UserList(userindex).Stats.TrofOro = CLng(UserFile.GetValue("STATS", "TrofOro"))
UserList(userindex).Stats.MedOro = CLng(UserFile.GetValue("STATS", "MedOro"))
UserList(userindex).Stats.TrofPlata = CLng(UserFile.GetValue("STATS", "TrofPlata"))
UserList(userindex).Stats.TrofBronce = CLng(UserFile.GetValue("STATS", "TrofBronce"))

UserList(userindex).Stats.MET = CInt(UserFile.GetValue("STATS", "MET"))
UserList(userindex).Stats.MaxHP = CInt(UserFile.GetValue("STATS", "MaxHP"))
UserList(userindex).Stats.MinHP = CInt(UserFile.GetValue("STATS", "MinHP"))

UserList(userindex).Stats.FIT = CInt(UserFile.GetValue("STATS", "FIT"))
UserList(userindex).Stats.MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
UserList(userindex).Stats.MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))

UserList(userindex).Stats.MaxMan = CInt(UserFile.GetValue("STATS", "MaxMAN"))
UserList(userindex).Stats.MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))

UserList(userindex).Stats.MaxHIT = CInt(UserFile.GetValue("STATS", "MaxHIT"))
UserList(userindex).Stats.MinHIT = CInt(UserFile.GetValue("STATS", "MinHIT"))

UserList(userindex).Stats.SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))

UserList(userindex).Stats.Repu = CLng(UserFile.GetValue("STATS", "Repu"))
UserList(userindex).Stats.Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
UserList(userindex).Stats.ELU = CLng(UserFile.GetValue("STATS", "ELU"))
UserList(userindex).Stats.ELV = CLng(UserFile.GetValue("STATS", "ELV"))


UserList(userindex).Stats.UsuariosMatados = CInt(UserFile.GetValue("MUERTES", "UserMuertes"))
UserList(userindex).Stats.CriminalesMatados = CInt(UserFile.GetValue("MUERTES", "CrimMuertes"))
UserList(userindex).Stats.NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))

'Mithrandir - Sistema de Consejos
UserList(userindex).ConsejoInfo.PertAlCons = CByte(UserFile.GetValue("CONSEJO", "PERTENECE"))
UserList(userindex).ConsejoInfo.LiderConsejo = CByte(UserFile.GetValue("CONSEJO", "LIDERCONSEJO"))
UserList(userindex).ConsejoInfo.PertAlConsCaos = CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS"))
UserList(userindex).ConsejoInfo.LiderConsejoCaos = CByte(UserFile.GetValue("CONSEJO", "LIDERCONSEJOCAOS"))
'Mithrandir - Sistema de Consejos

End Sub

'Mithrandir
Sub LoadUserStatus(ByVal userindex As Integer, ByRef UserFile As clsIniReader)
UserList(userindex).StatusMith.EsStatus = CByte(UserFile.GetValue("STATUS", "EsStatus"))
UserList(userindex).StatusMith.EligioStatus = CByte(UserFile.GetValue("STATUS", "Eligio"))
End Sub

Sub LoadUserReputacion(ByVal userindex As Integer, ByRef UserFile As clsIniReader)

UserList(userindex).Reputacion.AsesinoRep = CDbl(UserFile.GetValue("REP", "Asesino"))
UserList(userindex).Reputacion.BandidoRep = CDbl(UserFile.GetValue("REP", "Bandido"))
UserList(userindex).Reputacion.BurguesRep = CDbl(UserFile.GetValue("REP", "Burguesia"))
UserList(userindex).Reputacion.LadronesRep = CDbl(UserFile.GetValue("REP", "Ladrones"))
UserList(userindex).Reputacion.NobleRep = CDbl(UserFile.GetValue("REP", "Nobles"))
UserList(userindex).Reputacion.PlebeRep = CDbl(UserFile.GetValue("REP", "Plebe"))
UserList(userindex).Reputacion.Promedio = CDbl(UserFile.GetValue("REP", "Promedio"))

End Sub

Sub LoadUserInit(ByVal userindex As Integer, ByRef UserFile As clsIniReader)

Dim LoopC As Long
Dim ln As String

UserList(userindex).Faccion.ArmadaReal = CByte(UserFile.GetValue("FACCIONES", "EjercitoReal"))
UserList(userindex).Faccion.FuerzasCaos = CByte(UserFile.GetValue("FACCIONES", "EjercitoCaos"))
UserList(userindex).Faccion.CiudadanosMatados = CDbl(UserFile.GetValue("FACCIONES", "CiudMatados"))
UserList(userindex).Faccion.CriminalesMatados = CDbl(UserFile.GetValue("FACCIONES", "CrimMatados"))
UserList(userindex).Faccion.NeutralesMatados = CDbl(UserFile.GetValue("FACCIONES", "NeutrMatados"))
UserList(userindex).Faccion.RecibioArmaduraCaos = CByte(UserFile.GetValue("FACCIONES", "rArCaos"))
UserList(userindex).Faccion.RecibioArmaduraReal = CByte(UserFile.GetValue("FACCIONES", "rArReal"))
UserList(userindex).Faccion.RecibioExpInicialCaos = CByte(UserFile.GetValue("FACCIONES", "rExCaos"))
UserList(userindex).Faccion.RecibioExpInicialReal = CByte(UserFile.GetValue("FACCIONES", "rExReal"))
UserList(userindex).Faccion.RecompensasCaos = CLng(UserFile.GetValue("FACCIONES", "recCaos"))
UserList(userindex).Faccion.RecompensasReal = CLng(UserFile.GetValue("FACCIONES", "recReal"))
UserList(userindex).Faccion.Reenlistadas = CByte(UserFile.GetValue("FACCIONES", "Reenlistadas"))

UserList(userindex).flags.Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
'UserList(UserIndex).flags.EleDeFuego = CByte(UserFile.GetValue("FLAGS", "EleDeFuego"))
'UserList(UserIndex).flags.EleDeAgua = CByte(UserFile.GetValue("FLAGS", "EleDeAgua"))
'UserList(UserIndex).flags.EleDeTierra = CByte(UserFile.GetValue("FLAGS", "EleDeTierra"))
UserList(userindex).flags.EnTorneo = CByte(UserFile.GetValue("FLAGS", "EnTorneo"))
UserList(userindex).flags.Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))
UserList(userindex).flags.VIP = CByte(UserFile.GetValue("FLAGS", "VIP"))

UserList(userindex).flags.Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))

UserList(userindex).flags.Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
UserList(userindex).flags.Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
If UserList(userindex).flags.Paralizado = 1 Then
    UserList(userindex).Counters.Paralisis = IntervaloParalizado
End If
UserList(userindex).flags.Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
UserList(userindex).flags.Montando = CByte(UserFile.GetValue("FLAGS", "Montando"))
UserList(userindex).flags.PJerarquia = CByte(UserFile.GetValue("FLAGS", "PJerarquia"))
UserList(userindex).flags.SJerarquia = CByte(UserFile.GetValue("FLAGS", "SJerarquia"))
UserList(userindex).flags.TJerarquia = CByte(UserFile.GetValue("FLAGS", "TJerarquia"))
UserList(userindex).flags.CJerarquia = CByte(UserFile.GetValue("FLAGS", "CJerarquia"))
UserList(userindex).flags.CJerarquiaC = CByte(UserFile.GetValue("FLAGS", "CJerarquiaC"))
UserList(userindex).flags.Transformado = CByte(UserFile.GetValue("FLAGS", "Transformado"))



UserList(userindex).Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))

UserList(userindex).email = UserFile.GetValue("CONTACTO", "Email")
UserList(userindex).Pin = UserFile.GetValue("CONTACTO", "Pin")

UserList(userindex).Genero = UserFile.GetValue("INIT", "Genero")
UserList(userindex).Clase = UserFile.GetValue("INIT", "Clase")
UserList(userindex).Raza = UserFile.GetValue("INIT", "Raza")
UserList(userindex).Hogar = UserFile.GetValue("INIT", "Hogar")
UserList(userindex).Char.Heading = CInt(UserFile.GetValue("INIT", "Heading"))




UserList(userindex).OrigChar.Head = CInt(UserFile.GetValue("INIT", "Head"))
UserList(userindex).OrigChar.Body = CInt(UserFile.GetValue("INIT", "Body"))
UserList(userindex).OrigChar.WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
UserList(userindex).OrigChar.ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
UserList(userindex).OrigChar.CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))
UserList(userindex).OrigChar.Heading = eHeading.SOUTH

If UserList(userindex).flags.Muerto = 0 Then
    UserList(userindex).Char = UserList(userindex).OrigChar
Else
    UserList(userindex).Char.Body = iCuerpoMuerto
    UserList(userindex).Char.Head = iCabezaMuerto
    UserList(userindex).Char.WeaponAnim = NingunArma
    UserList(userindex).Char.ShieldAnim = NingunEscudo
    UserList(userindex).Char.CascoAnim = NingunCasco
End If


UserList(userindex).Desc = UserFile.GetValue("INIT", "Desc")


UserList(userindex).pos.Map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
UserList(userindex).pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
UserList(userindex).pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))

UserList(userindex).Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))

'[KEVIN]--------------------------------------------------------------------
'***********************************************************************************
UserList(userindex).BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))
'Lista de objetos del banco
For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
    ln = UserFile.GetValue("BancoInventory", "Obj" & LoopC)
    UserList(userindex).BancoInvent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
    UserList(userindex).BancoInvent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
Next LoopC
'------------------------------------------------------------------------------------
'[/KEVIN]*****************************************************************************


'Lista de objetos
For LoopC = 1 To MAX_INVENTORY_SLOTS
    ln = UserFile.GetValue("Inventory", "Obj" & LoopC)
    UserList(userindex).Invent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
    UserList(userindex).Invent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
    UserList(userindex).Invent.Object(LoopC).Equipped = CByte(ReadField(3, ln, 45))
Next LoopC

'Obtiene el indice-objeto del arma
UserList(userindex).Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))
If UserList(userindex).Invent.WeaponEqpSlot > 0 Then
    UserList(userindex).Invent.WeaponEqpObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.WeaponEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del armadura
UserList(userindex).Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))
If UserList(userindex).Invent.ArmourEqpSlot > 0 Then
    UserList(userindex).Invent.ArmourEqpObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.ArmourEqpSlot).ObjIndex
    UserList(userindex).flags.Desnudo = 0
Else
    UserList(userindex).flags.Desnudo = 1
End If

'Obtiene el indice-objeto del escudo
UserList(userindex).Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))
If UserList(userindex).Invent.EscudoEqpSlot > 0 Then
    UserList(userindex).Invent.EscudoEqpObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.EscudoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del casco
UserList(userindex).Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))
If UserList(userindex).Invent.CascoEqpSlot > 0 Then
    UserList(userindex).Invent.CascoEqpObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.CascoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto barco
UserList(userindex).Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))
If UserList(userindex).Invent.BarcoSlot > 0 Then
    UserList(userindex).Invent.BarcoObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.BarcoSlot).ObjIndex
End If

'Obtiene el indice-objeto municion
UserList(userindex).Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))
If UserList(userindex).Invent.MunicionEqpSlot > 0 Then
    UserList(userindex).Invent.MunicionEqpObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.MunicionEqpSlot).ObjIndex
End If

'[Alejo]
'Obtiene el indice-objeto herramienta
UserList(userindex).Invent.HerramientaEqpSlot = CInt(UserFile.GetValue("Inventory", "HerramientaSlot"))
If UserList(userindex).Invent.HerramientaEqpSlot > 0 Then
    UserList(userindex).Invent.HerramientaEqpObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.HerramientaEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto montura
UserList(userindex).Invent.MonturaSlot = CInt(UserFile.GetValue("Inventory", "MonturaSlot"))
If UserList(userindex).Invent.MonturaSlot > 0 Then
    UserList(userindex).Invent.MonturaObjIndex = UserList(userindex).Invent.Object(UserList(userindex).Invent.MonturaSlot).ObjIndex
End If

UserList(userindex).NroMacotas = 0

ln = UserFile.GetValue("Guild", "GUILDINDEX")
If IsNumeric(ln) Then
    UserList(userindex).GuildIndex = CInt(ln)
Else
    UserList(userindex).GuildIndex = 0
End If

End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = ""
  
sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
  
  
GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, file
  
GetVar = RTrim$(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()

If frmMain.Visible Then frmMain.txStatus.caption = "Cargando backup."

Dim Map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    NumMaps = Val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
      
    For Map = 1 To NumMaps
        
        If Val(GetVar(App.Path & MapPath & "Mapa" & Map & ".Dat", "Mapa" & Map, "BackUp")) <> 0 Then
            tFileName = App.Path & "\WorldBackUp\Mapa" & Map
        Else
            tFileName = App.Path & MapPath & "Mapa" & Map
        End If
        
        Call CargarMapa(Map, tFileName)
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next Map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)
 
End Sub

Sub LoadMapData()

If frmMain.Visible Then frmMain.txStatus.caption = "Cargando mapas..."

Dim Map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    NumMaps = Val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    Call InitAreas
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
      
    For Map = 1 To NumMaps
        
        tFileName = App.Path & MapPath & "Mapa" & Map
        Call CargarMapa(Map, tFileName)
        
        frmCargando.cargar.value = frmCargando.cargar.value + 1
        DoEvents
    Next Map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

End Sub

Public Sub CargarMapa(ByVal Map As Long, ByVal MAPFl As String)
On Error GoTo errh
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim npcfile As String
    Dim TempInt As Integer
      
    FreeFileMap = FreeFile
    
    Open MAPFl & ".map" For Binary As #FreeFileMap
    Seek FreeFileMap, 1
    
    FreeFileInf = FreeFile
    
    'inf
    Open MAPFl & ".inf" For Binary As #FreeFileInf
    Seek FreeFileInf, 1

    'map Header
    Get #FreeFileMap, , MapInfo(Map).MapVersion
    Get #FreeFileMap, , MiCabecera
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    Get #FreeFileMap, , TempInt
    
    'inf Header
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt
    Get #FreeFileInf, , TempInt

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            '.dat file
            Get FreeFileMap, , ByFlags

            If ByFlags And 1 Then
                MapData(Map, X, Y).Blocked = 1
            End If
            
            Get FreeFileMap, , MapData(Map, X, Y).Graphic(1)
            
            'Layer 2 used?
            If ByFlags And 2 Then Get FreeFileMap, , MapData(Map, X, Y).Graphic(2)
            
            'Layer 3 used?
            If ByFlags And 4 Then Get FreeFileMap, , MapData(Map, X, Y).Graphic(3)
            
            'Layer 4 used?
            If ByFlags And 8 Then Get FreeFileMap, , MapData(Map, X, Y).Graphic(4)
            
            'Trigger used?
            If ByFlags And 16 Then
                'Enums are 4 byte long in VB, so we make sure we only read 2
                Get FreeFileMap, , TempInt
                MapData(Map, X, Y).trigger = TempInt
            End If
            
            Get FreeFileInf, , ByFlags
            
            If ByFlags And 1 Then
                Get FreeFileInf, , MapData(Map, X, Y).TileExit.Map
                Get FreeFileInf, , MapData(Map, X, Y).TileExit.X
                Get FreeFileInf, , MapData(Map, X, Y).TileExit.Y
            End If
            
            If ByFlags And 2 Then
                'Get and make NPC
                Get FreeFileInf, , MapData(Map, X, Y).NpcIndex
                
                If MapData(Map, X, Y).NpcIndex > 0 Then
                    If MapData(Map, X, Y).NpcIndex > 499 Then
                        npcfile = DatPath & "NPCs-HOSTILES.dat"
                    Else
                        npcfile = DatPath & "NPCs.dat"
                    End If

                    'Si el npc debe hacer respawn en la pos
                    'original la guardamos
                    If Val(GetVar(npcfile, "NPC" & MapData(Map, X, Y).NpcIndex, "PosOrig")) = 1 Then
                        MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.Map = Map
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.X = X
                        Npclist(MapData(Map, X, Y).NpcIndex).Orig.Y = Y
                    Else
                        MapData(Map, X, Y).NpcIndex = OpenNPC(MapData(Map, X, Y).NpcIndex)
                    End If
                            
                    Npclist(MapData(Map, X, Y).NpcIndex).pos.Map = Map
                    Npclist(MapData(Map, X, Y).NpcIndex).pos.X = X
                    Npclist(MapData(Map, X, Y).NpcIndex).pos.Y = Y
                            
                    Call MakeNPCChar(SendTarget.ToMap, 0, 0, MapData(Map, X, Y).NpcIndex, 1, 1, 1)
                End If
            End If
            
            If ByFlags And 4 Then
                'Get and make Object
                Get FreeFileInf, , MapData(Map, X, Y).OBJInfo.ObjIndex
                Get FreeFileInf, , MapData(Map, X, Y).OBJInfo.Amount
            End If
        Next X
    Next Y
    
    
    Close FreeFileMap
    Close FreeFileInf
    
    MapInfo(Map).name = GetVar(MAPFl & ".dat", "Mapa" & Map, "Name")
    MapInfo(Map).Music = GetVar(MAPFl & ".dat", "Mapa" & Map, "MusicNum")
    MapInfo(Map).StartPos.Map = Val(ReadField(1, GetVar(MAPFl & ".dat", "Mapa" & Map, "StartPos"), Asc("-")))
    MapInfo(Map).StartPos.X = Val(ReadField(2, GetVar(MAPFl & ".dat", "Mapa" & Map, "StartPos"), Asc("-")))
    MapInfo(Map).StartPos.Y = Val(ReadField(3, GetVar(MAPFl & ".dat", "Mapa" & Map, "StartPos"), Asc("-")))
    MapInfo(Map).MagiaSinEfecto = Val(GetVar(MAPFl & ".dat", "Mapa" & Map, "MagiaSinEfecto"))
    MapInfo(Map).NoEncriptarMP = Val(GetVar(MAPFl & ".dat", "Mapa" & Map, "NoEncriptarMP"))
    
    If Val(GetVar(MAPFl & ".dat", "Mapa" & Map, "Pk")) = 0 Then
        MapInfo(Map).Pk = True
    Else
        MapInfo(Map).Pk = False
    End If
    
    
    MapInfo(Map).Terreno = GetVar(MAPFl & ".dat", "Mapa" & Map, "Terreno")
    MapInfo(Map).Zona = GetVar(MAPFl & ".dat", "Mapa" & Map, "Zona")
    MapInfo(Map).Restringir = GetVar(MAPFl & ".dat", "Mapa" & Map, "Restringir")
    MapInfo(Map).BackUp = Val(GetVar(MAPFl & ".dat", "Mapa" & Map, "BACKUP"))
Exit Sub

errh:
    Call LogError("Error cargando mapa: " & Map & "." & Err.Description)
End Sub

Sub LoadSini()

Dim Temporal As Long
Dim Temporal1 As Long
Dim LoopC As Integer

If frmMain.Visible Then frmMain.txStatus.caption = "Cargando info de inicio del server."

BootDelBackUp = Val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))

Puerto = Val(GetVar(IniPath & "Server.ini", "INIT", "Puerto"))
HideMe = Val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
AllowMultiLogins = Val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
IdleLimit = Val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
'Lee la version correcta del cliente
ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")

PuedeCrearPersonajes = Val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
CamaraLenta = Val(GetVar(IniPath & "Server.ini", "INIT", "CamaraLenta"))
ServerSoloGMs = Val(GetVar(IniPath & "Server.ini", "init", "ServerSoloGMs"))

'ArmaduraImperial1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial1"))

                                     'Imperial
'Bajos
GuerreroBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "GuerreroBajosArmadura1"))
MagoBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "MagoBajosArmadura1"))
PaladinBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "PaladinBajosArmadura1"))
ClerigoBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "ClerigoBajosArmadura1"))
BardoBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "BardoBajosArmadura1"))
AsesinoBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "AsesinoBajosArmadura1"))
DruidaBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "DruidaBajosArmadura1"))
CazadorBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CazadorBajosArmadura1"))
OtrasBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "OtrasBajosArmadura1"))

GuerreroBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "GuerreroBajosArmadura2"))
MagoBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "MagoBajosArmadura2"))
PaladinBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "PaladinBajosArmadura2"))
ClerigoBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "ClerigoBajosArmadura2"))
BardoBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "BardoBajosArmadura2"))
AsesinoBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "AsesinoBajosArmadura2"))
DruidaBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "DruidaBajosArmadura2"))
CazadorBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CazadorBajosArmadura2"))
OtrasBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "OtrasBajosArmadura2"))

GuerreroBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "GuerreroBajosArmadura3"))
MagoBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "MagoBajosArmadura3"))
PaladinBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "PaladinBajosArmadura3"))
ClerigoBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "ClerigoBajosArmadura3"))
BardoBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "BardoBajosArmadura3"))
AsesinoBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "AsesinoBajosArmadura3"))
DruidaBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "DruidaBajosArmadura3"))
CazadorBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CazadorBajosArmadura3"))
OtrasBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "OtrasBajosArmadura3"))

GuerreroBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "GuerreroBajosArmadura4"))
MagoBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "MagoBajosArmadura4"))
PaladinBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "PaladinBajosArmadura4"))
ClerigoBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "ClerigoBajosArmadura4"))
BardoBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "BardoBajosArmadura4"))
AsesinoBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "AsesinoBajosArmadura4"))
DruidaBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "DruidaBajosArmadura4"))
CazadorBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CazadorBajosArmadura4"))
OtrasBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "OtrasBajosArmadura4"))

'Altos
GuerreroAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "GuerreroAltosArmadura1"))
MagoAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "MagoAltosArmadura1"))
PaladinAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "PaladinAltosArmadura1"))
ClerigoAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "ClerigoAltosArmadura1"))
BardoAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "BardoAltosArmadura1"))
AsesinoAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "AsesinoAltosArmadura1"))
DruidaAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "DruidaAltosArmadura1"))
CazadorAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CazadorAltosArmadura1"))
OtrasAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "OtrasAltosArmadura1"))

GuerreroAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "GuerreroAltosArmadura2"))
MagoAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "MagoAltosArmadura2"))
PaladinAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "PaladinAltosArmadura2"))
ClerigoAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "ClerigoAltosArmadura2"))
BardoAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "BardoAltosArmadura2"))
AsesinoAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "AsesinoAltosArmadura2"))
DruidaAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "DruidaAltosArmadura2"))
CazadorAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CazadorAltosArmadura2"))
OtrasAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "OtrasAltosArmadura2"))

GuerreroAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "GuerreroAltosArmadura3"))
MagoAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "MagoAltosArmadura3"))
PaladinAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "PaladinAltosArmadura3"))
ClerigoAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "ClerigoAltosArmadura3"))
BardoAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "BardoAltosArmadura3"))
AsesinoAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "AsesinoAltosArmadura3"))
DruidaAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "DruidaAltosArmadura3"))
CazadorAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CazadorAltosArmadura3"))
OtrasAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "OtrasAltosArmadura3"))

GuerreroAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "GuerreroAltosArmadura4"))
MagoAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "MagoAltosArmadura4"))
PaladinAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "PaladinAltosArmadura4"))
ClerigoAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "ClerigoAltosArmadura4"))
BardoAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "BardoAltosArmadura4"))
AsesinoAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "AsesinoAltosArmadura4"))
DruidaAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "DruidaAltosArmadura4"))
CazadorAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CazadorAltosArmadura4"))
OtrasAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "OtrasAltosArmadura4"))

                                     'Caotico
'Bajos
CaosGuerreroBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosGuerreroBajosArmadura1"))
CaosMagoBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosMagoBajosArmadura1"))
CaosPaladinBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosPaladinBajosArmadura1"))
CaosClerigoBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosClerigoBajosArmadura1"))
CaosBardoBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosBardoBajosArmadura1"))
CaosAsesinoBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosAsesinoBajosArmadura1"))
CaosDruidaBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosDruidaBajosArmadura1"))
CaosCazadorBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosCazadorBajosArmadura1"))
CaosOtrasBajosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosOtrasBajosArmadura1"))

CaosGuerreroBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "GCaosuerreroBajosArmadura2"))
CaosMagoBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosMagoBajosArmadura2"))
CaosPaladinBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosPaladinBajosArmadura2"))
CaosClerigoBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosClerigoBajosArmadura2"))
CaosBardoBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosBardoBajosArmadura2"))
CaosAsesinoBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosAsesinoBajosArmadura2"))
CaosDruidaBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosDruidaBajosArmadura2"))
CaosCazadorBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosCazadorBajosArmadura2"))
CaosOtrasBajosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosOtrasBajosArmadura2"))

CaosGuerreroBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosGuerreroBajosArmadura3"))
CaosMagoBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosMagoBajosArmadura3"))
CaosPaladinBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosPaladinBajosArmadura3"))
CaosClerigoBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosClerigoBajosArmadura3"))
CaosBardoBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosBardoBajosArmadura3"))
CaosAsesinoBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosAsesinoBajosArmadura3"))
CaosDruidaBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosDruidaBajosArmadura3"))
CaosCazadorBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosCazadorBajosArmadura3"))
CaosOtrasBajosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosOtrasBajosArmadura3"))

CaosGuerreroBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosGuerreroBajosArmadura4"))
CaosMagoBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosMagoBajosArmadura4"))
CaosPaladinBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosPaladinBajosArmadura4"))
CaosClerigoBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosClerigoBajosArmadura4"))
CaosBardoBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosBardoBajosArmadura4"))
CaosAsesinoBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosAsesinoBajosArmadura4"))
CaosDruidaBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosDruidaBajosArmadura4"))
CaosCazadorBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosCazadorBajosArmadura4"))
CaosOtrasBajosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosOtrasBajosArmadura4"))

'Altos
CaosGuerreroAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosGuerreroAltosArmadura1"))
CaosMagoAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosMagoAltosArmadura1"))
CaosPaladinAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosPaladinAltosArmadura1"))
CaosClerigoAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosClerigoAltosArmadura1"))
CaosBardoAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosBardoAltosArmadura1"))
CaosAsesinoAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosAsesinoAltosArmadura1"))
CaosDruidaAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosDruidaAltosArmadura1"))
CaosCazadorAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosCazadorAltosArmadura1"))
CaosOtrasAltosArmadura1 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosOtrasAltosArmadura1"))

CaosGuerreroAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosGuerreroAltosArmadura2"))
CaosMagoAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosMagoAltosArmadura2"))
CaosPaladinAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosPaladinAltosArmadura2"))
CaosClerigoAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosClerigoAltosArmadura2"))
CaosBardoAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosBardoAltosArmadura2"))
CaosAsesinoAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosAsesinoAltosArmadura2"))
CaosDruidaAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosDruidaAltosArmadura2"))
CaosCazadorAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosCazadorAltosArmadura2"))
CaosOtrasAltosArmadura2 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosOtrasAltosArmadura2"))

CaosGuerreroAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosGuerreroAltosArmadura3"))
CaosMagoAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosMagoAltosArmadura3"))
CaosPaladinAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosPaladinAltosArmadura3"))
CaosClerigoAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosClerigoAltosArmadura3"))
CaosBardoAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosBardoAltosArmadura3"))
CaosAsesinoAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosAsesinoAltosArmadura3"))
CaosDruidaAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosDruidaAltosArmadura3"))
CaosCazadorAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosCazadorAltosArmadura3"))
CaosOtrasAltosArmadura3 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosOtrasAltosArmadura3"))

CaosGuerreroAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosGuerreroAltosArmadura4"))
CaosMagoAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosMagoAltosArmadura4"))
CaosPaladinAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosPaladinAltosArmadura4"))
CaosClerigoAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosClerigoAltosArmadura4"))
CaosBardoAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosBardoAltosArmadura4"))
CaosAsesinoAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosAsesinoAltosArmadura4"))
CaosDruidaAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosDruidaAltosArmadura4"))
CaosCazadorAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosCazadorAltosArmadura4"))
CaosOtrasAltosArmadura4 = Val(GetVar(IniPath & "Server.ini", "INIT", "CaosOtrasAltosArmadura4"))

MAPA_PRETORIANO = Val(GetVar(IniPath & "Server.ini", "INIT", "MapaPretoriano"))

ClientsCommandsQueue = Val(GetVar(IniPath & "Server.ini", "INIT", "ClientsCommandsQueue"))
EnTesting = Val(GetVar(IniPath & "Server.ini", "INIT", "Testing"))
EncriptarProtocolosCriticos = Val(GetVar(IniPath & "Server.ini", "INIT", "Encriptar"))

'Start pos
StartPos.Map = Val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
StartPos.X = Val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))
StartPos.Y = Val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "StartPos"), 45))

'Intervalos
SanaIntervaloSinDescansar = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar

StaminaIntervaloSinDescansar = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar

SanaIntervaloDescansar = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar

StaminaIntervaloDescansar = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar

IntervaloVeneno = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno

IntervaloParalizado = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado

IntervaloInvisible = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible

IntervaloFrio = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
FrmInterv.txtIntervaloFrio.Text = IntervaloFrio

IntervaloWavFx = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx

IntervaloInvocacion = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
FrmInterv.txtInvocacion.Text = IntervaloInvocacion

IntervaloParaConexion = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion

'&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&


IntervaloUserPuedeCastear = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear

frmMain.TIMER_AI.Interval = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))
FrmInterv.txtAI.Text = frmMain.TIMER_AI.Interval

frmMain.npcataca.Interval = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.Interval

IntervaloUserPuedeTrabajar = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar

IntervaloUserPuedeAtacar = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar


frmMain.CmdExec.Interval = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTimerExec"))
FrmInterv.txtCmdExec.Text = frmMain.CmdExec.Interval

MinutosWs = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))
If MinutosWs < 60 Then MinutosWs = 180

IntervaloCerrarConexion = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
IntervaloUserPuedeUsar = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))
IntervaloFlechasCazadores = Val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechasCazadores"))


'Ressurect pos
ResPos.Map = Val(ReadField(1, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.X = Val(ReadField(2, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
ResPos.Y = Val(ReadField(3, GetVar(IniPath & "Server.ini", "INIT", "ResPos"), 45))
  
recordusuarios = Val(GetVar(IniPath & "Server.ini", "INIT", "Record"))
  
'Max users
Temporal = Val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
If MaxUsers = 0 Then
    MaxUsers = Temporal
    ReDim UserList(1 To MaxUsers) As User
End If

Helkat.Map = GetVar(DatPath & "Ciudades.dat", "Helkat", "Mapa")
Helkat.X = GetVar(DatPath & "Ciudades.dat", "Helkat", "X")
Helkat.Y = GetVar(DatPath & "Ciudades.dat", "Helkat", "Y")

Runek.Map = GetVar(DatPath & "Ciudades.dat", "Runek", "Mapa")
Runek.X = GetVar(DatPath & "Ciudades.dat", "Runek", "X")
Runek.Y = GetVar(DatPath & "Ciudades.dat", "Runek", "Y")

Banderbill.Map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")

Lindos.Map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")

Call ConsultaPopular.LoadData

#If SeguridadAlkon Then
Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
#End If

LoadAntiCheat

End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, value, file
    
End Sub

Sub SaveUser(ByVal userindex As Integer, ByVal UserFile As String)
On Error GoTo errhandler

Dim OldUserHead As Long


'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
If UserList(userindex).Clase = "" Or UserList(userindex).Stats.ELV = 0 Then
    Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(userindex).name)
    Exit Sub
End If


If UserList(userindex).flags.Mimetizado = 1 Then
    UserList(userindex).Char.Body = UserList(userindex).CharMimetizado.Body
    UserList(userindex).Char.Head = UserList(userindex).CharMimetizado.Head
    UserList(userindex).Char.CascoAnim = UserList(userindex).CharMimetizado.CascoAnim
    UserList(userindex).Char.ShieldAnim = UserList(userindex).CharMimetizado.ShieldAnim
    UserList(userindex).Char.WeaponAnim = UserList(userindex).CharMimetizado.WeaponAnim
    UserList(userindex).Counters.Mimetismo = 0
    UserList(userindex).flags.Mimetizado = 0
End If



If FileExist(UserFile, vbNormal) Then
       If UserList(userindex).flags.Muerto = 1 Then
        OldUserHead = UserList(userindex).Char.Head
        UserList(userindex).Char.Head = CStr(GetVar(UserFile, "INIT", "Head"))
       End If
'       Kill UserFile
End If

Dim LoopC As Integer


Call WriteVar(UserFile, "FLAGS", "Muerto", CStr(UserList(userindex).flags.Muerto))
'Call WriteVar(UserFile, "FLAGS", "EleDeFuego", CStr(UserList(UserIndex).flags.EleDeFuego))
'Call WriteVar(UserFile, "FLAGS", "EleDeAgua", CStr(UserList(UserIndex).flags.EleDeAgua))
'Call WriteVar(UserFile, "FLAGS", "EleDeTierra", CStr(UserList(UserIndex).flags.EleDeTierra))
Call WriteVar(UserFile, "FLAGS", "EnTorneo", CStr(UserList(userindex).flags.EnTorneo))
Call WriteVar(UserFile, "FLAGS", "Escondido", CStr(UserList(userindex).flags.Escondido))
Call WriteVar(UserFile, "FLAGS", "VIP", CStr(UserList(userindex).flags.VIP))
Call WriteVar(UserFile, "FLAGS", "Desnudo", CStr(UserList(userindex).flags.Desnudo))
Call WriteVar(UserFile, "FLAGS", "Ban", CStr(UserList(userindex).flags.Ban))
Call WriteVar(UserFile, "FLAGS", "Navegando", CStr(UserList(userindex).flags.Navegando))
Call WriteVar(UserFile, "FLAGS", "Montando", CStr(UserList(userindex).flags.Montando))
Call WriteVar(UserFile, "FLAGS", "PJerarquia", CStr(UserList(userindex).flags.PJerarquia))
Call WriteVar(UserFile, "FLAGS", "SJerarquia", CStr(UserList(userindex).flags.SJerarquia))
Call WriteVar(UserFile, "FLAGS", "TJerarquia", CStr(UserList(userindex).flags.TJerarquia))
Call WriteVar(UserFile, "FLAGS", "CJerarquia", CStr(UserList(userindex).flags.CJerarquia))
Call WriteVar(UserFile, "FLAGS", "CJerarquiaC", CStr(UserList(userindex).flags.CJerarquiaC))
Call WriteVar(UserFile, "FLAGS", "Transformado", CStr(UserList(userindex).flags.Transformado))


Call WriteVar(UserFile, "FLAGS", "Envenenado", CStr(UserList(userindex).flags.Envenenado))
Call WriteVar(UserFile, "FLAGS", "Paralizado", CStr(UserList(userindex).flags.Paralizado))

'Mithrandir - Sistema de Consejos
Call WriteVar(UserFile, "CONSEJO", "PERTENECE", CStr(UserList(userindex).ConsejoInfo.PertAlCons))
Call WriteVar(UserFile, "CONSEJO", "LIDERCONSEJO", CStr(UserList(userindex).ConsejoInfo.LiderConsejo))
Call WriteVar(UserFile, "CONSEJO", "PERTENECECAOS", CStr(UserList(userindex).ConsejoInfo.PertAlConsCaos))
Call WriteVar(UserFile, "CONSEJO", "LIDERCONSEJOCAOS", CStr(UserList(userindex).ConsejoInfo.LiderConsejoCaos))
'Mithrandir - Sistema de Consejos


Call WriteVar(UserFile, "COUNTERS", "Pena", CStr(UserList(userindex).Counters.Pena))

Call WriteVar(UserFile, "FACCIONES", "EjercitoReal", CStr(UserList(userindex).Faccion.ArmadaReal))
Call WriteVar(UserFile, "FACCIONES", "EjercitoCaos", CStr(UserList(userindex).Faccion.FuerzasCaos))
Call WriteVar(UserFile, "FACCIONES", "CiudMatados", CStr(UserList(userindex).Faccion.CiudadanosMatados))
Call WriteVar(UserFile, "FACCIONES", "CrimMatados", CStr(UserList(userindex).Faccion.CriminalesMatados))
Call WriteVar(UserFile, "FACCIONES", "NeutrMatados", CStr(UserList(userindex).Faccion.NeutralesMatados))
Call WriteVar(UserFile, "FACCIONES", "rArCaos", CStr(UserList(userindex).Faccion.RecibioArmaduraCaos))
Call WriteVar(UserFile, "FACCIONES", "rArReal", CStr(UserList(userindex).Faccion.RecibioArmaduraReal))
Call WriteVar(UserFile, "FACCIONES", "rExCaos", CStr(UserList(userindex).Faccion.RecibioExpInicialCaos))
Call WriteVar(UserFile, "FACCIONES", "rExReal", CStr(UserList(userindex).Faccion.RecibioExpInicialReal))
Call WriteVar(UserFile, "FACCIONES", "recCaos", CStr(UserList(userindex).Faccion.RecompensasCaos))
Call WriteVar(UserFile, "FACCIONES", "recReal", CStr(UserList(userindex).Faccion.RecompensasReal))
Call WriteVar(UserFile, "FACCIONES", "Reenlistadas", CStr(UserList(userindex).Faccion.Reenlistadas))


'¿Fueron modificados los atributos del usuario?
If Not UserList(userindex).flags.TomoPocion Then
    For LoopC = 1 To UBound(UserList(userindex).Stats.UserAtributos)
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, CStr(UserList(userindex).Stats.UserAtributos(LoopC)))
    Next
Else
    For LoopC = 1 To UBound(UserList(userindex).Stats.UserAtributos)
        'UserList(UserIndex).Stats.UserAtributos(LoopC) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)
        Call WriteVar(UserFile, "ATRIBUTOS", "AT" & LoopC, CStr(UserList(userindex).Stats.UserAtributosBackUP(LoopC)))
    Next
End If

For LoopC = 1 To UBound(UserList(userindex).Stats.UserSkills)
    Call WriteVar(UserFile, "SKILLS", "SK" & LoopC, CStr(UserList(userindex).Stats.UserSkills(LoopC)))
Next


Call WriteVar(UserFile, "CONTACTO", "Email", UserList(userindex).email)
Call WriteVar(UserFile, "CONTACTO", "Pin", UserList(userindex).Pin)

Call WriteVar(UserFile, "INIT", "Genero", UserList(userindex).Genero)
Call WriteVar(UserFile, "INIT", "Raza", UserList(userindex).Raza)
Call WriteVar(UserFile, "INIT", "Hogar", UserList(userindex).Hogar)
Call WriteVar(UserFile, "INIT", "Clase", UserList(userindex).Clase)
Call WriteVar(UserFile, "INIT", "Password", UserList(userindex).PassWord)
Call WriteVar(UserFile, "INIT", "Desc", UserList(userindex).Desc)

Call WriteVar(UserFile, "INIT", "Heading", CStr(UserList(userindex).Char.Heading))

Call WriteVar(UserFile, "INIT", "Head", CStr(UserList(userindex).OrigChar.Head))

If UserList(userindex).flags.Muerto = 0 Then
    Call WriteVar(UserFile, "INIT", "Body", CStr(UserList(userindex).Char.Body))
End If

Call WriteVar(UserFile, "INIT", "Arma", CStr(UserList(userindex).Char.WeaponAnim))
Call WriteVar(UserFile, "INIT", "Escudo", CStr(UserList(userindex).Char.ShieldAnim))
Call WriteVar(UserFile, "INIT", "Casco", CStr(UserList(userindex).Char.CascoAnim))

Call WriteVar(UserFile, "INIT", "LastIP", UserList(userindex).ip)
Call WriteVar(UserFile, "INIT", "Position", UserList(userindex).pos.Map & "-" & UserList(userindex).pos.X & "-" & UserList(userindex).pos.Y)
Call WriteVar(UserFile, "INIT", "LastHD", UserList(userindex).hd)

Call WriteVar(UserFile, "STATS", "PuntosTorneo", CStr(UserList(userindex).Stats.PuntosTorneo))
Call WriteVar(UserFile, "STATS", "PuntosDonacion", CStr(UserList(userindex).Stats.PuntosDonacion))
Call WriteVar(UserFile, "STATS", "PuntosVIP", CStr(UserList(userindex).Stats.PuntosVIP))
Call WriteVar(UserFile, "STATS", "RetosGanados", CStr(UserList(userindex).Stats.RetosGanados))
Call WriteVar(UserFile, "STATS", "RetosPerdidos", CStr(UserList(userindex).Stats.RetosPerdidos))
Call WriteVar(UserFile, "STATS", "DuelosGanados", CStr(UserList(userindex).Stats.DuelosGanados))
Call WriteVar(UserFile, "STATS", "DuelosPerdidos", CStr(UserList(userindex).Stats.DuelosPerdidos))
Call WriteVar(UserFile, "STATS", "TrofOro", CStr(UserList(userindex).Stats.TrofOro))
Call WriteVar(UserFile, "STATS", "MedOro", CStr(UserList(userindex).Stats.MedOro))
Call WriteVar(UserFile, "STATS", "TrofPlata", CStr(UserList(userindex).Stats.TrofPlata))
Call WriteVar(UserFile, "STATS", "TrofBronce", CStr(UserList(userindex).Stats.TrofBronce))

Call WriteVar(UserFile, "STATS", "MET", CStr(UserList(userindex).Stats.MET))
Call WriteVar(UserFile, "STATS", "MaxHP", CStr(UserList(userindex).Stats.MaxHP))
Call WriteVar(UserFile, "STATS", "MinHP", CStr(UserList(userindex).Stats.MinHP))

Call WriteVar(UserFile, "STATS", "FIT", CStr(UserList(userindex).Stats.FIT))
Call WriteVar(UserFile, "STATS", "MaxSTA", CStr(UserList(userindex).Stats.MaxSta))
Call WriteVar(UserFile, "STATS", "MinSTA", CStr(UserList(userindex).Stats.MinSta))

Call WriteVar(UserFile, "STATS", "MaxMAN", CStr(UserList(userindex).Stats.MaxMan))
Call WriteVar(UserFile, "STATS", "MinMAN", CStr(UserList(userindex).Stats.MinMAN))

Call WriteVar(UserFile, "STATS", "MaxHIT", CStr(UserList(userindex).Stats.MaxHIT))
Call WriteVar(UserFile, "STATS", "MinHIT", CStr(UserList(userindex).Stats.MinHIT))

Call WriteVar(UserFile, "STATS", "SkillPtsLibres", CStr(UserList(userindex).Stats.SkillPts))

'STATUS - MITHRANDIR
Call WriteVar(UserFile, "STATUS", "EsStatus", CStr(UserList(userindex).StatusMith.EsStatus))
Call WriteVar(UserFile, "STATUS", "Eligio", CStr(UserList(userindex).StatusMith.EligioStatus))
'STATUS - MITHRANDIR
  
Call WriteVar(UserFile, "STATS", "Repu", CStr(UserList(userindex).Stats.Repu))
Call WriteVar(UserFile, "STATS", "EXP", CStr(UserList(userindex).Stats.Exp))
Call WriteVar(UserFile, "STATS", "ELV", CStr(UserList(userindex).Stats.ELV))


Dim tmpInt As Integer
For tmpInt = 1 To MAXUSERQUESTS
    Call WriteVar(UserFile, "Quests", "Q" & tmpInt, UserList(userindex).Stats.UserQuests(tmpInt).QuestIndex & "-" & UserList(userindex).Stats.UserQuests(tmpInt).NPCsKilled)
Next tmpInt
 
Call WriteVar(UserFile, "Quests", "UserQuestsDone", UserList(userindex).Stats.UserQuestsDone)
 



Call WriteVar(UserFile, "STATS", "ELU", CStr(UserList(userindex).Stats.ELU))
Call WriteVar(UserFile, "MUERTES", "UserMuertes", CStr(UserList(userindex).Stats.UsuariosMatados))
Call WriteVar(UserFile, "MUERTES", "CrimMuertes", CStr(UserList(userindex).Stats.CriminalesMatados))
Call WriteVar(UserFile, "MUERTES", "NpcsMuertes", CStr(UserList(userindex).Stats.NPCsMuertos))
  
'[KEVIN]----------------------------------------------------------------------------
'*******************************************************************************************
Call WriteVar(UserFile, "BancoInventory", "CantidadItems", Val(UserList(userindex).BancoInvent.NroItems))
Dim loopd As Integer
For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
    Call WriteVar(UserFile, "BancoInventory", "Obj" & loopd, UserList(userindex).BancoInvent.Object(loopd).ObjIndex & "-" & UserList(userindex).BancoInvent.Object(loopd).Amount)
Next loopd
'*******************************************************************************************
'[/KEVIN]-----------
  
'Save Inv
Call WriteVar(UserFile, "Inventory", "CantidadItems", Val(UserList(userindex).Invent.NroItems))

For LoopC = 1 To MAX_INVENTORY_SLOTS
    Call WriteVar(UserFile, "Inventory", "Obj" & LoopC, UserList(userindex).Invent.Object(LoopC).ObjIndex & "-" & UserList(userindex).Invent.Object(LoopC).Amount & "-" & UserList(userindex).Invent.Object(LoopC).Equipped)
Next

Call WriteVar(UserFile, "Inventory", "WeaponEqpSlot", str(UserList(userindex).Invent.WeaponEqpSlot))
Call WriteVar(UserFile, "Inventory", "ArmourEqpSlot", str(UserList(userindex).Invent.ArmourEqpSlot))
Call WriteVar(UserFile, "Inventory", "CascoEqpSlot", str(UserList(userindex).Invent.CascoEqpSlot))
Call WriteVar(UserFile, "Inventory", "EscudoEqpSlot", str(UserList(userindex).Invent.EscudoEqpSlot))
Call WriteVar(UserFile, "Inventory", "BarcoSlot", str(UserList(userindex).Invent.BarcoSlot))
Call WriteVar(UserFile, "Inventory", "MonturaSlot", str(UserList(userindex).Invent.MonturaSlot))
Call WriteVar(UserFile, "Inventory", "MunicionSlot", str(UserList(userindex).Invent.MunicionEqpSlot))
Call WriteVar(UserFile, "Inventory", "HerramientaSlot", str(UserList(userindex).Invent.HerramientaEqpSlot))


'Reputacion
Call WriteVar(UserFile, "REP", "Asesino", Val(UserList(userindex).Reputacion.AsesinoRep))
Call WriteVar(UserFile, "REP", "Bandido", Val(UserList(userindex).Reputacion.BandidoRep))
Call WriteVar(UserFile, "REP", "Burguesia", Val(UserList(userindex).Reputacion.BurguesRep))
Call WriteVar(UserFile, "REP", "Ladrones", Val(UserList(userindex).Reputacion.LadronesRep))
Call WriteVar(UserFile, "REP", "Nobles", Val(UserList(userindex).Reputacion.NobleRep))
Call WriteVar(UserFile, "REP", "Plebe", Val(UserList(userindex).Reputacion.PlebeRep))

Dim L As Long
L = (-UserList(userindex).Reputacion.AsesinoRep) + _
    (-UserList(userindex).Reputacion.BandidoRep) + _
    UserList(userindex).Reputacion.BurguesRep + _
    (-UserList(userindex).Reputacion.LadronesRep) + _
    UserList(userindex).Reputacion.NobleRep + _
    UserList(userindex).Reputacion.PlebeRep
L = L / 6
Call WriteVar(UserFile, "REP", "Promedio", Val(L))

Dim cad As String

For LoopC = 1 To MAXUSERHECHIZOS
    cad = UserList(userindex).Stats.UserHechizos(LoopC)
    Call WriteVar(UserFile, "HECHIZOS", "H" & LoopC, cad)
Next

Dim NroMascotas As Long
NroMascotas = UserList(userindex).NroMacotas

For LoopC = 1 To MAXMASCOTAS
    ' Mascota valida?
    If UserList(userindex).MascotasIndex(LoopC) > 0 Then
        ' Nos aseguramos que la criatura no fue invocada
        If Npclist(UserList(userindex).MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
            cad = UserList(userindex).MascotasType(LoopC)
        Else 'Si fue invocada no la guardamos
            cad = "0"
            NroMascotas = NroMascotas - 1
        End If
        Call WriteVar(UserFile, "MASCOTAS", "MAS" & LoopC, cad)
    End If

Next

Call WriteVar(UserFile, "MASCOTAS", "NroMascotas", str(NroMascotas))

'Devuelve el head de muerto
If UserList(userindex).flags.Muerto = 1 Then
    UserList(userindex).Char.Head = iCabezaMuerto
End If

Exit Sub

errhandler:
Call LogError("Error en SaveUser")

End Sub

'Newbie - O no eligio
Function Neutral(ByVal userindex As Integer) As Boolean
Neutral = UserList(userindex).StatusMith.EsStatus = 0
End Function
Function TransformadoVIP(ByVal userindex As Integer) As Boolean
TransformadoVIP = UserList(userindex).Stats.TransformadoVIP = 1
End Function
'Ciudadano
Function Ciudadano(ByVal userindex As Integer) As Boolean
Ciudadano = UserList(userindex).StatusMith.EsStatus = 1 Or UserList(userindex).StatusMith.EsStatus = 3
End Function
'Criminal
Function Criminal(ByVal userindex As Integer) As Boolean
Criminal = UserList(userindex).StatusMith.EsStatus = 2 Or UserList(userindex).StatusMith.EsStatus = 4
End Function

Sub BackUPnPc(NpcIndex As Integer)

Dim NpcNumero As Integer
Dim npcfile As String
Dim LoopC As Integer


NpcNumero = Npclist(NpcIndex).Numero

If NpcNumero > 499 Then
    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "bkNPCs.dat"
End If

'General
Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", Npclist(NpcIndex).name)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", Npclist(NpcIndex).Desc)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", Val(Npclist(NpcIndex).Char.Head))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", Val(Npclist(NpcIndex).Char.Body))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", Val(Npclist(NpcIndex).Char.Heading))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", Val(Npclist(NpcIndex).Movement))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", Val(Npclist(NpcIndex).Attackable))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", Val(Npclist(NpcIndex).Comercia))
Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", Val(Npclist(NpcIndex).TipoItems))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", Val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", Val(Npclist(NpcIndex).GiveEXP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", Val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", Val(Npclist(NpcIndex).InvReSpawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", Val(Npclist(NpcIndex).NPCtype))
Call WriteVar(npcfile, "NPC" & NpcNumero, "QuestNumber", Val(Npclist(NpcIndex).QuestNumber))
Call WriteVar(npcfile, "NPC" & NpcNumero, "TalkAfterQuest", Npclist(NpcIndex).TalkAfterQuest)
Call WriteVar(npcfile, "NPC" & NpcNumero, "TalkDuringQuest", Npclist(NpcIndex).TalkDuringQuest)


'Stats
Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", Val(Npclist(NpcIndex).Stats.Alineacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", Val(Npclist(NpcIndex).Stats.def))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", Val(Npclist(NpcIndex).Stats.MaxHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", Val(Npclist(NpcIndex).Stats.MaxHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", Val(Npclist(NpcIndex).Stats.MinHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", Val(Npclist(NpcIndex).Stats.MinHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", Val(Npclist(NpcIndex).Stats.UsuariosMatados))




'Flags
Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", Val(Npclist(NpcIndex).flags.Respawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", Val(Npclist(NpcIndex).flags.BackUp))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", Val(Npclist(NpcIndex).flags.Domable))

'Inventario
Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", Val(Npclist(NpcIndex).Invent.NroItems))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
   For LoopC = 1 To MAX_INVENTORY_SLOTS
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex & "-" & Npclist(NpcIndex).Invent.Object(LoopC).Amount)
   Next
End If


End Sub



Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)

'Status
If frmMain.Visible Then frmMain.txStatus.caption = "Cargando backup Npc"

Dim npcfile As String

If NpcNumber > 499 Then
    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
Else
    npcfile = DatPath & "bkNPCs.dat"
End If

Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
Npclist(NpcIndex).Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
Npclist(NpcIndex).Movement = Val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).NPCtype = Val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.Body = Val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = Val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).Char.Heading = Val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))

Npclist(NpcIndex).Attackable = Val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = Val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = Val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).GiveEXP = Val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP")) * 155

Npclist(NpcIndex).InvReSpawn = Val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))
Npclist(NpcIndex).QuestNumber = Val(GetVar(npcfile, "NPC" & NpcNumber, "QuestNumber"))
Npclist(NpcIndex).TalkAfterQuest = GetVar(npcfile, "NPC" & NpcNumber, "TalkAfterQuest")
Npclist(NpcIndex).TalkDuringQuest = GetVar(npcfile, "NPC" & NpcNumber, "TalkDuringQuest")

Npclist(NpcIndex).Stats.MaxHP = Val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = Val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHIT = Val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = Val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.def = Val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = Val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))


Dim LoopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = Val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = Val(ReadField(1, ln, 45))
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = Val(ReadField(2, ln, 45))
       
    Next LoopC
Else
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = 0
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = 0
    Next LoopC
End If


Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.UseAINow = False
Npclist(NpcIndex).flags.Respawn = Val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
Npclist(NpcIndex).flags.BackUp = Val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.Domable = Val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
Npclist(NpcIndex).flags.RespawnOrigPos = Val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = Val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

End Sub


Sub LogBan(ByVal BannedIndex As Integer, ByVal userindex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).name, "BannedBy", UserList(userindex).name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).name, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, UserList(BannedIndex).name
Close #mifile

End Sub


Sub LogBanFromName(ByVal BannedName As String, ByVal userindex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(userindex).name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub


Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)


'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub

Public Sub CargaApuestas()

    Apuestas.Ganancias = Val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = Val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = Val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

End Sub
