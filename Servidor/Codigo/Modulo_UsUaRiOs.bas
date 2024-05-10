Attribute VB_Name = "UsUaRiOs"
'Argentum Online 0.9.0.2
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)

Dim DaExp As Integer
Dim DaPT As Integer

DaExp = CInt(UserList(VictimIndex).Stats.ELV)

UserList(AttackerIndex).Stats.Exp = UserList(AttackerIndex).Stats.Exp + DaExp
If UserList(AttackerIndex).Stats.Exp > MAXEXP Then _
    UserList(AttackerIndex).Stats.Exp = MAXEXP

'Lo mata
Call SendData(SendTarget.toindex, AttackerIndex, 0, "||Has matado a " & UserList(VictimIndex).name & " (" & DaExp & ")" & FONTTYPE_AMARILLON)
Call LogUSER(UserList(AttackerIndex).name, "Mato a " & UserList(VictimIndex).name, False)
      
Call SendData(SendTarget.toindex, VictimIndex, 0, "||" & UserList(AttackerIndex).name & " te ha matado!" & FONTTYPE_FIGHT)

If TriggerZonaPelea(VictimIndex, AttackerIndex) <> TRIGGER6_PERMITE Then
    If Not Criminal(VictimIndex) And UserList(AttackerIndex).pos.Map <> 31 And UserList(AttackerIndex).pos.Map <> 32 And UserList(AttackerIndex).pos.Map <> 33 And UserList(AttackerIndex).pos.Map <> 34 Then
         UserList(AttackerIndex).Reputacion.AsesinoRep = UserList(AttackerIndex).Reputacion.AsesinoRep + vlASESINO * 2
         If UserList(AttackerIndex).Reputacion.AsesinoRep > MAXREP Then _
            UserList(AttackerIndex).Reputacion.AsesinoRep = MAXREP
         UserList(AttackerIndex).Reputacion.BurguesRep = 0
         UserList(AttackerIndex).Reputacion.NobleRep = 0
         UserList(AttackerIndex).Reputacion.PlebeRep = 0
    Else
         UserList(AttackerIndex).Reputacion.NobleRep = UserList(AttackerIndex).Reputacion.NobleRep + vlNoble
         If UserList(AttackerIndex).Reputacion.NobleRep > MAXREP Then _
            UserList(AttackerIndex).Reputacion.NobleRep = MAXREP
    End If
End If


Call UserDie(VictimIndex)

If UserList(AttackerIndex).Stats.UsuariosMatados < 32000 Then _
    UserList(AttackerIndex).Stats.UsuariosMatados = UserList(AttackerIndex).Stats.UsuariosMatados + 1
    
    'desafio
If MapInfo(72).NumUsers = 1 Then 'mapa de desafio
If UserList(AttackerIndex).flags.EnDesafio = 1 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " vencio a " & UserList(VictimIndex).name & ". " & FONTTYPE_INFO)
UserList(AttackerIndex).flags.rondas = UserList(AttackerIndex).flags.rondas + 1
UserList(VictimIndex).flags.EnDesafio = 0
UserList(VictimIndex).flags.Desafio = 0
End If
End If
 
If MapInfo(72).NumUsers = 1 Then 'mapa de desafio
If UserList(AttackerIndex).flags.rondas = 5 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 10 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 15 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 20 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 25 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 30 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 35 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 40 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 45 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 50 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 55 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 60 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 65 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 70 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 75 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 80 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 85 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 90 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 95 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 100 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 105 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 110 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 115 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 120 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 125 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 130 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 135 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 140 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 145 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 150 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 155 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 160 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 165 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 170 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 175 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 180 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 185 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 190 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 195 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
If UserList(AttackerIndex).flags.rondas = 200 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(AttackerIndex).name & " lleva " & UserList(AttackerIndex).flags.rondas & " rondas ganadas consecutivamente. " & FONTTYPE_GUILD)
End If
End If
 
If UserList(VictimIndex).flags.EnDesafio = 1 Then
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(VictimIndex).name & " fue vencido por " & UserList(AttackerIndex).name & ". " & FONTTYPE_GRISN)
'Call WarpUserChar(VictimIndex, 1, 50, 50, True)
Call WarpUserChar(AttackerIndex, 1, 64, 45, True)
UserList(AttackerIndex).flags.Desafio = 0
UserList(VictimIndex).flags.EnDesafio = 0
UserList(VictimIndex).flags.rondas = 0
 
End If


'Log
Call LogAsesinato(UserList(AttackerIndex).name & " asesino a " & UserList(VictimIndex).name)

End Sub


Sub RevivirUsuario(ByVal userindex As Integer)

UserList(userindex).flags.Muerto = 0
UserList(userindex).Stats.MinHP = 35

'No puede estar empollando
UserList(userindex).flags.EstaEmpo = 0
UserList(userindex).EmpoCont = 0

If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
End If

Call DarCuerpoDesnudo(userindex)
Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).OrigChar.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
Call SendUserStatsBox(userindex)


End Sub


Sub ChangeUserChar(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer, _
                    ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, _
                    ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)

    UserList(userindex).Char.Body = Body
    UserList(userindex).Char.Head = Head
    UserList(userindex).Char.Heading = Heading
    UserList(userindex).Char.WeaponAnim = Arma
    UserList(userindex).Char.ShieldAnim = Escudo
    UserList(userindex).Char.CascoAnim = Casco
    
    If sndRoute = SendTarget.ToMap Then
        Call SendToUserArea(userindex, "CP" & UserList(userindex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(userindex).Char.FX & "," & UserList(userindex).Char.loops & "," & Casco)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "CP" & UserList(userindex).Char.CharIndex & "," & Body & "," & Head & "," & Heading & "," & Arma & "," & Escudo & "," & UserList(userindex).Char.FX & "," & UserList(userindex).Char.loops & "," & Casco)
    End If
End Sub

Sub EnviarSubirNivel(ByVal userindex As Integer, ByVal puntos As Integer)
    Call SendData(SendTarget.toindex, userindex, 0, "UIVC" & puntos)
End Sub

Sub EnviarSkills(ByVal userindex As Integer)
    Dim i As Integer
    Dim cad As String
    
    For i = 1 To NUMSKILLS
       cad = cad & UserList(userindex).Stats.UserSkills(i) & ","
    Next i
    
    SendData SendTarget.toindex, userindex, 0, "LLSIKS" & cad$
End Sub

Sub EnviarFama(ByVal userindex As Integer)
    Dim cad As String
    
    cad = cad & UserList(userindex).Reputacion.AsesinoRep & ","
    cad = cad & UserList(userindex).Reputacion.BandidoRep & ","
    cad = cad & UserList(userindex).Reputacion.BurguesRep & ","
    cad = cad & UserList(userindex).Reputacion.LadronesRep & ","
    cad = cad & UserList(userindex).Reputacion.NobleRep & ","
    cad = cad & UserList(userindex).Reputacion.PlebeRep & ","
    
    Dim L As Long
    
    L = (-UserList(userindex).Reputacion.AsesinoRep) + _
        (-UserList(userindex).Reputacion.BandidoRep) + _
        UserList(userindex).Reputacion.BurguesRep + _
        (-UserList(userindex).Reputacion.LadronesRep) + _
        UserList(userindex).Reputacion.NobleRep + _
        UserList(userindex).Reputacion.PlebeRep
    L = L / 6
    
    UserList(userindex).Reputacion.Promedio = L
    
    cad = cad & UserList(userindex).Reputacion.Promedio
    
    SendData SendTarget.toindex, userindex, 0, "YGIJ" & cad
End Sub

Sub EnviarAtrib(ByVal userindex As Integer)
Dim i As Integer
Dim cad As String
For i = 1 To NUMATRIBUTOS
  cad = cad & UserList(userindex).Stats.UserAtributos(i) & ","
Next
Call SendData(SendTarget.toindex, userindex, 0, "KAJ" & cad)
End Sub

Public Sub EnviarMiniEstadisticas(ByVal userindex As Integer)
With UserList(userindex)
Call SendData(SendTarget.toindex, userindex, 0, "KIDX" & .Faccion.CiudadanosMatados & "," & _
.Faccion.CriminalesMatados & "," & .Stats.UsuariosMatados & "," & _
.Stats.NPCsMuertos & "," & .Clase & "," & .Counters.Pena & "," & .Faccion.NeutralesMatados)
End With
End Sub

Sub EraseUserChar(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, userindex As Integer)

On Error GoTo ErrorHandler
   
    CharList(UserList(userindex).Char.CharIndex) = 0
    
    If UserList(userindex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar <= 1 Then Exit Do
        Loop
    End If
    
    'Le mandamos el mensaje para que borre el personaje a los clientes que estén en el mismo mapa
    If sndRoute = SendTarget.ToMap Then
        Call SendToUserArea(userindex, "BP" & UserList(userindex).Char.CharIndex)
        Call QuitarUser(userindex, UserList(userindex).pos.Map)
    Else
        Call SendData(sndRoute, sndIndex, sndMap, "BP" & UserList(userindex).Char.CharIndex)
    End If
    
    MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).userindex = 0
    UserList(userindex).Char.CharIndex = 0
    
    NumChars = NumChars - 1
    
    Exit Sub
    
ErrorHandler:
        Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.Description)

End Sub

Sub MakeUserChar(ByVal sndRoute As SendTarget, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
On Local Error GoTo hayerror
    Dim CharIndex As Integer

    If InMapBounds(Map, X, Y) Then
        'If needed make a new character in list
        If UserList(userindex).Char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            UserList(userindex).Char.CharIndex = CharIndex
            CharList(CharIndex) = userindex
        End If
        
        'Place character on map
        MapData(Map, X, Y).userindex = userindex
        
        'Send make character command to clients
        Dim klan As String
        If UserList(userindex).GuildIndex > 0 Then
            klan = Guilds(UserList(userindex).GuildIndex).GuildName
        End If
        
        Dim bCr As Byte
        Dim bCnO As Byte
        Dim SendPrivilegios As Byte
       
        bCr = Criminal(userindex)
        bCnO = TransformadoVIP(userindex)

        If klan <> "" Then
            If sndRoute = SendTarget.toindex Then
#If SeguridadAlkon Then
                If EncriptarProtocolosCriticos Then
                    If UserList(userindex).flags.Privilegios > PlayerType.User Then
                        If UserList(userindex).showName Then
                            Call SendCryptedData(SendTarget.toindex, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & " <" & klan & ">" & "," & UserList(userindex).StatusMith.EsStatus & "," & IIf(UserList(userindex).flags.EsRolesMaster, 5, UserList(userindex).flags.Privilegios))
                        Else
                            'Hide the name and clan
                            Call SendCryptedData(SendTarget.toindex, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & ",," & UserList(userindex).StatusMith.EsStatus & "," & UserList(userindex).flags.Privilegios)
                        End If
                    Else
                        Call SendCryptedData(SendTarget.toindex, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & " <" & klan & ">" & "," & UserList(userindex).StatusMith.EsStatus & "," & UserList(userindex).flags.Privilegios)
                    End If
                Else
#End If
                    If UserList(userindex).flags.Privilegios > PlayerType.User Then
                        If UserList(userindex).showName Then
                            Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & " <" & klan & ">" & "," & UserList(userindex).StatusMith.EsStatus & "," & UserList(userindex).flags.Privilegios)
                        Else
                            'Hide the name and clan
                            Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & ",," & UserList(userindex).StatusMith.EsStatus & "," & UserList(userindex).flags.Privilegios)
                        End If
                    Else
                        Call SendData(sndRoute, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & " <" & klan & ">" & "," & UserList(userindex).StatusMith.EsStatus & "," & UserList(userindex).flags.Privilegios)
                    End If
#If SeguridadAlkon Then
                End If
#End If
            ElseIf sndRoute = SendTarget.ToMap Then
                Call AgregarUser(userindex, UserList(userindex).pos.Map)
                Call CheckUpdateNeededUser(userindex, USER_NUEVO)
            End If
        Else 'if tiene clan
            If sndRoute = SendTarget.toindex Then
#If SeguridadAlkon Then
                If EncriptarProtocolosCriticos Then
                    If UserList(userindex).flags.Privilegios > PlayerType.User Then
                        If UserList(userindex).showName Then
                            Call SendCryptedData(SendTarget.toindex, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & "," & UserList(userindex).StatusMith.EsStatus & "," & UserList(userindex).flags.Privilegios)
                        Else
                            'Hide the name
                            Call SendCryptedData(SendTarget.toindex, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & ",," & UserList(userindex).StatusMith.EsStatus & "," & UserList(userindex).flags.Privilegios)
                        End If
                    Else
                        Call SendCryptedData(SendTarget.toindex, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & "," & UserList(userindex).StatusMith.EsStatus & "," & UserList(userindex).flags.Privilegios)
                    End If
                Else
#End If
                    If UserList(userindex).flags.Privilegios > PlayerType.User Then
                        If UserList(userindex).showName Then
                            Call SendData(SendTarget.toindex, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & "," & UserList(userindex).StatusMith.EsStatus & "," & UserList(userindex).flags.Privilegios)
                        Else
                            Call SendData(SendTarget.toindex, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & ",," & UserList(userindex).StatusMith.EsStatus & "," & UserList(userindex).flags.Privilegios)
                        End If
                    Else
                        Call SendData(SendTarget.toindex, sndIndex, sndMap, "CC" & UserList(userindex).Char.Body & "," & UserList(userindex).Char.Head & "," & UserList(userindex).Char.Heading & "," & UserList(userindex).Char.CharIndex & "," & X & "," & Y & "," & UserList(userindex).Char.WeaponAnim & "," & UserList(userindex).Char.ShieldAnim & "," & UserList(userindex).Char.FX & "," & 999 & "," & UserList(userindex).Char.CascoAnim & "," & UserList(userindex).name & "," & UserList(userindex).StatusMith.EsStatus & "," & "," & UserList(userindex).flags.Privilegios)
                    End If
#If SeguridadAlkon Then
                End If
#End If
            ElseIf sndRoute = SendTarget.ToMap Then
                Call AgregarUser(userindex, UserList(userindex).pos.Map)
                Call CheckUpdateNeededUser(userindex, USER_NUEVO)
            End If
       End If   'if clan
    End If
Exit Sub

hayerror:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.Description)
    'Resume Next
    Call CloseSocket(userindex)
End Sub

Sub CheckUserLevel(ByVal userindex As Integer)

On Error GoTo errhandler

Dim Pts As Integer
Dim AumentoHIT As Integer
Dim AumentoMANA As Integer
Dim AumentoSTA As Integer
Dim WasNewbie As Boolean

'¿Alcanzo el maximo nivel?
If UserList(userindex).Stats.ELV >= STAT_MAXELV Then
    UserList(userindex).Stats.Exp = 0
    UserList(userindex).Stats.ELU = 0
    Exit Sub
End If

WasNewbie = EsNewbie(userindex)

'Si exp >= then Exp para subir de nivel entonce subimos el nivel
'If UserList(UserIndex).Stats.Exp >= UserList(UserIndex).Stats.ELU Then
Do While UserList(userindex).Stats.Exp >= UserList(userindex).Stats.ELU
    
    'Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_NIVEL)
    Call SendData(SendTarget.toindex, userindex, 0, "||¡Has subido de nivel!" & FONTTYPE_INFO)
    
    If UserList(userindex).Stats.ELV = 1 Then
        Pts = 10
    Else
        Pts = 10
    End If
    
    'UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts + Pts
    
    Call SendData(SendTarget.toindex, userindex, 0, "||Has ganado " & Pts & " skillpoints." & FONTTYPE_INFO)
       
    UserList(userindex).Stats.ELV = UserList(userindex).Stats.ELV + 1
    
    UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp - UserList(userindex).Stats.ELU
    
    If Not EsNewbie(userindex) And WasNewbie Then
        Call QuitarNewbieObj(userindex)
        If UCase$(MapInfo(UserList(userindex).pos.Map).Restringir) = "SI" Then
            Call WarpUserChar(userindex, 1, 59, 47, True)
            Call SendData(SendTarget.toindex, userindex, 0, "||Debes abandonar el Dungeon Newbie." & FONTTYPE_WARNING)
        End If
    End If

    If UserList(userindex).Stats.ELV < 50 Then
        UserList(userindex).Stats.ELU = UserList(userindex).Stats.ELU * 1.3
    End If

    Dim AumentoHP As Integer
    Select Case UCase$(UserList(userindex).Clase)
        Case "GUERRERO"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = RandomNumber(10, 12)
                Case 21
                    AumentoHP = RandomNumber(10, 11)
                Case 20
                    AumentoHP = RandomNumber(9, 11)
                Case 19, 18
                    AumentoHP = RandomNumber(8, 9)
                Case Else
                    AumentoHP = RandomNumber(6, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPGuerrero
            End Select
            
            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case "CAZADOR"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = RandomNumber(10, 11)
                Case 21
                    AumentoHP = RandomNumber(9, 11)
                Case 20
                    AumentoHP = RandomNumber(9, 10)
                Case 19
                    AumentoHP = RandomNumber(8, 9)
                Case 18
                    AumentoHP = RandomNumber(7, 9)
                Case Else
                    AumentoHP = RandomNumber(6, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select

            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case "PIRATA"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(9, 11)
                Case 20
                    AumentoHP = RandomNumber(7, 11)
                Case 18, 19
                    AumentoHP = RandomNumber(6, 11)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPGuerrero
            End Select
            
            AumentoHIT = 3
            AumentoSTA = AumentoSTDef
        
        Case "PALADIN"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = RandomNumber(10, 11)
                Case 21
                    AumentoHP = RandomNumber(9, 11)
                Case 20
                    AumentoHP = RandomNumber(9, 10)
                Case 19
                    AumentoHP = RandomNumber(8, 10)
                Case 18
                    AumentoHP = RandomNumber(8, 9)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) + AdicionalHPCazador
            End Select
            
            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "LADRON"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19, 18
                    AumentoHP = RandomNumber(5, 9)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select
            
            AumentoHIT = 1
            AumentoSTA = AumentoSTLadron
            
        Case "MAGO"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = RandomNumber(8, 9)
                Case 21
                    AumentoHP = RandomNumber(7, 9)
                Case 20
                    AumentoHP = RandomNumber(7, 8)
                Case 19
                    AumentoHP = RandomNumber(6, 8)
                Case 18
                    AumentoHP = RandomNumber(6, 7)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select
            If AumentoHP < 1 Then AumentoHP = 4
            
            AumentoHIT = 1
            AumentoMANA = 3.2 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTMago
        
        Case "LEÑADOR"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 9)
                Case 20
                    AumentoHP = RandomNumber(5, 9)
                Case 19, 18
                    AumentoHP = RandomNumber(4, 8)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTLeñador
        
        Case "MINERO"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 9)
                Case 20
                    AumentoHP = RandomNumber(5, 9)
                Case 19, 18
                    AumentoHP = RandomNumber(4, 8)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTMinero
        
        Case "PESCADOR"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 9)
                Case 20
                    AumentoHP = RandomNumber(5, 9)
                Case 19, 18
                    AumentoHP = RandomNumber(4, 8)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select
            
            AumentoHIT = 1
            AumentoSTA = AumentoSTPescador
        
        Case "CLERIGO"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = RandomNumber(9, 10)
                Case 21
                    AumentoHP = RandomNumber(8, 10)
                Case 20
                    AumentoHP = RandomNumber(8, 9)
                Case 19
                    AumentoHP = RandomNumber(7, 9)
                Case 18
                    AumentoHP = RandomNumber(7, 8)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select
            
            AumentoHIT = 2
            AumentoMANA = 1.8 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "DRUIDA"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = RandomNumber(8, 9)
                Case 21
                    AumentoHP = RandomNumber(8, 9)
                Case 20
                    AumentoHP = RandomNumber(7, 9)
                Case 19
                    AumentoHP = RandomNumber(7, 8)
                Case 18
                    AumentoHP = RandomNumber(6, 8)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "ASESINO"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = RandomNumber(9, 10)
                Case 21
                    AumentoHP = RandomNumber(8, 10)
                Case 20
                    AumentoHP = RandomNumber(8, 9)
                Case 19
                    AumentoHP = RandomNumber(7, 9)
                Case 18
                    AumentoHP = RandomNumber(7, 8)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select
            
            AumentoHIT = IIf(UserList(userindex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case "BARDO"
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 22
                    AumentoHP = RandomNumber(9, 10)
                Case 21
                    AumentoHP = RandomNumber(8, 10)
                Case 20
                    AumentoHP = RandomNumber(8, 9)
                Case 19
                    AumentoHP = RandomNumber(7, 9)
                Case 18
                    AumentoHP = RandomNumber(7, 8)
                Case Else
                    AumentoHP = RandomNumber(4, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2)
            End Select
            
            AumentoHIT = 2
            AumentoMANA = 2.1 * UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case Else
            Select Case UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion)
                Case 21
                    AumentoHP = RandomNumber(6, 9)
                Case 20
                    AumentoHP = RandomNumber(5, 9)
                Case 19, 18
                    AumentoHP = RandomNumber(4, 8)
                Case Else
                    AumentoHP = RandomNumber(5, UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) \ 2) - AdicionalHPCazador
            End Select

            AumentoHIT = 2
            AumentoSTA = AumentoSTDef
    End Select
    
    'Actualizamos HitPoints
    UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + AumentoHP
    If UserList(userindex).Stats.MaxHP > STAT_MAXHP Then _
        UserList(userindex).Stats.MaxHP = STAT_MAXHP
    'Actualizamos Stamina
    UserList(userindex).Stats.MaxSta = UserList(userindex).Stats.MaxSta + AumentoSTA
    If UserList(userindex).Stats.MaxSta > STAT_MAXSTA Then _
        UserList(userindex).Stats.MaxSta = STAT_MAXSTA
    'Actualizamos Mana
    UserList(userindex).Stats.MaxMan = UserList(userindex).Stats.MaxMan + AumentoMANA
    If UserList(userindex).Stats.ELV < 36 Then
        If UserList(userindex).Stats.MaxMan > STAT_MAXMAN Then _
            UserList(userindex).Stats.MaxMan = STAT_MAXMAN
    Else
        If UserList(userindex).Stats.MaxMan > 9999 Then _
            UserList(userindex).Stats.MaxMan = 9999
    End If
    
    'Actualizamos Golpe Máximo
    UserList(userindex).Stats.MaxHIT = UserList(userindex).Stats.MaxHIT + AumentoHIT
    If UserList(userindex).Stats.ELV < 36 Then
        If UserList(userindex).Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(userindex).Stats.MaxHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(userindex).Stats.MaxHIT > STAT_MAXHIT_OVER36 Then _
            UserList(userindex).Stats.MaxHIT = STAT_MAXHIT_OVER36
    End If
    
    'Actualizamos Golpe Mínimo
    UserList(userindex).Stats.MinHIT = UserList(userindex).Stats.MinHIT + AumentoHIT
    If UserList(userindex).Stats.ELV < 36 Then
        If UserList(userindex).Stats.MinHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(userindex).Stats.MinHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(userindex).Stats.MinHIT > STAT_MAXHIT_OVER36 Then _
            UserList(userindex).Stats.MinHIT = STAT_MAXHIT_OVER36
    End If
    
    'Notificamos al user
    If AumentoHP > 0 Then SendData SendTarget.toindex, userindex, 0, "||Has ganado " & AumentoHP & " puntos de vida." & FONTTYPE_INFO
    If AumentoSTA > 0 Then SendData SendTarget.toindex, userindex, 0, "||Has ganado " & AumentoSTA & " puntos de vitalidad." & FONTTYPE_INFO
    If AumentoMANA > 0 Then SendData SendTarget.toindex, userindex, 0, "||Has ganado " & AumentoMANA & " puntos de magia." & FONTTYPE_INFO
    If AumentoHIT > 0 Then
        SendData SendTarget.toindex, userindex, 0, "||Tu golpe maximo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
        SendData SendTarget.toindex, userindex, 0, "||Tu golpe minimo aumento en " & AumentoHIT & " puntos." & FONTTYPE_INFO
    End If
    
    Call LogDesarrollo(Date & " " & UserList(userindex).name & " paso a nivel " & UserList(userindex).Stats.ELV & " gano HP: " & AumentoHP)
    
    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
    
    Call EnviarSkills(userindex)
    Call EnviarSubirNivel(userindex, Pts)
   
    SendUserStatsBox userindex
    
Loop
'End If


Exit Sub

errhandler:
    LogError ("Error en la subrutina CheckUserLevel")
End Sub

Function PuedeAtravesarAgua(ByVal userindex As Integer) As Boolean

PuedeAtravesarAgua = _
  UserList(userindex).flags.Navegando = 1 Or _
  UserList(userindex).flags.Vuela = 1

End Function

Sub MoveUserChar(ByVal userindex As Integer, ByVal nHeading As Byte)


Dim nPos As WorldPos
Dim invpos As WorldPos
    nPos = UserList(userindex).pos
    Call HeadtoPos(nHeading, nPos)
    
    If LegalPos(UserList(userindex).pos.Map, nPos.X, nPos.Y, PuedeAtravesarAgua(userindex)) Then
        If MapInfo(UserList(userindex).pos.Map).NumUsers > 1 Then
#If SeguridadAlkon Then
            Call SendCryptedMoveChar(nPos.Map, userindex, nPos.X, nPos.Y)
#Else
            Call SendToUserAreaButindex(userindex, "+" & UserList(userindex).Char.CharIndex & "," & nPos.X & "," & nPos.Y)
#End If
        End If

          'Update map and user pos
        MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).userindex = 0
        UserList(userindex).pos = nPos
        UserList(userindex).Char.Heading = nHeading
        MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).userindex = userindex
        
        If HayTD Then Call TouchDown(userindex)
        
        If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).OBJInfo.ObjIndex = 1073 And UserList(userindex).flags.Muerto = 0 Then
        Call GetObj(userindex)
        End If
        
        If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).OBJInfo.ObjIndex = 1074 And UserList(userindex).flags.Muerto = 0 Then
        If UserList(userindex).Stats.MinHP <= 35 Then
        Call UserDie(userindex)
        Else
        UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - 35
        Call SendUserStatsBox(userindex)
        End If
        Call EraseObj(SendTarget.ToMap, 0, UserList(userindex).pos.Map, 1, UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y)
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & 55 & ",")
        End If
        
        If 0 = Val(GetVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho")) Then
        If MapData(mapainvo, mapainvoX1, mapainvoY1).userindex > 0 And MapData(mapainvo, mapainvoX2, mapainvoY2).userindex > 0 And MapData(mapainvo, mapainvoX3, mapainvoY3).userindex > 0 And MapData(mapainvo, mapainvoX4, mapainvoY4).userindex > 0 Then
        invpos.Map = mapainvo
        invpos.X = 29
        invpos.Y = 29
        Call SpawnNpc(RandomNumber(911, 935), invpos, True, False)
        Call WriteVar(DatPath & "InvocoBicho.dat", "INIT", "InvocoBicho", Val(1))
        End If
        End If
        
          If ZonaCura(userindex) Then Call AutoCuraUser(userindex)
          
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(userindex, nHeading)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "PU" & UserList(userindex).pos.X & "," & UserList(userindex).pos.Y)
    End If

    If UserList(userindex).Counters.Ocultando Then _
        UserList(userindex).Counters.Ocultando = UserList(userindex).Counters.Ocultando - 1
End Sub

Sub AutoCuraUser(ByVal userindex As Integer)
' Autor: Joan Calderón - SaturoS.
'Codigo: Sacerdotes automaticos.
If UserList(userindex).flags.Muerto = 1 Then
Call RevivirUsuario(userindex)
UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MaxSta
Call SendData(toindex, userindex, 0, "||El sacerdote te ha resucitado y curado." & FONTTYPE_INFO)
Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "TW20") ' este es el sonido cuando cura o resucita al personaje
Call SendUserStatsBox(userindex)
End If
 
If UserList(userindex).Stats.MinHP < UserList(userindex).Stats.MaxHP Then
UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
Call SendData(toindex, userindex, 0, "||El sacerdote te ha curado." & FONTTYPE_INFO)
Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "TW20") ' este es el sonido de cuando resucita o cura al personaje.
Call SendUserStatsBox(userindex)
End If
 
If UserList(userindex).flags.Envenenado = 1 Then UserList(userindex).flags.Envenenado = 0
 
 
End Sub

Sub ChangeUserInv(userindex As Integer, Slot As Byte, Object As UserOBJ)

    UserList(userindex).Invent.Object(Slot) = Object
    
    If Object.ObjIndex > 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "FBI" & Slot & "," & Object.ObjIndex & "," & ObjData(Object.ObjIndex).name & "," & Object.Amount & "," & Object.Equipped & "," & ObjData(Object.ObjIndex).GrhIndex & "," _
        & ObjData(Object.ObjIndex).OBJType & "," _
        & ObjData(Object.ObjIndex).MaxHIT & "," _
        & ObjData(Object.ObjIndex).MinHIT & "," _
        & ObjData(Object.ObjIndex).MaxDef)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "FBI" & Slot & "," & "0" & "," & "(Nada)" & "," & "0" & "," & "0")
    End If

End Sub


Function NextOpenCharIndex() As Integer
'Modificada por el oso para codificar los MP1234,2,1 en 2 bytes
'para lograrlo, el charindex no puede tener su bit numero 6 (desde 0) en 1
'y tampoco puede ser un charindex que tenga el bit 0 en 1.

On Local Error GoTo hayerror

Dim LoopC As Integer
    
    LoopC = 1
    
    While LoopC < MAXCHARS
        If CharList(LoopC) = 0 And Not ((LoopC And &HFFC0&) = 64) Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            If LoopC > LastChar Then LastChar = LoopC
            Exit Function
        Else
            LoopC = LoopC + 1
        End If
    Wend

Exit Function
hayerror:
LogError ("NextOpenCharIndex: num: " & Err.Number & " desc: " & Err.Description)

End Function
Function NextOpenUser() As Integer
    Dim LoopC As Long
   
    For LoopC = 1 To MaxUsers + 1
        If LoopC > MaxUsers Then Exit For
        If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
    Next LoopC
   
    NextOpenUser = LoopC
End Function
 
Sub SendUserHitBox(ByVal userindex As Integer)
Dim cosa As String
 
If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
        cosa = cosa & ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).MaxHIT & ","
    Else
        cosa = cosa & "0/0,"
    End If
 
    If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
        cosa = cosa & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MaxDef & ","
    Else
        cosa = cosa & "0/0,"
    End If
    
    If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
        cosa = cosa & ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MaxDef & ","
    Else
        cosa = cosa & "0/0,"
    End If
    
    If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
        cosa = cosa & ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).MaxDef & ","
    Else
        cosa = cosa & "0/0,"
    End If
    
    If UserList(userindex).Invent.HerramientaEqpObjIndex > 0 And UserList(userindex).Invent.CascoEqpObjIndex > 0 And UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
       cosa = cosa & (ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin + ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMin + ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMin) & "/" & (ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax + ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMax + ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMax)
        ElseIf UserList(userindex).Invent.HerramientaEqpObjIndex > 0 And UserList(userindex).Invent.CascoEqpObjIndex = 0 And UserList(userindex).Invent.ArmourEqpObjIndex = 0 Then
        cosa = cosa & ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin & "/" & ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax
        ElseIf UserList(userindex).Invent.HerramientaEqpObjIndex = 0 And UserList(userindex).Invent.CascoEqpObjIndex > 0 And UserList(userindex).Invent.ArmourEqpObjIndex = 0 Then
        cosa = cosa & ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMin & "/" & ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMax
        ElseIf UserList(userindex).Invent.HerramientaEqpObjIndex > 0 And UserList(userindex).Invent.CascoEqpObjIndex > 0 And UserList(userindex).Invent.ArmourEqpObjIndex = 0 Then
        cosa = cosa & (ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin + ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMin) & "/" & (ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax + ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
        ElseIf UserList(userindex).Invent.HerramientaEqpObjIndex = 0 And UserList(userindex).Invent.CascoEqpSlot = 0 And UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
        cosa = cosa & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMin & "/" & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMax
        ElseIf UserList(userindex).Invent.HerramientaEqpObjIndex > 0 And UserList(userindex).Invent.CascoEqpObjIndex = 0 And UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
        cosa = cosa & (ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin + ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMin) & "/" & (ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax + ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMax)
        ElseIf UserList(userindex).Invent.HerramientaEqpObjIndex = 0 And UserList(userindex).Invent.CascoEqpObjIndex > 0 And UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
        cosa = cosa & (ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMin + ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMin) & "/" & (ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMax + ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMax)
        Else
        cosa = cosa & "0/0"
    End If
    
SendData SendTarget.toindex, userindex, 0, "ARM" & cosa
 
End Sub

Sub SendUserStatsBox(ByVal userindex As Integer)
    Call SendData(SendTarget.toindex, userindex, 0, "WBP" & UserList(userindex).Stats.MaxHP & "," & UserList(userindex).Stats.MinHP & "," & UserList(userindex).Stats.MaxMan & "," & UserList(userindex).Stats.MinMAN & "," & UserList(userindex).Stats.MaxSta & "," & UserList(userindex).Stats.MinSta & "," & UserList(userindex).Stats.ELV & "," & UserList(userindex).Stats.ELU & "," & UserList(userindex).Stats.Exp & "," & UserList(userindex).BancoInvent.NroItems)
End Sub

Sub EnviarPuntos(ByVal userindex As Integer)
 Call SendData(SendTarget.toindex, userindex, 0, "PNT" & UserList(userindex).Stats.PuntosTorneo)
End Sub
Sub EnviarPuntosDonacion(ByVal userindex As Integer)
 Call SendData(SendTarget.toindex, userindex, 0, "DNC" & UserList(userindex).Stats.PuntosDonacion)
End Sub
Sub SendUserStatux(ByVal userindex As Integer)
On Error Resume Next
Dim Info As String
 
Call SendData(SendTarget.toindex, userindex, 0, "EZT" & "," & UserList(userindex).StatusMith.EsStatus)
 
Info = "NX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).StatusMith.EsStatus & "," & UserList(userindex).name
Call SendData(ToMap, userindex, UserList(userindex).pos.Map, (Info))
 
End Sub

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
Dim GuildI As Integer


    Call SendData(SendTarget.toindex, sendIndex, 0, "||Estadisticas de: " & UserList(userindex).name & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Nivel: " & UserList(userindex).Stats.ELV & "  EXP: " & UserList(userindex).Stats.Exp & "/" & UserList(userindex).Stats.ELU & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Vitalidad: " & UserList(userindex).Stats.FIT & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Salud: " & UserList(userindex).Stats.MinHP & "/" & UserList(userindex).Stats.MaxHP & "  Mana: " & UserList(userindex).Stats.MinMAN & "/" & UserList(userindex).Stats.MaxMan & "  Vitalidad: " & UserList(userindex).Stats.MinSta & "/" & UserList(userindex).Stats.MaxSta & FONTTYPE_INFO)
    
    If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
        Call SendData(SendTarget.toindex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(userindex).Stats.MinHIT & "/" & UserList(userindex).Stats.MaxHIT & " (" & ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).MaxHIT & ")" & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.toindex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & UserList(userindex).Stats.MinHIT & "/" & UserList(userindex).Stats.MaxHIT & FONTTYPE_INFO)
    End If
    
    If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
        Call SendData(SendTarget.toindex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: " & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).MaxDef & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.toindex, sendIndex, 0, "||(CUERPO) Min Def/Max Def: 0" & FONTTYPE_INFO)
    End If
    
    If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
        Call SendData(SendTarget.toindex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: " & ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(userindex).Invent.CascoEqpObjIndex).MaxDef & FONTTYPE_INFO)
    Else
        Call SendData(SendTarget.toindex, sendIndex, 0, "||(CABEZA) Min Def/Max Def: 0" & FONTTYPE_INFO)
    End If
    
    GuildI = UserList(userindex).GuildIndex
    If GuildI > 0 Then
        Call SendData(SendTarget.toindex, sendIndex, 0, "||Clan: " & Guilds(GuildI).GuildName & FONTTYPE_INFO)
        If UCase$(Guilds(GuildI).GetLeader) = UCase$(UserList(sendIndex).name) Then
            Call SendData(SendTarget.toindex, sendIndex, 0, "||Status: Lider" & FONTTYPE_INFO)
        End If
        'guildpts no tienen objeto
        'Call SendData(SendTarget.ToIndex, sendIndex, 0, "||User GuildPoints: " & UserList(UserIndex).GuildInfo.GuildPoints & FONTTYPE_INFO)
    End If
    
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Posicion: " & UserList(userindex).pos.X & "," & UserList(userindex).pos.Y & " en mapa " & UserList(userindex).pos.Map & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Puntos de Torneo: " & UserList(userindex).Stats.PuntosTorneo & (FONTTYPE_INFO))
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Retos Ganados: " & UserList(userindex).Stats.RetosGanados & (FONTTYPE_INFO))
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Retos Perdidos: " & UserList(userindex).Stats.RetosPerdidos & (FONTTYPE_INFO))
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Duelos Ganados: " & UserList(userindex).Stats.DuelosGanados & (FONTTYPE_INFO))
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Duelos Perdidos: " & UserList(userindex).Stats.DuelosPerdidos & (FONTTYPE_INFO))
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Dados: " & UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) & ", " & UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) & ", " & UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) & ", " & UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) & ", " & UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) & FONTTYPE_INFO)

End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
 
With UserList(userindex)
Call SendData(SendTarget.toindex, sendIndex, 0, "||Pj: " & .name & FONTTYPE_INFO)
Call SendData(SendTarget.toindex, sendIndex, 0, "||CiudadanosMatados: " & .Faccion.CiudadanosMatados & "CriminalesMatados: " & .Faccion.CriminalesMatados & "NeutralesMatados: " & .Faccion.NeutralesMatados & "UsuariosMatados: " & .Stats.UsuariosMatados & FONTTYPE_INFO)
Call SendData(SendTarget.toindex, sendIndex, 0, "||NPCsMuertos: " & .Stats.NPCsMuertos & FONTTYPE_INFO)
Call SendData(SendTarget.toindex, sendIndex, 0, "||Clase: " & .Clase & FONTTYPE_INFO)
Call SendData(SendTarget.toindex, sendIndex, 0, "||Pena: " & .Counters.Pena & FONTTYPE_INFO)
End With
 
End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
Dim CharFile As String
Dim Ban As String
Dim BanDetailPath As String

    BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & CharName & ".chr"
    
    If FileExist(CharFile) Then
        Call SendData(SendTarget.toindex, sendIndex, 0, "||Pj: " & CharName & FONTTYPE_INFO)
        ' 3 en uno :p
        Call SendData(SendTarget.toindex, sendIndex, 0, "||CiudadanosMatados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " UsuariosMatados: " & GetVar(CharFile, "MUERTES", "UserMuertes") & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, sendIndex, 0, "||NPCsMuertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes") & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, sendIndex, 0, "||Clase: " & GetVar(CharFile, "INIT", "Clase") & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, sendIndex, 0, "||Pena: " & GetVar(CharFile, "COUNTERS", "PENA") & FONTTYPE_INFO)
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call SendData(SendTarget.toindex, sendIndex, 0, "||Ban: " & Ban & FONTTYPE_INFO)
        If Ban = "1" Then
            Call SendData(SendTarget.toindex, sendIndex, 0, "||Ban por: " & GetVar(CharFile, CharName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, CharName, "Reason") & FONTTYPE_INFO)
        End If
    Else
        Call SendData(SendTarget.toindex, sendIndex, 0, "||El pj no existe: " & CharName & FONTTYPE_INFO)
    End If
    
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
On Error Resume Next

    Dim j As Long
    
    Call SendData(SendTarget.toindex, sendIndex, 0, "||" & UserList(userindex).name & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, sendIndex, 0, "|| Tiene " & UserList(userindex).Invent.NroItems & " objetos." & FONTTYPE_INFO)
    
    For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(userindex).Invent.Object(j).ObjIndex > 0 Then
            Call SendData(SendTarget.toindex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(UserList(userindex).Invent.Object(j).ObjIndex).name & " Cantidad:" & UserList(userindex).Invent.Object(j).Amount & FONTTYPE_INFO)
        End If
    Next j
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal CharName As String)
On Error Resume Next

    Dim j As Long
    Dim CharFile As String, Tmp As String
    Dim ObjInd As Long, ObjCant As Long
    
    CharFile = CharPath & CharName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call SendData(SendTarget.toindex, sendIndex, 0, "||" & CharName & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, sendIndex, 0, "|| Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos." & FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))
            If ObjInd > 0 Then
                Call SendData(SendTarget.toindex, sendIndex, 0, "|| Objeto " & j & " " & ObjData(ObjInd).name & " Cantidad:" & ObjCant & FONTTYPE_INFO)
            End If
        Next j
    Else
        Call SendData(SendTarget.toindex, sendIndex, 0, "||Usuario inexistente: " & CharName & FONTTYPE_INFO)
    End If
    
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal userindex As Integer)
On Error Resume Next
Dim j As Integer
Call SendData(SendTarget.toindex, sendIndex, 0, "||" & UserList(userindex).name & FONTTYPE_INFO)
For j = 1 To NUMSKILLS
    Call SendData(SendTarget.toindex, sendIndex, 0, "|| " & SkillsNames(j) & " = " & UserList(userindex).Stats.UserSkills(j) & FONTTYPE_INFO)
Next
Call SendData(SendTarget.toindex, sendIndex, 0, "|| SkillLibres:" & UserList(userindex).Stats.SkillPts & FONTTYPE_INFO)
End Sub

Function DameUserindex(SocketId As Integer) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Do Until UserList(LoopC).ConnID = SocketId

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserindex = 0
        Exit Function
    End If
    
Loop
  
DameUserindex = LoopC

End Function

Function DameUserIndexConNombre(ByVal Nombre As String) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Nombre = UCase$(Nombre)

Do Until UCase$(UserList(LoopC).name) = Nombre

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserIndexConNombre = 0
        Exit Function
    End If
    
Loop
  
DameUserIndexConNombre = LoopC

End Function


Function EsMascotaCiudadano(ByVal NpcIndex As Integer, ByVal userindex As Integer) As Boolean

If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascotaCiudadano = Not Criminal(Npclist(NpcIndex).MaestroUser)
        If EsMascotaCiudadano Then Call SendData(SendTarget.toindex, Npclist(NpcIndex).MaestroUser, 0, "||¡¡" & UserList(userindex).name & " esta atacando tu mascota!!" & FONTTYPE_FIGHT)
End If

End Function

Sub NpcAtacado(ByVal NpcIndex As Integer, ByVal userindex As Integer)

If Npclist(NpcIndex).pos.Map = MapCastilloN And Npclist(NpcIndex).NPCtype = ReyCastillo And Npclist(NpcIndex).Stats.MinHP > 15000 And Npclist(NpcIndex).Stats.MinHP <> 25000 Then Call SendData(SendTarget.ToDiosesYclan, GuildIndex(CastilloNorte), 0, "||El Rey del castillo Norte esta siendo atacado por el clan " & Guilds(UserList(userindex).GuildIndex).GuildName & "." & FONTTYPE_GUILD)
If Npclist(NpcIndex).pos.Map = MapCastilloN And Npclist(NpcIndex).NPCtype = ReyCastillo And Npclist(NpcIndex).Stats.MinHP > 500 And Npclist(NpcIndex).Stats.MinHP < 5000 Then Call SendData(SendTarget.ToDiosesYclan, GuildIndex(CastilloNorte), 0, "||El Rey del castillo Norte esta  a punto de caer en las manos del clan " & Guilds(UserList(userindex).GuildIndex).GuildName & "." & FONTTYPE_GUILD)
If Npclist(NpcIndex).pos.Map = MapCastilloS And Npclist(NpcIndex).NPCtype = ReyCastillo And Npclist(NpcIndex).Stats.MinHP > 15000 And Npclist(NpcIndex).Stats.MinHP <> 25000 Then Call SendData(SendTarget.ToDiosesYclan, GuildIndex(CastilloSur), 0, "||El Rey del castillo Sur esta siendo atacado por el clan " & Guilds(UserList(userindex).GuildIndex).GuildName & "." & FONTTYPE_GUILD)
If Npclist(NpcIndex).pos.Map = MapCastilloS And Npclist(NpcIndex).NPCtype = ReyCastillo And Npclist(NpcIndex).Stats.MinHP > 500 And Npclist(NpcIndex).Stats.MinHP < 5000 Then Call SendData(SendTarget.ToDiosesYclan, GuildIndex(CastilloSur), 0, "||El Rey del castillo Sur esta  a punto de caer en las manos del clan " & Guilds(UserList(userindex).GuildIndex).GuildName & "." & FONTTYPE_GUILD)
If Npclist(NpcIndex).pos.Map = MapCastilloE And Npclist(NpcIndex).NPCtype = ReyCastillo And Npclist(NpcIndex).Stats.MinHP > 15000 And Npclist(NpcIndex).Stats.MinHP <> 25000 Then Call SendData(SendTarget.ToDiosesYclan, GuildIndex(CastilloEste), 0, "||El Rey del castillo Este esta siendo atacado por el clan " & Guilds(UserList(userindex).GuildIndex).GuildName & "." & FONTTYPE_GUILD)
If Npclist(NpcIndex).pos.Map = MapCastilloE And Npclist(NpcIndex).NPCtype = ReyCastillo And Npclist(NpcIndex).Stats.MinHP > 500 And Npclist(NpcIndex).Stats.MinHP < 5000 Then Call SendData(SendTarget.ToDiosesYclan, GuildIndex(CastilloEste), 0, "||El Rey del castillo Este esta  a punto de caer en las manos del clan " & Guilds(UserList(userindex).GuildIndex).GuildName & "." & FONTTYPE_GUILD)
If Npclist(NpcIndex).pos.Map = MapCastilloO And Npclist(NpcIndex).NPCtype = ReyCastillo And Npclist(NpcIndex).Stats.MinHP > 15000 And Npclist(NpcIndex).Stats.MinHP <> 25000 Then Call SendData(SendTarget.ToDiosesYclan, GuildIndex(CastilloOeste), 0, "||El Rey del castillo Oeste esta siendo atacado por el clan " & Guilds(UserList(userindex).GuildIndex).GuildName & "." & FONTTYPE_GUILD)
If Npclist(NpcIndex).pos.Map = MapCastilloO And Npclist(NpcIndex).NPCtype = ReyCastillo And Npclist(NpcIndex).Stats.MinHP > 500 And Npclist(NpcIndex).Stats.MinHP < 5000 Then Call SendData(SendTarget.ToDiosesYclan, GuildIndex(CastilloOeste), 0, "||El Rey del castillo Oeste esta  a punto de caer en las manos del clan " & Guilds(UserList(userindex).GuildIndex).GuildName & "." & FONTTYPE_GUILD)

'Guardamos el usuario que ataco el npc
Npclist(NpcIndex).flags.AttackedBy = UserList(userindex).name

If Npclist(NpcIndex).MaestroUser > 0 Then Call AllMascotasAtacanUser(userindex, Npclist(NpcIndex).MaestroUser)

'Si atacaste mascota, te las picas de ciuda =D - Mithrandir
If EsMascotaCiudadano(NpcIndex, userindex) Then
Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
Npclist(NpcIndex).Hostile = 1
Else
'Reputacion
If Npclist(NpcIndex).Stats.Alineacion = 0 Then
 
'Si era guardia real, cagaste
If Ciudadano(userindex) Then
If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
End If
End If
 
'else :O
ElseIf Npclist(NpcIndex).Stats.Alineacion = 1 Then
UserList(userindex).Reputacion.PlebeRep = UserList(userindex).Reputacion.PlebeRep + vlCAZADOR / 2
If UserList(userindex).Reputacion.PlebeRep > MAXREP Then _
UserList(userindex).Reputacion.PlebeRep = MAXREP
End If
 
'hacemos que el npc se defienda
Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
Npclist(NpcIndex).Hostile = 1
End If

End Sub

Function PuedeApuñalar(ByVal userindex As Integer) As Boolean

If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
 PuedeApuñalar = _
 ((UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR) _
 And (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Apuñala = 1)) _
 Or _
  ((UCase$(UserList(userindex).Clase) = "ASESINO") And _
  (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).Apuñala = 1))
Else
 PuedeApuñalar = False
End If
End Function
Sub SubirSkill(ByVal userindex As Integer, ByVal Skill As Integer)

    Dim Aumenta As Integer
    Dim Prob As Integer
    
    If UserList(userindex).Stats.ELV <= 3 Then
        Prob = 25
    ElseIf UserList(userindex).Stats.ELV > 3 _
        And UserList(userindex).Stats.ELV < 6 Then
        Prob = 35
    ElseIf UserList(userindex).Stats.ELV >= 6 _
        And UserList(userindex).Stats.ELV < 10 Then
        Prob = 40
    ElseIf UserList(userindex).Stats.ELV >= 10 _
        And UserList(userindex).Stats.ELV < 20 Then
        Prob = 45
    Else
        Prob = 50
    End If
    
    Aumenta = 7
    
    Dim lvl As Integer
    lvl = UserList(userindex).Stats.ELV
    
    If lvl >= UBound(LevelSkill) Then Exit Sub
    If UserList(userindex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
    
    If Aumenta = 7 And UserList(userindex).Stats.UserSkills(Skill) < LevelSkill(lvl).LevelValue Then
        UserList(userindex).Stats.UserSkills(Skill) = UserList(userindex).Stats.UserSkills(Skill) + 1
        Call SendData(SendTarget.toindex, userindex, 0, "||¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(userindex).Stats.UserSkills(Skill) & " pts." & FONTTYPE_INFO)
        
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + 50
        If UserList(userindex).Stats.Exp > MAXEXP Then _
            UserList(userindex).Stats.Exp = MAXEXP
        
        Call SendData(SendTarget.toindex, userindex, 0, "||¡Has ganado 50 puntos de experiencia!" & FONTTYPE_AMARILLON)
        Call CheckUserLevel(userindex)
    End If

End Sub

Sub UserDie(ByVal userindex As Integer)
On Error GoTo ErrorHandler

    'Sonido
    If UCase$(UserList(userindex).Genero) = "MUJER" Then
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, e_SoundIndex.MUERTE_MUJER)
    Else
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, e_SoundIndex.MUERTE_HOMBRE)
    End If
    
    'Quitar el dialogo del user muerto
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "ULZ" & UserList(userindex).Char.CharIndex)
    
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbRed & "°" & "¡Aaaahhhh!" & "°" & str(UserList(userindex).Char.CharIndex))
    
    UserList(userindex).Stats.MinHP = 0
    UserList(userindex).Stats.MinSta = 0
    UserList(userindex).flags.AtacadoPorNpc = 0
    UserList(userindex).flags.AtacadoPorUser = 0
    UserList(userindex).flags.Envenenado = 0
    UserList(userindex).flags.Muerto = 1
    UserList(userindex).flags.Transformado = 0
    UserList(userindex).flags.TimeRevivir = 60
    
    
    Call SendUserHitBox(userindex)
    Dim aN As Integer
    
    aN = UserList(userindex).flags.AtacadoPorNpc
    
    If aN > 0 Then
        Npclist(aN).Movement = Npclist(aN).flags.OldMovement
        Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
        Npclist(aN).flags.AttackedBy = ""
    End If
    
    '<<<< Paralisis >>>>
    If UserList(userindex).flags.Paralizado = 1 Then
        UserList(userindex).flags.Paralizado = 0
        Call SendData(SendTarget.toindex, userindex, 0, "PARADOK")
    End If
    
    '<<< Estupidez >>>
    If UserList(userindex).flags.Estupidez = 1 Then
        UserList(userindex).flags.Estupidez = 0
        Call SendData(SendTarget.toindex, userindex, 0, "NESTUP")
    End If
    
    '<<<< Descansando >>>>
    If UserList(userindex).flags.Descansar Then
        UserList(userindex).flags.Descansar = False
        Call SendData(SendTarget.toindex, userindex, 0, "VGH")
    End If
    
    '<<<< Meditando >>>>
    If UserList(userindex).flags.Meditando Then
        UserList(userindex).flags.Meditando = False
        Call SendData(SendTarget.toindex, userindex, 0, "PEDOP")
    End If
    
    '<<<<< Seg Resu >>>>>
    If UserList(userindex).flags.SeguroResu = False Then
    Call SendData(SendTarget.toindex, userindex, 0, "SEGONR")
        UserList(userindex).flags.SeguroResu = True
    End If
    
    '<<<< Invisible >>>>
    If UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1 Then
        UserList(userindex).flags.Oculto = 0
        UserList(userindex).flags.Invisible = 0
        'no hace falta encriptar este NOVER
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
    End If
    
       If MapInfo(UserList(userindex).pos.Map).SeCaenItems = 0 Then
         If TriggerZonaPelea(userindex, userindex) <> TRIGGER6_PERMITE Then
                ' << Si es newbie no pierde el inventario >>
           If Not EsNewbie(userindex) Or Criminal(userindex) Then
                 Call TirarTodo(userindex)
           Else
                 If EsNewbie(userindex) Then Call TirarTodosLosItemsNoNewbies(userindex)
           End If
         End If
         
    End If
    
    
    ' DESEQUIPA TODOS LOS OBJETOS
    'desequipar armadura
    If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.ArmourEqpSlot)
    End If
    'desequipar arma
    If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)
    End If
    'desequipar casco
    If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.CascoEqpSlot)
    End If
    'desequipar herramienta
    If UserList(userindex).Invent.HerramientaEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.HerramientaEqpSlot)
    End If
    'desequipar municiones
    If UserList(userindex).Invent.MunicionEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.MunicionEqpSlot)
    End If
    'desequipar montura
    If UserList(userindex).flags.Montando = 1 Then
        UserList(userindex).flags.Montando = 0
        Call SendData(SendTarget.toindex, userindex, 0, "EQUIT")
        Call Desequipar(userindex, UserList(userindex).Invent.MonturaSlot)
    End If
    'desequipar escudo
    If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
        Call Desequipar(userindex, UserList(userindex).Invent.EscudoEqpSlot)
    End If

'---------------------------------------------
'Form al Morir
'---------------------------------------------
    If Not UserList(userindex).pos.Map = 1 Then
    If Not UserList(userindex).pos.Map = 14 Then
    If Not UserList(userindex).pos.Map = 66 Then
    If Not UserList(userindex).pos.Map = 72 Then
    If Not UserList(userindex).pos.Map = 54 Then
    If Not UserList(userindex).pos.Map = 20 Then
    If Not UserList(userindex).pos.Map = 8 Then
    If Not UserList(userindex).pos.Map = 31 Then
    If Not UserList(userindex).pos.Map = 32 Then
    If Not UserList(userindex).pos.Map = 33 Then
    If Not UserList(userindex).pos.Map = 34 Then
    If Not UserList(userindex).pos.Map = 81 Then
    If Not UserList(userindex).pos.Map = 105 Then
    If Not UserList(userindex).pos.Map = 104 Then
    If Not UserList(userindex).pos.Map = 70 Then
    If Not UserList(userindex).pos.Map = 120 Then
    Call SendData(SendTarget.toindex, userindex, 0, "FEERASD")
    End If '1
     End If '14
      End If '66
       End If '72
        End If '54
         End If '20
          End If '8
           End If '31
            End If '32
             End If '33
              End If '34
               End If '81
                End If '105
                 End If '104
                  End If '70
                   End If '120
'---------------------------------------------
'/Form al Morir
'---------------------------------------------


        If UserList(userindex).flags.EnDuelo Then
        Dim uDuelo1     As Integer
        Dim uDuelo2     As Integer
        
        uDuelo2 = NameIndex(UserList(userindex).flags.DueliandoContra)
        uDuelo1 = userindex
        
        Call TerminaDuelin(uDuelo1, uDuelo2)
    End If
    
    ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> casted - pareja 2vs2
    If HayPareja = True Then
    If UserList(Pareja.Jugador1).flags.EnPareja = True And UserList(Pareja.Jugador2).flags.EnPareja = True And UserList(Pareja.Jugador1).flags.Muerto = 1 And UserList(Pareja.Jugador2).flags.Muerto = 1 Then
        Call WarpUserChar(Pareja.Jugador1, PosUserPareja1.Map, PosUserPareja1.X, PosUserPareja1.Y)
        Call WarpUserChar(Pareja.Jugador2, PosUserPareja2.Map, PosUserPareja2.X, PosUserPareja2.Y)
        Call WarpUserChar(Pareja.Jugador3, PosUserPareja3.Map, PosUserPareja3.X, PosUserPareja3.Y)
        Call WarpUserChar(Pareja.Jugador4, PosUserPareja4.Map, PosUserPareja4.X, PosUserPareja4.Y)
        UserList(Pareja.Jugador1).flags.EnPareja = False
        UserList(Pareja.Jugador1).flags.EsperaPareja = False
        UserList(Pareja.Jugador1).flags.SuPareja = 0
        UserList(Pareja.Jugador2).flags.EnPareja = False
        UserList(Pareja.Jugador2).flags.EsperaPareja = False
        UserList(Pareja.Jugador2).flags.SuPareja = 0
        UserList(Pareja.Jugador3).flags.EnPareja = False
        UserList(Pareja.Jugador3).flags.EsperaPareja = False
        UserList(Pareja.Jugador3).flags.SuPareja = 0
        UserList(Pareja.Jugador4).flags.EnPareja = False
        UserList(Pareja.Jugador4).flags.EsperaPareja = False
        UserList(Pareja.Jugador4).flags.SuPareja = 0
        HayPareja = False
        Call SendData(SendTarget.toall, 0, 0, "||" & UserList(Pareja.Jugador3).name & " y " & UserList(Pareja.Jugador4).name & " ganaron el desafio." & FONTTYPE_BLANCON)
    End If
   
    If UserList(Pareja.Jugador3).flags.EnPareja = True And UserList(Pareja.Jugador4).flags.EnPareja = True And UserList(Pareja.Jugador3).flags.Muerto = 1 And UserList(Pareja.Jugador4).flags.Muerto = 1 Then
        Call WarpUserChar(Pareja.Jugador1, PosUserPareja1.Map, PosUserPareja1.X, PosUserPareja1.Y)
        Call WarpUserChar(Pareja.Jugador2, PosUserPareja2.Map, PosUserPareja2.X, PosUserPareja2.Y)
        Call WarpUserChar(Pareja.Jugador3, PosUserPareja3.Map, PosUserPareja3.X, PosUserPareja3.Y)
        Call WarpUserChar(Pareja.Jugador4, PosUserPareja4.Map, PosUserPareja4.X, PosUserPareja4.Y)
        UserList(Pareja.Jugador1).flags.EnPareja = False
        UserList(Pareja.Jugador1).flags.EsperaPareja = False
        UserList(Pareja.Jugador1).flags.SuPareja = 0
        UserList(Pareja.Jugador2).flags.EnPareja = False
        UserList(Pareja.Jugador2).flags.EsperaPareja = False
        UserList(Pareja.Jugador2).flags.SuPareja = 0
        UserList(Pareja.Jugador3).flags.EnPareja = False
        UserList(Pareja.Jugador3).flags.EsperaPareja = False
        UserList(Pareja.Jugador3).flags.SuPareja = 0
        UserList(Pareja.Jugador4).flags.EnPareja = False
        UserList(Pareja.Jugador4).flags.EsperaPareja = False
        UserList(Pareja.Jugador4).flags.SuPareja = 0
        HayPareja = False
        Call SendData(SendTarget.toall, 0, 0, "||" & UserList(Pareja.Jugador1).name & " y " & UserList(Pareja.Jugador2).name & " ganaron el desafio." & FONTTYPE_BLANCON)
    End If
End If

'-------------------Finales--------------
'-------------------1vs1--------------
If UserList(userindex).flags.DueleandoFinal = True Then
If Arena1 = True Then
    If UserList(Torne.Jugador1).flags.Muerto = 1 Then
        Call SendData(SendTarget.toall, 0, 0, "||Torneo: Final: " & UserList(Torne.Jugador2).name & " ganó el combate." & "~230~230~0~1~0")
        Arena1 = False
        UserList(Torne.Jugador1).flags.DueleandoFinal = False
        UserList(Torne.Jugador2).flags.DueleandoFinal = False
        Torne.Jugador1 = 0
        Torne.Jugador2 = 0
    ElseIf UserList(Torne.Jugador2).flags.Muerto = 1 Then
        Call SendData(SendTarget.toall, 0, 0, "||Torneo: Final: " & UserList(Torne.Jugador1).name & " ganó el combate." & "~230~230~0~1~0")
        Arena1 = False
        UserList(Torne.Jugador1).flags.DueleandoFinal = False
        UserList(Torne.Jugador2).flags.DueleandoFinal = False
        Torne.Jugador1 = 0
        Torne.Jugador2 = 0
        End If
    End If
End If
'----------------------------2VS2----------------------------
If UserList(userindex).flags.DueleandoFinal2 = True Then
If Arena1 = True Then
If UserList(Torne.Jugador1).flags.Muerto = 1 And UserList(Torne.Jugador2).flags.Muerto = 1 Then
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Final: " & UserList(Torne.Jugador3).name & " y " & UserList(Torne.Jugador4).name & " ganaron el combate." & "~230~230~0~1~0")
Arena1 = False
UserList(Torne.Jugador1).flags.DueleandoFinal2 = False
UserList(Torne.Jugador2).flags.DueleandoFinal2 = False
UserList(Torne.Jugador3).flags.DueleandoFinal2 = False
UserList(Torne.Jugador4).flags.DueleandoFinal2 = False
Torne.Jugador1 = 0
Torne.Jugador2 = 0
Torne.Jugador3 = 0
Torne.Jugador4 = 0
ElseIf UserList(Torne.Jugador3).flags.Muerto = 1 And UserList(Torne.Jugador4).flags.Muerto = 1 Then
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Final: " & UserList(Torne.Jugador1).name & " y " & UserList(Torne.Jugador2).name & " ganaron el combate." & "~230~230~0~1~0")
Arena1 = False
UserList(Torne.Jugador1).flags.DueleandoFinal2 = False
UserList(Torne.Jugador2).flags.DueleandoFinal2 = False
UserList(Torne.Jugador3).flags.DueleandoFinal2 = False
UserList(Torne.Jugador4).flags.DueleandoFinal2 = False
Torne.Jugador1 = 0
Torne.Jugador2 = 0
Torne.Jugador3 = 0
Torne.Jugador4 = 0
End If
End If
End If
'----------------------------3VS3----------------------------
 'Arena1

If UserList(userindex).flags.DueleandoFinal3 = True Then
If Arena1 = True Then
If UserList(Torne.Jugador1).flags.Muerto = 1 And UserList(Torne.Jugador2).flags.Muerto = 1 And UserList(Torne.Jugador3).flags.Muerto = 1 Then
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Final: " & UserList(Torne.Jugador4).name & ", " & UserList(Torne.Jugador5).name & " y " & UserList(Torne.Jugador6).name & " ganaron el combate." & "~230~230~0~1~0")
Arena1 = False
UserList(Torne.Jugador1).flags.DueleandoFinal3 = False
UserList(Torne.Jugador2).flags.DueleandoFinal3 = False
UserList(Torne.Jugador3).flags.DueleandoFinal3 = False
UserList(Torne.Jugador4).flags.DueleandoFinal3 = False
UserList(Torne.Jugador5).flags.DueleandoFinal3 = False
UserList(Torne.Jugador6).flags.DueleandoFinal3 = False
Torne.Jugador1 = 0
Torne.Jugador2 = 0
Torne.Jugador3 = 0
Torne.Jugador4 = 0
Torne.Jugador5 = 0
Torne.Jugador6 = 0
ElseIf UserList(Torne.Jugador4).flags.Muerto = 1 And UserList(Torne.Jugador5).flags.Muerto = 1 And UserList(Torne.Jugador6).flags.Muerto = 1 Then
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Final: " & UserList(Torne.Jugador1).name & ", " & UserList(Torne.Jugador2).name & " y " & UserList(Torne.Jugador3).name & " ganaron el combate." & "~230~230~0~1~0")
Arena1 = False
UserList(Torne.Jugador1).flags.DueleandoFinal3 = False
UserList(Torne.Jugador2).flags.DueleandoFinal3 = False
UserList(Torne.Jugador3).flags.DueleandoFinal3 = False
UserList(Torne.Jugador4).flags.DueleandoFinal3 = False
UserList(Torne.Jugador5).flags.DueleandoFinal3 = False
UserList(Torne.Jugador6).flags.DueleandoFinal3 = False
Torne.Jugador1 = 0
Torne.Jugador2 = 0
Torne.Jugador3 = 0
Torne.Jugador4 = 0
Torne.Jugador5 = 0
Torne.Jugador6 = 0
End If
End If
End If
'----------------------------------4VS4---------------------------------
If UserList(userindex).flags.DueleandoFinal4 = True Then
If Arena1 = True Then
If UserList(Torne.Jugador1).flags.Muerto = 1 And UserList(Torne.Jugador2).flags.Muerto = 1 And UserList(Torne.Jugador3).flags.Muerto = 1 And UserList(Torne.Jugador4).flags.Muerto = 1 Then
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Final: " & UserList(Torne.Jugador5).name & ", " & UserList(Torne.Jugador6).name & ", " & UserList(Torne.Jugador7).name & " y " & UserList(Torne.Jugador8).name & " ganaron el combate." & "~230~230~0~1~0")
Arena1 = False
UserList(Torne.Jugador1).flags.DueleandoFinal4 = False
UserList(Torne.Jugador2).flags.DueleandoFinal4 = False
UserList(Torne.Jugador3).flags.DueleandoFinal4 = False
UserList(Torne.Jugador4).flags.DueleandoFinal4 = False
UserList(Torne.Jugador5).flags.DueleandoFinal4 = False
UserList(Torne.Jugador6).flags.DueleandoFinal4 = False
UserList(Torne.Jugador7).flags.DueleandoFinal4 = False
UserList(Torne.Jugador8).flags.DueleandoFinal4 = False
Torne.Jugador1 = 0
Torne.Jugador2 = 0
Torne.Jugador3 = 0
Torne.Jugador4 = 0
Torne.Jugador5 = 0
Torne.Jugador6 = 0
Torne.Jugador7 = 0
Torne.Jugador8 = 0
ElseIf UserList(Torne.Jugador5).flags.Muerto = 1 And UserList(Torne.Jugador6).flags.Muerto = 1 And UserList(Torne.Jugador7).flags.Muerto = 1 And UserList(Torne.Jugador8).flags.Muerto = 1 Then
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Final: " & UserList(Torne.Jugador1).name & ", " & UserList(Torne.Jugador2).name & ", " & UserList(Torne.Jugador3).name & " y " & UserList(Torne.Jugador4).name & " ganaron el combate." & "~230~230~0~1~0")
Arena1 = False
UserList(Torne.Jugador1).flags.DueleandoFinal4 = False
UserList(Torne.Jugador2).flags.DueleandoFinal4 = False
UserList(Torne.Jugador3).flags.DueleandoFinal4 = False
UserList(Torne.Jugador4).flags.DueleandoFinal4 = False
UserList(Torne.Jugador5).flags.DueleandoFinal4 = False
UserList(Torne.Jugador6).flags.DueleandoFinal4 = False
UserList(Torne.Jugador7).flags.DueleandoFinal4 = False
UserList(Torne.Jugador8).flags.DueleandoFinal4 = False
Torne.Jugador1 = 0
Torne.Jugador2 = 0
Torne.Jugador3 = 0
Torne.Jugador4 = 0
Torne.Jugador5 = 0
Torne.Jugador6 = 0
Torne.Jugador7 = 0
Torne.Jugador8 = 0
End If
End If
End If

'-------------------1vs1--------------
'Arena1
If UserList(userindex).flags.DueleandoTorneo = True Then
If Arena1 = True Then
    If UserList(Torne.Jugador1).flags.Muerto = 1 Then
        Call WarpUserChar(Torne.Jugador1, 1, 50, 50)
        Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 1: " & UserList(Torne.Jugador2).name & " ganó el combate." & "~230~230~0~1~0")
        Arena1 = False
        UserList(Torne.Jugador1).flags.DueleandoTorneo = False
        UserList(Torne.Jugador2).flags.DueleandoTorneo = False
        Torne.Jugador1 = 0
        Torne.Jugador2 = 0
    ElseIf UserList(Torne.Jugador2).flags.Muerto = 1 Then
        Call WarpUserChar(Torne.Jugador2, 1, 50, 50)
        Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 1: " & UserList(Torne.Jugador1).name & " ganó el combate." & "~230~230~0~1~0")
        Arena1 = False
        UserList(Torne.Jugador1).flags.DueleandoTorneo = False
        UserList(Torne.Jugador2).flags.DueleandoTorneo = False
        Torne.Jugador1 = 0
        Torne.Jugador2 = 0
        End If
    End If
        
If Arena2 = True Then
    If UserList(Torne.Jugador3).flags.Muerto = 1 Then
        Call WarpUserChar(Torne.Jugador3, 1, 50, 50)
        Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 2: " & UserList(Torne.Jugador4).name & " ganó el combate." & "~230~230~0~1~0")
        Arena2 = False
        UserList(Torne.Jugador3).flags.DueleandoTorneo = False
        UserList(Torne.Jugador4).flags.DueleandoTorneo = False
        Torne.Jugador3 = 0
        Torne.Jugador4 = 0
    ElseIf UserList(Torne.Jugador4).flags.Muerto = 1 Then
        Call WarpUserChar(Torne.Jugador4, 1, 50, 50)
        Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 2: " & UserList(Torne.Jugador3).name & " ganó el combate." & "~230~230~0~1~0")
        Arena2 = False
        UserList(Torne.Jugador3).flags.DueleandoTorneo = False
        UserList(Torne.Jugador4).flags.DueleandoTorneo = False
        Torne.Jugador3 = 0
        Torne.Jugador4 = 0
        End If
     End If
        
If Arena3 = True Then
        If UserList(Torne.Jugador5).flags.Muerto = 1 Then
        Call WarpUserChar(Torne.Jugador5, 1, 50, 50)
        Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 3: " & UserList(Torne.Jugador6).name & " ganó el combate." & "~230~230~0~1~0")
        Arena3 = False
        UserList(Torne.Jugador5).flags.DueleandoTorneo = False
        UserList(Torne.Jugador6).flags.DueleandoTorneo = False
        Torne.Jugador5 = 0
        Torne.Jugador6 = 0
    ElseIf UserList(Torne.Jugador6).flags.Muerto = 1 Then
        Call WarpUserChar(Torne.Jugador6, 1, 50, 50)
        Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 3: " & UserList(Torne.Jugador5).name & " ganó el combate." & "~230~230~0~1~0")
        Arena3 = False
        UserList(Torne.Jugador5).flags.DueleandoTorneo = False
        UserList(Torne.Jugador6).flags.DueleandoTorneo = False
        Torne.Jugador5 = 0
        Torne.Jugador6 = 0
        End If
    End If
        
        
If Arena4 = True Then
        If UserList(Torne.Jugador7).flags.Muerto = 1 Then
        Call WarpUserChar(Torne.Jugador7, 1, 50, 50)
        Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 4: " & UserList(Torne.Jugador8).name & " ganó el combate." & "~230~230~0~1~0")
        Arena4 = False
        UserList(Torne.Jugador7).flags.DueleandoTorneo = False
        UserList(Torne.Jugador8).flags.DueleandoTorneo = False
        Torne.Jugador7 = 0
        Torne.Jugador8 = 0
    ElseIf UserList(Torne.Jugador8).flags.Muerto = 1 Then
        Call WarpUserChar(Torne.Jugador8, 1, 50, 50)
        Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 4: " & UserList(Torne.Jugador7).name & " ganó el combate." & "~230~230~0~1~0")
        Arena4 = False
        UserList(Torne.Jugador7).flags.DueleandoTorneo = False
        UserList(Torne.Jugador8).flags.DueleandoTorneo = False
        Torne.Jugador7 = 0
        Torne.Jugador8 = 0
        End If
    End If
    End If
 '----------------------------2VS2----------------------------
'Arena1
If UserList(userindex).flags.DueleandoTorneo2 = True Then
If Arena1 = True Then
If UserList(Torne.Jugador1).flags.Muerto = 1 And UserList(Torne.Jugador2).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador1, 1, 50, 50)
Call WarpUserChar(Torne.Jugador2, 1, 50, 51)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 1: " & UserList(Torne.Jugador3).name & " y " & UserList(Torne.Jugador4).name & " ganaron el combate." & "~230~230~0~1~0")
Arena1 = False
UserList(Torne.Jugador1).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador2).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador3).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador4).flags.DueleandoTorneo2 = False
Torne.Jugador1 = 0
Torne.Jugador2 = 0
Torne.Jugador3 = 0
Torne.Jugador4 = 0
ElseIf UserList(Torne.Jugador3).flags.Muerto = 1 And UserList(Torne.Jugador4).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador3, 1, 50, 50)
Call WarpUserChar(Torne.Jugador4, 1, 50, 51)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 1: " & UserList(Torne.Jugador1).name & " y " & UserList(Torne.Jugador2).name & " ganaron el combate." & "~230~230~0~1~0")
Arena1 = False
UserList(Torne.Jugador1).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador2).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador3).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador4).flags.DueleandoTorneo2 = False
Torne.Jugador1 = 0
Torne.Jugador2 = 0
Torne.Jugador3 = 0
Torne.Jugador4 = 0
End If
End If
'Arena2
If Arena2 = True Then
If UserList(Torne.Jugador5).flags.Muerto = 1 And UserList(Torne.Jugador6).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador5, 1, 50, 50)
Call WarpUserChar(Torne.Jugador6, 1, 50, 51)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 2: " & UserList(Torne.Jugador7).name & " y " & UserList(Torne.Jugador8).name & " ganaron el combate." & "~230~230~0~1~0")
Arena2 = False
UserList(Torne.Jugador5).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador6).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador7).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador8).flags.DueleandoTorneo2 = False
Torne.Jugador5 = 0
Torne.Jugador6 = 0
Torne.Jugador7 = 0
Torne.Jugador8 = 0
ElseIf UserList(Torne.Jugador7).flags.Muerto = 1 And UserList(Torne.Jugador8).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador7, 1, 50, 50)
Call WarpUserChar(Torne.Jugador8, 1, 50, 51)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 2: " & UserList(Torne.Jugador5).name & " y " & UserList(Torne.Jugador6).name & " ganaron el combate." & "~230~230~0~1~0")
Arena2 = False
UserList(Torne.Jugador5).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador6).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador7).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador8).flags.DueleandoTorneo2 = False
Torne.Jugador5 = 0
Torne.Jugador6 = 0
Torne.Jugador7 = 0
Torne.Jugador8 = 0
End If
End If

'Arena3
If Arena3 = True Then
If UserList(Torne.Jugador9).flags.Muerto = 1 And UserList(Torne.Jugador10).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador9, 1, 50, 50)
Call WarpUserChar(Torne.Jugador10, 1, 50, 51)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 3: " & UserList(Torne.Jugador11).name & " y " & UserList(Torne.Jugador12).name & " ganaron el combate." & "~230~230~0~1~0")
Arena3 = False
UserList(Torne.Jugador9).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador10).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador11).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador12).flags.DueleandoTorneo2 = False
Torne.Jugador9 = 0
Torne.Jugador10 = 0
Torne.Jugador11 = 0
Torne.Jugador12 = 0
ElseIf UserList(Torne.Jugador11).flags.Muerto = 1 And UserList(Torne.Jugador12).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador11, 1, 50, 50)
Call WarpUserChar(Torne.Jugador12, 1, 50, 51)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 3: " & UserList(Torne.Jugador9).name & " y " & UserList(Torne.Jugador10).name & " ganaron el combate." & "~230~230~0~1~0")
Arena3 = False
UserList(Torne.Jugador9).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador10).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador11).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador12).flags.DueleandoTorneo2 = False
Torne.Jugador9 = 0
Torne.Jugador10 = 0
Torne.Jugador11 = 0
Torne.Jugador12 = 0
End If
End If

'Arena4
If Arena4 = True Then
If UserList(Torne.Jugador13).flags.Muerto = 1 And UserList(Torne.Jugador14).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador13, 1, 50, 50)
Call WarpUserChar(Torne.Jugador14, 1, 50, 51)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 4: " & UserList(Torne.Jugador15).name & " y " & UserList(Torne.Jugador16).name & " ganaron el combate." & "~230~230~0~1~0")
Arena4 = False
UserList(Torne.Jugador13).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador14).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador15).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador16).flags.DueleandoTorneo2 = False
Torne.Jugador13 = 0
Torne.Jugador14 = 0
Torne.Jugador15 = 0
Torne.Jugador16 = 0
ElseIf UserList(Torne.Jugador15).flags.Muerto = 1 And UserList(Torne.Jugador16).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador15, 1, 50, 50)
Call WarpUserChar(Torne.Jugador16, 1, 50, 51)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 4: " & UserList(Torne.Jugador13).name & " y " & UserList(Torne.Jugador14).name & " ganaron el combate." & "~230~230~0~1~0")
Arena4 = False
UserList(Torne.Jugador13).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador14).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador15).flags.DueleandoTorneo2 = False
UserList(Torne.Jugador16).flags.DueleandoTorneo2 = False
Torne.Jugador13 = 0
Torne.Jugador14 = 0
Torne.Jugador15 = 0
Torne.Jugador16 = 0
End If
End If
End If
 '----------------------------3VS3----------------------------
 'Arena1
If UserList(userindex).flags.DueleandoTorneo3 = True Then
If Arena1 = True Then
If UserList(Torne.Jugador1).flags.Muerto = 1 And UserList(Torne.Jugador2).flags.Muerto = 1 And UserList(Torne.Jugador3).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador1, 1, 50, 50)
Call WarpUserChar(Torne.Jugador2, 1, 50, 51)
Call WarpUserChar(Torne.Jugador3, 1, 50, 52)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 1: " & UserList(Torne.Jugador4).name & ", " & UserList(Torne.Jugador5).name & " y " & UserList(Torne.Jugador6).name & " ganaron el combate." & "~230~230~0~1~0")
Arena1 = False
UserList(Torne.Jugador1).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador2).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador3).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador4).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador5).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador6).flags.DueleandoTorneo3 = False
Torne.Jugador1 = 0
Torne.Jugador2 = 0
Torne.Jugador3 = 0
Torne.Jugador4 = 0
Torne.Jugador5 = 0
Torne.Jugador6 = 0
ElseIf UserList(Torne.Jugador4).flags.Muerto = 1 And UserList(Torne.Jugador5).flags.Muerto = 1 And UserList(Torne.Jugador6).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador4, 1, 50, 50)
Call WarpUserChar(Torne.Jugador5, 1, 50, 51)
Call WarpUserChar(Torne.Jugador6, 1, 50, 52)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 1: " & UserList(Torne.Jugador1).name & ", " & UserList(Torne.Jugador2).name & " y " & UserList(Torne.Jugador3).name & " ganaron el combate." & "~230~230~0~1~0")
Arena1 = False
UserList(Torne.Jugador1).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador2).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador3).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador4).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador5).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador6).flags.DueleandoTorneo3 = False
Torne.Jugador1 = 0
Torne.Jugador2 = 0
Torne.Jugador3 = 0
Torne.Jugador4 = 0
Torne.Jugador5 = 0
Torne.Jugador6 = 0
End If
End If

'Arena2

If Arena2 = True Then
If UserList(Torne.Jugador7).flags.Muerto = 1 And UserList(Torne.Jugador8).flags.Muerto = 1 And UserList(Torne.Jugador9).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador7, 1, 50, 50)
Call WarpUserChar(Torne.Jugador8, 1, 50, 51)
Call WarpUserChar(Torne.Jugador9, 1, 50, 52)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 2: " & UserList(Torne.Jugador10).name & ", " & UserList(Torne.Jugador11).name & " y " & UserList(Torne.Jugador12).name & " ganaron el combate." & "~230~230~0~1~0")
Arena2 = False
UserList(Torne.Jugador7).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador8).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador9).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador10).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador11).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador12).flags.DueleandoTorneo3 = False
Torne.Jugador7 = 0
Torne.Jugador8 = 0
Torne.Jugador9 = 0
Torne.Jugador10 = 0
Torne.Jugador11 = 0
Torne.Jugador12 = 0
ElseIf UserList(Torne.Jugador10).flags.Muerto = 1 And UserList(Torne.Jugador11).flags.Muerto = 1 And UserList(Torne.Jugador12).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador10, 1, 50, 50)
Call WarpUserChar(Torne.Jugador11, 1, 50, 51)
Call WarpUserChar(Torne.Jugador12, 1, 50, 52)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 2: " & UserList(Torne.Jugador7).name & ", " & UserList(Torne.Jugador8).name & " y " & UserList(Torne.Jugador9).name & " ganaron el combate." & "~230~230~0~1~0")
Arena2 = False
UserList(Torne.Jugador7).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador8).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador9).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador10).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador11).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador12).flags.DueleandoTorneo3 = False
Torne.Jugador7 = 0
Torne.Jugador8 = 0
Torne.Jugador9 = 0
Torne.Jugador10 = 0
Torne.Jugador11 = 0
Torne.Jugador12 = 0
End If
End If

'Arena3
If Arena3 = True Then
If UserList(Torne.Jugador13).flags.Muerto = 1 And UserList(Torne.Jugador14).flags.Muerto = 1 And UserList(Torne.Jugador15).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador13, 1, 50, 50)
Call WarpUserChar(Torne.Jugador14, 1, 50, 51)
Call WarpUserChar(Torne.Jugador15, 1, 50, 52)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 3: " & UserList(Torne.Jugador16).name & ", " & UserList(Torne.Jugador17).name & " y " & UserList(Torne.Jugador18).name & " ganaron el combate." & "~230~230~0~1~0")
Arena3 = False
UserList(Torne.Jugador13).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador14).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador15).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador16).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador17).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador18).flags.DueleandoTorneo3 = False
Torne.Jugador13 = 0
Torne.Jugador14 = 0
Torne.Jugador15 = 0
Torne.Jugador16 = 0
Torne.Jugador17 = 0
Torne.Jugador18 = 0
ElseIf UserList(Torne.Jugador16).flags.Muerto = 1 And UserList(Torne.Jugador17).flags.Muerto = 1 And UserList(Torne.Jugador18).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador16, 1, 50, 50)
Call WarpUserChar(Torne.Jugador17, 1, 50, 51)
Call WarpUserChar(Torne.Jugador18, 1, 50, 52)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 3: " & UserList(Torne.Jugador13).name & ", " & UserList(Torne.Jugador14).name & " y " & UserList(Torne.Jugador15).name & " ganaron el combate." & "~230~230~0~1~0")
Arena3 = False
UserList(Torne.Jugador13).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador14).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador15).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador16).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador17).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador18).flags.DueleandoTorneo3 = False
Torne.Jugador13 = 0
Torne.Jugador14 = 0
Torne.Jugador15 = 0
Torne.Jugador16 = 0
Torne.Jugador17 = 0
Torne.Jugador18 = 0
End If
End If
'Arena4
If Arena4 = True Then
If UserList(Torne.Jugador19).flags.Muerto = 1 And UserList(Torne.Jugador20).flags.Muerto = 1 And UserList(Torne.Jugador21).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador19, 1, 50, 50)
Call WarpUserChar(Torne.Jugador20, 1, 50, 51)
Call WarpUserChar(Torne.Jugador21, 1, 50, 52)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 4: " & UserList(Torne.Jugador22).name & ", " & UserList(Torne.Jugador23).name & " y " & UserList(Torne.Jugador24).name & " ganaron el combate." & "~230~230~0~1~0")
Arena4 = False
UserList(Torne.Jugador19).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador20).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador21).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador22).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador23).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador24).flags.DueleandoTorneo3 = False
Torne.Jugador19 = 0
Torne.Jugador20 = 0
Torne.Jugador21 = 0
Torne.Jugador22 = 0
Torne.Jugador23 = 0
Torne.Jugador24 = 0
ElseIf UserList(Torne.Jugador22).flags.Muerto = 1 And UserList(Torne.Jugador23).flags.Muerto = 1 And UserList(Torne.Jugador24).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador22, 1, 50, 50)
Call WarpUserChar(Torne.Jugador23, 1, 50, 51)
Call WarpUserChar(Torne.Jugador24, 1, 50, 52)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 4: " & UserList(Torne.Jugador19).name & ", " & UserList(Torne.Jugador20).name & " y " & UserList(Torne.Jugador21).name & " ganaron el combate." & "~230~230~0~1~0")
Arena4 = False
UserList(Torne.Jugador19).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador20).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador21).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador22).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador23).flags.DueleandoTorneo3 = False
UserList(Torne.Jugador24).flags.DueleandoTorneo3 = False
Torne.Jugador19 = 0
Torne.Jugador20 = 0
Torne.Jugador21 = 0
Torne.Jugador22 = 0
Torne.Jugador23 = 0
Torne.Jugador24 = 0
End If
End If
End If

'----------------------------------4VS4---------------------------------
'Arena1
If UserList(userindex).flags.DueleandoTorneo4 = True Then
If Arena1 = True Then
If UserList(Torne.Jugador1).flags.Muerto = 1 And UserList(Torne.Jugador2).flags.Muerto = 1 And UserList(Torne.Jugador3).flags.Muerto = 1 And UserList(Torne.Jugador4).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador1, 1, 50, 50)
Call WarpUserChar(Torne.Jugador2, 1, 50, 51)
Call WarpUserChar(Torne.Jugador3, 1, 50, 52)
Call WarpUserChar(Torne.Jugador4, 1, 50, 53)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 1: " & UserList(Torne.Jugador5).name & ", " & UserList(Torne.Jugador6).name & ", " & UserList(Torne.Jugador7).name & " y " & UserList(Torne.Jugador8).name & " ganaron el combate." & "~230~230~0~1~0")
Arena1 = False
UserList(Torne.Jugador1).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador2).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador3).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador4).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador5).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador6).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador7).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador8).flags.DueleandoTorneo4 = False
Torne.Jugador1 = 0
Torne.Jugador2 = 0
Torne.Jugador3 = 0
Torne.Jugador4 = 0
Torne.Jugador5 = 0
Torne.Jugador6 = 0
Torne.Jugador7 = 0
Torne.Jugador8 = 0
ElseIf UserList(Torne.Jugador5).flags.Muerto = 1 And UserList(Torne.Jugador6).flags.Muerto = 1 And UserList(Torne.Jugador7).flags.Muerto = 1 And UserList(Torne.Jugador8).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador5, 1, 50, 50)
Call WarpUserChar(Torne.Jugador6, 1, 50, 51)
Call WarpUserChar(Torne.Jugador7, 1, 50, 52)
Call WarpUserChar(Torne.Jugador8, 1, 50, 53)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 1: " & UserList(Torne.Jugador1).name & ", " & UserList(Torne.Jugador2).name & ", " & UserList(Torne.Jugador3).name & " y " & UserList(Torne.Jugador4).name & " ganaron el combate." & "~230~230~0~1~0")
Arena1 = False
UserList(Torne.Jugador1).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador2).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador3).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador4).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador5).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador6).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador7).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador8).flags.DueleandoTorneo4 = False
Torne.Jugador1 = 0
Torne.Jugador2 = 0
Torne.Jugador3 = 0
Torne.Jugador4 = 0
Torne.Jugador5 = 0
Torne.Jugador6 = 0
Torne.Jugador7 = 0
Torne.Jugador8 = 0
End If
End If

'Arena2
If Arena2 = True Then
If UserList(Torne.Jugador9).flags.Muerto = 1 And UserList(Torne.Jugador10).flags.Muerto = 1 And UserList(Torne.Jugador11).flags.Muerto = 1 And UserList(Torne.Jugador12).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador9, 1, 50, 50)
Call WarpUserChar(Torne.Jugador10, 1, 50, 51)
Call WarpUserChar(Torne.Jugador11, 1, 50, 52)
Call WarpUserChar(Torne.Jugador12, 1, 50, 53)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 2: " & UserList(Torne.Jugador13).name & ", " & UserList(Torne.Jugador14).name & ", " & UserList(Torne.Jugador15).name & " y " & UserList(Torne.Jugador16).name & " ganaron el combate." & "~230~230~0~1~0")
Arena2 = False
UserList(Torne.Jugador9).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador10).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador11).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador12).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador13).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador14).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador15).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador16).flags.DueleandoTorneo4 = False
Torne.Jugador9 = 0
Torne.Jugador10 = 0
Torne.Jugador11 = 0
Torne.Jugador12 = 0
Torne.Jugador13 = 0
Torne.Jugador14 = 0
Torne.Jugador15 = 0
Torne.Jugador16 = 0
ElseIf UserList(Torne.Jugador13).flags.Muerto = 1 And UserList(Torne.Jugador14).flags.Muerto = 1 And UserList(Torne.Jugador15).flags.Muerto = 1 And UserList(Torne.Jugador16).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador13, 1, 50, 50)
Call WarpUserChar(Torne.Jugador14, 1, 50, 51)
Call WarpUserChar(Torne.Jugador15, 1, 50, 52)
Call WarpUserChar(Torne.Jugador16, 1, 50, 53)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 2: " & UserList(Torne.Jugador9).name & ", " & UserList(Torne.Jugador10).name & ", " & UserList(Torne.Jugador11).name & " y " & UserList(Torne.Jugador12).name & " ganaron el combate." & "~230~230~0~1~0")
Arena2 = False
UserList(Torne.Jugador9).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador10).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador11).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador12).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador13).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador14).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador15).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador16).flags.DueleandoTorneo4 = False
Torne.Jugador9 = 0
Torne.Jugador10 = 0
Torne.Jugador11 = 0
Torne.Jugador12 = 0
Torne.Jugador13 = 0
Torne.Jugador14 = 0
Torne.Jugador15 = 0
Torne.Jugador16 = 0
End If
End If

'Arena3
If Arena3 = True Then
If UserList(Torne.Jugador17).flags.Muerto = 1 And UserList(Torne.Jugador18).flags.Muerto = 1 And UserList(Torne.Jugador19).flags.Muerto = 1 And UserList(Torne.Jugador20).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador17, 1, 50, 50)
Call WarpUserChar(Torne.Jugador18, 1, 50, 51)
Call WarpUserChar(Torne.Jugador19, 1, 50, 52)
Call WarpUserChar(Torne.Jugador20, 1, 50, 53)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 3: " & UserList(Torne.Jugador21).name & ", " & UserList(Torne.Jugador22).name & ", " & UserList(Torne.Jugador23).name & " y " & UserList(Torne.Jugador24).name & " ganaron el combate." & "~230~230~0~1~0")
Arena3 = False
UserList(Torne.Jugador17).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador18).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador19).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador20).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador21).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador22).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador23).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador24).flags.DueleandoTorneo4 = False
Torne.Jugador17 = 0
Torne.Jugador18 = 0
Torne.Jugador19 = 0
Torne.Jugador20 = 0
Torne.Jugador21 = 0
Torne.Jugador22 = 0
Torne.Jugador23 = 0
Torne.Jugador24 = 0
ElseIf UserList(Torne.Jugador21).flags.Muerto = 1 And UserList(Torne.Jugador22).flags.Muerto = 1 And UserList(Torne.Jugador23).flags.Muerto = 1 And UserList(Torne.Jugador24).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador21, 1, 50, 50)
Call WarpUserChar(Torne.Jugador22, 1, 50, 51)
Call WarpUserChar(Torne.Jugador23, 1, 50, 52)
Call WarpUserChar(Torne.Jugador24, 1, 50, 53)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 3: " & UserList(Torne.Jugador17).name & ", " & UserList(Torne.Jugador18).name & ", " & UserList(Torne.Jugador19).name & " y " & UserList(Torne.Jugador20).name & " ganaron el combate." & "~230~230~0~1~0")
Arena3 = False
UserList(Torne.Jugador17).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador18).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador19).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador20).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador21).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador22).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador23).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador24).flags.DueleandoTorneo4 = False
Torne.Jugador17 = 0
Torne.Jugador18 = 0
Torne.Jugador19 = 0
Torne.Jugador20 = 0
Torne.Jugador21 = 0
Torne.Jugador22 = 0
Torne.Jugador23 = 0
Torne.Jugador24 = 0
End If
End If

'Arena4
If Arena4 = True Then
If UserList(Torne.Jugador25).flags.Muerto = 1 And UserList(Torne.Jugador26).flags.Muerto = 1 And UserList(Torne.Jugador27).flags.Muerto = 1 And UserList(Torne.Jugador28).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador25, 1, 50, 50)
Call WarpUserChar(Torne.Jugador26, 1, 50, 51)
Call WarpUserChar(Torne.Jugador27, 1, 50, 52)
Call WarpUserChar(Torne.Jugador28, 1, 50, 53)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 4: " & UserList(Torne.Jugador29).name & ", " & UserList(Torne.Jugador30).name & ", " & UserList(Torne.Jugador31).name & " y " & UserList(Torne.Jugador32).name & " ganaron el combate." & "~230~230~0~1~0")
Arena4 = False
UserList(Torne.Jugador25).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador26).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador27).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador28).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador29).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador30).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador31).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador32).flags.DueleandoTorneo4 = False
Torne.Jugador25 = 0
Torne.Jugador26 = 0
Torne.Jugador27 = 0
Torne.Jugador28 = 0
Torne.Jugador29 = 0
Torne.Jugador30 = 0
Torne.Jugador31 = 0
Torne.Jugador32 = 0
ElseIf UserList(Torne.Jugador29).flags.Muerto = 1 And UserList(Torne.Jugador30).flags.Muerto = 1 And UserList(Torne.Jugador31).flags.Muerto = 1 And UserList(Torne.Jugador32).flags.Muerto = 1 Then
Call WarpUserChar(Torne.Jugador29, 1, 50, 50)
Call WarpUserChar(Torne.Jugador30, 1, 50, 51)
Call WarpUserChar(Torne.Jugador31, 1, 50, 52)
Call WarpUserChar(Torne.Jugador32, 1, 50, 53)
Call SendData(SendTarget.toall, 0, 0, "||Torneo: Arena 4: " & UserList(Torne.Jugador25).name & ", " & UserList(Torne.Jugador26).name & ", " & UserList(Torne.Jugador27).name & " y " & UserList(Torne.Jugador28).name & " ganaron el combate." & "~230~230~0~1~0")
Arena4 = False
UserList(Torne.Jugador25).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador26).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador27).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador28).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador29).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador30).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador31).flags.DueleandoTorneo4 = False
UserList(Torne.Jugador32).flags.DueleandoTorneo4 = False
Torne.Jugador25 = 0
Torne.Jugador26 = 0
Torne.Jugador27 = 0
Torne.Jugador28 = 0
Torne.Jugador29 = 0
Torne.Jugador30 = 0
Torne.Jugador31 = 0
Torne.Jugador32 = 0
End If
 End If
  End If
  
  If HayTD = True And UserList(userindex).flags.TeamTD <> 0 Then
  Call MuereUserTD(userindex)
  End If

    'Corte turra guachin
    If UserList(userindex).flags.EstaDueleando = True Then
        Call TerminarDuelo(UserList(userindex).flags.Oponente, userindex)
    End If
    
    ' << Reseteamos los posibles FX sobre el personaje >>
    If UserList(userindex).Char.loops = LoopAdEternum Then
        UserList(userindex).Char.FX = 0
        UserList(userindex).Char.loops = 0
    End If
    
    If UserList(userindex).flags.Automaticop = True Then
Call Rondas_UsuarioMuerep(userindex)
End If
    
    If UserList(userindex).flags.automatico = True Then
Call Rondas_UsuarioMuere(userindex)
End If
    
    ' << Restauramos el mimetismo
    If UserList(userindex).flags.Mimetizado = 1 Then
        UserList(userindex).Char.Body = UserList(userindex).CharMimetizado.Body
        UserList(userindex).Char.Head = UserList(userindex).CharMimetizado.Head
        UserList(userindex).Char.CascoAnim = UserList(userindex).CharMimetizado.CascoAnim
        UserList(userindex).Char.ShieldAnim = UserList(userindex).CharMimetizado.ShieldAnim
        UserList(userindex).Char.WeaponAnim = UserList(userindex).CharMimetizado.WeaponAnim
        UserList(userindex).Counters.Mimetismo = 0
        UserList(userindex).flags.Mimetizado = 0
    End If
    
'<< Cambiamos la apariencia del char >>
If UserList(userindex).flags.Navegando = 0 Then
If Criminal(userindex) Then
UserList(userindex).Char.Body = iCuerpoMuertoCrimi
UserList(userindex).Char.Head = iCabezaMuertoCrimi
ElseIf Ciudadano(userindex) Then
UserList(userindex).Char.Body = iCuerpoMuerto
UserList(userindex).Char.Head = iCabezaMuerto
ElseIf Neutral(userindex) Then
UserList(userindex).Char.Body = iCuerpoMuertoNeutro
UserList(userindex).Char.Head = iCabezaMuertoNeutro
End If
UserList(userindex).Char.ShieldAnim = NingunEscudo
UserList(userindex).Char.WeaponAnim = NingunArma
UserList(userindex).Char.CascoAnim = NingunCasco
Else
UserList(userindex).Char.Body = iFragataFantasmal ';)
End If
    
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        
        If UserList(userindex).MascotasIndex(i) > 0 Then
               If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                    Call MuereNpc(UserList(userindex).MascotasIndex(i), 0)
               Else
                    Npclist(UserList(userindex).MascotasIndex(i)).MaestroUser = 0
                    Npclist(UserList(userindex).MascotasIndex(i)).Movement = Npclist(UserList(userindex).MascotasIndex(i)).flags.OldMovement
                    Npclist(UserList(userindex).MascotasIndex(i)).Hostile = Npclist(UserList(userindex).MascotasIndex(i)).flags.OldHostil
                    UserList(userindex).MascotasIndex(i) = 0
                    UserList(userindex).MascotasType(i) = 0
               End If
        End If
        
    Next i
    
    UserList(userindex).NroMacotas = 0
    
    
    'If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
    '        Dim MiObj As Obj
    '        Dim nPos As WorldPos
    '        MiObj.ObjIndex = RandomNumber(554, 555)
    '        MiObj.Amount = 1
    '        nPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    '        Dim ManchaSangre As New cGarbage
    '        ManchaSangre.Map = nPos.Map
    '        ManchaSangre.X = nPos.X
    '        ManchaSangre.Y = nPos.Y
    '        Call TrashCollector.Add(ManchaSangre)
    'End If
    
    '<< Actualizamos clientes >>
    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, Val(userindex), UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
    Call SendUserStatsBox(userindex)
    
    
    '<<Castigos por party>>
    If UserList(userindex).PartyIndex > 0 Then
        Call mdParty.ObtenerExito(userindex, UserList(userindex).Stats.ELV * -10 * mdParty.CantMiembros(userindex), UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y)
    End If
    
        If UserList(userindex).pos.Map = 72 Then 'mapa de desafio
    Call WarpUserChar(userindex, 1, 50, 50, True) 'Poner el mapa en donde salen
    End If
    
    If MapInfo(UserList(userindex).pos.Map).Transport.Activo = 1 Then
    Call WarpUserChar(userindex, MapInfo(UserList(userindex).pos.Map).Transport.Map, MapInfo(UserList(userindex).pos.Map).Transport.X, MapInfo(UserList(userindex).pos.Map).Transport.Y, True)
    End If

      If UserList(userindex).EnCvc Then
            'Dim ijaji As Integer
            'For ijaji = 1 To LastUser
                With UserList(userindex)
                    If Guilds(.GuildIndex).GuildName = Nombre1 Then
                        If .EnCvc = True Then
                            If .flags.Muerto Then
                                Call WarpUserChar(userindex, 1, 50, 50, False)
                                modGuilds.UsuariosEnCvcClan1 = modGuilds.UsuariosEnCvcClan1 - 1
                                UserList(userindex).EnCvc = False
                                If modGuilds.UsuariosEnCvcClan1 = 0 Then
                                    Call SendData(SendTarget.toall, userindex, 0, "||" & "El clan " & Nombre2 & " derrotó al clan " & Nombre1 & "." & "~255~255~255~1~0")
                                    CvcFunciona = False
                                    Call LlevarUsuarios
                                End If
                            End If
                         End If
                     End If
                      
                
                    If Guilds(.GuildIndex).GuildName = Nombre2 Then
                        If .EnCvc = True Then
                            If .flags.Muerto Then
                                Call WarpUserChar(userindex, 1, 50, 50, False)
                                modGuilds.UsuariosEnCvcClan2 = modGuilds.UsuariosEnCvcClan2 - 1
                                UserList(userindex).EnCvc = False
                                If modGuilds.UsuariosEnCvcClan2 = 0 Then
                                    Call SendData(SendTarget.toall, userindex, 0, "||" & "El clan " & Nombre1 & " derrotó al clan " & Nombre2 & "." & "~255~255~255~1~0")
                                    CvcFunciona = False
                                    Call LlevarUsuarios
                                End If
                            End If
                        End If
                    End If
                End With
            'Next ijaji
    End If
    
    If UserList(userindex).flags.EstaDueleandoxset = True Then
    Call WarpUserChar(userindex, 1, 50, 50, True)
    Call SendData(SendTarget.toindex, UserList(userindex).flags.Oponentexset, 0, "||Puedes agarrar los items que has ganado, cuando termines tipea /GANE." & FONTTYPE_INFO)
    End If

UserList(userindex).Stats.Repu = UserList(userindex).Stats.Repu - 20
Call SendData(toindex, userindex, 0, "SOUND") 'Aca tambien cortamos el sonido de mierda ese
    
Exit Sub


ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.Description)
End Sub


Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

    If EsNewbie(Muerto) Then Exit Sub
    
    If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
    
    If UserList(Atacante).pos.Map = MapCastilloS Then Exit Sub
    If UserList(Atacante).pos.Map = MapCastilloN Then Exit Sub
    If UserList(Atacante).pos.Map = MapCastilloE Then Exit Sub
    If UserList(Atacante).pos.Map = MapCastilloO Then Exit Sub
    
    If Criminal(Muerto) Then
        If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).name Then
            UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).name
            If UserList(Atacante).Faccion.CriminalesMatados < 65000 Then _
                UserList(Atacante).Faccion.CriminalesMatados = UserList(Atacante).Faccion.CriminalesMatados + 1
        End If
        
        If UserList(Atacante).Faccion.CriminalesMatados > MAXUSERMATADOS Then
            UserList(Atacante).Faccion.CriminalesMatados = 0
            UserList(Atacante).Faccion.RecompensasReal = 0
        End If
        
        If UserList(Atacante).Faccion.RecibioExpInicialCaos = 1 And UserList(Muerto).Faccion.FuerzasCaos = 1 Then
            UserList(Atacante).Faccion.Reenlistadas = 200  'jaja que trucho
            
            'con esto evitamos que se vuelva a reenlistar
        End If
    Else
        If UserList(Atacante).flags.LastCiudMatado <> UserList(Muerto).name Then
            UserList(Atacante).flags.LastCiudMatado = UserList(Muerto).name
            If UserList(Atacante).Faccion.CiudadanosMatados < 65000 Then _
                UserList(Atacante).Faccion.CiudadanosMatados = UserList(Atacante).Faccion.CiudadanosMatados + 1
        End If
        
        If UserList(Atacante).Faccion.CiudadanosMatados > MAXUSERMATADOS Then
UserList(Atacante).Faccion.CiudadanosMatados = 0
End If
If Neutral(Muerto) Then
If UserList(Atacante).flags.LastNeutrMatado <> UserList(Muerto).name Then
UserList(Atacante).flags.LastNeutrMatado = UserList(Muerto).name
If UserList(Atacante).Faccion.NeutralesMatados < 65000 Then _
UserList(Atacante).Faccion.NeutralesMatados = UserList(Atacante).Faccion.NeutralesMatados + 1
End If
End If
 
If UserList(Atacante).Faccion.NeutralesMatados > MAXUSERMATADOS Then
UserList(Atacante).Faccion.NeutralesMatados = 0
End If
End If

UserList(Atacante).Stats.Repu = UserList(Atacante).Stats.Repu + 20

End Sub

Sub Tilelibre(ByRef pos As WorldPos, ByRef nPos As WorldPos, ByRef Obj As Obj)
'Call LogTarea("Sub Tilelibre")

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
    hayobj = False
    nPos.Map = pos.Map
    
    Do While Not LegalPos(pos.Map, nPos.X, nPos.Y) Or hayobj
        
        If LoopC > 15 Then
            Notfound = True
            Exit Do
        End If
        
        For tY = pos.Y - LoopC To pos.Y + LoopC
            For tX = pos.X - LoopC To pos.X + LoopC
            
                If LegalPos(nPos.Map, tX, tY) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex > 0 And MapData(nPos.Map, tX, tY).OBJInfo.ObjIndex <> Obj.ObjIndex)
                    If Not hayobj Then _
                        hayobj = (MapData(nPos.Map, tX, tY).OBJInfo.Amount + Obj.Amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                        nPos.X = tX
                        nPos.Y = tY
                        tX = pos.X + LoopC
                        tY = pos.Y + LoopC
                    End If
                End If
            
            Next tX
        Next tY
        
        LoopC = LoopC + 1
        
    Loop
    
    If Notfound = True Then
        nPos.X = 0
        nPos.Y = 0
    End If

End Sub

Sub WarpUserChar(ByVal userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)

Dim OldMap As Integer
Dim OldX As Integer
Dim OldY As Integer

    'Quitar el dialogo
    Call SendToUserArea(userindex, "ULZ" & UserList(userindex).Char.CharIndex)
    Call SendData(SendTarget.toindex, userindex, UserList(userindex).pos.Map, "QTDL")
    
    OldMap = UserList(userindex).pos.Map
    OldX = UserList(userindex).pos.X
    OldY = UserList(userindex).pos.Y
    
    Call EraseUserChar(SendTarget.ToMap, 0, OldMap, userindex)
        
    If OldMap <> Map Then
        Call SendData(SendTarget.toindex, userindex, 0, "CM" & Map & "," & MapInfo(UserList(userindex).pos.Map).MapVersion)
        Call SendData(SendTarget.toindex, userindex, 0, "TM" & MapInfo(Map).Music)
        Call SendData(SendTarget.toindex, userindex, 0, "N~" & MapInfo(Map).name)
        Call SendData(SendTarget.toindex, userindex, 0, "N~" & MapInfo(Map).name)

        
        'Update new Map Users
        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
    
        'Update old Map Users
        MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
        If MapInfo(OldMap).NumUsers < 0 Then
            MapInfo(OldMap).NumUsers = 0
        End If
    End If
       If MapData(Map, X, Y).userindex > 0 Then
    X = X + 1
    End If
    UserList(userindex).pos.X = X
    UserList(userindex).pos.Y = Y
    UserList(userindex).pos.Map = Map
    
    Call MakeUserChar(SendTarget.ToMap, 0, Map, userindex, Map, X, Y)
    Call SendData(SendTarget.toindex, userindex, 0, "IP" & UserList(userindex).Char.CharIndex)
    
    'Seguis invisible al pasar de mapa
    If (UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1) And (Not UserList(userindex).flags.AdminInvisible = 1) Then
        Call SendToUserArea(userindex, "NOVER" & UserList(userindex).Char.CharIndex & ",1", EncriptarProtocolosCriticos)
    End If
    
    If FX And UserList(userindex).flags.AdminInvisible = 0 Then 'FX
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_WARP)
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXWARP & ",0")
    End If
    
    Call WarpMascotas(userindex)
End Sub

Sub UpdateUserMap(ByVal userindex As Integer)

Dim Map As Integer
Dim X As Integer
Dim Y As Integer

On Error GoTo 0

Map = UserList(userindex).pos.Map

For Y = YMinMapSize To YMaxMapSize
    For X = XMinMapSize To XMaxMapSize
        If MapData(Map, X, Y).userindex > 0 And userindex <> MapData(Map, X, Y).userindex Then
            Call MakeUserChar(SendTarget.toindex, userindex, 0, MapData(Map, X, Y).userindex, Map, X, Y)
#If SeguridadAlkon Then
            If EncriptarProtocolosCriticos Then
                If UserList(MapData(Map, X, Y).userindex).flags.Invisible = 1 Or UserList(MapData(Map, X, Y).userindex).flags.Oculto = 1 Then Call SendCryptedData(SendTarget.toindex, userindex, 0, "NOVER" & UserList(MapData(Map, X, Y).userindex).Char.CharIndex & ",1")
            Else
#End If
                If UserList(MapData(Map, X, Y).userindex).flags.Invisible = 1 Or UserList(MapData(Map, X, Y).userindex).flags.Oculto = 1 Then Call SendData(SendTarget.toindex, userindex, 0, "NOVER" & UserList(MapData(Map, X, Y).userindex).Char.CharIndex & ",1")
#If SeguridadAlkon Then
            End If
#End If
        End If

        If MapData(Map, X, Y).NpcIndex > 0 Then
            Call MakeNPCChar(SendTarget.toindex, userindex, 0, MapData(Map, X, Y).NpcIndex, Map, X, Y)
        End If

        If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
            If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType <> eOBJType.otArboles Then
                Call MakeObj(SendTarget.toindex, userindex, 0, MapData(Map, X, Y).OBJInfo, Map, X, Y)
                If ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                          Call Bloquear(SendTarget.toindex, userindex, 0, Map, X, Y, MapData(Map, X, Y).Blocked)
                          Call Bloquear(SendTarget.toindex, userindex, 0, Map, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
                End If
            End If
        End If
        
    Next X
Next Y

End Sub


Sub WarpMascotas(ByVal userindex As Integer)
Dim i As Integer

Dim UMascRespawn  As Boolean
Dim miflag As Byte, MascotasReales As Integer
Dim prevMacotaType As Integer

Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer, InvocadosMatados As Integer

NroPets = UserList(userindex).NroMacotas
InvocadosMatados = 0

    'Matamos los invocados
    '[Alejo 18-03-2004]
    For i = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasIndex(i) > 0 Then
            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada.
            If Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(UserList(userindex).MascotasIndex(i))
                UserList(userindex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
            End If
        End If
    Next i
    
    If InvocadosMatados > 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Pierdes el control de tus mascotas." & FONTTYPE_INFO)
        UserList(userindex).flags.EleDeAgua = 0
        UserList(userindex).flags.EleDeFuego = 0
        UserList(userindex).flags.EleDeTierra = 0
    End If
    
    For i = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasIndex(i) > 0 Then
            PetRespawn(i) = Npclist(UserList(userindex).MascotasIndex(i)).flags.Respawn = 0
            PetTypes(i) = UserList(userindex).MascotasType(i)
            PetTiempoDeVida(i) = Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia
            Call QuitarNPC(UserList(userindex).MascotasIndex(i))
        End If
    Next i
    
    For i = 1 To MAXMASCOTAS
        If PetTypes(i) > 0 Then
            UserList(userindex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(userindex).pos, False, PetRespawn(i))
            UserList(userindex).MascotasType(i) = PetTypes(i)
            'Controlamos que se sumoneo OK
            If UserList(userindex).MascotasIndex(i) = 0 Then
                UserList(userindex).MascotasIndex(i) = 0
                UserList(userindex).MascotasType(i) = 0
                If UserList(userindex).NroMacotas > 0 Then UserList(userindex).NroMacotas = UserList(userindex).NroMacotas - 1
                Exit Sub
            End If
            Npclist(UserList(userindex).MascotasIndex(i)).MaestroUser = userindex
            Npclist(UserList(userindex).MascotasIndex(i)).Movement = TipoAI.SigueAmo
            Npclist(UserList(userindex).MascotasIndex(i)).Target = 0
            Npclist(UserList(userindex).MascotasIndex(i)).TargetNPC = 0
            Npclist(UserList(userindex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
            Call FollowAmo(UserList(userindex).MascotasIndex(i))
        End If
    Next i
    
    UserList(userindex).NroMacotas = NroPets

End Sub


Sub RepararMascotas(ByVal userindex As Integer)
Dim i As Integer
Dim MascotasReales As Integer

    For i = 1 To MAXMASCOTAS
      If UserList(userindex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
    Next i
    
    If MascotasReales <> UserList(userindex).NroMacotas Then UserList(userindex).NroMacotas = 0

End Sub

Sub Cerrar_Usuario(ByVal userindex As Integer, Optional ByVal Tiempo As Integer = -1)

If UserList(userindex).flags.Stopped Then Exit Sub

    If Tiempo = -1 Then Tiempo = IntervaloCerrarConexion
    
    If UserList(userindex).flags.UserLogged And Not UserList(userindex).Counters.Saliendo Then
        UserList(userindex).Counters.Saliendo = True
        UserList(userindex).Counters.Salir = IIf(UserList(userindex).flags.Privilegios > PlayerType.User Or Not MapInfo(UserList(userindex).pos.Map).Pk, 0, Tiempo)
        
        
        Call SendData(SendTarget.toindex, userindex, 0, "||Cerrando...Se cerrará el juego en " & UserList(userindex).Counters.Salir & " segundos..." & FONTTYPE_INFO)
    End If
    
    If UserList(userindex).pos.Map = 72 Then 'cambiar el 12 por numero de mapa
If UserList(userindex).flags.EnDesafio = 1 And MapInfo(72).NumUsers = 2 Then 'cambiar el 12 por numero de mapa
Call WarpUserChar(Desafio.Primero, 1, 64, 45, True) 'mapa donde lleva al creador del desafio
Call WarpUserChar(Desafio.Segundo, 1, 67, 45, True) 'mapa donde llevar al retador
UserList(Desafio.Primero).flags.EnDesafio = 0
UserList(Desafio.Primero).flags.rondas = 0
UserList(Desafio.Segundo).flags.Desafio = 0
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(Desafio.Primero).name & " ha abandonado el desafio." & FONTTYPE_GRISN)
ElseIf UserList(userindex).flags.EnDesafio = 1 And MapInfo(72).NumUsers = 1 Then 'cambiar el 12 por numero de mapa Then
Call WarpUserChar(Desafio.Primero, 1, 64, 45, True) 'mapa donde lleva al creador del desafio
UserList(Desafio.Primero).flags.EnDesafio = 0
UserList(Desafio.Primero).flags.rondas = 0
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(Desafio.Primero).name & " ha abandonado el desafio." & FONTTYPE_GRISN)
Else
If UserList(userindex).flags.Desafio = 1 And MapInfo(72).NumUsers = 2 Then 'cambiar el 12 por numero de mapa Then
Call WarpUserChar(Desafio.Segundo, 1, 67, 45, True) 'mapa donde lleva al retador
UserList(Desafio.Segundo).flags.Desafio = 0
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(Desafio.Segundo).name & " ha abandonado el desafio." & FONTTYPE_INFO)
Exit Sub
End If
Exit Sub
End If
Exit Sub
End If

 'salio en duelo by Feer
        If UserList(userindex).flags.EnDuelo Then
        
        Dim uDuelo1     As Integer
        Dim uDuelo2     As Integer
        
        uDuelo2 = NameIndex(UserList(userindex).flags.DueliandoContra)
        uDuelo1 = userindex
        
        'Reset Duelo Usuario Perdedor
        UserList(uDuelo1).flags.EnDuelo = False
        UserList(uDuelo1).flags.DueliandoContra = ""
        UserList(uDuelo1).flags.LeMandaronDuelo = False
        UserList(uDuelo1).flags.UltimoEnMandarDuelo = ""
        'Reset Duelo Usuario Perdedor
        'Set Usuario Ganador
        UserList(uDuelo2).flags.EnDuelo = False
        UserList(uDuelo2).flags.DueliandoContra = ""
        'Set Usuario Ganador
        'Set Todo
        SendData SendTarget.toall, userindex, 0, "||Duelos: El duelo fue cancelado por la desconeccion de " & UserList(uDuelo1).name & "." & "~255~255~255~0~1"
        WarpUserChar uDuelo1, PosUserDuelo2.Map, PosUserDuelo2.X, PosUserDuelo2.Y, True 'No jodan, esta al revez porque a mi se me canta la chota ~ Feer~
        WarpUserChar uDuelo2, PosUserDuelo1.Map, PosUserDuelo1.X, PosUserDuelo1.Y, True 'No jodan, esta al revez porque a mi se me canta la chota ~ Feer~
    End If
    'salio en duelo by Feer

            'casted - pareja 2vs2
If UserList(userindex).pos.Map = 54 Then 'mapa de pareja
If MapInfo(54).NumUsers = 2 And UserList(userindex).flags.EnPareja = True Then 'mapa de duelos 2vs2
        Call WarpUserChar(Pareja.Jugador1, PosUserPareja1.Map, PosUserPareja1.X, PosUserPareja1.Y)
        Call WarpUserChar(Pareja.Jugador2, PosUserPareja2.Map, PosUserPareja2.X, PosUserPareja2.Y)
            Call SendData(SendTarget.toall, 0, 0, "||" & UserList(Pareja.Jugador1).name & " y " & UserList(Pareja.Jugador2).name & " abandonaron el duelo 2 vs 2." & FONTTYPE_GUILD)
            UserList(Pareja.Jugador1).flags.EnPareja = False
            UserList(Pareja.Jugador1).flags.EsperaPareja = False
            UserList(Pareja.Jugador1).flags.SuPareja = 0
            UserList(Pareja.Jugador2).flags.EnPareja = False
            UserList(Pareja.Jugador2).flags.EsperaPareja = False
            UserList(Pareja.Jugador2).flags.SuPareja = 0
            HayPareja = False
            Exit Sub
            End If
End If


If UserList(userindex).pos.Map = 54 Then 'mapa de pareja
If MapInfo(54).NumUsers = 4 And UserList(userindex).flags.EnPareja = True Then
        Call WarpUserChar(Pareja.Jugador1, PosUserPareja1.Map, PosUserPareja1.X, PosUserPareja1.Y)
        Call WarpUserChar(Pareja.Jugador2, PosUserPareja2.Map, PosUserPareja2.X, PosUserPareja2.Y)
        Call WarpUserChar(Pareja.Jugador3, PosUserPareja3.Map, PosUserPareja3.X, PosUserPareja3.Y)
        Call WarpUserChar(Pareja.Jugador4, PosUserPareja4.Map, PosUserPareja4.X, PosUserPareja4.Y)
        UserList(Pareja.Jugador1).flags.EnPareja = False
        UserList(Pareja.Jugador1).flags.EsperaPareja = False
        UserList(Pareja.Jugador1).flags.SuPareja = 0
        UserList(Pareja.Jugador2).flags.EnPareja = False
        UserList(Pareja.Jugador2).flags.EsperaPareja = False
        UserList(Pareja.Jugador2).flags.SuPareja = 0
        UserList(Pareja.Jugador3).flags.EnPareja = False
        UserList(Pareja.Jugador3).flags.EsperaPareja = False
        UserList(Pareja.Jugador3).flags.SuPareja = 0
        UserList(Pareja.Jugador4).flags.EnPareja = False
        UserList(Pareja.Jugador4).flags.EsperaPareja = False
        UserList(Pareja.Jugador4).flags.SuPareja = 0
        HayPareja = False
        Call SendData(SendTarget.toall, 0, 0, "||El duelo 2 vs 2 se ha cancelado por la desconeccion de algun usuario." & FONTTYPE_TALK)
            Exit Sub
        End If
        End If
    
    If UserList(userindex).flags.EstaDueleando = True Then
    Call DesconectarDuelo(UserList(userindex).flags.Oponente, userindex)
    End If
    
    If UserList(userindex).flags.EstaDueleandoxset = True Then
    Call WarpUserChar(userindex, 1, 50, 50, True)
    Call WarpUserChar(UserList(userindex).flags.Oponentexset, 1, 51, 48, True)
    Call DesconectarDueloxset(UserList(userindex).flags.Oponente, userindex)
    End If
    
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal userindex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
Dim ViejoNick As String
Dim ViejoCharBackup As String

If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
ViejoNick = UserList(UserIndexDestino).name

If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
    'hace un backup del char
    ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
    Name CharPath & ViejoNick & ".chr" As ViejoCharBackup
End If

End Sub

Public Sub Empollando(ByVal userindex As Integer)
If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).OBJInfo.ObjIndex > 0 Then
    UserList(userindex).flags.EstaEmpo = 1
Else
    UserList(userindex).flags.EstaEmpo = 0
    UserList(userindex).EmpoCont = 0
End If

End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal Nombre As String)

If FileExist(CharPath & Nombre & ".chr", vbArchive) = False Then
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Pj Inexistente" & FONTTYPE_INFO)
Else
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Estadisticas de: " & Nombre & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Nivel: " & GetVar(CharPath & Nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "elu") & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Vitalidad: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "maxSta") & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxHP") & "  Mana: " & GetVar(CharPath & Nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN") & FONTTYPE_INFO)
    
    Call SendData(SendTarget.toindex, sendIndex, 0, "||Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", "MaxHIT") & FONTTYPE_INFO)
    
End If
Exit Sub

End Sub

Sub LlevarUsuarios()
Dim ijaji As Integer
For ijaji = 1 To LastUser
If UserList(ijaji).pos.Map = 8 And UserList(ijaji).EnCvc = True Then
    Call WarpUserChar(ijaji, 1, RandomNumber(44, 35), RandomNumber(52, 41), False)
    UserList(ijaji).EnCvc = False
End If
Next ijaji
End Sub

Private Function encriptarMpUsuario(ByVal CharIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As String
Dim Key As Byte
Key = (X Xor Y) + 4
encriptarMpUsuario = Chr$((Int(CharIndex / 128) Xor X) + 16) & _
                Chr$((X Xor Key) + 4) & _
                    Chr$((Int(CharIndex Mod 128) Xor Y) + 16) & _
                        Chr$((Y Xor Key) + 4) & _
                            Chr$(Key Xor &HFD&)
End Function




