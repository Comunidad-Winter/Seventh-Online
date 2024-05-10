Attribute VB_Name = "Extra"
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

Public Function EsNewbie(ByVal userindex As Integer) As Boolean
EsNewbie = UserList(userindex).Stats.ELV <= LimiteNewbie
End Function



Public Sub DoTileEvents(ByVal userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

On Error GoTo errhandler

Dim nPos As WorldPos
Dim FxFlag As Boolean
'Controla las salidas
If InMapBounds(Map, X, Y) Then
    
    If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
        FxFlag = ObjData(MapData(Map, X, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otTeleport
    End If
    
    If MapData(Map, X, Y).TileExit.Map > 0 Then
        '¿Es mapa de newbies?
        If UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Restringir) = "SI" Then
            '¿El usuario es un newbie?
            If EsNewbie(userindex) Then
                If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(userindex)) Then
                    If FxFlag Then '¿FX?
                        Call WarpUserChar(userindex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                    Else
                        Call WarpUserChar(userindex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
                    End If
                Else
                    Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        If FxFlag Then
                            Call WarpUserChar(userindex, nPos.Map, nPos.X, nPos.Y, True)
                        Else
                            Call WarpUserChar(userindex, nPos.Map, nPos.X, nPos.Y)
                        End If
                    End If
                End If
            Else 'No es newbie
                Call SendData(SendTarget.toindex, userindex, 0, "||Mapa exclusivo para newbies." & FONTTYPE_INFO)
                Dim veces As Byte
                veces = 0
                Call ClosestStablePos(UserList(userindex).pos, nPos)

                If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(userindex, nPos.Map, nPos.X, nPos.Y)
                End If
            End If
        Else 'No es un mapa de newbies
            If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(userindex)) Then
                 If MapData(Map, X, Y).TileExit.Map = 31 Or MapData(Map, X, Y).TileExit.Map = 32 Or MapData(Map, X, Y).TileExit.Map = 33 Or MapData(Map, X, Y).TileExit.Map = 34 Then
                         If Not UserList(userindex).GuildIndex <> 0 Then
                                Call SendData(SendTarget.toindex, userindex, 0, "||Debes tener clan para entrar al castillo." & FONTTYPE_INFO)
                                Call ClosestStablePos(UserList(userindex).pos, nPos)
                             If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(userindex, nPos.Map, nPos.X, nPos.Y, True)
                             End If
                             Exit Sub
                         End If
                         End If
                If FxFlag Then
                    Call WarpUserChar(userindex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                Else
                    Call WarpUserChar(userindex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
                End If
            Else
                Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                If nPos.X <> 0 And nPos.Y <> 0 Then
                    If FxFlag Then
                        Call WarpUserChar(userindex, nPos.Map, nPos.X, nPos.Y, True)
                    Else
                        Call WarpUserChar(userindex, nPos.Map, nPos.X, nPos.Y)
                    End If
                End If
            End If
        End If
    End If
    
End If

Exit Sub

errhandler:
    Call LogError("Error en DotileEvents")

End Sub

Function InRangoVision(ByVal userindex As Integer, X As Integer, Y As Integer) As Boolean

If X > UserList(userindex).pos.X - MinXBorder And X < UserList(userindex).pos.X + MinXBorder Then
    If Y > UserList(userindex).pos.Y - MinYBorder And Y < UserList(userindex).pos.Y + MinYBorder Then
        InRangoVision = True
        Exit Function
    End If
End If
InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, X As Integer, Y As Integer) As Boolean

If X > Npclist(NpcIndex).pos.X - MinXBorder And X < Npclist(NpcIndex).pos.X + MinXBorder Then
    If Y > Npclist(NpcIndex).pos.Y - MinYBorder And Y < Npclist(NpcIndex).pos.Y + MinYBorder Then
        InRangoVisionNPC = True
        Exit Function
    End If
End If
InRangoVisionNPC = False

End Function


Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapBounds = False
Else
    InMapBounds = True
End If

End Function

Sub ClosestLegalPos(pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = pos.Map

Do While Not LegalPos(pos.Map, nPos.X, nPos.Y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
     For tY = pos.Y + LoopC To pos.Y + LoopC
        For tX = pos.X - LoopC To pos.X + LoopC
            
            If LegalPos(nPos.Map, tX, tY) Then
                nPos.X = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = pos.X + LoopC
                tY = pos.Y + LoopC
  
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

Sub ClosestStablePos(pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = pos.Map

Do While Not LegalPos(pos.Map, nPos.X, nPos.Y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = pos.Y + LoopC To pos.Y + LoopC
        For tX = pos.X - LoopC To pos.X + LoopC
            
            If LegalPos(nPos.Map, tX, tY) And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                nPos.X = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = pos.X + LoopC
                tY = pos.Y + LoopC
  
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

Function NameIndex(ByRef name As String) As Integer

Dim userindex As Integer
'¿Nombre valido?
If name = "" Then
    NameIndex = 0
    Exit Function
End If

name = UCase$(Replace(name, "+", " "))

userindex = 1
Do Until UCase$(UserList(userindex).name) = name
    
    userindex = userindex + 1
    
    If userindex > MaxUsers Then
        NameIndex = 0
        Exit Function
    End If
    
Loop
 
NameIndex = userindex
 
End Function



Function IP_Index(ByVal inIP As String) As Integer
 
Dim userindex As Integer
'¿Nombre valido?
If inIP = "" Then
    IP_Index = 0
    Exit Function
End If
  
userindex = 1
Do Until UserList(userindex).ip = inIP
    
    userindex = userindex + 1
    
    If userindex > MaxUsers Then
        IP_Index = 0
        Exit Function
    End If
    
Loop
 
IP_Index = userindex

Exit Function

End Function


Function CheckForSameIP(ByVal userindex As Integer, ByVal UserIP As String) As Boolean
Dim LoopC As Integer
For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged = True Then
        If UserList(LoopC).ip = UserIP And userindex <> LoopC Then
            CheckForSameIP = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameIP = False
End Function

Function CheckForSameName(ByVal userindex As Integer, ByVal name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
Dim LoopC As Long
For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged Then
        
        'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
        'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
        'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
        'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
        'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
        
        If UCase$(UserList(LoopC).name) = UCase$(name) Then
            CheckForSameName = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameName = False
End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef pos As WorldPos)
'*****************************************************************
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
Dim X As Integer
Dim Y As Integer
Dim tempVar As Single
Dim nX As Integer
Dim nY As Integer

X = pos.X
Y = pos.Y

If Head = eHeading.NORTH Then
    nX = X
    nY = Y - 1
End If

If Head = eHeading.SOUTH Then
    nX = X
    nY = Y + 1
End If

If Head = eHeading.EAST Then
    nX = X + 1
    nY = Y
End If

If Head = eHeading.WEST Then
    nX = X - 1
    nY = Y
End If

'Devuelve valores
pos.X = nX
pos.Y = nY

End Sub

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False) As Boolean

'¿Es un mapa valido?
If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
            LegalPos = False
Else
  
  If Not PuedeAgua Then
        LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                   (MapData(Map, X, Y).userindex = 0) And _
                   (MapData(Map, X, Y).NpcIndex = 0) And _
                   (Not HayAgua(Map, X, Y))
  Else
        LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                   (MapData(Map, X, Y).userindex = 0) And _
                   (MapData(Map, X, Y).NpcIndex = 0) And _
                   (HayAgua(Map, X, Y))
  End If
   
End If

End Function

Function LegalPosNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte) As Boolean

If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
    LegalPosNPC = False
Else

 If AguaValida = 0 Then
   LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And _
     (MapData(Map, X, Y).userindex = 0) And _
     (MapData(Map, X, Y).NpcIndex = 0) And _
     (MapData(Map, X, Y).trigger <> eTrigger.POSINVALIDA) _
     And Not HayAgua(Map, X, Y)
 Else
   LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And _
     (MapData(Map, X, Y).userindex = 0) And _
     (MapData(Map, X, Y).NpcIndex = 0) And _
     (MapData(Map, X, Y).trigger <> eTrigger.POSINVALIDA)
 End If
 
End If


End Function

Sub SendHelp(ByVal index As Integer)
Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = Val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call SendData(SendTarget.toindex, index, 0, "||" & GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC) & FONTTYPE_INFO)
Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal userindex As Integer)
    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & Npclist(NpcIndex).Expresiones(randomi) & "°" & Npclist(NpcIndex).Char.CharIndex & FONTTYPE_INFO)
    End If
End Sub

Sub LookatTile(ByVal userindex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

'Responde al click del usuario sobre el mapa
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim OBJType As Integer

'¿Posicion valida?
If InMapBounds(Map, X, Y) Then
    UserList(userindex).flags.TargetMap = Map
    UserList(userindex).flags.TargetX = X
    UserList(userindex).flags.TargetY = Y
    '¿Es un obj?
    If MapData(Map, X, Y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
        UserList(userindex).flags.TargetObjMap = Map
        UserList(userindex).flags.TargetObjX = X
        UserList(userindex).flags.TargetObjY = Y
        FoundSomething = 1
    ElseIf MapData(Map, X + 1, Y).OBJInfo.ObjIndex > 0 Then
        'Informa el nombre
        If ObjData(MapData(Map, X + 1, Y).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            UserList(userindex).flags.TargetObjMap = Map
            UserList(userindex).flags.TargetObjX = X + 1
            UserList(userindex).flags.TargetObjY = Y
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(userindex).flags.TargetObjMap = Map
            UserList(userindex).flags.TargetObjX = X + 1
            UserList(userindex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X, Y + 1).OBJInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, X, Y + 1).OBJInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(userindex).flags.TargetObjMap = Map
            UserList(userindex).flags.TargetObjX = X
            UserList(userindex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    End If
    
    If FoundSomething = 1 Then
        UserList(userindex).flags.TargetObj = MapData(Map, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex
        If MostrarCantidad(UserList(userindex).flags.TargetObj) Then
            Call SendData(SendTarget.toindex, userindex, 0, "||" & ObjData(UserList(userindex).flags.TargetObj).name & " - " & MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.Amount & "" & FONTTYPE_INFO)
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||" & ObjData(UserList(userindex).flags.TargetObj).name & FONTTYPE_INFO)
        End If
    
    End If
    '¿Es un personaje?
    If Y + 1 <= YMaxMapSize Then
        If MapData(Map, X, Y + 1).userindex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).userindex
            If UserList(TempCharIndex).showName Then    ' Es GM y pidió que se oculte su nombre??
                FoundChar = 1
            End If
        End If
        If MapData(Map, X, Y + 1).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
            FoundChar = 2
        End If
    End If
    '¿Es un personaje?
    If FoundChar = 0 Then
        If MapData(Map, X, Y).userindex > 0 Then
            TempCharIndex = MapData(Map, X, Y).userindex
            If UserList(TempCharIndex).showName Then    ' Es GM y pidió que se oculte su nombre??
                FoundChar = 1
            End If
        End If
        If MapData(Map, X, Y).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y).NpcIndex
            FoundChar = 2
        End If
    End If
    
    
    'Reaccion al personaje
    If FoundChar = 1 Then '  ¿Encontro un Usuario?
            
       If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(userindex).flags.Privilegios = PlayerType.Dios Then
            
                If EsNewbie(TempCharIndex) Then
                    Stat = Stat & " <NEWBIE>"
                End If
                
                If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                    Stat = Stat & " <Ejército Real> " & "<" & TituloReal(TempCharIndex) & ">"
                ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                    Stat = Stat & " <Legión Oscura> " & "<" & TituloCaos(TempCharIndex) & ">"
                End If
                
                If UserList(TempCharIndex).GuildIndex > 0 Then
                    Stat = Stat & " <" & Guilds(UserList(TempCharIndex).GuildIndex).GuildName & ">"
                End If
                
                If Len(UserList(TempCharIndex).Desc) > 1 Then
                    Stat = "Ves a " & UserList(TempCharIndex).name & Stat & " - " & UserList(TempCharIndex).Desc
                Else
                    Stat = "Ves a " & UserList(TempCharIndex).name & Stat
                End If
                
                If UserList(TempCharIndex).flags.Privilegios = 1 Then
                        Stat = Stat & " <VIP>"
                End If
                
                If UserList(TempCharIndex).flags.Muerto Then
                    Stat = Stat & " [Muerto]"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.2) Then
                    Stat = Stat & " [Agonizando]"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.45) Then
                    Stat = Stat & " [Gravemente herido]"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.75) Then
                    Stat = Stat & " [Medio herido]"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP) Then
                    Stat = Stat & " [Algo lastimado]"
                Else
                    Stat = Stat & " [Intacto]"
                End If
                
                'Mithrandir - Sistema de Consejos
If UserList(TempCharIndex).ConsejoInfo.PertAlCons > 0 Then
'Es lider?
If UserList(TempCharIndex).ConsejoInfo.LiderConsejo > 0 Then
Stat = Stat & " [Lider del Bien]" & FONTTYPE_CONSEJOVesA
'Si no es lider... es miembro
Else
Stat = Stat & " [Consejo de la Luz]" & FONTTYPE_CONSEJOVesA
End If
ElseIf UserList(TempCharIndex).ConsejoInfo.PertAlConsCaos > 0 Then
'Es lider?
If UserList(TempCharIndex).ConsejoInfo.LiderConsejoCaos > 0 Then
Stat = Stat & " [Lider del Caos]" & FONTTYPE_CONSEJOCAOSVesA
Else
Stat = Stat & " [Consejo de las Sombras]" & FONTTYPE_CONSEJOCAOSVesA
End If
End If
'Mithrandir - Sistema de Consejos
                
                If UserList(TempCharIndex).flags.Privilegios > 3 Then
                        Stat = Stat & " <Administrador> ~156~3~152~1~0"
                ElseIf UserList(TempCharIndex).flags.Privilegios > 2 Then
                        Stat = Stat & " <Dios> ~250~250~150~1~0"
                ElseIf UserList(TempCharIndex).flags.Privilegios > 1 Then
                        Stat = Stat & " <Event Master> ~30~255~150~1~0"
                ElseIf Criminal(TempCharIndex) Then
                        Stat = Stat & " <CRIMINAL> ~255~0~0~1~0"
                ElseIf Ciudadano(TempCharIndex) Then
                        Stat = Stat & " <CIUDADANO> ~60~94~255~1~0"
                ElseIf Neutral(TempCharIndex) Then
                        Stat = Stat & " <NEUTRAL> ~120~120~120~1~0"
                End If
                

            Else
                Stat = UserList(TempCharIndex).DescRM & " " & FONTTYPE_INFOBOLD
            End If
            
            If Len(Stat) > 0 Then _
                Call SendData(SendTarget.toindex, userindex, 0, "||" & Stat)

            FoundSomething = 1
            UserList(userindex).flags.TargetUser = TempCharIndex
            UserList(userindex).flags.TargetNPC = 0
            UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
       End If

    End If
    If FoundChar = 2 Then '¿Encontro un NPC?
            Dim estatus As String
            
            If UserList(userindex).flags.Privilegios >= PlayerType.SemiDios Then
                estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ")"
            Else
                If UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) <= 10 Then
                    estatus = "(Dudoso) "
                ElseIf UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) > 10 And UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) <= 20 Then
                    If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP / 2) Then
                        estatus = "(Herido) "
                    Else
                        estatus = "(Sano) "
                    End If
                ElseIf UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) > 20 And UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) <= 30 Then
                    If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                        estatus = "(Malherido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                        estatus = "(Herido) "
                    Else
                        estatus = "(Sano) "
                    End If
                ElseIf UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) > 30 And UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) <= 40 Then
                    If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.25) Then
                        estatus = "(Muy malherido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                        estatus = "(Herido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                        estatus = "(Levemente herido) "
                    Else
                        estatus = "(Sano) "
                    End If
                ElseIf UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) > 40 And UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) < 60 Then
                    If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.05) Then
                        estatus = "(Agonizando) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.1) Then
                        estatus = "(Casi muerto) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.25) Then
                        estatus = "(Muy Malherido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                        estatus = "(Herido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                        estatus = "(Levemente herido) "
                    ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP) Then
                        estatus = "(Sano) "
                    Else
                        estatus = "(Intacto) "
                    End If
                ElseIf UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) >= 60 Then
                    estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") "
                Else
                    estatus = "!error!"
                End If
            End If
            
            If Npclist(TempCharIndex).pos.Map = MapCastilloN And Npclist(TempCharIndex).NPCtype = ReyCastillo Then Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Estoy al servicio del clan " & CastilloNorte & "°" & Npclist(TempCharIndex).Char.CharIndex & FONTTYPE_INFO)
            If Npclist(TempCharIndex).pos.Map = MapCastilloS And Npclist(TempCharIndex).NPCtype = ReyCastillo Then Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Estoy al servicio del clan " & CastilloSur & "°" & Npclist(TempCharIndex).Char.CharIndex & FONTTYPE_INFO)
            If Npclist(TempCharIndex).pos.Map = MapCastilloE And Npclist(TempCharIndex).NPCtype = ReyCastillo Then Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Estoy al servicio del clan " & CastilloEste & "°" & Npclist(TempCharIndex).Char.CharIndex & FONTTYPE_INFO)
            If Npclist(TempCharIndex).pos.Map = MapCastilloO And Npclist(TempCharIndex).NPCtype = ReyCastillo Then Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Estoy al servicio del clan " & CastilloOeste & "°" & Npclist(TempCharIndex).Char.CharIndex & FONTTYPE_INFO)
            
            If Len(Npclist(TempCharIndex).Desc) > 1 Then
 
                If Npclist(TempCharIndex).QuestNumber Then
                    If UserTieneQuest(userindex, Npclist(TempCharIndex).QuestNumber) Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).TalkDuringQuest & "°" & Npclist(TempCharIndex).Char.CharIndex & FONTTYPE_INFO)
                    ElseIf UserHizoQuest(userindex, Npclist(TempCharIndex).QuestNumber) Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).TalkAfterQuest & "°" & Npclist(TempCharIndex).Char.CharIndex & FONTTYPE_INFO)
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).Char.CharIndex & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & Npclist(TempCharIndex).Desc & "°" & Npclist(TempCharIndex).Char.CharIndex & FONTTYPE_INFO)
                End If
            ElseIf TempCharIndex = CentinelaNPCIndex Then
                'Enviamos nuevamente el texto del centinela según quien pregunta
                Call modCentinela.CentinelaSendClave(userindex)
            Else
                If Npclist(TempCharIndex).MaestroUser > 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "|| " & estatus & Npclist(TempCharIndex).name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).name & FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "|| " & estatus & Npclist(TempCharIndex).name & "." & FONTTYPE_INFO)
                End If
                
            End If
            FoundSomething = 1
            UserList(userindex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
            UserList(userindex).flags.TargetNPC = TempCharIndex
            UserList(userindex).flags.TargetUser = 0
            UserList(userindex).flags.TargetObj = 0
        
    End If
    
    If FoundChar = 0 Then
        UserList(userindex).flags.TargetNPC = 0
        UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(userindex).flags.TargetUser = 0
    End If
    
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(userindex).flags.TargetNPC = 0
        UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(userindex).flags.TargetUser = 0
        UserList(userindex).flags.TargetObj = 0
        UserList(userindex).flags.TargetObjMap = 0
        UserList(userindex).flags.TargetObjX = 0
        UserList(userindex).flags.TargetObjY = 0
        
    End If

    If FoundSomething = 0 Then
        UserList(userindex).flags.TargetNPC = 0
        UserList(userindex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(userindex).flags.TargetUser = 0
        UserList(userindex).flags.TargetObj = 0
        UserList(userindex).flags.TargetObjMap = 0
        UserList(userindex).flags.TargetObjX = 0
        UserList(userindex).flags.TargetObjY = 0
        
    End If


End Sub

Function FindDirection(pos As WorldPos, Target As WorldPos) As eHeading
'*****************************************************************
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************
Dim X As Integer
Dim Y As Integer

X = pos.X - Target.X
Y = pos.Y - Target.Y

'NE
If Sgn(X) = -1 And Sgn(Y) = 1 Then
    FindDirection = eHeading.NORTH
    Exit Function
End If

'NW
If Sgn(X) = 1 And Sgn(Y) = 1 Then
    FindDirection = eHeading.WEST
    Exit Function
End If

'SW
If Sgn(X) = 1 And Sgn(Y) = -1 Then
    FindDirection = eHeading.WEST
    Exit Function
End If

'SE
If Sgn(X) = -1 And Sgn(Y) = -1 Then
    FindDirection = eHeading.SOUTH
    Exit Function
End If

'Sur
If Sgn(X) = 0 And Sgn(Y) = -1 Then
    FindDirection = eHeading.SOUTH
    Exit Function
End If

'norte
If Sgn(X) = 0 And Sgn(Y) = 1 Then
    FindDirection = eHeading.NORTH
    Exit Function
End If

'oeste
If Sgn(X) = 1 And Sgn(Y) = 0 Then
    FindDirection = eHeading.WEST
    Exit Function
End If

'este
If Sgn(X) = -1 And Sgn(Y) = 0 Then
    FindDirection = eHeading.EAST
    Exit Function
End If

'misma
If Sgn(X) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function

'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal index As Integer) As Boolean

ItemNoEsDeMapa = ObjData(index).OBJType <> eOBJType.otPuertas And _
            ObjData(index).OBJType <> eOBJType.otForos And _
            ObjData(index).OBJType <> eOBJType.otCarteles And _
            ObjData(index).OBJType <> eOBJType.otArboles And _
            ObjData(index).OBJType <> eOBJType.otYacimiento And _
            ObjData(index).OBJType <> eOBJType.otTeleport
End Function
'[/Barrin 30-11-03]

Public Function MostrarCantidad(ByVal index As Integer) As Boolean
MostrarCantidad = ObjData(index).OBJType <> eOBJType.otPuertas And _
            ObjData(index).OBJType <> eOBJType.otForos And _
            ObjData(index).OBJType <> eOBJType.otCarteles And _
            ObjData(index).OBJType <> eOBJType.otArboles And _
            ObjData(index).OBJType <> eOBJType.otYacimiento And _
            ObjData(index).OBJType <> eOBJType.otTeleport
End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean

EsObjetoFijo = OBJType = eOBJType.otForos Or _
               OBJType = eOBJType.otCarteles Or _
               OBJType = eOBJType.otArboles Or _
               OBJType = eOBJType.otYacimiento

End Function
