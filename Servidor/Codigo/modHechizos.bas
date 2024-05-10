Attribute VB_Name = "modHechizos"
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

Public Const HELEMENTAL_FUEGO As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const SUPERANILLO As Integer = 700

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal userindex As Integer, ByVal Spell As Integer)

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(userindex).flags.Invisible = 1 Or UserList(userindex).flags.Oculto = 1 Then Exit Sub

Npclist(NpcIndex).CanAttack = 0
Dim daño As Integer

If Hechizos(Spell).SubeHP = 1 Then

    daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Hechizos(Spell).WAV)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)

    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP + daño
    If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
    
    Call SendData(SendTarget.toindex, userindex, 0, "||" & Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_BORDON)
    Call SendUserStatsBox(Val(userindex))

ElseIf Hechizos(Spell).SubeHP = 2 Then
    
    If UserList(userindex).flags.Privilegios <= PlayerType.VIP Then
    
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        
        If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(userindex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
        End If
        
        If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).DefensaMagicaMax)
        End If
        
        If UserList(userindex).Invent.HerramientaEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(userindex).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
        End If
        
        If daño < 0 Then daño = 0
        
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
    
        UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - daño
        
        Call SendData(SendTarget.toindex, userindex, 0, "||" & Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_BORDON)
        Call SendUserStatsBox(Val(userindex))
        
        'Muere
        If UserList(userindex).Stats.MinHP < 1 Then
            UserList(userindex).Stats.MinHP = 0
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                RestarCriminalidad (userindex)
            End If
            
            '[Barrin 1-12-03]
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call ContarMuerte(userindex, Npclist(NpcIndex).MaestroUser)
                Call ActStats(userindex, Npclist(NpcIndex).MaestroUser)
            End If
            '[/Barrin]
        End If
    
    End If
    
End If

If Hechizos(Spell).Paraliza = 1 Then
     If UserList(userindex).flags.Paralizado = 0 Then
          Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Hechizos(Spell).WAV)
          Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
          
            If UserList(userindex).Invent.HerramientaEqpObjIndex = SUPERANILLO Then
                Call SendData(SendTarget.toindex, userindex, 0, "|| Tu anillo rechaza los efectos del hechizo." & FONTTYPE_FIGHT)
                Exit Sub
            End If
          UserList(userindex).flags.Paralizado = 1
          UserList(userindex).Counters.Paralisis = IntervaloParalizado

#If SeguridadAlkon Then
        If EncriptarProtocolosCriticos Then
            Call SendCryptedData(SendTarget.toindex, userindex, 0, "PARADOK")
        Else
#End If
            Call SendData(SendTarget.toindex, userindex, 0, "PARADOK")
#If SeguridadAlkon Then
        End If
#End If
     End If
     
     
End If


'Muere el usuario
If UserList(userindex).Stats.MinHP <= 0 Then
 
    Call SendData(SendTarget.toindex, userindex, 0, "6") ' Le informamos que ha muerto ;)
    
    Call UserDie(userindex)
    End If

End Sub


Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
'solo hechizos ofensivos!

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
Npclist(NpcIndex).CanAttack = 0

Dim daño As Integer

If Hechizos(Spell).SubeHP = 2 Then
    
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, Npclist(TargetNPC).pos.Map, "TW" & Hechizos(Spell).WAV)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, Npclist(TargetNPC).pos.Map, "XFC" & Npclist(TargetNPC).Char.CharIndex & "," & Hechizos(Spell).FXgrh & "," & Hechizos(Spell).loops)
        
        Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - daño
        
        'Muere
        If Npclist(TargetNPC).Stats.MinHP < 1 Then
            Npclist(TargetNPC).Stats.MinHP = 0
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
            Else
                Call MuereNpc(TargetNPC, 0)
            End If
        End If
    
End If
    
End Sub



Function TieneHechizo(ByVal i As Integer, ByVal userindex As Integer) As Boolean

On Error GoTo errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(userindex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
errhandler:

End Function

Sub AgregarHechizo(ByVal userindex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer
hIndex = ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).HechizoIndex

If Not TieneHechizo(hIndex, userindex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(userindex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(userindex).Stats.UserHechizos(j) <> 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No tenes espacio para mas hechizos." & FONTTYPE_INFO)
    Else
        UserList(userindex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, userindex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(userindex, CByte(Slot), 1)
    End If
Else
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya tenes ese hechizo." & FONTTYPE_INFO)
End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal s As String, ByVal userindex As Integer)
On Error Resume Next

    Dim ind As String
    ind = UserList(userindex).Char.CharIndex
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbCyan & "°" & s & "°" & ind)
    Exit Sub
End Sub

Function PuedeLanzar(ByVal userindex As Integer, ByVal HechizoIndex As Integer) As Boolean

If Len(Hechizos(HechizoIndex).ExclusivoClase) > 0 And Hechizos(HechizoIndex).ExclusivoClase <> UCase$(UserList(userindex).Clase) Then
Call SendData(SendTarget.toindex, userindex, 0, "||Tú clase no puede lanzar este hechizo." & FONTTYPE_INFO)
PuedeLanzar = False
Exit Function
End If

If UserList(userindex).flags.Muerto = 0 Then
    Dim wp2 As WorldPos
    wp2.Map = UserList(userindex).flags.TargetMap
    wp2.X = UserList(userindex).flags.TargetX
    wp2.Y = UserList(userindex).flags.TargetY
    
    If Hechizos(HechizoIndex).NeedStaff > 0 Then
        If UCase$(UserList(userindex).Clase) = "MAGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||Tu Báculo no es lo suficientemente poderoso para que puedas lanzar el conjuro." & FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes lanzar este conjuro sin la ayuda de un báculo." & FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
            End If
        End If
    End If
        
    If UserList(userindex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
        If UserList(userindex).Stats.UserSkills(eSkill.Magia) >= Hechizos(HechizoIndex).MinSkill Then
            If UserList(userindex).Stats.MinSta >= Hechizos(HechizoIndex).StaRequerido Then
                PuedeLanzar = True
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||Estás muy cansado para lanzar este hechizo." & FONTTYPE_INFO)
                PuedeLanzar = False
            End If
                
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||No tenes suficientes puntos de magia para lanzar este hechizo." & FONTTYPE_INFO)
            PuedeLanzar = False
        End If
    Else
            Call SendData(SendTarget.toindex, userindex, 0, "||No tenes suficiente mana." & FONTTYPE_INFO)
            PuedeLanzar = False
    End If
Else
   Call SendData(SendTarget.toindex, userindex, 0, "||No podes lanzar hechizos porque estas muerto." & FONTTYPE_INFO)
   PuedeLanzar = False
End If

End Function

Sub HechizoTerrenoEstado(ByVal userindex As Integer, ByRef b As Boolean)
Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim h As Integer
Dim TempX As Integer
Dim TempY As Integer


    PosCasteadaX = UserList(userindex).flags.TargetX
    PosCasteadaY = UserList(userindex).flags.TargetY
    PosCasteadaM = UserList(userindex).flags.TargetMap
    
    h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    
    If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
        b = True
        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).userindex > 0 Then
                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).userindex).flags.Invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).userindex).flags.AdminInvisible = 0 Then
                            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(MapData(PosCasteadaM, TempX, TempY).userindex).Char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)
                        End If
                    End If
                End If
            Next TempY
        Next TempX
    
        Call InfoHechizo(userindex)
    End If

End Sub

Sub HechizoInvocacion(ByVal userindex As Integer, ByRef b As Boolean)

If UserList(userindex).NroMacotas >= MAXMASCOTAS Then Exit Sub

'No permitimos se invoquen criaturas en zonas seguras
If MapInfo(UserList(userindex).pos.Map).Pk = False Or MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = eTrigger.ZONASEGURA Then
    Call SendData(SendTarget.toindex, userindex, 0, "||En zona segura no puedes invocar criaturas." & FONTTYPE_INFO)
    Exit Sub
End If

Dim h As Integer, j As Integer, ind As Integer, index As Integer
Dim TargetPos As WorldPos


TargetPos.Map = UserList(userindex).flags.TargetMap
TargetPos.X = UserList(userindex).flags.TargetX
TargetPos.Y = UserList(userindex).flags.TargetY

h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    If Hechizos(h).NumNpc = 94 Then
    If UserList(userindex).flags.EleDeTierra = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Ya has invocado un elemental de tierra." & FONTTYPE_INFO)
    Exit Sub
    End If
    End If
    If Hechizos(h).NumNpc = 92 Then
    If UserList(userindex).flags.EleDeAgua = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Ya has invocado un elemental de agua." & FONTTYPE_INFO)
    Exit Sub
    End If
    End If
    If Hechizos(h).NumNpc = 93 Then
    If UserList(userindex).flags.EleDeFuego = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Ya has invocado un elemental de fuego." & FONTTYPE_INFO)
    Exit Sub
    End If
    End If
For j = 1 To Hechizos(h).Cant
    
    If UserList(userindex).NroMacotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(h).NumNpc, TargetPos, True, False)
        If Hechizos(h).NumNpc = 92 Then
            UserList(userindex).flags.EleDeAgua = 1
        End If
        If Hechizos(h).NumNpc = 93 Then
            UserList(userindex).flags.EleDeFuego = 1
        End If
        If Hechizos(h).NumNpc = 94 Then
            UserList(userindex).flags.EleDeTierra = 1
        End If
        If ind > 0 Then
            UserList(userindex).NroMacotas = UserList(userindex).NroMacotas + 1
            
            index = FreeMascotaIndex(userindex)
            
            UserList(userindex).MascotasIndex(index) = ind
            UserList(userindex).MascotasType(index) = Npclist(ind).Numero
            
            Npclist(ind).MaestroUser = userindex
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
            
            Call FollowAmo(ind)
        End If
            
    Else
        Exit For
    End If
    
Next j


Call InfoHechizo(userindex)
b = True


End Sub

Sub HandleHechizoTerreno(ByVal userindex As Integer, ByVal uh As Integer)

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uInvocacion '
        Call HechizoInvocacion(userindex, b)
    Case TipoHechizo.uEstado
        Call HechizoTerrenoEstado(userindex, b)
    
End Select

If b Then
    Call SubirSkill(userindex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(userindex).Stats.MinSta < 0 Then UserList(userindex).Stats.MinSta = 0
    Call SendUserStatsBox(userindex)
End If


End Sub

Sub HandleHechizoUsuario(ByVal userindex As Integer, ByVal uh As Integer)

Dim b As Boolean
Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoUsuario(userindex, b)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
    'If UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP Then
    'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes curar teniendo la vida llena." & FONTTYPE_INFO)
    'Else
       Call HechizoPropUsuario(userindex, b)
       'End If
End Select

If b Then
    Call SubirSkill(userindex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(userindex).Stats.MinSta < 0 Then UserList(userindex).Stats.MinSta = 0
    Call SendUserStatsBox(userindex)
    Call SendUserStatsBox(UserList(userindex).flags.TargetUser)
    UserList(userindex).flags.TargetUser = 0
End If

End Sub

Sub HandleHechizoNPC(ByVal userindex As Integer, ByVal uh As Integer)

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
        Call HechizoEstadoNPC(UserList(userindex).flags.TargetNPC, uh, b, userindex)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
        Call HechizoPropNPC(uh, UserList(userindex).flags.TargetNPC, userindex, b)
End Select

If b Then
    Call SubirSkill(userindex, Magia)
    UserList(userindex).flags.TargetNPC = 0
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(userindex).Stats.MinMAN < 0 Then UserList(userindex).Stats.MinMAN = 0
    UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(userindex).Stats.MinSta < 0 Then UserList(userindex).Stats.MinSta = 0
    Call SendUserStatsBox(userindex)
End If

End Sub


Sub LanzarHechizo(index As Integer, userindex As Integer)

Dim uh As Integer
Dim exito As Boolean

uh = UserList(userindex).Stats.UserHechizos(index)

If PuedeLanzar(userindex, uh) Then
    Select Case Hechizos(uh).Target
        
        Case TargetType.uUsuarios
            If UserList(userindex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(userindex).flags.TargetUser).pos.Y - UserList(userindex).pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(userindex, uh)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||Este hechizo actua solo sobre usuarios." & FONTTYPE_INFO)
            End If
        Case TargetType.uNPC
            If UserList(userindex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(userindex).flags.TargetNPC).pos.Y - UserList(userindex).pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(userindex, uh)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||Este hechizo solo afecta a los npcs." & FONTTYPE_INFO)
            End If
        Case TargetType.uUsuariosYnpc
            If UserList(userindex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(userindex).flags.TargetUser).pos.Y - UserList(userindex).pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(userindex, uh)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            ElseIf UserList(userindex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(userindex).flags.TargetNPC).pos.Y - UserList(userindex).pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(userindex, uh)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos para lanzar este hechizo." & FONTTYPE_WARNING)
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||Target invalido." & FONTTYPE_INFO)
            End If
        Case TargetType.uTerreno
            Call HandleHechizoTerreno(userindex, uh)
        Case TargetType.uOnlyUsuario
        'Verifica si el objetivo es el user.
        If UserList(userindex).flags.TargetUser = userindex Then
        Call HandleHechizoUsuario(userindex, uh)
        Else
        'Si no es tira mensaje de error
        Call SendData(SendTarget.toindex, userindex, 0, "||Este hechizo solamente puede ser lanzado sobre ti mismo." & FONTTYPE_INFO)
        End If
    End Select
    
End If

If UserList(userindex).Counters.Trabajando Then _
    UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando - 1

If UserList(userindex).Counters.Ocultando Then _
    UserList(userindex).Counters.Ocultando = UserList(userindex).Counters.Ocultando - 1
    
    Call SendData(toindex, userindex, UserList(userindex).pos.Map, "PZ" & UserList(userindex).Stats.UserAtributos(Fuerza))
    Call SendData(toindex, userindex, UserList(userindex).pos.Map, "PX" & UserList(userindex).Stats.UserAtributos(Agilidad))
    
End Sub

Sub HechizoEstadoUsuario(ByVal userindex As Integer, ByRef b As Boolean)



Dim h As Integer, TU As Integer
h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
TU = UserList(userindex).flags.TargetUser


If Hechizos(h).Invisibilidad = 1 Then
   
    If UserList(TU).flags.Muerto = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||¡Está muerto!" & FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    'Si sos Ciudadano y el es Criminal no podes
If Criminal(TU) And Ciudadano(userindex) And TriggerZonaPelea(TU, userindex) <> TRIGGER6_PERMITE Then
Call SendData(SendTarget.toindex, userindex, 0, "||Los ciudadanos no pueden ayudar a criminales, escribe /RENUNCIAR." & FONTTYPE_FIGHT)
Exit Sub
End If
'Mithrandir - Sistema de Status

    
    UserList(TU).flags.Invisible = 1
#If SeguridadAlkon Then
    If EncriptarProtocolosCriticos Then
        Call SendCryptedData(SendTarget.ToMap, 0, UserList(TU).pos.Map, "NOVER" & UserList(TU).Char.CharIndex & ",1")
    Else
#End If
        Call SendData(SendTarget.ToMap, 0, UserList(TU).pos.Map, "NOVER" & UserList(TU).Char.CharIndex & ",1")
#If SeguridadAlkon Then
    End If
#End If
    Call InfoHechizo(userindex)
    b = True
End If

If Hechizos(h).Mimetiza = 1 Then
    If UserList(TU).flags.Muerto = 1 Then
        Exit Sub
    End If
    
    If UserList(TU).flags.Navegando = 1 Then
        Exit Sub
    End If
    If UserList(userindex).flags.Navegando = 1 Then
        Exit Sub
    End If
    
    If UserList(userindex).flags.Mimetizado = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Ya te encuentras transformado. El hechizo no ha tenido efecto" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    'copio el char original al mimetizado
    
    With UserList(userindex)
        .CharMimetizado.Body = .Char.Body
        .CharMimetizado.Head = .Char.Head
        .CharMimetizado.CascoAnim = .Char.CascoAnim
        .CharMimetizado.ShieldAnim = .Char.ShieldAnim
        .CharMimetizado.WeaponAnim = .Char.WeaponAnim
        
        .flags.Mimetizado = 1
        
        'ahora pongo local el del enemigo
        .Char.Body = UserList(TU).Char.Body
        .Char.Head = UserList(TU).Char.Head
        .Char.CascoAnim = UserList(TU).Char.CascoAnim
        .Char.ShieldAnim = UserList(TU).Char.ShieldAnim
        .Char.WeaponAnim = UserList(TU).Char.WeaponAnim
    
        Call ChangeUserChar(SendTarget.ToMap, 0, .pos.Map, userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    End With
   
   Call InfoHechizo(userindex)
   b = True
End If


If Hechizos(h).Envenena = 1 Then
        If Not PuedeAtacar(userindex, TU) Then Exit Sub
        If userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(userindex, TU)
        End If
        UserList(TU).flags.Envenenado = 1
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(h).CuraVeneno = 1 Then
        UserList(TU).flags.Envenenado = 0
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(h).Maldicion = 1 Then
        If Not PuedeAtacar(userindex, TU) Then Exit Sub
        If userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(userindex, TU)
        End If
        UserList(TU).flags.Maldicion = 1
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(h).RemoverMaldicion = 1 Then
        UserList(TU).flags.Maldicion = 0
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(h).Bendicion = 1 Then
        UserList(TU).flags.Bendicion = 1
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(h).Paraliza = 1 Or Hechizos(h).Inmoviliza = 1 Then
If userindex = TU Then
 Call SendData(SendTarget.toindex, userindex, 0, "||Target invalido." & FONTTYPE_INFO)
Exit Sub
End If
Dim klan As String
If UserList(userindex).flags.SeguroClan = True Then
If UserList(TU).GuildIndex > 0 Then
If UserList(userindex).GuildIndex > 0 Then
    klan = Guilds(UserList(userindex).GuildIndex).GuildName
   
      If Guilds(UserList(TU).GuildIndex).GuildName = Guilds(UserList(userindex).GuildIndex).GuildName And klan <> "" Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes atacar a tu propio Clan con el seguro activado, escribe /SEGCLAN para desactivarlo." & FONTTYPE_FIGHT)
        Exit Sub
      End If
    End If
End If
End If
     If UserList(TU).flags.Paralizado = 0 Then
            If Not PuedeAtacar(userindex, TU) Then Exit Sub
            
            If userindex <> TU Then
                Call UsuarioAtacadoPorUsuario(userindex, TU)
            End If
            
            Call InfoHechizo(userindex)
            b = True
            If UserList(TU).Invent.HerramientaEqpObjIndex = SUPERANILLO Then
                Call SendData(SendTarget.toindex, TU, 0, "|| Tu anillo rechaza los efectos del hechizo." & FONTTYPE_FIGHT)
                Call SendData(SendTarget.toindex, userindex, 0, "|| ¡El hechizo no tiene efecto!" & FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            UserList(TU).flags.Paralizado = 1
            UserList(TU).Counters.Paralisis = IntervaloParalizado
                Call SendData(SendTarget.toindex, TU, 0, "PARADOK")
                Call SendData(SendTarget.toindex, TU, 0, "PU" & UserList(TU).pos.X & "," & UserList(TU).pos.Y)

    End If
End If

If Hechizos(h).RemoverParalisis = 1 Then
If UserList(TU).flags.Paralizado = 1 Then
If Criminal(TU) And Ciudadano(userindex) And TriggerZonaPelea(TU, userindex) <> TRIGGER6_PERMITE Then
Call SendData(SendTarget.toindex, userindex, 0, "||Los ciudadanos no pueden ayudar a criminales, escribe /RENUNCIAR.." & FONTTYPE_INFO)
Exit Sub
End If
 
UserList(TU).flags.Paralizado = 0
'no need to crypt this
Call SendData(SendTarget.toindex, TU, 0, "PARADOK")
Call InfoHechizo(userindex)
b = True
End If
End If

If Hechizos(h).RemoverEstupidez = 1 Then
    If Not UserList(TU).flags.Estupidez = 0 Then
                UserList(TU).flags.Estupidez = 0
                'no need to crypt this
                Call SendData(SendTarget.toindex, TU, 0, "NESTUP")
                Call InfoHechizo(userindex)
                b = True
    End If
End If


If Hechizos(h).Revivir = 1 Then

                If UserList(userindex).pos.Map = 81 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes resucitar a nadie aquí." & FONTTYPE_INFO)
            Exit Sub
        End If
        
                        If UserList(userindex).pos.Map = 8 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes resucitar a nadie aquí." & FONTTYPE_INFO)
            Exit Sub
        End If
        
                        If UserList(userindex).pos.Map = 54 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes resucitar a nadie aquí." & FONTTYPE_INFO)
            Exit Sub
        End If
        
                                If UserList(userindex).pos.Map = 31 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes resucitar a nadie aquí." & FONTTYPE_INFO)
            Exit Sub
        End If
        
                                If UserList(userindex).pos.Map = 32 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes resucitar a nadie aquí." & FONTTYPE_INFO)
            Exit Sub
        End If
        
                                If UserList(userindex).pos.Map = 33 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes resucitar a nadie aquí." & FONTTYPE_INFO)
            Exit Sub
        End If
        
                                If UserList(userindex).pos.Map = 34 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes resucitar a nadie aquí." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(TU).flags.SeguroResu = True And UserList(TU).flags.Muerto = 1 Then
            Call SendData(toindex, UserList(userindex).flags.TargetUser, 0, "||" & UserList(userindex).name & " esta intentando resucitarte, si quieres que te resucite desactiva el seguro de resurreccion con la tecla X o escribiendo /SEGR" & FONTTYPE_INFO)
            Call SendData(SendTarget.toindex, userindex, 0, "||El espiritu no tiene intensiones de revivir." & FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
           If UserList(TU).flags.TimeRevivir > 0 Then
    SendData SendTarget.toindex, userindex, 0, "||Debes esperar " & UserList(TU).flags.TimeRevivir & " segundos para poder revivir al usuario." & FONTTYPE_INFO
    Exit Sub
    End If
    If UserList(TU).flags.Muerto = 1 Then
'Revivir?
If Criminal(TU) And Ciudadano(userindex) And TriggerZonaPelea(TU, userindex) <> TRIGGER6_PERMITE Then
Call SendData(SendTarget.toindex, userindex, 0, "||Los ciudadanos no pueden ayudar a criminales, escribe /RENUNCIAR." & FONTTYPE_FIGHT)
Exit Sub
End If
'Revivir?

        'revisamos si necesita vara
        If UCase$(UserList(userindex).Clase) = "MAGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(h).NeedStaff Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||Necesitas un mejor báculo para este hechizo" & FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
        End If

        '/Juan Maraxus
        If Not Criminal(TU) Then
            If TU <> userindex Then
                UserList(userindex).Reputacion.NobleRep = UserList(userindex).Reputacion.NobleRep + 500
                If UserList(userindex).Reputacion.NobleRep > MAXREP Then _
                    UserList(userindex).Reputacion.NobleRep = MAXREP
            End If
        End If
        UserList(TU).Stats.MinMAN = 0
        '/Pablo Toxic Waste
        
        b = True
        Call InfoHechizo(userindex)
        Call RevivirUsuario(TU)
    Else
        b = False
    End If

End If

If Hechizos(h).Ceguera = 1 Then
        If Not PuedeAtacar(userindex, TU) Then Exit Sub
        If userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(userindex, TU)
        End If
        UserList(TU).flags.Ceguera = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado / 3
#If SeguridadAlkon Then
        Call SendCryptedData(SendTarget.toindex, TU, 0, "CEGU")
#Else
        Call SendData(SendTarget.toindex, TU, 0, "CEGU")
#End If
        Call InfoHechizo(userindex)
        b = True
End If

If Hechizos(h).Estupidez = 1 Then
        If Not PuedeAtacar(userindex, TU) Then Exit Sub
        If userindex <> TU Then
            Call UsuarioAtacadoPorUsuario(userindex, TU)
        End If
        UserList(TU).flags.Estupidez = 1
        UserList(TU).Counters.Ceguera = IntervaloParalizado
#If SeguridadAlkon Then
        If EncriptarProtocolosCriticos Then
            Call SendCryptedData(SendTarget.toindex, TU, 0, "DUMB")
        Else
#End If
            Call SendData(SendTarget.toindex, TU, 0, "DUMB")
#If SeguridadAlkon Then
        End If
#End If
        Call InfoHechizo(userindex)
        b = True
End If

End Sub
Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal userindex As Integer)



If Hechizos(hIndex).Invisibilidad = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Invisible = 1
   b = True
End If

If Hechizos(hIndex).Envenena = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
        Exit Sub
   End If
   
   If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes atacar guardias reales, escribe /RENUNCIAR." & FONTTYPE_WARNING)
            Exit Sub
    End If
        
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Envenenado = 1
   b = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Envenenado = 0
   b = True
End If

If Hechizos(hIndex).Maldicion = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
        Exit Sub
   End If
   
   If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes atacar guardias reales, escribe /RENUNCIAR." & FONTTYPE_WARNING)
            Exit Sub
    End If
    
    Call InfoHechizo(userindex)
    Npclist(NpcIndex).flags.Maldicion = 1
    b = True
End If

If Hechizos(hIndex).RemoverMaldicion = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Maldicion = 0
   b = True
End If

If Hechizos(hIndex).Bendicion = 1 Then
   Call InfoHechizo(userindex)
   Npclist(NpcIndex).flags.Bendicion = 1
   b = True
End If

If Hechizos(hIndex).Paraliza = 1 Then
If Npclist(NpcIndex).Attackable = 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
Exit Sub
End If
If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
If UserList(userindex).flags.Seguro Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes atacar guardas, escribe /RENUNCIAR." & FONTTYPE_WARNING)
Exit Sub
Else
UserList(userindex).Reputacion.NobleRep = 0
UserList(userindex).Reputacion.PlebeRep = 0
UserList(userindex).Reputacion.AsesinoRep = UserList(userindex).Reputacion.AsesinoRep + 500
If UserList(userindex).Reputacion.AsesinoRep > MAXREP Then _
UserList(userindex).Reputacion.AsesinoRep = MAXREP
End If
 
'Si es Ciudadano lo hacemos Neutro
If Ciudadano(userindex) Then Exit Sub
'Mithrandir
End If
 
Call InfoHechizo(userindex)
Npclist(NpcIndex).flags.Paralizado = 1
Npclist(NpcIndex).flags.Inmovilizado = 0
Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
b = True
Else
Call SendData(SendTarget.toindex, userindex, 0, "||El npc es inmune a este hechizo." & FONTTYPE_FIGHT)
End If
End If

'[Barrin 16-2-04]
If Hechizos(hIndex).RemoverParalisis = 1 Then
   If Npclist(NpcIndex).flags.Paralizado = 1 And Npclist(NpcIndex).MaestroUser = userindex Then
            Call InfoHechizo(userindex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
   Else
      Call SendData(SendTarget.toindex, userindex, 0, "||Este hechizo solo afecta NPCs que tengan amo." & FONTTYPE_WARNING)
   End If
End If
'[/Barrin]
 
If Hechizos(hIndex).Inmoviliza = 1 Then
If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
If UserList(userindex).flags.Seguro Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes atacar guardias, escribe /RENUNCIAR." & FONTTYPE_WARNING)
Exit Sub
Else
UserList(userindex).Reputacion.NobleRep = 0
UserList(userindex).Reputacion.PlebeRep = 0
UserList(userindex).Reputacion.AsesinoRep = UserList(userindex).Reputacion.AsesinoRep + 500
If UserList(userindex).Reputacion.AsesinoRep > MAXREP Then _
UserList(userindex).Reputacion.AsesinoRep = MAXREP
 
'Si es Ciudadano lo hacemos Neutro
If Ciudadano(userindex) Then Exit Sub
'Mithrandir
 
End If
End If
 
Npclist(NpcIndex).flags.Inmovilizado = 1
Npclist(NpcIndex).flags.Paralizado = 0
Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
Call InfoHechizo(userindex)
b = True
Else
Call SendData(SendTarget.toindex, userindex, 0, "||El npc es inmune a este hechizo." & FONTTYPE_FIGHT)
End If
End If

End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal userindex As Integer, ByRef b As Boolean)

If Npclist(NpcIndex).pos.Map = 106 And Npclist(NpcIndex).Numero = 937 And GuardiasRey <= 3 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No puedes atacar al rey mientras sus guerreros sigan con vida!." & FONTTYPE_ORO)
Exit Sub
End If

Dim daño As Long

If Hechizos(hIndex).SubeHP > 1 Then
If Npclist(NpcIndex).NPCtype = ReyCastillo Then
        If (Npclist(NpcIndex).pos.Map = MapCastilloN Or Npclist(NpcIndex).pos.Map = MapCastilloS Or Npclist(NpcIndex).pos.Map = MapCastilloE Or Npclist(NpcIndex).pos.Map = MapCastilloO) Then
            Dim castiact As String
            If Npclist(NpcIndex).pos.Map = MapCastilloN Then castiact = CastilloNorte
            If Npclist(NpcIndex).pos.Map = MapCastilloS Then castiact = CastilloSur
            If Npclist(NpcIndex).pos.Map = MapCastilloE Then castiact = CastilloEste
            If Npclist(NpcIndex).pos.Map = MapCastilloO Then castiact = CastilloOeste
 
If Not UserList(userindex).GuildIndex <> 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes atacar al rey del castillo por que no perteneses a ningun clan!!" & FONTTYPE_FIGHT)
             Exit Sub
            End If
            If Guilds(UserList(userindex).GuildIndex).GuildName = castiact Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes atacar al rey de tu castillo " & FONTTYPE_FIGHT)
                Exit Sub
            End If
        End If
    End If
End If


'Salud
If Hechizos(hIndex).SubeHP = 1 Then
    daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(userindex).Stats.ELV)
    
    Call InfoHechizo(userindex)
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP + daño
    If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then _
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
    Call SendData(SendTarget.toindex, userindex, 0, "||Has curado " & daño & " puntos de salud a la criatura." & FONTTYPE_VERDEN)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbGreen & "°+" & daño & "°" & str(Npclist(NpcIndex).Char.CharIndex))
    b = True
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    
    If Npclist(NpcIndex).Attackable = 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No podes atacar a ese npc." & FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    If Npclist(NpcIndex).NPCtype = 2 And UserList(userindex).flags.Seguro Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Debes sacarte el seguro para atacar guardias del imperio." & FONTTYPE_FIGHT)
        b = False
        Exit Sub
    End If
    
    If Not PuedeAtacarNPC(userindex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    
    daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(userindex).Stats.ELV)
    
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbYellow & "°-" & daño & "°" & str(Npclist(NpcIndex).Char.CharIndex))


        If UCase$(UserList(userindex).Clase) = "BARDO" Then
    If UserList(userindex).Invent.WeaponEqpObjIndex = 893 = False Then
        daño = daño + 0
        Else
        daño = daño + (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus)
    End If
    End If

    
        If UCase$(UserList(userindex).Clase) = "DRUIDA" Then
    If UserList(userindex).Invent.WeaponEqpObjIndex = 947 = False Then
        daño = daño + 0
        Else
        daño = daño + (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus)
    End If
    End If
    
            If UCase$(UserList(userindex).Clase) = "DRUIDA" Then
    If UserList(userindex).Invent.WeaponEqpObjIndex = 946 = False Then
        daño = daño + 0
        Else
        daño = daño + (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus)
    End If
    End If
    
            If UCase$(UserList(userindex).Clase) = "DRUIDA" Then
    If UserList(userindex).Invent.WeaponEqpObjIndex = 658 = False Then
        daño = daño + 0
        Else
        daño = daño + (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus)
    End If
    End If

    If Hechizos(hIndex).StaffAffected Then
        If UCase$(UserList(userindex).Clase) = "MAGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                daño = (daño * (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                'Aumenta daño segun el staff-
                'Daño = (Daño* (80 + BonifBáculo)) / 100
            Else
                daño = daño * 0.7 'Baja daño a 80% del original
            End If
        End If
    End If

    Call InfoHechizo(userindex)
    b = True
    Call NpcAtacado(NpcIndex, userindex)
    If Npclist(NpcIndex).flags.Snd2 > 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Npclist(NpcIndex).flags.Snd2)
    
    If UserList(userindex).flags.Privilegios = PlayerType.Admin Then
Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - Npclist(NpcIndex).Stats.MaxHP
Else
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - daño
    End If
    SendData SendTarget.toindex, userindex, 0, "||Le has causado " & daño & " puntos de daño a la criatura!" & FONTTYPE_BORDON
    Call CheckPets(NpcIndex, userindex, False)
    Call CalcularDarExp(userindex, NpcIndex, daño)

    If Npclist(NpcIndex).Stats.MinHP < 1 Then
        Npclist(NpcIndex).Stats.MinHP = 0
        Call MuereNpc(NpcIndex, userindex)
    End If
End If

End Sub

Sub InfoHechizo(ByVal userindex As Integer)


    Dim h As Integer
    h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
    
    
    Call DecirPalabrasMagicas(Hechizos(h).PalabrasMagicas, userindex)
    
    If UserList(userindex).flags.TargetUser > 0 Then
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(UserList(userindex).flags.TargetUser).Char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)
        Call SendData(SendTarget.ToPCArea, UserList(userindex).flags.TargetUser, UserList(userindex).pos.Map, "TW" & Hechizos(h).WAV)
    ElseIf UserList(userindex).flags.TargetNPC > 0 Then
        Call SendData(SendTarget.ToNPCArea, UserList(userindex).flags.TargetNPC, Npclist(UserList(userindex).flags.TargetNPC).pos.Map, "XFC" & Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex & "," & Hechizos(h).FXgrh & "," & Hechizos(h).loops)
        Call SendData(SendTarget.ToNPCArea, UserList(userindex).flags.TargetNPC, UserList(userindex).pos.Map, "TW" & Hechizos(h).WAV)
    End If
    
    If UserList(userindex).flags.TargetUser > 0 Then
        If userindex <> UserList(userindex).flags.TargetUser Then
            Call SendData(SendTarget.toindex, userindex, 0, "||" & Hechizos(h).HechizeroMsg & " " & UserList(UserList(userindex).flags.TargetUser).name & FONTTYPE_FIGHT)
            Call SendData(SendTarget.toindex, UserList(userindex).flags.TargetUser, 0, "||" & UserList(userindex).name & " " & Hechizos(h).TargetMsg & FONTTYPE_FIGHT)
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||" & Hechizos(h).PropioMsg & FONTTYPE_FIGHT)
        End If
    ElseIf UserList(userindex).flags.TargetNPC > 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||" & Hechizos(h).HechizeroMsg & " " & "la criatura." & FONTTYPE_FIGHT)
    End If

End Sub

Sub HechizoPropUsuario(ByVal userindex As Integer, ByRef b As Boolean)

Dim h As Integer
Dim daño As Integer
Dim tempChr As Integer
    
    
h = UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)
tempChr = UserList(userindex).flags.TargetUser
      
'If UserList(UserIndex).Name = "EL OSO" Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| le tiro el hechizo " & H & " a " & UserList(tempChr).Name & FONTTYPE_VENENO)
'End If

If Hechizos(h).ActivaVIP = 1 Then
If UserList(userindex).Stats.TransformadoVIP = 1 Then
UserList(userindex).Stats.TransformadoVIP = 0
UserList(userindex).Char.CascoAnim = 0
Call SendData(SendTarget.toindex, userindex, 0, "||VIP desactivado." & FONTTYPE_TALK)
UserList(userindex).flags.Privilegios = 0
Call SendData(SendTarget.toindex, userindex, 0, "VIP")
Else
UserList(userindex).Stats.TransformadoVIP = 1
UserList(userindex).Char.CascoAnim = 32
Call SendData(SendTarget.toindex, userindex, 0, "||VIP activado." & FONTTYPE_TALK)
UserList(userindex).flags.Privilegios = 1
Call SendData(SendTarget.toindex, userindex, 0, "VIP")
End If
Call WarpUserChar(userindex, UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y, False)
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARVIPW & ",")
    Exit Sub
End If

If Hechizos(h).Lenteja = 1 Then
Call SendData(SendTarget.toindex, tempChr, 0, "KKW")
UserList(tempChr).Counters.LentejaTiempo = 3
    Exit Sub
End If

' <-------- Agilidad ---------->
If Hechizos(h).SubeAgilidad = 1 Then
    If Criminal(tempChr) And Ciudadano(userindex) And TriggerZonaPelea(tempChr, userindex) <> TRIGGER6_PERMITE Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Los ciudadanos no pueden ayudar a criminales, escribe /RENUNCIAR." & FONTTYPE_FIGHT)
            Exit Sub
        Else
            Call DisNobAuBan(userindex, UserList(userindex).Reputacion.NobleRep * 0.5, 10000)
        End If
    
    Call InfoHechizo(userindex)
    daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = 7000
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).flags.TomoPocion = True
    daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = 7000
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
    b = True
    
End If

' <-------- Fuerza ---------->
If Hechizos(h).SubeFuerza = 1 Then
    If Criminal(tempChr) And Ciudadano(userindex) And TriggerZonaPelea(tempChr, userindex) <> TRIGGER6_PERMITE Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Los ciudadanos no pueden ayudar a criminales, escribe /RENUNCIAR." & FONTTYPE_FIGHT)
            Exit Sub
        Else
            Call DisNobAuBan(userindex, UserList(userindex).Reputacion.NobleRep * 0.5, 10000)
        End If
    
    Call InfoHechizo(userindex)
    daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
    UserList(tempChr).flags.DuracionEfecto = 7000

    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) * 2)
    
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(h).SubeFuerza = 2 Then

    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).flags.TomoPocion = True
    
    daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = 7000
    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
    b = True
    
End If

'Salud
If Hechizos(h).SubeHP = 1 Then
    
    If Criminal(tempChr) And Ciudadano(userindex) And TriggerZonaPelea(tempChr, userindex) <> TRIGGER6_PERMITE Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Los ciudadanos no pueden ayudar a criminales, escribe /RENUNCIAR." & FONTTYPE_FIGHT)
            Exit Sub
        Else
            Call DisNobAuBan(userindex, UserList(userindex).Reputacion.NobleRep * 0.5, 10000)
        End If
    
    
    daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(userindex).Stats.ELV)
    
    'Call InfoHechizo(UserIndex)

            If userindex <> tempChr Then
        Call SendUserStatsBox(tempChr)
        With UserList(tempChr).Stats
            If .MinHP >= .MaxHP Then
                .MinHP = .MaxHP
                Call SendData(SendTarget.toindex, userindex, 0, "||El usuario que intentas curar tiene la vida al máximo." & FONTTYPE_VERDEN)
                Exit Sub
            Else
                Call InfoHechizo(userindex)
                Call SendData(SendTarget.toindex, userindex, 0, "||Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).name & FONTTYPE_VERDEN)
                Call SendData(SendTarget.toindex, tempChr, 0, "||" & UserList(userindex).name & " te ha restaurado " & daño & " puntos de vida." & FONTTYPE_VERDEN)
                UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP + daño
            End If
        End With
    Else
    Call SendUserStatsBox(userindex)
        With UserList(userindex).Stats
        
            If .MinHP >= .MaxHP Then
                .MinHP = .MaxHP
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes curarte tienes la vida al máximo." & FONTTYPE_VERDEN)
                Exit Sub
            Else
                Call InfoHechizo(userindex)
                Call SendData(SendTarget.toindex, userindex, 0, "||Te has restaurado " & daño & " puntos de vida." & FONTTYPE_VERDEN)
                UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP + daño
            End If
        End With
    End If

    If UserList(tempChr).Stats.MinHP > UserList(tempChr).Stats.MaxHP Then _
        UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP

    
    b = True
ElseIf Hechizos(h).SubeHP = 2 Then
    
    Dim klan As String
If UserList(userindex).flags.SeguroClan = True Then
If UserList(tempChr).GuildIndex > 0 Then
If UserList(userindex).GuildIndex > 0 Then
    klan = Guilds(UserList(userindex).GuildIndex).GuildName
   
      If Guilds(UserList(tempChr).GuildIndex).GuildName = Guilds(UserList(userindex).GuildIndex).GuildName And klan <> "" Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes atacar a tu propio Clan con el seguro activado, escribe /SEGCLAN para desactivarlo." & FONTTYPE_FIGHT)
        Exit Sub
      End If
    End If
End If
End If
    
    If userindex = tempChr Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No podes atacarte a vos mismo." & FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    

    
    daño = RandomNumber(Hechizos(h).MinHP, Hechizos(h).MaxHP)
    
'If UserList(UserIndex).Name = "EL OSO" Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| danio, minhp, maxhp " & daño & " " & Hechizos(H).MinHP & " " & Hechizos(H).MaxHP & FONTTYPE_VENENO)
'End If
    
    If UserList(userindex).Stats.ELV < 51 Then
daño = daño + Porcentaje(daño, 3 * UserList(userindex).Stats.ELV)
Else
daño = daño + Porcentaje(daño, 3 * 50)
End If
    
'If UserList(UserIndex).Name = "EL OSO" Then
'    Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| daño, ELV " & daño & " " & UserList(UserIndex).Stats.ELV & FONTTYPE_VENENO)
'End If
    
    

        If UCase$(UserList(userindex).Clase) = "BARDO" Then
    If UserList(userindex).Invent.WeaponEqpObjIndex = 893 = False Then
        daño = daño + 0
        Else
        daño = daño + (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus)
    End If
    End If

    
        If UCase$(UserList(userindex).Clase) = "DRUIDA" Then
    If UserList(userindex).Invent.WeaponEqpObjIndex = 946 = False Then
        daño = daño + 0
        Else
        daño = daño + (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus)
    End If
    End If
    
            If UCase$(UserList(userindex).Clase) = "DRUIDA" Then
    If UserList(userindex).Invent.WeaponEqpObjIndex = 947 = False Then
        daño = daño + 0
        Else
        daño = daño + (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus)
    End If
    End If
    
            If UCase$(UserList(userindex).Clase) = "DRUIDA" Then
    If UserList(userindex).Invent.WeaponEqpObjIndex = 659 = False Then
        daño = daño + 0
        Else
        daño = daño + (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus)
    End If
    End If
    
    If UserList(tempChr).flags.GemaActivada = "Celeste" Then
        daño = daño - (daño * 10 / 100)
    End If
    
    If Hechizos(h).StaffAffected Then
        If UCase$(UserList(userindex).Clase) = "MAGO" Then
            If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                daño = (daño * (ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            Else
                daño = daño * 0.7 'Baja daño a 70% del original
            End If
        End If
    End If
    
    'cascos antimagia
    If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)
    End If
    
    'Armaduras antimagia
     If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).DefensaMagicaMax)
    End If
    
    'anillos
    If (UserList(tempChr).Invent.HerramientaEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.HerramientaEqpObjIndex).DefensaMagicaMax)
    End If
    
    If daño < 0 Then daño = 0
    
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    

    
    Call InfoHechizo(userindex)
    
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - daño
    
    Call SendData(SendTarget.toindex, userindex, 0, "||Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).name & FONTTYPE_BORDON)
    Call SendData(SendTarget.toindex, tempChr, 0, "||" & UserList(userindex).name & " te ha quitado " & daño & " puntos de vida." & FONTTYPE_BORDON)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & &HFFFF& & "°" & "- " & daño & "" & "°" & str(UserList(tempChr).Char.CharIndex))
    
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
        Call ContarMuerte(tempChr, userindex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, userindex)
    End If
    
    b = True
End If

'Mana
If Hechizos(h).SubeMana = 1 Then
    
    Call InfoHechizo(userindex)
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + daño
    If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMan Then _
        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMan
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Le has restaurado " & daño & " puntos de mana a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.toindex, tempChr, 0, "||" & UserList(userindex).name & " te ha restaurado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||Te has restaurado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(h).SubeMana = 2 Then
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Le has quitado " & daño & " puntos de mana a " & UserList(tempChr).name & FONTTYPE_BORDON)
        Call SendData(SendTarget.toindex, tempChr, 0, "||" & UserList(userindex).name & " te ha quitado " & daño & " puntos de mana." & FONTTYPE_BORDON)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||Te has quitado " & daño & " puntos de mana." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - daño
    If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
    b = True
    
End If

'Stamina
If Hechizos(h).SubeSta = 1 Then
    Call InfoHechizo(userindex)
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + daño
    If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then _
        UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta
    If userindex <> tempChr Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Le has restaurado " & daño & " puntos de vitalidad a " & UserList(tempChr).name & FONTTYPE_FIGHT)
        Call SendData(SendTarget.toindex, tempChr, 0, "||" & UserList(userindex).name & " te ha restaurado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||Te has restaurado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(h).SubeMana = 2 Then
    If Not PuedeAtacar(userindex, tempChr) Then Exit Sub
    
    If userindex <> tempChr Then
        Call UsuarioAtacadoPorUsuario(userindex, tempChr)
    End If
    
    Call InfoHechizo(userindex)
    
    If userindex <> tempChr Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Le has quitado " & daño & " puntos de vitalidad a " & UserList(tempChr).name & FONTTYPE_BORDON)
        Call SendData(SendTarget.toindex, tempChr, 0, "||" & UserList(userindex).name & " te ha quitado " & daño & " puntos de vitalidad." & FONTTYPE_BORDON)
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||Te has quitado " & daño & " puntos de vitalidad." & FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - daño
    
    If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
    b = True
End If

End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal userindex As Integer, ByVal Slot As Byte)

'Call LogTarea("Sub UpdateUserHechizos")

Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(userindex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(userindex, Slot, UserList(userindex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(userindex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For LoopC = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(userindex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(userindex, LoopC, UserList(userindex).Stats.UserHechizos(LoopC))
        Else
            Call ChangeUserHechizo(userindex, LoopC, 0)
        End If

Next LoopC

End If

End Sub

Sub ChangeUserHechizo(ByVal userindex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

UserList(userindex).Stats.UserHechizos(Slot) = Hechizo


If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then

    Call SendData(SendTarget.toindex, userindex, 0, "ATK" & Slot & "," & Hechizo & "," & Hechizos(Hechizo).Nombre)

Else

    Call SendData(SendTarget.toindex, userindex, 0, "ATK" & Slot & "," & "0" & "," & "(Nada)")

End If


End Sub


Public Sub DesplazarHechizo(ByVal userindex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)

If Not (Dire >= 1 And Dire <= 2) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

If Dire = 1 Then 'Mover arriba
    If CualHechizo = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes mover el hechizo en esa direccion." & FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(userindex).Stats.UserHechizos(CualHechizo)
        UserList(userindex).Stats.UserHechizos(CualHechizo) = UserList(userindex).Stats.UserHechizos(CualHechizo - 1)
        UserList(userindex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo
        
        Call UpdateUserHechizos(False, userindex, CualHechizo - 1)
    End If
Else 'mover abajo
    If CualHechizo = MAXUSERHECHIZOS Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes mover el hechizo en esa direccion." & FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(userindex).Stats.UserHechizos(CualHechizo)
        UserList(userindex).Stats.UserHechizos(CualHechizo) = UserList(userindex).Stats.UserHechizos(CualHechizo + 1)
        UserList(userindex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo
        
        Call UpdateUserHechizos(False, userindex, CualHechizo + 1)
    End If
End If
Call UpdateUserHechizos(False, userindex, CualHechizo)

End Sub


Public Sub DisNobAuBan(ByVal userindex As Integer, NoblePts As Long, BandidoPts As Long)
'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos

    'Si estamos en la arena no hacemos nada
    If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = 6 Then Exit Sub
    
    'pierdo nobleza...
    UserList(userindex).Reputacion.NobleRep = UserList(userindex).Reputacion.NobleRep - NoblePts
    If UserList(userindex).Reputacion.NobleRep < 0 Then
        UserList(userindex).Reputacion.NobleRep = 0
    End If
    
    'gano bandido...
    UserList(userindex).Reputacion.BandidoRep = UserList(userindex).Reputacion.BandidoRep + BandidoPts
    UserList(userindex).Reputacion.BurguesRep = 0
    UserList(userindex).Reputacion.NobleRep = 0
    UserList(userindex).Reputacion.PlebeRep = 0
    If UserList(userindex).Reputacion.BandidoRep > MAXREP Then _
        UserList(userindex).Reputacion.BandidoRep = MAXREP
    If Criminal(userindex) Then If UserList(userindex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(userindex)
End Sub
