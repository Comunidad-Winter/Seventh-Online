Attribute VB_Name = "Trabajo"
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

Public Sub DoPermanecerOculto(ByVal userindex As Integer)
On Error GoTo errhandler
Dim Suerte As Integer
Dim res As Integer

If UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 10 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= -1 Then
                    Suerte = 35
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 20 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 11 Then
                    Suerte = 30
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 30 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 21 Then
                    Suerte = 28
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 40 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 31 Then
                    Suerte = 24
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 50 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 41 Then
                    Suerte = 22
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 60 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 51 Then
                    Suerte = 20
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 70 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 61 Then
                    Suerte = 18
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 80 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 71 Then
                    Suerte = 15
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 90 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 81 Then
                    Suerte = 10
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 100 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 91 Then
                    Suerte = 10     'Lo atamos con alambre.... en la 11.6 el sistema de ocultarse debería de estar bien hecho
End If

If UCase$(UserList(userindex).Clase) <> "LADRON" Then Suerte = Suerte + 50

'cazador con armadura de cazador oculto no se hace visible
If UCase$(UserList(userindex).Clase) = "CAZADOR" And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) > 90 Then
    If UserList(userindex).Invent.ArmourEqpObjIndex = 648 Or UserList(userindex).Invent.ArmourEqpObjIndex = 360 Then
        Exit Sub
    End If
End If


res = RandomNumber(1, Suerte)

If res > 9 Then
    UserList(userindex).flags.Oculto = 0
    If UserList(userindex).flags.Invisible = 0 Then
        'no hace falta encriptar este (se jode el gil que bypassea esto)
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
        Call SendData(SendTarget.toindex, userindex, 0, "||¡Has vuelto a ser visible!" & FONTTYPE_INFO)
    End If
End If


Exit Sub

errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub

Public Sub DoOcultarse(ByVal userindex As Integer)

On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 10 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= -1 Then
                    Suerte = 35
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 20 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 11 Then
                    Suerte = 30
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 30 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 21 Then
                    Suerte = 28
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 40 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 31 Then
                    Suerte = 24
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 50 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 41 Then
                    Suerte = 22
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 60 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 51 Then
                    Suerte = 20
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 70 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 61 Then
                    Suerte = 18
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 80 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 71 Then
                    Suerte = 15
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 90 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 81 Then
                    Suerte = 10
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) <= 100 _
   And UserList(userindex).Stats.UserSkills(eSkill.Ocultarse) >= 91 Then
                    Suerte = 7
End If

If UCase$(UserList(userindex).Clase) <> "LADRON" Then Suerte = Suerte + 50

res = RandomNumber(1, Suerte)

If res <= 5 Then
    UserList(userindex).flags.Oculto = 1
#If SeguridadAlkon Then
    If EncriptarProtocolosCriticos Then
        Call SendCryptedData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",1")
    Else
#End If
        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",1")
#If SeguridadAlkon Then
    End If
#End If
    Call SendData(SendTarget.toindex, userindex, 0, "||¡Te has escondido entre las sombras!" & FONTTYPE_INFO)
    Call SubirSkill(userindex, Ocultarse)
Else
    '[CDT 17-02-2004]
    If Not UserList(userindex).flags.UltimoMensaje = 4 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||¡No has logrado esconderte!" & FONTTYPE_INFO)
        UserList(userindex).flags.UltimoMensaje = 4
    End If
    '[/CDT]
End If

UserList(userindex).Counters.Ocultando = UserList(userindex).Counters.Ocultando + 1

Exit Sub

errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub


Public Sub DoNavega(ByVal userindex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)

Dim ModNave As Long
ModNave = ModNavegacion(UserList(userindex).Clase)

If UserList(userindex).Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No tenes suficientes conocimientos para usar este barco." & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, userindex, 0, "||Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion." & FONTTYPE_INFO)
    Exit Sub
End If

UserList(userindex).Invent.BarcoObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
UserList(userindex).Invent.BarcoSlot = Slot

If UserList(userindex).flags.Navegando = 0 Then
    
    UserList(userindex).Char.Head = 0
    
    If UserList(userindex).flags.Muerto = 0 Then
        UserList(userindex).Char.Body = Barco.Ropaje
    Else
        UserList(userindex).Char.Body = iFragataFantasmal
    End If
    
    UserList(userindex).Char.ShieldAnim = NingunEscudo
    UserList(userindex).Char.WeaponAnim = NingunArma
    UserList(userindex).Char.CascoAnim = NingunCasco
    UserList(userindex).flags.Navegando = 1
    
Else
    
    UserList(userindex).flags.Navegando = 0
    
    If UserList(userindex).flags.Muerto = 0 Then
        UserList(userindex).Char.Head = UserList(userindex).OrigChar.Head
        
        If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(userindex).Char.Body = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(userindex)
        End If
        
        If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(userindex).Char.ShieldAnim = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(userindex).Char.WeaponAnim = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(userindex).Char.CascoAnim = ObjData(UserList(userindex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(userindex).Char.Body = iCuerpoMuerto
        UserList(userindex).Char.Head = iCabezaMuerto
        UserList(userindex).Char.ShieldAnim = NingunEscudo
        UserList(userindex).Char.WeaponAnim = NingunArma
        UserList(userindex).Char.CascoAnim = NingunCasco
    End If
End If

Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
Call SendData(SendTarget.toindex, userindex, 0, "NAVEG")

End Sub

Public Sub FundirMineral(ByVal userindex As Integer)
'Call LogTarea("Sub FundirMineral")

If UserList(userindex).flags.TargetObjInvIndex > 0 Then
   
   If ObjData(UserList(userindex).flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData(UserList(userindex).flags.TargetObjInvIndex).MinSkill <= UserList(userindex).Stats.UserSkills(eSkill.Mineria) / ModFundicion(UserList(userindex).Clase) Then
        Call DoLingotes(userindex)
   Else
        Call SendData(SendTarget.toindex, userindex, 0, "||No tenes conocimientos de mineria suficientes para trabajar este mineral." & FONTTYPE_INFO)
   End If

End If

End Sub
Function TieneObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal userindex As Integer) As Boolean
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(userindex).Invent.Object(i).ObjIndex = ItemIndex Then
        Total = Total + UserList(userindex).Invent.Object(i).Amount
    End If
Next i

If Cant <= Total Then
    TieneObjetos = True
    Exit Function
End If
        
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal userindex As Integer) As Boolean
'Call LogTarea("Sub QuitarObjetos")

Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(userindex).Invent.Object(i).ObjIndex = ItemIndex Then
        
        Call Desequipar(userindex, i)
        
        UserList(userindex).Invent.Object(i).Amount = UserList(userindex).Invent.Object(i).Amount - Cant
        If (UserList(userindex).Invent.Object(i).Amount <= 0) Then
            Cant = Abs(UserList(userindex).Invent.Object(i).Amount)
            UserList(userindex).Invent.Object(i).Amount = 0
            UserList(userindex).Invent.Object(i).ObjIndex = 0
        Else
            Cant = 0
        End If
        
        Call UpdateUserInv(False, userindex, i)
        
        If (Cant = 0) Then
            QuitarObjetos = True
            Exit Function
        End If
    End If
Next i

End Function

Sub HerreroQuitarMateriales(ByVal userindex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, userindex)
    If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, userindex)
    If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, userindex)
End Sub

Sub CarpinteroQuitarMateriales(ByVal userindex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera, userindex)
End Sub

Function CarpinteroTieneMateriales(ByVal userindex As Integer, ByVal ItemIndex As Integer) As Boolean
    
    If ObjData(ItemIndex).Madera > 0 Then
            If Not TieneObjetos(Leña, ObjData(ItemIndex).Madera, userindex) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No tenes suficientes madera." & FONTTYPE_INFO)
                    CarpinteroTieneMateriales = False
                    Exit Function
            End If
    End If
    
    CarpinteroTieneMateriales = True

End Function
 
Function HerreroTieneMateriales(ByVal userindex As Integer, ByVal ItemIndex As Integer) As Boolean
    If ObjData(ItemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, userindex) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No tenes suficientes lingotes de hierro." & FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, userindex) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No tenes suficientes lingotes de plata." & FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO, userindex) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No tenes suficientes lingotes de oro." & FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    HerreroTieneMateriales = True
End Function

Public Function PuedeConstruir(ByVal userindex As Integer, ByVal ItemIndex As Integer) As Boolean
PuedeConstruir = HerreroTieneMateriales(userindex, ItemIndex) And UserList(userindex).Stats.UserSkills(eSkill.Herreria) >= _
 ObjData(ItemIndex).SkHerreria
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ArmasHerrero)
    If ArmasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
For i = 1 To UBound(ArmadurasHerrero)
    If ArmadurasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
PuedeConstruirHerreria = False
End Function


Public Sub HerreroConstruirItem(ByVal userindex As Integer, ByVal ItemIndex As Integer)
'Call LogTarea("Sub HerreroConstruirItem")
If PuedeConstruir(userindex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
    Call HerreroQuitarMateriales(userindex, ItemIndex)
    ' AGREGAR FX
    If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Has construido el arma!." & FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otESCUDO Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Has construido el escudo!." & FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCASCO Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Has construido el casco!." & FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Has construido la armadura!." & FONTTYPE_INFO)
    End If
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(userindex, MiObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
    End If
    Call SubirSkill(userindex, Herreria)
    Call UpdateUserInv(True, userindex, 0)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & MARTILLOHERRERO)
    
End If

UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando + 1

End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjCarpintero)
    If ObjCarpintero(i) = ItemIndex Then
        PuedeConstruirCarpintero = True
        Exit Function
    End If
Next i
PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal userindex As Integer, ByVal ItemIndex As Integer)

If CarpinteroTieneMateriales(userindex, ItemIndex) And _
   UserList(userindex).Stats.UserSkills(eSkill.Carpinteria) >= _
   ObjData(ItemIndex).SkCarpinteria And _
   PuedeConstruirCarpintero(ItemIndex) And _
   UserList(userindex).Invent.HerramientaEqpObjIndex = SERRUCHO_CARPINTERO Then

    Call CarpinteroQuitarMateriales(userindex, ItemIndex)
    Call SendData(SendTarget.toindex, userindex, 0, "||Has construido el objeto!" & FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(userindex, MiObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
    End If
    
    Call SubirSkill(userindex, Carpinteria)
    Call UpdateUserInv(True, userindex, 0)
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & LABUROCARPINTERO)
End If

UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando + 1

End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer
    Select Case Lingote
        Case iMinerales.HierroCrudo
            MineralesParaLingote = 13
        Case iMinerales.PlataCruda
            MineralesParaLingote = 25
        Case iMinerales.OroCrudo
            MineralesParaLingote = 50
        Case Else
            MineralesParaLingote = 10000
    End Select
End Function


Public Sub DoLingotes(ByVal userindex As Integer)
'    Call LogTarea("Sub DoLingotes")
Dim Slot As Integer
Dim obji As Integer

    Slot = UserList(userindex).flags.TargetObjInvSlot
    obji = UserList(userindex).Invent.Object(Slot).ObjIndex
    
    If UserList(userindex).Invent.Object(Slot).Amount < MineralesParaLingote(obji) Or _
        ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes minerales para hacer un lingote." & FONTTYPE_INFO)
            Exit Sub
    End If
    
    UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount - MineralesParaLingote(obji)
    If UserList(userindex).Invent.Object(Slot).Amount < 1 Then
        UserList(userindex).Invent.Object(Slot).Amount = 0
        UserList(userindex).Invent.Object(Slot).ObjIndex = 0
    End If
    Call SendData(SendTarget.toindex, userindex, 0, "||Has obtenido un lingote!!!" & FONTTYPE_INFO)
    Dim nPos As WorldPos
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = ObjData(UserList(userindex).flags.TargetObjInvIndex).LingoteIndex
    If Not MeterItemEnInventario(userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
    End If
    Call UpdateUserInv(False, userindex, Slot)
    Call SendData(SendTarget.toindex, userindex, 0, "||¡Has obtenido un lingote!" & FONTTYPE_INFO)
    


UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando + 1

End Sub

Function ModNavegacion(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "PIRATA"
        ModNavegacion = 1
    Case "PESCADOR"
        ModNavegacion = 1.2
    Case Else
        ModNavegacion = 2.3
End Select

End Function


Function ModFundicion(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "MINERO"
        ModFundicion = 1
    Case "HERRERO"
        ModFundicion = 1.2
    Case Else
        ModFundicion = 3
End Select

End Function

Function ModCarpinteria(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "CARPINTERO"
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
End Select

End Function

Function ModHerreriA(ByVal Clase As String) As Integer

Select Case UCase$(Clase)
    Case "HERRERO"
        ModHerreriA = 1
    Case "MINERO"
        ModHerreriA = 1.2
    Case Else
        ModHerreriA = 4
End Select

End Function

Function ModDomar(ByVal Clase As String) As Integer
    Select Case UCase$(Clase)
        Case "DRUIDA"
            ModDomar = 6
        Case "CAZADOR"
            ModDomar = 6
        Case "CLERIGO"
            ModDomar = 7
        Case Else
            ModDomar = 10
    End Select
End Function

Function CalcularPoderDomador(ByVal userindex As Integer) As Long
    With UserList(userindex).Stats
        CalcularPoderDomador = .UserAtributos(eAtributos.Carisma) _
            * (.UserSkills(eSkill.Domar) / ModDomar(UserList(userindex).Clase)) _
            + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3) _
            + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3) _
            + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3)
    End With
End Function

Function FreeMascotaIndex(ByVal userindex As Integer) As Integer
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(userindex).MascotasIndex(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next j
End Function

Sub DoDomar(ByVal userindex As Integer, ByVal NpcIndex As Integer)
'Call LogTarea("Sub DoDomar")

If UserList(userindex).NroMacotas < MAXMASCOTAS Then
    
    If Npclist(NpcIndex).MaestroUser = userindex Then
        Call SendData(SendTarget.toindex, userindex, 0, "||La criatura ya te ha aceptado como su amo." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||La criatura ya tiene amo." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(NpcIndex).flags.Domable <= CalcularPoderDomador(userindex) Then
        Dim index As Integer
        UserList(userindex).NroMacotas = UserList(userindex).NroMacotas + 1
        index = FreeMascotaIndex(userindex)
        UserList(userindex).MascotasIndex(index) = NpcIndex
        UserList(userindex).MascotasType(index) = Npclist(NpcIndex).Numero
        
        Npclist(NpcIndex).MaestroUser = userindex
        
        Call FollowAmo(NpcIndex)
        
        Call SendData(SendTarget.toindex, userindex, 0, "||La criatura te ha aceptado como su amo." & FONTTYPE_INFO)
        Call SubirSkill(userindex, Domar)
    Else
        If Not UserList(userindex).flags.UltimoMensaje = 5 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No has logrado domar la criatura." & FONTTYPE_INFO)
            UserList(userindex).flags.UltimoMensaje = 5
        End If
    End If
Else
    Call SendData(SendTarget.toindex, userindex, 0, "||No podes controlar mas criaturas." & FONTTYPE_INFO)
End If
End Sub

Sub DoAdminInvisible(ByVal userindex As Integer)
    
    If UserList(userindex).flags.AdminInvisible = 0 Then
        
        ' Sacamos el mimetizmo
        If UserList(userindex).flags.Mimetizado = 1 Then
            UserList(userindex).Char.Body = UserList(userindex).CharMimetizado.Body
            UserList(userindex).Char.Head = UserList(userindex).CharMimetizado.Head
            UserList(userindex).Char.CascoAnim = UserList(userindex).CharMimetizado.CascoAnim
            UserList(userindex).Char.ShieldAnim = UserList(userindex).CharMimetizado.ShieldAnim
            UserList(userindex).Char.WeaponAnim = UserList(userindex).CharMimetizado.WeaponAnim
            UserList(userindex).Counters.Mimetismo = 0
            UserList(userindex).flags.Mimetizado = 0
        End If
        
        UserList(userindex).flags.AdminInvisible = 1
        UserList(userindex).flags.Invisible = 1
        UserList(userindex).flags.Oculto = 1
        UserList(userindex).flags.OldBody = UserList(userindex).Char.Body
        UserList(userindex).flags.OldHead = UserList(userindex).Char.Head
        UserList(userindex).Char.Body = 0
        UserList(userindex).Char.Head = 0
        
    Else
        
        UserList(userindex).flags.AdminInvisible = 0
        UserList(userindex).flags.Invisible = 0
        UserList(userindex).flags.Oculto = 0
        UserList(userindex).Char.Body = UserList(userindex).flags.OldBody
        UserList(userindex).Char.Head = UserList(userindex).flags.OldHead
        
    End If
    
    'vuelve a ser visible por la fuerza
    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
    Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal userindex As Integer)

Dim Suerte As Byte
Dim exito As Byte
Dim raise As Byte
Dim Obj As Obj
Dim posMadera As WorldPos

If Not LegalPos(Map, X, Y) Then Exit Sub

With posMadera
    .Map = Map
    .X = X
    .Y = Y
End With

If Distancia(posMadera, UserList(userindex).pos) > 2 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Estás demasiado lejos para prender la fogata." & FONTTYPE_INFO)
    Exit Sub
End If

If UserList(userindex).flags.Muerto = 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes hacer fogatas estando muerto." & FONTTYPE_INFO)
    Exit Sub
End If

If MapData(Map, X, Y).OBJInfo.Amount < 3 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Necesitas por lo menos tres troncos para hacer una fogata." & FONTTYPE_INFO)
    Exit Sub
End If


If UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) < 6 Then
    Suerte = 3
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) >= 6 And UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) <= 34 Then
    Suerte = 2
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Supervivencia) >= 35 Then
    Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.ObjIndex = FOGATA_APAG
    Obj.Amount = MapData(Map, X, Y).OBJInfo.Amount \ 3
    
    Call SendData(SendTarget.toindex, userindex, 0, "||Has hecho " & Obj.Amount & " fogatas." & FONTTYPE_INFO)
    
    Call MakeObj(SendTarget.ToMap, 0, Map, Obj, Map, X, Y)
    
    'Seteamos la fogata como el nuevo TargetObj del user
    UserList(userindex).flags.TargetObj = FOGATA_APAG
Else
    '[CDT 17-02-2004]
    If Not UserList(userindex).flags.UltimoMensaje = 10 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No has podido hacer la fogata." & FONTTYPE_INFO)
        UserList(userindex).flags.UltimoMensaje = 10
    End If
    '[/CDT]
End If

Call SubirSkill(userindex, Supervivencia)


End Sub

Public Sub DoPescar(ByVal userindex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer


If UCase$(UserList(userindex).Clase) = "PESCADOR" Then
    Call QuitarSta(userindex, EsfuerzoPescarPescador)
Else
    Call QuitarSta(userindex, EsfuerzoPescarGeneral)
End If

If UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 10 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= -1 Then
                    Suerte = 35
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 20 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 11 Then
                    Suerte = 30
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 30 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 21 Then
                    Suerte = 28
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 40 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 31 Then
                    Suerte = 24
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 50 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 41 Then
                    Suerte = 22
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 60 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 51 Then
                    Suerte = 20
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 70 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 61 Then
                    Suerte = 18
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 80 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 71 Then
                    Suerte = 15
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 90 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 81 Then
                    Suerte = 13
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Pesca) <= 100 _
   And UserList(userindex).Stats.UserSkills(eSkill.Pesca) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    MiObj.Amount = 1
    MiObj.ObjIndex = Pescado
    
    If Not MeterItemEnInventario(userindex, MiObj) Then
        Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
    End If
    
    Call SendData(SendTarget.toindex, userindex, 0, "||¡Has pescado un lindo pez!" & FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(userindex).flags.UltimoMensaje = 6 Then
      Call SendData(SendTarget.toindex, userindex, 0, "||¡No has pescado nada!" & FONTTYPE_INFO)
      UserList(userindex).flags.UltimoMensaje = 6
    End If
    '[/CDT]
End If

Call SubirSkill(userindex, Pesca)

UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando + 1

Exit Sub

errhandler:
    Call LogError("Error en DoPescar")
End Sub

Public Sub DoPescarRed(ByVal userindex As Integer)
On Error GoTo errhandler

Dim iSkill As Integer
Dim Suerte As Integer
Dim res As Integer
Dim EsPescador As Boolean

If UCase(UserList(userindex).Clase) = "PESCADOR" Then
    Call QuitarSta(userindex, EsfuerzoPescarPescador)
    EsPescador = True
Else
    Call QuitarSta(userindex, EsfuerzoPescarGeneral)
    EsPescador = False
End If

iSkill = UserList(userindex).Stats.UserSkills(eSkill.Pesca)

' m = (60-11)/(1-10)
' y = mx - m*10 + 11

Select Case iSkill
Case 0:         Suerte = 0
Case 1 To 10:   Suerte = 60
Case 11 To 20:  Suerte = 54
Case 21 To 30:  Suerte = 49
Case 31 To 40:  Suerte = 43
Case 41 To 50:  Suerte = 38
Case 51 To 60:  Suerte = 32
Case 61 To 70:  Suerte = 27
Case 71 To 80:  Suerte = 21
Case 81 To 90:  Suerte = 16
Case 91 To 100: Suerte = 11
Case Else:      Suerte = 0
End Select

If Suerte > 0 Then
    res = RandomNumber(1, Suerte)
    
    If res < 6 Then
        Dim nPos As WorldPos
        Dim MiObj As Obj
        Dim PecesPosibles(1 To 4) As Integer
        
        PecesPosibles(1) = PESCADO1
        PecesPosibles(2) = PESCADO2
        PecesPosibles(3) = PESCADO3
        PecesPosibles(4) = PESCADO4
        
        If EsPescador = True Then
            MiObj.Amount = RandomNumber(1, 5)
        Else
            MiObj.Amount = 1
        End If
        MiObj.ObjIndex = PecesPosibles(RandomNumber(LBound(PecesPosibles), UBound(PecesPosibles)))
        
        If Not MeterItemEnInventario(userindex, MiObj) Then
            Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
        End If
        
        Call SendData(SendTarget.toindex, userindex, 0, "||¡Has pescado algunos peces!" & FONTTYPE_INFO)
        
    Else
        Call SendData(SendTarget.toindex, userindex, 0, "||¡No has pescado nada!" & FONTTYPE_INFO)
    End If
    
    Call SubirSkill(userindex, Pesca)
End If

Exit Sub

errhandler:
    Call LogError("Error en DoPescarRed")
End Sub

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)

If Not MapInfo(UserList(VictimaIndex).pos.Map).Pk Then Exit Sub

If UserList(LadrOnIndex).flags.Seguro Then
    Call SendData(SendTarget.toindex, LadrOnIndex, 0, "||Debes quitar el seguro para robar" & FONTTYPE_FIGHT)
    Exit Sub
End If

If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And UserList(LadrOnIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(SendTarget.toindex, LadrOnIndex, 0, "||No puedes robar a otros miembros de las fuerzas del caos" & FONTTYPE_FIGHT)
    Exit Sub
End If

If UserList(VictimaIndex).flags.Privilegios = PlayerType.User Then
    Dim Suerte As Integer
    Dim res As Integer
    
    If UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 10 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= -1 Then
                        Suerte = 35
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 20 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 11 Then
                        Suerte = 30
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 30 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 21 Then
                        Suerte = 28
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 40 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 31 Then
                        Suerte = 24
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 50 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 41 Then
                        Suerte = 22
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 60 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 51 Then
                        Suerte = 20
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 70 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 61 Then
                        Suerte = 18
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 80 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 71 Then
                        Suerte = 15
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 90 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 81 Then
                        Suerte = 10
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 100 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 91 Then
                        Suerte = 5
    End If
    res = RandomNumber(1, Suerte)
    
    If res < 3 Then 'Exito robo
       
        If (RandomNumber(1, 50) < 25) And (UCase$(UserList(LadrOnIndex).Clase) = "LADRON") Then
            If TieneObjetosRobables(VictimaIndex) Then
                Call RobarObjeto(LadrOnIndex, VictimaIndex)
            Else
                Call SendData(SendTarget.toindex, LadrOnIndex, 0, "||" & UserList(VictimaIndex).name & " no tiene objetos." & FONTTYPE_INFO)
            End If
    Else
        Call SendData(SendTarget.toindex, LadrOnIndex, 0, "||¡No has logrado robar nada!" & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).name & " ha intentado robarte!" & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, VictimaIndex, 0, "||¡" & UserList(LadrOnIndex).name & " es un criminal!" & FONTTYPE_INFO)
    End If
    End If

'Si es Ciudadano, lo hacemos neutral
If Ciudadano(LadrOnIndex) And Ciudadano(VictimaIndex) Then
Call SendData(SendTarget.toindex, LadrOnIndex, 0, "||No puedes robarle a ciudadanos, escribe /RENUNCIAR." & FONTTYPE_INFO)
Exit Sub
End If
'Por las dudas,, si elegio ser Criminal
'Mithrandir
    
    If UserList(LadrOnIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(LadrOnIndex)

    UserList(LadrOnIndex).Reputacion.LadronesRep = UserList(LadrOnIndex).Reputacion.LadronesRep + vlLadron
    If UserList(LadrOnIndex).Reputacion.LadronesRep > MAXREP Then _
        UserList(LadrOnIndex).Reputacion.LadronesRep = MAXREP
    Call SubirSkill(LadrOnIndex, Robar)
End If


End Sub


Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
' Agregué los barcos
' Esta funcion determina qué objetos son robables.

Dim OI As Integer

OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

ObjEsRobable = _
ObjData(OI).OBJType <> eOBJType.otLlaves And _
UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
ObjData(OI).Real = 0 And _
ObjData(OI).Caos = 0 And _
ObjData(OI).OBJType <> eOBJType.otMonturas And _
ObjData(OI).OBJType <> eOBJType.otBarcos

End Function

Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'Call LogTarea("Sub RobarObjeto")
Dim flag As Boolean
Dim i As Integer
flag = False

If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
    i = 1
    Do While Not flag And i <= MAX_INVENTORY_SLOTS
        'Hay objeto en este slot?
        If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
           If ObjEsRobable(VictimaIndex, i) Then
                 If RandomNumber(1, 10) < 4 Then flag = True
           End If
        End If
        If Not flag Then i = i + 1
    Loop
Else
    i = 20
    Do While Not flag And i > 0
      'Hay objeto en este slot?
      If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
         If ObjEsRobable(VictimaIndex, i) Then
               If RandomNumber(1, 10) < 4 Then flag = True
         End If
      End If
      If Not flag Then i = i - 1
    Loop
End If

If flag Then
    Dim MiObj As Obj
    Dim num As Byte
    'Cantidad al azar
    num = RandomNumber(1, 5)
                
    If num > UserList(VictimaIndex).Invent.Object(i).Amount Then
         num = UserList(VictimaIndex).Invent.Object(i).Amount
    End If
                
    MiObj.Amount = num
    MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    
    UserList(VictimaIndex).Invent.Object(i).Amount = UserList(VictimaIndex).Invent.Object(i).Amount - num
                
    If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
          Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
    End If
            
    Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).pos, MiObj)
    End If
    
    Call SendData(SendTarget.toindex, LadrOnIndex, 0, "||Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).name & FONTTYPE_INFO)
Else
    Call SendData(SendTarget.toindex, LadrOnIndex, 0, "||No has logrado robar un objetos." & FONTTYPE_INFO)
End If

End Sub
Public Sub DoApuñalar(ByVal userindex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)

Dim Suerte As Integer
Dim res As Integer

If UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 10 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= -1 Then
                    Suerte = 200
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 20 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 11 Then
                    Suerte = 190
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 30 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 21 Then
                    Suerte = 180
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 40 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 31 Then
                    Suerte = 170
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 50 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 41 Then
                    Suerte = 160
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 60 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 51 Then
                    Suerte = 150
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 70 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 61 Then
                    Suerte = 140
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 80 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 71 Then
                    Suerte = 130
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) <= 90 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 81 Then
                    Suerte = 120
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) < 100 _
   And UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) >= 91 Then
                    Suerte = 110
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Apuñalar) = 100 Then
                    Suerte = 100
End If

If UCase$(UserList(userindex).Clase) = "ASESINO" Then
    res = RandomNumber(0, Suerte / 1.1)
    If res < 25 Then res = 0
Else
    res = RandomNumber(0, Suerte * 1.35)
End If

If res < 15 Then
    If VictimUserIndex <> 0 Then
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - Int(daño * 1.5)
        Call SendData(SendTarget.toindex, userindex, 0, "||Has apuñalado a " & UserList(VictimUserIndex).name & " por " & Int(daño * 1.5) & FONTTYPE_FIGHT)
        Call SendData(SendTarget.toindex, VictimUserIndex, 0, "||Te ha apuñalado " & UserList(userindex).name & " por " & Int(daño * 1.5) & FONTTYPE_FIGHT)
        Call SendData(SendTarget.ToPCArea, VictimUserIndex, UserList(VictimUserIndex).pos.Map, "XFC" & UserList(VictimUserIndex).Char.CharIndex & "," & FXAPUÑALAR & "," & 0)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - Int(daño * 2)
        Call SendData(SendTarget.toindex, userindex, 0, "||Has apuñalado la criatura por " & Int(daño * 2) & FONTTYPE_FIGHT)
        Call SubirSkill(userindex, Apuñalar)
        '[Alejo]
        Call CalcularDarExp(userindex, VictimNpcIndex, Int(daño * 2))
    End If
Else
    Call SendData(SendTarget.toindex, userindex, 0, "||¡No has logrado apuñalar a tu enemigo!" & FONTTYPE_FIGHT)
End If

End Sub

Public Sub QuitarSta(ByVal userindex As Integer, ByVal Cantidad As Integer)
UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta - Cantidad
If UserList(userindex).Stats.MinSta < 0 Then UserList(userindex).Stats.MinSta = 0
End Sub

Public Sub DoTalar(ByVal userindex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer


If UCase$(UserList(userindex).Clase) = "LEÑADOR" Then
    Call QuitarSta(userindex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(userindex, EsfuerzoTalarGeneral)
End If

If UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 10 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= -1 Then
                    Suerte = 35
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 20 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 11 Then
                    Suerte = 30
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 30 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 21 Then
                    Suerte = 28
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 40 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 31 Then
                    Suerte = 24
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 50 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 41 Then
                    Suerte = 22
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 60 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 51 Then
                    Suerte = 20
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 70 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 61 Then
                    Suerte = 18
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 80 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 71 Then
                    Suerte = 15
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 90 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 81 Then
                    Suerte = 13
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Talar) <= 100 _
   And UserList(userindex).Stats.UserSkills(eSkill.Talar) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    If UCase$(UserList(userindex).Clase) = "LEÑADOR" Then
        MiObj.Amount = RandomNumber(1, 5)
    Else
        MiObj.Amount = 1
    End If
    
    MiObj.ObjIndex = Leña
    
    
    If Not MeterItemEnInventario(userindex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
        
    End If
    
    Call SendData(SendTarget.toindex, userindex, 0, "||¡Has conseguido algo de leña!" & FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(userindex).flags.UltimoMensaje = 8 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||¡No has obtenido leña!" & FONTTYPE_INFO)
        UserList(userindex).flags.UltimoMensaje = 8
    End If
    '[/CDT]
End If

Call SubirSkill(userindex, Talar)

UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando + 1

Exit Sub

errhandler:
    Call LogError("Error en DoTalar")

End Sub

Sub VolverCriminal(ByVal userindex As Integer)
'Mithrandir
If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = 6 Then Exit Sub
If UserList(userindex).pos.Map = MapCastilloN Then Exit Sub
If UserList(userindex).pos.Map = MapCastilloS Then Exit Sub
If UserList(userindex).pos.Map = MapCastilloE Then Exit Sub
If UserList(userindex).pos.Map = MapCastilloO Then Exit Sub
'Criminal
UserList(userindex).StatusMith.EsStatus = 2
UserList(userindex).StatusMith.EligioStatus = 1
Call SendUserStatux(userindex)
End Sub

Sub VolverCiudadano(ByVal userindex As Integer)
If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = 6 Then Exit Sub
'Ciudadano
If UserList(userindex).pos.Map = MapCastilloN Then Exit Sub
If UserList(userindex).pos.Map = MapCastilloS Then Exit Sub
If UserList(userindex).pos.Map = MapCastilloE Then Exit Sub
If UserList(userindex).pos.Map = MapCastilloO Then Exit Sub
UserList(userindex).StatusMith.EsStatus = 1
UserList(userindex).StatusMith.EligioStatus = 1
'Mithrandir
Call SendUserStatux(userindex)
End Sub
 
Sub VolverNeutral(ByVal userindex As Integer)
If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = 6 Then Exit Sub
'Neutral
If UserList(userindex).pos.Map = MapCastilloN Then Exit Sub
If UserList(userindex).pos.Map = MapCastilloS Then Exit Sub
If UserList(userindex).pos.Map = MapCastilloE Then Exit Sub
If UserList(userindex).pos.Map = MapCastilloO Then Exit Sub
UserList(userindex).StatusMith.EsStatus = 0
UserList(userindex).StatusMith.EligioStatus = 1
'Mithrandir
Call SendUserStatux(userindex)
End Sub

Public Sub DoMineria(ByVal userindex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer
Dim metal As Integer

If UCase$(UserList(userindex).Clase) = "MINERO" Then
    Call QuitarSta(userindex, EsfuerzoExcavarMinero)
Else
    Call QuitarSta(userindex, EsfuerzoExcavarGeneral)
End If

If UserList(userindex).Stats.UserSkills(eSkill.Mineria) <= 10 _
   And UserList(userindex).Stats.UserSkills(eSkill.Mineria) >= -1 Then
                    Suerte = 35
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Mineria) <= 20 _
   And UserList(userindex).Stats.UserSkills(eSkill.Mineria) >= 11 Then
                    Suerte = 30
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Mineria) <= 30 _
   And UserList(userindex).Stats.UserSkills(eSkill.Mineria) >= 21 Then
                    Suerte = 28
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Mineria) <= 40 _
   And UserList(userindex).Stats.UserSkills(eSkill.Mineria) >= 31 Then
                    Suerte = 24
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Mineria) <= 50 _
   And UserList(userindex).Stats.UserSkills(eSkill.Mineria) >= 41 Then
                    Suerte = 22
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Mineria) <= 60 _
   And UserList(userindex).Stats.UserSkills(eSkill.Mineria) >= 51 Then
                    Suerte = 20
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Mineria) <= 70 _
   And UserList(userindex).Stats.UserSkills(eSkill.Mineria) >= 61 Then
                    Suerte = 18
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Mineria) <= 80 _
   And UserList(userindex).Stats.UserSkills(eSkill.Mineria) >= 71 Then
                    Suerte = 15
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Mineria) <= 90 _
   And UserList(userindex).Stats.UserSkills(eSkill.Mineria) >= 81 Then
                    Suerte = 10
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Mineria) <= 100 _
   And UserList(userindex).Stats.UserSkills(eSkill.Mineria) >= 91 Then
                    Suerte = 7
End If
res = RandomNumber(1, Suerte)

If res <= 5 Then
    Dim MiObj As Obj
    Dim nPos As WorldPos
    
    If UserList(userindex).flags.TargetObj = 0 Then Exit Sub
    
    MiObj.ObjIndex = ObjData(UserList(userindex).flags.TargetObj).MineralIndex
    
    If UCase$(UserList(userindex).Clase) = "MINERO" Then
        MiObj.Amount = RandomNumber(1, 6)
    Else
        MiObj.Amount = 1
    End If
    
    If Not MeterItemEnInventario(userindex, MiObj) Then _
        Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
    
    Call SendData(SendTarget.toindex, userindex, 0, "||¡Has extraido algunos minerales!" & FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(userindex).flags.UltimoMensaje = 9 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||¡No has conseguido nada!" & FONTTYPE_INFO)
        UserList(userindex).flags.UltimoMensaje = 9
    End If
    '[/CDT]
End If

Call SubirSkill(userindex, Mineria)

UserList(userindex).Counters.Trabajando = UserList(userindex).Counters.Trabajando + 1

Exit Sub

errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub



Public Sub DoMeditar(ByVal userindex As Integer)

UserList(userindex).Counters.IdleCount = 0

Dim Suerte As Integer
Dim res As Integer
Dim Cant As Integer

'Barrin 3/10/03
'Esperamos a que se termine de concentrar
Dim TActual As Long
TActual = GetTickCount() And &H7FFFFFFF
If TActual - UserList(userindex).Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
    Exit Sub
End If

If UserList(userindex).Counters.bPuedeMeditar = False Then
    UserList(userindex).Counters.bPuedeMeditar = True
End If

If UserList(userindex).Stats.MinMAN >= UserList(userindex).Stats.MaxMan Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Has terminado de meditar." & FONTTYPE_INFO)
    Call SendData(SendTarget.toindex, userindex, 0, "PEDOP")
    Call SendData(SendTarget.toindex, userindex, 0, "SOUND")
    UserList(userindex).flags.Meditando = False
    UserList(userindex).Char.FX = 0
    UserList(userindex).Char.loops = 0
    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & 0 & "," & 0)
    Exit Sub
End If

If UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 10 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= -1 Then
                    Suerte = 35
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 20 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 11 Then
                    Suerte = 30
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 30 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 21 Then
                    Suerte = 28
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 40 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 31 Then
                    Suerte = 24
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 50 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 41 Then
                    Suerte = 22
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 60 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 51 Then
                    Suerte = 20
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 70 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 61 Then
                    Suerte = 18
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 80 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 71 Then
                    Suerte = 15
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 90 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 81 Then
                    Suerte = 10
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Meditar) <= 100 _
   And UserList(userindex).Stats.UserSkills(eSkill.Meditar) >= 91 Then
                    Suerte = 10
End If
res = RandomNumber(1, Suerte)

If res = 1 Then
    Cant = Porcentaje(UserList(userindex).Stats.MaxMan, 2)
    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN + Cant
    If UserList(userindex).Stats.MinMAN > UserList(userindex).Stats.MaxMan Then _
        UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MaxMan
    
    If Not UserList(userindex).flags.UltimoMensaje = 22 Then
        UserList(userindex).flags.UltimoMensaje = 22
    End If
    
    Call SendData(SendTarget.toindex, userindex, 0, "ASM" & UserList(userindex).Stats.MinMAN)
    Call SubirSkill(userindex, Meditar)
End If

End Sub



Public Sub Desarmar(ByVal userindex As Integer, ByVal VictimIndex As Integer)

Dim Suerte As Integer
Dim res As Integer

If UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 10 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= -1 Then
                    Suerte = 35
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 20 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 11 Then
                    Suerte = 30
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 30 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 21 Then
                    Suerte = 28
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 40 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 31 Then
                    Suerte = 24
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 50 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 41 Then
                    Suerte = 22
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 60 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 51 Then
                    Suerte = 20
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 70 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 61 Then
                    Suerte = 18
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 80 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 71 Then
                    Suerte = 15
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 90 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 81 Then
                    Suerte = 10
ElseIf UserList(userindex).Stats.UserSkills(eSkill.Wresterling) <= 100 _
   And UserList(userindex).Stats.UserSkills(eSkill.Wresterling) >= 91 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res <= 2 Then
        Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
        Call SendData(SendTarget.toindex, userindex, 0, "||Has logrado desarmar a tu oponente!" & FONTTYPE_FIGHT)
        If UserList(VictimIndex).Stats.ELV < 20 Then Call SendData(SendTarget.toindex, VictimIndex, 0, "||Tu oponente te ha desarmado!" & FONTTYPE_FIGHT)
    End If
End Sub

Public Sub DoEquita(ByVal userindex As Integer, ByRef Montura As ObjData, ByVal Slot As Integer)
 
Dim ModEqui As Long
 
ModEqui = ModEquitacion(UserList(userindex).Clase)
 
If UserList(userindex).Stats.UserSkills(Equitacion) / ModEqui < Montura.MinSkill Then
    Call SendData(toindex, userindex, 0, "||Para usar esta montura necesitas " & Montura.MinSkill * ModEqui & " puntos en equitación." & FONTTYPE_INFO)
    Exit Sub
End If
 
UserList(userindex).Invent.MonturaObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
UserList(userindex).Invent.MonturaSlot = Slot
 
If UserList(userindex).flags.Montando = 0 Then
    UserList(userindex).Char.Head = 0
    If UserList(userindex).flags.Muerto = 0 Then
        UserList(userindex).Char.Body = Montura.Ropaje
    Else
        UserList(userindex).Char.Body = iCuerpoMuerto
        UserList(userindex).Char.Head = iCabezaMuerto
    End If
    UserList(userindex).Char.Head = UserList(userindex).OrigChar.Head
    UserList(userindex).Char.ShieldAnim = NingunEscudo
    UserList(userindex).Char.WeaponAnim = NingunArma
    UserList(userindex).Char.CascoAnim = UserList(userindex).Char.CascoAnim
    UserList(userindex).flags.Montando = 1
Else
    UserList(userindex).flags.Montando = 0
    If UserList(userindex).flags.Muerto = 0 Then
        UserList(userindex).Char.Head = UserList(userindex).OrigChar.Head
        If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(userindex).Char.Body = ObjData(UserList(userindex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(userindex)
        End If
        If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then UserList(userindex).Char.ShieldAnim = ObjData(UserList(userindex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then UserList(userindex).Char.WeaponAnim = ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then UserList(userindex).Char.CascoAnim = ObjData(UserList(userindex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(userindex).Char.Body = iCuerpoMuerto
        UserList(userindex).Char.Head = iCabezaMuerto
        UserList(userindex).Char.ShieldAnim = NingunEscudo
        UserList(userindex).Char.WeaponAnim = NingunArma
        UserList(userindex).Char.CascoAnim = NingunCasco
    End If
End If
 
Call ChangeUserChar(ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
Call SendData(toindex, userindex, 0, "EQUIT")
 
End Sub
 
Function ModEquitacion(ByVal Clase As String) As Integer
 
Select Case UCase$(Clase)
    Case "CLERIGO"
        ModEquitacion = 1
    Case Else
        ModEquitacion = 1.5
End Select
 
End Function
