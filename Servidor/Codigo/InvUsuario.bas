Attribute VB_Name = "InvUsuario"
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

Public Function TieneObjetosRobables(ByVal userindex As Integer) As Boolean

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(userindex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And _
                ObjData(ObjIndex).OBJType <> eOBJType.otBarcos And ObjData(ObjIndex).Newbie = 0) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
   
    End If
Next i

End Function

Function ClasePuedeUsarItem(ByVal userindex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo manejador

'Call LogTarea("ClasePuedeUsarItem")

Dim flag As Boolean

If ObjData(ObjIndex).ClaseProhibida(1) <> "" Then
    
    Dim i As Integer
    For i = 1 To NUMCLASES
        If ObjData(ObjIndex).ClaseProhibida(i) = UCase$(UserList(userindex).Clase) Then
                ClasePuedeUsarItem = False
                Exit Function
        End If
    Next i
    
Else
    
    

End If

ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal userindex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(userindex).Invent.Object(j).ObjIndex > 0 Then
             
             If ObjData(UserList(userindex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(userindex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, userindex, j)
        
        End If
Next

'[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
'es transportado a su hogar de origen ;)
If UserList(userindex).pos.Map = 37 Then
    
    Dim DeDonde As WorldPos
    
    Select Case UCase$(UserList(userindex).Hogar)
        Case "LINDOS" 'Vamos a tener que ir por todo el desierto... uff!
            DeDonde = Lindos
        Case "Runek"
            DeDonde = Runek
        Case "BANDERBILL"
            DeDonde = Banderbill
        Case Else
            DeDonde = Helkat
    End Select
       
    Call WarpUserChar(userindex, DeDonde.Map, DeDonde.X, DeDonde.Y, True)

End If
'[/Barrin]

End Sub

Sub LimpiarInventario(ByVal userindex As Integer)


Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        UserList(userindex).Invent.Object(j).ObjIndex = 0
        UserList(userindex).Invent.Object(j).Amount = 0
        UserList(userindex).Invent.Object(j).Equipped = 0
        
Next

UserList(userindex).Invent.NroItems = 0

UserList(userindex).Invent.ArmourEqpObjIndex = 0
UserList(userindex).Invent.ArmourEqpSlot = 0

UserList(userindex).Invent.WeaponEqpObjIndex = 0
UserList(userindex).Invent.WeaponEqpSlot = 0

UserList(userindex).Invent.CascoEqpObjIndex = 0
UserList(userindex).Invent.CascoEqpSlot = 0

UserList(userindex).Invent.EscudoEqpObjIndex = 0
UserList(userindex).Invent.EscudoEqpSlot = 0

UserList(userindex).Invent.HerramientaEqpObjIndex = 0
UserList(userindex).Invent.HerramientaEqpSlot = 0

UserList(userindex).Invent.MunicionEqpObjIndex = 0
UserList(userindex).Invent.MunicionEqpSlot = 0

UserList(userindex).Invent.BarcoObjIndex = 0
UserList(userindex).Invent.BarcoSlot = 0

UserList(userindex).Invent.MonturaObjIndex = 0
UserList(userindex).Invent.MonturaSlot = 0

End Sub

Sub QuitarUserInvItem(ByVal userindex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)

Dim MiObj As Obj
'Desequipa
If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(userindex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(userindex, Slot)

'Quita un objeto
UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount - Cantidad
'¿Quedan mas?
If UserList(userindex).Invent.Object(Slot).Amount <= 0 Then
    UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
    UserList(userindex).Invent.Object(Slot).ObjIndex = 0
    UserList(userindex).Invent.Object(Slot).Amount = 0
End If
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal userindex As Integer, ByVal Slot As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(userindex).Invent.Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(userindex, Slot, UserList(userindex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(userindex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_INVENTORY_SLOTS

        'Actualiza el inventario
        If UserList(userindex).Invent.Object(LoopC).ObjIndex > 0 Then
            Call ChangeUserInv(userindex, LoopC, UserList(userindex).Invent.Object(LoopC))
        Else
            
            Call ChangeUserInv(userindex, LoopC, NullObj)
            
        End If

    Next LoopC

End If

End Sub

Sub DropObj(ByVal userindex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

If UserList(userindex).flags.Muerto = 0 Then
If UserList(userindex).Counters.TiraItem > 0 Then Call SendData(ToIndex, userindex, 0, "||Debes esperar " & UserList(userindex).Counters.TiraItem & " segundos para tirar otro item." & FONTTYPE_INFO): Exit Sub
            Else
           
End If
            
Dim Obj As Obj

If MapData(UserList(userindex).pos.Map, X, Y).OBJInfo.ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No hay espacio en el piso." & FONTTYPE_INFO)
Exit Sub
End If

If UserList(userindex).flags.Transformado = 1 Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes tirar items estando transformado, escribe /DESTRANSFORMAR." & FONTTYPE_INFO)
        Exit Sub
        End If


If num > 0 Then
  
  If num > UserList(userindex).Invent.Object(Slot).Amount Then num = UserList(userindex).Invent.Object(Slot).Amount
  
  'Check objeto en el suelo
  If MapData(UserList(userindex).pos.Map, X, Y).OBJInfo.ObjIndex = 0 Then
        If UserList(userindex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(userindex, Slot)
        Obj.ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
        
'        If ObjData(Obj.ObjIndex).Newbie = 1 And EsNewbie(UserIndex) Then
'            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes tirar el objeto." & FONTTYPE_INFO)
'            Exit Sub
'        End If
        
        If num + MapData(UserList(userindex).pos.Map, X, Y).OBJInfo.Amount > MAX_INVENTORY_OBJS Then
            num = MAX_INVENTORY_OBJS - MapData(UserList(userindex).pos.Map, X, Y).OBJInfo.Amount
        End If
        
        Obj.Amount = num
        
        Call MakeObj(SendTarget.ToMap, 0, Map, Obj, Map, X, Y)
If UserList(userindex).flags.Muerto = 0 Then
         UserList(userindex).Counters.TiraItem = RandomNumber(1, 4)
         Else
       
         End If
        Call QuitarUserInvItem(userindex, Slot, num)
        Call UpdateUserInv(False, userindex, Slot)
        
        If ObjData(Obj.ObjIndex).OBJType = eOBJType.otBarcos Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡ATENCION!! ¡ACABAS DE TIRAR TU BARCA!" & FONTTYPE_TALK)
        End If
        If ObjData(Obj.ObjIndex).Caos = 1 Or ObjData(Obj.ObjIndex).Real = 1 Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||¡ATENCION!! ¡¡ACABAS DE TIRAR TU ARMADURA FACCIONARIA!!" & FONTTYPE_TALK)
        End If
        
        If ObjData(Obj.ObjIndex).OBJType = eOBJType.otMonturas Then
            Call SendData(ToIndex, userindex, 0, "||¡¡ATENCION!! ¡ACABAS DE TIRAR TU MONTURA!" & FONTTYPE_TALK)
        End If
        
        If UserList(userindex).flags.Privilegios > PlayerType.User Then Call LogGM(UserList(userindex).name, "Tiro cantidad:" & num & " Objeto:" & ObjData(Obj.ObjIndex).name, False)
  Else
    Call SendData(SendTarget.ToIndex, userindex, 0, "||No hay espacio en el piso." & FONTTYPE_INFO)
  End If
    
End If

End Sub

Sub EraseObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, ByVal num As Integer, ByVal Map As Byte, ByVal X As Integer, ByVal Y As Integer)

MapData(Map, X, Y).OBJInfo.Amount = MapData(Map, X, Y).OBJInfo.Amount - num

If MapData(Map, X, Y).OBJInfo.Amount <= 0 Then
    MapData(Map, X, Y).OBJInfo.ObjIndex = 0
    MapData(Map, X, Y).OBJInfo.Amount = 0
    
    If sndRoute = SendTarget.ToMap Then
        Call SendToAreaByPos(Map, X, Y, "BO" & X & "," & Y)
   Else
        Call SendData(sndRoute, sndIndex, sndMap, "BO" & X & "," & Y)
    End If
End If

End Sub

Sub MakeObj(ByVal sndRoute As Byte, ByVal sndIndex As Integer, ByVal sndMap As Integer, Obj As Obj, Map As Integer, ByVal X As Integer, ByVal Y As Integer)

If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then

    If MapData(Map, X, Y).OBJInfo.ObjIndex = Obj.ObjIndex Then
        MapData(Map, X, Y).OBJInfo.Amount = MapData(Map, X, Y).OBJInfo.Amount + Obj.Amount
    Else
        MapData(Map, X, Y).OBJInfo = Obj
        
If sndRoute = SendTarget.ToMap Then
            Call ModAreas.SendToAreaByPos(Map, X, Y, "HO" & ObjData(Obj.ObjIndex).GrhIndex & "," & X & "," & Y & "," & ObjData(Obj.ObjIndex).name)
        Else
            Call SendData(sndRoute, sndIndex, sndMap, "HO" & ObjData(Obj.ObjIndex).GrhIndex & "," & X & "," & Y & "," & ObjData(Obj.ObjIndex).name)
        End If
    End If
End If

End Sub

Function MeterItemEnInventario(ByVal userindex As Integer, ByRef MiObj As Obj) As Boolean
On Error GoTo errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim X As Integer
Dim Y As Integer
Dim Slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(userindex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
         UserList(userindex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
    
'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(userindex).Invent.Object(Slot).ObjIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes cargar mas objetos." & FONTTYPE_FIGHT)
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(userindex).Invent.Object(Slot).Amount + MiObj.Amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(userindex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
   UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount + MiObj.Amount
Else
   UserList(userindex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, userindex, Slot)


Exit Function
errhandler:

End Function


Sub GetObj(ByVal userindex As Integer)

Dim Obj As ObjData
Dim MiObj As Obj

        If TieneObjetos(1062, 1, userindex) And MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).OBJInfo.ObjIndex = 1062 Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes agarrar un mapa del tesoro si ya tienes uno en el inventario." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).OBJInfo.ObjIndex = 1062 And MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).OBJInfo.Amount >= 2 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Hay mas de 1 mapa en el piso, no pueden ser agarrados." & FONTTYPE_INFO)
        Exit Sub
        End If

'¿Hay algun obj?
If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).OBJInfo.ObjIndex > 0 Then
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).OBJInfo.ObjIndex).Agarrable <> 1 Then
        Dim X As Integer
        Dim Y As Integer
        Dim Slot As Byte
        
        X = UserList(userindex).pos.X
        Y = UserList(userindex).pos.Y
        Obj = ObjData(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).OBJInfo.ObjIndex)
        MiObj.Amount = MapData(UserList(userindex).pos.Map, X, Y).OBJInfo.Amount
        MiObj.ObjIndex = MapData(UserList(userindex).pos.Map, X, Y).OBJInfo.ObjIndex
        
        If MapData(UserList(userindex).pos.Map, X, Y).OBJInfo.ObjIndex = 1073 Then
        UserList(userindex).Char.Aura = 20248
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
        Call SendData(SendTarget.ToTD, 0, 0, "||" & UserList(userindex).name & " lleva la pelota!." & "~255~0~0~1~0")
        End If
        
        If Not MeterItemEnInventario(userindex, MiObj) Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedo cargar mas objetos." & FONTTYPE_INFO)
        Else
            'Quitamos el objeto
            Call EraseObj(SendTarget.ToMap, 0, UserList(userindex).pos.Map, MapData(UserList(userindex).pos.Map, X, Y).OBJInfo.Amount, UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y)
            If UserList(userindex).flags.Privilegios > PlayerType.User Then Call LogGM(UserList(userindex).name, "Agarro:" & MiObj.Amount & " Objeto:" & ObjData(MiObj.ObjIndex).name, False)
        End If
        
    End If
Else
    Call SendData(SendTarget.ToIndex, userindex, 0, "||No hay nada aqui." & FONTTYPE_INFO)
End If

End Sub

Sub Desequipar(ByVal userindex As Integer, ByVal Slot As Byte)
'Desequipa el item slot del inventario
Dim Obj As ObjData


If (Slot < LBound(UserList(userindex).Invent.Object)) Or (Slot > UBound(UserList(userindex).Invent.Object)) Then
    Exit Sub
ElseIf UserList(userindex).Invent.Object(Slot).ObjIndex = 0 Then
    Exit Sub
End If

Obj = ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex)

Select Case Obj.OBJType
    Case eOBJType.otWeapon
    If UserList(userindex).flags.Montando = 1 Then Exit Sub
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.WeaponEqpObjIndex = 0
        UserList(userindex).Invent.WeaponEqpSlot = 0
        If Not UserList(userindex).flags.Mimetizado = 1 Then
            UserList(userindex).Char.WeaponAnim = NingunArma
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
        End If
    
        UserList(userindex).Char.Aura = 0
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
    
    Case eOBJType.otFlechas
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.MunicionEqpObjIndex = 0
        UserList(userindex).Invent.MunicionEqpSlot = 0
    
    Case eOBJType.otHerramientas
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.HerramientaEqpObjIndex = 0
        UserList(userindex).Invent.HerramientaEqpSlot = 0
    
    Case eOBJType.otArmadura
    If UserList(userindex).flags.Montando = 1 Then Exit Sub
    If UserList(userindex).flags.Montando = 1 Then Exit Sub
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.ArmourEqpObjIndex = 0
        UserList(userindex).Invent.ArmourEqpSlot = 0
        Call DarCuerpoDesnudo(userindex, UserList(userindex).flags.Mimetizado = 1)
        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
        UserList(userindex).Char.Aura = 0
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
            
    Case eOBJType.otCASCO
    If UserList(userindex).flags.Montando = 1 Then Exit Sub
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.CascoEqpObjIndex = 0
        UserList(userindex).Invent.CascoEqpSlot = 0
        UserList(userindex).Char.Aura = 0
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
        If Not UserList(userindex).flags.Mimetizado = 1 Then
            UserList(userindex).Char.CascoAnim = NingunCasco
            Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
        End If
    
    Case eOBJType.otESCUDO
    If UserList(userindex).flags.Montando = 1 Then Exit Sub
        UserList(userindex).Invent.Object(Slot).Equipped = 0
        UserList(userindex).Invent.EscudoEqpObjIndex = 0
        UserList(userindex).Invent.EscudoEqpSlot = 0
        UserList(userindex).Char.Aura = 0
        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
        If Not UserList(userindex).flags.Mimetizado = 1 Then
        UserList(userindex).Char.ShieldAnim = NingunEscudo
        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
        End If
End Select

Call SendUserStatsBox(userindex)
Call UpdateUserInv(False, userindex, Slot)
Call SendUserHitBox(userindex)

End Sub

Function SexoPuedeUsarItem(ByVal userindex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo errhandler

If ObjData(ObjIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(userindex).Genero) <> "HOMBRE"
ElseIf ObjData(ObjIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UCase$(UserList(userindex).Genero) <> "MUJER"
Else
    SexoPuedeUsarItem = True
End If

Exit Function
errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal userindex As Integer, ByVal ObjIndex As Integer) As Boolean

If ObjData(ObjIndex).Real = 1 Then
    If Not Criminal(userindex) Then
        FaccionPuedeUsarItem = (UserList(userindex).Faccion.ArmadaReal = 1)
    Else
        FaccionPuedeUsarItem = False
    End If
ElseIf ObjData(ObjIndex).Caos = 1 Then
    If Criminal(userindex) Then
        FaccionPuedeUsarItem = (UserList(userindex).Faccion.FuerzasCaos = 1)
    Else
        FaccionPuedeUsarItem = False
    End If
Else
    FaccionPuedeUsarItem = True
End If

End Function

Sub EquiparInvItem(ByVal userindex As Integer, ByVal Slot As Byte)
On Error GoTo errhandler

If HayTD = True And TieneObjetos(1073, 1, userindex) Then Exit Sub
 
'Equipa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer
 
ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
Obj = ObjData(ObjIndex)
 
If Obj.Newbie = 1 And Not EsNewbie(userindex) Then
     Call SendData(SendTarget.ToIndex, userindex, 0, "||Solo los newbies pueden usar este objeto." & FONTTYPE_INFO)
     Exit Sub
End If

If Obj.SoloVIP = 1 And UserList(userindex).flags.Privilegios = PlayerType.User Then
Call SendData(SendTarget.ToIndex, userindex, 0, "||¡Este item solo puede ser utilizado por un usuario VIP!." & FONTTYPE_INFO)
Exit Sub
End If
       
Select Case Obj.OBJType
    Case eOBJType.otWeapon
    If UserList(userindex).flags.Montando = 1 Then Exit Sub
       If ClasePuedeUsarItem(userindex, ObjIndex) And _
          FaccionPuedeUsarItem(userindex, ObjIndex) Or UserList(userindex).flags.Privilegios >= PlayerType.Dios Then
                'Si esta equipado lo quita
                If UserList(userindex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(userindex, Slot)
                    'Animacion por defecto
                    If UserList(userindex).flags.Mimetizado = 1 Then
                        UserList(userindex).CharMimetizado.WeaponAnim = NingunArma
                    Else
                        UserList(userindex).Char.WeaponAnim = NingunArma
                        Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                    End If
                    Exit Sub
                End If
               
                'Quitamos el elemento anterior
                If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)
                End If
       
                UserList(userindex).Invent.Object(Slot).Equipped = 1
                UserList(userindex).Invent.WeaponEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
                UserList(userindex).Invent.WeaponEqpSlot = Slot
               
                'Sonido
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_SACARARMA)
                
                UserList(userindex).Char.Aura = Obj.Aura
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)

       
                If UserList(userindex).flags.Mimetizado = 1 Then
                    UserList(userindex).CharMimetizado.WeaponAnim = Obj.WeaponAnim
                Else
                    UserList(userindex).Char.WeaponAnim = Obj.WeaponAnim
                    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                End If
       Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
       End If
   
    Case eOBJType.otHerramientas
       If ClasePuedeUsarItem(userindex, ObjIndex) And _
          FaccionPuedeUsarItem(userindex, ObjIndex) Or UserList(userindex).flags.Privilegios >= PlayerType.Dios Then
                'Si esta equipado lo quita
                If UserList(userindex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(userindex, Slot)
                    Exit Sub
                End If
               
                'Quitamos el elemento anterior
                If UserList(userindex).Invent.HerramientaEqpObjIndex > 0 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.HerramientaEqpSlot)
                End If
       
                UserList(userindex).Invent.Object(Slot).Equipped = 1
                UserList(userindex).Invent.HerramientaEqpObjIndex = ObjIndex
                UserList(userindex).Invent.HerramientaEqpSlot = Slot
               
       Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
       End If
   
    Case eOBJType.otFlechas
       If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) And _
          FaccionPuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) Or UserList(userindex).flags.Privilegios >= PlayerType.Dios Then
               
                'Si esta equipado lo quita
                If UserList(userindex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(userindex, Slot)
                    Exit Sub
                End If
               
                'Quitamos el elemento anterior
                If UserList(userindex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(userindex, UserList(userindex).Invent.MunicionEqpSlot)
                End If
       
                UserList(userindex).Invent.Object(Slot).Equipped = 1
                UserList(userindex).Invent.MunicionEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
                UserList(userindex).Invent.MunicionEqpSlot = Slot
               
       Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
       End If
   
    Case eOBJType.otArmadura
        If UserList(userindex).flags.Montando = 1 Then Exit Sub
        If UserList(userindex).flags.Navegando = 1 Then Exit Sub
        
        'Nos aseguramos que puede usarla
        If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) And _
           SexoPuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) And _
           CheckRazaUsaRopa(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) And _
           FaccionPuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) _
           Or UserList(userindex).flags.Privilegios >= PlayerType.Dios Then
           
           'Si esta equipado lo quita
            If UserList(userindex).Invent.Object(Slot).Equipped Then
                Call Desequipar(userindex, Slot)
                Call DarCuerpoDesnudo(userindex, UserList(userindex).flags.Mimetizado = 1)
                If Not UserList(userindex).flags.Mimetizado = 1 Then
                    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                End If
                Exit Sub
            End If
            
            'Quita el anterior
            If UserList(userindex).Invent.ArmourEqpObjIndex > 0 Then
                Call Desequipar(userindex, UserList(userindex).Invent.ArmourEqpSlot)
            End If
            
            'Lo equipa
            UserList(userindex).Invent.Object(Slot).Equipped = 1
            UserList(userindex).Invent.ArmourEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
            UserList(userindex).Invent.ArmourEqpSlot = Slot
        
            If UserList(userindex).flags.Mimetizado = 1 Then
                UserList(userindex).CharMimetizado.Body = Obj.Ropaje
                                UserList(userindex).Char.Aura = Obj.Aura
Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "AAU")
Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "AUR" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
            Else
                UserList(userindex).Char.Body = Obj.Ropaje
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                                UserList(userindex).Char.Aura = Obj.Aura
Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "AAU")
Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "AUR" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
            End If
            UserList(userindex).flags.Desnudo = 0
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Tu clase,genero o raza no puede usar este objeto." & FONTTYPE_INFO)
        End If
   
    Case eOBJType.otCASCO
    If UserList(userindex).flags.Montando = 1 Then Exit Sub
        If UserList(userindex).flags.Navegando = 1 Then Exit Sub
        If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) _
        Or UserList(userindex).flags.Privilegios >= PlayerType.Dios Then
            'Si esta equipado lo quita
            If UserList(userindex).Invent.Object(Slot).Equipped Then
                Call Desequipar(userindex, Slot)
                If UserList(userindex).flags.Mimetizado = 1 Then
                    UserList(userindex).CharMimetizado.CascoAnim = NingunCasco
                Else
                    UserList(userindex).Char.CascoAnim = NingunCasco
                    Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                End If
                Exit Sub
            End If
   
            'Quita el anterior
            If UserList(userindex).Invent.CascoEqpObjIndex > 0 Then
                Call Desequipar(userindex, UserList(userindex).Invent.CascoEqpSlot)
            End If
   
            'Lo equipa
           
            UserList(userindex).Invent.Object(Slot).Equipped = 1
            UserList(userindex).Invent.CascoEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
            UserList(userindex).Invent.CascoEqpSlot = Slot
            UserList(userindex).Char.Aura = Obj.Aura
Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "AAU")
Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "AUR" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
            If UserList(userindex).flags.Mimetizado = 1 Then
                UserList(userindex).CharMimetizado.CascoAnim = Obj.CascoAnim
            Else
                UserList(userindex).Char.CascoAnim = Obj.CascoAnim
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
            End If
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
        End If
   
    Case eOBJType.otESCUDO
    If UserList(userindex).flags.Montando = 1 Then Exit Sub
        If UserList(userindex).flags.Navegando = 1 Then Exit Sub
         If ClasePuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) And _
             FaccionPuedeUsarItem(userindex, UserList(userindex).Invent.Object(Slot).ObjIndex) _
                Or UserList(userindex).flags.Privilegios >= PlayerType.Dios Then
 
             'Si esta equipado lo quita
             If UserList(userindex).Invent.Object(Slot).Equipped Then
                 Call Desequipar(userindex, Slot)
                 If UserList(userindex).flags.Mimetizado = 1 Then
                     UserList(userindex).CharMimetizado.ShieldAnim = NingunEscudo
                 Else
                     UserList(userindex).Char.ShieldAnim = NingunEscudo
                     Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
                 End If
                 Exit Sub
             End If
     
             'Quita el anterior
             If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
                 Call Desequipar(userindex, UserList(userindex).Invent.EscudoEqpSlot)
             End If
     
             'Lo equipa
             
             UserList(userindex).Invent.Object(Slot).Equipped = 1
             UserList(userindex).Invent.EscudoEqpObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
             UserList(userindex).Invent.EscudoEqpSlot = Slot
             
             UserList(userindex).Char.Aura = Obj.Aura
Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "AAU")
Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "AUR" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
             
             If UserList(userindex).flags.Mimetizado = 1 Then
                 UserList(userindex).CharMimetizado.ShieldAnim = Obj.ShieldAnim
             Else
                 UserList(userindex).Char.ShieldAnim = Obj.ShieldAnim
                 
                 Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
             End If
         Else
             Call SendData(SendTarget.ToIndex, userindex, 0, "||Tu clase no puede usar este objeto." & FONTTYPE_INFO)
         End If
End Select
 
         If UserList(userindex).flags.EnTD = True And UserList(userindex).flags.TeamTD = 1 Then
        UserList(userindex).Char.Body = 320
        End If
   
        If UserList(userindex).flags.EnTD = True And UserList(userindex).flags.TeamTD = 2 Then
        UserList(userindex).Char.Body = 322
        End If
 
'Actualiza ~ Feer
Call SendUserStatsBox(userindex)
Call UpdateUserInv(False, userindex, Slot)
Call SendUserHitBox(userindex)
 
Exit Sub
errhandler:
Call LogError("EquiparInvItem Slot:" & Slot)
End Sub

Private Function CheckRazaUsaRopa(ByVal userindex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo errhandler

'Verifica si la raza puede usar la ropa
If UserList(userindex).Raza = "Humano" Or _
   UserList(userindex).Raza = "Elfo" Or _
   UserList(userindex).Raza = "Elfo Oscuro" Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
End If


Exit Function
errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal userindex As Integer, ByVal Slot As Byte)

'Usa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As Obj

If UserList(userindex).Invent.Object(Slot).Amount = 0 Then Exit Sub

Obj = ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex)

If Obj.Newbie = 1 And Not EsNewbie(userindex) Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Solo los newbies pueden usar estos objetos." & FONTTYPE_INFO)
    Exit Sub
End If

If Obj.OBJType = eOBJType.otWeapon Then
If UserList(userindex).Lac.LUsar.Puedo = False Then Exit Sub
    If Obj.proyectil = 1 Then
        'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
        If Not IntervaloPermiteUsarArcos(userindex, False) Then Exit Sub
    Else
        'dagas
        If Not IntervaloPermiteUsar(userindex) Then Exit Sub
    End If
Else
    If Not IntervaloPermiteUsar(userindex) Then Exit Sub
End If

ObjIndex = UserList(userindex).Invent.Object(Slot).ObjIndex
UserList(userindex).flags.TargetObjInvIndex = ObjIndex
UserList(userindex).flags.TargetObjInvSlot = Slot

Select Case Obj.OBJType

    Case eOBJType.otGuita
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(userindex).Invent.Object(Slot).Amount = 0
        UserList(userindex).Invent.Object(Slot).ObjIndex = 0
        UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems - 1
        
        Call UpdateUserInv(False, userindex, Slot)
        Call SendUserStatsBox(userindex)
        
    Case eOBJType.otWeapon
        If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
        End If
        
        If ObjData(ObjIndex).proyectil = 1 Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "T01" & Proyectiles)
        Else
            If UserList(userindex).flags.TargetObj = 0 Then Exit Sub
            
            '¿El target-objeto es leña?
            If UserList(userindex).flags.TargetObj = Leña Then
                If UserList(userindex).Invent.Object(Slot).ObjIndex = DAGA Then
                    Call TratarDeHacerFogata(UserList(userindex).flags.TargetObjMap, _
                         UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY, userindex)
                End If
            End If
        End If
    
    Case eOBJType.otPociones
    If UserList(userindex).Lac.LPociones.Puedo = False Then Exit Sub '[Loopzer]
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If


        
        UserList(userindex).flags.TomoPocion = True
        UserList(userindex).flags.TipoPocion = Obj.TipoPocion
                
        Select Case UserList(userindex).flags.TipoPocion
        
            Case 1 'Modif la agilidad
                UserList(userindex).flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then _
                    UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                If UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) > 2 * UserList(userindex).Stats.UserAtributosBackUP(Agilidad) Then UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = 2 * UserList(userindex).Stats.UserAtributosBackUP(Agilidad)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_BEBER)
                Call SendData(ToIndex, userindex, UserList(userindex).pos.Map, "PX" & UserList(userindex).Stats.UserAtributos(Agilidad))
        
            Case 2 'Modif la fuerza
                UserList(userindex).flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
                    UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                If UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) > 2 * UserList(userindex).Stats.UserAtributosBackUP(Fuerza) Then UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = 2 * UserList(userindex).Stats.UserAtributosBackUP(Fuerza)
                
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_BEBER)
                Call SendData(ToIndex, userindex, UserList(userindex).pos.Map, "PZ" & UserList(userindex).Stats.UserAtributos(Fuerza))

                
            Case 3 'Pocion roja, restaura HP
                'Usa el item
                UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(userindex).Stats.MinHP > UserList(userindex).Stats.MaxHP Then _
                    UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_BEBER)
            
            Case 4 'Pocion azul, restaura MANA
                'Usa el item
                UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN + Porcentaje(UserList(userindex).Stats.MaxMan, 5)
                If UserList(userindex).Stats.MinMAN > UserList(userindex).Stats.MaxMan Then _
                    UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MaxMan
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_BEBER)
                
            Case 5 ' Pocion violeta
                If UserList(userindex).flags.Envenenado = 1 Then
                    UserList(userindex).flags.Envenenado = 0
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Te has curado del envenenamiento." & FONTTYPE_INFO)
                End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_BEBER)
            Case 6  ' Pocion Negra
                If UserList(userindex).flags.Privilegios = PlayerType.User Then
                    Call QuitarUserInvItem(userindex, Slot, 1)
                    Call UserDie(userindex)
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||Sientes un gran mareo y pierdes el conocimiento." & FONTTYPE_FIGHT)
                End If
Case 7 ' Poción de Energia, cura energia.
    UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MinSta + Porcentaje(UserList(userindex).Stats.MaxSta, 5)
        If UserList(userindex).Stats.MinSta > UserList(userindex).Stats.MaxSta Then _
    UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MaxSta
    'Quitamos del inv el item
                Call QuitarUserInvItem(userindex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_BEBER)
       End Select
       Call SendUserStatsBox(userindex)
       Call UpdateUserInv(False, userindex, Slot)
    
    Case eOBJType.otLlaves
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(userindex).flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(UserList(userindex).flags.TargetObj)
        '¿El objeto clickeado es una puerta?
        If TargObj.OBJType = eOBJType.otPuertas Then
            '¿Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '¿Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.clave = Obj.clave Then
         
                        MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerrada
                        UserList(userindex).flags.TargetObj = MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Has abierto la puerta." & FONTTYPE_INFO)
                        Exit Sub
                     Else
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||La llave no sirve." & FONTTYPE_INFO)
                        Exit Sub
                     End If
                  Else
                     If TargObj.clave = Obj.clave Then
                        MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex _
                        = ObjData(MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex).IndexCerradaLlave
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||Has cerrado con llave la puerta." & FONTTYPE_INFO)
                        UserList(userindex).flags.TargetObj = MapData(UserList(userindex).flags.TargetObjMap, UserList(userindex).flags.TargetObjX, UserList(userindex).flags.TargetObjY).OBJInfo.ObjIndex
                        Exit Sub
                     Else
                        Call SendData(SendTarget.ToIndex, userindex, 0, "||La llave no sirve." & FONTTYPE_INFO)
                        Exit Sub
                     End If
                  End If
            Else
                  Call SendData(SendTarget.ToIndex, userindex, 0, "||No esta cerrada." & FONTTYPE_INFO)
                  Exit Sub
            End If
            
        End If
               
    Case eOBJType.otBolsaTesoro
                If UserList(userindex).flags.Muerto = 1 Then
                   Call SendData(ToIndex, userindex, 0, "||¡Estas Muerto!." & FONTTYPE_INFO)
                Exit Sub
                End If
                Dim Tes As Integer
                Tes = RandomNumber(1, 2)
                Dim Obj1 As Obj
                Dim Obj2 As Obj
                Dim Obj3 As Obj
                Dim Obj4 As Obj
               
                Obj1.Amount = 1
                Obj1.ObjIndex = 1053 'escudo tortuga
               
                Obj2.Amount = 1
                Obj2.ObjIndex = 665 'arco cazador
               
                Obj3.Amount = 500
                Obj3.ObjIndex = 1053 'flecha
               
                Obj4.Amount = 1
                Obj4.ObjIndex = 478 'arco simple
                Call QuitarUserInvItem(userindex, Slot, 1)
                If Tes = 1 Then
                    If Not MeterItemEnInventario(userindex, Obj1) Then
                        Call TirarItemAlPiso(UserList(userindex).pos, Obj1)
                    End If
                    Call SendData(ToIndex, userindex, 0, "||¡Has recibido una Espada Mata Dragones como recompensa!." & FONTTYPE_INFO)
                End If
                'Quitamos el item
                Call UpdateUserInv(False, userindex, Slot)
        
        Case eOBJType.otBebidas
        If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas muerto, Solo podes usar items cuando estas vivo." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        Case 29
   
        If MapInfo(UserList(userindex).pos.Map).Pk = False Then
            SendData SendTarget.ToIndex, userindex, 0, "||No podés activar la gema acá." & "~255~0~0~1~0"
            Exit Sub
        End If
       
        If UserList(userindex).flags.Muerto = 1 Then
            SendData SendTarget.ToIndex, userindex, 0, "||Estas muerto, Solo podés usar items cuando estas vivo." & FONTTYPE_SERVER
            Exit Sub
        End If
       
        Dim GemaName As String
       
        Select Case Obj.name
            Case "Gema Roja"
                GemaName = "Roja"
            Case "Gema Naranja"
                GemaName = "Naranja"
            Case "Gema Verde"
                GemaName = "Verde"
            Case "Gema Azul"
                GemaName = "Azul"
            Case "Gema Plateada"
                GemaName = "Plateada"
            Case "Gema Celeste"
                GemaName = "Celeste"
            Case "Gema Violeta"
                GemaName = "Violeta"
            Case "Gema Lila"
                GemaName = "Lila"
        End Select
       
        If UserList(userindex).flags.ActivoGema = 1 Then
        With UserList(userindex).flags
            .ActivoGema = 0
            .GemaActivada = ""
            .ActivoGema = 1
            .GemaActivada = GemaName
            .TimeGema = 45
        End With
            SendData SendTarget.ToIndex, userindex, 0, "||El efecto de la Gema ha terminado." & "~255~0~0~1~0"
            SendData SendTarget.ToIndex, userindex, 0, "||Obtuviste el poder de la " & Obj.name & "." & FONTTYPE_GANAR
        Else
        With UserList(userindex).flags
            .ActivoGema = 1
            .GemaActivada = GemaName
            .TimeGema = 45
        End With
        SendData SendTarget.ToIndex, userindex, 0, "||Obtuviste el poder de la " & Obj.name & "." & FONTTYPE_GANAR
        End If
    
        Case eOBJType.otBotellaVacia
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If Not HayAgua(UserList(userindex).pos.Map, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY) Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||No hay agua allí." & FONTTYPE_INFO)
                Exit Sub
            End If
            MiObj.Amount = 1
            MiObj.ObjIndex = ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).IndexAbierta
            Call QuitarUserInvItem(userindex, Slot, 1)
            If Not MeterItemEnInventario(userindex, MiObj) Then
                Call TirarItemAlPiso(UserList(userindex).pos, MiObj)
            End If
            
            Call UpdateUserInv(False, userindex, Slot)
            
        Case eOBJType.otHerramientas
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If Not UserList(userindex).Stats.MinSta > 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Estas muy cansado" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(userindex).Invent.Object(Slot).Equipped = 0 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Antes de usar la herramienta deberias equipartela." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            UserList(userindex).Reputacion.PlebeRep = UserList(userindex).Reputacion.PlebeRep + vlProleta
            If UserList(userindex).Reputacion.PlebeRep > MAXREP Then _
                UserList(userindex).Reputacion.PlebeRep = MAXREP
            
            Select Case ObjIndex
                Case CAÑA_PESCA, RED_PESCA
                    Call SendData(SendTarget.ToIndex, userindex, 0, "T01" & Pesca)
                Case HACHA_LEÑADOR
                    Call SendData(SendTarget.ToIndex, userindex, 0, "T01" & Talar)
                Case PIQUETE_MINERO
                    Call SendData(SendTarget.ToIndex, userindex, 0, "T01" & Mineria)
                Case MARTILLO_HERRERO
                    Call SendData(SendTarget.ToIndex, userindex, 0, "T01" & Herreria)
                Case SERRUCHO_CARPINTERO
                    Call EnivarObjConstruibles(userindex)
                    Call SendData(SendTarget.ToIndex, userindex, 0, "SFC")

            End Select
        
        Case eOBJType.otPergaminos
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            
If Hechizos(ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).HechizoIndex).ExclusivoClase = UCase$(UserList(userindex).Clase) Or _
Len(Hechizos(ObjData(UserList(userindex).Invent.Object(Slot).ObjIndex).HechizoIndex).ExclusivoClase) = 0 Then
Call AgregarHechizo(userindex, Slot)
Call UpdateUserInv(False, userindex, Slot)
Else
Call SendData(SendTarget.ToIndex, userindex, 0, "||Tú clase no puede aprender este hechizo." & FONTTYPE_INFO)
End If
       
       Case eOBJType.otMinerales
           If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
           End If
           Call SendData(SendTarget.ToIndex, userindex, 0, "T01" & FundirMetal)
       
       Case eOBJType.otInstrumentos
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & Obj.Snd1)
        
        Case eOBJType.otMonturas
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(ToIndex, userindex, 0, "||¡¡Estas muerto!! Los muertos no dominan los animales. " & FONTTYPE_INFO)
                Exit Sub
            End If
            Call DoEquita(userindex, Obj, Slot)
       
       Case eOBJType.otBarcos
    'Verifica si esta aproximado al agua antes de permitirle navegar
        If UserList(userindex).Stats.ELV < 25 Then
            If UCase$(UserList(userindex).Clase) <> "PESCADOR" And UCase$(UserList(userindex).Clase) <> "PIRATA" Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||Para recorrer los mares debes ser nivel 25 o superior." & FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        If UserList(userindex).flags.Montando = 1 Then
               Call SendData(SendTarget.ToIndex, userindex, 0, "EQUIT")
                    Call SendData(SendTarget.ToIndex, userindex, 0, "||No podes navegar estando montado." & FONTTYPE_INFO)
        Exit Sub
        End If
        
                If UserList(userindex).flags.Transformado = 1 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes navegar estando transformado, escribe /DESTRANSFORMAR" & FONTTYPE_INFO)
        Exit Sub
        End If
        
        If ((LegalPos(UserList(userindex).pos.Map, UserList(userindex).pos.X - 1, UserList(userindex).pos.Y, True) Or _
            LegalPos(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y - 1, True) Or _
            LegalPos(UserList(userindex).pos.Map, UserList(userindex).pos.X + 1, UserList(userindex).pos.Y, True) Or _
            LegalPos(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y + 1, True)) And _
            UserList(userindex).flags.Navegando = 0) _
            Or UserList(userindex).flags.Navegando = 1 Then
           Call DoNavega(userindex, Obj, Slot)
        Else
            Call SendData(SendTarget.ToIndex, userindex, 0, "||¡Debes aproximarte al agua para usar el barco!" & FONTTYPE_INFO)
        End If
           
End Select

'Actualiza
'Call SendUserStatsBox(UserIndex)
'Call UpdateUserInv(False, UserIndex, Slot)

End Sub

Sub EnivarArmasConstruibles(ByVal userindex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ArmasHerrero)
    If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(userindex).Stats.UserSkills(eSkill.Herreria) \ ModHerreriA(UserList(userindex).Clase) Then
        If ObjData(ArmasHerrero(i)).OBJType = eOBJType.otWeapon Then
        '[DnG!]
            cad$ = cad$ & ObjData(ArmasHerrero(i)).name & " (" & ObjData(ArmasHerrero(i)).LingH & "-" & ObjData(ArmasHerrero(i)).LingP & "-" & ObjData(ArmasHerrero(i)).LingO & ")" & "," & ArmasHerrero(i) & ","
        '[/DnG!]
        Else
            cad$ = cad$ & ObjData(ArmasHerrero(i)).name & "," & ArmasHerrero(i) & ","
        End If
    End If
Next i

Call SendData(SendTarget.ToIndex, userindex, 0, "LAH" & cad$)

End Sub
 
Sub EnivarObjConstruibles(ByVal userindex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ObjCarpintero)
    If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(userindex).Stats.UserSkills(eSkill.Carpinteria) / ModCarpinteria(UserList(userindex).Clase) Then _
        cad$ = cad$ & ObjData(ObjCarpintero(i)).name & " (" & ObjData(ObjCarpintero(i)).Madera & ")" & "," & ObjCarpintero(i) & ","
Next i

Call SendData(SendTarget.ToIndex, userindex, 0, "OBR" & cad$)

End Sub

Sub EnivarArmadurasConstruibles(ByVal userindex As Integer)

Dim i As Integer, cad$

For i = 1 To UBound(ArmadurasHerrero)
    If ObjData(ArmadurasHerrero(i)).SkHerreria <= UserList(userindex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(userindex).Clase) Then
        '[DnG!]
        cad$ = cad$ & ObjData(ArmadurasHerrero(i)).name & " (" & ObjData(ArmadurasHerrero(i)).LingH & "-" & ObjData(ArmadurasHerrero(i)).LingP & "-" & ObjData(ArmadurasHerrero(i)).LingO & ")" & "," & ArmadurasHerrero(i) & ","
        '[/DnG!]
    End If
Next i

Call SendData(SendTarget.ToIndex, userindex, 0, "LAR" & cad$)

End Sub


                   

Sub TirarTodo(ByVal userindex As Integer)
On Error Resume Next

If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = 6 Then Exit Sub

Call TirarTodosLosItems(userindex)

End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean

ItemSeCae = (ObjData(index).Real <> 1 Or ObjData(index).NoSeCae = 0) And _
            (ObjData(index).Caos <> 1 Or ObjData(index).NoSeCae = 0) And _
            ObjData(index).OBJType <> eOBJType.otLlaves And _
            ObjData(index).OBJType <> eOBJType.otBarcos And _
            ObjData(index).OBJType <> eOBJType.otMonturas And _
            ObjData(index).NoSeCae = 0


End Function

Sub TirarTodosLosItems(ByVal userindex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = 6 Then Exit Sub
 
For i = 1 To MAX_INVENTORY_SLOTS
If UserList(userindex).Invent.Object(i).ObjIndex = SacriIndex Then
If DropSacri = 0 Then
NuevaPos.X = 0: NuevaPos.Y = 0
MiObj.Amount = UserList(userindex).Invent.Object(i).Amount: MiObj.ObjIndex = SacriIndex
Call Tilelibre(UserList(userindex).pos, NuevaPos, MiObj)
If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call DropObj(userindex, i, 1, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
Else
Call QuitarUserInvItem(userindex, i, 1)
Call UpdateUserInv(False, userindex, i)
End If
Exit Sub
End If
Next i
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(userindex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
             If ItemSeCae(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                
                'Creo el Obj
                MiObj.Amount = UserList(userindex).Invent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex
                
                Tilelibre UserList(userindex).pos, NuevaPos, MiObj
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
             End If
        End If
    Next i
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal userindex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer
If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = 6 Then Exit Sub
 
For i = 1 To MAX_INVENTORY_SLOTS
If UserList(userindex).Invent.Object(i).ObjIndex = SacriIndex Then
If DropSacri = 0 Then
NuevaPos.X = 0: NuevaPos.Y = 0
MiObj.Amount = UserList(userindex).Invent.Object(i).Amount: MiObj.ObjIndex = SacriIndex
Call Tilelibre(UserList(userindex).pos, NuevaPos, MiObj)
If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call DropObj(userindex, i, 1, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
Else
Call QuitarUserInvItem(userindex, i, 1)
Call UpdateUserInv(False, userindex, i)
End If
Exit Sub
End If
Next i

If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = 6 Then Exit Sub

For i = 1 To MAX_INVENTORY_SLOTS
    ItemIndex = UserList(userindex).Invent.Object(i).ObjIndex
    If ItemIndex > 0 Then
        If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
            NuevaPos.X = 0
            NuevaPos.Y = 0
            
            'Creo MiObj
            MiObj.Amount = UserList(userindex).Invent.Object(i).ObjIndex
            MiObj.ObjIndex = ItemIndex
            
            Tilelibre UserList(userindex).pos, NuevaPos, MiObj
            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).OBJInfo.ObjIndex = 0 Then Call DropObj(userindex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
            End If
        End If
    End If
Next i

End Sub
