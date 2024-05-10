Attribute VB_Name = "Comercio"
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

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%          MODULO DE COMERCIO NPC-USER              %%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


Function UserCompraObj(ByVal userindex As Integer, ByVal ObjIndex As Integer, ByVal NpcIndex As Integer, ByVal Cantidad As Integer) As Boolean
On Error GoTo errorh
        
    Dim Slot As Integer
    Dim obji As Integer
    Dim Encontre As Boolean
    
    UserCompraObj = False
    
    If (Npclist(UserList(userindex).flags.TargetNPC).Invent.Object(ObjIndex).Amount <= 0) Then Exit Function
    
    obji = Npclist(UserList(userindex).flags.TargetNPC).Invent.Object(ObjIndex).ObjIndex
    
    If ObjData(obji).OBJType = eOBJType.otMapaTesoro And Cantidad >= 2 Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Solo puedes llevar 1 mapa del tesoro a la ves." & FONTTYPE_INFO)
    Exit Function
    End If
    
    If ObjData(obji).OBJType = eOBJType.otMapaTesoro And TieneObjetos(1062, 1, userindex) Then
    Call SendData(SendTarget.ToIndex, userindex, 0, "||Solo puedes llevar 1 mapa del tesoro a la ves." & FONTTYPE_INFO)
    Exit Function
    End If
    
    'es una armadura real y el tipo no es faccion?
    If ObjData(obji).Real = 1 Then
        If Npclist(NpcIndex).name <> "SR" Then
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "Lo siento, la ropa faccionaria solo es para muestra, no tengo autorización para venderla. Diríjete al sastre de tu ejército." & "°" & str(Npclist(NpcIndex).Char.CharIndex))
            Exit Function
        End If
    End If
    
    If ObjData(obji).Caos = 1 Then
        If Npclist(NpcIndex).name <> "SC" Then
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "Lo siento, la ropa faccionaria solo es para muestra, no tengo autorización para venderla. Diríjete al sastre de tu ejército." & "°" & str(Npclist(NpcIndex).Char.CharIndex))
            Exit Function
        End If
    End If

    '¿Ya tiene un objeto de este tipo?
    Slot = 1
    Do Until UserList(userindex).Invent.Object(Slot).ObjIndex = obji And _
       UserList(userindex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
        
        Slot = Slot + 1
        If Slot > MAX_INVENTORY_SLOTS Then
            Exit Do
        End If
    Loop
    
    'Sino se fija por un slot vacio
    If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until UserList(userindex).Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1
            
            If Slot > MAX_INVENTORY_SLOTS Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||No podés tener mas objetos." & FONTTYPE_INFO)
                Exit Function
            End If
        Loop
        UserList(userindex).Invent.NroItems = UserList(userindex).Invent.NroItems + 1
    End If
    
    'desde aca para abajo se realiza la transaccion
    UserCompraObj = True
    'Mete el obj en el slot
    If UserList(userindex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
        'Menor que MAX_INV_OBJS
        UserList(userindex).Invent.Object(Slot).ObjIndex = obji
        UserList(userindex).Invent.Object(Slot).Amount = UserList(userindex).Invent.Object(Slot).Amount + Cantidad
        
        
        'tal vez suba el skill comerciar ;-)
        Call SubirSkill(userindex, Comerciar)
        
        Call SendData(SendTarget.ToIndex, userindex, 0, "||Has comprado: " & ObjData(obji).name & "." & FONTTYPE_INFO)
        If ObjData(obji).OBJType = eOBJType.otLlaves Then Call logVentaCasa(UserList(userindex).name & " compro " & ObjData(obji).name)
        
        Call QuitarNpcInvItem(UserList(userindex).flags.TargetNPC, CByte(ObjIndex), Cantidad)
    Else
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No podés tener mas objetos." & FONTTYPE_INFO)
    End If
Exit Function

errorh:
Call LogError("Error en USERCOMPRAOBJ. " & Err.Description)
End Function


Sub NpcCompraObj(ByVal userindex As Integer, ByVal ObjIndex As Integer, ByVal Cantidad As Integer)
On Error GoTo errorh
    Dim Slot As Integer
    Dim obji As Integer
    Dim NpcIndex As Integer
          
    If Cantidad < 1 Then Exit Sub
    
    If UserList(userindex).flags.Privilegios = PlayerType.Dios Then Exit Sub
    If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then Exit Sub
    
    NpcIndex = UserList(userindex).flags.TargetNPC
    obji = UserList(userindex).Invent.Object(ObjIndex).ObjIndex
    
    If ObjData(obji).Newbie = 1 Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||No comercio objetos para newbies." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(NpcIndex).TipoItems <> eOBJType.otCualquiera Then
        '¿Son los items con los que comercia el npc?
        If Npclist(NpcIndex).TipoItems <> ObjData(obji).OBJType Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||El npc no esta interesado en comprar ese objeto." & FONTTYPE_WARNING)
            Exit Sub
        End If
    End If
    
    If obji = iORO Then
        Call SendData(SendTarget.ToIndex, userindex, 0, "||El npc no esta interesado en comprar ese objeto." & FONTTYPE_WARNING)
        Exit Sub
    End If
    
    '¿Ya tiene un objeto de este tipo?
    Slot = 1
    Do Until (Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = obji _
      And Npclist(NpcIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS)
        
        Slot = Slot + 1
        
        If Slot > MAX_INVENTORY_SLOTS Then Exit Do
    Loop
    
    'Sino se fija por un slot vacio antes del slot devuelto
    If Slot > MAX_INVENTORY_SLOTS Then
        Slot = 1
        Do Until Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = 0
            Slot = Slot + 1
            If Slot > MAX_INVENTORY_SLOTS Then Exit Do
        Loop
        If Slot <= MAX_INVENTORY_SLOTS Then Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1
    End If
    
    If Slot <= MAX_INVENTORY_SLOTS Then 'Slot valido
        'Mete el obj en el slot
        Npclist(NpcIndex).Invent.Object(Slot).ObjIndex = obji
        If Npclist(NpcIndex).Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            Npclist(NpcIndex).Invent.Object(Slot).Amount = Npclist(NpcIndex).Invent.Object(Slot).Amount + Cantidad
        Else
            Npclist(NpcIndex).Invent.Object(Slot).Amount = MAX_INVENTORY_OBJS
        End If
    End If
    
    Call QuitarUserInvItem(userindex, CByte(ObjIndex), Cantidad)
    
    'tal vez suba el skill comerciar ;-)
    Call SubirSkill(userindex, Comerciar)
Exit Sub

errorh:
    Call LogError("Error en NPCCOMPRAOBJ. " & Err.Description)
End Sub

Sub IniciarCOmercioNPC(ByVal userindex As Integer)
On Error GoTo errhandler
    'Mandamos el Inventario
    Call EnviarNpcInv(userindex, UserList(userindex).flags.TargetNPC)
    'Hacemos un Update del inventario del usuario
    Call UpdateUserInv(True, userindex, 0)
    'Atcualizamos el dinero
    Call SendUserStatsBox(userindex)
    'Mostramos la ventana pa' comerciar y ver ladear la osamenta. jajaja
    UserList(userindex).flags.Comerciando = True
    SendData SendTarget.ToIndex, userindex, 0, "INITCOM"
Exit Sub

errhandler:
    Dim str As String
    str = "Error en IniciarComercioNPC. UI=" & userindex
    If userindex > 0 Then
        str = str & ".Nombre: " & UserList(userindex).name & " IP:" & UserList(userindex).ip & " comerciando con "
        If UserList(userindex).flags.TargetNPC > 0 Then
            str = str & Npclist(UserList(userindex).flags.TargetNPC).name
        Else
            str = str & "<NPCINDEX 0>"
        End If
    Else
        str = str & "<USERINDEX 0>"
    End If
End Sub

Sub NPCVentaItem(ByVal userindex As Integer, ByVal i As Integer, ByVal Cantidad As Integer, ByVal NpcIndex As Integer)
'listindex+1, cantidad
On Error GoTo errhandler
    
    If Cantidad < 1 Then Exit Sub
    
    'NPC VENDE UN OBJ A UN USUARIO
    Call SendUserStatsBox(userindex)
    
    If i > MAX_INVENTORY_SLOTS Then
        Call SendData(SendTarget.ToAdmins, 0, 0, "Posible intento de romper el sistema de comercio. Usuario: " & UserList(userindex).name & FONTTYPE_WARNING)
        Exit Sub
    End If
    
    If Cantidad > MAX_INVENTORY_OBJS Then
        Call SendData(SendTarget.ToAll, 0, 0, UserList(userindex).name & " ha sido baneado por el sistema anti-cheats." & FONTTYPE_FIGHT)
        Call Ban(UserList(userindex).name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio " & Cantidad)
        UserList(userindex).flags.Ban = 1
        Call SendData(SendTarget.ToIndex, userindex, 0, "ERRHas sido baneado por el sistema anti cheats")
        Call CloseSocket(userindex)
        Exit Sub
    End If
    
    
        If Npclist(UserList(userindex).flags.TargetNPC).Invent.Object(i).Amount > 0 Then
            If Cantidad > Npclist(UserList(userindex).flags.TargetNPC).Invent.Object(i).Amount Then Cantidad = Npclist(UserList(userindex).flags.TargetNPC).Invent.Object(i).Amount
            'Agregamos el obj que compro al inventario
            If Not UserCompraObj(userindex, CInt(i), UserList(userindex).flags.TargetNPC, Cantidad) Then
                Call SendData(SendTarget.ToIndex, userindex, 0, "||No puedes comprar este ítem." & FONTTYPE_INFO)
            End If
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, userindex, 0)
            'Actualizamos la ventana de comercio
            Call EnviarNpcInv(userindex, UserList(userindex).flags.TargetNPC)
            Call UpdateVentanaComercio(i, 0, userindex)
        End If
Exit Sub

errhandler:
    Call LogError("Error en comprar item: " & Err.Description)
End Sub

Sub NPCCompraItem(ByVal userindex As Integer, ByVal Item As Integer, ByVal Cantidad As Integer)
On Error GoTo errhandler
    Dim NpcIndex As Integer
    
    NpcIndex = UserList(userindex).flags.TargetNPC
    
    If UserList(userindex).flags.Privilegios = PlayerType.Dios Then Exit Sub
    If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then Exit Sub
    
    'Si es una armadura faccionaria vemos que la está intentando vender al sastre
    If ObjData(UserList(userindex).Invent.Object(Item).ObjIndex).Real = 1 Then
        If Npclist(NpcIndex).name <> "SR" Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Las armaduras faccionarias sólo las puedes vender a sus respectivos Sastres" & FONTTYPE_WARNING)
            
            'Actualizamos la ventana de comercio
            Call UpdateVentanaComercio(Item, 1, userindex)
            Call EnviarNpcInv(userindex, UserList(userindex).flags.TargetNPC)
            Exit Sub
        End If
    ElseIf ObjData(UserList(userindex).Invent.Object(Item).ObjIndex).Caos = 1 Then
        If Npclist(NpcIndex).name <> "SC" Then
            Call SendData(SendTarget.ToIndex, userindex, 0, "||Las armaduras faccionarias sólo las puedes vender a sus respectivos Sastres" & FONTTYPE_WARNING)
            
            'Actualizamos la ventana de comercio
            Call UpdateVentanaComercio(Item, 1, userindex)
            Call EnviarNpcInv(userindex, UserList(userindex).flags.TargetNPC)
            Exit Sub
        End If
    End If
    
    'NPC COMPRA UN OBJ A UN USUARIO
    Call SendUserStatsBox(userindex)
   
    If UserList(userindex).Invent.Object(Item).Amount > 0 And UserList(userindex).Invent.Object(Item).Equipped = 0 Then
        If Cantidad > 0 And Cantidad > UserList(userindex).Invent.Object(Item).Amount Then Cantidad = UserList(userindex).Invent.Object(Item).Amount
        'Agregamos el obj que compro al inventario
        Call NpcCompraObj(userindex, CInt(Item), Cantidad)
        'Actualizamos el inventario del usuario
        Call UpdateUserInv(True, userindex, 0)
        'Actualizamos el oro
        Call SendUserStatsBox(userindex)
        
        Call EnviarNpcInv(userindex, UserList(userindex).flags.TargetNPC)
        'Actualizamos la ventana de comercio
        Call UpdateVentanaComercio(Item, 1, userindex)
    End If
Exit Sub

errhandler:
    Call LogError("Error en vender item: " & Err.Description)
End Sub

Sub UpdateVentanaComercio(ByVal Slot As Integer, ByVal NpcInv As Byte, ByVal userindex As Integer)
    Call SendData(SendTarget.ToIndex, userindex, 0, "TRANSOK" & Slot & "," & NpcInv)
End Sub

Sub EnviarNpcInv(ByVal userindex As Integer, ByVal NpcIndex As Integer)
    'Enviamos el inventario del npc con el cual el user va a comerciar...
    Dim i As Integer

    
    For i = 1 To MAX_INVENTORY_SLOTS
        If Npclist(NpcIndex).Invent.Object(i).ObjIndex > 0 Then
            'Calculamos el porc de inflacion del npc
            SendData SendTarget.ToIndex, userindex, 0, "EFYK" & _
            ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).name _
            & "," & Npclist(NpcIndex).Invent.Object(i).Amount & _
            "," _
            & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).GrhIndex _
            & "," & Npclist(NpcIndex).Invent.Object(i).ObjIndex _
            & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).OBJType _
            & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).MaxHIT _
            & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).MinHIT _
            & "," & ObjData(Npclist(NpcIndex).Invent.Object(i).ObjIndex).MaxDef
        Else
            SendData SendTarget.ToIndex, userindex, 0, "EFYK" & _
                        "Nada" _
                        & "," & 0 _
                        & "," & 0 _
                        & "," & 0 _
                        & "," & 0 _
                        & "," & 0 _
                        & "," & 0 _
                        & "," & 0 _
                        & "," & 0 _
                        & "," & 0 _
                        & "," & 0 _
                        & "," & 0 _
                        & "," & 0 _
                        & "," & 0 _
                        & "," & 0 _
                        & "," & 0
        End If
    Next i
End Sub
