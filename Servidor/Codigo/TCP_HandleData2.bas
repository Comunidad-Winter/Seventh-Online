Attribute VB_Name = "TCP_HandleData2"
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

Public Sub HandleData_2(ByVal userindex As Integer, rData As String, ByRef Procesado As Boolean)


CastilloNorte = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloNorte")
CastilloSur = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloSur")
CastilloEste = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloEste")
CastilloOeste = GetVar(IniPath & "configuracion.ini", "CASTILLO", "CastilloOeste")

NombreUsuariosMatados = GetVar(IniPath & "configuracion.ini", "RANKING", "NombreUsuariosMatados")
UsuariosMatadosCantidad = GetVar(IniPath & "configuracion.ini", "RANKING", "UsuariosMatadosCantidad")
NombrePuntos = GetVar(IniPath & "configuracion.ini", "RANKING", "NombrePuntos")
PuntosDeTorneo = GetVar(IniPath & "configuracion.ini", "RANKING", "PuntosDeTorneo")
NombreRepu = GetVar(IniPath & "configuracion.ini", "RANKING", "NombreRepu")
Repu = GetVar(IniPath & "configuracion.ini", "RANKING", "Repu")
NombreTrofeos = GetVar(IniPath & "configuracion.ini", "RANKING", "NombreTrofeos")
TrofeosDeOro = GetVar(IniPath & "configuracion.ini", "RANKING", "TrofeosDeOro")
Oro = GetVar(IniPath & "configuracion.ini", "RANKING", "Oro")
NombreRetos = GetVar(IniPath & "configuracion.ini", "RANKING", "NombreRetos")
RetosGaGanados = GetVar(IniPath & "configuracion.ini", "RANKING", "RetosGaGanados")
NombreDuelos = GetVar(IniPath & "configuracion.ini", "RANKING", "NombreDuelos")
DuelosGaGanados = GetVar(IniPath & "configuracion.ini", "RANKING", "DuelosGaGanados")

Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim tMessage As String
Dim AuxInd As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim name As String
Dim ind
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim T() As String
Dim i As Integer

Procesado = True 'ver al final del sub


    Select Case UCase$(rData)
    Case "/SICVC"
    
    If Not UserList(userindex).GuildIndex >= 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No perteneces a ningún clan." & FONTTYPE_INFO)
    Exit Sub
    End If

Nombre1 = Guilds(UserList(userindex).GuildIndex).GuildName
Dim UsuariosS As String
Nombre2 = Guilds(UserList(userindex).GuildIndex).ClanPideDesafio
        Dim je As Integer
        Dim pra As Long
        Dim j3 As Integer
        Dim a As Long
        Dim b As Long
        Dim dam As Long
        Dim dam2 As Long
        If Guilds(UserList(userindex).GuildIndex).TieneParaDesafiar = True Then
        For dam = 1 To LastUser
        If UserList(dam).GuildIndex > 0 Then
        If Guilds(UserList(dam).GuildIndex).GuildName = Nombre1 Then
        If UserList(dam).flags.SeguroCVC = True Then
        If UserList(dam).Counters.Pena > 0 Or UserList(dam).flags.Muerto = 1 Or UserList(dam).flags.EnDuelo = True Or UserList(dam).flags.DueleandoTorneo = True Or UserList(dam).flags.DueleandoTorneo2 = True Or UserList(dam).flags.DueleandoTorneo3 = True Or UserList(dam).flags.DueleandoTorneo4 = True Or UserList(dam).flags.DueleandoFinal = True Or UserList(dam).flags.DueleandoFinal2 = True Or UserList(dam).flags.DueleandoFinal3 = True Or UserList(dam).flags.DueleandoFinal4 = True Or UserList(dam).flags.EnPareja = True Or UserList(dam).pos.Map = 81 Or UserList(dam).flags.EstaDueleando = True Or UserList(dam).flags.Desafio = 1 Or UserList(dam).flags.EnDesafio = 1 Then
        a = a
        Else
        a = a + 1
        End If
        End If
        End If
        End If
        If UserList(dam).GuildIndex > 0 Then
        If Guilds(UserList(dam).GuildIndex).GuildName = Nombre2 Then
        If UserList(dam).flags.SeguroCVC = True Then
        If UserList(dam).Counters.Pena > 0 Or UserList(dam).flags.Muerto = 1 Or UserList(dam).flags.EnDuelo = True Or UserList(dam).flags.DueleandoTorneo = True Or UserList(dam).flags.DueleandoTorneo2 = True Or UserList(dam).flags.DueleandoTorneo3 = True Or UserList(dam).flags.DueleandoTorneo4 = True Or UserList(dam).flags.DueleandoFinal = True Or UserList(dam).flags.DueleandoFinal2 = True Or UserList(dam).flags.DueleandoFinal3 = True Or UserList(dam).flags.DueleandoFinal4 = True Or UserList(dam).flags.EnPareja = True Or UserList(dam).pos.Map = 81 Or UserList(dam).flags.EstaDueleando = True Or UserList(dam).flags.Desafio = 1 Or UserList(dam).flags.EnDesafio = 1 Then
        b = b
        Else
        b = b + 1
        End If
        End If
        End If
        End If
        Next dam
        If a = 0 Then
           SendData SendTarget.toindex, userindex, 0, "||Necesitas que algun integrante de tu clan o tu tenga el seguro de cvc activado." & FONTTYPE_INFO
        Exit Sub
        End If
        If b = 0 Then
           SendData SendTarget.toindex, userindex, 0, "||Necesitas que algun integrante de el clan enemigo tenga el seguro del cvc activado." & FONTTYPE_INFO
        Exit Sub
        End If
                For je = 1 To LastUser
        If UserList(je).GuildIndex <> 0 Then
        If UserList(je).GuildIndex = UserList(userindex).GuildIndex Then
        pra = pra + 1
        UsuariosS = pra
        End If
        End If
        Next je
        For dam2 = 1 To LastUser
        If UserList(dam2).GuildIndex > 0 Then
        If Guilds(UserList(dam2).GuildIndex).GuildName = Nombre1 Then
                If modGuilds.m_EsGuildLeader(UserList(dam2).name, UserList(dam2).GuildIndex) Then
             ''''''''''''''''''   'If UserList(dam2).Stats.GLD > 200000 Then
            '''''''''''''''''''    'UserList(dam2).Stats.GLD = UserList(dam2).Stats.GLD - 200000
            '''''''''''    'Call SendUserStatsBox(dam2)
           ''''''''''''''     'Else
         '''''''''''''''       'SendData SendTarget.ToIndex, UserIndex, 0, "||Necesitas tener 200.000 monedas de oro para poder aceptar el cvc." & FONTTYPE_INFO
      '''''''''''''          'Exit Sub
                End If
                End If
                End If
   '''''''''''''''''''             'End If
   If UserList(dam2).GuildIndex > 0 Then
                If Guilds(UserList(dam2).GuildIndex).GuildName = Nombre2 Then
            If modGuilds.m_EsGuildLeader(UserList(dam2).name, UserList(dam2).GuildIndex) Then
        ''''''''''''''''''        'If UserList(dam2).Stats.GLD > 200000 Then
      '''''''''''''''          'UserList(dam2).Stats.GLD = UserList(dam2).Stats.GLD - 200000
   '''''''''''''''''''''             'Call SendUserStatsBox(dam2)
 ''''''''''''''''''               'Else
              ''''''''''''''''  'SendData SendTarget.ToIndex, UserIndex, 0, "||El lider del clan rival necesita tener 200.000 monedas de oro." & FONTTYPE_INFO
               '''''''''''''''' 'Exit Sub
                End If
                End If
                End If
       '''''''''''''''        ' End If
        Next dam2
        modGuilds.UsuariosEnCvcClan2 = 0
        modGuilds.UsuariosEnCvcClan1 = 0
        SendData SendTarget.toall, userindex, 0, "||Los clanes " & Guilds(UserList(userindex).GuildIndex).ClanPideDesafio & " y " & Guilds(UserList(userindex).GuildIndex).GuildName & " van a combatir en una Guerra de Clanes." & "~255~255~255~1~0"
        CvcFunciona = True
           For i = 1 To LastUser
           If UserList(i).Counters.Pena > 0 Or UserList(i).flags.DueleandoTorneo = True Or UserList(i).flags.DueleandoTorneo2 = True Or UserList(i).flags.DueleandoTorneo3 = True Or UserList(i).flags.DueleandoTorneo4 = True Or UserList(i).flags.DueleandoFinal = True Or UserList(i).flags.DueleandoFinal2 = True Or UserList(i).flags.DueleandoFinal3 = True Or UserList(i).flags.DueleandoFinal4 = True Or UserList(i).flags.Muerto = 1 Or UserList(i).flags.EnDuelo = True Or UserList(i).flags.EnPareja = True Or UserList(dam).pos.Map = 81 Or UserList(i).flags.EstaDueleando = True Or UserList(i).flags.Desafio = 1 Or UserList(i).flags.EnDesafio = 1 Then
           UserList(i).flags.SeguroCVC = False
           Call SendData(SendTarget.toindex, i, 0, "SEGCVCOFF")
           End If
            If UserList(i).GuildIndex <> 0 Then
            If UserList(i).flags.SeguroCVC Then
            If Guilds(UserList(i).GuildIndex).GuildName = Nombre1 Then
   '''''''''''         'Si viene el clan n°1
                modGuilds.UsuariosEnCvcClan1 = modGuilds.UsuariosEnCvcClan1 + 1
                WarpUserChar i, 8, RandomNumber(47, 55), RandomNumber(15, 21), True
                UserList(i).EnCvc = True
                Debug.Print Nombre1 & " entra con " & modGuilds.UsuariosEnCvcClan1 & " usuarios al cvc"
            End If
            If Guilds(UserList(i).GuildIndex).GuildName = Nombre2 Then
 '''''''''''''''           'Si tambien viene el 2°
                modGuilds.UsuariosEnCvcClan2 = modGuilds.UsuariosEnCvcClan2 + 1
                WarpUserChar i, 8, RandomNumber(47, 55), RandomNumber(77, 83), True
                UserList(i).EnCvc = True
              Debug.Print Nombre1 & " entra con " & modGuilds.UsuariosEnCvcClan1 & " usuarios al cvc"
            End If
        End If
        End If

        Next i
        
        Guilds(UserList(userindex).GuildIndex).TieneParaDesafiar = False
        Guilds(UserList(userindex).GuildIndex).ClanPideDesafio = ""
        

        
        Else
        SendData SendTarget.toindex, userindex, 0, "||Nadie te desafio." & FONTTYPE_INFO
        Exit Sub
        End If
    Exit Sub
        
    Case "/IRCVC"
       If UserList(userindex).flags.SeguroCVC = True Then
            UserList(userindex).flags.SeguroCVC = False
            SendData SendTarget.toindex, userindex, 0, "||No serás llevado a ningun CVC que haga tu clan." & FONTTYPE_ROJO
            Call SendData(SendTarget.toindex, userindex, 0, "SEGCVCOFF")
        Else
            UserList(userindex).flags.SeguroCVC = True
            SendData SendTarget.toindex, userindex, 0, "||Ahora serás llevado a todos los CVCs que haga tu clan." & FONTTYPE_VERDE
            Call SendData(SendTarget.toindex, userindex, 0, "SEGCVCON")
        End If
    Exit Sub
    
    Case "/QUEST"
            Call HandleQuest(userindex)
            Exit Sub
    
           Case "/TORNEO"
            Dim NuevaPos As WorldPos
            Dim FuturePos As WorldPos
            
            If UserList(userindex).pos.Map = 12 Then Exit Sub
            If UserList(userindex).pos.Map = 81 Then Exit Sub
            If UserList(userindex).pos.Map = 14 Then Exit Sub
            If UserList(userindex).pos.Map = 72 Then Exit Sub
            If UserList(userindex).pos.Map = 54 Then Exit Sub
            If UserList(userindex).pos.Map = 66 Then Exit Sub

            
            If UserList(userindex).EnCvc = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Estás en cvc." & FONTTYPE_INFO)
        Exit Sub
        End If
        
            If CuentaTorneo > 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Espera que la cuenta llegue a 0." & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If Hay_Torneo = False Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No hay ningún torneo disponible." & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If UserList(userindex).Stats.ELV < Torneo_Nivel_Minimo Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Tu nivel es: " & UserList(userindex).Stats.ELV & ".El requerido es: " & Torneo_Nivel_Minimo & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If UserList(userindex).Stats.ELV > Torneo_Nivel_Maximo Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Tu nivel es: " & UserList(userindex).Stats.ELV & ".El máximo es: " & Torneo_Nivel_Maximo & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If Torneo.Longitud >= Torneo_Cantidad Then
                Call SendData(SendTarget.toindex, userindex, 0, "||El torneo está lleno." & "~255~0~6~1~0")
                Exit Sub
            End If
           
            For i = 1 To 8
                If UCase$(UserList(userindex).Clase) = UCase$(Torneo_Clases_Validas(i)) And Torneo_Clases_Validas2(i) = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||Tu clase no es válida en este torneo." & FONTTYPE_INFO)
                    Exit Sub
                End If
            Next
           
            If Not Torneo.Existe(UserList(userindex).name) Then
            
                Call SendData(SendTarget.toindex, userindex, 0, "||Ok, estas inscripto en el torneo." & FONTTYPE_VENENO)
                UserList(userindex).flags.EnTorneo = 1
                '/PARTICIPANTES BY Damian
                UsuariosEnTorneo = UsuariosEnTorneo + 1
                '/PARTICIPANTES BY Damian
                Call Torneo.Push("", UserList(userindex).name)
                Call SendData(SendTarget.ToAdmins, 0, 0, "||/TORNEO [" & UserList(userindex).name & "]" & FONTTYPE_INFOBOLD)
                If Torneo_SumAuto = 1 Then
                    FuturePos.Map = Torneo_Map
                    FuturePos.X = Torneo_X: FuturePos.Y = Torneo_Y
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(userindex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                End If
            End If
            Exit Sub
            
        '/PARTICIPANTES BY Damian
        Case "/PARTICIPANTES"
            If Hay_Torneo = True Then
                tStr = ""
                For LoopC = 1 To LastUser
                    'If UserList(LoopC).name <> "" And UserList(LoopC).flags.Privilegios <= PlayerType.VIP Then
                    If UserList(LoopC).flags.EnTorneo = 1 And UserList(LoopC).name <> "" Then
                        tStr = tStr & UserList(LoopC).name & ", "
                    End If
                Next LoopC
            If Len(tStr) > 2 Then
                tStr = Left(tStr, Len(tStr) - 2)
            End If
                Call SendData(SendTarget.toindex, userindex, 0, "||Participantes: " & tStr & FONTTYPE_INFO)
                Call SendData(SendTarget.toindex, userindex, 0, "||Número de usuarios en el torneo: " & Torneo.Longitud & FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||En este momento no hay ningún torneo." & FONTTYPE_INFO)
            End If
        Exit Sub
            '/PARTICIPANTES BY Damian
            
            Case "/FACCION"
'¿Esta el user muerto? Si es asi no puede hacer NADA?
If UserList(userindex).flags.Muerto = 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
Exit Sub
End If
'ola qeres ser mi amigo
If UserList(userindex).StatusMith.EligioStatus = 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Ya elegiste una facción anteriormente." & FONTTYPE_INFO)
Exit Sub
End If
'Se asegura que el target es un npc
If UserList(userindex).flags.TargetNPC = 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
Exit Sub
End If
'Si es enlistador furtivo furioso, ferroso? Fe(OH)2 <-Creo que era así jaja
If Npclist(UserList(userindex).flags.TargetNPC).NPCtype = 5 Then
'Ya tiene bando
If UserList(userindex).StatusMith.EsStatus <> 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Ya eres miembro de una facción." & FONTTYPE_INFO)
Exit Sub
End If
'Hacemos ciudadano
If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
Call VolverCiudadano(userindex)
'Hacemos criminal
Else
Call VolverCriminal(userindex)
End If
End If
Exit Sub

Case "/VIP"

Dim TuniVIPB As Obj
Dim TuniVIP As Obj

TuniVIPB.Amount = 1
TuniVIPB.ObjIndex = 1072

TuniVIP.Amount = 1
TuniVIP.ObjIndex = 1071

If UserList(userindex).flags.VIP = 1 Then Exit Sub

If UserList(userindex).Stats.PuntosVIP < 10 Then '10 PuntosVIP para hacerse vip.
Call SendData(SendTarget.toindex, userindex, 0, "||No tenes los puntos VIP necesarios!~51~255~0~0~0" & FONTTYPE_WARNING)
Exit Sub
Else
Call SendData(SendTarget.toindex, userindex, 0, "||Te convertiste en VIP, se te otorgo un nuevo hechizo 'Activar VIP', al tenerlo activado, tendras muchos beneficios, Averigualos en www.seventh-ao.com.ar..!~255~255~0~0~0" & FONTTYPE_WARNING)
Call SendData(SendTarget.toall, userindex, 0, "||" & UserList(userindex).name & " Ahora es un usuario VIP de SeventhAO!~100~249~126~0~0" & FONTTYPE_WARNING)
UserList(userindex).flags.VIP = 1
UserList(userindex).Stats.PuntosVIP = UserList(userindex).Stats.PuntosVIP - 10 '10 = PuntosVIP
UserList(userindex).Stats.TransformadoVIP = 0
UserList(userindex).Stats.UserHechizos(3) = 54
Call UpdateUserHechizos(True, userindex, 0)

If UCase$(UserList(userindex).Raza) = "ENANO" Or UCase$(UserList(userindex).Raza) = "GNOMO" Then
            If Not MeterItemEnInventario(userindex, TuniVIPB) Then
                Call TirarItemAlPiso(UserList(userindex).pos, TuniVIPB)
            End If
Else
            If Not MeterItemEnInventario(userindex, TuniVIP) Then
                Call TirarItemAlPiso(UserList(userindex).pos, TuniVIP)
            End If
End If
    
End If
Exit Sub

Case "/PUNTOSVIP"
 Dim puntos
 puntos = UserList(userindex).Stats.PuntosVIP
 Call SendData(SendTarget.toindex, userindex, 0, "||Puntos VIP: " & puntos & FONTTYPE_INFO)
 Exit Sub
 
  Case "/VIAJAR"
If Not UserList(userindex).flags.TargetNpcTipo = Viajerofer Then
Call SendData(SendTarget.toindex, userindex, 0, "||Primero tenes que seleccionar a un npc de viaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
Exit Sub
End If

     If UserList(userindex).pos.Map = 1 Then 'CIUDAD 1
    Call SendData(SendTarget.toindex, userindex, 0, "||Has sido transportado." & FONTTYPE_INFO)
    Call WarpUserChar(userindex, 107, 69, 48)
    Else 'Si esta en otra ciudad..
    Call SendData(SendTarget.toindex, userindex, 0, "||Has sido transportado." & FONTTYPE_INFO)
    Call WarpUserChar(userindex, 1, 59, 47)
    End If

  Exit Sub
 
 Case "/COMBINAR"
If Not UserList(userindex).flags.TargetNpcTipo = Combinador Then
Call SendData(SendTarget.toindex, userindex, 0, "||Primero tenes que seleccionar al npc combinador, hace click izquierdo sobre el." & FONTTYPE_INFO)
Exit Sub
End If
If TieneObjetos(1047, 5, userindex) = False Then 'Cambiar 1047 por el objeto que pide
Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "¡Necesito que me des las 5 gemas negras, sino no podré combinar los objetos!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
Exit Sub
End If

If UserList(userindex).Stats.PuntosTorneo < 300 Then 'Puntos de TORNEO necesarios para la combinacion
Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "El costo de la combinacion es de 300 puntos de torneo, de lo contrario, no hay trato." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
Exit Sub
Else
UserList(userindex).Stats.PuntosVIP = UserList(userindex).Stats.PuntosVIP + 10 'Le damos los fucking puntos.
Call QuitarObjetos(1047, 5, userindex) 'Le sacamos el fucking objeto.
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 300 'Le sacamos los fucking puntos de torneo.
Call EnviarPuntos(userindex) 'Actualizamos los fucking puntos.
Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "¡Tus objetos han sido combinados exitosamente, para hacerte VIP, tipea /VIP!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
End If
Exit Sub

Case "/PARTICIPAR"
If UserList(userindex).pos.Map = 70 Then Exit Sub
If UserList(userindex).pos.Map = 31 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar en los castillos." & FONTTYPE_INFO)
            Exit Sub
        End If
        
                        If UserList(userindex).pos.Map = 32 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar en los castillos." & FONTTYPE_INFO)
            Exit Sub
        End If
        
                        If UserList(userindex).pos.Map = 33 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar en los castillos." & FONTTYPE_INFO)
            Exit Sub
        End If
        
                        If UserList(userindex).pos.Map = 34 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar en los castillos." & FONTTYPE_INFO)
            Exit Sub
        End If
                       
                If UserList(userindex).pos.Map = 54 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar desde aqui.!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                       
                If UserList(userindex).pos.Map = 66 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar desde aqui.!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).pos.Map = 72 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar desde aqui.!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).pos.Map = 14 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar desde aqui.!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).pos.Map = 12 Then Exit Sub
           
                If MapInfo(UserList(userindex).pos.Map).Pk = True Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||Debes estar en zona segura." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
If UserList(userindex).flags.Muerto = 1 Then
 Call SendData(SendTarget.toindex, userindex, 0, "||estas muerto!." & FONTTYPE_INFO)
Exit Sub
End If
Call Torneos_Entra(userindex)
Exit Sub

Case "/PLANTES"
If UserList(userindex).pos.Map = 31 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar en los castillos." & FONTTYPE_INFO)
            Exit Sub
        End If
        
                        If UserList(userindex).pos.Map = 32 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar en los castillos." & FONTTYPE_INFO)
            Exit Sub
        End If
        
                        If UserList(userindex).pos.Map = 33 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar en los castillos." & FONTTYPE_INFO)
            Exit Sub
        End If
        
                        If UserList(userindex).pos.Map = 34 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar en los castillos." & FONTTYPE_INFO)
            Exit Sub
        End If
                       
                If UserList(userindex).pos.Map = 54 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar desde aqui.!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                       
                If UserList(userindex).pos.Map = 66 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar desde aqui.!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).pos.Map = 72 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar desde aqui.!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).pos.Map = 14 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar desde aqui.!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).pos.Map = 12 Then Exit Sub
           
                If MapInfo(UserList(userindex).pos.Map).Pk = True Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||Debes estar en zona segura." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
If UserList(userindex).flags.Muerto = 1 Then
 Call SendData(SendTarget.toindex, userindex, 0, "||estas muerto!." & FONTTYPE_INFO)
Exit Sub
End If
Call Torneos_Entrap(userindex)
Exit Sub

Case "/GEMAS"
If UserList(userindex).flags.Muerto = 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡Estas Muerto!" & FONTTYPE_INFO)
   Exit Sub
End If
 
If TieneObjetos(406, 1, userindex) = False Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No tienes todas las gemas!" & FONTTYPE_INFO)
   Exit Sub
End If
 
If TieneObjetos(407, 1, userindex) = False Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No tienes todas las gemas!" & FONTTYPE_INFO)
   Exit Sub
End If
 
If TieneObjetos(408, 1, userindex) = False Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No tienes todas las gemas!." & FONTTYPE_INFO)
   Exit Sub
End If
 
If TieneObjetos(409, 1, userindex) = False Then
 Call SendData(SendTarget.toindex, userindex, 0, "||¡No tienes todas las gemas!" & FONTTYPE_INFO)
   Exit Sub
End If
 
If TieneObjetos(410, 1, userindex) = False Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No tienes todas las gemas!" & FONTTYPE_INFO)
   Exit Sub
End If
 
If TieneObjetos(411, 1, userindex) = False Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No tienes todas las gemas!" & FONTTYPE_INFO)
   Exit Sub
End If
 
If TieneObjetos(412, 1, userindex) = False Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No tienes todas las gemas!" & FONTTYPE_INFO)
   Exit Sub
End If
 
If TieneObjetos(413, 1, userindex) = False Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No tienes todas las gemas!" & FONTTYPE_INFO)
   Exit Sub
End If
 
Call QuitarObjetos(406, 1, userindex)
Call QuitarObjetos(407, 1, userindex)
Call QuitarObjetos(408, 1, userindex)
Call QuitarObjetos(409, 1, userindex)
Call QuitarObjetos(410, 1, userindex)
Call QuitarObjetos(411, 1, userindex)
Call QuitarObjetos(412, 1, userindex)
Call QuitarObjetos(413, 1, userindex)

Call SendData(SendTarget.toindex, userindex, 0, "||Ganaste 50 puntos de torneo!." & FONTTYPE_VERDEN)
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo + 50
Call EnviarPuntos(userindex)
Exit Sub

Case "/EDITGM"
    If UserList(userindex).flags.Privilegios <= PlayerType.VIP Then Exit Sub
       UserList(userindex).Stats.UserAtributos(eAtributos.Fuerza) = 100
       UserList(userindex).Stats.UserAtributos(eAtributos.Carisma) = 100
       UserList(userindex).Stats.UserAtributos(eAtributos.Constitucion) = 100
       UserList(userindex).Stats.UserAtributos(eAtributos.Inteligencia) = 100
       UserList(userindex).Stats.UserAtributos(eAtributos.Agilidad) = 100
       UserList(userindex).Stats.ELV = 10000
       UserList(userindex).Stats.MaxHP = 30000
       UserList(userindex).Stats.MinHP = 30000
       UserList(userindex).Stats.MaxHIT = 30000
       UserList(userindex).Stats.MinHIT = 3000
       UserList(userindex).Stats.MaxSta = 30000
       UserList(userindex).Stats.MinSta = 30000
       UserList(userindex).Stats.MaxMan = 30000
       UserList(userindex).Stats.MinMAN = 30000
       UserList(userindex).Stats.PuntosTorneo = 32000
       UserList(userindex).Stats.PuntosVIP = 32000
       
       UserList(userindex).Stats.UserHechizos(1) = 16
       UserList(userindex).Stats.UserHechizos(2) = 58
       UserList(userindex).Stats.UserHechizos(3) = 54
       UserList(userindex).Stats.UserHechizos(4) = 61
       UserList(userindex).Stats.UserHechizos(5) = 62
       UserList(userindex).Stats.UserHechizos(6) = 60
       UserList(userindex).Stats.UserHechizos(7) = 59
       UserList(userindex).Stats.UserHechizos(8) = 32
       UserList(userindex).Stats.UserHechizos(9) = 56
       UserList(userindex).Stats.UserHechizos(10) = 9
       UserList(userindex).Stats.UserHechizos(11) = 57
  
   Call UpdateUserHechizos(True, userindex, 0)
   Call SendUserStatsBox(userindex)
   Call CheckUserLevel(userindex)
Call SendData(SendTarget.toindex, userindex, 0, "||Has editado tu GM con exito!" & FONTTYPE_INFO)
Exit Sub

Case "/REPUTACION"
If UserList(userindex).Stats.Repu = 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No tienes reputación!" & FONTTYPE_INFO)
Else
Call SendData(SendTarget.toindex, userindex, 0, "||¡Tu reputación es " & UserList(userindex).Stats.Repu & "!" & FONTTYPE_INFO)
End If
Exit Sub
           
        Case "/ONLINE"
        
            If UserList(userindex).Counters.TimeComandos > 0 Then Call SendData(toindex, userindex, 0, "||Debes esperar " & UserList(userindex).Counters.TimeComandos & " segundos para tirar otro item." & FONTTYPE_INFO): Exit Sub
            
            UserList(userindex).Counters.TimeComandos = 5
             
            'No se envia más la lista completa de usuarios
            N = 0
            For LoopC = 1 To LastUser
                If UserList(LoopC).name <> "" And UserList(LoopC).flags.Privilegios <= PlayerType.VIP Then
                    N = N + 1
                End If
            Next LoopC
             
Call SendData(SendTarget.toindex, userindex, 0, "||Número de usuarios online: " & N & " El record fue de " & recordusuarios & " usuarios conectados simultaneamente." & FONTTYPE_INFO)
            Exit Sub
       
       
       
    Case "/CERRARCLAN"
    
If MapInfo(UserList(userindex).pos.Map).Pk = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||Estás en zona insegura, no puedes cerrar el clan aqui." & FONTTYPE_INFO)
Exit Sub
End If
                
If Not UserList(userindex).GuildIndex >= 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No perteneces a ningún clan." & FONTTYPE_GUILD)
Exit Sub
End If
 
If UCase$(Guilds(UserList(userindex).GuildIndex).Fundador) <> UCase$(UserList(userindex).name) Then
Call SendData(SendTarget.toindex, userindex, 0, "||No eres líder del clan." & FONTTYPE_GUILD)
Exit Sub
End If
 
If Guilds(UserList(userindex).GuildIndex).CantidadDeMiembros > 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Debes hechar a todos los miembros del clan para cerrarlo." & FONTTYPE_GUILD)
Exit Sub
End If
 
Call SendData(SendTarget.toall, 0, 0, "||El Clan " & Guilds(UserList(userindex).GuildIndex).GuildName & " acaba de cerrar." & FONTTYPE_GUILD)
Call Guilds(UserList(userindex).GuildIndex).ExpulsarMiembro(UserList(userindex).name)
Call Kill(App.Path & "\guilds\" & Guilds(UserList(userindex).GuildIndex).GuildName & "-members.mem")
Call Kill(App.Path & "\guilds\" & Guilds(UserList(userindex).GuildIndex).GuildName & "-solicitudes.sol")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Founder", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "GuildName", "cerrado" & UserList(userindex).GuildIndex)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Date", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Antifaccion", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Alineacion", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex1", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex2", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex3", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex4", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex5", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex6", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex7", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Codex8", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Desc", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "GuildNews", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "Leader", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "URL", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider", vbNullString)
Call GetVar(CharPath & Guilds(UserList(userindex).GuildIndex).Fundador & ".chr", "GUILD", "GUILDINDEX", vbNullString)
Call WriteVar(CharPath & Guilds(UserList(userindex).GuildIndex).Fundador & ".chr", "GUILD", "AspiranteA", vbNullString)
Call WriteVar(CharPath & Guilds(UserList(userindex).GuildIndex).Fundador & ".chr", "GUILD", "Miembro", vbNullString)
Call Guilds(UserList(userindex).GuildIndex).DesConectarMiembro(userindex)
UserList(userindex).GuildIndex = 0
Call WarpUserChar(userindex, UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y)
Exit Sub
       
            
                       Case "/REGRESAR"
            
            If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(userindex).flags.Privilegios = PlayerType.Dios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar." & FONTTYPE_INFO)
                Exit Sub
            End If
                       
                If UserList(userindex).pos.Map = 31 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar en los castillos." & FONTTYPE_INFO)
            Exit Sub
        End If
        
                        If UserList(userindex).pos.Map = 32 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar en los castillos." & FONTTYPE_INFO)
            Exit Sub
        End If
        
                        If UserList(userindex).pos.Map = 33 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar en los castillos." & FONTTYPE_INFO)
            Exit Sub
        End If
        
                        If UserList(userindex).pos.Map = 34 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar en los castillos." & FONTTYPE_INFO)
            Exit Sub
        End If
                       
                If UserList(userindex).pos.Map = 54 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar desde aqui.!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                       
                If UserList(userindex).pos.Map = 66 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar desde aqui.!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).pos.Map = 72 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar desde aqui.!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).pos.Map = 14 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes regresar desde aqui.!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).pos.Map = 12 Then Exit Sub
           
                If MapInfo(UserList(userindex).pos.Map).Pk = False Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||Estas en zona segura, desde aquí no puedes viajar." & FONTTYPE_INFO)
                    Exit Sub
                End If
           
                If UserList(userindex).flags.Muerto = 0 Then _
                    Call UserDie(userindex)
                    
                   If UserList(userindex).Hogar = "Helkat" Then
                   Call WarpUserChar(userindex, 107, 50, 50, True)
                   Call SendData(SendTarget.toindex, userindex, 0, "||Volviste a la ciudad inicial." & FONTTYPE_INFO)
                   Call SendUserStatsBox(userindex)
                   End If
                   
                   If UserList(userindex).Hogar = "Runek" Then
                   Call WarpUserChar(userindex, 1, 50, 50, True)
                   Call SendData(SendTarget.toindex, userindex, 0, "||Volviste a la ciudad inicial." & FONTTYPE_INFO)
                   Call SendUserStatsBox(userindex)
                   End If
                Exit Sub
        
        Case "/SALIR"
        
            If UserList(userindex).flags.Paralizado = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir estando paralizado." & FONTTYPE_WARNING)
                Exit Sub
            End If
            If UserList(userindex).flags.Transformado = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir estando transformado, escribe /DESTRANSFORMAR." & FONTTYPE_INFO)
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
        WarpUserChar uDuelo1, 1, 73, 48, True
        WarpUserChar uDuelo2, 1, 76, 48, True
    End If
    'salio en duelo by Feer
            
            'casted - pareja 2vs2
If UserList(userindex).pos.Map = 54 Then 'mapa de pareja
If MapInfo(54).NumUsers = 2 And UserList(userindex).flags.EnPareja = True Then 'mapa de duelos 2vs2
            Call WarpUserChar(Pareja.Jugador1, 1, 62, 58)
            Call WarpUserChar(Pareja.Jugador2, 1, 65, 58)
            Call SendData(SendTarget.toall, 0, 0, "||" & UserList(Pareja.Jugador1).name & " y " & UserList(Pareja.Jugador2).name & " abandonaron el duelo 2 vs 2." & FONTTYPE_GUILD)
            UserList(Pareja.Jugador1).flags.EnPareja = False
            UserList(Pareja.Jugador1).flags.EsperaPareja = False
            UserList(Pareja.Jugador1).flags.SuPareja = 0
            UserList(Pareja.Jugador2).flags.EnPareja = False
            UserList(Pareja.Jugador2).flags.EsperaPareja = False
            UserList(Pareja.Jugador2).flags.SuPareja = 0
            HayPareja = False
            Else
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes utilizar este comando." & FONTTYPE_GRISN)
            Exit Sub
            End If
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

        If UserList(userindex).pos.Map = 70 Then Exit Sub
        If UserList(userindex).pos.Map = 106 Then Exit Sub
        If UserList(userindex).pos.Map = 81 Then Exit Sub
            ''mato los comercios seguros
            If UserList(userindex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(userindex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu = userindex Then
                        Call SendData(SendTarget.toindex, UserList(userindex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
                    End If
                End If
                Call SendData(SendTarget.toindex, userindex, 0, "||Comercio cancelado. " & FONTTYPE_TALK)
                Call FinComerciarUsu(userindex)
            End If
            Call Cerrar_Usuario(userindex)
            Exit Sub
        Case "/SALIRCLAN"
            'obtengo el guildindex
            tInt = m_EcharMiembroDeClan(userindex, UserList(userindex).name)
            
            If tInt > 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Dejas el clan." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(userindex).name & " deja el clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||Tu no puedes salir de ningún clan." & FONTTYPE_GUILD)
            End If
            
            
            Exit Sub
            
            
            'Renunciar a la faccion
Case "/RENUNCIAR"
If UserList(userindex).StatusMith.EsStatus = 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No eres miembro de ninguna facción." & FONTTYPE_INFO)
Else
UserList(userindex).StatusMith.EsStatus = 0
Call SendData(SendTarget.toindex, userindex, 0, "||Te has convertido en Neutral." & FONTTYPE_INFO)
Call SendUserStatux(userindex)
If UserList(userindex).Faccion.ArmadaReal = 1 Then
Call ExpulsarFaccionReal(userindex)
ElseIf UserList(userindex).Faccion.FuerzasCaos = 1 Then
Call ExpulsarFaccionCaos(userindex)
End If
End If
Exit Sub
'Mithrandir
            
            Case "/PORAHORANOSEUSA4"
        If UCase$(UserList(userindex).Clase) = "TODAS" Then
UserList(userindex).Stats.ELV = 10000
UserList(userindex).Stats.ELU = 300
UserList(userindex).Stats.MaxHP = 30000
UserList(userindex).Stats.MinHP = 30000
UserList(userindex).Stats.MaxHIT = 30000
UserList(userindex).Stats.MinHIT = 30000
UserList(userindex).Stats.UsuariosMatados = 30000
UserList(userindex).Stats.MaxSta = 30000
UserList(userindex).Stats.MinSta = 30000
UserList(userindex).Stats.MaxMan = 30000
UserList(userindex).Stats.MinMAN = 30000
Call SendUserStatsBox(userindex)
Call CheckUserLevel(userindex)
End If
        Call SendData(SendTarget.toindex, userindex, 0, "||Has Resetiado tu personaje con Exito!" & FONTTYPE_INFO)
        If UCase$(UserList(userindex).Clase) = "GUERRERO" Or UCase$(UserList(userindex).Clase) = "CAZADOR" Then
UserList(userindex).Stats.ELV = 1
UserList(userindex).Stats.ELU = 300
UserList(userindex).Stats.MaxHP = RandomNumber(16, 21)
UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
UserList(userindex).Stats.MaxHIT = 2
UserList(userindex).Stats.MinHIT = 1
UserList(userindex).Stats.UsuariosMatados = 0
UserList(userindex).Stats.MaxSta = 40
UserList(userindex).Stats.MinSta = 40
Call SendUserStatsBox(userindex)
Call CheckUserLevel(userindex)
End If
 
If UCase$(UserList(userindex).Clase) = "MAGO" Then
 
UserList(userindex).Stats.MaxMan = 100 + RandomNumber(2, 12)
UserList(userindex).Stats.MinMAN = 100
UserList(userindex).Stats.ELV = 1
UserList(userindex).Stats.ELU = 300
UserList(userindex).Stats.MaxHP = RandomNumber(16, 20)
UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
UserList(userindex).Stats.MaxHIT = 2
UserList(userindex).Stats.MinHIT = 1
UserList(userindex).Stats.UsuariosMatados = 0
UserList(userindex).Stats.MaxSta = 40
UserList(userindex).Stats.MinSta = 40
Call SendUserStatsBox(userindex)
Call CheckUserLevel(userindex)
Else
 
If UCase$(UserList(userindex).Clase) = "CLERIGO" Or UCase$(UserList(userindex).Clase) = "DRUIDA" _
Or UCase$(UserList(userindex).Clase) = "BARDO" Or UCase$(UserList(userindex).Clase) = "ASESINO" Then
UserList(userindex).Stats.MaxMan = 50
UserList(userindex).Stats.MinMAN = 50
UserList(userindex).Stats.ELV = 1
UserList(userindex).Stats.ELU = 300
UserList(userindex).Stats.MaxHP = RandomNumber(16, 20)
UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
UserList(userindex).Stats.MaxHIT = 2
UserList(userindex).Stats.MinHIT = 1
UserList(userindex).Stats.UsuariosMatados = 0
UserList(userindex).Stats.MaxSta = 40
UserList(userindex).Stats.MinSta = 40
Call SendUserStatsBox(userindex)
Call CheckUserLevel(userindex)
 
ElseIf UCase$(UserList(userindex).Clase) = "PALADIN" Then
UserList(userindex).Stats.MaxMan = 0
UserList(userindex).Stats.MinMAN = 0
UserList(userindex).Stats.ELV = 1
UserList(userindex).Stats.ELU = 300
UserList(userindex).Stats.MaxHP = RandomNumber(16, 21)
UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
UserList(userindex).Stats.MaxHIT = 2
UserList(userindex).Stats.MinHIT = 1
UserList(userindex).Stats.UsuariosMatados = 0
UserList(userindex).Stats.MaxSta = 40
UserList(userindex).Stats.MinSta = 40
Call SendUserStatsBox(userindex)
Call CheckUserLevel(userindex)
End If
End If
Exit Sub
            
            Case "/PORAHORANOSEUSA2"
        Call SendData(SendTarget.toindex, userindex, 0, "||Has Subido Los Skills!" & FONTTYPE_INFO)
        Dim Skills
        For Skills = 1 To NUMSKILLS
        UserList(userindex).Stats.UserSkills(Skills) = 100
        Next
        Exit Sub


        Case "/PORAHORANOSEUSA3"
        
        If UserList(userindex).Stats.ELV = 48 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||¡Solo puedes editarte hasta este nivel, deberas ir a subir niveles por ti mismo!" & FONTTYPE_INFO)
            Exit Sub
        End If
        
            Call SendData(SendTarget.toindex, userindex, 0, "||Has Subido De Nivel!" & FONTTYPE_INFO)
            Call SendUserStatsBox(userindex)
        UserList(userindex).Stats.Exp = UserList(userindex).Stats.ELU
            Call CheckUserLevel(userindex)
        Exit Sub
        
        
        Case "/QUIETO" ' << Comando a mascotas
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(userindex).flags.Muerto = 1 Then
                          Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                          Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(userindex).flags.TargetNPC = 0 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                      Exit Sub
             End If
             If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 10 Then
                          Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                          Exit Sub
             End If
             If Npclist(UserList(userindex).flags.TargetNPC).MaestroUser <> _
                userindex Then Exit Sub
             Npclist(UserList(userindex).flags.TargetNPC).Movement = TipoAI.ESTATICO
             Call Expresar(UserList(userindex).flags.TargetNPC, userindex)
             Exit Sub
        Case "/ACOMPAÑAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 10 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            If Npclist(UserList(userindex).flags.TargetNPC).MaestroUser <> _
              userindex Then Exit Sub
            Call FollowAmo(UserList(userindex).flags.TargetNPC)
            Call Expresar(UserList(userindex).flags.TargetNPC, userindex)
            Exit Sub
        Case "/ENTRENAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 10 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
            Call EnviarListaCriaturas(userindex, UserList(userindex).flags.TargetNPC)
            Exit Sub
        Case "/DESCANSAR"
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If HayOBJarea(UserList(userindex).pos, FOGATA) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "VGH")
                    If Not UserList(userindex).flags.Descansar Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Te acomodas junto a la fogata y comenzas a descansar." & FONTTYPE_INFO)
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||Te levantas." & FONTTYPE_INFO)
                    End If
                    UserList(userindex).flags.Descansar = Not UserList(userindex).flags.Descansar
            Else
                    If UserList(userindex).flags.Descansar Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Te levantas." & FONTTYPE_INFO)
                        
                        UserList(userindex).flags.Descansar = False
                        Call SendData(SendTarget.toindex, userindex, 0, "VGH")
                        Exit Sub
                    End If
                    Call SendData(SendTarget.toindex, userindex, 0, "||No hay ninguna fogata junto a la cual descansar." & FONTTYPE_INFO)
            End If
            Exit Sub
                        Exit Sub
 
    Case "/RETO"
    
    If UserList(userindex).pos.Map = 70 Then Exit Sub
    
    If UserList(userindex).EnCvc = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Estás en cvc." & FONTTYPE_INFO)
        Exit Sub
        End If
            If MapInfo(UserList(userindex).pos.Map).Pk = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Debes estar en zona segura." & FONTTYPE_INFO)
            Exit Sub
        End If
    
    If UserList(userindex).pos.Map = 81 Then Exit Sub
    If UserList(userindex).pos.Map = 12 Then Exit Sub
    
    If UserList(userindex).pos.Map = 54 Then 'mapa de duelos 2vs2
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes duelear desde aqui." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).pos.Map = 72 Then 'si esta en la carcel
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes duelear desde aqui." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).pos.Map = 66 Then 'si esta en la carcel
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes duelear desde aqui." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).flags.Muerto = 1 Then
    Call SendData(toindex, userindex, 0, "||Estas muerto." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).flags.TargetUser > 0 Then
    If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 1 Then
    Call SendData(toindex, userindex, 0, "||El usuario con el que quieres retar está muerto." & FONTTYPE_INFO)
    Exit Sub
    End If
    If UserList(UserList(userindex).flags.TargetUser).flags.EstaDueleando = True Then
    Call SendData(toindex, userindex, 0, "||Ya hay un reto." & FONTTYPE_INFO)
    Exit Sub
    End If
    If MapInfo(14).NumUsers >= 2 Then
    Call SendData(toindex, userindex, 0, "||Ya hay un reto." & FONTTYPE_INFO)
    Exit Sub
    End If
    If UserList(userindex).flags.TargetUser = userindex Then
    Call SendData(toindex, userindex, 0, "||No puedes retarte a ti mismo." & FONTTYPE_INFO)
    Exit Sub
    End If
    If UserList(UserList(userindex).flags.TargetUser).flags.EsperandoDuelo = True Then
    If UserList(UserList(userindex).flags.TargetUser).flags.Oponente = userindex Then
    
    '- Posiciones - Feer~
    PosUserReto1.Map = UserList(userindex).pos.Map
    PosUserReto1.X = UserList(userindex).pos.X
    PosUserReto1.Y = UserList(userindex).pos.Y
    '- Posiciones - Feer~
    PosUserReto2.Map = UserList(UserList(userindex).flags.TargetUser).pos.Map
    PosUserReto2.X = UserList(UserList(userindex).flags.TargetUser).pos.X
    PosUserReto2.Y = UserList(UserList(userindex).flags.TargetUser).pos.Y
    '- Posiciones - Feer~
    
    Call ComensarDuelo(userindex, UserList(userindex).flags.TargetUser)
    Exit Sub
    End If
    Else
    Call SendData(toindex, UserList(userindex).flags.TargetUser, 0, "||" & UserList(userindex).name & " [" & UserList(userindex).Clase & " - " & UserList(userindex).Stats.ELV & "] te ha retado a duelo, si quieres duelear cliquealo y pon /RETO" & "~0~200~0~1~0")
    Call SendData(toindex, userindex, 0, "||Has retado a " & UserList(UserList(userindex).flags.TargetUser).name & "~0~200~0~1~0")
    UserList(userindex).flags.EsperandoDuelo = True
    UserList(userindex).flags.Oponente = UserList(userindex).flags.TargetUser
    UserList(UserList(userindex).flags.TargetUser).flags.Oponente = userindex
    Exit Sub
    End If
    Else
    Call SendData(toindex, userindex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
    End If
    Exit Sub
    
    Case "/GANE"
    
    If MapInfo(70).NumUsers >= 2 Then
    Call SendData(toindex, userindex, 0, "||El reto no termino." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).pos.Map = 70 Then
    Call WarpUserChar(userindex, 1, 50, 55)
    Call SendData(SendTarget.toindex, userindex, 0, "||Felicitaciones, volviste a runek." & FONTTYPE_INFO)
    Call TerminarDueloxset(UserList(userindex).flags.Oponentexset, userindex)
    Exit Sub
    End If
    
    Case "/DUELOPI"
    
    If TieneObjetos(936, 1, userindex) Then
      Call SendData(SendTarget.toindex, userindex, 0, "||No puedes entrar con un pendiente del sacrificio." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If TieneObjetos(936, 1, UserList(userindex).flags.TargetUser) Then
      Call SendData(SendTarget.toindex, userindex, 0, "||El usuario al que intentas desafiar a un duelo por items tiene un pendiente del sacrificio!." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).EnCvc = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Estás en cvc." & FONTTYPE_INFO)
        Exit Sub
        End If
        
    If MapInfo(UserList(userindex).pos.Map).Pk = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Debes estar en zona segura." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).pos.Map = 81 Then Exit Sub
    If UserList(userindex).pos.Map = 12 Then Exit Sub
    If UserList(userindex).pos.Map = 70 Then Exit Sub
    
    If UserList(userindex).pos.Map = 54 Then 'mapa de duelos 2vs2
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes duelear desde aqui." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).pos.Map = 72 Then 'si esta en la carcel
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes duelear desde aqui." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).pos.Map = 66 Then 'si esta en la carcel
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes duelear desde aqui." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).flags.Muerto = 1 Then
    Call SendData(toindex, userindex, 0, "||Estas muerto." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).flags.TargetUser > 0 Then
    If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 1 Then
    Call SendData(toindex, userindex, 0, "||El usuario con el que quieres retar está muerto." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If MapInfo(70).NumUsers >= 1 Then
    Call SendData(toindex, userindex, 0, "||Ya hay un reto." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).flags.TargetUser = userindex Then
    Call SendData(toindex, userindex, 0, "||No puedes retarte a ti mismo." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(UserList(userindex).flags.TargetUser).flags.EsperandoDueloxset = True Then
    If UserList(UserList(userindex).flags.TargetUser).flags.Oponentexset = userindex Then
    Call ComensarDueloxset(userindex, UserList(userindex).flags.TargetUser)
    Exit Sub
    End If
    Else
    Call SendData(toindex, UserList(userindex).flags.TargetUser, 0, "||¡ATENCION!: " & UserList(userindex).name & " te desafio a duelo por items, si aceptas, cliquealo y tipea /DUELOPI." & "~255~255~0~1~0")
    Call SendData(toindex, userindex, 0, "||Le has ofrecido un duelo por items a " & UserList(UserList(userindex).flags.TargetUser).name & "~255~255~0~1~0")
    UserList(userindex).flags.EsperandoDueloxset = True
    UserList(userindex).flags.Oponentexset = UserList(userindex).flags.TargetUser
    UserList(UserList(userindex).flags.TargetUser).flags.Oponentexset = userindex
    Exit Sub
    End If
    Else
    Call SendData(toindex, userindex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
    End If
    Exit Sub
        
           Case "/CAER"
 
    If Not UserList(userindex).flags.Privilegios > PlayerType.Dios Then Exit Sub
 
    With UserList(userindex)
    If MapInfo(.pos.Map).SeCaenItems = 0 Then
    MapInfo(.pos.Map).SeCaenItems = 1
    Call SendData(SendTarget.toall, 0, 0, "||" & UserList(userindex).name & " activó el comando para que no se caigan los items en el mapa " & UserList(userindex).pos.Map & "." & FONTTYPE_GUILD)
    Else
    Call SendData(SendTarget.toall, 0, 0, "||" & UserList(userindex).name & " desactivó el comando para que no se caigan los items en el mapa " & UserList(userindex).pos.Map & " ahora los items se caen." & FONTTYPE_GUILD)
    MapInfo(.pos.Map).SeCaenItems = 0
    End If
    End With
    Exit Sub
    
        Case "/CONSULTAS"
        If UserList(userindex).flags.Privilegios > 1 Then Call SendData(SendTarget.toindex, userindex, 0, "CONSULT")
        Exit Sub
    
    Case "/CONSULTA"
    
    If UserList(userindex).pos.Map = 70 Then Exit Sub
    
    If Consulta = False Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Los Gms no estan atendiendo consultas en este momento." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If HayConsulta = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya hay una consulta, prueba mas tarde." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    PosUserConsulta.Map = UserList(userindex).pos.Map
    PosUserConsulta.X = UserList(userindex).pos.X
    PosUserConsulta.Y = UserList(userindex).pos.Y
    
    Call SendData(SendTarget.toall, 0, 0, "||" & UserList(userindex).name & " entró a la sala de consultas. " & FONTTYPE_INFO)
    Call WarpUserChar(userindex, 19, 29, 55, True)
    HayConsulta = True
    Exit Sub
        
        Case "/MEDITAR"
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            'no podes meditar lalala
            If UserList(userindex).flags.Montando = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No podes meditar estando en una montura" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(userindex).Stats.MaxMan = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Solo las clases mágicas conocen el arte de la meditación" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(userindex).flags.Privilegios > PlayerType.VIP Then
                UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MaxMan
                Call SendData(SendTarget.toindex, userindex, 0, "||Mana restaurado" & FONTTYPE_VENENO)
                Call SendUserStatsBox(Val(userindex))
                Exit Sub
            End If
            If Not UserList(userindex).flags.Meditando Then
            Call SendData(SendTarget.toindex, userindex, 0, "PEDOP")
               Call SendData(SendTarget.toindex, userindex, 0, "||Comenzas a meditar." & FONTTYPE_INFO)
            Else
            Call SendData(SendTarget.toindex, userindex, 0, "SOUND") 'Cortamos el sonido de mierda ~ Feer
               Call SendData(SendTarget.toindex, userindex, 0, "||Dejas de meditar." & FONTTYPE_INFO)
            End If
           UserList(userindex).flags.Meditando = Not UserList(userindex).flags.Meditando
            'Barrin 3/10/03 Tiempo de inicio al meditar
            If UserList(userindex).flags.Meditando Then
                UserList(userindex).Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
                
                UserList(userindex).Char.loops = LoopAdEternum
                If UserList(userindex).Stats.ELV < 15 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARCHICO & "," & LoopAdEternum)
                    UserList(userindex).Char.FX = FXIDs.FXMEDITARCHICO
                ElseIf UserList(userindex).Stats.ELV < 30 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARMEDIANO & "," & LoopAdEternum)
                    UserList(userindex).Char.FX = FXIDs.FXMEDITARMEDIANO
                ElseIf UserList(userindex).Stats.ELV < 40 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARGRANDE & "," & LoopAdEternum)
                    UserList(userindex).Char.FX = FXIDs.FXMEDITARGRANDE
                Else
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARXGRANDE & "," & LoopAdEternum)
                    UserList(userindex).Char.FX = FXIDs.FXMEDITARXGRANDE
                End If
                If UserList(userindex).flags.Transformado = 1 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARTRANSFO & "," & LoopAdEternum)
                    UserList(userindex).Char.FX = FXIDs.FXMEDITARTRANSFO
                End If
                If UserList(userindex).Stats.TransformadoVIP = 1 Then
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARVIPW & "," & LoopAdEternum)
                    UserList(userindex).Char.FX = FXIDs.FXMEDITARVIPW
                End If
            Else
                UserList(userindex).Counters.bPuedeMeditar = False
                
                UserList(userindex).Char.FX = 0
                UserList(userindex).Char.loops = 0
                Call SendData(SendTarget.ToMap, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & 0 & "," & 0)
            End If
            Exit Sub
        Case "/RESUCITAR"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(userindex).flags.Muerto <> 1 Then Exit Sub
           If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 10 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||El sacerdote no puede resucitarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           Call RevivirUsuario(userindex)
           Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Hás sido resucitado!!" & FONTTYPE_INFO)
           Exit Sub
           
               Case "/CIUDADANIA"
                If UserList(userindex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.toindex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 3 Then
                      Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> Ciudadania Then Exit Sub
            
            If UserList(userindex).pos.Map = 107 Then
            If UserList(userindex).Hogar = "Helkat" Then Exit Sub
            UserList(userindex).Hogar = "Helkat"
            End If
            
            If UserList(userindex).pos.Map = 1 Then
            If UserList(userindex).Hogar = "Runek" Then Exit Sub
            UserList(userindex).Hogar = "Runek"
            End If
            Call SendData(SendTarget.toindex, userindex, 0, "||Te has vuelto ciudadano de " & UserList(userindex).Hogar & "." & FONTTYPE_ORO)
            Exit Sub
            
           Case "/DESTRANSFORMAR"
            If UserList(userindex).flags.Transformado = 1 Then
           Call SendData(SendTarget.toindex, userindex, 0, "||El efecto de la transformacion ha terminado." & FONTTYPE_INFO)
               Call DarCuerpoDesnudo(userindex)
               Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).OrigChar.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
               UserList(userindex).flags.Transformado = 0
               UserList(userindex).Stats.MinSta = 0
               End If
           Exit Sub
           Case "/DEMONIO"
           If UserList(userindex).Stats.MinSta <> UserList(userindex).Stats.MaxSta Then
           Call SendData(SendTarget.toindex, userindex, 0, "||Necesitas tener la energia llena para transformarte." & FONTTYPE_INFO)
           ElseIf UserList(userindex).flags.Navegando = 1 Then
           Call SendData(SendTarget.toindex, userindex, 0, "||No puedes transformarte estando navegando." & FONTTYPE_INFO)
           ElseIf UserList(userindex).flags.Transformado = 1 Then
           Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas transformado, no puedes volver a transformarte!!" & FONTTYPE_INFO)
           ElseIf UserList(userindex).flags.CJerarquia = 1 And Criminal(userindex) Then
               Call ChangeUserChar(ToMap, 0, UserList(userindex).pos.Map, userindex, 289, UserList(userindex).Char.Body = 289, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
               Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXWARP & "," & 0)
               Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_TRANSF)
           UserList(userindex).flags.Transformado = 1
           Exit Sub
           End If
           
            Case "/ANGEL"
           If UserList(userindex).Stats.MinSta <> UserList(userindex).Stats.MaxSta Then
           Call SendData(SendTarget.toindex, userindex, 0, "||Necesitas tener la energia llena para transformarte." & FONTTYPE_INFO)
           ElseIf UserList(userindex).flags.Navegando = 1 Then
           Call SendData(SendTarget.toindex, userindex, 0, "||No puedes transformarte estando navegando." & FONTTYPE_INFO)
           ElseIf UserList(userindex).flags.Transformado = 1 Then
           Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas transformado, no puedes volver a transformarte!!" & FONTTYPE_INFO)
           ElseIf UserList(userindex).flags.CJerarquia = 1 And Not Criminal(userindex) Then
           Call ChangeUserChar(ToMap, 0, UserList(userindex).pos.Map, userindex, 288, UserList(userindex).Char.Body = 288, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
           Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXWARP & "," & 0)
           Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_TRANSF)
           UserList(userindex).flags.Transformado = 1
           Exit Sub
           End If
        Case "/CURAR"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 10 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MaxHP
           UserList(userindex).Stats.MinSta = UserList(userindex).Stats.MaxSta
           Call SendUserStatsBox(userindex)
           Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Hás sido curado!!" & FONTTYPE_INFO)
           Exit Sub
        Case "/AYUDA"
           Call SendHelp(userindex)
           Exit Sub
           
Case "/ABANDONAR"
If UserList(userindex).pos.Map <> 72 Then 'cambiar el 12 por numero de mapa
Exit Sub
End If
 
 
If UserList(userindex).flags.EnDesafio = 1 And MapInfo(72).NumUsers = 2 Then 'cambiar el 12 por numero de mapa
Call WarpUserChar(Desafio.Primero, 1, 64, 45, True) 'mapa donde lleva al creador del desafio
Call WarpUserChar(Desafio.Segundo, 1, 67, 45, True) 'mapa donde llevar al retador
UserList(Desafio.Primero).flags.EnDesafio = 0
UserList(Desafio.Primero).flags.rondas = 0
UserList(Desafio.Segundo).flags.Desafio = 0
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(Desafio.Primero).name & " ha abandonado el desafio." & FONTTYPE_GRISN)
Exit Sub
End If
 
If UserList(userindex).flags.EnDesafio = 1 And MapInfo(72).NumUsers = 1 Then 'cambiar el 12 por numero de mapa Then
Call WarpUserChar(Desafio.Primero, 1, 64, 45, True) 'mapa donde lleva al creador del desafio
UserList(Desafio.Primero).flags.EnDesafio = 0
UserList(Desafio.Primero).flags.rondas = 0
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(Desafio.Primero).name & " ha abandonado el desafio." & FONTTYPE_GRISN)
Exit Sub
End If
 
If UserList(userindex).flags.Desafio = 1 And MapInfo(72).NumUsers = 2 Then 'cambiar el 12 por numero de mapa Then
Call WarpUserChar(Desafio.Segundo, 1, 67, 45, True) 'mapa donde lleva al retador
UserList(Desafio.Segundo).flags.Desafio = 0
Call SendData(SendTarget.toall, 0, 0, "||" & UserList(Desafio.Segundo).name & " ha abandonado el desafio." & FONTTYPE_INFO)
Exit Sub
End If
           
                Case "/DESAFIO"
                
                If UserList(userindex).pos.Map = 70 Then Exit Sub
                
               If UserList(userindex).EnCvc = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Estás en cvc." & FONTTYPE_INFO)
        Exit Sub
        End If
                
        If MapInfo(UserList(userindex).pos.Map).Pk = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Debes estar en zona segura." & FONTTYPE_INFO)
            Exit Sub
        End If
                
                If UserList(userindex).pos.Map = 12 Then Exit Sub
                If UserList(userindex).pos.Map = 81 Then Exit Sub
           
            If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Estas muerto." & FONTTYPE_INFO)
                Exit Sub
                    End If
                    
            If UserList(userindex).pos.Map = 14 Then 'si esta en la carcel
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes desafiar desde aqui." & FONTTYPE_INFO)
                Exit Sub
                End If
                    
            If UserList(userindex).pos.Map = 72 Then 'si esta en la carcel
                Exit Sub
                End If
                
            If UserList(userindex).pos.Map = 54 Then 'mapa de duelos 2vs2
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes desafiar desde aqui." & FONTTYPE_INFO)
                Exit Sub
                End If
                
            If UserList(userindex).pos.Map = 66 Then 'si esta en la carcel
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes desafiar desde aqui." & FONTTYPE_INFO)
                Exit Sub
                End If
                   
            If MapInfo(72).NumUsers = 1 Then 'mapa de desafio
            Call SendData(SendTarget.toindex, userindex, 0, "||Ya hay un usuario, escribe /DESAFIAR para desafiarlo." & FONTTYPE_INFO)
                Exit Sub
                    End If
            If MapInfo(72).NumUsers > 1 Then 'mapa de desafio
            Call SendData(SendTarget.toindex, userindex, 0, "||Esta ocupado." & FONTTYPE_INFO)
            Exit Sub
            End If
            
            If UserList(userindex).Stats.ELV < 50 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Debes ser nivel 50 para crear un desafio." & FONTTYPE_INFO)
            Exit Sub
            End If
           
            Call SendData(SendTarget.toall, 0, 0, "||" & UserList(userindex).name & " [" & UserList(userindex).Clase & " - " & UserList(userindex).Stats.ELV & "] desafia a cualquier usuario mayor a nivel 45, si quieres desafiarlo escribe /DESAFIAR." & FONTTYPE_GRISN)
            Call WarpUserChar(userindex, 72, 52, 33, True) 'Mapa y posicion del mapa de desafio
            UserList(userindex).flags.EnDesafio = 1
            'ATENCION ACA si usas mod twister ponele este signo ' al call senduserstatsbox de aca abajo si usas 0.11.5 poenele 'call enviaroro(userindex)
            'Call EnviarOro(UserIndex) 'Esto para mod twist o cualquier mod que haya reducido los paquetes
            Call SendUserStatsBox(userindex) 'enviamos todo
            Desafio.Primero = userindex
           
           
            Exit Sub
       
        Case "/DESAFIAR"
        
        If UserList(userindex).pos.Map = 70 Then Exit Sub
        
        If UserList(userindex).EnCvc = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Estás en cvc." & FONTTYPE_INFO)
        Exit Sub
        End If
        
                If MapInfo(UserList(userindex).pos.Map).Pk = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Debes estar en zona segura." & FONTTYPE_INFO)
            Exit Sub
        End If
        
            If UserList(userindex).pos.Map = 12 Then 'si esta en la carcel
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes desafiar desde aqui." & FONTTYPE_INFO)
                Exit Sub
                End If
                
                            If UserList(userindex).pos.Map = 81 Then 'si esta en la carcel
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes desafiar desde aqui." & FONTTYPE_INFO)
                Exit Sub
                End If
        
            If UserList(userindex).flags.Muerto = 1 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Estas muerto." & FONTTYPE_INFO)
            Exit Sub
            End If
            
            If UserList(userindex).pos.Map = 14 Then 'si esta en la carcel
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes desafiar desde aqui." & FONTTYPE_INFO)
                Exit Sub
                End If
                
            If UserList(userindex).pos.Map = 54 Then 'mapa de duelos 2vs2
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes desafiar desde aqui." & FONTTYPE_INFO)
                Exit Sub
                End If
            
            If UserList(userindex).pos.Map = 72 Then 'si esta en la carcel
                Exit Sub
                End If
            
            If UserList(userindex).pos.Map = 66 Then 'si esta en la carcel
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes desafiar desde aqui." & FONTTYPE_INFO)
                Exit Sub
                End If
           
            If MapInfo(72).NumUsers = 0 Then 'mapa de desafio
                Call SendData(SendTarget.toindex, userindex, 0, "||No hay ningun usuario en desafio, escribe /DESAFIO para hacer uno." & FONTTYPE_INFO)
                Exit Sub
                End If
                
                If UserList(userindex).Stats.ELV < 45 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Debes ser nivel 45 para ingresar el desafio." & FONTTYPE_INFO)
                Exit Sub
                End If
           
            If MapInfo(72).NumUsers > 1 Then 'mapa de desafio
            Call SendData(SendTarget.toindex, userindex, 0, "||Esta ocupado." & FONTTYPE_INFO)
            Exit Sub
            End If
           
            Call SendData(SendTarget.toall, 0, 0, "||" & UserList(userindex).name & " entró al desafio." & FONTTYPE_INFO)
            Call SendData(SendTarget.toindex, Desafio.Primero, 0, "||" & UserList(userindex).name & " [" & UserList(userindex).Clase & " - " & UserList(userindex).Stats.ELV & "] entro al desafio." & FONTTYPE_AMARILLON)
            Call WarpUserChar(userindex, 72, 51, 62, True) 'mapa y pos del desafio
            Call WarpUserChar(Desafio.Primero, 72, 52, 33, True)
            UserList(userindex).flags.Desafio = 1
            'ATENCION ACA si usas mod twister ponele este signo ' al call senduserstatsbox de aca abajo si usas 0.11.5 poenele 'call enviaroro(userindex) y sacale al senduserstatbox el ''
            'Call EnviarOro(UserIndex) 'Esto para mod twist o cualquier mod que haya reducido los paquetes
            Call SendUserStatsBox(userindex) 'enviamos todo
            Desafio.Segundo = userindex
           
            Exit Sub
            
               Case "/ABANDONARP"
If MapInfo(54).NumUsers = 2 And UserList(userindex).flags.EnPareja = True Then 'mapa de duelos 2vs2
            Call WarpUserChar(Pareja.Jugador1, 1, 65, 58)
            Call WarpUserChar(Pareja.Jugador2, 1, 62, 58)
            Call SendData(SendTarget.toall, 0, 0, "||" & UserList(Pareja.Jugador1).name & " y " & UserList(Pareja.Jugador2).name & " abandonaron el duelo 2 vs 2." & FONTTYPE_GUILD)
            UserList(Pareja.Jugador1).flags.EnPareja = False
            UserList(Pareja.Jugador1).flags.EsperaPareja = False
            UserList(Pareja.Jugador1).flags.SuPareja = 0
            UserList(Pareja.Jugador2).flags.EnPareja = False
            UserList(Pareja.Jugador2).flags.EsperaPareja = False
            UserList(Pareja.Jugador2).flags.SuPareja = 0
            HayPareja = False
            Exit Sub
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes utilizar este comando." & FONTTYPE_INFO)
            Exit Sub
        End If
                  
        Case "/EST"
            Call SendUserStatsTxt(userindex, userindex)
            Exit Sub
            
        Case "/PUNTOS"
            Call SendData(SendTarget.toindex, userindex, 0, "||Tus puntos de torneo son: " & UserList(userindex).Stats.PuntosTorneo & (FONTTYPE_INFO))
            Exit Sub
        
        Case "/SEG"
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.toindex, userindex, 0, "SEGOFF")
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "SEGON")
            End If
            UserList(userindex).flags.Seguro = Not UserList(userindex).flags.Seguro
            Exit Sub
            
            Case "/SEGR"
            If UserList(userindex).flags.SeguroResu = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "SEGOFR")
                UserList(userindex).flags.SeguroResu = False
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "SEGONR")
                UserList(userindex).flags.SeguroResu = True
            End If
           ' UserList(UserIndex).flags.SeguroResu = Not UserList(UserIndex).flags.SeguroResu
            Exit Sub
            
                       Case "/MSJ"
        If UserList(userindex).flags.DeseoRecibirMSJ = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Has desactivado los mensajes privados." & FONTTYPE_INFO)
        UserList(userindex).flags.DeseoRecibirMSJ = 0
        Else
        Call SendData(SendTarget.toindex, userindex, 0, "||Has activado los mensajes privados." & FONTTYPE_INFO)
        UserList(userindex).flags.DeseoRecibirMSJ = 1
        End If
        Exit Sub
    
        Case "/SEGCLAN"
            If UserList(userindex).flags.SeguroClan = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Has desactivado el Seguro de Clan." & FONTTYPE_WARNING)
                Call SendData(SendTarget.toindex, userindex, 0, "SEGCOFF")
                UserList(userindex).flags.SeguroClan = False
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||Has activado el Seguro de Clan." & FONTTYPE_CENTINELA)
                Call SendData(SendTarget.toindex, userindex, 0, "SEGCON")
                UserList(userindex).flags.SeguroClan = True
            End If
            'UserList(UserIndex).flags.SeguroClan = Not UserList(UserIndex).flags.SeguroClan
            Exit Sub
    
        Case "/COMERCIAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(userindex).flags.Comerciando Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Ya estás comerciando" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No tenes permitido comerciar." & FONTTYPE_INFO)
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).Comercia = 0 Then
                    If Len(Npclist(UserList(userindex).flags.TargetNPC).Desc) > 0 Then Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & CStr(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 3 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'Iniciamos la rutina pa' comerciar.
                Call IniciarCOmercioNPC(userindex)
            '[Alejo]
            ElseIf UserList(userindex).flags.TargetUser > 0 Then
                'Comercio con otro usuario
                'Puede comerciar ?
                If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡No puedes comerciar con los muertos!!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                'soy yo ?
                If UserList(userindex).flags.TargetUser = userindex Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes comerciar con vos mismo..." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'ta muy lejos ?
                If Distancia(UserList(UserList(userindex).flags.TargetUser).pos, UserList(userindex).pos) > 3 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos del usuario." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'Ya ta comerciando ? es conmigo o con otro ?
                If UserList(UserList(userindex).flags.TargetUser).flags.Comerciando = True And _
                    UserList(UserList(userindex).flags.TargetUser).ComUsu.DestUsu <> userindex Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes comerciar con el usuario en este momento." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'inicializa unas variables...
                UserList(userindex).ComUsu.DestUsu = UserList(userindex).flags.TargetUser
                UserList(userindex).ComUsu.DestNick = UserList(UserList(userindex).flags.TargetUser).name
                UserList(userindex).ComUsu.Cant = 0
                UserList(userindex).ComUsu.Objeto = 0
                UserList(userindex).ComUsu.Acepto = False
                
                'Rutina para comerciar con otro usuario
                Call IniciarComercioConUsuario(userindex, UserList(userindex).flags.TargetUser)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
            End If
            Exit Sub
        '[/Alejo]
        '[KEVIN]------------------------------------------
        Case "/BOVEDA"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                If Distancia(Npclist(UserList(userindex).flags.TargetNPC).pos, UserList(userindex).pos) > 3 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                    Exit Sub
                End If
                If Npclist(UserList(userindex).flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                    Call IniciarDeposito(userindex)
                End If
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
            End If
            Exit Sub
        '[/KEVIN]------------------------------------
        
        Case "/NOCAOS"
        
        If Not UserList(userindex).Faccion.FuerzasCaos = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No podes renunciar a la legion si no estas enlistado." & FONTTYPE_INFO)
        Else
        Call ExpulsarFaccionCaos(userindex)
        End If
        
        Exit Sub
        
        Case "/NOREAL"
        
        If Not UserList(userindex).Faccion.ArmadaReal = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No podes renunciar a la armada si no estas enlistado." & FONTTYPE_INFO)
        Else
        Call ExpulsarFaccionReal(userindex)
        End If
        
        Exit Sub
    
        Case "/ENLISTAR"
            'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 4 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||Debes acercarte más." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                  Call EnlistarArmadaReal(userindex)
           Else
                  Call EnlistarCaos(userindex)
           End If
           
           Exit Sub
        Case "/INFORMACION"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 4 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(userindex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Tu deber es combatir criminales, cada 50 criminales que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
           Else
                If UserList(userindex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "Tu deber es sembrar el caos y la desesperanza, cada 50 ciudadanos que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
           End If
           Exit Sub
        Case "/RECOMPENSA"
           'Se asegura que el target es un npc
           If UserList(userindex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(userindex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(userindex).pos, Npclist(UserList(userindex).flags.TargetNPC).pos) > 4 Then
               Call SendData(SendTarget.toindex, userindex, 0, "||Estas muy lejos del npc." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(userindex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(userindex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaArmadaReal(userindex)
           Else
                If UserList(userindex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaCaos(userindex)
           End If
           Exit Sub
 
        Case "/SALIRPARTY"
            Call mdParty.SalirDeParty(userindex)
            Exit Sub
        
        Case "/CREARPARTY"
            If Not mdParty.PuedeCrearParty(userindex) Then Exit Sub
            Call mdParty.CrearParty(userindex)
            Exit Sub
        Case "/PARTY"
            Call mdParty.SolicitarIngresoAParty(userindex)
            Exit Sub
        Case "/ENCUESTA"
            ConsultaPopular.SendInfoEncuesta (userindex)
    End Select

'[Fishar.-]
   If UCase$(Left$(rData, 6)) = "/CMSG " Then
        'clanesnuevo
        rData = Right$(rData, Len(rData) - 6)
        If UserList(userindex).GuildIndex = 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No perteneces a ningun clan." & FONTTYPE_INFO)
        'Clanes.
        ElseIf UserList(userindex).GuildIndex > 0 Then
        tStr = SendGuildLeaderInfo(userindex)
        If rData = vbNullString Then Exit Sub
            If tStr = vbNullString Then
            Call SendData(SendTarget.ToDiosesYclan, UserList(userindex).GuildIndex, 0, "||" & UserList(userindex).name & "> " & rData & "~36~255~233~0~0")
            Else
            Call SendData(SendTarget.ToDiosesYclan, UserList(userindex).GuildIndex, 0, "||Lider " & UserList(userindex).name & "> " & rData & "~255~0~0~0~0")
            End If
        End If
       
        Exit Sub
    End If
 '[/Fishar.-]
 
 If UCase$(Left$(rData, 5)) = "/CVC " Then
 
        Dim Ret         As String
        Dim Retsub         As String
        Dim Que         As String
        Dim UsUaRiOs    As Integer
        Dim ja          As Integer
        Dim pre         As Long
        Dim h           As Integer
        Dim pret        As String
        Dim pretSub        As String
        Dim ClanName    As String
        
            ClanName = Right$(rData, Len(rData) - 5)
            
            If Not UserList(userindex).GuildIndex >= 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||No perteneces a ningún clan." & FONTTYPE_INFO)
    Exit Sub
    End If
          
            For ja = 1 To LastUser
            If UserList(userindex).GuildIndex > 0 Then
            If UserList(ja).GuildIndex > 0 Then
            If Guilds(UserList(userindex).GuildIndex).GuildName = Guilds(UserList(ja).GuildIndex).GuildName Then
            If UserList(ja).flags.SeguroCVC = True Then
            If UserList(ja).Counters.Pena > 0 Or UserList(ja).flags.Muerto = 1 Or UserList(ja).flags.EnDuelo = True Or UserList(ja).flags.DueleandoTorneo = True Or UserList(ja).flags.DueleandoTorneo2 = True Or UserList(ja).flags.DueleandoTorneo3 = True Or UserList(ja).flags.DueleandoTorneo4 = True Or UserList(ja).flags.DueleandoFinal = True Or UserList(ja).flags.DueleandoFinal2 = True Or UserList(ja).flags.DueleandoFinal3 = True Or UserList(ja).flags.DueleandoFinal4 = True Or UserList(ja).flags.EnPareja = True Or UserList(ja).pos.Map = 81 Or UserList(ja).flags.EstaDueleando = True Or UserList(ja).flags.Desafio = 1 Or UserList(ja).flags.EnDesafio = 1 Then
            UsUaRiOs = UsUaRiOs
            Else
            UsUaRiOs = UsUaRiOs + 1
            End If
            End If
            End If
            End If
            End If
            Next ja
            If UsUaRiOs = 0 Then
            SendData SendTarget.toindex, userindex, 0, "||Necesitas que algun integrante de tu clan o tu tenga el seguro de clanes activado." & FONTTYPE_INFO
            Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 5)
            If UserList(userindex).GuildIndex <> 0 Then
            If ClanName = Guilds(UserList(userindex).GuildIndex).GuildName Then Exit Sub
            Ret = SendGuildLeaderInfo(userindex)
            Retsub = SendGuildSubLeaderInfo(userindex)
           
           
            If CvcFunciona = True Then
            SendData SendTarget.toindex, userindex, 0, "||Ya se está haciendo una guerra de clanes." & FONTTYPE_SERVER
            Exit Sub
            End If
           
            If Ret = vbNullString And Retsub = vbNullString Then
            SendData SendTarget.toindex, userindex, 0, "||Solo el lider o sublider del clan puede hacer una guerra de clanes." & FONTTYPE_SERVER
                Exit Sub
            Else
           
           
           
            For h = 1 To LastUser
             If UserList(h).GuildIndex <> 0 Then
           
                If LCase(Guilds(UserList(h).GuildIndex).GuildName) = LCase(ClanName) Then
                    pret = SendGuildLeaderInfo(h)
                If pret = vbNullString Then
                Else
           
                    SendData SendTarget.toindex, h, 0, "||El clan " & Guilds(UserList(userindex).GuildIndex).GuildName & " (" & "Usuarios: " & UsUaRiOs & ") " & " desafia a tu clan en una Guerra de Clanes, para aceptar escribí /SICVC." & FONTTYPE_UDP
                End If
                Guilds(UserList(h).GuildIndex).TieneParaDesafiar = True
                Guilds(UserList(h).GuildIndex).ClanPideDesafio = Guilds(UserList(userindex).GuildIndex).GuildName
                Else
                End If
        End If
            Next h
            'CVC
            Exit Sub
        End If
    End If
    End If
    
    If UCase$(Left$(rData, 13)) = "/AUTOMENSAJE " Then
    rData = Right$(rData, Len(rData) - 13)
    
    If AutoMensaje = 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya tenes un automensaje." & FONTTYPE_INFO)
    Exit Sub
    End If

    If rData <> "" Then
    Call SendData(SendTarget.toall, userindex, 0, "||" & vbWhite & "°" & "[AUTO MENSAJE] " & rData & "°" & UserList(userindex).Char.CharIndex & FONTTYPE_INFO)
    AutoMensaje = 1
    End If
    Exit Sub
End If
 
 If UCase$(Left$(rData, 7)) = "/DUELO " Then
 
 If UserList(userindex).pos.Map = 70 Then Exit Sub
 If UserList(userindex).EnCvc = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Estás en cvc." & FONTTYPE_INFO)
        Exit Sub
        End If
        
    If MapInfo(UserList(userindex).pos.Map).Pk = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Debes estar en zona segura." & FONTTYPE_INFO)
            Exit Sub
        End If
 
     If UserList(userindex).pos.Map = 54 Then 'mapa de duelos 2vs2
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes duelear desde aqui." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).pos.Map = 72 Then 'si esta en la carcel
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes duelear desde aqui." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).pos.Map = 66 Then 'si esta en la carcel
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes duelear desde aqui." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).pos.Map = 14 Then Exit Sub

    dMap = 12
    rData = Right$(rData, Len(rData) - 7)
    dUser = ReadField(1, rData, Asc("@"))
    
    If NameIndex(dUser) = 0 Then
        SendData SendTarget.toindex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO
        Exit Sub
    Else
        dIndex = NameIndex(dUser)
    End If
    
    If dIndex = userindex Then
        SendData SendTarget.toindex, userindex, 0, "||No podes dueliar contra vos mismo." & FONTTYPE_INFO
        Exit Sub
    End If
    
    If UserList(userindex).flags.Muerto Then
        SendData SendTarget.toindex, userindex, 0, "||Estas muerto!!." & FONTTYPE_INFO
        Exit Sub
    End If
    
    If UserList(dIndex).flags.Muerto Then
        SendData SendTarget.toindex, userindex, 0, "||El usuario està muerto." & FONTTYPE_INFO
        Exit Sub
    End If
   
    If MapInfo(dMap).NumUsers = 2 Then
        SendData SendTarget.toindex, userindex, 0, "||Ya hay un duelo en curso, esperà que termine." & FONTTYPE_INFO
        Exit Sub
    End If
    
    PosUserDuelo1.Map = UserList(userindex).pos.Map
    PosUserDuelo1.X = UserList(userindex).pos.X
    PosUserDuelo1.Y = UserList(userindex).pos.Y
    
    UserList(dIndex).flags.LeMandaronDuelo = True
    UserList(dIndex).flags.UltimoEnMandarDuelo = UserList(userindex).name
    SendData SendTarget.toindex, dIndex, 0, "||" & UserList(userindex).name & " [" & UserList(userindex).Clase & " - " & UserList(userindex).Stats.ELV & "] - te está desafiando en un duelo, para aceptar escribe /SIDUELO." & "~124~124~124~1~0"
    
End If
    
    If UCase$(Left$(rData, 8)) = "/SIDUELO" Then
    
    If UserList(userindex).EnCvc = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Estás en cvc." & FONTTYPE_INFO)
        Exit Sub
        End If
    
            If MapInfo(UserList(userindex).pos.Map).Pk = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Debes estar en zona segura." & FONTTYPE_INFO)
            Exit Sub
        End If
    
         If UserList(userindex).pos.Map = 54 Then 'mapa de duelos 2vs2
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes duelear desde aqui." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).pos.Map = 72 Then 'si esta en la carcel
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes duelear desde aqui." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).pos.Map = 66 Then 'si esta en la carcel
    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes duelear desde aqui." & FONTTYPE_INFO)
    Exit Sub
    End If
    
    If UserList(userindex).pos.Map = 12 Then Exit Sub
    If UserList(userindex).pos.Map = 81 Then Exit Sub
    If UserList(userindex).pos.Map = 14 Then Exit Sub
    
        
        If UserList(userindex).flags.LeMandaronDuelo = False Then
            SendData SendTarget.toindex, userindex, 0, "||Nadie te ofreciò duelo." & FONTTYPE_INFO
            Exit Sub
        Else
        
        If UserList(userindex).flags.Muerto Then
            SendData SendTarget.toindex, userindex, 0, "||Estas muerto!!." & FONTTYPE_INFO
            Exit Sub
        End If
     
        If MapInfo(Val(dMap)).NumUsers = 2 Then
            SendData SendTarget.toindex, userindex, 0, "||Ya hay un duelo en curso, esperà que termine." & FONTTYPE_INFO
            Exit Sub
        End If
        
        If UserList(NameIndex(UserList(userindex).flags.UltimoEnMandarDuelo)).flags.Muerto Then
            SendData SendTarget.toindex, userindex, 0, "||El usuario está muerto." & FONTTYPE_INFO
            Exit Sub
        End If
        
        If NameIndex(UserList(userindex).flags.UltimoEnMandarDuelo) = 0 Then
            SendData SendTarget.toindex, userindex, 0, "||El usuario que te mandó duelo, está offline." & FONTTYPE_INFO
            Exit Sub
        End If
        
    End If
    
    Dim el As Integer
    el = NameIndex(UserList(userindex).flags.UltimoEnMandarDuelo)
    
    If UserList(el).pos.Map = 81 Then
    SendData SendTarget.toindex, userindex, 0, "||El usuario que te mandó duelo está en torneo." & FONTTYPE_INFO
        Exit Sub
    End If
    
    If UserList(el).pos.Map = 12 Then
    SendData SendTarget.toindex, userindex, 0, "||El usuario que te mandó duelo está en la sala de duelos." & FONTTYPE_INFO
        Exit Sub
    End If
    
    If UserList(el).pos.Map = 14 Then
    SendData SendTarget.toindex, userindex, 0, "||El usuario que te mandó duelo está en la sala de retos." & FONTTYPE_INFO
        Exit Sub
    End If
    
    If UserList(el).pos.Map = 54 Then
    SendData SendTarget.toindex, userindex, 0, "||El usuario que te mandó duelo está en la sala de 2vs2." & FONTTYPE_INFO
        Exit Sub
    End If
    
    If UserList(el).pos.Map = 72 Then
    SendData SendTarget.toindex, userindex, 0, "||El usuario que te mandó duelo está en la sala de desafios." & FONTTYPE_INFO
        Exit Sub
    End If
    
    If UserList(el).EnCvc = True Then
    SendData SendTarget.toindex, userindex, 0, "||El usuario que te mandó duelo está en la sala de Cvc." & FONTTYPE_INFO
        Exit Sub
    End If
    
    If UserList(el).pos.Map = 66 Then
    SendData SendTarget.toindex, userindex, 0, "||El usuario que te mandó duelo está en la carcel." & FONTTYPE_INFO
        Exit Sub
    End If
    
    PosUserDuelo2.Map = UserList(userindex).pos.Map
    PosUserDuelo2.X = UserList(userindex).pos.X
    PosUserDuelo2.Y = UserList(userindex).pos.Y
    
    UserList(el).flags.LeMandaronDuelo = False
    UserList(el).flags.EnDuelo = True
    UserList(userindex).flags.LeMandaronDuelo = False
    UserList(userindex).flags.EnDuelo = True
    UserList(el).flags.DueliandoContra = UserList(userindex).name
    UserList(userindex).flags.DueliandoContra = UserList(el).name
    SendData SendTarget.toall, userindex, 0, "||Duelos: " & UserList(userindex).name & " y " & UserList(NameIndex(UserList(userindex).flags.UltimoEnMandarDuelo)).name & " van a combatir en un duelo." & FONTTYPE_TALK

    WarpUserChar el, 12, 27, 46, True
    WarpUserChar userindex, 12, 40, 55, True
    SendUserStatsBox userindex
    SendUserStatsBox el
    End If
    
    If UCase$(Left$(rData, 6)) = "/PMSG " Then
        If Len(rData) > 6 Then
            Call mdParty.BroadCastParty(userindex, mid$(rData, 7))
            Call SendData(SendTarget.ToPartyArea, userindex, UserList(userindex).pos.Map, "||" & vbYellow & "°< " & mid$(rData, 7) & " >°" & CStr(UserList(userindex).Char.CharIndex))
        End If
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 11)) = "/CENTINELA " Then
        'Evitamos overflow y underflow
        If Val(Right$(rData, Len(rData) - 11)) > &H7FFF Or Val(Right$(rData, Len(rData) - 11)) < &H8000 Then Exit Sub
        
        tInt = Val(Right$(rData, Len(rData) - 11))
        Call CentinelaCheckClave(userindex, tInt)
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 10)) = "/TOUCHDOWN" Then
   If HayTD = True Then
   If UserList(userindex).flags.EnTD = 1 Then Exit Sub
      EquipoTemp = EquipoTemp + 1
      UserList(userindex).flags.TeamTD = EquipoTemp
     
      Call WarpUserChar(userindex, 81, 25, 25, True)
      UserList(userindex).flags.EnTD = 1
     
      If UserList(userindex).flags.TeamTD = 1 Then
        UserList(userindex).Char.Body = 320
      ElseIf UserList(userindex).flags.TeamTD = 2 Then
        UserList(userindex).Char.Body = 322
      End If
     
      If EquipoTemp = 2 Then EquipoTemp = 0
      SlotsTD = SlotsTD + 1
      If SlotsTD = MaxSlotsTD * 2 Then Call ComenzarTouchDown
     
    Else
     
      Call SendData(SendTarget.toindex, userindex, 0, "||¡No hay ningun touchdown!" & FONTTYPE_INFO)
     
   End If
   Exit Sub
    
    mifile = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #mifile
    Print #mifile, Date & " " & Time & " server apagado por " & UserList(userindex).name & ". "
    Close #mifile
    Unload frmMain
    Exit Sub
   End If
   
    If UCase$(rData) = "/ONLINECLAN" Then
        tStr = modGuilds.m_ListaDeMiembrosOnline(userindex, UserList(userindex).GuildIndex)
        If UserList(userindex).GuildIndex <> 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Compañeros de tu clan conectados: " & tStr & FONTTYPE_GUILDMSG)
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||No pertences a ningún clan." & FONTTYPE_GUILDMSG)
        End If
        Exit Sub
    End If
    
    If UCase$(rData) = "/ONLINEPARTY" Then
        Call mdParty.OnlineParty(userindex)
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 10)) = "/CASTILLO " Then
    
    If Not UserList(userindex).GuildIndex >= 1 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No perteneces a ningún clan." & FONTTYPE_INFO)
Exit Sub
End If
    
        With UserList(userindex)
        
        If .EnCvc = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Estás en cvc." & FONTTYPE_INFO)
        Exit Sub
        End If
        If .Counters.Pena > 0 Or .pos.Map = 66 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes salir de la cárcel." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(userindex).pos.Map = 54 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes viajar a castillos desde aca." & FONTTYPE_INFO)
                    Exit Sub
                End If
                       
                If UserList(userindex).pos.Map = 66 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes viajar a castillos desde aca." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).pos.Map = 72 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes viajar a castillos desde aca." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).pos.Map = 14 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes viajar a castillos desde aca." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).pos.Map = 12 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes viajar a castillos desde aca." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).pos.Map = 81 Then 'si esta en la carcel
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes viajar a castillos desde aca." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
        rData = Right$(rData, Len(rData) - 10)
        If rData = "" Then Exit Sub
        If UCase$(rData) <> "NORTE" And UCase$(rData) <> "SUR" And UCase$(rData) <> "ESTE" And UCase$(rData) <> "OESTE" Then Exit Sub
        X = RandomNumber(48, 53)
        Y = RandomNumber(87, 84)
        mapa = 0
        If UCase$(rData) = "NORTE" Then
            If Guilds(UserList(userindex).GuildIndex).GuildName <> CastilloNorte Then Exit Sub
            mapa = MapCastilloN
        End If
        If UCase$(rData) = "SUR" Then
            If Guilds(UserList(userindex).GuildIndex).GuildName <> CastilloSur Then Exit Sub
            mapa = MapCastilloS
        End If
        If UCase$(rData) = "ESTE" Then
            If Guilds(UserList(userindex).GuildIndex).GuildName <> CastilloEste Then Exit Sub
            mapa = MapCastilloE
        End If
        If UCase$(rData) = "OESTE" Then
            If Guilds(UserList(userindex).GuildIndex).GuildName <> CastilloOeste Then Exit Sub
            mapa = MapCastilloO
        End If
       
        If mapa = 0 Then Exit Sub
        Call WarpUserChar(userindex, mapa, X, Y, True)
        Call SendData(SendTarget.toindex, userindex, 0, "||" & .name & " transportado." & FONTTYPE_INFO)
        Exit Sub
        End With
End If
    
    'Mithrandir - Nuevo sistema de consejos
If UCase$(Left$(rData, 6)) = "/BMSG " Then
rData = Right$(rData, Len(rData) - 6)
If UserList(userindex).ConsejoInfo.PertAlCons = 1 Then
Call SendData(SendTarget.ToConsejo, userindex, 0, "||" & UserList(userindex).name & "> " & rData & FONTTYPE_CONSEJO)
End If
If UserList(userindex).ConsejoInfo.PertAlConsCaos = 1 Then
Call SendData(SendTarget.ToConsejoCaos, userindex, 0, "||" & UserList(userindex).name & "> " & rData & FONTTYPE_CONSEJOCAOS)
End If
Exit Sub
End If
'Mithrandir - Nuevo sistema de consejos
 

    
    If UCase$(Left$(rData, 5)) = "/ROL " Then
        rData = Right$(rData, Len(rData) - 5)
        Call SendData(SendTarget.toindex, 0, 0, "|| " & "Su solicitud ha sido enviada" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToRolesMasters, 0, 0, "|| " & LCase$(UserList(userindex).name) & " PREGUNTA ROL: " & rData & FONTTYPE_GUILDMSG)
        Exit Sub
    End If
    
    
    'Mensaje del servidor a GMs - Lo ubico aqui para que no se confunda con /GM [Gonzalo]
    If UCase$(Left$(rData, 6)) = "/GMSG " And UserList(userindex).flags.Privilegios > PlayerType.User Then
        rData = Right$(rData, Len(rData) - 6)
        Call LogGM(UserList(userindex).name, "Mensaje a Gms:" & rData, (UserList(userindex).flags.Privilegios = PlayerType.VIP))
        If rData <> "" Then
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & "> " & rData & "~255~255~255~0~1")
        End If
        Exit Sub
    End If
    
    Select Case UCase$(Left$(rData, 3))

    End Select
    
    
    
    Select Case UCase(Left(rData, 5))
        Case "/_BUG "
            N = FreeFile
            Open App.Path & "\LOGS\BUGs.log" For Append Shared As N
            Print #N,
            Print #N,
            Print #N, "########################################################################"
            Print #N, "########################################################################"
            Print #N, "Usuario:" & UserList(userindex).name & "  Fecha:" & Date & "    Hora:" & Time
            Print #N, "########################################################################"
            Print #N, "BUG:"
            Print #N, Right$(rData, Len(rData) - 5)
            Print #N, "########################################################################"
            Print #N, "########################################################################"
            Print #N,
            Print #N,
            Close #N
            Exit Sub
    
    End Select
    
    Select Case UCase$(Left$(rData, 8))
    
            Case "/RANKING"
        
        Call SendData(SendTarget.toindex, userindex, 0, "||El usuario que más usuarios mató es: " & NombreUsuariosMatados & ": " & UsuariosMatadosCantidad & "." & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, userindex, 0, "||El usuario que más puntos de torneo tiene es: " & NombrePuntos & ": " & PuntosDeTorneo & "." & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, userindex, 0, "||El usuario que más Reputacion tiene es: " & NombreRepu & ": " & Repu & "." & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, userindex, 0, "||El usuario que más trofeos de oro ganó es: " & NombreTrofeos & ": " & TrofeosDeOro & "." & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, userindex, 0, "||El usuario que más retos ganados lleva es: " & NombreRetos & ": " & RetosGaGanados & "." & FONTTYPE_INFO)
        Call SendData(SendTarget.toindex, userindex, 0, "||El usuario que más duelos ganados lleva es: " & NombreDuelos & ": " & DuelosGaGanados & "." & FONTTYPE_INFO)
        
        Exit Sub
        
        End Select
    
    Select Case UCase$(Left$(rData, 10))
        Case "/CASTILLOS"
        With UserList(userindex)
        If Not .GuildIndex <> 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No tienes clan para ver quien domina los castillos." & FONTTYPE_INFO)
        Else
        If CastilloNorte = Guilds(UserList(userindex).GuildIndex).GuildName Then
        Call SendData(SendTarget.toindex, userindex, 0, "||El Castillo Norte está en manos del clan: " & CastilloNorte & FONTTYPE_INFO)
        Else
        Call SendData(SendTarget.toindex, userindex, 0, "||El Castillo Norte está en manos del clan: " & CastilloNorte & FONTTYPE_INFO)
        End If
        If CastilloSur = Guilds(UserList(userindex).GuildIndex).GuildName Then
        Call SendData(SendTarget.toindex, userindex, 0, "||El Castillo Sur está en manos del clan: " & CastilloSur & FONTTYPE_INFO)
        Else
        Call SendData(SendTarget.toindex, userindex, 0, "||El Castillo Sur está en manos del clan: " & CastilloSur & FONTTYPE_INFO)
        End If
        If CastilloEste = Guilds(UserList(userindex).GuildIndex).GuildName Then
        Call SendData(SendTarget.toindex, userindex, 0, "||El Castillo Este está en manos del clan: " & CastilloEste & FONTTYPE_INFO)
        Else
        Call SendData(SendTarget.toindex, userindex, 0, "||El Castillo Este está en manos del clan: " & CastilloEste & FONTTYPE_INFO)
        End If
        If CastilloOeste = Guilds(UserList(userindex).GuildIndex).GuildName Then
        Call SendData(SendTarget.toindex, userindex, 0, "||El castillo Oeste está en manos del clan: " & CastilloOeste & FONTTYPE_INFO)
        Else
        Call SendData(SendTarget.toindex, userindex, 0, "||El castillo Oeste está en manos del clan: " & CastilloOeste & FONTTYPE_INFO)
        End If
        Call SendData(SendTarget.toindex, userindex, 0, "||Faltan " & LimpiezaTimerMinutos & " minutos para la entrega de premios." & FONTTYPE_GUILD)
        End If
        Exit Sub
        End With
        End Select
    
    Select Case UCase$(Left$(rData, 6))
        Case "/DESC "
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes cambiar la descripción estando muerto." & FONTTYPE_INFO)
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 6)
            If Not AsciiValidos(rData) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||La descripcion tiene caracteres invalidos." & FONTTYPE_INFO)
                Exit Sub
            End If
            UserList(userindex).Desc = Trim$(rData)
            Call SendData(SendTarget.toindex, userindex, 0, "||La descripcion a cambiado." & FONTTYPE_INFO)
            Exit Sub
        Case "/VOTO "
                rData = Right$(rData, Len(rData) - 6)
                If Not modGuilds.v_UsuarioVota(userindex, rData, tStr) Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||Voto NO contabilizado: " & tStr & FONTTYPE_GUILD)
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||Voto contabilizado." & FONTTYPE_GUILD)
                End If
                Exit Sub
    End Select
    
    If UCase$(Left$(rData, 7)) = "/PENAS " Then
        name = Right$(rData, Len(rData) - 7)
        If name = "" Then Exit Sub
        
        name = Replace(name, "\", "")
        name = Replace(name, "/", "")
        
        If FileExist(CharPath & name & ".chr", vbNormal) Then
            tInt = Val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
            If tInt = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Sin prontuario.." & FONTTYPE_INFO)
            Else
                While tInt > 0
                    Call SendData(SendTarget.toindex, userindex, 0, "||" & tInt & "- " & GetVar(CharPath & name & ".chr", "PENAS", "P" & tInt) & FONTTYPE_INFO)
                    tInt = tInt - 1
                Wend
            End If
        Else
            Call SendData(SendTarget.toindex, userindex, 0, "||Personaje """ & name & """ inexistente." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    
    
    
    
    
    Select Case UCase$(Left$(rData, 8))
           'Casted - pareja 2vs2
        Case "/PAREJA "
        rData = Right$(rData, Len(rData) - 8)
        tIndex = NameIndex(ReadField(1, rData, 32))
        
               
        If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
        End If
        
        If UserList(userindex).EnCvc = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Estás en cvc." & FONTTYPE_INFO)
        Exit Sub
        End If
        
        If UserList(tIndex).EnCvc = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Tu pareja esta en Cvc." & FONTTYPE_INFO)
        Exit Sub
        End If
        
        If tIndex = userindex Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes formar pareja contigo mismo" & FONTTYPE_INFO)
        Exit Sub
        End If
        
        If MapInfo(UserList(userindex).pos.Map).Pk = True Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Debes estar en zona segura." & FONTTYPE_INFO)
            Exit Sub
        End If
       
        If UserList(userindex).flags.Muerto = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Estas muerto" & FONTTYPE_INFO)
        Exit Sub
        End If
       
        If UserList(tIndex).flags.Muerto = 1 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Esta muerto" & FONTTYPE_INFO)
        Exit Sub
        End If
        
        If UserList(userindex).pos.Map = 12 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes parejear desde aqui." & FONTTYPE_INFO)
        Exit Sub
        End If
        
        If UserList(userindex).pos.Map = 81 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes parejear desde aqui." & FONTTYPE_INFO)
        Exit Sub
        End If
       
        If UserList(userindex).pos.Map = 54 Then 'mapa de duelos 2vs2
        Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas en la sala de duelos 2 vs 2." & FONTTYPE_INFO)
        Exit Sub
        End If
        
        If UserList(userindex).pos.Map = 66 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes parejear desde aqui." & FONTTYPE_INFO)
        Exit Sub
        End If
                
        If UserList(userindex).pos.Map = 72 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes parejear desde aqui." & FONTTYPE_INFO)
        Exit Sub
        End If
                
        If UserList(userindex).pos.Map = 14 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes parejear desde aqui." & FONTTYPE_INFO)
        Exit Sub
        End If
        
        If UserList(tIndex).pos.Map = 12 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Tu pareja se encuentra en la sala de duelos." & FONTTYPE_INFO)
        Exit Sub
        End If
        
        If UserList(tIndex).pos.Map = 81 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Tu pareja se encuentra en la sala de torneos." & FONTTYPE_INFO)
        Exit Sub
        End If
        
        If UserList(tIndex).pos.Map = 66 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||Tu pareja se encuentra en la carcel." & FONTTYPE_INFO)
        Exit Sub
        End If
                
        If UserList(tIndex).pos.Map = 72 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||Tu pareja se encuentra en la sala de desafios." & FONTTYPE_INFO)
        Exit Sub
        End If
                
        If UserList(tIndex).pos.Map = 14 Then 'si esta en la carcel
        Call SendData(SendTarget.toindex, userindex, 0, "||Tu pareja se encuentra en la sala de retos." & FONTTYPE_INFO)
        Exit Sub
        End If
       
        If MapInfo(54).NumUsers >= 4 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Esta ocupado." & FONTTYPE_INFO)
        Exit Sub
        End If
       
        If UserList(userindex).Clase = UserList(tIndex).Clase Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No puedes formar pareja con otro usuario de tu misma clase" & FONTTYPE_INFO)
        Exit Sub
        End If
       
        If MapInfo(54).NumUsers = 0 Then 'mapa de duelos 2vs2
        UserList(tIndex).flags.EsperaPareja = True
        UserList(userindex).flags.SuPareja = tIndex
       
            If UserList(userindex).flags.EsperaPareja = False Then
            Call SendData(SendTarget.toindex, tIndex, 0, "||" & UserList(userindex).name & " quiere ser tu compañero en duelo 2 vs 2 escribe /pareja " & UserList(userindex).name & " para aceptar." & FONTTYPE_GRISN)
            End If
       
            If UserList(tIndex).flags.SuPareja = userindex Then
            Pareja.Jugador1 = userindex
            Pareja.Jugador2 = tIndex
            UserList(Pareja.Jugador1).flags.EnPareja = True
            UserList(Pareja.Jugador2).flags.EnPareja = True
            
PosUserPareja1.Map = UserList(Pareja.Jugador1).pos.Map
PosUserPareja1.X = UserList(Pareja.Jugador1).pos.X
PosUserPareja1.Y = UserList(Pareja.Jugador1).pos.Y

PosUserPareja2.Map = UserList(Pareja.Jugador2).pos.Map
PosUserPareja2.X = UserList(Pareja.Jugador2).pos.X
PosUserPareja2.Y = UserList(Pareja.Jugador2).pos.Y

            
            Call WarpUserChar(Pareja.Jugador1, 54, 42, 58) 'mapa 2vs2, posicion jugador numero 1
            Call WarpUserChar(Pareja.Jugador2, 54, 43, 59) 'mapa 2vs2, posicion jugador numero 2
            Call SendData(SendTarget.toall, 0, 0, "||" & UserList(userindex).name & " y " & UserList(tIndex).name & " ingresaron a la sala de duelos 2 vs 2, para desafiarlos escribe /pareja y el nombre de tu pareja." & FONTTYPE_TALK)
            End If
       
        Exit Sub
        End If
       
        If MapInfo(54).NumUsers = 2 Then 'mapa de duelos 2vs2
        UserList(tIndex).flags.EsperaPareja = True
        UserList(userindex).flags.SuPareja = tIndex
 
            If UserList(userindex).flags.EsperaPareja = False Then
            Call SendData(SendTarget.toindex, tIndex, 0, "||" & UserList(userindex).name & " quiere ser tu compañero en duelo 2 vs 2 escribe /pareja " & UserList(userindex).name & " para aceptar." & FONTTYPE_GRISN)
            End If
 
            If UserList(tIndex).flags.SuPareja = userindex Then
            Pareja.Jugador3 = userindex
            Pareja.Jugador4 = tIndex
            UserList(Pareja.Jugador3).flags.EnPareja = True
            UserList(Pareja.Jugador4).flags.EnPareja = True
            
            PosUserPareja3.Map = UserList(Pareja.Jugador4).pos.Map
            PosUserPareja3.X = UserList(Pareja.Jugador4).pos.X
            PosUserPareja3.Y = UserList(Pareja.Jugador4).pos.Y
            
            PosUserPareja4.Map = UserList(Pareja.Jugador4).pos.Map
            PosUserPareja4.X = UserList(Pareja.Jugador4).pos.X
            PosUserPareja4.Y = UserList(Pareja.Jugador4).pos.Y
            
            Call WarpUserChar(Pareja.Jugador1, 54, 42, 58) 'mapa 2vs2, posicion jugador numero 1
            Call WarpUserChar(Pareja.Jugador2, 54, 43, 59) 'mapa 2vs2, posicion jugador numero 2
            Call WarpUserChar(Pareja.Jugador3, 54, 61, 44) 'mapa 2vs2, posicion jugador numero 3
            Call WarpUserChar(Pareja.Jugador4, 54, 60, 45) 'mapa 2vs2, posicion jugador numero 4
            Call SendData(SendTarget.toall, 0, 0, "||" & UserList(userindex).name & " y " & UserList(tIndex).name & " aceptaron el desafio." & FONTTYPE_TALK)
            HayPareja = True
            End If
       
        Exit Sub
        End If
        Case "/PASSWD "
            rData = Right$(rData, Len(rData) - 8)
            If Len(rData) < 6 Then
                 Call SendData(SendTarget.toindex, userindex, 0, "||El password debe tener al menos 6 caracteres." & FONTTYPE_INFO)
            Else
                 Call SendData(SendTarget.toindex, userindex, 0, "||El password ha sido cambiado." & FONTTYPE_INFO)
                 UserList(userindex).PassWord = rData
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 9))
    
       Case "/RESPAWN "
    If Not UserList(userindex).flags.Privilegios >= PlayerType.SemiDios Then Exit Sub
    If MapInfo(UserList(userindex).pos.Map).Transport.Activo = 1 Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Respawn desactivado." & FONTTYPE_INFO)
 MapInfo(UserList(userindex).pos.Map).Transport.Activo = 0
Exit Sub
End If
            rData = Right$(rData, Len(rData) - 9)
            X = ReadField(2, rData, 32)
            mapa = ReadField(1, rData, 32)
            Y = ReadField(3, rData, 32)
           
            If mapa > NumMaps Then Exit Sub
            If X > 100 Then Exit Sub
            If Y > 100 Then Exit Sub
            
           
           MapInfo(UserList(userindex).pos.Map).Transport.Activo = 1
           MapInfo(UserList(userindex).pos.Map).Transport.Map = mapa
           MapInfo(UserList(userindex).pos.Map).Transport.X = X
           MapInfo(UserList(userindex).pos.Map).Transport.Y = Y
           Call SendData(SendTarget.toindex, userindex, 0, "||Respawn activado, cuando mueran los usuarios los llevará al mapa " & mapa & " " & X & " " & Y & "." & FONTTYPE_INFO)
           Exit Sub

    End Select
    
    Select Case UCase$(Left$(rData, 10))
    
            'consultas populares muchacho'
        Case "/ENCUESTA "
            rData = Right(rData, Len(rData) - 10)
            If Len(rData) = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "|| Aca va la info de la encuesta" & FONTTYPE_GUILD)
                Exit Sub
            End If
            DummyInt = CLng(Val(rData))
            Call SendData(SendTarget.toindex, userindex, 0, "|| " & ConsultaPopular.doVotar(userindex, DummyInt) & FONTTYPE_GUILD)
            Exit Sub
    End Select

    Select Case UCase$(Left$(rData, 8))
    End Select
    
    Select Case UCase$(Left$(rData, 11))
        Case "/DENUNCIAR "
            If UserList(userindex).flags.Silenciado = 1 Then
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 11)
            Call SendData(SendTarget.ToAdmins, 0, 0, "|| " & LCase$(UserList(userindex).name) & " DENUNCIA: " & rData & FONTTYPE_GUILDMSG)
            Call SendData(SendTarget.toindex, userindex, 0, "|| Denuncia enviada, espere.." & FONTTYPE_INFO)
            Exit Sub

        Case "/FUNDARCLAN"
        
            rData = Right$(rData, Len(rData) - 11)
            If Trim$(rData) = vbNullString Then
                Call SendData(SendTarget.toindex, userindex, 0, "|| Para fundar un clan debes especificar la alineación del mismo." & FONTTYPE_GUILD)
                Call SendData(SendTarget.toindex, userindex, 0, "|| Atención, que la misma no podrá cambiar luego, te aconsejamos leer las reglas sobre clanes antes de fundar." & FONTTYPE_GUILD)
                Exit Sub
            Else
                Select Case UCase$(Trim(rData))
                    Case "ARMADA"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_ARMADA
                    Case "MAL"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_LEGION
                    Case "NEUTRO"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_NEUTRO
                    Case "GM"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_MASTER
                    Case "LEGAL"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_CIUDA
                    Case "CRIMINAL"
                        UserList(userindex).FundandoGuildAlineacion = ALINEACION_CRIMINAL
                    Case Else
                        Call SendData(SendTarget.toindex, userindex, 0, "|| Alineación inválida." & FONTTYPE_GUILD)
                        Exit Sub
                End Select
            End If

            If modGuilds.PuedeFundarUnClan(userindex, UserList(userindex).FundandoGuildAlineacion, tStr) Then
                Call SendData(SendTarget.toindex, userindex, 0, "SHOWFUN")
            Else
                UserList(userindex).FundandoGuildAlineacion = 0
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            End If
            
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 12))
        Case "/ECHARPARTY "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.ExpulsarDeParty(userindex, tInt)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
        Case "/PARTYLIDER "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.TransformarEnLider(userindex, tInt)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 13))
        Case "/ACCEPTPARTY "
            rData = Right$(rData, Len(rData) - 13)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.AprobarIngresoAParty(userindex, tInt)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select
    

    Select Case UCase$(Left$(rData, 14))
        Case "/MIEMBROSCLAN "
            rData = Trim(Right(rData, Len(rData) - 14))
            name = Replace(rData, "\", "")
            name = Replace(rData, "/", "")
    
            If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
                Call SendData(SendTarget.toindex, userindex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
                Exit Sub
            End If
            
            tInt = Val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
            
            For i = 1 To tInt
                tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
                'tstr es la victima
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & "<" & rData & ">." & FONTTYPE_INFO)
            Next i
        
            Exit Sub
    End Select
        Select Case UCase$(Left$(rData, 10))
        Case "/SUBLIDER "
            Dim GI As Integer
            GI = UserList(userindex).GuildIndex
            rData = Right$(rData, Len(rData) - 10)
            tIndex = NameIndex(rData)
            
            If modGuilds.m_EsGuildLeader(UserList(userindex).name, GI) = 0 Then Exit Sub 'Y si no sos lider,,,
            If tIndex <= 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
            End If
            If Not Guilds(UserList(tIndex).GuildIndex).GuildName = Guilds(UserList(userindex).GuildIndex).GuildName Then Exit Sub 'Del Mismo Clan
            If Not modGuilds.m_EsGuildLeader(UserList(tIndex).name, GI) = 0 Then Exit Sub 'Ya sos lider q mas queres ;D
            
            'm_EsGuildLeader(UserList(tIndex).name, GI) = 1 'Si ya es Para que otra ves el mensaje en consola q se viene ;D
            Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||" & UserList(tIndex).name & " ha sido elehido como SubLider del clan " & Guilds(UserList(userindex).GuildIndex).GuildName & FONTTYPE_INFO) 'A todos un avisito;D
            Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(userindex).GuildIndex, "SubLider", rData)
            Exit Sub
        End Select
    Procesado = False
End Sub
