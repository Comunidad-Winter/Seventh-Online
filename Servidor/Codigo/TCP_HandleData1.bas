Attribute VB_Name = "TCP_HandleData1"
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

Public Sub HandleData_1(ByVal userindex As Integer, rData As String, ByRef Procesado As Boolean)


Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim iStr As String
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

    Select Case UCase$(Left$(rData, 1))
    
              Case "X"        ' >>> Sistema Consultas
            rData = Right$(rData, Len(rData) - 1)
            Dim Usuario As Integer
            Dim texto As String
            Usuario = NameIndex(ReadField(1, rData, Asc("*")))
            texto = ReadField(2, rData, Asc("*"))
            Call SendData(SendTarget.toindex, Usuario, 0, "||Tu pregunta ha sido respondida, para verla, escribì /GM y clickea el boton 'Respuesta'." & "~255~255~255~1~0")
            Call SendData(SendTarget.toindex, Usuario, 0, "RESPUES" & texto)
            Exit Sub
        Case "#"       ' >>> Sistema Consultas
        Debug.Print "Me llego SOS"
            rData = Right$(rData, Len(rData) - 1)
            Dim TipoConsulta As Byte
            Dim rDatax As String
            TipoConsulta = ReadField(1, rData, Asc(","))
            rDatax = ReadField(2, rData, Asc(","))
   
            If UserList(userindex).flags.Silenciado = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Estás silenciado." & FONTTYPE_INFO)
                Exit Sub
            End If
       
            If TipoConsulta = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||Nuevo mensaje S.o.S" & "~195~250~35~1~0")
                    Call SendData(SendTarget.ToAdmins, 0, 0, "NEWSOSM[Pregunta]," & UserList(userindex).name & "," & rDatax)
            ElseIf TipoConsulta = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||Nuevo mensaje S.o.S" & "~195~250~35~1~0")
                    Call SendData(SendTarget.ToAdmins, 0, 0, "NEWSOSM[Descargo]," & UserList(userindex).name & "," & rDatax)
            ElseIf TipoConsulta = 2 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||Nuevo mensaje S.o.S" & "~195~250~35~1~0")
                    Call SendData(SendTarget.ToAdmins, 0, 0, "NEWSOSM[Denuncia / Acusacion]," & UserList(userindex).name & "," & rDatax)
            ElseIf TipoConsulta = 3 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||Nuevo mensaje S.o.S" & "~195~250~35~1~0")
                    Call SendData(SendTarget.ToAdmins, 0, 0, "NEWSOSM[Sugerencia]," & UserList(userindex).name & "," & rDatax)
            ElseIf TipoConsulta = 4 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "NEWSOSM[Bug]," & UserList(userindex).name & "," & rDatax)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||Nuevo mensaje S.o.S" & "~195~250~35~1~0")
            End If
            Exit Sub
    
        Case ";" 'Hablar
            rData = Right$(rData, Len(rData) - 1)
            If InStr(rData, "°") Then
                Exit Sub
            End If
            
            If AutoMensaje = 1 Then Exit Sub
        
            '[VIPs]
            If UserList(userindex).flags.Privilegios = PlayerType.VIP Then
                Call LogGM(UserList(userindex).name, "Dijo: " & rData, True)
            End If
            
            ind = UserList(userindex).Char.CharIndex
            
            'piedra libre para todos los compas!
            If UserList(userindex).flags.Oculto > 0 Then
                UserList(userindex).flags.Oculto = 0
                If UserList(userindex).flags.Invisible = 0 Then
                    Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡Has vuelto a ser visible!" & FONTTYPE_INFO)
                End If
            End If
            
If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToDeadArea, userindex, UserList(userindex).pos.Map, "||12632256°" & rData & "°" & CStr(ind))
            Else
            If UserList(userindex).flags.Privilegios > VIP Then
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbYellow & "°" & rData & "°" & CStr(ind))
                If Not rData = vbNullString Then
                End If
            ElseIf UserList(userindex).flags.Privilegios = VIP Then
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & rData & "°" & CStr(ind))
                If Not rData = vbNullString Then
                End If
            ElseIf UserList(userindex).flags.Privilegios = User Then
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & rData & "°" & CStr(ind))
                If Not rData = vbNullString Then
                End If
            End If
            End If
            Exit Sub
        Case "-" 'Gritar
            If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
                    Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 1)
            If InStr(rData, "°") Then
                Exit Sub
            End If
            '[VIPs]
            If UserList(userindex).flags.Privilegios = PlayerType.VIP Then
                Call LogGM(UserList(userindex).name, "Grito: " & rData, True)
            End If
    
            'piedra libre para todos los compas!
            If UserList(userindex).flags.Oculto > 0 Then
                UserList(userindex).flags.Oculto = 0
                If UserList(userindex).flags.Invisible = 0 Then
                    Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡Has vuelto a ser visible!" & FONTTYPE_INFO)
                End If
            End If
    
    
            ind = UserList(userindex).Char.CharIndex
            Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbRed & "°" & rData & "°" & str(ind))
            Exit Sub
       Case "\" 'Susurrar al oido
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 1)
            tName = ReadField(1, rData, 32)
            
            'A los dioses y admins no vale susurrarles si no sos uno vos mismo (así no pueden ver si están conectados o no)
            If (EsDios(tName) Or EsAdministrador(tName)) And UserList(userindex).flags.Privilegios < PlayerType.Dios Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes susurrarle a los Dioses y Admins." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            'A los VIPs y SemiDioses no vale susurrarles si sos un PJ común.
            If UserList(userindex).flags.Privilegios = PlayerType.User And (EsSemiDios(tName)) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes susurrarle a los GMs" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            tIndex = NameIndex(tName)
            If tIndex <> 0 Then
                If Len(rData) <> Len(tName) Then
                    tMessage = Right$(rData, Len(rData) - (1 + Len(tName)))
                Else
                    tMessage = " "
                End If
                ind = UserList(userindex).Char.CharIndex
                If InStr(tMessage, "°") Then
                    Exit Sub
                End If
                
            If UserList(tIndex).flags.DeseoRecibirMSJ = 0 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||El usuario no desea recibir ningun mensaje." & FONTTYPE_INFO)
            Exit Sub
            End If
                
                '[VIPs]
                If UserList(userindex).flags.Privilegios = PlayerType.VIP Then
                    Call LogGM(UserList(userindex).name, "Le dijo a '" & UserList(tIndex).name & "' " & tMessage, True)
                End If
                
                    Call SendData(SendTarget.toindex, userindex, UserList(userindex).pos.Map, "||" & "Le dijiste a " & UserList(tIndex).name & "> " & tMessage & FONTTYPE_VERDE)
                    Call SendData(SendTarget.toindex, tIndex, UserList(userindex).pos.Map, "||" & UserList(userindex).name & " te dice> " & tMessage & FONTTYPE_AMARILLON)
                Exit Sub
            End If
            Call SendData(SendTarget.toindex, userindex, 0, "||Usuario inexistente. " & FONTTYPE_INFO)
            Exit Sub
    
            
        Case "M" 'Moverse

            rData = Right$(rData, Len(rData) - 1)
            
            If UserList(userindex).flags.Stopped Then Exit Sub
            
            If AutoMensaje = 1 Then
            Call SendData(SendTarget.toall, userindex, 0, "||" & vbWhite & "° °" & UserList(userindex).Char.CharIndex & FONTTYPE_INFO)
            AutoMensaje = 0
            End If
            
            'salida parche
            If UserList(userindex).Counters.Saliendo Then
                Call SendData(SendTarget.toindex, userindex, 0, "||/salir cancelado." & FONTTYPE_WARNING)
                UserList(userindex).Counters.Saliendo = False
                UserList(userindex).Counters.Salir = 0
            End If
            
            If UserList(userindex).flags.Paralizado = 0 Then
            If Not UserList(userindex).flags.Descansar And Not UserList(userindex).flags.Meditando Then
                    Call MoveUserChar(userindex, Val(rData))
                ElseIf UserList(userindex).flags.Descansar Then
                  UserList(userindex).flags.Descansar = False
                  Call SendData(toindex, userindex, 0, "VGH")
                  Call SendData(toindex, userindex, 0, "||Has dejado de descansar." & FONTTYPE_INFO)
                  Call MoveUserChar(userindex, Val(rData))
                ElseIf UserList(userindex).flags.Meditando Then
                  'Call SendData(ToIndex, UserIndex, 0, "PRE53")
                  UserList(userindex).flags.Meditando = False
                  Call SendData(toindex, userindex, 0, "PEDOP")
                  Call SendData(toindex, userindex, 0, "SOUND")
                  Call SendData(toindex, userindex, 0, "||Dejas de meditar." & FONTTYPE_INFO)
                  UserList(userindex).Char.FX = 0
                  UserList(userindex).Char.loops = 0
                  Call SendData(ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & 0 & "," & 0)
                  Call MoveUserChar(userindex, Val(rData))
                End If
            Else    'paralizado
                '[CDT 17-02-2004] (<- emmmmm ?????)
                If Not UserList(userindex).flags.UltimoMensaje = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No podes moverte porque estas paralizado." & FONTTYPE_INFO)
                    UserList(userindex).flags.UltimoMensaje = 1
                End If
                '[/CDT]
   
            End If
            
            If UserList(userindex).flags.Oculto = 1 Then
                If UCase$(UserList(userindex).Clase) <> "LADRON" Then
                    UserList(userindex).flags.Oculto = 0
                    If UserList(userindex).flags.Invisible = 0 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
                    End If
                End If
            End If
            
            If UserList(userindex).flags.Muerto = 1 Then
                Call Empollando(userindex)
            Else
                UserList(userindex).flags.EstaEmpo = 0
                UserList(userindex).EmpoCont = 0
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(rData)
        
    Case "ACTPT"
    Call EnviarPuntos(userindex)
    Exit Sub
    
    Case "FEERMANDA"
    Call EnviarPuntosDonacion(userindex)
    Exit Sub

        Case "RPU" 'Pedido de actualizacion de la posicion
            Call SendData(SendTarget.toindex, userindex, 0, "PU" & UserList(userindex).pos.X & "," & UserList(userindex).pos.Y)
            Exit Sub
        Case "AT"
        If UserList(userindex).Lac.LPegar.Puedo = False Then Exit Sub
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||¡¡No podes atacar a nadie porque estas muerto!!. " & FONTTYPE_INFO)
                Exit Sub
            End If

                If UserList(userindex).Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||No podés usar asi esta arma." & FONTTYPE_INFO)
                        Exit Sub
                    End If

                Call UsuarioAtaca(userindex)
                
                'piedra libre para todos los compas!
                If UserList(userindex).flags.Oculto > 0 And UserList(userindex).flags.AdminInvisible = 0 Then
                    UserList(userindex).flags.Oculto = 0
                    If UserList(userindex).flags.Invisible = 0 Then
                        Call SendData(SendTarget.ToMap, 0, UserList(userindex).pos.Map, "NOVER" & UserList(userindex).Char.CharIndex & ",0")
                        Call SendData(SendTarget.toindex, userindex, 0, "||¡Has vuelto a ser visible!" & FONTTYPE_INFO)
                    End If
                End If
                
            End If
            Exit Sub
        Case "AG"
            If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!! Los muertos no pueden tomar objetos. " & FONTTYPE_INFO)
                    Exit Sub
            End If
            Call GetObj(userindex)
            Exit Sub
        Case "SEG" 'Activa / desactiva el seguro
            If UserList(userindex).flags.Seguro Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Escribe /SEG para quitar el seguro" & FONTTYPE_FIGHT)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "SEGON")
                UserList(userindex).flags.Seguro = Not UserList(userindex).flags.Seguro
            End If
            Exit Sub
        Case "ACTUALIZAR"
            Call SendData(SendTarget.toindex, userindex, 0, "PU" & UserList(userindex).pos.X & "," & UserList(userindex).pos.Y)
            Exit Sub
        Case "GLINFO"
        Call LoadGuildsClanes
        Dim GI As Integer
            GI = UserList(userindex).GuildIndex
            tStr = SendGuildLeaderInfo(userindex)
            iStr = SendGuildSubLeaderInfo(userindex)
            If tStr = vbNullString And iStr = vbNullString Then
                Call SendData(SendTarget.toindex, userindex, 0, "GL" & SendGuildsList(userindex))
            Else
            If m_EsGuildSubLeader(UserList(userindex).name, GI) Then
            Call SendData(SendTarget.toindex, userindex, 0, "LEADSUB" & iStr)
            Else
            Call SendData(SendTarget.toindex, userindex, 0, "IREDAEL" & tStr)
            End If
            End If
           Exit Sub
        Case "ATRI"
            Call EnviarAtrib(userindex)
            Exit Sub
        Case "YGIJ"
            Call EnviarFama(userindex)
            Exit Sub
        Case "ESKI"
            Call EnviarSkills(userindex)
            Exit Sub
        Case "FEST" 'Mini estadisticas :)
            Call EnviarMiniEstadisticas(userindex)
            Exit Sub
        '[Alejo]
        Case "FINCOM"
            'User sale del modo COMERCIO
            UserList(userindex).flags.Comerciando = False
            Call SendData(SendTarget.toindex, userindex, 0, "FINCOMOK")
            Exit Sub
        Case "FINCOMUSU"
            'Sale modo comercio Usuario
            If UserList(userindex).ComUsu.DestUsu > 0 And _
                UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu = userindex Then
                Call SendData(SendTarget.toindex, UserList(userindex).ComUsu.DestUsu, 0, "||" & UserList(userindex).name & " ha dejado de comerciar con vos." & FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
            End If
            
            Call FinComerciarUsu(userindex)
            Exit Sub
        '[KEVIN]---------------------------------------
        '******************************************************
               Case "INIBOV"
            Call SendUserStatsBox(userindex)
            Call IniciarDeposito(userindex)
            Exit Sub
        Case "FINBAN"
            'User sale del modo BANCO
            UserList(userindex).flags.Comerciando = False
            Call SendData(SendTarget.toindex, userindex, 0, "FINBANOK")
            Exit Sub
        '-------------------------------------------------------
        '[/KEVIN]**************************************
        Case "COMUSUOK"
            'Aceptar el cambio
            Call AceptarComercioUsu(userindex)
            Exit Sub
        Case "COMUSUNO"
            'Rechazar el cambio
            If UserList(userindex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(userindex).ComUsu.DestUsu).flags.UserLogged Then
                    Call SendData(SendTarget.toindex, UserList(userindex).ComUsu.DestUsu, 0, "||" & UserList(userindex).name & " ha rechazado tu oferta." & FONTTYPE_TALK)
                    Call FinComerciarUsu(UserList(userindex).ComUsu.DestUsu)
                End If
            End If
            Call SendData(SendTarget.toindex, userindex, 0, "||Has rechazado la oferta del otro usuario." & FONTTYPE_TALK)
            Call FinComerciarUsu(userindex)
            Exit Sub
        '[/Alejo]
    
    
    End Select
    
    Select Case UCase$(Left$(rData, 6))
    
Case "DONA01"
Dim MonturaDorada As Obj
MonturaDorada.Amount = 1
MonturaDorada.ObjIndex = 1054
 
If UserList(userindex).Stats.PuntosDonacion < 1500 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de donacion!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosDonacion = UserList(userindex).Stats.PuntosDonacion - 1500
Call SendData(SendTarget.toindex, userindex, 0, "||Tus puntos actuales son " & UserList(userindex).Stats.PuntosDonacion & FONTTYPE_INFO)
Call EnviarPuntosDonacion(userindex)
Call MeterItemEnInventario(userindex, MonturaDorada)
End If
Exit Sub

Case "DONA02"
Dim MonturaRoja As Obj
MonturaRoja.Amount = 1
MonturaRoja.ObjIndex = 1054
 
If UserList(userindex).Stats.PuntosDonacion < 1000 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de donacion!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosDonacion = UserList(userindex).Stats.PuntosDonacion - 1000
Call SendData(SendTarget.toindex, userindex, 0, "||Tus puntos actuales son " & UserList(userindex).Stats.PuntosDonacion & FONTTYPE_INFO)
Call EnviarPuntosDonacion(userindex)
Call MeterItemEnInventario(userindex, MonturaRoja)
End If
Exit Sub

Case "DONA03"
Dim TunicaChamp As Obj
TunicaChamp.Amount = 1
TunicaChamp.ObjIndex = 1055
 
If UserList(userindex).Stats.PuntosDonacion < 250 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de donacion!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosDonacion = UserList(userindex).Stats.PuntosDonacion - 250
Call SendData(SendTarget.toindex, userindex, 0, "||Tus puntos actuales son " & UserList(userindex).Stats.PuntosDonacion & FONTTYPE_INFO)
Call EnviarPuntosDonacion(userindex)
Call MeterItemEnInventario(userindex, TunicaChamp)
End If
Exit Sub

Case "DONA04"
Dim TunicaChampB As Obj
TunicaChampB.Amount = 1
TunicaChampB.ObjIndex = 1056
 
If UserList(userindex).Stats.PuntosDonacion < 250 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de donacion!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosDonacion = UserList(userindex).Stats.PuntosDonacion - 250
Call SendData(SendTarget.toindex, userindex, 0, "||Tus puntos actuales son " & UserList(userindex).Stats.PuntosDonacion & FONTTYPE_INFO)
Call EnviarPuntosDonacion(userindex)
Call MeterItemEnInventario(userindex, TunicaChampB)
End If
Exit Sub

Case "DONA05"
Dim TunicaHeroes As Obj
TunicaHeroes.Amount = 1
TunicaHeroes.ObjIndex = 1011
 
If UserList(userindex).Stats.PuntosDonacion < 500 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de donacion!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosDonacion = UserList(userindex).Stats.PuntosDonacion - 500
Call SendData(SendTarget.toindex, userindex, 0, "||Tus puntos actuales son " & UserList(userindex).Stats.PuntosDonacion & FONTTYPE_INFO)
Call EnviarPuntosDonacion(userindex)
Call MeterItemEnInventario(userindex, TunicaHeroes)
End If
Exit Sub

Case "DONA06"
Dim TunicaHeroesB As Obj
TunicaHeroesB.Amount = 1
TunicaHeroesB.ObjIndex = 1010
 
If UserList(userindex).Stats.PuntosDonacion < 500 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de donacion!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosDonacion = UserList(userindex).Stats.PuntosDonacion - 500
Call SendData(SendTarget.toindex, userindex, 0, "||Tus puntos actuales son " & UserList(userindex).Stats.PuntosDonacion & FONTTYPE_INFO)
Call EnviarPuntosDonacion(userindex)
Call MeterItemEnInventario(userindex, TunicaHeroesB)
End If
Exit Sub

Case "DONA07"
Dim TunicadelaLuz As Obj
TunicadelaLuz.Amount = 1
TunicadelaLuz.ObjIndex = 1050
 
If UserList(userindex).Stats.PuntosDonacion < 750 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de donacion!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosDonacion = UserList(userindex).Stats.PuntosDonacion - 750
Call SendData(SendTarget.toindex, userindex, 0, "||Tus puntos actuales son " & UserList(userindex).Stats.PuntosDonacion & FONTTYPE_INFO)
Call EnviarPuntosDonacion(userindex)
Call MeterItemEnInventario(userindex, TunicadelaLuz)
End If
Exit Sub

Case "DONA08"
Dim TunicadelaLuzB As Obj
TunicadelaLuzB.Amount = 1
TunicadelaLuzB.ObjIndex = 1049
 
If UserList(userindex).Stats.PuntosDonacion < 750 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de donacion!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosDonacion = UserList(userindex).Stats.PuntosDonacion - 750
Call SendData(SendTarget.toindex, userindex, 0, "||Tus puntos actuales son " & UserList(userindex).Stats.PuntosDonacion & FONTTYPE_INFO)
Call EnviarPuntosDonacion(userindex)
Call MeterItemEnInventario(userindex, TunicadelaLuzB)
End If
Exit Sub

Case "DONA09"
Dim TunicadelaOscuridad As Obj
TunicadelaOscuridad.Amount = 1
TunicadelaOscuridad.ObjIndex = 1051
 
If UserList(userindex).Stats.PuntosDonacion < 750 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de donacion!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosDonacion = UserList(userindex).Stats.PuntosDonacion - 750
Call SendData(SendTarget.toindex, userindex, 0, "||Tus puntos actuales son " & UserList(userindex).Stats.PuntosDonacion & FONTTYPE_INFO)
Call EnviarPuntosDonacion(userindex)
Call MeterItemEnInventario(userindex, TunicadelaOscuridad)
End If
Exit Sub

Case "DONA10"
Dim TunicadelaOscuridadB As Obj
TunicadelaOscuridadB.Amount = 1
TunicadelaOscuridadB.ObjIndex = 1052
 
If UserList(userindex).Stats.PuntosDonacion < 750 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de donacion!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosDonacion = UserList(userindex).Stats.PuntosDonacion - 750
Call SendData(SendTarget.toindex, userindex, 0, "||Tus puntos actuales son " & UserList(userindex).Stats.PuntosDonacion & FONTTYPE_INFO)
Call EnviarPuntosDonacion(userindex)
Call MeterItemEnInventario(userindex, TunicadelaOscuridadB)
End If
Exit Sub

Case "DONA11"
If UserList(userindex).Stats.PuntosDonacion < 2500 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de donacion!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosDonacion = UserList(userindex).Stats.PuntosDonacion - 2500
UserList(userindex).Stats.PuntosVIP = UserList(userindex).Stats.PuntosVIP + 10
Call SendData(SendTarget.toindex, userindex, 0, "||Tipea /VIP para completar la donacion." & FONTTYPE_INFO)
Call SendData(SendTarget.toindex, userindex, 0, "||Tus puntos actuales son " & UserList(userindex).Stats.PuntosDonacion & FONTTYPE_INFO)
Call EnviarPuntosDonacion(userindex)
End If
Exit Sub
End Select
    


    
    Select Case UCase$(Left$(rData, 2))
    '    Case "/Z"
    '        Dim Pos As WorldPos, Pos2 As WorldPos
    '        Dim O As Obj
    '
    '        For LoopC = 1 To 100
    '            Pos = UserList(UserIndex).Pos
    '            O.Amount = 1
    '            O.ObjIndex = iORO
    '            'Exit For
    '            Call TirarOro(100000, UserIndex)
    '            'Call Tilelibre(Pos, Pos2)
    '            'If Pos2.x = 0 Or Pos2.y = 0 Then Exit For
    '
    '            'Call MakeObj(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, O, Pos2.Map, Pos2.x, Pos2.y)
    '        Next LoopC
    '
    '        Exit Sub
    
        Case "TI" 'Tirar item
                If UserList(userindex).flags.Navegando = 1 Or _
                   UserList(userindex).flags.Muerto = 1 Or _
                   UserList(userindex).flags.Montando = 1 Or _
                   (UserList(userindex).flags.EsRolesMaster) Or _
                   UserList(userindex).flags.EnTD = 1 Or _
                   (UserList(userindex).flags.Privilegios = PlayerType.SemiDios And Not UserList(userindex).flags.EsRolesMaster) Then Exit Sub
                   '[VIPs]
                
                rData = Right$(rData, Len(rData) - 2)
                Arg1 = ReadField(1, rData, 44)
                Arg2 = ReadField(2, rData, 44)
                    If Val(Arg1) <= MAX_INVENTORY_SLOTS And Val(Arg1) > 0 Then
                        If UserList(userindex).Invent.Object(Val(Arg1)).ObjIndex = 0 Then
                                Exit Sub
                        End If
                        Call DropObj(userindex, Val(Arg1), Val(Arg2), UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y)
                    Else
                        Exit Sub
                    End If
                Exit Sub
        Case "DH" ' Lanzar hechizo
        If UserList(userindex).Lac.LLanzar.Puedo = False Then Exit Sub
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!!." & FONTTYPE_INFO)
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 2)
            UserList(userindex).flags.Hechizo = Val(rData)
            Exit Sub
        Case "LC" 'Click izquierdo
            rData = Right$(rData, Len(rData) - 2)
            Arg1 = ReadField(1, rData, 44)
            Arg2 = ReadField(2, rData, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
            X = CInt(Arg1)
            Y = CInt(Arg2)
            Call LookatTile(userindex, UserList(userindex).pos.Map, X, Y)
            Exit Sub
        Case "RC" 'Click derecho
            rData = Right$(rData, Len(rData) - 2)
            Arg1 = ReadField(1, rData, 44)
            Arg2 = ReadField(2, rData, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
            X = CInt(Arg1)
            Y = CInt(Arg2)
            Call Accion(userindex, UserList(userindex).pos.Map, X, Y)
            Exit Sub
        Case "YX"
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!!." & FONTTYPE_INFO)
                Exit Sub
            End If
    
            rData = Right$(rData, Len(rData) - 2)
            Select Case Val(rData)
                Case Robar
                    Call SendData(SendTarget.toindex, userindex, 0, "T01" & Robar)
                Case Magia
                    Call SendData(SendTarget.toindex, userindex, 0, "T01" & Magia)
                Case Domar
                    Call SendData(SendTarget.toindex, userindex, 0, "T01" & Domar)
                Case Ocultarse
                    If UserList(userindex).flags.Navegando = 1 Then
                        '[CDT 17-02-2004]
                        If Not UserList(userindex).flags.UltimoMensaje = 3 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "||No podes ocultarte si estas navegando." & FONTTYPE_INFO)
                            UserList(userindex).flags.UltimoMensaje = 3
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                    
                    If UserList(userindex).flags.Montando = 1 Then
                              If Not UserList(userindex).flags.UltimoMensaje = 3 Then
                                Call SendData(toindex, userindex, 0, "||No podes ocultarte si estas sobre una montura." & FONTTYPE_INFO)
                                UserList(userindex).flags.UltimoMensaje = 3
                              End If
                          Exit Sub
                    End If
                    
                    If UserList(userindex).flags.Oculto = 1 Then
                        '[CDT 17-02-2004]
                        If Not UserList(userindex).flags.UltimoMensaje = 2 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas oculto." & FONTTYPE_INFO)
                            UserList(userindex).flags.UltimoMensaje = 2
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                    
                    Call DoOcultarse(userindex)
            End Select
            Exit Sub
    
    End Select
    
    Select Case UCase$(Left$(rData, 3))
         Case "UMH" ' Usa macro de hechizos
            Call SendData(SendTarget.ToAdmins, userindex, 0, "||" & UserList(userindex).name & " fue expulsado por Anti-macro de hechizos " & FONTTYPE_VENENO)
            Call SendData(SendTarget.toindex, userindex, 0, "ERR Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros" & FONTTYPE_INFO)
            Call CloseSocket(userindex)
            Exit Sub
        Case "KLQ"
            rData = Right$(rData, Len(rData) - 3)
            If Val(rData) <= MAX_INVENTORY_SLOTS And Val(rData) > 0 Then
                If UserList(userindex).Invent.Object(Val(rData)).ObjIndex = 0 Then Exit Sub
            Else
                Exit Sub
            End If
            Call UseInvItem(userindex, Val(rData))
            Exit Sub
        Case "CNS" ' Construye herreria
            rData = Right$(rData, Len(rData) - 3)
            X = CInt(rData)
            If X < 1 Then Exit Sub
            If ObjData(X).SkHerreria = 0 Then Exit Sub
            Call HerreroConstruirItem(userindex, X)
            Exit Sub
        Case "CNC" ' Construye carpinteria
            rData = Right$(rData, Len(rData) - 3)
            X = CInt(rData)
            If X < 1 Or ObjData(X).SkCarpinteria = 0 Then Exit Sub
            Call CarpinteroConstruirItem(userindex, X)
            Exit Sub
        Case "WLC" 'Click izquierdo en modo trabajo
            rData = Right$(rData, Len(rData) - 3)
            Arg1 = ReadField(1, rData, 44)
            Arg2 = ReadField(2, rData, 44)
            Arg3 = ReadField(3, rData, 44)
            If Arg3 = "" Or Arg2 = "" Or Arg1 = "" Then Exit Sub
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Or Not Numeric(Arg3) Then Exit Sub
            
            X = CInt(Arg1)
            Y = CInt(Arg2)
            tLong = CInt(Arg3)
            
            If UserList(userindex).flags.Muerto = 1 Or _
               UserList(userindex).flags.Descansar Or _
               UserList(userindex).flags.Meditando Or _
               Not InMapBounds(UserList(userindex).pos.Map, X, Y) Then Exit Sub
            
            If Not InRangoVision(userindex, X, Y) Then
                Call SendData(SendTarget.toindex, userindex, 0, "PU" & UserList(userindex).pos.X & "," & UserList(userindex).pos.Y)
                Exit Sub
            End If
            
            Select Case tLong
            
            Case Proyectiles
                Dim TU As Integer, tN As Integer
                'Nos aseguramos que este usando un arma de proyectiles
                If Not IntervaloPermiteAtacar(userindex, False) Or Not IntervaloPermiteUsarArcos(userindex) Then
                    Exit Sub
                End If

                DummyInt = 0

                If UserList(userindex).Invent.WeaponEqpObjIndex = 0 Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.WeaponEqpSlot < 1 Or UserList(userindex).Invent.WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.MunicionEqpSlot < 1 Or UserList(userindex).Invent.MunicionEqpSlot > MAX_INVENTORY_SLOTS Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.MunicionEqpObjIndex = 0 Then
                    DummyInt = 1
                ElseIf ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).proyectil <> 1 Then
                    DummyInt = 2
                ElseIf ObjData(UserList(userindex).Invent.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
                    DummyInt = 1
                ElseIf UserList(userindex).Invent.Object(UserList(userindex).Invent.MunicionEqpSlot).Amount < 1 Then
                    DummyInt = 1
                End If
                
                If DummyInt <> 0 Then
                    If DummyInt = 1 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||No tenes municiones." & FONTTYPE_INFO)
                    End If
                    Call Desequipar(userindex, UserList(userindex).Invent.MunicionEqpSlot)
                    Call Desequipar(userindex, UserList(userindex).Invent.WeaponEqpSlot)
                    Exit Sub
                End If
                
                DummyInt = 0
                'Quitamos stamina
                If UserList(userindex).Stats.MinSta >= 10 Then
                     Call QuitarSta(userindex, RandomNumber(1, 10))
                Else
                     Call SendData(SendTarget.toindex, userindex, 0, "||Estas muy cansado para luchar." & FONTTYPE_INFO)
                     Exit Sub
                End If
                 
                Call LookatTile(userindex, UserList(userindex).pos.Map, Arg1, Arg2)
                
                TU = UserList(userindex).flags.TargetUser
                tN = UserList(userindex).flags.TargetNPC
                
                'Sólo permitimos atacar si el otro nos puede atacar también
                If TU > 0 Then
                    If Abs(UserList(UserList(userindex).flags.TargetUser).pos.Y - UserList(userindex).pos.Y) > RANGO_VISION_Y Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos para atacar." & FONTTYPE_WARNING)
                        Exit Sub
                    End If
                ElseIf tN > 0 Then
                    If Abs(Npclist(UserList(userindex).flags.TargetNPC).pos.Y - UserList(userindex).pos.Y) > RANGO_VISION_Y Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos para atacar." & FONTTYPE_WARNING)
                        Exit Sub
                    End If
                End If
                
                
                If TU > 0 Then
                    'Previene pegarse a uno mismo
                    If TU = userindex Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes atacarte a vos mismo!" & FONTTYPE_INFO)
                        DummyInt = 1
                        Exit Sub
                    End If
                End If
    
                If DummyInt = 0 Then
                    'Saca 1 flecha
                    DummyInt = UserList(userindex).Invent.MunicionEqpSlot
                    Call QuitarUserInvItem(userindex, UserList(userindex).Invent.MunicionEqpSlot, 1)
                    If DummyInt < 1 Or DummyInt > MAX_INVENTORY_SLOTS Then Exit Sub
                    If UserList(userindex).Invent.Object(DummyInt).Amount > 0 Then
                        UserList(userindex).Invent.Object(DummyInt).Equipped = 1
                        UserList(userindex).Invent.MunicionEqpSlot = DummyInt
                        UserList(userindex).Invent.MunicionEqpObjIndex = UserList(userindex).Invent.Object(DummyInt).ObjIndex
                        Call UpdateUserInv(False, userindex, UserList(userindex).Invent.MunicionEqpSlot)
                    Else
                        Call UpdateUserInv(False, userindex, DummyInt)
                        UserList(userindex).Invent.MunicionEqpSlot = 0
                        UserList(userindex).Invent.MunicionEqpObjIndex = 0
                    End If
                    '-----------------------------------
                End If

                If tN > 0 Then
                    If Npclist(tN).Attackable <> 0 Then
                        Call UsuarioAtacaNpc(userindex, tN)
                    End If
                ElseIf TU > 0 Then
                        If Ciudadano(TU) And Ciudadano(userindex) And TriggerZonaPelea(TU, userindex) <> TRIGGER6_PERMITE And UserList(userindex).pos.Map <> 31 And UserList(userindex).pos.Map <> 32 And UserList(userindex).pos.Map <> 33 And UserList(userindex).pos.Map <> 34 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes atacar ciudadanos, escribe /RENUNCIAR." & FONTTYPE_FIGHT)
                            Exit Sub
                        End If
                    Call UsuarioAtacaUsuario(userindex, TU)
                End If
                
            Case Magia
                If MapInfo(UserList(userindex).pos.Map).MagiaSinEfecto > 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||Una fuerza oscura te impide canalizar tu energía" & FONTTYPE_FIGHT)
                    Exit Sub
                End If
                Call LookatTile(userindex, UserList(userindex).pos.Map, X, Y)
                
                'MmMmMmmmmM
                Dim wp2 As WorldPos
                wp2.Map = UserList(userindex).pos.Map
                wp2.X = X
                wp2.Y = Y
                                
                If UserList(userindex).flags.Hechizo > 0 Then
                    If IntervaloPermiteLanzarSpell(userindex) Then
                        Call LanzarHechizo(UserList(userindex).flags.Hechizo, userindex)
                    '    UserList(UserIndex).flags.PuedeLanzarSpell = 0
                        UserList(userindex).flags.Hechizo = 0
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡Primero selecciona el hechizo que quieres lanzar!" & FONTTYPE_INFO)
                End If
                
                'If Distancia(UserList(UserIndex).Pos, wp2) > 10 Then
                If (Abs(UserList(userindex).pos.X - wp2.X) > 9 Or Abs(UserList(userindex).pos.Y - wp2.Y) > 8) Then
                    Dim txt As String
                    txt = "Ataque fuera de rango de " & UserList(userindex).name & "(" & UserList(userindex).pos.Map & "/" & UserList(userindex).pos.X & "/" & UserList(userindex).pos.Y & ") ip: " & UserList(userindex).ip & " a la posicion (" & wp2.Map & "/" & wp2.X & "/" & wp2.Y & ") "
                    If UserList(userindex).flags.Hechizo > 0 Then
                        txt = txt & ". Hechizo: " & Hechizos(UserList(userindex).Stats.UserHechizos(UserList(userindex).flags.Hechizo)).Nombre
                    End If
                    If MapData(wp2.Map, wp2.X, wp2.Y).userindex > 0 Then
                        txt = txt & " hacia el usuario: " & UserList(MapData(wp2.Map, wp2.X, wp2.Y).userindex).name
                    ElseIf MapData(wp2.Map, wp2.X, wp2.Y).NpcIndex > 0 Then
                        txt = txt & " hacia el NPC: " & Npclist(MapData(wp2.Map, wp2.X, wp2.Y).NpcIndex).name
                    End If
                    
                    Call LogCheating(txt)
                End If
                
            
            
            
            Case Pesca
                        
                AuxInd = UserList(userindex).Invent.HerramientaEqpObjIndex
                If AuxInd = 0 Then Exit Sub
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub
                
                If AuxInd <> CAÑA_PESCA And AuxInd <> RED_PESCA Then
                    'Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                'Basado en la idea de Barrin
                'Comentario por Barrin: jah, "basado", caradura ! ^^
                If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes pescar desde donde te encuentras." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If HayAgua(UserList(userindex).pos.Map, X, Y) Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_PESCAR)
                    
                    Select Case AuxInd
                    Case CAÑA_PESCA
                        Call DoPescar(userindex)
                    Case RED_PESCA
                        With UserList(userindex)
                            wpaux.Map = .pos.Map
                            wpaux.X = X
                            wpaux.Y = Y
                        End With
                        
                        If Distancia(UserList(userindex).pos, wpaux) > 2 Then
                            Call SendData(SendTarget.toindex, userindex, 0, "||Estás demasiado lejos para pescar." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        Call DoPescarRed(userindex)
                    End Select
    
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||No hay agua donde pescar busca un lago, rio o mar." & FONTTYPE_INFO)
                End If
                
            Case Robar
               If MapInfo(UserList(userindex).pos.Map).Pk Then
                    'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                    If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub
                    
                    Call LookatTile(userindex, UserList(userindex).pos.Map, X, Y)
                    
                    If UserList(userindex).flags.TargetUser > 0 And UserList(userindex).flags.TargetUser <> userindex Then
                       If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 0 Then
                            wpaux.Map = UserList(userindex).pos.Map
                            wpaux.X = Val(ReadField(1, rData, 44))
                            wpaux.Y = Val(ReadField(2, rData, 44))
                            If Distancia(wpaux, UserList(userindex).pos) > 2 Then
                                Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                                Exit Sub
                            End If
                            '17/09/02
                            'No aseguramos que el trigger le permite robar
                            If MapData(UserList(UserList(userindex).flags.TargetUser).pos.Map, UserList(UserList(userindex).flags.TargetUser).pos.X, UserList(UserList(userindex).flags.TargetUser).pos.Y).trigger = eTrigger.ZONASEGURA Then
                                Call SendData(SendTarget.toindex, userindex, 0, "||No podes robar aquí." & FONTTYPE_WARNING)
                                Exit Sub
                            End If
                            If MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).trigger = eTrigger.ZONASEGURA Then
                                Call SendData(SendTarget.toindex, userindex, 0, "||No podes robar aquí." & FONTTYPE_WARNING)
                                Exit Sub
                            End If
                            
                            Call DoRobar(userindex, UserList(userindex).flags.TargetUser)
                       End If
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||No a quien robarle!." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡No podes robarle en zonas seguras!." & FONTTYPE_INFO)
                End If
            Case Talar
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub
                
                If UserList(userindex).Invent.HerramientaEqpObjIndex = 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||Deberías equiparte el hacha." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(userindex).Invent.HerramientaEqpObjIndex <> HACHA_LEÑADOR Then
                    ' Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                AuxInd = MapData(UserList(userindex).pos.Map, X, Y).OBJInfo.ObjIndex
                If AuxInd > 0 Then
                    wpaux.Map = UserList(userindex).pos.Map
                    wpaux.X = X
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(userindex).pos) > 2 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Barrin 29/9/03
                    If Distancia(wpaux, UserList(userindex).pos) = 0 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||No podes talar desde allí." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    '¿Hay un arbol donde clickeo?
                    If ObjData(AuxInd).OBJType = eOBJType.otArboles Then
                        Call SendData(SendTarget.ToPCArea, CInt(userindex), UserList(userindex).pos.Map, "TW" & SND_TALAR)
                        Call DoTalar(userindex)
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||No hay ningun arbol ahi." & FONTTYPE_INFO)
                End If
            Case Mineria
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub
                                
                If UserList(userindex).Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                
                If UserList(userindex).Invent.HerramientaEqpObjIndex <> PIQUETE_MINERO Then
                    ' Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                Call LookatTile(userindex, UserList(userindex).pos.Map, X, Y)
                
                AuxInd = MapData(UserList(userindex).pos.Map, X, Y).OBJInfo.ObjIndex
                If AuxInd > 0 Then
                    wpaux.Map = UserList(userindex).pos.Map
                    wpaux.X = X
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(userindex).pos) > 2 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                    '¿Hay un yacimiento donde clickeo?
                    If ObjData(AuxInd).OBJType = eOBJType.otYacimiento Then
                        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TW" & SND_MINERO)
                        Call DoMineria(userindex)
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||Ahi no hay ningun yacimiento." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||Ahi no hay ningun yacimiento." & FONTTYPE_INFO)
                End If
            Case Domar
              'Modificado 25/11/02
              'Optimizado y solucionado el bug de la doma de
              'criaturas hostiles.
              Dim CI As Integer
              
              Call LookatTile(userindex, UserList(userindex).pos.Map, X, Y)
              CI = UserList(userindex).flags.TargetNPC
              
              If CI > 0 Then
                       If Npclist(CI).flags.Domable > 0 Then
                            wpaux.Map = UserList(userindex).pos.Map
                            wpaux.X = X
                            wpaux.Y = Y
                            If Distancia(wpaux, Npclist(UserList(userindex).flags.TargetNPC).pos) > 2 Then
                                  Call SendData(SendTarget.toindex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                                  Exit Sub
                            End If
                            If Npclist(CI).flags.AttackedBy <> "" Then
                                  Call SendData(SendTarget.toindex, userindex, 0, "||No podés domar una criatura que está luchando con un jugador." & FONTTYPE_INFO)
                                  Exit Sub
                            End If
                            Call DoDomar(userindex, CI)
                        Else
                            Call SendData(SendTarget.toindex, userindex, 0, "||No podes domar a esa criatura." & FONTTYPE_INFO)
                        End If
              Else
                     Call SendData(SendTarget.toindex, userindex, 0, "||No hay ninguna criatura alli!." & FONTTYPE_INFO)
              End If
              
            Case FundirMetal
                'Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                If Not IntervaloPermiteTrabajar(userindex) Then Exit Sub
                
                If UserList(userindex).flags.TargetObj > 0 Then
                    If ObjData(UserList(userindex).flags.TargetObj).OBJType = eOBJType.otFragua Then
                        ''chequeamos que no se zarpe duplicando oro
                        If UserList(userindex).Invent.Object(UserList(userindex).flags.TargetObjInvSlot).ObjIndex <> UserList(userindex).flags.TargetObjInvIndex Then
                            If UserList(userindex).Invent.Object(UserList(userindex).flags.TargetObjInvSlot).ObjIndex = 0 Or UserList(userindex).Invent.Object(UserList(userindex).flags.TargetObjInvSlot).Amount = 0 Then
                                Call SendData(SendTarget.toindex, userindex, 0, "||No tienes mas minerales" & FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                    
                            Call SendData(SendTarget.toindex, userindex, 0, "ERRHas sido expulsado por el sistema anti cheats. Reconéctate.")
                            Call CloseSocket(userindex)
                            Exit Sub
                        End If
                        Call FundirMineral(userindex)
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||Ahi no hay ninguna fragua." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||Ahi no hay ninguna fragua." & FONTTYPE_INFO)
                End If
                
            Case Herreria
                Call LookatTile(userindex, UserList(userindex).pos.Map, X, Y)
                
                If UserList(userindex).flags.TargetObj > 0 Then
                    If ObjData(UserList(userindex).flags.TargetObj).OBJType = eOBJType.otYunque Then
                        Call EnivarArmasConstruibles(userindex)
                        Call EnivarArmadurasConstruibles(userindex)
                        Call SendData(SendTarget.toindex, userindex, 0, "SFH")
                    Else
                        Call SendData(SendTarget.toindex, userindex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)
                End If
                
            End Select
            
            'UserList(UserIndex).flags.PuedeTrabajar = 0
            Exit Sub
        Case "CIG"
            rData = Right$(rData, Len(rData) - 3)
            
            If modGuilds.CrearNuevoClan(rData, userindex, UserList(userindex).FundandoGuildAlineacion, tStr) Then
                Call SendData(SendTarget.toall, 0, 0, "||" & UserList(userindex).name & " fundó el clan " & Guilds(UserList(userindex).GuildIndex).GuildName & " de alineación " & Alineacion2String(Guilds(UserList(userindex).GuildIndex).Alineacion) & "." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            End If
            
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rData, 4))
    Case "PCCC" 'Te veo el caption jaja esa eM
            Dim caption As String
            rData = Right$(rData, Len(rData) - 4)
            caption = ReadField(1, rData, 44)
            tIndex = ReadField(2, rData, 44)
            Call SendData(SendTarget.toindex, tIndex, 0, "PCCC" & caption & "," & UserList(userindex).name)
            Exit Sub
        Case "DRAG"
            rData = Right$(rData, Len(rData) - 4)
            ObjSlot1 = ReadField(1, rData, 44)
            ObjSlot2 = ReadField(2, rData, 44)
            DragObjects (userindex)
            Exit Sub
        Case "INFS" 'Informacion del hechizo
                rData = Right$(rData, Len(rData) - 4)
                If Val(rData) > 0 And Val(rData) < MAXUSERHECHIZOS + 1 Then
                    Dim h As Integer
                    h = UserList(userindex).Stats.UserHechizos(Val(rData))
                    If h > 0 And h < NumeroHechizos + 1 Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & FONTTYPE_INFO)
                        Call SendData(SendTarget.toindex, userindex, 0, "||Nombre:" & Hechizos(h).Nombre & FONTTYPE_INFO)
                        Call SendData(SendTarget.toindex, userindex, 0, "||Descripcion:" & Hechizos(h).Desc & FONTTYPE_INFO)
                        Call SendData(SendTarget.toindex, userindex, 0, "||Skill requerido: " & Hechizos(h).MinSkill & " de magia." & FONTTYPE_INFO)
                        Call SendData(SendTarget.toindex, userindex, 0, "||Mana necesario: " & Hechizos(h).ManaRequerido & FONTTYPE_INFO)
                        Call SendData(SendTarget.toindex, userindex, 0, "||Stamina necesaria: " & Hechizos(h).StaRequerido & FONTTYPE_INFO)
                        Call SendData(SendTarget.toindex, userindex, 0, "||%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%" & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡Primero selecciona el hechizo.!" & FONTTYPE_INFO)
                End If
                Exit Sub
        Case "KHEV"
                If UserList(userindex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                    Exit Sub
                End If
                rData = Right$(rData, Len(rData) - 4)
                If Val(rData) <= MAX_INVENTORY_SLOTS And Val(rData) > 0 Then
                     If UserList(userindex).Invent.Object(Val(rData)).ObjIndex = 0 Then Exit Sub
                Else
                    Exit Sub
                End If
                Call EquiparInvItem(userindex, Val(rData))
                Exit Sub
        Case "CHEA" 'Cambiar Heading ;-)
            rData = Right$(rData, Len(rData) - 4)
            If Val(rData) > 0 And Val(rData) < 5 Then
                UserList(userindex).Char.Heading = rData
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(userindex).pos.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
            End If
            Exit Sub
        Case "SKSE" 'Modificar skills
            Dim sumatoria As Integer
            Dim incremento As Integer
            rData = Right$(rData, Len(rData) - 4)
            
            'Codigo para prevenir el hackeo de los skills
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            For i = 1 To NUMSKILLS
                incremento = Val(ReadField(i, rData, 44))
                
                If incremento < 0 Then
                    'Call SendData(SendTarget.ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                    Call LogHackAttemp(UserList(userindex).name & " IP:" & UserList(userindex).ip & " trato de hackear los skills.")
                    UserList(userindex).Stats.SkillPts = 0
                    Call CloseSocket(userindex)
                    Exit Sub
                End If
                
                sumatoria = sumatoria + incremento
            Next i
            
            If sumatoria > UserList(userindex).Stats.SkillPts Then
                'UserList(UserIndex).Flags.AdministrativeBan = 1
                'Call SendData(SendTarget.ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                Call LogHackAttemp(UserList(userindex).name & " IP:" & UserList(userindex).ip & " trato de hackear los skills.")
                Call CloseSocket(userindex)
                Exit Sub
            End If
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            
            For i = 1 To NUMSKILLS
                incremento = Val(ReadField(i, rData, 44))
                UserList(userindex).Stats.SkillPts = UserList(userindex).Stats.SkillPts - incremento
                UserList(userindex).Stats.UserSkills(i) = UserList(userindex).Stats.UserSkills(i) + incremento
                If UserList(userindex).Stats.UserSkills(i) > 100 Then UserList(userindex).Stats.UserSkills(i) = 100
            Next i
            Exit Sub
        Case "ENTR" 'Entrena hombre!
            
            If UserList(userindex).flags.TargetNPC = 0 Then Exit Sub
            
            If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 3 Then Exit Sub
            
            rData = Right$(rData, Len(rData) - 4)
            
            If Npclist(UserList(userindex).flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
                If Val(rData) > 0 And Val(rData) < Npclist(UserList(userindex).flags.TargetNPC).NroCriaturas + 1 Then
                        Dim SpawnedNpc As Integer
                        SpawnedNpc = SpawnNpc(Npclist(UserList(userindex).flags.TargetNPC).Criaturas(Val(rData)).NpcIndex, Npclist(UserList(userindex).flags.TargetNPC).pos, True, False)
                        If SpawnedNpc > 0 Then
                            Npclist(SpawnedNpc).MaestroNpc = UserList(userindex).flags.TargetNPC
                            Npclist(UserList(userindex).flags.TargetNPC).Mascotas = Npclist(UserList(userindex).flags.TargetNPC).Mascotas + 1
                        End If
                End If
            Else
                Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & vbWhite & "°" & "No puedo traer mas criaturas, mata las existentes!" & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
            End If
            
            Exit Sub
        Case "COMP"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).Comercia = 0 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 5)
            'User compra el item del slot rdata
            If UserList(userindex).flags.Comerciando = False Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No estas comerciando " & FONTTYPE_INFO)
                Exit Sub
            End If
            'listindex+1, cantidad
            Call NPCVentaItem(userindex, Val(ReadField(1, rData, 44)), Val(ReadField(2, rData, 44)), UserList(userindex).flags.TargetNPC)
            Exit Sub
        '[KEVIN]*********************************************************************
        '------------------------------------------------------------------------------------
        Case "RETI"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(userindex).flags.Muerto = 1 Then
                       Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                       Exit Sub
             End If
             '¿El target es un NPC valido?
             If UserList(userindex).flags.TargetNPC > 0 Then
                   '¿Es el banquero?
                   If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> 4 Then
                       Exit Sub
                   End If
             Else
               Exit Sub
             End If
             rData = Right(rData, Len(rData) - 5)
             'User retira el item del slot rdata
             Call UserRetiraItem(userindex, Val(ReadField(1, rData, 44)), Val(ReadField(2, rData, 44)))
             Exit Sub
        '-----------------------------------------------------------------------------------
        '[/KEVIN]****************************************************************************
        Case "VEND"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 5)
            '¿El target es un NPC valido?
            tInt = Val(ReadField(1, rData, 44))
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).Comercia = 0 Then
                    Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(userindex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
'           rdata = Right$(rdata, Len(rdata) - 5)
            'User compra el item del slot rdata
            Call NPCCompraItem(userindex, Val(ReadField(1, rData, 44)), Val(ReadField(2, rData, 44)))
            Exit Sub
        '[KEVIN]-------------------------------------------------------------------------
        '****************************************************************************************
        Case "DEPO"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(userindex).flags.Muerto = 1 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(userindex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(userindex).flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            rData = Right(rData, Len(rData) - 5)
            'User deposita el item del slot rdata
            Call UserDepositaItem(userindex, Val(ReadField(1, rData, 44)), Val(ReadField(2, rData, 44)))
            Exit Sub
        '****************************************************************************************
        '[/KEVIN]---------------------------------------------------------------------------------
    End Select

    Select Case UCase$(Left$(rData, 5))
     Case "JKNCM"
        UserList(userindex).flags.ClienteValido = 1
        Exit Sub
        
        Case "DEMSG"
            If UserList(userindex).flags.TargetObj > 0 Then
            rData = Right$(rData, Len(rData) - 5)
            Dim f As String, Titu As String, msg As String, f2 As String
            f = App.Path & "\foros\"
            f = f & UCase$(ObjData(UserList(userindex).flags.TargetObj).ForoID) & ".for"
            Titu = ReadField(1, rData, 176)
            msg = ReadField(2, rData, 176)
            Dim n2 As Integer, loopme As Integer
            If FileExist(f, vbNormal) Then
                Dim num As Integer
                num = Val(GetVar(f, "INFO", "CantMSG"))
                If num > MAX_MENSAJES_FORO Then
                    For loopme = 1 To num
                        Kill App.Path & "\foros\" & UCase$(ObjData(UserList(userindex).flags.TargetObj).ForoID) & loopme & ".for"
                    Next
                    Kill App.Path & "\foros\" & UCase$(ObjData(UserList(userindex).flags.TargetObj).ForoID) & ".for"
                    num = 0
                End If
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & num + 1 & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", num + 1)
            Else
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & "1" & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", 1)
            End If
            Close #n2
            End If
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rData, 6))
       'Standelf Viajes:
    Case "TRAVEL"
        rData = Right(rData, Len(rData) - 6)
                Dim Destino As String
                    Dim DestMapa As Integer
                Dim DestX As Integer
            Dim DestY As Integer
            
            If UserList(userindex).Counters.Pena > 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No puedes utilizar este comando si estas en la carcel." & FONTTYPE_INFO)
                Exit Sub
            End If
            If MapInfo(UserList(userindex).pos.Map).Pk = True Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Estas en zona insegura, desde aquí no puedes viajar." & FONTTYPE_WARNING)
                Exit Sub
            End If
            
        If Val(rData) = 1 Then
                Destino = "Runek"
                    DestMapa = 1
                DestX = 58
            DestY = 45
            
        ElseIf Val(rData) = 2 Then
                Destino = "Banderbill"
                    DestMapa = 59
                DestX = 50
            DestY = 50
            
        ElseIf Val(rData) = 3 Then
                Destino = "Lindos"
                    DestMapa = 62
                DestX = 72
            DestY = 41
            
        ElseIf Val(rData) = 4 Then
                Destino = "Helkat"
                    DestMapa = 34
                DestX = 44
            DestY = 88
        End If
            
                    Call SendUserStatsBox(userindex)
                    Call SendData(SendTarget.toindex, userindex, 0, "||Has viajado a " & Destino & FONTTYPE_INFO)
                Call WarpUserChar(userindex, DestMapa, DestX, DestY, True)
        Exit Sub
        
        Case "TNOBLE"
If UserList(userindex).Stats.TransformadoVIP = 1 Then
UserList(userindex).Stats.TransformadoVIP = 0
UserList(userindex).Char.CascoAnim = 0
Call SendData(SendTarget.toindex, userindex, 0, "||VIP desactivado." & FONTTYPE_TALK)
UserList(userindex).flags.Privilegios = 0
Else
UserList(userindex).Stats.TransformadoVIP = 1
UserList(userindex).Char.CascoAnim = 32
Call SendData(SendTarget.toindex, userindex, 0, "||VIP activado." & FONTTYPE_TALK)
UserList(userindex).flags.Privilegios = 1
End If
Call WarpUserChar(userindex, UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y, False)
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "XFC" & UserList(userindex).Char.CharIndex & "," & FXIDs.FXMEDITARVIPW & ",")
Exit Sub

        Case "DESPHE" 'Mover Hechizo de lugar
            rData = Right(rData, Len(rData) - 6)
            Call DesplazarHechizo(userindex, CInt(ReadField(1, rData, 44)), CInt(ReadField(2, rData, 44)))
            Exit Sub
        Case "DESCOD" 'Informacion del hechizo
                rData = Right$(rData, Len(rData) - 6)
                Call modGuilds.ActualizarCodexYDesc(rData, UserList(userindex).GuildIndex)
                Exit Sub
    End Select
    
    '[Alejo]
    
        Select Case UCase$(Left$(rData, 7))
    
     Case "NANVAME"
            rData = Right(rData, Len(rData) - 7)

            Call SendData(SendTarget.ToAdmins, 0, 0, "||Anti Cheat> " & UserList(userindex).name & " hay posibilidades de que este usando Speed Hack, revisenlon! " & FONTTYPE_SERVER)

            Exit Sub
            
            
     Case "BANEAME"
            rData = Right(rData, Len(rData) - 7)
            tStr = UserList(userindex).name 'Nick
            h = FreeFile
            Open App.Path & "\LOGS\CHEATERS.log" For Append Shared As h
            
            Print #h, "########################################################################"
            Print #h, "Usuario: " & UserList(userindex).name
            Print #h, "Fecha: " & Date
            Print #h, "Hora: " & Time
            Print #h, "CHEAT: " & rData
            Print #h, "########################################################################"
            Print #h, " "
            Close #h
            
            UserList(userindex).flags.Ban = 1
            
            tInt = Val(GetVar(CharPath & tStr & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "Cant", tInt + 1)
            Call WriteVar(CharPath & tStr & ".chr", "PENAS", "P" & tInt + 1, "Anti Cheat> te ha baneado por uso de " & rData & " " & Date & " " & Time)
        
            'Avisamos a los admins
            Call SendData(SendTarget.ToAdmins, 0, 0, "||Anti Cheat> " & UserList(userindex).name & " ha sido Baneado por uso de " & rData & FONTTYPE_SERVER)
            Call CloseSocket(userindex)
            Exit Sub
            
Case "CANJE01"
Dim PremioObj As Obj
PremioObj.Amount = 1
PremioObj.ObjIndex = 868

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 100 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 100
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 100 Pts. de Torneo y ahora tienes una Corona!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub
 
Case "CANJE02"
PremioObj.Amount = 1
PremioObj.ObjIndex = 865

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 200 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 200
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 200 Pts. de Torneo y ahora tienes un Manto Alado!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub
 
Case "CANJE03"
PremioObj.Amount = 1
PremioObj.ObjIndex = 864

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 200 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 200
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 200 Pts. de Torneo y ahora tienes un Manto Alado (E/G)!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub
 
Case "CANJE04"
PremioObj.Amount = 1
PremioObj.ObjIndex = 859

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 180 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 180
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 180 Pts. de Torneo y ahora tienes una Túnica Apocaliptica!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub
 
Case "CANJE05"
PremioObj.Amount = 1
PremioObj.ObjIndex = 862

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If

 
If UserList(userindex).Stats.PuntosTorneo < 180 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 180
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 180 Pts. de Torneo y ahora tienes una Túnica Apocaliptica (E/G)!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE06"
PremioObj.Amount = 1
PremioObj.ObjIndex = 945

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 120 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 120
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 120 Pts. de Torneo y ahora tienes un Báculo Divino!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE07"
PremioObj.Amount = 1
PremioObj.ObjIndex = 857

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 140 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 140
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 140 Pts. de Torneo y ahora tienes una Espada Barlog!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE08"
PremioObj.Amount = 1
PremioObj.ObjIndex = 888

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 90 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 90
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 90 Pts. de Torneo y ahora tienes una Espada Argentum!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE09"
PremioObj.Amount = 1
PremioObj.ObjIndex = 903

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 90 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 90
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 90 Pts. de Torneo y ahora tienes una Daga Infernal!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE10"
PremioObj.Amount = 1
PremioObj.ObjIndex = 854

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 60 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 60
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 60 Pts. de Torneo y ahora tienes una Espada de las Almas!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE11"
PremioObj.Amount = 1
PremioObj.ObjIndex = 863

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 120 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 120
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 120 Pts. de Torneo y ahora tienes un Arco Èlfico!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE12"
PremioObj.Amount = 1
PremioObj.ObjIndex = 946

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 150 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 150
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 150 Pts. de Torneo y ahora tienes un Cetro Perfecto!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE13"
PremioObj.Amount = 1
PremioObj.ObjIndex = 925

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 150 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 150
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 150 Pts. de Torneo y ahora tienes una Armadura Ancestral!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE14"
PremioObj.Amount = 1
PremioObj.ObjIndex = 924

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 150 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 150
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 150 Pts. de Torneo y ahora tienes una Armadura Ancestral (E/G)!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE15"
PremioObj.Amount = 1
PremioObj.ObjIndex = 889

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 190 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 190
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 190 Pts. de Torneo y ahora tienes una Coraza del Mal!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE16"
PremioObj.Amount = 1
PremioObj.ObjIndex = 949

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 190 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 190
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 190 Pts. de Torneo y ahora tienes una Coraza del Mal (E/G)!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE17"
PremioObj.Amount = 1
PremioObj.ObjIndex = 931

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 230 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 230
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 230 Pts. de Torneo y ahora tienes una Armadura Diabólica (E/G)!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE18"
PremioObj.Amount = 1
PremioObj.ObjIndex = 930

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 210 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 210
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 210 Pts. de Torneo y ahora tienes una Armadura Extrema (E/G)!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE19"
PremioObj.Amount = 1
PremioObj.ObjIndex = 879

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 150 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 150
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 150 Pts. de Torneo y ahora tienes un Anillo Divino!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE20"
PremioObj.Amount = 1
PremioObj.ObjIndex = 939

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 350 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 350
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 350 Pts. de Torneo y ahora tienes un Talisman de Lider!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE21"
PremioObj.Amount = 1
PremioObj.ObjIndex = 936

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 50 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 50
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 50 Pts. de Torneo y ahora tienes un Pendiente del Sacrificio!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE22"
PremioObj.Amount = 1
PremioObj.ObjIndex = 1026

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 110 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 110
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 110 Pts. de Torneo y ahora tienes un Escudo de Dragón!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub

Case "CANJE23"
PremioObj.Amount = 1
PremioObj.ObjIndex = 1027

If UserList(userindex).flags.Privilegios = PlayerType.SemiDios Then
            Call SendData(SendTarget.toindex, userindex, 0, "||No puedes canjear items." & FONTTYPE_INFO)
            Exit Sub
            End If
 
If UserList(userindex).Stats.PuntosTorneo < 140 Then
Call SendData(SendTarget.toindex, userindex, 0, "||No tienes suficientes puntos de torneo!." & FONTTYPE_INFO)
Else
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 140
Call SendData(SendTarget.toindex, userindex, 0, "||Se te han descontado 140 Pts. de Torneo y ahora tienes un Casco Siniestro!." & FONTTYPE_INFO)
Call EnviarPuntos(userindex)
    If Not MeterItemEnInventario(userindex, PremioObj) Then
                    Call TirarItemAlPiso(UserList(userindex).pos, PremioObj)
    End If
End If
Exit Sub
 
End Select
    
    Select Case UCase$(Left$(rData, 7))
    
    Case "OFRECER"
            rData = Right$(rData, Len(rData) - 7)
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))

            If Val(Arg1) <= 0 Or Val(Arg2) <= 0 Then
                Exit Sub
            End If
            If UserList(UserList(userindex).ComUsu.DestUsu).flags.UserLogged = False Then
                'sigue vivo el usuario ?
                Call FinComerciarUsu(userindex)
                Exit Sub
            Else
                'esta vivo ?
                If UserList(UserList(userindex).ComUsu.DestUsu).flags.Muerto = 1 Then
                    Call FinComerciarUsu(userindex)
                    Exit Sub
                End If
                '//Tiene la cantidad que ofrece ??//'
                    'inventario
                    If Val(Arg2) > UserList(userindex).Invent.Object(Val(Arg1)).Amount Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
                        Exit Sub
                    End If
                End If
                If UserList(userindex).ComUsu.Objeto > 0 Then
                    Call SendData(SendTarget.toindex, userindex, 0, "||No puedes cambiar tu oferta." & FONTTYPE_TALK)
                    Exit Sub
                End If
                'No permitimos vender barcos mientras están equipados (no podés desequiparlos y causa errores)
                If UserList(userindex).flags.Navegando = 1 Then
                    If UserList(userindex).Invent.BarcoSlot = Val(Arg1) Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||No podés vender tu barco mientras lo estés usando." & FONTTYPE_TALK)
                        Exit Sub
                    End If
                End If
                
                UserList(userindex).ComUsu.Objeto = Val(Arg1)
                UserList(userindex).ComUsu.Cant = Val(Arg2)
                If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.DestUsu <> userindex Then
                    Call FinComerciarUsu(userindex)
                    Exit Sub
                Else
                    '[CORREGIDO]
                    If UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.Acepto = True Then
                        'NO NO NO vos te estas pasando de listo...
                        UserList(UserList(userindex).ComUsu.DestUsu).ComUsu.Acepto = False
                        Call SendData(SendTarget.toindex, UserList(userindex).ComUsu.DestUsu, 0, "||" & UserList(userindex).name & " ha cambiado su oferta." & FONTTYPE_TALK)
                    End If
                    '[/CORREGIDO]
                    'Es la ofrenda de respuesta :)
                    Call EnviarObjetoTransaccion(UserList(userindex).ComUsu.DestUsu)
                End If
            Exit Sub
    End Select
    '[/Alejo]

    Select Case UCase$(Left$(rData, 8))
        'clanesnuevo
        Case "ACEPPEAT" 'aceptar paz
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_AceptarPropuestaDePaz(userindex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||Tu clan ha firmado la paz con " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||Tu clan ha firmado la paz con " & UserList(userindex).name & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "RECPALIA" 'rechazar alianza
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_RechazarPropuestaDeAlianza(userindex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||Tu clan rechazado la propuesta de alianza de " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(userindex).name & " ha rechazado nuestra propuesta de alianza con su clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "RECPPEAT" 'rechazar propuesta de paz
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_RechazarPropuestaDePaz(userindex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||Tu clan rechazado la propuesta de paz de " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(userindex).name & " ha rechazado nuestra propuesta de paz con su clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ACEPALIA" 'aceptar alianza
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_AceptarPropuestaDeAlianza(userindex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||Tu clan ha firmado la alianza con " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||Tu clan ha firmado la paz con " & UserList(userindex).name & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "PEACEOFF"
            'un clan solicita propuesta de paz a otro
            rData = Right$(rData, Len(rData) - 8)
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))
            If modGuilds.r_ClanGeneraPropuesta(userindex, Arg1, PAZ, Arg2, Arg3) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Propuesta de paz enviada" & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||" & Arg3 & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ALLIEOFF" 'un clan solicita propuesta de alianza a otro
            rData = Right$(rData, Len(rData) - 8)
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))
            If modGuilds.r_ClanGeneraPropuesta(userindex, Arg1, ALIADOS, Arg2, Arg3) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Propuesta de alianza enviada" & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "||" & Arg3 & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ALLIEDET"
            'un clan pide los detalles de una propuesta de ALIANZA
            rData = Right$(rData, Len(rData) - 8)
            tStr = modGuilds.r_VerPropuesta(userindex, rData, ALIADOS, Arg1)
            If tStr = vbNullString Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & Arg1 & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "ALLIEDE" & tStr)
            End If
            Exit Sub
        Case "PEACEDET" '-"ALLIEDET"
            'un clan pide los detalles de una propuesta de paz
            rData = Right$(rData, Len(rData) - 8)
            tStr = modGuilds.r_VerPropuesta(userindex, rData, PAZ, Arg1)
            If tStr = vbNullString Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & Arg1 & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "PEACEDE" & tStr)
            End If
            Exit Sub
        Case "ENVCOMEN"
            rData = Trim$(Right$(rData, Len(rData) - 8))
            If rData = vbNullString Then Exit Sub
            tStr = modGuilds.a_DetallesAspirante(userindex, rData)
            If tStr = vbNullString Then
                Call SendData(SendTarget.toindex, userindex, 0, "|| El personaje no ha mandado solicitud, o no estás habilitado para verla." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "PETICIO" & tStr)
            End If
            Exit Sub
        Case "ENVALPRO" 'enviame la lista de propuestas de alianza
            tIndex = modGuilds.r_CantidadDePropuestas(userindex, ALIADOS)
            tStr = "ALLIEPR" & tIndex & ","
            If tIndex > 0 Then
                tStr = tStr & modGuilds.r_ListaDePropuestas(userindex, ALIADOS)
            End If
            Call SendData(SendTarget.toindex, userindex, 0, tStr)
            Exit Sub
        Case "ENVPROPP" 'enviame la lista de propuestas de paz
            tIndex = modGuilds.r_CantidadDePropuestas(userindex, PAZ)
            tStr = "PEACEPR" & tIndex & ","
            If tIndex > 0 Then
                tStr = tStr & modGuilds.r_ListaDePropuestas(userindex, PAZ)
            End If
            Call SendData(SendTarget.toindex, userindex, 0, tStr)
            Exit Sub
        Case "DECGUERR" 'declaro la guerra
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_DeclararGuerra(userindex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                'WAR shall be!
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "|| TU CLAN HA ENTRADO EN GUERRA CON " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "|| " & UserList(userindex).name & " LE DECLARA LA GUERRA A TU CLAN" & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "NEWWEBSI"
            rData = Right$(rData, Len(rData) - 8)
            Call modGuilds.ActualizarWebSite(userindex, rData)
            Exit Sub
        Case "ACEPTARI"
            rData = Right$(rData, Len(rData) - 8)
            If Guilds(UserList(userindex).GuildIndex).CantidadDeMiembros >= 15 Then
            Call SendData(SendTarget.toindex, userindex, 0, "||El clan esta lleno." & FONTTYPE_GUILD)
            Exit Sub
            End If
            If Not modGuilds.a_AceptarAspirante(userindex, rData, tStr) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                tInt = NameIndex(rData)
                If tInt > 0 Then
                    Call modGuilds.m_ConectarMiembroAClan(tInt, UserList(userindex).GuildIndex)
                End If
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||" & rData & " ha sido aceptado como miembro del clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "RECHAZAR"
            rData = Trim$(Right$(rData, Len(rData) - 8))
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))
            If Not modGuilds.a_RechazarAspirante(userindex, Arg1, Arg2, Arg3) Then
                Call SendData(SendTarget.toindex, userindex, 0, "|| " & Arg3 & FONTTYPE_GUILD)
            Else
                tInt = NameIndex(Arg1)
                tStr = Arg3 & ": " & Arg2       'el mensaje de rechazo
                If tInt > 0 Then
                    Call SendData(SendTarget.toindex, tInt, 0, "|| " & tStr & FONTTYPE_GUILD)
                Else
                    'hay que grabar en el char su rechazo
                    Call modGuilds.a_RechazarAspiranteChar(Arg1, UserList(userindex).GuildIndex, Arg2)
                End If
            End If
            Exit Sub
        
        Case "ECHARCLA"
            'el lider echa de clan a alguien
            rData = Trim$(Right$(rData, Len(rData) - 8))
            tInt = modGuilds.m_EcharMiembroDeClan(userindex, rData)
            If tInt > 0 Then
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & rData & " fue expulsado del clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "|| No puedes expulsar ese personaje del clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ACTGNEWS"
            rData = Right$(rData, Len(rData) - 8)
            Call modGuilds.ActualizarNoticias(userindex, rData)
            Exit Sub
        Case "1HRINFO<"
            rData = Right$(rData, Len(rData) - 8)
            If Trim$(rData) = vbNullString Then Exit Sub
            tStr = modGuilds.a_DetallesPersonaje(userindex, rData, Arg1)
            If tStr = vbNullString Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & Arg1 & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.toindex, userindex, 0, "CHRINFO" & tStr)
            End If
            Exit Sub
        Case "ABREELEC"
            If Not modGuilds.v_AbrirElecciones(userindex, tStr) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(userindex).GuildIndex, 0, "||¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & UserList(userindex).name & FONTTYPE_GUILD)
            End If
            Exit Sub
    End Select
    

    Select Case UCase$(Left$(rData, 9))
 

        Case "SOLICITUD"
             rData = Right$(rData, Len(rData) - 9)
             Arg1 = ReadField(1, rData, Asc(","))
             Arg2 = ReadField(2, rData, Asc(","))
             If Not modGuilds.a_NuevoAspirante(userindex, Arg1, Arg2, tStr) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||" & tStr & FONTTYPE_GUILD)
             Else
                Call SendData(SendTarget.toindex, userindex, 0, "||Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & Arg1 & "." & FONTTYPE_GUILD)
             End If
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 11))
        Case "CLANDETAILS"
            Dim GII As Integer
            GII = UserList(userindex).GuildIndex
            rData = Right$(rData, Len(rData) - 11)
            If Trim$(rData) = vbNullString Then Exit Sub
            'If m_EsGuildSubLeader(UserList(UserIndex).name, GII) Then
            'Call SendData(SendTarget.ToIndex, UserIndex, 0, "CLANDETSUB" & modGuilds.SendGuildDetails(rData))
            'Else
            Call SendData(SendTarget.toindex, userindex, 0, "CLANDET" & modGuilds.SendGuildDetails(rData))
            'End If
            Exit Sub
    End Select
    
Procesado = False
    
End Sub
