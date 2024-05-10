Attribute VB_Name = "TorneosAutomaticosPlantes"
Option Explicit
' Codigo: Torneos Automaticos 100%
' Autor: Joan Calderón - SaturoS.
Public Torneo_Activop As Boolean
Public Torneo_Esperandop As Boolean
Private Torneo_Rondasp As Integer
Private Torneo_Luchadoresp() As Integer
 
Private Const mapatorneop As Integer = 81
' esquinas superior isquierda del ring
Private Const esquina1xp As Integer = 49
Private Const esquina1yp As Integer = 73
' esquina inferior derecha del ring
Private Const esquina2xp As Integer = 50
Private Const esquina2yp As Integer = 73
' Donde esperan los tios
Private Const esperaxp As Integer = 57
Private Const esperayp As Integer = 77
' Mapa desconecta
Private Const mapa_fuerap As Integer = 1
Private Const fueraesperayp As Integer = 50
Private Const fueraesperaxp As Integer = 50
 ' estas son las pocisiones de las 2 esquinas de la zona de espera, en su mapa tienen que tener en la misma posicion las 2 esquinas.
Private Const X1P As Integer = 57
Private Const X2P As Integer = 58
Private Const Y1P As Integer = 77
Private Const Y2P As Integer = 77
 
Sub Torneoautop_Cancela()
On Error GoTo errorh:
    If (Not Torneo_Activo And Not Torneo_Esperando) Then Exit Sub
    Torneo_Activop = False
    Torneo_Esperandop = False
    Call SendData(SendTarget.toall, 0, 0, "||Torneo de Plantes: Torneo cancelado por falta de participantes." & FONTTYPE_VENENO)
    Dim i As Integer
     For i = LBound(Torneo_Luchadoresp) To UBound(Torneo_Luchadoresp)
                If (Torneo_Luchadoresp(i) <> -1) Then
                 Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = mapa_fuerap
                    FuturePos.X = fueraesperaxp: FuturePos.Y = fueraesperayp
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadoresp(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                      UserList(Torneo_Luchadoresp(i)).flags.Automaticop = False
                End If
        Next i
errorh:
End Sub
Sub Rondas_Cancelap()
On Error GoTo errorh
    If (Not Torneo_Activop And Not Torneo_Esperandop) Then Exit Sub
    Torneo_Activop = False
    Torneo_Esperandop = False
    Call SendData(SendTarget.toall, 0, 0, "||Torneo de Plantes: Torneo cancelado por Game Master" & FONTTYPE_VENENO)
    Dim i As Integer
    For i = LBound(Torneo_Luchadoresp) To UBound(Torneo_Luchadoresp)
                If (Torneo_Luchadoresp(i) <> -1) Then
                        Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = mapa_fuerap
                    FuturePos.X = fueraesperaxp: FuturePos.Y = fueraesperayp
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadoresp(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                    UserList(Torneo_Luchadoresp(i)).flags.Automaticop = False
                End If
        Next i
errorh:
End Sub
Sub Rondas_UsuarioMuerep(ByVal userindex As Integer, Optional Real As Boolean = True, Optional CambioMapa As Boolean = False)
On Error GoTo rondas_usuariomuerep_errorh
        Dim i As Integer, pos As Integer, j As Integer
        Dim combate As Integer, LI1 As Integer, LI2 As Integer
        Dim UI1 As Integer, UI2 As Integer
If (Not Torneo_Activop) Then
                Exit Sub
            ElseIf (Torneo_Activop And Torneo_Esperandop) Then
                For i = LBound(Torneo_Luchadoresp) To UBound(Torneo_Luchadoresp)
                    If (Torneo_Luchadoresp(i) = userindex) Then
                        Torneo_Luchadoresp(i) = -1
                        Call WarpUserChar(userindex, mapa_fuerap, fueraesperayp, fueraesperaxp, True)
                         UserList(userindex).flags.Automaticop = False
                        Exit Sub
                    End If
                Next i
                Exit Sub
            End If
 
        For pos = LBound(Torneo_Luchadoresp) To UBound(Torneo_Luchadoresp)
                If (Torneo_Luchadoresp(pos) = userindex) Then Exit For
        Next pos
 
        ' si no lo ha encontrado
        If (Torneo_Luchadoresp(pos) <> userindex) Then Exit Sub
       
 '  Ojo con esta parte, aqui es donde verifica si el usuario esta en la posicion de espera del torneo, en estas cordenadas tienen que fijarse al crear su Mapa de torneos.
 
If UserList(userindex).pos.X >= X1P And UserList(userindex).pos.X <= X2P And UserList(userindex).pos.Y >= Y1P And UserList(userindex).pos.Y <= Y2P Then
Call SendData(SendTarget.toall, 0, 0, "||Torneo de Plantes: " & UserList(userindex).name & " se fue del torneo mientras esperaba pelear.!" & FONTTYPE_VENENO)
Call WarpUserChar(userindex, mapa_fuerap, fueraesperaxp, fueraesperayp, True)
UserList(userindex).flags.Automaticop = False
Torneo_Luchadoresp(pos) = -1
Exit Sub
End If
 
        combate = 1 + (pos - 1) \ 2
 
        'ponemos li1 y li2 (luchador index) de los que combatian
        LI1 = 2 * (combate - 1) + 1
        LI2 = LI1 + 1
 
        'se informa a la gente
        If (Real) Then
                Call SendData(SendTarget.toall, 0, 0, "||Torneo de Plantes: " & UserList(userindex).name & " pierde el plante!" & FONTTYPE_VENENO)
        Else
                Call SendData(SendTarget.toall, 0, 0, "||Torneo de Plantes: " & UserList(userindex).name & " se fue del plante!" & FONTTYPE_VENENO)
        End If
 
        'se le teleporta fuera si murio
        If (Real) Then
                Call WarpUserChar(userindex, mapa_fuerap, fueraesperaxp, fueraesperayp, True)
                 UserList(userindex).flags.Automaticop = False
        ElseIf (Not CambioMapa) Then
             
                 Call WarpUserChar(userindex, mapa_fuerap, fueraesperaxp, fueraesperayp, True)
                  UserList(userindex).flags.Automaticop = False
        End If
 
        'se le borra de la lista y se mueve el segundo a li1
        If (Torneo_Luchadoresp(LI1) = userindex) Then
                Torneo_Luchadoresp(LI1) = Torneo_Luchadoresp(LI2) 'cambiamos slot
                Torneo_Luchadoresp(LI2) = -1
        Else
                Torneo_Luchadoresp(LI2) = -1
        End If
 
    'si es la ultima ronda
    If (Torneo_Rondasp = 1) Then
        Call WarpUserChar(Torneo_Luchadoresp(LI1), mapa_fuerap, 51, 51, True)
        Call SendData(SendTarget.toall, 0, 0, "||Ganador del Torneo de Plantes: " & UserList(Torneo_Luchadoresp(LI1)).name & FONTTYPE_ROJO)
        Call SendData(SendTarget.toall, 0, 0, "||Premio: 25 Puntos de Torneo." & FONTTYPE_VENENO)
        UserList(Torneo_Luchadoresp(LI1)).Stats.PuntosTorneo = UserList(Torneo_Luchadoresp(LI1)).Stats.PuntosTorneo + 25
         UserList(Torneo_Luchadoresp(LI1)).flags.Automaticop = False
       Call SendUserStatsBox(Torneo_Luchadoresp(LI1))
        Torneo_Activop = False
        Exit Sub
    Else
        'a su compañero se le teleporta dentro, condicional por seguridad
        Call WarpUserChar(Torneo_Luchadoresp(LI1), mapatorneop, esperaxp, esperayp, True)
    End If
 
               
        'si es el ultimo combate de la ronda
        If (2 ^ Torneo_Rondasp = 2 * combate) Then
 
                Call SendData(SendTarget.toall, 0, 0, "||Torneo de Plantes: Siguiente Plante!" & FONTTYPE_VENENO)
                Torneo_Rondasp = Torneo_Rondasp - 1
 
        'antes de llamar a la proxima ronda hay q copiar a los tipos
        For i = 1 To 2 ^ Torneo_Rondasp
                UI1 = Torneo_Luchadoresp(2 * (i - 1) + 1)
                UI2 = Torneo_Luchadoresp(2 * i)
                If (UI1 = -1) Then UI1 = UI2
                Torneo_Luchadoresp(i) = UI1
        Next i
ReDim Preserve Torneo_Luchadoresp(1 To 2 ^ Torneo_Rondasp) As Integer
        Call Rondas_Combatep(1)
        Exit Sub
        End If
 
        'vamos al siguiente combate
        Call Rondas_Combatep(combate + 1)
rondas_usuariomuerep_errorh:
 
End Sub
 
 
 
Sub Rondas_UsuarioDesconectap(ByVal userindex As Integer)
On Error GoTo errorh
Call SendData(SendTarget.toall, 0, 0, "||Torneo: " & UserList(userindex).name & " se ha desconectado en Torneo de Plantes, se le penaliza con 2 puntos de torneo!" & FONTTYPE_TALK)
 If UserList(userindex).Stats.PuntosTorneo >= 2 Then
UserList(userindex).Stats.PuntosTorneo = UserList(userindex).Stats.PuntosTorneo - 2
End If
Call SendUserStatsBox(userindex)
        Call Rondas_UsuarioMuerep(userindex, False, False)
errorh:
End Sub
 
 
 
Sub Rondas_UsuarioCambiamapap(ByVal userindex As Integer)
On Error GoTo errorh
        Call Rondas_UsuarioMuerep(userindex, False, True)
errorh:
End Sub
 
Sub torneos_autop(ByVal rondasp As Integer)
On Error GoTo errorh
If (Torneo_Activop) Then
               
                Exit Sub
        End If
        Call SendData(SendTarget.toall, 0, 0, "||Torneo de Plantes: Esta empezando un nuevo torneo de plantes para " & Val(2 ^ rondasp) & " participantes!! para participar pon /PLANTES - (No cae inventario)" & FONTTYPE_ROJO)
        Call SendData(SendTarget.toall, 0, 0, "TW48")
       
        Torneo_Rondasp = rondasp
        Torneo_Activop = True
        Torneo_Esperandop = True
 
        ReDim Torneo_Luchadoresp(1 To 2 ^ rondasp) As Integer
        Dim i As Integer
        For i = LBound(Torneo_Luchadoresp) To UBound(Torneo_Luchadoresp)
                Torneo_Luchadoresp(i) = -1
        Next i
errorh:
End Sub
 
Sub Torneos_Iniciap(ByVal userindex As Integer, ByVal rondasp As Integer)
On Error GoTo errorh
        If (Torneo_Activop) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Ya hay un torneo!." & FONTTYPE_INFO)
                Exit Sub
        End If
        Call SendData(SendTarget.toall, 0, 0, "||Torneo de Plantes: Esta empezando un nuevo torneo de plantes para " & Val(2 ^ rondasp) & " participantes!! para participar pon /PLANTES - (No cae inventario)" & FONTTYPE_VENENO)
        Call SendData(SendTarget.toall, 0, 0, "TW48")
       
        Torneo_Rondasp = rondasp
        Torneo_Activop = True
        Torneo_Esperandop = True
 
        ReDim Torneo_Luchadoresp(1 To 2 ^ rondasp) As Integer
        Dim i As Integer
        For i = LBound(Torneo_Luchadoresp) To UBound(Torneo_Luchadoresp)
                Torneo_Luchadoresp(i) = -1
        Next i
errorh:
End Sub
 
 
 
Sub Torneos_Entrap(ByVal userindex As Integer)
On Error GoTo errorh
        Dim i As Integer
       
        If (Not Torneo_Activop) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No hay ningun torneo!." & FONTTYPE_INFO)
                Exit Sub
        End If
       
        If (Not Torneo_Esperandop) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||El torneo ya ha empezado, te quedaste fuera!." & FONTTYPE_INFO)
                Exit Sub
        End If
       
        For i = LBound(Torneo_Luchadoresp) To UBound(Torneo_Luchadoresp)
                If (Torneo_Luchadoresp(i) = userindex) Then
                        Call SendData(SendTarget.toindex, userindex, 0, "||Ya estas dentro!" & FONTTYPE_WARNING)
                        Exit Sub
                End If
        Next i
 
        For i = LBound(Torneo_Luchadoresp) To UBound(Torneo_Luchadoresp)
        If (Torneo_Luchadoresp(i) = -1) Then
                Torneo_Luchadoresp(i) = userindex
                 Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.Map = mapatorneop
                    FuturePos.X = esperaxp: FuturePos.Y = esperayp
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                   
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadoresp(i), NuevaPos.Map, NuevaPos.X, NuevaPos.Y, True)
                 UserList(Torneo_Luchadoresp(i)).flags.Automaticop = True
                 
                Call SendData(SendTarget.toindex, userindex, 0, "||Estas dentro del torneo!" & FONTTYPE_INFO)
               
                Call SendData(SendTarget.toall, 0, 0, "||Torneo de Plantes: El usuario " & UserList(userindex).name & " entro al torneo." & FONTTYPE_VENENO)
                If (i = UBound(Torneo_Luchadoresp)) Then
                Call SendData(SendTarget.toall, 0, 0, "||Torneo de Plantes: Empieza el torneo!" & FONTTYPE_ROJO)
                Torneo_Esperandop = False
                Call Rondas_Combatep(1)
     
                End If
                  Exit Sub
        End If
        Next i
errorh:
End Sub
 
 
Sub Rondas_Combatep(combate As Integer)
On Error GoTo errorh
Dim UI1 As Integer, UI2 As Integer
    UI1 = Torneo_Luchadoresp(2 * (combate - 1) + 1)
    UI2 = Torneo_Luchadoresp(2 * combate)
   
    If (UI2 = -1) Then
        UI2 = Torneo_Luchadoresp(2 * (combate - 1) + 1)
        UI1 = Torneo_Luchadoresp(2 * combate)
    End If
   
    If (UI1 = -1) Then
        Call SendData(SendTarget.toall, 0, 0, "||Torneo de Plantes: Plante anulado porque un participante involucrado se desconecto" & FONTTYPE_TALK)
        If (Torneo_Rondasp = 1) Then
            If (UI2 <> -1) Then
                Call SendData(SendTarget.toall, 0, 0, "||Torneo de Plantes: Torneo terminado. Ganador del torneo por eliminacion: " & UserList(UI2).name & FONTTYPE_VENENO)
                UserList(UI2).flags.Automaticop = False
                ' dale_recompensa()
                Torneo_Activop = False
                Exit Sub
            End If
            Call SendData(SendTarget.toall, 0, 0, "||Torneo: Torneo terminado. No hay ganador porque todos se fueron :(" & FONTTYPE_VENENO)
            Exit Sub
        End If
        If (UI2 <> -1) Then _
            Call SendData(SendTarget.toall, 0, 0, "||Torneo de Plantes: " & UserList(UI2).name & " pasa a la siguiente ronda!" & FONTTYPE_TALK)
   
        If (2 ^ Torneo_Rondasp = 2 * combate) Then
            Call SendData(SendTarget.toall, 0, 0, "||Torneo: Siguiente Plante!" & FONTTYPE_ROJO)
            Torneo_Rondasp = Torneo_Rondasp - 1
            'antes de llamar a la proxima ronda hay q copiar a los tipos
            Dim i As Integer, j As Integer
            For i = 1 To 2 ^ Torneo_Rondasp
                UI1 = Torneo_Luchadoresp(2 * (i - 1) + 1)
                UI2 = Torneo_Luchadoresp(2 * i)
                If (UI1 = -1) Then UI1 = UI2
                Torneo_Luchadoresp(i) = UI1
            Next i
            ReDim Preserve Torneo_Luchadoresp(1 To 2 ^ Torneo_Rondasp) As Integer
            Call Rondas_Combatep(1)
            Exit Sub
        End If
        Call Rondas_Combatep(combate + 1)
        Exit Sub
    End If
 
    Call SendData(SendTarget.toall, 0, 0, "||Torneo de Plantes: " & UserList(UI1).name & " versus " & UserList(UI2).name & ". Peleen!" & FONTTYPE_GRISN)
 
    Call WarpUserChar(UI1, mapatorneop, esquina1xp, esquina1yp, True)
    Call WarpUserChar(UI2, mapatorneop, esquina2xp, esquina2yp, True)
errorh:
End Sub

