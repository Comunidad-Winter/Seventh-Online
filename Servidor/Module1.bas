Attribute VB_Name = "Module1"
Sub Deathmatch_Ingresa(ByVal userindex As Integer)
On Error GoTo errorh
       
        If (Not Deathmatch_Creado) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||No hay ningun torneo!." & FONTTYPE_INFO)
                Exit Sub
        End If
       
       If (Not Deathmatch_Esperando) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||El torneo ya ha empezado, te quedaste fuera!." & FONTTYPE_INFO)
                Exit Sub
        End If
       
   If Not UserList(userindex).flags.EnDM = True Then
   If Team1_Inscriptos < 4 Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Entraste al deathmatch en el Team 1." & FONTTYPE_INFO)
Call WarpUserChar(userindex, 1, 50, 50, True)
      UserList(userindex).flags.DeathMatch = True
                Team1_Inscriptos = Team1_Inscriptos + 1
Else
                Call SendData(SendTarget.toindex, userindex, 0, "||Entraste al deathmatch en el Team 2." & FONTTYPE_INFO)
Call WarpUserChar(userindex, 1, 50, 50, True)
      UserList(userindex).flags.DeathMatch = True
                Team2_Inscriptos = Team2_Inscriptos + 1
 End If
 End If

    If Team1_Inscriptos = 4 And Team2_Inscriptos = 4 Then
                Call Iniciar_DeathMatch(Team1_Inscriptos, Team2_Inscriptos)
                  Exit Sub
        End If

errorh:
End Sub


   Public Sub Iniciar_DeathMatch(ByVal Team1_Inscriptos As Byte, ByVal Team2_Inscriptos As Byte)
   
   Dim feer1 As Byte
   Dim Feer2 As Byte
    
   For feer1 = 1 To Team1_Inscriptos
   For Feer2 = 1 To Team2_Inscriptos

    UserList(feer1).flags.YaestaJugando = True
    Call WarpUserChar(feer1, 14, 27, 46)

    UserList(Feer2).flags.YaestaJugando = True
    Call WarpUserChar(Feer2, 14, 27, 46)

    Call SendData(toall, 0, 0, "||El DeathMatch esta Completo, Team1 vs Team2." & "~0~200~0~0~0")
    End Sub
Sub IniciarDeath(ByVal userindex As Integer, ByVal puntos As Integer)
On Error GoTo errorh
        If (Deathmatch_Creado) Then
                Call SendData(SendTarget.toindex, userindex, 0, "||Ya hay un deathmatch en curso!." & FONTTYPE_INFO)
                Exit Sub
        End If

        Call SendData(SendTarget.toall, 0, 0, "||" & UserList(userindex).name & " creo un deathmatch de " & Val(puntos) & " puntos, para entrar tipea /INGRESAR (No caen Items) " & FONTTYPE_GUILD)
        Call SendData(SendTarget.toall, 0, 0, "TW48")
       
        Deathmatch_puntos = puntos
        Deathmatch_Creado = True
        Deathmatch_Esperando = True

errorh:
End Sub
