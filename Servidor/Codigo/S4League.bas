Attribute VB_Name = "S4League"
'Todo esto fue programado por Feer~ :$

Sub TouchDown(userindex As Integer)
 
If HayTD = True Then
        If UserList(userindex).pos.Map = 120 Then
            If TieneObjetos(1073, 1, userindex) Then
                    ' TEAM Alpha!
                    If UserList(userindex).pos.X = 40 And UserList(userindex).pos.Y = 17 And UserList(userindex).flags.TeamTD = 2 Then
                        TD_A = TD_A + 1
                        Call SendData(SendTarget.toall, 0, 0, "||TouchDown <Team Alpha> By " & UserList(userindex).name & FONTTYPE_INFO)
                        Call SendData(toall, 0, 0, "TD" & TD_A)
                        Call QuitarObjetos(1073, 1, userindex)
                        UserList(userindex).Char.Aura = 0
                        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
                        
                        If TD_A = 5 Or TD_B = 5 Then
                            Call EntreTiempo
                            Call QuitarObjetos(1073, 1, userindex)
                        Exit Sub
                        End If
                        
                        If TD_A = 10 Or TD_B = 10 Then
                            Call TerminarPartido
                         Else
                            Call ReComenzarTouchDown
                        End If
                        
                    ElseIf UserList(userindex).pos.X = 40 And UserList(userindex).pos.Y = 18 And UserList(userindex).flags.TeamTD = 2 Then
                        TD_A = TD_A + 1
                        Call SendData(SendTarget.toall, 0, 0, "||TouchDown <Team Alpha> By " & UserList(userindex).name & FONTTYPE_INFO)
                        Call SendData(toall, 0, 0, "TD" & TD_A)
                        Call QuitarObjetos(1073, 1, userindex)
                        UserList(userindex).Char.Aura = 0
                        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
                        If TD_A = 5 Or TD_B = 5 Then
                            Call EntreTiempo
                            Call QuitarObjetos(1073, 1, userindex)
                        Exit Sub
                        End If
                        
                        If TD_A = 10 Or TD_B = 10 Then
                            Call TerminarPartido
                         Else
                            Call ReComenzarTouchDown
                        End If
                        
                    ElseIf UserList(userindex).pos.X = 41 And UserList(userindex).pos.Y = 17 And UserList(userindex).flags.TeamTD = 2 Then
                        TD_A = TD_A + 1
                        Call SendData(SendTarget.toall, 0, 0, "||TouchDown <Team Alpha> By " & UserList(userindex).name & FONTTYPE_INFO)
                        Call SendData(toall, 0, 0, "TD" & TD_A)
                        Call QuitarObjetos(1073, 1, userindex)
                        UserList(userindex).Char.Aura = 0
                        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
                        If TD_A = 5 Or TD_B = 5 Then
                            Call EntreTiempo
                            Call QuitarObjetos(1073, 1, userindex)
                        Exit Sub
                        End If
                        
                        If TD_A = 10 Or TD_B = 10 Then
                            Call TerminarPartido
                         Else
                            Call ReComenzarTouchDown
                        End If
                        
                    ElseIf UserList(userindex).pos.X = 41 And UserList(userindex).pos.Y = 18 And UserList(userindex).flags.TeamTD = 2 Then
                        TD_A = TD_A + 1
                        Call SendData(SendTarget.toall, 0, 0, "||TouchDown <Team Alpha> By " & UserList(userindex).name & FONTTYPE_INFO)
                        Call SendData(toall, 0, 0, "TD" & TD_A)
                        Call QuitarObjetos(1073, 1, userindex)
                        UserList(userindex).Char.Aura = 0
                        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
                        If TD_A = 5 Or TD_B = 5 Then
                            Call EntreTiempo
                            Call QuitarObjetos(1073, 1, userindex)
                        Exit Sub
                        End If
                        
                        If TD_A = 10 Or TD_B = 10 Then
                            Call TerminarPartido
                         Else
                            Call ReComenzarTouchDown
                        End If
                        
                    ' TEAM Beta!
                    ElseIf UserList(userindex).pos.X = 39 And UserList(userindex).pos.Y = 83 And UserList(userindex).flags.TeamTD = 1 Then
                        TD_B = TD_B + 1
                        Call SendData(SendTarget.toall, 0, 0, "||TouchDown <Team Beta> By " & UserList(userindex).name & FONTTYPE_INFO)
                        Call SendData(toall, 0, 0, "TF" & TD_B)
                        Call QuitarObjetos(1073, 1, userindex)
                        UserList(userindex).Char.Aura = 0
                        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
                        If TD_A = 5 Or TD_B = 5 Then
                            Call EntreTiempo
                            Call QuitarObjetos(1073, 1, userindex)
                        Exit Sub
                        End If
                        
                        If TD_A = 10 Or TD_B = 10 Then
                            Call TerminarPartido
                         Else
                            Call ReComenzarTouchDown
                        End If
                    ElseIf UserList(userindex).pos.X = 39 And UserList(userindex).pos.Y = 84 And UserList(userindex).flags.TeamTD = 1 Then
                        TD_B = TD_B + 1
                        Call SendData(SendTarget.toall, 0, 0, "||TouchDown <Team Beta> By " & UserList(userindex).name & FONTTYPE_INFO)
                        Call SendData(toall, 0, 0, "TF" & TD_B)
                        Call QuitarObjetos(1073, 1, userindex)
                        UserList(userindex).Char.Aura = 0
                        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
                        If TD_A = 5 Or TD_B = 5 Then
                            Call EntreTiempo
                            Call QuitarObjetos(1073, 1, userindex)
                        Exit Sub
                        End If
                        
                        If TD_A = 10 Or TD_B = 10 Then
                            Call TerminarPartido
                         Else
                            Call ReComenzarTouchDown
                        End If
                    ElseIf UserList(userindex).pos.X = 40 And UserList(userindex).pos.Y = 83 And UserList(userindex).flags.TeamTD = 1 Then
                        TD_B = TD_B + 1
                        Call SendData(SendTarget.toall, 0, 0, "||TouchDown <Team Beta> By " & UserList(userindex).name & FONTTYPE_INFO)
                        Call SendData(toall, 0, 0, "TF" & TD_B)
                        Call QuitarObjetos(1073, 1, userindex)
                        UserList(userindex).Char.Aura = 0
                        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
                        If TD_A = 5 Or TD_B = 5 Then
                            Call EntreTiempo
                            Call QuitarObjetos(1073, 1, userindex)
                        Exit Sub
                        End If
                        
                        If TD_A = 10 Or TD_B = 10 Then
                            Call TerminarPartido
                         Else
                            Call ReComenzarTouchDown
                        End If
                    ElseIf UserList(userindex).pos.X = 40 And UserList(userindex).pos.Y = 84 And UserList(userindex).flags.TeamTD = 1 Then
                        TD_B = TD_B + 1
                        Call SendData(SendTarget.toall, 0, 0, "||TouchDown <Team Beta> By " & UserList(userindex).name & FONTTYPE_INFO)
                        Call SendData(toall, 0, 0, "TF" & TD_B)
                        Call QuitarObjetos(1073, 1, userindex)
                        UserList(userindex).Char.Aura = 0
                        Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
                        If TD_A = 5 Or TD_B = 5 Then
                            Call EntreTiempo
                            Call QuitarObjetos(1073, 1, userindex)
                        Exit Sub
                        End If
                        
                        If TD_A = 10 Or TD_B = 10 Then
                            Call TerminarPartido
                         Else
                            Call ReComenzarTouchDown
                        End If
                    End If
                End If
            End If
        End If
 
End Sub
Public Sub ComenzarTouchDown()
    Dim i As Integer, Pelota As Obj, Bomba As Obj
    Pelota.ObjIndex = 1073
    Pelota.Amount = 1
    
    Bomba.ObjIndex = 1074
    Bomba.Amount = 1
   
  For i = 1 To LastUser
         If UserList(i).flags.TeamTD = 1 Then
               Call WarpUserChar(i, 120, 57, 29)
         ElseIf UserList(i).flags.TeamTD = 2 Then
               Call WarpUserChar(i, 120, 24, 72)
         End If
  Next i
 
  Call MakeObj(toall, 0, 0, Pelota, 120, 40, 50)
    'Bombas Alpha
  Call MakeObj(toall, 0, 0, Bomba, 120, 40, 29)
  Call MakeObj(toall, 0, 0, Bomba, 120, 39, 27)
  Call MakeObj(toall, 0, 0, Bomba, 120, 42, 27)
  Call MakeObj(toall, 0, 0, Bomba, 120, 41, 24)
  Call MakeObj(toall, 0, 0, Bomba, 120, 39, 23)
  Call MakeObj(toall, 0, 0, Bomba, 120, 39, 21)
  Call MakeObj(toall, 0, 0, Bomba, 120, 42, 20)
  Call MakeObj(toall, 0, 0, Bomba, 120, 40, 19)
  
  'Bombas Beta
  Call MakeObj(toall, 0, 0, Bomba, 120, 39, 71)
  Call MakeObj(toall, 0, 0, Bomba, 120, 38, 73)
  Call MakeObj(toall, 0, 0, Bomba, 120, 41, 73)
  Call MakeObj(toall, 0, 0, Bomba, 120, 39, 76)
  Call MakeObj(toall, 0, 0, Bomba, 120, 41, 77)
  Call MakeObj(toall, 0, 0, Bomba, 120, 41, 79)
  Call MakeObj(toall, 0, 0, Bomba, 120, 38, 80)
  Call MakeObj(toall, 0, 0, Bomba, 120, 40, 82)
  HayTD = True
End Sub
Public Sub ReComenzarTouchDown()
    Dim i As Integer, Pelota As Obj, Bomba As Obj
    Pelota.ObjIndex = 1073
    Pelota.Amount = 1
    
    Bomba.ObjIndex = 1074
    Bomba.Amount = 1
    
  For i = 1 To LastUser
         If UserList(i).flags.TeamTD = 1 Then
               Call WarpUserChar(i, 120, 57, 29)
         ElseIf UserList(i).flags.TeamTD = 2 Then
            Call WarpUserChar(i, 120, 24, 72)
         End If
  Next i
 
  Call MakeObj(toall, 0, 0, Pelota, 120, 40, 50)
  
  'Bombas Alpha
  Call MakeObj(toall, 0, 0, Bomba, 120, 40, 29)
  Call MakeObj(toall, 0, 0, Bomba, 120, 39, 27)
  Call MakeObj(toall, 0, 0, Bomba, 120, 42, 27)
  Call MakeObj(toall, 0, 0, Bomba, 120, 41, 24)
  Call MakeObj(toall, 0, 0, Bomba, 120, 39, 23)
  Call MakeObj(toall, 0, 0, Bomba, 120, 39, 21)
  Call MakeObj(toall, 0, 0, Bomba, 120, 42, 20)
  Call MakeObj(toall, 0, 0, Bomba, 120, 40, 19)
  
  'Bombas Beta
  Call MakeObj(toall, 0, 0, Bomba, 120, 39, 71)
  Call MakeObj(toall, 0, 0, Bomba, 120, 38, 73)
  Call MakeObj(toall, 0, 0, Bomba, 120, 41, 73)
  Call MakeObj(toall, 0, 0, Bomba, 120, 39, 76)
  Call MakeObj(toall, 0, 0, Bomba, 120, 41, 77)
  Call MakeObj(toall, 0, 0, Bomba, 120, 41, 79)
  Call MakeObj(toall, 0, 0, Bomba, 120, 38, 80)
  Call MakeObj(toall, 0, 0, Bomba, 120, 40, 82)
End Sub
Public Sub TerminarPartido()
    Dim i As Integer
    For i = 1 To LastUser
         If UserList(i).flags.TeamTD <> 0 Then
               UserList(i).flags.TeamTD = 0
               Call WarpUserChar(i, 1, 48, 51)
         End If
    Next i
    
   Dim k As Byte
   For k = 1 To LastUser
   UserList(k).flags.EnTD = 0
   Next k
   
    Call SendData(SendTarget.toall, 0, 0, "||Termino el touchdown el resultado es: " & FONTTYPE_INFO)
    Call SendData(SendTarget.toall, 0, 0, "||Alpha: " & TD_A & FONTTYPE_INFO)
    Call SendData(SendTarget.toall, 0, 0, "||Beta: " & TD_B & FONTTYPE_INFO)
    Call SendData(SendTarget.toall, 0, 0, "||¡Felicidades al equipo ganador!" & FONTTYPE_INFO)
    Call SendData(SendTarget.toall, 0, 0, "AGF")
                       
    TD_A = 0
    TD_B = 0
    HayTD = False
End Sub
Sub MuereUserTD(userindex As Integer)

Dim Pelota As Obj
Pelota.ObjIndex = 1073
Pelota.Amount = 1

If TieneObjetos(1073, 1, userindex) Then
Call QuitarObjetos(1073, 1, userindex)
Call TirarItemAlPiso(UserList(userindex).pos, Pelota)
End If

If UserList(userindex).flags.Muerto = 1 And UserList(userindex).flags.TeamTD = 1 Then
UserList(userindex).Char.Aura = 0
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
UserList(userindex).Counters.MuereEnTD = 7
End If

If UserList(userindex).flags.Muerto = 1 And UserList(userindex).flags.TeamTD = 2 Then
UserList(userindex).Char.Aura = 0
Call SendData(SendTarget.ToPCArea, userindex, UserList(userindex).pos.Map, "TLX" & UserList(userindex).Char.CharIndex & "," & UserList(userindex).Char.Aura)
UserList(userindex).Counters.MuereEnTD = 7
End If
End Sub

Sub EntreTiempo()
Dim Feerpro As Long
  For Feerpro = 1 To LastUser
         If UserList(Feerpro).flags.TeamTD <> 0 Then
            Call WarpUserChar(Feerpro, 81, 22, 57)
            UserList(Feerpro).Counters.EntreTiempo = 20
            Call SendData(SendTarget.toindex, Feerpro, 0, "||Comenzo el entretiempo, el juego se renaudara en 20 segundos." & FONTTYPE_INFO)
         End If
  Next Feerpro
End Sub
