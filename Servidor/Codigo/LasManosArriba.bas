Attribute VB_Name = "LasManosArriba"
   Public Sub ComensarDuelo(ByVal userindex As Integer, ByVal tIndex As Integer)
    UserList(userindex).flags.EstaDueleando = True
    UserList(userindex).flags.Oponente = tIndex
    UserList(tIndex).flags.EstaDueleando = True
    Call WarpUserChar(tIndex, 14, 27, 46)
    UserList(tIndex).flags.Oponente = userindex
    Call WarpUserChar(userindex, 14, 40, 55)
    Call SendData(ToAll, 0, 0, "||Retos: " & UserList(tIndex).name & " y " & UserList(userindex).name & " van a jugar un reto." & "~0~200~0~0~0")
    End Sub
    Public Sub ResetDuelo(ByVal userindex As Integer, ByVal tIndex As Integer)
    UserList(userindex).flags.EsperandoDuelo = False
    UserList(userindex).flags.Oponente = 0
    UserList(userindex).flags.EstaDueleando = False
    Call WarpUserChar(userindex, PosUserReto2.Map, PosUserReto2.X, PosUserReto2.Y) 'Esto tambien, lo pongo alreves porque se me canta la verga - Feer~
    Call WarpUserChar(tIndex, PosUserReto1.Map, PosUserReto1.X, PosUserReto1.Y) 'Esto tambien, lo pongo alreves porque se me canta la verga - Feer~
    UserList(tIndex).flags.EsperandoDuelo = False
    UserList(tIndex).flags.Oponente = 0
    UserList(tIndex).flags.EstaDueleando = False
    End Sub
    Public Sub TerminarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(ToAll, Ganador, 0, "||Retos: " & UserList(Ganador).name & " venció a " & UserList(Perdedor).name & " en un reto." & "~0~200~0~0~1")
    UserList(Ganador).Stats.RetosGanados = UserList(Ganador).Stats.RetosGanados + 1
    UserList(Perdedor).Stats.RetosPerdidos = UserList(Perdedor).Stats.RetosPerdidos + 1
    Call ResetDuelo(Ganador, Perdedor)
    End Sub
    Public Sub DesconectarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(ToAll, Ganador, 0, "||Retos: El reto ha sido cancelado por la desconexión de " & UserList(Perdedor).name & "." & "~0~200~0~0~1")
    Call ResetDuelo(Ganador, Perdedor)
    End Sub

