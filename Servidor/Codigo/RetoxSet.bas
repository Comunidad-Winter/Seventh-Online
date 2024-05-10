Attribute VB_Name = "RetoxSet"
   Public Sub ComensarDueloxset(ByVal userindex As Integer, ByVal tIndex As Integer)
    UserList(userindex).flags.EstaDueleandoxset = True
    UserList(userindex).flags.Oponentexset = tIndex
    UserList(tIndex).flags.EstaDueleandoxset = True
    Call WarpUserChar(tIndex, 70, 41, 56)
    UserList(tIndex).flags.Oponentexset = userindex
    Call WarpUserChar(userindex, 70, 62, 42)
    Call SendData(ToAll, 0, 0, "||" & UserList(tIndex).name & " y " & UserList(userindex).name & " van a jugar un duelo por items!." & "~255~255~0~0~1")
    End Sub
    Public Sub TerminarDueloxset(ByVal Ganadorxset As Integer, ByVal Perdedorxset As Integer)
    Call SendData(ToAll, Ganadorxset, 0, "|| " & UserList(Perdedorxset).name & " venció a " & UserList(Ganadorxset).name & " en un duelo por items!." & "~255~255~0~1")
    UserList(Ganadorxset).flags.EsperandoDueloxset = False
    UserList(Ganadorxset).flags.Oponentexset = 0
    UserList(Ganadorxset).flags.EstaDueleandoxset = False
    UserList(Perdedorxset).flags.EsperandoDueloxset = False
    UserList(Perdedorxset).flags.Oponentexset = 0
    UserList(Perdedorxset).flags.EstaDueleandoxset = False
    End Sub
    Public Sub DesconectarDueloxset(ByVal Ganadorxset As Integer, ByVal Perdedorxset As Integer)
    Call SendData(ToAll, Ganadorxset, 0, "||El duelo por items ha sido cancelado por la desconexión de " & UserList(Perdedorxset).name & "." & "~255~255~0~0~1")
    End Sub
