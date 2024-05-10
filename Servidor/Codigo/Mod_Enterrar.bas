Attribute VB_Name = "Mod_Enterrar"
Option Explicit
Public FeerMap As Integer
Public FeerX As Integer
Public FeerY As Integer
Public ObjPremio As Obj
Public Sub Enterrar(ByVal userindex As Integer)
    
If ObjData(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).OBJInfo.ObjIndex).Agarrable <> 1 Then
        Obj = ObjData(MapData(UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y).OBJInfo.ObjIndex)
        MiObj.Amount = MapData(UserList(userindex).pos.Map, X, Y).OBJInfo.Amount
        MiObj.ObjIndex = MapData(UserList(userindex).pos.Map, X, Y).OBJInfo.ObjIndex
    
    'Quitamos el objeto
    Call EraseObj(SendTarget.ToMap, 0, UserList(userindex).pos.Map, MapData(UserList(userindex).pos.Map, X, Y).OBJInfo.Amount, UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y)
    
    FeerMap = UserList(userindex).pos.Map
    FeerX = UserList(userindex).pos.X
    FeerY = UserList(userindex).pos.Y

    ObjEnterrado = False
End If
End Sub
