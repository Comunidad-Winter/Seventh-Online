Attribute VB_Name = "TCP_HandleData3"
Public Sub HandleData_3(ByVal userindex As Integer, rData As String, ByRef Procesado As Boolean)


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

If UserList(userindex).flags.Privilegios = PlayerType.User Then
    UserList(userindex).Counters.IdleCount = IdleCountBackup
End If

'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<
If UserList(userindex).flags.Privilegios <= VIP Then Exit Sub
'>>>>>>>>>>>>>>>>>>>>>> SOLO ADMINISTRADORES <<<<<<<<<<<<<<<<<<<

If UCase$(Left$(rData, 9)) = "/CUATRO4 " Then
If UserList(userindex).flags.EsRolesMaster Then Exit Sub
rData = Right$(rData, Len(rData) - 9)
Dim wIndex As Integer
Dim w2Index As Integer
Dim w3Index As Integer
Dim w4Index As Integer
Dim w5Index As Integer
Dim w6Index As Integer
Dim w7index As Integer
Dim w8index As Integer
wIndex = NameIndex(ReadField(1, rData, 64))
w2Index = NameIndex(ReadField(2, rData, 64))
w3Index = NameIndex(ReadField(3, rData, 64))
w4Index = NameIndex(ReadField(4, rData, 64))
w5Index = NameIndex(ReadField(5, rData, 64))
w6Index = NameIndex(ReadField(6, rData, 64))
w7index = NameIndex(ReadField(7, rData, 64))
w8index = NameIndex(ReadField(8, rData, 64))
If Arena4 = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 4." & FONTTYPE_INFO)
    Exit Sub
    End If
If wIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf w2Index <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf w3Index <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El tercer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf w4Index <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El cuarto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf w5Index <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El quinto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf w6Index <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El sexto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf w7index <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El septimo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf w8index <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El octavo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
End If
If wIndex = w2Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y segundo)" & FONTTYPE_INFO)
Exit Sub
ElseIf wIndex = w3Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y tercero)" & FONTTYPE_INFO)
Exit Sub
ElseIf wIndex = w4Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf wIndex = w5Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf wIndex = w6Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf w2Index = w7index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y tercer)" & FONTTYPE_INFO)
Exit Sub
ElseIf w2Index = w4Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf w2Index = w5Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf w2Index = w6Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf w2Index = w7index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y septimo)" & FONTTYPE_INFO)
Exit Sub
ElseIf w2Index = w8index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y octavo)" & FONTTYPE_INFO)
Exit Sub
ElseIf w3Index = w4Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf w3Index = w5Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf w3Index = w6Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf w4Index = w5Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf w4Index = w6Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf w5Index = w6Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (quinto y sexto)" & FONTTYPE_INFO)
Exit Sub
End If
Call SendData(ToAll, userindex, 0, "||Torneo: Arena 4: " & UserList(wIndex).name & ", " & UserList(w2Index).name & ", " & UserList(w3Index).name & " y " & UserList(w4Index).name & " VS " & UserList(w5Index).name & ", " & UserList(w6Index).name & ", " & UserList(w7index).name & " y " & UserList(w8index).name & "~230~230~0~1~0")
Call WarpUserChar(wIndex, 81, 69, 44, True)
Call WarpUserChar(w2Index, 81, 69, 45, True)
Call WarpUserChar(w3Index, 81, 70, 44, True)
Call WarpUserChar(w4Index, 81, 70, 45, True)
Call WarpUserChar(w5Index, 81, 87, 62, True)
Call WarpUserChar(w6Index, 81, 87, 63, True)
Call WarpUserChar(w7index, 81, 88, 62, True)
Call WarpUserChar(w8index, 81, 88, 63, True)
CuentaArena = 4
Arena4 = True
Torne.Jugador25 = wIndex
Torne.Jugador26 = w2Index
Torne.Jugador27 = w3Index
Torne.Jugador28 = w4Index
Torne.Jugador29 = w5Index
Torne.Jugador30 = w6Index
Torne.Jugador31 = w7index
Torne.Jugador32 = w8index
UserList(Torne.Jugador25).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador26).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador27).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador28).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador29).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador30).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador31).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador32).flags.DueleandoTorneo4 = True
Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/CUATRO3 " Then
If UserList(userindex).flags.EsRolesMaster Then Exit Sub
rData = Right$(rData, Len(rData) - 9)
Dim mIndex As Integer
Dim m2Index As Integer
Dim m3Index As Integer
Dim m4Index As Integer
Dim m5Index As Integer
Dim m6Index As Integer
Dim m7index As Integer
Dim m8index As Integer
mIndex = NameIndex(ReadField(1, rData, 64))
m2Index = NameIndex(ReadField(2, rData, 64))
m3Index = NameIndex(ReadField(3, rData, 64))
m4Index = NameIndex(ReadField(4, rData, 64))
m5Index = NameIndex(ReadField(5, rData, 64))
m6Index = NameIndex(ReadField(6, rData, 64))
m7index = NameIndex(ReadField(7, rData, 64))
m8index = NameIndex(ReadField(8, rData, 64))
If Arena3 = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 3." & FONTTYPE_INFO)
    Exit Sub
    End If
If mIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf m2Index <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf m3Index <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El tercer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf m4Index <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El cuarto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf m5Index <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El quinto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf m6Index <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El sexto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf m7index <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El septimo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf m8index <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El octavo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
End If
If mIndex = m2Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y segundo)" & FONTTYPE_INFO)
Exit Sub
ElseIf mIndex = m3Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y tercero)" & FONTTYPE_INFO)
Exit Sub
ElseIf mIndex = m4Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf mIndex = m5Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf mIndex = m6Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf m2Index = m7index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y tercer)" & FONTTYPE_INFO)
Exit Sub
ElseIf m2Index = m4Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf m2Index = m5Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf m2Index = m6Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf m2Index = m7index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y septimo)" & FONTTYPE_INFO)
Exit Sub
ElseIf m2Index = m8index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y octavo)" & FONTTYPE_INFO)
Exit Sub
ElseIf m3Index = m4Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf m3Index = m5Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf m3Index = m6Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf m4Index = m5Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf m4Index = m6Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf m5Index = m6Index Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (quinto y sexto)" & FONTTYPE_INFO)
Exit Sub
End If
Call SendData(ToAll, userindex, 0, "||Torneo: Arena 3: " & UserList(mIndex).name & ", " & UserList(m2Index).name & ", " & UserList(m3Index).name & " y " & UserList(m4Index).name & " VS " & UserList(m5Index).name & ", " & UserList(m6Index).name & ", " & UserList(m7index).name & " y " & UserList(m8index).name & "~230~230~0~1~0")
Call WarpUserChar(mIndex, 81, 40, 44, True)
Call WarpUserChar(m2Index, 81, 40, 45, True)
Call WarpUserChar(m3Index, 81, 41, 44, True)
Call WarpUserChar(m4Index, 81, 41, 45, True)
Call WarpUserChar(m5Index, 81, 58, 62, True)
Call WarpUserChar(m6Index, 81, 58, 63, True)
Call WarpUserChar(m7index, 81, 59, 62, True)
Call WarpUserChar(m8index, 81, 59, 63, True)
CuentaArena = 4
Arena3 = True
Torne.Jugador17 = mIndex
Torne.Jugador18 = m2Index
Torne.Jugador19 = m3Index
Torne.Jugador20 = m4Index
Torne.Jugador21 = m5Index
Torne.Jugador22 = m6Index
Torne.Jugador23 = m7index
Torne.Jugador24 = m8index
UserList(Torne.Jugador17).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador18).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador19).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador20).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador21).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador22).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador23).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador24).flags.DueleandoTorneo4 = True
Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/CUATRO2 " Then
If UserList(userindex).flags.EsRolesMaster Then Exit Sub
rData = Right$(rData, Len(rData) - 9)
Dim uIndex As Integer
Dim uuIndex As Integer
Dim uuuIndex As Integer
Dim uuuuIndex As Integer
Dim uuuuuIndex As Integer
Dim uuuuuuIndex As Integer
Dim uuuuuuuindex As Integer
Dim uuuuuuuuindex As Integer
uIndex = NameIndex(ReadField(1, rData, 64))
uuIndex = NameIndex(ReadField(2, rData, 64))
uuuIndex = NameIndex(ReadField(3, rData, 64))
uuuuIndex = NameIndex(ReadField(4, rData, 64))
uuuuuIndex = NameIndex(ReadField(5, rData, 64))
uuuuuuIndex = NameIndex(ReadField(6, rData, 64))
uuuuuuuindex = NameIndex(ReadField(7, rData, 64))
uuuuuuuuindex = NameIndex(ReadField(8, rData, 64))
If Arena2 = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 2." & FONTTYPE_INFO)
    Exit Sub
    End If
If uIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf uuIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf uuuIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El tercer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf uuuuIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El cuarto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf uuuuuIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El quinto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf uuuuuuIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El sexto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf uuuuuuuindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El septimo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf uuuuuuuuindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El octavo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
End If
If uIndex = uuIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y segundo)" & FONTTYPE_INFO)
Exit Sub
ElseIf uIndex = uuuIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y tercero)" & FONTTYPE_INFO)
Exit Sub
ElseIf uIndex = uuuuIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf uIndex = uuuuuIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf uIndex = uuuuuuIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf uuIndex = uuuuuuuindex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y tercer)" & FONTTYPE_INFO)
Exit Sub
ElseIf uuIndex = uuuuIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf uuIndex = uuuuuIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf uuIndex = uuuuuuIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf uuIndex = uuuuuuuindex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y septimo)" & FONTTYPE_INFO)
Exit Sub
ElseIf uuIndex = uuuuuuuindex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y octavo)" & FONTTYPE_INFO)
Exit Sub
ElseIf uuuIndex = uuuuIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf uuuIndex = uuuuuIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf uuuIndex = uuuuuuIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf uuuuIndex = uuuuuIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf uuuuIndex = uuuuuuIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf uuuuuIndex = uuuuuuIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (quinto y sexto)" & FONTTYPE_INFO)
Exit Sub
End If
Call SendData(ToAll, userindex, 0, "||Torneo: Arena 2: " & UserList(uIndex).name & ", " & UserList(uuIndex).name & ", " & UserList(uuuIndex).name & " y " & UserList(uuuuIndex).name & " VS " & UserList(uuuuuIndex).name & ", " & UserList(uuuuuuIndex).name & ", " & UserList(uuuuuuuindex).name & " y " & UserList(uuuuuuuuindex).name & "~230~230~0~1~0")
Call WarpUserChar(uIndex, 81, 69, 15, True)
Call WarpUserChar(uuIndex, 81, 69, 16, True)
Call WarpUserChar(uuuIndex, 81, 70, 15, True)
Call WarpUserChar(uuuuIndex, 81, 70, 16, True)
Call WarpUserChar(uuuuuIndex, 81, 87, 34, True)
Call WarpUserChar(uuuuuuIndex, 81, 87, 35, True)
Call WarpUserChar(uuuuuuuindex, 81, 88, 34, True)
Call WarpUserChar(uuuuuuuuindex, 81, 88, 35, True)
CuentaArena = 4
Arena2 = True
Torne.Jugador9 = uIndex
Torne.Jugador10 = uuIndex
Torne.Jugador11 = uuuIndex
Torne.Jugador12 = uuuuIndex
Torne.Jugador13 = uuuuuIndex
Torne.Jugador14 = uuuuuuIndex
Torne.Jugador15 = uuuuuuuindex
Torne.Jugador16 = uuuuuuuuindex
UserList(Torne.Jugador9).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador10).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador11).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador12).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador13).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador14).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador15).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador16).flags.DueleandoTorneo4 = True
Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/CUATRO1 " Then
If UserList(userindex).flags.EsRolesMaster Then Exit Sub
rData = Right$(rData, Len(rData) - 9)
Dim unoindex As Integer
Dim dosIndex As Integer
Dim tresIndex As Integer
Dim cuatroIndex As Integer
Dim cincoIndex As Integer
Dim seisIndex As Integer
Dim sieteindex As Integer
Dim ochoindex As Integer
unoindex = NameIndex(ReadField(1, rData, 64))
dosIndex = NameIndex(ReadField(2, rData, 64))
tresIndex = NameIndex(ReadField(3, rData, 64))
cuatroIndex = NameIndex(ReadField(4, rData, 64))
cincoIndex = NameIndex(ReadField(5, rData, 64))
seisIndex = NameIndex(ReadField(6, rData, 64))
sieteindex = NameIndex(ReadField(7, rData, 64))
ochoindex = NameIndex(ReadField(8, rData, 64))
If Arena1 = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 1." & FONTTYPE_INFO)
    Exit Sub
    End If
If unoindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf dosIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf tresIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El tercer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf cuatroIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El cuarto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf cincoIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El quinto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf seisIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El sexto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf sieteindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El septimo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf ochoindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El octavo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
End If
If unoindex = dosIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y segundo)" & FONTTYPE_INFO)
Exit Sub
ElseIf unoindex = tresIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y tercero)" & FONTTYPE_INFO)
Exit Sub
ElseIf unoindex = cuatroIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf unoindex = cincoIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf unoindex = seisIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf dosIndex = sieteindex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y tercer)" & FONTTYPE_INFO)
Exit Sub
ElseIf dosIndex = cuatroIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf dosIndex = cincoIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf dosIndex = seisIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf dosIndex = sieteindex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y septimo)" & FONTTYPE_INFO)
Exit Sub
ElseIf dosIndex = ochoindex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y octavo)" & FONTTYPE_INFO)
Exit Sub
ElseIf tresIndex = cuatroIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tresIndex = cincoIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tresIndex = seisIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf cuatroIndex = cincoIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf cuatroIndex = seisIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf cincoIndex = seisIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (quinto y sexto)" & FONTTYPE_INFO)
Exit Sub
End If
Call SendData(ToAll, userindex, 0, "||Torneo: Arena 1: " & UserList(unoindex).name & ", " & UserList(dosIndex).name & ", " & UserList(tresIndex).name & " y " & UserList(cuatroIndex).name & " VS " & UserList(cincoIndex).name & ", " & UserList(seisIndex).name & ", " & UserList(sieteindex).name & " y " & UserList(ochoindex).name & "~230~230~0~1~0")
Call WarpUserChar(unoindex, 81, 41, 15, True)
Call WarpUserChar(dosIndex, 81, 41, 16, True)
Call WarpUserChar(tresIndex, 81, 40, 15, True)
Call WarpUserChar(cuatroIndex, 81, 40, 16, True)
Call WarpUserChar(cincoIndex, 81, 58, 34, True)
Call WarpUserChar(seisIndex, 81, 58, 35, True)
Call WarpUserChar(sieteindex, 81, 59, 34, True)
Call WarpUserChar(ochoindex, 81, 59, 35, True)
CuentaArena = 4
Arena1 = True
Torne.Jugador1 = unoindex
Torne.Jugador2 = dosIndex
Torne.Jugador3 = tresIndex
Torne.Jugador4 = cuatroIndex
Torne.Jugador5 = cincoIndex
Torne.Jugador6 = seisIndex
Torne.Jugador7 = sieteindex
Torne.Jugador8 = ochoindex
UserList(Torne.Jugador1).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador2).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador3).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador4).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador5).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador6).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador7).flags.DueleandoTorneo4 = True
UserList(Torne.Jugador8).flags.DueleandoTorneo4 = True
Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/CUATROF " Then
If UserList(userindex).flags.EsRolesMaster Then Exit Sub
rData = Right$(rData, Len(rData) - 9)
unoindex = NameIndex(ReadField(1, rData, 64))
dosIndex = NameIndex(ReadField(2, rData, 64))
tresIndex = NameIndex(ReadField(3, rData, 64))
cuatroIndex = NameIndex(ReadField(4, rData, 64))
cincoIndex = NameIndex(ReadField(5, rData, 64))
seisIndex = NameIndex(ReadField(6, rData, 64))
sieteindex = NameIndex(ReadField(7, rData, 64))
ochoindex = NameIndex(ReadField(8, rData, 64))
If Arena1 = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 1." & FONTTYPE_INFO)
    Exit Sub
    End If
If unoindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf dosIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf tresIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El tercer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf cuatroIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El cuarto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf cincoIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El quinto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf seisIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El sexto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf sieteindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El septimo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf ochoindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El octavo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
End If
If unoindex = dosIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y segundo)" & FONTTYPE_INFO)
Exit Sub
ElseIf unoindex = tresIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y tercero)" & FONTTYPE_INFO)
Exit Sub
ElseIf unoindex = cuatroIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf unoindex = cincoIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf unoindex = seisIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf dosIndex = sieteindex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y tercer)" & FONTTYPE_INFO)
Exit Sub
ElseIf dosIndex = cuatroIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf dosIndex = cincoIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf dosIndex = seisIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf dosIndex = sieteindex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y septimo)" & FONTTYPE_INFO)
Exit Sub
ElseIf dosIndex = ochoindex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y octavo)" & FONTTYPE_INFO)
Exit Sub
ElseIf tresIndex = cuatroIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tresIndex = cincoIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tresIndex = seisIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf cuatroIndex = cincoIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf cuatroIndex = seisIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf cincoIndex = seisIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (quinto y sexto)" & FONTTYPE_INFO)
Exit Sub
End If
Call SendData(ToAll, userindex, 0, "||Torneo: Final: " & UserList(unoindex).name & ", " & UserList(dosIndex).name & ", " & UserList(tresIndex).name & " y " & UserList(cuatroIndex).name & " VS " & UserList(cincoIndex).name & ", " & UserList(seisIndex).name & ", " & UserList(sieteindex).name & " y " & UserList(ochoindex).name & "~230~230~0~1~0")
Call WarpUserChar(unoindex, 81, 41, 15, True)
Call WarpUserChar(dosIndex, 81, 41, 16, True)
Call WarpUserChar(tresIndex, 81, 40, 15, True)
Call WarpUserChar(cuatroIndex, 81, 40, 16, True)
Call WarpUserChar(cincoIndex, 81, 58, 34, True)
Call WarpUserChar(seisIndex, 81, 58, 35, True)
Call WarpUserChar(sieteindex, 81, 59, 34, True)
Call WarpUserChar(ochoindex, 81, 59, 35, True)
CuentaArena = 4
Arena1 = True
Torne.Jugador1 = unoindex
Torne.Jugador2 = dosIndex
Torne.Jugador3 = tresIndex
Torne.Jugador4 = cuatroIndex
Torne.Jugador5 = cincoIndex
Torne.Jugador6 = seisIndex
Torne.Jugador7 = sieteindex
Torne.Jugador8 = ochoindex
UserList(Torne.Jugador1).flags.DueleandoFinal4 = True
UserList(Torne.Jugador2).flags.DueleandoFinal4 = True
UserList(Torne.Jugador3).flags.DueleandoFinal4 = True
UserList(Torne.Jugador4).flags.DueleandoFinal4 = True
UserList(Torne.Jugador5).flags.DueleandoFinal4 = True
UserList(Torne.Jugador6).flags.DueleandoFinal4 = True
UserList(Torne.Jugador7).flags.DueleandoFinal4 = True
UserList(Torne.Jugador8).flags.DueleandoFinal4 = True
Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/TRES1 " Then
If UserList(userindex).flags.EsRolesMaster Then Exit Sub
rData = Right$(rData, Len(rData) - 7)
Dim ttIndex As Integer
Dim tttIndex As Integer
Dim ttttIndex As Integer
Dim tttttIndex As Integer
Dim ttttttIndex As Integer
tIndex = NameIndex(ReadField(1, rData, 64))
ttIndex = NameIndex(ReadField(2, rData, 64))
tttIndex = NameIndex(ReadField(3, rData, 64))
ttttIndex = NameIndex(ReadField(4, rData, 64))
tttttIndex = NameIndex(ReadField(5, rData, 64))
ttttttIndex = NameIndex(ReadField(6, rData, 64))
If Arena1 = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 1." & FONTTYPE_INFO)
    Exit Sub
    End If
If tIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf ttIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf tttIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El tercer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf ttttIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El cuarto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf tttttIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El quinto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf ttttttIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El sexto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
End If
If tIndex = ttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y segundo)" & FONTTYPE_INFO)
Exit Sub
ElseIf tIndex = tttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y tercero)" & FONTTYPE_INFO)
Exit Sub
ElseIf tIndex = ttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tIndex = tttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tIndex = ttttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ttIndex = tttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y tercer)" & FONTTYPE_INFO)
Exit Sub
ElseIf ttIndex = ttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ttIndex = tttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ttIndex = ttttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tttIndex = ttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tttIndex = tttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tttIndex = ttttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ttttIndex = tttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ttttIndex = ttttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tttttIndex = ttttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (quinto y sexto)" & FONTTYPE_INFO)
Exit Sub
End If
Call SendData(ToAll, userindex, 0, "||Torneo: Arena 1: " & UserList(tIndex).name & ", " & UserList(ttIndex).name & " y " & UserList(tttIndex).name & " VS " & UserList(ttttIndex).name & ", " & UserList(tttttIndex).name & " y " & UserList(ttttttIndex).name & "~230~230~0~1~0")
Call WarpUserChar(tIndex, 81, 41, 15, True)
Call WarpUserChar(ttIndex, 81, 41, 16, True)
Call WarpUserChar(tttIndex, 81, 40, 16, True)
Call WarpUserChar(ttttIndex, 81, 59, 34, True)
Call WarpUserChar(tttttIndex, 81, 59, 35, True)
Call WarpUserChar(ttttttIndex, 81, 58, 35, True)
CuentaArena = 4
Arena1 = True
Torne.Jugador1 = tIndex
Torne.Jugador2 = ttIndex
Torne.Jugador3 = tttIndex
Torne.Jugador4 = ttttIndex
Torne.Jugador5 = tttttIndex
Torne.Jugador6 = ttttttIndex
UserList(Torne.Jugador1).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador2).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador3).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador4).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador5).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador6).flags.DueleandoTorneo3 = True
Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/TRESF " Then
If UserList(userindex).flags.EsRolesMaster Then Exit Sub
rData = Right$(rData, Len(rData) - 7)
tIndex = NameIndex(ReadField(1, rData, 64))
ttIndex = NameIndex(ReadField(2, rData, 64))
tttIndex = NameIndex(ReadField(3, rData, 64))
ttttIndex = NameIndex(ReadField(4, rData, 64))
tttttIndex = NameIndex(ReadField(5, rData, 64))
ttttttIndex = NameIndex(ReadField(6, rData, 64))
If Arena1 = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 1." & FONTTYPE_INFO)
    Exit Sub
    End If
If tIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf ttIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf tttIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El tercer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf ttttIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El cuarto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf tttttIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El quinto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf ttttttIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El sexto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
End If
If tIndex = ttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y segundo)" & FONTTYPE_INFO)
Exit Sub
ElseIf tIndex = tttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y tercero)" & FONTTYPE_INFO)
Exit Sub
ElseIf tIndex = ttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tIndex = tttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tIndex = ttttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ttIndex = tttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y tercer)" & FONTTYPE_INFO)
Exit Sub
ElseIf ttIndex = ttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ttIndex = tttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ttIndex = ttttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tttIndex = ttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tttIndex = tttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tttIndex = ttttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ttttIndex = tttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ttttIndex = ttttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf tttttIndex = ttttttIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (quinto y sexto)" & FONTTYPE_INFO)
Exit Sub
End If
Call SendData(ToAll, userindex, 0, "||Torneo: Final: " & UserList(tIndex).name & ", " & UserList(ttIndex).name & " y " & UserList(tttIndex).name & " VS " & UserList(ttttIndex).name & ", " & UserList(tttttIndex).name & " y " & UserList(ttttttIndex).name & "~230~230~0~1~0")
Call WarpUserChar(tIndex, 81, 41, 15, True)
Call WarpUserChar(ttIndex, 81, 41, 16, True)
Call WarpUserChar(tttIndex, 81, 40, 16, True)
Call WarpUserChar(ttttIndex, 81, 59, 34, True)
Call WarpUserChar(tttttIndex, 81, 59, 35, True)
Call WarpUserChar(ttttttIndex, 81, 58, 35, True)
CuentaArena = 4
Arena1 = True
Torne.Jugador1 = tIndex
Torne.Jugador2 = ttIndex
Torne.Jugador3 = tttIndex
Torne.Jugador4 = ttttIndex
Torne.Jugador5 = tttttIndex
Torne.Jugador6 = ttttttIndex
UserList(Torne.Jugador1).flags.DueleandoFinal3 = True
UserList(Torne.Jugador2).flags.DueleandoFinal3 = True
UserList(Torne.Jugador3).flags.DueleandoFinal3 = True
UserList(Torne.Jugador4).flags.DueleandoFinal3 = True
UserList(Torne.Jugador5).flags.DueleandoFinal3 = True
UserList(Torne.Jugador6).flags.DueleandoFinal3 = True
Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/TRES2 " Then
If UserList(userindex).flags.EsRolesMaster Then Exit Sub
rData = Right$(rData, Len(rData) - 7)
Dim xindex As Integer
Dim xxIndex As Integer
Dim xxxIndex As Integer
Dim xxxxIndex As Integer
Dim xxxxxIndex As Integer
Dim xxxxxxIndex As Integer
xindex = NameIndex(ReadField(1, rData, 64))
xxIndex = NameIndex(ReadField(2, rData, 64))
xxxIndex = NameIndex(ReadField(3, rData, 64))
xxxxIndex = NameIndex(ReadField(4, rData, 64))
xxxxxIndex = NameIndex(ReadField(5, rData, 64))
xxxxxxIndex = NameIndex(ReadField(6, rData, 64))
If Arena2 = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 2." & FONTTYPE_INFO)
    Exit Sub
    End If
If xindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf xxIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf xxxIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El tercer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf xxxxIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El cuarto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf xxxxxIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El quinto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf xxxxxxIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El sexto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
End If
If xindex = xxIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y segundo)" & FONTTYPE_INFO)
Exit Sub
ElseIf xindex = xxxIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y tercero)" & FONTTYPE_INFO)
Exit Sub
ElseIf xindex = xxxxIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf xindex = xxxxxIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf xindex = xxxxxxIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf xxIndex = xxxIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y tercer)" & FONTTYPE_INFO)
Exit Sub
ElseIf xxIndex = xxxxIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf xxIndex = xxxxxIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf xxIndex = xxxxxxIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf xxxIndex = xxxxIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf xxxIndex = xxxxxIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf xxxIndex = xxxxxxIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf xxxxIndex = xxxxxIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf xxxxIndex = xxxxxxIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf xxxxxIndex = xxxxxxIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (quinto y sexto)" & FONTTYPE_INFO)
Exit Sub
End If
Call SendData(ToAll, userindex, 0, "||Torneo: Arena 2: " & UserList(xindex).name & ", " & UserList(xxIndex).name & " y " & UserList(xxxIndex).name & " VS " & UserList(xxxxIndex).name & ", " & UserList(xxxxxIndex).name & " y " & UserList(xxxxxxIndex).name & "~230~230~0~1~0")
Call WarpUserChar(xindex, 81, 70, 15, True)
Call WarpUserChar(xxIndex, 81, 70, 16, True)
Call WarpUserChar(xxxIndex, 81, 69, 16, True)
Call WarpUserChar(xxxxIndex, 81, 88, 34, True)
Call WarpUserChar(xxxxxIndex, 81, 88, 35, True)
Call WarpUserChar(xxxxxxIndex, 81, 87, 35, True)
CuentaArena = 4
Arena2 = True
Torne.Jugador7 = xindex
Torne.Jugador8 = xxIndex
Torne.Jugador9 = xxxIndex
Torne.Jugador10 = xxxxIndex
Torne.Jugador11 = xxxxxIndex
Torne.Jugador12 = xxxxxxIndex
UserList(Torne.Jugador7).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador8).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador9).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador10).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador11).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador12).flags.DueleandoTorneo3 = True
Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/TRES3 " Then
If UserList(userindex).flags.EsRolesMaster Then Exit Sub
rData = Right$(rData, Len(rData) - 7)
Dim qindex As Integer
Dim qqIndex As Integer
Dim qqqIndex As Integer
Dim qqqqIndex As Integer
Dim qqqqqIndex As Integer
Dim qqqqqqIndex As Integer
qindex = NameIndex(ReadField(1, rData, 64))
qqIndex = NameIndex(ReadField(2, rData, 64))
qqqIndex = NameIndex(ReadField(3, rData, 64))
qqqqIndex = NameIndex(ReadField(4, rData, 64))
qqqqqIndex = NameIndex(ReadField(5, rData, 64))
qqqqqqIndex = NameIndex(ReadField(6, rData, 64))
If Arena3 = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 3." & FONTTYPE_INFO)
    Exit Sub
    End If
If qindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf qqIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf qqqIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El tercer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf qqqqIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El cuarto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf qqqqqIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El quinto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf qqqqqqIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El sexto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
End If
If qindex = qqIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y segundo)" & FONTTYPE_INFO)
Exit Sub
ElseIf qindex = qqqIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y tercero)" & FONTTYPE_INFO)
Exit Sub
ElseIf qindex = qqqqIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf qindex = qqqqqIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf qindex = qqqqqqIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf qqIndex = qqqIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y tercer)" & FONTTYPE_INFO)
Exit Sub
ElseIf qqIndex = qqqqIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf qqIndex = qqqqqIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf qqIndex = qqqqqqIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf qqqIndex = qqqqIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf qqqIndex = qqqqqIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf qqqIndex = qqqqqqIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf qqqqIndex = qqqqqIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf qqqqIndex = qqqqqqIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf qqqqqIndex = qqqqqqIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (quinto y sexto)" & FONTTYPE_INFO)
Exit Sub
End If
Call SendData(ToAll, userindex, 0, "||Torneo: Arena 3: " & UserList(qindex).name & ", " & UserList(qqIndex).name & "y " & UserList(qqqIndex).name & " VS " & UserList(qqqqIndex).name & ", " & UserList(qqqqqIndex).name & " y " & UserList(qqqqqqIndex).name & "~230~230~0~1~0")
Call WarpUserChar(qindex, 81, 41, 44, True)
Call WarpUserChar(qqIndex, 81, 41, 45, True)
Call WarpUserChar(qqqIndex, 81, 40, 45, True)
Call WarpUserChar(qqqqIndex, 81, 59, 62, True)
Call WarpUserChar(qqqqqIndex, 81, 59, 63, True)
Call WarpUserChar(qqqqqqIndex, 81, 58, 63, True)
CuentaArena = 4
Arena3 = True
Torne.Jugador13 = qindex
Torne.Jugador14 = qqIndex
Torne.Jugador15 = qqqIndex
Torne.Jugador16 = qqqqIndex
Torne.Jugador17 = qqqqqIndex
Torne.Jugador18 = qqqqqqIndex
UserList(Torne.Jugador13).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador14).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador15).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador16).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador17).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador18).flags.DueleandoTorneo3 = True
Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/TRES4 " Then
If UserList(userindex).flags.EsRolesMaster Then Exit Sub
rData = Right$(rData, Len(rData) - 7)
Dim cindex As Integer
Dim ccIndex As Integer
Dim cccIndex As Integer
Dim ccccIndex As Integer
Dim cccccIndex As Integer
Dim ccccccIndex As Integer
cindex = NameIndex(ReadField(1, rData, 64))
ccIndex = NameIndex(ReadField(2, rData, 64))
cccIndex = NameIndex(ReadField(3, rData, 64))
ccccIndex = NameIndex(ReadField(4, rData, 64))
cccccIndex = NameIndex(ReadField(5, rData, 64))
ccccccIndex = NameIndex(ReadField(6, rData, 64))
If Arena4 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 4." & FONTTYPE_INFO)
 Exit Sub
End If
If cindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf ccIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf cccIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El tercer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf ccccIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El cuarto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf cccccIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El quinto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf ccccccIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El sexto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
End If
If cindex = ccIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y segundo)" & FONTTYPE_INFO)
Exit Sub
ElseIf cindex = cccIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y tercero)" & FONTTYPE_INFO)
Exit Sub
ElseIf cindex = ccccIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf cindex = cccccIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf cindex = ccccccIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ccIndex = cccIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y tercer)" & FONTTYPE_INFO)
Exit Sub
ElseIf ccIndex = ccccIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ccIndex = cccccIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ccIndex = ccccccIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf cccIndex = ccccIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf cccIndex = cccccIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf cccIndex = ccccccIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ccccIndex = cccccIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y quinto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ccccIndex = ccccccIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (cuarto y sexto)" & FONTTYPE_INFO)
Exit Sub
ElseIf cccccIndex = ccccccIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (quinto y sexto)" & FONTTYPE_INFO)
Exit Sub
End If
Call SendData(ToAll, userindex, 0, "||Torneo: Arena 4: " & UserList(cindex).name & ", " & UserList(ccIndex).name & " y " & UserList(cccIndex).name & " VS " & UserList(ccccIndex).name & ", " & UserList(cccccIndex).name & " y " & UserList(ccccccIndex).name & "~230~230~0~1~0")
Call WarpUserChar(cindex, 81, 70, 44, True)
Call WarpUserChar(ccIndex, 81, 70, 45, True)
Call WarpUserChar(cccIndex, 81, 69, 45, True)
Call WarpUserChar(ccccIndex, 81, 88, 62, True)
Call WarpUserChar(cccccIndex, 81, 88, 63, True)
Call WarpUserChar(ccccccIndex, 81, 87, 63, True)
CuentaArena = 4
Arena4 = True
Torne.Jugador19 = cindex
Torne.Jugador20 = ccIndex
Torne.Jugador21 = cccIndex
Torne.Jugador22 = ccccIndex
Torne.Jugador23 = cccccIndex
Torne.Jugador24 = ccccccIndex
UserList(Torne.Jugador19).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador20).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador21).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador22).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador23).flags.DueleandoTorneo3 = True
UserList(Torne.Jugador24).flags.DueleandoTorneo3 = True
Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/VERSUSF " Then
If UserList(userindex).flags.EsRolesMaster Then Exit Sub
rData = Right$(rData, Len(rData) - 9)
Dim gIndex As Integer
Dim ggindex As Integer
Dim gggIndex As Integer
Dim ggggIndex As Integer
gIndex = NameIndex(ReadField(1, rData, 64))
ggindex = NameIndex(ReadField(2, rData, 64))
gggIndex = NameIndex(ReadField(3, rData, 64))
ggggIndex = NameIndex(ReadField(4, rData, 64))
If Arena1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 1." & FONTTYPE_INFO)
 Exit Sub
End If
If gIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf ggindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf gggIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El tercer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf ggggIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El cuarto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
End If
If gIndex = ggindex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y segundo)" & FONTTYPE_INFO)
Exit Sub
ElseIf gIndex = gggIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y tercero)" & FONTTYPE_INFO)
Exit Sub
End If
If gIndex = ggggIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ggindex = gggIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y tercer)" & FONTTYPE_INFO)
Exit Sub
ElseIf ggindex = ggggIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf gggIndex = ggggIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y cuarto)" & FONTTYPE_INFO)
Exit Sub
End If
Call SendData(ToAll, userindex, 0, "||Torneo: Final: " & UserList(gIndex).name & " y " & UserList(ggindex).name & " VS " & UserList(gggIndex).name & " y " & UserList(ggggIndex).name & "." & "~230~230~0~1~0")
Call WarpUserChar(gIndex, 81, 40, 15, True)
Call WarpUserChar(ggindex, 81, 41, 16, True)
Call WarpUserChar(gggIndex, 81, 59, 34, True)
Call WarpUserChar(ggggIndex, 81, 58, 35, True)
CuentaArena = 4
Arena1 = True
Torne.Jugador1 = gIndex
Torne.Jugador2 = ggindex
Torne.Jugador3 = gggIndex
Torne.Jugador4 = ggggIndex
UserList(Torne.Jugador1).flags.DueleandoFinal2 = True
UserList(Torne.Jugador2).flags.DueleandoFinal2 = True
UserList(Torne.Jugador3).flags.DueleandoFinal2 = True
UserList(Torne.Jugador4).flags.DueleandoFinal2 = True
Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/VERSUS1 " Then
If UserList(userindex).flags.EsRolesMaster Then Exit Sub
rData = Right$(rData, Len(rData) - 9)
Dim eIndex As Integer
Dim eeindex As Integer
Dim eeeIndex As Integer
Dim eeeeIndex As Integer
eIndex = NameIndex(ReadField(1, rData, 64))
eeindex = NameIndex(ReadField(2, rData, 64))
eeeIndex = NameIndex(ReadField(3, rData, 64))
eeeeIndex = NameIndex(ReadField(4, rData, 64))
If Arena1 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 1." & FONTTYPE_INFO)
 Exit Sub
End If
If eIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf eeindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf eeeIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El tercer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf eeeeIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El cuarto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
End If
If eIndex = eeindex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y segundo)" & FONTTYPE_INFO)
Exit Sub
ElseIf eIndex = eeeIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y tercero)" & FONTTYPE_INFO)
Exit Sub
End If
If eIndex = eeeeIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf eeindex = eeeIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y tercer)" & FONTTYPE_INFO)
Exit Sub
ElseIf eeindex = eeeeIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf eeeIndex = eeeeIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y cuarto)" & FONTTYPE_INFO)
Exit Sub
End If
Call SendData(ToAll, userindex, 0, "||Torneo: Arena 1: " & UserList(eIndex).name & " y " & UserList(eeindex).name & " VS " & UserList(eeeIndex).name & " y " & UserList(eeeeIndex).name & "." & "~230~230~0~1~0")
Call WarpUserChar(eIndex, 81, 40, 15, True)
Call WarpUserChar(eeindex, 81, 41, 16, True)
Call WarpUserChar(eeeIndex, 81, 59, 34, True)
Call WarpUserChar(eeeeIndex, 81, 58, 35, True)
CuentaArena = 4
Arena1 = True
Torne.Jugador1 = eIndex
Torne.Jugador2 = eeindex
Torne.Jugador3 = eeeIndex
Torne.Jugador4 = eeeeIndex
UserList(Torne.Jugador1).flags.DueleandoTorneo2 = True
UserList(Torne.Jugador2).flags.DueleandoTorneo2 = True
UserList(Torne.Jugador3).flags.DueleandoTorneo2 = True
UserList(Torne.Jugador4).flags.DueleandoTorneo2 = True
Exit Sub
End If



If UCase$(Left$(rData, 9)) = "/VERSUS2 " Then
If UserList(userindex).flags.EsRolesMaster Then Exit Sub
rData = Right$(rData, Len(rData) - 9)
Dim zindex As Integer
Dim zzindex As Integer
Dim zzzIndex As Integer
Dim zzzzIndex As Integer
zindex = NameIndex(ReadField(1, rData, 64))
zzindex = NameIndex(ReadField(2, rData, 64))
zzzIndex = NameIndex(ReadField(3, rData, 64))
zzzzIndex = NameIndex(ReadField(4, rData, 64))
If Arena2 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 2." & FONTTYPE_INFO)
 Exit Sub
End If
If zindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf zzindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf zzzIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El tercer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf zzzzIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El cuarto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
End If
If zindex = zzindex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y segundo)" & FONTTYPE_INFO)
Exit Sub
ElseIf zindex = zzzIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y tercero)" & FONTTYPE_INFO)
Exit Sub
End If
If zindex = zzzzIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf zzindex = zzzIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y tercer)" & FONTTYPE_INFO)
Exit Sub
ElseIf zzindex = zzzzIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf zzzIndex = zzzzIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y cuarto)" & FONTTYPE_INFO)
Exit Sub
End If
Call SendData(ToAll, userindex, 0, "||Torneo: Arena 2: " & UserList(zindex).name & " y " & UserList(zzindex).name & " VS " & UserList(zzzIndex).name & " y " & UserList(zzzzIndex).name & "." & "~230~230~0~1~0")
Call WarpUserChar(zindex, 81, 69, 15, True)
Call WarpUserChar(zzindex, 81, 70, 16, True)
Call WarpUserChar(zzzIndex, 81, 88, 34, True)
Call WarpUserChar(zzzzIndex, 81, 87, 35, True)
CuentaArena = 4
Arena2 = True
Torne.Jugador5 = zindex
Torne.Jugador6 = zzindex
Torne.Jugador7 = zzzIndex
Torne.Jugador8 = zzzzIndex
UserList(Torne.Jugador5).flags.DueleandoTorneo2 = True
UserList(Torne.Jugador6).flags.DueleandoTorneo2 = True
UserList(Torne.Jugador7).flags.DueleandoTorneo2 = True
UserList(Torne.Jugador8).flags.DueleandoTorneo2 = True
Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/VERSUS3 " Then
If UserList(userindex).flags.EsRolesMaster Then Exit Sub
rData = Right$(rData, Len(rData) - 9)
Dim yindex As Integer
Dim yyindex As Integer
Dim yyyIndex As Integer
Dim yyyyIndex As Integer
yindex = NameIndex(ReadField(1, rData, 64))
yyindex = NameIndex(ReadField(2, rData, 64))
yyyIndex = NameIndex(ReadField(3, rData, 64))
yyyyIndex = NameIndex(ReadField(4, rData, 64))
If Arena3 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 3." & FONTTYPE_INFO)
 Exit Sub
End If
If yindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf yyindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf yyyIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El tercer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf yyyyIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El cuarto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
End If
If yindex = yyindex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y segundo)" & FONTTYPE_INFO)
Exit Sub
ElseIf yindex = yyyIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y tercero)" & FONTTYPE_INFO)
Exit Sub
End If
If yindex = yyyyIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf yyindex = yyyIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y tercer)" & FONTTYPE_INFO)
Exit Sub
ElseIf yyindex = yyyyIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf yyyIndex = yyyyIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y cuarto)" & FONTTYPE_INFO)
Exit Sub
End If
Call SendData(ToAll, userindex, 0, "||Torneo: Arena 3: " & UserList(yindex).name & " y " & UserList(yyindex).name & " VS " & UserList(yyyIndex).name & " y " & UserList(yyyyIndex).name & "." & "~230~230~0~1~0")
Call WarpUserChar(yindex, 81, 41, 44, True)
Call WarpUserChar(yyindex, 81, 40, 45, True)
Call WarpUserChar(yyyIndex, 81, 59, 62, True)
Call WarpUserChar(yyyyIndex, 81, 58, 63, True)
CuentaArena = 4
Arena3 = True
Torne.Jugador9 = yindex
Torne.Jugador10 = yyindex
Torne.Jugador11 = yyyIndex
Torne.Jugador12 = yyyyIndex
UserList(Torne.Jugador9).flags.DueleandoTorneo2 = True
UserList(Torne.Jugador10).flags.DueleandoTorneo2 = True
UserList(Torne.Jugador11).flags.DueleandoTorneo2 = True
UserList(Torne.Jugador12).flags.DueleandoTorneo2 = True
Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/VERSUS4 " Then
If UserList(userindex).flags.EsRolesMaster Then Exit Sub
rData = Right$(rData, Len(rData) - 9)
Dim fIndex As Integer
Dim ffindex As Integer
Dim fffIndex As Integer
Dim ffffIndex As Integer
fIndex = NameIndex(ReadField(1, rData, 64))
ffindex = NameIndex(ReadField(2, rData, 64))
fffIndex = NameIndex(ReadField(3, rData, 64))
ffffIndex = NameIndex(ReadField(4, rData, 64))
If Arena4 = True Then
Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 4." & FONTTYPE_INFO)
 Exit Sub
End If
If fIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf ffindex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf fffIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El tercer usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
ElseIf ffffIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||El cuarto usuario tipeado a sumonear no se encuentra online." & FONTTYPE_INFO)
Exit Sub
End If
If fIndex = ffindex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y segundo)" & FONTTYPE_INFO)
Exit Sub
ElseIf fIndex = fffIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y tercero)" & FONTTYPE_INFO)
Exit Sub
End If
If fIndex = ffffIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (primero y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf ffindex = fffIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y tercer)" & FONTTYPE_INFO)
Exit Sub
ElseIf ffindex = ffffIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (segundo y cuarto)" & FONTTYPE_INFO)
Exit Sub
ElseIf fffIndex = ffffIndex Then
Call SendData(SendTarget.toindex, userindex, 0, "||¡No puedes sumonear al mismo usuario dos veces! (tercero y cuarto)" & FONTTYPE_INFO)
Exit Sub
End If
Call SendData(ToAll, userindex, 0, "||Torneo: Arena 4: " & UserList(fIndex).name & " y " & UserList(ffindex).name & " VS " & UserList(fffIndex).name & " y " & UserList(ffffIndex).name & "." & "~230~230~0~1~0")
Call WarpUserChar(fIndex, 81, 70, 44, True)
Call WarpUserChar(ffindex, 81, 69, 45, True)
Call WarpUserChar(fffIndex, 81, 88, 62, True)
Call WarpUserChar(ffffIndex, 81, 87, 63, True)
CuentaArena = 4
Arena4 = True
Torne.Jugador13 = fIndex
Torne.Jugador14 = ffindex
Torne.Jugador15 = fffIndex
Torne.Jugador16 = ffffIndex
UserList(Torne.Jugador13).flags.DueleandoTorneo2 = True
UserList(Torne.Jugador14).flags.DueleandoTorneo2 = True
UserList(Torne.Jugador15).flags.DueleandoTorneo2 = True
UserList(Torne.Jugador16).flags.DueleandoTorneo2 = True
Exit Sub
End If

If UCase$(Left$(rData, 8)) = "/PELEA1 " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 8)
    
    Dim ptindex As Integer
    tIndex = NameIndex(rData)
    ptindex = NameIndex(rData)
 
    tIndex = NameIndex(ReadField(1, rData, 64))
    ptindex = NameIndex(ReadField(2, rData, 64))
 
     If tIndex = ptindex Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No puede combatir un usuario contra sí mismo." & FONTTYPE_INFO)
        Exit Sub
    End If
 
    If ptindex <= 0 And tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Los usuarios tipeados no estan online." & FONTTYPE_INFO)
        Exit Sub
  End If
 
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
 
     If ptindex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Arena1 = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 1." & FONTTYPE_INFO)
    Exit Sub
    End If
 
   ' If tIndex = ttIndex Then
   '     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puede combatir un único usuario." & FONTTYPE_INFO)
   '     Exit Sub
   ' End If
 
    Call SendData(ToAll, userindex, 0, "||Torneo: Arena 1: " & UserList(tIndex).name & " VS " & UserList(ptindex).name & "." & "~230~230~0~1~0")
 
    Call LogGM(UserList(userindex).name, "/Pelea " & UserList(tIndex).name & " - " & UserList(ptindex).name, False)
 
    Call WarpUserChar(tIndex, 81, 40, 15, True)   'El primero en el comando
    Call WarpUserChar(ptindex, 81, 59, 34, True) 'El segundo en el comando
    CuentaArena = 4
    Arena1 = True
    Torne.Jugador1 = tIndex
    Torne.Jugador2 = ptindex
    UserList(Torne.Jugador1).flags.DueleandoTorneo = True
    UserList(Torne.Jugador2).flags.DueleandoTorneo = True
Exit Sub
End If
 

If UCase$(Left$(rData, 8)) = "/PELEAF " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 8)
    
    tIndex = NameIndex(rData)
    ptindex = NameIndex(rData)
 
    tIndex = NameIndex(ReadField(1, rData, 64))
    ptindex = NameIndex(ReadField(2, rData, 64))
 
     If tIndex = ptindex Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No puede combatir un usuario contra sí mismo." & FONTTYPE_INFO)
        Exit Sub
    End If
 
    If ptindex <= 0 And tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Los usuarios tipeados no estan online." & FONTTYPE_INFO)
        Exit Sub
  End If
 
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
 
     If ptindex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Arena1 = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 1." & FONTTYPE_INFO)
    Exit Sub
    End If
 
   ' If tIndex = ttIndex Then
   '     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puede combatir un único usuario." & FONTTYPE_INFO)
   '     Exit Sub
   ' End If
 
    Call SendData(ToAll, userindex, 0, "||Torneo: Final: " & UserList(tIndex).name & " VS " & UserList(ptindex).name & "." & "~230~230~0~1~0")
 
    Call LogGM(UserList(userindex).name, "/Pelea " & UserList(tIndex).name & " - " & UserList(ptindex).name, False)
 
    Call WarpUserChar(tIndex, 81, 40, 15, True)   'El primero en el comando
    Call WarpUserChar(ptindex, 81, 59, 34, True) 'El segundo en el comando
    CuentaArena = 4
    Arena1 = True
    Torne.Jugador1 = tIndex
    Torne.Jugador2 = ptindex
    UserList(Torne.Jugador1).flags.DueleandoFinal = True
    UserList(Torne.Jugador2).flags.DueleandoFinal = True
Exit Sub
End If

If UCase$(Left$(rData, 8)) = "/PELEA2 " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 8)
    
    tIndex = NameIndex(rData)
    ptindex = NameIndex(rData)
 
    tIndex = NameIndex(ReadField(1, rData, 64))
    ptindex = NameIndex(ReadField(2, rData, 64))
 
     If tIndex = ptindex Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No puede combatir un usuario contra sí mismo." & FONTTYPE_INFO)
        Exit Sub
    End If
 
    If ptindex <= 0 And tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Los usuarios tipeados no estan online." & FONTTYPE_INFO)
        Exit Sub
  End If
 
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
 
     If ptindex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Arena2 = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 2." & FONTTYPE_INFO)
    Exit Sub
    End If
 
   ' If tIndex = ttIndex Then
   '     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puede combatir un único usuario." & FONTTYPE_INFO)
   '     Exit Sub
   ' End If
 
    Call SendData(ToAll, userindex, 0, "||Torneo: Arena 2: " & UserList(tIndex).name & " VS " & UserList(ptindex).name & "." & "~230~230~0~1~0")
 
    Call LogGM(UserList(userindex).name, "/Pelea " & UserList(tIndex).name & " - " & UserList(ptindex).name, False)
 
    Call WarpUserChar(tIndex, 81, 69, 15, True)   'El primero en el comando
    Call WarpUserChar(ptindex, 81, 88, 34, True) 'El segundo en el comando
    CuentaArena = 4
    Arena2 = True
    Torne.Jugador3 = tIndex
    Torne.Jugador4 = ptindex
    UserList(Torne.Jugador3).flags.DueleandoTorneo = True
    UserList(Torne.Jugador4).flags.DueleandoTorneo = True
Exit Sub
End If

If UCase$(Left$(rData, 8)) = "/PELEA3 " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 8)
    
    tIndex = NameIndex(rData)
    ptindex = NameIndex(rData)
 
    tIndex = NameIndex(ReadField(1, rData, 64))
    ptindex = NameIndex(ReadField(2, rData, 64))
 
     If tIndex = ptindex Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No puede combatir un usuario contra sí mismo." & FONTTYPE_INFO)
        Exit Sub
    End If
 
    If ptindex <= 0 And tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Los usuarios tipeados no estan online." & FONTTYPE_INFO)
        Exit Sub
  End If
 
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
 
     If ptindex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Arena3 = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 3." & FONTTYPE_INFO)
    Exit Sub
    End If
 
   ' If tIndex = ttIndex Then
   '     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puede combatir un único usuario." & FONTTYPE_INFO)
   '     Exit Sub
   ' End If
 
    Call SendData(ToAll, userindex, 0, "||Torneo: Arena 3: " & UserList(tIndex).name & " VS " & UserList(ptindex).name & "." & "~230~230~0~1~0")
 
    Call LogGM(UserList(userindex).name, "/Pelea " & UserList(tIndex).name & " - " & UserList(ptindex).name, False)
 
    Call WarpUserChar(tIndex, 81, 40, 44, True)   'El primero en el comando
    Call WarpUserChar(ptindex, 81, 59, 62, True) 'El segundo en el comando
    CuentaArena = 4
    Arena3 = True
    Torne.Jugador5 = tIndex
    Torne.Jugador6 = ptindex
    UserList(Torne.Jugador5).flags.DueleandoTorneo = True
    UserList(Torne.Jugador6).flags.DueleandoTorneo = True
Exit Sub
End If

If UCase$(Left$(rData, 8)) = "/PELEA4 " Then
    If UserList(userindex).flags.EsRolesMaster Then Exit Sub
    rData = Right$(rData, Len(rData) - 8)
    
    tIndex = NameIndex(rData)
    ptindex = NameIndex(rData)
 
    tIndex = NameIndex(ReadField(1, rData, 64))
    ptindex = NameIndex(ReadField(2, rData, 64))
 
     If tIndex = ptindex Then
        Call SendData(SendTarget.toindex, userindex, 0, "||No puede combatir un usuario contra sí mismo." & FONTTYPE_INFO)
        Exit Sub
    End If
 
    If ptindex <= 0 And tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Los usuarios tipeados no estan online." & FONTTYPE_INFO)
        Exit Sub
  End If
 
    If tIndex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||El primer usuario tipeado no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
 
     If ptindex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||El segundo usuario tipeado no esta online." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Arena4 = True Then
    Call SendData(SendTarget.toindex, userindex, 0, "||Ya están combatiendo en la Arena 4." & FONTTYPE_INFO)
    Exit Sub
    End If
 
   ' If tIndex = ttIndex Then
   '     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puede combatir un único usuario." & FONTTYPE_INFO)
   '     Exit Sub
   ' End If
 
    Call SendData(ToAll, userindex, 0, "||Torneo: Arena 4: " & UserList(tIndex).name & " VS " & UserList(ptindex).name & "." & "~230~230~0~1~0")
 
    Call LogGM(UserList(userindex).name, "/Pelea " & UserList(tIndex).name & " - " & UserList(ptindex).name, False)
 
    Call WarpUserChar(tIndex, 81, 69, 44, True)   'El primero en el comando
    Call WarpUserChar(ptindex, 81, 88, 62, True) 'El segundo en el comando
    CuentaArena = 4
    Arena4 = True
    Torne.Jugador7 = tIndex
    Torne.Jugador8 = ptindex
    UserList(Torne.Jugador7).flags.DueleandoTorneo = True
    UserList(Torne.Jugador8).flags.DueleandoTorneo = True
Exit Sub
End If

If UCase$(Left$(rData, 9)) = "/SUMTODOS" Then
rData = Right$(rData, Len(rData) - 9)
For i = 1 To NumUsers
Call WarpUserChar(i, UserList(userindex).pos.Map, UserList(userindex).pos.X, UserList(userindex).pos.Y, False)
Call SendData(SendTarget.ToAll, userindex, 0, "||Servidor> " & UserList(userindex).name & " sumoneo a todos los usuarios.." & FONTTYPE_INFO)
Next i
Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/ARENA1" Then
rData = Right$(rData, Len(rData) - 7)
Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & " desbugeo la Arena 1, ahora está lista para usarse." & FONTTYPE_INFO)
Arena1 = False
Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/CVCOFF" Then
rData = Right$(rData, Len(rData) - 7)
Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & " desbugeo el CVC." & FONTTYPE_INFO)
CvcFunciona = False
Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/ARENA2" Then
rData = Right$(rData, Len(rData) - 7)
Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & " desbugeo la Arena 2, ahora está lista para usarse." & FONTTYPE_INFO)
Arena2 = False
Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/ARENA3" Then
rData = Right$(rData, Len(rData) - 7)
Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & " desbugeo la Arena 3, ahora está lista para usarse." & FONTTYPE_INFO)
Arena3 = False
Exit Sub
End If

If UCase$(Left$(rData, 7)) = "/ARENA4" Then
rData = Right$(rData, Len(rData) - 7)
Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(userindex).name & " desbugeo la Arena 4, ahora está lista para usarse." & FONTTYPE_INFO)
Arena4 = False
Exit Sub
End If

If UCase$(Left$(rData, 16)) = "/ACTIVARCONSULTA" Then
If UserList(userindex).flags.Privilegios = PlayerType.VIP Then Exit Sub
rData = Right$(rData, Len(rData) - 16)
Call SendData(SendTarget.ToAll, 0, 0, "||" & UserList(userindex).name & " está atendiendo consultas, escribe /CONSULTA." & FONTTYPE_GUILD)
Call WarpUserChar(userindex, 19, 29, 53, True)
Consulta = True
HayConsulta = False
Exit Sub
End If

If UCase(Left(rData, 6)) = "/SOLU " Then
If UserList(userindex).flags.Privilegios = PlayerType.VIP Then Exit Sub
rData = Right$(rData, Len(rData) - 6)
Dim fex As String
fex = NameIndex(rData)

        If fex <= 0 Then
        Call SendData(SendTarget.toindex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
        End If
        
        If Consulta = False Then Exit Sub

Call SendData(SendTarget.ToAll, 0, 0, "||" & UserList(userindex).name & " está atendiendo consultas, escribe /CONSULTA." & FONTTYPE_GUILD)
HayConsulta = False
    Call WarpUserChar(fex, PosUserConsulta.Map, PosUserConsulta.X, PosUserConsulta.Y)
Exit Sub
End If

If UCase$(Left$(rData, 19)) = "/DESACTIVARCONSULTA" Then
If UserList(userindex).flags.Privilegios = PlayerType.VIP Then Exit Sub
rData = Right$(rData, Len(rData) - 19)
Call SendData(SendTarget.ToAll, 0, 0, "||" & UserList(userindex).name & " desactivó las consultas." & FONTTYPE_GUILD)
Consulta = False
Exit Sub
End If

'LIDER
If UCase$(Left$(rData, 5)) = "/LID " Then
rData = Right$(rData, Len(rData) - 5)
tIndex = NameIndex(ReadField(1, rData, 64))
Arg1 = ReadField(2, rData, 64)
'Offline
If tIndex <= 0 Then
Call SendData(SendTarget.toindex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
Exit Sub
End If
'Es posible?
If Val(Arg1) <= 0 Or Val(Arg1) > 2 Then
 
Call SendData(SendTarget.toindex, userindex, 0, "||Utilizava lores de 1 a 2. (1 Luz - 2 Sombras)" & FONTTYPE_WARNING)
Exit Sub
End If
 
Select Case Val(Arg1)
Case 1
UserList(tIndex).ConsejoInfo.LiderConsejo = 1
 
Call SendData(SendTarget.ToAll, 0, 0, "||" & UserList(tIndex).name & " es el nuevo lider del Consejo de la Luz." & FONTTYPE_CELESTEN)
Call SendData(SendTarget.toindex, tIndex, 0, "||Has sido convertido en lider del consejo de La Luz. " & FONTTYPE_INFO)
Call WriteVar(CharPath & UserList(tIndex).name & ".chr", "CONSEJO", "LIDERCONSEJO", UserList(tIndex).ConsejoInfo.LiderConsejo)
Case 2
UserList(tIndex).ConsejoInfo.LiderConsejoCaos = 1
 
Call SendData(SendTarget.ToAll, 0, 0, "||" & UserList(tIndex).name & " es el nuevo lider del Concillio del Caos." & FONTTYPE_ROJON)
Call SendData(SendTarget.toindex, tIndex, 0, "||Has sido convertido en lider del consejo de Las Sombras. " & FONTTYPE_INFO)
Call WriteVar(CharPath & UserList(tIndex).name & ".chr", "CONSEJO", "LIDERCONSEJOCAOS", UserList(tIndex).ConsejoInfo.LiderConsejoCaos)
End Select
End If
'LIDER

    Procesado = False
End Sub
