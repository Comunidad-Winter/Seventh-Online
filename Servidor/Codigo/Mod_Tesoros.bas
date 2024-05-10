Attribute VB_Name = "Mod_Tesoros"
'******************************************************************************
'Zefron AO v0.1.0
'Mod_Tesoros.bas
'You can contact the programmer of Zefron AO at gabi.13@live.com.ar
'******************************************************************************
 
Public MapaTesoro As Integer
Public RecompenzaTesoro As Integer
Public MapaTesoroMap As Integer
Public MapaTesoroX As Integer
Public MapaTesoroY As Integer
Public TiempoTesoro As Integer
Public TesoroContando As Boolean
Public SepuedeDesenterrar As Boolean
Public Const LlaveTesoro As Integer = 1062 'num del mapa
Public ObjetoT As Obj
Public objetoCofreAbierto As Obj
 
Public Sub Tesoros()
       
    ObjetoT.Amount = 1
    ObjetoT.ObjIndex = 11 'Cofre Cerrado
   
    objetoCofreAbierto.Amount = 1
    objetoCofreAbierto.ObjIndex = 10 'Cofre abierto
    
MapaTesoro = RandomNumber(91, 95)
 
If MapaTesoro = 91 Then
    MapaTesoroMap = 91 ' mapa . Les dejo este ejemplo para que se guien
    MapaTesoroX = RandomNumber(62, 89) '  rango de posicion de X. Les dejo este ejemplo para que se guien
    MapaTesoroY = RandomNumber(22, 61) '  rango de posicion de Y. Les dejo este ejemplo para que se guien
ElseIf MapaTesoro = 92 Then
    MapaTesoroMap = 92 ' mapa . Les dejo este ejemplo para que se guien
    MapaTesoroX = RandomNumber(18, 69)  '  rango de posicion de X. Les dejo este ejemplo para que se guien
    MapaTesoroY = RandomNumber(58, 93) '  rango de posicion de Y. Les dejo este ejemplo para que se guien
ElseIf MapaTesoro = 93 Then
    MapaTesoroMap = 93 ' mapa . Les dejo este ejemplo para que se guien
    MapaTesoroX = RandomNumber(19, 42) '  rango de posicion de X. Les dejo este ejemplo para que se guien
    MapaTesoroY = RandomNumber(19, 35) '  rango de posicion de Y. Les dejo este ejemplo para que se guien
ElseIf MapaTesoro = 94 Then
    MapaTesoroMap = 94 ' mapa . Les dejo este ejemplo para que se guien
    MapaTesoroX = RandomNumber(39, 61) '  rango de posicion de X. Les dejo este ejemplo para que se guien
    MapaTesoroY = RandomNumber(85, 62) '  rango de posicion de Y. Les dejo este ejemplo para que se guien
ElseIf MapaTesoro = 95 Then
    MapaTesoroMap = 95 ' mapa . Les dejo este ejemplo para que se guien
    MapaTesoroX = RandomNumber(30, 73) '  rango de posicion de X. Les dejo este ejemplo para que se guien
    MapaTesoroY = RandomNumber(62, 85) '  rango de posicion de Y. Les dejo este ejemplo para que se guien
End If

    SepuedeDesenterrar = False
    TesoroContando = False
    TiempoTesoro = 30
    Call SendData(SendTarget.toall, 0, 0, "||Rondan noticias que hay un tesoro enterrado en el mapa " & MapaTesoroMap & " en las coordenadas " & MapaTesoroX & ", " & MapaTesoroY & "." & FONTTYPE_INFO)
End Sub
Public Sub DondeTesoros()
    Call SendData(SendTarget.toall, 0, 0, "||El tesoro se encuentra en el mapa " & MapaTesoro & " en las coordenadas " & MapaTesoroX & ", " & MapaTesoroY & "." & FONTTYPE_INFO)
End Sub
Public Sub CofreAbierto()
Call EraseObj(SendTarget.ToMap, userindex, MapaTesoroMap, 10000, MapaTesoroMap, MapaTesoroX, MapaTesoroY)
Call MakeObj(SendTarget.ToMap, 0, MapaTesoroMap, objetoCofreAbierto, MapaTesoroMap, MapaTesoroX, MapaTesoroY)
End Sub
