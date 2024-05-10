Attribute VB_Name = "DrawPJenPicture"

Sub DibujaPJ(Grh As Grh, ByVal x As Integer, ByVal y As Integer, Index As Integer)
On Error Resume Next
Dim iGrhIndex As Integer
If Grh.grhindex <= 0 Then Exit Sub
iGrhIndex = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)

Call engine.GrhRenderToHdc(iGrhIndex, frmCuent.PJ(Index).hdc, x, y, True)

frmCuent.PJ(Index).Refresh

End Sub

Sub dibujamuerto(Index As Integer)

End Sub
Sub DibujarTodo(ByVal Index As Integer, Body As Integer, Head As Integer, Casco As Integer, Shield As Integer, Weapon As Integer, Baned As Integer, Nombre As String, LVL As Integer, Clase As String, Muerto As Integer)

Dim Grh As Grh
Dim Pos As Integer
Dim LoopC As Integer
Dim r As RECT
Dim r2 As RECT

Dim YBody As Integer
Dim YYY As Integer
Dim XBody As Integer
Dim BBody As Integer


With r2
    .Left = 0
  .Top = 0
   .Right = 150
  .bottom = 150
End With

'BackBufferSurface.BltColorFill r, 0



frmCuent.Nombre(Index).Caption = Nombre
frmCuent.CP(Index).Visible = False
frmCuent.Nombre(Index).Visible = True
frmCuent.Label2(Index).Visible = True
If LVL > 50 Then
    frmCuent.Label2(Index).Caption = "Nivel: 50 + " & LVL - 50
Else
    frmCuent.Label2(Index).Caption = "Nivel: " & LVL
End If
XBody = 29
YBody = 20
BBody = 35

'XBody = 12
'YBody = 15
'BBody = 17

If Muerto = 1 Then
    Body = 8
    Head = 500
    Arma = 2
    Shield = 2
    Weapon = 2
XBody = 28
YBody = 38
BBody = 35
    Call dibujamuerto(Index)
End If

Grh = BodyData(Body).Walk(3)
    
Call DibujaPJ(Grh, XBody, YBody, Index)

If Muerto = 0 Then YYY = BodyData(Body).HeadOffset.y
If Muerto = 1 Then YYY = -9

Pos = YYY + GrhData(GrhData(Grh.grhindex).Frames(Grh.FrameCounter)).pixelHeight
Grh = HeadData(Head).Head(3)
    If Baned = 1 Then
    Call dibujaban(Index, vbBlack)
    Call dibujaban(Index, vbRed)
      
End If


Call DibujaPJ(Grh, BBody, Pos + 2, Index)
    
If Casco <> 2 And Casco > 0 Then
    Grh = CascoAnimData(Casco).Head(3)
    Call DibujaPJ(Grh, BBody, Pos + 2, Index)
End If

If Weapon <> 2 And Weapon > 0 Then
    Grh = WeaponAnimData(Weapon).WeaponWalk(3)
    Call DibujaPJ(Grh, XBody, YBody, Index)
End If

If Shield <> 2 And Shield > 0 Then
    Grh = ShieldAnimData(Shield).ShieldWalk(3)
    Call DibujaPJ(Grh, XBody + 4, BBody - 13, Index)
End If

End Sub



Sub dibujaban(Index As Integer, color As Long)

Dim r2 As RECT, auxr As RECT

With r2
   .Left = 0
   .Top = 0
   .Right = 20
   .bottom = 20
End With

With auxr
    .Left = 0
  .Top = 0
   .Right = 150
  .bottom = 150
End With



End Sub

