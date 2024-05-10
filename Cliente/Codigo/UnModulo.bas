Attribute VB_Name = "UnModulo"
Option Explicit
 
Public MiCuerpo As Integer, MiCabeza As Integer
 
Private Sub DrawGrafico(Grh As Grh, ByVal x As Byte, ByVal Y As Byte)
 
Dim r2 As RECT, auxr As RECT
Dim iGrhIndex As Integer
 
    If Grh.GrhIndex <= 0 Then Exit Sub
    
    iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
        
    With r2
        .Left = GrhData(iGrhIndex).sX
        .Top = GrhData(iGrhIndex).sY
        .Right = .Left + GrhData(iGrhIndex).pixelWidth
        .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
    End With
    
    With auxr
        .Left = 0
        .Top = 0
        .Right = 50
        .Bottom = 65
    End With
    
    BackBufferSurface.BltFast x, Y, SurfaceDB.Surface(GrhData(iGrhIndex).FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    Call BackBufferSurface.BltToDC(frmCrearPersonaje.PlayerView.hDC, auxr, auxr)
 
End Sub
 
Sub DibujarCPJ(ByVal MyBody As Integer, ByVal MyHead As Integer)
 
Dim Grh As Grh
Dim Pos As Integer
Dim r2 As RECT
 
    With r2
        .Left = 0
        .Top = 0
        .Right = 50
        .Bottom = 65
    End With
    
    BackBufferSurface.BltColorFill r2, vbBlack
    
    Grh = BodyData(MyBody).Walk(3)
    Call DrawGrafico(Grh, 12, 15)
    
    Pos = BodyData(MyBody).HeadOffset.Y + GrhData(GrhData(Grh.GrhIndex).Frames(1)).pixelHeight
    Grh = HeadData(MyHead).Head(3)
    Call DrawGrafico(Grh, 17, Pos)
    
    frmCrearPersonaje.PlayerView.Refresh
    
End Sub
 
Sub DameOpciones()
 
Dim i As Integer
 
If frmCrearPersonaje.lstGenero.listIndex < 0 Or frmCrearPersonaje.lstRaza.listIndex < 0 Then
    frmCrearPersonaje.cabeza.Enabled = False
ElseIf frmCrearPersonaje.lstGenero.listIndex <> -1 And frmCrearPersonaje.lstRaza.listIndex <> -1 Then
    frmCrearPersonaje.cabeza.Enabled = True
End If
 
frmCrearPersonaje.cabeza.Clear
    
Select Case frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.listIndex)
   Case "Hombre"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.listIndex)
            Case "Humano"
                For i = 1 To 30
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 1
            Case "Elfo"
                For i = 101 To 112
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 2
            Case "Elfo Oscuro"
                For i = 201 To 209
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 3
            Case "Enano"
                For i = 301 To 305
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 52
            Case "Gnomo"
                For i = 401 To 406
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 52
            Case Else
                UserHead = 1
                MiCuerpo = 1
        End Select
   Case "Mujer"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.listIndex)
            Case "Humano"
                For i = 70 To 76
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 1
            Case "Elfo"
                For i = 170 To 176
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 2
            Case "Elfo Oscuro"
                For i = 271 To 280
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 3
            Case "Gnomo"
                For i = 470 To 474
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 52
            Case "Enano"
                For i = 370 To 372
                    frmCrearPersonaje.cabeza.AddItem i
                Next i
                MiCuerpo = 52
            Case Else
                frmCrearPersonaje.cabeza.AddItem "70"
                MiCuerpo = 1
        End Select
End Select
 
frmCrearPersonaje.PlayerView.Cls
 
End Sub

