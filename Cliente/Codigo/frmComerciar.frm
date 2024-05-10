VERSION 5.00
Begin VB.Form frmComerciar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmComerciar.frx":0000
   ScaleHeight     =   510
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrNumber 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   405
      Top             =   405
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3240
      TabIndex        =   8
      Text            =   "1"
      Top             =   6960
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   885
      ScaleHeight     =   480
      ScaleWidth      =   465
      TabIndex        =   2
      Top             =   1605
      Width           =   465
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3930
      Index           =   1
      Left            =   3735
      TabIndex        =   1
      Top             =   2595
      Width           =   2490
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3930
      Index           =   0
      Left            =   750
      TabIndex        =   0
      Top             =   2580
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   1620
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   2880
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image cmdMasMenos 
      Height          =   390
      Index           =   1
      Left            =   3870
      Top             =   6900
      Width           =   165
   End
   Begin VB.Image cmdMasMenos 
      Height          =   390
      Index           =   0
      Left            =   2955
      Top             =   6885
      Width           =   165
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   6480
      Top             =   165
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   420
      Index           =   1
      Left            =   4245
      MouseIcon       =   "frmComerciar.frx":0046
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6855
      Width           =   2130
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   0
      Left            =   585
      MouseIcon       =   "frmComerciar.frx":0198
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6840
      Width           =   2145
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   7395
      TabIndex        =   7
      Top             =   1170
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   7200
      TabIndex        =   6
      Top             =   765
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   5775
      TabIndex        =   5
      Top             =   1530
      Width           =   60
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   5685
      TabIndex        =   4
      Top             =   1905
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   1620
      TabIndex        =   3
      Top             =   1530
      Width           =   45
   End
End
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*****************************
'*****************************
'*****      Samke       ******
'*****************************
'**************************************************
'**************************************************
'*****      SoHnsalxixon_u2@hotmail.com      ******
'**************************************************
'**************************************************

Private m_Interval As Integer
Private m_Number As Integer
Private m_Increment As Integer
Public LastIndex1 As Integer
Public LastIndex2 As Integer

Private Sub cantidad_Change()
    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
    
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = 1
    End If
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdMasMenos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Call Audio.PlayWave("click.wav")

Select Case Index
    Case 0
        cmdMasMenos(Index).Picture = LoadPicture(App.Path & "\Graficos\Interfaces\menos-down.jpg")
        cmdMasMenos(Index).Tag = "1"
        cantidad.Text = Str((Val(cantidad.Text) - 1))
        m_Increment = -1
    Case 1
        cmdMasMenos(Index).Picture = LoadPicture(App.Path & "\Graficos\Interfaces\mas-down.jpg")
        cmdMasMenos(Index).Tag = "1"
        m_Increment = 1
End Select

tmrNumber.Interval = 30
tmrNumber.Enabled = True

End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub
Private Sub Form_Load()
If frmRendimiento.Transp.value = 0 Then
Else
    Transparencia Me.hWnd, 150
End If
Me.Picture = LoadPicture(App.Path & "\Graficos\Interfaces\Comercio.jpg")
m_Number = 1
m_Interval = 30
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Image1(0).Tag = "1" Then
    Image1(0).Picture = Nothing
    Image1(0).Tag = "0"
End If

If Image1(1).Tag = "1" Then
    Image1(1).Picture = Nothing
    Image1(1).Tag = "0"
End If

If cmdMasMenos(0).Tag = "1" Then
    cmdMasMenos(0).Picture = Nothing
    cmdMasMenos(0).Tag = "0"
End If

If cmdMasMenos(1).Tag = "1" Then
    cmdMasMenos(1).Picture = Nothing
    cmdMasMenos(1).Tag = "0"
End If

If Image2.Tag = "1" Then
    Image2.Picture = Nothing
    Image2.Tag = "0"
End If

End Sub

Private Sub Image1_Click(Index As Integer)

Call Audio.PlayWave("click.wav")

If List1(Index).List(List1(Index).listIndex) = "Nada" Or _
   List1(Index).listIndex < 0 Then Exit Sub

Select Case Index
    Case 0
        frmComerciar.List1(0).SetFocus
        LastIndex1 = List1(0).listIndex
        If UserGLD >= NPCInventory(List1(0).listIndex + 1).Valor * Val(cantidad) Then
                SendData ("COMP" & "," & List1(0).listIndex + 1 & "," & cantidad.Text)
                
        Else
            AddtoRichTextBox frmMain.RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1
            Exit Sub
        End If
   Case 1
        LastIndex2 = List1(1).listIndex
        If Not Inventario.Equipped(List1(1).listIndex + 1) Then
            SendData ("VEND" & "," & List1(1).listIndex + 1 & "," & cantidad.Text)
        Else
            AddtoRichTextBox frmMain.RecTxt, "No podes vender el item porque lo estas usando.", 2, 51, 223, 1, 1
            Exit Sub
        End If
                
End Select
List1(0).Clear

List1(1).Clear

NPCInvDim = 0
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Index = 0 Then
    If Image1(Index).Tag = "0" Then
        Image1(Index).Picture = LoadPicture(App.Path & "\Graficos\Interfaces\comprar-over.jpg")
        Image1(Index).Tag = "1"
    End If
ElseIf Index = 1 Then
    If Image1(Index).Tag = "0" Then
        Image1(Index).Picture = LoadPicture(App.Path & "\Graficos\Interfaces\vender-over.jpg")
        Image1(Index).Tag = "1"
    End If
End If

End Sub

Private Sub cantidad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Audio.PlayWave("click.wav")
Image2.Picture = LoadPicture(App.Path & "\Graficos\Interfaces\salir-down.jpg")
Image2.Tag = "1"

End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
SendData ("FINCOM")
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Image2.Tag = "0" Then
    Image2.Picture = LoadPicture(App.Path & "\Graficos\Interfaces\salir-over.jpg")
    Image2.Tag = "1"
End If

End Sub

Private Sub List1_Click(Index As Integer)
Dim SR As RECT, DR As RECT, GrhIndex As Integer

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.Bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.Bottom = 32

Select Case Index
    Case 0
        Label1(0).Caption = NPCInventory(List1(0).listIndex + 1).Name
        Label1(1).Caption = NPCInventory(List1(0).listIndex + 1).Valor
        Label1(2).Caption = NPCInventory(List1(0).listIndex + 1).Amount
        GrhIndex = NPCInventory(List1(0).listIndex + 1).GrhIndex
        Select Case NPCInventory(List1(0).listIndex + 1).OBJType
            Case 2
                Label1(5).Caption = "Max Golpe: " & NPCInventory(List1(0).listIndex + 1).MaxHit
                Label1(6).Caption = "Min Golpe: " & NPCInventory(List1(0).listIndex + 1).MinHit
                Label1(5).Visible = True
                Label1(6).Visible = True
            Case 3
                Label1(5).Visible = False
                Label1(6).Caption = "Defensa: " & NPCInventory(List1(0).listIndex + 1).Def
                Label1(6).Visible = True
            Case 16
                Label1(5).Visible = False
                Label1(6).Caption = "Defensa: " & NPCInventory(List1(0).listIndex + 1).Def
                Label1(6).Visible = True
            Case 17
                Label1(5).Visible = False
                Label1(6).Caption = "Defensa: " & NPCInventory(List1(0).listIndex + 1).Def
                Label1(6).Visible = True
        End Select
        Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hDC, NPCInventory(List1(0).listIndex + 1).GrhIndex, SR, DR)
    Case 1
        Label1(0).Caption = Inventario.ItemName(List1(1).listIndex + 1)
        Label1(1).Caption = Inventario.Valor(List1(1).listIndex + 1)
        Label1(2).Caption = Inventario.Amount(List1(1).listIndex + 1)
        GrhIndex = Inventario.GrhIndex(List1(1).listIndex + 1)
        Select Case Inventario.OBJType(List1(1).listIndex + 1)
            Case 2
                Label1(5).Caption = "Max Golpe: " & Inventario.MaxHit(List1(1).listIndex + 1)
                Label1(6).Caption = "Min Golpe: " & Inventario.MinHit(List1(1).listIndex + 1)
                Label1(5).Visible = True
                Label1(6).Visible = True
            Case 3
                Label1(5).Visible = False
                Label1(6).Caption = "Defensa: " & Inventario.Def(List1(1).listIndex + 1)
                Label1(6).Visible = True
            Case 16
                Label1(5).Visible = False
                Label1(6).Caption = "Defensa: " & Inventario.Def(List1(1).listIndex + 1)
                Label1(6).Visible = True
            Case 17
                Label1(5).Visible = False
                Label1(6).Caption = "Defensa: " & Inventario.Def(List1(1).listIndex + 1)
                Label1(6).Visible = True
        End Select
        Call DrawGrhtoHdc(Picture1.hWnd, Picture1.hDC, Inventario.GrhIndex(List1(1).listIndex + 1), SR, DR)
End Select

Picture1.Refresh

End Sub

Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub tmrNumber_Timer()

Const MIN_NUMBER = 1
Const MAX_NUMBER = 10000

    m_Number = m_Number + m_Increment
    If m_Number < MIN_NUMBER Then
        m_Number = MIN_NUMBER
    ElseIf m_Number > MAX_NUMBER Then
        m_Number = MAX_NUMBER
    End If

    cantidad.Text = Format$(m_Number)
    
    If m_Interval > 1 Then
        m_Interval = m_Interval - 1
        tmrNumber.Interval = m_Interval
    End If
    
End Sub

Private Sub cmdMasMenos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
tmrNumber.Enabled = False
End Sub

Private Sub cmdMasMenos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
    Case 0
        If cmdMasMenos(Index).Tag = "0" Then
            cmdMasMenos(Index).Picture = LoadPicture(App.Path & "\Graficos\Interfaces\menos-over.jpg")
            cmdMasMenos(Index).Tag = "1"
        End If
    Case 1
        If cmdMasMenos(Index).Tag = "0" Then
            cmdMasMenos(Index).Picture = LoadPicture(App.Path & "\Graficos\Interfaces\mas-over.jpg")
            cmdMasMenos(Index).Tag = "1"
        End If
End Select

End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Index = 0 Then
    Image1(Index).Picture = LoadPicture(App.Path & "\Graficos\Interfaces\comprar-down.jpg")
    Image1(Index).Tag = "1"
ElseIf Index = 1 Then
    Image1(Index).Picture = LoadPicture(App.Path & "\Graficos\Interfaces\vender-down.jpg")
    Image1(Index).Tag = "1"
End If

End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub

