VERSION 5.00
Begin VB.Form frmCanjesDonacion 
   BorderStyle     =   0  'None
   Caption         =   "Canjes Nightmare"
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCanjesDonacion.frx":0000
   ScaleHeight     =   5985
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   3345
      Left            =   780
      TabIndex        =   5
      Top             =   840
      Width           =   1970
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   485
      Left            =   5280
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   600
      Width           =   485
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "¡Canjeo de Donaciones!"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label DescripcionCanje 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   3720
      TabIndex        =   8
      Top             =   4620
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1800
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   360
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Hit 
      Alignment       =   2  'Center
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
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   2460
      Width           =   855
   End
   Begin VB.Label DM 
      Alignment       =   2  'Center
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
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   3375
      Width           =   855
   End
   Begin VB.Label lblptos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1845
      TabIndex        =   3
      Top             =   5340
      Width           =   735
   End
   Begin VB.Label RM 
      Alignment       =   2  'Center
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
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   2910
      Width           =   855
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
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
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   3855
      Width           =   855
   End
   Begin VB.Label Stats 
      Alignment       =   2  'Center
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
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   0
      Top             =   1965
      Width           =   855
   End
End
Attribute VB_Name = "frmCanjesDonacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Image1_Click()
If List1.Text = "Montura de Dragon Dorado" Then
If MsgBox("¿Está seguro que desea canjear una Montura de Dragon Dorado?", vbYesNo) = vbYes Then
Call SendData("DONA01")
End If
End If
If List1.Text = "Montura de Dragon Rojo" Then
If MsgBox("¿Está seguro que desea canjear una Montura de Dragon Rojo?.", vbYesNo) = vbYes Then
Call SendData("DONA02")
End If
End If
If List1.Text = "Túnica de los Campeones" Then
If MsgBox("¿Está seguro que desea canjear una Túnica de los Campeones?", vbYesNo) = vbYes Then
Call SendData("DONA03")
End If
End If
If List1.Text = "Túnica de los Campeones (E/G)" Then
If MsgBox("¿Está seguro que desea canjear una Túnica de los Campeones (E/G)?", vbYesNo) = vbYes Then
Call SendData("DONA04")
End If
End If
If List1.Text = "Túnica de los Heroes" Then
If MsgBox("¿Está seguro que desea canjear una Túnica de los Heroes?", vbYesNo) = vbYes Then
Call SendData("DONA05")
End If
End If
If List1.Text = "Túnica de los Heroes (E/G)" Then
If MsgBox("¿Está seguro que desea canjear una Túnica de los Heroes (E/G)?", vbYesNo) = vbYes Then
Call SendData("DONA06")
End If
End If
If List1.Text = "Túnica de la Luz" Then
If MsgBox("¿Está seguro que desea canjear una Túnica de la Luz?", vbYesNo) = vbYes Then
Call SendData("DONA07")
End If
End If
If List1.Text = "Túnica de la Luz (E/G)" Then
If MsgBox("¿Está seguro que desea canjear una Túnica de la Luz (E/G)?", vbYesNo) = vbYes Then
Call SendData("DONA08")
End If
End If
If List1.Text = "Túnica de la Oscuridad" Then
If MsgBox("¿Está seguro que desea canjear una Túnica de la Oscuridad?", vbYesNo) = vbYes Then
Call SendData("DONA09")
End If
End If
If List1.Text = "Túnica de la Oscuridad (E/G)" Then
If MsgBox("¿Está seguro que desea canjear una Túnica de la Oscuridad (E/G)?", vbYesNo) = vbYes Then
Call SendData("DONA10")
End If
End If
If List1.Text = "VIP" Then
If MsgBox("¿Está seguro que desea canjear los puntos vip?", vbYesNo) = vbYes Then
Call SendData("DONA11")
End If
End If
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If frmRendimiento.Transp.value = 0 Then
Else
    Transparencia Me.hwnd, 150
End If
Picture = LoadPicture(App.Path & "\Graficos\Interfaces\Canjes.jpg")
List1.AddItem "Montura de Dragon Dorado"
List1.AddItem "Montura de Dragon Rojo"
List1.AddItem "Túnica de los Campeones"
List1.AddItem "Túnica de los Campeones (E/G)"
List1.AddItem "Túnica de los Heroes"
List1.AddItem "Túnica de los Heroes (E/G)"
List1.AddItem "Túnica de la Luz"
List1.AddItem "Túnica de la Luz (E/G)"
List1.AddItem "Túnica de la Oscuridad"
List1.AddItem "Túnica de la Oscuridad (E/G)"
List1.AddItem "VIP"
End Sub
Private Sub List1_Click()
If List1.Text = "Montura de Dragon Dorado" Then
Picture1.Picture = LoadPicture(DirGraficos & "12152.bmp")
Puntos.Caption = "1500"
Stats.Caption = "8/8"
RM.Caption = "10/10"
DM.Caption = "N/A"
Hit.Caption = "20/20"
End If
If List1.Text = "Montura de Dragon Rojo" Then
Picture1.Picture = LoadPicture(DirGraficos & "17903.bmp")
Puntos.Caption = "1000"
Stats.Caption = "10/10"
RM.Caption = "10/10"
DM.Caption = "N/A"
Hit.Caption = "10/10"
End If
If List1.Text = "Túnica de los Campeones" Then
Picture1.Picture = LoadPicture(DirGraficos & "13659.bmp")
Puntos.Caption = "250"
Stats.Caption = "45/50"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Túnica de los Campeones (E/G)" Then
Picture1.Picture = LoadPicture(DirGraficos & "13671.bmp")
Puntos.Caption = "250"
Stats.Caption = "40/45"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Túnica de los Heroes" Then
Picture1.Picture = LoadPicture(DirGraficos & "13641.bmp")
Puntos.Caption = "500"
Stats.Caption = "40/45"
RM.Caption = "10/15"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Túnica de los Heroes (E/G)" Then
Picture1.Picture = LoadPicture(DirGraficos & "13804.bmp")
Puntos.Caption = "500"
Stats.Caption = "45/50"
RM.Caption = "10/15"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Túnica de la Luz" Then
Picture1.Picture = LoadPicture(DirGraficos & "13880.bmp")
Puntos.Caption = "750"
Stats.Caption = "48/53"
RM.Caption = "15/20"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Túnica de la Luz (E/G)" Then
Picture1.Picture = LoadPicture(DirGraficos & "13880.bmp")
Puntos.Caption = "750"
Stats.Caption = "48/53"
RM.Caption = "15/20"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Túnica de la Oscuridad" Then
Picture1.Picture = LoadPicture(DirGraficos & "13882.bmp")
Puntos.Caption = "750"
Stats.Caption = "48/53"
RM.Caption = "15/20"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Túnica de la Oscuridad (E/G)" Then
Picture1.Picture = LoadPicture(DirGraficos & "13882.bmp")
Puntos.Caption = "750"
Stats.Caption = "48/53"
RM.Caption = "15/20"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "VIP" Then
Picture1.Picture = LoadPicture(DirGraficos & "13820.bmp")
Puntos.Caption = "2500"
Stats.Caption = "N/A"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If

'descripcion by zaiko

Select Case List1.Text
 
Case Is = "Montura de Dragon Dorado"
       DescripcionCanje.Caption = "Item exclusivo para donadores."
Case Is = "Montura de Dragon Rojo"
       DescripcionCanje.Caption = "Item exclusivo para donadores."
Case Is = "Túnica de los Campeones"
       DescripcionCanje.Caption = "Item exclusivo para donadores."
Case Is = "Túnica de los Campeones (E/G)"
       DescripcionCanje.Caption = "Item exclusivo para donadores."
Case Is = "Túnica de los Heroes"
       DescripcionCanje.Caption = "Item exclusivo para donadores."
Case Is = "Túnica de los Heroes (E/G)"
       DescripcionCanje.Caption = "Item exclusivo para donadores."
Case Is = "Túnica de la Luz"
       DescripcionCanje.Caption = "Item exclusivo para donadores."
Case Is = "Túnica de la Luz (E/G)"
       DescripcionCanje.Caption = "Item exclusivo para donadores."
Case Is = "Túnica de la Oscuridad"
       DescripcionCanje.Caption = "Item exclusivo para donadores."
Case Is = "Túnica de la Oscuridad (E/G)"
       DescripcionCanje.Caption = "Item exclusivo para donadores."
Case Is = "VIP"
       DescripcionCanje.Caption = "Canjeando esto, se te otorgaran los Puntos VIP necesarios para hacerse vip."
       
End Select

End Sub
