VERSION 5.00
Begin VB.Form frmMuerte 
   Caption         =   "Has sido Asesinado.."
   ClientHeight    =   2880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   Picture         =   "frmMuerte.frx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Seguir como fantasma"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Volver a la ciudad inicial"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   720
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "¡Has sido Asesinado!"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmMuerte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call SendData("/REGRESAR")
Me.Visible = False
End Sub

Private Sub Command2_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()
    Transparencia Me.hwnd, 150
End Sub
