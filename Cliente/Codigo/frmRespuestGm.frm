VERSION 5.00
Begin VB.Form frmRespuestaGM 
   BorderStyle     =   0  'None
   Caption         =   "Respuesta GM"
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRespuestGm.frx":0000
   ScaleHeight     =   3960
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   375
      Left            =   4680
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   4215
   End
End
Attribute VB_Name = "frmRespuestaGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
TieneParaResponder = True
frmWriteMSG.Show
End Sub

Private Sub Image1_Click()
Unload Me
End Sub
