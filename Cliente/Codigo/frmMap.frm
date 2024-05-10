VERSION 5.00
Begin VB.Form frmMap 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   DrawStyle       =   5  'Transparent
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   Picture         =   "frmMap.frx":0000
   ScaleHeight     =   4725
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   150
      Left            =   840
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   150
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5760
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Transparencia Me.hwnd, 150

    Call MovemosUserMapa
    Picture = LoadPicture(App.Path & "\Graficos\Interfaces\seventhmapa.JPG")
End Sub
