VERSION 5.00
Begin VB.Form frmViajes 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   Picture         =   "frmViajes.frx":0000
   ScaleHeight     =   4470.199
   ScaleMode       =   0  'User
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Exit 
      Height          =   255
      Left            =   4130
      Top             =   120
      Width           =   210
   End
   Begin VB.Label Info 
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
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   825
      Width           =   3975
   End
   Begin VB.Image Travel 
      Height          =   555
      Index           =   4
      Left            =   585
      Top             =   3644
      Width           =   3345
   End
   Begin VB.Image Travel 
      Height          =   555
      Index           =   3
      Left            =   585
      Top             =   2880
      Width           =   3345
   End
   Begin VB.Image Travel 
      Height          =   555
      Index           =   2
      Left            =   585
      Top             =   2230
      Width           =   3345
   End
   Begin VB.Image Travel 
      Height          =   555
      Index           =   1
      Left            =   585
      Top             =   1530
      Width           =   3350
   End
End
Attribute VB_Name = "frmViajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************************
'Standelf - Menu de Viajes.
'******************************************************************************************************
'Menu de Viajes.
'******************************************************************************************************
'Otras Caracteristicas:
'******************************************************************************************************
'Standelf || 29-Octubre-2008
'******************************************************************************************************

Option Explicit

'Acomodar Form
Private Sub Form_Load()
    With frmViajes
        .Width = 4500
        .Height = 4500
        .Info.Caption = "Bienvenido al centro de Viajes. Haciendo click Izquierdo sobre un destino vera la información de cada lugar, Con Click Derecho podra Viajar."
    End With
End Sub

'Boton de Destino
Private Sub Travel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Call SendData("TRAVEL" & Index)
        Unload Me
    Else
        Call Mensaje(Index)
    End If
End Sub

'Acomodar Info
Private Function Mensaje(ByVal i As Integer)
    If Val(i) = 1 Then
        Info.Caption = "Ullathorpe es una pequeña ciudad en el centro del Mundo. Valor del Viaje: 500 Monedas."
    ElseIf Val(i) = 2 Then
        Info.Caption = "Banderbill es la ciudad mas grande del Mundo. Se encuentra fuertemente custodiada por Los Guardias reales. Valor del Viaje: 800 Monedas"
    ElseIf Val(i) = 3 Then
        Info.Caption = "Lindos es una Pequeña Ciudad, Con una gran variedad de casas, y una abadia. Valor del Viaje: 500 Monedas."
    ElseIf Val(i) = 4 Then
        Info.Caption = "Nix, es una Ciudad ubicada en un extremo del Mundo, Tiene un pequeño puerto Donde Nunca faltan pescadores. Valor del Viaje: 600 Monedas.."
    End If
End Function

'Boton Salir
Private Sub Exit_Click()
    Unload Me
End Sub

'Cerrar el Form
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        Unload Me
    End If
End Sub
