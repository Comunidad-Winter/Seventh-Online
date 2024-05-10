VERSION 5.00
Begin VB.Form frmPanelTorneo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Panel Torneo"
   ClientHeight    =   2580
   ClientLeft      =   105
   ClientTop       =   375
   ClientWidth     =   5145
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Iniciar cuenta regresiva para las inscripciones ..."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Crear Torneo / Evento"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.ListBox Evento_Faccionario 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   705
         Left            =   1560
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   1310
         Width           =   3255
      End
      Begin VB.TextBox Nivel_Minimo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Text            =   "1"
         Top             =   960
         Width           =   495
      End
      Begin VB.ListBox Cupos_Maximos 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   1560
         TabIndex        =   3
         Top             =   680
         Width           =   495
      End
      Begin VB.ListBox Torneo_Tipo 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label4 
         Caption         =   "« Clickeà antes de empezar ..."
         Height          =   255
         Left            =   2080
         TabIndex        =   9
         Top             =   675
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Faccionario"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   1245
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel Mìnimo"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cupos Màximos"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   680
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmPanelTorneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command6_Click()
If Val(Cupos_Maximo) > 0 Then
MsgBox "Elija un maximo de cupos por torneo.", vbCritical, "Error #5"
Exit Sub
End If


            Torneos.Cupos = Val(Cupos_Maximos)
Torneos.Nivel = Val(Nivel_Minimo.Text)
Select Case Torneo_Tipo
    Case Is = "Deathmacht - Torneo todos contra todos."
        Torneos.TIPO = "DM"
    Case Is = "Simple Fight - Torneo usuario versus usuario."
        Torneos.TIPO = "1V1"
    Case Is = "Partner Fight - Torneo en pareja (2 versus 2)"
        Torneos.TIPO = "2V2"
    Case Is = "Trilpe Trouble - Torneo en parejas de 3 versus 3"
        Torneos.TIPO = "3V3"
    Case Is = "Guilds Event - Torneo de Clanes"
        Torneos.TIPO = "CE"
    Case Else
        Call MsgBox("Error #3 - Elije un tipo de torneo, P.D.: No olvides de clickearlo.", vbInformation, "Error #3")
    Exit Sub
End Select
Call SendData("/TOR " & Torneos.TIPO & "," & Torneos.Cupos & "," & Torneos.Nivel & "," & Torneos.AutoSum & "," & Torneos.m & "," & Torneos.X & "," & Torneos.Y)
End Sub

Private Sub Form_Load()
'TORNEOS
Torneo_Tipo.AddItem "Deathmacht - Torneo todos contra todos."
Torneo_Tipo.AddItem "Simple Fight - Torneo usuario versus usuario."
Torneo_Tipo.AddItem "Partner Fight - Torneo en pareja (2 versus 2)"
Torneo_Tipo.AddItem "Trilpe Trouble - Torneo en parejas de 3 versus 3"
Torneo_Tipo.AddItem "Guilds Event - Torneo de Clanes"
Cupos_Maximos.AddItem "8"
Cupos_Maximos.AddItem "16"
Cupos_Maximos.AddItem "32"
Cupos_Maximos.AddItem "64"
End Sub
