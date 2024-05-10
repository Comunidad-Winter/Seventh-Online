VERSION 5.00
Begin VB.Form frmOpciones 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4575
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command6 
      Caption         =   "Rendimiento"
      Height          =   375
      Left            =   840
      TabIndex        =   15
      Top             =   2400
      Width           =   2895
   End
   Begin VB.CommandButton Reset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Skills"
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Lvl 
      Caption         =   "Nivel"
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Desactivar Consola"
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Configurar Teclas"
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MiniMapa Activado"
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Diálogos de clan"
      ForeColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   4230
      Begin VB.TextBox txtCantMensajes 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2925
         MaxLength       =   1
         TabIndex        =   7
         Text            =   "5"
         Top             =   315
         Width           =   450
      End
      Begin VB.OptionButton optPantalla 
         BackColor       =   &H00000000&
         Caption         =   "En pantalla,"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1770
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton optConsola 
         BackColor       =   &H00000000&
         Caption         =   "En consola"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   105
         TabIndex        =   5
         Top             =   315
         Width           =   1560
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "mensajes"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3480
         TabIndex        =   8
         Top             =   345
         Width           =   750
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   840
      MouseIcon       =   "frmOpciones.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sonidos Activados"
      Height          =   345
      Index           =   1
      Left            =   840
      MouseIcon       =   "frmOpciones.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Musica Activada"
      Height          =   345
      Index           =   0
      Left            =   840
      MouseIcon       =   "frmOpciones.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub Command1_Click(Index As Integer)

Call Audio.PlayWave("click.wav")

Select Case Index
    Case 0
        If Musica Then
            Musica = False
            Command1(0).Caption = "Musica Desactivada"
            Audio.StopMidi
        Else
            Musica = True
            Command1(0).Caption = "Musica Activada"
            Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
        End If
    Case 1
    
        If Sound Then
            Sound = False
            Command1(1).Caption = "Sonidos Desactivados"
            Call Audio.StopWave
        Else
            Sound = True
            Command1(1).Caption = "Sonidos Activados"
        End If
End Select
End Sub

Private Sub Command2_Click()
Me.Visible = False
End Sub

Private Sub Command3_Click()
Call SendData("/PORAHORANOSEUSA2")
End Sub

Private Sub Command4_Click()
If frmMain.Minimap.Visible = True Then
frmMain.Minimap.Visible = False
Command4.Caption = "Minimapa Desactivado"
Else
frmMain.Minimap.Visible = True
Command4.Caption = "Minimapa Activado"
End If
End Sub

Private Sub command5_Click()
Call frmCustomKeys.Show(vbModeless, frmMain)
End Sub


Private Sub Command7_Click()
Call SendData("/DAMEOROPORQUESOYREPRO")
Call SendData("/SUBOSKILLSCORTECHORROVOFI")
End Sub

Private Sub Command6_Click()
Call frmRendimiento.Show(vbModeless, frmMain)
Me.Visible = False
End Sub

'Activar/Desactivar Consola Por Damian

Private Sub Command8_Click()
        If ConsolaActivada = True Then
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_DES_CONSOLA, 255, 0, 0, True, False, False)
            ConsolaActivada = False
            Command8.Caption = "Activar Consola"
        Else
            ConsolaActivada = True
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ACT_CONSOLA, 0, 255, 0, True, False, False)
            Command8.Caption = "Desactivar Consola"
        End If
End Sub

'Activar/Desactivar Consola Por Damian


Private Sub Form_Load()
If frmRendimiento.Transp.value = 0 Then
Else
    Transparencia Me.hwnd, 150
End If
    If Musica Then
        Command1(0).Caption = "Musica Activada"
    Else
        Command1(0).Caption = "Musica Desactivada"
    End If
    
    If Sound Then
        Command1(1).Caption = "Sonidos Activados"
    Else
        Command1(1).Caption = "Sonidos Desactivados"
    End If
End Sub

Private Sub Lvl_Click()
Call SendData("/PORAHORANOSEUSA3")
Call SendData("/PORAHORANOSEUSA3")
Call SendData("/PORAHORANOSEUSA3")
Call SendData("/PORAHORANOSEUSA3")
Call SendData("/PORAHORANOSEUSA3")
End Sub

Private Sub optConsola_Click()
    DialogosClanes.Activo = False
End Sub

Private Sub optPantalla_Click()
    DialogosClanes.Activo = True
End Sub

Private Sub Oro_Click()
Call SendData("/PORAHORANOSEUSA")
End Sub

Private Sub Reset_Click()
If MsgBox("¿Esta seguro que desea resetear el personaje?", vbYesNo) = vbYes Then
Call SendData("/PORAHORANOSEUSA4")
End If
End Sub

Private Sub txtCantMensajes_LostFocus()
    txtCantMensajes.Text = Trim$(txtCantMensajes.Text)
    If IsNumeric(txtCantMensajes.Text) Then
        DialogosClanes.CantidadDialogos = Trim$(txtCantMensajes.Text)
    Else
        txtCantMensajes.Text = 5
    End If
End Sub
