VERSION 5.00
Begin VB.Form frmWriteMSG 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ayuda GM"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Respuesta"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Consulta regular"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Value           =   -1  'True
      Width           =   1635
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Reportar Bug"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   1920
      TabIndex        =   6
      Top             =   1200
      Width           =   1515
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Sugerencia"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1920
      TabIndex        =   5
      Top             =   1680
      Width           =   1515
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Denunciar"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1515
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Descargo"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmWriteMSG.frx":0000
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   $"frmWriteMSG.frx":002E
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmWriteMSG.frx":013F
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmWriteMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If InStr(1, Text1.Text, ",") Then
    frmMensaje.Show
frmMensaje.msg.Caption = "Imposible mandar un mensaje GM con el signo ',' (Coma) dentro del mensaje, edita el mensaje y volve a enviarlo."
    Exit Sub
End If

If Text1.Text = "" Then
    Call frmMensaje.Show
frmMensaje.msg.Caption = "Debes escribir tu mensaje"
    Exit Sub
ElseIf DarIndiceElegido = -1 Then
        frmMensaje.Show
frmMensaje.msg.Caption = "Debes elegir el motivo de tu consulta"
    Exit Sub
Else
    Call SendData("#" & DarIndiceElegido & "," & Text1.Text)
    Debug.Print "Mande SOS"
    Unload Me
End If

End Sub

Private Function DarIndiceElegido() As Integer

Dim I As Integer

For I = 0 To 4
    If optConsulta(I).value = True Then
        DarIndiceElegido = I
        Exit Function
    End If
Next I

DarIndiceElegido = -1

End Function

Private Sub Command3_Click()
frmRespuestaGM.Show
Me.Hide
End Sub

Private Sub CommandButton1_Click()
frmManualseti.Show
End Sub

Private Sub Form_Load()
If TieneParaResponder = False Then
    Command3.Enabled = False
Else
    Command3.Enabled = True
End If

End Sub

Private Sub Text1_Click()
If Text1.Text = "Si mandas un GM inadecuado serás penado ..." & vbNewLine & "" Then
    Text1.Text = ""
End If
End Sub
