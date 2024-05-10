VERSION 5.00
Begin VB.Form frmConsultas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5100
   ClientLeft      =   75
   ClientTop       =   375
   ClientWidth     =   5130
   ControlBox      =   0   'False
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
   ScaleHeight     =   5100
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Peleas"
      Height          =   4935
      Left            =   10440
      TabIndex        =   6
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton Command14 
         Caption         =   "¿Online?"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command13 
         Caption         =   "¿Online?"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Caption         =   "<< Menos"
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
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Iniciar Pelea"
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
         TabIndex        =   10
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Ingresá el Nick del segundo Usuario"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Ingresá el Nick del primer Usuario"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "@"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   1920
         Width           =   255
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   4815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Preguntas de usuarios y respuestas"
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton Command8 
         Caption         =   "X"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         ToolTipText     =   "Borrar S.O.S Selecionado."
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Enviar Respuesta"
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         ToolTipText     =   "Enviar Respuesta"
         Top             =   4680
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   3720
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reponder al Usuario:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Info 
         AutoSize        =   -1  'True
         Caption         =   "Selecione un mensaje s.o.s, para ver el contenido ..."
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   4815
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Menu cSalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command7_Click()

On Error Resume Next

Call SendData("X" & EsUsuario & "*" & Text3.Text)

EsUsuario = ""

If List1.listIndex < 0 Then Exit Sub

List1.RemoveItem (List1.listIndex)

Info.Caption = "Selecione un mensaje s.o.s, para ver el contenido ..."

Call SendData("SOSDONE" & MensajesSOS(List1.listIndex + 1).Autor)

MensajesNumber = MensajesNumber - 1

Frame3.Caption = "Preguntas de usuarios y respuestas - (" & MensajesNumber & "/500)"

End Sub

Private Sub Command8_Click()

On Error Resume Next

If List1.listIndex < 0 Then Exit Sub

List1.RemoveItem (List1.listIndex)

Info.Caption = "Selecione un mensaje s.o.s, para ver el contenido ..."

Call SendData("SOSDONE" & MensajesSOS(List1.listIndex + 1).Autor)

MensajesNumber = MensajesNumber - 1

Frame3.Caption = "Preguntas de usuarios y respuestas - (" & MensajesNumber & "/500)"

End Sub

Private Sub cSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
If frmRendimiento.Transp.value = 0 Then
Else
    Transparencia Me.hWnd, 150
End If
Frame3.Caption = "Preguntas de usuarios y respuestas - (" & MensajesNumber & "/500)"

Dim ocx As Integer
For ocx = 1 To MensajesNumber

List1.AddItem "[" & MensajesSOS(ocx).Autor & "] - " & MensajesSOS(ocx).TIPO

Next ocx

End Sub

Private Sub List1_Click()

EsUsuario = MensajesSOS(List1.listIndex + 1).Autor

Info.Caption = "[" & MensajesSOS(List1.listIndex + 1).Autor & "] - " & MensajesSOS(List1.listIndex + 1).TIPO & " dice: " & vbNewLine & MensajesSOS(List1.listIndex + 1).Contenido

End Sub
