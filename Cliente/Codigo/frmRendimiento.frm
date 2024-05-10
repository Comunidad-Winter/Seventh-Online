VERSION 5.00
Begin VB.Form frmRendimiento 
   BackColor       =   &H80000012&
   Caption         =   "Rendimiento SeventhAO"
   ClientHeight    =   3075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Auris 
      BackColor       =   &H00000000&
      Caption         =   "Transparencias en  Auras"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   2040
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.CheckBox Temuestroelcartelitoarre 
      BackColor       =   &H00000000&
      Caption         =   "Form al Morir"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Bordes en Nombres"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.CheckBox TechosActivados 
      BackColor       =   &H80000007&
      Caption         =   "Techos con Transparencia"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox Nomedigas 
      BackColor       =   &H80000007&
      Caption         =   "Nombre de los objetos al pasar el mouse"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.CheckBox Transp 
      BackColor       =   &H80000007&
      Caption         =   "Forms con Transparencia"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmRendimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Visible = False
End Sub
