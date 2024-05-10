VERSION 5.00
Begin VB.Form FrmEscape 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmEscape.frx":0000
   ScaleHeight     =   2640
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image4 
      Height          =   375
      Index           =   2
      Left            =   0
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Image Image4 
      Height          =   375
      Index           =   1
      Left            =   0
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Image Image4 
      Height          =   375
      Index           =   0
      Left            =   0
      Top             =   960
      Width           =   3735
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Top             =   2280
      Width           =   3735
   End
End
Attribute VB_Name = "FrmEscape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Picture = LoadPicture(App.Path & "\Graficos\Interfaces\Opciones.jpg")
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    HookSurfaceHwnd Me
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image2_Click()
Call SendData("/SALIR")
End Sub

Private Sub Image3_Click()
Call SendData("/SALIR")
End Sub

Private Sub Image4_Click(Index As Integer)
    Call Audio.PlayWave("click.wav")

    Select Case Index
        Case 0
            '[MatuX] : 01 de Abril del 2002
                Call frmOpciones.Show(vbModeless, frmMain)
            '[END]
        Case 1
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            SendData "ATRI"
            SendData "ESKI"
            SendData "FEST"
            SendData "YGIJ"
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
        Case 2
            If Not frmGuildLeader.Visible Then _
                Call SendData("GLINFO")
    End Select
End Sub

