VERSION 5.00
Begin VB.Form frmNobleza 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Noble"
   ClientHeight    =   7380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3910
      Picture         =   "frmNobleza.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   4100
      Width           =   480
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1170
      Picture         =   "frmNobleza.frx":0844
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   4100
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3960
      Picture         =   "frmNobleza.frx":1088
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   630
      Width           =   480
   End
   Begin VB.ListBox lstReq 
      BackColor       =   &H00000040&
      ForeColor       =   &H00C0FFFF&
      Height          =   2010
      Index           =   3
      ItemData        =   "frmNobleza.frx":18CC
      Left            =   3050
      List            =   "frmNobleza.frx":190C
      TabIndex        =   4
      Top             =   4680
      Width           =   2260
   End
   Begin VB.ListBox lstReq 
      BackColor       =   &H00000040&
      ForeColor       =   &H00C0FFFF&
      Height          =   2010
      Index           =   2
      ItemData        =   "frmNobleza.frx":1957
      Left            =   280
      List            =   "frmNobleza.frx":1997
      TabIndex        =   3
      Top             =   4680
      Width           =   2260
   End
   Begin VB.ListBox lstReq 
      BackColor       =   &H00000040&
      ForeColor       =   &H00C0FFFF&
      Height          =   2010
      Index           =   1
      ItemData        =   "frmNobleza.frx":19E2
      Left            =   3050
      List            =   "frmNobleza.frx":1A22
      TabIndex        =   2
      Top             =   1200
      Width           =   2260
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1180
      Picture         =   "frmNobleza.frx":1A6D
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   630
      Width           =   480
   End
   Begin VB.ListBox lstReq 
      BackColor       =   &H00000040&
      ForeColor       =   &H00C0FFFF&
      Height          =   2010
      Index           =   0
      ItemData        =   "frmNobleza.frx":22B1
      Left            =   280
      List            =   "frmNobleza.frx":22F1
      TabIndex        =   0
      Top             =   1200
      Width           =   2260
   End
   Begin VB.Image Image4 
      Height          =   435
      Left            =   3000
      Top             =   6720
      Width           =   2340
   End
   Begin VB.Image Image3 
      Height          =   435
      Left            =   240
      Top             =   6720
      Width           =   2340
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   240
      Top             =   3240
      Width           =   2340
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   3000
      Top             =   3240
      Width           =   2340
   End
   Begin VB.Image ImageSalir 
      Height          =   375
      Left            =   5280
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmNobleza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()
SendData ("/ARMADURA")
End Sub

Private Sub Image1_Click()
SendData ("/ESPADA")
End Sub

Private Sub Image3_Click()
SendData ("/ESCUDO")
End Sub

Private Sub Image4_Click()
SendData ("/ANILLO")
End Sub

Private Sub form_load()
Me.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_main.jpg")
Image1.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruirn.jpg")
Image2.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruir2n.jpg")
Image3.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruir4n.jpg")
Image4.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruir3n.jpg")

If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))

End Sub

Private Sub ImageSalir_Click()
Unload Me
End Sub

Private Sub image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruira.jpg")
End Sub

Private Sub image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruiri.jpg")
End Sub

Private Sub image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruir2a.jpg")
End Sub

Private Sub image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruir2i.jpg")
End Sub

Private Sub image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruir4i.jpg")
End Sub

Private Sub image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruir4a.jpg")
End Sub

Private Sub image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruir3i.jpg")
End Sub

Private Sub image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruir3a.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruirn.jpg")
Image2.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruir2n.jpg")
Image3.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruir4n.jpg")
Image4.Picture = LoadPicture(DirInterfaces & "Principal\Nobleza_bconstruir3n.jpg")
End Sub

Private Sub lstReq_Click(Index As Integer)

End Sub
