VERSION 5.00
Begin VB.Form frmCanjes 
   BorderStyle     =   0  'None
   Caption         =   "Canjes Nightmare"
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCanjes.frx":0000
   ScaleHeight     =   5985
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   3345
      Left            =   780
      TabIndex        =   5
      Top             =   840
      Width           =   1970
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   485
      Left            =   5280
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   600
      Width           =   485
   End
   Begin VB.Label lblptosdonacion 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Donaciones"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.Label DescripcionCanje 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   3720
      TabIndex        =   8
      Top             =   4620
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1800
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   360
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Hit 
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
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   2460
      Width           =   855
   End
   Begin VB.Label DM 
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
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   3375
      Width           =   855
   End
   Begin VB.Label lblptos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1845
      TabIndex        =   3
      Top             =   5340
      Width           =   735
   End
   Begin VB.Label RM 
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
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   2910
      Width           =   855
   End
   Begin VB.Label Puntos 
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
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   3855
      Width           =   855
   End
   Begin VB.Label Stats 
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
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   0
      Top             =   1965
      Width           =   855
   End
End
Attribute VB_Name = "frmCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Image1_Click()
If List1.Text = "Corona" Then
If MsgBox("�Est� seguro que desea canjear una Corona?", vbYesNo) = vbYes Then
Call SendData("CANJE01")
End If
End If
If List1.Text = "Manto Alado" Then
If MsgBox("�Est� seguro que desea canjear un Manto Alado?.", vbYesNo) = vbYes Then
Call SendData("CANJE02")
End If
End If
If List1.Text = "Manto Alado (E/G)" Then
If MsgBox("�Est� seguro que desea canjear un Manto Alado (E/G)?", vbYesNo) = vbYes Then
Call SendData("CANJE03")
End If
End If
If List1.Text = "T�nica Apocaliptica" Then
If MsgBox("�Est� seguro que desea canjear una T�nica Apocaliptica?", vbYesNo) = vbYes Then
Call SendData("CANJE04")
End If
End If
If List1.Text = "T�nica Apocaliptica (E/G)" Then
If MsgBox("�Est� seguro que desea canjear una Tunica Apocaliptica (E/G)?", vbYesNo) = vbYes Then
Call SendData("CANJE05")
End If
End If
If List1.Text = "B�culo Divino" Then
If MsgBox("�Est� seguro que desea canjear un B�culo Divino?", vbYesNo) = vbYes Then
Call SendData("CANJE06")
End If
End If
If List1.Text = "Espada Barlog" Then
If MsgBox("�Est� seguro que desea canjear una Espada Barlog?", vbYesNo) = vbYes Then
Call SendData("CANJE07")
End If
End If
If List1.Text = "Espada Argentum" Then
If MsgBox("�Est� seguro que desea canjear una Espada Argentum?", vbYesNo) = vbYes Then
Call SendData("CANJE08")
End If
End If
If List1.Text = "Daga Infernal" Then
If MsgBox("�Est� seguro que desea canjear una Daga Infernal?", vbYesNo) = vbYes Then
Call SendData("CANJE09")
End If
End If
If List1.Text = "Espada de las Almas" Then
If MsgBox("�Est� seguro que desea canjear una Espada de las Almas?", vbYesNo) = vbYes Then
Call SendData("CANJE10")
End If
End If
If List1.Text = "Arco �lfico" Then
If MsgBox("�Est� seguro que desea canjear un Arco �lfico?", vbYesNo) = vbYes Then
Call SendData("CANJE11")
End If
End If
If List1.Text = "Cetro Perfecto" Then
If MsgBox("�Est� seguro que desea canjear un Cetro Perfecto?", vbYesNo) = vbYes Then
Call SendData("CANJE12")
End If
End If
If List1.Text = "Armadura Ancestral" Then
If MsgBox("�Est� seguro que desea canjear una Armadura Ancestral?", vbYesNo) = vbYes Then
Call SendData("CANJE13")
End If
End If
If List1.Text = "Armadura Ancestral (E/G)" Then
If MsgBox("�Est� seguro que desea canjear una Armadura Ancestral (E/G)?", vbYesNo) = vbYes Then
Call SendData("CANJE14")
End If
End If
If List1.Text = "Coraza del Mal" Then
If MsgBox("�Est� seguro que desea canjear una Coraza del Mal?", vbYesNo) = vbYes Then
Call SendData("CANJE15")
End If
End If
If List1.Text = "Coraza del Mal (E/G)" Then
If MsgBox("�Est� seguro que desea canjear una Coraza del Mal (E/G)?", vbYesNo) = vbYes Then
Call SendData("CANJE16")
End If
End If
If List1.Text = "Armadura Diab�lica (E/G)" Then
If MsgBox("�Est� seguro que desea canjear una Armadura Diab�lica (E/G)?", vbYesNo) = vbYes Then
Call SendData("CANJE17")
End If
End If
If List1.Text = "Armadura Extrema (E/G)" Then
If MsgBox("�Est� seguro que desea canjear una Armadura Extrema (E/G)?", vbYesNo) = vbYes Then
Call SendData("CANJE18")
End If
End If
If List1.Text = "Anillo Divino" Then
If MsgBox("�Est� seguro que desea canjear un Anillo Divino?", vbYesNo) = vbYes Then
Call SendData("CANJE19")
End If
End If
If List1.Text = "Talisman del Lider" Then
If MsgBox("��Est� seguro que desea canjear un Talisman del Lider!?", vbYesNo) = vbYes Then
Call SendData("CANJE20")
End If
End If
If List1.Text = "Pendiente del Sacrificio" Then
If MsgBox("�Est� seguro que desea canjear un Pendiente del Sacrificio?", vbYesNo) = vbYes Then
Call SendData("CANJE21")
End If
End If
If List1.Text = "Escudo de Drag�n" Then
If MsgBox("�Est� seguro que desea canjear un Escudo de Drag�n?", vbYesNo) = vbYes Then
Call SendData("CANJE22")
End If
End If
If List1.Text = "Casco Siniestro" Then
If MsgBox("�Est� seguro que desea canjear un Casco Siniestro?", vbYesNo) = vbYes Then
Call SendData("CANJE23")
End If
End If
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If frmRendimiento.Transp.value = 0 Then
Else
    Transparencia Me.hwnd, 150
End If
Picture = LoadPicture(App.Path & "\Graficos\Interfaces\Canjes.jpg")
List1.AddItem "Corona"
List1.AddItem "Manto Alado"
List1.AddItem "Manto Alado (E/G)"
List1.AddItem "T�nica Apocaliptica"
List1.AddItem "T�nica Apocaliptica (E/G)"
List1.AddItem "B�culo Divino"
List1.AddItem "Espada Barlog"
List1.AddItem "Espada Argentum"
List1.AddItem "Daga Infernal"
List1.AddItem "Espada de las Almas"
List1.AddItem "Arco �lfico"
List1.AddItem "Cetro Perfecto"
List1.AddItem "Armadura Ancestral"
List1.AddItem "Armadura Ancestral (E/G)"
List1.AddItem "Coraza del Mal"
List1.AddItem "Coraza del Mal (E/G)"
List1.AddItem "Armadura Diab�lica (E/G)"
List1.AddItem "Armadura Extrema (E/G)"
List1.AddItem "Anillo Divino"
List1.AddItem "Escudo de Drag�n"
List1.AddItem "Casco Siniestro"
List1.AddItem "Talisman del Lider"
List1.AddItem "Pendiente del Sacrificio"
End Sub

Private Sub Label1_Click()
Call SendData("FEERMANDA")
frmCanjesDonacion.Show vbModal
End Sub

Private Sub lblptosdonacion_Click()
Call SendData("FEERMANDA")
frmCanjesDonacion.Show vbModal
Unload Me
End Sub

Private Sub List1_Click()
If List1.Text = "Corona" Then
Picture1.Picture = LoadPicture(DirGraficos & "13197.bmp")
Puntos.Caption = "100"
Stats.Caption = "8/8"
RM.Caption = "22/25"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Manto Alado" Then
Picture1.Picture = LoadPicture(DirGraficos & "13589.bmp")
Puntos.Caption = "200"
Stats.Caption = "45/50"
RM.Caption = "10/15"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Manto Alado (E/G)" Then
Picture1.Picture = LoadPicture(DirGraficos & "13607.bmp")
Puntos.Caption = "200"
Stats.Caption = "45/50"
RM.Caption = "10/15"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "T�nica Apocaliptica" Then
Picture1.Picture = LoadPicture(DirGraficos & "13615.bmp")
Puntos.Caption = "180"
Stats.Caption = "15/25"
RM.Caption = "15/20"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "T�nica Apocaliptica (E/G)" Then
Picture1.Picture = LoadPicture(DirGraficos & "13617.bmp")
Puntos.Caption = "180"
Stats.Caption = "15/25"
RM.Caption = "15/20"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "B�culo Divino" Then
Picture1.Picture = LoadPicture(DirGraficos & "13133.bmp")
Puntos.Caption = "120"
Stats.Caption = "N/A"
RM.Caption = "N/A"
DM.Caption = "37/40"
Hit.Caption = "N/A"
End If
If List1.Text = "Espada Barlog" Then
Picture1.Picture = LoadPicture(DirGraficos & "13111.bmp")
Puntos.Caption = "140"
Stats.Caption = "N/A"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "24/26"
End If
If List1.Text = "Espada Argentum" Then
Picture1.Picture = LoadPicture(DirGraficos & "13122.bmp")
Puntos.Caption = "90"
Stats.Caption = "N/A"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "22/25"
End If
If List1.Text = "Daga Infernal" Then
Picture1.Picture = LoadPicture(DirGraficos & "13103.bmp")
Puntos.Caption = "90"
Stats.Caption = "N/A"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "10/12"
End If
If List1.Text = "Espada de las Almas" Then
Picture1.Picture = LoadPicture(DirGraficos & "13139.bmp")
Puntos.Caption = "60"
Stats.Caption = "N/A"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "17/20"
End If
If List1.Text = "Arco �lfico" Then
Picture1.Picture = LoadPicture(DirGraficos & "13118.bmp")
Puntos.Caption = "120"
Stats.Caption = "N/A"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "15/17"
End If
If List1.Text = "Cetro Perfecto" Then
Picture1.Picture = LoadPicture(DirGraficos & "13116.bmp")
Puntos.Caption = "150"
Stats.Caption = "N/A"
RM.Caption = "N/A"
DM.Caption = "45/50"
Hit.Caption = "N/A"
End If
If List1.Text = "Armadura Ancestral" Then
Picture1.Picture = LoadPicture(DirGraficos & "13736.bmp")
Puntos.Caption = "150"
Stats.Caption = "67/70"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Armadura Ancestral (E/G)" Then
Picture1.Picture = LoadPicture(DirGraficos & "13795.bmp")
Puntos.Caption = "150"
Stats.Caption = "67/70"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Coraza del Mal" Then
Picture1.Picture = LoadPicture(DirGraficos & "13633.bmp")
Puntos.Caption = "190"
Stats.Caption = "69/72"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Coraza del Mal (E/G)" Then
Picture1.Picture = LoadPicture(DirGraficos & "13797.bmp")
Puntos.Caption = "190"
Stats.Caption = "69/72"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Armadura Diab�lica (E/G)" Then
Picture1.Picture = LoadPicture(DirGraficos & "13593.bmp")
Puntos.Caption = "230"
Stats.Caption = "74/76"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Armadura Extrema (E/G)" Then
Picture1.Picture = LoadPicture(DirGraficos & "13591.bmp")
Puntos.Caption = "210"
Stats.Caption = "73/74"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Anillo Divino" Then
Picture1.Picture = LoadPicture(DirGraficos & "13465.bmp")
Puntos.Caption = "150"
Stats.Caption = "N/A"
RM.Caption = "15/18"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Escudo de Drag�n" Then
Picture1.Picture = LoadPicture(DirGraficos & "13235.bmp")
Puntos.Caption = "110"
Stats.Caption = "19/24"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Casco Siniestro" Then
Picture1.Picture = LoadPicture(DirGraficos & "13179.bmp")
Puntos.Caption = "140"
Stats.Caption = "30/32"
RM.Caption = "2/5"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Talisman del Lider" Then
Picture1.Picture = LoadPicture(DirGraficos & "13463.bmp")
Puntos.Caption = "350"
Stats.Caption = "N/A"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If
If List1.Text = "Pendiente del Sacrificio" Then
Picture1.Picture = LoadPicture(DirGraficos & "13464.bmp")
Puntos.Caption = "50"
Stats.Caption = "N/A"
RM.Caption = "N/A"
DM.Caption = "N/A"
Hit.Caption = "N/A"
End If

'descripcion by zaiko

Select Case List1.Text
 
Case Is = "Corona"
       DescripcionCanje.Caption = "Corona que les otorgaban a los antiguos principes, muy efectiva para resistir ataques m�gicos."
Case Is = "Manto Alado"
       DescripcionCanje.Caption = "Los dioses griegos han creado este maravilloso item, con una muy buena defensa y resistencia m�gica, para luchar contra sus enemigos."
Case Is = "Manto Alado (E/G)"
       DescripcionCanje.Caption = "Los dioses griegos han creado este maravilloso item, con una muy buena defensa y resistencia m�gica, para luchar contra sus enemigos."
Case Is = "T�nica Apocaliptica"
       DescripcionCanje.Caption = "T�nica que contiene en sus telas fragmentos de fuego, por eso puede aguantar hasta los mas poderosos hechizos."
Case Is = "T�nica Apocaliptica (E/G)"
       DescripcionCanje.Caption = "T�nica que contiene en sus telas fragmentos de fuego, por eso puede aguantar hasta los mas poderosos hechizos."
Case Is = "B�culo Divino"
       DescripcionCanje.Caption = "Uno de los tantos b�culos que ha fundado el gran hechizero Gandalf, cuenta la leyenda que lo us� en algunas de sus historias."
Case Is = "Espada Barlog"
       DescripcionCanje.Caption = "La leyenda dice que fu� forjada en lo m�s profundo del inframundo con un mineral que nadie conoce, por eso la espada est� encendida en llamas."
Case Is = "Espada Argentum"
       DescripcionCanje.Caption = "Espada antes llamada, espada angelical, pero antiguos cl�rigos que se revelaron contra sus aliados la mancharon de sangre y destrucci�n."
Case Is = "Daga Infernal"
       DescripcionCanje.Caption = "La m�s filosa daga, forjada en el infierno, podr� atravesar hasta la mas dura armadura."
Case Is = "Espada de las Almas"
       DescripcionCanje.Caption = "En esta espada se encuentran las almas de los mas valientes guerreros que ocuparon estas tierras."
Case Is = "Arco �lfico"
       DescripcionCanje.Caption = "Arco fundado en la comarca de elfos del bosque, es el arco que m�s r�pido dispara sus flechas, eso causa un gran da�o a quien las reciba."
Case Is = "Cetro Perfecto"
       DescripcionCanje.Caption = "Domadores de la naturaleza crearon este item para domar hasta la m�s peligrosa criatura, tambi�n posee un excelente da�o m�gico."
Case Is = "Armadura Ancestral"
       DescripcionCanje.Caption = "Armadura ligera, los que la posean se sentir�n agusto con el peso, forjada con minerales de plata y oro."
Case Is = "Armadura Ancestral (E/G)"
       DescripcionCanje.Caption = "Armadura ligera, los que la posean se sentir�n agusto con el peso, forjada con minerales de plata y oro."
Case Is = "Coraza del Mal"
       DescripcionCanje.Caption = "Esta es una armadura que fue hecha con la piel del mas tenebroso drag�n que existi� en estas tierras."
Case Is = "Coraza del Mal (E/G)"
       DescripcionCanje.Caption = "Esta es una armadura que fue hecha con la piel del mas tenebroso drag�n que existi� en estas tierras."
Case Is = "Armadura Diab�lica (E/G)"
       DescripcionCanje.Caption = "Forjada con un mineral que todav�a nadie sabe de donde se sac�, su corteza es la mas gruesa de todas."
Case Is = "Armadura Extrema (E/G)"
       DescripcionCanje.Caption = "Para los r�pidos y veloces cazadores, el metal que tiene esta armadura es tan fuerte que puede aguantar hasta la m�s filosa daga."
Case Is = "Anillo Divino"
       DescripcionCanje.Caption = "Los magos antiguos se centraban en la resistencia m�gica, por eso crearon este poderoso anillo capaz de resistir fuertes ataques m�gicos."
Case Is = "Talisman del Lider"
       DescripcionCanje.Caption = "Talisman necesario para fundar clan."
Case Is = "Pendiente del Sacrificio"
       DescripcionCanje.Caption = "Con este item, al morir no se caer�n los items, solo puedes utilizarlo 1 vez."
Case Is = "Escudo de Drag�n"
       DescripcionCanje.Caption = "Lleva en sus puntas las garras del drag�n dorado, tambi�n est� compuesto por la piel del drag�n oscuro."
Case Is = "Casco Siniestro"
       DescripcionCanje.Caption = "Hecho con artes oscuras, el casco eleva la resistencia m�gica y la defensa."
       
End Select

End Sub
