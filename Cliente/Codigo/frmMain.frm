VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9015
   ClientLeft      =   360
   ClientTop       =   270
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":0CCA
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   7080
      Top             =   2520
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer tsControl 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   2520
   End
   Begin VB.Timer WorkMacro 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   6120
      Top             =   2520
   End
   Begin VB.Timer AntiExternos 
      Interval        =   64000
      Left            =   5640
      Top             =   2520
   End
   Begin VB.ListBox hlst 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2205
      Left            =   9000
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.PictureBox Minimap 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   6735
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   5
      Top             =   180
      Width           =   1500
      Begin VB.Shape UserM 
         BorderColor     =   &H000000C0&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   45
         Left            =   600
         Shape           =   3  'Circle
         Top             =   720
         Width           =   45
      End
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   120
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1800
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.Timer FPS 
      Interval        =   1000
      Left            =   4680
      Top             =   2520
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   4200
      Top             =   2520
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   5160
      Top             =   2520
   End
   Begin VB.TextBox SendCMSTXT 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   120
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1800
      Visible         =   0   'False
      Width           =   8160
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   3240
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6600
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RecTxt 
      CausesValidation=   0   'False
      Height          =   1500
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   165
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":309FD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2160
      Left            =   9000
      ScaleHeight     =   144
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   15
      Top             =   2880
      Width           =   2400
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   10320
      Top             =   8220
      Width           =   495
   End
   Begin VB.Image Image8 
      Height          =   375
      Left            =   10815
      Top             =   8400
      Width           =   495
   End
   Begin VB.Image Image7 
      Height          =   375
      Left            =   11355
      Top             =   8205
      Width           =   495
   End
   Begin VB.Image Image6 
      Height          =   360
      Left            =   10815
      Top             =   7920
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel Máximo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   9855
      TabIndex        =   27
      Top             =   1455
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image cmdQuests 
      Height          =   255
      Index           =   3
      Left            =   10275
      Top             =   7260
      Width           =   1575
   End
   Begin VB.Label rm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   570
      TabIndex        =   26
      Top             =   8730
      Width           =   735
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   10680
      Top             =   150
      Width           =   570
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   11280
      Top             =   165
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   2
      Left            =   10320
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   1
      Left            =   10320
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   10320
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   255
      Index           =   0
      Left            =   10320
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label HpBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8520
      TabIndex        =   18
      Top             =   7365
      Width           =   1395
   End
   Begin VB.Label StaBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8520
      TabIndex        =   20
      Top             =   6315
      Width           =   1395
   End
   Begin VB.Label ManaBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8520
      TabIndex        =   19
      Top             =   6825
      Width           =   1395
   End
   Begin VB.Label ItemName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9000
      TabIndex        =   25
      Top             =   5100
      Width           =   2295
   End
   Begin VB.Label Arma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4965
      TabIndex        =   24
      Top             =   8730
      Width           =   735
   End
   Begin VB.Label Escudo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   8730
      Width           =   735
   End
   Begin VB.Label Armadura 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1590
      TabIndex        =   22
      Top             =   8730
      Width           =   735
   End
   Begin VB.Label Casco 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   21
      Top             =   8730
      Width           =   735
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   150
      Left            =   8520
      Top             =   6840
      Width           =   1395
   End
   Begin VB.Shape Hpshp 
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   8490
      Top             =   7395
      Width           =   1410
   End
   Begin VB.Label Fuerza 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   17
      Top             =   8685
      Width           =   255
   End
   Begin VB.Label Agilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   8700
      Width           =   255
   End
   Begin VB.Image CmdLanzar 
      Height          =   300
      Left            =   9000
      MouseIcon       =   "frmMain.frx":30A7B
      MousePointer    =   99  'Custom
      Top             =   5220
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Image cmdInfo 
      Height          =   285
      Left            =   10560
      MouseIcon       =   "frmMain.frx":30BCD
      MousePointer    =   99  'Custom
      Top             =   5220
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   1
      Left            =   11400
      MouseIcon       =   "frmMain.frx":30D1F
      MousePointer    =   99  'Custom
      Top             =   3720
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   375
      Index           =   0
      Left            =   11400
      MouseIcon       =   "frmMain.frx":30E71
      MousePointer    =   99  'Custom
      Top             =   3360
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   10320
      MouseIcon       =   "frmMain.frx":30FC3
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   8640
      MouseIcon       =   "frmMain.frx":31115
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   2280
      Width           =   1485
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   10200
      TabIndex        =   11
      Top             =   1455
      Width           =   345
   End
   Begin VB.Label lblPorcLvl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[0%]"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   10080
      TabIndex        =   10
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9480
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label LvlLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8940
      TabIndex        =   8
      Top             =   1440
      Width           =   105
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8820
      TabIndex        =   7
      Top             =   1200
      Width           =   465
   End
   Begin VB.Shape ExpShp 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   9420
      Top             =   1470
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   9120
      TabIndex        =   6
      Top             =   690
      Width           =   2145
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "(000,00,00)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   150
      Left            =   8910
      TabIndex        =   3
      Top             =   8700
      Width           =   750
   End
   Begin VB.Label lblMapaName 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "mapa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8460
      TabIndex        =   4
      Top             =   8475
      Width           =   1575
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00404040&
      Height          =   6165
      Left            =   0
      Top             =   2400
      Width           =   8205
   End
   Begin VB.Image InvEqu 
      Height          =   3675
      Left            =   8640
      Top             =   2160
      Width           =   3195
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H00008080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      Height          =   150
      Left            =   8550
      Top             =   6330
      Width           =   1365
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Public InvMouseBoton As Long

Public InvMouseLanzar As Long

Public InvMousePantalla As Long

Private TiempoActual As Long
Private Contador As Integer

Public ActualSecond As Long
Public lastSecond As Long
Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long

Dim gDSB As DirectSoundBuffer
Dim gD As DSBUFFERDESC
Dim gW As WAVEFORMATEX
Dim gFileName As String
Dim dsE As DirectSoundEnum
Dim Pos(0) As DSBPOSITIONNOTIFY
Public IsPlaying As Byte

Dim endEvent As Long
Dim PuedeMacrear As Boolean


Implements DirectXEvent

Private Sub CmdLanzar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
InvMouseLanzar = Button
If Not GetAsyncKeyState(Button) < 0 Then InvMouseLanzar = 0
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)
If hlst.listIndex = -1 Then Exit Sub

Select Case Index
Case 0 'subir
    If hlst.listIndex = 0 Then Exit Sub
Case 1 'bajar
    If hlst.listIndex = hlst.ListCount - 1 Then Exit Sub
End Select

Call SendData("DESPHE" & Index + 1 & "," & hlst.listIndex + 1)

Select Case Index
Case 0 'subir
    hlst.listIndex = hlst.listIndex - 1
Case 1 'bajar
    hlst.listIndex = hlst.listIndex + 1
End Select

End Sub

Private Sub cmdQuests_Click(Index As Integer)
 Call SendData("QLR") 'Quest List Request
End Sub

Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)

End Sub

Private Sub CreateEvent()
     endEvent = DirectX.CreateEvent(Me)
End Sub

Public Sub ActivarMacroHechizos()
    If Not hlst.Visible Then
        Call AddtoRichTextBox(frmMain.RecTxt, "Debes tener seleccionado el hechizo para activar el auto-lanzar", 0, 200, 200, False, True, False)
        Exit Sub
    End If
    TrainingMacro.Interval = 2788
    TrainingMacro.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos activado", 0, 200, 200, False, True, False)
End Sub

Public Sub DesactivarMacroHechizos()
        TrainingMacro.Enabled = False
        SecuenciaMacroHechizos = 0
        Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, False)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
InvMousePantalla = Button
If Not GetAsyncKeyState(Button) < 0 Then InvMousePantalla = 0

    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If endEvent Then
        DirectX.DestroyEvent endEvent
    End If
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub FPS_Timer()

If PocionesZAO >= 15 Then MsgBox "¡Cheat Detectado!", vbCritical, "SeventhAO v3.0": End
PocionesZAO = 0

If logged And Not frmMain.Visible Then
    Unload frmConnect
    frmMain.Show
End If
    
End Sub

Private Sub Image2_Click()
Call SendData("ACTPT")
Call SendData("FEERMANDA")
frmCanjes.Show vbModal
End Sub


Private Sub Image6_Click()
If MsgBox("¿Esta seguro que desea viajar al Castillo Norte?", vbYesNo) = vbYes Then
Call SendData("/CASTILLO NORTE")
End If
End Sub

Private Sub Image7_Click()
If MsgBox("¿Esta seguro que desea viajar al Castillo Este?", vbYesNo) = vbYes Then
Call SendData("/CASTILLO ESTE")
End If
End Sub

Private Sub Image8_Click()
If MsgBox("¿Esta seguro que desea viajar al Castillo Sur?", vbYesNo) = vbYes Then
Call SendData("/CASTILLO SUR")
End If
End Sub

Private Sub Image9_Click()
If MsgBox("¿Esta seguro que desea viajar al Castillo Oeste?", vbYesNo) = vbYes Then
Call SendData("/CASTILLO OESTE")
End If
End Sub


Private Sub Label2_Click()
frmCanjes.Show vbModal
End Sub

Private Sub Image4_Click()
If MsgBox("¿Estas seguro que quieres salir?", vbYesNo) = vbYes Then
Call SendData("/SALIR")
End If
End Sub

Private Sub Image5_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Label5_Click()

End Sub

Private Sub lblMapaName_Click()
AddtoRichTextBox frmMain.RecTxt, "Este es el nombre del mapa en cual te encuentras.", 255, 255, 255, False, False, False
End Sub

Private Sub Macro_Timer()
    PuedeMacrear = True
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    SendData "LC" & tX & "," & tY
    SendData "/COMERCIAR"
End Sub

Private Sub mnuNpcDesc_Click()
    SendData "LC" & tX & "," & tY
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub Coord_Click()
    AddtoRichTextBox frmMain.RecTxt, "Estas coordenadas son tu ubicación en el mapa.", 255, 255, 255, False, False, False
End Sub


Private Sub picInv_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
InvMouseBoton = Button
If Not GetAsyncKeyState(Button) < 0 Then InvMouseBoton = 0
End Sub

Private Sub Second_Timer()
    ActualSecond = mid(Time, 7, 2)
    ActualSecond = ActualSecond + 1
    If ActualSecond = lastSecond Then End
    lastSecond = ActualSecond
    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub

Private Sub TirarItem()
    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            SendData "TI" & Inventario.SelectedItem & "," & 1
        Else
           If Inventario.Amount(Inventario.SelectedItem) > 1 Then
            frmCantidad.Show , frmMain
           End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    SendData "AG"
End Sub

Private Sub UsarItem()
    If TrainingMacro.Enabled Then DesactivarMacroHechizos
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then SendData "KLQ" & Inventario.SelectedItem: PocionesZAO = PocionesZAO + 1
End Sub

Private Sub EquiparItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        SendData "KHEV" & Inventario.SelectedItem
End Sub


''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub TrainingMacro_Timer()
    If Not hlst.Visible Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    If Comerciando Then Exit Sub
    Select Case SecuenciaMacroHechizos
        Case 0
            If hlst.List(hlst.listIndex) <> "(Nada)" And UserCanAttack = 1 Then
                Call SendData("DH" & hlst.listIndex + 1)
                Call SendData("YX" & Magia)
                'UserCanAttack = 0
            End If
            SecuenciaMacroHechizos = 1
        Case 1
            Call ConvertCPtoTP(MainViewShp.Left, MainViewShp.Top, MouseX, MouseY, tX, tY)
            If (UsingSkill = Magia Or UsingSkill = Proyectiles) And UserCanAttack = 0 Then Exit Sub
            SendData "WLC" & tX & "," & tY & "," & UsingSkill
            If UsingSkill = Magia Or UsingSkill = Proyectiles Then UserCanAttack = 0
            UsingSkill = 0
            SecuenciaMacroHechizos = 0
        Case Else
            DesactivarMacroHechizos
    End Select
    
End Sub


Private Sub cmdLanzar_Click()

 If InvMouseLanzar = 0 Then Exit Sub
    If hlst.List(hlst.listIndex) <> "(Nada)" And UserCanAttack = 1 Then
        Call SendData("DH" & hlst.listIndex + 1)
        Call SendData("YX" & Magia)
        UsaMacro = True

    End If
      InvMouseLanzar = 0
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub


Private Sub CmdInfo_Click()
    Call SendData("INFS" & hlst.listIndex + 1)
End Sub


Private Sub Form_Click()
Detectar RecTxt.hwnd, Me.hwnd
    If InvMousePantalla = 0 Then Exit Sub
    
    If Cartel Then Cartel = False

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(MouseBoton, True)
#End If

    If Not Comerciando Then
        Call ConvertCPtoTP(MainViewShp.Left, MainViewShp.Top, MouseX, MouseY, tX, tY)

        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                        If CnTd = 3 Then
                            SendData "UMH"
                            CnTd = 0
                        End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    SendData "LC" & tX & "," & tY
               
                Else
                    frmMain.MousePointer = vbDefault
                    If (UsingSkill = Magia Or UsingSkill = Proyectiles) And UserCanAttack = 0 Then Exit Sub
                    If TrainingMacro.Enabled Then DesactivarMacroHechizos
                    SendData "WLC" & tX & "," & tY & "," & UsingSkill
                    If UsingSkill = Magia Or UsingSkill = Proyectiles Then UserCanAttack = 0
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If MouseShift = vbLeftButton Then
                Call SendData("/TELEP YO " & UserMap & " " & tX & " " & tY)
            End If
        End If
    End If
        InvMousePantalla = 0
    
End Sub

Private Sub Form_DblClick()
 If picInv = False Then
    If Not frmForo.Visible Then
        SendData "RC" & tX & "," & tY
    End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
 
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
        
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    If Not Audio.PlayingMusic Then
                        Musica = True
                        Audio.PlayMIDI CStr(currentMidi) & ".mid"
                    Else
                        Musica = False
                        Audio.StopMidi
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatMode)
          
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMap)
                    Call frmMap.Show(vbModeless, frmMain)
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    Call SendData("YX" & Domar)
                
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    Call SendData("YX" & Robar)
                            
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    Call SendData("YX" & Ocultarse)
                
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                        Call UsarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If UserPuedeRefrescar Then
                        Call SendData("RPU")
                        UserPuedeRefrescar = False
                        Beep
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                     AddtoRichTextBox frmMain.RecTxt, "Para activar o desactivar el seguro tipea /SEG.", 255, 255, 255, False, False, False
                
                Case CustomKeys.BindedKey(eKeyType.mKeyFoto)
                     Dim I As Integer
For I = 1 To 1000
    If Not FileExist(App.Path & "\Fotos\Foto" & I & ".bmp", vbNormal) Then Exit For
           Next
        Call Capturar_Guardar(App.Path & "/Fotos/Foto" & I & ".bmp")
Call AddtoRichTextBox(frmMain.RecTxt, "Foto" & I & ".bmp Guardada en la Carpeta Fotos", 255, 150, 50, False, False, False)
                
                Case CustomKeys.BindedKey(eKeyType.mKeySeguroResu)
                    Call SendData("/SEGR")
                
                Case CustomKeys.BindedKey(eKeyType.mKeySeguroCvc)
                    Call SendData("/ircvc")
                    SeguroCvc = Not SeguroCvc
                Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)
                 If frmMain.WorkMacro.Enabled = True Then
                    frmMain.WorkMacro.Enabled = False
                   AddtoRichTextBox frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255
                Else
                    frmMain.WorkMacro.Enabled = True
                    AddtoRichTextBox frmMain.RecTxt, "Macro de Trabajo Activado.", 255, 255, 255
                End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeySeguroClan)
                    Call SendData("/SEGCLAN")
            End Select
        Else
 
        End If
    End If
    
    Select Case KeyCode
        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
                If SendTxt.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendCMSTXT.Visible = True
                    SendCMSTXT.SetFocus
                End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(vbModeless, frmMain)
        
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            Call SendData("/MEDITAR")
        
        Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)
                If TrainingMacro.Enabled Then
                    DesactivarMacroHechizos
                Else
                    ActivarMacroHechizos
                End If
                
        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
            Call SendData("/SALIR")
            
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
            If (UserCanAttack = 1) And _
                   (Not UserDescansar) And _
                   (Not UserMeditar) Then
                        SendData "AT"
                        UserCanAttack = 0
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
                If SendCMSTXT.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendTxt.Visible = True
                SendTxt.SetFocus
                End If
            
    End Select
End Sub


Private Sub Form_Load()

    Call SetWindowLong(RecTxt.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    
    Picture = LoadPicture(App.Path & "\Graficos\Interfaces\Principal.jpg")
    
    InvEqu.Picture = LoadPicture(App.Path & _
    "\Graficos\Interfaces\Inventario.jpg")
    
   Me.Left = 0
   Me.Top = 0

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MouseX = x
    MouseY = Y
    
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub Image1_Click(Index As Integer)
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

Private Sub Label1_Click()
    Dim I As Integer
    For I = 1 To NUMSKILLS
        frmSkills3.Text1(I).Caption = UserSkills(I)
    Next I
    Alocados = SkillPoints
    frmSkills3.Puntos.Caption = "Puntos:" & SkillPoints
    frmSkills3.Show , frmMain
End Sub

Private Sub Label4_Click()

    Call Audio.PlayWave("click.wav")

    InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Interfaces\Inventario.jpg")

    'DespInv(0).Visible = True
    'DespInv(1).Visible = True
    picInv.Visible = True

    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    
    ItemName.Visible = True
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
End Sub

Private Sub Label7_Click()
    
    Call Audio.PlayWave("click.wav")

    InvEqu.Picture = LoadPicture(App.Path & "\Graficos\Interfaces\Hechizos.jpg")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    'DespInv(0).Visible = False
    'DespInv(1).Visible = False
    picInv.Visible = False
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
    
    ItemName.Visible = False
    
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
End Sub

Private Sub picInv_DblClick()
If InvMouseBoton = 0 Then Exit Sub
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    Call UsarItem
     InvMouseBoton = 0
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call Audio.PlayWave("click.wav")
End Sub

Private Sub RecTxt_Change()
    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf Me.SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
    Else
      If (Not frmComerciar.Visible) And _
         (Not frmSkills3.Visible) And _
         (Not frmMSG.Visible) And _
         (Not frmForo.Visible) And _
         (Not frmEstadisticas.Visible) And _
         (Not frmCantidad.Visible) And _
         (picInv.Visible) Then
            picInv.SetFocus
      End If
    End If
    On Error GoTo 0
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim I As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For I = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, I, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next I
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If Left$(stxtbuffer, 1) = "/" Then
            If UCase(Left$(stxtbuffer, 8)) = "/PASSWD " Then
                    Dim j As String
#If SeguridadAlkon Then
                    j = md5.GetMD5String(Right$(stxtbuffer, Len(stxtbuffer) - 8))
                    Call md5.MD5Reset
#Else
                    j = (Right$(stxtbuffer, Len(stxtbuffer) - 8))
#End If
                    stxtbuffer = "/PASSWD " & j
            ElseIf UCase$(stxtbuffer) = "/HACERTORNEO" Then
                frmConsolaTorneo.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                Exit Sub
            ElseIf UCase$(stxtbuffer) = "/FUNDARCLAN" Then
                frmEligeAlineacion.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
            ElseIf UCase$(stxtbuffer) = "/GM" Then
                frmWriteMSG.Show , Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
            ElseIf UCase$(stxtbuffer) = "/PASSWD" Then
                FrmPass.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
            ElseIf UCase$(stxtbuffer) = "/VERQUEST" Then
            stxtbuffer = ""
            SendTxt.Text = ""
            Call SendData("QLR")
            KeyCode = 0
            SendTxt.Visible = False
                Exit Sub
            End If
                
            
            Call SendData(stxtbuffer)
    
       'Shout
        ElseIf Left$(stxtbuffer, 1) = "-" Then
            Call SendData("-" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

        'Whisper
        ElseIf Left$(stxtbuffer, 1) = "\" Then
            Call SendData("\" & Right$(stxtbuffer, Len(stxtbuffer) - 1))

'Borrar cartel
ElseIf stxtbuffer = "" Then
Call SendData(";" & " ")

ElseIf stxtbuffer = "" Then
Call SendData(";" & " ")
'***************
'*****kHALED****
'***************
'Funcion hace que borre mensaje con doble enter
        'Say
        ElseIf stxtbuffer <> "" Then
            Call SendData(";" & stxtbuffer)

        End If

        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
    End If
End Sub


Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call SendData("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
    End If
End Sub


Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub


Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.Text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub


''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then

Private Sub Socket1_Connect()

    Second.Enabled = True

    If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call SendData("FjCnXaKdNcZmS")

    ElseIf EstadoLogin = E_MODO.Normal Then
        Call SendData("FjCnXaKdNcZmS")

    ElseIf EstadoLogin = E_MODO.Dados Then
        Call SendData("FjCnXaKdNcZmS")


    End If
End Sub


Private Sub Socket1_Disconnect()
    Dim I As Long
    
    lastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For I = 0 To Forms.Count - 1
        If Forms(I).Name <> Me.Name And Forms(I).Name <> frmConnect.Name Then
            Unload Forms(I)
        End If
    Next I
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False
    
#If SegudidadAlkon Then
    LOGGING = False
    LOGSTRING = False
    LastPressed = 0
    LastMouse = False
    LastAmount = 0
#End If

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    UserPin = ""
    
    For I = 1 To NUMSKILLS
        UserSkills(I) = 0
    Next I

    For I = 1 To NUMATRIBUTOS
        UserAtributos(I) = 0
    Next I

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Response = 0
    lastSecond = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect
    
    If frmConnect.Visible Then
        frmConnect.Visible = False
    End If

    If Not frmCrearPersonaje.Visible Then
            frmConnect.Show
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer
    
    Socket1.Read RD, DataLength
    
    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        'Call LogCustom("HandleData: " & rBuffer(loopc))
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub


#End If

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).charindex > 0 Then
        If charlist(MapData(tX, tY).charindex).invisible = False Then
        
            Dim I As Long
            Dim m As New frmMenuseFashion
            
            Load m
            m.SetCallback Me
            m.SetMenuId 1
            m.ListaInit 2, False
            
            If charlist(MapData(tX, tY).charindex).Nombre <> "" Then
                m.ListaSetItem 0, charlist(MapData(tX, tY).charindex).Nombre, True
            Else
                m.ListaSetItem 0, "<NPC>", True
            End If
            m.ListaSetItem 1, "Comerciar"
            
            m.ListaFin
            m.Show , Me

        End If
    End If
End If

#End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId

Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call TirarItem
    Case 3 'Usar
        If Not NoPuedeUsar Then
            NoPuedeUsar = True
            Call UsarItem
        End If
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        SendData "LC" & tX & "," & tY
    Case 1 'Comerciar
        Call SendData("LC" & tX & "," & tY)
        Call SendData("/COMERCIAR")
    End Select
End Select
End Sub
'
' -------------------
'    W I N S O C K
' -------------------
'

Private Sub web_Click(Index As Integer)

End Sub

Private Sub tsControl_Timer()
Call ShControl
End Sub

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    Dim I As Long
    
    Debug.Print "WInsock Close"
    
    lastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For I = 0 To Forms.Count - 1
        If Forms(I).Name <> Me.Name And Forms(I).Name <> frmConnect.Name Then
            Unload Forms(I)
        End If
    Next I
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For I = 1 To NUMSKILLS
        UserSkills(I) = 0
    Next I

    For I = 1 To NUMATRIBUTOS
        UserAtributos(I) = 0
    Next I

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()

    Second.Enabled = True
    
  
   
     If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call SendData("FjCnXaKdNcZmS")

    ElseIf EstadoLogin = E_MODO.Normal Then
        Call SendData("FjCnXaKdNcZmS")
    ElseIf EstadoLogin = E_MODO.Dados Then
        Call SendData("FjCnXaKdNcZmS")
 End If
    
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer

    Debug.Print "Winsock DataArrival"
    
    'Socket1.Read RD, DataLength
    Winsock1.GetData RD

    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For loopc = 1 To Len(RD)

        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = loopc - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = loopc + 1
        End If

    Next loopc

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For loopc = 1 To CR
        Call HandleData(rBuffer(loopc))
    Next loopc
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    lastSecond = 0
    Second.Enabled = False

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    If frmOldPersonaje.Visible Then
        frmOldPersonaje.Visible = False
    End If

    If Not frmCrearPersonaje.Visible Then
        If Not frmBorrar.Visible And Not frmRecuperar.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

#End If


Private Sub Minimap_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then Call SendData("/TELEP YO " & UserMap & " " & CByte(x) & " " & CByte(Y))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then ' Al precionar esta tecla abre el form
       FrmEscape.Show
    End If
End Sub


Private Sub AntiExternos_Timer()

If FindWindow(vbNullString, UCase$("MACRO FOWL")) Then
    Call CheatExterno("MACRO FOWL")
 ElseIf FindWindow(vbNullString, UCase$("WPE PRO")) Then
   Call CheatExterno("WPE PRO")
    ElseIf FindWindow(vbNullString, UCase$("Macro B C")) Then
   Call CheatExterno("Macro B C")
    ElseIf FindWindow(vbNullString, UCase$("Macro BC")) Then
   Call CheatExterno("Macro BC")
ElseIf FindWindow(vbNullString, UCase$("MINI MACRO BY FOWL WWW.XTREME-ZONE.NET")) Then
    Call CheatExterno("MINI MACRO BY FOWL WWW.XTREME-ZONE.NET")
ElseIf FindWindow(vbNullString, UCase$("MACROSARAZA")) Then
    Call CheatExterno("MACROSARAZA")
ElseIf FindWindow(vbNullString, UCase$("Macroncmurd")) Then
    Call CheatExterno("Macroncmurd")
ElseIf FindWindow(vbNullString, UCase$("AUTOTRAINING")) Then
    Call CheatExterno("AUTOTRAINING")
ElseIf FindWindow(vbNullString, UCase$("0RK4M Version 1.5")) Then
    Call CheatExterno("0RK4M Version 1.5")
ElseIf FindWindow(vbNullString, UCase$("cmd")) Then
    Call CheatExterno("cmd")
ElseIf FindWindow(vbNullString, UCase$("X-Z MULTIMACRO VERSION II BY THEGABYX WWW.XTREME-ZONE.NET")) Then
    Call CheatExterno("X-Z MULTIMACRO VERSION II BY THEGABYX WWW.XTREME-ZONE.NET")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1.1")) Then
    Call CheatExterno("CHEAT ENGINE 5.1.1")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.0")) Then
    Call CheatExterno("CHEAT ENGINE 5.0")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 6.0")) Then
    Call CheatExterno("CHEAT ENGINE 6.0")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 6.1")) Then
    Call CheatExterno("CHEAT ENGINE 6.1")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.4")) Then
    Call CheatExterno("CHEAT ENGINE 5.4")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.5")) Then
    Call CheatExterno("CHEAT ENGINE 5.5")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.6")) Then
    Call CheatExterno("CHEAT ENGINE 5.6")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.7")) Then
    Call CheatExterno("CHEAT ENGINE 5.7")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.9")) Then
    Call CheatExterno("CHEAT ENGINE 5.9")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.8")) Then
    Call CheatExterno("CHEAT ENGINE 5.8")
ElseIf FindWindow(vbNullString, UCase$("Pts")) Then
    Call CheatExterno("Auto Pots")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.2")) Then
    Call CheatExterno("CHEAT ENGINE 5.2")
ElseIf FindWindow(vbNullString, UCase$("SoLocoVo?")) Then
    Call CheatExterno("SOLOCOVO?")
ElseIf FindWindow(vbNullString, UCase$("-=[ANUBYS RADAR]=-")) Then
    Call CheatExterno("-=[ANUBYS RADAR]=-")
ElseIf FindWindow(vbNullString, UCase$("CRAZY SPEEDER 1.05")) Then
    Call CheatExterno("CRAZY SPEEDER 1.05")
ElseIf FindWindow(vbNullString, UCase$("SET !XSPEED.NET")) Then
    Call CheatExterno("SET !XSPEED.NET")
ElseIf FindWindow(vbNullString, UCase$("SPEEDERXP V1.80 - UNREGISTERED")) Then
    Call CheatExterno("SPEEDERXP V1.80 - UNREGISTERED")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.3")) Then
    Call CheatExterno("CHEAT ENGINE 5.3")
ElseIf FindWindow(vbNullString, UCase$("CHEAT ENGINE 5.1")) Then
    Call CheatExterno("CHEAT ENGINE 5.1")
ElseIf FindWindow(vbNullString, UCase$("A SPEEDER")) Then
    Call CheatExterno("A SPEEDER")
ElseIf FindWindow(vbNullString, UCase$("MEMO :P")) Then
    Call CheatExterno("MEMO :P")
ElseIf FindWindow(vbNullString, UCase$("ORK4M VERSION 1.5")) Then
    Call CheatExterno("ORK4M VERSION 1.5")
ElseIf FindWindow(vbNullString, UCase$("BY FEDEX")) Then
    Call CheatExterno("By Fedex")
ElseIf FindWindow(vbNullString, UCase$("!XSPEED.NET +4.59")) Then
    Call CheatExterno("!Xspeeder")
ElseIf FindWindow(vbNullString, UCase$("CAMBIA TITULOS DE CHEATS BY FEDEX")) Then
    Call CheatExterno("Cambia titulos")
ElseIf FindWindow(vbNullString, UCase$("NEWENG OCULTO")) Then
    Call CheatExterno("NEWENG OCULTO")
ElseIf FindWindow(vbNullString, UCase$("SERBIO ENGINE")) Then
    Call CheatExterno("SERBIO ENGINE")
ElseIf FindWindow(vbNullString, UCase$("REYMIX ENGINE 5.3 PUBLIC")) Then
    Call CheatExterno("REYMIX ENGINE 5.3 PUBLIC")
ElseIf FindWindow(vbNullString, UCase$("REY ENGINE 5.2")) Then
    Call CheatExterno("REY ENGINE 5.2")
ElseIf FindWindow(vbNullString, UCase$("AUTOCLICK - BY NIO_SHOOTER")) Then
    Call CheatExterno("AUTOCLICK - BY NIO_SHOOTER")
ElseIf FindWindow(vbNullString, UCase$("TONNER MINER! :D [REG][SKLOV] 2.0")) Then
    Call CheatExterno("TONNER MINER! :D [REG][SKLOV] 2.0")
ElseIf FindWindow(vbNullString, UCase$("Buffy The vamp Slayer")) Then
    Call CheatExterno("Buffy The vamp Slayer")
ElseIf FindWindow(vbNullString, UCase$("Blorb Slayer 1.12.552 (BETA)")) Then
    Call CheatExterno("Blorb Slayer 1.12.552 (BETA)")
ElseIf FindWindow(vbNullString, UCase$("PumaEngine3.0")) Then
    Call CheatExterno("PumaEngine3.0")
ElseIf FindWindow(vbNullString, UCase$("Vicious Engine 5.0")) Then
    Call CheatExterno("Vicious Engine 5.0")
ElseIf FindWindow(vbNullString, UCase$("AkumaEngine33")) Then
    Call CheatExterno("AkumaEngine33")
ElseIf FindWindow(vbNullString, UCase$("Spuc3ngine")) Then
    Call CheatExterno("Spuc3ngine")
ElseIf FindWindow(vbNullString, UCase$("Ultra Engine")) Then
    Call CheatExterno("Ultra Engine")
ElseIf FindWindow(vbNullString, UCase$("Engine")) Then
    Call CheatExterno("Engine")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V5.6")) Then
    Call CheatExterno("Cheat Engine V5.6")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V5.5")) Then
    Call CheatExterno("Cheat Engine V5.5")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4")) Then
    Call CheatExterno("Cheat Engine V4.4")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.4 German Add-On")) Then
    Call CheatExterno("Cheat Engine V4.4 German Add-On")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.3")) Then
    Call CheatExterno("Cheat Engine V4.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.2")) Then
    Call CheatExterno("Cheat Engine V4.2")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V4.1.1")) Then
    Call CheatExterno("Cheat Engine V4.1.1")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.3")) Then
    Call CheatExterno("Cheat Engine V3.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.2")) Then
    Call CheatExterno("Cheat Engine V3.2")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine V3.1")) Then
    Call CheatExterno("Cheat Engine V3.1")
ElseIf FindWindow(vbNullString, UCase$("Cheat Engine")) Then
    Call CheatExterno("Cheat Engine")
ElseIf FindWindow(vbNullString, UCase$("danza engine 5.2.150")) Then
    Call CheatExterno("danza engine 5.2.150")
ElseIf FindWindow(vbNullString, UCase$("zenx engine")) Then
    Call CheatExterno("zenx engine")
ElseIf FindWindow(vbNullString, UCase$("MACROMAKER")) Then
    Call CheatExterno("Macro Maker")
ElseIf FindWindow(vbNullString, UCase$("MACREOMAKER - EDIT MACRO")) Then
    Call CheatExterno("Macro Maker")
ElseIf FindWindow(vbNullString, UCase$("By Fedex")) Then
    Call CheatExterno("Macro Fedex")
ElseIf FindWindow(vbNullString, UCase$("Macro Mage 1.0")) Then
    Call CheatExterno("Macro Mage")
ElseIf FindWindow(vbNullString, UCase$("Auto* v0.4 (c) 2001 Pete Powa")) Then
    Call CheatExterno("Macro Fisher")
ElseIf FindWindow(vbNullString, UCase$("Kizsada")) Then
    Call CheatExterno("Macro K33")
ElseIf FindWindow(vbNullString, UCase$("Makro K33")) Then
    Call CheatExterno("Macro K33")
ElseIf FindWindow(vbNullString, UCase$("Super Saiyan")) Then
    Call CheatExterno("El Chit del Geri")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete")) Then
    Call CheatExterno("Piringulete")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete 2003")) Then
    Call CheatExterno("Piringulete 2003")
ElseIf FindWindow(vbNullString, UCase$("TUKY2005")) Then
    Call CheatExterno("Makro Tuky")
ElseIf FindWindow(vbNullString, UCase$("Volks")) Then
    Call CheatExterno("TURBINAS")
ElseIf FindWindow(vbNullString, UCase$("Turbinas")) Then
    Call CheatExterno("TURBINAS")
ElseIf FindWindow(vbNullString, UCase$("msn")) Then
    Call CheatExterno("msn")
ElseIf FindWindow(vbNullString, UCase$("Volks")) Then
    Call CheatExterno("TURBINAS")
ElseIf FindWindow(vbNullString, UCase$("MacroSaraza[BETA]")) Then
    Call CheatExterno("MacroSaraza[BETA]")
ElseIf FindWindow(vbNullString, UCase$("Shell_TrayWnd")) Then
    Call CheatExterno("Shell_TrayWnd")
ElseIf FindWindow(vbNullString, UCase$("mmen")) Then
    Call CheatExterno("Cheat")
ElseIf FindWindow(vbNullString, UCase$("heat Celtic AO By Fowl")) Then
    Call CheatExterno("Cheat Celtic AO By Fowl")
ElseIf FindWindow(vbNullString, UCase$("Project1")) Then
    Call CheatExterno("Project1")
ElseIf FindWindow(vbNullString, UCase$("VB6")) Then
    Call CheatExterno("VB6")
ElseIf FindWindow(vbNullString, UCase$("Cheat_Celtic_AO_By_Fowl")) Then
    Call CheatExterno("Cheat_Celtic_AO_By_Fowl")
ElseIf FindWindow(vbNullString, UCase$("Auto Remo")) Then
    Call CheatExterno("Auto Remo")
ElseIf FindWindow(vbNullString, UCase$("Auto Remo")) Then
    Call CheatExterno("Auto Remo")
ElseIf FindWindow(vbNullString, UCase$("Auto Remo By Francohhh (www.neo-zone.activoforo.com)")) Then
    Call CheatExterno("Auto Remo By Francohhh (www.neo-zone.activoforo.com)")
ElseIf FindWindow(vbNullString, UCase$("Macro Configurable")) Then
    Call CheatExterno("Macro Configurable")
ElseIf FindWindow(vbNullString, UCase$("Mega Macro By Francohhh")) Then
    Call CheatExterno("Mega Macro By Francohhh")
ElseIf FindWindow(vbNullString, UCase$("MegaMacro By Francohhh (www.neo-zone.activoforo.com)")) Then
    Call CheatExterno("MegaMacro By Francohhh (www.neo-zone.activoforo.com)")
ElseIf FindWindow(vbNullString, UCase$("By FaKiTa!.-")) Then
    Call CheatExterno("By FaKiTa!.-")
ElseIf FindWindow(vbNullString, UCase$("Macro b53!")) Then
    Call CheatExterno("Macro b53!")
ElseIf FindWindow(vbNullString, UCase$("Borrar...")) Then
    Call CheatExterno("Borrar...")
ElseIf FindWindow(vbNullString, UCase$("Ares.exe")) Then
    Call CheatExterno("Ares.exe")
ElseIf FindWindow(vbNullString, UCase$("Crown Makro")) Then
    Call CheatExterno("Crown Makro")
ElseIf FindWindow(vbNullString, UCase$("AutoPots")) Then
    Call CheatExterno("AutoPots")
ElseIf FindWindow(vbNullString, UCase$("FaKiTa")) Then
    Call CheatExterno("AutoPots")
ElseIf FindWindow(vbNullString, UCase$("FaKiTa.-")) Then
    Call CheatExterno("AutoPots")
ElseIf FindWindow(vbNullString, UCase$("FaKiTa!.-")) Then
    Call CheatExterno("AutoPots")
ElseIf FindWindow(vbNullString, UCase$("msnmsgr")) Then
    Call CheatExterno("msnmsgr")
ElseIf FindWindow(vbNullString, UCase$("MacroSaraza1.3.3")) Then
    Call CheatExterno("MacroSaraza1.3.3")
ElseIf FindWindow(vbNullString, UCase$("MacroSaraza[BETA]")) Then
    Call CheatExterno("MacroSaraza[BETA]")
ElseIf FindWindow(vbNullString, UCase$("Macro-ilanchus")) Then
    Call CheatExterno("Macro-ilanchus")
ElseIf FindWindow(vbNullString, UCase$("MacroSaraza[BETA] ")) Then
    Call CheatExterno("MacroSaraza[BETA] ")
ElseIf FindWindow(vbNullString, UCase$("Autopotear")) Then
    Call CheatExterno("Autopotear")
ElseIf FindWindow(vbNullString, UCase$("MacroSaraza")) Then
    Call CheatExterno("MacroSaraza")
ElseIf FindWindow(vbNullString, UCase$("SpeederXP")) Then
    Call CheatExterno("SpeederXP")
ElseIf FindWindow(vbNullString, UCase$("MLEngine")) Then
    Call CheatExterno("MLEngine")
ElseIf FindWindow(vbNullString, UCase$("Makro-Piringulete")) Then
    Call CheatExterno("Makro-Piringulete")
ElseIf FindWindow(vbNullString, UCase$("MoonLight Engine 1129.1 by llvMoney A.K.A FaaF")) Then
    Call CheatExterno("MoonLight Engine 1129.1 by llvMoney A.K.A FaaF")
ElseIf FindWindow(vbNullString, UCase$("vb6")) Then
    Call CheatExterno("vb6")
ElseIf FindWindow(vbNullString, UCase$("VB6")) Then
    Call CheatExterno("VB6")
ElseIf FindWindow(vbNullString, UCase$("msmsgs")) Then
    Call CheatExterno("msmsgs")
ElseIf FindWindow(vbNullString, UCase$("Macro Magic")) Then
    Call CheatExterno("Macro Magic")
ElseIf FindWindow(vbNullString, UCase$("Iolo Macro Magic")) Then
    Call CheatExterno("Iolo Macro Magic")
ElseIf FindWindow(vbNullString, UCase$("AO Macro II 1.0.2")) Then
    Call CheatExterno("AO Macro II 1.0.2")
ElseIf FindWindow(vbNullString, UCase$("0rk4M")) Then
    Call CheatExterno("0rk4M")
ElseIf FindWindow(vbNullString, UCase$("AOFlechas")) Then
    Call CheatExterno("AOFlechas")
ElseIf FindWindow(vbNullString, UCase$("Auto remo By FaKiTa")) Then
    Call CheatExterno("Auto remo By FaKiTa")
ElseIf FindWindow(vbNullString, UCase$("AutoClick")) Then
    Call CheatExterno("AutoClick")
ElseIf FindWindow(vbNullString, UCase$("Borrar Cartel")) Then
    Call CheatExterno("Borrar Cartel")
ElseIf FindWindow(vbNullString, UCase$("Borrar Cartel 1.0 by BRASUkA!.-")) Then
    Call CheatExterno("Borrar Cartel 1.0 by BRASUkA!.-")
ElseIf FindWindow(vbNullString, UCase$("Cheat By The PePoH!")) Then
    Call CheatExterno("Cheat By The PePoH!")
ElseIf FindWindow(vbNullString, UCase$("Cheat By The PePoH!!!")) Then
    Call CheatExterno("Cheat By The PePoH!!!")
ElseIf FindWindow(vbNullString, UCase$("dddr")) Then
    Call CheatExterno("dddr")
ElseIf FindWindow(vbNullString, UCase$("Fedex")) Then
    Call CheatExterno("Fedex")
ElseIf FindWindow(vbNullString, UCase$("Flooder By FaKiTa")) Then
    Call CheatExterno("Flooder By FaKiTa")
ElseIf FindWindow(vbNullString, UCase$("Flooder")) Then
    Call CheatExterno("Flooder")
ElseIf FindWindow(vbNullString, UCase$("Full Cheat")) Then
    Call CheatExterno("Full Cheat")
ElseIf FindWindow(vbNullString, UCase$("Argentum-Pesca 0.2b Por Manchess")) Then
    Call CheatExterno("Argentum-Pesca 0.2b Por Manchess")
ElseIf FindWindow(vbNullString, UCase$("Macro_b53___By_Daaai")) Then
    Call CheatExterno("Macro_b53___By_Daaai")
ElseIf FindWindow(vbNullString, UCase$("MacroCrack")) Then
    Call CheatExterno("MacroCrack")
ElseIf FindWindow(vbNullString, UCase$("Macro-Resucitar")) Then
    Call CheatExterno("Macro-Resucitar")
ElseIf FindWindow(vbNullString, UCase$("Macro-Resucitar 1.0 | By Super Culd")) Then
    Call CheatExterno("Macro-Resucitar 1.0 | By Super Culd")
ElseIf FindWindow(vbNullString, UCase$("MakroK33")) Then
    Call CheatExterno("MakroK33")
ElseIf FindWindow(vbNullString, UCase$("Mega_Macro_By_Francohhh")) Then
    Call CheatExterno("Mega_Macro_By_Francohhh")
ElseIf FindWindow(vbNullString, UCase$("Contraseña")) Then
    Call CheatExterno("Contraseña")
ElseIf FindWindow(vbNullString, UCase$("MegaCheat")) Then
    Call CheatExterno("MegaCheat")
ElseIf FindWindow(vbNullString, UCase$("Eleji el cheat")) Then
    Call CheatExterno("Eleji el cheat")
ElseIf FindWindow(vbNullString, UCase$("Sacar letras hechiz By FaKiTa")) Then
    Call CheatExterno("Sacar letras hechiz By FaKiTa")
ElseIf FindWindow(vbNullString, UCase$("sh")) Then
    Call CheatExterno("sh")
ElseIf FindWindow(vbNullString, UCase$("Turbinas By Francohhh")) Then
    Call CheatExterno("Turbinas By Francohhh")
ElseIf FindWindow(vbNullString, UCase$("Auto Pots By Santeh")) Then
    Call CheatExterno("Auto Pots By Santeh")
ElseIf FindWindow(vbNullString, UCase$("ByAxeII")) Then
    Call CheatExterno("ByAxeII")
ElseIf FindWindow(vbNullString, UCase$("Cheat_By_Santeh_1.3")) Then
    Call CheatExterno("Cheat_By_Santeh_1.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat By Santeh 1.3")) Then
    Call CheatExterno("Cheat By Santeh 1.3")
ElseIf FindWindow(vbNullString, UCase$("Cheat 1.0 [By Santeh]")) Then
    Call CheatExterno("Cheat 1.0 [By Santeh]")
ElseIf FindWindow(vbNullString, UCase$("Auto_Floder__By_Santeh_")) Then
    Call CheatExterno("Auto_Floder__By_Santeh_")
ElseIf FindWindow(vbNullString, UCase$("Auto Floder [By Santeh]")) Then
    Call CheatExterno("Auto Floder [By Santeh]")
ElseIf FindWindow(vbNullString, UCase$("Cheat_By_Santeh_1.4")) Then
    Call CheatExterno("Cheat_By_Santeh_1.4")
ElseIf FindWindow(vbNullString, UCase$("Cheat By Santeh 1.4")) Then
    Call CheatExterno("Cheat By Santeh 1.4")
ElseIf FindWindow(vbNullString, UCase$("Macro  V1.0.0 - TheFranK - www.TheFranK-Cheats.com.ar")) Then
    Call CheatExterno("Macro  V1.0.0 - TheFranK - www.TheFranK-Cheats.com.ar")
ElseIf FindWindow(vbNullString, UCase$("!xSpeed.net -1.41")) Then
     Call CheatExterno("!xSpeed.net -1.41")
ElseIf FindWindow(vbNullString, UCase$("Ccleaner")) Then
     Call CheatExterno("Macro")
ElseIf FindWindow(vbNullString, UCase$("Ccleaner")) Then
     Call CheatExterno("Macro")
     ElseIf FindWindow(vbNullString, UCase$("CCLEANER")) Then
     Call CheatExterno("Macro")
ElseIf FindWindow(vbNullString, UCase$("Visual Basic 6.0")) Then
     Call CheatExterno("Visual Basic")
ElseIf FindWindow(vbNullString, UCase$("vb6")) Then
     Call CheatExterno("VB6")
ElseIf FindWindow(vbNullString, UCase$("Easy AO Makro - V 0.9 Beta")) Then
     Call CheatExterno("Easy AO Makro - V 0.9 Beta")
ElseIf FindWindow(vbNullString, UCase$("Piringulete")) Then
     Call CheatExterno("Piringulete")
ElseIf FindWindow(vbNullString, UCase$("MAKRO K33")) Then
     Call CheatExterno("MAKRO K33")
ElseIf FindWindow(vbNullString, UCase$("MAKRO-PIRINGULETE")) Then
     Call CheatExterno("MAKRO-PIRINGULETE")
ElseIf FindWindow(vbNullString, UCase$(".:::MAXICHIN")) Then
     Call CheatExterno(".:::MAXICHIN")
ElseIf FindWindow(vbNullString, UCase$("CHUPAS A LO PEDOS Y TE REMOVES VITH")) Then
     Call CheatExterno("CHUPAS A LO PEDOS Y TE REMOVES VITH")
ElseIf FindWindow(vbNullString, UCase$("A SPEEDER V2.1")) Then
     Call CheatExterno("A SPEEDER V2.1")
ElseIf FindWindow(vbNullString, UCase$("A SPEEDER")) Then
     Call CheatExterno("A SPEEDER")
ElseIf FindWindow(vbNullString, UCase$("SPEEDER - UNREGISTERED")) Then
     Call CheatExterno("SPEEDER - UNREGISTERED")
ElseIf FindWindow(vbNullString, UCase$("SPEEDERXP V1.60 - UNREGISTERED")) Then
     Call CheatExterno("SPEEDERXP V1.60 - UNREGISTERED")
ElseIf FindWindow(vbNullString, UCase$("SPEEDERXP V1.60 - REGISTERED")) Then
     Call CheatExterno("SPEEDERXP V1.60 - REGISTERED")
ElseIf FindWindow(vbNullString, UCase$("MACRO MAGE 1.0")) Then
     Call CheatExterno("MACRO MAGE 1.0")
ElseIf FindWindow(vbNullString, UCase$("AOITEMS - BY TAIKU - V1.0")) Then
     Call CheatExterno("AOITEMS - BY TAIKU - V1.0")
ElseIf FindWindow(vbNullString, UCase$("RADAR SILVERAO")) Then
     Call CheatExterno("RADAR SILVERAO")
ElseIf FindWindow(vbNullString, UCase$("MACRO 2005")) Then
     Call CheatExterno("MACRO 2005")
ElseIf FindWindow(vbNullString, UCase$("SPEEDER - REGISTERED")) Then
     Call CheatExterno("SPEEDER - REGISTERED")
ElseIf FindWindow(vbNullString, UCase$("PIRINGULETE")) Then
     Call CheatExterno("PIRINGULETE")
ElseIf FindWindow(vbNullString, UCase$("MACRO")) Then
     Call CheatExterno("MACRO")
ElseIf FindWindow(vbNullString, UCase$("MACRO-PIRINGULETE 2003")) Then
     Call CheatExterno("MACRO-PIRINGULETE 2003")
ElseIf FindWindow(vbNullString, UCase$("ARGENTUM FALSE V 0.2")) Then
     Call CheatExterno("ARGENTUM FALSE V 0.2")
ElseIf FindWindow(vbNullString, UCase$("SH")) Then
     Call CheatExterno("SH")
ElseIf FindWindow(vbNullString, UCase$("SPEEDER")) Then
     Call CheatExterno("SPEEDER")
ElseIf FindWindow(vbNullString, UCase$("SPEED")) Then
     Call CheatExterno("SPEED")
ElseIf FindWindow(vbNullString, UCase$("KORVEN")) Then
     Call CheatExterno("KORVEN")
ElseIf FindWindow(vbNullString, UCase$("EASY AO MAKRO - V 0.9 BETA")) Then
     Call CheatExterno("EASY AO MAKRO - V 0.9 BETA")
ElseIf FindWindow(vbNullString, UCase$("SOLOCOVO  ?")) Then
     Call CheatExterno("SOLOCOVO  ?")
ElseIf FindWindow(vbNullString, UCase$("CHITEO")) Then
     Call CheatExterno("CHITEO")
ElseIf FindWindow(vbNullString, UCase$("MacroCrack <gonza_vi@hotmail.com>")) Then
     Call CheatExterno("MacroCrack <gonza_vi@hotmail.com>")
ElseIf FindWindow(vbNullString, UCase$("MacroCrack <gonza_vi@hotmail.com> ")) Then
     Call CheatExterno("MacroCrack <gonza_vi@hotmail.com>")
End If
 
End Sub

Private Sub WorkMacro_Timer()

If Me.ItemName.Caption = "Hacha de Leñador" Or Me.ItemName.Caption = "Piquete de Minero" Or Me.ItemName.Caption = "Caña de Pescar" Then
    SendData "KLQ" & Inventario.SelectedItem
    SendData "WLC" & tX & "," & tY & "," & UsingSkill
Else
    AddtoRichTextBox frmMain.RecTxt, "No Puedes Usar el Macro Con Este item!", 255, 255, 255
    frmMain.WorkMacro.Enabled = False
    AddtoRichTextBox frmMain.RecTxt, "Macro de Trabajo Desactivado.", 255, 255, 255
    Exit Sub
End If

End Sub
